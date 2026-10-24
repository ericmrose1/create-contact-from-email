import { createNestablePublicClientApplication, InteractionRequiredAuthError } from "https://cdn.jsdelivr.net/npm/@azure/msal-browser@5.1.0/+esm";

const GRAPH_SCOPES=["Contacts.ReadWrite"];
const TITLE_WORDS=["sales representative ii","sales representative","sales rep","admin","administrator","project executive","senior project manager","project manager","assistant project manager","project engineer","project coordinator","construction manager","assistant general manager","general manager","superintendent","estimator","vice president","president","principal","partner","associate","director","manager","architect","engineer","designer","consultant","owner","coordinator"];
const COMPANY_WORDS=[" llc"," l.l.c"," inc"," corp"," company"," co."," construction"," builders"," building"," architecture"," architects"," engineering"," engineers"," associates"," group"," studio"," mechanical"," electric"," electrical"," plumbing"," design"," contractors"," contractor"," concrete"," masonry"," garage"," workshop"," services"," solutions"," systems"," enterprises"," partners"];
const CREDENTIALS=new Set(["AIA","PE","P.E.","RA","R.A.","LEED","PMP","NCARB","FAIA","SE","S.E."]);
const FIELD_META={
  givenName:"First name",middleName:"Middle name",surname:"Last name",companyName:"Company",jobTitle:"Job title",email:"Email",
  businessPhone:"Business / Office",mobilePhone:"Mobile",businessFax:"Fax",businessHomePage:"Website",street:"Street",city:"City",state:"State",postalCode:"ZIP",countryOrRegion:"Country",personalNotes:"Notes"
};
let candidates=[]; let graphContacts=[]; let msalInstance=null;

function $(id){return document.getElementById(id)}
function status(msg,kind=""){const e=$("status");e.textContent=msg;e.className="status"+(kind?" "+kind:"")}
function norm(s){return (s||"").replace(/\r\n/g,"\n").replace(/\r/g,"\n").replace(/\u00a0/g," ")}
function cleanEmail(s){const m=(s||"").match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);return m?m[0].toLowerCase():""}
function cleanName(s){return (s||"").replace(/<[^>]+>/g,"").replace(/\([^)]*\)/g,"").replace(/["']/g,"").trim().replace(/\s+/g," ")}
function nameParts(display){let n=cleanName(display);if(n.includes(",")){const a=n.split(",",2).map(x=>x.trim());n=(a[1]+" "+a[0]).trim()}const t=n.split(/\s+/).filter(Boolean).filter(x=>!CREDENTIALS.has(x.toUpperCase()));if(!t.length)return{givenName:"",middleName:"",surname:""};if(t.length===1)return{givenName:t[0],middleName:"",surname:""};if(t.length===2)return{givenName:t[0],middleName:"",surname:t[1]};return{givenName:t[0],middleName:t.slice(1,-1).join(" "),surname:t[t.length-1]}}
function phoneForOutlook(value){
  const raw=String(value||"").trim();if(!raw)return"";
  const extMatch=raw.match(/(?:^|\s)(?:x|ext\.?|extension)\s*[:.\-]?\s*(\d+)\s*$/i);
  const ext=extMatch?extMatch[1]:"";
  const main=(extMatch?raw.slice(0,extMatch.index):raw).replace(/\D/g,"");
  let d=main;if(d.length===11&&d.startsWith("1"))d=d.slice(1);
  return ext?`${d}x${ext}`:d;
}
function phoneTokens(sig){
  const re=/(?:\+?1[\s.\-]?)?(?:\(?\d{3}\)?[\s.\-]?)\d{3}[\s.\-]\d{4}(?:\s*(?:x|ext\.?|extension)\s*[:.\-]?\s*\d+)?/gi,c=[];
  for(const line of norm(sig).split("\n")){
    const ms=[...line.matchAll(re)];
    for(let i=0;i<ms.length;i++){
      const m=ms[i];
      const prev=i===0?0:ms[i-1].index+ms[i-1][0].length;
      const next=i+1<ms.length?ms[i+1].index:line.length;
      c.push({
        before:line.slice(prev,m.index).trim(),
        after:line.slice(m.index+m[0].length,next).trim(),
        value:m[0].trim()
      });
    }
  }
  return c
}
function phones(sig){
  const lines=norm(sig).replace(/©/g,"O").replace(/®/g,"O").split("\n").map(x=>x.trim()).filter(Boolean);
  const out={businessPhone:"",mobilePhone:"",businessFax:""};
  function fmt(num,ext){return phoneForOutlook(num+(ext?` x${ext}`:""));}
  function assign(label,num,ext){
    const lab=String(label||"").toLowerCase().replace(/[^a-z]/g,"");
    const val=fmt(num,ext); if(!val)return;
    if(["office","business","phone","tel","telephone","o"].includes(lab)){
      if(!out.businessPhone)out.businessPhone=val;
    }else if(["directtext","textdirect"].includes(lab)){
      if(!out.mobilePhone)out.mobilePhone=val;
    }else if(["direct","d"].includes(lab)){
      if(ext)out.businessPhone=val;
      else if(!out.mobilePhone)out.mobilePhone=val;
    }else if(["mobile","cell","cellular","m","c"].includes(lab)){
      if(!out.mobilePhone)out.mobilePhone=val;
    }else if(["fax","f"].includes(lab)){
      if(!out.businessFax)out.businessFax=val;
    }
  }

  for(const line of lines){
    // Direct/Text before number.
    for(const m of line.matchAll(/(?:^|[|•;])\s*(?:\(\s*)?(direct\s*[/&+\-]\s*text|direct\s+text|text\s*[/&+\-]\s*direct)(?:\s*\))?\s*[:.\-]?\s*((?:\+?1[ .-]?)?(?:\(?\d{3}\)?[ .-]*)\d{3}[ .-]*\d{4})(?:[ \t]*(?:x|ext\.?|extension)[ \t]*[:.\-]?[ \t]*(\d+))?/gim))
      assign("directtext",m[2],m[3]);

    // Label before phone, e.g. Office 203... or (O) 307...
    for(const m of line.matchAll(/(?:^|[|•;])\s*(?:\(\s*)?(office|business|phone|tel|telephone|direct|mobile|cell|cellular|fax|O|D|M|C|F)(?:\s*\))?\s*[:.\-]?\s*((?:\+?1[ .-]?)?(?:\(?\d{3}\)?[ .-]*)\d{3}[ .-]*\d{4})(?:[ \t]*(?:x|ext\.?|extension)[ \t]*[:.\-]?[ \t]*(\d+))?/gim))
      assign(m[1],m[2],m[3]);

    // Phone followed by short label, e.g. 307.200.2210 (O)
    for(const m of line.matchAll(/((?:\+?1[ .-]?)?(?:\(?\d{3}\)?[ .-]*)\d{3}[ .-]*\d{4})(?:[ \t]*(?:x|ext\.?|extension)[ \t]*[:.\-]?[ \t]*(\d+))?[ \t]*\(\s*(O|D|M|C|F)\s*\)/gim))
      assign(m[3],m[1],m[2]);

    // Phone followed by word label.
    for(const m of line.matchAll(/((?:\+?1[ .-]?)?(?:\(?\d{3}\)?[ .-]*)\d{3}[ .-]*\d{4})(?:[ \t]*(?:x|ext\.?|extension)[ \t]*[:.\-]?[ \t]*(\d+))?[ \t]+(Office|Business|Direct\/Text|Direct|Mobile|Cell|Fax)\b/gim))
      assign(m[3].replace(/\//g,""),m[1],m[2]);
  }

  if(!out.businessPhone&&!out.mobilePhone){
    const vals=phoneTokens(lines.join("\n")).map(x=>phoneForOutlook(x.value)).filter(Boolean);
    const unique=[...new Set(vals)];
    if(unique.length===1)out.businessPhone=unique[0];
  }
  if(out.businessPhone===out.mobilePhone)out.mobilePhone="";
  if(out.businessPhone===out.businessFax)out.businessFax="";
  return out;
}
function decodeProofpointUrl(v){
  const raw=String(v||"").trim();
  if(!/urldefense\.proofpoint\.com\/v2\/url/i.test(raw))return raw;
  try{
    const u=new URL(raw);
    let enc=u.searchParams.get("u")||"";
    if(!enc)return raw;
    // Proofpoint v2 encodes punctuation as -HH and slashes as underscores.
    enc=enc.replace(/-([0-9A-Fa-f]{2})/g,(_,h)=>String.fromCharCode(parseInt(h,16))).replace(/_/g,"/");
    return decodeURIComponent(enc);
  }catch(_){return raw}
}
function normalizeWebsiteUrl(v){
  let raw=decodeProofpointUrl(v).trim().replace(/[),.;]+$/g,"");
  if(!raw)return"";
  if(!/^https?:\/\//i.test(raw))raw="https://"+raw;
  try{
    const u=new URL(raw);
    if(/^www\./i.test(u.hostname)||u.hostname.split(".").length>=2){
      const host=u.hostname.toLowerCase();
      return "https://"+host.replace(/^www\./,"");
    }
  }catch(_){ }
  return raw;
}
function isJunkResourceUrl(v){return /(?:\.(?:png|jpe?g|gif|svg|webp|bmp|ico)(?:[?#]|$)|^cid:|^data:|\/image\/|\/images\/|\/logo[s]?\/|safelinks\.protection\.outlook\.com|google\.[^/]+\/maps|maps\.google\.|maps\.apple\.|bing\.com\/maps|goo\.gl\/maps)/i.test(v||"")}
function website(sig,senderEmail){
  const proof=/https?:\/\/urldefense\.proofpoint\.com\/v2\/url\?[^\s|]+/i;
  const normal=/\b(?:https?:\/\/)?(?:www\.)?[a-z0-9][a-z0-9.-]+\.[a-z]{2,}(?:\/[^\s|]*)?/i;
  for(const line of norm(sig).split("\n")){
    if(line.includes("@")&&!proof.test(line))continue;
    let m=line.match(proof)||line.match(normal);if(!m)continue;
    let v=m[0];
    if(isJunkResourceUrl(v)&&!/urldefense\.proofpoint\.com/i.test(v))continue;
    v=normalizeWebsiteUrl(v);
    if(v&&!/urldefense\.proofpoint\.com/i.test(v))return v;
  }
  const email=cleanEmail(senderEmail);const domain=email.split("@")[1]||"";
  const personal=/^(gmail|outlook|hotmail|live|icloud|me|aol|yahoo|protonmail|msn)\./i;
  if(domain&&!personal.test(domain))return "https://"+domain.toLowerCase();
  return "";
}

function isClosingPhrase(line){
  const v=String(line||"").toLowerCase()
    .replace(/[.,!;:]+$/g,"")
    .replace(/\s+/g," ")
    .trim();
  if(!v)return false;
  return new Set([
    "thank you","thanks","many thanks","thank you very much","thank you again",
    "best","best regards","kind regards","warm regards","regards",
    "sincerely","respectfully","cheers","all the best","with appreciation"
  ]).has(v);
}


function signatureNameFromLine(line){
  const raw=String(line||"").replace(/\s+/g," ").trim();
  if(!raw||isClosingPhrase(raw))return "";
  if(/\d|@|https?:|www\.|\b(?:street|st\.?|road|rd\.?|ave\.?|avenue|blvd\.?|boulevard|suite|ste\.?|drive|dr\.?|lane|ln\.?|loop|way|court|ct\.?|place|pl\.?)\b/i.test(raw))return "";
  if(/^(?:wire transfer|payment options?|please\b|good (?:morning|afternoon|evening)\b|see attached\b|attached\b|invoice\b|message\b|hello\b|hi\b)/i.test(raw))return "";
  if(/[!?]$/.test(raw))return "";
  const split=splitNameTitleLine(raw);
  if(split&&split.name)return split.name;
  return looksLikePersonName(raw)?raw:"";
}
function splitNameTitleLine(line){
  const raw=String(line||"").replace(/\s+/g," ").trim();
  if(!raw || isClosingPhrase(raw))return null;

  // Delimited name/title: "Linda Shin | Associate"
  const pieces=raw.split(/\s*(?:\||•|·|—|–)\s*/).map(x=>x.trim()).filter(Boolean);
  if(pieces.length>1 && looksLikePersonName(pieces[0])){
    const rest=pieces.slice(1).join(" | ").trim();
    if(rest && TITLE_WORDS.some(t=>rest.toLowerCase().includes(t))){
      return {name:pieces[0],title:rest};
    }
  }

  // Undelimited name/title: "Alex Athanasopoulos   Admin & Project Support Coordinator"
  const words=raw.split(/\s+/).filter(Boolean);
  for(let n=2;n<=Math.min(4,words.length-1);n++){
    const possibleName=words.slice(0,n).join(" ");
    const possibleTitle=words.slice(n).join(" ");
    if(looksLikePersonName(possibleName) &&
       TITLE_WORDS.some(t=>possibleTitle.toLowerCase().includes(t))){
      return {name:possibleName,title:possibleTitle};
    }
  }
  return null;
}

function headerIdentity(value){
  const raw=String(value||"").trim();
  const email=cleanEmail(raw);
  let name=raw;
  if(email)name=name.replace(new RegExp(email.replace(/[.*+?^${}()|[\]\\]/g,"\\$&"),"i")," ");
  name=name
    .replace(/<[^>]*>/g," ")
    .replace(/\[[^\]]*mailto:[^\]]*\]/ig," ")
    .replace(/\bmailto:\S+/ig," ")
    .replace(/[<>]/g," ")
    .replace(/\s+/g," ")
    .trim();
  return {name:cleanName(name),email};
}

function title(sig,authoritativeName=""){
  const lines=norm(sig).split("\n").map(x=>x.trim()).filter(Boolean);
  const auth=nameParts(authoritativeName||"");
  const authSurname=String(auth.surname||"").toLowerCase();

  const cleanPart=v=>String(v||"")
    .replace(/^[-|•·—–\s]+|[-|•·—–\s]+$/g,"")
    .replace(/\s+/g," ")
    .trim();

  const hasTitleWord=v=>TITLE_WORDS.some(t=>String(v||"").toLowerCase().includes(t));

  function stripAuthoritativeOnly(v){
    const original=String(v||"").trim();
    let plain=cleanPart(original);
    if(!plain)return "";
    const authName=cleanName(authoritativeName||"");
    if(authName){
      const low=plain.toLowerCase(),alow=authName.toLowerCase();
      if(low.startsWith(alow+" "))return cleanPart(plain.slice(authName.length));
    }
    if(authSurname){
      const low=plain.toLowerCase();
      if(low===authSurname)return "";
      if(low.startsWith(authSurname+" ")){
        const rem=cleanPart(plain.slice(auth.surname.length));
        if(hasTitleWord(rem)||rem.endsWith("&"))return rem;
      }
    }
    return original;
  }

  function stripKnownNamePrefix(v){
    let raw=cleanPart(v);
    if(!raw)return "";

    const split=splitNameTitleLine(raw);
    if(split&&split.title)return cleanPart(split.title);

    const authName=cleanName(authoritativeName||"");
    if(authName){
      const low=raw.toLowerCase(),alow=authName.toLowerCase();
      if(low.startsWith(alow+" ")){
        const rem=cleanPart(raw.slice(authName.length));
        if(hasTitleWord(rem))return rem;
      }
    }

    // Outlook sometimes wraps first and last name separately. If the title line begins
    // with the authoritative surname, remove that surname before evaluating the title.
    if(authSurname){
      const low=raw.toLowerCase();
      if(low===authSurname)return "";
      if(low.startsWith(authSurname+" ")){
        const rem=cleanPart(raw.slice(auth.surname.length));
        if(hasTitleWord(rem)||rem.endsWith("&"))return rem;
      }
    }

    // Last-resort same-line nickname/full-name split.
    const words=raw.split(/\s+/).filter(Boolean);
    for(let n=2;n<=Math.min(4,words.length-1);n++){
      const possibleName=words.slice(0,n).join(" ");
      const rem=words.slice(n).join(" ");
      if(looksLikePersonName(possibleName)&&hasTitleWord(rem))return cleanPart(rem);
    }
    return raw;
  }

  const joinParts=parts=>{
    let out="";
    for(const raw of parts){
      const rawText=String(raw||"").trim();
      const part=cleanPart(raw);if(!part)continue;
      if(!out){out=part;continue}
      if(/^\|/.test(rawText)){out=out+" | "+part;continue}
      if(out.endsWith("&"))out=out.replace(/\s*&\s*$/,"")+" & "+part.replace(/^&\s*/,"");
      else if(part.startsWith("&"))out=out+" & "+part.replace(/^&\s*/,"");
      else out=out+" "+part;
    }
    return out.replace(/\s+/g," ").trim();
  };

  const canContinue=(current,next)=>{
    const c=String(current||"").trim(),n=String(next||"").trim(),nl=n.toLowerCase();
    if(!n)return false;
    if(c.endsWith("&")||n.startsWith("&"))return true;
    if(/^\s*[|•·—–]/.test(n)&&hasTitleWord(nl))return true;
    // A short first title fragment can wrap before another recognizable title phrase.
    if(c.split(/\s+/).length<=3&&hasTitleWord(c)&&hasTitleWord(nl))return true;
    return false;
  };

  for(let i=0;i<lines.length;i++){
    const line=lines[i];
    if(isClosingPhrase(line)||line.length>160)continue;

    const stripped=stripKnownNamePrefix(line);
    if(stripped!==cleanPart(line) && (hasTitleWord(stripped)||stripped.endsWith("&"))){
      const parts=[stripped];
      if(i+1<lines.length&&canContinue(parts[0],lines[i+1]))parts.push(stripAuthoritativeOnly(lines[i+1]));
      return joinParts(parts);
    }

    const split=splitNameTitleLine(line);
    if(split&&split.title){
      const parts=[split.title];
      if(i+1<lines.length&&canContinue(parts[0],lines[i+1]))parts.push(stripAuthoritativeOnly(lines[i+1]));
      return joinParts(parts);
    }

    const low=line.toLowerCase();
    if(hasTitleWord(low)&&!COMPANY_WORDS.some(w=>(" "+low).includes(w))){
      const first=stripKnownNamePrefix(line);
      if(!first||!hasTitleWord(first))continue;
      const parts=[first];
      if(i+1<lines.length&&canContinue(parts[0],lines[i+1]))parts.push(stripAuthoritativeOnly(lines[i+1]));
      return joinParts(parts);
    }
  }
  return "";
}
function looksLikePersonName(line){
  const v=(line||"").trim();
  if(!v||v.length<4||v.length>70||isClosingPhrase(v))return false;
  if(cleanEmail(v)||/\d|https?:|www\.|@|\b(?:street|st\.?|road|rd\.?|ave\.?|avenue|blvd\.?|boulevard|suite|ste\.?|drive|dr\.?|lane|ln\.?|city|inc\.?|llc|corp\.?|company|garage|workshop)\b/i.test(v))return false;
  const low=v.toLowerCase();
  if(TITLE_WORDS.some(t=>low.includes(t))||COMPANY_WORDS.some(w=>(" "+low).includes(w)))return false;
  const words=v.replace(/[,]/g," ").split(/\s+/).filter(Boolean);
  if(words.length<2||words.length>5)return false;
  return words.every(w=>/^[A-Za-z][A-Za-z'.-]*$/.test(w));
}
function personNameFromLine(line){
  const raw=(line||"").trim();
  if(isClosingPhrase(raw))return "";
  if(looksLikePersonName(raw))return raw;
  const split=splitNameTitleLine(raw);
  return split?split.name:"";
}
function inferPersonName(sig){
  const lines=norm(sig).split("\n").map(x=>x.trim()).filter(Boolean);
  for(let i=0;i<Math.min(lines.length,12);i++){
    if(isClosingPhrase(lines[i]))continue;
    const split=splitNameTitleLine(lines[i]);
    if(split&&split.name)return split.name;
    const found=personNameFromLine(lines[i]);
    if(found){
      const next=(lines[i+1]||"").toLowerCase();
      if(!next||TITLE_WORDS.some(t=>next.includes(t))||COMPANY_WORDS.some(w=>(" "+next).includes(w))||lines[i].includes("|")||lines[i].includes("•"))return found;
    }
  }
  return "";
}
function companyFromDomain(senderEmail,site){
  let domain=cleanEmail(senderEmail).split("@")[1]||"";
  if(!domain&&site){try{domain=new URL(/^https?:\/\//i.test(site)?site:"https://"+site).hostname}catch(_){domain=""}}
  domain=domain.toLowerCase().replace(/^www\./,"");
  if(!domain||/^(gmail|outlook|hotmail|live|icloud|me|aol|yahoo|protonmail|msn)\./i.test(domain))return"";
  let stem=domain.split(".")[0].replace(/[-_]+/g," ").trim();
  if(!stem)return"";
  if(!/\s/.test(stem)){
    const suffixes=["construction","concrete","electric","electrical","engineering","architects","architecture","design","builders","building","mechanical","plumbing","contractors","contractor","services","solutions","systems","garage","workshop"];
    for(const s of suffixes){
      if(stem.toLowerCase().endsWith(s) && stem.length>s.length+1){
        stem=stem.slice(0,-s.length)+" "+stem.slice(-s.length);
        break;
      }
    }
  }
  return stem.split(/\s+/).map(w=>/^\d+$/.test(w)?w:(w.charAt(0).toUpperCase()+w.slice(1))).join(" ");
}
function company(sig,name,senderEmail,site){
  const lines=norm(sig).split("\n").map(x=>x.trim()).filter(Boolean),lowName=(name||"").toLowerCase();
  const noise=/\b(?:payment|invoice|billing|remitted|remit|check|cheque|wire transfer|credit card|echeck|customer id|pay online|please see attached|please reach|thank you)\b/i;
  const addressLike=/\b(?:street|st\.?|road|rd\.?|avenue|ave\.?|boulevard|blvd\.?|drive|dr\.?|lane|ln\.?|loop|suite|ste\.?|bldg|building)\b/i;

  for(const line of lines){
    const low=" "+line.toLowerCase();
    if(line.length<=2||line.length>=90||noise.test(line))continue;
    if(addressLike.test(line)&&/\d/.test(line))continue;
    if(COMPANY_WORDS.some(w=>low.includes(w))){
      if(/\b(?:at|to)\s+\d{1,6}\b/i.test(line))continue;
      return line;
    }
  }

  const domainCompany=companyFromDomain(senderEmail,site);
  if(domainCompany)return domainCompany;

  for(const line of lines.slice(0,10)){
    const low=line.toLowerCase();
    if(noise.test(line)||low===lowName||isClosingPhrase(line)||looksLikePersonName(line)||low.includes("@")||/\d{3}[\s.\-]\d{3}/.test(line)||/^https?:|^www\./i.test(line))continue;
    if(TITLE_WORDS.some(t=>low.includes(t)))continue;
    if(/^from:|^sent:|^to:|^cc:|^subject:/i.test(line))continue;
    if(addressLike.test(line)||/\b[A-Z]{2}\s+\d{5}(?:-\d{4})?\b/.test(line))continue;
    if(/\d/.test(line)||/[.!?]\s*$/.test(line)||line.split(/\s+/).length>7)continue;
    if(line.length>=3&&line.length<=60)return line;
  }
  return "";
}
const US_STATES=new Set(["AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN","IA","KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV","NH","NJ","NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN","TX","UT","VT","VA","WA","WV","WI","WY","DC"]);
function usCountry(state){return US_STATES.has(String(state||"").toUpperCase())?"United States":""}
function address(sig){
  const raw=norm(sig).replace(/[\u200B-\u200D\uFEFF]/g,"");
  const lines=raw.split("\n").map(x=>x.replace(/\s+/g," ").trim()).filter(Boolean);
  const streetWord=/\b(street|st\.?|road|rd\.?|avenue|ave\.?|boulevard|blvd\.?|drive|dr\.?|lane|ln\.?|way|court|ct\.?|highway|hwy\.?|parkway|pkwy\.?|place|pl\.?|trail|trl\.?|circle|cir\.?|square|sq\.?|loop|terrace|ter\.?|suite|ste\.?|floor|fl\.?|bldg|building|p\.?o\.?\s*box)\b/i;
  const cityStateZip=/([A-Za-z][A-Za-z .'-]{1,60}?),?\s+([A-Z]{2})\s+(\d{5}(?:-\d{4})?)/g;

  // Strong combined pattern for signatures flattened into one line, e.g.
  // "2657 Aero Loop Sheridan, WY 82801". This avoids treating the street as part of the city.
  const flatAll=raw.replace(/[\n•|]+/g," ").replace(/\s+/g," ").trim();
  const combined=flatAll.match(/(\d{1,6}\s+[A-Za-z0-9 .,'#&\/-]+?\b(?:street|st\.?|road|rd\.?|avenue|ave\.?|boulevard|blvd\.?|drive|dr\.?|lane|ln\.?|way|court|ct\.?|highway|hwy\.?|parkway|pkwy\.?|place|pl\.?|trail|trl\.?|circle|cir\.?|square|sq\.?|loop|terrace|ter\.?)(?:\s+(?:suite|ste\.?|bldg|building|floor|fl\.?)\s*[A-Za-z0-9-]+)?)\s+([A-Za-z][A-Za-z .'-]{1,50}?),?\s+([A-Z]{2})\s+(\d{5}(?:-\d{4})?)/i);
  if(combined && US_STATES.has(combined[3].toUpperCase())){
    return {street:combined[1].trim(),city:combined[2].trim(),state:combined[3].toUpperCase(),postalCode:combined[4],countryOrRegion:"United States"};
  }

  // First, find a city/state/ZIP anywhere in the signature. Outlook HTML often collapses
  // what visually appear to be separate address lines into one line of text.
  let location=null;
  for(const line of lines){
    cityStateZip.lastIndex=0;
    let m;
    while((m=cityStateZip.exec(line))){
      if(US_STATES.has(m[2].toUpperCase())){
        location={line,city:m[1].trim().replace(/^[,;|•\s]+|[,;|•\s]+$/g,""),state:m[2].toUpperCase(),postalCode:m[3],index:m.index};
        break;
      }
    }
    if(location)break;
  }

  // If line-by-line matching failed, search the complete signature with separators normalized.
  if(!location){
    const flat=raw.replace(/[•|]/g," \n ").replace(/\s+/g," ").trim();
    cityStateZip.lastIndex=0;
    let m;
    while((m=cityStateZip.exec(flat))){
      if(US_STATES.has(m[2].toUpperCase())){
        location={line:flat,city:m[1].trim().replace(/^[,;|•\s]+|[,;|•\s]+$/g,""),state:m[2].toUpperCase(),postalCode:m[3],index:m.index,flat:true};
        break;
      }
    }
  }

  if(!location){
    // Very loose fallback: find a US state+ZIP anywhere, then derive city and street independently.
    const flat=raw.replace(/[\n•|]+/g," ").replace(/\s+/g," ").trim();
    const sz=[...flat.matchAll(/\b([A-Z]{2})\s+(\d{5}(?:-\d{4})?)\b/g)].find(m=>US_STATES.has(m[1]));
    if(sz){
      const state=sz[1],postalCode=sz[2],before=flat.slice(0,sz.index).replace(/[,;\s]+$/g,"");
      const streetMatch=before.match(/(\d{1,6}\s+[A-Za-z0-9 .,'#&\/-]+?\b(?:street|st\.?|road|rd\.?|avenue|ave\.?|boulevard|blvd\.?|drive|dr\.?|lane|ln\.?|way|court|ct\.?|highway|hwy\.?|parkway|pkwy\.?|place|pl\.?|trail|trl\.?|circle|cir\.?|square|sq\.?|loop|terrace|ter\.?)(?:\s+(?:suite|ste\.?|bldg|building|floor|fl\.?)\s*[A-Za-z0-9-]+)?)/ig);
      let street="",city="";
      if(streetMatch&&streetMatch.length){street=streetMatch[streetMatch.length-1].trim();const pos=before.toLowerCase().lastIndexOf(street.toLowerCase());city=before.slice(pos+street.length).replace(/^[,;\s]+|[,;\s]+$/g,"").trim();}
      if(!city){const cm=before.match(/([A-Za-z][A-Za-z .'-]{1,40})$/);if(cm)city=cm[1].trim();}
      return{street,city,state,postalCode,countryOrRegion:"United States"};
    }
    return{street:"",city:"",state:"",postalCode:"",countryOrRegion:""};
  }

  let street="";
  const locLineIndex=lines.findIndex(l=>l===location.line);

  // Same-line street, e.g. "2657 Aero Loop Sheridan, WY 82801".
  if(!location.flat && location.index>0){
    const before=location.line.slice(0,location.index).replace(/[•|,;\s]+$/g,"").trim();
    if(/\d/.test(before)&&streetWord.test(before))street=before;
  }

  // Normal multi-line signature: walk backward from city/state/ZIP for the street line(s).
  if(!street && locLineIndex>=0){
    const picked=[];
    for(let j=locLineIndex-1;j>=Math.max(0,locLineIndex-4);j--){
      const prev=lines[j];
      if(/^from:|^sent:|^to:|^cc:|^bcc:|^subject:/i.test(prev))break;
      const addressLike=(/^\d{1,6}\s+/.test(prev)||/p\.?o\.?\s*box/i.test(prev))&&(streetWord.test(prev)||/\d/.test(prev));
      if(addressLike)picked.unshift(prev);
      else if(picked.length)break;
    }
    street=picked.join("\n");
  }

  // Last-resort extraction from flattened HTML text immediately before the city/state/ZIP.
  if(!street){
    const flat=raw.replace(/[\n•|]+/g," ").replace(/\s+/g," ").trim();
    const locRx=new RegExp(location.city.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")+",?\\s+"+location.state+"\\s+"+location.postalCode.replace("-","\\-"),"i");
    const lm=flat.match(locRx);
    if(lm){
      const prefix=flat.slice(0,lm.index).slice(-180);
      const sm=prefix.match(/(\d{1,6}\s+[A-Za-z0-9 .,'#&\/-]+?\b(?:street|st\.?|road|rd\.?|avenue|ave\.?|boulevard|blvd\.?|drive|dr\.?|lane|ln\.?|way|court|ct\.?|highway|hwy\.?|parkway|pkwy\.?|place|pl\.?|trail|trl\.?|circle|cir\.?|square|sq\.?|loop|terrace|ter\.?)\b(?:\s+(?:suite|ste\.?|bldg|building|floor|fl\.?)\s*[A-Za-z0-9-]+)?)[,;\s]*$/i);
      if(sm)street=sm[1].trim();
    }
  }

  return{street,city:location.city,state:location.state,postalCode:location.postalCode,countryOrRegion:"United States"};
}
function signatureScore(line){let s=0;if(cleanEmail(line))s+=3;if(phoneTokens(line).length)s+=2;if(/\b(?:www\.|https?:\/\/)/i.test(line))s+=2;if(/\b[A-Z]{2}\s+\d{5}(?:-\d{4})?\b/.test(line))s+=2;if(/^\d{1,6}\s+/.test(line))s+=1;const low=" "+line.toLowerCase();if(COMPANY_WORDS.some(w=>low.includes(w)))s+=2;if(TITLE_WORDS.some(w=>low.includes(w)))s+=1;if(looksLikePersonName(line))s+=1;return s}


function identitySignatureOccurrences(text,identityName,identityEmail){
  const lines=norm(text).split("\n").map(x=>x.trim());
  const email=cleanEmail(identityEmail);
  const np=nameParts(identityName||"");
  const surname=String(np.surname||"").toLowerCase();
  if(!email||!surname)return [];
  const isHeader=l=>/^\s*(from|sent|to|cc|bcc|subject):\s*/i.test(l)||/^[-_]{5,}$/.test(l);
  const isDisclaimer=l=>/confidential|privileged|intended recipient|virus|disclaimer|please consider the environment/i.test(l);
  const hasSurname=l=>new RegExp(`(?:^|\\s)${surname.replace(/[.*+?^${}()|[\\]\\]/g,"\\$&")}(?:$|\\s|[,|•·—–-])`,`i`).test(String(l||""));
  const out=[];

  for(let i=0;i<lines.length;i++){
    const line=lines[i];
    if(!line||isHeader(line)||isDisclaimer(line)||isClosingPhrase(line)||!hasSurname(line))continue;

    // A genuine signature-name line must be followed shortly by this person's email
    // and by other signature evidence. This keeps ordinary body mentions out.
    let emailLine=-1,score=0;
    const maxEnd=Math.min(lines.length-1,i+12);
    for(let j=i;j<=maxEnd;j++){
      const v=lines[j];
      if(j>i&&(isHeader(v)||isDisclaimer(v)))break;
      if(j>i&&hasSurname(v)&&j!==i)break;
      if(sameEmail(cleanEmail(v),email)){emailLine=j;score+=8;break}
      score+=signatureScore(v);
    }
    if(emailLine<0)continue;

    let start=i;
    // Outlook sometimes breaks "Alex Athanasopoulos" into two lines. Include a plausible
    // one-word first-name fragment immediately above the surname/title line.
    if(i>0){
      const prev=lines[i-1];
      if(prev&&!isHeader(prev)&&!isClosingPhrase(prev)&&/^[A-Za-z][A-Za-z'.-]*$/.test(prev))start=i-1;
    }

    let end=emailLine;
    // Include a clean website-only line immediately following the email if Outlook split it.
    for(let j=emailLine+1;j<=Math.min(lines.length-1,emailLine+2);j++){
      const v=lines[j];
      if(!v||isHeader(v)||isDisclaimer(v)||hasSurname(v)||cleanEmail(v))break;
      if(/^https?:|^www\./i.test(v)){end=j;continue}
      break;
    }

    const sig=cleanSignatureText(lines.slice(start,end+1).filter(Boolean).join("\n"));
    const sigLines=norm(sig).split("\n").filter(Boolean);
    if(!sig||sigLines.length<3||sigLines.length>12)continue;
    const ph=phones(sig),ad=address(sig);
    const evidence=(ph.businessPhone||ph.mobilePhone?2:0)+(ad.state&&ad.postalCode?2:0)+(title(sig,identityName)?2:0);
    if(evidence<2)continue;
    out.push({sig,start,end,emailLine,occurrence:out.length,score:100+score+evidence});
    i=end;
  }
  return out;
}

function sanitizeJobTitle(value,identityName,signatureText=""){
  let t=String(value||"").replace(/\s+/g," ").trim();
  const fromSig=title(signatureText||"",identityName||"");
  if(fromSig)t=String(fromSig).replace(/\s+/g," ").trim();
  if(!t)return "";

  const split=splitNameTitleLine(t);
  if(split&&split.title)t=split.title;

  const np=nameParts(identityName||"");
  const full=cleanName(identityName||"");
  if(full&&t.toLowerCase().startsWith(full.toLowerCase()+" "))t=t.slice(full.length).trim();
  if(np.surname){
    const sr=np.surname.replace(/[.*+?^${}()|[\]\\]/g,"\\$&");
    t=t.replace(new RegExp(`^[A-Za-z][A-Za-z'.-]*\\s+${sr}\\s+`,`i`),"");
    t=t.replace(new RegExp(`^${sr}\\s+`,`i`),"");
  }
  return t.replace(/^[-|•·—–\s]+|[-|•·—–\s]+$/g,"").replace(/\s+/g," ").trim();
}

function tightNotesSignature(text,identityName,identityEmail){
  const identityOcc=identitySignatureOccurrences(text,identityName,identityEmail);
  if(identityOcc.length)return identityOcc[0].sig;
  const emailOcc=signatureOccurrences(text,identityEmail);
  if(emailOcc.length)return emailOcc[0].sig;

  // Last-resort safety: never save a large email/message body as Notes.
  const cleaned=cleanSignatureText(text||"");
  const lines=norm(cleaned).split("\n").map(x=>x.trim()).filter(Boolean);
  if(lines.length>12||/^(from|sent|to|cc|subject):/im.test(cleaned)||/\b(?:payment options?|please see attached|please reach out|wire transfer|customer id)\b/i.test(cleaned))return "";
  return cleaned;
}

function signatureOccurrences(segmentText,senderEmail){
  const lines=norm(segmentText).split("\n").map(x=>x.trim());
  const email=cleanEmail(senderEmail);
  if(!email)return [];
  const isHeader=l=>/^\s*(from|sent|to|cc|bcc|subject):\s*/i.test(l)||/^[-_]{5,}$/.test(l);
  const isDisclaimer=l=>/confidential|privileged|intended recipient|virus|disclaimer|please consider the environment/i.test(l);
  const anchors=[];
  for(let i=0;i<lines.length;i++)if(!isHeader(lines[i])&&sameEmail(cleanEmail(lines[i]),email))anchors.push(i);
  const out=[];
  for(let occurrence=0;occurrence<anchors.length;occurrence++){
    const a=anchors[occurrence];
    let start=-1;
    for(let j=a-1;j>=Math.max(0,a-9);j--){
      const v=lines[j];
      if(isHeader(v)||isDisclaimer(v))break;
      if(isClosingPhrase(v))continue;
      if(signatureNameFromLine(v)){start=j;break}
    }
    if(start<0)continue;

    let end=a;
    // Allow a website/company line immediately after the email, but never cross into a
    // second signature occurrence or a new message/header.
    for(let j=a+1;j<Math.min(lines.length,a+4);j++){
      const v=lines[j];
      if(!v||isHeader(v)||isDisclaimer(v)||signatureNameFromLine(v)||cleanEmail(v))break;
      if(signatureScore(v)>0||/^https?:|^www\./i.test(v)){end=j;continue}
      break;
    }
    const sig=cleanSignatureText(lines.slice(start,end+1).filter(Boolean).join("\n"));
    if(!sig)continue;
    const score=norm(sig).split("\n").reduce((t,l)=>t+signatureScore(l),0)+12;
    if(score<8)continue;
    out.push({sig,score,start,end,anchor:a,occurrence});
  }
  return out;
}

function isolateSignature(segmentText,senderEmail){
  const rawLines=norm(segmentText).split("\n").map(x=>x.trim());
  const isHeader=l=>/^\s*(from|sent|to|cc|subject):\s*/i.test(l)||/^[-_]{5,}$/.test(l);
  const isDisclaimer=l=>/confidential|privileged|intended recipient|virus|disclaimer|please consider the environment/i.test(l);
  const email=(senderEmail||"").toLowerCase();
  const anchors=[];
  for(let i=0;i<rawLines.length;i++){
    if(email&&cleanEmail(rawLines[i])===email&&!isHeader(rawLines[i]))anchors.push(i);
  }
  // Strongest rule: a sender's own email inside the message body almost always sits in the signature.
  // Build a compact block around it and stop at blank lines / quoted-message headers so email prose
  // cannot leak into Notes or contact fields.
  if(anchors.length){
    let best=null;
    for(const a of anchors){
      let lo=a,hi=a;
      while(lo>0 && a-lo<8){
        const p=rawLines[lo-1];
        if(!p||isHeader(p)||isDisclaimer(p))break;
        lo--;
      }
      while(hi+1<rawLines.length && hi-a<5){
        const n=rawLines[hi+1];
        if(!n||isHeader(n)||isDisclaimer(n))break;
        hi++;
      }
      let chunk=rawLines.slice(lo,hi+1).filter(Boolean);
      // If there was no blank line, trim leading prose until the block becomes signature-like.
      while(chunk.length>5 && signatureScore(chunk[0])===0 && signatureScore(chunk[1])===0)chunk.shift();
      const score=chunk.reduce((t,l)=>t+signatureScore(l),0)+8;
      if(!best||score>best.score)best={score,chunk};
    }
    if(best&&best.chunk.length)return best.chunk.join("\n");
  }
  // No sender-email anchor: score compact blocks, but heavily penalize prose.
  let lines=rawLines.filter((l,i)=>l && !isDisclaimer(l));
  if(lines.length>60)lines=lines.slice(-60);
  let best={score:-999,chunk:[]};
  for(let end=0;end<lines.length;end++){
    if(isHeader(lines[end]))continue;
    for(let len=3;len<=12;len++){
      const start=Math.max(0,end-len+1),chunk=lines.slice(start,end+1);
      if(chunk.some(isHeader))continue;
      let score=chunk.reduce((t,l)=>t+signatureScore(l),0)+end/Math.max(1,lines.length);
      score-=chunk.filter(l=>l.length>90||(/[.!?]$/.test(l)&&l.split(/\s+/).length>10)).length*5;
      if(score>best.score)best={score,chunk};
    }
  }
  if(best.score>=3)return best.chunk.join("\n");
  return "";
}

const US_STATE_NAMES={
  alabama:"AL",alaska:"AK",arizona:"AZ",arkansas:"AR",california:"CA",colorado:"CO",connecticut:"CT",delaware:"DE",florida:"FL",georgia:"GA",hawaii:"HI",idaho:"ID",illinois:"IL",indiana:"IN",iowa:"IA",kansas:"KS",kentucky:"KY",louisiana:"LA",maine:"ME",maryland:"MD",massachusetts:"MA",michigan:"MI",minnesota:"MN",mississippi:"MS",missouri:"MO",montana:"MT",nebraska:"NE",nevada:"NV","new hampshire":"NH","new jersey":"NJ","new mexico":"NM","new york":"NY","north carolina":"NC","north dakota":"ND",ohio:"OH",oklahoma:"OK",oregon:"OR",pennsylvania:"PA","rhode island":"RI","south carolina":"SC","south dakota":"SD",tennessee:"TN",texas:"TX",utah:"UT",vermont:"VT",virginia:"VA",washington:"WA","west virginia":"WV",wisconsin:"WI",wyoming:"WY","district of columbia":"DC"
};

function structuredStateAbbrev(v){
  const raw=String(v||"").trim();
  if(!raw)return"";
  const up=raw.toUpperCase();
  if(US_STATES.has(up))return up;
  return US_STATE_NAMES[raw.toLowerCase()]||"";
}

function structuredPhoneObservations(sig){
  const observations=[];
  const re=/(?:\+?1[\s.\-]?)?(?:\(?\d{3}\)?[\s.\-]*)\d{3}[\s.\-]*\d{4}(?:\s*(?:x|ext\.?|extension)\s*[:.\-]?\s*\d+)?/gi;

  const tokenLabel=token=>{
    const t=String(token||"").toLowerCase().replace(/[().]/g,"").replace(/\s+/g," ").trim();
    if(/^(?:direct\s*[/&+\-]\s*text|direct text|text\s*[/&+\-]\s*direct)$/.test(t))return"directtext";
    if(/^(?:fax|f)$/.test(t))return"fax";
    if(/^(?:mobile|cell|cellular|m|c)$/.test(t))return"mobile";
    if(/^(?:office|business|work|main|telephone|phone|tel|p|o)$/.test(t))return"business";
    if(/^(?:direct|d)$/.test(t))return"direct";
    return"";
  };

  const prefixLabel=segment=>{
    const s=String(segment||"");
    const m=s.match(/(?:^|[|•;,]\s*|\s+)(direct\s*[/&+\-]\s*text|direct text|text\s*[/&+\-]\s*direct|office|business|work|main|telephone|phone|tel|mobile|cell|cellular|fax|direct|[POMCFD])\s*[:.\-]?\s*$/i);
    return m?tokenLabel(m[1]):"";
  };

  const suffixLabel=segment=>{
    const s=String(segment||"");
    let m=s.match(/^\s*\(\s*(direct\s*[/&+\-]\s*text|direct text|office|business|work|main|telephone|phone|tel|mobile|cell|cellular|fax|direct|[OMCFD])\s*\)/i);
    if(m)return tokenLabel(m[1]);
    m=s.match(/^\s*(direct\s*[/&+\-]\s*text|direct text|mobile|cell|cellular|fax|direct)\b(?!\s*:)/i);
    return m?tokenLabel(m[1]):"";
  };

  for(const rawLine of norm(sig).split("\n")){
    const line=String(rawLine||"").replace(/\u00a0/g," ").trim();
    if(!line)continue;
    const matches=[...line.matchAll(re)];
    for(let k=0;k<matches.length;k++){
      const m=matches[k];
      const value=phoneForOutlook(m[0]); if(!value)continue;
      const start=m.index||0,end=start+m[0].length;
      const prevEnd=k?((matches[k-1].index||0)+matches[k-1][0].length):0;
      const nextStart=k+1<matches.length?(matches[k+1].index||line.length):line.length;
      const before=line.slice(prevEnd,start);
      const after=line.slice(end,nextStart);

      // A parenthesized/word suffix belongs to the number just before it.
      // Otherwise use the nearest prefix in the local segment. This prevents
      // the "M:" before a second number from labeling the first number.
      const label=suffixLabel(after)||prefixLabel(before);
      observations.push({value,label,hasExtension:/x\d+$/i.test(value),line});
    }
  }

  const rank={fax:7,business:6,mobile:6,directtext:6,direct:5,"":1};
  const by=new Map();
  for(const o of observations){
    const prev=by.get(o.value);
    if(!prev||rank[o.label]>rank[prev.label])by.set(o.value,o);
  }
  return [...by.values()];
}
function structuredPhones(sig){
  const obs=structuredPhoneObservations(sig);
  const out={businessPhone:"",mobilePhone:"",businessFax:""};
  const take=(kind,o)=>{if(o&&!out[kind])out[kind]=o.value};
  const fax=obs.filter(o=>o.label==="fax");
  const bus=obs.filter(o=>o.label==="business");
  const mob=obs.filter(o=>o.label==="mobile"||o.label==="directtext");
  const direct=obs.filter(o=>o.label==="direct");
  const unlabeled=obs.filter(o=>!o.label);
  take("businessFax",fax[0]);
  take("businessPhone",bus[0]);
  take("mobilePhone",mob[0]);

  // Direct with an extension is a company/direct-dial number, so Business.
  for(const o of direct.filter(x=>x.hasExtension))take("businessPhone",o);

  // Assign one unlabeled number as Business before interpreting a bare Direct.
  // This makes an unlabeled main number + Direct number map Business + Mobile.
  for(const o of unlabeled)take("businessPhone",o);

  for(const o of direct.filter(x=>!x.hasExtension)){
    if(out.businessPhone && out.businessPhone!==o.value)take("mobilePhone",o);
    else if(!out.mobilePhone && (bus.length||unlabeled.some(x=>x.value!==o.value)))take("mobilePhone",o);
    else take("businessPhone",o);
  }

  // If only an explicit mobile existed, a remaining unlabeled number is Business.
  if(!out.businessPhone){for(const o of unlabeled)take("businessPhone",o)}

  // Preserve old parser only as a gap filler, never overriding relational assignments.
  const legacy=phones(sig);
  if(!out.businessPhone)out.businessPhone=legacy.businessPhone||"";
  if(!out.mobilePhone)out.mobilePhone=legacy.mobilePhone||"";
  if(!out.businessFax)out.businessFax=legacy.businessFax||"";
  if(out.businessPhone===out.mobilePhone)out.mobilePhone="";
  if(out.businessPhone===out.businessFax)out.businessFax="";
  return out;
}

function structuredAddress(sig){
  const legacy=address(sig);
  if(legacy.street&&legacy.city&&legacy.state&&legacy.postalCode)return legacy;
  const lines=norm(sig).split("\n").map(x=>String(x||"").replace(/\u00a0/g," ").replace(/\s+/g," ").trim()).filter(Boolean);
  const stateNames=Object.keys(US_STATE_NAMES).sort((a,b)=>b.length-a.length).map(x=>x.replace(/ /g,"\\s+")).join("|");
  const locRx=new RegExp("([A-Za-z][A-Za-z .'-]{1,60}?),?\\s+("+stateNames+"|[A-Z]{2})\\s+(\\d{5}(?:-\\d{4})?)","i");
  const streetLike=/\b(?:p\.?o\.?\s*box|street|st\.?|road|rd\.?|avenue|ave\.?|boulevard|blvd\.?|drive|dr\.?|lane|ln\.?|way|court|ct\.?|highway|hwy\.?|parkway|pkwy\.?|place|pl\.?|trail|trl\.?|circle|cir\.?|square|sq\.?|loop|terrace|ter\.?|broadway|suite|ste\.?)\b/i;
  for(let i=0;i<lines.length;i++){
    const line=lines[i]; const m=line.match(locRx); if(!m)continue;
    const state=structuredStateAbbrev(m[2]); if(!state)continue;
    let city=m[1].trim().replace(/^[,;|•\s]+|[,;|•\s]+$/g,"");
    let street="";
    const prefix=line.slice(0,m.index).replace(/[,;|•\s]+$/g,"").trim();
    if(prefix&&(/\d/.test(prefix)||streetLike.test(prefix)))street=prefix;
    if(!street&&i>0){
      const prev=lines[i-1];
      if((/\d/.test(prev)||streetLike.test(prev))&&!cleanEmail(prev)&&!/^https?:|^www\./i.test(prev))street=prev;
    }
    // A one-line pipe address may have several street/address pieces before city.
    if(street&&line.includes("|")&&m.index>0)street=prefix;
    return {street,city,state,postalCode:m[3],countryOrRegion:"United States"};
  }
  return legacy;
}

function structuredCompanyEvidence(line){
  const v=String(line||"").trim(); if(!v)return false;

  // Whole-word company evidence only. Avoid job-title collisions:
  // Designer != Design; Engineer != Engineers; Architect != Architects.
  if(/\b(?:llc|l\.l\.c\.?|inc\.?|corp\.?|corporation|pllc|lp|llp|ltd\.?|company|construction|builders|building|architecture|architects|engineering|engineers|associates|group|studio|mechanical|electric|electrical|plumbing|design|contractors|concrete|masonry|garage|workshop|services|solutions|systems|enterprises|partners)\b/i.test(v))return true;
  if(/(?:^|\s)co\.(?:\s|$)/i.test(v))return true;
  if(v.length<=70&&/[A-Z]/.test(v)&&v===v.toUpperCase()&&v.split(/\s+/).length>=2&&!/\d{3}/.test(v)&&!TITLE_WORDS.some(t=>v.toLowerCase().includes(t)))return true;
  return false;
}

function structuredDomainKey(senderEmail,site){
  let domain=cleanEmail(senderEmail).split("@")[1]||"";
  if(!domain&&site){
    try{domain=new URL(/^https?:\/\//i.test(site)?site:"https://"+site).hostname}catch(_){domain=""}
  }
  domain=String(domain||"").toLowerCase().replace(/^www\./,"").split(".")[0];
  return domain.replace(/[^a-z0-9]/g,"");
}

function structuredCompanyLineKey(line){
  return String(line||"")
    .toLowerCase()
    .replace(/\b(?:llc|l\.l\.c\.?|inc\.?|corp\.?|corporation|pllc|lp|llp|ltd\.?|company|co\.?)\b/g," ")
    .replace(/[^a-z0-9]/g,"");
}

function structuredCompanyMatchesDomain(line,senderEmail,site){
  const a=structuredCompanyLineKey(line);
  const b=structuredDomainKey(senderEmail,site);
  if(!a||!b||a.length<3||b.length<3)return false;
  if(a===b)return true;
  // Allow a modest legal/branding suffix difference, but not generic title substrings.
  return a.length>=6&&b.length>=6&&(a.startsWith(b)||b.startsWith(a));
}

function structuredWebsiteLike(line){
  const v=String(line||"").trim();
  return /^(?:https?:\/\/|www\.)/i.test(v)||
    /\b[a-z0-9][a-z0-9-]*(?:\.[a-z0-9-]+)*\.(?:com|net|org|co|us|io|biz|info)\b/i.test(v);
}


function structuredLocationOnlyLine(line){
  const v=String(line||"").replace(/\u00a0/g," ").replace(/\s+/g," ").trim();
  if(!v||/\d/.test(v))return false;
  const parts=v.split(/\s*(?:\||•|·)\s*/).map(x=>x.trim()).filter(Boolean);
  if(!parts.length)return false;
  return parts.every(p=>{
    const m=p.match(/^([A-Za-z][A-Za-z .'-]{1,60}),\s*([A-Za-z][A-Za-z .]{1,30})$/);
    return !!(m&&structuredStateAbbrev(m[2]));
  });
}
function structuredDelimitedParts(line){
  return String(line||"").split(/\s*(?:\||•|·|—|–)\s*/).map(x=>x.trim()).filter(Boolean);
}

function structuredPlausibleTitleText(v){
  const x=String(v||"").trim();
  if(!x||x.length>90||isClosingPhrase(x)||cleanEmail(x)||structuredWebsiteLike(x))return false;
  if(structuredPhoneObservations(x).length||structuredLocationOnlyLine(x))return false;
  if(/\d{5}(?:-\d{4})?/.test(x))return false;
  if(/\b(?:street|st\.?|road|rd\.?|avenue|ave\.?|boulevard|blvd\.?|drive|dr\.?|lane|ln\.?|way|court|ct\.?|broadway|suite|ste\.?|p\.?o\.?\s*box)\b/i.test(x)&&/\d/.test(x))return false;
  if(/[.!?]$/.test(x)&&x.split(/\s+/).length>7)return false;
  return true;
}

function structuredDiscoverSignatures(body,currentName,currentEmail,myEmail){
  if(!body)return[];
  let xs=allPhysicalSignatureCandidates(body,currentName||"",currentEmail||"",myEmail);
  if(!xs.length)xs=recoverCurrentSenderSignatureCandidates(body,currentName||"",currentEmail||"",myEmail);
  if(!xs.length)xs=recoverCurrentSenderNameCompanyCandidates(body,currentName||"",currentEmail||"",myEmail);
  return xs;
}

function structuredCandidateMeta(body,items,currentName,currentEmail){
  const segments=allMessageSegments(body,currentName||"",currentEmail||"");
  const counters=new Map();
  return [...items].sort((a,b)=>(a.globalStart??0)-(b.globalStart??0)).map(c=>{
    const gs=c.globalStart??0;
    let segmentIndex=0;
    for(let i=0;i<segments.length;i++){
      if(gs>=segments[i].start&&gs<=Math.max(segments[i].start,segments[i].end)){segmentIndex=i;break}
      if(gs>segments[i].end)segmentIndex=i;
    }
    const identity=cleanEmail(c.email)||String(c.name||"").toLowerCase().replace(/\s+/g," ").trim()||"unknown";
    const base=segmentIndex+"|"+identity;
    const occurrence=(counters.get(base)||0)+1;
    counters.set(base,occurrence);
    return {candidate:c,key:base+"|"+occurrence,segmentIndex,occurrence,identity};
  });
}

function mergeSignatureCandidateRepresentations(textBody,textItems,htmlBody,htmlItems,currentName,currentEmail){
  const ta=structuredCandidateMeta(textBody,textItems,currentName,currentEmail);
  const ha=structuredCandidateMeta(htmlBody,htmlItems,currentName,currentEmail);
  const map=new Map();
  for(const x of ta)map.set(x.key,{...x,sourceRank:0});
  for(const x of ha)if(!map.has(x.key))map.set(x.key,{...x,sourceRank:1});
  return [...map.values()]
    .sort((a,b)=>a.segmentIndex-b.segmentIndex||a.occurrence-b.occurrence||a.sourceRank-b.sourceRank)
    .map(x=>x.candidate);
}


const V4_CREDENTIAL_EXCLUSIONS=new Set([
  "LLC","L.L.C.","INC","INC.","CORP","CORP.","PLLC","LP","LLP","LTD","LTD.","CO","CO."
]);

function v4DisplayText(value){
  return String(value||"")
    .replace(/\u0091/g,"‘").replace(/\u0092/g,"’")
    .replace(/\u0093/g,"“").replace(/\u0094/g,"”")
    .replace(/\u0095/g,"•").replace(/\u0096/g,"–")
    .replace(/\u0097/g,"—").replace(/\u0085/g,"…");
}
function v4VisibleSignatureBlock(signatureText){
  return v4DisplayText(signatureDiscoveryText(norm(signatureText||"")))
    .split("\n")
    .map(x=>String(x||"").replace(/\u00a0/g," ").replace(/[ \t]+$/g,""))
    .join("\n")
    .replace(/^\s*\n+|\n+\s*$/g,"");
}

function v4LooksCredentialPart(value){
  const raw=String(value||"").trim().replace(/^[|•·—–,\s]+|[|•·—–,\s]+$/g,"");
  if(!raw||raw.length>32)return false;
  const up=raw.toUpperCase();
  if(V4_CREDENTIAL_EXCLUSIONS.has(up))return false;
  if(/^(?:AIA|FAIA|RA|R\.A\.|PE|P\.E\.|SE|S\.E\.|NCARB|PMP|CPA|CCM|DBIA|CSI|CDT|RCDD|CEM|CPHC|CPHT|JD|J\.D\.|MBA|M\.ARCH\.?|B\.ARCH\.?|PHD|PH\.D\.|ESQ\.?|LEED(?:\s+(?:AP(?:\s+(?:BD\+C|ID\+C|O\+M|ND|HOMES))?|GA))?|WELL\s+AP)$/i.test(raw))
    return true;
  if(structuredCompanyEvidence(raw))return false;
  if(/^(?:[A-Z]\.?){2,7}$/.test(raw) && !V4_CREDENTIAL_EXCLUSIONS.has(up))return true;
  return false;
}

function v4CredentialList(value){
  const raw=String(value||"").trim().replace(/^[,;|\s]+|[,;|\s]+$/g,"");
  if(!raw)return[];
  const parts=raw.split(/\s*,\s*/).map(x=>x.trim()).filter(Boolean);
  if(parts.length&&parts.every(v4LooksCredentialPart))return parts;
  if(v4LooksCredentialPart(raw))return[raw];
  return[];
}

function v4FindCredentials(lines,resolvedName){
  const found=[];
  const add=x=>{for(const c of v4CredentialList(x))if(!found.some(y=>y.toLowerCase()===c.toLowerCase()))found.push(c)};
  const rn=cleanName(resolvedName||"");
  const np=nameParts(rn);
  const surname=np.surname||"";
  const nameRx=rn?new RegExp("^\\s*"+rn.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")+"\\b","i"):null;
  const surnameRx=surname?new RegExp("\\b"+surname.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")+"\\b","i"):null;

  const senderComma=String(resolvedName||"").indexOf(",");
  if(senderComma>0)add(String(resolvedName).slice(senderComma+1));

  for(let i=0;i<lines.length;i++){
    const line=String(lines[i]||"").trim();
    if(!line)continue;
    const parts=structuredDelimitedParts(line);

    if(parts.length>1 && (nameRx?.test(parts[0]) || (surnameRx?.test(parts[0])&&personNameFromLine(parts[0])))){
      for(const p of parts.slice(1))add(p);
    }

    if(nameRx?.test(line)){
      const rest=line.replace(nameRx,"").replace(/^[\s,]+/,"");
      if(rest)add(rest);
    }else if(surnameRx?.test(line)&&personNameFromLine(line)){
      const comma=line.indexOf(",");
      if(comma>0)add(line.slice(comma+1));
    }

    if(v4CredentialList(line).length){
      const prev=String(lines[i-1]||"");
      if(nameRx?.test(prev)||surnameRx?.test(prev))add(line);
    }
  }
  return found;
}

function v4RecognizedNonFieldLine(line){
  const v=v4DisplayText(String(line||"")).replace(/\u00a0/g," ").replace(/\s+/g," ").trim();
  if(!v)return"blank";
  if(/^[,;:|•·\s]+$/.test(v)||/^[-_=|•·\s]+$/.test(v)||/^=+\s*[A-Z]?\s*=+$/.test(v))return"decorative";
  if(/^(?:T|P|O|M|C|D|F|E|W)\s*:$/i.test(v))return"contact label";
  if(/^\[?external\]?$/i.test(v))return"system banner";
  if(/^(?:ig|instagram|facebook|linkedin|twitter|x|houzz|vimeo)\b/i.test(v))return"social media";
  if(/\b(?:instagram|facebook|linkedin|vimeo|houzz)\b/i.test(v)&&/[|•▪]/.test(v))return"social media";
  if(/^licensed in\b/i.test(v))return"licensure note";
  if(/^(?:apply for|click here|follow us|connect with|learn more|visit us|schedule|book now)\b/i.test(v))return"marketing/action";
  if(/\b(?:credit account|making it happen for our customers)\b/i.test(v))return"marketing/action";
  if(/^24\s*\/\s*7\s*\/\s*365$/i.test(v))return"availability/tagline";
  if(structuredLocationOnlyLine(v))return"office locations";
  const locParts=v.split(/\s*(?:\||•|▪|·)\s*/).map(x=>x.trim()).filter(Boolean);
  if(locParts.length>=2&&!/\d|@|https?:|www\./i.test(v)&&
     locParts.every(p=>/^[A-Za-z][A-Za-z .'-]{1,40}$/.test(p))&&
     (v===v.toUpperCase()||locParts.length>=3))return"office locations";
  if(/^[“"'‘].*[”"'’]?(?:\s*[-—].*)?$/.test(v)&&v.split(/\s+/).length>=6)return"quotation/tagline";
  if(v.split(/\s+/).length>=8&&/[.!?]$/.test(v)&&
     !cleanEmail(v)&&!structuredPhoneObservations(v).length&&!structuredWebsiteLike(v))return"tagline";
  return"";
}

function v4RemoveCI(haystack,needle){
  const n=String(needle||"").trim();
  if(!n)return String(haystack||"");
  return String(haystack||"").replace(new RegExp(n.replace(/[.*+?^${}()|[\]\\]/g,"\\$&"),"ig")," ");
}
function v4RemoveWordCI(haystack,needle){
  const n=String(needle||"").trim();
  if(!n)return String(haystack||"");
  return String(haystack||"").replace(new RegExp("\\b"+n.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")+"\\b","ig")," ");
}

function v4AuditSignature(signatureText,parsed,resolvedName,credentials){
  const original=v4VisibleSignatureBlock(signatureText);
  const lines=norm(original).split("\n").map(x=>String(x||"").trim()).filter(Boolean);
  const recognized=[];
  const unaccounted=[];
  const phoneRx=/(?:\+?1[\s.\-]?)?(?:\(?\d{3}\)?[\s.\-]*)\d{3}[\s.\-]*\d{4}(?:\s*(?:x|ext\.?|extension)\s*[:.\-]?\s*\d+)?/gi;
  const emailRx=/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi;
  const urlRx=/(?:https?:\/\/|www\.)[^\s|]+|\b[a-z0-9][a-z0-9-]*(?:\.[a-z0-9-]+)*\.(?:com|net|org|co|us|io|biz|info)\b/gi;

  for(const line of lines){
    const nonField=v4RecognizedNonFieldLine(line);
    if(nonField){recognized.push({text:line,category:nonField});continue}
    let residue=line;

    // Strip machine-identifiable spans before textual fields so a company
    // name inside an email/domain is not removed first and turned into residue.
    residue=residue.replace(emailRx," ").replace(urlRx," ").replace(phoneRx," ");

    residue=v4RemoveCI(residue,resolvedName);
    const possibleName=personNameFromLine(line)||signatureNameFromLine(line);
    const auditSurname=nameParts(resolvedName||"").surname||"";
    if(possibleName && auditSurname &&
       new RegExp("\\b"+auditSurname.replace(/[.*+?^${}()|[\\]\\]/g,"\\$&")+"\\b","i").test(line))
      residue=v4RemoveCI(residue,possibleName);

    for(const field of ["jobTitle","companyName","street","city","postalCode"]){
      if(parsed[field])residue=v4RemoveCI(residue,parsed[field]);
    }
    if(parsed.state){
      const full=Object.entries(US_STATE_NAMES).find(([,abbr])=>abbr===parsed.state)?.[0]||"";
      if(full)residue=v4RemoveWordCI(residue,full);
      residue=v4RemoveWordCI(residue,parsed.state);
    }
    if(parsed.countryOrRegion)residue=v4RemoveCI(residue,parsed.countryOrRegion);
    for(const c of credentials||[])residue=v4RemoveCI(residue,c);
    residue=residue
      .replace(/\b(?:office|business|work|main|telephone|phone|tel|direct\s*\/\s*text|direct|mobile|cell|cellular|fax|email|website|web|ext(?:ension)?|IG)\b\s*[:.\-]?/gi," ")
      .replace(/(?:^|[\s|•;,])(?:P|O|M|C|D|F|E|W)\s*:\s*/gi," ")
      .replace(/\(\s*(?:P|O|M|C|D|F|E|W|Direct|Office|Mobile|Cell|Fax)\s*\)/gi," ")
      .replace(/\bx\d+\b/gi," ");
    if(structuredPhoneObservations(line).length)
      residue=residue.replace(/\b(?:P|O|M|C|D|F)\b\s*:?\s*/gi," ");
    if(cleanEmail(line))
      residue=residue.replace(/^\s*E\s*:?\s*/i," ");
    if(structuredWebsiteLike(line))
      residue=residue.replace(/^\s*W\s*:?\s*/i," ");

    residue=residue.replace(/[|•▪·—–,:;()[\]{}<>/\\]+/g," ").replace(/[“”"'‘’]/g," ").replace(/\s+/g," ").trim();
    residue=residue.replace(/^(?:and|the|of|at|for)\b\s*/i,"").replace(/\b(?:and|the|of|at|for)$/i,"").trim();

    if(residue&&/[A-Za-z0-9]/.test(residue))unaccounted.push({line,text:residue});
  }

  const unique=[];
  for(const u of unaccounted){
    const key=u.text.toLowerCase();
    if(!unique.some(x=>x.text.toLowerCase()===key))unique.push(u);
  }
  return{recognizedNonField:recognized,unaccounted:unique,unaccountedText:unique.map(x=>x.text).join("\n")};
}


function v4CompanyFromDomain(senderEmail,site){
  let domain=cleanEmail(senderEmail).split("@")[1]||"";
  if(!domain&&site){
    try{domain=new URL(/^https?:\/\//i.test(site)?site:"https://"+site).hostname}catch(_){domain=""}
  }
  domain=String(domain||"").toLowerCase().replace(/^www\./,"");
  if(!domain||/^(?:gmail|outlook|hotmail|live|icloud|me|aol|yahoo|protonmail|msn)\./i.test(domain))return"";
  let stem=(domain.split(".")[0]||"").replace(/[-_]+/g," ").trim();
  if(!stem)return"";

  // Domains often concatenate brand + industry word. Split a broad but finite
  // set of common organization suffixes while preserving the brand stem.
  if(!/\s/.test(stem)){
    const suffixes=[
      "technologies","technology","environmental","engineering","architecture",
      "construction","contractors","contractor","concrete","electrical","electric",
      "mechanical","plumbing","planning","solutions","services","systems","rentals",
      "equipment","builders","building","design","studio","consulting","consultants",
      "management","development","developers","properties","industries","industrial",
      "workshop","garage"
    ].sort((a,b)=>b.length-a.length);
    for(const s of suffixes){
      if(stem.endsWith(s)&&stem.length>s.length+1){
        stem=stem.slice(0,-s.length)+" "+s;
        break;
      }
    }
  }
  return stem.split(/\s+/).filter(Boolean).map(w=>{
    if(w.length<=2)return w.toUpperCase();
    return w.charAt(0).toUpperCase()+w.slice(1);
  }).join(" ");
}

function v4ResidualTitleCompany(parsed,signatureText,resolvedName,credentials){
  const sig=v4VisibleSignatureBlock(signatureText);
  const lines=norm(sig).split("\n").map(x=>String(x||"").replace(/\s+/g," ").trim()).filter(Boolean);
  const email=parsed.email||"",site=parsed.businessHomePage||"";
  const np=nameParts(resolvedName||""),surname=np.surname||"";

  if(parsed.jobTitle&&(structuredLocationOnlyLine(parsed.jobTitle)||v4RecognizedNonFieldLine(parsed.jobTitle)||
     cleanEmail(parsed.jobTitle)||structuredPhoneObservations(parsed.jobTitle).length))parsed.jobTitle="";

  if(parsed.jobTitle){
    const kept=structuredDelimitedParts(parsed.jobTitle).filter(x=>!v4CredentialList(x).length);
    parsed.jobTitle=kept.join(" | ").trim();
  }

  const atoms=[];
  lines.forEach((line,i)=>{const parts=structuredDelimitedParts(line);for(const p of(parts.length?parts:[line]))atoms.push({i,text:p,line})});

  const isIdentity=a=>{
    const t=a.text;
    if(resolvedName&&t.toLowerCase().includes(resolvedName.toLowerCase()))return true;
    if(surname&&new RegExp("\\b"+surname.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")+"\\b","i").test(t)&&personNameFromLine(t))return true;
    return false;
  };
  const machine=a=>{
    const t=a.text;
    return cleanEmail(t)||structuredPhoneObservations(t).length||structuredWebsiteLike(t)||/\d{5}(?:-\d{4})?/.test(t)||
      (parsed.street&&t.toLowerCase().includes(parsed.street.toLowerCase()))||
      structuredLocationOnlyLine(t)||v4RecognizedNonFieldLine(t)||v4CredentialList(t).length;
  };
  const candidates=atoms.filter(a=>!isIdentity(a)&&!machine(a)&&structuredPlausibleTitleText(a.text));

  const companies=candidates.map(a=>{
    let score=0;
    if(structuredCompanyMatchesDomain(a.text,email,site))score+=12;
    if(structuredCompanyEvidence(a.text))score+=6;
    if(a.text===a.text.toUpperCase()&&a.text.split(/\s+/).length>=2)score+=2;
    return{...a,score};
  }).filter(a=>a.score>0).sort((a,b)=>b.score-a.score||a.i-b.i);

  // Credentials and role text are never company material.
  if(parsed.companyName&&v4CredentialList(parsed.companyName).length)parsed.companyName="";
  if(parsed.companyName&&parsed.jobTitle&&
     parsed.companyName.trim().toLowerCase()===parsed.jobTitle.trim().toLowerCase())
    parsed.companyName="";

  let currentCompany=String(parsed.companyName||"");
  let currentStrong=structuredCompanyEvidence(currentCompany)||structuredCompanyMatchesDomain(currentCompany,email,site);

  // If the current company looks like a human role and has no organization/domain
  // evidence, discard it and let the whole-block evidence choose again.
  if(currentCompany&&!currentStrong&&structuredPlausibleTitleText(currentCompany)){
    parsed.companyName="";
    currentCompany="";
    currentStrong=false;
  }

  if(companies.length&&(!currentCompany||!currentStrong)){
    parsed.companyName=companies[0].line.includes("|")&&structuredCompanyMatchesDomain(companies[0].line,email,site)
      ? companies[0].line : companies[0].text;
  }

  // Some signatures use only a logo/website for the company. If no explicit
  // company survives, infer the organization from the sender/site domain.
  if(!parsed.companyName){
    const inferredCompany=v4CompanyFromDomain(email,site);
    if(inferredCompany)parsed.companyName=inferredCompany;
  }

  if(parsed.jobTitle){
    const current=parsed.jobTitle.toLowerCase();
    const richer=candidates.filter(a=>{
      const low=a.text.toLowerCase();
      if(structuredCompanyEvidence(a.text)||structuredCompanyMatchesDomain(a.text,email,site))return false;
      if(v4CredentialList(a.text).length||v4RecognizedNonFieldLine(a.text))return false;
      return low!==current && low.includes(current) && a.text.length>parsed.jobTitle.length+2;
    }).sort((a,b)=>a.i-b.i);
    if(richer.length)parsed.jobTitle=richer[0].text;
  }

  if(!parsed.jobTitle){
    const companyLine=lines.findIndex(x=>parsed.companyName&&x.toLowerCase().includes(parsed.companyName.toLowerCase()));
    const titles=candidates.filter(a=>{
      if(parsed.companyName&&(a.text.toLowerCase()===parsed.companyName.toLowerCase()||a.line.toLowerCase()===parsed.companyName.toLowerCase()))return false;
      if(structuredCompanyEvidence(a.text)||structuredCompanyMatchesDomain(a.text,email,site))return false;
      return true;
    }).map(a=>{
      let score=1;
      if(a.i<=2)score+=2;
      if(companyLine>=0&&a.i<companyLine)score+=4;
      return{...a,score};
    }).sort((a,b)=>b.score-a.score||a.i-b.i);
    if(titles.length)parsed.jobTitle=titles[0].text;
  }
  return parsed;
}


function v4FillPartialAddress(parsed,signatureText){
  if(parsed.street&&parsed.city&&parsed.state)return parsed;
  const lines=norm(signatureText).split("\n")
    .map(x=>String(x||"").replace(/\u00a0/g," ").replace(/\s+/g," ").trim())
    .filter(Boolean);

  const isSeparator=x=>/^[,;|•·]+$/.test(String(x||"").trim())||/^[-_=]{5,}$/.test(String(x||"").trim());
  const meaningfulBefore=i=>{
    const out=[];
    for(let j=i-1;j>=0&&out.length<4;j--){
      if(!lines[j]||isSeparator(lines[j]))continue;
      out.push(lines[j]);
    }
    return out;
  };
  const meaningfulAfter=i=>{
    const out=[];
    for(let j=i+1;j<lines.length&&out.length<3;j++){
      if(!lines[j]||isSeparator(lines[j]))continue;
      out.push(lines[j]);
    }
    return out;
  };

  // Fragmented Outlook address layout:
  //   255 E. Anderson Street
  //   ,
  //   Idaho Falls
  //   ,
  //   ID 83401
  //   ,
  //   United States
  for(let i=0;i<lines.length;i++){
    const m=lines[i].match(/^([A-Za-z]{2,30})\s+(\d{5}(?:-\d{4})?)$/);
    if(!m)continue;
    const st=structuredStateAbbrev(m[1]);
    if(!st)continue;
    const prev=meaningfulBefore(i);
    const city=prev[0]||"";
    const street=prev[1]||"";
    const streetLike=/\b(?:street|st\.?|road|rd\.?|avenue|ave\.?|boulevard|blvd\.?|drive|dr\.?|lane|ln\.?|way|court|ct\.?|place|pl\.?|loop|terrace|ter\.?|highway|hwy\.?|broadway|p\.?o\.?\s*box)\b/i;
    if(street&&city&&/\d/.test(street)&&streetLike.test(street)&&
       /^[A-Za-z][A-Za-z .'-]{1,60}$/.test(city)){
      parsed.street=parsed.street||street.trim();
      parsed.city=parsed.city||city.trim();
      parsed.state=parsed.state||st;
      parsed.postalCode=parsed.postalCode||m[2];
      const after=meaningfulAfter(i);
      if(after.some(x=>/^United States(?: of America)?$/i.test(x))||st)
        parsed.countryOrRegion=parsed.countryOrRegion||"United States";
      if(parsed.street&&parsed.city&&parsed.state)return parsed;
    }
  }

  for(const line of lines){
    if(cleanEmail(line)||structuredPhoneObservations(line).length||structuredWebsiteLike(line))continue;

    // Standard no-ZIP form: 123 Main Street, Orem, UT
    let m=line.match(/^(.+?),\s*([A-Za-z][A-Za-z .'-]{1,60}),\s*([A-Za-z]{2,30})$/);
    if(m){
      const st=structuredStateAbbrev(m[3]);
      if(st&&(/\d/.test(m[1])||/\b(?:street|st\.?|road|rd\.?|avenue|ave\.?|boulevard|blvd\.?|drive|dr\.?|lane|ln\.?|way|court|ct\.?|place|pl\.?|loop|terrace|ter\.?|highway|hwy\.?)\b/i.test(m[1]))){
        parsed.street=parsed.street||m[1].trim();
        parsed.city=parsed.city||m[2].trim();
        parsed.state=parsed.state||st;
        parsed.countryOrRegion=parsed.countryOrRegion||"United States";
        return parsed;
      }
    }

    // Western grid-address form without a separator before city:
    // 48 N 1330 W Orem, UT
    m=line.match(/^(\d+\s+(?:N|S|E|W)\s+\d+\s+(?:N|S|E|W))\s+([A-Za-z][A-Za-z .'-]{1,60}),\s*([A-Za-z]{2,30})$/i);
    if(m){
      const st=structuredStateAbbrev(m[3]);
      if(st){
        parsed.street=parsed.street||m[1].trim();
        parsed.city=parsed.city||m[2].trim();
        parsed.state=parsed.state||st;
        parsed.countryOrRegion=parsed.countryOrRegion||"United States";
        return parsed;
      }
    }

    // Street-suffix form with city immediately following the street and only
    // the state separated by a comma.
    m=line.match(/^(.+?\b(?:street|st\.?|road|rd\.?|avenue|ave\.?|boulevard|blvd\.?|drive|dr\.?|lane|ln\.?|way|court|ct\.?|place|pl\.?|loop|terrace|ter\.?|highway|hwy\.?))\s+([A-Za-z][A-Za-z .'-]{1,60}),\s*([A-Za-z]{2,30})$/i);
    if(m&&/\d/.test(m[1])){
      const st=structuredStateAbbrev(m[3]);
      if(st){
        parsed.street=parsed.street||m[1].trim();
        parsed.city=parsed.city||m[2].trim();
        parsed.state=parsed.state||st;
        parsed.countryOrRegion=parsed.countryOrRegion||"United States";
        return parsed;
      }
    }
  }
  return parsed;
}

function v4ContactFromSignature(senderName,senderEmail,signatureText){
  const original=v4VisibleSignatureBlock(signatureText);
  const base=structuredContactFromSignature(senderName,senderEmail,original);
  v4FillPartialAddress(base,original);
  const senderLooksHuman=looksLikePersonName(senderName||"");
  const inferred=inferPersonName(original);
  const resolvedName=senderLooksHuman?cleanName(senderName):(inferred||"");
  const credentials=v4FindCredentials(norm(original).split("\n"),resolvedName);

  v4ResidualTitleCompany(base,original,resolvedName,credentials);

  const np=nameParts(resolvedName||[base.givenName,base.middleName,base.surname].filter(Boolean).join(" "));
  if(np.givenName)base.givenName=np.givenName;
  if(np.middleName)base.middleName=np.middleName;
  if(np.surname)base.surname=np.surname+(credentials.length?", "+credentials.join(", "):"");

  base.signature=original;
  base.personalNotes=original;

  let audit=v4AuditSignature(original,base,resolvedName,credentials);
  if(audit.unaccounted.length&&(!base.jobTitle||!base.companyName)){
    v4ResidualTitleCompany(base,original,resolvedName,credentials);
    audit=v4AuditSignature(original,base,resolvedName,credentials);
  }

  base._audit={credentials,recognizedNonField:audit.recognizedNonField,unaccounted:audit.unaccounted,unaccountedText:audit.unaccountedText};
  base._debug={...(base._debug||{}),v4HolisticParser:true,credentials,unaccountedLines:audit.unaccounted};
  return base;
}


function v4LikelyAddressContinuation(lines,index){
  const clean=i=>String(lines[i]||"").replace(/\u00a0/g," ").replace(/\s+/g," ").trim();
  const v=clean(index);
  if(!v)return false;
  if(/^[,;]+$/.test(v))return true;
  if(/^United States(?: of America)?$/i.test(v))return true;
  if(/^([A-Za-z]{2,30})\\s+\\d{5}(?:-\\d{4})?$/.test(v) &&
     structuredStateAbbrev(v.match(/^([A-Za-z]{2,30})/)[1]))return true;
  if(/\\d/.test(v)&&/\\b(?:street|st\\.?|road|rd\\.?|avenue|ave\\.?|boulevard|blvd\\.?|drive|dr\\.?|lane|ln\\.?|way|court|ct\\.?|place|pl\\.?|loop|terrace|ter\\.?|highway|hwy\\.?|broadway|p\\.?o\\.?\\s*box)\\b/i.test(v))return true;

  // City fragment between a street line and a state/ZIP line.
  if(/^[A-Za-z][A-Za-z .'-]{1,60}$/.test(v)){
    let prev="",next="";
    for(let j=index-1;j>=0;j--){
      const x=clean(j); if(!x||/^[,;]+$/.test(x))continue; prev=x; break;
    }
    for(let j=index+1;j<lines.length;j++){
      const x=clean(j); if(!x||/^[,;]+$/.test(x))continue; next=x; break;
    }
    const prevStreet=/\\d/.test(prev)&&/\\b(?:street|st\\.?|road|rd\\.?|avenue|ave\\.?|boulevard|blvd\\.?|drive|dr\\.?|lane|ln\\.?|way|court|ct\\.?|place|pl\\.?|loop|terrace|ter\\.?|highway|hwy\\.?|broadway|p\\.?o\\.?\\s*box)\\b/i.test(prev);
    const nextState=/^([A-Za-z]{2,30})\\s+\\d{5}(?:-\\d{4})?$/.test(next) &&
      structuredStateAbbrev(next.match(/^([A-Za-z]{2,30})/)[1]);
    if(prevStreet&&nextState)return true;
  }
  return false;
}

function v4SignatureTailCategory(line){
  const v=String(line||"").trim();
  if(!v)return"";
  if(v4RecognizedNonFieldLine(v))return"recognized";
  if(/^https?:\/\/|^www\./i.test(v))return"website";
  return"";
}

function v4AttachOriginalSignatureBlock(body,candidate){
  if(!candidate?.parsed||candidate.candidateType==="referral")return candidate;
  const lines=norm(body||"").split("\n");
  let start=Number.isInteger(candidate.globalStart)?candidate.globalStart:-1;
  let end=Number.isInteger(candidate.globalEnd)?candidate.globalEnd:-1;
  if(start<0||end<start||end>=lines.length){
    candidate.rawSignatureBlock=v4VisibleSignatureBlock(candidate.parsed.signature||candidate.text||"");
    return candidate;
  }

  // Some email-anchored candidates begin at company/title even though the
  // person's linked display name is visibly just above them. Recover that
  // human-name line for the preserved original block.
  const identityName=candidate.name||[candidate.parsed?.givenName,candidate.parsed?.middleName,candidate.parsed?.surname].filter(Boolean).join(" ");
  const idSurname=nameParts(identityName||"").surname||"";
  for(let j=start-1;j>=Math.max(0,start-10);j--){
    const v=v4DisplayText(signatureDiscoveryText(String(lines[j]||""))).trim();
    if(!v)continue;
    if(/^\s*(?:from|sent|to|cc|bcc|subject):\s*/i.test(v)||/^[-_]{5,}$/.test(v))break;
    const surnameHit=idSurname&&new RegExp("\\b"+idSurname.replace(/[.*+?^${}()|[\\]\\]/g,"\\$&")+"\\b","i").test(v);
    if((identityName&&nameMatchesIdentity(v,identityName))||(surnameHit&&personNameFromLine(v))){
      start=j;break;
    }
    if(isClosingPhrase(v))break;
  }

  let finalEnd=end;
  for(let j=end+1;j<Math.min(lines.length,end+28);j++){
    const v=String(lines[j]||"").trim();
    if(!v)continue;
    if(/^\s*(?:from|sent|to|cc|bcc|subject):\s*/i.test(v))break;

    // A decorative divider immediately before a quoted-message header is not
    // part of the signature. Otherwise decorative lines may be intentional
    // signature design and can precede a website/logo footer.
    if(/^[-_]{5,}$/.test(v)){
      let next="";
      for(let k=j+1;k<Math.min(lines.length,j+8);k++){
        const x=String(lines[k]||"").trim(); if(!x)continue; next=x; break;
      }
      if(/^\s*(?:from|sent|to|cc|bcc|subject):\s*/i.test(next))break;
      finalEnd=j;
      continue;
    }

    if(v4SignatureTailCategory(v)||v4LikelyAddressContinuation(lines,j)){
      finalEnd=j;
      continue;
    }
    break;
  }
  const block=signatureDiscoveryText(lines.slice(start,finalEnd+1).join("\n")).replace(/^\s*\n+|\n+\s*$/g,"");
  candidate.rawSignatureBlock=block||v4VisibleSignatureBlock(candidate.parsed.signature||candidate.text||"");
  return candidate;
}


function structuredContactFromSignature(senderName,senderEmail,signatureText){
  const sig=cleanSignatureText(signatureText);
  const lines=norm(sig).split("\n")
    .map(x=>String(x||"").replace(/\u00a0/g," ").replace(/\s+/g," ").trim())
    .filter(Boolean);

  const senderLooksHuman=looksLikePersonName(senderName||"");
  const inferred=inferPersonName(sig);
  const resolvedName=senderLooksHuman?cleanName(senderName):(inferred||"");
  const email=cleanEmail(senderEmail)||cleanEmail(sig);
  const site=website(sig,email);
  const addr=structuredAddress(sig);
  const ph=structuredPhones(sig);
  const np=nameParts(resolvedName||"");

  const assigned=new Set();
  let nameLine=-1,companyLine=-1,titleLine=-1;
  let companyValue="",titleValue="";
  const socialNoise=/^(?:ig|instagram|facebook|linkedin|twitter|x)\b/i;

  // Locate the human name line, allowing authoritative full names to map
  // to nickname signatures via the surname.
  for(let i=0;i<lines.length;i++){
    const v=lines[i];
    if(resolvedName&&v.toLowerCase().startsWith(resolvedName.toLowerCase())){
      nameLine=i;assigned.add(i);break;
    }
    if(np.surname){
      const rx=new RegExp("\\b"+np.surname.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")+"\\b","i");
      if(rx.test(v)&&personNameFromLine(v)){nameLine=i;assigned.add(i);break}
    }
  }

  // Name-line semantics. Handle:
  //   Lana Kozlovskaya | Designer
  //   Susan Johnson | SJ Planning Solutions LLC
  //   Jason Ro | Principal | RA, AIA
  //   Tyson Cox | Sunbelt Rentals, Inc. | Outside sales manager
  if(nameLine>=0){
    const raw=lines[nameLine];
    const pieces=structuredDelimitedParts(raw);
    if(pieces.length>1){
      const rest=pieces.slice(1);
      let companyIdx=-1,best=-1;
      rest.forEach((p,i)=>{
        let score=0;
        if(structuredCompanyMatchesDomain(p,email,site))score+=10;
        if(structuredCompanyEvidence(p))score+=5;
        if(score>best){best=score;companyIdx=score>0?i:-1}
      });
      if(companyIdx>=0){
        companyValue=rest[companyIdx];
        const titleParts=rest.filter((_,i)=>i!==companyIdx).filter(structuredPlausibleTitleText);
        if(titleParts.length)titleValue=titleParts.join(" | ");
      }else{
        const titleParts=rest.filter(structuredPlausibleTitleText);
        if(titleParts.length)titleValue=titleParts.join(" | ");
      }
    }else{
      const split=splitNameTitleLine(raw);
      if(split&&split.title&&structuredPlausibleTitleText(split.title))titleValue=split.title;
    }
  }

  // Machine-identifiable lines and location-only office lists are assigned
  // before residual reasoning, so "New York, NY | White Plains, NY" can
  // never become a job title.
  lines.forEach((v,i)=>{
    if(i===nameLine)return;
    if(cleanEmail(v)||structuredPhoneObservations(v).length||structuredWebsiteLike(v))assigned.add(i);
    if(addr.postalCode&&v.includes(addr.postalCode))assigned.add(i);
    if(addr.street&&v.toLowerCase().includes(addr.street.toLowerCase()))assigned.add(i);
    if(socialNoise.test(v)||structuredLocationOnlyLine(v))assigned.add(i);
  });

  // Composite non-name lines. This is the important Eric Zagar case:
  //   Technical Associate | CUONO ENGINEERING, PLLC
  // but preserves brand forms such as:
  //   RO | ROCKETT DESIGN
  for(let i=0;i<lines.length;i++){
    if(i===nameLine||assigned.has(i))continue;
    const v=lines[i],parts=structuredDelimitedParts(v);
    if(parts.length<2)continue;

    const companyParts=[];
    parts.forEach((p,j)=>{
      let score=0;
      if(structuredCompanyMatchesDomain(p,email,site))score+=10;
      if(structuredCompanyEvidence(p))score+=5;
      if(score>0)companyParts.push({j,p,score});
    });
    companyParts.sort((a,b)=>b.score-a.score);

    if(companyParts.length){
      const cp=companyParts[0];
      const wholeMatches=structuredCompanyMatchesDomain(v,email,site);
      const otherParts=parts.filter((_,j)=>j!==cp.j);

      // A short brand prefix plus an organization fragment is one company
      // name (e.g. RO | ROCKETT DESIGN), not Title | Company.
      const shortBrand=wholeMatches&&otherParts.length&&otherParts.every(x=>x.length<=4&&!/\s/.test(x));
      if(!companyValue)companyValue=shortBrand?v:cp.p;
      if(!titleValue&&!shortBrand){
        const tp=otherParts.filter(structuredPlausibleTitleText);
        if(tp.length)titleValue=tp.join(" | ");
      }
      companyLine=i;assigned.add(i);
    }
  }

  // Legacy title recognition is accepted only if it is not a location or
  // company line. This helps same-line name/title cases such as Alex/E1.
  if(!titleValue){
    const legacyTitle=title(sig,resolvedName)||"";
    if(legacyTitle&&structuredPlausibleTitleText(legacyTitle)&&
       !structuredLocationOnlyLine(legacyTitle)&&
       !structuredCompanyEvidence(legacyTitle)&&
       !structuredCompanyMatchesDomain(legacyTitle,email,site)){
      titleValue=legacyTitle;
    }
  }

  // Company pass: exact domain agreement first, then organization evidence.
  if(!companyValue){
    const cc=[];
    for(let i=0;i<lines.length;i++){
      if(i===nameLine||assigned.has(i)||i===titleLine)continue;
      const v=lines[i];
      if(!v||isClosingPhrase(v)||cleanEmail(v)||structuredPhoneObservations(v).length||
         structuredWebsiteLike(v)||socialNoise.test(v)||structuredLocationOnlyLine(v))continue;
      if(/\d{5}(?:-\d{4})?/.test(v)||/\b(?:street|st\.?|road|rd\.?|avenue|ave\.?|boulevard|blvd\.?|drive|dr\.?|lane|ln\.?|way|court|ct\.?|broadway|suite|ste\.?|p\.?o\.?\s*box)\b/i.test(v))continue;
      let score=0;
      if(structuredCompanyMatchesDomain(v,email,site))score+=10;
      if(structuredCompanyEvidence(v))score+=5;
      if(v===v.toUpperCase()&&v.split(/\s+/).length>=2)score+=2;
      if(nameLine>=0&&i>nameLine&&i<=nameLine+3)score+=1;
      if(score>0)cc.push({i,v,score});
    }
    cc.sort((a,b)=>b.score-a.score||a.i-b.i);
    if(cc.length){
      companyValue=cc[0].v;companyLine=cc[0].i;assigned.add(companyLine);
    }
  }

  // Mark exact standalone title/company lines as assigned.
  if(titleValue){
    for(let i=0;i<lines.length;i++){
      if(i===nameLine)continue;
      if(lines[i].toLowerCase()===titleValue.toLowerCase()){
        titleLine=i;assigned.add(i);break;
      }
    }
  }
  if(companyValue&&companyLine<0){
    for(let i=0;i<lines.length;i++){
      if(i===nameLine)continue;
      if(lines[i].toLowerCase()===companyValue.toLowerCase()){
        companyLine=i;assigned.add(i);break;
      }
    }
  }

  // SECOND PASS: inspect every unexplained line for a missing title.
  // Unassigned text is evidence, not something to silently discard.
  if(!titleValue){
    const tc=[];
    for(let i=0;i<lines.length;i++){
      if(assigned.has(i)||i===companyLine||i===nameLine)continue;
      const v=lines[i];
      if(!structuredPlausibleTitleText(v))continue;
      if(structuredCompanyMatchesDomain(v,email,site)||structuredCompanyEvidence(v))continue;
      let score=1;
      if(nameLine>=0&&i>nameLine)score+=2;
      if(companyLine>=0&&i<companyLine)score+=4;
      if(i<=2)score+=1;
      tc.push({i,v,score});
    }
    tc.sort((a,b)=>b.score-a.score||a.i-b.i);
    if(tc.length){
      titleValue=tc[0].v;titleLine=tc[0].i;assigned.add(titleLine);
    }
  }

  // THIRD PASS: inspect remaining unexplained text for company identity.
  if(!companyValue){
    const cc=[];
    for(let i=0;i<lines.length;i++){
      if(assigned.has(i)||i===nameLine||i===titleLine)continue;
      const v=lines[i];
      if(!structuredPlausibleTitleText(v))continue;
      let score=0;
      if(structuredCompanyMatchesDomain(v,email,site))score+=10;
      if(structuredCompanyEvidence(v))score+=5;
      if(titleLine>=0&&i>titleLine&&i<=titleLine+2)score+=2;
      if(nameLine>=0&&i>nameLine&&i<=nameLine+3)score+=1;
      if(score>0)cc.push({i,v,score});
    }
    cc.sort((a,b)=>b.score-a.score||a.i-b.i);
    if(cc.length){
      companyValue=cc[0].v;companyLine=cc[0].i;assigned.add(companyLine);
    }
  }

  companyValue=companyValue||companyFromDomain(email,site)||"";

  const result=Object.assign({},nameParts(resolvedName),{
    companyName:companyValue,
    jobTitle:String(titleValue||"").replace(/^[-|•·—–\s]+|[-|•·—–\s]+$/g,"").trim(),
    email
  },ph,{businessHomePage:site},addr,{signature:sig,personalNotes:sig});

  result._debug={
    structuredParser:true,
    unassignedLines:lines.filter((_,i)=>!assigned.has(i))
  };
  return result;
}
function parseContactFromSignature(senderName,senderEmail,signatureText){
  const sig=cleanSignatureText(signatureText);
  const inferred=inferPersonName(sig);
  const senderLooksHuman=looksLikePersonName(senderName||"");
  const resolvedName=senderLooksHuman?cleanName(senderName):(inferred||"");
  const site=website(sig,senderEmail);
  return Object.assign(
    {},
    nameParts(resolvedName),
    {companyName:company(sig,resolvedName,senderEmail,site),jobTitle:title(sig,resolvedName),email:cleanEmail(senderEmail)},
    phones(sig),
    {businessHomePage:site},
    address(sig),
    {signature:sig,personalNotes:sig}
  );
}

function parseContact(senderName,senderEmail,segmentText){
  return parseContactFromSignature(senderName,senderEmail,isolateSignature(segmentText,senderEmail));
}
function sameEmail(a,b){return cleanEmail(a)&&cleanEmail(a)===cleanEmail(b)}
function cleanSignatureText(sig){
  const cleanLine=line=>{
    let v=String(line||"").replace(/[\u00ad\u200b-\u200d\ufeff]/g,"").trim();
    if(!v)return "";
    if(/^(?:https?:\/\/)?[^\s]+\.(?:png|jpe?g|gif|svg|webp|bmp|ico)(?:[?#].*)?$/i.test(v))return "";
    if(/^cid:|^data:image/i.test(v))return "";
    v=v.replace(/<https?:\/\/[^>]*(?:urldefense\.proofpoint\.com|safelinks\.protection\.outlook\.com)[^>]*>/ig," ");
    v=v.replace(/https?:\/\/[^\s]*(?:urldefense\.proofpoint\.com|safelinks\.protection\.outlook\.com)[^\s]*/ig," ");
    v=v.replace(/<mailto:[^>]+>/ig," ");
    v=v.replace(/\s+/g," ").trim();
    if(/(?:google\.[^/]+\/maps|maps\.google\.|maps\.apple\.|bing\.com\/maps|goo\.gl\/maps)/i.test(v)&&!cleanEmail(v))return "";
    return v;
  };
  let lines=norm(sig).split("\n").map(cleanLine).filter(Boolean);
  while(lines.length&&isClosingPhrase(lines[0]))lines.shift();
  return lines.join("\n");
}
function splitMessage(body,currentName,currentEmail){
  const text=norm(body),lines=text.split("\n"),headers=[];
  for(let i=0;i<lines.length;i++){
    const m=lines[i].match(/^\s*From:\s*(.+)$/i);
    if(!m)continue;
    const id=headerIdentity(m[1]);
    if(id.email)headers.push({i,name:id.name,email:id.email});
  }

  const firstHeader=headers.length?headers[0].i:lines.length;
  const out=[{name:cleanName(currentName||"Current sender"),email:cleanEmail(currentEmail),text:lines.slice(0,firstHeader).join("\n"),headerAuthoritative:true}];

  for(let h=0;h<headers.length;h++){
    const a=headers[h],b=headers[h+1]?headers[h+1].i:lines.length;
    out.push({name:a.name,email:a.email,text:lines.slice(a.i+1,b).join("\n"),headerAuthoritative:true});
  }

  const bestByKey=new Map();
  for(const seg of out){
    const key=seg.email||String(seg.name||"").toLowerCase();
    if(!key)continue;
    const sig=isolateSignature(seg.text,seg.email);
    let score=norm(sig).split("\n").reduce((t,l)=>t+signatureScore(l),0);
    if(seg.email&&norm(sig).toLowerCase().includes(seg.email.toLowerCase()))score+=6;
    if(seg.headerAuthoritative&&seg.name&&looksLikePersonName(seg.name))score+=30;
    const prev=bestByKey.get(key);
    if(!prev||score>prev.score)bestByKey.set(key,{score,seg});
  }
  return [...bestByKey.values()].map(x=>x.seg);
}


function plainTextOccurrenceCandidates(body,currentName,currentEmail,myEmail){
  const lines=norm(body).split("\n").map(x=>x.trim());
  const identities=[];
  const seenIdentity=new Set();

  for(const line of lines){
    const m=line.match(/^\s*From:\s*(.+)$/i);
    if(!m)continue;
    const id=headerIdentity(m[1]);
    const key=cleanEmail(id.email);
    if(key&&!sameEmail(key,myEmail)&&looksLikePersonName(id.name||"")&&!seenIdentity.has(key)){
      seenIdentity.add(key);identities.push({name:id.name,email:key,authoritative:true});
    }
  }

  const ce=cleanEmail(currentEmail);
  if(ce&&!sameEmail(ce,myEmail)&&looksLikePersonName(currentName||"")&&!seenIdentity.has(ce)){
    seenIdentity.add(ce);identities.push({name:cleanName(currentName),email:ce,authoritative:true});
  }

  const out=[];
  const identityEmails=new Set();
  for(const id of identities){
    const occs=identitySignatureOccurrences(body,id.name,id.email);
    if(!occs.length)continue;
    identityEmails.add(id.email);
    for(const o of occs){
      const parsed=parseContactFromSignature(id.name,id.email,o.sig);
      const np=nameParts(id.name);
      parsed.givenName=np.givenName;parsed.middleName=np.middleName;parsed.surname=np.surname;
      parsed.email=id.email;
      parsed.jobTitle=sanitizeJobTitle(parsed.jobTitle,id.name,o.sig);
      parsed.signature=o.sig;
      parsed.personalNotes=o.sig;
      parsed._debug={source:"Bounded signature occurrence + forwarded From identity",headerAuthoritative:true,headerName:id.name,occurrence:o.occurrence,textSignature:o.sig,htmlSignature:""};
      out.push({score:o.score,name:id.name,email:id.email,text:o.sig,parsed,headerAuthoritative:true,occurrenceKey:`identity:${id.email}:${o.start}:${o.end}`,occurrence:o.occurrence});
    }
  }

  // Generic fallback for people whose forwarded From header is unavailable. Keep occurrences
  // separate rather than collapsing identical email addresses.
  const isHeader=l=>/^\s*(from|sent|to|cc|bcc|subject):\s*/i.test(l)||/^[-_]{5,}$/.test(l);
  const isDisclaimer=l=>/confidential|privileged|intended recipient|virus|disclaimer|please consider the environment/i.test(l);
  let occurrence=0;
  for(let i=0;i<lines.length;i++){
    if(isHeader(lines[i]))continue;
    const email=cleanEmail(lines[i]);
    if(!email||sameEmail(email,myEmail)||identityEmails.has(email))continue;
    let start=-1;
    for(let j=i-1;j>=Math.max(0,i-10);j--){
      const v=lines[j];
      if(isHeader(v)||isDisclaimer(v))break;
      if(!v||isClosingPhrase(v))continue;
      if(signatureNameFromLine(v)){start=j;break}
    }
    if(start<0)continue;
    let end=i;
    const sig=cleanSignatureText(lines.slice(start,end+1).filter(Boolean).join("\n"));
    if(!sig)continue;
    const name=signatureNameFromLine(lines[start])||inferPersonName(sig);
    if(!name||!looksLikePersonName(name))continue;
    const parsed=parseContactFromSignature(name,email,sig);
    parsed.jobTitle=sanitizeJobTitle(parsed.jobTitle,name,sig);
    parsed.signature=sig;parsed.personalNotes=sig;
    out.push({score:50,name,email,text:sig,parsed,headerAuthoritative:false,occurrenceKey:`generic:${i}:${occurrence}`,occurrence});
    occurrence++;
  }
  return out;
}

function headerSignatureCandidates(body,currentName,currentEmail,myEmail){
  const out=[];
  let segmentIndex=0;
  for(const seg of splitMessage(body,currentName,currentEmail)){
    if(!seg.email||sameEmail(seg.email,myEmail)||!looksLikePersonName(seg.name||"")){segmentIndex++;continue}
    let occs=signatureOccurrences(seg.text,seg.email);
    if(!occs.length){
      const fallback=cleanSignatureText(isolateSignature(seg.text,seg.email));
      if(fallback)occs=[{sig:fallback,score:norm(fallback).split("\n").reduce((t,l)=>t+signatureScore(l),0),occurrence:0}];
    }
    for(const o of occs){
      const sig=cleanSignatureText(o.sig);
      if(!sig)continue;
      const parsed=parseContactFromSignature(seg.name,seg.email,sig);
      const np=nameParts(seg.name);
      parsed.givenName=np.givenName;
      parsed.middleName=np.middleName;
      parsed.surname=np.surname;
      parsed.email=seg.email;
      parsed.signature=sig;
      parsed.personalNotes=sig;
      parsed._debug={
        source:"Forwarded From header authoritative — signature occurrence "+(o.occurrence+1),
        headerAuthoritative:true,
        headerName:seg.name,
        occurrence:o.occurrence,
        htmlSignature:"",
        textSignature:sig
      };
      out.push({
        score:100+(o.score||0),name:seg.name,email:seg.email,text:sig,parsed,
        headerAuthoritative:true,occurrenceKey:`header:${segmentIndex}:${o.occurrence}`,
        occurrence:o.occurrence
      });
    }
    segmentIndex++;
  }
  return out;
}


function inferAnchoredName(sig,email){
  const lines=norm(sig).split("\n").map(x=>x.trim()).filter(Boolean);
  const emailIdx=lines.findIndex(l=>cleanEmail(l)===cleanEmail(email));
  const stop=emailIdx>=0?emailIdx:lines.length;
  for(let i=stop-1;i>=Math.max(0,stop-8);i--){const found=personNameFromLine(lines[i]);if(found)return found}
  return inferPersonName(sig);
}

function currentSenderSignatureFromBody(body,currentName){
  const lines=norm(body).split("\n").map(x=>x.trim());
  const wanted=String(currentName||"").toLowerCase().replace(/[^a-z0-9]+/g," ").trim();
  if(!wanted)return "";
  const isHeader=l=>/^\s*(from|sent|to|cc|bcc|subject):\s*/i.test(l)||/^[-_]{5,}$/.test(l);
  const isDisclaimer=l=>/confidential|privileged|intended recipient|virus|disclaimer|please consider the environment/i.test(l);
  const normKey=v=>String(v||"").toLowerCase().replace(/[^a-z0-9]+/g," ").trim();
  const hits=[];
  for(let i=0;i<lines.length;i++){
    const key=normKey(lines[i]);
    // Accept "Linda Shin | Associate", "Linda Shin - Architect", etc.
    if(key===wanted || key.startsWith(wanted+" "))hits.push(i);
  }
  let best=null;
  for(const start of hits){
    const kept=[];
    let blankRun=0;
    for(let j=start;j<lines.length && kept.length<18;j++){
      const v=lines[j];
      if(j>start && (isHeader(v)||isDisclaimer(v)))break;
      if(!v){
        blankRun++;
        if(blankRun>=2 && kept.length>=4)break;
        continue;
      }
      blankRun=0;
      kept.push(v);
      // Once we have a substantial signature, stop before obvious prose that follows it.
      if(kept.length>=6 && v.length>120 && /[.!?]$/.test(v))break;
    }
    const sig=cleanSignatureText(kept.join("\n"));
    const parsed=parseContactFromSignature(currentName,"",sig);
    const evidence=[
      !!(parsed.companyName),
      !!(parsed.businessPhone||parsed.mobilePhone),
      !!(parsed.state&&parsed.postalCode),
      !!parsed.businessHomePage,
      !!parsed.jobTitle
    ].filter(Boolean).length;
    const score=signatureScore(lines[start]||"")+evidence*3+norm(sig).split("\n").filter(Boolean).length/10;
    if(sig && (!best||score>best.score))best={score,sig};
  }
  return best?best.sig:"";
}

function anchoredSignatureCandidates(body,currentName,currentEmail,myEmail){
  const lines=norm(body).split("\n").map(x=>x.trim());
  const isHeader=l=>/^\s*(from|sent|to|cc|bcc|subject):\s*/i.test(l)||/^[-_]{5,}$/.test(l);
  const isDisclaimer=l=>/confidential|privileged|intended recipient|virus|disclaimer|please consider the environment/i.test(l);

  const authoritative=headerSignatureCandidates(body,currentName,currentEmail,myEmail);
  const authoritativeEmails=new Set(authoritative.map(x=>cleanEmail(x.email)).filter(Boolean));
  const byEmail=new Map();

  // Current Outlook sender (non-user) when there is no forwarded-header candidate for it.
  if(currentEmail&&!sameEmail(currentEmail,myEmail)&&!authoritativeEmails.has(cleanEmail(currentEmail))&&looksLikePersonName(currentName||"")){
    const directSig=currentSenderSignatureFromBody(body,currentName);
    if(directSig){
      const parsed=parseContactFromSignature(currentName,currentEmail,directSig);
      const np=nameParts(currentName);
      parsed.givenName=np.givenName;parsed.middleName=np.middleName;parsed.surname=np.surname;
      parsed.email=cleanEmail(currentEmail);
      parsed.signature=cleanSignatureText(directSig);parsed.personalNotes=parsed.signature;
      parsed._debug={source:"Current sender — Outlook From header + visible name anchor",htmlSignature:"",textSignature:parsed.signature};
      const directScore=norm(parsed.signature).split("\n").reduce((t,l)=>t+signatureScore(l),0)+10;
      byEmail.set(cleanEmail(currentEmail),{score:directScore,name:currentName,email:cleanEmail(currentEmail),text:parsed.signature,parsed});
    }
  }

  // Generic visible-email discovery for messages without an authoritative forwarded From header.
  for(let i=0;i<lines.length;i++){
    if(isHeader(lines[i]))continue;
    const email=cleanEmail(lines[i]);
    if(!email||sameEmail(email,myEmail)||authoritativeEmails.has(email))continue;

    const occs=signatureOccurrences(lines.join("\n"),email).filter(o=>o.anchor===i);
    let sig="";
    if(occs.length)sig=occs[0].sig;
    if(!sig){
      let lo=i,hi=i,blankBudget=1;
      for(let j=i-1;j>=0&&i-j<=10;j--){const v=lines[j];if(isHeader(v)||isDisclaimer(v))break;if(!v){if(blankBudget--<=0)break;continue}lo=j}
      blankBudget=1;
      for(let j=i+1;j<lines.length&&j-i<=8;j++){const v=lines[j];if(isHeader(v)||isDisclaimer(v))break;if(!v){if(blankBudget--<=0)break;continue}hi=j}
      let chunk=lines.slice(lo,hi+1).filter(Boolean);
      let emailLocal=chunk.findIndex(l=>cleanEmail(l)===email);if(emailLocal<0)emailLocal=chunk.length-1;
      let nameLocal=-1;
      for(let j=emailLocal-1;j>=Math.max(0,emailLocal-8);j--){if(!isClosingPhrase(chunk[j])&&signatureNameFromLine(chunk[j])){nameLocal=j;break}}
      if(nameLocal>=0)chunk=chunk.slice(nameLocal);else while(chunk.length>5&&signatureScore(chunk[0])===0)chunk.shift();
      sig=cleanSignatureText(chunk.join("\n"));
    }
    const score=norm(sig).split("\n").reduce((t,l)=>t+signatureScore(l),0);
    if(score<5)continue;
    const inferred=inferAnchoredName(sig,email);
    const name=(sameEmail(email,currentEmail)&&looksLikePersonName(currentName||""))?currentName:(inferred||email.split("@")[0]);
    const parsed=parseContactFromSignature(name,email,sig);
    parsed.signature=sig;parsed.personalNotes=sig;
    const existing=byEmail.get(email);
    if(!existing||score>existing.score)byEmail.set(email,{score,name,email,text:sig,parsed});
  }

  return [...authoritative,...byEmail.values()].map(x=>({
    name:x.name,email:x.email,text:x.text,parsed:x.parsed,
    headerAuthoritative:!!x.headerAuthoritative,occurrenceKey:x.occurrenceKey||"",occurrence:x.occurrence
  }));
}


function allMessageSegments(body,currentName,currentEmail){
  const lines=norm(body).split("\n");
  const headers=[];
  for(let i=0;i<lines.length;i++){
    const m=lines[i].match(/^\s*From:\s*(.+)$/i);
    if(!m)continue;
    const id=headerIdentity(m[1]);
    if(id.email||id.name)headers.push({i,name:id.name,email:cleanEmail(id.email)});
  }
  const segments=[];
  const firstHeader=headers.length?headers[0].i:lines.length;
  segments.push({start:0,end:firstHeader-1,name:cleanName(currentName||""),email:cleanEmail(currentEmail),authoritative:true,text:lines.slice(0,firstHeader).join("\n")});
  for(let h=0;h<headers.length;h++){
    const a=headers[h],b=headers[h+1]?headers[h+1].i:lines.length;
    segments.push({start:a.i+1,end:b-1,name:cleanName(a.name||""),email:cleanEmail(a.email),authoritative:true,text:lines.slice(a.i+1,b).join("\n")});
  }
  return segments;
}
function signatureEvidenceScore(lines){
  const text=lines.join("\n");let score=0;
  if(cleanEmail(text))score+=4;
  if(/\b(?:office|direct|mobile|cell|fax|tel|phone)\b/i.test(text)||/\d{3}[\s.\-]\d{3}[\s.\-]\d{4}/.test(text))score+=3;
  if(/\b[A-Z]{2}\s+\d{5}(?:-\d{4})?\b/.test(text))score+=3;
  if(/\b(?:street|st\.?|road|rd\.?|avenue|ave\.?|boulevard|blvd\.?|drive|dr\.?|lane|ln\.?|loop|suite|ste\.?)\b/i.test(text))score+=2;
  if(/https?:\/\/|www\./i.test(text))score+=2;
  if(TITLE_WORDS.some(t=>text.toLowerCase().includes(t)))score+=2;
  if(COMPANY_WORDS.some(w=>(" "+text.toLowerCase()).includes(w)))score+=1;
  return score;
}
function nameMatchesIdentity(line,identityName){
  const raw=cleanName(signatureNameFromLine(line)||personNameFromLine(line)||"");
  const wanted=cleanName(identityName||"");
  if(!raw||!wanted)return false;
  const key=v=>String(v||"").toLowerCase().replace(/[^a-z0-9]+/g," ").trim();
  const a=key(raw),b=key(wanted);if(a===b)return true;
  const ap=a.split(/\s+/).filter(Boolean),bp=b.split(/\s+/).filter(Boolean);
  return ap.length>=2&&bp.length>=2&&ap[ap.length-1]===bp[bp.length-1];
}

function authoritativeSignatureStart(lines,j,identityName){
  const parts=nameParts(identityName||"");
  const surname=String(parts.surname||"").toLowerCase();
  const line=String(lines[j]||"").replace(/[\u00ad\u200b-\u200d\ufeff]/g,"").trim();
  if(!line)return -1;

  const split=splitNameTitleLine(line);
  if(split&&split.name){
    const sp=nameParts(split.name);
    if(!surname||String(sp.surname||"").toLowerCase()===surname)return j;
  }
  if(looksLikePersonName(line)){
    const p=nameParts(line);
    if(!surname||String(p.surname||"").toLowerCase()===surname)return j;
  }

  if(surname){
    const next=String(lines[j+1]||"").replace(/[\u00ad\u200b-\u200d\ufeff]/g,"").trim();
    if(/^[A-Za-z][A-Za-z'.-]*$/.test(line)&&next.toLowerCase().startsWith(surname+" ")){
      const rem=next.slice(parts.surname.length).trim();
      if(TITLE_WORDS.some(t=>rem.toLowerCase().includes(t))||rem.endsWith("&"))return j;
    }
    const low=line.toLowerCase();
    if(low.startsWith(surname+" ")){
      const rem=line.slice(parts.surname.length).trim();
      if(TITLE_WORDS.some(t=>rem.toLowerCase().includes(t))||rem.endsWith("&")){
        const prev=String(lines[j-1]||"").replace(/[\u00ad\u200b-\u200d\ufeff]/g,"").trim();
        if(/^[A-Za-z][A-Za-z'.-]*$/.test(prev))return j-1;
        return j;
      }
    }
  }
  return -1;
}

function boundedSignatureAtEmail(lines,emailIndex,identityName,identityEmail){
  const isHeader=l=>/^\s*(from|sent|to|cc|bcc|subject):\s*/i.test(l)||/^[-_]{5,}$/.test(l);
  const isDisclaimer=l=>/confidential|privileged|intended recipient|virus|disclaimer|please consider the environment/i.test(l);
  let start=-1;
  for(let j=emailIndex-1;j>=Math.max(0,emailIndex-12);j--){
    const v=String(lines[j]||"").replace(/[\u00ad\u200b-\u200d\ufeff]/g,"").trim();
    if(isHeader(v)||isDisclaimer(v))break;
    if(!v)continue;
    const a=authoritativeSignatureStart(lines,j,identityName);
    if(a>=0){start=a;break}
    if(isClosingPhrase(v))break;
    const generic=signatureNameFromLine(v);
    if(generic){start=j;break}
  }
  if(start<0)return null;
  let end=emailIndex;
  for(let j=emailIndex+1;j<Math.min(lines.length,emailIndex+5);j++){
    const v=String(lines[j]||"").replace(/[\u00ad\u200b-\u200d\ufeff]/g,"").trim();
    if(!v||isHeader(v)||isDisclaimer(v)||isClosingPhrase(v))break;
    if(cleanEmail(v)&&!sameEmail(cleanEmail(v),identityEmail))break;
    if(authoritativeSignatureStart(lines,j,identityName)>=0||signatureNameFromLine(v))break;
    if(/^https?:|^www\./i.test(v)||signatureScore(v)>0){end=j;continue}
    break;
  }
  const raw=lines.slice(start,end+1).map(x=>String(x||"").replace(/[\u00ad\u200b-\u200d\ufeff]/g,"").trim()).filter(Boolean);
  const sig=cleanSignatureText(raw.join("\n"));
  if(!sig||signatureEvidenceScore(raw)<5)return null;
  return {start,end,sig};
}
function boundedSignatureAtName(lines,nameIndex,identityName,identityEmail){
  const isHeader=l=>/^\s*(from|sent|to|cc|bcc|subject):\s*/i.test(l)||/^[-_]{5,}$/.test(l);
  const isDisclaimer=l=>/confidential|privileged|intended recipient|virus|disclaimer|please consider the environment/i.test(l);
  const raw=[];
  let lastIncluded=nameIndex-1;
  let blankRun=0;

  // Outlook often inserts blank lines between every visual line of a signature.
  // Blank lines therefore are not, by themselves, a signature boundary.
  for(let j=nameIndex;j<Math.min(lines.length,nameIndex+22);j++){
    const v=String(lines[j]||"")
      .replace(/[\u00ad\u200b-\u200d\ufeff]/g,"")
      .trim();

    if(j>nameIndex&&(isHeader(v)||isDisclaimer(v)))break;

    if(!v){
      blankRun++;
      // Three consecutive blank lines after a substantial block are a safe stop.
      if(blankRun>=3&&raw.length>=4)break;
      continue;
    }
    blankRun=0;

    if(j>nameIndex&&isClosingPhrase(v))break;
    if(j>nameIndex&&nameMatchesIdentity(v,identityName))break;

    // A horizontal divider ends the physical signature block.
    if(/^[-_]{5,}$/.test(v))break;

    raw.push(v);
    lastIncluded=j;
  }

  const sig=cleanSignatureText(raw.join("\n"));
  if(!sig||signatureEvidenceScore(raw)<5)return null;
  return {start:nameIndex,end:lastIncluded,sig};
}
function signatureOccurrencesInSegment(segment,myEmail){
  const identityEmail=cleanEmail(segment.email),identityName=cleanName(segment.name||"");
  if(identityEmail&&sameEmail(identityEmail,myEmail))return [];
  const lines=norm(segment.text||"").split("\n"),out=[],used=[];
  const overlaps=(a,b)=>used.some(r=>!(b<r.start||a>r.end));
  const add=(block,source)=>{
    if(!block||overlaps(block.start,block.end))return;
    const parsed=parseContactFromSignature(identityName,identityEmail,block.sig);
    if(identityName&&looksLikePersonName(identityName)){const np=nameParts(identityName);parsed.givenName=np.givenName;parsed.middleName=np.middleName;parsed.surname=np.surname}
    parsed.email=identityEmail||cleanEmail(parsed.email);parsed.jobTitle=sanitizeJobTitle(parsed.jobTitle,identityName||[parsed.givenName,parsed.surname].filter(Boolean).join(" "),block.sig);parsed.signature=block.sig;parsed.personalNotes=block.sig;parsed._debug={source,physicalOccurrence:true};
    used.push({start:block.start,end:block.end});
    out.push({score:100,name:identityName||inferPersonName(block.sig),email:identityEmail||cleanEmail(parsed.email),text:block.sig,parsed,physicalOccurrence:true,occurrenceKey:`segment:${segment.start}:${block.start}:${block.end}`,localStart:block.start,localEnd:block.end});
  };
  if(identityEmail){for(let i=0;i<lines.length;i++)if(sameEmail(cleanEmail(lines[i]),identityEmail))add(boundedSignatureAtEmail(lines,i,identityName,identityEmail),"Physical signature occurrence — email anchor")}
  if(identityName){for(let i=0;i<lines.length;i++)if(nameMatchesIdentity(lines[i],identityName))add(boundedSignatureAtName(lines,i,identityName,identityEmail),"Physical signature occurrence — name anchor")}
  return out;
}
function genericPhysicalSignatureOccurrences(body,myEmail,claimedGlobalRanges){
  const lines=norm(body).split("\n");const out=[];const claimed=claimedGlobalRanges||[];
  const isHeader=l=>/^\s*(from|sent|to|cc|bcc|subject):\s*/i.test(l)||/^[-_]{5,}$/.test(l);
  const isClaimed=(a,b)=>claimed.some(r=>!(b<r.start||a>r.end));
  for(let i=0;i<lines.length;i++){
    if(isHeader(lines[i]))continue;const email=cleanEmail(lines[i]);if(!email||sameEmail(email,myEmail))continue;
    let name="",start=-1;for(let j=i-1;j>=Math.max(0,i-12);j--){if(isHeader(lines[j]))break;const n=signatureNameFromLine(lines[j])||personNameFromLine(lines[j]);if(n&&!isClosingPhrase(lines[j])){name=n;start=j;break}}
    if(start<0)continue;const block=boundedSignatureAtEmail(lines,i,name,email);if(!block||isClaimed(block.start,block.end))continue;
    const parsed=parseContactFromSignature(name,email,block.sig);parsed.signature=block.sig;parsed.personalNotes=block.sig;parsed.jobTitle=sanitizeJobTitle(parsed.jobTitle,name,block.sig);
    out.push({score:60,name,email,text:block.sig,parsed,physicalOccurrence:true,occurrenceKey:`generic:${block.start}:${block.end}`,globalStart:block.start,globalEnd:block.end});claimed.push({start:block.start,end:block.end});
  }
  return out;
}
function allPhysicalSignatureCandidates(body,currentName,currentEmail,myEmail){
  const segments=allMessageSegments(body,currentName,currentEmail),out=[],ranges=[];
  for(const seg of segments){
    for(const c of signatureOccurrencesInSegment(seg,myEmail)){
      c.globalStart=seg.start+(c.localStart||0);c.globalEnd=seg.start+(c.localEnd||0);ranges.push({start:c.globalStart,end:c.globalEnd});out.push(c);
    }
  }
  out.push(...genericPhysicalSignatureOccurrences(body,myEmail,ranges));
  out.sort((a,b)=>(a.globalStart??0)-(b.globalStart??0));
  return out;
}

function referralLineText(v){
  return String(v||"")
    .replace(/[\u00ad\u200b-\u200d\ufeff]/g,"")
    .replace(/\u00a0/g," ")
    .replace(/\s+/g," ")
    .trim();
}

function referralEmailLine(line){
  const raw=referralLineText(line);
  const email=cleanEmail(raw);
  if(!email)return "";
  const residue=raw.replace(new RegExp(email.replace(/[.*+?^${}()|[\]\\]/g,"\\$&"),"i"),"")
    .replace(/^(?:e|email)\s*[:\-]?\s*/i,"")
    .replace(/[<>\s]/g,"");
  return residue?"":email;
}

function referralPersonPhoneLine(line){
  const raw=referralLineText(line);
  // Conservative form: "First Last: phone".  At least two name tokens.
  const m=raw.match(/^([A-Z][A-Za-z'’.\-]+(?:\s+(?:[A-Z]\.?|[A-Z][A-Za-z'’.\-]+)){1,4})\s*:\s*(.+)$/);
  if(!m)return null;

  const name=cleanName(m[1]);
  const phone=phoneForOutlook(m[2]);
  const digitsOnly=phone.replace(/\D/g,"");
  if(!looksLikePersonName(name)||digitsOnly.length<10||digitsOnly.length>15)return null;
  return {name,phone};
}

function plausibleReferralCompanyLine(line){
  const raw=referralLineText(line);
  if(!raw||raw.length<2||raw.length>80)return false;
  if(cleanEmail(raw)||/^https?:|^www\./i.test(raw)||/\d{3}[\s().\-]\d{3}/.test(raw))return false;
  if(/^\s*(?:from|sent|to|cc|bcc|subject)\s*:/i.test(raw))return false;
  if(isClosingPhrase(raw))return false;
  if(/[!?]$/.test(raw))return false;
  if(raw.split(/\s+/).length>8)return false;
  if(/^(?:hi|hello|dear|thanks?|thank you|regards?|best|references?|please|attached|updated|talk|worked|yes|no)\b/i.test(raw))return false;
  return /[A-Za-z]/.test(raw);
}

function nextReferralNonblank(lines,start,maxLook=4){
  for(let i=start;i<Math.min(lines.length,start+maxLook);i++){
    if(referralLineText(lines[i]))return i;
  }
  return -1;
}

function referredContactCandidates(body,myEmail){
  const lines=norm(body).split("\n");
  const out=[];

  for(let i=0;i<lines.length;i++){
    const company=referralLineText(lines[i]);
    if(!plausibleReferralCompanyLine(company))continue;

    const personIdx=nextReferralNonblank(lines,i+1,4);
    if(personIdx<0)continue;
    const pp=referralPersonPhoneLine(lines[personIdx]);
    if(!pp)continue;

    const emailIdx=nextReferralNonblank(lines,personIdx+1,4);
    if(emailIdx<0)continue;
    const email=referralEmailLine(lines[emailIdx]);
    if(!email||sameEmail(email,myEmail))continue;

    // Keep this deliberately narrow: the three logical lines are the complete
    // referred-contact block. Do not infer title, address, or homepage.
    const block=[company,referralLineText(lines[personIdx]),referralLineText(lines[emailIdx])].join("\n");
    const np=nameParts(pp.name);
    const parsed={
      givenName:np.givenName,
      middleName:np.middleName,
      surname:np.surname,
      companyName:company,
      jobTitle:"",
      email,
      businessPhone:pp.phone,
      mobilePhone:"",
      businessFax:"",
      businessHomePage:"",
      street:"",
      city:"",
      state:"",
      postalCode:"",
      countryOrRegion:"",
      signature:block,
      personalNotes:block,
      _debug:{source:"Contact information found in message",referralContact:true}
    };

    out.push({
      score:100,
      name:pp.name,
      email,
      text:block,
      parsed,
      candidateType:"referral",
      physicalOccurrence:true,
      occurrenceKey:`referral:${i}:${personIdx}:${emailIdx}`,
      globalStart:i,
      globalEnd:emailIdx
    });
  }

  return out;
}

function candidateOccurrenceLabels(items){
  const totals=new Map();for(const c of items){const k=cleanEmail(c.email)||String(c.name||"").toLowerCase();totals.set(k,(totals.get(k)||0)+1)}
  const seen=new Map();return items.map(c=>{const k=cleanEmail(c.email)||String(c.name||"").toLowerCase();const n=(seen.get(k)||0)+1;seen.set(k,n);return{n,total:totals.get(k)||1}});
}

function renderCandidates(){
  const box=$("candidates");box.innerHTML="";const labels=candidateOccurrenceLabels(candidates);
  candidates.forEach((c,i)=>{
    const lab=labels[i];
    const occ=c.candidateType==="referral"
      ? `<div class="muted">Contact information found in message</div>`
      : (lab.total>1?`<div class="muted">Signature occurrence ${lab.n} of ${lab.total}</div>`:`<div class="muted">Signature occurrence</div>`);
    const el=document.createElement("div");el.className="candidate";
    el.innerHTML=`<div class="candidate-head"><input type="checkbox" data-candidate="${i}"><div><div class="candidate-name">${html(c.name||"Unknown sender")}</div><div class="muted">${html(c.email||"No email found")}</div>${occ}</div></div><div class="preview">${html(c.parsed.signature||"No contact block confidently found")}</div>`;
    const cb=el.querySelector('input[data-candidate]');cb.addEventListener("change",()=>{if(!cb.checked)return;document.querySelectorAll('input[data-candidate]').forEach(x=>{if(x!==cb)x.checked=false})});box.appendChild(el)
  });
  $("candidateSection").hidden=false
}
function html(s){return String(s||"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[c]))}
function visibleSignatureTextFromNode(node){
  try{
    const clone=node.cloneNode(true);
    clone.querySelectorAll("script,style,noscript").forEach(n=>n.remove());
    clone.querySelectorAll("a").forEach(a=>{
      const href=(a.getAttribute("href")||"").trim();
      let visible=(a.textContent||"").replace(/\s+/g," ").trim();
      if(!visible){
        if(/^mailto:/i.test(href)) visible=decodeURIComponent(href.replace(/^mailto:/i,"").split("?")[0]);
        else if(/^tel:/i.test(href)) visible=decodeURIComponent(href.replace(/^tel:/i,"").split("?")[0]);
        else if(/^https?:/i.test(href)){const decoded=decodeProofpointUrl(href);if(!isJunkResourceUrl(decoded)) visible=normalizeWebsiteUrl(decoded)}
      }
      a.replaceWith(clone.ownerDocument.createTextNode(visible?` ${visible} `:" "));
    });
    clone.querySelectorAll("img").forEach(img=>{
      const alt=(img.getAttribute("alt")||img.getAttribute("title")||"").replace(/\s+/g," ").trim();
      const safeAlt=isJunkResourceUrl(alt)?"":alt;
      img.replaceWith(clone.ownerDocument.createTextNode(safeAlt?` ${safeAlt} `:" "));
    });
    clone.querySelectorAll("br").forEach(br=>br.replaceWith(clone.ownerDocument.createTextNode("\n")));
    clone.querySelectorAll("tr,p,div,li,table,blockquote").forEach(el=>el.appendChild(clone.ownerDocument.createTextNode("\n")));
    // Separate table cells with a space rather than a blank line; this preserves same-row phone labels.
    clone.querySelectorAll("td,th").forEach(el=>el.appendChild(clone.ownerDocument.createTextNode(" ")));
    return norm(clone.textContent||"")
      .replace(/[\u200B-\u200D\uFEFF]/g,"")
      .split("\n")
      .map(x=>x.replace(/[ \t]+/g," ").trim())
      .filter(Boolean)
      .join("\n");
  }catch(_){return ""}
}
function htmlBodyToText(htmlText){
  try{
    const doc=new DOMParser().parseFromString(htmlText||"","text/html");
    return visibleSignatureTextFromNode(doc.body);
  }catch(_){return ""}
}
function signatureContainerScore(text,email){
  const t=norm(text);
  if(!t||!sameEmail(cleanEmail(t),email) && !t.toLowerCase().includes(String(email||"").toLowerCase()))return -999;
  const lines=t.split("\n").map(x=>x.trim()).filter(Boolean);
  let score=lines.reduce((sum,l)=>sum+signatureScore(l),0);
  if(address(t).state)score+=4;
  const ph=phones(t);if(ph.businessPhone||ph.mobilePhone)score+=4;
  if(inferPersonName(t))score+=2;
  if(lines.length>=4&&lines.length<=18)score+=3;
  if(lines.length>30)score-=10;
  if(/^(from|sent|to|cc|subject):/im.test(t))score-=12;
  return score;
}
function findBestHtmlSignatureContainer(anchor,email){
  // Prefer a complete signature table. Most Outlook signatures are HTML tables and the
  // mailto link may live several nested cells below the name/address/phones.
  const tables=[];
  let t=anchor.closest?anchor.closest("table"):null;
  let hops=0;
  while(t&&hops<4){
    const text=visibleSignatureTextFromNode(t);
    const lines=norm(text).split("\n").map(x=>x.trim()).filter(Boolean);
    const score=signatureContainerScore(text,email);
    if(score>-999 && lines.length>=3 && lines.length<=30)tables.push({node:t,text,score:score+8});
    t=t.parentElement?t.parentElement.closest("table"):null;
    hops++;
  }
  if(tables.length){
    tables.sort((a,b)=>b.score-a.score || b.text.length-a.text.length);
    return tables[0];
  }

  // Fallback for signatures built from DIV/SPAN blocks instead of tables. Prefer the
  // broadest compact ancestor that still looks signature-like, rather than the smallest
  // ancestor containing only the email link.
  let node=anchor,best=null;
  for(let depth=0;node&&depth<10;depth++,node=node.parentElement){
    if(!node||!node.textContent)continue;
    const tag=(node.tagName||"").toLowerCase();
    if(["body","html"].includes(tag))break;
    const text=visibleSignatureTextFromNode(node);
    const lines=norm(text).split("\n").map(x=>x.trim()).filter(Boolean);
    const score=signatureContainerScore(text,email);
    if(score>-999 && lines.length<=30){
      const completeness=(phones(text).businessPhone||phones(text).mobilePhone?5:0)+(address(text).state?5:0)+(inferPersonName(text)?2:0);
      const candidate={node,text,score:score+completeness,lines:lines.length};
      if(!best || candidate.score>best.score || (candidate.score===best.score&&candidate.lines>best.lines))best=candidate;
    }
  }
  return best;
}


function forwardedIdentityMap(...texts){
  const out=new Map();
  for(const text of texts){
    for(const line of norm(text||"").split("\n")){
      const m=line.match(/^\s*From:\s*(.+)$/i);
      if(!m)continue;
      const id=headerIdentity(m[1]);
      const e=cleanEmail(id.email);
      if(e&&looksLikePersonName(id.name||""))out.set(e,cleanName(id.name));
    }
  }
  return out;
}

function htmlBoundedSignatureForAnchor(anchor,email){
  let best=null;
  let node=anchor;
  for(let depth=0;node&&depth<12;depth++,node=node.parentElement){
    if(!node||!node.textContent)continue;
    const tag=(node.tagName||"").toLowerCase();
    if(tag==="body"||tag==="html")break;
    const text=cleanSignatureText(visibleSignatureTextFromNode(node));
    const lines=norm(text).split("\n").map(x=>x.trim()).filter(Boolean);
    if(!text||lines.length<3||lines.length>14)continue;
    if(!lines.some(l=>sameEmail(cleanEmail(l),email)))continue;
    if(/^(?:from|sent|to|cc|bcc|subject):/im.test(text))continue;
    if(/\b(?:good afternoon|good morning|good evening|payment options?|please see attached|please reach out|wire transfer|customer id|billing zip)\b/i.test(text))continue;
    const ph=phones(text),ad=address(text),ttl=title(text,"");
    const evidence=(ph.businessPhone||ph.mobilePhone?3:0)+(ad.state&&ad.postalCode?3:0)+(ttl?2:0)+(inferPersonName(text)?2:0);
    if(evidence<4)continue;
    // Prefer the smallest complete block. It is much less likely to contain message prose or a second signature.
    const score=evidence*20-lines.length;
    if(!best||score>best.score)best={score,text};
  }
  return best?best.text:"";
}

function htmlMailtoOccurrenceCandidates(html,plainText,currentName,currentEmail,myEmail){
  const out=[];
  try{
    const doc=new DOMParser().parseFromString(html||"","text/html");
    const visible=visibleSignatureTextFromNode(doc.body);
    const ids=forwardedIdentityMap(plainText||"",visible||"");
    const currentE=cleanEmail(currentEmail);
    if(currentE&&!sameEmail(currentE,myEmail)&&looksLikePersonName(currentName||""))ids.set(currentE,cleanName(currentName));

    const anchors=[...doc.querySelectorAll('a[href^="mailto:" i]')];
    let serial=0;
    for(const a of anchors){
      const href=(a.getAttribute("href")||"");
      const email=cleanEmail(decodeURIComponent(href.replace(/^mailto:/i,"").split("?")[0])||a.textContent||"");
      if(!email||sameEmail(email,myEmail))continue;

      const sig=htmlBoundedSignatureForAnchor(a,email);
      if(!sig)continue;

      const authoritativeName=ids.get(email)||"";
      const inferred=inferAnchoredName(sig,email)||inferPersonName(sig);
      const name=looksLikePersonName(authoritativeName)?authoritativeName:(inferred||"");
      if(!name||!looksLikePersonName(name))continue;

      const parsed=parseContactFromSignature(name,email,sig);
      if(authoritativeName){
        const np=nameParts(authoritativeName);
        parsed.givenName=np.givenName;parsed.middleName=np.middleName;parsed.surname=np.surname;
      }
      parsed.email=email;
      parsed.jobTitle=sanitizeJobTitle(parsed.jobTitle,name,sig);
      parsed.signature=sig;
      parsed.personalNotes=sig;
      parsed._debug={source:"HTML mailto signature occurrence",headerAuthoritative:!!authoritativeName,headerName:authoritativeName||name,occurrence:serial,textSignature:"",htmlSignature:sig};
      out.push({score:120,name,email,text:sig,parsed,headerAuthoritative:!!authoritativeName,occurrenceKey:`html-mailto:${serial}`,occurrence:serial});
      serial++;
    }
  }catch(_){ }
  return out;
}

function htmlSignatureCandidates(html,currentName,currentEmail,myEmail){
  const out=new Map();
  try{
    const doc=new DOMParser().parseFromString(html||"","text/html");
    const anchors=[...doc.querySelectorAll('a[href^="mailto:" i]')];
    for(const a of anchors){
      const href=(a.getAttribute("href")||"");
      const email=cleanEmail(decodeURIComponent(href.replace(/^mailto:/i,"").split("?")[0])||a.textContent||"");
      if(!email||sameEmail(email,myEmail))continue;
      const best=findBestHtmlSignatureContainer(a,email);
      if(!best||best.score<7)continue;
      const sig=cleanSignatureText(best.text);
      const inferred=inferAnchoredName(sig,email);
      const name=(sameEmail(email,currentEmail)&&looksLikePersonName(currentName||""))?currentName:(inferred||email.split("@")[0]);
      const parsed=parseContactFromSignature(name,email,sig);
      const prev=out.get(email);
      if(!prev||best.score>prev.score)out.set(email,{score:best.score,name,email,text:sig,parsed});
    }
  }catch(_){ }
  return [...out.values()].map(x=>({name:x.name,email:x.email,text:x.text,parsed:x.parsed}));
}
function parsedCompleteness(p){
  if(!p)return 0;
  let score=0;
  for(const k of ["street","city","state","postalCode","countryOrRegion","businessPhone","mobilePhone","companyName","businessHomePage","jobTitle"])if(String(p[k]||"").trim())score+=1;
  const sigLines=norm(p.signature||"").split("\n").map(x=>x.trim()).filter(Boolean).length;
  score+=Math.min(sigLines,10)/10;
  return score;
}
function mergeParsed(primary,secondary,sourceLabel){
  const a={...(primary||{})},b=secondary||{};
  const fill=["givenName","middleName","surname","companyName","jobTitle","email","businessPhone","mobilePhone","businessFax","businessHomePage","street","city","state","postalCode","countryOrRegion"];
  for(const k of fill)if(!String(a[k]||"").trim()&&String(b[k]||"").trim())a[k]=b[k];
  // Country must be inferred independently whenever either parser found a valid U.S. state.
  if(!a.countryOrRegion){const st=(a.state||b.state||"").toUpperCase();if(US_STATES.has(st))a.countryOrRegion="United States"}
  // Prefer the more complete visible signature for Notes. This prevents a small HTML cell
  // containing only the email address from replacing a richer text signature.
  const aSig=cleanSignatureText(a.signature||""),bSig=cleanSignatureText(b.signature||"");
  const aLines=aSig.split("\n").filter(Boolean).length,bLines=bSig.split("\n").filter(Boolean).length;
  const aScore=parsedCompleteness({...a,signature:aSig}),bScore=parsedCompleteness({...b,signature:bSig});
  if(bSig && (bLines>aLines || bScore>aScore+0.5))a.signature=bSig; else a.signature=aSig;
  a.personalNotes=a.signature||"";
  a._debug={source:sourceLabel||"merged",htmlSignature:primary?.signature||"",textSignature:secondary?.signature||"",htmlScore:aScore,textScore:bScore};
  return a;
}
function diagnosticRows(parsed){
  const vals=[
    ["Company",parsed.companyName],["Job title",parsed.jobTitle],["Business / Office",parsed.businessPhone],["Mobile",parsed.mobilePhone],["Fax",parsed.businessFax],
    ["Street",parsed.street],["City",parsed.city],["State",parsed.state],["ZIP",parsed.postalCode],["Country inferred",parsed.countryOrRegion]
  ];
  return vals.map(([l,v])=>`<div class="diag-row"><b>${html(l)}</b><span>${html(v||"(not detected)")}</span></div>`).join("");
}
function diagnosticPanel(parsed){
  const d=parsed._debug||{};
  return `<details class="diagnostic"><summary>Show diagnostic</summary><div class="diag-source">Source used: ${html(d.source||"email signature")}</div>${diagnosticRows(parsed)}<div class="diag-block"><b>Signature text used for Notes</b><pre>${html(parsed.signature||"(none)")}</pre></div>${d.htmlSignature&&d.textSignature&&d.htmlSignature!==d.textSignature?`<div class="diag-block"><b>HTML signature candidate</b><pre>${html(d.htmlSignature)}</pre><b>Plain-text signature candidate</b><pre>${html(d.textSignature)}</pre></div>`:""}</details>`;
}



// v4.0.0 — stricter image-signature OCR classification.
// Normal text/HTML parsing still runs first. OCR is invoked only when no usable
// contact candidate was found, which keeps ordinary emails fast.
let __tesseractPromise=null;
function loadTesseract(){
  if(window.Tesseract)return Promise.resolve(window.Tesseract);
  if(__tesseractPromise)return __tesseractPromise;
  __tesseractPromise=new Promise((resolve,reject)=>{
    const sc=document.createElement("script");
    sc.src="https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js";
    sc.async=true;
    sc.onload=()=>window.Tesseract?resolve(window.Tesseract):reject(new Error("OCR library loaded but was unavailable."));
    sc.onerror=()=>reject(new Error("Unable to load the OCR library."));
    document.head.appendChild(sc);
  });
  return __tesseractPromise;
}
function imageMimeFromName(name){
  const n=String(name||"").toLowerCase();
  if(n.endsWith(".png"))return "image/png";
  if(n.endsWith(".gif"))return "image/gif";
  if(n.endsWith(".webp"))return "image/webp";
  if(n.endsWith(".bmp"))return "image/bmp";
  return "image/jpeg";
}
function attachmentImageDataUrl(item,att){
  return new Promise(resolve=>{
    try{
      item.getAttachmentContentAsync(att.id,r=>{
        if(r.status!==Office.AsyncResultStatus.Succeeded)return resolve("");
        const c=r.value||{},content=c.content||"";
        if(!content)return resolve("");
        if(/^data:image\//i.test(content))return resolve(content);
        // File attachments/inline images are normally returned as base64.
        if(String(c.format||"").toLowerCase().includes("base64") || /^[A-Za-z0-9+/=\r\n]+$/.test(content.slice(0,200))){
          return resolve(`data:${imageMimeFromName(att.name)};base64,${content.replace(/\s+/g,"")}`);
        }
        resolve("");
      });
    }catch(_){resolve("")}
  });
}
function htmlImageHints(htmlText){
  const out=[];
  try{
    const doc=new DOMParser().parseFromString(htmlText||"","text/html");
    for(const img of [...doc.querySelectorAll("img")]){
      const src=(img.getAttribute("src")||"").trim();
      if(!src)continue;
      const w=parseInt(img.getAttribute("width")||img.style?.width||"0",10)||0;
      const h=parseInt(img.getAttribute("height")||img.style?.height||"0",10)||0;
      // Signature cards are usually wide; keep unknown-size images too because Outlook
      // frequently strips width/height attributes from inline images.
      const likely=(w===0||h===0||(w>=220&&h>=55)||(w/h>=2.2&&w>=180));
      if(!likely)continue;
      if(/^data:image\//i.test(src) || /^https?:\/\//i.test(src))out.push(src);
    }
  }catch(_){ }
  return [...new Set(out)];
}
async function imageSignatureSources(item,htmlText){
  const sources=htmlImageHints(htmlText);
  const atts=Array.from(item.attachments||[]).filter(a=>{
    const n=String(a.name||"").toLowerCase();
    return !!a.isInline || /\.(png|jpe?g|gif|webp|bmp)$/i.test(n);
  }).slice(0,8);
  for(const a of atts){const d=await attachmentImageDataUrl(item,a);if(d)sources.push(d)}
  return [...new Set(sources)].slice(0,8);
}
function smartTitleCase(v){
  const x=String(v||"").trim();
  if(!x)return"";
  if(x===x.toUpperCase())return x.toLowerCase().replace(/\b[a-z]/g,c=>c.toUpperCase());
  return x;
}
function ocrJobTitle(text,currentName){
  const lines=norm(text).split("\n").map(x=>x.trim()).filter(Boolean);
  const target=cleanName(currentName||"").toLowerCase();
  let start=0;
  if(target){const i=lines.findIndex(l=>cleanName(l).toLowerCase()===target||cleanName(l).toLowerCase().startsWith(target+" "));if(i>=0)start=i+1;}
  for(let i=start;i<Math.min(lines.length,start+5);i++){
    const line=lines[i];
    if(cleanEmail(line)||/\d{3}[\s.\-]*\d{3}/.test(line)||/\b\d{5}(?:-\d{4})?\b/.test(line)||/^www\.|https?:/i.test(line))break;
    const low=line.toLowerCase();
    if(TITLE_WORDS.some(t=>low.includes(t))&&!COMPANY_WORDS.some(w=>(" "+low).includes(w)))return smartTitleCase(line.replace(/^[-|•·—–\s]+|[-|•·—–\s]+$/g,""));
  }
  return"";
}

function strictOcrPhones(text){
  const raw=norm(text).replace(/©/g,"O").replace(/®/g,"O");
  const out={businessPhone:"",mobilePhone:"",businessFax:""};
  function fmt(num,ext){return phoneForOutlook(num+(ext?` x${ext}`:""));}

  const dt=raw.match(/(?:^|[\n|•;\s])(?:direct\s*[/&+\-]\s*text|direct\s+text|text\s*[/&+\-]\s*direct)\s*[:.\-]?\s*((?:\+?1[\s.\-]?)?(?:\(?\d{3}\)?[\s.\-]*)\d{3}[\s.\-]*\d{4})(?:\s*(?:x|ext\.?|extension)\s*[:.\-]?\s*(\d+))?/im);
  if(dt)out.mobilePhone=fmt(dt[1],dt[2]);

  const office=raw.match(/(?:^|[\n|•;\s])(?:office|business|phone|tel|telephone|O|0)\s*[:.\-]?\s*((?:\+?1[\s.\-]?)?(?:\(?\d{3}\)?[\s.\-]*)\d{3}[\s.\-]*\d{4})(?:\s*(?:x|ext\.?|extension)\s*[:.\-]?\s*(\d+))?/im);
  if(office)out.businessPhone=fmt(office[1],office[2]);

  const direct=raw.match(/(?:^|[\n|•;\s])(?:direct|D)\s*[:.\-]?\s*((?:\+?1[\s.\-]?)?(?:\(?\d{3}\)?[\s.\-]*)\d{3}[\s.\-]*\d{4})(?:\s*(?:x|ext\.?|extension)\s*[:.\-]?\s*(\d+))?/im);
  if(direct){
    const val=fmt(direct[1],direct[2]);
    if(direct[2])out.businessPhone=val;
    else if(!out.mobilePhone)out.mobilePhone=val;
  }

  if(!out.mobilePhone){
    const mobile=raw.match(/(?:^|[\n|•;\s])(?:mobile|cell|cellular|M|C)\s*[:.\-]?\s*((?:\+?1[\s.\-]?)?(?:\(?\d{3}\)?[\s.\-]*)\d{3}[\s.\-]*\d{4})(?:\s*(?:x|ext\.?|extension)\s*[:.\-]?\s*(\d+))?/im);
    if(mobile)out.mobilePhone=fmt(mobile[1],mobile[2]);
  }

  const fax=raw.match(/(?:^|[\n|•;\s])(?:fax|F)\s*[:.\-]?\s*((?:\+?1[\s.\-]?)?(?:\(?\d{3}\)?[\s.\-]*)\d{3}[\s.\-]*\d{4})(?:\s*(?:x|ext\.?|extension)\s*[:.\-]?\s*(\d+))?/im);
  if(fax)out.businessFax=fmt(fax[1],fax[2]);

  if(!out.businessPhone && !out.mobilePhone){
    const matches=[...raw.matchAll(/((?:\+?1[\s.\-]?)?(?:\(?\d{3}\)?[\s.\-]*)\d{3}[\s.\-]*\d{4})(?:\s*(?:x|ext\.?|extension)\s*[:.\-]?\s*(\d+))?/gi)];
    if(matches.length===1)out.businessPhone=fmt(matches[0][1],matches[0][2]);
  }

  if(out.businessPhone===out.mobilePhone)out.mobilePhone="";
  if(out.businessPhone===out.businessFax)out.businessFax="";
  return out;
}
function ocrCompanyFromText(text,currentEmail,site){
  const lines=norm(text).split("\n").map(x=>x.replace(/\s+/g," ").trim()).filter(Boolean);
  for(let i=0;i<lines.length;i++){
    const line=lines[i],low=" "+line.toLowerCase();
    if(cleanEmail(line)||/^www\.|https?:/i.test(line)||/\d{3}[\s.\-]*\d{3}/.test(line))continue;
    if(COMPANY_WORDS.some(w=>low.includes(w))){
      if(/^(concrete|construction|design|engineering|electric|electrical|builders|architects|architecture)$/i.test(line) && i>0){
        const prev=lines[i-1];
        if(prev && prev.length<35 && !cleanEmail(prev) && !/\d/.test(prev) && !TITLE_WORDS.some(t=>prev.toLowerCase().includes(t))){
          return smartTitleCase(prev+" "+line);
        }
      }
      return smartTitleCase(line);
    }
  }
  return companyFromDomain(currentEmail,site);
}
function sourceToDataUrl(src){
  return new Promise((resolve)=>{
    if(/^data:image\//i.test(src||""))return resolve(src);
    try{
      const im=new Image();im.crossOrigin="anonymous";
      im.onload=()=>{
        try{
          const c=document.createElement("canvas");c.width=im.naturalWidth;c.height=im.naturalHeight;
          c.getContext("2d").drawImage(im,0,0);resolve(c.toDataURL("image/png"));
        }catch(_){resolve("")}
      };
      im.onerror=()=>resolve("");
      im.src=src;
    }catch(_){resolve("")}
  });
}
function cropSignatureTextSide(src){
  return new Promise(async resolve=>{
    try{
      const data=await sourceToDataUrl(src);if(!data)return resolve("");
      const im=new Image();
      im.onload=()=>{
        try{
          if(!im.naturalWidth||!im.naturalHeight||im.naturalWidth/im.naturalHeight<2.2)return resolve("");
          const c=document.createElement("canvas");
          c.width=Math.max(1,Math.floor(im.naturalWidth*0.58));c.height=im.naturalHeight;
          c.getContext("2d").drawImage(im,0,0,c.width,im.naturalHeight,0,0,c.width,im.naturalHeight);
          resolve(c.toDataURL("image/png"));
        }catch(_){resolve("")}
      };
      im.onerror=()=>resolve("");
      im.src=data;
    }catch(_){resolve("")}
  });
}

function repairOcrStateAndCountry(parsed,text){
  const out={...parsed};
  let st=String(out.state||"").toUpperCase().trim();
  const zip=String(out.postalCode||"").trim() || ((norm(text).match(/\b\d{5}(?:-\d{4})?\b/)||[])[0]||"");
  if(!out.postalCode&&zip)out.postalCode=zip;
  if(!US_STATES.has(st) && zip){
    const z=parseInt(zip.slice(0,3),10);
    // Common OCR confusion for ID -> 10/1D/IO. Idaho ZIP prefixes are 832-838.
    if(z>=832&&z<=838)st="ID";
    else {
      const near=norm(text).match(new RegExp(`[,\\s]+([A-Z0-9]{2})\\s+${zip.replace(/[-]/g,"\\-")}`,'i'));
      if(near){const tok=near[1].toUpperCase().replace(/^1D$/,"ID").replace(/^I0$/,"ID").replace(/^10$/,"ID");if(US_STATES.has(tok))st=tok;}
    }
  }
  if(st&&US_STATES.has(st))out.state=st;
  // A recognized US state OR a normal five-digit ZIP is enough to infer country.
  if(!out.countryOrRegion && (US_STATES.has(String(out.state||"").toUpperCase()) || /^\d{5}(?:-\d{4})?$/.test(zip)))out.countryOrRegion="United States";
  return out;
}
function repairOcrParsed(parsed,text,currentName,currentEmail){
  let out={...parsed};
  const exactTitle=ocrJobTitle(text,currentName);
  if(exactTitle)out.jobTitle=exactTitle;
  const strictPhones=strictOcrPhones(text);
  out.businessPhone=strictPhones.businessPhone;
  out.mobilePhone=strictPhones.mobilePhone;
  out.businessFax=strictPhones.businessFax;
  out=repairOcrStateAndCountry(out,text);
  const site=website(text,currentEmail)||out.businessHomePage||"";
  out.businessHomePage=site;
  const co=ocrCompanyFromText(text,currentEmail,site);
  if(co)out.companyName=co;
  // Outlook's From header remains authoritative for the sender identity/email.
  const np=nameParts(currentName||[out.givenName,out.surname].filter(Boolean).join(" "));
  if(currentName)Object.assign(out,np);
  if(currentEmail)out.email=String(currentEmail).toLowerCase();
  return out;
}

function ocrSignatureScore(text,currentName,currentEmail){
  const t=cleanSignatureText(text||"");if(!t)return -999;
  let score=t.split("\n").reduce((n,l)=>n+signatureScore(l),0);
  const nm=cleanName(currentName||"").toLowerCase();
  if(nm&&t.toLowerCase().includes(nm))score+=8;
  if(currentEmail&&t.toLowerCase().includes(String(currentEmail).toLowerCase()))score+=7;
  const ph=phones(t),ad=address(t);
  if(ph.businessPhone||ph.mobilePhone)score+=5;
  if(ad.state&&ad.postalCode)score+=6;
  if(website(t,currentEmail))score+=3;
  if(t.split("\n").filter(Boolean).length>=4)score+=2;
  return score;
}
async function imageSignatureCandidate(item,htmlText,currentName,currentEmail,myEmail){
  if(!currentEmail || sameEmail(currentEmail,myEmail))return null;
  const sources=await imageSignatureSources(item,htmlText);
  if(!sources.length)return null;
  status(`No text signature found. Reading ${sources.length} signature image${sources.length===1?"":"s"}…`);
  const T=await loadTesseract();
  let worker=null,best=null;
  try{
    worker=await T.createWorker("eng");
    try{await worker.setParameters({tessedit_pageseg_mode:"6"});}catch(_){}
    let jobs=[];
    for(const src of sources){
      jobs.push({src,kind:"full"});
      const crop=await cropSignatureTextSide(src);
      if(crop)jobs.push({src:crop,kind:"text-side crop"});
    }
    for(let i=0;i<jobs.length;i++){
      status(`Reading signature image ${Math.floor(i/2)+1}…`);
      try{
        const r=await worker.recognize(jobs[i].src);
        const text=cleanSignatureText(r?.data?.text||"");
        let score=ocrSignatureScore(text,currentName,currentEmail);
        if(jobs[i].kind==="text-side crop")score+=3;
        if(!best||score>best.score)best={score,text,kind:jobs[i].kind};
      }catch(_){ }
    }
  }finally{try{if(worker)await worker.terminate()}catch(_){ }}
  if(!best || best.score<8 || !best.text)return null;
  let parsed=parseContactFromSignature(currentName||"Current sender",String(currentEmail||"").toLowerCase(),best.text);
  parsed=repairOcrParsed(parsed,best.text,currentName,currentEmail);
  parsed.signature=best.text;
  parsed.personalNotes=best.text;
  parsed._debug={source:`Image OCR fallback — ${best.kind}`,htmlSignature:"",textSignature:best.text,ocrScore:best.score};
  return {name:currentName||[parsed.givenName,parsed.surname].filter(Boolean).join(" "),email:parsed.email,text:best.text,parsed};
}

function readBody(item){
  return new Promise((resolve,reject)=>{
    let html="",plain="",htmlDone=false,textDone=false,htmlErr=null,textErr=null;
    const finish=()=>{
      if(!htmlDone||!textDone)return;
      if(!plain.trim()&&html.trim())plain=htmlBodyToText(html);
      if(!plain.trim()&&!html.trim())return reject(new Error(textErr?.message||htmlErr?.message||"Unable to read the email body."));
      resolve({html,plainText:plain,text:plain});
    };
    item.body.getAsync(Office.CoercionType.Html,r=>{
      htmlDone=true;
      if(r.status===Office.AsyncResultStatus.Succeeded)html=r.value||"";
      else htmlErr=r.error;
      finish();
    });
    item.body.getAsync(Office.CoercionType.Text,r=>{
      textDone=true;
      if(r.status===Office.AsyncResultStatus.Succeeded)plain=r.value||"";
      else textErr=r.error;
      finish();
    });
  });
}

function mergeCandidatePair(h,t){
  if(h&&t){
    if(t.headerAuthoritative||t.parsed?._debug?.headerAuthoritative){
      const parsed=mergeParsed(t.parsed,h.parsed,"Forwarded header/plain text authoritative; HTML fills blanks");
      parsed.givenName=t.parsed.givenName;
      parsed.middleName=t.parsed.middleName;
      parsed.surname=t.parsed.surname;
      parsed.email=t.parsed.email;
      parsed.jobTitle=t.parsed.jobTitle||parsed.jobTitle;
      parsed.signature=t.parsed.signature;
      parsed.personalNotes=t.parsed.personalNotes;
      parsed._debug={...(parsed._debug||{}),source:"Forwarded From header authoritative",headerAuthoritative:true,headerName:t.name};
      return {...t,parsed,text:parsed.signature,headerAuthoritative:true};
    }
    const parsed=mergeParsed(h.parsed,t.parsed,"HTML + plain-text merged");
    return {...h,parsed,text:parsed.signature};
  }
  if(h){
    h.parsed._debug={source:"HTML only",htmlSignature:h.parsed.signature||"",textSignature:""};
    return h;
  }
  if(t){
    t.parsed._debug={...(t.parsed._debug||{}),source:t.headerAuthoritative?"Forwarded From header authoritative":"Plain text only",htmlSignature:"",textSignature:t.parsed.signature||""};
    return t;
  }
  return null;
}


function chooseOccurrenceCandidates(textCandidates,htmlTextCandidates){
  const byEmail=arr=>{
    const m=new Map();
    for(const c of arr||[]){const k=cleanEmail(c.email);if(!k)continue;if(!m.has(k))m.set(k,[]);m.get(k).push(c)}
    return m;
  };
  const t=byEmail(textCandidates),h=byEmail(htmlTextCandidates),out=[];
  const keys=new Set([...t.keys(),...h.keys()]);
  for(const k of keys){
    const ta=t.get(k)||[],ha=h.get(k)||[];
    // Prefer the source that found more distinct bounded occurrences. On a tie prefer Outlook Text.
    const chosen=ha.length>ta.length?ha:ta;
    out.push(...chosen);
  }
  return out;
}



function lineHasPhone(line,phone){
  const target=phoneForOutlook(phone||"").replace(/\D/g,"").slice(-10);
  if(!target)return false;
  return String(line||"").replace(/\D/g,"").includes(target);
}

function looseSignatureOccurrences(text,base){
  const p=base?.parsed||base||{};
  const name=base?.name||[p.givenName,p.middleName,p.surname].filter(Boolean).join(" ");
  const np=nameParts(name||"");
  const surname=String(np.surname||p.surname||"").trim();
  const email=cleanEmail(base?.email||p.email||"");
  if(!surname)return [];
  const lines=norm(text||"").split("\n").map(x=>String(x||"").replace(/[\u00ad\u200b-\u200d\ufeff]/g,"").trim());
  const esc=surname.replace(/[.*+?^${}()|[\]\\]/g,"\\$&");
  const surnameRx=new RegExp(`(?:^|\\s)${esc}(?:$|\\s|[,|•·—–-])`,`i`);
  const isHeader=l=>/^\s*(from|sent|to|cc|bcc|subject):\s*/i.test(l)||/^[-_]{5,}$/.test(l);
  const out=[];
  for(let i=0;i<lines.length;i++){
    const line=lines[i];
    if(!line||isHeader(line)||isClosingPhrase(line)||!surnameRx.test(line))continue;
    const max=Math.min(lines.length-1,i+14);
    let score=0,emailAt=-1,lastEvidence=i;
    let sawBusiness=false,sawMobile=false,sawZip=false,sawStreet=false,sawTitle=false;
    for(let j=i;j<=max;j++){
      const v=lines[j];
      if(j>i&&isHeader(v))break;
      if(j>i&&surnameRx.test(v))break;
      if(email&&String(v).toLowerCase().includes(email)){score+=4;emailAt=j;lastEvidence=j}
      if(!sawBusiness&&p.businessPhone&&lineHasPhone(v,p.businessPhone)){score+=2;sawBusiness=true;lastEvidence=j}
      if(!sawMobile&&p.mobilePhone&&lineHasPhone(v,p.mobilePhone)){score+=2;sawMobile=true;lastEvidence=j}
      if(!sawZip&&p.postalCode&&String(v).includes(String(p.postalCode))){score+=1;sawZip=true;lastEvidence=j}
      if(!sawStreet&&p.street){
        const first=String(p.street).toLowerCase().split(/\s+/).filter(Boolean).slice(0,2).join(" ");
        if(first&&String(v).toLowerCase().includes(first)){score+=1;sawStreet=true;lastEvidence=j}
      }
      if(!sawTitle&&p.jobTitle){
        const words=String(p.jobTitle).toLowerCase().split(/[^a-z0-9]+/).filter(w=>w.length>3);
        const hits=words.filter(w=>String(v).toLowerCase().includes(w)).length;
        if(hits>=Math.min(2,words.length)){score+=2;sawTitle=true;lastEvidence=j}
      }
    }
    if(score<5)continue;
    let start=i;
    if(i>0&&/^[A-Za-z][A-Za-z'.-]*$/.test(lines[i-1])&&!isClosingPhrase(lines[i-1]))start=i-1;
    let end=Math.max(lastEvidence,emailAt>=0?emailAt:i);
    for(let j=end+1;j<=Math.min(lines.length-1,end+2);j++){
      const v=lines[j];
      if(!v||isHeader(v)||surnameRx.test(v))break;
      if(email&&String(v).toLowerCase().includes(email)){end=j;continue}
      if(/^https?:|^www\./i.test(v)){end=j;continue}
      break;
    }
    const sig=cleanSignatureText(lines.slice(start,end+1).filter(Boolean).join("\n"));
    if(!sig)continue;
    if(out.length&&start<=out[out.length-1].end)continue;
    out.push({sig,start,end,occurrence:out.length,score:200+score});
    i=end;
  }
  return out;
}

function canonicalNotesFromParsed(p,identityName=""){
  const lines=[];
  const name=cleanName(identityName||[p.givenName,p.middleName,p.surname].filter(Boolean).join(" "));
  if(name)lines.push(name);
  if(p.jobTitle)lines.push(p.jobTitle);
  if(p.companyName)lines.push(p.companyName);
  if(p.street)lines.push(p.street);
  const cityLine=[p.city,[p.state,p.postalCode].filter(Boolean).join(" ")].filter(Boolean).join(", ");
  if(cityLine)lines.push(cityLine);
  if(p.businessPhone)lines.push(`Office: ${phoneForOutlook(p.businessPhone)}`);
  if(p.mobilePhone)lines.push(`Mobile: ${phoneForOutlook(p.mobilePhone)}`);
  if(p.businessFax)lines.push(`Fax: ${phoneForOutlook(p.businessFax)}`);
  if(p.email)lines.push(cleanEmail(p.email));
  if(p.businessHomePage)lines.push(p.businessHomePage);
  return lines.filter(Boolean).join("\n");
}

function safeNotesSignature(raw,p,identityName=""){
  const cleaned=cleanSignatureText(raw||"");
  const lines=norm(cleaned).split("\n").map(x=>x.trim()).filter(Boolean);
  const bodyNoise=/\b(?:good afternoon|good morning|good evening|please see attached|payment options?|remitted|wire transfer|customer id|billing zip|please reach out|attached invoices?)\b/i;
  const hasHeader=/^(?:from|sent|to|cc|bcc|subject):/im.test(cleaned);
  const plausible=cleaned&&lines.length>=2&&lines.length<=12&&!bodyNoise.test(cleaned)&&!hasHeader;
  if(plausible)return cleaned;
  return canonicalNotesFromParsed(p,identityName);
}

function enforceDuplicateMultiplicity(input,plainText,htmlVisible){
  const groups=new Map();
  for(const c of input||[]){
    const key=cleanEmail(c.email||c.parsed?.email)||`name:${String(c.name||"").toLowerCase()}`;
    if(!groups.has(key))groups.set(key,[]);
    groups.get(key).push(c);
  }
  const out=[];
  for(const group of groups.values()){
    const base=group[0];
    const textLoose=looseSignatureOccurrences(plainText||"",base);
    const htmlLoose=looseSignatureOccurrences(htmlVisible||"",base);
    const occs=htmlLoose.length>textLoose.length?htmlLoose:textLoose;
    const wanted=Math.max(group.length,occs.length);
    if(wanted<=group.length){out.push(...group);continue}
    out.push(...group);
    for(let i=group.length;i<wanted;i++){
      const occ=occs[i]||occs[0];
      const clone=JSON.parse(JSON.stringify(base));
      clone.occurrence=i;
      clone.occurrenceKey=`fingerprint:${cleanEmail(base.email||base.parsed?.email)}:${i}:${occ?.start??i}:${occ?.end??i}`;
      clone.duplicateTotal=wanted;
      if(occ?.sig){clone.text=occ.sig;clone.parsed.signature=occ.sig;clone.parsed.personalNotes=occ.sig}
      clone.parsed._debug={...(clone.parsed._debug||{}),source:"Fingerprint duplicate occurrence fallback",occurrence:i};
      out.push(clone);
    }
  }
  return out;
}

function candidateDefaultChecked(total,index){
  // Unique contacts stay selected as before. Duplicate signature occurrences require an explicit user choice.
  return total<=1;
}

function rebuildCandidateFromOccurrence(base,occ,index,total){
  const name=base.name||[base.parsed?.givenName,base.parsed?.middleName,base.parsed?.surname].filter(Boolean).join(" ");
  const email=cleanEmail(base.email||base.parsed?.email||"");
  const parsed=parseContactFromSignature(name,email,occ.sig);
  if(looksLikePersonName(name)){
    const np=nameParts(name);
    parsed.givenName=np.givenName;parsed.middleName=np.middleName;parsed.surname=np.surname;
  }
  parsed.email=email;
  parsed.jobTitle=sanitizeJobTitle(parsed.jobTitle,name,occ.sig);
  parsed.signature=occ.sig;
  parsed.personalNotes=occ.sig;
  parsed._debug={source:"Final bounded signature occurrence",headerAuthoritative:!!base.headerAuthoritative,headerName:name,occurrence:index,textSignature:occ.sig,htmlSignature:""};
  return {...base,name,email,text:occ.sig,parsed,occurrence:index,occurrenceKey:`final:${email}:${index}:${occ.start??index}:${occ.end??index}`,duplicateTotal:total};
}

function expandCandidatesByBodyOccurrences(input,plainText,htmlVisible){
  const groups=new Map();
  for(const c of input||[]){
    const key=cleanEmail(c.email||c.parsed?.email)||`name:${String(c.name||"").toLowerCase()}`;
    if(!groups.has(key))groups.set(key,[]);
    groups.get(key).push(c);
  }
  const out=[];
  for(const group of groups.values()){
    const base=group[0];
    const name=base.name||[base.parsed?.givenName,base.parsed?.middleName,base.parsed?.surname].filter(Boolean).join(" ");
    const email=cleanEmail(base.email||base.parsed?.email||"");
    let textOcc=email&&name?identitySignatureOccurrences(plainText||"",name,email):[];
    let htmlOcc=email&&name?identitySignatureOccurrences(htmlVisible||"",name,email):[];
    let occs=htmlOcc.length>textOcc.length?htmlOcc:textOcc;

    // If an earlier parser already found more distinct occurrences, preserve those rather than collapse them.
    if(group.length>occs.length){
      const bounded=group.filter(x=>{
        const s=String(x.parsed?.signature||"");
        const lines=norm(s).split("\n").filter(Boolean).length;
        return s&&lines<=14&&!/^(?:from|sent|to|cc|bcc|subject):/im.test(s)&&!/\b(?:payment options?|please see attached|please reach out|wire transfer|customer id|billing zip)\b/i.test(s);
      });
      if(bounded.length===group.length&&bounded.length>occs.length){out.push(...bounded);continue}
    }

    if(occs.length){
      for(let i=0;i<occs.length;i++)out.push(rebuildCandidateFromOccurrence(base,occs[i],i,occs.length));
    }else{
      out.push(...group);
    }
  }
  return out;
}

function finalizeCandidate(c){
  if(!c||!c.parsed)return c;
  const p=c.parsed;
  const identity=c.name||[p.givenName,p.middleName,p.surname].filter(Boolean).join(" ");
  p.jobTitle=sanitizeJobTitle(p.jobTitle,identity,p.signature||"");
  let tight=tightNotesSignature(p.signature||c.text||"",identity,p.email||c.email||"");
  tight=safeNotesSignature(tight,p,identity);
  p.signature=tight;
  p.personalNotes=tight;
  c.text=tight;
  return c;
}



function recoverCurrentSenderNameCompanyCandidates(bodyText,currentName,currentEmail,myEmail){
  const name=cleanName(currentName||"");
  const email=cleanEmail(currentEmail||"");
  if(!bodyText||!name||!email||sameEmail(email,myEmail)||!looksLikePersonName(name))return [];

  const segments=allMessageSegments(bodyText,name,email);
  const current=segments[0];
  if(!current?.text)return [];

  const lines=norm(current.text).split("\n");
  const cleanLine=v=>String(v||"")
    .replace(/[\u00ad\u200b-\u200d\ufeff]/g,"")
    .replace(/\u00a0/g," ")
    .replace(/\s+/g," ")
    .trim();

  const nameRx=new RegExp(
    "^"+name.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")+"\\s*(?:\\||•|·|—|–)\\s*(.+)$",
    "i"
  );
  const isHeader=l=>/^\s*(from|sent|to|cc|bcc|subject):\s*/i.test(l);
  const out=[];

  for(let i=0;i<lines.length;i++){
    const first=cleanLine(lines[i]);
    if(!first)continue;

    const m=first.match(nameRx);
    if(!m)continue;

    const companyPart=String(m[1]||"").trim();
    const companyLow=" "+companyPart.toLowerCase();

    // This fallback is specifically for "Name | Company", not "Name | Title".
    // Require clear company evidence to avoid reclassifying ordinary delimited text.
    if(!companyPart||!COMPANY_WORDS.some(w=>companyLow.includes(w)))continue;

    const raw=[first];
    let end=i;
    let nonblank=1;

    for(let j=i+1;j<Math.min(lines.length,i+25)&&nonblank<12;j++){
      const v=cleanLine(lines[j]);
      if(!v)continue;
      if(isHeader(v))break;
      if(/^[-_=]{5,}$/.test(v))continue;

      // A second matching signature start is a new physical occurrence.
      if(j>i&&nameRx.test(v))break;

      raw.push(v);
      end=j;
      nonblank++;

      // Stop once we have enough complete contact evidence and the next
      // meaningful content would be outside this compact signature.
      if(raw.length>=4){
        const sigNow=cleanSignatureText(raw.join("\n"));
        const ph=phones(sigNow);
        const ad=address(sigNow);
        if((ph.businessPhone||ph.mobilePhone) && ad.state && ad.postalCode){
          // Keep looking only until a message header; this allows a website
          // after the phone/address without leaking into quoted history.
          continue;
        }
      }
    }

    const sig=cleanSignatureText(raw.join("\n"));
    if(!sig||signatureEvidenceScore(raw)<5)continue;

    const parsed=parseContactFromSignature(name,email,sig);
    parsed.companyName=companyPart;
    parsed.signature=sig;
    parsed.personalNotes=sig;
    parsed._debug={...(parsed._debug||{}),source:"Current sender name + company recovery"};

    if(!parsed.givenName||!parsed.surname)continue;
    if(!parsed.businessPhone&&!parsed.mobilePhone&&!parsed.street&&!parsed.businessHomePage)continue;

    out.push({
      score:100,
      name,
      email,
      text:sig,
      parsed,
      candidateType:"signature",
      physicalOccurrence:true,
      occurrenceKey:`current-name-company:${current.start||0}:${i}:${end}`,
      globalStart:(current.start||0)+i,
      globalEnd:(current.start||0)+end
    });
  }

  return out;
}


function recoverCurrentSenderSignatureCandidates(bodyText,currentName,currentEmail,myEmail){
  const email=cleanEmail(currentEmail||"");
  if(!bodyText||!email||sameEmail(email,myEmail))return [];

  const segments=allMessageSegments(bodyText,currentName||"",email);
  const current=segments[0];
  if(!current?.text)return [];

  const lines=norm(current.text).split("\n");
  const cleanLine=v=>String(v||"")
    .replace(/[\u00ad\u200b-\u200d\ufeff]/g,"")
    .replace(/\u00a0/g," ")
    .replace(/\s+/g," ")
    .trim();
  const isMessageHeader=l=>/^\s*(from|sent|to|cc|bcc|subject):\s*/i.test(l);

  const out=[];

  for(let i=0;i<lines.length;i++){
    const v=cleanLine(lines[i]);
    if(!v||cleanEmail(v)!==email)continue;

    // Find a human signer above the current sender email. A graphic/divider
    // line may sit between the company name and phone/email, so do not treat
    // underscores/dashes as an Outlook message boundary here.
    let nameIdx=-1;
    for(let j=i-1;j>=Math.max(0,i-24);j--){
      const b=cleanLine(lines[j]);
      if(!b)continue;
      if(isMessageHeader(b))break;
      if(/^[-_=]{5,}$/.test(b))continue;
      if(isClosingPhrase(b))continue;
      if(looksLikePersonName(b)){
        nameIdx=j;
        break;
      }
    }
    if(nameIdx<0)continue;

    let end=i;
    let nonblankAfterEmail=0;
    for(let j=i+1;j<Math.min(lines.length,i+40);j++){
      const f=cleanLine(lines[j]);
      if(!f)continue;
      if(isMessageHeader(f))break;
      if(/^[-_=]{5,}$/.test(f))continue;
      if(cleanEmail(f)&&cleanEmail(f)!==email)break;

      end=j;
      nonblankAfterEmail++;

      const candidateText=lines.slice(nameIdx,end+1)
        .map(cleanLine)
        .filter(x=>x&&!/^[-_=]{5,}$/.test(x))
        .join("\n");
      const ad=address(candidateText);
      if(ad.street&&ad.state&&ad.postalCode)break;

      if(nonblankAfterEmail>=10&&f.length>100&&/[.!?]$/.test(f))break;
    }

    const raw=lines.slice(nameIdx,end+1)
      .map(cleanLine)
      .filter(x=>x&&!/^[-_=]{5,}$/.test(x));
    const sig=cleanSignatureText(raw.join("\n"));
    if(!sig)continue;

    const parsed=parseContactFromSignature("",email,sig);
    if(!parsed.givenName||!parsed.surname)continue;
    if(!parsed.businessPhone&&!parsed.mobilePhone&&!parsed.businessHomePage&&!parsed.street)continue;

    parsed.personalNotes=parsed.signature||sig;
    parsed._debug={...(parsed._debug||{}),source:"Current sender signature recovery"};

    out.push({
      score:100,
      name:[parsed.givenName,parsed.middleName,parsed.surname].filter(Boolean).join(" "),
      email,
      text:parsed.signature||sig,
      parsed,
      candidateType:"signature",
      physicalOccurrence:true,
      occurrenceKey:`current-recovery:${current.start||0}:${nameIdx}:${i}`,
      globalStart:(current.start||0)+nameIdx,
      globalEnd:(current.start||0)+end
    });
  }
  return out;
}



function signatureDiscoveryText(body){
  return norm(body).split("\n").map(line=>{
    let v=String(line||"");
    v=v.replace(/\s*<\s*(?:https?:\/\/|mailto:|tel:)[^>\r\n]*>\s*/gi," ");
    v=v.replace(/\[([^\]]+)\]\((?:https?:\/\/|mailto:|tel:)[^)]+\)/gi,"$1");
    v=v.replace(/\s*\(\s*(?:https?:\/\/|mailto:|tel:)[^)\r\n]*\)\s*/gi," ");
    return v.replace(/[ \t]+$/g,"");
  }).join("\n");
}

function enrichMissingExplicitAddressFromBody(candidate,bodyText){
  if(!candidate?.parsed)return candidate;
  const p=candidate.parsed;
  if(p.street||p.city||p.state||p.postalCode||p.countryOrRegion)return candidate;

  const email=cleanEmail(candidate.email||p.email||"");
  if(!email||!bodyText)return candidate;

  const lines=norm(bodyText).split("\n");
  const isHeader=l=>/^\s*(from|sent|to|cc|bcc|subject):\s*/i.test(l)||/^[-_]{5,}$/.test(l);
  const matches=[];

  const cleanLine=v=>String(v||"")
    .replace(/[\u00ad\u200b-\u200d\ufeff]/g,"")
    .replace(/\u00a0/g," ")
    .replace(/\s+/g," ")
    .trim();

  const companyKey=String(p.companyName||"").toLowerCase()
    .replace(/[^a-z0-9]+/g," ")
    .split(/\s+/).filter(w=>w.length>3).slice(0,3);

  const identity=candidate.name||[p.givenName,p.middleName,p.surname].filter(Boolean).join(" ");
  const surname=String(nameParts(identity||"").surname||"").toLowerCase();

  for(let i=0;i<lines.length;i++){
    const line=cleanLine(lines[i]);
    if(!line)continue;
    const found=cleanEmail(line);
    if(!found||found!==email)continue;

    const tail=[];
    let nonblank=0;
    for(let j=i+1;j<lines.length&&j<=i+45&&nonblank<10;j++){
      const v=cleanLine(lines[j]);
      if(!v)continue;
      if(isHeader(v))break;
      if(cleanEmail(v)&&cleanEmail(v)!==email)break;
      nonblank++;
      tail.push(v);

      const combined=cleanSignatureText(
        [p.signature||candidate.text||"",...tail].filter(Boolean).join("\n")
      );
      const ad=address(combined);
      if(ad.street&&ad.state&&ad.postalCode){
        let score=0;
        const before=[];
        for(let k=i-1;k>=0&&k>=i-30&&before.length<12;k--){
          const b=cleanLine(lines[k]);
          if(!b)continue;
          if(isHeader(b))break;
          before.unshift(b);
        }
        const nearby=before.join(" ").toLowerCase();
        if(p.businessPhone&&lineHasPhone(nearby,p.businessPhone))score+=4;
        if(p.mobilePhone&&lineHasPhone(nearby,p.mobilePhone))score+=3;
        if(surname&&nearby.includes(surname))score+=2;
        if(companyKey.length&&companyKey.some(w=>nearby.includes(w)))score+=2;

        matches.push({
          score,
          ad,
          signature:combined,
          addressKey:[ad.street,ad.city,ad.state,ad.postalCode].join("|").toLowerCase()
        });
        break;
      }

      // Do not wander into normal prose after a signature.
      if(v.length>120&&/[.!?]$/.test(v))break;
    }
  }

  if(!matches.length)return candidate;
  matches.sort((a,b)=>b.score-a.score);
  const topScore=matches[0].score;
  const top=matches.filter(x=>x.score===topScore);
  const uniqueAddresses=new Set(top.map(x=>x.addressKey));

  // If equally plausible occurrences point to different addresses, do nothing.
  if(uniqueAddresses.size>1)return candidate;

  const best=top[0];
  p.street=best.ad.street;
  p.city=best.ad.city;
  p.state=best.ad.state;
  p.postalCode=best.ad.postalCode;
  p.countryOrRegion=best.ad.countryOrRegion||usCountry(best.ad.state);

  // The discovered address is part of the same explicit signature tail, so keep
  // Notes/preview complete as well.
  p.signature=best.signature;
  p.personalNotes=best.signature;
  candidate.text=best.signature;
  p._debug={...(p._debug||{}),addressEnrichedFromNearbyExplicitSignature:true};
  return candidate;
}

async function scan(){try{
  status("Reading the current email chain…");$("reviewSection").hidden=true;
  const item=Office.context.mailbox.item;if(!item||item.itemType!==Office.MailboxEnums.ItemType.Message)throw new Error("Open or select an email message first.");
  const from=item.from||{},body=await readBody(item);const myEmail=(Office.context.mailbox.userProfile?.emailAddress||"").toLowerCase();

  const textBody=body.text||body.plainText||"";
  const textSignatureBody=signatureDiscoveryText(textBody);
  const textSignatures=structuredDiscoverSignatures(
    textSignatureBody,from.displayName||"",from.emailAddress||"",myEmail
  ).map(c=>v4AttachOriginalSignatureBlock(textSignatureBody,c));

  const htmlVisible=body.html?htmlBodyToText(body.html):"";
  const htmlSignatureVisible=signatureDiscoveryText(htmlVisible);
  const htmlSignatures=htmlSignatureVisible
    ? structuredDiscoverSignatures(htmlSignatureVisible,from.displayName||"",from.emailAddress||"",myEmail)
        .map(c=>v4AttachOriginalSignatureBlock(htmlSignatureVisible,c))
    : [];

  // Never let one representation replace the other wholesale. Preserve the
  // Text candidate when both representations see the same physical occurrence,
  // and add only occurrences visible in HTML but missing from Text.
  let signatureCandidates=mergeSignatureCandidateRepresentations(
    textSignatureBody,textSignatures,
    htmlSignatureVisible,htmlSignatures,
    from.displayName||"",from.emailAddress||""
  );

  // Referred contacts remain an independent candidate source.
  let referralCandidates=referredContactCandidates(textBody,myEmail);
  if(!referralCandidates.length&&htmlVisible)referralCandidates=referredContactCandidates(htmlVisible,myEmail);

  if(!signatureCandidates.length){
    try{
      const ocr=await imageSignatureCandidate(item,body.html||"",from.displayName||"",from.emailAddress||"",myEmail);
      if(ocr)signatureCandidates=[ocr];
    }catch(ocrErr){console.warn("Image signature OCR fallback failed",ocrErr)}
  }

  const addressSource=textSignatureBody||htmlSignatureVisible||"";
  signatureCandidates=signatureCandidates.map(c=>enrichMissingExplicitAddressFromBody(c,addressSource));

  candidates=[...signatureCandidates,...referralCandidates];

  candidates=candidates.map(c=>{
    if(!c?.parsed)return c;
    if(c.candidateType!=="referral"){
      const identity=c.name||[c.parsed.givenName,c.parsed.middleName,c.parsed.surname].filter(Boolean).join(" ");
      const sourceBlock=c.rawSignatureBlock||c.parsed.signature||c.text||"";
      const reparsed=v4ContactFromSignature(identity,c.email||c.parsed.email||"",sourceBlock);
      reparsed._debug={...(c.parsed._debug||{}),...(reparsed._debug||{})};
      c.parsed=reparsed;
    }
    if(c.candidateType==="referral")c.parsed.personalNotes=c.parsed.signature||"";
    return c
  });

  renderCandidates();
  status(`Found ${candidates.length} contact choice${candidates.length===1?"":"s"}. Select one to process. Your own messages/signature are ignored.`,candidates.length?"ok":"error");
}catch(e){status(e.message||String(e),"error")}}
const APP_CLIENT_ID="aaaad2a2-1ca4-4d45-8b8d-21c56f2c4137";
function clientId(){return APP_CLIENT_ID}
function showAuthSetup(){$("authSetup").hidden=true;$("clientId").value=APP_CLIENT_ID}
async function initMsal(){const id=clientId();if(!id)throw new Error("Microsoft Contacts access is not configured yet. Enter the Application (client) ID first.");if(!msalInstance){msalInstance=await createNestablePublicClientApplication({auth:{clientId:id,authority:"https://login.microsoftonline.com/common"},cache:{cacheLocation:"localStorage"}})}return msalInstance}
async function accessToken(){const pca=await initMsal();const request={scopes:GRAPH_SCOPES,loginHint:Office.context.mailbox.userProfile.emailAddress};try{return (await pca.acquireTokenSilent(request)).accessToken}catch(err){if(err instanceof InteractionRequiredAuthError || /interaction|consent|login/i.test(String(err?.errorCode||err?.message||err))){return (await pca.acquireTokenPopup(request)).accessToken}throw err}}
async function graph(path,opts={}){const token=await accessToken();const r=await fetch("https://graph.microsoft.com/v1.0"+path,{...opts,headers:{Authorization:`Bearer ${token}`,"Content-Type":"application/json",...(opts.headers||{})}});if(!r.ok){const t=await r.text();throw new Error(`Microsoft Contacts error ${r.status}: ${t.slice(0,350)}`)}if(r.status===204)return null;return r.json()}
async function loadContacts(){let url="/me/contacts?$top=250&$select=id,displayName,givenName,middleName,surname,companyName,jobTitle,emailAddresses,businessPhones,mobilePhone,businessHomePage,businessAddress,homeAddress,otherAddress,personalNotes";const all=[];while(url){const data=await graph(url.replace("https://graph.microsoft.com/v1.0",""));all.push(...(data.value||[]));url=data["@odata.nextLink"]||""}return all}
function n(s){return String(s||"").trim().toLowerCase().replace(/[^a-z0-9]+/g," ").trim()}
function digits(s){return String(s||"").replace(/\D/g,"").slice(-10)}
function addressHasData(a){return !!(a&&(a.street||a.city||a.state||a.postalCode||a.countryOrRegion))}
function preferredExistingAddress(c){
  if(addressHasData(c.businessAddress))return {a:c.businessAddress,type:"Business"};
  if(addressHasData(c.homeAddress))return {a:c.homeAddress,type:"Home"};
  if(addressHasData(c.otherAddress))return {a:c.otherAddress,type:"Other"};
  return {a:{},type:""};
}
function existingFlat(c){const pref=preferredExistingAddress(c),a=pref.a||{};return{givenName:c.givenName||"",middleName:c.middleName||"",surname:c.surname||"",companyName:c.companyName||"",jobTitle:c.jobTitle||"",email:(c.emailAddresses?.[0]?.address)||"",businessPhone:(c.businessPhones?.[0])||"",mobilePhone:c.mobilePhone||"",businessFax:"",businessHomePage:c.businessHomePage||"",street:a.street||"",city:a.city||"",state:a.state||"",postalCode:a.postalCode||"",countryOrRegion:a.countryOrRegion||"",personalNotes:c.personalNotes||"",_addressType:pref.type,_allBusinessPhones:(c.businessPhones||[]).join(" | ")}}
function matchContact(p){const email=n(p.email);if(email){const m=graphContacts.find(c=>(c.emailAddresses||[]).some(e=>n(e.address)===email));if(m)return{contact:m,confidence:"Exact email"}}
  const full=n([p.givenName,p.middleName,p.surname].filter(Boolean).join(" "));const companyN=n(p.companyName);const ph=[digits(p.businessPhone),digits(p.mobilePhone)].filter(Boolean);
  const candidates2=graphContacts.filter(c=>{const f=n([c.givenName,c.middleName,c.surname].filter(Boolean).join(" "));if(!full||f!==full)return false;const cf=n(c.companyName);if(companyN&&cf&&companyN===cf)return true;const cp=[...(c.businessPhones||[]),c.mobilePhone||""].map(digits).filter(Boolean);return ph.some(x=>cp.includes(x))});
  return candidates2.length===1?{contact:candidates2[0],confidence:"Name + company/phone"}:null}
function fieldDiffs(parsed,existing){const out=[];for(const [key,label] of Object.entries(FIELD_META)){const nv=(parsed[key]||"").trim(),ov=(existing[key]||"").trim();if(!nv)continue;if(n(nv)===n(ov))continue;out.push({key,label,old:ov,new:nv,defaultChecked:!ov})}return out}
function graphPayload(p,keys=null,existing=null){const use=k=>!keys||keys.has(k);const o={};if(use("givenName"))o.givenName=p.givenName||"";if(use("middleName"))o.middleName=p.middleName||"";if(use("surname"))o.surname=p.surname||"";if(use("companyName"))o.companyName=p.companyName||"";if(use("jobTitle"))o.jobTitle=p.jobTitle||"";if(use("email")&&p.email)o.emailAddresses=[{address:p.email,name:[p.givenName,p.surname].filter(Boolean).join(" ")||p.email}];if(use("businessPhone")&&p.businessPhone)o.businessPhones=[phoneForOutlook(p.businessPhone)];if(use("mobilePhone")&&p.mobilePhone)o.mobilePhone=phoneForOutlook(p.mobilePhone);if(use("businessHomePage")&&p.businessHomePage)o.businessHomePage=p.businessHomePage;if(use("personalNotes"))o.personalNotes=p.personalNotes||"";
  const addrNames=["street","city","state","postalCode","countryOrRegion"];
  const addressKeys=addrNames.filter(use);
  if(addressKeys.length){const base=existing||{};o.businessAddress={};for(const k of addrNames)o.businessAddress[k]=(use(k)&&p[k])?p[k]:(base[k]||"");}
  return o}
function editableFields(p,idx,existing=null){
  return `<div class="grid">${Object.entries(FIELD_META).map(([k,l])=>{
    const initial=(p[k]||"");
    if(k==="personalNotes")return `<div class="field full notes-field"><div class="notes-label-row"><label>${html(l)} <span class="muted">— editable before saving</span></label><div class="notes-actions"><button type="button" class="mini-btn" data-notes-action="clear" data-index="${idx}">Clear Notes</button><button type="button" class="mini-btn" data-notes-action="restore" data-index="${idx}">Restore Signature</button></div></div><textarea data-review="${idx}" data-field="${k}" spellcheck="true" aria-label="Editable Notes">${html(initial)}</textarea><div class="muted notes-help">Delete, shorten, or rewrite this text. Outlook will receive exactly what remains in this box.</div></div>`;
    return `<div class="field ${["email","street"].includes(k)?"full":""}"><label>${html(l)}</label><input data-review="${idx}" data-field="${k}" value="${html(initial)}"></div>`
  }).join("")}</div>`
}
function comparisonRows(parsed,existing){
  return Object.entries(FIELD_META).map(([key,label])=>{
    const nv=(parsed[key]||"").trim(),ov=(existing[key]||"").trim();
    const same=nv&&ov&&n(nv)===n(ov);
    const state=!nv?"No value found in email":same?"Same":"Different / new";
    const cls=!nv?"missing":same?"same":"changed";
    return `<div class="compare-row ${cls}"><div class="compare-label">${html(label)}</div><div><span class="compare-head">Existing Outlook</span><div>${html(ov||"(blank)")}</div></div><div><span class="compare-head">From email</span><div>${html(nv||"(not found)")}</div></div><div class="compare-state">${html(state)}</div></div>`
  }).join("")
}

function unaccountedSignaturePanel(p){
  const value=String(p?._audit?.unaccountedText||"").trim();
  const shown=value||"None";
  const cls=value?"unaccounted has-content":"unaccounted clear";
  return `<div class="${cls}"><div class="unaccounted-title">Unaccounted signature information</div><textarea readonly aria-label="Unaccounted signature information">${html(shown)}</textarea><div class="muted">${value?"This text was not confidently assigned to an Outlook field. Review the editable fields below and place it manually if appropriate. It is not saved automatically.":"Every meaningful part of the signature was either assigned to an Outlook field or recognized as intentional non-field signature material."}</div></div>`;
}

function renderReviews(items){
  const box=$("reviews");box.innerHTML="";
  items.forEach((it,idx)=>{
    const p=it.parsed,m=it.match,existing=m?existingFlat(m.contact):null,diffs=existing?fieldDiffs(p,existing):[];
    const el=document.createElement("div");el.className="review";el.dataset.reviewCard=idx;
    const badge=m?`<span class="badge existing">Existing contact — ${html(m.confidence)}</span>`:`<span class="badge">New contact</span>`;

    let actionArea="";
    if(m && diffs.length){
      actionArea=`<div class="diffs"><b>Choose fields to update</b>${diffs.map(d=>`<label class="diff"><input type="checkbox" data-diff="${idx}" data-key="${d.key}" ${d.defaultChecked?"checked":""}><span>${html(d.label)}</span><span class="vals"><span class="old">Existing: ${html(d.old||"(blank)")}</span><br><span class="new">From email: ${html(d.new)}</span></span></label>`).join("")}</div><div class="toolbar"><button class="btn primary" data-action="update" data-index="${idx}">Update Existing Contact</button><button class="btn" data-action="skip" data-index="${idx}">Skip</button></div>`;
    }else if(m){
      actionArea=`<div class="toolbar"><span class="muted">No different nonblank fields were found to update.</span><button class="btn" data-action="skip" data-index="${idx}">Done</button></div>`;
    }else{
      actionArea=`<div class="toolbar"><button class="btn primary" data-action="create" data-index="${idx}">Create Contact</button><button class="btn" data-action="skip" data-index="${idx}">Skip</button></div>`;
    }

    el.innerHTML=`${badge}<div class="candidate-name" style="margin-top:6px">${html([p.givenName,p.middleName,p.surname].filter(Boolean).join(" ")||it.name)}</div><div class="muted">${html(p.email||it.email||"")}</div><div class="signature-sticky"><div class="sig-title">Original signature block from email — preserved in Notes</div><pre>${html(p.signature||"No signature block confidently found")}</pre></div>${unaccountedSignaturePanel(p)}${diagnosticPanel(p)}${editableFields(p,idx,null)}${actionArea}`;
    box.appendChild(el);
  });
  $("reviewSection").hidden=false;
  box.querySelectorAll("button[data-action]").forEach(b=>b.addEventListener("click",handleAction));
  box.querySelectorAll("button[data-notes-action]").forEach(b=>b.addEventListener("click",e=>{
    const idx=Number(e.currentTarget.dataset.index);
    const ta=document.querySelector(`textarea[data-review="${idx}"][data-field="personalNotes"]`);
    if(!ta)return;
    if(e.currentTarget.dataset.notesAction==="clear")ta.value="";
    else ta.value=(window.__reviewItems[idx]?.parsed?.signature||"");
    ta.focus();
  }));
}
function currentParsed(idx){const d={};document.querySelectorAll(`[data-review="${idx}"][data-field]`).forEach(i=>d[i.dataset.field]=i.value.trim());d.businessPhone=phoneForOutlook(d.businessPhone);d.mobilePhone=phoneForOutlook(d.mobilePhone);d.businessFax=phoneForOutlook(d.businessFax);return d}
function finishCard(idx,label){const c=document.querySelector(`[data-review-card="${idx}"]`);c.classList.add("done");c.querySelectorAll("button").forEach(b=>b.disabled=true);const span=document.createElement("span");span.className="badge done";span.textContent=label;c.prepend(span)}
async function handleAction(ev){const action=ev.currentTarget.dataset.action,idx=Number(ev.currentTarget.dataset.index),item=window.__reviewItems[idx];if(action==="skip"){finishCard(idx,"Skipped");return}try{ev.currentTarget.disabled=true;status(action==="create"?"Creating Outlook contact…":"Updating Outlook contact…");const p=currentParsed(idx);if(action==="create"){await graph("/me/contacts",{method:"POST",body:JSON.stringify(graphPayload(p))});finishCard(idx,"Created");status("Contact created in Outlook Contacts.","ok")}else{const selected=new Set([...document.querySelectorAll(`input[data-diff="${idx}"]:checked`)].map(x=>x.dataset.key));if(!selected.size)throw new Error("Check at least one field to update.");await graph(`/me/contacts/${encodeURIComponent(item.match.contact.id)}`,{method:"PATCH",body:JSON.stringify(graphPayload(p,selected,existingFlat(item.match.contact)))});finishCard(idx,"Updated");status("Existing Outlook contact updated.","ok")}}catch(e){ev.currentTarget.disabled=false;status(e.message||String(e),"error")}}
async function compareSelected(){try{
  const selected=[...document.querySelectorAll("input[data-candidate]:checked")].map(x=>candidates[Number(x.dataset.candidate)]);
  if(selected.length!==1)throw new Error("Select exactly one signature block to process.");
  if(!clientId()){showAuthSetup();throw new Error("Complete the one-time Microsoft Contacts setup first.")}
  status("Signing in to Microsoft and checking Outlook Contacts…");graphContacts=await loadContacts();const items=selected.map(c=>({...c,match:matchContact(c.parsed)}));window.__reviewItems=items;renderReviews(items);status(`Compared the selected signature with ${graphContacts.length} Outlook contact${graphContacts.length===1?"":"s"}.`,"ok")
}catch(e){status(e.message||String(e),"error")}}
Office.onReady(async info=>{if(info.host!==Office.HostType.Outlook){status("This page must be opened from the Outlook add-in.","error");return}showAuthSetup();$("saveClientId").addEventListener("click",()=>showAuthSetup());$("selectAll").addEventListener("click",()=>document.querySelectorAll("input[data-candidate]").forEach(x=>x.checked=false));$("scanAgain").addEventListener("click",scan);$("compareSelected").addEventListener("click",compareSelected);await scan()});



