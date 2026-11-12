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
function structuredResolvedName(senderName,signatureText){
  const sender=cleanName(senderName||"");
  const inferred=cleanName(inferPersonName(signatureText)||"");
  if(!inferred)return looksLikePersonName(sender)?sender:"";
  if(!looksLikePersonName(sender))return inferred;
  const tokens=v=>cleanName(v).toLowerCase().match(/[a-z0-9]+/g)||[];
  const a=tokens(sender),b=tokens(inferred);
  // The visible signature is authoritative for name order when it contains
  // the same name tokens as Outlook's sender metadata (which is often
  // stored/displayed as Last First).
  if(a.length>=2&&a.length===b.length&&[...a].sort().join("|")===[...b].sort().join("|"))return inferred;
  return sender;
}
function phoneForOutlook(value,country=""){
  const raw=String(value||"").trim();if(!raw)return"";
  const extMatch=raw.match(/(?:^|\s)(?:x|ext\.?|extension)\s*[:.\-]?\s*(\d+)\s*$/i);
  const ext=extMatch?extMatch[1]:"";
  const original=(extMatch?raw.slice(0,extMatch.index):raw).trim();
  const main=original.replace(/\D/g,"");
  if(/^\+44/.test(original)){const national=main.slice(2).replace(/^0/,"");return "+44"+national+(ext?"x"+ext:"");}
  if(country==="United Kingdom"&&/^0\d{10}$/.test(main))return "+44"+main.slice(1)+(ext?"x"+ext:"");
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

function structuredPhoneObservations(sig,country=""){
  const observations=[];
  const re=/(?<!\d)(?:\+44\s*(?:\(0\)\s*|0\s*)?|0)(?:\d[\s().-]*){8,10}\d(?:\s*(?:x|ext\.?|extension)\s*[:.\-]?\s*\d+)?(?!\d)|(?:\+?1[\s.\-]?)?(?:\(?\d{3}\)?[\s.\-]*)\d{3}[\s.\-]*\d{4}(?:\s*(?:x|ext\.?|extension)\s*[:.\-]?\s*\d+)?/gi;

  const tokenLabel=token=>{
    const t=String(token||"").toLowerCase().replace(/[().]/g,"").replace(/\s+/g," ").trim();
    if(/^(?:direct\s*[/&+\-]\s*text|direct text|text\s*[/&+\-]\s*direct)$/.test(t))return"directtext";
    if(/^(?:fax|f)$/.test(t))return"fax";
    if(/^(?:mobile|cell|cellular|m|c)$/.test(t))return"mobile";
    if(/^(?:office|business|work|main|telephone|phone|tel|p|o|b|w)$/.test(t))return"business";
    if(/^(?:direct|d|dd|ddi)$/.test(t))return"direct";
    return"";
  };

  const prefixLabel=segment=>{
    const s=String(segment||"");
    const m=s.match(/(?:^|[|•;,]\s*|\s+)(direct\s*[/&+\-]\s*text|direct text|text\s*[/&+\-]\s*direct|office|business|work|main|telephone|phone|tel|mobile|cell|cellular|fax|direct|ddi|dd|[PBOMCFDW])\s*[:.\-]?\s*$/i);
    return m?tokenLabel(m[1]):"";
  };

  const suffixLabel=segment=>{
    const s=String(segment||"");
    let m=s.match(/^\s*(?:[|•;,]\s*)?\(\s*(direct\s*[/&+\-]\s*text|direct text|office|business|work|main|telephone|phone|tel|mobile|cell|cellular|fax|direct|ddi|dd|[BOMCFDW])\s*\)/i);
    if(m)return tokenLabel(m[1]);
    // A bare word after one number is often the prefix for the NEXT number
    // ("Mobile 970... Office 720...").  Treat a word as a suffix only when
    // a visible separator binds it to the preceding number ("954... | Mobile").
    m=s.match(/^\s*[|•;,]\s*(direct\s*[/&+\-]\s*text|direct text|office|business|work|main|telephone|phone|tel|mobile|cell|cellular|fax|direct|[BOMCFDW])\b(?!\s*:)/i);
    return m?tokenLabel(m[1]):"";
  };

  for(const rawLine of norm(sig).split("\n")){
    const line=String(rawLine||"").replace(/\u00a0/g," ").trim();
    if(!line)continue;
    const matches=[...line.matchAll(re)];
    for(let k=0;k<matches.length;k++){
      const m=matches[k];
      const value=phoneForOutlook(m[0],country); if(!value)continue;
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
function structuredPhones(sig,country=""){
  const obs=structuredPhoneObservations(sig,country);
  const out={businessPhone:"",mobilePhone:"",businessFax:""};
  const take=(kind,o)=>{if(o&&!out[kind])out[kind]=o.value};
  const fax=obs.filter(o=>o.label==="fax");
  const bus=obs.filter(o=>o.label==="business");
  const mob=obs.filter(o=>o.label==="mobile");
  const directText=obs.filter(o=>o.label==="directtext");
  const direct=obs.filter(o=>o.label==="direct");
  const unlabeled=obs.filter(o=>!o.label);
  take("businessFax",fax[0]);
  take("businessPhone",bus[0]);
  take("mobilePhone",mob[0]||directText[0]);

  // Direct with an extension is a company/direct-dial number, so Business.
  for(const o of direct.filter(x=>x.hasExtension))take("businessPhone",o);

  // Assign one unlabeled number as Business before interpreting a bare Direct.
  // This makes an unlabeled main number + Direct number map Business + Mobile.
  for(const o of unlabeled)take("businessPhone",o);

  for(const o of direct.filter(x=>!x.hasExtension)){
    // Use Direct Dial as Mobile only when no explicitly identified Mobile exists.
    if(!out.mobilePhone)take("mobilePhone",o);
    else if(!out.businessPhone)take("businessPhone",o);
  }

  // If only an explicit mobile existed, a remaining unlabeled number is Business.
  if(!out.businessPhone){for(const o of unlabeled)take("businessPhone",o)}

  // Explicit/relational observations are authoritative. Never let the legacy
  // parser reclassify a number that already carried an Office/Business,
  // Mobile/Cell, or Fax label. Use legacy parsing only when no structured
  // phone observation was possible at all.
  if(!obs.length){
    const legacy=phones(sig);
    out.businessPhone=legacy.businessPhone||"";
    out.mobilePhone=legacy.mobilePhone||"";
    out.businessFax=legacy.businessFax||"";
  }
  if(out.businessPhone===out.mobilePhone)out.mobilePhone="";
  if(out.businessPhone===out.businessFax)out.businessFax="";
  return out;
}

// A labeled office address may contain both a physical street and a postal
// P.O. Box on the same line. Parse complete address segments before assigning
// any of them to the single Outlook businessAddress.street field.
function v4CompoundOfficeAddress(sig){
  for(const raw of norm(sig).split("\n")){
    const line=String(raw||"").replace(/\u00a0/g," ").replace(/\s+/g," ").trim();
    if(!line.includes("|"))continue;
    const match=line.match(/^([A-Za-z][A-Za-z .'’-]{1,50}?\s+(?:Office|Branch|Location))\s*:\s*(.+)$/i);
    const body=match?match[2]:line;
    const segments=body.split(/\s*\|\s*/).map(x=>x.trim()).filter(Boolean);
    if(segments.length<2||segments.length>4)continue;
    const addresses=[];
    for(const part of segments){
      const m=part.match(/^(.+),\s*([A-Za-z][A-Za-z .'-]{1,60}),\s*([A-Z]{2})\s*,?\s*(\d{5}(?:-\d{4})?)$/i);
      if(!m)break;
      const street=m[1].trim(),city=m[2].trim(),state=structuredStateAbbrev(m[3]);
      if(!state||!/\d/.test(street)||!city)break;
      addresses.push({street,city,state,postalCode:m[4]});
    }
    if(addresses.length!==segments.length)continue;
    const first=addresses[0];
    if(!addresses.every(a=>a.city.toLowerCase()===first.city.toLowerCase()&&a.state===first.state&&a.postalCode===first.postalCode))continue;
    const streetLines=[...new Set(addresses.map(a=>a.street))];
    return {street:streetLines.join("\n"),city:first.city,state:first.state,postalCode:first.postalCode,
      countryOrRegion:"United States",_sourceLine:line,_officeLabel:match?match[1]:""};
  }
  return null;
}

function structuredAddress(sig){
  const compound=v4CompoundOfficeAddress(sig);
  if(compound)return compound;
  // Royal Mail postcode syntax; never invent an American state for a UK address.
  const ukLines=norm(sig).split("\n").map(x=>String(x||"").replace(/\u00a0/g," ").trim()).filter(Boolean);
  const ukPost=/(?:\bGIR\s?0AA\b|\b[A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2}\b)/i;
  for(let i=0;i<ukLines.length;i++){
    const hit=ukLines[i].match(ukPost);if(!hit)continue;
    const postalCode=hit[0].toUpperCase().replace(/\s+/g,"").replace(/(.{3})$/, " $1");
    const before=ukLines[i].slice(0,hit.index).replace(/[?,;\s]+$/g,"");
    const components=before.split(",").map(x=>x.replace(/^[?\s]+/g,"").trim()).filter(Boolean);
    let city=components.length?components[components.length-1]:"";
    let street=components.length>=2?components[components.length-2]:"";
    if(!city&&i>0){city=ukLines[i-1].trim();if(i>1)street=ukLines[i-2].trim();}
    else if(!street&&i>0)street=ukLines[i-1].trim();
    if(city&&/^(?:united kingdom|uk|england|scotland|wales|northern ireland)$/i.test(city))city="";
    if(street&&/[@<>]|https?:|^www\./i.test(street))street="";
    return {street,city,state:"",postalCode,countryOrRegion:"United Kingdom"};
  }
  // A common Outlook line wrap places "Street, City," immediately before
  // a separate "ST 12345" line. Keep the street rather than dropping it.
  for(let i=1;i<ukLines.length;i++){
    const stateZip=ukLines[i].match(/^([A-Z]{2})\s*,?\s*(\d{5}(?:-\d{4})?)$/);
    if(!stateZip||!structuredStateAbbrev(stateZip[1]))continue;
    const preceding=ukLines[i-1].replace(/[,\s]+$/g,"");
    const parts=preceding.split(",").map(x=>x.trim()).filter(Boolean);
    if(parts.length<2)continue;
    const street=parts.slice(0,-1).join(", ");
    const city=parts[parts.length-1];
    if(!street||!city||!(/\d/.test(street)||/\b(?:street|road|rd|avenue|ave|drive|plaza|place|lane|boulevard)\b/i.test(street)))continue;
    return {street,city,state:stateZip[1],postalCode:stateZip[2],countryOrRegion:"United States"};
  }
  const legacy=address(sig);
  if(legacy.street&&legacy.city&&legacy.state&&legacy.postalCode)return legacy;
  const lines=norm(sig).split("\n").map(x=>String(x||"").replace(/\u00a0/g," ").replace(/\s+/g," ").trim()).filter(Boolean);
  const stateNames=Object.keys(US_STATE_NAMES).sort((a,b)=>b.length-a.length).map(x=>x.replace(/ /g,"\\s+")).join("|");
  const locRx=new RegExp("([A-Za-z][A-Za-z .'-]{1,60}?),?\\s+("+stateNames+"|[A-Z]{2})\\s*,?\\s*(\\d{5}(?:-\\d{4})?)","i");
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

function structuredLegalCompanyEvidence(line){
  const v=String(line||"").trim(); if(!v)return false;
  // Legal/entity designators at the end of a line are definitive company
  // evidence. Covers common U.S. and international forms.
  return /(?:^|[,\s])(?:llc|l\.l\.c\.?|inc\.?|incorporated|corp\.?|corporation|pllc|lp|llp|ltd\.?|limited|plc|gmbh|ag|s\.?a\.?|sarl|s\.?a\.?r\.?l\.?|sas|b\.?v\.?|n\.?v\.?|pty\.?\s+ltd\.?|co\.?)\s*\.?$/i.test(v);
}

function structuredCompanyStrength(line,senderEmail,site){
  if(structuredLegalCompanyEvidence(line))return 3;
  if(structuredCompanyMatchesDomain(line,senderEmail,site))return 2;
  if(structuredCompanyEvidence(line))return 1;
  return 0;
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
  // Stage 1 only: locate and freeze complete visible signature blocks.
  // Do not parse fields here and do not use missing Outlook fields to change a boundary.
  return captureSignatureBlocks(body,currentName||"",currentEmail||"",myEmail||"");
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
// Notes are plain text. Keep the full immutable signature for discovery,
// parsing and preview, but remove layout-only empty lines in the Notes copy.
// Every nonblank line remains in its original order with its text intact.
function compactSignatureNotes(signatureText){
  return norm(signatureText||"").split("\n")
    .map(line=>String(line).replace(/[\u00ad\u200b-\u200d\ufeff]/g,"").trim())
    .filter(Boolean)
    .join("\n");
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

  const addressDecision=parsed._addressDecision||null;
  const normAuditLine=v=>String(v||"").replace(/\u00a0/g," ").replace(/\s+/g," ").trim();

  for(const line of lines){
    if(parsed._sourceLine&&normAuditLine(line).toLowerCase()===normAuditLine(parsed._sourceLine).toLowerCase()){
      recognized.push({text:line,category:"complete physical and postal address"});
      continue;
    }
    if(addressDecision?.branchLine&&
       normAuditLine(line).toLowerCase()===normAuditLine(addressDecision.branchLine).toLowerCase()){
      recognized.push({text:line,category:"branch/address selector"});
      continue;
    }

    const alternate=addressDecision?.alternatives?.find(a=>
      normAuditLine(a.rawLine).toLowerCase()===normAuditLine(line).toLowerCase()
    );
    if(alternate){
      unaccounted.push({line,text:line});
      continue;
    }

    const nonField=v4RecognizedNonFieldLine(line);
    if(nonField){recognized.push({text:line,category:nonField});continue}
    let residue=line;

    const selectedAddress=addressDecision?.selected;
    if(selectedAddress&&selectedAddress.label&&
       normAuditLine(selectedAddress.rawLine).toLowerCase()===normAuditLine(line).toLowerCase())
      residue=v4RemoveCI(residue,selectedAddress.label);

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
      "workshop","garage","sealants","sealant","waterproofing"
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

  const companies=candidates.filter(a=>
    !(v4RecognizedNonFieldLine(a.line)&&!structuredCompanyMatchesDomain(a.line,email,site)&&!structuredLegalCompanyEvidence(a.text))
  ).map(a=>{
    let score=0;
    if(structuredLegalCompanyEvidence(a.text))score+=100;
    if(structuredCompanyMatchesDomain(a.text,email,site))score+=50;
    if(structuredCompanyEvidence(a.text))score+=10;
    if(a.text===a.text.toUpperCase()&&a.text.split(/\s+/).length>=2)score+=2;
    return{...a,score};
  }).filter(a=>a.score>0).sort((a,b)=>b.score-a.score||a.i-b.i);

  // Credentials, office-location footers, marketing lines and role text
  // are never company material.
  if(parsed.companyName&&v4CredentialList(parsed.companyName).length)parsed.companyName="";
  if(parsed.companyName&&v4RecognizedNonFieldLine(parsed.companyName)&&
     !structuredCompanyMatchesDomain(parsed.companyName,email,site))parsed.companyName="";
  if(parsed.companyName&&parsed.jobTitle&&
     parsed.companyName.trim().toLowerCase()===parsed.jobTitle.trim().toLowerCase())
    parsed.companyName="";

  let currentCompany=String(parsed.companyName||"");
  let currentStrength=structuredCompanyStrength(currentCompany,email,site);

  // If the current company looks like a human role and has no organization/domain
  // evidence, discard it and let the whole-block evidence choose again.
  if(currentCompany&&!currentStrength&&structuredPlausibleTitleText(currentCompany)){
    parsed.companyName="";
    currentCompany="";
    currentStrength=0;
  }

  if(companies.length){
    const best=companies[0];
    const bestStrength=structuredCompanyStrength(best.text,email,site);
    // A legal entity suffix outranks domain inference; domain agreement outranks
    // generic organization words such as "Solutions" or "Services".
    if(!currentCompany||bestStrength>currentStrength){
      parsed.companyName=best.line.includes("|")&&structuredCompanyMatchesDomain(best.line,email,site)
        ? best.line : best.text;
    }
  }

  // Some signatures use only a logo/website for the company. If no explicit
  // company survives, infer the organization from the sender/site domain.
  const inferredCompany=v4CompanyFromDomain(email,site);
  if(!parsed.companyName){
    if(inferredCompany)parsed.companyName=inferredCompany;
  }else if(inferredCompany){
    const companyKey=structuredCompanyLineKey(parsed.companyName);
    const inferredKey=structuredCompanyLineKey(inferredCompany);
    const explicitCompanyLine=lines.some(line=>{
      if(cleanEmail(line)||structuredWebsiteLike(line)||structuredPhoneObservations(line).length)return false;
      const parts=structuredDelimitedParts(line);
      return structuredCompanyLineKey(line)===companyKey ||
        parts.some(part=>structuredCompanyLineKey(part)===companyKey);
    });
    if(!explicitCompanyLine&&companyKey&&companyKey===inferredKey&&
       !/[\s|,&]/.test(parsed.companyName))
      parsed.companyName=inferredCompany;
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
      if(v4RecognizedNonFieldLine(a.line))return false;
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


function v402NormalizeOfficeKey(value){
  return String(value||"").toLowerCase().replace(/[^a-z0-9]+/g," ").trim();
}

function v402OfficeAddressCandidates(signatureText){
  const lines=norm(v4VisibleSignatureBlock(signatureText)).split("\n")
    .map(x=>String(x||"").replace(/\u00a0/g," ").replace(/\s+/g," ").trim());

  const stateNames=Object.keys(US_STATE_NAMES)
    .sort((a,b)=>b.length-a.length)
    .map(x=>x.replace(/ /g,"\\s+"))
    .join("|");
  const commaRx=new RegExp(
    "^(.+?)\\s*,\\s*([A-Za-z][A-Za-z .'-]{1,60})\\s*,?\\s*("+stateNames+"|[A-Z]{2})\\s+(\\d{5}(?:-\\d{4})?)\\s*$",
    "i"
  );
  const suffixRx=new RegExp(
    "^(.+?\\b(?:street|st\\.?|road|rd\\.?|avenue|ave\\.?|boulevard|blvd\\.?|drive|dr\\.?|lane|ln\\.?|way|court|ct\\.?|place|pl\\.?|loop|terrace|ter\\.?|highway|hwy\\.?))\\s+([A-Za-z][A-Za-z .'-]{1,60})\\s*,\\s*("+stateNames+"|[A-Z]{2})\\s+(\\d{5}(?:-\\d{4})?)\\s*$",
    "i"
  );

  const out=[];
  for(let i=0;i<lines.length;i++){
    const line=lines[i];
    if(!line||cleanEmail(line)||structuredPhoneObservations(line).length)continue;
    let m=line.match(commaRx);
    if(!m)m=line.match(suffixRx);
    if(!m)continue;
    const state=structuredStateAbbrev(m[3]);
    if(!state)continue;

    const prefix=m[1].trim().replace(/[,\s]+$/g,"");
    const start=prefix.search(/(?:p\.?o\.?\s*box|\d)/i);
    if(start<0)continue;

    const label=prefix.slice(0,start).trim().replace(/[\s|:;,-]+$/g,"");
    const street=prefix.slice(start).trim().replace(/[,\s]+$/g,"");
    if(!street||!/\d/.test(street))continue;

    out.push({
      lineIndex:i,rawLine:line,label,street,
      city:m[2].trim(),state,postalCode:m[4],countryOrRegion:"United States"
    });
  }
  return out;
}

function v402BranchContext(signatureText,resolvedName){
  const lines=norm(v4VisibleSignatureBlock(signatureText)).split("\n")
    .map(x=>String(x||"").replace(/\u00a0/g," ").replace(/\s+/g," ").trim());

  let nameIndex=-1;
  const surname=nameParts(resolvedName||"").surname||"";
  for(let i=0;i<lines.length;i++){
    const v=lines[i];
    if(!v)continue;
    if(resolvedName&&nameMatchesIdentity(v,resolvedName)){nameIndex=i;break}
    if(surname&&new RegExp("\\b"+surname.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")+"\\b","i").test(v)&&personNameFromLine(v)){
      nameIndex=i;break;
    }
  }

  const start=nameIndex>=0?nameIndex+1:0;
  const end=Math.min(lines.length,start+8);
  for(let i=start;i<end;i++){
    const v=lines[i];
    if(!v)continue;
    const m=v.match(/^(.{2,50}?)\s+(?:Branch|Office|Location)$/i);
    if(m)return {line:v,name:m[1].trim(),key:v402NormalizeOfficeKey(m[1])};
  }
  return {line:"",name:"",key:""};
}

function v402SelectOfficeAddress(parsed,signatureText,resolvedName){
  if(parsed._sourceLine)return parsed;
  const candidates=v402OfficeAddressCandidates(signatureText);
  if(candidates.length<2)return parsed;

  const branch=v402BranchContext(signatureText,resolvedName);
  let selected=null;

  if(branch.key){
    const scored=candidates.map(c=>{
      const lk=v402NormalizeOfficeKey(c.label);
      const ck=v402NormalizeOfficeKey(c.city);
      let score=0;
      if(lk===branch.key)score=100;
      else if(lk&&lk.split(" ").includes(branch.key))score=95;
      else if(lk&&lk.includes(branch.key))score=90;
      else if(ck===branch.key)score=80;
      return {c,score};
    }).sort((a,b)=>b.score-a.score||a.c.lineIndex-b.c.lineIndex);
    if(scored[0]?.score>0)selected=scored[0].c;
  }

  if(!selected)selected=candidates.find(c=>/\bmain\b/i.test(c.label))||null;

  if(!selected&&parsed.street){
    const ps=v402NormalizeOfficeKey(parsed.street);
    selected=candidates.find(c=>v402NormalizeOfficeKey(c.street)===ps)||null;
  }

  if(!selected)selected=candidates[0];

  parsed.street=selected.street;
  parsed.city=selected.city;
  parsed.state=selected.state;
  parsed.postalCode=selected.postalCode;
  parsed.countryOrRegion=selected.countryOrRegion;
  parsed._addressDecision={
    branchLine:branch.line,branchName:branch.name,selected,
    alternatives:candidates.filter(c=>c!==selected)
  };
  return parsed;
}


function v4ContactFromSignature(senderName,senderEmail,signatureText){
  const original=v4VisibleSignatureBlock(signatureText);
  const base=structuredContactFromSignature(senderName,senderEmail,original);
  v4FillPartialAddress(base,original);
  const resolvedName=structuredResolvedName(senderName,original);
  const credentials=v4FindCredentials(norm(original).split("\n"),resolvedName);

  v402SelectOfficeAddress(base,original,resolvedName);
  v4ResidualTitleCompany(base,original,resolvedName,credentials);

  const np=nameParts(resolvedName||[base.givenName,base.middleName,base.surname].filter(Boolean).join(" "));
  if(np.givenName)base.givenName=np.givenName;
  if(np.middleName)base.middleName=np.middleName;
  if(np.surname)base.surname=np.surname+(credentials.length?", "+credentials.join(", "):"");

  base.signature=original;
  base.personalNotes=compactSignatureNotes(original);

  let audit=v4AuditSignature(original,base,resolvedName,credentials);
  if(audit.unaccounted.length&&(!base.jobTitle||!base.companyName)){
    v4ResidualTitleCompany(base,original,resolvedName,credentials);
    audit=v4AuditSignature(original,base,resolvedName,credentials);
  }

  base._audit={credentials,recognizedNonField:audit.recognizedNonField,unaccounted:audit.unaccounted,unaccountedText:audit.unaccountedText};
  base._debug={...(base._debug||{}),v4HolisticParser:true,credentials,unaccountedLines:audit.unaccounted};
  return base;
}


function structuredContactFromSignature(senderName,senderEmail,signatureText){
  const sig=cleanSignatureText(signatureText);
  const lines=norm(sig).split("\n")
    .map(x=>String(x||"").replace(/\u00a0/g," ").replace(/\s+/g," ").trim())
    .filter(Boolean);

  const resolvedName=structuredResolvedName(senderName,sig);
  const email=cleanEmail(senderEmail)||cleanEmail(sig);
  const site=website(sig,email);
  const addr=structuredAddress(sig);
  const ph=structuredPhones(sig,addr.countryOrRegion);
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
        if(structuredLegalCompanyEvidence(p))score+=100;
        if(structuredCompanyMatchesDomain(p,email,site))score+=50;
        if(structuredCompanyEvidence(p))score+=10;
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
      if(structuredLegalCompanyEvidence(p))score+=100;
      if(structuredCompanyMatchesDomain(p,email,site))score+=50;
      if(structuredCompanyEvidence(p))score+=10;
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
      if(structuredLegalCompanyEvidence(v))score+=100;
      if(structuredCompanyMatchesDomain(v,email,site))score+=50;
      if(structuredCompanyEvidence(v))score+=10;
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
      if(structuredLegalCompanyEvidence(v))score+=100;
      if(structuredCompanyMatchesDomain(v,email,site))score+=50;
      if(structuredCompanyEvidence(v))score+=10;
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

// -----------------------------------------------------------------------------
// Stage 1: immutable signature capture
// -----------------------------------------------------------------------------
// A captured block is the source of truth for both Notes and all downstream
// parsing.  Field detection is deliberately forbidden from changing the block.
function signatureCaptureLine(v){
  const out=v4DisplayText(signatureDiscoveryText(String(v||"")))
    .replace(/[\u00ad\u200b-\u200d\ufeff]/g,"")
    .replace(/\u00a0/g," ")
    .replace(/[ \t]+$/g,"");
  // Outlook plain text can expose inline-logo attachment placeholders that are
  // not visible signature text.  Do not copy those transport artifacts to Notes.
  if(/^\s*<image\d+\.(?:png|jpe?g|gif|bmp|webp)>\s*$/i.test(out))return "";
  return out;
}
function signatureHardBoundary(line){
  const v=String(line||"").trim();
  if(!v)return false;
  if(/^\s*(?:from|sent|to|cc|bcc|subject):\s*/i.test(v))return true;
  if(/^(?:confidentiality notice|confidential notice|legal notice|disclaimer)\b/i.test(v))return true;
  if(/\b(?:this communication|this e-?mail).{0,45}\bconfidential\b/i.test(v))return true;
  return false;
}
function visibleNameAtLine(line,identityName=""){
  const direct=cleanName(signatureNameFromLine(line)||personNameFromLine(line)||"");
  if(direct&&(!identityName||nameMatchesIdentity(line,identityName)))return direct;
  return "";
}
function signatureAnchorEvidence(lines,start,identityEmail){
  const probe=[];
  let nonblank=0;
  for(let j=start;j<lines.length&&j<start+24&&nonblank<12;j++){
    const v=signatureCaptureLine(lines[j]);
    if(signatureHardBoundary(v)&&j>start)break;
    if(!v.trim())continue;
    probe.push(v.trim());nonblank++;
  }
  const t=probe.join("\n");
  if(identityEmail&&sameEmail(cleanEmail(t),identityEmail))return true;
  if(phoneTokens(t).length)return true;
  if(/\b[A-Z]{2}\s*,?\s*\d{5}(?:-\d{4})?\b/i.test(t))return true;
  if(/\b(?:inc\.?|incorporated|llc|l\.l\.c\.?|corp\.?|corporation|ltd\.?|limited|llp|pllc|gmbh|s\.a\.?|s\.r\.l\.?|p\.c\.?)\b/i.test(t))return true;
  if(/https?:\/\/|www\.|@[a-z0-9.-]+\.[a-z]{2,}/i.test(t))return true;
  if(TITLE_WORDS.some(x=>t.toLowerCase().includes(x)))return true;
  return false;
}
function captureSignatureFromStart(lines,start,identityName,identityEmail){
  if(start<0||start>=lines.length)return null;
  let end=start;
  let blankRun=0,nonblank=0;
  for(let j=start;j<lines.length&&j<start+60&&nonblank<32;j++){
    const raw=signatureCaptureLine(lines[j]);
    const v=raw.trim();
    if(j>start&&signatureHardBoundary(v))break;
    if(j>start&&v&&nameMatchesIdentity(v,identityName))break;
    if(!v){
      blankRun++;
      // Outlook often inserts a blank line between every signature row.  Only a
      // long visual gap is considered a boundary, and only after content exists.
      if(blankRun>=3&&nonblank>=2&&!signatureContinuationAhead(lines,j))break;
      end=j;
      continue;
    }
    blankRun=0;
    // A long prose paragraph after signature evidence is not signature material.
    if(j>start&&nonblank>=2&&v.length>180&&/[.!?]$/.test(v)&&!cleanEmail(v)&&!phoneTokens(v).length)break;
    end=j;nonblank++;
  }
  if(!signatureAnchorEvidence(lines,start,identityEmail))return null;
  const block=v4VisibleSignatureBlock(lines.slice(start,end+1).map(signatureCaptureLine).join("\n"));
  if(!block)return null;
  return {start,end,block};
}

// Find signature owners from the visible block, not just the From header.
// Some Outlook plain-text bodies split a visually single-line name across
// separated first/last-name lines. Evidence from title/company/email is required.
function captureVisibleNameInfo(lines,start){
  const first=signatureCaptureLine(lines[start]).trim();
  if(!first||signatureHardBoundary(first)||isClosingPhrase(first)||/^(?:hi|hello|dear|hey|good morning|good afternoon)\b/i.test(first))return null;
  const single=v=>/^[A-Z][A-Za-z'’\-]{1,34}$/.test(v);
  const following=[];
  for(let j=start+1;j<lines.length&&j<=start+18&&following.length<8;j++){
    const v=signatureCaptureLine(lines[j]).trim();
    if(signatureHardBoundary(v))break;
    if(v)following.push({index:j,value:v});
  }
  let name=cleanName(signatureNameFromLine(first)||personNameFromLine(first)||"");
  let nameEnd=start;
  if(!name&&single(first)&&following.length&&single(following[0].value)){
    const combined=first+" "+following[0].value;
    if(looksLikePersonName(combined)){
      name=combined;
      nameEnd=following[0].index;
    }
  }
  if(!name)return null;
  const after=following.filter(x=>x.index>nameEnd).slice(0,6).map(x=>x.value);
  const shortField=v=>v.length<=75&&v.split(/\s+/).length<=9&&!/[!?]$/.test(v)&&!cleanEmail(v);
  const titleIndex=after.slice(0,2).findIndex(v=>shortField(v)&&TITLE_WORDS.some(t=>v.toLowerCase().includes(t)));
  if(titleIndex<0)return null;
  const hasCompany=after.slice(titleIndex+1,titleIndex+4).some(v=>shortField(v)&&(structuredLegalCompanyEvidence(v)||structuredCompanyEvidence(v)));
  if(!hasCompany)return null;
  return {name,start};
}
function captureVisibleNameAnchors(lines,myEmail){
  const anchors=[];
  for(let i=0;i<lines.length;i++){
    const candidate=captureVisibleNameInfo(lines,i);
    if(!candidate)continue;
    let email="";
    for(let j=i+1;j<lines.length&&j<i+85;j++){
      const v=signatureCaptureLine(lines[j]);
      if(signatureHardBoundary(v))break;
      const e=cleanEmail(v);
      if(e&&!sameEmail(e,myEmail)){email=e;break}
    }
    if(email)anchors.push({...candidate,email});
  }
  return anchors;
}
function signatureContinuationAhead(lines,at){
  // Preserve layout gaps inside signatures but not email headers/disclaimers.
  for(let j=at+1;j<lines.length&&j<=at+22;j++){
    const v=signatureCaptureLine(lines[j]).trim();
    if(signatureHardBoundary(v))return false;
    if(!v)continue;
    if(cleanEmail(v)||structuredPhoneObservations(v).length||/\b[A-Z]{2}\s*\d{5}(?:-\d{4})?\b/.test(v)||
       /\b(?:[A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})\b/i.test(v)||/\bwww\.|https?:\/\//i.test(v))return true;
  }
  return false;
}

function captureSegmentSignatureBlocks(segment,myEmail){
  const identityEmail=cleanEmail(segment.email||"");
  const identityName=cleanName(segment.name||"");
  if(identityEmail&&sameEmail(identityEmail,myEmail))return [];
  const lines=norm(segment.text||"").split("\n");
  const starts=[];

  // Best anchor: the person's visibly written name.  This outranks directory
  // metadata and preserves the order the signer actually used.
  if(identityName){
    for(let i=0;i<lines.length;i++)if(nameMatchesIdentity(signatureCaptureLine(lines[i]),identityName))starts.push(i);
  }

  // Fallback: a visible signature email with a nearby human name above it.
  if(!starts.length&&identityEmail){
    for(let i=0;i<lines.length;i++){
      if(!sameEmail(cleanEmail(signatureCaptureLine(lines[i])),identityEmail))continue;
      for(let j=i-1;j>=Math.max(0,i-18);j--){
        const v=signatureCaptureLine(lines[j]).trim();
        if(signatureHardBoundary(v))break;
        if(!v||isClosingPhrase(v))continue;
        const n=cleanName(signatureNameFromLine(v)||personNameFromLine(v)||"");
        if(n){starts.push(j);break}
      }
    }
  }

  const localNames=new Map();
  for(const a of captureVisibleNameAnchors(lines,myEmail)){
    starts.push(a.start);localNames.set(a.start,a);
  }
  const out=[],used=[];
  for(const start of [...new Set(starts)].sort((a,b)=>a-b)){
    if(used.some(r=>start>=r.start&&start<=r.end))continue;
    const cap=captureSignatureFromStart(lines,start,identityName,identityEmail);
    if(!cap)continue;
    const visibleName=localNames.get(start)?.name||visibleNameAtLine(lines[start],identityName)||inferPersonName(cap.block)||identityName;
    const email=localNames.get(start)?.email||cleanEmail(cap.block)||identityEmail;
    const parsed=v4ContactFromSignature(visibleName,email,cap.block);
    // Freeze the full Stage 1 source. Notes is a separate compact text copy.
    parsed.signature=cap.block;
    parsed.personalNotes=compactSignatureNotes(cap.block);
    parsed._debug={...(parsed._debug||{}),source:"Stage 1 immutable signature capture",capturedStart:cap.start,capturedEnd:cap.end};
    out.push({
      score:100,name:visibleName||identityName,email,text:cap.block,rawSignatureBlock:cap.block,
      parsed,candidateType:"signature",physicalOccurrence:true,
      occurrenceKey:`capture:${segment.start}:${cap.start}:${cap.end}`,
      globalStart:segment.start+cap.start,globalEnd:segment.start+cap.end
    });
    used.push({start:cap.start,end:cap.end});
  }
  return out;
}
function captureGenericSignatureBlocks(body,myEmail,claimed){
  const lines=norm(body||"").split("\n"),out=[];
  const ranges=claimed||[];
  const overlaps=(a,b)=>ranges.some(r=>!(b<r.start||a>r.end));
  for(let i=0;i<lines.length;i++){
    const current=signatureCaptureLine(lines[i]);
    if(signatureHardBoundary(current))continue;
    const email=cleanEmail(current);
    if(!email||sameEmail(email,myEmail))continue;
    let start=-1,name="";
    for(let j=i-1;j>=Math.max(0,i-18);j--){
      const v=signatureCaptureLine(lines[j]).trim();
      if(signatureHardBoundary(v))break;
      if(!v||isClosingPhrase(v))continue;
      const n=cleanName(signatureNameFromLine(v)||personNameFromLine(v)||"");
      if(n){start=j;name=n;break}
    }
    if(start<0)continue;
    const cap=captureSignatureFromStart(lines,start,name,email);
    if(!cap||overlaps(cap.start,cap.end))continue;
    const parsed=v4ContactFromSignature(name,email,cap.block);
    parsed.signature=cap.block;parsed.personalNotes=compactSignatureNotes(cap.block);
    parsed._debug={...(parsed._debug||{}),source:"Stage 1 generic immutable signature capture",capturedStart:cap.start,capturedEnd:cap.end};
    out.push({score:70,name,email,text:cap.block,rawSignatureBlock:cap.block,parsed,candidateType:"signature",physicalOccurrence:true,occurrenceKey:`generic-capture:${cap.start}:${cap.end}`,globalStart:cap.start,globalEnd:cap.end});
    ranges.push({start:cap.start,end:cap.end});
  }
  return out;
}
function captureSignatureBlocks(body,currentName,currentEmail,myEmail){
  const segments=allMessageSegments(body,currentName,currentEmail),out=[],ranges=[];
  for(const seg of segments){
    for(const c of captureSegmentSignatureBlocks(seg,myEmail)){
      out.push(c);ranges.push({start:c.globalStart,end:c.globalEnd});
    }
  }
  out.push(...captureGenericSignatureBlocks(body,myEmail,ranges));
  out.sort((a,b)=>(a.globalStart??0)-(b.globalStart??0));
  return out;
}

function nameMatchesIdentity(line,identityName){
  const raw=cleanName(signatureNameFromLine(line)||personNameFromLine(line)||"");
  const wanted=cleanName(identityName||"");
  if(!raw||!wanted)return false;
  const tokens=v=>String(v||"").toLowerCase().match(/[a-z0-9]+/g)||[];
  const a=tokens(raw),b=tokens(wanted);
  if(a.join(" ")===b.join(" "))return true;

  // Outlook/Exchange frequently exposes directory names as "Last First".
  // Name order is therefore not authoritative metadata.  Match the same visible
  // name tokens regardless of order, while tolerating a middle initial on one side.
  const significant=x=>x.filter(t=>t.length>1);
  const as=significant(a),bs=significant(b);
  const aset=new Set(as),bset=new Set(bs);
  if(as.length>=2&&bs.length>=2&&as.every(t=>bset.has(t))&&bs.every(t=>aset.has(t)))return true;
  if(as.length>=2&&bs.length>=2&&(as.every(t=>bset.has(t))||bs.every(t=>aset.has(t))))return true;
  return false;
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
    const logoInfo=c._imageEvidence
      ? `<div class="muted">Signature image: ${html(c._imageEvidence.company||"detected — company not yet verified")}</div>`
      : "";
    const el=document.createElement("div");el.className="candidate";
    el.innerHTML=`<div class="candidate-head"><input type="checkbox" data-candidate="${i}"><div><div class="candidate-name">${html(c.name||"Unknown sender")}</div><div class="muted">${html(c.email||"No email found")}</div>${occ}${logoInfo}</div></div><div class="preview">${html(c.parsed.signature||"No contact block confidently found")}</div>`;
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
      // Image ALT text alone is not authoritative company data.
      // Its image is inventoried and inspected separately.
      img.replaceWith(clone.ownerDocument.createTextNode(" "));
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
// A company whose name exists only in a signature logo is not fully recoverable
// from plain text. OCR is a conservative optional enrichment. Accept only a
// multiword organization whose initials or literal spelling agree with the
// candidate's own email domain; unrelated logos in forwarded chains are ignored.
function logoCompanyFromOcrText(ocrText,email,site=""){
  const emailDomain=cleanEmail(email).split("@")[1]||"";
  if(!emailDomain||/^(?:gmail|outlook|hotmail|yahoo|icloud|aol|live)\./i.test(emailDomain))return "";
  const key=structuredDomainKey(email,site);
  const stem=key.replace(/(?:law|legal|group|llp|inc|company|co)$/i,"");
  const domainKeys=[key,stem].filter(k=>k.length>=3);
  const lines=norm(ocrText||"").split("\n").map(x=>x.replace(/[^A-Za-z\s'’-]/g," ").replace(/\s+/g," ").trim())
    .filter(x=>x.length>=3&&x.length<=70&&/^[A-Za-z][A-Za-z\s'’-]+$/.test(x)).slice(0,14);
  const matches=[];
  for(let i=0;i<lines.length;i++){
    for(let j=i;j<Math.min(i+4,lines.length);j++){
      const phrase=lines.slice(i,j+1).join(" ").trim();
      const words=phrase.split(/\s+/).filter(Boolean);
      if(words.length<2||words.length>5||words.some(w=>w.length<2))continue;
      const normalized=words.join("").toLowerCase().replace(/[^a-z]/g,"");
      const initials=words.map(w=>w[0]).join("").toLowerCase();
      const matched=domainKeys.some(k=>k===normalized||(k===initials&&words.length>=3));
      if(!matched)continue;
      const name=words.map(w=>w.toLowerCase().replace(/^[a-z]/,c=>c.toUpperCase())).join(" ");
      matches.push({name,wordCount:words.length});
    }
  }
  matches.sort((a,b)=>b.wordCount-a.wordCount);
  return matches[0]?.name||"";
}
function companyFromPlainDomainOnly(candidate){
  const p=candidate?.parsed||{};
  if(candidate?.candidateType==="referral"||!candidate?.email)return false;
  const name=String(p.companyName||"").trim();
  const inferred=v4CompanyFromDomain(candidate.email,p.businessHomePage||"");
  return !!inferred&&(!name||structuredCompanyLineKey(name)===structuredCompanyLineKey(inferred));
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

function signatureDiscoveryText(body){
  return norm(body).split("\n").map(line=>{
    let v=String(line||"");
    v=v.replace(/\s*<\s*(?:https?:\/\/|mailto:|tel:)[^>\r\n]*>\s*/gi," ");
    v=v.replace(/\[([^\]]+)\]\((?:https?:\/\/|mailto:|tel:)[^)]+\)/gi,"$1");
    v=v.replace(/\s*\(\s*(?:https?:\/\/|mailto:|tel:)[^)\r\n]*\)\s*/gi," ");
    return v.replace(/[ \t]+$/g,"");
  }).join("\n");
}

async function scan(){try{
  status("Reading the current email chain…");$("reviewSection").hidden=true;
  const item=Office.context.mailbox.item;if(!item||item.itemType!==Office.MailboxEnums.ItemType.Message)throw new Error("Open or select an email message first.");
  const from=item.from||{},body=await readBody(item);const myEmail=(Office.context.mailbox.userProfile?.emailAddress||"").toLowerCase();

  const textBody=body.text||body.plainText||"";
  // Stage 1 inventories signature images before deciding on contact fields.
  const imageInventory=window.SignatureImages?.inventory(body.html||"",from.displayName||"",from.emailAddress||"",myEmail,{
    visibleSignatureTextFromNode,captureSignatureBlocks,norm
  });
  const textSignatureBody=signatureDiscoveryText(textBody);
  const textSignatures=structuredDiscoverSignatures(
    textSignatureBody,from.displayName||"",from.emailAddress||"",myEmail
  );

  const htmlVisible=body.html?htmlBodyToText(body.html):"";
  const htmlSignatureVisible=signatureDiscoveryText(htmlVisible);
  const htmlSignatures=htmlSignatureVisible
    ? structuredDiscoverSignatures(htmlSignatureVisible,from.displayName||"",from.emailAddress||"",myEmail)
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

  candidates=[...signatureCandidates,...referralCandidates];
  window.SignatureImages?.associate(candidates,imageInventory);
  // OCR belongs to image inventory, not to a failure of text recognition.
  if(candidates.some(c=>c._signatureImages?.length)){
    status("Inspecting images within identified signatures…");
    try{await window.SignatureImages?.recognize(item,candidates,{
      loadTesseract,attachmentImageDataUrl,sourceToDataUrl,logoCompanyFromOcrText
    })}catch(err){console.warn("Signature image inspection unavailable",err)}
  }

  candidates=candidates.map(c=>{
    if(!c?.parsed)return c;
    if(c.candidateType!=="referral"){
      // Stage 2: parse the already-frozen Stage 1 block.  Parsing may populate
      // fields, but it must not alter the preserved source signature.
      const sourceBlock=c.rawSignatureBlock||c.parsed.signature||c.text||"";
      const visibleIdentity=(c.name&&looksLikePersonName(c.name)?c.name:"")||inferPersonName(sourceBlock)||c.name||[c.parsed.givenName,c.parsed.middleName,c.parsed.surname].filter(Boolean).join(" ");
      const reparsed=v4ContactFromSignature(visibleIdentity,c.email||c.parsed.email||"",sourceBlock);
      reparsed.signature=sourceBlock;
      reparsed.personalNotes=compactSignatureNotes(sourceBlock);
      reparsed._debug={...(c.parsed._debug||{}),...(reparsed._debug||{}),source:"Stage 2 parse of immutable signature block"};
      c.name=visibleIdentity||c.name;
      c.text=sourceBlock;
      c.parsed=reparsed;
    }
    if(c.candidateType==="referral")c.parsed.personalNotes=compactSignatureNotes(c.parsed.signature||"");
    return c
  });

  window.SignatureImages?.apply(candidates,{
    companyFromPlainDomainOnly,structuredCompanyLineKey
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
function imageEvidencePanel(p){
  const e=p?._imageEvidence;
  if(!e?.count)return "";
  const message=e.status==="company-mismatch"
   ? "The logo suggests "+e.company+", but the company field differs. Please review."
   :e.company
     ? "Company identified in signature logo: "+e.company
     :e.status==="examined"
       ? "Signature image examined, but company could not be verified. Please check."
       : "Signature image found but could not be read. Please verify the company.";
  return `<div class="muted">${html(message)}</div>`;
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

    el.innerHTML=`${badge}<div class="candidate-name" style="margin-top:6px">${html([p.givenName,p.middleName,p.surname].filter(Boolean).join(" ")||it.name)}</div><div class="muted">${html(p.email||it.email||"")}</div><div class="signature-sticky"><div class="sig-title">Original signature block from email — compacted in Notes</div><pre>${html(p.signature||"No signature block confidently found")}</pre></div>${imageEvidencePanel(p)}${unaccountedSignaturePanel(p)}${diagnosticPanel(p)}${editableFields(p,idx,null)}${actionArea}`;
    box.appendChild(el);
  });
  $("reviewSection").hidden=false;
  box.querySelectorAll("button[data-action]").forEach(b=>b.addEventListener("click",handleAction));
  box.querySelectorAll("button[data-notes-action]").forEach(b=>b.addEventListener("click",e=>{
    const idx=Number(e.currentTarget.dataset.index);
    const ta=document.querySelector(`textarea[data-review="${idx}"][data-field="personalNotes"]`);
    if(!ta)return;
    if(e.currentTarget.dataset.notesAction==="clear")ta.value="";
    else ta.value=(window.__reviewItems[idx]?.parsed?.personalNotes||"");
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



