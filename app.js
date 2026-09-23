import { createNestablePublicClientApplication, InteractionRequiredAuthError } from "https://cdn.jsdelivr.net/npm/@azure/msal-browser@5.1.0/+esm";

const GRAPH_SCOPES=["Contacts.ReadWrite"];
const TITLE_WORDS=["project executive","senior project manager","project manager","assistant project manager","project engineer","project coordinator","construction manager","assistant general manager","general manager","superintendent","estimator","vice president","president","principal","partner","associate","director","manager","architect","engineer","designer","consultant","owner","coordinator"];
const COMPANY_WORDS=[" llc"," l.l.c"," inc"," corp"," company"," co."," construction"," builders"," building"," architecture"," architects"," engineering"," engineers"," associates"," group"," studio"," mechanical"," electric"," electrical"," plumbing"," design"," contractors"," contractor"," concrete"," masonry"," garage"," workshop"," services"," solutions"," systems"," enterprises"," partners"];
const CREDENTIALS=new Set(["AIA","PE","P.E.","RA","R.A.","LEED","PMP","NCARB","FAIA","SE","S.E."]);
const FIELD_META={
  givenName:"First name",middleName:"Middle name",surname:"Last name",companyName:"Company",jobTitle:"Job title",email:"Email",
  businessPhone:"Business / direct",mobilePhone:"Mobile",businessFax:"Fax",businessHomePage:"Website",street:"Street",city:"City",state:"State",postalCode:"ZIP",countryOrRegion:"Country",personalNotes:"Notes"
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
  // First pass: honor explicit labels. This is especially important for OCR, where
  // a business-card image may contain Office, Direct, Mobile/Cell and Fax on one line.
  const labeled={office:"",direct:"",mobile:"",fax:""};
  const labelText=norm(sig).replace(/©/g,"O").replace(/®/g,"O");
  const labelRx=/(?:^|[|•;\s])\b(office|business|phone|tel|telephone|direct|mobile|cell|fax|o|d|m|c|f)\b\s*[:.\-]?\s*((?:\+?1[\s.\-]?)?(?:\(?\d{3}\)?[\s.\-]*)\d{3}[\s.\-]*\d{4})(?:\s*(?:x|ext\.?|extension)\s*[:.\-]?\s*(\d+))?/gi;
  for(const line of labelText.split("\n")){
    labelRx.lastIndex=0; let m;
    while((m=labelRx.exec(line))){
      const lab=m[1].toLowerCase();
      const val=phoneForOutlook(m[2]+(m[3]?` x${m[3]}`:""));
      if(!val)continue;
      if(/^(office|business|phone|tel|telephone|o)$/.test(lab) && !labeled.office)labeled.office=val;
      else if(/^(direct|d)$/.test(lab) && !labeled.direct)labeled.direct=val;
      else if(/^(mobile|cell|m|c)$/.test(lab) && !labeled.mobile)labeled.mobile=val;
      else if(/^(fax|f)$/.test(lab) && !labeled.fax)labeled.fax=val;
    }
  }

  const c=phoneTokens(labelText);
  const has=(x,rx)=>rx.test((x.before+" "+x.after).trim());
  const mobileRx=/\b(mobile|cell|cellular)\b|(?:^|[|•;\s])\(?\s*(?:m|c)\s*\)?\s*[:.\-]?\s*(?:$|[|•;])/i;
  const faxRx=/\bfax\b|(?:^|[|•;\s])\(?\s*f\s*\)?\s*[:.\-]?\s*(?:$|[|•;])/i;
  const directRx=/\bdirect\b|(?:^|[|•;\s])\(?\s*d\s*\)?\s*[:.\-]?\s*(?:$|[|•;])/i;
  const officeRx=/\b(office|business|phone|tel|telephone)\b|(?:^|[|•;\s])\(?\s*(?:o|p|t)\s*\)?\s*[:.\-]?\s*(?:$|[|•;])/i;
  const mobile=labeled.mobile||((c.find(x=>has(x,mobileRx))||{}).value||"");
  const fax=labeled.fax||((c.find(x=>has(x,faxRx))||{}).value||"");
  const direct=labeled.direct||((c.find(x=>has(x,directRx))||{}).value||"");
  const office=labeled.office||((c.find(x=>has(x,officeRx))||{}).value||"");
  let business=direct||office;
  if(!business){const u=c.find(x=>phoneForOutlook(x.value)!==phoneForOutlook(mobile)&&phoneForOutlook(x.value)!==phoneForOutlook(fax));business=u?u.value:""}
  const result={businessPhone:phoneForOutlook(business),mobilePhone:phoneForOutlook(mobile),businessFax:phoneForOutlook(fax)};
  // Never duplicate the same OCR number into Office and Mobile. If OCR confused an
  // O label with C/M, the explicit business label is the safer interpretation.
  if(result.businessPhone && result.mobilePhone && result.businessPhone===result.mobilePhone)result.mobilePhone="";
  if(result.businessPhone && result.businessFax && result.businessPhone===result.businessFax && !labeled.fax)result.businessFax="";
  return result;
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
function title(sig){
  const lines=norm(sig).split("\n").map(x=>x.trim()).filter(Boolean);
  for(let i=0;i<lines.length;i++){
    const line=lines[i],low=line.toLowerCase();
    if(line.length>100||!TITLE_WORDS.some(t=>low.includes(t))||COMPANY_WORDS.some(w=>(" "+low).includes(w)))continue;
    let out=line;
    const person=personNameFromLine(line);
    if(person&&out.toLowerCase().startsWith(person.toLowerCase()))out=out.slice(person.length).replace(/^[-|•·—–\s]+/,"");
    out=out.replace(/^[-|•·—–\s]+|[-|•·—–\s]+$/g,"");
    const next=lines[i+1]||"",nextLow=next.toLowerCase();
    if(next&&/^\s*[|•·—–]/.test(next)&&TITLE_WORDS.some(t=>nextLow.includes(t))){
      const extra=next.replace(/^[-|•·—–\s]+|[-|•·—–\s]+$/g,"");
      if(extra)out=(out?out+" | ":"")+extra;
    }
    if(out)return out;
  }
  return"";
}
function looksLikePersonName(line){
  const v=(line||"").trim();if(!v||v.length<4||v.length>55)return false;
  if(cleanEmail(v)||/\d|https?:|www\.|@|\b(?:street|st\.?|road|rd\.?|ave\.?|avenue|blvd\.?|boulevard|suite|ste\.?|drive|dr\.?|lane|ln\.?|city|inc\.?|llc|corp\.?|company|garage|workshop)\b/i.test(v))return false;
  const low=v.toLowerCase();if(TITLE_WORDS.some(t=>low.includes(t))||COMPANY_WORDS.some(w=>(" "+low).includes(w)))return false;
  const words=v.replace(/[,]/g," ").split(/\s+/).filter(Boolean);if(words.length<2||words.length>5)return false;
  return words.every(w=>/^[A-Za-z][A-Za-z'.-]*$/.test(w));
}
function personNameFromLine(line){
  const raw=(line||"").trim();
  if(looksLikePersonName(raw))return raw;
  // Many signatures put the name and title on the same visual line:
  // "Linda Shin | Associate" or "Jason Ro • Principal".
  const pieces=raw.split(/\s*(?:\||•|·|—|–)\s*/).map(x=>x.trim()).filter(Boolean);
  if(pieces.length>1 && looksLikePersonName(pieces[0]))return pieces[0];
  return "";
}
function inferPersonName(sig){
  const lines=norm(sig).split("\n").map(x=>x.trim()).filter(Boolean);
  for(let i=0;i<Math.min(lines.length,12);i++){
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
  const stem=domain.split(".")[0].replace(/[-_]+/g," ").trim();
  if(!stem)return"";
  return stem.split(/\s+/).map(w=>/^\d+$/.test(w)?w:(w.charAt(0).toUpperCase()+w.slice(1))).join(" ");
}
function company(sig,name,senderEmail,site){
  const lines=norm(sig).split("\n").map(x=>x.trim()).filter(Boolean),lowName=(name||"").toLowerCase();
  // First trust explicit company-like text in the signature.
  for(const line of lines){
    const low=" "+line.toLowerCase();
    if(line.length>2&&line.length<110&&COMPANY_WORDS.some(w=>low.includes(w)))return line;
  }
  // If the company is in a logo and not exposed as text, the business domain is safer
  // than guessing from ordinary email prose.
  const domainCompany=companyFromDomain(senderEmail,site);
  if(domainCompany)return domainCompany;
  // Last-resort generic text fallback, deliberately conservative.
  for(const line of lines.slice(0,10)){
    const low=line.toLowerCase();
    if(low===lowName||looksLikePersonName(line)||low.includes("@")||/\d{3}[\s.\-]\d{3}/.test(line)||/^https?:|^www\./i.test(line))continue;
    if(TITLE_WORDS.some(t=>low.includes(t)))continue;
    if(/^from:|^sent:|^to:|^cc:|^subject:/i.test(line))continue;
    if(/\b(?:p\.?o\.?\s*box|box|street|st\.?|road|rd\.?|avenue|ave\.?|boulevard|blvd\.?|drive|dr\.?|lane|ln\.?|loop|suite|ste\.?|bldg|building)\b/i.test(line)||/\b[A-Z]{2}\s+\d{5}(?:-\d{4})?\b/.test(line))continue;
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
function parseContact(senderName,senderEmail,segmentText){
  const rawSig=isolateSignature(segmentText,senderEmail);
  const sig=cleanSignatureText(rawSig);
  const inferred=inferPersonName(sig);
  const senderLooksHuman=looksLikePersonName(senderName||"");
  const resolvedName=senderLooksHuman?senderName:(inferred||senderName||"");
  const site=website(sig,senderEmail);return Object.assign({},nameParts(resolvedName),{companyName:company(sig,resolvedName,senderEmail,site),jobTitle:title(sig),email:senderEmail||""},phones(sig),{businessHomePage:site},address(sig),{signature:sig,personalNotes:sig})
}

function parseContactFromSignature(senderName,senderEmail,signatureText){
  const sig=cleanSignatureText(signatureText);
  const inferred=inferPersonName(sig);
  const senderLooksHuman=looksLikePersonName(senderName||"");
  const resolvedName=senderLooksHuman?senderName:(inferred||senderName||"");
  const site=website(sig,senderEmail);
  return Object.assign({},nameParts(resolvedName),{companyName:company(sig,resolvedName,senderEmail,site),jobTitle:title(sig),email:senderEmail||""},phones(sig),{businessHomePage:site},address(sig),{signature:sig,personalNotes:sig});
}

function sameEmail(a,b){return cleanEmail(a)&&cleanEmail(a)===cleanEmail(b)}
function cleanSignatureText(sig){
  return norm(sig).split("\n").map(x=>x.trim()).filter(Boolean).filter(line=>{
    if(/^(?:https?:\/\/)?[^\s]+\.(?:png|jpe?g|gif|svg|webp|bmp|ico)(?:[?#].*)?$/i.test(line))return false;
    if(/^cid:|^data:image/i.test(line))return false;
    if(/(?:google\.[^/]+\/maps|maps\.google\.|maps\.apple\.|bing\.com\/maps|goo\.gl\/maps|safelinks\.protection\.outlook\.com|urldefense\.proofpoint\.com\/v2\/url\?)/i.test(line))return false;
    return true;
  }).join("\n")
}
function splitMessage(body,currentName,currentEmail){const text=norm(body);const lines=text.split("\n");const headers=[];for(let i=0;i<lines.length;i++){const m=lines[i].match(/^\s*From:\s*(.+)$/i);if(!m)continue;const val=m[1].trim();const email=cleanEmail(val);let name=cleanName(val.replace(email,""));if(!name&&email)name=email.split("@")[0];headers.push({i,name,email})}
  const firstHeader=headers.length?headers[0].i:lines.length;const out=[];out.push({name:currentName||"Current sender",email:(currentEmail||"").toLowerCase(),text:lines.slice(0,firstHeader).join("\n")});
  for(let h=0;h<headers.length;h++){const a=headers[h],b=headers[h+1]?headers[h+1].i:lines.length;out.push({name:a.name,email:a.email,text:lines.slice(a.i+1,b).join("\n")})}
  // A person may appear more than once in a chain. Keep the segment that contains the
  // strongest sender-specific signature rather than simply keeping the newest occurrence.
  const bestByKey=new Map();
  for(const seg of out){
    const key=seg.email||seg.name.toLowerCase();if(!key)continue;
    const sig=isolateSignature(seg.text,seg.email);
    let score=norm(sig).split("\n").reduce((t,l)=>t+signatureScore(l),0);
    if(seg.email && norm(sig).toLowerCase().includes(seg.email.toLowerCase()))score+=6;
    const prev=bestByKey.get(key);
    if(!prev||score>prev.score)bestByKey.set(key,{score,seg});
  }
  return [...bestByKey.values()].map(x=>x.seg)}



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
  const byEmail=new Map();
  // v2.5.7: treat the current Outlook sender separately from generic chain detection.
  // Outlook already gives us an authoritative From name + email. If that sender's visible
  // name appears in the body, build the signature directly from that point forward even
  // when the signature itself contains no printed email address or mailto link.
  if(currentEmail&&!sameEmail(currentEmail,myEmail)&&looksLikePersonName(currentName||"")){
    const directSig=currentSenderSignatureFromBody(body,currentName);
    if(directSig){
      const parsed=parseContactFromSignature(currentName,currentEmail,directSig);
      const np=nameParts(currentName);
      parsed.givenName=np.givenName;parsed.middleName=np.middleName;parsed.surname=np.surname;
      parsed.email=currentEmail.toLowerCase();
      parsed.signature=directSig;parsed.personalNotes=directSig;
      parsed._debug={source:"Current sender — Outlook From header + visible name anchor",htmlSignature:"",textSignature:directSig};
      const directScore=norm(directSig).split("\n").reduce((t,l)=>t+signatureScore(l),0)+10;
      byEmail.set(currentEmail.toLowerCase(),{score:directScore,name:currentName,email:currentEmail.toLowerCase(),text:directSig,parsed});
    }
  }
  for(let i=0;i<lines.length;i++){
    if(isHeader(lines[i]))continue;
    const email=cleanEmail(lines[i]);
    if(!email||sameEmail(email,myEmail))continue;
    // Build a tight block around the visible email address. This is much more reliable in
    // forwarded/replied chains than depending on Outlook's reconstructed From: headers.
    let lo=i,hi=i,blankBudget=1;
    for(let j=i-1;j>=0 && i-j<=10;j--){
      const v=lines[j];
      if(isHeader(v)||isDisclaimer(v))break;
      if(!v){if(blankBudget--<=0)break;continue}
      lo=j;
    }
    blankBudget=1;
    for(let j=i+1;j<lines.length && j-i<=8;j++){
      const v=lines[j];
      if(isHeader(v)||isDisclaimer(v))break;
      if(!v){if(blankBudget--<=0)break;continue}
      hi=j;
    }
    let chunk=lines.slice(lo,hi+1).filter(Boolean);
    // Prefer the nearest plausible person's name above the anchored email. This strips
    // ordinary message prose even when there is no blank line before the signature.
    let emailLocal=chunk.findIndex(l=>cleanEmail(l)===email);
    if(emailLocal<0)emailLocal=chunk.length-1;
    let nameLocal=-1;
    for(let j=emailLocal-1;j>=Math.max(0,emailLocal-8);j--){if(looksLikePersonName(chunk[j])){nameLocal=j;break}}
    if(nameLocal>=0)chunk=chunk.slice(nameLocal);
    else{
      while(chunk.length>5 && signatureScore(chunk[0])===0)chunk.shift();
    }
    while(chunk.length>5 && signatureScore(chunk[chunk.length-1])===0 && signatureScore(chunk[chunk.length-2])===0)chunk.pop();
    const sig=cleanSignatureText(chunk.join("\n"));
    const score=norm(sig).split("\n").reduce((t,l)=>t+signatureScore(l),0);
    if(score<5)continue;
    const inferred=inferAnchoredName(sig,email);
    const name=(sameEmail(email,currentEmail)&&looksLikePersonName(currentName||""))?currentName:(inferred||email.split("@")[0]);
    const parsed=parseContact(name,email,sig);
    // parseContact will isolate again; preserve the exact anchored signature we just selected.
    parsed.signature=sig; parsed.personalNotes=sig;
    const existing=byEmail.get(email);
    if(!existing||score>existing.score)byEmail.set(email,{score,name,email,text:sig,parsed});
  }
  // Current-sender fallback when the signature itself does not print an email address.
  // Outlook's From name/address are authoritative for the current message. Do NOT require
  // the signature parser to rediscover the sender's name exactly; visually similar HTML
  // signatures can put "Name | Title" in one cell and would otherwise be rejected.
  if(currentEmail&&!sameEmail(currentEmail,myEmail)&&!byEmail.has(currentEmail.toLowerCase())){
    const legacy=splitMessage(body,currentName,currentEmail).find(x=>sameEmail(x.email,currentEmail));
    if(legacy){
      let rawSig=isolateSignature(legacy.text,legacy.email);
      // If Outlook gives us a trustworthy current-sender name, use the matching visible
      // name line as the beginning of the signature. This preserves lines such as
      // "Linda Shin | Associate" that a score-only fallback can otherwise drop.
      const rawLines=norm(legacy.text).split("\n").map(x=>x.trim());
      const wanted=String(currentName||legacy.name||"").toLowerCase().replace(/[^a-z0-9]+/g," ").trim();
      let nameAt=-1;
      for(let j=0;j<rawLines.length;j++){
        const pn=personNameFromLine(rawLines[j]);
        const key=String(pn||"").toLowerCase().replace(/[^a-z0-9]+/g," ").trim();
        if(wanted&&key===wanted){nameAt=j;break}
      }
      if(nameAt>=0){
        const kept=[];
        for(let j=nameAt;j<rawLines.length&&kept.length<16;j++){
          const v=rawLines[j];
          if(j>nameAt && /^\s*(from|sent|to|cc|bcc|subject):\s*/i.test(v))break;
          if(j>nameAt && /confidential|privileged|intended recipient|virus|disclaimer/i.test(v))break;
          if(v)kept.push(v);
        }
        if(kept.length>=3)rawSig=kept.join("\n");
      }
      const sig=cleanSignatureText(rawSig);
      const parsed=parseContactFromSignature(currentName||legacy.name,legacy.email,sig);
      const sigScore=norm(sig).split("\n").reduce((t,l)=>t+signatureScore(l),0);
      const ph=phones(sig),addr=address(sig),site=website(sig,legacy.email);
      const evidence=[
        !!(ph.businessPhone||ph.mobilePhone),
        !!(addr.state&&addr.postalCode),
        !!site,
        !!company(sig,currentName||legacy.name,legacy.email,site),
        !!title(sig)
      ].filter(Boolean).length;
      // Require a genuinely signature-like block, but trust Outlook for who the current
      // sender is. This keeps the safety check without rejecting signatures lacking email.
      if(sig && looksLikePersonName(currentName||legacy.name||"") && sigScore>=5 && evidence>=2){
        const np=nameParts(currentName||legacy.name);
        parsed.givenName=np.givenName; parsed.middleName=np.middleName; parsed.surname=np.surname;
        parsed.email=legacy.email;
        parsed.signature=sig;
        parsed.personalNotes=sig;
        parsed._debug={source:"Current sender fallback (Outlook From header)",htmlSignature:"",textSignature:sig};
        byEmail.set(currentEmail.toLowerCase(),{score:sigScore,name:currentName||legacy.name,email:legacy.email,text:sig,parsed});
      }
    }
  }
  return [...byEmail.values()].map(x=>({name:x.name,email:x.email,text:x.text,parsed:x.parsed}));
}

function renderCandidates(){const box=$("candidates");box.innerHTML="";candidates.forEach((c,i)=>{const el=document.createElement("div");el.className="candidate";el.innerHTML=`<div class="candidate-head"><input type="checkbox" data-candidate="${i}" checked><div><div class="candidate-name">${html(c.name||"Unknown sender")}</div><div class="muted">${html(c.email||"No email found")}</div></div></div><div class="preview">${html(c.parsed.signature||"No signature block confidently found")}</div>`;box.appendChild(el)});$("candidateSection").hidden=false}
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
    ["Company",parsed.companyName],["Job title",parsed.jobTitle],["Business / direct",parsed.businessPhone],["Mobile",parsed.mobilePhone],["Fax",parsed.businessFax],
    ["Street",parsed.street],["City",parsed.city],["State",parsed.state],["ZIP",parsed.postalCode],["Country inferred",parsed.countryOrRegion]
  ];
  return vals.map(([l,v])=>`<div class="diag-row"><b>${html(l)}</b><span>${html(v||"(not detected)")}</span></div>`).join("");
}
function diagnosticPanel(parsed){
  const d=parsed._debug||{};
  return `<details class="diagnostic" open><summary>Parser diagnostic — what the add-in actually detected</summary><div class="diag-source">Source used: ${html(d.source||"email signature")}</div>${diagnosticRows(parsed)}<div class="diag-block"><b>Signature text used for Notes</b><pre>${html(parsed.signature||"(none)")}</pre></div>${d.htmlSignature&&d.textSignature&&d.htmlSignature!==d.textSignature?`<div class="diag-block"><b>HTML signature candidate</b><pre>${html(d.htmlSignature)}</pre><b>Plain-text signature candidate</b><pre>${html(d.textSignature)}</pre></div>`:""}</details>`;
}



// v2.6.1 — OCR fallback and image-signature repair.
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
  Object.assign(out,phones(text));
  out=repairOcrStateAndCountry(out,text);
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
    for(let i=0;i<sources.length;i++){
      status(`Reading signature image ${i+1} of ${sources.length}…`);
      try{
        const r=await worker.recognize(sources[i]);
        const text=cleanSignatureText(r?.data?.text||"");
        const score=ocrSignatureScore(text,currentName,currentEmail);
        if(!best||score>best.score)best={score,text};
      }catch(_){ }
    }
  }finally{try{if(worker)await worker.terminate()}catch(_){ }}
  if(!best || best.score<8 || !best.text)return null;
  let parsed=parseContactFromSignature(currentName||"Current sender",String(currentEmail||"").toLowerCase(),best.text);
  parsed=repairOcrParsed(parsed,best.text,currentName,currentEmail);
  parsed.signature=best.text;
  parsed.personalNotes=best.text;
  parsed._debug={source:"Image OCR fallback",htmlSignature:"",textSignature:best.text,ocrScore:best.score};
  return {name:currentName||[parsed.givenName,parsed.surname].filter(Boolean).join(" "),email:parsed.email,text:best.text,parsed};
}

function readBody(item){
  return new Promise((resolve,reject)=>{
    item.body.getAsync(Office.CoercionType.Html,r=>{
      if(r.status===Office.AsyncResultStatus.Succeeded){
        const html=r.value||"";
        const text=htmlBodyToText(html);
        if(text.trim())return resolve({html,text});
      }
      item.body.getAsync(Office.CoercionType.Text,r2=>r2.status===Office.AsyncResultStatus.Succeeded?resolve({html:"",text:r2.value||""}):reject(new Error(r2.error?.message||"Unable to read the email body.")))
    })
  })
}
async function scan(){try{
  status("Reading the current email chain…");$("reviewSection").hidden=true;
  const item=Office.context.mailbox.item;if(!item||item.itemType!==Office.MailboxEnums.ItemType.Message)throw new Error("Open or select an email message first.");
  const from=item.from||{},body=await readBody(item);
  const myEmail=(Office.context.mailbox.userProfile?.emailAddress||"").toLowerCase();
  const htmlCandidates=body.html?htmlSignatureCandidates(body.html,from.displayName||"",from.emailAddress||"",myEmail):[];
  const textCandidates=anchoredSignatureCandidates(body.text,from.displayName||"",from.emailAddress||"",myEmail);
  // Compare the HTML and plain-text readings for the same person instead of blindly
  // letting HTML win. Outlook often puts the mailto link in a small nested table cell while
  // the plain-text rendering contains the complete visible address and phones.
  const merged=new Map();
  const htmlMap=new Map(htmlCandidates.map(c=>[(c.email||"").toLowerCase(),c]));
  const textMap=new Map(textCandidates.map(c=>[(c.email||"").toLowerCase(),c]));
  const keys=new Set([...htmlMap.keys(),...textMap.keys()]);
  for(const key of keys){
    const h=htmlMap.get(key),t=textMap.get(key);
    if(h&&t){
      const parsed=mergeParsed(h.parsed,t.parsed,"HTML + plain-text merged");
      merged.set(key,{...h,parsed,text:parsed.signature});
    }else if(h){h.parsed._debug={source:"HTML only",htmlSignature:h.parsed.signature||"",textSignature:""};merged.set(key,h)}
    else if(t){t.parsed._debug={source:"Plain text only",htmlSignature:"",textSignature:t.parsed.signature||""};merged.set(key,t)}
  }
  candidates=[...merged.values()];
  // Image-only signatures can coexist with quoted text from older messages. In that case
  // the generic text scanner may create a false current-sender candidate from somebody
  // else's quoted signature. If the current sender is not actually visible in the selected
  // text signature, OCR the likely signature image and prefer that result for this sender.
  try{
    const currentEmail=String(from.emailAddress||"").toLowerCase();
    const currentName=cleanName(from.displayName||"");
    const current=candidates.find(c=>sameEmail(c.email,currentEmail));
    const sig=String(current?.parsed?.signature||"");
    const inferred=cleanName(inferPersonName(sig)).toLowerCase();
    const hasIdentity=!!current && ((currentEmail&&sig.toLowerCase().includes(currentEmail)) || (currentName&&inferred===currentName.toLowerCase()));
    if(!hasIdentity){
      const ocr=await imageSignatureCandidate(item,body.html||"",from.displayName||"",from.emailAddress||"",myEmail);
      if(ocr){
        candidates=candidates.filter(c=>!sameEmail(c.email,currentEmail));
        candidates.unshift(ocr);
      }
    }
  }catch(ocrErr){
    console.warn("Image signature OCR fallback failed",ocrErr);
  }
  renderCandidates();
  status(`Found ${candidates.length} possible contact${candidates.length===1?"":"s"}. Your own messages/signature are ignored. Select the people you want to process.`,candidates.length?"ok":"error")
}catch(e){status(e.message||String(e),"error")}}

function clientId(){return localStorage.getItem("ccfe_client_id")||""}
function showAuthSetup(){const id=clientId();$("authSetup").hidden=!!id;$("clientId").value=id}
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
    const initial=(p[k]||(!p[k]&&existing?existing[k]:"")||"");
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
function renderReviews(items){const box=$("reviews");box.innerHTML="";items.forEach((it,idx)=>{const p=it.parsed,m=it.match,existing=m?existingFlat(m.contact):null,diffs=existing?fieldDiffs(p,existing):[];const el=document.createElement("div");el.className="review";el.dataset.reviewCard=idx;const badge=m?`<span class="badge existing">Existing contact — ${html(m.confidence)}</span>`:`<span class="badge">New contact</span>`;const existingSummary=m?`<div class="existing-panel"><b>Existing Outlook contact data</b>${existing._addressType?`<div class="muted">Address currently stored by Outlook as ${html(existing._addressType)} Address.</div>`:""}<div class="compare-table">${comparisonRows(p,existing)}</div></div>`:"";el.innerHTML=`${badge}<div class="candidate-name" style="margin-top:6px">${html([p.givenName,p.middleName,p.surname].filter(Boolean).join(" ")||it.name)}</div><div class="muted">${html(p.email||it.email||"")}</div><div class="signature-sticky"><div class="sig-title">Signature block from email</div><pre>${html(p.signature||"No signature block confidently found")}</pre></div>${diagnosticPanel(p)}${existingSummary}${editableFields(p,idx,existing)}${m?`<div class="diffs"><b>Fields available to update</b>${diffs.length?diffs.map(d=>`<label class="diff"><input type="checkbox" data-diff="${idx}" data-key="${d.key}" ${d.defaultChecked?"checked":""}><span>${html(d.label)}</span><span class="vals"><span class="old">Existing: ${html(d.old||"(blank)")}</span><br><span class="new">From email: ${html(d.new)}</span></span></label>`).join(""):"<div class=\"muted\">The email did not provide any new or different nonblank values. Existing Outlook data is shown above and will not be erased.</div>"}</div><div class="toolbar"><button class="btn primary" data-action="update" data-index="${idx}" ${diffs.length?"":"disabled"}>Update Existing Contact</button><button class="btn" data-action="skip" data-index="${idx}">Skip</button></div>`:`<div class="toolbar"><button class="btn primary" data-action="create" data-index="${idx}">Create Contact</button><button class="btn" data-action="skip" data-index="${idx}">Skip</button></div>`}`;box.appendChild(el)});$("reviewSection").hidden=false;box.querySelectorAll("button[data-action]").forEach(b=>b.addEventListener("click",handleAction));box.querySelectorAll("button[data-notes-action]").forEach(b=>b.addEventListener("click",e=>{const idx=Number(e.currentTarget.dataset.index);const ta=document.querySelector(`textarea[data-review="${idx}"][data-field="personalNotes"]`);if(!ta)return;if(e.currentTarget.dataset.notesAction==="clear")ta.value="";else ta.value=(window.__reviewItems[idx]?.parsed?.signature||"");ta.focus();}));}
function currentParsed(idx){const d={};document.querySelectorAll(`[data-review="${idx}"][data-field]`).forEach(i=>d[i.dataset.field]=i.value.trim());d.businessPhone=phoneForOutlook(d.businessPhone);d.mobilePhone=phoneForOutlook(d.mobilePhone);d.businessFax=phoneForOutlook(d.businessFax);return d}
function finishCard(idx,label){const c=document.querySelector(`[data-review-card="${idx}"]`);c.classList.add("done");c.querySelectorAll("button").forEach(b=>b.disabled=true);const span=document.createElement("span");span.className="badge done";span.textContent=label;c.prepend(span)}
async function handleAction(ev){const action=ev.currentTarget.dataset.action,idx=Number(ev.currentTarget.dataset.index),item=window.__reviewItems[idx];if(action==="skip"){finishCard(idx,"Skipped");return}try{ev.currentTarget.disabled=true;status(action==="create"?"Creating Outlook contact…":"Updating Outlook contact…");const p=currentParsed(idx);if(action==="create"){await graph("/me/contacts",{method:"POST",body:JSON.stringify(graphPayload(p))});finishCard(idx,"Created");status("Contact created in Outlook Contacts.","ok")}else{const selected=new Set([...document.querySelectorAll(`input[data-diff="${idx}"]:checked`)].map(x=>x.dataset.key));if(!selected.size)throw new Error("Check at least one field to update.");await graph(`/me/contacts/${encodeURIComponent(item.match.contact.id)}`,{method:"PATCH",body:JSON.stringify(graphPayload(p,selected,existingFlat(item.match.contact)))});finishCard(idx,"Updated");status("Existing Outlook contact updated.","ok")}}catch(e){ev.currentTarget.disabled=false;status(e.message||String(e),"error")}}
async function compareSelected(){try{const selected=[...document.querySelectorAll("input[data-candidate]:checked")].map(x=>candidates[Number(x.dataset.candidate)]);if(!selected.length)throw new Error("Select at least one person first.");if(!clientId()){showAuthSetup();throw new Error("Complete the one-time Microsoft Contacts setup first.")}status("Signing in to Microsoft and checking Outlook Contacts…");graphContacts=await loadContacts();const items=selected.map(c=>({...c,match:matchContact(c.parsed)}));window.__reviewItems=items;renderReviews(items);status(`Compared ${items.length} selected contact${items.length===1?"":"s"} with ${graphContacts.length} Outlook contact${graphContacts.length===1?"":"s"}.`,"ok")}catch(e){status(e.message||String(e),"error")}}

Office.onReady(async info=>{if(info.host!==Office.HostType.Outlook){status("This page must be opened from the Outlook add-in.","error");return}showAuthSetup();$("saveClientId").addEventListener("click",()=>{const id=$("clientId").value.trim();if(!/^[0-9a-f-]{36}$/i.test(id)){status("That does not look like a Microsoft Application (client) ID.","error");return}localStorage.setItem("ccfe_client_id",id);msalInstance=null;showAuthSetup();status("Client ID saved. You can now compare contacts.","ok")});$("selectAll").addEventListener("click",()=>document.querySelectorAll("input[data-candidate]").forEach(x=>x.checked=true));$("selectNone").addEventListener("click",()=>document.querySelectorAll("input[data-candidate]").forEach(x=>x.checked=false));$("scanAgain").addEventListener("click",scan);$("compareSelected").addEventListener("click",compareSelected);await scan()});
