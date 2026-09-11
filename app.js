import { createNestablePublicClientApplication, InteractionRequiredAuthError } from "https://cdn.jsdelivr.net/npm/@azure/msal-browser@5.1.0/+esm";

const GRAPH_SCOPES=["Contacts.ReadWrite"];
const TITLE_WORDS=["project executive","senior project manager","project manager","assistant project manager","project engineer","project coordinator","construction manager","assistant general manager","general manager","superintendent","estimator","vice president","president","principal","partner","director","manager","architect","engineer","designer","consultant","owner","coordinator"];
const COMPANY_WORDS=[" llc"," l.l.c"," inc"," corp"," company"," co."," construction"," builders"," building"," architecture"," architects"," engineering"," engineers"," associates"," group"," studio"," mechanical"," electric"," electrical"," plumbing"," design"," contractors"," contractor"," garage"," workshop"," services"," solutions"," systems"," enterprises"," partners"];
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
  const extMatch=raw.match(/(?:^|\s)(?:x|ext\.?|extension)\s*(\d+)\s*$/i);
  const ext=extMatch?extMatch[1]:"";
  const main=(extMatch?raw.slice(0,extMatch.index):raw).replace(/\D/g,"");
  let d=main;if(d.length===11&&d.startsWith("1"))d=d.slice(1);
  return ext?`${d}x${ext}`:d;
}
function phoneTokens(sig){
  const re=/(?:\+?1[\s.\-]?)?(?:\(?\d{3}\)?[\s.\-]?)\d{3}[\s.\-]\d{4}(?:\s*(?:x|ext\.?|extension)\s*\d+)?/gi,c=[];
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
  const c=phoneTokens(sig);
  const has=(x,rx)=>rx.test((x.before+" "+x.after).trim());
  const mobileRx=/\b(mobile|cell|cellular)\b|(?:^|[|•;\s])\(?\s*(?:m|c)\s*\)?\s*[:.\-]?\s*(?:$|[|•;])/i;
  const faxRx=/\bfax\b|(?:^|[|•;\s])\(?\s*f\s*\)?\s*[:.\-]?\s*(?:$|[|•;])/i;
  const directRx=/\bdirect\b|(?:^|[|•;\s])\(?\s*d\s*\)?\s*[:.\-]?\s*(?:$|[|•;])/i;
  const officeRx=/\b(office|business|phone|tel|telephone)\b|(?:^|[|•;\s])\(?\s*(?:o|p|t)\s*\)?\s*[:.\-]?\s*(?:$|[|•;])/i;
  const mobile=(c.find(x=>has(x,mobileRx))||{}).value||"";
  const fax=(c.find(x=>has(x,faxRx))||{}).value||"";
  const direct=(c.find(x=>has(x,directRx))||{}).value||"";
  const office=(c.find(x=>has(x,officeRx))||{}).value||"";
  let business=direct||office;
  if(!business){const u=c.find(x=>x.value!==mobile&&x.value!==fax);business=u?u.value:""}
  return{businessPhone:phoneForOutlook(business),mobilePhone:phoneForOutlook(mobile),businessFax:phoneForOutlook(fax)}
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
function title(sig){for(const line of norm(sig).split("\n").map(x=>x.trim()).filter(Boolean)){const low=line.toLowerCase();if(line.length<=100&&TITLE_WORDS.some(t=>low.includes(t))&&!COMPANY_WORDS.some(w=>(" "+low).includes(w)))return line.replace(/^[-|•\s]+|[-|•\s]+$/g,"")}return""}
function looksLikePersonName(line){
  const v=(line||"").trim();if(!v||v.length<4||v.length>55)return false;
  if(cleanEmail(v)||/\d|https?:|www\.|@|\b(?:street|st\.?|road|rd\.?|ave\.?|avenue|blvd\.?|boulevard|suite|ste\.?|drive|dr\.?|lane|ln\.?|city|inc\.?|llc|corp\.?|company|garage|workshop)\b/i.test(v))return false;
  const low=v.toLowerCase();if(TITLE_WORDS.some(t=>low.includes(t))||COMPANY_WORDS.some(w=>(" "+low).includes(w)))return false;
  const words=v.replace(/[,]/g," ").split(/\s+/).filter(Boolean);if(words.length<2||words.length>5)return false;
  return words.every(w=>/^[A-Za-z][A-Za-z'.-]*$/.test(w));
}
function inferPersonName(sig){
  const lines=norm(sig).split("\n").map(x=>x.trim()).filter(Boolean);
  for(let i=0;i<Math.min(lines.length,12);i++){
    if(looksLikePersonName(lines[i])){
      const next=(lines[i+1]||"").toLowerCase();
      if(!next||TITLE_WORDS.some(t=>next.includes(t))||COMPANY_WORDS.some(w=>(" "+next).includes(w)))return lines[i];
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
  const lines=norm(sig).split("\n").map(x=>x.trim()).filter(Boolean);
  const streetWord=/\b(street|st\.?|road|rd\.?|avenue|ave\.?|boulevard|blvd\.?|drive|dr\.?|lane|ln\.?|way|court|ct\.?|highway|hwy\.?|parkway|pkwy\.?|place|pl\.?|trail|trl\.?|circle|cir\.?|square|sq\.?|loop|terrace|ter\.?|suite|ste\.?|floor|fl\.?|bldg|building|p\.?o\.?\s*box)\b/i;
  const cityStateZip=/^([A-Za-z .'-]+?),?\s+([A-Z]{2})\s+(\d{5}(?:-\d{4})?)$/;
  for(let i=0;i<lines.length;i++){
    const line=lines[i];
    const parts=line.split(/\s*[•|]\s*/).map(x=>x.trim()).filter(Boolean);
    if(parts.length>=2){
      const last=parts[parts.length-1].match(cityStateZip);
      if(last){
        const streetParts=parts.slice(0,-1).filter(x=>/\d/.test(x)||streetWord.test(x));
        return{street:streetParts.join("\n"),city:last[1].trim(),state:last[2],postalCode:last[3],countryOrRegion:usCountry(last[2])};
      }
    }
    let m=line.match(/^(.*\d.*?),\s*([A-Za-z .'-]+),\s*([A-Z]{2})\s+(\d{5}(?:-\d{4})?)$/);
    if(m&&streetWord.test(m[1]))return{street:m[1].trim(),city:m[2].trim(),state:m[3],postalCode:m[4],countryOrRegion:usCountry(m[3])};
    m=line.match(/^(.*?)\s*[|•]\s*([A-Za-z .'-]+?),?\s+([A-Z]{2})\s+(\d{5}(?:-\d{4})?)$/);
    if(m&&/\d/.test(m[1]))return{street:m[1].trim(),city:m[2].trim(),state:m[3],postalCode:m[4],countryOrRegion:usCountry(m[3])};
    m=line.match(cityStateZip);
    if(m){
      let street="";
      for(let j=i-1;j>=Math.max(0,i-3);j--){
        const prev=lines[j];
        if(/^from:|^sent:|^to:|^cc:|^subject:/i.test(prev))break;
        if((/^\d{1,6}\s+/.test(prev)||/p\.?o\.?\s*box/i.test(prev))&&(streetWord.test(prev)||/\d/.test(prev))){street=prev+(street?"\n"+street:"");}
        else if(street)break;
      }
      return{street,city:m[1].trim(),state:m[2],postalCode:m[3],countryOrRegion:usCountry(m[2])};
    }
  }
  return{street:"",city:"",state:"",postalCode:"",countryOrRegion:""}
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
  for(let i=stop-1;i>=Math.max(0,stop-8);i--){if(looksLikePersonName(lines[i]))return lines[i]}
  return inferPersonName(sig);
}

function anchoredSignatureCandidates(body,currentName,currentEmail,myEmail){
  const lines=norm(body).split("\n").map(x=>x.trim());
  const isHeader=l=>/^\s*(from|sent|to|cc|bcc|subject):\s*/i.test(l)||/^[-_]{5,}$/.test(l);
  const isDisclaimer=l=>/confidential|privileged|intended recipient|virus|disclaimer|please consider the environment/i.test(l);
  const byEmail=new Map();
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
  // If the current sender has no email printed in the signature, preserve the legacy segment
  // method only as a fallback for that sender.
  if(currentEmail&&!sameEmail(currentEmail,myEmail)&&!byEmail.has(currentEmail.toLowerCase())){
    const legacy=splitMessage(body,currentName,currentEmail).find(x=>sameEmail(x.email,currentEmail));
    if(legacy){
      const parsed=parseContact(legacy.name,legacy.email,legacy.text);
      const actualSigName=inferAnchoredName(parsed.signature,"");
      const sigName=String(actualSigName||"").toLowerCase().replace(/[^a-z0-9]+/g," ").trim();
      const curName=String(currentName||"").toLowerCase().replace(/[^a-z0-9]+/g," ").trim();
      // Only keep this weaker fallback if the signature itself clearly belongs to the
      // current sender; otherwise it is safer to omit it than attach someone else's block.
      if(parsed.signature && curName && sigName===curName)byEmail.set(currentEmail.toLowerCase(),{score:1,name:legacy.name,email:legacy.email,text:legacy.text,parsed});
    }
  }
  return [...byEmail.values()].map(x=>({name:x.name,email:x.email,text:x.text,parsed:x.parsed}));
}

function renderCandidates(){const box=$("candidates");box.innerHTML="";candidates.forEach((c,i)=>{const el=document.createElement("div");el.className="candidate";el.innerHTML=`<div class="candidate-head"><input type="checkbox" data-candidate="${i}" checked><div><div class="candidate-name">${html(c.name||"Unknown sender")}</div><div class="muted">${html(c.email||"No email found")}</div></div></div><div class="preview">${html(c.parsed.signature||"No signature block confidently found")}</div>`;box.appendChild(el)});$("candidateSection").hidden=false}
function html(s){return String(s||"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[c]))}
function htmlBodyToText(htmlText){
  try{
    const doc=new DOMParser().parseFromString(htmlText||"","text/html");
    doc.querySelectorAll("script,style,noscript").forEach(n=>n.remove());
    doc.querySelectorAll("a").forEach(a=>{
      const href=(a.getAttribute("href")||"").trim();
      let visible=(a.textContent||"").replace(/\s+/g," ").trim();
      if(!visible){
        if(/^mailto:/i.test(href)) visible=decodeURIComponent(href.replace(/^mailto:/i,"").split("?")[0]);
        else if(/^tel:/i.test(href)) visible=decodeURIComponent(href.replace(/^tel:/i,"").split("?")[0]);
        else if(/^https?:/i.test(href)){ const decoded=decodeProofpointUrl(href); if(!isJunkResourceUrl(decoded)) visible=normalizeWebsiteUrl(decoded); }
      }
      if(isJunkResourceUrl(visible))visible="";
      a.replaceWith(doc.createTextNode(visible?` ${visible} `:" "));
    });
    doc.querySelectorAll("img").forEach(img=>{
      const alt=(img.getAttribute("alt")||img.getAttribute("title")||"").trim();
      const safeAlt=isJunkResourceUrl(alt)?"":alt;
      img.replaceWith(doc.createTextNode(safeAlt?` ${safeAlt} `:" "));
    });
    doc.querySelectorAll("br").forEach(br=>br.replaceWith(doc.createTextNode("\n")));
    doc.querySelectorAll("p,div,li,tr,table,blockquote,td").forEach(el=>{el.appendChild(doc.createTextNode("\n"))});
    return norm(doc.body.textContent||"").replace(/[\u200B-\u200D\uFEFF]/g,"").replace(/\n[ \t]+/g,"\n").replace(/\n{3,}/g,"\n\n");
  }catch(_){return ""}
}
function readBody(item){
  return new Promise((resolve,reject)=>{
    item.body.getAsync(Office.CoercionType.Html,r=>{
      if(r.status===Office.AsyncResultStatus.Succeeded){const t=htmlBodyToText(r.value||"");if(t.trim())return resolve(t)}
      item.body.getAsync(Office.CoercionType.Text,r2=>r2.status===Office.AsyncResultStatus.Succeeded?resolve(r2.value||""):reject(new Error(r2.error?.message||"Unable to read the email body.")))
    })
  })
}
async function scan(){try{
  status("Reading the current email chain…");$("reviewSection").hidden=true;
  const item=Office.context.mailbox.item;if(!item||item.itemType!==Office.MailboxEnums.ItemType.Message)throw new Error("Open or select an email message first.");
  const from=item.from||{},body=await readBody(item);
  const myEmail=(Office.context.mailbox.userProfile?.emailAddress||"").toLowerCase();
  candidates=anchoredSignatureCandidates(body,from.displayName||"",from.emailAddress||"",myEmail);
  renderCandidates();
  status(`Found ${candidates.length} possible contact${candidates.length===1?"":"s"}. Your own messages/signature are ignored. Select the people you want to process.`,"ok")
}catch(e){status(e.message||String(e),"error")}}

function clientId(){return localStorage.getItem("ccfe_client_id")||""}
function showAuthSetup(){const id=clientId();$("authSetup").hidden=!!id;$("clientId").value=id}
async function initMsal(){const id=clientId();if(!id)throw new Error("Microsoft Contacts access is not configured yet. Enter the Application (client) ID first.");if(!msalInstance){msalInstance=await createNestablePublicClientApplication({auth:{clientId:id,authority:"https://login.microsoftonline.com/common"},cache:{cacheLocation:"localStorage"}})}return msalInstance}
async function accessToken(){const pca=await initMsal();const request={scopes:GRAPH_SCOPES,loginHint:Office.context.mailbox.userProfile.emailAddress};try{return (await pca.acquireTokenSilent(request)).accessToken}catch(err){if(err instanceof InteractionRequiredAuthError || /interaction|consent|login/i.test(String(err?.errorCode||err?.message||err))){return (await pca.acquireTokenPopup(request)).accessToken}throw err}}
async function graph(path,opts={}){const token=await accessToken();const r=await fetch("https://graph.microsoft.com/v1.0"+path,{...opts,headers:{Authorization:`Bearer ${token}`,"Content-Type":"application/json",...(opts.headers||{})}});if(!r.ok){const t=await r.text();throw new Error(`Microsoft Contacts error ${r.status}: ${t.slice(0,350)}`)}if(r.status===204)return null;return r.json()}
async function loadContacts(){let url="/me/contacts?$top=250&$select=id,displayName,givenName,middleName,surname,companyName,jobTitle,emailAddresses,businessPhones,mobilePhone,businessHomePage,businessAddress,personalNotes";const all=[];while(url){const data=await graph(url.replace("https://graph.microsoft.com/v1.0",""));all.push(...(data.value||[]));url=data["@odata.nextLink"]||""}return all}
function n(s){return String(s||"").trim().toLowerCase().replace(/[^a-z0-9]+/g," ").trim()}
function digits(s){return String(s||"").replace(/\D/g,"").slice(-10)}
function existingFlat(c){const a=c.businessAddress||{};return{givenName:c.givenName||"",middleName:c.middleName||"",surname:c.surname||"",companyName:c.companyName||"",jobTitle:c.jobTitle||"",email:(c.emailAddresses?.[0]?.address)||"",businessPhone:(c.businessPhones?.[0])||"",mobilePhone:c.mobilePhone||"",businessFax:"",businessHomePage:c.businessHomePage||"",street:a.street||"",city:a.city||"",state:a.state||"",postalCode:a.postalCode||"",countryOrRegion:a.countryOrRegion||"",personalNotes:c.personalNotes||""}}
function matchContact(p){const email=n(p.email);if(email){const m=graphContacts.find(c=>(c.emailAddresses||[]).some(e=>n(e.address)===email));if(m)return{contact:m,confidence:"Exact email"}}
  const full=n([p.givenName,p.middleName,p.surname].filter(Boolean).join(" "));const companyN=n(p.companyName);const ph=[digits(p.businessPhone),digits(p.mobilePhone)].filter(Boolean);
  const candidates2=graphContacts.filter(c=>{const f=n([c.givenName,c.middleName,c.surname].filter(Boolean).join(" "));if(!full||f!==full)return false;const cf=n(c.companyName);if(companyN&&cf&&companyN===cf)return true;const cp=[...(c.businessPhones||[]),c.mobilePhone||""].map(digits).filter(Boolean);return ph.some(x=>cp.includes(x))});
  return candidates2.length===1?{contact:candidates2[0],confidence:"Name + company/phone"}:null}
function fieldDiffs(parsed,existing){const out=[];for(const [key,label] of Object.entries(FIELD_META)){const nv=(parsed[key]||"").trim(),ov=(existing[key]||"").trim();if(!nv)continue;if(n(nv)===n(ov))continue;out.push({key,label,old:ov,new:nv,defaultChecked:!ov})}return out}
function graphPayload(p,keys=null){const use=k=>!keys||keys.has(k);const o={};if(use("givenName"))o.givenName=p.givenName||"";if(use("middleName"))o.middleName=p.middleName||"";if(use("surname"))o.surname=p.surname||"";if(use("companyName"))o.companyName=p.companyName||"";if(use("jobTitle"))o.jobTitle=p.jobTitle||"";if(use("email")&&p.email)o.emailAddresses=[{address:p.email,name:[p.givenName,p.surname].filter(Boolean).join(" ")||p.email}];if(use("businessPhone")&&p.businessPhone)o.businessPhones=[phoneForOutlook(p.businessPhone)];if(use("mobilePhone"))o.mobilePhone=phoneForOutlook(p.mobilePhone||"");if(use("businessHomePage"))o.businessHomePage=p.businessHomePage||"";if(use("personalNotes"))o.personalNotes=p.personalNotes||"";const addressKeys=["street","city","state","postalCode","countryOrRegion"].filter(use);if(addressKeys.length)o.businessAddress={street:p.street||"",city:p.city||"",state:p.state||"",postalCode:p.postalCode||"",countryOrRegion:p.countryOrRegion||""};return o}
function editableFields(p,idx){return `<div class="grid">${Object.entries(FIELD_META).map(([k,l])=>k==="personalNotes"?`<div class="field full notes-field"><div class="notes-label-row"><label>${html(l)} <span class="muted">— editable before saving</span></label><div class="notes-actions"><button type="button" class="mini-btn" data-notes-action="clear" data-index="${idx}">Clear Notes</button><button type="button" class="mini-btn" data-notes-action="restore" data-index="${idx}">Restore Signature</button></div></div><textarea data-review="${idx}" data-field="${k}" spellcheck="true" aria-label="Editable Notes">${html(p[k]||"")}</textarea><div class="muted notes-help">Delete, shorten, or rewrite this text. Outlook will receive exactly what remains in this box.</div></div>`:`<div class="field ${["email","street"].includes(k)?"full":""}"><label>${html(l)}</label><input data-review="${idx}" data-field="${k}" value="${html(p[k]||"")}"></div>`).join("")}</div>`}
function renderReviews(items){const box=$("reviews");box.innerHTML="";items.forEach((it,idx)=>{const p=it.parsed,m=it.match,existing=m?existingFlat(m.contact):null,diffs=existing?fieldDiffs(p,existing):[];const el=document.createElement("div");el.className="review";el.dataset.reviewCard=idx;const badge=m?`<span class="badge existing">Existing contact — ${html(m.confidence)}</span>`:`<span class="badge">New contact</span>`;el.innerHTML=`${badge}<div class="candidate-name" style="margin-top:6px">${html([p.givenName,p.middleName,p.surname].filter(Boolean).join(" ")||it.name)}</div><div class="muted">${html(p.email||it.email||"")}</div><div class="signature-sticky"><div class="sig-title">Signature block from email</div><pre>${html(p.signature||"No signature block confidently found")}</pre></div>${editableFields(p,idx)}${m?`<div class="diffs"><b>New or different information</b>${diffs.length?diffs.map(d=>`<label class="diff"><input type="checkbox" data-diff="${idx}" data-key="${d.key}" ${d.defaultChecked?"checked":""}><span>${html(d.label)}</span><span class="vals"><span class="old">Existing: ${html(d.old||"(blank)")}</span><br><span class="new">From email: ${html(d.new)}</span></span></label>`).join(""):"<div class=\"muted\">No new information was found.</div>"}</div><div class="toolbar"><button class="btn primary" data-action="update" data-index="${idx}" ${diffs.length?"":"disabled"}>Update Existing Contact</button><button class="btn" data-action="skip" data-index="${idx}">Skip</button></div>`:`<div class="toolbar"><button class="btn primary" data-action="create" data-index="${idx}">Create Contact</button><button class="btn" data-action="skip" data-index="${idx}">Skip</button></div>`}`;box.appendChild(el)});$("reviewSection").hidden=false;box.querySelectorAll("button[data-action]").forEach(b=>b.addEventListener("click",handleAction));box.querySelectorAll("button[data-notes-action]").forEach(b=>b.addEventListener("click",e=>{const idx=Number(e.currentTarget.dataset.index);const ta=document.querySelector(`textarea[data-review="${idx}"][data-field="personalNotes"]`);if(!ta)return;if(e.currentTarget.dataset.notesAction==="clear")ta.value="";else ta.value=(window.__reviewItems[idx]?.parsed?.signature||"");ta.focus();}));}
function currentParsed(idx){const d={};document.querySelectorAll(`[data-review="${idx}"][data-field]`).forEach(i=>d[i.dataset.field]=i.value.trim());d.businessPhone=phoneForOutlook(d.businessPhone);d.mobilePhone=phoneForOutlook(d.mobilePhone);d.businessFax=phoneForOutlook(d.businessFax);return d}
function finishCard(idx,label){const c=document.querySelector(`[data-review-card="${idx}"]`);c.classList.add("done");c.querySelectorAll("button").forEach(b=>b.disabled=true);const span=document.createElement("span");span.className="badge done";span.textContent=label;c.prepend(span)}
async function handleAction(ev){const action=ev.currentTarget.dataset.action,idx=Number(ev.currentTarget.dataset.index),item=window.__reviewItems[idx];if(action==="skip"){finishCard(idx,"Skipped");return}try{ev.currentTarget.disabled=true;status(action==="create"?"Creating Outlook contact…":"Updating Outlook contact…");const p=currentParsed(idx);if(action==="create"){await graph("/me/contacts",{method:"POST",body:JSON.stringify(graphPayload(p))});finishCard(idx,"Created");status("Contact created in Outlook Contacts.","ok")}else{const selected=new Set([...document.querySelectorAll(`input[data-diff="${idx}"]:checked`)].map(x=>x.dataset.key));if(!selected.size)throw new Error("Check at least one field to update.");await graph(`/me/contacts/${encodeURIComponent(item.match.contact.id)}`,{method:"PATCH",body:JSON.stringify(graphPayload(p,selected))});finishCard(idx,"Updated");status("Existing Outlook contact updated.","ok")}}catch(e){ev.currentTarget.disabled=false;status(e.message||String(e),"error")}}
async function compareSelected(){try{const selected=[...document.querySelectorAll("input[data-candidate]:checked")].map(x=>candidates[Number(x.dataset.candidate)]);if(!selected.length)throw new Error("Select at least one person first.");if(!clientId()){showAuthSetup();throw new Error("Complete the one-time Microsoft Contacts setup first.")}status("Signing in to Microsoft and checking Outlook Contacts…");graphContacts=await loadContacts();const items=selected.map(c=>({...c,match:matchContact(c.parsed)}));window.__reviewItems=items;renderReviews(items);status(`Compared ${items.length} selected contact${items.length===1?"":"s"} with ${graphContacts.length} Outlook contact${graphContacts.length===1?"":"s"}.`,"ok")}catch(e){status(e.message||String(e),"error")}}

Office.onReady(async info=>{if(info.host!==Office.HostType.Outlook){status("This page must be opened from the Outlook add-in.","error");return}showAuthSetup();$("saveClientId").addEventListener("click",()=>{const id=$("clientId").value.trim();if(!/^[0-9a-f-]{36}$/i.test(id)){status("That does not look like a Microsoft Application (client) ID.","error");return}localStorage.setItem("ccfe_client_id",id);msalInstance=null;showAuthSetup();status("Client ID saved. You can now compare contacts.","ok")});$("selectAll").addEventListener("click",()=>document.querySelectorAll("input[data-candidate]").forEach(x=>x.checked=true));$("selectNone").addEventListener("click",()=>document.querySelectorAll("input[data-candidate]").forEach(x=>x.checked=false));$("scanAgain").addEventListener("click",scan);$("compareSelected").addEventListener("click",compareSelected);await scan()});
