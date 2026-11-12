/* Signature image inventory. Runs before contact fields are populated.
 * All image evidence is scoped to one physical signature occurrence.
 * No external URLs or binaries are copied into contact Notes.
 */
(function(root){
"use strict";
const marker=i=>"SIGIMAGE"+i+"ENDMARKER";
const same=(a,b)=>String(a||"").trim().toLowerCase()===String(b||"").trim().toLowerCase();
function inventory(html,fromName,fromEmail,myEmail,api){
 const images=[],occurrences=[];
 if(!html||!root.DOMParser)return{images,occurrences};
 try{
  const doc=new root.DOMParser().parseFromString(html,"text/html");
  const body=doc.body.cloneNode(true);
  [...body.querySelectorAll("img")].forEach((img,i)=>{
   const src=String(img.getAttribute("src")||"").trim();
   const alt=String(img.getAttribute("alt")||img.getAttribute("title")||"").trim();
   const w=parseInt(img.getAttribute("width")||img.style?.width||"0",10)||0;
   const h=parseInt(img.getAttribute("height")||img.style?.height||"0",10)||0;
   const usable=/^(?:cid:|data:image\/|https?:\/\/)/i.test(src) && !(w&&h&&w<=24&&h<=24);
   images.push({id:i,src,alt,width:w,height:h,usable});
   img.replaceWith(body.ownerDocument.createTextNode("\n"+marker(i)+"\n"));
  });
  if(!images.length)return{images,occurrences};
  const marked=api.visibleSignatureTextFromNode(body);
  const lines=api.norm(marked).split("\n");
  const found=api.captureSignatureBlocks(marked,fromName,fromEmail,myEmail);
  for(const candidate of found){
   const refs=[];
   const start=Math.max(0,(candidate.globalStart||0)-5);
   const end=Math.min(lines.length-1,candidate.globalEnd??0);
   for(let j=start;j<=end;j++){
    if(j<(candidate.globalStart||0) && lines.slice(j+1,candidate.globalStart).some(v=>v.trim()))continue;
    const match=String(lines[j]||"").trim().match(/^SIGIMAGE(\d+)ENDMARKER$/);
    if(!match)continue;
    const im=images[Number(match[1])];
    if(im?.usable&&!refs.some(x=>x.id===im.id))refs.push(im);
   }
   occurrences.push({email:String(candidate.email||"").toLowerCase(),images:refs});
  }
 }catch(err){console.warn("Signature image inventory failed",err)}
 return{images,occurrences};
}
function associate(candidates,inventory){
 const groups=new Map();
 for(const occurrence of inventory?.occurrences||[]){
  if(!occurrence.email)continue;
  if(!groups.has(occurrence.email))groups.set(occurrence.email,[]);
  groups.get(occurrence.email).push(occurrence);
 }
 for(const [email,occurrences] of groups){
  const targets=(candidates||[]).filter(c=>c.candidateType==="signature"&&same(c.email,email));
  // Ambiguous occurrence matching is safer left unresolved than attributed
  // to the wrong person in a forwarded chain.
  if(occurrences.length!==targets.length)continue;
  targets.forEach((candidate,i)=>{
   if(!occurrences[i].images.length)return;
   candidate._signatureImages=occurrences[i].images;
   candidate._imageEvidence={count:occurrences[i].images.length,status:"detected",company:""};
  });
 }
}
function attachmentForCid(item,src){
 if(!/^cid:/i.test(src))return null;
 const cid=String(src).replace(/^cid:/i,"").replace(/^<|>$/g,"").toLowerCase();
 const local=cid.split("@")[0];
 const inline=Array.from(item?.attachments||[]).filter(a=>a.isInline);
 const exact=inline.filter(a=>same(String(a.contentId||"").replace(/^<|>$/g,""),cid));
 if(exact.length===1)return exact[0];
 const named=inline.filter(a=>same(a.name,local));
 return named.length===1?named[0]:null;
}
async function imageBytes(item,image,api){
 if(/^cid:/i.test(image.src)){
  const att=attachmentForCid(item,image.src);
  return att?await api.attachmentImageDataUrl(item,att):"";
 }
 if(/^(?:https?:\/\/|data:image\/)/i.test(image.src))
  return api.sourceToDataUrl(image.src);
 return "";
}
async function recognize(item,candidates,api){
 const pending=(candidates||[]).filter(c=>c._signatureImages?.length);
 if(!pending.length)return;
 const cache=new Map();
 let worker=null,requests=0;
 try{
  const OCR=await api.loadTesseract();
  worker=await OCR.createWorker("eng");
  try{await worker.setParameters({tessedit_pageseg_mode:"6"})}catch(_){}
  for(const c of pending){
   const result=c._imageEvidence;
   for(const image of c._signatureImages.slice(0,3)){
    if(requests>=12)break;
    let content=cache.get(image.src);
    if(content===undefined){
     requests++;
     content="";
     const imageData=await imageBytes(item,image,api);
     if(imageData){
      try{content=String((await worker.recognize(imageData))?.data?.text||"")}catch(_){}
     }
     cache.set(image.src,content);
    }
    if(!content)continue;
    result.status="examined";
    const company=api.logoCompanyFromOcrText(content,c.email,c.parsed?.businessHomePage||"");
    if(company){
     result.company=company;
     result.status="company-verified";
     break;
    }
   }
  }
 }catch(err){console.warn("Signature image recognition unavailable",err)}
 finally{try{if(worker)await worker.terminate()}catch(_){}}
}
function apply(candidates,api){
 for(const c of candidates||[]){
  const e=c._imageEvidence;
  if(!e||!c.parsed)continue;
  if(e.company){
   const previous=String(c.parsed.companyName||"").trim();
   if(!previous||api.companyFromPlainDomainOnly(c))c.parsed.companyName=e.company;
   else if(api.structuredCompanyLineKey(previous)!==api.structuredCompanyLineKey(e.company))
    e.status="company-mismatch";
   const note="Company logo identified: "+e.company;
   if(!String(c.parsed.personalNotes||"").includes(note))
    c.parsed.personalNotes=[c.parsed.personalNotes,note].filter(Boolean).join("\n");
  }
  c.parsed._imageEvidence={count:e.count,status:e.status,company:e.company};
 }
}
root.SignatureImages={inventory,associate,recognize,apply,attachmentForCid};
})(window);
