/* Signature image inventory. Runs before contact fields are populated.
 * All image evidence is scoped to one physical signature occurrence.
 * No external URLs or binaries are copied into contact Notes.
 */
(function(root){
"use strict";
const marker=i=>"SIGIMAGE"+i+"ENDMARKER";
const same=(a,b)=>String(a||"").trim().toLowerCase()===String(b||"").trim().toLowerCase();
// Step 1: discover text signatures and their boundaries. At this point no
// image is classified or treated as company evidence.
function inventory(html,fromName,fromEmail,myEmail,api){
 const occurrences=[],images=[];
 if(!html||!root.DOMParser)return{images,occurrences};
 try{
  const doc=new root.DOMParser().parseFromString(html,"text/html");
  const visible=api.visibleSignatureTextFromNode(doc.body);
  const signatures=api.captureSignatureBlocks(visible,fromName,fromEmail,myEmail);
  if(!signatures.length)return{images,occurrences};

  // Step 2: make a second structural view to locate images relative to the
  // already established signatures. Markers represent DOM positions only:
  // they never enter the discovery or contact parsing input.
  const clone=doc.body.cloneNode(true);
  const byIndex=new Map();
  [...clone.querySelectorAll("img")].forEach((img,i)=>{
   const src=String(img.getAttribute("src")||"").trim();
   const alt=String(img.getAttribute("alt")||img.getAttribute("title")||"").trim();
   const width=parseInt(img.getAttribute("width")||img.style?.width||"0",10)||0;
   const height=parseInt(img.getAttribute("height")||img.style?.height||"0",10)||0;
   const usable=/^(?:cid:|data:image\/|https?:\/\/)/i.test(src)&&
     !(width&&height&&width<=24&&height<=24);
   byIndex.set(i,{id:i,src,alt,width,height,usable});
   img.replaceWith(clone.ownerDocument.createTextNode("\n"+marker(i)+"\n"));
  });
  if(!byIndex.size){
   for(const c of signatures)occurrences.push({email:String(c.email||"").toLowerCase(),images:[]});
   return{images,occurrences};
  }
  const marked=api.visibleSignatureTextFromNode(clone);
  const originalLines=api.norm(visible).split("\n");
  const markedLines=api.norm(marked).split("\n");
  const plainToMarked=[];
  let originalIndex=0,aligned=true;
  for(let i=0;i<markedLines.length;i++){
   if(/^SIGIMAGE\d+ENDMARKER$/.test(markedLines[i].trim()))continue;
   if(String(originalLines[originalIndex]||"").trim()!==markedLines[i].trim()){
    aligned=false;break;
   }
   plainToMarked.push(i);
   originalIndex++;
  }
  if(!aligned||originalIndex!==originalLines.length){
   // In complex HTML, a logo may split an inline text node. Without a
   // proven mapping to a specific signature, do not attribute any image.
   for(const c of signatures)occurrences.push({email:String(c.email||"").toLowerCase(),images:[]});
   return{images,occurrences};
  }
  const isMarker=i=>String(markedLines[i]||"").trim().match(/^SIGIMAGE(\d+)ENDMARKER$/);
  const isBoundary=v=>/^\s*(?:From:|Sent:|To:|Cc:|Subject:|Confidential(?:ity)? Notice:|Disclaimer:)/i.test(v||"");
  for(const candidate of signatures){
   const start=plainToMarked[candidate.globalStart??-1];
   const end=plainToMarked[candidate.globalEnd??-1];
   const refs=[];
   if(start!==undefined&&end!==undefined&&end>=start){
    // A text capture may include trailing short prose. Only treat an image as
    // demonstrably in a signature when it precedes the last contact-field
    // evidence, or when it is directly adjacent to a hard message boundary.
    const contactEvidence=v=>/@[a-z0-9.-]+\.[a-z]{2,}/i.test(v)||
      /(?:www\.|https?:\/\/)/i.test(v)||
      /(?:^|\s)(?:tel|phone|mobile|cell|fax|direct)\s*:/i.test(v)||
      /\b\d{3}[.\s-]?\d{3}[.\s-]?\d{4}\b/.test(v)||
      /\b[A-Z]{2}\s+\d{5}(?:-\d{4})?\b/.test(v);
    let lastEvidence=start;
    for(let k=start;k<=end;k++){
      if(!isMarker(k)&&contactEvidence(markedLines[k]))lastEvidence=k;
    }
    // Treat after-last-contact images exactly like trailing images below.
    for(let k=start;k<=end;k++){
     const m=isMarker(k);if(!m||k>lastEvidence)continue;
     const image=byIndex.get(Number(m[1]));
     if(image?.usable&&!refs.some(x=>x.id===image.id))refs.push(image);
    }
    // A logo between the final contact line and the next hard boundary can
    // still belong to the signature. A prose sentence means it is ambiguous.
    const last=lastEvidence;
    let tailAt=last+1,tail=[];
    while(tailAt<markedLines.length&&isMarker(tailAt)&&tail.length<3){
      const image=byIndex.get(Number(isMarker(tailAt)[1]));
      if(image?.usable)tail.push(image);
      tailAt++;
    }
    if(tail.length&&(tailAt===markedLines.length||isBoundary(markedLines[tailAt]))){
      for(const image of tail)if(!refs.some(x=>x.id===image.id))refs.push(image);
    }

    // A contiguous logo immediately after the final signature text belongs
    // to the block only when followed by a hard message/disclaimer boundary,
    // or the document ends. Otherwise its membership is ambiguous.
    let k=end+1,trailing=[];
    while(k<markedLines.length&&isMarker(k)&&trailing.length<3){
     const image=byIndex.get(Number(isMarker(k)[1]));
     if(image?.usable)trailing.push(image);
     k++;
    }
    if(trailing.length&&(k===markedLines.length||isBoundary(markedLines[k]))){
     for(const image of trailing)if(!refs.some(x=>x.id===image.id))refs.push(image);
    }

    // A logo directly between a closing phrase and the signature owner's
    // name may precede the text anchor. Do not reach into normal email prose.
    k=start-1;const leading=[];
    while(k>=0&&isMarker(k)&&leading.length<3){
     const image=byIndex.get(Number(isMarker(k)[1]));
     if(image?.usable)leading.push(image);
     k--;
    }
    if(leading.length&&k>=0&&
      /^(?:kind regards|best regards|regards|sincerely|thank you|thanks)[,!.\s]*$/i.test(markedLines[k].trim())){
     for(const image of leading)if(!refs.some(x=>x.id===image.id))refs.push(image);
    }
   }
   occurrences.push({email:String(candidate.email||"").toLowerCase(),images:refs});
   for(const image of refs)if(!images.some(x=>x.id===image.id))images.push(image);
  }
 }catch(err){console.warn("Signature boundary image inventory skipped",err)}
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
  // If HTML and Text disagree about the occurrence count, do not guess.
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
