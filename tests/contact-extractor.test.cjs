// Tests use the actual production app.js with a mocked Office host.
// Fixture names, email addresses, and phone numbers are synthetic.
const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict');
const {test}=require('node:test'),path=require('node:path');
const source=fs.readFileSync(path.join(__dirname,'..','app.js'),'utf8').replace(/^import\s+[^\n]+\n/,'');
const ctx={console,URL,Office:{onReady(){},AsyncResultStatus:{Succeeded:'succeeded'}},setTimeout,clearTimeout};
ctx.window={Tesseract:{createWorker:async()=>({
 setParameters:async()=>{},
 recognize:async()=>({data:{text:'ALDER\nNORTH\nWEST'}}),
 terminate:async()=>{}
})}};
vm.runInNewContext(source+'\n;globalThis.TEST_API={captureSignatureBlocks,structuredAddress,structuredPhones,signatureDiscoveryText,phoneForOutlook,graphPayload,fieldDiffs,compactSignatureNotes,v4CompoundOfficeAddress,logoCompanyFromOcrText,v4ContactFromSignature,companyFromPlainDomainOnly,structuredCompanyLineKey,norm};',ctx,{filename:'app.js'});

const imageModule=fs.readFileSync(path.join(__dirname,'..','signature-images.js'),'utf8');
vm.runInNewContext(imageModule,ctx,{filename:'signature-images.js'});
const images=ctx.window.SignatureImages;
const imageApi={
 loadTesseract:async()=>ctx.window.Tesseract,
 attachmentImageDataUrl:async()=> 'data:image/jpeg;base64,AA',
 sourceToDataUrl:async()=> 'data:image/jpeg;base64,AA',
 logoCompanyFromOcrText:ctx.TEST_API.logoCompanyFromOcrText
};
const applyApi={
 companyFromPlainDomainOnly:ctx.TEST_API.companyFromPlainDomainOnly,
 structuredCompanyLineKey:ctx.TEST_API.structuredCompanyLineKey
};
const p=ctx.TEST_API, own='owner@example.com';
const first=[
'Anna','','Marsh','','Sales & Marketing Director','','Northbridge US Inc',
'','','','','','','','','(717)-305-4694','','+44 01922 714082 (DD)',
'','','anna@example.com','','www.northbridge.example','','1 International Plaza, Philadelphia,',' PA 19153'
].join('\n');
const second=[
'Peter Hale','','Business Development Director','','Sales','','Northbridge Limited',
'','01922 711185','','01922 714077 (DD)','','peter@example.com',
'','Walsall Road, Aldridge, WS9 0SW'
].join('\n');
const message=[
'From: Northbridge USA <desk@example.com>',
'Sent: Tuesday','Subject: Example correspondence','','Hi,','Thanks for the information.','',first,'',
'From: Owner <owner@example.com>','Sent: Monday','Subject: RE: Example correspondence','','Acknowledged.','',
'From: Northbridge USA <desk@example.com>','Sent: Monday','Subject: Example correspondence','',first,'',
'From: Peter Hale <peter@example.com>','Sent: Sunday','Subject: Example correspondence','',second
].join('\n');
const candidates=()=>p.captureSignatureBlocks(p.signatureDiscoveryText(message),'Northbridge USA','desk@example.com',own);
test('finds two Anna and one Peter signatures under a shared sender',()=>{
 const c=candidates();
 assert.equal(c.filter(x=>x.name==='Anna Marsh').length,2);
 assert.equal(c.filter(x=>x.name==='Peter Hale').length,1);
 assert(c.every(x=>x.email!==own));
});
test('Notes remove empty layout lines but retain every real signature line',()=>{
 const c=candidates().filter(x=>x.name==='Anna Marsh');
 for(const x of c){
  assert(x.parsed.personalNotes.includes('anna@example.com'));
  assert(x.parsed.personalNotes.includes('1 International Plaza'));
  assert(x.parsed.personalNotes.includes('+44 01922 714082'));
  assert(x.rawSignatureBlock.includes('\n\n'), 'source preview remains unmodified');
  assert(!/\n\s*\n/.test(x.parsed.personalNotes), 'Notes has no blank lines');
  assert.equal(x.parsed.personalNotes,p.compactSignatureNotes(x.rawSignatureBlock));
  assert.equal(x.parsed.personalNotes.split('\n').length,
    x.rawSignatureBlock.split('\n').filter(line=>line.trim()).length);
 }
});
test('signature name and email override shared sender metadata',()=>{
 for(const x of candidates().filter(x=>x.name==='Anna Marsh')){
  assert.equal(x.email,'anna@example.com');
  assert.equal(x.parsed.givenName,'Anna');
  assert.equal(x.parsed.surname,'Marsh');
 }
});
test('wrapped US postal street address stays intact',()=>{
 const a=p.structuredAddress(first);
 assert.equal(a.street,'1 International Plaza');
 assert.equal(a.city,'Philadelphia');
 assert.equal(a.state,'PA');
 assert.equal(a.postalCode,'19153');
 assert.equal(a.countryOrRegion,'United States');
});
test('Royal Mail postcode and UK post town have no fabricated US state',()=>{
 const a=p.structuredAddress(second);
 assert.equal(a.street,'Walsall Road');
 assert.equal(a.city,'Aldridge');
 assert.equal(a.postalCode,'WS9 0SW');
 assert.equal(a.state,'');
 assert.equal(a.countryOrRegion,'United Kingdom');
});
test('UK international and US phone normalization',()=>{
 assert.equal(p.phoneForOutlook('+44 01922 714082'),'+441922714082');
 assert.equal(p.phoneForOutlook('01922 714077','United Kingdom'),'+441922714077');
 assert.equal(p.phoneForOutlook('(717)-305-4694'),'7173054694');
});
test('explicit Mobile wins over DD',()=>{
 const q=p.structuredPhones('DD: 01922 714077\nMobile: 07123 456789','United Kingdom');
 assert.equal(q.mobilePhone,'+447123456789');
 assert.equal(q.businessPhone,'+441922714077');
});
test('DD alone maps to Mobile; DD with extension maps to Business',()=>{
 const q=p.structuredPhones('DD: 01922 714077','United Kingdom');
 assert.equal(q.mobilePhone,'+441922714077');
 const ext=p.structuredPhones('DD: 01922 714077 ext 50','United Kingdom');
 assert.equal(ext.businessPhone,'+441922714077x50');
 assert.equal(ext.mobilePhone,'');
});
test('Graph payload retains international numbers, postcode, and Notes',()=>{
 const c=candidates().find(x=>x.name==='Peter Hale');
 assert(c);
 const body=p.graphPayload(c.parsed);
 assert.equal(body.mobilePhone,'+441922714077');
 assert.equal(body.businessPhones[0],'+441922711185');
 assert.equal(body.businessAddress.postalCode,'WS9 0SW');
 assert.equal(body.personalNotes,p.compactSignatureNotes(c.rawSignatureBlock));
});
test('comparison does not report identical name or email as differences',()=>{
 const old={givenName:'Anna',surname:'Marsh',email:'anna@example.com'};
 const changes=p.fieldDiffs({...old,companyName:'Northbridge US Inc'},old);
 assert.equal(changes.filter(x=>x.key==='companyName').length,1);
 assert(!changes.some(x=>x.key==='givenName'||x.key==='email'));
});
test('prose Good News is not falsely treated as a person',()=>{
 const bad=['From: Team <desk@example.com>','The schedule is below.',
  'Good','News','This week we have a delivery.','Email: team@example.com'].join('\n');
 const found=p.captureSignatureBlocks(bad,'Team','desk@example.com',own);
 assert(!found.some(x=>x.name==='Good News'));
});

test('Notes formatter trims whitespace-only and zero-width lines, never merges content',()=>{
 const raw='  Name   \r\n \r\n\u200b\r\n  Job Title  \r\n\r\n  +44 01922 714082 (DD)  \r\n\r\n  person@example.com  ';
 const result=p.compactSignatureNotes(raw);
 assert.equal(result,'Name\nJob Title\n+44 01922 714082 (DD)\nperson@example.com');
 assert.equal(p.compactSignatureNotes(result),result,'idempotent formatting');
 assert.equal(raw.includes('\r\n\r\n'),true,'source remains untouched');
});
test('UK signature and Graph Notes contain no extra empty lines',()=>{
 const x=candidates().find(y=>y.name==='Peter Hale');
 assert(x);
 assert(!/\n\s*\n/.test(x.parsed.personalNotes));
 assert.equal(p.graphPayload(x.parsed).personalNotes,x.parsed.personalNotes);
 assert(x.rawSignatureBlock.includes('\n\n'));
});

test('a labeled line with physical and postal addresses uses two Street lines',()=>{
 const sig=['Mona Pierce','Administrative Assistant',
   'Jackson Office: 130 E. Cedar, Suite 103, Jackson, WY 83001 | P.O. Box 555, Jackson, WY 83001',
   'Tel: 307.734.1908 | Fax: 307.733.3752','www.anw-law.com'].join('\n');
 const parsed=p.v4ContactFromSignature('Mona Pierce','mona@anw-law.com',sig);
 assert.equal(parsed.street,'130 E. Cedar, Suite 103\nP.O. Box 555');
 assert.equal(parsed.city,'Jackson');
 assert.equal(parsed.state,'WY');
 assert.equal(parsed.postalCode,'83001');
 assert.equal(parsed.countryOrRegion,'United States');
 assert(!parsed.street.includes('Jackson Office'));
 assert(!parsed._audit.unaccountedText.includes('P.O. Box'));
 assert(!parsed._audit.unaccountedText.includes('130 E. Cedar'));
 assert.equal(p.graphPayload(parsed).businessAddress.street,'130 E. Cedar, Suite 103\nP.O. Box 555');
});
test('addresses in different cities are not silently combined',()=>{
 const input='Jackson Office: 130 E. Cedar, Jackson, WY 83001 | P.O. Box 555, Cheyenne, WY 82001';
 assert.equal(p.v4CompoundOfficeAddress(input),null);
});
test('logo OCR requires confirmation by matching email domain initials',()=>{
 assert.equal(p.logoCompanyFromOcrText('ALDER\nNORTH\nWEST','mona@anw-law.com'),'Alder North West');
 assert.equal(p.logoCompanyFromOcrText('RANDOM DESIGN STUDIO','mona@anw-law.com'),'');
 assert.equal(p.logoCompanyFromOcrText('ALDER NORTH WEST','mona@gmail.com'),'');
});

test('image logo enriches the right signature and preserves original raw block',async()=>{
 const sig=['Mona Pierce','Administrative Assistant',
  'Jackson Office: 130 E. Cedar, Suite 103, Jackson, WY 83001 | P.O. Box 555, Jackson, WY 83001',
  'Tel: 307.734.1908 | Fax: 307.733.3752','www.anw-law.com'].join('\n');
 const parsed=p.v4ContactFromSignature('Mona Pierce','mona@anw-law.com',sig);
 const original=parsed.signature,priorNotes=parsed.personalNotes;
 const candidate={candidateType:'signature',email:'mona@anw-law.com',parsed,
  _signatureImages:[{src:'cid:company-logo',alt:'logo',width:320,height:90}],
  _imageEvidence:{count:1,status:'detected',company:''}};
 const item={attachments:[{id:'logo',contentId:'company-logo',name:'logo.jpg',isInline:true}]};
 await images.recognize(item,[candidate],imageApi);
 images.apply([candidate],applyApi);
 assert.equal(parsed.companyName,'Alder North West');
 assert.equal(parsed.signature,original);
 assert(parsed.personalNotes.startsWith(priorNotes));
 assert(parsed.personalNotes.endsWith('Company logo identified: Alder North West'));
 assert.equal(p.graphPayload(parsed).companyName,'Alder North West');
 assert.equal(p.graphPayload(parsed).businessAddress.street,'130 E. Cedar, Suite 103\nP.O. Box 555');
});
test('explicit legal firm name is not overwritten by a conflicting logo',async()=>{
 const sig='Martin Finch\nDesigner\nPine Tree Design LLC\nmfinch@anw-law.com';
 const parsed=p.v4ContactFromSignature('Martin Finch','mfinch@anw-law.com',sig);
 assert.equal(parsed.companyName,'Pine Tree Design LLC');
 const candidate={candidateType:'signature',email:'mfinch@anw-law.com',parsed,
  _signatureImages:[{src:'cid:company-logo',width:300,height:100}],
  _imageEvidence:{count:1,status:'detected',company:''}};
 const item={attachments:[{id:'logo',contentId:'company-logo',name:'logo.jpg',isInline:true}]};
 await images.recognize(item,[candidate],imageApi);
 images.apply([candidate],applyApi);
 assert.equal(parsed.companyName,'Pine Tree Design LLC');
 assert.equal(candidate._imageEvidence.status,'company-mismatch');
});
test('content IDs match a unique inline attachment, but never a neighboring image',()=>{
 const item={attachments:[
  {id:'a',isInline:true,contentId:'logo-aaa',name:'image001.png'},
  {id:'b',isInline:true,contentId:'logo-bbb',name:'image002.png'}
 ]};
 assert.equal(images.attachmentForCid(item,'cid:logo-bbb').id,'b');
 assert.equal(images.attachmentForCid(item,'cid:image001.png@abc').id,'a');
 assert.equal(images.attachmentForCid(item,'cid:other'),null);
 const ambiguous={attachments:[
  {id:'a',isInline:true,name:'same.png'},
  {id:'b',isInline:true,name:'same.png'}
 ]};
 assert.equal(images.attachmentForCid(ambiguous,'cid:same.png@abc'),null);
});
test('a forwarded-chain image is never attributed by email domain alone',()=>{
 const cand=[{candidateType:'signature',email:'mona@anw-law.com'},
  {candidateType:'signature',email:'peter@wcb.com'}];
 images.associate(cand,{occurrences:[{email:'mona@anw-law.com',images:[{src:'cid:mona'}]}]});
 assert(cand[0]._signatureImages?.length===1);
 assert(!cand[1]._signatureImages);
});
test('ambiguous repeated signatures are left unassigned rather than guessing',()=>{
 const cand=[{candidateType:'signature',email:'mona@anw-law.com'},
  {candidateType:'signature',email:'mona@anw-law.com'}];
 images.associate(cand,{occurrences:[{email:'mona@anw-law.com',images:[{src:'cid:logo'}]}]});
 assert(!cand[0]._signatureImages);
 assert(!cand[1]._signatureImages);
});
test('image inventory belongs to stage one and runs before stage two parsing',()=>{
 const scan=source.slice(source.indexOf('async function scan()'));
 assert(scan.indexOf('SignatureImages?.inventory(')>=0);
 assert(scan.indexOf('SignatureImages?.inventory(')<scan.indexOf('candidates=candidates.map('));
 assert(scan.indexOf('SignatureImages?.recognize(')<scan.indexOf('candidates=candidates.map('));
 assert(scan.indexOf('SignatureImages?.apply(')>scan.indexOf('candidates=candidates.map('));
});
