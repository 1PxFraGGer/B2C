import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import {randomUUID} from 'node:crypto';
const png='iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6wAAAAABJRU5ErkJggg==';
// Service doubles exist only in tests; the production backend uses Google's APIs.
function fixture({replacement=false,fail=false,timeoutAfterCommit=false}={}){
 const tables={
  Team_Members:{headers:['Member ID','Employee Name','Team','Active Status'],rows:[{_row:2,'Member ID':'MEM-SAHIL','Employee Name':'Sahil',Team:'USP, Ranking & Logo Uploadation','Active Status':'Active'}]},
  USP_Ranking_Assignments:{headers:['Assignment ID','Date','University Name','Assigned To','Member ID','Version','Last Updated','Completed At'],rows:[{_row:2,'Assignment ID':1,Date:'2026-10-14','University Name':'Test University','Assigned To':'Sahil','Member ID':'MEM-SAHIL',Version:1}]},
  USP_Ranking_Items:{headers:['Item ID','Assignment ID','University Name','Assigned To','Type','Item Name','Description','Status','Image File ID','Image URL','Image Preview','Remarks','Last Updated','Member ID','Version','Completed At','Item Folder ID'],rows:[{_row:2,'Item ID':'ITEM-001-NAAC','Assignment ID':1,'University Name':'Test University','Assigned To':'Sahil',Type:'USP','Item Name':'NAAC',Status:replacement?'Done':'Pending','Image File ID':replacement?'old-file':'','Image URL':replacement?'https://drive.google.com/file/d/old-file/view':'',Remarks:'original','Member ID':'MEM-SAHIL',Version:1,'Item Folder ID':replacement?'old-folder':''}]},
  Content_Verification:{headers:['Assignment ID'],rows:[]},
  Activity_Log:{headers:['Timestamp','Employee','Team','University','Task','Previous Value','Updated Value','Action','Event ID','Assignment ID','Item ID','Member ID','Request ID'],rows:[]}
 };
 const ids={Team_Members:3,USP_Ranking_Assignments:1,USP_Ranking_Items:2,Content_Verification:0,Activity_Log:4};
 const files=[],folders=[],props=new Map();
 function folder(id,name,parent){return {getId:()=>id,isTrashed:()=>false,getSharingAccess:()=> 'PRIVATE',getParents:()=>iterator(parent?[parent]:[]),createFolder(n){const f=folder('folder-'+(folders.length+1),n,this);folders.push({name:n,folder:f});return f;},createFile(blob){const file={id:'file-'+(files.length+1),trashed:false,getId(){return this.id;},getUrl(){return 'https://drive.google.com/file/d/'+this.id+'/view';},setTrashed(v){this.trashed=v;},isTrashed(){return this.trashed;},getSize:()=>100,getParents:()=>iterator([this]),getBlob:()=>blob};files.push(file);return file;}};}
 function iterator(arr){let i=0;return {hasNext:()=>i<arr.length,next:()=>arr[i++]};}
 const root=folder('root','Root'),oldFolder=folder('old-folder','NAAC',root),old={id:'old-file',trashed:false,getId(){return this.id;},getSize:()=>100,isTrashed(){return this.trashed;},setTrashed(v){this.trashed=v;},getParents:()=>iterator([oldFolder])};
 const context=vm.createContext({console,Date,Set,Map,JSON,Number,String,Array,Object,Math,RegExp,Error,
  PropertiesService:{getScriptProperties:()=>({getProperty:k=>props.get(k)||null,setProperty:(k,v)=>props.set(k,v),getProperties:()=>Object.fromEntries(props),deleteProperty:k=>props.delete(k)})},
  Utilities:{getUuid:randomUUID,computeDigest:(algorithm,v)=>[...Buffer.from(String(v)).subarray(0,30)],DigestAlgorithm:{SHA_256:'sha256'},Charset:{UTF_8:'utf8'},base64Decode:v=>[...Buffer.from(v,'base64')],newBlob:(bytes,mime,name)=>({bytes,mime,name,getDataAsString:()=>Buffer.from(bytes).toString()})},
  DriveApp:{Access:{ANYONE:'ANYONE',ANYONE_WITH_LINK:'ANYONE_WITH_LINK'},getFolderById:id=>id==='1FJ1jHU-koTlzKdo9qKw7-lbU88uFzjJH'?root:folders.find(f=>f.folder.getId()===id)?.folder||oldFolder,getFileById:id=>id==='old-file'?old:files.find(f=>f.id===id)},
  Sheets:{Spreadsheets:{batchUpdate({requests}){
   if(fail)throw Error('Sheets unavailable');
   for(const request of requests){if(request.updateCells){const r=request.updateCells,t=Object.values(tables).find(t=>t.sheet.getSheetId()===r.range.sheetId),record=t.rows.find(row=>row._row===r.range.startRowIndex+1);r.rows[0].values.forEach((cell,c)=>{const v=cell.userEnteredValue;record[t.headers[c]]=v.numberValue??v.stringValue??'';});}else if(request.appendCells){const r=request.appendCells,t=tables.Activity_Log;for(const row of r.rows)t.rows.push(Object.fromEntries(t.headers.map((h,i)=>[h,row.values[i].userEnteredValue.stringValue])));}}
   if(timeoutAfterCommit)throw Error('Response lost after commit');return {};
  }}}
 });
 vm.runInContext(fs.readFileSync(new URL('../apps-script/Code.gs',import.meta.url),'utf8'),context);
 for(const [name,t] of Object.entries(tables))t.sheet={getSheetId:()=>ids[name],getRange(row){const r=t.rows.find(r=>r._row===row);return {getValues:()=>[t.headers.map(h=>r?.[h]??'')],getFormulas:()=>[t.headers.map(()=> '')]};}};
 context.table_=name=>tables[name];
 const user={id:'MEM-SAHIL',memberId:'MEM-SAHIL',name:'Sahil',role:'employee',team:'USP, Ranking & Logo Uploadation'};
 const payload={itemId:'ITEM-001-NAAC',version:1,status:'Done',remarks:'verified',operationId:randomUUID(),image:{name:'logo.png',extension:'png',mime:'image/png',base64:png}};
 return {context,tables,user,payload,files,folders,old};
}
test('each upload creates the university/type/item folders and commits image + log together',()=>{
 const f=fixture(),result=f.context.saveItem_(f.user,f.payload);
 assert.equal(result.item.status,'Done');assert.equal(result.item.imageFileId,'file-1');assert.deepEqual(f.folders.map(f=>f.name),['Test University','USP','NAAC']);
 assert.equal(f.tables.Activity_Log.rows.length,2);assert.equal(f.tables.USP_Ranking_Items.rows[0].Version,2);
 assert.equal(f.context.saveItem_(f.user,f.payload).duplicate,true);assert.equal(f.files.length,1);
});
test('failed replacement preserves original reference, status, remarks, and file',()=>{
 const f=fixture({replacement:true,fail:true});assert.throws(()=>f.context.saveItem_(f.user,f.payload));
 const item=f.tables.USP_Ranking_Items.rows[0];assert.equal(item['Image File ID'],'old-file');assert.equal(item.Remarks,'original');assert.equal(item.Version,1);assert.equal(f.old.trashed,false);assert.equal(f.files[0].trashed,true);
});
test('a lost response after commit must not trash the now-linked replacement',()=>{
 const f=fixture({replacement:true,timeoutAfterCommit:true}),r=f.context.saveItem_(f.user,f.payload);assert.equal(r.recovered,true);assert.equal(f.files[0].trashed,false);assert.equal(f.old.trashed,false);
});
test('ownership, university mismatch, stale versions, and missing image are rejected',()=>{
 const f=fixture();assert.throws(()=>f.context.saveItem_({...f.user,memberId:'MEM-OTHER'},f.payload),e=>e.status===403);
 assert.throws(()=>f.context.saveItem_(f.user,{...f.payload,version:0}),e=>e.status===409);
 assert.throws(()=>f.context.saveItem_(f.user,{...f.payload,image:null}),e=>e.status===422);
 f.tables.USP_Ranking_Items.rows[0]['University Name']='Different University';assert.throws(()=>f.context.saveItem_(f.user,f.payload),e=>e.status===409);
});
test('a successful replacement keeps the original private file for recovery',()=>{
 const f=fixture({replacement:true}),r=f.context.saveItem_(f.user,f.payload);assert.equal(r.item.imageFileId,'file-1');assert.equal(f.old.trashed,false);
});

test('upload sets Done even when submitted as Pending',()=>{
 const f=fixture(),r=f.context.saveItem_(f.user,{...f.payload,status:'Pending'});assert.equal(r.item.status,'Done');
});
test('removing image clears only references and resets Pending; original stays recoverable',()=>{
 const f=fixture({replacement:true}),r=f.context.saveItem_(f.user,{...f.payload,image:null,removeImage:true});assert.equal(r.item.status,'Pending');assert.equal(r.item.imageFileId,'');assert.equal(r.item.imageUrl,'');assert.equal(r.item.folderId,'');assert.equal(f.old.trashed,false);assert.equal(f.files.length,0);
});
test('failed image removal preserves original image and Done status',()=>{
 const f=fixture({replacement:true,fail:true});assert.throws(()=>f.context.saveItem_(f.user,{...f.payload,image:null,removeImage:true}));assert.equal(f.tables.USP_Ranking_Items.rows[0]['Image File ID'],'old-file');assert.equal(f.tables.USP_Ranking_Items.rows[0].Status,'Done');assert.equal(f.old.trashed,false);
});
test('content status follows nonblank remarks and clearing remarks resets Pending',()=>{
 const f=fixture(),names=['Overview','Program & Fees','Admission','Placement','Student Support','FAQs'];
 f.tables.Content_Verification.headers=['Assignment ID','Date','University Name','Assigned To','Member ID','Version','Last Updated','Completed At',...names.flatMap(n=>[n+' Status',n+' Remarks'])];
 f.tables.Content_Verification.rows=[{_row:2,'Assignment ID':'CONTENT-1',Date:'2026-10-15','University Name':'Test University','Assigned To':'Sahil','Member ID':'MEM-SAHIL',Version:1}];
 let r=f.context.saveContent_(f.user,{assignmentId:'CONTENT-1',version:1,operationId:randomUUID(),pages:names.map((name,i)=>({name,status:'Pending',remarks:i===0?'Checked content':''}))});assert.equal(r.assignment.pages[0].status,'Done');assert.equal(r.assignment.completed,1);
 r=f.context.saveContent_(f.user,{assignmentId:'CONTENT-1',version:2,operationId:randomUUID(),pages:names.map(name=>({name,status:'Done',remarks:'  '}))});assert.equal(r.assignment.completed,0);assert(r.assignment.pages.every(p=>p.status==='Pending'));
});


test('exam reviews require all nine sections and save independent exam columns',()=>{
 const f=fixture(),names=['Overview','Eligibility','Application Form','Admit Card','Exam Day Guidelines','Exam Pattern','Syllabus','Highlights','PYQs'];
 const row={_row:2,'Assignment ID':'EXAM-2026-0001',Date:'2026-10-22','University Name':'JEE Main','Assigned To':'Sahil','Member ID':'MEM-SAHIL',Version:1,'Content Type':'Exam','Overview Remarks':'keep university field'};
 f.tables.Content_Verification.headers=['Assignment ID','Date','University Name','Assigned To','Member ID','Version','Last Updated','Completed At','Content Type','Overview Remarks',...names.flatMap(n=>['Exam '+n+' Status','Exam '+n+' Remarks'])];
 f.tables.Content_Verification.rows=[row];
 const payload={assignmentId:row['Assignment ID'],version:1,operationId:randomUUID(),pages:names.map(name=>({name,status:'Pending',remarks:'Verified '+name}))};
 assert.throws(()=>f.context.saveContent_(f.user,{...payload,pages:payload.pages.slice(0,6)}),/All 9 subpages/);
 assert.throws(()=>f.context.saveContent_(f.user,{...payload,pages:payload.pages.map((p,i)=>i===8?{...p,name:'Overview'}:p)}),/All 9 subpages/);
 assert.throws(()=>f.context.saveContent_(f.user,{...payload,pages:payload.pages.map((p,i)=>i===8?{...p,name:'Placement'}:p)}),/Invalid subpage/);
 assert.throws(()=>f.context.saveContent_({...f.user,memberId:'OTHER'},payload),e=>e.status===403);
 const r=f.context.saveContent_(f.user,payload);
 assert.equal(r.assignment.contentType,'Exam');assert.equal(r.assignment.total,9);assert.equal(r.assignment.completed,9);assert.equal(r.assignment.progress,1);assert.equal(r.assignment.status,'Completed');
 assert.equal(row['Exam PYQs Remarks'],'Verified PYQs');assert.equal(row['Overview Remarks'],'keep university field');
 assert.equal(f.context.saveContent_(f.user,payload).duplicate,true);
 assert.throws(()=>f.context.saveContent_(f.user,{...payload,operationId:randomUUID()}),e=>e.status===409);
 const cleared=f.context.saveContent_(f.user,{...payload,version:2,operationId:randomUUID(),pages:payload.pages.map((p,i)=>({...p,remarks:i===8?'':p.remarks}))});
 assert.equal(cleared.assignment.completed,8);assert.equal(cleared.assignment.status,'In Progress');assert.equal(row['Exam PYQs Status'],'Pending');
});
