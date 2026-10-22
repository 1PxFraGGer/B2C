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
  Backlink_Assignments:{headers:['Assignment ID','Date','Project','Source URL','Assigned To','Member ID','Required Platforms','Completed Platforms','Progress Percentage','Overall Status','Last Updated','Version','Completed At','Request ID'],rows:[]},
  Backlink_Submissions:{headers:['Submission ID','Assignment ID','Member ID','Platform','Submission URL','Remarks','Status','Last Updated','Version','Completed At'],rows:[]},
  Activity_Log:{headers:['Timestamp','Employee','Team','University','Task','Previous Value','Updated Value','Action','Event ID','Assignment ID','Item ID','Member ID','Request ID'],rows:[]}
 };
 const ids={Team_Members:3,USP_Ranking_Assignments:1,USP_Ranking_Items:2,Content_Verification:0,Activity_Log:4,Backlink_Assignments:6,Backlink_Submissions:7};
 const files=[],folders=[],props=new Map();
 function folder(id,name,parent){return {getId:()=>id,isTrashed:()=>false,getSharingAccess:()=> 'PRIVATE',getParents:()=>iterator(parent?[parent]:[]),createFolder(n){const f=folder('folder-'+(folders.length+1),n,this);folders.push({name:n,folder:f});return f;},createFile(blob){const file={id:'file-'+(files.length+1),trashed:false,getId(){return this.id;},getUrl(){return 'https://drive.google.com/file/d/'+this.id+'/view';},setTrashed(v){this.trashed=v;},isTrashed(){return this.trashed;},getSize:()=>100,getParents:()=>iterator([this]),getBlob:()=>blob};files.push(file);return file;}};}
 function iterator(arr){let i=0;return {hasNext:()=>i<arr.length,next:()=>arr[i++]};}
 const root=folder('root','Root'),oldFolder=folder('old-folder','NAAC',root),old={id:'old-file',trashed:false,getId(){return this.id;},getSize:()=>100,isTrashed(){return this.trashed;},setTrashed(v){this.trashed=v;},getParents:()=>iterator([oldFolder])};
 const context=vm.createContext({console,Date,Set,Map,JSON,Number,String,Array,Object,Math,RegExp,Error,
  PropertiesService:{getScriptProperties:()=>({getProperty:k=>props.get(k)||null,setProperty:(k,v)=>props.set(k,v),getProperties:()=>Object.fromEntries(props),deleteProperty:k=>props.delete(k)})},
  Utilities:{getUuid:randomUUID,formatDate:d=>d.toISOString().slice(0,10),computeDigest:(algorithm,v)=>[...Buffer.from(String(v)).subarray(0,30)],DigestAlgorithm:{SHA_256:'sha256'},Charset:{UTF_8:'utf8'},base64Decode:v=>[...Buffer.from(v,'base64')],newBlob:(bytes,mime,name)=>({bytes,mime,name,getDataAsString:()=>Buffer.from(bytes).toString()})},
  DriveApp:{Access:{ANYONE:'ANYONE',ANYONE_WITH_LINK:'ANYONE_WITH_LINK'},getFolderById:id=>id==='1FJ1jHU-koTlzKdo9qKw7-lbU88uFzjJH'?root:folders.find(f=>f.folder.getId()===id)?.folder||oldFolder,getFileById:id=>id==='old-file'?old:files.find(f=>f.id===id)},
  Sheets:{Spreadsheets:{batchUpdate({requests}){
   if(fail)throw Error('Sheets unavailable');
   for(const request of requests){if(request.updateCells){const r=request.updateCells,t=Object.values(tables).find(t=>t.sheet.getSheetId()===r.range.sheetId),record=t.rows.find(row=>row._row===r.range.startRowIndex+1);r.rows[0].values.forEach((cell,c)=>{const v=cell.userEnteredValue;record[t.headers[c]]=v.numberValue??v.stringValue??'';});}else if(request.appendCells){const r=request.appendCells,t=Object.values(tables).find(t=>t.sheet.getSheetId()===r.sheetId);for(const row of r.rows)t.rows.push({_row:t.rows.length+2,...Object.fromEntries(t.headers.map((h,i)=>[h,row.values[i].userEnteredValue.numberValue??row.values[i].userEnteredValue.stringValue??'']))});}}
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

function backlinkFixture(options){
 const f=fixture(options);
 for(const name of ['Gauranshi','Anukriti'])f.tables.Team_Members.rows.push({_row:f.tables.Team_Members.rows.length+2,'Member ID':'MEM-'+name.toUpperCase(),'Employee Name':name,Team:'Backlink Team','Active Status':'Active'});
 f.admin={id:'ADMIN',role:'admin',name:'Admin'};
 f.gauranshi={id:'MEM-GAURANSHI',memberId:'MEM-GAURANSHI',name:'Gauranshi',role:'employee',team:'Backlink Team'};
 f.anukriti={id:'MEM-ANUKRITI',memberId:'MEM-ANUKRITI',name:'Anukriti',role:'employee',team:'Backlink Team'};
 f.create=(memberId='MEM-GAURANSHI',sourceUrl='https://www.geduconnect.in/university/test')=>f.context.createBacklink_(f.admin,{memberId,sourceUrl,operationId:randomUUID()}).assignment;
 return f;
}
test('admin assigns one source link and exactly the member-specific three platforms',()=>{
 const f=backlinkFixture(),op=randomUUID(),p={memberId:'MEM-GAURANSHI',sourceUrl:'https://www.geduconnect.in/university/test',operationId:op};
 assert.throws(()=>f.context.createBacklink_(f.gauranshi,p),e=>e.status===403);
 assert.throws(()=>f.context.createBacklink_(f.admin,{...p,sourceUrl:'https://other.example/test'}),/GEDUCONNECT/);
 const first=f.context.createBacklink_(f.admin,p);
 assert.deepEqual(Array.from(first.assignment.platforms),['LinkedIn','X','Threads']);assert.equal(first.assignment.total,3);assert.equal(first.assignment.completed,0);
 assert.equal(f.tables.Backlink_Assignments.rows.length,1);assert.equal(f.tables.Backlink_Submissions.rows.length,3);
 assert.equal(f.context.createBacklink_(f.admin,p).duplicate,true);assert.equal(f.tables.Backlink_Assignments.rows.length,1);
 assert.throws(()=>f.context.createBacklink_(f.admin,{...p,operationId:randomUUID()}),e=>e.status===409);
 const second=f.create('MEM-ANUKRITI');
 assert.deepEqual(Array.from(second.platforms),['Diigo','Blogger','Tumblr']);assert.equal(f.tables.Backlink_Submissions.rows.length,6);
 assert.throws(()=>f.context.createBacklink_(f.admin,{...p,memberId:'MEM-SAHIL',operationId:randomUUID()}));
});
test('each member can only read and edit their own assignments and allowed platforms',()=>{
 const f=backlinkFixture(),a=f.create(),b=f.create('MEM-ANUKRITI');
 const p={assignmentId:a.id,platform:'LinkedIn',submissionUrl:'https://www.linkedin.com/posts/example',remarks:'',version:1,operationId:randomUUID()};
 assert.throws(()=>f.context.saveBacklink_(f.anukriti,p),e=>e.status===403);
 assert.throws(()=>f.context.saveBacklink_(f.gauranshi,{...p,platform:'Blogger'}),e=>e.status===403);
 assert.throws(()=>f.context.saveBacklink_(f.gauranshi,{...p,submissionUrl:'javascript:alert(1)'}));
 assert.throws(()=>f.context.saveBacklink_(f.gauranshi,{...p,submissionUrl:'https://user:pass@example.com/post'}));
 assert.throws(()=>f.context.saveBacklink_(f.gauranshi,{...p,submissionUrl:'https://www.geduconnect.in/university/test'}));
 assert.equal(f.context.workspace_(f.gauranshi,false).backlinks.length,1);
 assert.equal(f.context.workspace_(f.gauranshi,false).backlinks[0].id,a.id);
 assert.equal(f.context.workspace_(f.anukriti,false).backlinks[0].id,b.id);
 assert.equal(f.context.workspace_(f.admin,false).backlinks.length,2);
});
test('three distinct platform submissions complete a source and clearing one reopens it',()=>{
 const f=backlinkFixture();let a=f.create();
 for(const [i,platform]of ['LinkedIn','X','Threads'].entries()){
  const p={assignmentId:a.id,platform,submissionUrl:'https://example.com/submission/'+i,remarks:'Published',version:a.version,operationId:randomUUID()};
  a=f.context.saveBacklink_(f.gauranshi,p).assignment;assert.equal(a.completed,i+1);
  assert.equal(f.context.saveBacklink_(f.gauranshi,p).duplicate,true);
  assert.throws(()=>f.context.saveBacklink_(f.gauranshi,{...p,operationId:randomUUID()}),e=>e.status===409);
 }
 assert.equal(a.status,'Completed');assert.equal(a.progress,1);assert(a.completedAt);
 a=f.context.saveBacklink_(f.admin,{assignmentId:a.id,platform:'X',submissionUrl:'',remarks:'Needs a replacement',version:a.version,operationId:randomUUID()}).assignment;
 assert.equal(a.completed,2);assert.equal(a.status,'In Progress');assert.equal(a.completedAt,'');
 assert.equal(f.tables.Backlink_Assignments.rows[0]['Completed Platforms'],2);
 assert.equal(f.tables.Backlink_Submissions.rows.find(r=>r.Platform==='X').Status,'Pending');
 assert.equal(f.context.workspace_(f.admin,true).teams.backlink.inProgress,1);
});
test('failed backlink writes leave the source and all submissions unchanged',()=>{
 const f=backlinkFixture({fail:true});assert.throws(()=>f.create());assert.equal(f.tables.Backlink_Assignments.rows.length,0);assert.equal(f.tables.Backlink_Submissions.rows.length,0);
});
