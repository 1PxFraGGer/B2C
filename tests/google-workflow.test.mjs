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
  Backlink_Assignments:{headers:['Assignment ID','Date','Project','Source URL','Assigned To','Member ID','Required Platforms','Completed Platforms','Progress Percentage','Overall Status','Last Updated','Version','Completed At','Request ID','University Name','Page Name','Assignment Type','Exam Name','Section','Link Number','Exam Group ID'],rows:[]},
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
 for(const name of ['Gauranshi','Payal','Yash','Arhan'])f.tables.Team_Members.rows.push({_row:f.tables.Team_Members.rows.length+2,'Member ID':'MEM-'+name.toUpperCase(),'Employee Name':name,Team:'Backlink Team','Active Status':'Active'});
 f.admin={id:'ADMIN',role:'admin',name:'Admin'};
 f.gauranshi={id:'MEM-GAURANSHI',memberId:'MEM-GAURANSHI',name:'Gauranshi',role:'employee',team:'Backlink Team'};
 f.payal={id:'MEM-PAYAL',memberId:'MEM-PAYAL',name:'Payal',role:'employee',team:'Backlink Team'};
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
 const second=f.create('MEM-PAYAL');
 assert.deepEqual(Array.from(second.platforms),['Diigo','Blogger','Tumblr']);assert.equal(f.tables.Backlink_Submissions.rows.length,6);
 assert.throws(()=>f.context.createBacklink_(f.admin,{...p,memberId:'MEM-SAHIL',operationId:randomUUID()}));
});
test('each member can only read and edit their own assignments and allowed platforms',()=>{
 const f=backlinkFixture(),a=f.create(),b=f.create('MEM-PAYAL');
 const p={assignmentId:a.id,platform:'LinkedIn',submissionUrl:'https://www.linkedin.com/posts/example',remarks:'',version:1,operationId:randomUUID()};
 assert.throws(()=>f.context.saveBacklink_(f.payal,p),e=>e.status===403);
 assert.throws(()=>f.context.saveBacklink_(f.gauranshi,{...p,platform:'Blogger'}),e=>e.status===403);
 assert.throws(()=>f.context.saveBacklink_(f.gauranshi,{...p,submissionUrl:'javascript:alert(1)'}));
 assert.throws(()=>f.context.saveBacklink_(f.gauranshi,{...p,submissionUrl:'https://user:pass@example.com/post'}));
 assert.throws(()=>f.context.saveBacklink_(f.gauranshi,{...p,submissionUrl:'https://www.geduconnect.in/university/test'}));
 assert.equal(f.context.workspace_(f.gauranshi,false).backlinks.length,1);
 assert.equal(f.context.workspace_(f.gauranshi,false).backlinks[0].id,a.id);
 assert.equal(f.context.workspace_(f.payal,false).backlinks[0].id,b.id);
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

test('a university creates all six exact .com links for both members atomically',()=>{
 const f=backlinkFixture(),p={university:'VGU Online',sourceUrl:'https://geduconnect.com/university/vgu-online',date:'2026-10-22',operationId:randomUUID()};
 assert.throws(()=>f.context.createBacklinkUniversity_(f.gauranshi,p),e=>e.status===403);
 const result=f.context.createBacklinkUniversity_(f.admin,p);assert.equal(result.assignments.length,12);assert.equal(f.tables.Backlink_Submissions.rows.length,36);
 const suffixes=['','/program-fees','/admission-process','/placement-career-support','/student-support','/faqs'];
 for(const id of ['MEM-GAURANSHI','MEM-PAYAL']){const rows=result.assignments.filter(r=>r.memberId===id);assert.deepEqual(Array.from(rows,r=>r.sourceUrl),suffixes.map(s=>p.sourceUrl+s));assert(rows.every(r=>r.university==='VGU Online'&&r.date===p.date));assert.equal(new Set(rows.map(r=>r.page)).size,6);}
 assert.equal(f.context.createBacklinkUniversity_(f.admin,p).duplicate,true);assert.equal(f.tables.Backlink_Assignments.rows.length,12);
 assert.throws(()=>f.context.createBacklinkUniversity_(f.admin,{...p,operationId:randomUUID()}),e=>e.status===409);
 assert.throws(()=>f.context.createBacklinkUniversity_(f.admin,{...p,sourceUrl:'https://geduconnect.com.evil.example/university/vgu-online',operationId:randomUUID()}));
 assert.throws(()=>f.context.createBacklinkUniversity_(f.admin,{...p,sourceUrl:p.sourceUrl+'/faqs',operationId:randomUUID()}));
 assert.throws(()=>f.context.createBacklinkUniversity_(f.admin,{...p,date:'2026-02-30',operationId:randomUUID()}));
});
test('daily allocation permits two universities and rejects a third without partial rows',()=>{
 const f=backlinkFixture();for(const slug of ['one','two'])f.context.createBacklinkUniversity_(f.admin,{university:slug,sourceUrl:'https://geduconnect.com/university/'+slug,date:'2026-10-22',operationId:randomUUID()});
 assert.equal(f.tables.Backlink_Assignments.rows.length,24);assert.equal(f.tables.Backlink_Submissions.rows.length,72);
 assert.throws(()=>f.context.createBacklinkUniversity_(f.admin,{university:'three',sourceUrl:'https://geduconnect.com/university/three',date:'2026-10-22',operationId:randomUUID()}),/Two universities/);assert.equal(f.tables.Backlink_Assignments.rows.length,24);
 f.context.createBacklinkUniversity_(f.admin,{university:'three',sourceUrl:'https://geduconnect.com/university/three',date:'2026-10-23',operationId:randomUUID()});assert.equal(f.tables.Backlink_Assignments.rows.length,36);
});
test('university creation fails as one batch and .com sources retain the original platform save protections',()=>{
 const f=backlinkFixture({fail:true});assert.throws(()=>f.context.createBacklinkUniversity_(f.admin,{university:'VGU Online',sourceUrl:'https://geduconnect.com/university/vgu-online',date:'2026-10-22',operationId:randomUUID()}));assert.equal(f.tables.Backlink_Assignments.rows.length,0);assert.equal(f.tables.Backlink_Submissions.rows.length,0);
 const ok=backlinkFixture(),r=ok.create('MEM-GAURANSHI','https://geduconnect.com/university/vgu-online');assert.equal(r.total,3);assert.equal(r.sourceUrl,'https://geduconnect.com/university/vgu-online');
});

function examBacklinkPayload(){return {exam:'JEE Main',date:'2026-10-22',articles:Array.from({length:6},(_,i)=>'https://example.com/article/'+(i+1)),news:Array.from({length:4},(_,i)=>'https://example.com/news/'+(i+1)),operationId:randomUUID()};}
test('exam assignment creates 6 Articles and 4 News for each member with 30 and 34 submissions',()=>{
 const f=backlinkFixture(),p=examBacklinkPayload();assert.throws(()=>f.context.createBacklinkExam_(f.gauranshi,p),e=>e.status===403);
 const r=f.context.createBacklinkExam_(f.admin,p);assert.equal(r.assignments.length,30);assert.equal(f.tables.Backlink_Submissions.rows.length,98);
 for(const [id,count] of [['MEM-YASH',30],['MEM-ARHAN',34],['MEM-PAYAL',34]]){const rows=r.assignments.filter(r=>r.memberId===id);assert.equal(rows.length,10);assert.equal(rows.filter(r=>r.section==='Articles').length,6);assert.equal(rows.filter(r=>r.section==='News').length,4);assert.equal(rows.reduce((n,r)=>n+r.total,0),count);assert(rows.every(r=>r.assignmentType==='Exam'&&r.exam==='JEE Main'));}
 assert.equal(new Set(r.assignments.map(r=>r.examGroupId)).size,1);
 assert(r.assignments.filter(r=>r.memberId==='MEM-YASH').every(r=>!r.platforms.includes('Classified Ads')));
 assert(r.assignments.filter(r=>r.memberId==='MEM-ARHAN'&&r.section==='Articles').every(r=>!r.platforms.includes('Classified Ads')));
 assert(r.assignments.filter(r=>r.memberId==='MEM-ARHAN'&&r.section==='News').every(r=>r.platforms.includes('Classified Ads')));
 assert.equal(f.context.createBacklinkExam_(f.admin,p).duplicate,true);assert.equal(f.tables.Backlink_Assignments.rows.length,30);
 assert.throws(()=>f.context.createBacklinkExam_(f.admin,{...p,operationId:randomUUID()}),e=>e.status===409);
});
test('Classified Ads is rejected for Articles and Yash; Arhan News needs all four submissions',()=>{
 const f=backlinkFixture(),r=f.context.createBacklinkExam_(f.admin,examBacklinkPayload()),arhan={memberId:'MEM-ARHAN',role:'employee',team:'Backlink Team',name:'Arhan'},yash={...arhan,memberId:'MEM-YASH',name:'Yash'};
 const article=r.assignments.find(r=>r.memberId===arhan.memberId&&r.section==='Articles'),news=r.assignments.find(r=>r.memberId===arhan.memberId&&r.section==='News'),yashNews=r.assignments.find(r=>r.memberId===yash.memberId&&r.section==='News');
 const payload=(row,platform,version=1)=>({assignmentId:row.id,platform,version,submissionUrl:'https://published.example/'+platform.replaceAll(' ','-'),remarks:'',operationId:randomUUID()});
 assert.throws(()=>f.context.saveBacklink_(arhan,payload(article,'Classified Ads')),e=>e.status===403);assert.throws(()=>f.context.saveBacklink_(yash,payload(yashNews,'Classified Ads')),e=>e.status===403);assert.throws(()=>f.context.saveBacklink_(yash,payload(news,'Blogger')),e=>e.status===403);
 let updated;for(const [i,platform] of ['Tumblr','Blogger','Diigo'].entries())updated=f.context.saveBacklink_(arhan,payload(news,platform,i+1)).assignment;
 assert.equal(updated.completed,3);assert.equal(updated.total,4);assert.equal(updated.status,'In Progress');updated=f.context.saveBacklink_(arhan,payload(news,'Classified Ads',4)).assignment;assert.equal(updated.status,'Completed');assert.equal(updated.completed,4);
 assert.equal(f.context.saveBacklink_(arhan,{...payload(news,'Classified Ads',5),submissionUrl:''}).assignment.status,'In Progress');
});
test('exam links reject empty or oversized sections, duplicate sources, invalid dates and unsafe URLs before writing',()=>{
 const f=backlinkFixture(),p=examBacklinkPayload();for(const invalid of [{...p,articles:[]},{...p,news:[]},{...p,articles:[...p.articles,'https://example.com/extra']},{...p,news:[...p.news,'https://example.com/extra']},{...p,news:[p.articles[0],...p.news.slice(1)]},{...p,date:'2026-02-30'},{...p,articles:['javascript:alert(1)',...p.articles.slice(1)]}])assert.throws(()=>f.context.createBacklinkExam_(f.admin,invalid));assert.equal(f.tables.Backlink_Assignments.rows.length,0);
 const failed=backlinkFixture({fail:true});assert.throws(()=>failed.context.createBacklinkExam_(failed.admin,p));assert.equal(failed.tables.Backlink_Assignments.rows.length,0);assert.equal(failed.tables.Backlink_Submissions.rows.length,0);
});

test('CLAT with five articles and three news links produces no empty slots or inflated totals',()=>{
 const f=backlinkFixture(),p=examBacklinkPayload();p.exam='CLAT';p.articles=p.articles.slice(0,5);p.news=p.news.slice(0,3);
 const r=f.context.createBacklinkExam_(f.admin,p);assert.equal(r.assignments.length,24);assert.equal(f.tables.Backlink_Submissions.rows.length,78);
 for(const [id,total] of [['MEM-YASH',24],['MEM-ARHAN',27],['MEM-PAYAL',27]]){const rows=r.assignments.filter(r=>r.memberId===id);assert.equal(rows.filter(r=>r.section==='Articles').length,5);assert.equal(rows.filter(r=>r.section==='News').length,3);assert.equal(rows.reduce((n,r)=>n+r.total,0),total);assert(rows.every(r=>r.sourceUrl));}
 assert.equal(f.context.createBacklinkExam_(f.admin,p).duplicate,true);assert.equal(f.tables.Backlink_Submissions.rows.length,78);
});

function customUniversityPages(){return [['Overview','https://geduconnect.com/university/custom-main'],['Program & Fees','https://geduconnect.com/custom-fees'],['Admission process','https://geduconnect.com/admissions/custom'],['Placement & Career Support','https://geduconnect.com/career-services'],['Student Support','https://geduconnect.com/help/custom'],['FAQs','https://geduconnect.com/questions-custom']].map(([name,url])=>({name,url}));}
test('custom university source links are saved exactly and daily counting uses university identity',()=>{
 const f=backlinkFixture(),p={university:'Custom University',date:'2026-10-23',pages:customUniversityPages(),operationId:randomUUID()};
 const r=f.context.createBacklinkUniversity_(f.admin,p);assert.equal(r.assignments.length,12);for(const id of ['MEM-GAURANSHI','MEM-PAYAL'])assert.deepEqual(Array.from(r.assignments.filter(r=>r.memberId===id),r=>r.sourceUrl),p.pages.map(p=>p.url));
 f.context.createBacklinkUniversity_(f.admin,{university:'Second',date:p.date,sourceUrl:'https://geduconnect.com/university/second',operationId:randomUUID()});
 assert.throws(()=>f.context.createBacklinkUniversity_(f.admin,{university:'Third',date:p.date,sourceUrl:'https://geduconnect.com/university/third',operationId:randomUUID()}),/Two universities/);
});
test('admin link corrections update both members, retain submissions, and reject stale or unauthorized edits',()=>{
 const f=backlinkFixture(),created=f.context.createBacklinkUniversity_(f.admin,{university:'VGU Online',date:'2026-10-23',sourceUrl:'https://geduconnect.com/university/vgu-online',operationId:randomUUID()}).assignments;
 const a=created[0];f.context.saveBacklink_(f.gauranshi,{assignmentId:a.id,platform:'LinkedIn',version:1,submissionUrl:'https://linkedin.com/posts/existing',remarks:'Keep this',operationId:randomUUID()});
 const p={assignmentId:a.id,pages:customUniversityPages(),versions:f.tables.Backlink_Assignments.rows.map(r=>({id:r['Assignment ID'],version:r.Version})),operationId:randomUUID()};
 assert.throws(()=>f.context.saveUniversitySourceLinks_(f.gauranshi,p),e=>e.status===403);
 assert.throws(()=>f.context.saveUniversitySourceLinks_(f.admin,{...p,versions:p.versions.map(v=>({...v,version:0}))}),e=>e.status===409);
 assert.equal(f.context.saveUniversitySourceLinks_(f.admin,p).updated,12);assert(f.tables.Backlink_Submissions.rows.some(r=>r['Submission URL']==='https://linkedin.com/posts/existing'&&r.Remarks==='Keep this'));
 for(const id of ['MEM-GAURANSHI','MEM-PAYAL'])assert.deepEqual(f.tables.Backlink_Assignments.rows.filter(r=>r['Member ID']===id).map(r=>r['Source URL']),p.pages.map(p=>p.url));
 assert.equal(f.context.saveUniversitySourceLinks_(f.admin,p).duplicate,true);
 assert.throws(()=>f.context.saveUniversitySourceLinks_(f.admin,{...p,operationId:randomUUID()}),e=>e.status===409);
});
test('source link corrections reject duplicate, incomplete and non-GEDUCONNECT URLs before writing',()=>{
 const f=backlinkFixture(),created=f.context.createBacklinkUniversity_(f.admin,{university:'VGU Online',date:'2026-10-23',sourceUrl:'https://geduconnect.com/university/vgu-online',operationId:randomUUID()}).assignments;
 const p={assignmentId:created[0].id,versions:created.map(r=>({id:r.id,version:r.version})),pages:customUniversityPages(),operationId:randomUUID()};
 for(const pages of [p.pages.slice(1),p.pages.map((r,i)=>i===1?{...r,url:p.pages[0].url}:r),p.pages.map((r,i)=>i===1?{...r,url:'https://geduconnect.com.evil.example/test'}:r)])assert.throws(()=>f.context.saveUniversitySourceLinks_(f.admin,{...p,pages}));
 assert.equal(f.tables.Backlink_Assignments.rows[0]['Source URL'],'https://geduconnect.com/university/vgu-online');
});

test('Payal News supports Classified Ads and retains three university platforms',()=>{const f=backlinkFixture(),r=f.context.createBacklinkExam_(f.admin,examBacklinkPayload()),own=r.assignments.filter(r=>r.memberId==='MEM-PAYAL');assert(own.filter(r=>r.section==='Articles').every(r=>r.total===3&&!r.platforms.includes('Classified Ads')));const news=own.find(r=>r.section==='News');assert.equal(news.total,4);const result=f.context.saveBacklink_(f.payal,{assignmentId:news.id,platform:'Classified Ads',submissionUrl:'https://example.com/classified',remarks:'',version:1,operationId:randomUUID()});assert.equal(result.assignment.completed,1);assert.equal(f.create('MEM-PAYAL','https://geduconnect.com/university/retained').total,3);});

function examEditFixture(){const f=backlinkFixture(),created=f.context.createBacklinkExam_(f.admin,examBacklinkPayload()).assignments;f.context.saveBacklink_(f.payal,{assignmentId:created.find(r=>r.memberId==='MEM-PAYAL').id,platform:'Diigo',version:1,submissionUrl:'https://example.com/existing',remarks:'Keep this',operationId:randomUUID()});const links=created.filter(r=>r.memberId==='MEM-YASH').map(r=>({section:r.section,linkNumber:r.linkNumber,url:r.sourceUrl}));links[0].url='https://example.com/corrected-article';return {...f,p:{assignmentId:created[0].id,links,versions:f.tables.Backlink_Assignments.rows.map(r=>({id:r['Assignment ID'],version:r.Version})),operationId:randomUUID()}};}
test('admin exam source correction updates matching slots for all owners and preserves submissions and progress',()=>{const f=examEditFixture(),before=JSON.stringify(f.tables.Backlink_Submissions.rows),progress=f.tables.Backlink_Assignments.rows.map(r=>[r['Completed Platforms'],r['Overall Status']]);assert.throws(()=>f.context.saveExamSourceLinks_(f.payal,f.p),e=>e.status===403);assert.equal(f.context.saveExamSourceLinks_(f.admin,f.p).updated,3);assert.equal(JSON.stringify(f.tables.Backlink_Submissions.rows),before);assert.deepEqual(f.tables.Backlink_Assignments.rows.map(r=>[r['Completed Platforms'],r['Overall Status']]),progress);assert.equal(f.tables.Backlink_Assignments.rows.filter(r=>r['Source URL']===f.p.links[0].url).length,3);assert.equal(f.context.saveExamSourceLinks_(f.admin,f.p).duplicate,true);assert.throws(()=>f.context.saveExamSourceLinks_(f.admin,{...f.p,operationId:randomUUID()}),e=>e.status===409);});
test('exam source correction rejects incomplete, duplicate, unsafe and stale input without partial writes',()=>{for(const change of [p=>({...p,links:p.links.slice(1)}),p=>({...p,links:p.links.map((l,i)=>i===1?{...l,url:p.links[0].url}:l)}),p=>({...p,links:p.links.map((l,i)=>i===0?{...l,url:'javascript:alert(1)'}:l)}),p=>({...p,versions:p.versions.map(v=>({...v,version:0}))})]){const f=examEditFixture(),before=JSON.stringify(f.tables.Backlink_Assignments.rows);assert.throws(()=>f.context.saveExamSourceLinks_(f.admin,change(f.p)));assert.equal(JSON.stringify(f.tables.Backlink_Assignments.rows),before);}});

test('member sessions use the current team and inactive accounts are rejected',()=>{const f=backlinkFixture(),row=f.tables.Team_Members.rows.find(r=>r['Member ID']==='MEM-PAYAL');row['Auth Subject']='MEM-PAYAL';row['Login ID']='payal@example.com';const identity={id:'MEM-PAYAL',role:'employee',loginId:'payal@example.com'};assert.equal(f.context.profile_(identity).team,'Backlink Team');row['Active Status']='Inactive';assert.throws(()=>f.context.profile_(identity),e=>e.status===403);});
