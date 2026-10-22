/* University dashboard Google backend. Deploy as owner; only HMAC-signed POSTs are accepted. */
const SHEET_ID = '160hjCLAHaSTU_TedLYla_BqrMXXgbcaiJNaoEL-_l7k';
const ROOT_FOLDER_ID = '1FJ1jHU-koTlzKdo9qKw7-lbU88uFzjJH';
const MAX_BYTES = 2 * 1024 * 1024;
let requestTables_ = {}, requestBook_ = null;
const PAGES = ['Overview','Program & Fees','Admission','Placement','Student Support','FAQs'];
const EXAM_PAGES = ['Overview','Eligibility','Application Form','Admit Card','Exam Day Guidelines','Exam Pattern','Syllabus','Highlights','PYQs'];
function isExam_(row){return row['Content Type']==='Exam';}
function contentPages_(row){return isExam_(row)?EXAM_PAGES:PAGES;}
function contentColumn_(row,name,suffix){return (isExam_(row)?'Exam ':'')+name+' '+suffix;}
function failure_(message,status){const error=new Error(message);error.status=status||400;throw error;}
function json_(value){return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);}
function doGet(){return json_({ok:false,error:'Authenticated POST required',status:405});}
function doPost(e){
  requestTables_ = {}; requestBook_ = null;
  let lock;
  try{
    if(!e||!e.postData||e.postData.contents.length>3000000)failure_('Invalid request',413);
    const envelope=JSON.parse(e.postData.contents),secret=PropertiesService.getScriptProperties().getProperty('BRIDGE_SECRET');
    if(!secret||secret.length<32)failure_('Backend secret is not configured',503);
    if(typeof envelope.body!=='string'||typeof envelope.signature!=='string')failure_('Unauthorized',401);
    const expected=Utilities.base64Encode(Utilities.computeHmacSha256Signature(envelope.body,secret,Utilities.Charset.UTF_8));
    if(!equal_(expected,envelope.signature))failure_('Unauthorized',401);
    const request=JSON.parse(envelope.body);
    if(!Number.isFinite(request.timestamp)||Math.abs(Date.now()-request.timestamp)>120000||!/^[-\w]{20,80}$/.test(request.requestId||''))failure_('Expired request',401);
    const candidate=LockService.getScriptLock();if(!candidate.tryLock(10000))failure_('Another save is in progress. Please retry.',503);lock=candidate;
    const cache=CacheService.getScriptCache(),nonce='nonce:'+request.requestId;
    if(cache.get(nonce))failure_('Request already processed',409);cache.put(nonce,'1',180);
    const action=request.action,p=request.payload||{};
    // Reserve the nonce atomically, then allow independent reads and previews concurrently.
    if(['sessionRead','data','reports','image'].includes(action)){lock.releaseLock();lock=null;}
    const names=['Team_Members','Admin_Users'];
    if(['data','reports','saveItem','saveContent','createBacklink','saveBacklink','createBacklinkUniversity','createBacklinkExam'].includes(action))names.push('Content_Verification','USP_Ranking_Assignments','USP_Ranking_Items','Activity_Log','Backlink_Assignments','Backlink_Submissions');
    else if(action==='image')names.push('USP_Ranking_Assignments','USP_Ranking_Items');
    if(action!=='loginRate')primeTables_(names);
    if(action==='loginRate')return json_({ok:true,data:rate_(p.loginId)});
    if(action==='sessionCreate')return json_({ok:true,data:createSession_(p)});
    const user=session_(request.session);
    let data;
    if(action==='sessionRead')data=user;
    else if(action==='sessionDelete'){PropertiesService.getScriptProperties().deleteProperty('session:'+digest_(request.session));data={ok:true};}
    else if(action==='data'||action==='reports'){if(action==='reports'&&user.role!=='admin')failure_('Admin access required',403);data=workspace_(user,action==='reports');}
    else if(action==='saveItem')data=saveItem_(user,p);
    else if(action==='saveContent')data=saveContent_(user,p);
    else if(action==='createBacklink')data=createBacklink_(user,p);
    else if(action==='createBacklinkUniversity')data=createBacklinkUniversity_(user,p);
    else if(action==='createBacklinkExam')data=createBacklinkExam_(user,p);
    else if(action==='saveBacklink')data=saveBacklink_(user,p);
    else if(action==='image')data=image_(user,p.itemId);
    else failure_('Unknown action',404);
    return json_({ok:true,data});
  }catch(error){console.error(error.status?error.message:'Backend operation failed');return json_({ok:false,error:error.status?error.message:'Google operation failed. Refresh before retrying a save.',status:error.status||500});}
  finally{if(lock)lock.releaseLock();}
}
function equal_(a,b){if(a.length!==b.length)return false;let result=0;for(let i=0;i<a.length;i++)result|=a.charCodeAt(i)^b.charCodeAt(i);return result===0;}
function digest_(value){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(value),Utilities.Charset.UTF_8).map(b=>('0'+(b&255).toString(16)).slice(-2)).join('');}
function workbook_(){return requestBook_ || (requestBook_=SpreadsheetApp.openById(SHEET_ID));}
function primeTables_(names){
  const missing=names.filter(name=>!requestTables_[name]);if(!missing.length)return;
  const result=Sheets.Spreadsheets.Values.batchGet(SHEET_ID,{ranges:missing.map(name=>"'"+name+"'!A:"+(name==='Content_Verification'?'AR':'Z')),valueRenderOption:'UNFORMATTED_VALUE',dateTimeRenderOption:'SERIAL_NUMBER'});
  missing.forEach((name,i)=>{requestTables_[name]=tableValues_(name,result.valueRanges[i].values||[]);});
}
function tableValues_(name,values,sheet){
  const headers=values[0]||[];
  const rows=values.slice(1).map((values,i)=>{const row={_row:i+2};headers.forEach((h,j)=>{let v=values[j]===undefined?'':values[j];if(h==='Date'&&typeof v==='number')v=new Date(Math.round((v-25569)*86400000)).toISOString().slice(0,10);row[h]=v instanceof Date?v.toISOString():v;});return row;}).filter(row=>headers.some(h=>row[h]!==''));
  const table={headers,rows};let resolved=sheet;
  Object.defineProperty(table,'sheet',{get(){return resolved || (resolved=workbook_().getSheetByName(name));}});
  return table;
}
function table_(name){
  if(requestTables_[name])return requestTables_[name];
  const sheet=workbook_().getSheetByName(name);if(!sheet)failure_('Missing sheet: '+name,503);
  const values=sheet.getRange(1,1,Math.max(sheet.getLastRow(),1),sheet.getLastColumn()).getValues();
  return requestTables_[name]=tableValues_(name,values,sheet);
}
function required_(record,fields){return fields.every(key=>record[key]!==''&&record[key]!==null&&record[key]!==undefined);}
function key_(value){return String(value==null?'':value);}
function unique_(rows,id,key){const matches=rows.filter(row=>key_(row[key])===key_(id));if(matches.length!==1)failure_(matches.length?'Duplicate record IDs must be corrected':'Record not found',matches.length?409:404);return matches[0];}
function member_(row,members){const found=members.find(m=>key_(m['Member ID'])===key_(row['Member ID'])||!row['Member ID']&&m['Employee Name']===row['Assigned To']);return found?found['Member ID']:'';}
function profile_(identity){
  if(identity.role==='admin'){
    const row=unique_(table_('Admin_Users').rows,identity.id,'Auth Subject');
    if(row['Active Status']!=='Active'||String(row.Email).toLowerCase()!==identity.loginId)failure_('Account is inactive',403);
    return {...identity,name:row['Display Name'],memberId:null};
  }
  if(identity.role!=='employee')failure_('Invalid account',403);
  const row=unique_(table_('Team_Members').rows,identity.id,'Auth Subject');
  if(row['Active Status']!=='Active'||String(row['Login ID']).toLowerCase()!==identity.loginId)failure_('Account is inactive',403);
  return {...identity,name:row['Employee Name'],memberId:row['Member ID'],team:row.Team};
}
function cleanupProperties_(){const props=PropertiesService.getScriptProperties(),all=props.getProperties();Object.keys(all).forEach(k=>{if(/^(session:|rate:)/.test(k)){let r;try{r=JSON.parse(all[k]);}catch{return;}if(r.expires<Date.now())props.deleteProperty(k);}});}
function rate_(loginId){
  if(typeof loginId!=='string'||loginId.length>254)failure_('Invalid login ID');cleanupProperties_();
  const props=PropertiesService.getScriptProperties();
  for(const [key,limit] of [['rate:'+digest_(loginId),12],['rate:global',150]]){
    let r=JSON.parse(props.getProperty(key)||'null');if(!r||r.expires<Date.now())r={count:0,expires:Date.now()+900000};
    if(r.count>=limit)failure_('Too many login attempts. Try again in 15 minutes.',429);r.count++;props.setProperty(key,JSON.stringify(r));
  }return {ok:true};
}
function createSession_(p){
  if(!/^[a-f0-9]{64}$/.test(p.token||''))failure_('Invalid session');
  const identity=profile_(p.identity||{});cleanupProperties_();
  PropertiesService.getScriptProperties().setProperty('session:'+digest_(p.token),JSON.stringify({identity,expires:Date.now()+28800000}));return {ok:true};
}
function session_(token){
  if(typeof token!=='string'||!/^[a-f0-9]{64}$/.test(token))failure_('Please sign in',401);
  const r=JSON.parse(PropertiesService.getScriptProperties().getProperty('session:'+digest_(token))||'null');
  if(!r||r.expires<Date.now())failure_('Session expired. Please sign in again.',401);return profile_(r.identity);
}
function visible_(row,user,members){return user.role==='admin'||member_(row,members)===user.memberId;}
function contentView_(r){
  const pages=contentPages_(r).map(name=>({name,status:r[contentColumn_(r,name,'Status')]==='Done'?'Done':'Pending',remarks:String(r[contentColumn_(r,name,'Remarks')]||'')}));
  const completed=pages.filter(p=>p.status==='Done').length;
  return {id:key_(r['Assignment ID']),date:r.Date,university:r['University Name'],liveLink:String(r['Live Link']||'').trim(),universityCheckLink:String(r['University Check Link']||'').trim(),assignedTo:r['Assigned To'],memberId:r['Member ID'],version:Number(r.Version)||0,contentType:isExam_(r)?'Exam':'University',completed,total:pages.length,progress:completed/pages.length,status:completed===pages.length?'Completed':completed?'In Progress':'Pending',lastUpdated:r['Last Updated'],completedAt:r['Completed At'],pages};
}
function itemView_(r){return {id:key_(r['Item ID']),assignmentId:key_(r['Assignment ID']),university:r['University Name'],assignedTo:r['Assigned To'],type:r.Type,name:r['Item Name'],description:String(r.Description||''),status:r.Status==='Done'?'Done':'Pending',imageFileId:key_(r['Image File ID']),imageUrl:String(r['Image URL']||''),remarks:String(r.Remarks||''),version:Number(r.Version)||0,lastUpdated:r['Last Updated'],completedAt:r['Completed At'],folderId:key_(r['Item Folder ID'])};}
function assignmentView_(r,items){
  const linked=items.filter(i=>i.assignmentId===key_(r['Assignment ID'])),completed=linked.filter(i=>i.status==='Done'&&i.imageFileId).length,total=linked.length;
  return {id:key_(r['Assignment ID']),date:r.Date,university:r['University Name'],assignedTo:r['Assigned To'],memberId:r['Member ID'],version:Number(r.Version)||0,items:linked,completed,total,usp:linked.filter(i=>i.type==='USP').length,ranking:linked.filter(i=>i.type==='Ranking').length,progress:total?completed/total:0,status:total&&completed===total?'Completed':completed?'In Progress':'Pending',lastUpdated:r['Last Updated'],completedAt:r['Completed At']};
}
function totals_(rows){const completed=rows.filter(r=>r.status==='Completed').length;return {assigned:rows.length,completed,pending:rows.filter(r=>r.status==='Pending').length,inProgress:rows.filter(r=>r.status==='In Progress').length,progress:rows.length?completed/rows.length:0};}
function dataTables_(){
  const members=table_('Team_Members').rows.filter(r=>required_(r,['Member ID','Employee Name']));
  const content=table_('Content_Verification').rows.filter(r=>required_(r,['Assignment ID','Date','University Name','Assigned To']));
  const assignments=table_('USP_Ranking_Assignments').rows.filter(r=>required_(r,['Assignment ID','Date','University Name','Assigned To']));
  const items=table_('USP_Ranking_Items').rows.filter(r=>required_(r,['Item ID','Assignment ID','University Name','Assigned To','Type','Item Name'])&&['USP','Ranking'].includes(r.Type));
  return {members,content,assignments,items};
}
function workspace_(user,reports){
  const d=dataTables_(),b=backlinkTables_(),allContent=d.content.map(contentView_),allItems=d.items.map(itemView_),allAssignments=d.assignments.map(r=>assignmentView_(r,allItems)),allBacklinks=b.assignments.map(r=>backlinkView_(r,b.submissions));
  const employeeStats=d.members.map(m=>{const isContent=m.Team==='Content Verification',isBacklink=m.Team==='Backlink Team',records=(isBacklink?b.assignments:isContent?d.content:d.assignments).filter(r=>(isBacklink?key_(r['Member ID']):member_(r,d.members))===m['Member ID']);return {id:m['Member ID'],name:m['Employee Name'],team:m.Team,active:m['Active Status']==='Active',platforms:backlinkPlatforms_(m['Member ID']),...totals_(records.map(r=>isBacklink?backlinkView_(r,b.submissions):isContent?contentView_(r):assignmentView_(r,allItems)))};});
  const activity=table_('Activity_Log').rows.filter(r=>r.Timestamp&&(user.role==='admin'||r['Member ID']===user.memberId)).map(r=>({timestamp:r.Timestamp,employee:r.Employee,team:r.Team,university:r.University,task:r.Task,action:r.Action,assignmentId:key_(r['Assignment ID']),itemId:key_(r['Item ID']),previous:r['Previous Value'],updated:r['Updated Value']}));
  return {user,members:employeeStats,stats:totals_([...allContent,...allAssignments,...allBacklinks]),teams:{content:totals_(allContent),usp:totals_(allAssignments),backlink:totals_(allBacklinks)},
    content:d.content.filter(r=>visible_(r,user,d.members)).map(contentView_),assignments:d.assignments.filter(r=>visible_(r,user,d.members)).map(r=>assignmentView_(r,allItems)),backlinks:allBacklinks.filter(r=>user.role==='admin'||user.team==='Backlink Team'&&r.memberId===user.memberId),activity:reports?activity:activity.slice(-30).reverse()};
}
function itemContext_(user,id){
  const d=dataTables_(),item=unique_(d.items,id,'Item ID'),assignment=unique_(d.assignments,item['Assignment ID'],'Assignment ID');
  if(!visible_(assignment,user,d.members)||!visible_(item,user,d.members))failure_('You cannot edit this assignment',403);
  if(item['University Name']!==assignment['University Name']||item['Assigned To']!==assignment['Assigned To'])failure_('Item university or employee does not match its assignment',409);
  return {d,item,assignment};
}
function safeName_(name){return String(name).normalize('NFKC').replace(/[\\/\x00-\x1f<>:"|?*]/g,'-').replace(/\s+/g,' ').trim().slice(0,90)||'Item';}
function privateRoot_(){const root=DriveApp.getFolderById(ROOT_FOLDER_ID),access=root.getSharingAccess();if(access===DriveApp.Access.ANYONE||access===DriveApp.Access.ANYONE_WITH_LINK)failure_('Set the upload root folder to Restricted before uploading private images.',403);return root;}
function folder_(parent,name,identity){
  const props=PropertiesService.getScriptProperties(),key='folder:'+digest_(parent.getId()+'|'+identity),stored=props.getProperty(key);
  if(stored){try{const f=DriveApp.getFolderById(stored);if(!f.isTrashed()){const parents=f.getParents();while(parents.hasNext())if(parents.next().getId()===parent.getId())return f;}}catch{}}
  const created=parent.createFolder(safeName_(name));props.setProperty(key,created.getId());return created;
}
function destination_(assignment,item){
  const root=privateRoot_(),universityId=assignment['University ID']||'UNI-'+assignment['Assignment ID'];
  const university=folder_(root,assignment['University Name'],'university:'+universityId);
  const type=folder_(university,item.Type,'type:'+item.Type);
  return folder_(type,item['Item Name'],'item:'+item['Item ID']);
}
function validateImage_(image){
  if(!image||typeof image.base64!=='string'||image.base64.length>Math.ceil(MAX_BYTES/3)*4+4||!/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(image.base64))failure_('Invalid image');
  const bytes=Utilities.base64Decode(image.base64),u=bytes.map(b=>b&255),extension=String(image.extension||'').toLowerCase();
  if(!u.length||u.length>MAX_BYTES)failure_('Image exceeds 2 MB',413);
  let mime='';
  if(extension==='png'&&[137,80,78,71,13,10,26,10].every((b,i)=>u[i]===b))mime='image/png';
  if(['jpg','jpeg'].includes(extension)&&u[0]===255&&u[1]===216&&u[u.length-2]===255&&u[u.length-1]===217)mime='image/jpeg';
  if(extension==='webp'&&String.fromCharCode(...u.slice(0,4))==='RIFF'&&String.fromCharCode(...u.slice(8,12))==='WEBP')mime='image/webp';
  if(extension==='svg'){
    const text=Utilities.newBlob(bytes).getDataAsString('UTF-8');
    if(/<!DOCTYPE|<!ENTITY|<script|<foreignObject|\bon\w+\s*=|\bstyle\s*=|javascript:|data:|https?:\/\/(?!www\.w3\.org\/)/i.test(text))failure_('Unsafe SVG');
    let doc;try{doc=XmlService.parse(text);}catch{failure_('Invalid SVG');}
    if(doc.getRootElement().getName()!=='svg')failure_('Invalid SVG');mime='image/svg+xml';
  }
  if(!mime||mime!==image.mime)failure_('Unsupported image format');return {bytes,mime,extension};
}
function existingImage_(item){
  privateRoot_();if(!item['Image File ID']||!item['Item Folder ID'])failure_('Save an image before marking the item Done',422);
  let file;try{file=DriveApp.getFileById(item['Image File ID']);}catch{failure_('Saved image is unavailable',422);}
  if(file.isTrashed()||file.getSize()>MAX_BYTES)failure_('Saved image is unavailable',422);
  const parents=file.getParents();let belongs=false;while(parents.hasNext())if(parents.next().getId()===item['Item Folder ID'])belongs=true;
  if(!belongs)failure_('Image does not belong to this item',403);return file;
}
function image_(user,id){const {item}=itemContext_(user,id),file=existingImage_(item),blob=file.getBlob();return {base64:Utilities.base64Encode(blob.getBytes()),mime:blob.getContentType()};}
function operation_(id){if(typeof id!=='string'||!/^[-\w]{20,80}$/.test(id))failure_('Invalid save operation');}
function seen_(user,id,itemId,assignmentId){return table_('Activity_Log').rows.some(r=>r['Request ID']===id&&r['Member ID']===(user.memberId||user.id)&&key_(r['Item ID'])===key_(itemId)&&key_(r['Assignment ID'])===key_(assignmentId));}
function version_(record,supplied){if(!Number.isInteger(supplied)||supplied!==(Number(record.Version)||0))failure_('This task changed since you opened it. Refresh before saving.',409);}
function rowRequest_(name,record,changes){
  const t=table_(name),data=t.sheet.getRange(record._row,1,1,t.headers.length).getValues()[0],formulas=t.sheet.getRange(record._row,1,1,t.headers.length).getFormulas()[0];
  const values=t.headers.map((h,i)=>{const value=Object.prototype.hasOwnProperty.call(changes,h)?changes[h]:formulas[i]?{formula:formulas[i]}:data[i];
    if(value&&value.formula)return {userEnteredValue:{formulaValue:value.formula}};
    if(value instanceof Date){const local=Utilities.formatDate(value,workbook_().getSpreadsheetTimeZone(),'yyyy-MM-dd HH:mm:ss');return {userEnteredValue:{numberValue:Date.parse(local.replace(' ','T')+'Z')/86400000+25569}};}
    if(typeof value==='number')return {userEnteredValue:{numberValue:value}};
    return {userEnteredValue:{stringValue:String(value==null?'':value)}};});
  return {updateCells:{range:{sheetId:t.sheet.getSheetId(),startRowIndex:record._row-1,endRowIndex:record._row,startColumnIndex:0,endColumnIndex:t.headers.length},rows:[{values}],fields:'userEnteredValue'}};
}
function logRequest_(user,university,task,previous,updated,action,assignmentId,itemId,operationId){
  const t=table_('Activity_Log'),record={'Timestamp':new Date().toISOString(),'Employee':user.name,'Team':user.team||'Admin','University':university,'Task':task,'Previous Value':JSON.stringify(previous),'Updated Value':JSON.stringify(updated),'Action':action,'Event ID':Utilities.getUuid(),'Assignment ID':key_(assignmentId),'Item ID':key_(itemId),'Member ID':user.memberId||user.id,'Request ID':operationId};
  return {appendCells:{sheetId:t.sheet.getSheetId(),rows:[{values:t.headers.map(h=>({userEnteredValue:{stringValue:String(record[h]||'')}}))}],fields:'userEnteredValue'}};
}
function saveItem_(user,p){
  const ctx=itemContext_(user,p.itemId),{item,assignment,d}=ctx;operation_(p.operationId);
  if(seen_(user,p.operationId,item['Item ID'],assignment['Assignment ID']))return {item:itemView_(item),duplicate:true};
  version_(item,p.version);if(!['Pending','Done'].includes(p.status)||typeof p.remarks!=='string'||p.remarks.length>5000)failure_('Invalid status or remarks');
  if(p.removeImage&&p.image)failure_('Choose either image removal or replacement');
    let newFile,folder;const now=new Date().toISOString();
  try{
    if(p.image){const image=validateImage_(p.image);folder=destination_(assignment,item);const filename=safeName_(item['Item Name'])+'-'+safeName_(item['Item ID'])+'-'+Utilities.getUuid()+'.'+image.extension;newFile=folder.createFile(Utilities.newBlob(image.bytes,image.mime,filename));}
    if(p.removeImage)p.status='Pending';else if(newFile)p.status='Done';
      if(p.status==='Done'&&!newFile)existingImage_(item);
    const changes={'Status':p.status,'Remarks':p.remarks,'Last Updated':now,'Member ID':member_(assignment,d.members),'Version':(Number(item.Version)||0)+1,'Completed At':p.status==='Done'?(item.Status==='Done'?item['Completed At']||now:now):''};
    if(p.removeImage)Object.assign(changes,{'Image File ID':'','Image URL':'','Image Preview':'','Item Folder ID':''});
      if(newFile)Object.assign(changes,{'Image File ID':newFile.getId(),'Image URL':newFile.getUrl(),'Image Preview':{formula:'=HYPERLINK(J'+item._row+',"Open saved image")'},'Item Folder ID':folder.getId()});
    const next={...item,...changes},oldItems=d.items.map(itemView_),nextItems=d.items.map(r=>itemView_(key_(r['Item ID'])===key_(item['Item ID'])?next:r));
    const previousAssignment=assignmentView_(assignment,oldItems),updatedAssignment=assignmentView_(assignment,nextItems);
    const assignmentChanges={'Last Updated':now,'Version':(Number(assignment.Version)||0)+1,'Member ID':member_(assignment,d.members),'Completed At':updatedAssignment.status==='Completed'?(previousAssignment.status==='Completed'?assignment['Completed At']||now:now):''};
    const requests=[rowRequest_('USP_Ranking_Items',item,changes),rowRequest_('USP_Ranking_Assignments',assignment,assignmentChanges),logRequest_(user,item['University Name'],item['Item Name'],itemView_(item),itemView_(next),'Item Updated',assignment['Assignment ID'],item['Item ID'],p.operationId)];
    if(previousAssignment.status!==updatedAssignment.status)requests.push(logRequest_(user,assignment['University Name'],'University assignment',{status:previousAssignment.status},{status:updatedAssignment.status},updatedAssignment.status==='Completed'?'Assignment Completed':'Assignment Status Changed',assignment['Assignment ID'],'',p.operationId));
    Sheets.Spreadsheets.batchUpdate({requests},SHEET_ID);
    // The old file remains available for recovery and audit after a successful replacement.
    return {item:itemView_(next),assignment:{...updatedAssignment,lastUpdated:now,completedAt:assignmentChanges['Completed At']}};
  }catch(error){
    if(newFile){
      // An API timeout can happen after Google committed the batch. Verify before cleanup.
      let definitelyUncommitted=false;
      try{
        delete requestTables_['USP_Ranking_Items'];delete requestTables_['Activity_Log'];
        const current=unique_(table_('USP_Ranking_Items').rows,item['Item ID'],'Item ID');
        if(current['Image File ID']===newFile.getId()&&seen_(user,p.operationId,item['Item ID'],assignment['Assignment ID']))return {item:itemView_(current),recovered:true};
        definitelyUncommitted=key_(current['Image File ID'])===key_(item['Image File ID']);
      }catch{console.error('Upload outcome requires refresh; uploaded file retained');}
      if(definitelyUncommitted){try{newFile.setTrashed(true);}catch{console.error('Failed upload file needs owner cleanup');}}
    }throw error;
  }
}
function saveContent_(user,p){
  const d=dataTables_(),row=unique_(d.content,p.assignmentId,'Assignment ID');if(!visible_(row,user,d.members))failure_('You cannot edit this assignment',403);operation_(p.operationId);
  if(seen_(user,p.operationId,'',row['Assignment ID']))return {assignment:contentView_(row),duplicate:true};version_(row,p.version);
  const names=contentPages_(row);
  if(!Array.isArray(p.pages)||p.pages.length!==names.length||new Set(p.pages.map(p=>p&&p.name)).size!==names.length)failure_('All '+names.length+' subpages are required');
  const changes={'Last Updated':new Date().toISOString(),'Version':(Number(row.Version)||0)+1,'Member ID':member_(row,d.members)};
  p.pages.forEach(page=>{if(!page||!names.includes(page.name)||!['Pending','Done'].includes(page.status)||typeof page.remarks!=='string'||page.remarks.length>5000)failure_('Invalid subpage');changes[contentColumn_(row,page.name,'Status')]=page.remarks.trim()?'Done':'Pending';changes[contentColumn_(row,page.name,'Remarks')]=page.remarks;});
  const before=contentView_(row),after=contentView_({...row,...changes});changes['Completed At']=after.status==='Completed'?(before.status==='Completed'?row['Completed At']||changes['Last Updated']:changes['Last Updated']):'';
  const requests=[rowRequest_('Content_Verification',row,changes),logRequest_(user,row['University Name'],'Content verification',before,after,'Content Updated',row['Assignment ID'],'',p.operationId)];
  if(before.status!==after.status)requests.push(logRequest_(user,row['University Name'],'University assignment',{status:before.status},{status:after.status},after.status==='Completed'?'Assignment Completed':'Assignment Status Changed',row['Assignment ID'],'',p.operationId));
  Sheets.Spreadsheets.batchUpdate({requests},SHEET_ID);return {assignment:contentView_({...row,...changes})};
}
/** Run once in the editor after setting BRIDGE_SECRET and enabling the Sheets service. */
function checkSetup(){
  if((PropertiesService.getScriptProperties().getProperty('BRIDGE_SECRET')||'').length<32)failure_('Set BRIDGE_SECRET in Script Properties first');
  const book=workbook_();['Content_Verification','USP_Ranking_Assignments','USP_Ranking_Items','Team_Members','Activity_Log','Admin_Users'].forEach(name=>{if(!book.getSheetByName(name))failure_('Missing sheet: '+name);});
  privateRoot_();Sheets.Spreadsheets.get(SHEET_ID,{fields:'spreadsheetId'});console.log('Google backend configuration verified.');
}

const BACKLINK_PLATFORMS = {'MEM-GAURANSHI':['LinkedIn','X','Threads'],'MEM-ANUKRITI':['Diigo','Blogger','Tumblr'],'MEM-YASH':['LinkedIn','X','Threads'],'MEM-ARHAN':['Tumblr','Blogger','Diigo']};
function backlinkPlatforms_(memberId,row){const platforms=BACKLINK_PLATFORMS[memberId]||[];return memberId==='MEM-ARHAN'&&row&&row['Assignment Type']==='Exam'&&row.Section==='News'?[...platforms,'Classified Ads']:platforms;}
function backlinkUrl_(value,required){
  if(typeof value!=='string'||value.length>2048)failure_('Enter a valid web link');
  const url=value.trim();if(!url&&!required)return '';
  if(!/^https?:\/\/[a-z0-9](?:[a-z0-9.-]*[a-z0-9])?(?::[0-9]{1,5})?(?:[/?#][^\s]*)?$/i.test(url)||/[\u0000-\u0020<>"\\]/.test(url))failure_('Enter a complete http or https link');
  return url;
}
function backlinkView_(row,submissions){
  const memberId=key_(row['Member ID']),platforms=backlinkPlatforms_(memberId,row);
  const items=platforms.map(platform=>{
    const matches=submissions.filter(s=>key_(s['Assignment ID'])===key_(row['Assignment ID'])&&s.Platform===platform&&key_(s['Member ID'])===memberId);
    if(matches.length>1)failure_('Duplicate backlink submissions must be corrected',409);
    const s=matches[0]||{};
    return {platform,submissionUrl:String(s['Submission URL']||''),remarks:String(s.Remarks||''),status:s['Submission URL']?'Done':'Pending'};
  });
  const completed=items.filter(s=>s.status==='Done').length,total=platforms.length;
  return {id:key_(row['Assignment ID']),date:row.Date,project:row.Project||'GEDUCONNECT',sourceUrl:String(row['Source URL']||''),university:row['University Name']||'',page:row['Page Name']||'',assignmentType:row['Assignment Type']||'University',exam:row['Exam Name']||'',section:row.Section||'',linkNumber:Number(row['Link Number'])||0,examGroupId:row['Exam Group ID']||'',assignedTo:row['Assigned To'],memberId,version:Number(row.Version)||0,platforms,submissions:items,completed,total,progress:total?completed/total:0,status:total&&completed===total?'Completed':completed?'In Progress':'Pending',lastUpdated:row['Last Updated']||'',completedAt:row['Completed At']||''};
}
function backlinkTables_(){return {assignments:table_('Backlink_Assignments').rows.filter(r=>required_(r,['Assignment ID','Source URL','Member ID'])),submissions:table_('Backlink_Submissions').rows.filter(r=>required_(r,['Submission ID','Assignment ID','Member ID','Platform']))};}
function appendRecord_(name,record){
  const t=table_(name);return {appendCells:{sheetId:t.sheet.getSheetId(),rows:[{values:t.headers.map(h=>({userEnteredValue:typeof record[h]==='number'?{numberValue:record[h]}:{stringValue:String(record[h]??'')}}))}],fields:'userEnteredValue'}};
}
function backlinkOwner_(row,user){
  if(user.role!=='admin'&&(user.team!=='Backlink Team'||key_(row['Member ID'])!==user.memberId))failure_('You cannot edit this backlink assignment',403);
}
function createBacklink_(user,p){
  if(user.role!=='admin')failure_('Only administrators can assign backlinks',403);
  operation_(p.operationId);
  const member=unique_(table_('Team_Members').rows,p.memberId,'Member ID'),platforms=backlinkPlatforms_(p.memberId);
  if(!['MEM-GAURANSHI','MEM-ANUKRITI'].includes(p.memberId)||member.Team!=='Backlink Team'||member['Active Status']!=='Active'||platforms.length!==3)failure_('Select an active backlink member with configured platforms');
  const sourceUrl=backlinkUrl_(p.sourceUrl,true);
  if(!/^https?:\/\/(?:www\.)?geduconnect\.(?:com|in)(?:[/?#]|$)/i.test(sourceUrl))failure_('Enter a GEDUCONNECT source link');
  const d=backlinkTables_(),previous=d.assignments.find(r=>r['Request ID']===p.operationId);
  if(previous)return {assignment:backlinkView_(previous,d.submissions),duplicate:true};
  if(d.assignments.some(r=>r['Member ID']===p.memberId&&String(r['Source URL']).replace(/\/$/,'')===sourceUrl.replace(/\/$/,'')))failure_('This link is already assigned to this member',409);
  const now=new Date().toISOString(),id='BL-'+Utilities.getUuid(),record={'Assignment ID':id,Date:Utilities.formatDate(new Date(),'Asia/Kolkata','yyyy-MM-dd'),Project:'GEDUCONNECT','Source URL':sourceUrl,'Assigned To':member['Employee Name'],'Member ID':p.memberId,'Required Platforms':3,'Completed Platforms':0,'Progress Percentage':0,'Overall Status':'Pending','Last Updated':now,Version:1,'Completed At':'','Request ID':p.operationId};
  const items=platforms.map((platform,i)=>({'Submission ID':id+'-'+(i+1),'Assignment ID':id,'Member ID':p.memberId,Platform:platform,'Submission URL':'',Remarks:'',Status:'Pending','Last Updated':now,Version:1,'Completed At':''}));
  const actor={...user,memberId:p.memberId,name:user.name,team:'Backlink Team'};
  Sheets.Spreadsheets.batchUpdate({requests:[appendRecord_('Backlink_Assignments',record),...items.map(s=>appendRecord_('Backlink_Submissions',s)),logRequest_(actor,sourceUrl,'Backlink assignment',{},record,'Assignment Created',id,'',p.operationId)]},SHEET_ID);
  return {assignment:backlinkView_(record,items)};
}
function saveBacklink_(user,p){
  const d=backlinkTables_(),row=unique_(d.assignments,p.assignmentId,'Assignment ID');
  backlinkOwner_(row,user);operation_(p.operationId);
  const actor={...user,memberId:row['Member ID'],team:'Backlink Team'};
  if(seen_(actor,p.operationId,'',row['Assignment ID']))return {assignment:backlinkView_(row,d.submissions),duplicate:true};
  version_(row,p.version);
  if(!backlinkPlatforms_(row['Member ID'],row).includes(p.platform))failure_('This platform is not assigned to this member',403);
  if(typeof p.remarks!=='string'||p.remarks.length>5000)failure_('Remarks must be under 5000 characters');
  const submissionUrl=backlinkUrl_(p.submissionUrl,false);
  if(submissionUrl&&submissionUrl.replace(/\/$/,'')===String(row['Source URL']).replace(/\/$/,''))failure_('Enter the published submission link, not the source link');
  const matches=d.submissions.filter(s=>key_(s['Assignment ID'])===key_(row['Assignment ID'])&&s.Platform===p.platform&&s['Member ID']===row['Member ID']);
  if(matches.length!==1)failure_('The platform submission row must be corrected',409);
  const item=matches[0],now=new Date().toISOString(),changes={'Submission URL':submissionUrl,Remarks:p.remarks,Status:submissionUrl?'Done':'Pending','Last Updated':now,Version:(Number(item.Version)||0)+1,'Completed At':submissionUrl?(item['Completed At']||now):''};
  const next={...item,...changes},before=backlinkView_(row,d.submissions),after=backlinkView_(row,d.submissions.map(s=>s===item?next:s));
  const assignmentChanges={'Completed Platforms':after.completed,'Progress Percentage':after.progress,'Overall Status':after.status,'Last Updated':now,Version:(Number(row.Version)||0)+1,'Completed At':after.status==='Completed'?(row['Completed At']||now):''};
  const requests=[rowRequest_('Backlink_Submissions',item,changes),rowRequest_('Backlink_Assignments',row,assignmentChanges),logRequest_(actor,row['Source URL'],p.platform,before,after,'Backlink Updated',row['Assignment ID'],'',p.operationId)];
  if(before.status!==after.status)requests.push(logRequest_(actor,row['Source URL'],'Backlink assignment',{status:before.status},{status:after.status},after.status==='Completed'?'Assignment Completed':'Assignment Status Changed',row['Assignment ID'],'',p.operationId));
  Sheets.Spreadsheets.batchUpdate({requests},SHEET_ID);
  return {assignment:backlinkView_({...row,...assignmentChanges},d.submissions.map(s=>s===item?next:s))};
}

const BACKLINK_UNIVERSITY_PAGES = [['Overview',''],['Program & Fees','program-fees'],['Admission process','admission-process'],['Placement & Career Support','placement-career-support'],['Student Support','student-support'],['FAQs','faqs']];
function backlinkUniversityBase_(value){
  const url=backlinkUrl_(value,true).replace(/\/$/,'');
  if(!/^https:\/\/(?:www\.)?geduconnect\.com\/university\/[a-z0-9]+(?:-[a-z0-9]+)*$/i.test(url))failure_('Enter the GEDUCONNECT university overview link, such as https://geduconnect.com/university/vgu-online');
  return url.replace(/^https:\/\/(?:www\.)?/i,'https://').replace('GEDUCONNECT.COM','geduconnect.com');
}
function createBacklinkUniversity_(user,p){
  if(user.role!=='admin')failure_('Only administrators can assign universities',403);
  operation_(p.operationId);
  if(typeof p.university!=='string'||!p.university.trim()||p.university.length>200)failure_('Enter a university name');
  const base=backlinkUniversityBase_(p.sourceUrl),date=p.date;
  if(typeof date!=='string'||!/^\d{4}-\d{2}-\d{2}$/.test(date)||isNaN(new Date(date+'T00:00:00Z').getTime())||new Date(date+'T00:00:00Z').toISOString().slice(0,10)!==date)failure_('Enter a valid assignment date');
  const d=backlinkTables_(),previous=d.assignments.filter(r=>r['Request ID']===p.operationId);
  if(previous.length)return {assignments:previous.map(r=>backlinkView_(r,d.submissions)),duplicate:true};
  const members=['MEM-GAURANSHI','MEM-ANUKRITI'].map(id=>unique_(table_('Team_Members').rows,id,'Member ID'));
  for(const m of members){
    if(m.Team!=='Backlink Team'||m['Active Status']!=='Active'||backlinkPlatforms_(m['Member ID']).length!==3)failure_('Both backlink members must be active');
    const own=d.assignments.filter(r=>r['Member ID']===m['Member ID']);
    const sources=BACKLINK_UNIVERSITY_PAGES.map(page=>base+(page[1]?'/'+page[1]:''));
    if(own.some(r=>sources.includes(String(r['Source URL']).replace(/\/$/,''))))failure_('This university is already assigned to '+m['Employee Name'],409);
    const daily=new Set(own.filter(r=>String(r.Date).slice(0,10)===date).map(r=>String(r['Source URL']).replace(/(\/university\/[^/?#]+).*$/,'$1')));
    if(daily.size>=2)failure_('Two universities are already assigned to '+m['Employee Name']+' for this date',409);
  }
  const now=new Date().toISOString(),records=[],items=[],requests=[];
  for(const m of members)for(const page of BACKLINK_UNIVERSITY_PAGES){
    const id='BL-'+Utilities.getUuid(),sourceUrl=base+(page[1]?'/'+page[1]:''),record={'Assignment ID':id,Date:date,Project:'GEDUCONNECT','Source URL':sourceUrl,'Assigned To':m['Employee Name'],'Member ID':m['Member ID'],'Required Platforms':3,'Completed Platforms':0,'Progress Percentage':0,'Overall Status':'Pending','Last Updated':now,Version:1,'Completed At':'','Request ID':p.operationId,'University Name':p.university.trim(),'Page Name':page[0]};
    records.push(record);requests.push(appendRecord_('Backlink_Assignments',record));
    backlinkPlatforms_(m['Member ID']).forEach((platform,i)=>{const s={'Submission ID':id+'-'+(i+1),'Assignment ID':id,'Member ID':m['Member ID'],Platform:platform,'Submission URL':'',Remarks:'',Status:'Pending','Last Updated':now,Version:1,'Completed At':''};items.push(s);requests.push(appendRecord_('Backlink_Submissions',s));});
    requests.push(logRequest_({...user,memberId:m['Member ID'],team:'Backlink Team'},p.university.trim(),page[0],{},record,'Assignment Created',id,'',p.operationId));
  }
  Sheets.Spreadsheets.batchUpdate({requests},SHEET_ID);
  return {assignments:records.map(r=>backlinkView_(r,items))};
}

function createBacklinkExam_(user,p){
  if(user.role!=='admin')failure_('Only administrators can assign exams',403);
  operation_(p.operationId);
  if(typeof p.exam!=='string'||!p.exam.trim()||p.exam.length>200)failure_('Enter an exam name');
  if(typeof p.date!=='string'||!/^\d{4}-\d{2}-\d{2}$/.test(p.date)||isNaN(new Date(p.date+'T00:00:00Z').getTime())||new Date(p.date+'T00:00:00Z').toISOString().slice(0,10)!==p.date)failure_('Enter a valid assignment date');
  if(!Array.isArray(p.articles)||p.articles.length!==6||!Array.isArray(p.news)||p.news.length!==4)failure_('Provide exactly 6 article links and 4 news links');
  const sections=[['Articles',p.articles.map(v=>backlinkUrl_(v,true))],['News',p.news.map(v=>backlinkUrl_(v,true))]];
  const sources=sections.flatMap(s=>s[1]);if(new Set(sources.map(v=>v.replace(/\/$/,''))).size!==10)failure_('All 10 source links must be different');
  const d=backlinkTables_(),previous=d.assignments.filter(r=>r['Request ID']===p.operationId);
  if(previous.length)return {assignments:previous.map(r=>backlinkView_(r,d.submissions)),duplicate:true};
  const members=['MEM-YASH','MEM-ARHAN'].map(id=>unique_(table_('Team_Members').rows,id,'Member ID'));
  for(const m of members){
    if(m.Team!=='Backlink Team'||m['Active Status']!=='Active')failure_('Yash and Arhan must both be active backlink members');
    const own=d.assignments.filter(r=>r['Member ID']===m['Member ID']);
    if(own.some(r=>r['Assignment Type']==='Exam'&&String(r['Exam Name']).trim().toLowerCase()===p.exam.trim().toLowerCase()&&String(r.Date).slice(0,10)===p.date))failure_('This exam is already assigned to '+m['Employee Name']+' for this date',409);
    if(own.some(r=>sources.map(v=>v.replace(/\/$/,'')).includes(String(r['Source URL']).replace(/\/$/,''))))failure_('A source link is already assigned to '+m['Employee Name'],409);
  }
  const now=new Date().toISOString(),groupId='EXBL-'+Utilities.getUuid(),records=[],items=[],requests=[];
  for(const m of members)for(const section of sections)section[1].forEach((sourceUrl,i)=>{
    const id='BL-'+Utilities.getUuid(),record={'Assignment ID':id,Date:p.date,Project:'Exam Backlinks','Source URL':sourceUrl,'Assigned To':m['Employee Name'],'Member ID':m['Member ID'],'Completed Platforms':0,'Progress Percentage':0,'Overall Status':'Pending','Last Updated':now,Version:1,'Completed At':'','Request ID':p.operationId,'University Name':'','Page Name':section[0]+' '+(i+1),'Assignment Type':'Exam','Exam Name':p.exam.trim(),Section:section[0],'Link Number':i+1,'Exam Group ID':groupId};
    const platforms=backlinkPlatforms_(m['Member ID'],record);record['Required Platforms']=platforms.length;records.push(record);requests.push(appendRecord_('Backlink_Assignments',record));
    platforms.forEach((platform,j)=>{const s={'Submission ID':id+'-'+(j+1),'Assignment ID':id,'Member ID':m['Member ID'],Platform:platform,'Submission URL':'',Remarks:'',Status:'Pending','Last Updated':now,Version:1,'Completed At':''};items.push(s);requests.push(appendRecord_('Backlink_Submissions',s));});
    requests.push(logRequest_({...user,memberId:m['Member ID'],team:'Backlink Team'},p.exam.trim(),record['Page Name'],{},record,'Assignment Created',id,'',p.operationId));
  });
  Sheets.Spreadsheets.batchUpdate({requests},SHEET_ID);
  return {assignments:records.map(r=>backlinkView_(r,items))};
}
