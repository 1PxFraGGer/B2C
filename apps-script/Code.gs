/* University dashboard Google backend. Deploy as owner; only HMAC-signed POSTs are accepted. */
const SHEET_ID = '160hjCLAHaSTU_TedLYla_BqrMXXgbcaiJNaoEL-_l7k';
const ROOT_FOLDER_ID = '1FJ1jHU-koTlzKdo9qKw7-lbU88uFzjJH';
const MAX_BYTES = 2 * 1024 * 1024;
const PAGES = ['Overview','Program & Fees','Admission','Placement','Student Support','FAQs'];
function failure_(message,status){const error=new Error(message);error.status=status||400;throw error;}
function json_(value){return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);}
function doGet(){return json_({ok:false,error:'Authenticated POST required',status:405});}
function doPost(e){
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
    lock=LockService.getScriptLock();if(!lock.tryLock(10000))failure_('Another save is in progress. Please retry.',503);
    const cache=CacheService.getScriptCache(),nonce='nonce:'+request.requestId;
    if(cache.get(nonce))failure_('Request already processed',409);cache.put(nonce,'1',180);
    const action=request.action,p=request.payload||{};
    if(action==='loginRate')return json_({ok:true,data:rate_(p.loginId)});
    if(action==='sessionCreate')return json_({ok:true,data:createSession_(p)});
    const user=session_(request.session);
    let data;
    if(action==='sessionRead')data=user;
    else if(action==='sessionDelete'){PropertiesService.getScriptProperties().deleteProperty('session:'+digest_(request.session));data={ok:true};}
    else if(action==='data'||action==='reports'){if(action==='reports'&&user.role!=='admin')failure_('Admin access required',403);data=workspace_(user,action==='reports');}
    else if(action==='saveItem')data=saveItem_(user,p);
    else if(action==='saveContent')data=saveContent_(user,p);
    else if(action==='image')data=image_(user,p.itemId);
    else failure_('Unknown action',404);
    return json_({ok:true,data});
  }catch(error){console.error(error.status?error.message:'Backend operation failed');return json_({ok:false,error:error.status?error.message:'Google operation failed. Refresh before retrying a save.',status:error.status||500});}
  finally{if(lock)lock.releaseLock();}
}
function equal_(a,b){if(a.length!==b.length)return false;let result=0;for(let i=0;i<a.length;i++)result|=a.charCodeAt(i)^b.charCodeAt(i);return result===0;}
function digest_(value){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(value),Utilities.Charset.UTF_8).map(b=>('0'+(b&255).toString(16)).slice(-2)).join('');}
function workbook_(){return SpreadsheetApp.openById(SHEET_ID);}
function table_(name){
  const sheet=workbook_().getSheetByName(name);if(!sheet)failure_('Missing sheet: '+name,503);
  const last=sheet.getLastRow(),width=sheet.getLastColumn();
  const values=sheet.getRange(1,1,Math.max(last,1),width).getValues(),headers=values[0];
  const rows=values.slice(1).map((values,i)=>{const row={_row:i+2};headers.forEach((h,j)=>{row[h]=values[j] instanceof Date?values[j].toISOString():values[j];});return row;});
  return {sheet,headers,rows};
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
  const pages=PAGES.map(name=>({name,status:r[name+' Status']==='Done'?'Done':'Pending',remarks:String(r[name+' Remarks']||'')}));
  const completed=pages.filter(p=>p.status==='Done').length;
  return {id:key_(r['Assignment ID']),date:r.Date,university:r['University Name'],liveLink:String(r['Live Link']||'').trim(),assignedTo:r['Assigned To'],memberId:r['Member ID'],version:Number(r.Version)||0,completed,total:6,progress:completed/6,status:completed===6?'Completed':completed?'In Progress':'Pending',lastUpdated:r['Last Updated'],completedAt:r['Completed At'],pages};
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
  const d=dataTables_(),allContent=d.content.map(contentView_),allItems=d.items.map(itemView_),allAssignments=d.assignments.map(r=>assignmentView_(r,allItems));
  const employeeStats=d.members.map(m=>{const isContent=m.Team==='Content Verification',records=(isContent?d.content:d.assignments).filter(r=>member_(r,d.members)===m['Member ID']);return {id:m['Member ID'],name:m['Employee Name'],team:m.Team,active:m['Active Status']==='Active',...totals_(records.map(r=>isContent?contentView_(r):assignmentView_(r,allItems)))};});
  const activity=table_('Activity_Log').rows.filter(r=>r.Timestamp&&(user.role==='admin'||r['Member ID']===user.memberId)).map(r=>({timestamp:r.Timestamp,employee:r.Employee,team:r.Team,university:r.University,task:r.Task,action:r.Action,assignmentId:key_(r['Assignment ID']),itemId:key_(r['Item ID']),previous:r['Previous Value'],updated:r['Updated Value']}));
  return {user,members:employeeStats,stats:totals_([...allContent,...allAssignments]),teams:{content:totals_(allContent),usp:totals_(allAssignments)},
    content:d.content.filter(r=>visible_(r,user,d.members)).map(contentView_),assignments:d.assignments.filter(r=>visible_(r,user,d.members)).map(r=>assignmentView_(r,allItems)),activity:reports?activity:activity.slice(-30).reverse()};
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
  let newFile,folder;const now=new Date().toISOString();
  try{
    if(p.image){const image=validateImage_(p.image);folder=destination_(assignment,item);const filename=safeName_(item['Item Name'])+'-'+safeName_(item['Item ID'])+'-'+Utilities.getUuid()+'.'+image.extension;newFile=folder.createFile(Utilities.newBlob(image.bytes,image.mime,filename));}
    if(p.status==='Done'&&!newFile)existingImage_(item);
    const changes={'Status':p.status,'Remarks':p.remarks,'Last Updated':now,'Member ID':member_(assignment,d.members),'Version':(Number(item.Version)||0)+1,'Completed At':p.status==='Done'?(item.Status==='Done'?item['Completed At']||now:now):''};
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
  if(!Array.isArray(p.pages)||p.pages.length!==6||new Set(p.pages.map(p=>p.name)).size!==6)failure_('All six subpages are required');
  const changes={'Last Updated':new Date().toISOString(),'Version':(Number(row.Version)||0)+1,'Member ID':member_(row,d.members)};
  p.pages.forEach(page=>{if(!PAGES.includes(page.name)||!['Pending','Done'].includes(page.status)||typeof page.remarks!=='string'||page.remarks.length>5000)failure_('Invalid subpage');changes[page.name+' Status']=page.status;changes[page.name+' Remarks']=page.remarks;});
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
