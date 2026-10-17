import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import {randomUUID} from 'node:crypto';
const source=fs.readFileSync(new URL('../apps-script/Code.gs',import.meta.url),'utf8');
function fixture({busy=false}={}){
 let held=false,batches=0,opens=0,releases=0;const nonces=new Map();
 const context=vm.createContext({console,Date,Set,Map,JSON,Number,String,Array,Object,Math,RegExp,Error,
  ContentService:{MimeType:{JSON:'json'},createTextOutput:text=>({setMimeType:()=>JSON.parse(text)})},
  PropertiesService:{getScriptProperties:()=>({getProperty:()=> 's'.repeat(40)})},
  Utilities:{Charset:{UTF_8:'utf8'},computeHmacSha256Signature:()=>[],base64Encode:()=> 'signature'},
  CacheService:{getScriptCache:()=>({get:k=>nonces.get(k),put:(k,v)=>nonces.set(k,v)})},
  LockService:{getScriptLock:()=>({tryLock:()=>{if(busy)return false;held=true;return true;},releaseLock:()=>{assert(held);held=false;releases++;}})},
  SpreadsheetApp:{openById:()=>{opens++;return {getSheetByName:name=>({name})};}},
  Sheets:{Spreadsheets:{Values:{batchGet:(_id,options)=>{batches++;return {valueRanges:options.ranges.map(()=>({values:[['Assignment ID','Date','University Name'],['ID-1',46311,'University'],[],['ID-2',46312,'Second university']]}))};}}}}
 });vm.runInContext(source,context);
 context.session_=()=>({role:'employee',memberId:'MEM-PAYAL'});
 const invoke=action=>context.doPost({postData:{contents:JSON.stringify({body:JSON.stringify({action,payload:{},session:'token',timestamp:Date.now(),requestId:randomUUID()}),signature:'signature'})}});
 return {context,invoke,held:()=>held,batches:()=>batches,opens:()=>opens,releases:()=>releases};
}
test('workspace and preview reads release the global lock before doing slow work',()=>{
 const f=fixture();f.context.workspace_=()=>{assert.equal(f.held(),false);return {ok:true};};f.context.image_=()=>{assert.equal(f.held(),false);return {mime:'image/png'};};
 assert(f.invoke('data').ok);assert(f.invoke('image').ok);assert.equal(f.releases(),2);assert.equal(f.batches(),2);
});
test('writes remain serialized and release their lock after completion',()=>{
 const f=fixture();f.context.saveItem_=()=>{assert.equal(f.held(),true);return {saved:true};};assert(f.invoke('saveItem').ok);assert.equal(f.held(),false);assert.equal(f.releases(),1);
});
test('lock contention returns a clean retryable error without releasing an unowned lock',()=>{
 const f=fixture({busy:true}),r=f.invoke('saveItem');assert.equal(r.status,503);assert.equal(f.releases(),0);
});
test('batched reads reuse tables, preserve row numbers and normalize calendar dates',()=>{
 const f=fixture();f.context.primeTables_(['USP_Ranking_Assignments','USP_Ranking_Items']);
 const a=f.context.table_('USP_Ranking_Assignments');assert.equal(a.rows[0].Date,'2026-10-16');assert.equal(a.rows[1]._row,4);assert.equal(f.context.table_('USP_Ranking_Assignments'),a);
 f.context.primeTables_(['USP_Ranking_Assignments']);assert.equal(f.batches(),1);assert.equal(f.opens(),0);void a.sheet;void a.sheet;assert.equal(f.opens(),1);
});
test('each request refreshes its table snapshot instead of reusing another request data',()=>{
 const f=fixture();f.context.workspace_=()=>({});f.invoke('data');f.invoke('data');assert.equal(f.batches(),2);
});
