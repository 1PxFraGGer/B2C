import test from 'node:test';
import assert from 'node:assert/strict';
import handler from '../api/index.mjs';
import {cookie,csrf} from '../server/auth.mjs';
test('remembered cookie persists; ordinary login uses browser session cookie',()=>{
 assert.ok(!cookie('token').includes('Max-Age'));
 assert.ok(cookie('token',28800).includes('Max-Age=28800'));
 assert.ok(cookie('',0).includes('Max-Age=0'));
});
test('logout clears an expired session instead of rejecting it',async()=>{
 const oldFetch=globalThis.fetch;
 const previous={...process.env};
 process.env.APP_ORIGIN='https://example.com';process.env.BRIDGE_SECRET='x'.repeat(32);
 process.env.APPS_SCRIPT_URL='https://script.google.com/macros/s/test/exec';
 globalThis.fetch=async()=>({json:async()=>({ok:false,status:401,error:'Session expired'})});
 const token='a'.repeat(64),headers={};let status,body;
 const res={setHeader:(k,v)=>headers[k]=v,status(n){status=n;return this;},json(v){body=v;return this;}};
 try{
  await handler({method:'POST',query:{action:'logout'},headers:{origin:'https://example.com',cookie:'university_session='+token,'x-csrf-token':csrf(token)}},res);
  assert.equal(status,200);assert.equal(body.ok,true);assert.ok(headers['Set-Cookie'].includes('Max-Age=0'));
  await handler({method:'POST',query:{action:'logout'},headers:{origin:'https://evil.com',cookie:'university_session='+token,'x-csrf-token':csrf(token)}},res);
  assert.equal(status,403);
 }finally{globalThis.fetch=oldFetch;for(const key of ['APP_ORIGIN','BRIDGE_SECRET','APPS_SCRIPT_URL'])if(previous[key]===undefined)delete process.env[key];else process.env[key]=previous[key];}
});
