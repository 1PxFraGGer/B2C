import test from 'node:test';
import assert from 'node:assert/strict';
import {accounts} from '../server/auth.mjs';
const base={id:'existing',loginId:'existing@example.com',role:'employee',salt:'a'.repeat(32),passwordHash:'b'.repeat(128),active:true};
const added={...base,id:'new',loginId:'new@example.com'};
const encode=value=>Buffer.from(JSON.stringify(value)).toString('base64');
test('additional accounts preserve existing credentials and reject ambiguous configuration',()=>{
 const savedBase=process.env.AUTH_ACCOUNTS_BASE64,savedAdditional=process.env.AUTH_ADDITIONAL_ACCOUNTS_BASE64;
 try{
  process.env.AUTH_ACCOUNTS_BASE64=encode([base]);
  delete process.env.AUTH_ADDITIONAL_ACCOUNTS_BASE64;
  assert.deepEqual(accounts(),[base]);
  process.env.AUTH_ADDITIONAL_ACCOUNTS_BASE64=encode([added]);
  assert.deepEqual(accounts(),[base,added]);
  for(const invalid of [[base],{},[{...added,passwordHash:'invalid'}]]){
   process.env.AUTH_ADDITIONAL_ACCOUNTS_BASE64=encode(invalid);
   assert.throws(()=>accounts(),/Invalid login configuration/);
  }
  delete process.env.AUTH_ACCOUNTS_BASE64;
  assert.throws(()=>accounts(),/not configured/);
 }finally{
  for(const [key,value] of Object.entries({AUTH_ACCOUNTS_BASE64:savedBase,AUTH_ADDITIONAL_ACCOUNTS_BASE64:savedAdditional})){
   if(value===undefined)delete process.env[key];else process.env[key]=value;
  }
 }
});
