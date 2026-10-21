import { scrypt, timingSafeEqual, randomBytes, createHmac } from 'node:crypto';
import {promisify} from 'node:util';
import {google,ApiError} from './google.mjs';
const derive=promisify(scrypt);
export function accounts() {
  if(!process.env.AUTH_ACCOUNTS_BASE64)throw new ApiError('Login accounts are not configured on the server.',503);
  let users;
  try {
    users=JSON.parse(Buffer.from(process.env.AUTH_ACCOUNTS_BASE64,'base64').toString());
    if(!Array.isArray(users)||!users.length)throw new Error('Invalid base accounts');
    if(process.env.AUTH_ADDITIONAL_ACCOUNTS_BASE64){
      const additional=JSON.parse(Buffer.from(process.env.AUTH_ADDITIONAL_ACCOUNTS_BASE64,'base64').toString());
      if(!Array.isArray(additional))throw new Error('Invalid additional accounts');
      users=[...users,...additional];
    }
  }catch{throw new ApiError('Invalid login configuration',503);}
  if(!Array.isArray(users)||!users.length||new Set(users.map(u=>u.loginId)).size!==users.length)throw new ApiError('Invalid login configuration',503);
  for(const u of users)if(!u.id||!['employee','admin'].includes(u.role)||!/^[a-f0-9]{32}$/.test(u.salt)||!/^[a-f0-9]{128}$/.test(u.passwordHash))throw new ApiError('Invalid login configuration',503);
  return users;
}
export function sessionToken(req) {
  const cookies=String(req.headers.cookie||'').split(';').map(c=>c.trim().split('='));
  return cookies.find(c=>c[0]==='university_session')?.[1]||null;
}
export function csrf(token) {return createHmac('sha256',process.env.BRIDGE_SECRET||'').update('csrf:'+token).digest('hex');}
export function assertOrigin(req) {
  let expected;try{expected=new URL(process.env.APP_ORIGIN).origin;}catch{throw new ApiError('APP_ORIGIN must be configured on the backend.',503);}
  if(req.headers.origin!==expected)throw new ApiError('Request origin is not allowed',403);
}
export function assertCsrf(req,token) {
  assertOrigin(req);const received=String(req.headers['x-csrf-token']||'');const expected=csrf(token||'');
  if(received.length!==expected.length||!timingSafeEqual(Buffer.from(received),Buffer.from(expected)))throw new ApiError('Invalid request token. Refresh and try again.',403);
}
export function cookie(token,maxAge=null) {
  return `university_session=${token}; Path=/; HttpOnly; SameSite=Strict${maxAge===null?'':`; Max-Age=${maxAge}`}${process.env.NODE_ENV==='production'?'; Secure':''}`;
}
export async function login(req,body) {
  assertOrigin(req);
  if(typeof body.loginId!=='string'||typeof body.password!=='string'||body.loginId.length>254||body.password.length>256)throw new ApiError('Enter your login ID and password');
  const loginId=body.loginId.trim().toLowerCase();
  // The shared Google limiter persists across Vercel instances and applies to unknown IDs too.
  await google('loginRate',{loginId});
  const users=accounts(),user=users.find(u=>u.loginId===loginId),candidate=user||users[0];
  const actual=await derive(body.password,Buffer.from(candidate.salt,'hex'),64,{N:32768,r:8,p:1,maxmem:64*1024*1024});
  if(!timingSafeEqual(actual,Buffer.from(candidate.passwordHash,'hex'))||!user?.active)throw new ApiError('Invalid login ID or password',401);
  const token=randomBytes(32).toString('hex');
  const identity={id:user.id,loginId:user.loginId,name:user.name,role:user.role,memberId:user.memberId,team:user.team};
  await google('sessionCreate',{token,identity});
  return {token,identity};
}
export async function requireSession(req) {
  const token=sessionToken(req);if(!token||!/^[a-f0-9]{64}$/.test(token))throw new ApiError('Please sign in',401);
  return {token,identity:await google('sessionRead',{},token)};
}
