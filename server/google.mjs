import { createHmac, randomUUID } from 'node:crypto';
export class ApiError extends Error { constructor(message,status=400){super(message);this.status=status;} }
export async function google(action,payload={},session=null) {
  const url=process.env.APPS_SCRIPT_URL, secret=process.env.BRIDGE_SECRET;
  if(!url || !secret || secret.length<32) throw new ApiError('Google backend is not configured. Complete the Apps Script setup.',503);
  if(!/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(url))throw new ApiError('Invalid Apps Script deployment URL',503);
  const body=JSON.stringify({action,payload,session,requestId:randomUUID(),timestamp:Date.now()});
  const signature=createHmac('sha256',secret).update(body).digest('base64');
  let response;
  try { response=await fetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({body,signature}),signal:AbortSignal.timeout(50000)}); }
  catch {throw new ApiError('Google backend did not respond. Refresh before retrying a save.',504);}
  let result;try {result=await response.json();}catch{throw new ApiError('Google backend returned an invalid response. Check the web app deployment.',502);}
  if(!result.ok)throw new ApiError(result.error||'Google operation failed',result.status||502);
  return result.data;
}
