import {google,ApiError} from '../server/google.mjs';
import {login,requireSession,cookie,csrf,assertCsrf,assertOrigin,sessionToken} from '../server/auth.mjs';
import {validateImage} from '../server/files.mjs';
const writeActions=new Set(['saveItem','saveContent']);
export default async function handler(req,res) {
  res.setHeader('Cache-Control','no-store');res.setHeader('X-Content-Type-Options','nosniff');
  const action=req.query?.action;
  try{
    const body=req.body||{};
    if(action==='login'){
      if(req.method!=='POST')throw new ApiError('Method not allowed',405);
      const result=await login(req,body);res.setHeader('Set-Cookie',cookie(result.token,body.rememberMe===true?28800:null));return res.status(200).json({user:result.identity,csrf:csrf(result.token)});
    }
    if(action==='session'){
      if(req.method!=='GET')throw new ApiError('Method not allowed',405);
      const s=await requireSession(req);return res.status(200).json({user:s.identity,csrf:csrf(s.token)});
    }
    if(action==='logout'){
      if(req.method!=='POST')throw new ApiError('Method not allowed',405);
      assertOrigin(req);
      const token=sessionToken(req);
      if(token){
        assertCsrf(req,token);
        // Expired sessions are already revoked; clearing their cookie still succeeds.
        try{await google('sessionDelete',{},token);}catch(error){if(error.status!==401)throw error;}
      }
      res.setHeader('Set-Cookie',cookie('',0));return res.status(200).json({ok:true});
    }
    // Google validates the session and current Sheet profile inside each operation.
    const token=sessionToken(req);if(!token||!/^[a-f0-9]{64}$/.test(token))throw new ApiError('Please sign in',401);
    const s={token};
    if(action==='data'||action==='reports'){
      if(req.method!=='GET')throw new ApiError('Method not allowed',405);
      return res.status(200).json(await google(action,{},s.token));
    }
    if(writeActions.has(action)){
      if(req.method!=='POST')throw new ApiError('Method not allowed',405);assertCsrf(req,s.token);
      if(body.image)body.image=validateImage(body.image);
      return res.status(200).json(await google(action,body,s.token));
    }
    if(action==='image'){
      if(req.method!=='GET')throw new ApiError('Method not allowed',405);
      const file=await google('image',{itemId:String(req.query.itemId||'')},s.token);
      res.setHeader('Content-Type',file.mime);res.setHeader('Content-Security-Policy',"default-src 'none'; sandbox");res.setHeader('Content-Disposition','inline');return res.status(200).send(Buffer.from(file.base64,'base64'));
    }
    throw new ApiError('Endpoint not found',404);
  }catch(error){return res.status(error.status||500).json({error:error.status?error.message:'The server could not complete this request.'});}
}
