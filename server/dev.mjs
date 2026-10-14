import http from 'node:http';
import {createServer} from 'vite';
import {existsSync} from 'node:fs';
import handler from '../api/index.mjs';
if(existsSync('.env.local'))process.loadEnvFile('.env.local');
const server=http.createServer(async(req,res)=>{
  const url=new URL(req.url,'http://localhost:3001');req.query=Object.fromEntries(url.searchParams);req.query.action=url.pathname.split('/').pop();
  res.status=function(code){this.statusCode=code;return this;};res.json=function(body){this.setHeader('Content-Type','application/json');this.end(JSON.stringify(body));return this;};res.send=function(body){this.end(body);return this;};
  try{
    if(req.method==='POST'){let size=0,chunks=[];for await(const chunk of req){size+=chunk.length;if(size>3000000){res.status(413).json({error:'Request exceeds the upload limit'});return;}chunks.push(chunk);}req.body=JSON.parse(Buffer.concat(chunks).toString()||'{}');}
    await handler(req,res);
  }catch{res.status(400).json({error:'Invalid request'});}
});
server.listen(3001,'127.0.0.1');const vite=await createServer();await vite.listen();vite.printUrls();
function stop(){server.close();vite.close().then(()=>process.exit(0));}process.on('SIGINT',stop);process.on('SIGTERM',stop);
