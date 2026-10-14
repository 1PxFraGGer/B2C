export function safeLiveLink(value){
 if(typeof value!=='string'||value.length>2048)return '';
 const raw=value.trim();if(!/^https?:\/\//i.test(raw))return '';
 try{const url=new URL(raw);return ['http:','https:'].includes(url.protocol)&&!url.username&&!url.password?url.href:'';}catch{return '';}
}
