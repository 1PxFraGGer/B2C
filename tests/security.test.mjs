import test from 'node:test';
import assert from 'node:assert/strict';
import {validateImage,validateSvg,MAX_IMAGE_BYTES} from '../server/files.mjs';
import {assertOrigin,assertCsrf,csrf,cookie} from '../server/auth.mjs';
const png='iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6wAAAAABJRU5ErkJggg==';
test('file contents, extension, and size are checked',()=>{
 assert.equal(validateImage({name:'badge.png',base64:png}).mime,'image/png');
 assert.throws(()=>validateImage({name:'badge.jpg',base64:png}));
 assert.throws(()=>validateImage({name:'badge.png',base64:Buffer.alloc(MAX_IMAGE_BYTES+1).toString('base64')}));
 assert.throws(()=>validateImage({name:'badge.png',base64:'not-base64!'}));
});
test('SVG permits vector shapes but rejects script, external links, style, and entities',()=>{
 const svg=inside=>`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">${inside}</svg>`;
 assert.equal(validateSvg(svg('<path d="M0 0L100 100" fill="#087d69"/>')),svg('<path d="M0 0L100 100" fill="#087d69"/>'));
 for(const unsafe of ['<script>alert(1)</script>','<image href="https://example.com/a.png"/>','<path style="fill:red"/>','<use href="data:image/svg+xml,abc"/>','<rect onload="alert(1)"/>','<path fill="url(https://example.com/a)"/>'])assert.throws(()=>validateSvg(svg(unsafe)));
 assert.throws(()=>validateSvg('<!DOCTYPE svg [<!ENTITY x SYSTEM "file:///secret">]>'+svg('&x;')));
});
test('same-origin and CSRF checks cannot be bypassed by other origins',()=>{
 process.env.APP_ORIGIN='https://university.example';process.env.BRIDGE_SECRET='test-secret-for-unit-tests-only-long';
 const token='a'.repeat(64),valid={headers:{origin:'https://university.example','x-csrf-token':csrf(token)}};
 assert.doesNotThrow(()=>assertCsrf(valid,token));
 assert.throws(()=>assertOrigin({headers:{origin:'https://attacker.example'}}));
 assert.throws(()=>assertCsrf({headers:{origin:'https://university.example','x-csrf-token':'wrong'}},token));
 const original=process.env.NODE_ENV;process.env.NODE_ENV='production';assert.match(cookie(token),/HttpOnly; SameSite=Strict/);assert.match(cookie(token),/; Secure$/);process.env.NODE_ENV=original||'test';
});

