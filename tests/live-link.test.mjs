import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import {safeLiveLink} from '../src/live-link.mjs';
test('review links only permit absolute HTTP/S links without embedded credentials',()=>{
 assert.equal(safeLiveLink(' https://example.com/university?tab=fees '),'https://example.com/university?tab=fees');
 for(const value of ['',null,'javascript:alert(1)','data:text/html,hi','//evil.test','https://user:pass@example.com','https://','ftp://example.com'])assert.equal(safeLiveLink(value),'');
});
test('Google content responses include sheet review link, including blank rows',()=>{
 const context=vm.createContext({console});vm.runInContext(fs.readFileSync(new URL('../apps-script/Code.gs',import.meta.url),'utf8'),context);
 assert.equal(context.contentView_({'Assignment ID':1,'Live Link':' https://example.com/review ','University Check Link':' https://university.example.edu '}).liveLink,'https://example.com/review');
 assert.equal(context.contentView_({'University Check Link':' https://university.example.edu '}).universityCheckLink,'https://university.example.edu');
 assert.equal(context.contentView_({'Assignment ID':2}).universityCheckLink,'');
 assert.equal(context.contentView_({'Assignment ID':2}).liveLink,'');
});
