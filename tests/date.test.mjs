import test from 'node:test';
import assert from 'node:assert/strict';
import {dateLabel} from '../src/date.mjs';
test('calendar dates and ISO assignment dates keep their assigned calendar day',()=>{
 const previous=process.env.TZ;process.env.TZ='America/Los_Angeles';
 try{assert.equal(dateLabel('2026-10-16'),'16 Oct 2026');assert.equal(dateLabel('2026-10-16T00:00:00.000Z'),'16 Oct 2026');}finally{if(previous===undefined)delete process.env.TZ;else process.env.TZ=previous;}
});
test('dashboard header accepts Date objects and missing dates remain readable',()=>{
 assert.equal(dateLabel(new Date(2026,9,16,12)),'16 Oct 2026');assert.equal(dateLabel(''),'—');assert.equal(dateLabel('invalid'),'—');
});
