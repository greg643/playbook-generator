import assert from 'node:assert/strict';
import test from 'node:test';
import vm from 'node:vm';
import { readFileSync } from 'node:fs';

const script=readFileSync(new URL('../../dashboard/apple-finishing.js',import.meta.url),'utf8');
function run(fetch) {
  const elements={spinner:{hidden:false},heading:{textContent:'Finishing sign-in…'},status:{textContent:'Please wait'},retry:{hidden:true}};
  let removed=false;
  const redirects=[];
  const ready=vm.runInNewContext(script,{fetch,URL,AbortSignal,
    location:{origin:'https://playbook.example',replace:p=>redirects.push(p)},
    document:{querySelector:()=>({content:'encrypted-handoff',remove(){removed=true;}}),getElementById:id=>elements[id]}});
  return {ready,elements,redirects,get removed(){return removed;}};
}
test('finishing screen submits once and stays in progress throughout a slow exchange',async()=>{
  let release,calls=0;
  const p=run(async(path,options)=>{
    calls++;assert.equal(path,'/api/auth/apple/complete');
    assert.equal(options.credentials,'same-origin');
    assert.deepEqual(JSON.parse(options.body),{handoff:'encrypted-handoff'});
    return new Promise(resolve=>{release=()=>resolve(Response.json({redirect:'/editor'}));});
  });
  assert.equal(p.removed,true);
  assert.equal(p.elements.retry.hidden,true);
  assert.equal(p.elements.spinner.hidden,false);
  assert.equal(calls,1);
  release();await p.ready;
  assert.deepEqual(p.redirects,['/editor']);assert.equal(calls,1);
});
test('failed or unsafe completion offers a fresh sign-in, never an automatic resubmission',async()=>{
  for(const response of [new Response('',{status:503}),Response.json({redirect:'https://evil.example/editor'}),
    Response.json({redirect:'/api/auth/delete-account'}),Response.json({})]) {
    let calls=0;const p=run(async()=>{calls++;return response;});await p.ready;
    assert.equal(p.elements.retry.hidden,false);assert.equal(p.elements.spinner.hidden,true);
    assert.equal(p.redirects.length,0);assert.equal(calls,1);
  }
  const p=run(async()=>{throw Error('timeout');});await p.ready;
  assert.equal(p.elements.retry.hidden,false);
});
