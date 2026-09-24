import assert from 'node:assert/strict';
import test from 'node:test';
import { appleConfigured, appleKey, callbackApple, findAppleAccount, finishApple, pendingApple, sealApple, startApple } from '../../dashboard/functions/_lib/apple.js';
import { createSessionCookie, emailKey, getUser, generatePasswordResetToken, sha256Hex } from '../../dashboard/functions/_lib/auth.js';
import { onRequestPost as deleteAccount } from '../../dashboard/functions/api/auth/delete-account.js';
import { onRequestPost as register } from '../../dashboard/functions/api/auth/register.js';
import { onRequestGet as getPlays } from '../../dashboard/functions/api/plays.js';
import { onRequestPost as recoveryCode } from '../../dashboard/functions/api/auth/recovery-code.js';
import { onRequestPost as login } from '../../dashboard/functions/api/auth/login.js';
import { onRequestPost as resetRequest } from '../../dashboard/functions/api/auth/password-reset/request.js';
import { onRequestPost as resetComplete } from '../../dashboard/functions/api/auth/password-reset/complete.js';

const ORIGIN = 'https://playbook.example';
const APPLE_ID = '0123456789abcdef';
const UID = '11111111-1111-4111-8111-111111111111';
class Bucket {
  records = new Map(); sequence = 0; beforePut = null;
  async get(key) {
    const entry = this.records.get(key);
    return entry ? { key, etag: entry.etag, uploaded: new Date(), json: async () => JSON.parse(entry.text), text: async () => entry.text } : null;
  }
  async head(key) { return this.get(key); }
  async put(key, text, { onlyIf } = {}) {
    if (this.beforePut) await this.beforePut(key);
    const previous = this.records.get(key);
    if (onlyIf?.etagDoesNotMatch === '*' && previous) return null;
    if (onlyIf?.etagMatches && previous?.etag !== onlyIf.etagMatches) return null;
    this.records.set(key, { text, etag: String(++this.sequence) });
    return this.get(key);
  }
  async delete(keys) { for (const key of Array.isArray(keys) ? keys : [keys]) this.records.delete(key); }
  async list({ prefix = '' } = {}) { return { objects: [...this.records.keys()].filter(k => k.startsWith(prefix)).map(key => ({key})), truncated: false }; }
}
const signingKeys = await crypto.subtle.generateKey({ name:'ECDSA', namedCurve:'P-256' }, true, ['sign','verify']);
const pem = '-----BEGIN PRIVATE KEY-----\n' + Buffer.from(await crypto.subtle.exportKey('pkcs8', signingKeys.privateKey)).toString('base64') + '\n-----END PRIVATE KEY-----';
function environment() {
  return { PLAYBOOK_BUCKET:new Bucket(), JOBS_BUCKET:new Bucket(), AUTH_STATE_BUCKET:new Bucket(), SESSION_SECRET:'a'.repeat(64), APPLE_ENABLED:'true',
    APPLE_SERVICES_ID:'com.gss.web', APPLE_TEAM_ID:'TEAM', APPLE_KEY_ID:'KEY', APPLE_PRIVATE_KEY:pem,
    PUBLIC_ORIGIN:ORIGIN, GSS_API_BASE:'https://gss.example' };
}
function request(path, cookie = '', body, origin = ORIGIN) {
  return new Request(ORIGIN + path, { method:body === undefined ? 'GET' : 'POST', headers:{cookie,origin,'Content-Type':'application/json'},
    ...(body === undefined ? {} : {body:JSON.stringify(body)}) });
}
function sessionCookie(response) { return response.headers.getSetCookie().find(c => c.startsWith('pb_session=')).split(';')[0]; }
async function pending(env, changes = {}) {
  return '__Host-pb_apple_pending=' + await sealApple(env, {kind:'pending',origin:ORIGIN,iat:Math.floor(Date.now()/1000),
    nonce:generatePasswordResetToken(),appleAccountKey:APPLE_ID,displayName:'Coach Example',returnTo:'/editor',...changes});
}
async function legacy(env) {
  const email = 'coach@example.com';
  const key = await emailKey(email);
  const record = {userId:UID,email,sessionVersion:1,salt:'b'.repeat(32),hash:'c'.repeat(64),iterations:100000,
    recoveryHash:'d'.repeat(64),recoverySalt:'b'.repeat(32),recoveryIterations:100000,passwordReset:{tokenHash:'x'}};
  await env.PLAYBOOK_BUCKET.put(key,JSON.stringify(record));
  const cookie = (await createSessionCookie(UID,email,env)).split(';')[0];
  return {key,record,cookie};
}

test('Apple configuration fails closed on previews or absent settings', () => {
  const env = environment();
  assert.equal(appleConfigured(env,request('/')),true);
  assert.equal(appleConfigured(env,new Request('https://preview.example/')),false);
  assert.equal(appleConfigured({...env,APPLE_ENABLED:'false'},request('/')),false);
  assert.equal(appleConfigured({...env,AUTH_STATE_BUCKET:null},request('/')),false);
});

test('Apple start uses a fixed callback, unpredictable state, nonce, and safe return path', async () => {
  const env = environment();
  const response = await startApple({env,request:request('/api/auth/apple/start?returnTo=https://evil.example/')});
  const url = new URL(response.headers.get('location'));
  assert.equal(url.origin,'https://appleid.apple.com');
  assert.equal(url.searchParams.get('redirect_uri'),ORIGIN+'/api/auth/apple/callback');
  assert.match(url.searchParams.get('state'),/^[\w-]{43}$/);
  assert.match(url.searchParams.get('nonce'),/^[a-f0-9]{64}$/);
  assert.equal(url.searchParams.get('scope'),'name');
  assert.match(response.headers.get('set-cookie'),/HttpOnly; Secure; SameSite=None/);
  const other = await startApple({env,request:request('/api/auth/apple/start')});
  assert.notEqual(new URL(other.headers.get('location')).searchParams.get('state'),url.searchParams.get('state'));
});

test('new Apple accounts need no email or password and can read only their own saved book', async () => {
  const env = environment(); const cookie = await pending(env);
  const response = await finishApple({env,request:request('/api/auth/apple/finish',cookie,{mode:'new'})});
  assert.equal(response.status,200);
  const user = await getUser(request('/',sessionCookie(response)),env);
  assert.equal(user.authMethod,'apple'); assert.equal(user.email,'');
  assert.equal(user.account.appleAccountKey,APPLE_ID); assert.equal(user.account.hash,undefined);
  const book = await getPlays({env,request:request('/api/plays',sessionCookie(response))});
  assert.equal(book.status,200);
  const replay = await finishApple({env,request:request('/api/auth/apple/finish',cookie,{mode:'new'})});
  assert.equal(replay.status,409);
});

test('explicit Apple linking preserves playbooks and password recovery while revoking old sessions and reset links', async () => {
  const env=environment(); const existing=await legacy(env);
  const bookKey=`accounts/${UID}/playbook.json`;
  await env.PLAYBOOK_BUCKET.put(bookKey,JSON.stringify({schema:2,offense:[{name:'Saved play'}],defense:[]}));
  const cookie=await pending(env)+'; '+existing.cookie;
  const response=await finishApple({env,request:request('/api/auth/apple/finish',cookie,{mode:'link',userId:UID})});
  assert.equal(response.status,200);
  const user=await getUser(request('/',sessionCookie(response)),env);
  assert.equal(user.userId,UID); assert.equal(user.email,existing.record.email);
  for(const field of ['hash','salt','recoveryHash','recoverySalt']) assert.equal(user.account[field],existing.record[field]);
  assert.equal(user.account.passwordReset,undefined);
  assert.equal(await getUser(request('/',existing.cookie),env),null);
  assert.equal((await (await env.PLAYBOOK_BUCKET.get(bookKey)).json()).offense[0].name,'Saved play');
  const linked=await findAppleAccount(env,APPLE_ID);
  assert.equal(linked.credentialKey,existing.key);
});

test('linking requires both identities, matching local account, and same-origin JSON', async () => {
  const env=environment(); const old=await legacy(env); const proof=await pending(env);
  assert.equal((await finishApple({env,request:request('/api/auth/apple/finish',proof,{mode:'link',userId:UID})})).status,401);
  assert.equal((await finishApple({env,request:request('/api/auth/apple/finish',proof+'; '+old.cookie,{mode:'link',userId:'other'})})).status,401);
  assert.equal((await finishApple({env,request:request('/api/auth/apple/finish',proof+'; '+old.cookie,{mode:'link',userId:UID},'https://evil.example')})).status,403);
  assert.equal((await finishApple({env,request:request('/api/auth/apple/finish',proof+'; '+old.cookie,{mode:'new'})})).status,409);
  assert.equal((await (await env.PLAYBOOK_BUCKET.get(old.key)).json()).hash,old.record.hash);
  assert.equal(await env.PLAYBOOK_BUCKET.get(appleKey(APPLE_ID)),null);
});

test('Apple proof expires, is origin-bound, and rejects tampering', async () => {
  const env=environment();
  for(const cookie of [await pending(env,{iat:Math.floor(Date.now()/1000)-601}), await pending(env,{origin:'https://other.example'}), (await pending(env))+'x']) {
    assert.equal((await pendingApple({env,request:request('/',cookie)})).status,401);
    assert.equal((await finishApple({env,request:request('/',cookie,{mode:'new'})})).status,401);
  }
});

test('concurrent first Apple visits produce only one account', async () => {
  const env=environment();
  const proofs=await Promise.all([pending(env),pending(env)]);
  const results=await Promise.all(proofs.map(cookie=>finishApple({env,request:request('/',cookie,{mode:'new'})})));
  assert.deepEqual(results.map(r=>r.status).sort(),[200,409]);
  assert.equal(env.PLAYBOOK_BUCKET.records.size,1);
});

test('failed link CAS keeps old credentials and allows explicit retry', async () => {
  const env=environment(); const old=await legacy(env);
  env.PLAYBOOK_BUCKET.beforePut=async key=>{
    if(key!==old.key)return;
    env.PLAYBOOK_BUCKET.beforePut=null;
    await env.PLAYBOOK_BUCKET.put(old.key,JSON.stringify({...old.record,otherUpdate:true}));
  };
  const first=await finishApple({env,request:request('/',await pending(env)+'; '+old.cookie,{mode:'link',userId:UID})});
  assert.equal(first.status,409);
  assert.equal((await (await env.PLAYBOOK_BUCKET.get(old.key)).json()).hash,old.record.hash);
  assert.equal(await findAppleAccount(env,APPLE_ID),null);
  const retry=await finishApple({env,request:request('/',await pending(env)+'; '+old.cookie,{mode:'link',userId:UID})});
  assert.equal(retry.status,200);
});

test('a different Apple account cannot silently take over an existing linked account', async () => {
  const env=environment(); const old=await legacy(env);
  const response=await finishApple({env,request:request('/',await pending(env)+'; '+old.cookie,{mode:'link',userId:UID})});
  const newProof=await pending(env,{appleAccountKey:'fedcba9876543210'});
  assert.equal((await finishApple({env,request:request('/',newProof+'; '+sessionCookie(response),{mode:'link',userId:UID})})).status,409);
  assert.equal((await findAppleAccount(env,APPLE_ID)).record.userId,UID);
});

test('Apple account deletion removes playbooks and mapping without affecting the GSS app', async () => {
  const env=environment(); const old=await legacy(env);
  await env.PLAYBOOK_BUCKET.put(`accounts/${UID}/playbook.json`,JSON.stringify({offense:[],defense:[]}));
  const response=await finishApple({env,request:request('/',await pending(env)+'; '+old.cookie,{mode:'link',userId:UID})});
  const cookie=sessionCookie(response);
  assert.equal((await deleteAccount({env,request:request('/api/auth/delete-account',cookie,{userId:UID,confirm:'wrong'})})).status,403);
  assert.equal((await deleteAccount({env,request:request('/api/auth/delete-account',cookie,{userId:UID,confirm:'DELETE'})})).status,200);
  assert.equal(await env.PLAYBOOK_BUCKET.get(`accounts/${UID}/playbook.json`),null);
  assert.equal(await findAppleAccount(env,APPLE_ID),null);
  assert.equal(await getUser(request('/',cookie),env),null);
  const fresh=await finishApple({env,request:request('/',await pending(env),{mode:'new'})});
  assert.equal(fresh.status,200);
  assert.notEqual((await getUser(request('/',sessionCookie(fresh)),env)).userId,UID);
});

test('enabling Apple keeps new email/password registrations available', async () => {
  const env=environment();
  assert.equal((await register({env,request:request('/api/auth/register','',{email:'new@example.com',password:'password'})})).status,200);
  assert.equal(env.PLAYBOOK_BUCKET.records.size,1);
});

test('a linked coach can use either login and reset a password without losing Apple or playbooks', async () => {
  const env=environment(), email='dual@example.com', password='first-test-password';
  const registered=await register({env,request:request('/api/auth/register','',{email,password})});
  assert.equal(registered.status,200);
  const original=await registered.clone().json();
  const linked=await finishApple({env,request:request('/',await pending(env)+'; '+sessionCookie(registered),{mode:'link',userId:original.userId})});
  assert.equal(linked.status,200);
  const appleSession=sessionCookie(linked);
  const passwordLogin=await login({env,request:request('/api/auth/login','',{email,password})});
  assert.equal(passwordLogin.status,200);
  assert.equal((await getUser(request('/',sessionCookie(passwordLogin)),env)).userId,original.userId);
  assert.equal((await getUser(request('/',appleSession),env)).userId,original.userId);
  assert.equal((await recoveryCode({env,request:request('/',sessionCookie(passwordLogin),{})})).status,200);
  let delivered;
  env.EMAIL_SERVICE={fetch:async(url,options)=>{
    if(url.endsWith('/permit')) return Response.json({allowed:true});
    delivered=JSON.parse(options.body);
    return Response.json({ok:true});
  }};
  assert.equal((await resetRequest({env,request:request('/api/auth/password-reset/request','',{email})})).status,202);
  assert.equal(delivered.email,email);
  const reset=await resetComplete({env,request:request('/api/auth/password-reset/complete','',{email,token:delivered.token,newPassword:'replacement-password'})});
  assert.equal(reset.status,200);
  assert.equal(await getUser(request('/',appleSession),env),null);
  assert.equal((await findAppleAccount(env,APPLE_ID)).record.userId,original.userId);
  const after=await login({env,request:request('/api/auth/login','',{email,password:'replacement-password'})});
  assert.equal(after.status,200);
  assert.equal((await getUser(request('/',sessionCookie(after)),env)).userId,original.userId);
  const removed=await deleteAccount({env,request:request('/api/auth/delete-account',sessionCookie(after),{userId:original.userId,password:'replacement-password'})});
  assert.equal(removed.status,200);
  assert.equal(await findAppleAccount(env,APPLE_ID),null);
});

test('callback exchanges the code server-side and delegates token and nonce verification to GSS exactly once', async t => {
  const env=environment(); const state=generatePasswordResetToken(),rawNonce=generatePasswordResetToken();
  const txn=await sealApple(env,{kind:'transaction',origin:ORIGIN,iat:Math.floor(Date.now()/1000),state,rawNonce,returnTo:'/converter'});
  const calls=[];
  t.mock.method(globalThis,'fetch',async (url,options)=>{
    calls.push(url);
    if(url==='https://appleid.apple.com/auth/token') {
      assert.equal(options.body.get('code'),'one-use-code');
      const [head,payload,sig]=options.body.get('client_secret').split('.');
      assert.equal(JSON.parse(Buffer.from(payload,'base64url')).sub,env.APPLE_SERVICES_ID);
      assert.equal(await crypto.subtle.verify({name:'ECDSA',hash:'SHA-256'},signingKeys.publicKey,Buffer.from(sig,'base64url'),new TextEncoder().encode(head+'.'+payload)),true);
      return Response.json({id_token:'server-exchanged-token'});
    }
    const body=JSON.parse(options.body);
    assert.equal(body.identity_token,'server-exchanged-token');
    assert.equal(body.display_name,undefined,'repeat consent must not overwrite the GSS name with a fallback');
    assert.equal(await sha256Hex(body.raw_nonce),await sha256Hex(rawNonce));
    return Response.json({account_key:APPLE_ID,session_token:'verified-gss-session',display_name:'Coach'});
  });
  const makeRequest=()=>new Request(ORIGIN+'/api/auth/apple/callback',{method:'POST',headers:{cookie:'__Host-pb_apple_txn='+txn,'Content-Type':'application/x-www-form-urlencoded'},body:new URLSearchParams({state,code:'one-use-code',id_token:'attacker-controlled-token'})});
  const response=await callbackApple({env,request:makeRequest()});
  assert.equal(response.headers.get('location'),'/apple-welcome');
  assert.equal(env.PLAYBOOK_BUCKET.records.size,0);
  assert.equal(calls.length,2);
  const replay=await callbackApple({env,request:makeRequest()});
  assert.equal(replay.headers.get('location'),'/?apple=failed&apple_step=transaction_used');
  assert.equal(calls.length,2);
});

test('invalid callback state never reaches an identity provider', async t => {
  const env=environment(); let calls=0;
  t.mock.method(globalThis,'fetch',async()=>{calls++; throw new Error('unexpected');});
  const response=await callbackApple({env,request:new Request(ORIGIN+'/api/auth/apple/callback',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:'state=wrong&code=wrong'})});
  assert.equal(response.headers.get('location'),'/?apple=failed&apple_step=transaction_missing'); assert.equal(calls,0);
});

test('callback diagnostics identify failed stages without disclosing provider bodies or credentials', async t => {
  const logs=[];
  t.mock.method(console,'error',value=>logs.push(value));
  let scenario;
  const secretMarker='DO-NOT-LOG-provider-token-or-key';
  t.mock.method(globalThis,'fetch',async url=>{
    const isApple=url==='https://appleid.apple.com/auth/token';
    if(scenario==='network') throw new Error(secretMarker);
    if(scenario==='client') return Response.json({error:'invalid_client',error_description:secretMarker},{status:400});
    if(scenario==='code') return Response.json({error:'invalid_grant',error_description:secretMarker},{status:400});
    if(scenario==='unknown') return Response.json({error:secretMarker},{status:500});
    if(scenario==='oversized') return new Response(secretMarker.repeat(3000));
    if(scenario==='bad_json') return new Response(secretMarker,{status:502});
    if(scenario==='null') return Response.json(null);
    if(isApple) return Response.json({id_token:secretMarker});
    return Response.json({error:{code:secretMarker}},{status:scenario==='rate' ? 429 : 401});
  });
  for(const [kind,expected] of [
    ['key','signing_key'],['network','apple_exchange'],['client','apple_client'],['code','apple_code'],
    ['unknown','apple_response'],['oversized','apple_response'],['bad_json','apple_response'],
    ['null','apple_response'],['gss','gss_rejected'],['rate','gss_rate_limit'],
  ]) {
    scenario=kind;
    const env=environment();
    if(kind==='key') env.APPLE_PRIVATE_KEY=secretMarker;
    const state=generatePasswordResetToken();
    const txn=await sealApple(env,{kind:'transaction',origin:ORIGIN,iat:Math.floor(Date.now()/1000),state,rawNonce:generatePasswordResetToken(),returnTo:'/editor'});
    const response=await callbackApple({env,request:new Request(ORIGIN+'/api/auth/apple/callback',{
      method:'POST',headers:{cookie:'__Host-pb_apple_txn='+txn,'Content-Type':'application/x-www-form-urlencoded'},
      body:new URLSearchParams({state,code:secretMarker}),
    })});
    assert.equal(response.headers.get('location'),'/?apple=failed&apple_step='+expected,kind);
    assert.deepEqual(JSON.parse(logs.at(-1)),{event:'apple_callback_failed',code:expected});
    assert.equal(JSON.stringify([...response.headers]).includes(secretMarker),false);
    assert.equal(env.PLAYBOOK_BUCKET.records.size,0);
  }
  assert.equal(logs.join('').includes(secretMarker),false);
});

test('expired and mismatched callback transactions return safe actionable diagnostics', async t => {
  t.mock.method(console,'error',()=>{});
  let providerCalls=0;
  t.mock.method(globalThis,'fetch',async()=>{providerCalls++;throw new Error('must not call');});
  for(const [kind,expected] of [['expired','transaction_invalid'],['mismatch','state_mismatch']]) {
    const env=environment(), state=generatePasswordResetToken();
    const txn=await sealApple(env,{kind:'transaction',origin:ORIGIN,iat:Math.floor(Date.now()/1000)-(kind==='expired'?601:0),state,rawNonce:generatePasswordResetToken(),returnTo:'/editor'});
    const response=await callbackApple({env,request:new Request(ORIGIN+'/api/auth/apple/callback',{
      method:'POST',headers:{cookie:'__Host-pb_apple_txn='+txn,'Content-Type':'application/x-www-form-urlencoded'},
      body:new URLSearchParams({state:kind==='mismatch'?'wrong':state,code:'code'}),
    })});
    assert.equal(response.headers.get('location'),'/?apple=failed&apple_step='+expected);
    assert.equal(env.AUTH_STATE_BUCKET.records.size,0);
  }
  assert.equal(providerCalls,0);
});

test('invalid setup JSON and oversized bodies are rejected before consuming Apple proof', async () => {
  const env=environment(), cookie=await pending(env);
  for(const body of ['{broken', 'null', '[]']) {
    const req=new Request(ORIGIN+'/api/auth/apple/finish',{method:'POST',headers:{cookie,origin:ORIGIN,'Content-Type':'application/json'},body});
    assert.equal((await finishApple({env,request:req})).status,400);
  }
  assert.equal((await finishApple({env,request:request('/',cookie,{mode:'new',padding:'x'.repeat(2048)})})).status,413);
  assert.equal(env.AUTH_STATE_BUCKET.records.size,0);
});

test('Apple sessions cannot issue recovery credentials and honor account revocation', async () => {
  const env=environment();
  const result=await finishApple({env,request:request('/',await pending(env),{mode:'new'})});
  const cookie=sessionCookie(result);
  assert.equal((await recoveryCode({env,request:request('/',cookie,{})})).status,409);
  const key=appleKey(APPLE_ID), account=await (await env.PLAYBOOK_BUCKET.get(key)).json();
  await env.PLAYBOOK_BUCKET.put(key,JSON.stringify({...account,sessionVersion:account.sessionVersion+1}));
  assert.equal(await getUser(request('/',cookie),env),null);
});

test('interrupted migrated-account deletion retains the Apple mapping for authenticated retry', async t => {
  const env=environment(), old=await legacy(env);
  const linked=await finishApple({env,request:request('/',await pending(env)+'; '+old.cookie,{mode:'link',userId:UID})});
  const cookie=sessionCookie(linked);
  assert.equal((await deleteAccount({env,request:request('/',cookie,{userId:UID,confirm:'DELETE'},'https://evil.example')})).status,403);
  env.PLAYBOOK_BUCKET.beforePut=async key=>{
    const current=key===old.key && await (await env.PLAYBOOK_BUCKET.get(key)).json();
    if(current?.disabledAt) throw new Error('simulated storage outage during finalization');
  };
  t.mock.method(console,'error',()=>{});
  const failed=await deleteAccount({env,request:request('/',cookie,{userId:UID,confirm:'DELETE'})});
  assert.equal(failed.status,503);
  assert.ok((await findAppleAccount(env,APPLE_ID)).record.disabledAt);
  assert.equal(await getUser(request('/',cookie),env),null);
  env.PLAYBOOK_BUCKET.beforePut=null;
  assert.equal((await deleteAccount({env,request:request('/',cookie,{userId:UID,confirm:'DELETE'})})).status,200);
  assert.equal(await findAppleAccount(env,APPLE_ID),null);
});

test('linking requires fresh local authentication, not merely a valid month-long session', async t => {
  const env=environment(), clock=Date.now;
  t.mock.method(Date,'now',()=>clock()-11*60*1000);
  const old=await legacy(env);
  t.mock.restoreAll();
  const response=await finishApple({env,request:request('/',await pending(env)+'; '+old.cookie,{mode:'link',userId:UID})});
  assert.equal(response.status,401);
  assert.equal((await (await env.PLAYBOOK_BUCKET.get(old.key)).json()).hash,old.record.hash);
});

test('known Apple identities return to their requested tool; disabled identities can only resume deletion', async t => {
  const env=environment();
  await finishApple({env,request:request('/',await pending(env),{mode:'new'})});
  t.mock.method(globalThis,'fetch',async url=>url==='https://appleid.apple.com/auth/token'
    ? Response.json({id_token:'server-token'})
    : Response.json({account_key:APPLE_ID,session_token:'verified-gss-session'}));
  for (const disabled of [false,true]) {
    if(disabled) {
      const key=appleKey(APPLE_ID), record=await (await env.PLAYBOOK_BUCKET.get(key)).json();
      await env.PLAYBOOK_BUCKET.put(key,JSON.stringify({...record,disabledAt:new Date().toISOString()}));
    }
    const state=generatePasswordResetToken();
    const txn=await sealApple(env,{kind:'transaction',origin:ORIGIN,iat:Math.floor(Date.now()/1000),state,rawNonce:generatePasswordResetToken(),returnTo:'/converter'});
    const response=await callbackApple({env,request:new Request(ORIGIN+'/api/auth/apple/callback',{method:'POST',headers:{cookie:'__Host-pb_apple_txn='+txn,'Content-Type':'application/x-www-form-urlencoded'},body:new URLSearchParams({state,code:'code'})})});
    assert.equal(response.headers.get('location'),disabled ? '/apple-delete' : '/converter');
    const user=await getUser(request('/',sessionCookie(response)),env);
    assert.equal(Boolean(user),!disabled);
    assert.equal((await getUser(request('/',sessionCookie(response)),env,{allowDisabled:true})).authMethod,'apple');
  }
});
