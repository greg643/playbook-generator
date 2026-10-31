import assert from 'node:assert/strict';
import test from 'node:test';
import { createSessionCookie, emailKey, getUser, generatePasswordResetToken } from '../../dashboard/functions/_lib/auth.js';
import { appleKey, finishApple, sealApple } from '../../dashboard/functions/_lib/apple.js';
import { onRequestPost as login } from '../../dashboard/functions/api/auth/login.js';
import { onRequestPost as register } from '../../dashboard/functions/api/auth/register.js';
import { onRequestPost as recover } from '../../dashboard/functions/api/auth/recover.js';
import { onRequestPost as rotate } from '../../dashboard/functions/api/auth/recovery-code.js';
import { onRequestPost as resetRequest } from '../../dashboard/functions/api/auth/password-reset/request.js';
import { onRequestPost as resetComplete } from '../../dashboard/functions/api/auth/password-reset/complete.js';
import { onRequestPost as deleteAccount } from '../../dashboard/functions/api/auth/delete-account.js';
import { onRequestGet as me } from '../../dashboard/functions/api/auth/me.js';
import { onRequestGet as plays } from '../../dashboard/functions/api/plays.js';
import { onRequestPost as upload } from '../../dashboard/functions/api/upload.js';
import { onRequestPost as generate } from '../../dashboard/functions/api/generate.js';

const ORIGIN = 'https://playbook.example';
const APPLE = '0123456789abcdef';
const UID = '11111111-1111-4111-8111-111111111111';
class Bucket {
  records = new Map(); sequence = 0;
  async get(key) { const r=this.records.get(key); return r && {etag:r.etag,json:async()=>JSON.parse(r.text),text:async()=>r.text}; }
  async head(key) { return this.get(key); }
  async put(key,text,{onlyIf}={}) {
    const previous=this.records.get(key);
    if(onlyIf?.etagDoesNotMatch==='*'&&previous)return null;
    if(onlyIf?.etagMatches&&previous?.etag!==onlyIf.etagMatches)return null;
    this.records.set(key,{text,etag:String(++this.sequence)});return this.get(key);
  }
  async delete(keys) { for(const key of Array.isArray(keys)?keys:[keys])this.records.delete(key); }
  async list({prefix=''}={}) {return {objects:[...this.records.keys()].filter(k=>k.startsWith(prefix)).map(key=>({key})),truncated:false};}
}
function environment() { return {PLAYBOOK_BUCKET:new Bucket(),JOBS_BUCKET:new Bucket(),AUTH_STATE_BUCKET:new Bucket(),
  SESSION_SECRET:'a'.repeat(64),APPLE_ENABLED:'true',APPLE_SERVICES_ID:'com.gss.web',APPLE_TEAM_ID:'TEAM',
  APPLE_KEY_ID:'KEY',APPLE_PRIVATE_KEY:'configured',PUBLIC_ORIGIN:ORIGIN,GSS_API_BASE:'https://gss.example'}; }
function request(cookie='',body) {return new Request(ORIGIN+'/api/test',{headers:{cookie,origin:ORIGIN,'Content-Type':'application/json'},
  ...(body===undefined?{}:{method:'POST',body:JSON.stringify(body)})});}
async function seed(env, linked=true) {
  const email='coach@example.com',key=await emailKey(email);
  await env.PLAYBOOK_BUCKET.put(key,JSON.stringify({userId:UID,email,sessionVersion:1,hash:'legacy-hash',...(linked?{appleAccountKey:APPLE}:{})}));
  return {email,key};
}

test('Apple-only defaults retire all six legacy entry points before reading bodies, storage, or email', async () => {
  const forbidden=new Proxy({}, {get(){throw Error('Legacy endpoint touched storage/email');}});
  for(const value of [undefined,'false',true,'TRUE','']) {
    for(const handler of [login,register,recover,rotate,resetRequest,resetComplete]) {
      const request={text(){throw Error('Body read');},body:new ReadableStream()};
      const response=await handler({request,env:{LEGACY_EMAIL_AUTH_ENABLED:value,PLAYBOOK_BUCKET:forbidden,EMAIL_SERVICE:forbidden}});
      assert.equal(response.status,410);
      assert.equal((await response.json()).code,'apple_only');
      assert.match(response.headers.get('cache-control'),/no-store/);
      assert.equal(response.headers.get('set-cookie'),null);
    }
  }
});

test('old password sessions cannot access books, identity, uploads, generation, or deletion even for linked accounts', async () => {
  for(const linked of [false,true]) {
    const env=environment();const {email}=await seed(env,linked);
    const cookie=(await createSessionCookie(UID,email,env)).split(';')[0];
    assert.equal(await getUser(request(cookie),env),null);
    for(const handler of [me,plays,upload,generate]) {
      const response=await handler({env,request:request(cookie)});
      assert.equal(response.status,401);
    }
    const deleted=await deleteAccount({env,request:request(cookie,{userId:UID,email,password:'unused'})});
    assert.equal(deleted.status,401);
    assert.equal(env.PLAYBOOK_BUCKET.records.size,1);
  }
});

test('existing linked Apple sessions keep the same playbooks but expose no password recovery option', async () => {
  const env=environment();const {email,key}=await seed(env);
  const book=JSON.stringify({schema:2,offense:[],defense:[],defaultPlayersPerSide:5});
  await env.PLAYBOOK_BUCKET.put(`accounts/${UID}/playbook.json`,book);
  const cookie=(await createSessionCookie(UID,email,env,1,{credentialKey:key,appleAccountKey:APPLE})).split(';')[0];
  const response=await me({env,request:request(cookie)});
  assert.equal(response.status,200);
  assert.equal((await response.json()).passwordAvailable,false);
  assert.equal((await getUser(request(cookie),env)).userId,UID);
  assert.equal((await plays({env,request:request(cookie)})).status,200);
  assert.equal(await (await env.PLAYBOOK_BUCKET.get(`accounts/${UID}/playbook.json`)).text(),book);
  const wrong=(await createSessionCookie(UID,email,env,1,{credentialKey:key,appleAccountKey:'fedcba9876543210'})).split(';')[0];
  assert.equal(await getUser(request(wrong),env),null);
});

test('first-time Apple signup works in default policy, rejects replay and keeps legacy linking retired', async () => {
  const env=environment();
  const cookie='__Host-pb_apple_pending='+await sealApple(env,{kind:'pending',origin:ORIGIN,iat:Math.floor(Date.now()/1000),
    nonce:generatePasswordResetToken(),appleAccountKey:APPLE,displayName:'Coach',returnTo:'/converter'});
  const linking=await finishApple({env,request:request(cookie,{mode:'link',userId:UID})});
  assert.equal(linking.status,410);
  assert.equal(env.AUTH_STATE_BUCKET.records.size,0);
  const created=await finishApple({env,request:request(cookie,{mode:'new'})});
  assert.equal(created.status,200);
  assert.equal((await created.json()).returnTo,'/converter');
  const session=created.headers.getSetCookie().find(c=>c.startsWith('pb_session=')).split(';')[0];
  const user=await getUser(request(session),env);
  assert.equal(user.authMethod,'apple');
  assert.equal(user.account.hash,undefined);
  assert.equal(user.credentialKey,appleKey(APPLE));
  assert.equal((await finishApple({env,request:request(cookie,{mode:'new'})})).status,409);
});

test('Apple-only account deletion still requires recent Apple proof and typed confirmation', async () => {
  const env=environment();const {email,key}=await seed(env);
  await env.PLAYBOOK_BUCKET.put(appleKey(APPLE),JSON.stringify({kind:'link',credentialKey:key,userId:UID}));
  const cookie=(await createSessionCookie(UID,email,env,1,{credentialKey:key,appleAccountKey:APPLE})).split(';')[0];
  assert.equal((await deleteAccount({env,request:request(cookie,{userId:UID,confirm:'no'})})).status,403);
  assert.equal((await deleteAccount({env,request:request(cookie,{userId:UID,confirm:'DELETE'})})).status,200);
  assert.ok((await (await env.PLAYBOOK_BUCKET.get(key)).json()).deletedAt);
  assert.equal(await getUser(request(cookie),env),null);
});
