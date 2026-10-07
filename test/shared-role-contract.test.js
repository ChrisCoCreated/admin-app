const test = require('node:test');
const assert = require('node:assert/strict');
const { ROLES, validateAssignment } = require('../api/_lib/shared-roles');
const { loadSnapshot, profileFor, saveAssignment, listAssignments } = require('../api/_lib/shared-role-store');

test('shared storage enforces eight roles, explicit revocation, pagination and rollout controls', async t => {
  const prior = {...process.env}, priorFetch=global.fetch;
  Object.assign(process.env,{SUPABASE_URL:'https://roles.example.com',SUPABASE_SERVICE_ROLE_KEY:'test',SUPER_USER_EMAIL:'root@example.com'});
  t.after(()=>{global.fetch=priorFetch;for(const key of Object.keys(process.env))if(!(key in prior))delete process.env[key];Object.assign(process.env,prior);});
  let rows=Array.from({length:1002},(_,i)=>({email:`user${String(i).padStart(4,'0')}@example.com`,roles:['care']}));
  rows.push({email:'revoked@example.com',roles:[]},{email:'manager@example.com',roles:['manager']});
  let state={access_enabled:true,writes_enabled:true}, failing=false, written;
  global.fetch=async(url,options={})=>{
    if(failing)return new Response('{}',{status:503});
    const target=new URL(url);
    assert.equal(options.headers.Authorization,'Bearer test');
    if(target.pathname.endsWith('/app_role_management_state'))return new Response(JSON.stringify([state]));
    if(target.pathname.endsWith('/app_role_aliases'))return new Response(JSON.stringify(Number(target.searchParams.get('offset')) ? [] : [{alias_email:'alias@example.com',canonical_email:'manager@example.com'}]));
    if(options.method==='POST'){written=JSON.parse(options.body);return new Response(null,{status:204});}
    const offset=Number(target.searchParams.get('offset'));
    // Simulate a REST row cap smaller than the requested page size.
    return new Response(JSON.stringify(rows.slice(offset,offset+200)));
  };
  assert.deepEqual(Object.keys(ROLES),['superadmin','manager','financeManager','consultant','careCoordinator','care','marketing','hr']);
  const snapshot=await loadSnapshot();
  assert.equal(Object.keys(snapshot.assignments).length,1004);
  assert.deepEqual(profileFor('revoked@example.com',snapshot).roles,[]);
  assert.deepEqual(profileFor('unknown@example.com',snapshot).roles,[]);
  assert.deepEqual(profileFor(' ROOT@example.com ',snapshot).roles,['superadmin']);
  assert.equal(profileFor('manager@example.com',snapshot).isSuperadmin,false);
  assert.equal(profileFor('alias@example.com',snapshot).email,'manager@example.com');
  assert.deepEqual(profileFor('alias@example.com',snapshot).roles,['manager']);
  await assert.rejects(saveAssignment({email:'manager@example.com',roles:['superadmin']},'manager@example.com'),{status:403});
  await assert.rejects(saveAssignment({email:'root@example.com',roles:[]},'root@example.com'),{status:409});
  await saveAssignment({email:' FINANCE@example.com ',roles:['financeManager','care','financeManager']},'root@example.com');
  assert.deepEqual(written.roles,['financeManager','care']);
  assert.equal(written.email,'finance@example.com');
  state.writes_enabled=false;
  await assert.rejects(saveAssignment({email:'new@example.com',roles:['care']},'root@example.com'),/paused/);
  assert.equal((await listAssignments()).writesEnabled,false);
  state.access_enabled=false;
  await assert.rejects(loadSnapshot(),{status:503});
  assert.ok(await loadSnapshot({requireActive:false}));
  failing=true;
  await assert.rejects(loadSnapshot(),{status:503});
  failing=false;state.access_enabled=true;rows=[{email:'broken@example.com',roles:['admin']}];
  await assert.rejects(loadSnapshot(),/migration review/);
});

test('validation rejects obsolete roles and retains explicit empty assignments',()=>{
  for(const body of [{email:'bad',roles:[]},{email:'x@example.com',roles:['finance']},{email:'x@example.com',roles:'manager'},{email:'x@example.com',roles:[null]}])assert.throws(()=>validateAssignment(body,'root@example.com'),{status:400});
  assert.deepEqual(validateAssignment({email:'x@example.com',roles:[]},'root@example.com'),{email:'x@example.com',roles:[]});
  assert.throws(()=>validateAssignment({email:'self@example.com',roles:['manager']},'root@example.com','self@example.com'),{status:409});
});

test('minimal insert success accepts an empty 201 response', async () => {
  const previousFetch = global.fetch;
  const previousUrl = process.env.SUPABASE_URL, previousKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
  process.env.SUPABASE_URL = 'https://roles.example.com';
  process.env.SUPABASE_SERVICE_ROLE_KEY = 'test';
  global.fetch = async () => new Response(null, { status: 201 });
  try { assert.equal(await require('../api/_lib/shared-role-store').request('app_role_assignments', {}, { method: 'POST' }), null); }
  finally {
    global.fetch = previousFetch;
    if (previousUrl === undefined) delete process.env.SUPABASE_URL; else process.env.SUPABASE_URL = previousUrl;
    if (previousKey === undefined) delete process.env.SUPABASE_SERVICE_ROLE_KEY; else process.env.SUPABASE_SERVICE_ROLE_KEY = previousKey;
  }
});
