const test = require("node:test");
const assert = require("node:assert/strict");
const crypto = require("node:crypto");
const fs = require("node:fs");
const vm = require("node:vm");
const handler = require("../api/roles");
const me = require("../api/auth/me");
const { requireApiAuth } = require("../api/_lib/require-api-auth");
const { requireGraphAuth } = require("../api/_lib/require-graph-auth");

test("verified roles enforce finance boundaries, combinations, live edits and Graph revocations", async t => {
  const prior={...process.env}, priorFetch=global.fetch;
  Object.assign(process.env,{SUPABASE_URL:"https://roles.example.com",SUPABASE_SERVICE_ROLE_KEY:"test",SUPER_USER_EMAIL:"root@example.com",AZURE_TENANT_ID:"test-tenant",AZURE_API_AUDIENCE:"test-api",AZURE_REQUIRED_SCOPE:"client.read",ACCESS_FULL_EMAILS:"unassigned@example.com"});
  t.after(()=>{global.fetch=priorFetch;for(const key of Object.keys(process.env))if(!(key in prior))delete process.env[key];Object.assign(process.env,prior);});
  const {privateKey,publicKey}=crypto.generateKeyPairSync("rsa",{modulusLength:2048});
  let rows=[{email:"manager@example.com",roles:["manager"]},{email:"finance@example.com",roles:["financeManager"]},{email:"care@example.com",roles:["care"]},{email:"combo@example.com",roles:["consultant","financeManager"]}];
  let outage=false,graphEmail="finance@example.com",writes=0;
  global.fetch=async(url,options={})=>{
    const target=String(url);
    if(target.includes("openid-configuration"))return new Response(JSON.stringify({jwks_uri:"https://identity.example.com/keys"}));
    if(target.endsWith("/keys"))return new Response(JSON.stringify({keys:[{...publicKey.export({format:"jwk"}),kid:"test-key"}]}));
    if(target.includes("graph.microsoft.com"))return new Response(JSON.stringify({mail:graphEmail,userPrincipalName:graphEmail}));
    if(outage)return new Response('{}',{status:503});
    if(target.includes('app_role_management_state'))return new Response(JSON.stringify([{access_enabled:true,writes_enabled:true}]));
    if(target.includes('app_role_aliases'))return new Response('[]');
    if(options.method==='POST'){const row=JSON.parse(options.body);rows=rows.filter(r=>r.email!==row.email).concat(row);writes++;return new Response(null,{status:204});}
    const offset=Number(new URL(target).searchParams.get('offset'));
    return new Response(JSON.stringify(rows.slice(offset,offset+500)));
  };
  function token(email){const header=Buffer.from(JSON.stringify({alg:"RS256",kid:"test-key"})).toString('base64url');const payload=Buffer.from(JSON.stringify({preferred_username:email,aud:"test-api",iss:"https://login.microsoftonline.com/test-tenant/v2.0",scp:"client.read",exp:Math.floor(Date.now()/1000)+600})).toString('base64url');const input=`${header}.${payload}`;return `${input}.${crypto.sign('RSA-SHA256',Buffer.from(input),privateKey).toString('base64url')}`;}
  const res=()=>({statusCode:200,setHeader(){},status(code){this.statusCode=code;return this;},json(data){this.data=data;return this;}});
  const req=(email,method='GET',body={})=>({method,headers:{authorization:`Bearer ${token(email)}`},body});
  async function call(email,method='GET',body={},route=handler){const output=res();await route(req(email,method,body),output);return output;}
  for(const email of ['manager@example.com','finance@example.com','care@example.com','unassigned@example.com'])assert.equal((await call(email)).statusCode,403);
  assert.equal(writes,0);
  assert.equal((await call('root@example.com')).statusCode,200);
  assert.equal((await call('root@example.com','PUT',{email:'root@example.com',roles:[]})).statusCode,409);
  async function allowed(email,allowedRoles){const output=res();const input=req(email);const claims=await requireApiAuth(input,output,{allowedRoles});return {claims,output,input};}
  assert.ok((await allowed('finance@example.com',['finance'])).claims);
  assert.equal((await allowed('finance@example.com',['admin'])).output.statusCode,403);
  assert.equal((await allowed('care@example.com',['finance'])).output.statusCode,403);
  assert.ok((await allowed('combo@example.com',['consultant'])).claims);
  assert.ok((await allowed('combo@example.com',['finance'])).claims);
  assert.ok((await allowed('manager@example.com',['admin'])).claims);
  assert.equal((await call('finance@example.com','GET',{},me)).data.isSuperadmin,false);
  assert.equal((await call('root@example.com','PUT',{email:'finance@example.com',roles:['superadmin','financeManager']})).statusCode,200);
  assert.equal((await call('finance@example.com')).statusCode,200);
  assert.equal((await call('finance@example.com','PUT',{email:'finance@example.com',roles:['financeManager']})).statusCode,409);
  assert.equal((await call('root@example.com','PUT',{email:'finance@example.com',roles:[]})).statusCode,200);
  assert.equal((await call('finance@example.com')).statusCode,403);
  const graphRes=res();assert.equal(await requireGraphAuth(req('finance@example.com'),graphRes),null);assert.equal(graphRes.statusCode,403);
  outage=true;
  assert.equal((await call('root@example.com')).statusCode,503);
});

test("navigation exposes role management only with the superadmin capability", () => {
  const code = fs.readFileSync(require.resolve("../navigation.js"), "utf8")
    .replace(/^import[\s\S]*?from "\.\/role-preview[^\n]+\n/, "").replaceAll("export function", "function");
  const storage = new Map();
  const context = vm.createContext({ window: { sessionStorage: { getItem: (key) => storage.get(key) } } });
  vm.runInContext(code, context);
  assert.equal(context.canAccessPage("admin", "rolemanagement"), false);
  storage.set("thrive.access.superadmin", "true");
  assert.equal(context.canAccessPage("admin", "rolemanagement"), true);
  assert.equal(context.canAccessPage("marketing", "rolemanagement"), false);
  assert.equal(context.canAccessPage("logged_in", "rolemanagement"), false);
  assert.equal(context.canAccessPage("pages:rolemanagement", "rolemanagement"), false);
});
