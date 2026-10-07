// Integration check using real bearer verification and an isolated mocked shared database.
const path=require('node:path'),assert=require('node:assert/strict'),crypto=require('node:crypto');
const associates=path.resolve(process.argv[2] || '../Associates App');
const adminRoles=require('../api/roles'),adminMe=require('../api/auth/me'),assocRoles=require(path.join(associates,'api/roles')),assocMe=require(path.join(associates,'api/auth/me'));
Object.assign(process.env,{SUPABASE_URL:'https://roles.example.com',SUPABASE_SERVICE_ROLE_KEY:'test',SUPER_USER_EMAIL:'root@example.com',AZURE_TENANT_ID:'test-tenant',AZURE_API_AUDIENCE:'test-api',AZURE_REQUIRED_SCOPE:'client.read'});
const {privateKey,publicKey}=crypto.generateKeyPairSync('rsa',{modulusLength:2048});
let rows=[{email:'rebecca@thrivehomecare.co.uk',roles:['manager']},{email:'person@example.com',roles:['care']}];
global.fetch=async(url,options={})=>{
 const target=String(url);
 if(target.includes('openid-configuration'))return new Response(JSON.stringify({jwks_uri:'https://identity.example.com/keys'}));
 if(target.endsWith('/keys'))return new Response(JSON.stringify({keys:[{...publicKey.export({format:'jwk'}),kid:'shared-test'}]}));
 if(target.includes('app_role_management_state'))return new Response(JSON.stringify([{access_enabled:true,writes_enabled:true}]));
 const offset=Number(new URL(target).searchParams.get('offset'));
 if(target.includes('app_role_aliases'))return new Response(JSON.stringify(offset?[]:[{alias_email:'rebecca@planwithcare.co.uk',canonical_email:'rebecca@thrivehomecare.co.uk'}]));
 if(options.method==='POST'){const row=JSON.parse(options.body);rows=rows.filter(item=>item.email!==row.email).concat(row);return new Response(null,{status:204});}
 return new Response(JSON.stringify(rows.slice(offset,offset+500)));
};
function token(email){const header=Buffer.from(JSON.stringify({alg:'RS256',kid:'shared-test'})).toString('base64url');const payload=Buffer.from(JSON.stringify({preferred_username:email,aud:'test-api',iss:'https://login.microsoftonline.com/test-tenant/v2.0',scp:'client.read',exp:Math.floor(Date.now()/1000)+600})).toString('base64url');const input=`${header}.${payload}`;return `${input}.${crypto.sign('RSA-SHA256',Buffer.from(input),privateKey).toString('base64url')}`;}
async function call(handler,email,method='GET',body={}){const output={code:200,setHeader(){},status(code){this.code=code;return this;},json(data){this.data=data;return this;}};await handler({method,headers:{authorization:`Bearer ${token(email)}`},body},output);return output;}
(async()=>{
 assert.equal((await call(adminRoles,'root@example.com','PUT',{email:'person@example.com',roles:['financeManager','hr']})).code,200);
 assert.deepEqual((await call(assocMe,'person@example.com')).data.roles,['financeManager','care','hr']);
 assert.equal((await call(assocRoles,'root@example.com','PUT',{email:'person@example.com',roles:['consultant','marketing']})).code,200);
 assert.deepEqual((await call(adminMe,'person@example.com')).data.roles,['consultant','marketing']);
 assert.equal((await call(adminRoles,'root@example.com','PUT',{email:'rebecca@planwithcare.co.uk',roles:['hr']})).code,200);
 assert.deepEqual((await call(assocMe,'rebecca@thrivehomecare.co.uk')).data.roles,['care','hr']);
 assert.deepEqual((await call(assocMe,'rebecca@planwithcare.co.uk')).data.roles,['care','hr']);
 assert.equal((await call(assocRoles,'root@example.com','PUT',{email:'rebecca@planwithcare.co.uk',roles:[]})).code,200);
 assert.equal((await call(adminMe,'rebecca@planwithcare.co.uk')).code,403);
 assert.deepEqual((await call(assocMe,'rebecca@thrivehomecare.co.uk')).data.roles,['care']);
 assert.deepEqual((await call(assocMe,'unassigned@example.com')).data.roles,['care']);
 assert.equal((await call(adminMe,'unassigned@example.com')).code,403);
 assert.equal((await call(assocRoles,'unassigned@example.com')).code,403);
 console.log('Passed: edits from either app apply in the other, linked identities share assignments, and cleared roles block Admin while Associates retains only baseline Care.');
})().catch(error=>{console.error(error);process.exitCode=1;});
