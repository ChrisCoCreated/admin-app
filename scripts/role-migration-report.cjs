const fs = require('node:fs');
const path = require('node:path');
const { ROLES, normalizeEmail } = require('../api/_lib/shared-roles');
const { getAuthorizedUsersMap } = require('../api/_lib/authorized-users');
const mapped = { superadmin:'superadmin',admin:'manager',finance:'financeManager',operations:'careCoordinator',care_manager:'careCoordinator',consultant:'consultant',photo_layout:'consultant',enquiries_only:'consultant',marketing:'marketing',hr_only:'hr',logged_in:'care' };
function readEnv(file) {
  const env = {};
  for (const line of fs.readFileSync(file,'utf8').split(/\r?\n/)) {
    const match = /^\s*([A-Z_0-9]+)\s*=\s*(.*)$/.exec(line);
    if (match) env[match[1]] = match[2].replace(/^(['"])(.*)\1$/, '$2');
  }
  return env;
}
function buildReport({ adminUsers, associatesAccess, existing = [], sourcesVerified = false, protectedEmail, aliases = [] }) {
  const users = new Map();
  const userFor = value => {
    const email = normalizeEmail(value);
    if (!users.has(email)) users.set(email, {email,legacyAdminRole:null,associatesRoles:[],databaseRoles:null,roles:[],requiresReview:false,reviewed:false});
    return users.get(email);
  };
  for (const [email,role] of adminUsers) {
    const user=userFor(email);user.legacyAdminRole=role;
    if (mapped[role]) user.roles.push(mapped[role]); else user.requiresReview=true;
  }
  for (const [key,role] of Object.entries({managerEmails:'manager',consultantEmails:'consultant',careCoordinatorEmails:'careCoordinator',careEmails:'care',financeAdminEmails:'financeManager'})) {
    for (const email of associatesAccess[key] || []) {const user=userFor(email);user.associatesRoles.push(role);user.roles.push(role);}
  }
  for (const row of existing) {
    const user=userFor(row.email);
    if (Array.isArray(row.roles)) {
      user.databaseRoles=row.roles;user.roles=[...row.roles];
      if(row.roles.some(role=>!Object.hasOwn(ROLES,role)))user.requiresReview=true;
    } else if (row.role === 'revoked') {user.databaseRoles=[];user.roles=[];user.requiresReview=false;}
    else if (mapped[row.role]) {user.databaseRoles=[mapped[row.role]];user.roles=[mapped[row.role]];}
    else {user.legacyAdminRole=row.role;user.requiresReview=true;}
  }
  protectedEmail=normalizeEmail(protectedEmail || associatesAccess.superUserEmail || 'chris@planwithcare.co.uk');
  const root=userFor(protectedEmail);root.roles.push('superadmin');
  for (const alias of aliases) {
    const source=users.get(normalizeEmail(alias.aliasEmail));
    const target=userFor(alias.canonicalEmail);
    if (!source || source === target) continue;
    if (source.databaseRoles?.length === 0 || target.databaseRoles?.length === 0) {
      target.roles=[];target.requiresReview=true;
    } else target.roles.push(...source.roles);
    target.legacyAdminRole ||= source.legacyAdminRole;
    target.associatesRoles.push(...source.associatesRoles);
    target.requiresReview ||= source.requiresReview;
    users.delete(source.email);
  }
  return {version:1,sourcesVerified,protectedEmail,aliases,users:[...users.values()].sort((a,b)=>a.email.localeCompare(b.email)).map(user=>({...user,roles:Object.keys(ROLES).filter(role=>user.roles.includes(role)),reviewed:false}))};
}
function validateReviewedReport(report) {
  if(report.version!==1 || report.sourcesVerified!==true)throw new Error('Confirm both production inventories and set sourcesVerified to true before seeding.');
  if(!Array.isArray(report.users)||report.users.some(user=>user.reviewed!==true))throw new Error('Every user must be reviewed before seeding.');
  const { validateAssignment }=require('../api/_lib/shared-roles');
  const seen=new Set();
  const rows = report.users.map(user=>{
    const assignment=validateAssignment(user,report.protectedEmail);
    if(seen.has(assignment.email))throw new Error('Duplicate email in migration review.');seen.add(assignment.email);
    return assignment;
  });
  if (!rows.some(row=>row.email===normalizeEmail(report.protectedEmail) && row.roles.includes('superadmin'))) throw new Error('The protected superadmin must be included in the review.');
  const aliases = new Set();
  for (const alias of report.aliases || []) {
    const email = normalizeEmail(alias.aliasEmail), canonical = normalizeEmail(alias.canonicalEmail);
    if (!/^[^\s@,]+@[^\s@,]+\.[^\s@,]+$/.test(email) || email !== alias.aliasEmail || canonical !== alias.canonicalEmail || !seen.has(canonical) || seen.has(email) || aliases.has(email) || email === normalizeEmail(report.protectedEmail)) throw new Error('Review valid, unique aliases pointing to canonical assignments.');
    aliases.add(email);
  }
  return rows;
}
if(require.main===module){
  const [adminEnv,associatesConfig,existingFile,out,aliasesFile]=process.argv.slice(2);
  if(!out)throw new Error('Usage: node scripts/role-migration-report.cjs ADMIN_ENV ASSOCIATES_CONFIG_JSON EXISTING_ROWS_JSON OUTPUT_JSON');
  const original={...process.env};
  for(const key of Object.keys(process.env))if(key.startsWith('ACCESS_'))delete process.env[key];
  Object.assign(process.env,readEnv(adminEnv));
  const report=buildReport({adminUsers:getAuthorizedUsersMap(),associatesAccess:JSON.parse(fs.readFileSync(associatesConfig)).access,existing:JSON.parse(fs.readFileSync(existingFile)),aliases:aliasesFile?JSON.parse(fs.readFileSync(aliasesFile)):[]});
  for(const key of Object.keys(process.env))if(!(key in original))delete process.env[key];Object.assign(process.env,original);
  fs.writeFileSync(out,JSON.stringify(report,null,2)+'\n');
  const markdown=['# Shared role migration review','','Sources: local Admin environment (production confirmation required), production Associates frontend configuration, and current database export.','','Confirm the production Admin inventory and explicitly review every row before seeding. Suggested roles do not constitute approved assignments.','','| Email | Legacy Admin role | Associates roles | Suggested shared roles | Decision required |','|---|---|---|---|---|',...report.users.map(user=>`| ${user.email} | ${user.legacyAdminRole||'—'} | ${user.associatesRoles.join(', ')||'—'} | ${user.roles.join(', ')||'No access'} | ${user.requiresReview?'Choose replacement for legacy role':'Confirm assignment'} |`)];
  if(report.aliases.length)markdown.push('','Confirmed aliases:',...report.aliases.map(alias=>`- ${alias.aliasEmail} → ${alias.canonicalEmail}`));
  fs.writeFileSync(path.join(path.dirname(out),'role-migration-review.md'),markdown.join('\n')+'\n');
  console.log(`Prepared ${report.users.length} users; ${report.users.filter(user=>user.requiresReview).length} need replacement-role decisions. No assignments changed.`);
}
module.exports={buildReport,validateReviewedReport};
