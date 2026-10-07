const test=require('node:test');
const assert=require('node:assert/strict');
const {buildReport,validateReviewedReport}=require('../scripts/role-migration-report.cjs');
test('migration preserves finance, multiple roles and revocations and blocks unreviewed rollout',()=>{
  const report=buildReport({adminUsers:new Map([['finance@example.com','finance'],['restricted@example.com','time_hr_clients'],['revoked@example.com','admin']]),associatesAccess:{managerEmails:['finance@example.com'],financeAdminEmails:['finance@example.com'],careEmails:['revoked@example.com']},existing:[{email:'revoked@example.com',roles:[]}]});
  assert.deepEqual(report.users.find(u=>u.email==='finance@example.com').roles,['manager','financeManager']);
  assert.deepEqual(report.users.find(u=>u.email==='revoked@example.com').roles,[]);
  assert.equal(report.users.find(u=>u.email==='restricted@example.com').requiresReview,true);
  assert.throws(()=>validateReviewedReport(report),/production inventories/);
  report.sourcesVerified=true;
  assert.throws(()=>validateReviewedReport(report),/Every user/);
  for(const user of report.users)user.reviewed=true;
  assert.ok(validateReviewedReport(report).length);
});
