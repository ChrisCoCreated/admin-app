// Run only after the reviewed report matches both production inventories.
const fs = require('node:fs');
const { validateReviewedReport } = require('./role-migration-report.cjs');
const { request, loadSnapshot } = require('../api/_lib/shared-role-store');
async function seed(report) {
  const rows = validateReviewedReport(report);
  const snapshot = await loadSnapshot({ requireActive:false });
  if (snapshot.state.access_enabled || snapshot.state.writes_enabled) throw new Error('Seeding requires access and management writes to remain disabled.');
  // Existing database-only assignments must be reviewed too, never silently omitted.
  const reviewed = new Set(rows.map(row=>row.email));
  if (Object.keys(snapshot.assignments).some(email=>!reviewed.has(email))) throw new Error('The database changed after review. Regenerate the report before seeding.');
  for (const user of report.users) {
    if (user.databaseRoles !== null && user.databaseRoles !== undefined && JSON.stringify(snapshot.assignments[user.email]) !== JSON.stringify(user.databaseRoles)) throw new Error('An assignment changed after review. Preserve the new decision and regenerate the report.');
  }
  await request('app_role_assignments',{on_conflict:'email'},{
    method:'POST',headers:{'Content-Type':'application/json',Prefer:'resolution=merge-duplicates,return=minimal'},
    body:JSON.stringify(rows.map(row=>({...row,updated_by:'reviewed-shared-role-migration',updated_at:new Date().toISOString()}))),
  });
  if (report.aliases?.length) {
    await request('app_role_aliases',{on_conflict:'alias_email'},{method:'POST',headers:{'Content-Type':'application/json',Prefer:'resolution=merge-duplicates,return=minimal'},body:JSON.stringify(report.aliases.map(alias=>({alias_email:alias.aliasEmail,canonical_email:alias.canonicalEmail})))});
  }
  console.log(`Seeded ${rows.length} reviewed assignments. Access and editing remain disabled.`);
}
if(require.main===module)seed(JSON.parse(fs.readFileSync(process.argv[2],'utf8'))).catch(error=>{console.error(error.message);process.exitCode=1;});
module.exports={seed};
