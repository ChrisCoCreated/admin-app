# Shared roles rollout

Production project: `zfdogbvgahmkjtstaklu`, confirmed from both Vercel configurations and approved by the user on 7 October 2026. Both compatible versions are deployed and the reviewed assignments are migrated. The initially prepared tables in `jvloxlvgbtrgzmpqlhdi` remain inactive. The steps below are the rollout and recovery runbook.

## Review and preserve existing access

1. Export existing `app_role_assignments` rows and record both deployed Git revisions before any database changes. The pre-migration inspection of production project `zfdogbvgahmkjtstaklu` found no role tables; recheck before any future migration.
2. Obtain the current production Admin access lists securely. The generated `private .role-migration/role-migration-review.json` currently uses local Admin configuration and production Associates configuration; the production inventory and assignments were approved by the user on 7 October 2026.
3. Regenerate the report with `node scripts/role-migration-report.cjs ADMIN_ENV ASSOCIATES_CONFIG_JSON EXISTING_ROWS_JSON OUTPUT_JSON ALIASES_JSON`. Never commit environment exports or service-role keys.
4. Review every user's `roles`. Marketing and HR are shared roles, and Consultant includes Photo Layout and Enquiries. Review restricted combinations explicitly. The user confirmed that both Rebecca addresses belong to one person; .role-migration/role-aliases-review.json links the Plan With Care address to the Thrive address.
5. Mark each reviewed row `reviewed: true`, and set `sourcesVerified: true` only after both production inventories have been confirmed. Keep explicit revoked users as `roles: []`.
6. Use the same `SUPER_USER_EMAIL` in both deployments. The default protected account is `chris@planwithcare.co.uk`. No historical full-access or finance email list will grant runtime roles after activation.

## Migrate and deploy

1. Finish and verify both repositories before rollout, including making the Associates migration identical to `database/role-management.sql`. Disable any old role-management writers before changing the schema.
2. Apply the canonical migration once. It preserves legacy evidence, widens the role-array constraint and creates `app_role_management_state` with access and writes disabled. Exported legacy evidence must remain available throughout rollout.
3. With Supabase server settings configured, run `node scripts/seed-shared-roles.cjs REVIEWED_REPORT_JSON`. The script rejects incomplete reviews and changed database inventories. Do not seed while access is active.
4. Deploy both compatible app versions. They use `GET/PUT /api/roles`; Admin retains `/api/role-management` as an alias. Associates also serves authenticated `GET /api/auth/me`. Public configuration must contain no assignment directory or finance recipient list.
5. Verify schema permissions and both deployments' code versions, then set `access_enabled = true` on state row `id = 'shared'`, leaving `writes_enabled = false`.
6. Smoke-test both apps with Superadmin, Manager, Finance Manager, Care Coordinator, Consultant, Care, Marketing, HR and revoked accounts. Verify no role-management access for ordinary roles, no general management access for Finance Manager, and no access for an empty assignment.
7. Set `writes_enabled = true` only after both production checks pass. Edit a designated test account from each app and verify the next authenticated request in the other app reflects the change. Restore its reviewed roles afterward. Do not send test notification emails.

## Rollback

Pause management writes first. Preserve all assignments modified since the initial export, especially revocations. Do not restore environment-only authorization or a single-role app version against this shared schema: either would lose shared-role decisions. Roll back to a compatible build; otherwise disable shared access until a compatible fix is deployed. Keep deployment revisions and the approved migration report with the release record.
