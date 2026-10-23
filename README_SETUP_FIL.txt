UIMS — INTEGRATED MVP PACKAGE
Google Apps Script + Google Sheets

IMPORTANT
This package brings the planned phases into one integrated MVP, but it is NOT a certified production system and must be tested on dummy data first. Apps Script + Sheets has quota, concurrency, and security limitations. For sensitive/large business deployments, use a proper database/backend and professional security review.

FILES
- Code.gs: server functions, sheet setup, username/password login, server sessions, tenant checks, roles/permissions, inventory, FIFO layers, transfers, stock counts, reports and audit logging.
- Index.html: app layout and screens.
- Styles.html: responsive UI.
- JavaScript.html: client behavior.
- appsscript.json: runtime configuration.

INSTALL
1. Go to https://script.google.com/ and create a NEW Apps Script project for UIMS. Keep your Phase 1 project as backup; do not overwrite it before testing this package.
2. Replace Code.gs content with this package's Code.gs.
3. Add HTML files named exactly Index, Styles, JavaScript and paste their respective contents. The editor file names omit .html.
4. Project Settings > enable “Show appsscript.json manifest file in editor”, then replace manifest with this package's appsscript.json.
5. Save all files.
6. In the function dropdown, choose setupUIMSDatabase and click Run. Approve Google authorization.
7. Open Execution log. Copy the UIMS database URL for the owner and the ONE-TIME INITIAL ADMIN SETUP TOKEN. Never share the token publicly. The token is deleted after successful first-admin setup. Kung nawala ang token sa log, patakbuhin ang showSetupToken (hangga't wala pang admin).
8. Deploy > New deployment > Web app. Execute as: Me (the script owner). Who has access: 'Anyone' kung gagamitin ng staff na walang access sa Google account mo (may sariling username/password login ang app, at iyon ang gate). Kung ikaw lang ang gagamit, piliin ang 'Only myself'. Huwag i-share nang malawak ang URL.
9. Open the deployed web app, enter the one-time token and create the business + initial Business Admin. The admin password must be 10+ characters and include letters and numbers.
10. Sign in with the admin username/password. The initial admin must change the temporary password on first login (the first password set in setup is treated as temporary).
11. In Users & access, create staff accounts and grant only the permissions they need. Give temporary credentials privately.
12. Create Units and Locations first, then Categories and Items. For each item, Base Unit ID must be the ID shown in the Units table. Para sa ibang unit (hal. 1 Box = 12 Piece): Items page > Unit conversions. Ilagay kung ilang BASE unit ang laman ng 1 alternate unit.

SECURITY AND FEATURE NOTES
- Custom username/password login; no Google OAuth Client ID. Passwords are salted SHA-256 hashes, not plain text. SHA-256 is not a purpose-built password hashing algorithm; use a proper auth provider/strong password KDF for production.
- Server methods validate a cached session and BusinessID for data access. Session uses Apps Script CacheService and expires; users must sign in again after expiry. CacheService is not a durable session store.
- No public sign-up. Initial setup uses a one-time token printed to the execution log.
- Admin-created users are forced to change password on first login.
- User permissions are server checked for major functions; admin role bypasses permissions. Review permissions and all access paths before production.
- The app permits negative on-hand. Negative issues are not automatically inventory loss. If FIFO layers are exhausted, the negative portion has no cost layer, so valuation can be understated until reconciled. Review negative balances and FIFO results.
- FIFO is a basic implementation. Transfers (request, dispatch, receive) at stock counts (count entry, post, cancel) ay may UI na; MVP pa rin ang mga workflow.
- Expiry/batch columns are supported, but expiry alerts and full batch-level allocation are not complete.
- CSV export is implemented; PDF/print report layouts, offline queue synchronization, full approval UI, reversal/correction workflows, detailed item-unit conversion UI, and complete audit/report filters need a tested follow-up pass.
- Posted transactions should not be edited directly in Sheets. This MVP does not yet provide a complete reversal transaction UI; keep backups and use test data.
- Never store actual sensitive business records until all features and security have been tested. Make a copy of the database before upgrades.

BASIC TEST PLAN (DUMMY DATA ONLY)
1. Setup first admin; verify token cannot be reused.
2. Create a second user with limited permissions. Verify restricted server calls fail.
3. Create 2 locations, units, category and an item.
4. Receive 10 base units at cost 20; verify on-hand 10.
5. Issue 12; verify on-hand -2 and investigate FIFO valuation shortage.
6. Receive 5; verify on-hand 3.
7. Create a second business only after multi-business registration is implemented and tested; current bootstrap creates one business. Do not claim multi-business onboarding is complete until this is implemented.
8. Test logout, cache expiry, disabled account, password change, wrong credentials and concurrent transactions.

WHAT'S CORRECTED IN v1.1
- StockBalances ay binabasa na nang tama (dati laging 0 ang inventory/dashboard dahil maling column ang pinagbabasehan).
- FIFO valuation: tamang item/location columns na ang ginagamit.
- Auth screen ay nawawala na pagkatapos mag-login ([hidden] CSS fix).
- Date values ay hindi na ibinabalik bilang Date object (nagdudulot ng null sa google.script.run); plain-text ang SKU/barcode/petsa/hash columns.
- Password hash ay hex na (iwas problema kapag nagsimula sa + o = ang value sa Sheets).
- Walang partial posting: bina-validate muna lahat ng lines bago magsulat.
- FIFO cost ay per BASE unit (tama kahit may unit conversion); dala na ng transfers ang cost; ang approved na adjustment (+/-) ay tama na ang direksyon.
- allow_negative_stock setting ay ipinapatupad na (Settings sheet, BusinessID ng business: TRUE/FALSE).
- Login lockout: 5 maling subok = 10 minutong lock. Owner-only na ang setup functions. Hindi puwedeng gawing admin / i-disable ang admin / palitan ang sariling permissions ng non-admin na may manage_users.
- Bagong UI: multi-line transactions + SKU scan, approvals, unit conversions, transfers list (dispatch/receive), stock count entry/post/cancel, edit permissions, reset password, inventory location filter, report date filter.
- Walang audit log viewer pa (sa AuditLog sheet makikita), at wala pang Settings screen: i-edit ang Settings sheet para sa allow_negative_stock.

