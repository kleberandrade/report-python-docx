# Playtest web forms

A Portuguese academic evaluation site accompanying the Python DOCX generator. Built with Next.js App Router 16.4.6, React 19.3.1, Zod 4.6.7, and the Neon serverless Postgres driver 1.2.1. Requires Node.js 22. Source remains MIT, with Kleber de Oliveira Andrade's authorship retained.

## Local use

```sh
cd web
npm ci
cp .env.example .env.local
# Configure server-only values in .env.local with your own dedicated test database.
npm run db:migrate
npm run dev
```

Open `http://127.0.0.1:8793`. The public pages and complete synthetic report previews work without credentials. A submission without DATABASE_URL returns a visible error and preserves form values; it is never recorded in memory or localStorage as if it were a saved response.

Routes:

- `/`: select the evaluation stage.
- `/avaliar/alpha`: originality, fun, market, interest in playing, willingness to pay, and optional comments.
- `/avaliar/beta`: graphics, audio, HUD, gameplay, story, difficulty, age, gender, and optional strong/weak points.
- `/relatorios`: complete one-page alpha and beta report previews and DOCX downloads. All examples are fictional. Clicking a preview opens its original-resolution image.
- `/resultados`: teacher-only export form. The page is public; individual responses are only retrieved after a valid server-side export key check.

## Database and server-only configuration

Create a **dedicated** Neon test project. Do not reuse a database from another application or import the historical course workbook into a public demonstration.

- `DATABASE_URL`: the Neon connection string, kept on the server. No `NEXT_PUBLIC_` prefix.
- `ADMIN_PASSWORD`: a unique random key of at least 24 characters for CSV export. Keep it in a password manager and share it privately only with the responsible teaching staff. Export is disabled if it is absent or too short.

`npm run db:migrate` reads `.env.local` when present and applies `scripts/schema.sql` in a transaction. It creates `class_evaluations` and its stage/date index only if missing. Reexecution does not delete responses. Do not set up a broad database migration endpoint in the deployed application.

Responses are validated before parameterized SQL inserts. A UUID identifies a submission; retrying the same UUID does not add another row. The database stores stage, normalized response JSON, and creation time. No name, email, IP address, or browser fingerprint is collected by application code; hosting providers may maintain their own operational logs. Age and gender remain part of the beta schema, including `Prefiro não informar` for gender. Avoid identifiable comments.

The server accepts same-origin POST requests only and caps JSON bodies at 16 KB. Validation restricts scores, lengths, price labels, difficulty, and demographics. Runtime failures return generic messages without database details. The browser retains unsent answers in the current form after a failure; refreshing or closing the page clears them. The export key is sent in the Authorization header, never in a URL, database record, or browser storage. There are no emails, newsletters, or tracking scripts.

## Export to the Python generator

Enter the teacher key on `/resultados`, choose the stage, and download the CSV. No export deletes database records. Empty stages display a message. Exports currently stop at 10,000 responses per stage rather than silently truncating them.

Install the Python project from the repository root, then run:

```sh
class-report alpha-evaluation.csv -o reports/alpha
class-report beta-evaluation.csv --stage beta -o reports/beta
```

Exports use the exact original Portuguese column names, a UTF-8 BOM, comma separators, numeric score fields, and quoted multiline comments. Spreadsheet formula prefixes in text are neutralized with a leading apostrophe for safe opening in spreadsheet software; that apostrophe is intentionally preserved by the Python importer. No XLSX authoring dependency or document renderer is bundled into the web runtime.

The web application collects responses and exports data. DOCX generation runs in the existing Python package on the teacher's computer; there is no server-side DOCX download for newly collected evaluations in this first web version. Public example DOCX files are static assets.

## Vercel and Neon test deployment

After explicit user approval, a dedicated Neon **Free** database named `class-report-forms` was provisioned in São Paulo (`gru1`). Auth and Realtime were disabled. Resource identifiers and account console links are intentionally kept out of public documentation.

The test website is published at [class-report-forms.vercel.app](https://class-report-forms.vercel.app), with the production deployment marked Ready. The user selected the existing hosting workspace. No plan changes, upgrades, or payments were attempted.

The dedicated database is connected to Production, with connection variables stored as secrets. After the account owner completed Neon verification, the schema was applied successfully to the dedicated database. Fictional alpha and beta evaluations submitted through the public forms both displayed the saved confirmation; a direct database count confirmed one persisted response per stage. The account owner saved the teacher export password as a Production secret. After direct owner verification, a new deployment completed with Ready status and the current production domain. Authenticated cloud export is not yet verified because the teacher key was not entered into the export form; no secret was retrieved for testing. Complete alpha and beta report previews are publicly available at `/relatorios`.

Deployment procedure (the import, database connection, and initial deployment are complete; schema activation, collection checks, and the Production teacher secret are complete; authenticated export verification remains):

1. Import `kleberandrade/report-python-docx` into Vercel, set **Root Directory** to `web`, framework to Next.js, and Node.js to 22. The committed `vercel.json` pins `npm ci`, the build command, and São Paulo compute region.
2. Create or connect a dedicated Neon **free** test database through the existing integration. If free provisioning is unavailable or terms/payment changes are requested, stop and obtain the user's decision.
3. Configure `DATABASE_URL` and `ADMIN_PASSWORD` as server-only environment variables for the intended deployment environment. Do not expose values in Git, screenshots, logs, chat messages, or public client bundles.
4. Apply `scripts/schema.sql` to that database. The configured database region should be close to the function region where supported.
5. Deploy and test a fictional alpha and beta submission. Export using the teacher key and verify CSV-to-DOCX behavior against the newly saved records before claiming full integration testing.

`npm run build` requires no database access and does not run migrations. Preview and production deployments should use separate database branches/credentials; do not let previews silently write into a real course's collection.

This is a first test release: no respondent login, unique-person enforcement, CAPTCHA, global rate limiting, classroom/campaign management, automated retention policy, or per-teacher accounts are implemented. A UUID prevents accidental retry duplicates, not deliberately repeated evaluations by one person. Use a limited test audience; review abuse controls and access requirements before a larger public classroom rollout.

## Architecture and validation

- `lib/evaluations.ts`: the discriminated alpha/beta schemas and shared choice lists.
- `lib/database.ts`: parameterized Neon storage, idempotent inserts, export query, and constant-time key verification.
- `lib/http.ts`: bounded request parsing and origin comparison, including the Host header behind Next.js's internal URL normalization.
- `lib/export.ts`: safe CSV serialization with legacy headers.
- `app/api`: submission and restricted export handlers.
- `components`: client-side form state, preserving answers after errors and confirming only persisted submissions.
- `app`: server-rendered navigation, questionnaire routes, preview and teacher pages. No credential crosses into a client component.

```sh
npm test
npm run format:check
npm run build
npm run typecheck
npm start
```

Locally, 26 web tests passed with `CLASS_REPORT_PYTHON` set to the project Python executable. This includes two real JavaScript export → Python CLI → DOCX integration tests. Without that environment variable, those two tests are explicitly skipped. CI installs the Python package and sets it to `python`, so the round trips run there too.

The test suite covers payload schema boundaries, CSV content and formula neutralization, origin checks, oversized payloads, missing storage, successful/failed save outcomes, protected/empty exports, and key validation. Storage is mocked in endpoint tests; they do **not** certify a live Neon connection.

Browser QA on the actual local production build checked the landing page, alpha submission errors with answers retained, beta on a 390 × 844 viewport, beta error state, unauthorized teacher export, and the complete report preview page. A real origin mismatch was reproduced in-browser, pinned by a failing regression test, and corrected. Live browser tests also confirmed successful alpha and beta submissions and their persistence in the dedicated Neon database. The live export rejected an invalid key as expected. Authenticated live export remains unverified. No physical phone, Microsoft Word, screen reader, or Lighthouse score is claimed.

Official references: [Next.js deployment on Vercel](https://vercel.com/docs/frameworks/full-stack/nextjs), [Neon with Next.js](https://neon.com/docs/guides/nextjs).
