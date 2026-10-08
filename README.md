# Academic Marks Management System — V32.4

## V32.4 updates
- Individual student analysis now shows an automatic result-based performance observation.
- Teachers/admins can optionally save a custom remark against a student from the analysis page.
- Admin Settings has a fully themed dark UI with readable labels, inputs and lists.
- Admin can delete individual semester/test types and their linked marks.
- Admin can delete all test/assessment types and their linked marks in one action.
- Admin/teacher can delete an individual subject and its linked assignments/marks.
- Admin/teacher can delete all subjects, assignments and related marks while preserving students and attendance.
- Result/analysis class subject charts use a stable horizontal layout with adaptive height and readable labels to prevent overlapping/glitching labels.
- Existing large-file import, chunked-save, reset-request controls, and themed import guidance are retained.

## Run
1. Create `.env` from `.env.example`.
2. Install `requirements.txt`.
3. Ensure MySQL is running and the configured database is available.
4. Run `python app.py`.

## Security and deployment

1. Copy `.env.example` to `.env`.
2. Set a new random `SECRET_KEY` and private `ADMIN_RECOVERY_CODE`.
3. Set the database credentials in `.env`.
4. Keep `.env` out of Git.
5. Set `FLASK_DEBUG=0` in production.
6. Serve behind HTTPS and set `SESSION_COOKIE_SECURE=1`.
7. Run `pip install -r requirements.txt`.
8. Use a production WSGI server (for example, Gunicorn) instead of Flask's development server.

The repository intentionally contains no real credentials.

## Large academic imports

The application is configured for large college datasets. The default request upload limit is **512 MB** and can be changed with `MAX_CONTENT_LENGTH`.

There is no artificial maximum-student setting. Practical capacity depends on the database and server resources.

For very large datasets, use batch/chunk processing and transactional database writes. Do not treat one HTTP request as an unlimited-memory operation.

Recommended planning target:
- 10,000+ students: suitable for a properly sized deployment.
- Hundreds of thousands of marks rows: process in batches/transactions.
- Extremely large files: split imports or increase server resources.

Do not set an unlimited upload size in production.

## Security note

All state-changing HTML forms and AJAX requests use Flask-WTF CSRF protection. Login forms include the token explicitly, and JSON import requests send the token in the `X-CSRFToken` header.

## Large upload deployment

The application default request limit is **512 MB** and is configurable with `MAX_CONTENT_LENGTH`.

Long imports can exceed the timeout of the web server or reverse proxy even when Flask accepts the file. For production, configure the WSGI server/proxy timeout to at least the expected import duration. The application exposes `IMPORT_REQUEST_TIMEOUT` as a deployment setting for operators; the actual timeout must also be configured in the WSGI server/proxy.

If a college imports very large datasets, prefer CSV/XLSX files of reasonable size and split exceptionally large imports when the hosting provider has strict request-time limits.


### Password reset request behavior
- Each teacher/student account can have at most one pending password-reset request.
- Administrators can reset or delete/dismiss pending requests.
- Deleting a pending request allows that account to submit a new request.


## V32.4 performance fixes
- Database connection pooling reduces connection overhead on every page/query.
- Request-scoped user/settings/class caching removes repeated queries during page rendering.
- Large import preview renders only 60 rows at a time instead of thousands of DOM inputs.
- Import edits/deletions are stored as small change sets and finalized from the original server-side import file.
- Marks and attendance saves use batched catalog/student/relationship operations and database upserts.
- Stale temporary import files are cleaned automatically.
- Login/setup pages are no-cache to prevent stale CSRF tokens.
- Import guidance and login borders are explicitly dark-themed to prevent legacy white-surface CSS conflicts.
