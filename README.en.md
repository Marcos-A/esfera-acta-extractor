[Llegiu-lo en català](README.md)

# Esfer@ Acta Extractor

Application for converting Esfer@ PDF grade reports into Excel files. The project includes:

- a web app that accepts a single PDF or a ZIP containing multiple PDFs
- an admin area at `/admin` to inspect jobs, failures, and retained debug artifacts
- the original CLI workflow for local batch processing

The recommended way to use and deploy the project is the web version with Docker.

## What it does

The web app accepts:

- a single `.pdf` file and returns a single `.xlsx`
- a `.zip` file containing multiple PDFs and returns a `.zip` containing one `.xlsx` per converted PDF

The conversion pipeline reuses the original project logic to:

- extract tables from Esfer@ PDFs
- identify students, MP codes, RA, and EM grades
- generate an Excel workbook matching the structure historically produced in `02_extracted_data`

Literal `NA` values are treated as empty cells.

## Main features

### Web app

- Public interface in Catalan
- PDF or ZIP upload
- Asynchronous conversion flow
- Progress bar with intermediate status messages
- User-friendly public error messages
- Download of the converted file when processing completes

### Administration

- Username/password-protected `/admin` area
- Dashboard for recent jobs and per-file results
- Download links for retained failed source files and `failure.log`
- Manual deletion of retained failed artifacts
- Persistent SQLite audit trail

### Operations and maintenance

- Automatic deletion of successful uploads
- Failed uploads retained only for debugging
- Failure notifications through Telegram, SMTP, or webhook
- Cleanup script based on age and/or total retained size
- Docker-first deployment behind a reverse proxy

## Functional flow

1. The user uploads a PDF or ZIP file.
2. The app stores it in a temporary working directory.
3. A background thread performs the conversion.
4. The frontend polls the status API and updates the progress bar.
5. On success:
   - the download is prepared
   - temporary files are deleted after the response is served
6. On failure:
   - the source file and any partial output are retained in `FAILURE_ROOT/<job_id>`
   - a `failure.log` file is generated
   - the job is recorded in SQLite
   - a notification is sent if configured

## Generated outputs

The primary output is an Excel workbook matching the same kind of structure the project has historically generated in `02_extracted_data`.

The legacy CLI path can also generate summary workbooks in `03_final_grade_summaries`.

## Requirements

- Docker
- Python 3.9+ only if you want to run local scripts outside the container

## Project structure

```text
.
├── app.py                                # Flask web application
├── Dockerfile                            # Docker image for the web app
├── requirements.txt                      # Python dependencies
├── templates/                            # Public and admin HTML templates
│   ├── index.html                        # Public conversion interface
│   ├── admin_login.html                  # Admin login page
│   └── admin_dashboard.html              # Admin dashboard
├── src/
│   ├── audit.py                          # SQLite audit logging
│   ├── conversion_service.py             # Shared conversion workflow
│   ├── notifier.py                       # Failure notifications
│   ├── pdf_processor.py                  # PDF table and metadata extraction
│   ├── data_processor.py                 # Data cleanup and transformation
│   ├── grade_processor.py                # Grade-specific logic
│   ├── excel_processor.py                # Main Excel generation
│   └── summary_generator.py              # Legacy CLI summary generation
├── scripts/
│   ├── cleanup_failed_uploads.py         # Automatic cleanup for retained failures
│   ├── run-local-web.sh.example          # Local startup template
│   ├── run-retention-cleanup.sh.example  # Server cleanup helper
│   ├── install-retention-cron.sh.example # Cron installer for retention cleanup
│   └── esfera2excel-retention.logrotate.example # Logrotate rule for the cleanup log
├── .env.local.example                    # Example local environment file
├── data/                                 # SQLite database and persisted runtime data
├── failed_uploads/                       # Temporarily retained failed uploads for debugging
├── 01_source_pdfs/                       # Input directory for legacy CLI mode
├── 02_extracted_data/                    # Legacy CLI Excel output
├── 03_final_grade_summaries/             # Legacy CLI summary output
├── original_pdf_files/                   # Optional manual test PDFs, gitignored
└── esfera-acta-extractor.py              # Entry point for the legacy CLI workflow
```

## Quick start with Docker

### Recommended option

1. Create the private local env file:

```bash
cp .env.local.example .env.local
```

2. Create the private startup script:

```bash
cp scripts/run-local-web.sh.example run-local-web.sh
chmod +x run-local-web.sh
```

3. Run it:

```bash
./run-local-web.sh
```

4. Open:

- `http://127.0.0.1:8000/`
- `http://127.0.0.1:8000/admin/login`

First-run behavior:

- if `TELEGRAM_BOT_TOKEN` or `TELEGRAM_CHAT_ID` still contain placeholder values, the script prompts for them and saves them into `.env.local`
- the local container created by this script is named `esfera2excel-web-local`

Both `run-local-web.sh` and `.env.local` are ignored by Git.

### Manual Docker run

You can also start the app without the helper script:

```bash
docker build -t esfera-acta-extractor-web .

docker run --rm -p 8000:8000 \
  -e SECRET_KEY=change-this-secret \
  -e ADMIN_USERNAME=marcos \
  -e ADMIN_PASSWORD=change-this-password \
  -e TELEGRAM_BOT_TOKEN=123456:ABCDEF \
  -e TELEGRAM_CHAT_ID=123456789 \
  -e AUDIT_DB_PATH=/app/data/conversion_audit.sqlite3 \
  -e FAILURE_ROOT=/app/failed_uploads \
  -v "$(pwd)/data:/app/data" \
  -v "$(pwd)/failed_uploads:/app/failed_uploads" \
  esfera-acta-extractor-web
```

## Configuration

Main settings:

| Variable | Description | Default |
|---|---|---|
| `SECRET_KEY` | Flask session key for `/admin` | `change-me-before-production` |
| `ADMIN_USERNAME` | Username for `/admin` | `admin` |
| `ADMIN_PASSWORD` | Password for `/admin` | `change-me` |
| `MAX_UPLOAD_SIZE_MB` | Maximum upload size | `50` |
| `AUDIT_DB_PATH` | SQLite database path | `./data/conversion_audit.sqlite3` |
| `FAILURE_ROOT` | Directory where failed jobs are retained | `./failed_uploads` |
| `UPLOAD_ROOT` | Temporary working directory root | `/tmp/esfera-acta-extractor` |

Notifications:

| Variable | Purpose |
|---|---|
| `TELEGRAM_BOT_TOKEN` | Telegram bot token |
| `TELEGRAM_CHAT_ID` | Destination chat ID |
| `SMTP_HOST`, `SMTP_PORT`, `SMTP_USERNAME`, `SMTP_PASSWORD` | SMTP settings |
| `SMTP_USE_TLS`, `SMTP_USE_SSL` | SMTP transport options |
| `ALERT_FROM_EMAIL`, `ALERT_TO_EMAIL` | Email sender and recipient |
| `ALERT_WEBHOOK_URL` | Generic webhook fallback |

Retention policy:

| Variable | Purpose | Recommended |
|---|---|---|
| `FAILURE_RETENTION_DAYS` | Maximum number of days to keep retained failures | `30` |
| `FAILURE_MAX_SIZE_MB` | Maximum size of `failed_uploads` | `1024` |

## How the web app behaves

### For a single PDF

- the file is received
- it is converted into one Excel workbook
- the user downloads the `.xlsx`
- the uploaded file and working directory are then deleted

### For a ZIP file

- only valid PDFs are extracted
- each PDF is converted to Excel
- a final ZIP is generated with all `.xlsx` files
- the original ZIP and the temporary working directory are then deleted

### On failure

- the public UI shows a generic user-friendly message
- the technical details remain available in `/admin`
- a debug directory is retained containing:
  - the failed source file
  - any partial output
  - `failure.log`

## Admin area

The admin route is `/admin`.

It includes:

- recent job summary
- per-file conversion records
- timestamps for creation and return
- error details
- download links for retained failed source files
- download links for `failure.log`
- manual deletion of retained artifacts

Notes:

- the admin UI is in English
- stored timestamps are UTC

## Failure notifications

Telegram is the recommended option.

Minimum setup flow:

1. create a bot with `@BotFather`
2. get the bot token
3. send a message to the bot
4. retrieve the `chat_id`
5. set `TELEGRAM_BOT_TOKEN` and `TELEGRAM_CHAT_ID`

When a conversion fails, the app can notify with:

- `job_id`
- source filename
- request type
- main error
- debug path

## Retained failure cleanup

### Manual cleanup

From `/admin`, you can delete retained artifacts for a failed job once you no longer need them.

### Cleanup script

Run:

```bash
python3 scripts/cleanup_failed_uploads.py --retention-days 30 --max-size-mb 1024
```

Dry run:

```bash
python3 scripts/cleanup_failed_uploads.py --dry-run
```

### Run cleanup inside the production container

```bash
docker exec esfera2excel-web python scripts/cleanup_failed_uploads.py --retention-days 30 --max-size-mb 1024
```

### Helper and cron

Server helper:

```bash
cp scripts/run-retention-cleanup.sh.example /usr/local/bin/esfera2excel-retention-cleanup
chmod +x /usr/local/bin/esfera2excel-retention-cleanup
```

Cron installer using the same `.env.local` file:

```bash
cp scripts/install-retention-cron.sh.example /usr/local/bin/esfera2excel-install-retention-cron
chmod +x /usr/local/bin/esfera2excel-install-retention-cron
ENV_FILE=/path/to/.env.local CONTAINER_NAME=esfera2excel-web /usr/local/bin/esfera2excel-install-retention-cron
```

Example daily cron entry:

```cron
30 3 * * * CONTAINER_NAME=esfera2excel-web FAILURE_RETENTION_DAYS=30 FAILURE_MAX_SIZE_MB=1024 /usr/local/bin/esfera2excel-retention-cleanup >> /var/log/esfera2excel-retention.log 2>&1
```

Operational recommendation:

- use `FAILURE_RETENTION_DAYS` as the main rule
- use `FAILURE_MAX_SIZE_MB` as a safety limit
- keep manual deletion for cases already reviewed

To prevent the cron log itself from growing without bounds, you can also install a `logrotate` rule:

```bash
cp scripts/esfera2excel-retention.logrotate.example /etc/logrotate.d/esfera2excel-retention
```

## Recommended deployment

### Goal

Deploy the web app at `https://esfera2excel.marcos-a.com` behind a reverse proxy, with the container listening on `127.0.0.1:8000`.

### Persistent storage

Mount persistent volumes for:

- `/app/data`
- `/app/failed_uploads`

Also set:

- `AUDIT_DB_PATH=/app/data/conversion_audit.sqlite3`
- `FAILURE_ROOT=/app/failed_uploads`

### Reverse proxy

Minimal Nginx example:

```nginx
server {
    server_name esfera2excel.marcos-a.com;

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

The proxy should:

- serve `https://esfera2excel.marcos-a.com/`
- also serve `https://esfera2excel.marcos-a.com/admin`
- terminate TLS
- forward `X-Forwarded-For`

### Example production container run

```bash
docker run -d \
  --name esfera2excel-web \
  -p 127.0.0.1:8000:8000 \
  --log-opt max-size=10m \
  --log-opt max-file=5 \
  -e SECRET_KEY=change-this-secret \
  -e ADMIN_USERNAME=admin \
  -e ADMIN_PASSWORD=change-this-password \
  -e TELEGRAM_BOT_TOKEN=... \
  -e TELEGRAM_CHAT_ID=... \
  -e FAILURE_RETENTION_DAYS=30 \
  -e FAILURE_MAX_SIZE_MB=1024 \
  -e AUDIT_DB_PATH=/app/data/conversion_audit.sqlite3 \
  -e FAILURE_ROOT=/app/failed_uploads \
  -v /persistent/path/data:/app/data \
  -v /persistent/path/failed_uploads:/app/failed_uploads \
  esfera2excel-web
```

Good practice:

- do not bake secrets into the image
- do not expose the app directly to the public internet if a reverse proxy is available
- always change `SECRET_KEY` and `ADMIN_PASSWORD`
- limit Docker logs with `--log-opt max-size` and `--log-opt max-file`

## Useful HTTP routes

| Route | Purpose |
|---|---|
| `/` | Public upload interface |
| `/health` | Basic health check |
| `/convert` | Start a conversion job |
| `/convert/status/<job_id>` | Poll status and progress |
| `/convert/download/<job_id>` | Download the conversion result |
| `/admin/login` | Admin login |
| `/admin` | Admin dashboard |

## Troubleshooting

### The container runs but `/admin` login fails

Check:

- `ADMIN_USERNAME`
- `ADMIN_PASSWORD`
- `SECRET_KEY`

### Telegram alerts are not arriving

Check:

- the bot exists
- you have sent a message to the bot
- `TELEGRAM_BOT_TOKEN` is correct
- `TELEGRAM_CHAT_ID` is correct

### Failed files never disappear from the server

Check:

- whether scheduled cleanup is configured
- whether `FAILURE_RETENTION_DAYS` and `FAILURE_MAX_SIZE_MB` are set
- whether cron is actually running

## Legacy CLI workflow

The original CLI mode is still available for local batch processing.

Inputs and outputs:

- input: `01_source_pdfs`
- main output: `02_extracted_data`
- summaries: `03_final_grade_summaries`

Run with Docker:

```bash
docker build -t esfera-acta-extractor .
docker run --rm -v "$(pwd):/app" esfera-acta-extractor python esfera-acta-extractor.py
```

The CLI path is useful for local technical workflows and testing, but it is not the recommended deployment model.

## Development

Main libraries:

- `Flask`
- `gunicorn`
- `pandas`
- `pdfplumber`
- `openpyxl`
- `rich`

## Contributing

1. Fork the repository.
2. Create a working branch.
3. Make your changes.
4. Open a pull request.

## License

This project is distributed under GNU GPL-3.0. See [`LICENSE`](LICENSE).
