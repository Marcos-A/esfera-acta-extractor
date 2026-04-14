# Beta Deployment For `perf/excel-pipeline-optimization`

This branch can be run in parallel with production as a temporary beta instance at:

- `https://esfera2excel-perf.marcos-a.com`

The beta instance is designed to stay isolated from production by using:

- a different image tag
- a different container name
- a different host port
- separate host storage directories
- a separate SQLite audit database
- a separate upload temp root
- separate admin credentials
- performance instrumentation enabled only in beta

## Proposed Beta Instance Settings

Use these names and paths unless your host already reserves them:

| Item | Production | Beta |
|---|---|---|
| Image tag | `esfera2excel-web` | `esfera2excel-web:perf-excel-pipeline-optimization` |
| Container name | `esfera2excel-web` | `esfera2excel-web-perf` |
| Host bind | `127.0.0.1:8000->8000` | `127.0.0.1:8001->8000` |
| Subdomain | `esfera2excel.marcos-a.com` | `esfera2excel-perf.marcos-a.com` |
| Host data dir | production-specific | `/srv/data/esfera2excel-perf/data` |
| Host failed upload dir | production-specific | `/srv/data/esfera2excel-perf/failed_uploads` |
| Host temp/upload dir | container `/tmp` default | `/srv/data/esfera2excel-perf/upload_tmp` |
| Container audit DB | `/app/data/conversion_audit.sqlite3` | `/app/data/conversion_audit.sqlite3` |
| Container failure root | `/app/failed_uploads` | `/app/failed_uploads` |
| Container upload root | `/tmp/esfera-acta-extractor` default | `/app/upload_tmp` |
| Perf timing | disabled | `PERF_TIMING_ENABLED=true` |

## Files In This Folder

- `.env.perf.example`
- `nginx.esfera2excel-perf.conf.example`
- `Caddyfile.esfera2excel-perf.example`

No Docker Compose file is included because the current project deployment is documented as a direct `docker run` flow rather than Compose.

## Deploy Commands

Run these commands from the checked-out branch:

```bash
cd /srv/apps/esfera2excel-web
git checkout perf/excel-pipeline-optimization

sudo mkdir -p /srv/data/esfera2excel-perf/data
sudo mkdir -p /srv/data/esfera2excel-perf/failed_uploads
sudo mkdir -p /srv/data/esfera2excel-perf/upload_tmp
sudo chown -R "$USER":"$USER" /srv/data/esfera2excel-perf

cp deploy/perf/.env.perf.example .env.perf
$EDITOR .env.perf

docker build -t esfera2excel-web:perf-excel-pipeline-optimization .

docker rm -f esfera2excel-web-perf 2>/dev/null || true

docker run -d \
  --name esfera2excel-web-perf \
  --env-file .env.perf \
  -p 127.0.0.1:8001:8000 \
  --log-opt max-size=10m \
  --log-opt max-file=5 \
  -v /srv/data/esfera2excel-perf/data:/app/data \
  -v /srv/data/esfera2excel-perf/failed_uploads:/app/failed_uploads \
  -v /srv/data/esfera2excel-perf/upload_tmp:/app/upload_tmp \
  esfera2excel-web:perf-excel-pipeline-optimization
```

## Reverse Proxy

Choose one reverse proxy setup and point it at `127.0.0.1:8001`.

### Nginx

Install the example file:

```bash
sudo cp deploy/perf/nginx.esfera2excel-perf.conf.example /etc/nginx/sites-available/esfera2excel-perf.marcos-a.com.conf
sudo ln -s /etc/nginx/sites-available/esfera2excel-perf.marcos-a.com.conf /etc/nginx/sites-enabled/esfera2excel-perf.marcos-a.com.conf
sudo nginx -t
sudo systemctl reload nginx
```

### Caddy

Append the example block from `deploy/perf/Caddyfile.esfera2excel-perf.example` to your active Caddy config, then reload:

```bash
sudo caddy validate --config /etc/caddy/Caddyfile
sudo systemctl reload caddy
```

Make sure DNS for `esfera2excel-perf.marcos-a.com` points to the same host as production.

## Verify The Beta Instance

Check the container and local health endpoint:

```bash
docker ps --filter name=esfera2excel-web-perf
docker logs --tail 100 esfera2excel-web-perf
curl -fsS http://127.0.0.1:8001/health
```

Check the public beta subdomain:

```bash
curl -I https://esfera2excel-perf.marcos-a.com/health
curl -I https://esfera2excel-perf.marcos-a.com/
```

Optional side-by-side comparison:

```bash
curl -I https://esfera2excel.marcos-a.com/health
curl -I https://esfera2excel-perf.marcos-a.com/health
```

## Test Checklist

- Confirm `/` loads on both production and beta.
- Confirm `/admin/login` works with the beta-specific credentials.
- Upload a representative PDF to beta and confirm download success.
- Upload a multi-file batch to beta and confirm ZIP download success.
- Confirm beta logs contain `[PERF]` timing lines and production logs do not.
- Confirm beta writes only under `/srv/data/esfera2excel-perf`.

## Removal / Rollback

To remove the beta instance cleanly without touching production:

```bash
docker rm -f esfera2excel-web-perf
docker image rm esfera2excel-web:perf-excel-pipeline-optimization
rm -f /srv/apps/esfera2excel-web/.env.perf
```

Then remove the reverse proxy entry:

### Nginx removal

```bash
sudo rm -f /etc/nginx/sites-enabled/esfera2excel-perf.marcos-a.com.conf
sudo rm -f /etc/nginx/sites-available/esfera2excel-perf.marcos-a.com.conf
sudo nginx -t
sudo systemctl reload nginx
```

### Caddy removal

Remove the `esfera2excel-perf.marcos-a.com` block from the active Caddy config, then reload:

```bash
sudo caddy validate --config /etc/caddy/Caddyfile
sudo systemctl reload caddy
```

If you also want to delete beta data:

```bash
sudo rm -rf /srv/data/esfera2excel-perf
```

## Risks And Caveats

- Do not point the beta instance at production `data`, `failed_uploads`, or temp paths.
- Keep beta notifications disabled unless you explicitly want test failures to alert a real channel.
- The beta audit database is still SQLite, so it needs its own mounted data directory.
- `UPLOAD_ROOT` should remain writable by the container user; mounting `/srv/data/esfera2excel-perf/upload_tmp` avoids mixing beta temp files with production.
- The proxy snippet assumes TLS is already handled by your existing Nginx or Caddy setup.
- If production cleanup cron jobs target `esfera2excel-web`, do not reuse them for `esfera2excel-web-perf` unless you intentionally create a separate beta cleanup job.
