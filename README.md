[Read this in English](README.en.md)

# Esfer@ Acta Extractor

Aplicació per convertir actes d'Esfer@ en PDF a fitxers Excel. El projecte inclou:

- una aplicació web per pujar un PDF o un ZIP amb diversos PDFs
- una àrea d'administració a `/admin` per consultar conversions, errors i artefactes conservats
- una via CLI antiga per processar lots locals des del repositori

La forma recomanada d'ús i desplegament és la versió web amb Docker.

## Què fa

La web accepta:

- un únic fitxer `.pdf`, i retorna un únic `.xlsx`
- un fitxer `.zip` amb diversos PDFs, i retorna un `.zip` amb un `.xlsx` per cada PDF convertit

La conversió reutilitza la lògica històrica del projecte per:

- extreure taules del PDF d'Esfer@
- identificar alumnes, codis MP, RA i EM
- generar un Excel amb l'estructura del directori `02_extracted_data`

Els valors literals `NA` es tracten com a cel·les buides.

## Característiques principals

### Aplicació web

- Interfície pública en català
- Pujada de PDF o ZIP
- Seguiment asíncron del progrés de conversió
- Missatges d'estat intermedis, com ara descompressió, conversió i compressió final
- Missatges d'error públics pensats per a usuaris finals
- Descàrrega automàtica del fitxer convertit quan el procés acaba

### Administració

- Accés protegit amb usuari i contrasenya a `/admin`
- Tauler amb històric recent de treballs i fitxers convertits
- Descàrrega del fitxer font fallit i del `failure.log`
- Eliminació manual dels artefactes conservats d'un error
- Registre persistent en SQLite

### Operació i manteniment

- Eliminació automàtica dels fitxers pujats quan la conversió té èxit
- Conservació dels errors només per a depuració
- Notificacions d'error via Telegram, SMTP o webhook
- Script de neteja automàtica per antiguitat i/o mida total
- Desplegament pensat per a contenidor Docker darrere d'un proxy invers

## Arquitectura funcional

1. L'usuari puja un PDF o un ZIP.
2. La web desa temporalment el fitxer en un directori de treball.
3. Un fil en segon pla processa la conversió.
4. La UI consulta l'estat via API i mostra una barra de progrés.
5. En èxit:
   - es prepara la descàrrega
   - el fitxer temporal es neteja després de servir-lo
6. En error:
   - es conserva una còpia del fitxer font i qualsevol sortida parcial a `FAILURE_ROOT/<job_id>`
   - es genera un `failure.log`
   - es registra el cas a SQLite
   - s'envia una notificació al propietari si està configurada

## Sortides generades

La sortida principal és un fitxer Excel amb el mateix tipus d'estructura que el projecte ja generava al directori `02_extracted_data`.

Addicionalment, el codi històric encara pot generar resums de qualificacions al directori `03_final_grade_summaries` quan s'executa per CLI.

## Requisits

- Docker
- Python 3.9+ només si vols executar scripts locals fora del contenidor

## Estructura del projecte

```text
.
├── app.py                                # Aplicació web Flask
├── Dockerfile                            # Imatge Docker de la web
├── requirements.txt                      # Dependències Python
├── templates/                            # Plantilles HTML de la web i /admin
│   ├── index.html                        # Interfície pública de conversió
│   ├── admin_login.html                  # Pantalla de login d'administració
│   └── admin_dashboard.html              # Tauler d'administració
├── src/
│   ├── audit.py                          # Registre SQLite de treballs i fitxers
│   ├── conversion_service.py             # Flux reutilitzable de conversió
│   ├── notifier.py                       # Notificacions d'error
│   ├── pdf_processor.py                  # Extracció de taules i metadades del PDF
│   ├── data_processor.py                 # Neteja i transformació de dades
│   ├── grade_processor.py                # Lògica de qualificacions
│   ├── excel_processor.py                # Generació del fitxer Excel principal
│   └── summary_generator.py              # Generació de resums per CLI
├── scripts/
│   ├── cleanup_failed_uploads.py         # Neteja automàtica de fallades conservades
│   ├── run-local-web.sh.example          # Plantilla d'arrencada local
│   ├── run-retention-cleanup.sh.example  # Helper de neteja per a servidor
│   └── install-retention-cron.sh.example # Instal·lador de cron per a neteja
├── .env.local.example                    # Variables d'entorn locals de mostra
├── data/                                 # Base SQLite i dades persistides en local o al servidor
├── failed_uploads/                       # Errors conservats temporalment per a depuració
├── 01_source_pdfs/                       # Entrada per a la via CLI antiga
├── 02_extracted_data/                    # Sortida Excel de la via CLI antiga
├── 03_final_grade_summaries/             # Resums de la via CLI antiga
├── original_pdf_files/                   # Carpeta opcional per a proves manuals, ignorada per Git
└── esfera-acta-extractor.py              # Punt d'entrada de la via CLI antiga
```

## Inici ràpid local amb Docker

### Opció recomanada

1. Crea el fitxer local privat:

```bash
cp .env.local.example .env.local
```

2. Crea el script local privat:

```bash
cp scripts/run-local-web.sh.example run-local-web.sh
chmod +x run-local-web.sh
```

3. Executa'l:

```bash
./run-local-web.sh
```

4. Obre:

- `http://127.0.0.1:8000/`
- `http://127.0.0.1:8000/admin/login`

Comportament del primer llançament:

- si `TELEGRAM_BOT_TOKEN` o `TELEGRAM_CHAT_ID` encara tenen el valor de plantilla, l'script te'ls demanarà i els desarà a `.env.local`
- el contenidor local creat per aquest script s'anomena `esfera2excel-web-local`

Els fitxers `run-local-web.sh` i `.env.local` estan ignorats per Git.

### Execució Docker manual

També pots arrencar la web sense el helper:

```bash
docker build -t esfera-acta-extractor-web .

docker run --rm -p 8000:8000 \
  -e SECRET_KEY=canvia-aquesta-clau \
  -e ADMIN_USERNAME=marcos \
  -e ADMIN_PASSWORD=canvia-aquesta-contrasenya \
  -e TELEGRAM_BOT_TOKEN=123456:ABCDEF \
  -e TELEGRAM_CHAT_ID=123456789 \
  -e AUDIT_DB_PATH=/app/data/conversion_audit.sqlite3 \
  -e FAILURE_ROOT=/app/failed_uploads \
  -v "$(pwd)/data:/app/data" \
  -v "$(pwd)/failed_uploads:/app/failed_uploads" \
  esfera-acta-extractor-web
```

## Configuració

Variables principals:

| Variable | Descripció | Valor per defecte |
|---|---|---|
| `SECRET_KEY` | Clau de sessió de Flask per a `/admin` | `change-me-before-production` |
| `ADMIN_USERNAME` | Usuari de l'àrea `/admin` | `admin` |
| `ADMIN_PASSWORD` | Contrasenya de l'àrea `/admin` | `change-me` |
| `MAX_UPLOAD_SIZE_MB` | Mida màxima de pujada | `50` |
| `AUDIT_DB_PATH` | Ruta de la base SQLite | `./data/conversion_audit.sqlite3` |
| `FAILURE_ROOT` | Directori on es conserven errors | `./failed_uploads` |
| `UPLOAD_ROOT` | Directori temporal de treball | `/tmp/esfera-acta-extractor` |

Notificacions:

| Variable | Ús |
|---|---|
| `TELEGRAM_BOT_TOKEN` | Token del bot de Telegram |
| `TELEGRAM_CHAT_ID` | Chat ID receptor |
| `SMTP_HOST`, `SMTP_PORT`, `SMTP_USERNAME`, `SMTP_PASSWORD` | Configuració SMTP |
| `SMTP_USE_TLS`, `SMTP_USE_SSL` | Transport SMTP |
| `ALERT_FROM_EMAIL`, `ALERT_TO_EMAIL` | Remitent i destinatari de correu |
| `ALERT_WEBHOOK_URL` | Webhook genèric alternatiu |

Política de retenció:

| Variable | Ús | Valor recomanat |
|---|---|---|
| `FAILURE_RETENTION_DAYS` | Dies màxims de conservació | `30` |
| `FAILURE_MAX_SIZE_MB` | Límit de mida total de `failed_uploads` | `1024` |

## Com funciona la web

### Per a un PDF

- es rep el fitxer
- es converteix a un únic Excel
- l'usuari descarrega el `.xlsx`
- el fitxer pujat i el directori temporal s'eliminen després

### Per a un ZIP

- es descomprimeixen només els PDFs vàlids
- es converteix cada PDF a Excel
- es genera un ZIP final amb tots els `.xlsx`
- el ZIP original i el directori temporal s'eliminen després

### En cas d'error

- la interfície pública mostra un missatge genèric i amigable
- l'error tècnic queda registrat a l'àrea `/admin`
- es conserva un directori de depuració amb:
  - el fitxer font fallit
  - qualsevol sortida parcial
  - `failure.log`

## Àrea d'administració

La ruta és `/admin`.

Inclou:

- resum de treballs i fitxers
- estat de les conversions
- hora de creació i retorn
- errors registrats
- descàrrega del fitxer font conservat
- descàrrega del `failure.log`
- eliminació manual dels artefactes d'un error

Observacions:

- la interfície d'administració està en anglès
- els timestamps registrats es desen en UTC

## Notificacions d'error

L'opció recomanada és Telegram.

Flux mínim:

1. crea un bot amb `@BotFather`
2. obtén el token
3. envia un missatge al bot
4. obtén el `chat_id`
5. defineix `TELEGRAM_BOT_TOKEN` i `TELEGRAM_CHAT_ID`

Quan falla una conversió, l'aplicació pot enviar:

- `job_id`
- nom del fitxer
- tipus de petició
- error principal
- ruta del directori de depuració

## Conservació i neteja de fallades

### Esborrat manual

Des de `/admin` pots esborrar els artefactes conservats d'un error una vegada revisat.

### Script de neteja

Pots executar:

```bash
python3 scripts/cleanup_failed_uploads.py --retention-days 30 --max-size-mb 1024
```

Mode de prova:

```bash
python3 scripts/cleanup_failed_uploads.py --dry-run
```

### Execució dins del contenidor en producció

```bash
docker exec esfera-acta-extractor-web python scripts/cleanup_failed_uploads.py --retention-days 30 --max-size-mb 1024
```

### Helper i cron

Helper per a servidor:

```bash
cp scripts/run-retention-cleanup.sh.example /usr/local/bin/esfera2excel-retention-cleanup
chmod +x /usr/local/bin/esfera2excel-retention-cleanup
```

Instal·lador de `cron` a partir del mateix `.env.local`:

```bash
cp scripts/install-retention-cron.sh.example /usr/local/bin/esfera2excel-install-retention-cron
chmod +x /usr/local/bin/esfera2excel-install-retention-cron
ENV_FILE=/ruta/al/.env.local CONTAINER_NAME=esfera-acta-extractor-web /usr/local/bin/esfera2excel-install-retention-cron
```

Exemple de cron diari:

```cron
30 3 * * * CONTAINER_NAME=esfera-acta-extractor-web FAILURE_RETENTION_DAYS=30 FAILURE_MAX_SIZE_MB=1024 /usr/local/bin/esfera2excel-retention-cleanup >> /var/log/esfera2excel-retention.log 2>&1
```

Recomanació operativa:

- usa `FAILURE_RETENTION_DAYS` com a regla principal
- usa `FAILURE_MAX_SIZE_MB` com a límit de seguretat
- conserva l'esborrat manual per als casos que ja hagis investigat

## Desplegament recomanat

### Objectiu

Desplegar la web a `https://esfera2excel.marcos-a.com` darrere d'un proxy invers, amb el contenidor escoltant a `127.0.0.1:8000`.

### Persistència

Munta com a volums persistents:

- `/app/data`
- `/app/failed_uploads`

Configura també:

- `AUDIT_DB_PATH=/app/data/conversion_audit.sqlite3`
- `FAILURE_ROOT=/app/failed_uploads`

### Proxy invers

Exemple mínim amb Nginx:

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

El proxy ha de:

- servir `https://esfera2excel.marcos-a.com/`
- servir també `https://esfera2excel.marcos-a.com/admin`
- acabar TLS
- propagar `X-Forwarded-For`

### Comanda tipus de producció

```bash
docker run -d \
  --name esfera-acta-extractor-web \
  -p 127.0.0.1:8000:8000 \
  -e SECRET_KEY=canvia-aquesta-clau \
  -e ADMIN_USERNAME=admin \
  -e ADMIN_PASSWORD=canvia-aquesta-contrasenya \
  -e TELEGRAM_BOT_TOKEN=... \
  -e TELEGRAM_CHAT_ID=... \
  -e FAILURE_RETENTION_DAYS=30 \
  -e FAILURE_MAX_SIZE_MB=1024 \
  -e AUDIT_DB_PATH=/app/data/conversion_audit.sqlite3 \
  -e FAILURE_ROOT=/app/failed_uploads \
  -v /ruta/persistent/data:/app/data \
  -v /ruta/persistent/failed_uploads:/app/failed_uploads \
  esfera-acta-extractor-web
```

Bones pràctiques:

- no incrustis secrets a la imatge
- no exposis el port directament a internet si hi ha proxy invers
- canvia sempre `SECRET_KEY` i `ADMIN_PASSWORD`

## Rutes HTTP útils

| Ruta | Ús |
|---|---|
| `/` | Interfície pública |
| `/health` | Comprovació bàsica de salut |
| `/convert` | Inici de la conversió |
| `/convert/status/<job_id>` | Estat i progrés |
| `/convert/download/<job_id>` | Descàrrega del resultat |
| `/admin/login` | Accés d'administració |
| `/admin` | Tauler d'administració |

## Resolució de problemes

### El contenidor arrenca però no es pot iniciar sessió a `/admin`

Comprova:

- `ADMIN_USERNAME`
- `ADMIN_PASSWORD`
- `SECRET_KEY`

### Les alertes de Telegram no arriben

Comprova:

- que el bot existeix
- que li has enviat un missatge
- que `TELEGRAM_BOT_TOKEN` és correcte
- que `TELEGRAM_CHAT_ID` és correcte

### Els errors no desapareixen mai del servidor

Comprova:

- si has programat la neteja automàtica
- si `FAILURE_RETENTION_DAYS` i `FAILURE_MAX_SIZE_MB` estan definits
- si el `cron` s'està executant realment

## Via CLI antiga

El projecte manté la via CLI original per processar carpetes locals.

Entrades i sortides:

- entrada: `01_source_pdfs`
- sortida principal: `02_extracted_data`
- resums: `03_final_grade_summaries`

Execució amb Docker:

```bash
docker build -t esfera-acta-extractor .
docker run --rm -v "$(pwd):/app" esfera-acta-extractor python esfera-acta-extractor.py
```

La via CLI és útil per a proves o ús local tècnic, però no és la via recomanada per a desplegament.

## Desenvolupament

Llibreries principals:

- `Flask`
- `gunicorn`
- `pandas`
- `pdfplumber`
- `openpyxl`
- `rich`

## Contribució

1. Fes un fork del repositori.
2. Crea una branca de treball.
3. Fes els canvis.
4. Obre una pull request.

## Llicència

Aquest projecte es distribueix sota la llicència GNU GPL-3.0. Consulta [`LICENSE`](LICENSE).
