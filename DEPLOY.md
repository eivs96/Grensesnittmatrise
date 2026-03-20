# Publisering på eget domene

Denne appen er en Node/Express-app med SQLite-lagring. Det betyr at den må kjores pa en server eller host som kan:

- kjore `node server.js`
- eksponere en port til web
- lagre filer permanent i `data/projects.db`

## Enkel oppskrift

1. Legg prosjektet i Git eller last det opp til serveren din.
2. Start appen med `npm install` og `npm start`, eller bruk Docker.
3. Pek domenet ditt til serveren med DNS.
4. Legg en reverse proxy foran appen for HTTPS, for eksempel Nginx eller hostens innebygde domenestyring.

## Viktig for denne appen

- Frontend og backend serveres fra samme app.
- API-et bruker relative stier som `/api/projects`, sa domenet ditt vil fungere uten kodeendringer sa lenge hele appen ligger bak samme domene.
- SQLite-databasen lagres lokalt i `data/projects.db`, sa du ma bruke en host med persistent lagring.
- Du kan overstyre lagringsplassering med `DATA_DIR` eller `DB_PATH` hvis hosten din krever en bestemt mount-path.

## Kjoring uten Docker

```bash
npm install
npm start
```

Appen starter pa port `3000` som standard, eller pa porten i miljovariabelen `PORT`.

Eksempel med egendefinert datamappe:

```bash
DATA_DIR=/var/lib/grensesnittmatrise npm start
```

## Kjoring med Docker

Bygg image:

```bash
docker build -t grensesnittmatrise .
```

Start container med vedvarende lagring:

```bash
docker run -d ^
  --name grensesnittmatrise ^
  -p 3000:3000 ^
  -v grensesnittmatrise-data:/app/data ^
  grensesnittmatrise
```

Hvis du heller vil peke mot en eksplisitt filsti:

```bash
docker run -d ^
  --name grensesnittmatrise ^
  -p 3000:3000 ^
  -e DB_PATH=/app/data/projects.db ^
  -v grensesnittmatrise-data:/app/data ^
  grensesnittmatrise
```

## Domenekobling

Hvis du bruker rot-domenet, for eksempel `dittdomene.no`:

- legg inn en `A`-record mot serverens offentlige IP

Hvis du bruker subdomene, for eksempel `app.dittdomene.no`:

- legg inn en `CNAME` til hostnavnet du far fra leverandoren, eller `A`-record til server-IP

## HTTPS

For at siden skal vare trygg og uten advarsler i nettleseren, ma domenet ha SSL/TLS.

Vanlige oppsett:

- Nginx + Let's Encrypt pa egen VPS
- Domenestyring/SSL direkte hos hostingleverandoren

## Anbefalt hostingtype

For denne appen passer disse typene hosting:

- VPS med Node.js
- Docker-hosting med persistent volume
- Plattform som stotter langkjorende Node-app og vedvarende disk

Ren statisk hosting er ikke nok, fordi appen har backend og lagrer data.
