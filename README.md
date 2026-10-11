# Winsol Presentatievertaler

Online tool: sleep een PowerPoint (.pptx, max. 30 MB) in de tool, kies de
taal, en download de vertaalde presentatie. De opmaak blijft 100% behouden:
enkel de tekst in de bestaande tekstvakken wordt vervangen. De
Winsol-woordenlijst (NL/FR/DE/EN/ES) is daarbij bindend.

## Architectuur

Zelfde patroon als Feedbackloop: **frontend op GitHub Pages, enkel de API
als Worker op Cloudflare.**

- `index.html` + `app.js`: statische frontend (GitHub Pages). Pakt de PPTX
  **in de browser** uit (JSZip), haalt per alinea de tekst-runs eruit, zet
  de vertaling terug in exact dezelfde runs, past de taalcode aan
  (spellingcontrole) en pakt het bestand opnieuw in, in dezelfde volgorde.
  Het bestand verlaat de browser nooit.
- `worker.js`: Cloudflare Worker. Krijgt enkel losse teksten binnen, voegt
  de woordenlijst-termen toe die in die teksten voorkomen en laat Claude
  vertalen. De Anthropic API-sleutel staat **alleen** hier (Worker secret).
- Login en rollen zoals bij Feedbackloop: users en sessies in Workers KV,
  wachtwoorden als `sha256(wachtwoord + PASSWORD_PEPPER)`.

### Wat wordt bewaard?

| Wat | Waar | Hoe lang |
|---|---|---|
| Presentaties, vertalingen | **nergens** (enkel in het geheugen van het browsertabblad) | tot het tabblad sluit |
| Gebruikers (naam, hash, rol) | Workers KV `user:*` | tot een admin ze overschrijft |
| Sessies | Workers KV `session:*` | 12 u, of 30 dagen met "Blijf aangemeld" |
| Woordenlijst | Workers KV `glossary` | tot een admin een nieuwe oplaadt |

De teksten gaan wel, voor de duur van de vertaling, via de Worker naar de
Anthropic API. Daar gelden de voorwaarden van het Anthropic-account (API-data
wordt niet gebruikt om modellen te trainen; bewaartermijnen: zie de
voorwaarden van het account, of vraag Zero Data Retention aan via Anthropic).

### Rollen

- **user**: vertalen.
- **admin**: vertalen + woordenlijst opladen/downloaden + gebruikers opladen.

### Hoe de vertaling de opmaak bewaart

1. Elke alinea wordt doorgestuurd als lijst van "segmenten" (de runs van
   PowerPoint, elk met eigen opmaak). Claude moet exact evenveel segmenten
   teruggeven; de tool zet segment *i* terug in run *i*.
2. Regeleindes, velden (dianummer, datum), symbolen (Wingdings-pijltjes) en
   segmenten zonder letters worden nooit aangeraakt. Spaties/tabs vooraan en
   achteraan elk segment blijven behouden.
3. Krijgt een alinea een fout aantal segmenten terug, dan wordt ze één keer
   opnieuw vertaald. Lukt dat nog niet, dan komt de volledige vertaling in
   het eerste segment (tekst klopt, opmaak van losse woorden gaat verloren) —
   dit staat in het rapport.
4. Identieke alinea's (bv. de infokaders "Categorie(ën): …") worden één keer
   vertaald en overal hetzelfde ingevuld.
5. Er wordt niets herschaald of verschoven. Tekst die merkbaar langer werd
   (> 15%), staat in het rapport om na te kijken.

Wat wordt vertaald: dia's, tabellen, SmartArt, groepen, zichtbare tekst in
lay-outs/masters, en (optioneel) sprekersnotities. Niet: tekst in
afbeeldingen, grafiekdata, verborgen voorbeeldtekst in lay-outs.

### Grenzen (kost)

- Bestand max. **30 MB** (`MAX_FILE_MB` in `app.js`).
- Tekst max. **150.000 tekens** per presentatie (`MAX_TOTAL_CHARS`). Dit
  bepaalt de AI-kost, niet de bestandsgrootte: afbeeldingen worden nooit
  verstuurd. Ter vergelijking: R&D Vertelt Q3 (32 dia's) bevat ±5.000 tekens.
- Per aanroep naar de Worker max. 120 alinea's / 12.000 tekens
  (`worker.js`); de frontend splitst zelf op (80 / 8.000) en stuurt 4
  aanroepen tegelijk.

## Opzetten (eenmalig)

### 1. Frontend — GitHub Pages

Repo (public) → **Settings → Pages** → Source: "Deploy from a branch" →
Branch: `main`, map `/ (root)` → Save. Na een paar minuten is de tool
bereikbaar op `https://winsol-belgie.github.io/pptx-vertaler/`.

> Er staan geen geheimen in de code. Woordenlijst en gebruikers staan in de
> KV, niet in de repo (`.gitignore` sluit `*.xlsx` en `*.pptx` uit).

### 2. API — Cloudflare Worker

```bash
npm install

# KV-namespace aanmaken en de id in wrangler.toml zetten (REPLACE_WITH_KV_ID)
npx wrangler kv namespace create VERTALER_KV

npx wrangler deploy
# → toont de *.workers.dev-URL

npx wrangler secret put ANTHROPIC_API_KEY
npx wrangler secret put PASSWORD_PEPPER
# → PEPPER: een lange willekeurige tekst. Bewaar hem veilig: wijzig je hem
#   later, dan werken alle bestaande wachtwoorden niet meer.
```

### 3. Frontend koppelen aan de Worker

Zet de URL uit stap 2 in `app.js`, bovenaan, in `API_URL`. Commit en push.

### 4. Eerste admin aanmaken

Daarna beheer je gebruikers via de tool zelf, maar de allereerste admin
moet rechtstreeks in de KV:

```bash
node scripts/hash-password.js "<wachtwoord>" "<pepper>"
# → kopieer de hash
```

Maak een bestand `admin.json` (niet committen) met:

```json
{"username":"GwennV","role":"admin","passwordHash":"<hash>"}
```

```bash
npx wrangler kv key put "user:GwennV" --path admin.json --binding VERTALER_KV --remote
```

Verwijder `admin.json` daarna. Log in op de tool, en laad in de
admin-sectie de woordenlijst en de Users-Excel op.

## Beheer in de tool (admin)

- **Woordenlijst beheren**: Excel met één kolom per taal (koppen
  *Nederlands, Frans, Duits, Engels, Spaans*; volgorde vrij, lege rijen en
  extra spaties worden genegeerd). De opgeladen lijst **vervangt** de
  vorige. De tool waarschuwt bij dubbele termen met een andere vertaling.
  Met "Huidige lijst downloaden" krijg je de actieve lijst terug als Excel.
- **Gebruikers beheren**: Excel met kolommen `USER`, `PW`, `ROLE`
  (`admin`/`user`) — identiek aan Feedbackloop. Bestaande gebruikers die
  niet in het bestand staan, blijven ongewijzigd. Iemand verwijderen kan
  (voorlopig) enkel via
  `npx wrangler kv key delete "user:<naam>" --binding VERTALER_KV --remote`.

Bewaar de Users-Excel met leesbare wachtwoorden niet op een gedeelde plek.
Dit is een tussenoplossing tot de SSO-koppeling.

## Bij een nieuwe versie

- Enkel frontend gewijzigd: `git push`.
- **`app.js` gewijzigd?** Verhoog ook het versienummer in
  `<script src="app.js?v=...">` onderaan `index.html`.
- `worker.js` gewijzigd: `npx wrangler deploy`.

## Lokaal testen

```bash
npx wrangler dev
```

Zet `ANTHROPIC_API_KEY` en `PASSWORD_PEPPER` in een lokaal `.dev.vars`,
zet `API_URL` in `app.js` tijdelijk op `http://localhost:8787` en open
`index.html` in de browser. Voor de eerste admin lokaal: zelfde
`kv key put`-commando zonder `--remote`.
