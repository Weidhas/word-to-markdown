# Word2Markdown (Browser-only)

Eine statische Website, die `.docx` lokal im Browser in Markdown umwandelt.

## Features

- Nur `.docx` als Eingabe
- Umwandlung im Browser (kein Upload der Datei)
- Markdown-Ausgabe im Textfeld
- Copy-to-Clipboard
- Download als `.md`
- Design-fokussierte responsive Oberfläche
- Direkt als Azure Static Web App nutzbar

## Datenschutz

Die Dokumentkonvertierung läuft ausschließlich im Browser.

Hinweis: In dieser Version werden Bibliotheken per ESM-CDN geladen. Es werden keine Dokumentinhalte hochgeladen, aber die Seite lädt Bibliotheken über das Netzwerk. Für streng isolierte Umgebungen sollten die Bibliotheken lokal gebündelt und mit ausgeliefert werden.

## Lokales Starten

Da es eine statische App ist, genügt ein lokaler HTTP-Server.

Option A (VS Code Live Server):
- Projektordner öffnen
- `index.html` mit Live Server starten

Option B (Python, falls vorhanden):
- `python -m http.server 4173`
- Dann `http://localhost:4173` öffnen

## Deployment auf Azure Static Web Apps

1. Repository nach GitHub pushen.
2. In Azure Portal: Static Web App erstellen und mit dem Repository verbinden.
3. Build-Vorgaben:
   - App location: `/`
   - Api location: (leer lassen)
   - Output location: `/`
4. Azure erstellt automatisch einen Workflow.
5. Nach Deployment die URL öffnen und testen.

## Struktur

- `index.html` UI-Struktur
- `app.css` Styling und responsive Design
- `app.js` DOCX->HTML->Markdown Pipeline, Copy/Download
- `staticwebapp.config.json` Azure Static Web Apps Routing und Header

## Technische Notizen

- DOCX-Extraktion: `mammoth`
- HTML->Markdown: `turndown` + `turndown-plugin-gfm`
- Bilder werden als eingebettete Referenz markiert (`#embedded-image`) statt binärer Daten zu serialisieren.
