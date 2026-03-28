# Word To Markdown (Browser-only)

Eine statische Web-App, die `.docx` lokal im Browser in Markdown umwandelt.

## Live-Version

Das Projekt ist öffentlich erreichbar unter:

https://w2m.lurz.at

## Aktueller Funktionsumfang

- Nur `.docx` als Eingabe
- Drag-and-Drop oder Dateiauswahl
- Dateigrößenlimit (derzeit 12 MB)
- Clientseitige Umwandlung (`.docx` -> HTML -> Markdown)
- Markdown-Ausgabe im Editorfeld
- Kopieren in die Zwischenablage
- Download als `.md` (UTF-8 mit BOM)
- Umschaltbare Markdown-Vorschau (sanitiziertes HTML)
- Responsive Oberfläche inkl. Footer/Attribution
- Azure Static Web Apps kompatibel

## Änderungen seit Erstellung

Basierend auf der bisherigen Commit-Historie:

1. **Initiale Version** (`first commit`)
   - Grundstruktur als statische Browser-App
   - DOCX-zu-Markdown Konvertierung lokal im Browser
2. **CI-Setup für Azure Static Web Apps**
   - Automatisch erzeugte SWA-Workflow-Datei hinzugefügt
3. **Refactor von Deployment/UI**
   - SWA-Workflow-Datei wieder entfernt
   - UI-Elemente überarbeitet
4. **Markdown-Vorschau ergänzt**
   - Toggle-Button für Vorschauansicht hinzugefügt
   - Rendering-Pipeline für Vorschau integriert
5. **Footer + Bild-Handling verbessert**
   - Footer mit Attribution ergänzt
   - Bildbehandlung in der Markdown-Konvertierung angepasst

## Datenschutz

Die Dokumentkonvertierung läuft ausschließlich im Browser.

Hinweis: In der aktuellen Version werden Bibliotheken per ESM-CDN geladen. Es werden keine Dokumentinhalte hochgeladen, aber die Seite lädt Bibliotheken über das Netzwerk. Für streng isolierte Umgebungen sollten die Bibliotheken lokal gebündelt und mit ausgeliefert werden.

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
2. Im Azure-Portal eine Static Web App erstellen und mit dem Repository verbinden.
3. Build-Vorgaben:
   - App location: `/`
   - Api location: (leer lassen)
   - Output location: `/`
4. Deployment ausführen und URL testen.

## Projektstruktur

- `index.html`: UI-Struktur, Meta-Tags, Einstiegspunkt
- `app.css`: Styling und Responsive Layout
- `app.js`: DOCX->HTML->Markdown Pipeline, Preview, Copy/Download
- `staticwebapp.config.json`: SPA-Fallback, Sicherheitsheader, MIME-Typen
- `robots.txt`: Crawler-Vorgaben

## Technische Notizen

- DOCX-Extraktion: `mammoth`
- HTML->Markdown: `turndown` + `turndown-plugin-gfm`
- Markdown-Rendering in der Vorschau: `marked`
- Sanitizing der Vorschau: `DOMPurify`
- Bilder werden nicht binär in Markdown übernommen, sondern durch einen Hinweistext ersetzt.
