# fritzbox-phonebook-converter

Fritzbox Telefonbuch Generator — rein clientseitige Web-App (HTML/JS, keine Server-Komponente).

- Namen und Telefonnummern direkt im Browser eintragen und verwalten
- Export als formatierte Excel-Datei (.xlsx)
- Import bestehender Excel-Dateien
- Export als FRITZ!Box-Telefonbuch-XML (Telefonbuch → Telefonbuch-Import)

Alle Daten bleiben ausschließlich im Browser (localStorage) und werden nirgendwo hochgeladen.

## Nutzung

Einfach `index.html` öffnen oder über GitHub Pages aufrufen.

## Lokal testen

```bash
python3 -m http.server 8090
```

Dann `http://127.0.0.1:8090` öffnen.
