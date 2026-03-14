"""
DocFormatter — Standardisiert .docx-Formatierung per natürlichsprachigen Regeln.
Nutzt Gemini Flash zur Regelkompilierung, dann deterministische Anwendung.
"""

import os
import json
from flask import Flask, request, jsonify, send_file, render_template_string
from werkzeug.utils import secure_filename
from formatter import format_document, get_output_filename
import io

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", os.urandom(24).hex())

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# Gemini API key: env var takes precedence
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

COMPILE_PROMPT = '''Du bist ein Formatierungsregel-Compiler für Hörspiel- und Drehbuch-Skripte.

Der Benutzer gibt dir Formatierungsregeln in natürlicher Sprache. Du wandelst diese in ein strukturiertes JSON-Format um.

WICHTIG:
- Antworte NUR mit validem JSON. Keine Erklärungen, kein Markdown, keine Code-Blöcke.
- Regex-Patterns müssen valides Python-Regex sein.
- Alle Farben als #RRGGBB Hex-Code.
- Fehlende Werte sinnvoll mit Defaults füllen.

JSON-Schema:
{
  "defaults": {
    "font": "Schriftart (string)",
    "size": "Schriftgröße in pt (number)",
    "color": "#RRGGBB",
    "bold": false,
    "italic": false,
    "alignment": "left|center|right|justify",
    "line_spacing": 1.5
  },
  "filename_suffix_date": "true wenn Datum im Dateinamen gewünscht (boolean)",
  "filename_date_format": "DDMM|MMDD|YYYYMMDD",
  "page_numbers": {
    "enabled": true/false,
    "size": "pt (number)",
    "alignment": "left|center|right"
  },
  "scene_blank_lines": "Leerzeilen zwischen Szenen (number, 0 = keine extra)",
  "paragraph_rules": [
    {
      "name": "Beschreibender Name der Regel",
      "match_type": "first_contains|contains|starts_with|regex",
      "pattern": "Python-Regex-Pattern",
      "case_insensitive": true,
      "format": {
        "size": "pt (number, optional)",
        "bold": "boolean (optional)",
        "italic": "boolean (optional)",
        "uppercase": "boolean (optional)",
        "color": "#RRGGBB (optional)",
        "alignment": "left|center|right (optional)"
      }
    }
  ],
  "dialogue_rules": {
    "enabled": true/false,
    "detection_pattern": "Regex mit Capture-Group für den Rollennamen, z.B. ^([A-ZÄÖÜ][A-ZÄÖÜ\\\\s.\\\\-]+?)\\\\s*:",
    "name_format": {
      "bold": true/false,
      "uppercase": true/false
    },
    "character_colors": {
      "CHARAKTERNAME": {
        "name_color": "#RRGGBB (Farbe des Rollennamens)",
        "text_color": "#RRGGBB (optional, Farbe des Dialotexts, default = defaults.color)",
        "text_italic": "boolean (optional)",
        "text_indent_cm": "number (optional, Einrückung in cm)"
      }
    },
    "case_character_color": "#RRGGBB (optional, Farbe für den Fall-Charakter)",
    "default_color": "#RRGGBB (Farbe für unbekannte Nebencharaktere)",
    "secondary_shades": ["#RRGGBB", "..."]
  },
  "inline_rules": [
    {
      "name": "Beschreibender Name",
      "pattern": "Python-Regex-Pattern",
      "format": {
        "bold": "boolean (optional)",
        "italic": "boolean (optional)",
        "uppercase": "boolean (optional)",
        "color": "#RRGGBB (optional)"
      }
    }
  ]
}

HINWEISE ZUR INTERPRETATION:
- "Versalien" = uppercase: true
- "fett" = bold: true
- "kursiv" = italic: true
- "eingerückt" = text_indent_cm: 1.27 (Standard-Einrückung)
- "linksbündig/rechtsbündig/zentriert" = alignment

WICHTIG - SKRIPTFORMAT:
- In Hörspiel-Skripten stehen Rollennamen auf EIGENEN Zeilen (z.B. "ERZÄHLER" allein auf einer Zeile), der Dialogtext folgt in den nächsten Zeilen.
- Das ist NICHT das Format "NAME: Text" auf einer Zeile!
- Der Formatter erkennt Rollennamen automatisch anhand der character_colors-Liste.
- name_color = Farbe des Rollennamens (die Zeile mit dem Namen)
- text_color = Farbe des Dialogs (die Zeilen NACH dem Namen). Default = defaults.color (schwarz).
- text_italic, text_indent_cm = gelten für den Dialog-Text nach dem Namen.

FARB-INTERPRETATION:
- Wenn die Regeln sagen "die Rolle steht in folgender Farbe" → name_color setzen.
- Wenn die Regeln sagen "der Text/Redebeitrag steht in folgender Farbe" → text_color setzen.
- Lies die Regeln genau: "steht die Rolle, nicht aber ihr Text, in folgender Farbe" = name_color.
- "steht nicht die Rolle, aber ihr Text in folgender Farbe" = text_color.

PARAGRAPH-RULES NAMING:
- Verwende beschreibende deutsche Namen die folgende Schlüsselwörter enthalten:
  - "Serientitel" oder "Titel" für den Seriennamen
  - "Folge" oder "Episode" für Folgenbezeichnungen
  - "Szene" oder "Scene" für Szenenüberschriften
  - "Zeit" oder "Kumuliert" für Zeitangaben
- paragraph_rules werden in Reihenfolge geprüft, erste Übereinstimmung gewinnt.
- "first_contains" = nur die erste Stelle im Dokument die passt.
- Für inline_rules: Patterns sollten Textteile INNERHALB von Absätzen matchen.
- SFX/ATM und Regieanweisungen in eckigen Klammern [] werden vom Formatter automatisch erkannt und kursiv gesetzt.

Benutzer-Regeln:
---
{rules}
---'''

HTML = r"""
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>DocFormatter</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', system-ui, sans-serif;
            background: #0f0f0f;
            color: #e0e0e0;
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            align-items: center;
            padding: 2rem 1rem;
        }

        h1 {
            font-size: 1.8rem;
            font-weight: 600;
            margin-bottom: 0.3rem;
            color: #fff;
        }

        .subtitle {
            color: #888;
            margin-bottom: 2rem;
            font-size: 0.95rem;
        }

        .container {
            width: 100%;
            max-width: 700px;
            display: flex;
            flex-direction: column;
            gap: 1.5rem;
        }

        .card {
            background: #1a1a1a;
            border: 1px solid #2a2a2a;
            border-radius: 12px;
            padding: 1.5rem;
        }

        .card h2 {
            font-size: 1rem;
            font-weight: 600;
            margin-bottom: 1rem;
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }

        .card h2 .step {
            background: #333;
            color: #aaa;
            width: 24px;
            height: 24px;
            border-radius: 50%;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            font-size: 0.75rem;
            flex-shrink: 0;
        }

        /* Settings toggle */
        .settings-toggle {
            color: #666;
            font-size: 0.8rem;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 0.3rem;
            margin-bottom: 1rem;
        }

        .settings-toggle:hover { color: #999; }

        .settings-panel {
            display: none;
            margin-bottom: 1rem;
        }

        .settings-panel.open { display: block; }

        .settings-panel label {
            font-size: 0.8rem;
            color: #888;
            display: block;
            margin-bottom: 0.3rem;
        }

        .settings-panel input {
            width: 100%;
            padding: 0.5rem 0.7rem;
            background: #222;
            border: 1px solid #333;
            border-radius: 6px;
            color: #e0e0e0;
            font-size: 0.85rem;
            font-family: monospace;
        }

        .settings-panel input:focus {
            outline: none;
            border-color: #6366f1;
        }

        .settings-hint {
            font-size: 0.7rem;
            color: #555;
            margin-top: 0.3rem;
        }

        /* Rules textarea */
        .rules-area {
            width: 100%;
            min-height: 250px;
            padding: 0.8rem;
            background: #111;
            border: 1px solid #333;
            border-radius: 8px;
            color: #e0e0e0;
            font-size: 0.85rem;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', system-ui, sans-serif;
            line-height: 1.5;
            resize: vertical;
        }

        .rules-area:focus {
            outline: none;
            border-color: #6366f1;
        }

        .rules-area::placeholder {
            color: #555;
        }

        .btn-row {
            display: flex;
            align-items: center;
            gap: 0.8rem;
            margin-top: 0.8rem;
        }

        .btn {
            padding: 0.55rem 1.2rem;
            border: none;
            border-radius: 8px;
            font-size: 0.85rem;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.15s;
            font-family: inherit;
        }

        .btn-primary {
            background: #6366f1;
            color: #fff;
        }

        .btn-primary:hover { background: #5558e6; }

        .btn-primary:disabled {
            opacity: 0.5;
            cursor: not-allowed;
        }

        .compile-status {
            font-size: 0.8rem;
            color: #888;
        }

        .compile-status.ok { color: #4ade80; }
        .compile-status.error { color: #f87171; }
        .compile-status.loading { color: #818cf8; }

        /* Rules status badge */
        .rules-badge {
            font-size: 0.75rem;
            padding: 0.15rem 0.5rem;
            border-radius: 4px;
            margin-left: auto;
        }

        .rules-badge.active {
            background: #0d2818;
            color: #4ade80;
            border: 1px solid #1a4d2e;
        }

        .rules-badge.inactive {
            background: #2a1a0a;
            color: #f59e0b;
            border: 1px solid #4d3319;
        }

        /* Drop zone */
        .dropzone {
            border: 2px dashed #333;
            border-radius: 12px;
            padding: 3rem 1.5rem;
            text-align: center;
            transition: all 0.2s;
            cursor: pointer;
        }

        .dropzone.active {
            border-color: #6366f1;
            background: #1a1a2e;
        }

        .dropzone.disabled {
            opacity: 0.4;
            pointer-events: none;
        }

        .dropzone-icon {
            font-size: 2.5rem;
            margin-bottom: 0.8rem;
            opacity: 0.6;
        }

        .dropzone p {
            color: #888;
            font-size: 0.9rem;
        }

        .dropzone p strong { color: #6366f1; }

        /* File list */
        .file-list {
            display: flex;
            flex-direction: column;
            gap: 0.5rem;
            margin-top: 1rem;
        }

        .file-item {
            display: flex;
            align-items: center;
            gap: 0.8rem;
            padding: 0.7rem 0.8rem;
            background: #222;
            border-radius: 8px;
            font-size: 0.85rem;
        }

        .file-item .fname { flex: 1; color: #ccc; }

        .file-item .status {
            font-size: 0.8rem;
            padding: 0.2rem 0.6rem;
            border-radius: 4px;
        }

        .status.processing { background: #1e1b4b; color: #818cf8; }
        .status.done { background: #0d2818; color: #4ade80; cursor: pointer; text-decoration: none; }
        .status.done:hover { background: #134e2a; }
        .status.error { background: #2a0a0a; color: #f87171; }

        .spinner {
            width: 14px; height: 14px;
            border: 2px solid #333;
            border-top-color: #818cf8;
            border-radius: 50%;
            animation: spin 0.6s linear infinite;
        }

        @keyframes spin { to { transform: rotate(360deg); } }

        .footer {
            margin-top: 2rem;
            color: #555;
            font-size: 0.8rem;
        }

        /* How-To Card */
        .howto-card {
            background: #1a1a1a;
            border: 1px solid #2a2a2a;
            border-radius: 12px;
            padding: 1rem 1.5rem;
        }

        .howto-toggle {
            display: flex;
            align-items: center;
            gap: 0.5rem;
            cursor: pointer;
            color: #ccc;
            font-size: 0.95rem;
            font-weight: 600;
        }

        .howto-toggle:hover { color: #fff; }

        .howto-toggle .arrow {
            font-size: 0.7rem;
            color: #888;
            transition: transform 0.2s;
        }

        .howto-content {
            display: none;
            margin-top: 1rem;
            padding-top: 1rem;
            border-top: 1px solid #2a2a2a;
        }

        .howto-content.open { display: block; }

        .howto-content h3 {
            font-size: 0.85rem;
            font-weight: 600;
            color: #e0e0e0;
            margin: 1rem 0 0.4rem 0;
        }

        .howto-content h3:first-child { margin-top: 0; }

        .howto-content p,
        .howto-content li {
            font-size: 0.82rem;
            color: #999;
            line-height: 1.6;
        }

        .howto-content ol,
        .howto-content ul {
            padding-left: 1.2rem;
            margin: 0.3rem 0;
        }

        .howto-content li { margin-bottom: 0.3rem; }

        .howto-content code {
            background: #222;
            padding: 0.1rem 0.4rem;
            border-radius: 3px;
            font-size: 0.78rem;
            color: #ccc;
        }

        .howto-content .tip {
            background: #1e1b4b;
            border: 1px solid #312e81;
            border-radius: 8px;
            padding: 0.6rem 0.8rem;
            font-size: 0.8rem;
            color: #a5b4fc;
            margin-top: 0.8rem;
        }

        /* Compiled rules preview */
        .compiled-preview {
            display: none;
            margin-top: 0.8rem;
        }

        .compiled-preview.open { display: block; }

        .compiled-toggle {
            font-size: 0.75rem;
            color: #555;
            cursor: pointer;
            margin-top: 0.5rem;
        }

        .compiled-toggle:hover { color: #888; }

        .compiled-json {
            width: 100%;
            min-height: 120px;
            max-height: 300px;
            padding: 0.6rem;
            background: #0a0a0a;
            border: 1px solid #222;
            border-radius: 6px;
            color: #888;
            font-size: 0.7rem;
            font-family: 'SF Mono', 'Fira Code', monospace;
            line-height: 1.4;
            resize: vertical;
            overflow: auto;
        }
    </style>
</head>
<body>
    <h1>DocFormatter</h1>
    <p class="subtitle">Formatiere Skripte einheitlich per Drag & Drop</p>

    <div class="container">
        <!-- How-To -->
        <div class="howto-card">
            <div class="howto-toggle" onclick="toggleHowTo()">
                <span class="arrow" id="howtoArrow">▶</span>
                So funktioniert's
            </div>
            <div id="howtoContent" class="howto-content">
                <h3>In 3 Schritten zum formatierten Skript</h3>
                <ol>
                    <li><strong>Regeln schreiben</strong> — Beschreibe im Textfeld, wie dein Dokument aussehen soll. Schriftart, Schriftgröße, Farben, Zeilenabstand — alles in normalem Deutsch.</li>
                    <li><strong>"Regeln kompilieren" klicken</strong> — Eine KI wandelt deine Beschreibung einmalig in technische Formatierungsregeln um. Das passiert nur einmal, nicht bei jeder Datei.</li>
                    <li><strong>.docx-Dateien reinziehen</strong> — Zieh deine Word-Dateien in das Feld unten oder klick drauf. Die formatierte Version wird automatisch heruntergeladen.</li>
                </ol>

                <h3>Regeln ändern?</h3>
                <p>Einfach den Text oben anpassen und nochmal "Regeln kompilieren" klicken. Deine Regeln und Einstellungen werden im Browser gespeichert — beim nächsten Besuch ist alles noch da.</p>

                <h3>API-Key Hinweis</h3>
                <p>Die App nutzt ein inklusives Kontingent zum Kompilieren der Regeln. Falls du eine Meldung bekommst, dass das Limit erreicht ist, kannst du unter "Einstellungen" einen eigenen kostenlosen Gemini API Key eintragen. Den bekommst du hier:</p>
                <p><a href="https://aistudio.google.com/apikey" target="_blank" style="color: #818cf8;">aistudio.google.com/apikey</a> — Google-Konto reicht, kostet nichts.</p>

                <div class="tip">
                    <strong>Tipp:</strong> Die KI versteht auch komplexe Regeln — z.B. unterschiedliche Farben pro Rolle, Szenenüberschriften in Versalien, oder kursive Regieanweisungen. Je genauer du beschreibst, desto besser das Ergebnis.
                </div>
            </div>
        </div>

        <!-- Settings (collapsible) -->
        <div class="card" style="padding: 1rem 1.5rem;">
            <div class="settings-toggle" onclick="toggleSettings()">
                <span id="settingsArrow">▶</span> Einstellungen
            </div>
            <div id="settingsPanel" class="settings-panel">
                <label>Eigener Gemini API Key (optional)</label>
                <input type="password" id="apiKey" placeholder="AIza..." value="">
                <p class="settings-hint">
                    Nur nötig, wenn das inklusive Kontingent aufgebraucht ist.<br>
                    Kostenlos erstellen: <a href="https://aistudio.google.com/apikey" target="_blank" style="color: #818cf8;">aistudio.google.com/apikey</a>
                    {% if has_env_key %}<br><span style="color:#4ade80;">Inklusiver Server-Key aktiv — dieses Feld ist optional.</span>{% endif %}
                </p>
            </div>
        </div>

        <!-- Step 1: Rules -->
        <div class="card">
            <h2>
                <span class="step">1</span> Formatierungsregeln
                <span id="rulesBadge" class="rules-badge inactive">Nicht kompiliert</span>
            </h2>
            <textarea id="rulesText" class="rules-area" placeholder="Formatierungsregeln hier eingeben...

Beispiel:
Der gesamte Text ist in der Schriftart Calibri geschrieben.
Die Standardschriftgröße ist 12.
Alles ist linksbündig.
Alles ist in schwarz geschrieben.
Der Zeilenabstand ist 1.5.
..."></textarea>
            <div class="btn-row">
                <button id="compileBtn" class="btn btn-primary" onclick="compileRules()">
                    Regeln kompilieren
                </button>
                <span id="compileStatus" class="compile-status"></span>
            </div>
            <div class="compiled-toggle" onclick="togglePreview()">
                <span id="previewArrow">▶</span> Kompilierte Regeln anzeigen
            </div>
            <div id="compiledPreview" class="compiled-preview">
                <pre id="compiledJson" class="compiled-json">Noch keine Regeln kompiliert.</pre>
            </div>
        </div>

        <!-- Step 2: Format files -->
        <div class="card">
            <h2><span class="step">2</span> Skripte formatieren</h2>
            <div id="dropzone" class="dropzone disabled">
                <div class="dropzone-icon">📄</div>
                <p><strong>.docx Dateien hierher ziehen</strong><br>oder klicken zum Auswählen</p>
                <input type="file" id="fileInput" accept=".docx" multiple style="display:none">
            </div>
            <div id="fileList" class="file-list"></div>
        </div>
    </div>

    <p class="footer">Texte bleiben erhalten — nur die Formatierung wird angepasst.</p>

    <script>
        const STORAGE_RULES = 'docfmt_rules_text';
        const STORAGE_COMPILED = 'docfmt_compiled';
        const STORAGE_APIKEY = 'docfmt_apikey';

        // --- Init ---
        function init() {
            // Load saved state
            const savedRules = localStorage.getItem(STORAGE_RULES);
            if (savedRules) document.getElementById('rulesText').value = savedRules;

            const savedKey = localStorage.getItem(STORAGE_APIKEY);
            if (savedKey) document.getElementById('apiKey').value = savedKey;

            const savedCompiled = localStorage.getItem(STORAGE_COMPILED);
            if (savedCompiled) {
                try {
                    JSON.parse(savedCompiled);
                    document.getElementById('compiledJson').textContent = savedCompiled;
                    enableDropzone();
                    document.getElementById('rulesBadge').className = 'rules-badge active';
                    document.getElementById('rulesBadge').textContent = 'Aktiv';
                } catch (e) {}
            }

            // Auto-save rules text
            document.getElementById('rulesText').addEventListener('input', () => {
                localStorage.setItem(STORAGE_RULES, document.getElementById('rulesText').value);
                // Mark as needing recompile
                document.getElementById('rulesBadge').className = 'rules-badge inactive';
                document.getElementById('rulesBadge').textContent = 'Geändert — neu kompilieren';
            });

            // Auto-save API key
            document.getElementById('apiKey').addEventListener('input', () => {
                localStorage.setItem(STORAGE_APIKEY, document.getElementById('apiKey').value);
            });
        }

        init();

        // --- How-To toggle ---
        function toggleHowTo() {
            const content = document.getElementById('howtoContent');
            const arrow = document.getElementById('howtoArrow');
            content.classList.toggle('open');
            arrow.textContent = content.classList.contains('open') ? '▼' : '▶';
        }

        // --- Settings toggle ---
        function toggleSettings() {
            const panel = document.getElementById('settingsPanel');
            const arrow = document.getElementById('settingsArrow');
            panel.classList.toggle('open');
            arrow.textContent = panel.classList.contains('open') ? '▼' : '▶';
        }

        function togglePreview() {
            const panel = document.getElementById('compiledPreview');
            const arrow = document.getElementById('previewArrow');
            panel.classList.toggle('open');
            arrow.textContent = panel.classList.contains('open') ? '▼' : '▶';
        }

        // --- Compile rules ---
        async function compileRules() {
            const btn = document.getElementById('compileBtn');
            const status = document.getElementById('compileStatus');
            const rules = document.getElementById('rulesText').value.trim();
            const apiKey = document.getElementById('apiKey').value.trim();

            if (!rules) {
                status.textContent = 'Bitte Regeln eingeben.';
                status.className = 'compile-status error';
                return;
            }

            btn.disabled = true;
            status.textContent = 'Kompiliere...';
            status.className = 'compile-status loading';

            try {
                const res = await fetch('/compile-rules', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ rules, api_key: apiKey }),
                });

                const data = await res.json();

                if (!res.ok) {
                    throw new Error(data.error || 'Kompilierung fehlgeschlagen');
                }

                // Save compiled rules
                const compiled = JSON.stringify(data.compiled, null, 2);
                localStorage.setItem(STORAGE_COMPILED, compiled);
                document.getElementById('compiledJson').textContent = compiled;

                status.textContent = 'Erfolgreich kompiliert!';
                status.className = 'compile-status ok';

                document.getElementById('rulesBadge').className = 'rules-badge active';
                document.getElementById('rulesBadge').textContent = 'Aktiv';

                enableDropzone();
            } catch (err) {
                status.textContent = err.message;
                status.className = 'compile-status error';
            } finally {
                btn.disabled = false;
            }
        }

        function enableDropzone() {
            document.getElementById('dropzone').classList.remove('disabled');
        }

        // --- Drag & drop ---
        const dropzone = document.getElementById('dropzone');
        const fileInput = document.getElementById('fileInput');
        const fileList = document.getElementById('fileList');

        dropzone.addEventListener('click', () => {
            if (!dropzone.classList.contains('disabled')) fileInput.click();
        });

        ['dragenter', 'dragover'].forEach(evt => {
            dropzone.addEventListener(evt, (e) => {
                e.preventDefault();
                if (!dropzone.classList.contains('disabled')) dropzone.classList.add('active');
            });
        });

        ['dragleave', 'drop'].forEach(evt => {
            dropzone.addEventListener(evt, () => dropzone.classList.remove('active'));
        });

        dropzone.addEventListener('drop', (e) => {
            e.preventDefault();
            if (!dropzone.classList.contains('disabled')) handleFiles(e.dataTransfer.files);
        });

        fileInput.addEventListener('change', (e) => {
            handleFiles(e.target.files);
            fileInput.value = '';
        });

        function handleFiles(files) {
            for (const file of files) {
                if (!file.name.endsWith('.docx')) continue;
                processFile(file);
            }
        }

        async function processFile(file) {
            const id = 'f-' + Date.now() + '-' + Math.random().toString(36).slice(2, 6);
            const compiled = localStorage.getItem(STORAGE_COMPILED);

            if (!compiled) {
                alert('Bitte zuerst Regeln kompilieren.');
                return;
            }

            const item = document.createElement('div');
            item.className = 'file-item';
            item.id = id;
            item.innerHTML = `
                <span class="fname">${file.name}</span>
                <div class="spinner"></div>
                <span class="status processing">Formatiere...</span>
            `;
            fileList.prepend(item);

            const fd = new FormData();
            fd.append('file', file);
            fd.append('compiled_rules', compiled);

            try {
                const res = await fetch('/format', { method: 'POST', body: fd });

                if (!res.ok) {
                    const err = await res.json();
                    throw new Error(err.error || 'Fehler');
                }

                // Get filename from Content-Disposition header
                const disposition = res.headers.get('Content-Disposition');
                let outName = file.name.replace('.docx', '_formatted.docx');
                if (disposition) {
                    const match = disposition.match(/filename\*?=(?:UTF-8''|"?)([^";]+)/);
                    if (match) outName = decodeURIComponent(match[1].replace(/"/g, ''));
                }

                const blob = await res.blob();
                const url = URL.createObjectURL(blob);

                item.innerHTML = `
                    <span class="fname">${file.name}</span>
                    <a class="status done" href="${url}" download="${outName}">Herunterladen</a>
                `;

                // Auto-download
                const a = document.createElement('a');
                a.href = url;
                a.download = outName;
                a.click();
            } catch (err) {
                item.innerHTML = `
                    <span class="fname">${file.name}</span>
                    <span class="status error">${err.message}</span>
                `;
            }
        }
    </script>
</body>
</html>
"""


@app.route("/")
def index():
    has_env_key = bool(GEMINI_API_KEY)
    return render_template_string(HTML, has_env_key=has_env_key)


@app.route("/compile-rules", methods=["POST"])
def compile_rules():
    """Send rules to Gemini Flash, get compiled JSON back."""
    data = request.get_json()
    rules_text = data.get("rules", "").strip()
    api_key = data.get("api_key", "").strip() or GEMINI_API_KEY

    if not rules_text:
        return jsonify({"error": "Keine Regeln angegeben"}), 400

    if not api_key:
        return jsonify({"error": "Kein API-Key — bitte in den Einstellungen hinterlegen oder als GEMINI_API_KEY Environment Variable setzen"}), 400

    try:
        from google import genai

        client = genai.Client(api_key=api_key)

        prompt = COMPILE_PROMPT.replace("{rules}", rules_text)

        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=prompt,
        )

        # Extract JSON from response
        response_text = response.text.strip()

        # Strip markdown code blocks if present
        if response_text.startswith("```"):
            lines = response_text.split("\n")
            # Remove first and last ``` lines
            if lines[0].startswith("```"):
                lines = lines[1:]
            if lines and lines[-1].strip() == "```":
                lines = lines[:-1]
            response_text = "\n".join(lines)

        compiled = json.loads(response_text)

        # Validate basic structure
        if "defaults" not in compiled:
            compiled["defaults"] = {
                "font": "Arial", "size": 12, "color": "#000000",
                "bold": False, "italic": False, "alignment": "left", "line_spacing": 1.5
            }

        return jsonify({"ok": True, "compiled": compiled})

    except json.JSONDecodeError as e:
        return jsonify({"error": f"Gemini hat kein valides JSON zurückgegeben: {str(e)}"}), 500
    except ImportError:
        return jsonify({"error": "google-genai Paket nicht installiert. Bitte 'pip install google-genai' ausführen."}), 500
    except Exception as e:
        error_msg = str(e)
        if "API_KEY" in error_msg.upper() or "401" in error_msg or "403" in error_msg:
            return jsonify({"error": "Ungültiger API-Key. Bitte prüfen."}), 401
        return jsonify({"error": f"Fehler bei der Kompilierung: {error_msg}"}), 500


@app.route("/format", methods=["POST"])
def format_file():
    """Format a .docx file using compiled rules."""
    f = request.files.get("file")
    if not f or not f.filename.endswith(".docx"):
        return jsonify({"error": "Bitte eine .docx-Datei hochladen"}), 400

    compiled_json = request.form.get("compiled_rules", "{}")
    try:
        compiled_rules = json.loads(compiled_json)
    except json.JSONDecodeError:
        return jsonify({"error": "Ungültige kompilierte Regeln"}), 400

    if not compiled_rules.get("defaults"):
        return jsonify({"error": "Regeln nicht kompiliert — bitte zuerst kompilieren"}), 400

    # Save temp file
    tmp_path = os.path.join(UPLOAD_DIR, "tmp_" + secure_filename(f.filename))
    try:
        f.save(tmp_path)
        result_bytes = format_document(tmp_path, compiled_rules)
        output_name = get_output_filename(f.filename, compiled_rules)

        return send_file(
            io.BytesIO(result_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name=output_name,
        )
    except Exception as e:
        return jsonify({"error": f"Formatierung fehlgeschlagen: {str(e)}"}), 500
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5009))
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(host="0.0.0.0", port=port, debug=debug)
