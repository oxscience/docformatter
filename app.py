"""
DocFormatter — Standardisiert .docx-Formatierung anhand von Regeln.
"""

import os
from flask import Flask, request, jsonify, send_file, render_template_string, session
from werkzeug.utils import secure_filename
from formatter import format_document, DEFAULT_RULES
import io

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", os.urandom(24))

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

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
            padding: 2rem;
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
            max-width: 640px;
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

        /* Rules form */
        .rules-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 0.8rem;
        }

        .rule-group {
            display: flex;
            flex-direction: column;
            gap: 0.3rem;
        }

        .rule-group.full-width {
            grid-column: 1 / -1;
        }

        .rule-group label {
            font-size: 0.8rem;
            color: #888;
            font-weight: 500;
        }

        .rule-group input,
        .rule-group select {
            padding: 0.5rem 0.7rem;
            background: #222;
            border: 1px solid #333;
            border-radius: 6px;
            color: #e0e0e0;
            font-size: 0.85rem;
            font-family: inherit;
        }

        .rule-group input:focus,
        .rule-group select:focus {
            outline: none;
            border-color: #6366f1;
        }

        .rule-group select {
            cursor: pointer;
        }

        .checkboxes {
            grid-column: 1 / -1;
            display: flex;
            gap: 1.2rem;
            flex-wrap: wrap;
            margin-top: 0.3rem;
        }

        .checkbox-item {
            display: flex;
            align-items: center;
            gap: 0.4rem;
            font-size: 0.85rem;
            cursor: pointer;
        }

        .checkbox-item input[type="checkbox"] {
            width: 16px;
            height: 16px;
            accent-color: #6366f1;
            cursor: pointer;
        }

        .save-hint {
            grid-column: 1 / -1;
            font-size: 0.75rem;
            color: #555;
            margin-top: 0.2rem;
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

        .dropzone-icon {
            font-size: 2.5rem;
            margin-bottom: 0.8rem;
            opacity: 0.6;
        }

        .dropzone p {
            color: #888;
            font-size: 0.9rem;
        }

        .dropzone p strong {
            color: #6366f1;
        }

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

        .status.processing {
            background: #1e1b4b;
            color: #818cf8;
        }

        .status.done {
            background: #0d2818;
            color: #4ade80;
            cursor: pointer;
            text-decoration: none;
        }

        .status.done:hover { background: #134e2a; }

        .status.error {
            background: #2a0a0a;
            color: #f87171;
        }

        .spinner {
            width: 14px;
            height: 14px;
            border: 2px solid #333;
            border-top-color: #818cf8;
            border-radius: 50%;
            animation: spin 0.6s linear infinite;
        }

        @keyframes spin {
            to { transform: rotate(360deg); }
        }

        .footer {
            margin-top: 2rem;
            color: #555;
            font-size: 0.8rem;
        }
    </style>
</head>
<body>
    <h1>DocFormatter</h1>
    <p class="subtitle">Formatiere Skripte einheitlich per Drag & Drop</p>

    <div class="container">
        <!-- Step 1: Rules -->
        <div class="card">
            <h2><span class="step">1</span> Formatierungsregeln</h2>
            <div class="rules-grid">
                <div class="rule-group">
                    <label>Schriftart</label>
                    <select id="fontName">
                        <option value="Arial">Arial</option>
                        <option value="Times New Roman">Times New Roman</option>
                        <option value="Calibri">Calibri</option>
                        <option value="Helvetica">Helvetica</option>
                        <option value="Georgia">Georgia</option>
                        <option value="Verdana">Verdana</option>
                        <option value="Courier New">Courier New</option>
                        <option value="Garamond">Garamond</option>
                        <option value="Cambria">Cambria</option>
                        <option value="Palatino Linotype">Palatino Linotype</option>
                    </select>
                </div>
                <div class="rule-group">
                    <label>Zeilenabstand</label>
                    <select id="lineSpacing">
                        <option value="1.0">1.0</option>
                        <option value="1.15">1.15</option>
                        <option value="1.5" selected>1.5</option>
                        <option value="2.0">2.0</option>
                    </select>
                </div>
                <div class="rule-group">
                    <label>Fließtext (pt)</label>
                    <input type="number" id="bodySize" value="12" min="8" max="36">
                </div>
                <div class="rule-group">
                    <label>Titel (pt)</label>
                    <input type="number" id="titleSize" value="20" min="8" max="48">
                </div>
                <div class="rule-group">
                    <label>Überschrift 1 (pt)</label>
                    <input type="number" id="h1Size" value="16" min="8" max="48">
                </div>
                <div class="rule-group">
                    <label>Überschrift 2 (pt)</label>
                    <input type="number" id="h2Size" value="14" min="8" max="48">
                </div>
                <div class="checkboxes">
                    <label class="checkbox-item">
                        <input type="checkbox" id="removeColors" checked>
                        Farben entfernen
                    </label>
                    <label class="checkbox-item">
                        <input type="checkbox" id="removeBold">
                        Fett entfernen (Fließtext)
                    </label>
                    <label class="checkbox-item">
                        <input type="checkbox" id="removeItalic">
                        Kursiv entfernen
                    </label>
                </div>
                <p class="save-hint">Einstellungen werden im Browser gespeichert.</p>
            </div>
        </div>

        <!-- Step 2: Convert -->
        <div class="card">
            <h2><span class="step">2</span> Skripte formatieren</h2>
            <div id="dropzone" class="dropzone">
                <div class="dropzone-icon">📄</div>
                <p><strong>.docx Dateien hierher ziehen</strong><br>oder klicken zum Auswählen</p>
                <input type="file" id="fileInput" accept=".docx" multiple style="display:none">
            </div>
            <div id="fileList" class="file-list"></div>
        </div>
    </div>

    <p class="footer">Texte bleiben erhalten — nur die Formatierung wird angepasst.</p>

    <script>
        const dropzone = document.getElementById('dropzone');
        const fileInput = document.getElementById('fileInput');
        const fileList = document.getElementById('fileList');

        // --- Settings persistence ---
        const STORAGE_KEY = 'docformatter_rules';
        const fields = {
            fontName: 'select',
            lineSpacing: 'select',
            bodySize: 'input',
            titleSize: 'input',
            h1Size: 'input',
            h2Size: 'input',
            removeColors: 'checkbox',
            removeBold: 'checkbox',
            removeItalic: 'checkbox',
        };

        function loadSettings() {
            try {
                const saved = JSON.parse(localStorage.getItem(STORAGE_KEY));
                if (!saved) return;
                for (const [id, type] of Object.entries(fields)) {
                    const el = document.getElementById(id);
                    if (!el || !(id in saved)) continue;
                    if (type === 'checkbox') el.checked = saved[id];
                    else el.value = saved[id];
                }
            } catch (e) {}
        }

        function saveSettings() {
            const data = {};
            for (const [id, type] of Object.entries(fields)) {
                const el = document.getElementById(id);
                if (type === 'checkbox') data[id] = el.checked;
                else data[id] = el.value;
            }
            localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
        }

        // Auto-save on any change
        for (const id of Object.keys(fields)) {
            document.getElementById(id).addEventListener('change', saveSettings);
        }
        loadSettings();

        function getRules() {
            return {
                font_name: document.getElementById('fontName').value,
                body_size: parseInt(document.getElementById('bodySize').value) || 12,
                title_size: parseInt(document.getElementById('titleSize').value) || 20,
                heading1_size: parseInt(document.getElementById('h1Size').value) || 16,
                heading2_size: parseInt(document.getElementById('h2Size').value) || 14,
                line_spacing: parseFloat(document.getElementById('lineSpacing').value) || 1.5,
                remove_colors: document.getElementById('removeColors').checked,
                remove_bold: document.getElementById('removeBold').checked,
                remove_italic: document.getElementById('removeItalic').checked,
            };
        }

        // --- Drag & drop ---
        dropzone.addEventListener('click', () => fileInput.click());

        ['dragenter', 'dragover'].forEach(evt => {
            dropzone.addEventListener(evt, (e) => {
                e.preventDefault();
                dropzone.classList.add('active');
            });
        });

        ['dragleave', 'drop'].forEach(evt => {
            dropzone.addEventListener(evt, () => {
                dropzone.classList.remove('active');
            });
        });

        dropzone.addEventListener('drop', (e) => {
            e.preventDefault();
            handleFiles(e.dataTransfer.files);
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
            fd.append('rules', JSON.stringify(getRules()));

            try {
                const res = await fetch('/format', { method: 'POST', body: fd });

                if (!res.ok) {
                    const err = await res.json();
                    throw new Error(err.error || 'Fehler');
                }

                const blob = await res.blob();
                const url = URL.createObjectURL(blob);
                const outName = file.name.replace('.docx', '_formatted.docx');

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
    return render_template_string(HTML)


@app.route("/format", methods=["POST"])
def format_file():
    import json

    f = request.files.get("file")
    if not f or not f.filename.endswith(".docx"):
        return jsonify({"error": "Bitte eine .docx-Datei hochladen"}), 400

    # Parse rules from request
    rules_json = request.form.get("rules", "{}")
    try:
        rules = json.loads(rules_json)
    except json.JSONDecodeError:
        rules = {}

    # Merge with defaults
    from formatter import DEFAULT_RULES
    merged = DEFAULT_RULES.copy()
    merged.update(rules)

    # Validate numeric fields
    for key in ["body_size", "title_size", "heading1_size", "heading2_size"]:
        merged[key] = max(8, min(48, int(merged.get(key, 12))))
    merged["line_spacing"] = max(0.5, min(3.0, float(merged.get("line_spacing", 1.5))))

    # Save temp file
    tmp_path = os.path.join(UPLOAD_DIR, "tmp_" + secure_filename(f.filename))
    try:
        f.save(tmp_path)
        result_bytes = format_document(tmp_path, merged)
        return send_file(
            io.BytesIO(result_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name=f.filename.replace(".docx", "_formatted.docx"),
        )
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5009))
    app.run(host="0.0.0.0", port=port, debug=os.environ.get("FLASK_DEBUG", "0") == "1")
