"""
Hedda & Pico Skript-Formatter — Web-App.
Drag & Drop .docx → formatiertes .docx herunterladen.
"""

import os
import io
from flask import Flask, request, jsonify, send_file, render_template_string
from werkzeug.utils import secure_filename
from formatter import format_document, get_output_filename

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", os.urandom(24).hex())

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)


HTML = r"""
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Hedda & Pico Skript-Formatter</title>
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

        .logo {
            font-size: 2.2rem;
            font-weight: 700;
            margin-bottom: 0.2rem;
            color: #fff;
            letter-spacing: -0.5px;
        }
        .logo span { color: #FF8C00; }

        .subtitle {
            color: #888;
            margin-bottom: 2rem;
            font-size: 0.9rem;
        }

        .container {
            width: 100%;
            max-width: 600px;
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

        /* How-To */
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
        .howto-content p, .howto-content li {
            font-size: 0.82rem;
            color: #999;
            line-height: 1.6;
        }
        .howto-content ol, .howto-content ul {
            padding-left: 1.2rem;
            margin: 0.3rem 0;
        }
        .howto-content li { margin-bottom: 0.3rem; }

        .color-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 0.3rem 1rem;
            margin: 0.5rem 0;
        }
        .color-item {
            display: flex;
            align-items: center;
            gap: 0.4rem;
            font-size: 0.78rem;
            color: #999;
        }
        .color-dot {
            width: 10px;
            height: 10px;
            border-radius: 50%;
            flex-shrink: 0;
        }

        .tip {
            background: #1e1b4b;
            border: 1px solid #312e81;
            border-radius: 8px;
            padding: 0.6rem 0.8rem;
            font-size: 0.8rem;
            color: #a5b4fc;
            margin-top: 0.8rem;
        }

        /* Case character input */
        .case-section {
            display: flex;
            align-items: center;
            gap: 0.8rem;
            margin-bottom: 1rem;
        }
        .case-section label {
            font-size: 0.85rem;
            color: #aaa;
            white-space: nowrap;
        }
        .case-input {
            flex: 1;
            padding: 0.5rem 0.7rem;
            background: #111;
            border: 1px solid #333;
            border-radius: 8px;
            color: #e0e0e0;
            font-size: 0.85rem;
            font-family: inherit;
        }
        .case-input:focus {
            outline: none;
            border-color: #800080;
        }
        .case-input::placeholder { color: #555; }
        .case-hint {
            font-size: 0.7rem;
            color: #666;
            margin-top: -0.5rem;
            margin-bottom: 1rem;
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
            border-color: #FF8C00;
            background: #1a1500;
        }
        .dropzone-icon {
            font-size: 2.5rem;
            margin-bottom: 0.8rem;
            opacity: 0.6;
        }
        .dropzone p { color: #888; font-size: 0.9rem; }
        .dropzone p strong { color: #FF8C00; }

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
            border-top-color: #FF8C00;
            border-radius: 50%;
            animation: spin 0.6s linear infinite;
        }
        @keyframes spin { to { transform: rotate(360deg); } }

        .footer {
            margin-top: 2rem;
            color: #555;
            font-size: 0.8rem;
        }
    </style>
</head>
<body>
    <div class="logo">Hedda <span>&</span> Pico</div>
    <p class="subtitle">Skript-Formatter</p>

    <div class="container">
        <!-- How-To -->
        <div class="card" style="padding: 1rem 1.5rem;">
            <div class="howto-toggle" onclick="toggleHowTo()">
                <span class="arrow" id="howtoArrow">▶</span>
                So funktioniert's
            </div>
            <div id="howtoContent" class="howto-content">
                <h3>Skript reinziehen — fertig</h3>
                <p>Zieh deine .docx-Datei in das Feld unten (oder klick drauf). Die formatierte Version wird automatisch heruntergeladen. Das Datum wird an den Dateinamen angehängt.</p>

                <h3>Was wird formatiert?</h3>
                <ul>
                    <li><strong>Titel</strong> — „HEDDA & PICO" in 48pt, fett, Versalien</li>
                    <li><strong>Folge</strong> — in 26pt, fett</li>
                    <li><strong>Szenen</strong> — in 13pt, fett, unterstrichen, Versalien. 1 Leerzeile davor</li>
                    <li><strong>Rollen</strong> — fett, Versalien, farbig. Text klebt eng dran (0,5 Zeilenabstand)</li>
                    <li><strong>Erzähler</strong> — Text kursiv und eingerückt</li>
                    <li><strong>Regieanweisungen</strong> — kursiv, 9pt. Nach Rolle: auf gleicher Zeile</li>
                    <li><strong>SFX/ATM</strong> — kursiv, 9pt, eigene Zeile</li>
                    <li><strong>LEIT-OBJEKT</strong> — „LEIT-OBJEKT" fett + kursiv</li>
                    <li><strong>Kumulierte Zeit</strong> — kursiv, 9pt, rechtsbündig</li>
                    <li><strong>Klammern</strong> — eckige [] werden zu runden ()</li>
                    <li><strong>Seitenzahlen</strong> — 9pt, rechtsbündig</li>
                </ul>

                <h3>Farben der Rollen</h3>
                <div class="color-grid">
                    <div class="color-item"><span class="color-dot" style="background:#000000;border:1px solid #555;"></span> Erzähler — schwarz</div>
                    <div class="color-item"><span class="color-dot" style="background:#00B050;"></span> Wendt — grün</div>
                    <div class="color-item"><span class="color-dot" style="background:#FF8C00;"></span> Hedda — orange</div>
                    <div class="color-item"><span class="color-dot" style="background:#C88A00;"></span> Frau Fischer — ocker</div>
                    <div class="color-item"><span class="color-dot" style="background:#008B8B;"></span> Herr Novak — dunkeltürkis</div>
                    <div class="color-item"><span class="color-dot" style="background:#8B4513;"></span> Herr Hassan — braun</div>
                    <div class="color-item"><span class="color-dot" style="background:#FF69B4;"></span> Oma Stein — pink</div>
                    <div class="color-item"><span class="color-dot" style="background:#FF0000;"></span> Dr. Khoury — rot</div>
                    <div class="color-item"><span class="color-dot" style="background:#800080;"></span> Fall-Charakter — lila</div>
                    <div class="color-item"><span class="color-dot" style="background:#0000FF;"></span> Andere — blau (Abstufungen)</div>
                </div>

                <h3>Datenschutz</h3>
                <p>Deine Skripte werden <strong>nicht</strong> gespeichert. Die Dateien werden nur kurz verarbeitet und sofort gelöscht. Nichts bleibt auf dem Server.</p>

                <div class="tip">
                    <strong>Tipp:</strong> Trag unten den Fall-Charakter ein (z.B. „Max" oder „Nele"), damit er in Lila erscheint. Alle anderen unbekannten Rollen werden automatisch in Blautönen eingefärbt.
                </div>
            </div>
        </div>

        <!-- Drop zone -->
        <div class="card">
            <div class="case-section">
                <label>Fall-Charakter:</label>
                <input type="text" id="caseChar" class="case-input"
                       placeholder="z.B. Max, Nele ..." value="">
            </div>
            <p class="case-hint">Name der Person, um die es im Fall geht — wird in <strong style="color:#800080;">lila</strong> dargestellt.</p>

            <div id="dropzone" class="dropzone">
                <div class="dropzone-icon">📄</div>
                <p><strong>.docx Dateien hierher ziehen</strong><br>oder klicken zum Auswählen</p>
                <input type="file" id="fileInput" accept=".docx" multiple style="display:none">
            </div>
            <div id="fileList" class="file-list"></div>
        </div>
    </div>

    <p class="footer">Texte bleiben erhalten — nur die Formatierung wird angepasst.<br>v2.0</p>

    <script>
        // Persist case character
        const caseInput = document.getElementById('caseChar');
        const saved = localStorage.getItem('hp_case_char');
        if (saved) caseInput.value = saved;
        caseInput.addEventListener('input', () => {
            localStorage.setItem('hp_case_char', caseInput.value);
        });

        function toggleHowTo() {
            const c = document.getElementById('howtoContent');
            const a = document.getElementById('howtoArrow');
            c.classList.toggle('open');
            a.textContent = c.classList.contains('open') ? '▼' : '▶';
        }

        const dropzone = document.getElementById('dropzone');
        const fileInput = document.getElementById('fileInput');
        const fileList = document.getElementById('fileList');

        dropzone.addEventListener('click', () => fileInput.click());

        ['dragenter', 'dragover'].forEach(e => {
            dropzone.addEventListener(e, ev => { ev.preventDefault(); dropzone.classList.add('active'); });
        });
        ['dragleave', 'drop'].forEach(e => {
            dropzone.addEventListener(e, () => dropzone.classList.remove('active'));
        });

        dropzone.addEventListener('drop', e => {
            e.preventDefault();
            for (const f of e.dataTransfer.files) {
                if (f.name.endsWith('.docx')) processFile(f);
            }
        });

        fileInput.addEventListener('change', e => {
            for (const f of e.target.files) {
                if (f.name.endsWith('.docx')) processFile(f);
            }
            fileInput.value = '';
        });

        async function processFile(file) {
            const item = document.createElement('div');
            item.className = 'file-item';
            item.innerHTML = `
                <span class="fname">${file.name}</span>
                <div class="spinner"></div>
                <span class="status processing">Formatiere...</span>
            `;
            fileList.prepend(item);

            const fd = new FormData();
            fd.append('file', file);

            // Send case character(s)
            const caseVal = document.getElementById('caseChar').value.trim();
            if (caseVal) fd.append('case_characters', caseVal);

            try {
                const res = await fetch('/format', { method: 'POST', body: fd });

                if (!res.ok) {
                    const ct = res.headers.get('content-type') || '';
                    let msg = 'Fehler';
                    if (ct.includes('application/json')) {
                        const err = await res.json();
                        msg = err.error || msg;
                    }
                    throw new Error(msg);
                }

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
    f = request.files.get("file")
    if not f or not f.filename.endswith(".docx"):
        return jsonify({"error": "Bitte eine .docx-Datei hochladen"}), 400

    # Parse case characters (comma-separated)
    case_raw = request.form.get("case_characters", "")
    case_characters = set()
    if case_raw:
        for name in case_raw.split(","):
            name = name.strip()
            if name:
                case_characters.add(name.upper())

    tmp_path = os.path.join(UPLOAD_DIR, "tmp_" + secure_filename(f.filename))
    try:
        f.save(tmp_path)
        result_bytes = format_document(tmp_path, case_characters=case_characters)
        output_name = get_output_filename(f.filename)

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
