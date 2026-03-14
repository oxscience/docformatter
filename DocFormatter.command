#!/bin/bash
# ============================================
#  DocFormatter — Doppelklick zum Starten
# ============================================

# In das Verzeichnis des Skripts wechseln
cd "$(dirname "$0")"

clear
echo ""
echo "  ╔══════════════════════════════════════╗"
echo "  ║         DocFormatter starten         ║"
echo "  ╚══════════════════════════════════════╝"
echo ""

# ---------- Python prüfen ----------
PYTHON=""
if command -v python3 &>/dev/null; then
    PYTHON="python3"
elif command -v python &>/dev/null; then
    PYTHON="python"
fi

if [ -z "$PYTHON" ]; then
    echo "  ❌ Python wurde nicht gefunden."
    echo ""
    echo "  ┌──────────────────────────────────────────────────┐"
    echo "  │  So installierst du Python (dauert 2 Minuten):   │"
    echo "  │                                                   │"
    echo "  │  1. Öffne Safari und geh auf:                    │"
    echo "  │     python.org/downloads                          │"
    echo "  │                                                   │"
    echo "  │  2. Klick den gelben Button                      │"
    echo "  │     \"Download Python 3.x.x\"                      │"
    echo "  │                                                   │"
    echo "  │  3. Öffne die heruntergeladene .pkg Datei        │"
    echo "  │     (im Downloads-Ordner)                        │"
    echo "  │                                                   │"
    echo "  │  4. Klick dich durch den Installer               │"
    echo "  │     (immer \"Fortfahren\" / \"Installieren\")        │"
    echo "  │                                                   │"
    echo "  │  5. Danach dieses Skript nochmal doppelklicken   │"
    echo "  └──────────────────────────────────────────────────┘"
    echo ""

    # Anbieten, die Download-Seite direkt zu öffnen
    echo "  Soll ich die Download-Seite jetzt öffnen? (j/n)"
    read -n 1 REPLY
    echo ""
    if [ "$REPLY" = "j" ] || [ "$REPLY" = "J" ] || [ "$REPLY" = "y" ]; then
        open "https://www.python.org/downloads/"
        echo "  ✓ Safari wurde geöffnet."
    fi

    echo ""
    echo "  Drücke eine Taste zum Schließen..."
    read -n 1
    exit 1
fi

echo "  ✓ Python gefunden: $($PYTHON --version)"

# ---------- Erstinstallation ----------
if [ ! -d "venv" ]; then
    echo ""
    echo "  Erstinstallation — wird einmalig eingerichtet..."
    echo "  (Das kann 1–2 Minuten dauern)"
    echo ""

    $PYTHON -m venv venv
    if [ $? -ne 0 ]; then
        echo "  ❌ Konnte virtuelle Umgebung nicht erstellen."
        echo "  Drücke eine Taste zum Schließen..."
        read -n 1
        exit 1
    fi

    source venv/bin/activate
    pip install --quiet flask python-docx
    if [ $? -ne 0 ]; then
        echo "  ❌ Pakete konnten nicht installiert werden."
        echo "  Drücke eine Taste zum Schließen..."
        read -n 1
        exit 1
    fi

    echo "  ✓ Installation abgeschlossen!"
else
    source venv/bin/activate
    echo "  ✓ Umgebung geladen"
fi

echo ""
echo "  ➜  App startet auf: http://127.0.0.1:5009"
echo "  ➜  Browser öffnet sich gleich..."
echo ""
echo "  Zum Beenden: dieses Fenster schließen"
echo "  ──────────────────────────────────────"
echo ""

# ---------- Browser öffnen (kurz warten bis Server da ist) ----------
(
    sleep 2
    open "http://127.0.0.1:5009"
) &

# ---------- App starten ----------
venv/bin/python app.py 2>&1
