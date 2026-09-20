#!/usr/bin/env bash
# FOTOTAGE MUSTERSTADT 2026 — Umgebung einrichten
# Legt eine lokale Python-Umgebung (.venv) an und installiert openpyxl + reportlab.
#
#   ./setup.sh
#
# Danach:
#   ./.venv/bin/python scripts/menue.py
# oder Umgebung aktivieren:
#   source .venv/bin/activate

set -euo pipefail
cd "$(dirname "$0")"

PY="${PYTHON:-python3}"
echo "==> Python: $($PY --version)"

if [ ! -d .venv ]; then
  echo "==> Lege virtuelle Umgebung .venv an"
  "$PY" -m venv .venv
fi

PIP=(./.venv/bin/python -m pip)

# pip-Installation. Der zweite Versuch mit --use-deprecated=legacy-certs hilft
# hinter Firmen-Proxys, deren TLS-Zertifikat der macOS-Systemspeicher kennt,
# der pip-eigene aber nicht.
pip_install() {
  "${PIP[@]}" install "$@" || "${PIP[@]}" install --use-deprecated=legacy-certs "$@"
}

echo "==> Aktualisiere pip"
pip_install --upgrade pip || echo "   (pip-Upgrade übersprungen — nicht kritisch)"

echo "==> Installiere Abhängigkeiten aus requirements.txt"
pip_install -r requirements.txt

echo "==> Prüfe Installation"
./.venv/bin/python - <<'EOF'
import openpyxl, reportlab
print("  openpyxl ", openpyxl.__version__)
print("  reportlab", reportlab.Version)
EOF

echo
echo "Fertig. Test:"
echo "  ./.venv/bin/python scripts/menue.py"
