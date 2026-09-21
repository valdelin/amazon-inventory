#!/usr/bin/env bash
# coleta_diaria.sh — coleta o mês corrente (01/MM → hoje) e regenera o dashboard.
#
# Chamado pelo systemd user timer `amazon-inventory-diario.timer`.
# NOTA: `inventario.py --mensal` sem janela já resolve p/ 1º do mês até hoje
# (mesma regra do resolve_window), então a coleta diária NÃO passa --start/--end.
set -euo pipefail

cd "$(dirname "$0")"
PY=".venv/bin/python"
LOG_DIR="$HOME/.local/state/amazon-inventory"
OUT_DIR="."
mkdir -p "$LOG_DIR"

"$PY" inventario.py --mensal > "$LOG_DIR/diario.log" 2>&1
"$PY" dashboard.py --out "$OUT_DIR/analytics.html" >> "$LOG_DIR/diario.log" 2>&1

echo "OK — coleta diária concluída ($(date '+%Y-%m-%d %H:%M'))" | tee -a "$LOG_DIR/diario.log"
