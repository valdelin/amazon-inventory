#!/usr/bin/env bash
# coleta_semanal.sh — fecha a semana: regenera o dashboard sem bater na API.
#
# Chamado pelo systemd user timer `amazon-inventory-semanal.timer`.
# Decisão (vault 2026-09-21): semanal NÃO gera xlsx próprio — só regenera o
# analytics.html a partir do cache do mês (snapshot de fechamento semanal).
set -euo pipefail

cd "$(dirname "$0")"
PY=".venv/bin/python"
LOG_DIR="$HOME/.local/state/amazon-inventory"
OUT_DIR="."
mkdir -p "$LOG_DIR"

"$PY" dashboard.py --out "$OUT_DIR/analytics.html" >> "$LOG_DIR/semanal.log" 2>&1

echo "OK — fechamento semanal do dashboard ($(date '+%Y-%m-%d %H:%M'))" | tee -a "$LOG_DIR/semanal.log"
