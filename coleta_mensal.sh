#!/usr/bin/env bash
# coleta_mensal.sh — coleta o mês fechado e regenera o dashboard.
#
# Chamado pelo systemd user timer `amazon-inventory-mensal.timer`.
# Usa o refresh token persistente no .env (não renova manualmente).
#
# Uso manual:
#   ./coleta_mensal.sh                    # mês passado (padrão)
#   ./coleta_mensal.sh --start ... --end  # janela explícita
set -euo pipefail

cd "$(dirname "$0")"
PY=".venv/bin/python"
OUT_DIR="."
LOG_DIR="$HOME/.local/state/amazon-inventory"
mkdir -p "$LOG_DIR"

# Mês passado por padrão (fechado, exige --start/--end na API Orders).
START="${1:-$(date -d '1 month ago' +%Y-%m-01)}"

if [[ $# -ge 1 && "$1" == "--mensal" ]]; then
    "$PY" inventario.py --mensal > "$LOG_DIR/mensal.log" 2>&1
else
    END="$(date -d "$START + 1 month - 1 day" +%Y-%m-%d)"
    "$PY" inventario.py --mensal --start "$START" --end "$END" \
        >> "$LOG_DIR/mensal.log" 2>&1
fi

"$PY" dashboard.py --out "$OUT_DIR/analytics.html" >> "$LOG_DIR/mensal.log" 2>&1

echo "OK — coleta mensal concluída ($(date '+%Y-%m-%d %H:%M'))" | tee -a "$LOG_DIR/mensal.log"
