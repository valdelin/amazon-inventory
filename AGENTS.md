# AGENTS.md — amazon-inventory

Scripts Python para coletar inventário e volume de vendas da Amazon via SP-API e gerar XLSX.

## Como executar

- Use sempre o venv: `.venv/bin/python`, nunca `python`.
- Comandos prontos no `README.md` (parte de referência deste doc).
- Rodou coisa nova? Confirme o arquivo XLSX gerado existe e abre (`--out` padrão imprime o caminho).

## Convenções que o README não confessa

- `inventario.py --mensal` cobre 1º do mês até hoje; mês fechado precisa de `--start`/`--end`
  explícitos (a Amazon só devolve dados recentes pela janela padrão).
- Relatórios de mês antigo (>30 dias) usam a Orders API (`--historico`) e são paginados —
  podem demorar vários minutos.
- Maio/2026 foi o mês `cache_mensal_202605.json` original — a regra do cache é
  `cache_mensal_<aaaamm>.json`; um mês só está completo se esse arquivo existir.
- `evolucao.py` relê todos os `relatorio_mensal_*.xlsx`; rode-o DEPOIS de cada `--mensal`.

## Santos que não podem vazar

- `.env` (LWA creds + refresh token) e qualquer `relatorio/cache` com dados reais
  são ignorados no git. Sempre verificar `git status` antes de commitar.

## Venv

- Python 3.14 no venv; recriar com `python3 -m venv .venv` se quebrar.