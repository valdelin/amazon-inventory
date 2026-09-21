#!/usr/bin/env python3
"""Gera analytics.html (dashboard estatico + Chart.js) a partir dos caches.

Le os caches do inventario.py (cache_historico_*.json + cache_mensal_*.json)
e consolida em um unico HTML autocontido com:

- cards: faturamento total, pedidos, unidades, meses cobertos, ticket medio
- grafico de evolucao mensal (faturamento; pedidos + unidades)
- top SKUs/ASIN por faturamento
- taxas Amazon (settlement) por tipo, quando disponiveis no cache
- estoque corrente (do inventario/listagens do ultimo --mensal) por valor

Uso:  .venv/bin/python dashboard.py [--out analytics.html]
"""

import argparse
import glob
import json
import os
import re
from collections import defaultdict
from datetime import date

from openpyxl import load_workbook

from inventario import (CACHE_DIR, build_sales_rows, parse_flat_file,
                        parse_settlement_fees, summarise_orders)

MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho",
         "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]

MONTH_RE = re.compile(r"(\d{4})(\d{2})")
MESES_INDEX = {nome: i for i, nome in enumerate(MESES)}


def month_label(key):
    ano, mes = key.split("-")
    return f"{MESES[int(mes) - 1].title()}/{ano}"


def cache_month(path):
    m = MONTH_RE.search(path.rsplit("/", 1)[-1])
    return f"{m.group(1)}-{m.group(2)}" if m else None


def load_cache_data(path):
    """Le um cache e devolve (start, end, summary, sales_rows, settlements, inventario)."""
    payload = json.loads(open(path, encoding="utf-8").read())
    settlements = payload.get("settlements") or []
    if "rows" in payload:
        header, rows = payload["header"], payload["rows"]
    else:
        header, rows = parse_flat_file(payload["pedidos"])
    summary = summarise_orders(header, rows)
    sales_rows = build_sales_rows(header, rows)
    inv_header, inv_rows = ([], [])
    if payload.get("inventario"):
        inv_header, inv_rows = parse_flat_file(payload["inventario"])
    return (payload.get("start"), payload.get("end"), summary, sales_rows,
            settlements, inv_header, inv_rows)


def consolidate(caches):
    """Junta os meses em um dataset unico.""" 
    months = {}
    for path in caches:
        key = cache_month(path)
        if not key:
            continue
        start, end, summary, sales_rows, settlements, \
            inv_header, inv_rows = load_cache_data(path)
        months[key] = {
            "start": start, "end": end,
            "pedidos": summary["pedidos"],
            "unidades": summary["unidades"],
            "total": summary["total"],
            "products": summary["products"],
            "daily": summary["daily"],
            "sales_rows": sales_rows,
            "settlements": settlements,
            "inv_header": inv_header,
            "inv_rows": inv_rows,
        }
    ordered = [months[k] for k in sorted(months)]
    return months, ordered


def fee_breakdown(settlements):
    """{tipo: total} a partir dos textos de settlement do cache."""
    totals = defaultdict(float)
    for texts in settlements:
        fees = parse_settlement_fees([texts])
        for order_fees in fees.values():
            for fee_type, value in order_fees.items():
                totals[fee_type] += value
    return dict(sorted(totals.items(), key=lambda kv: -abs(kv[1])))


def inventory_summary(rows, header):
    """Por SKU: nome, preco, qtd, valor a partir do relatorio de listagens."""
    idx = {name: i for i, name in enumerate(header)}
    i_sku, i_nome = idx.get("seller-sku"), idx.get("item-name")
    i_preco, i_qtd = idx.get("price"), idx.get("quantity")
    i_canal, i_status = idx.get("fulfillment-channel"), idx.get("status")
    out = []
    for r in rows:
        price = _f(r[i_preco]) if i_preco is not None else 0.0
        qty = _f(r[i_qtd]) if i_qtd is not None else 0.0
        out.append({
            "sku": r[i_sku] if i_sku is not None else "",
            "nome": (r[i_nome] or "")[:40] if i_nome is not None else "",
            "preco": price, "qtd": qty, "valor": round(price * qty, 2),
            "canal": r[i_canal] if i_canal is not None else "",
            "status": r[i_status] if i_status is not None else "",
        })
    ativos = [o for o in out if str(o["status"]).lower() == "active"]
    # sem quantidade no relatorio de listagens (vem dos FBA/AFN, nao autorizados),
    # ordenamos por preco unitario e o "valor" vira preco (catalogo/precificacao)
    ativos.sort(key=lambda o: -o["preco"])
    return ativos


def _f(value):
    try:
        return float(str(value).strip()) if str(value).strip() else 0.0
    except ValueError:
        return 0.0


def render_html(data):
    import html as H
    meses = [month_label(k) for k in data["chave_meses"]]
    js_meses = json.dumps(meses)
    js_total = json.dumps([m["total"] for m in data["meses"]])
    js_pedidos = json.dumps([m["pedidos"] for m in data["meses"]])
    js_unidades = json.dumps([m["unidades"] for m in data["meses"]])
    js_top = json.dumps(data["top_skus"])
    js_fees = json.dumps(data["fees"])
    js_estoque = json.dumps(data["estoque_top"])

    cards = data["cards"]
    fee_card = ""
    if data["fees"]:
        fee_total = sum(abs(v) for v in data["fees"].values())
        fee_card = (
            f"<div class='card'><h3>Taxas Amazon</h3>"
            f"<p class='num'>{_brl(fee_total)}</p>"
            f"<p class='sub'>janela dos settlements (~90 dias)</p></div>")
    estoque_card = ""
    if data["cards"].get("estoque"):
        estoque_card = (
            f"<div class='card'><h3>Estoque estimado</h3>"
            f"<p class='num'>{_brl(data['cards']['estoque'])}</p>"
            f"<p class='sub'>{data['cards'].get('skus_ativos', 0)} SKUs ativos</p></div>")

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Amazon — Analytics</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
  :root {{ --bg:#0f1420; --card:#161d2e; --fg:#e6ecf5; --mut:#8b96ab; --acc:#38bdf8; }}
  * {{ box-sizing:border-box; }}
  body {{ margin:0; background:var(--bg); color:var(--fg);
         font:14px/1.5 system-ui,Segoe UI,Roboto,sans-serif; padding:24px; }}
  h1 {{ font-size:20px; margin:0 0 4px; }}
  .muted {{ color:var(--mut); font-size:12px; margin-bottom:16px; }}
  .grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(190px,1fr)); gap:12px; }}
  .card {{ background:var(--card); border:1px solid #22304b; border-radius:10px; padding:14px; }}
  .card h3 {{ margin:0 0 6px; font-size:12px; color:var(--mut); font-weight:600;
             text-transform:uppercase; letter-spacing:.5px; }}
  .num {{ font-size:22px; font-weight:700; margin:0; }}
  .sub {{ margin:4px 0 0; font-size:11px; color:var(--mut); }}
  .panels {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(420px,1fr)); gap:12px; margin-top:16px; }}
  .panel {{ background:var(--card); border:1px solid #22304b; border-radius:10px; padding:14px; }}
  .panel h2 {{ margin:0 0 8px; font-size:14px; }}
  .panel p.hint {{ margin:0 0 8px; font-size:11px; color:var(--mut); }}
  canvas {{ max-height:340px; }}
  table {{ width:100%; border-collapse:collapse; font-size:12px; }}
  th,td {{ text-align:left; padding:5px 8px; border-bottom:1px solid #22304b; }}
  th {{ color:var(--mut); font-weight:600; }}
  td.num,th.num {{ text-align:right; }}
  .two-col {{ display:grid; grid-template-columns:1fr 1fr; gap:12px; }}
  @media (max-width:900px){{ .two-col {{ grid-template-columns:1fr; }} }}
</style>
</head>
<body>
<h1>Amazon — Analytics da loja</h1>
<p class="muted">Gerado em {date.today().isoformat()} · dados das vendas {data['cards']['meses_cobertos']} meses
 ({data['cards']['primeiro']} → {data['cards']['ultimo']})</p>

<div class="grid">
  <div class="card"><h3>Faturamento total</h3><p class="num">{_brl(cards['faturamento'])}</p>
      <p class="sub">{cards['pedidos']} pedidos · {cards['unidades']} unidades</p></div>
  <div class="card"><h3>Ticket médio</h3><p class="num">{_brl(cards['ticket_medio'])}</p>
      <p class="sub">por pedido</p></div>
  <div class="card"><h3>Melhor mês</h3><p class="num">{cards['melhor_mes']}</p>
      <p class="sub">{_brl(cards['melhor_valor'])}</p></div>
  {fee_card}
  {estoque_card}
</div>

<div class="panels">
  <div class="panel"><h2>Evolução do faturamento</h2>
      <p class="hint">R$ líquido por mês (itens, sem taxas/shipping)</p>
      <canvas id="c-faturamento"></canvas></div>
  <div class="panel"><h2>Pedidos e unidades</h2>
      <p class="hint">barras = pedidos · linha = unidades</p>
      <canvas id="c-pedidos"></canvas></div>
</div>

<div class="two-col">
  <div class="panel"><h2>Top SKUs (por faturamento)</h2>
      <table><tr><th>SKU</th><th>Produto</th><th class="num">Total</th></tr>
      {H.unescape(''.join(
          f"<tr><td>{_esc(r[0])}</td><td>{_esc(r[2])}</td>"
          f"<td class='num'>{_brl(r[5])}</td></tr>" for r in data['top_10']))
        if data['top_10'] else '<tr><td colspan=3>sem dados</td></tr>'}
      </table></div>
  <div class="panel"><h2>Taxas Amazon por tipo</h2>
      <p class="hint">soma por tipo de taxa do settlement (negativo = custo)</p>
      <canvas id="c-fees"></canvas></div>
</div>

<div class="panels">
  <div class="panel"><h2>Estoque corrente</h2>
      <p class="hint">top 15 SKUs por valor em estoque (listagens do último --mensal)</p>
      <canvas id="c-estoque"></canvas></div>
</div>

<script>
const meses = {js_meses};
const fat = {js_total};
const ped = {js_pedidos};
const uni = {js_unidades};
const topSku = {js_top};
const fees = {js_fees};
const estoque = {js_estoque};
const COL = '#38bdf8', COL2 = '#f59e0b', COL3 = '#34d399';

new Chart(document.getElementById('c-faturamento'), {{
  type:'bar',
  data:{{ labels:meses, datasets:[{{ label:'Faturamento', data:fat, backgroundColor:COL }}] }},
  options:{{ responsive:true, plugins:{{ legend:{{display:false}} }},
            scales:{{ y:{{ ticks:{{ callback:v=>'R$ '+v.toLocaleString('pt-BR') }} }} }} }}
}});

new Chart(document.getElementById('c-pedidos'), {{
  data:{{ labels:meses,
    datasets:[
      {{ label:'Pedidos', data:ped, type:'bar', backgroundColor:COL2, yAxisID:'y' }},
      {{ label:'Unidades', data:uni, type:'line', borderColor:COL, backgroundColor:COL, yAxisID:'y1' }} ] }},
  options:{{ responsive:true,
    scales:{{ y:{{ beginAtZero:true, position:'left' }}, y1:{{ beginAtZero:true, position:'right', grid:{{drawOnChartArea:false}} }} }} }}
}});

new Chart(document.getElementById('c-fees'), {{
  type:'bar',
  data:{{ labels:Object.keys(fees), datasets:[{{ label:'Taxa', data:Object.values(fees), backgroundColor:COL3 }}] }},
  options:{{ indexAxis:'y', responsive:true, plugins:{{ legend:{{display:false}} }},
            scales:{{ x:{{ ticks:{{ callback:v=>'R$ '+v.toLocaleString('pt-BR') }} }} }} }}
}});

new Chart(document.getElementById('c-estoque'), {{
  type:'bar',
  data:{{ labels:estoque.map(r=>r.sku), datasets:[{{ label:'Valor em estoque (R$)', data:estoque.map(r=>r.valor), backgroundColor:COL }}] }},
  options:{{ indexAxis:'y', responsive:true, plugins:{{ legend:{{display:false}} }},
            scales:{{ x:{{ ticks:{{ callback:v=>'R$ '+v.toLocaleString('pt-BR') }} }} }} }}
}});
</script>
</body>
</html>"""


def _esc(v):
    import html as H
    return H.escape(str(v))


def _brl(v):
    try:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return "R$ 0,00"


def main():
    parser = argparse.ArgumentParser(description="Gera analytics.html a partir dos caches")
    parser.add_argument("--out", default="analytics.html")
    args = parser.parse_args()

    caches = sorted(glob.glob(str(CACHE_DIR / "cache_*_*.json")))
    months, data = consolidate(caches)
    if not data:
        print("Nenhum cache encontrado em", CACHE_DIR)
        return 1

    fat_total = sum(m["total"] for m in data)
    ped_total = sum(m["pedidos"] for m in data)
    uni_total = sum(m["unidades"] for m in data)
    chaves = sorted(months)
    # agrega taxas de todos os caches
    all_settlements = [s for m in data for s in m["settlements"]]
    fees = fee_breakdown(all_settlements)
    # top SKUs (ASIN agregado)
    top_by_asin = defaultdict(
        lambda: {"sku": "", "titulo": "", "unidades": 0.0, "total": 0.0})
    for m in data:
        for p in m["products"]:
            key = p[1] or p[0]
            t = top_by_asin[key]
            t["sku"] = t["sku"] or p[0]
            t["titulo"] = t["titulo"] or p[2]
            t["unidades"] += p[3]
            t["total"] += p[5]
    top_10 = sorted(top_by_asin.values(), key=lambda v: -v["total"])[:10]
    top_skus_labels = [t["sku"] or (t["titulo"] or "")[:12] for t in top_10]
    # estoque: usa o inventario do cache mais recente
    # estoque: usa o inventario do cache mais recente que o tenha
    estoque_top = []
    inventario_m = None
    for key in reversed(chaves):
        mm = months[key]
        if mm["inv_rows"]:
            inventario_m = mm
            break
    if inventario_m is not None:
        estoque_top = inventory_summary(inventario_m["inv_rows"],
                                        inventario_m["inv_header"])[:15]

    melhor_key = max(chaves, key=lambda k: months[k]["total"]) if chaves else ""
    cards = {
        "faturamento": fat_total,
        "pedidos": ped_total,
        "unidades": uni_total,
        "meses_cobertos": len(data),
        "primeiro": month_label(chaves[0]) if chaves else "-",
        "ultimo": month_label(chaves[-1]) if chaves else "-",
        "ticket_medio": fat_total / ped_total if ped_total else 0.0,
        "melhor_mes": month_label(melhor_key) if melhor_key else "-",
        "melhor_valor": months[melhor_key]["total"] if melhor_key else 0.0,
    }
    if estoque_top:
        cards["estoque"] = sum(r["valor"] for r in estoque_top)
        cards["skus_ativos"] = len(estoque_top)

    payload = {
        "chave_meses": chaves,
        "meses": data,
        "cards": cards,
        "top_skus": top_skus_labels,
        "top_10": [[t["sku"], "", t["titulo"], "", "", t["total"]] for t in top_10],
        "fees": fees,
        "estoque_top": estoque_top,
    }
    html = render_html(payload)
    open(args.out, "w", encoding="utf-8").write(html)
    print(args.out)


if __name__ == "__main__":
    main()

def gerar_planilha_custos(caches):
    """Exporta nao-destrutivo: inventario (SKU/ASIN/Titulo/Unidades/Pedidos/Total) + coluna
    'Custo unitario (R$) -- preencher' VAZIA, em cache/planilha_inventario_custos.xlsx.
    Voce preenche os custos; o dashboard le de volta p/ margem."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    months, data = consolidate([str(c) for c in caches])
    from collections import defaultdict
    agr = defaultdict(lambda: {"sku":"","asin":"","titulo":"","unidades":0,"pedidos":0,"total":0.0})
    for m in data:
        for p in m.get("products", []):
            key = p[0].strip() or p[1].strip()
            d = agr[key]
            d["sku"] = d["sku"] or p[0]; d["asin"] = d["asin"] or p[1]
            d["titulo"] = d["titulo"] or p[2]; d["unidades"] += int(p[3])
            d["pedidos"] += int(p[4]); d["total"] += float(p[5])
    wb = Workbook(); ws = wb.active; ws.title = "Inventario"
    hdr = ["SKU","ASIN","Titulo","Unidades","Pedidos","Total (R$)","Custo unitario (R$) -- preencher"]
    ws.append(hdr)
    for c in ws[1]:
        c.font = Font(bold=True); c.fill = PatternFill("solid", fgColor="DDEEFF")
        c.alignment = Alignment(horizontal="center")
    for sku in sorted(agr, key=lambda s: -agr[s]["total"]):
        d = agr[sku]
        ws.append([d["sku"], d["asin"], d["titulo"], d["unidades"], d["pedidos"], round(d["total"],2), None])
    out = CACHE_DIR / "planilha_inventario_custos.xlsx"
    wb.save(out); return out
