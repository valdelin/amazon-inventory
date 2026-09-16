#!/usr/bin/env python3
"""Gera evolucao_de_vendas.xlsx a partir dos relatorios mensais.

Le todos os relatorio_mensal_<mes>_<ano>.xlsx da pasta, agrega por mes
e monta graficos de evolucao (colunas/combos, compativeis com OnlyOffice).

Uso:
    .venv/bin/python evolucao.py [--out evolucao_de_vendas.xlsx]
"""
import argparse
import glob
import re

import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font

MONTH_PT = {1: "jan", 2: "fev", 3: "mar", 4: "abr", 5: "mai", 6: "jun",
            7: "jul", 8: "ago", 9: "set", 10: "out", 11: "nov", 12: "dez"}
TITLE_RE = re.compile(r"(\d{2})/(\d{4})")


def read_monthly_totals(report_dir="."):
    """Retorna lista ordenada por data de (label, pedidos, unidades, total)."""
    rows = []
    for f in glob.glob(f"{report_dir}/relatorio_mensal_*.xlsx"):
        wb = openpyxl.load_workbook(f)
        ws = wb["Vendas Mensais"]
        title = ws.cell(row=1, column=1).value or ""
        vals = {r[0]: r[1]
                for r in ws.iter_rows(min_row=3, max_row=5, values_only=True)
                if r[0] in ("Pedidos (mês)", "Unidades (mês)", "Total (mês)")}
        m = TITLE_RE.search(title)
        if not m:
            continue
        month, year = int(m.group(1)), int(m.group(2))
        label = f"{MONTH_PT[month]}/{year}"
        rows.append((year * 12 + month, label,
                     vals.get("Pedidos (mês)") or 0,
                     vals.get("Unidades (mês)") or 0,
                     float(vals.get("Total (mês)") or 0)))
    rows.sort(key=lambda r: r[0])
    return [(l, p, u, t) for _, l, p, u, t in rows]


def dash_inside(bar, name):
    """Remove borda interna das barras (ferramentas antigas ignoram)."""
    for s in bar.series:
        if s.graphicalProperties.line.color is not None:
            s.graphicalProperties.line.color.value = None


def build_workbook(rows, out_path):
    out = openpyxl.Workbook()
    ws = out.active
    ws.title = "Evolucao de Vendas"
    ws.append(["Mes", "Pedidos", "Unidades", "Total (R$)", "Var %"])

    prev = None
    for label, pedidos, unidades, total in rows:
        var = None
        if prev:
            var = round((total - prev) / prev * 100, 1)
        ws.append([label, pedidos, unidades, round(total, 2), var])
        prev = total

    last = len(rows) + 1
    totals_r = last + 1
    ws.cell(row=totals_r, column=1, value="TOTAL")
    ws.cell(row=totals_r, column=2, value=sum(r[1] for r in rows))
    ws.cell(row=totals_r, column=3, value=sum(r[2] for r in rows))
    ws.cell(row=totals_r, column=4, value=sum(r[3] for r in rows))

    for c in ws[1]:
        c.font = Font(bold=True)
    for col, w in {"A": 9, "B": 9, "C": 10, "D": 13, "E": 9}.items():
        ws.column_dimensions[col].width = w
    for col in "BCDE":
        ws[f"{col}{totals_r}"].font = Font(bold=True)
    for r in range(2, totals_r + 1):
        ws.cell(row=r, column=4).number_format = '#,##0.00'
        cell = ws.cell(row=r, column=5)
        if cell.value is not None:
            cell.number_format = '0.0"%"'
    ws.freeze_panes = "A2"

    cats = Reference(ws, min_col=1, min_row=2, max_row=last)

    # Grafico 1: Total (R$) em colunas agrupadas
    bar_total = BarChart()
    bar_total.type = "col"
    bar_total.grouping = "clustered"
    bar_total.title = "Evolucao das Vendas (R$)"
    bar_total.y_axis.title = "Total (R$)"
    bar_total.x_axis.title = "Mes"
    bar_total.height = 11
    bar_total.width = 23
    dash_inside(bar_total, "total")
    bar_total.add_data(Reference(ws, min_col=4, min_row=1, max_row=last),
                       titles_from_data=True)
    bar_total.set_categories(cats)
    bar_total.legend = None
    ws.add_chart(bar_total, "F2")

    # Grafico 2: Pedidos (colunas) + Unidades (linha) - combo
    bar_ped = BarChart()
    bar_ped.type = "col"
    bar_ped.grouping = "stacked"
    bar_ped.title = "Pedidos e Unidades por Mes"
    bar_ped.y_axis.title = "Quantidade"
    bar_ped.x_axis.title = "Mes"
    bar_ped.height = 11
    bar_ped.width = 23
    dash_inside(bar_ped, "ped")
    bar_ped.add_data(Reference(ws, min_col=2, min_row=1, max_row=last),
                     titles_from_data=True)
    bar_ped.set_categories(cats)
    line_uni = LineChart()
    line_uni.grouping = "stacked"
    line_uni.add_data(Reference(ws, min_col=3, min_row=1, max_row=last),
                      titles_from_data=True)
    line_uni.set_categories(cats)
    line_uni.series[0].smooth = False
    line_uni.series[0].graphicalProperties.line.solidFill = "C0504D"
    bar_ped += line_uni
    ws.add_chart(bar_ped, "F18")

    # Aba: Faturamento Acumulado
    ws2 = out.create_sheet("Faturamento Acumulado")
    ws2.append(["Mes", "Total (R$)", "Acumulado (R$)"])
    cum = 0
    for label, pedidos, unidades, total in rows:
        cum += total
        ws2.append([label, round(total, 2), round(cum, 2)])
    row_t = len(rows) + 2
    ws2.cell(row=row_t, column=1, value="TOTAL")
    ws2.cell(row=row_t, column=2, value=round(cum, 2))
    ws2.cell(row=row_t, column=3, value=round(cum, 2))
    for c in ws2[1]:
        c.font = Font(bold=True)
    for col, w in {"A": 9, "B": 13, "C": 16}.items():
        ws2.column_dimensions[col].width = w
    for r in range(2, row_t + 1):
        ws2.cell(row=r, column=2).number_format = '#,##0.00'
        ws2.cell(row=r, column=3).number_format = '#,##0.00'
    ws2[f"C{row_t}"].font = Font(bold=True)

    # Grafico 3: acumulado em linha reta
    line_cum = LineChart()
    line_cum.title = "Faturamento Acumulado (R$)"
    line_cum.y_axis.title = "Acumulado (R$)"
    line_cum.height = 11
    line_cum.width = 24
    line_cum.add_data(Reference(ws2, min_col=3, min_row=1, max_row=row_t - 1),
                      titles_from_data=True)
    line_cum.set_categories(Reference(ws2, min_col=1, min_row=2,
                                       max_row=row_t - 1))
    line_cum.legend = None
    line_cum.series[0].smooth = False
    line_cum.series[0].graphicalProperties.line.solidFill = "2F5597"
    ws2.add_chart(line_cum, "E2")

    out.save(out_path)
    return len(rows), sum(r[1] for r in rows), sum(r[2] for r in rows), sum(r[3] for r in rows)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default="evolucao_de_vendas.xlsx",
                        help="arquivo de saida")
    args = parser.parse_args()

    rows = read_monthly_totals()
    if not rows:
        raise SystemExit("Nenhum relatorio_mensal_*.xlsx encontrado.")
    n, p, u, t = build_workbook(rows, args.out)
    print(f"Ok -> {args.out} ({n} meses)")
    print(f"Pedidos: {p} | Unidades: {u} | Total: R$ {t:,.2f}")


if __name__ == "__main__":
    main()