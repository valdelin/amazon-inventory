#!/usr/bin/env python3
"""Coleta o inventário atual do Seller Central via SP-API e grava em XLSX.

Configuração pelas variáveis de ambiente (ou arquivo .env na mesma pasta):
  LWA_APP_ID            Client ID LWA
  LWA_CLIENT_SECRET     Client Secret LWA
  SP_API_REFRESH_TOKEN  Refresh token do vendedor autorizado
  SP_API_PRIVATE_KEY    Caminho do PEM com a chave privada EC do cadastro
  SP_API_DEFAULT_MARKETPLACE  (opcional, default: BR)

Uso:
  python inventario.py [--out saida.xlsx]
                       [--report afn myi reservado]
                       [--report-type GET_...] (avançado)
"""

import argparse
import base64
import csv
import json
import logging
import os
import sys
import time
import zlib
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

import httpx

from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import ec
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from sp_api.api import DataKiosk, Orders, Reports
from sp_api.base import Marketplaces

log = logging.getLogger("inventario")

REPORT_TYPE_DEFAULT = "GET_MERCHANT_LISTINGS_ALL_DATA"
REPORT_CHOICES = {
    "listagens": "GET_MERCHANT_LISTINGS_ALL_DATA",
    "afn": "GET_AFN_INVENTORY_DATA",
    "myi": "GET_FBA_MYI_UNSUPPRESSED_INVENTORY_DATA",
    "reservado": "GET_RESERVED_INVENTORY_DATA",
    "saude": "GET_FBA_INVENTORY_PLANNING_DATA",
    "pedidos": "GET_FLAT_FILE_ALL_ORDERS_DATA_BY_ORDER_DATE_GENERAL",
    "pedidos-update": "GET_FLAT_FILE_ALL_ORDERS_DATA_BY_LAST_UPDATE_GENERAL",
    "arquivado": "GET_FLAT_FILE_ARCHIVED_ORDERS_DATA_BY_ORDER_DATE",
}
SHEET_TITLES = {
    "GET_MERCHANT_LISTINGS_ALL_DATA": "Listagens",
    "GET_AFN_INVENTORY_DATA": "AFN (disponivel)",
    "GET_FBA_MYI_UNSUPPRESSED_INVENTORY_DATA": "Manage Inv (MYI)",
    "GET_RESERVED_INVENTORY_DATA": "Reservado",
    "GET_FBA_INVENTORY_PLANNING_DATA": "Saude do inv.",
    "GET_FLAT_FILE_ALL_ORDERS_DATA_BY_ORDER_DATE_GENERAL": "Pedidos (data)",
    "GET_FLAT_FILE_ALL_ORDERS_DATA_BY_LAST_UPDATE_GENERAL": "Pedidos (update)",
    "GET_FLAT_FILE_ARCHIVED_ORDERS_DATA_BY_ORDER_DATE": "Pedidos arquivados",
}
DATE_RANGE_REPORTS = {
    "GET_FLAT_FILE_ALL_ORDERS_DATA_BY_ORDER_DATE_GENERAL",
    "GET_FLAT_FILE_ALL_ORDERS_DATA_BY_LAST_UPDATE_GENERAL",
    "GET_FLAT_FILE_ARCHIVED_ORDERS_DATA_BY_ORDER_DATE",
}
DATA_KIOSK_REPORT_TYPES = {
    "GET_FBA_INVENTORY_DATA",
}
ORDER_TYPE_MONTHLY = "GET_FLAT_FILE_ALL_ORDERS_DATA_BY_ORDER_DATE_GENERAL"
ORDER_TYPE_ARCHIVED = "GET_FLAT_FILE_ARCHIVED_ORDERS_DATA_BY_ORDER_DATE"
SETTLEMENT_TYPE = "GET_V2_SETTLEMENT_REPORT_DATA_FLAT_FILE_V2"
SETTLEMENT_FEE_LABELS = {
    "Commission": "Taxa Comissão (referral)",
    "FBAPerUnitFulfillmentFee": "Taxa FBA (por unidade)",
    "FBAWeightBasedFee": "Taxa FBA (peso)",
    "FBAShippingFee": "Taxa FBA (frete)",
    "Flexible Customer Financing fee": "Taxa Financing",
    "ShippingChargeback": "Chargeback frete",
    "RefundCommission": "Comissão devolvida",
    "ClosingFee": "Taxa de fechamento",
    "Subscription Fee": "Assinatura",
}
MARKETPLACE_ID_BR = "A2Q3Y263D00KWC"
HISTORICO_HEADER = [
    "amazon-order-id", "merchant-order-id", "purchase-date", "order-status",
    "fulfillment-channel", "sales-channel", "ship-service-level", "sku", "asin",
    "product-name", "item-status", "quantity", "currency", "item-price",
    "item-tax", "shipping-price", "shipping-tax", "gift-wrap-price",
    "gift-wrap-tax", "item-promotion-discount", "ship-promotion-discount",
    "payment-method-details", "order-item-id", "ship-state",
]
HISTORICO_LIMIT_PAGES = 600
CACHE_DIR = Path("cache")
CACHE_EXT = ".json"
FBA_REPORT_TYPES = {
    "GET_AFN_INVENTORY_DATA",
    "GET_FBA_MYI_UNSUPPRESSED_INVENTORY_DATA",
    "GET_FBA_MYI_ALL_INVENTORY_DATA",
    "GET_RESERVED_INVENTORY_DATA",
    "GET_FBA_INVENTORY_PLANNING_DATA",
}
END_REPORT_MARKER = "EndOfReport"


def load_dotenv(path=".env"):
    if not os.path.isfile(path):
        return
    with open(path, encoding="utf-8") as fh:
        for raw in fh:
            line = raw.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key:
                os.environ.setdefault(key, value)


def require_env(name):
    value = os.environ.get(name)
    if not value:
        raise RuntimeError(f"Falta a variável de ambiente {name}")
    return value


def create_inventory_report(report_type, data_start_time=None, data_end_time=None):
    marketplace_name = os.environ.get("SP_API_DEFAULT_MARKETPLACE", "BR")
    try:
        marketplace = getattr(Marketplaces, marketplace_name)
    except AttributeError:
        raise RuntimeError(f"Marketplace inválido: {marketplace_name}")
    reports = Reports(marketplace=marketplace)
    kwargs = {"reportType": report_type}
    if data_start_time:
        kwargs["dataStartTime"] = data_start_time
    if data_end_time:
        kwargs["dataEndTime"] = data_end_time
    res = reports.create_report(**kwargs)
    report_id = res.payload["reportId"]
    log.info("Relatório %s criado: id=%s (marketplace=%s)", report_type, report_id, marketplace_name)
    return reports, report_id


def wait_for_report(reports, report_id, timeout_min, report_type=None):
    deadline = time.monotonic() + timeout_min * 60
    while time.monotonic() < deadline:
        payload = reports.get_report(report_id).payload
        status = payload["processingStatus"]
        if status == "DONE":
            return payload.get("reportDocumentId")
        if status in {"CANCELLED", "FATAL"}:
            if status == "CANCELLED" and report_type == ORDER_TYPE_ARCHIVED:
                raise RuntimeError(
                    "Relatório de pedidos arquivados cancelado: nenhum pedido "
                    "elegível no período. A Amazon arquiva pedidos algum tempo "
                    "após a venda; lojas novas ou períodos sem vendas retornam vazio."
                )
            raise RuntimeError(f"Relatório falhou: status={status}")
        log.info("status=%s, aguardando...", status)
        time.sleep(10)
    raise TimeoutError(f"Relatório não terminou em {timeout_min} min "
                       f"(status atual: {status})")


def load_private_key(pem_path):
    with open(pem_path, "rb") as fh:
        return serialization.load_pem_private_key(fh.read(), password=None)


def derive_aes_key(private_key, peer_public_der_b64):
    peer_public_key = serialization.load_der_public_key(
        base64.b64decode(peer_public_der_b64)
    )
    secret = private_key.exchange(ec.ECDH(), peer_public_key)
    digest = hashes.Hash(hashes.SHA256())
    digest.update(secret)
    return digest.finalize()


def aes_decrypt(encryption_details, content, private_key):
    iv = base64.b64decode(encryption_details["initializationVector"])
    key = derive_aes_key(private_key, encryption_details["key"])
    decryptor = Cipher(algorithms.AES(key), modes.CBC(iv)).decryptor()
    plain = decryptor.update(content) + decryptor.finalize()
    pad = plain[-1]
    if pad < 1 or pad > 16 or plain[-pad:] != bytes([pad]) * pad:
        raise RuntimeError("Falha ao remover padding PKCS#7")
    return plain[:-pad]


def fetch_document(payload, private_key):
    with httpx.Client(timeout=120) as client:
        content = client.get(payload["url"]).content
    if payload.get("compressionAlgorithm"):
        content = zlib.decompress(content, 15 + 32)
    if payload.get("encryptionDetails"):
        content = aes_decrypt(payload["encryptionDetails"], content, private_key)
    try:
        return content.decode("utf-8-sig")
    except UnicodeDecodeError:
        return content.decode("latin-1").lstrip("\ufeff")


def create_data_kiosk_query(report_type):
    marketplace_name = os.environ.get("SP_API_DEFAULT_MARKETPLACE", "BR")
    marketplace = getattr(Marketplaces, marketplace_name)
    kiosk = DataKiosk(marketplace=marketplace)
    res = kiosk.create_query(query=report_type)
    query_id = res.payload["queryId"]
    log.info("Query Data Kiosk %s criada: id=%s (marketplace=%s)",
             report_type, query_id, marketplace_name)
    return kiosk, query_id


def wait_for_data_kiosk(kiosk, query_id, timeout_min):
    failure = {"CANCELLED", "CANCELLING", "FAILED", "EXPIRED"}
    deadline = time.monotonic() + timeout_min * 60
    while time.monotonic() < deadline:
        payload = kiosk.get_query(query_id).payload
        status = payload["processingStatus"]
        if status == "DONE":
            return payload.get("documentId")
        if status in failure:
            raise RuntimeError(f"Query falhou: status={status}")
        log.info("status=%s, aguardando...", status)
        time.sleep(10)
    raise TimeoutError(f"Query não terminou em {timeout_min} min (status: {status})")


def download_data_kiosk_document(kiosk, document_id, private_key):
    payload = kiosk.get_document(document_id).payload
    if not private_key:
        raise RuntimeError(
            "Documentos do Data Kiosk (GET_FBA_INVENTORY_DATA) são sempre "
            "criptografados e exigem a chave privada EC. Defina "
            "SP_API_PRIVATE_KEY=/caminho/do/pem no .env."
        )
    text = fetch_document(payload, private_key)
    log.info("Documento baixado (%d bytes)", len(text))
    return text


def download_report(reports, document_id, private_key=None):
    payload = reports.get_report_document(document_id).payload
    if payload.get("encryptionDetails") and not private_key:
        raise RuntimeError("Documento criptografado exige a chave privada "
                           "(defina SP_API_PRIVATE_KEY no .env)")
    text = fetch_document(payload, private_key)
    log.info("Documento baixado (%d bytes)", len(text))
    return text


def fetch_report_text(report_type, private_key, timeout_min,
                      data_start_time=None, data_end_time=None):
    if report_type in DATA_KIOSK_REPORT_TYPES:
        kiosk, query_id = create_data_kiosk_query(report_type)
        try:
            document_id = wait_for_data_kiosk(kiosk, query_id, timeout_min)
            return download_data_kiosk_document(kiosk, document_id, private_key)
        finally:
            kiosk.close()
    reports, report_id = create_inventory_report(
        report_type, data_start_time, data_end_time
    )
    try:
        document_id = wait_for_report(reports, report_id, timeout_min, report_type)
        return download_report(reports, document_id, private_key)
    finally:
        reports.close()


def detect_delimiter(header_line):
    return max(["\t", ",", ";", "|"], key=header_line.count)


def parse_flat_file(text):
    lines = [ln for ln in text.splitlines() if ln.strip()]
    if not lines:
        raise RuntimeError("Relatório vazio")
    delimiter = detect_delimiter(lines[0])
    rows = list(csv.reader(lines, delimiter=delimiter))
    header = rows[0]
    data = [r for r in rows[1:]
            if r
            and any(c.strip() for c in r)
            and END_REPORT_MARKER not in {c.strip() for c in r}]
    log.info("Cabeçalho com %d colunas, %d linhas de dados", len(header), len(data))
    return header, data


def _cell(row, idx):
    if idx is None or idx >= len(row):
        return ""
    return row[idx]


def _to_float(value):
    try:
        return float(str(value).strip() or 0)
    except ValueError:
        return 0.0


def _to_int(value):
    try:
        return int(float(str(value).strip() or 0))
    except ValueError:
        return 0


def _parse_money(value):
    """Converte valor monetário, inclusive formato BR ('1.234,56' / '-20,05')."""
    text = str(value).strip()
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        pass
    try:
        return float(text.replace(".", "").replace(",", "."))
    except ValueError:
        return 0.0


def summarise_orders(header, rows):
    """Agrega pedidos: totais do mês, por dia e por produto (ASIN). Exclui cancelados."""
    idx = {name: i for i, name in enumerate(header)}
    cancelled = {"cancelled", "cancelorder", "invalid"}
    idx_status = idx.get("item-status")
    daily = defaultdict(lambda: {"pedidos": 0, "unidades": 0, "total": 0.0})
    products = defaultdict(lambda: {"sku": "", "titulo": "",
                                    "unidades": 0, "pedidos": 0, "total": 0.0})
    for row in rows:
        if (_cell(row, idx_status) or "").strip().lower() in cancelled:
            continue
        qty = _to_int(_cell(row, idx.get("quantity")))
        total_linha = qty * _to_float(_cell(row, idx.get("item-price"))) \
            - _to_float(_cell(row, idx.get("item-promotion-discount")))

        dia = (_cell(row, idx.get("purchase-date")) or "")[:10] or "(sem data)"
        daily[dia]["pedidos"] += 1
        daily[dia]["unidades"] += qty
        daily[dia]["total"] += total_linha

        asin = _cell(row, idx.get("asin")) or "(sem ASIN)"
        produto = products[asin]
        if not produto["sku"]:
            produto["sku"] = _cell(row, idx.get("sku")) or ""
        if not produto["titulo"]:
            produto["titulo"] = _cell(row, idx.get("product-name")) or ""
        produto["pedidos"] += 1
        produto["unidades"] += qty
        produto["total"] += total_linha

    daily_rows = [
        [dia, d["pedidos"], d["unidades"], round(d["total"], 2)]
        for dia, d in sorted(daily.items())
    ]
    product_rows = [
        [p["sku"], asin, p["titulo"], p["unidades"], p["pedidos"], round(p["total"], 2)]
        for asin, p in sorted(products.items(), key=lambda kv: kv[1]["total"],
                              reverse=True)
    ]
    summary = {
        "pedidos": sum(d["pedidos"] for d in daily.values()),
        "unidades": sum(d["unidades"] for d in daily.values()),
        "total": round(sum(d["total"] for d in daily.values()), 2),
        "daily": daily_rows,
        "products": product_rows,
    }
    log.info("Vendas resumidas: %d pedidos, %d unidades, R$ %.2f",
             summary["pedidos"], summary["unidades"], summary["total"])
    return summary


def build_sales_rows(header, rows):
    """Agrupa vendas por pedido: id, data, valor pago, estado, pagamento. Exclui cancelados."""
    idx = {name: i for i, name in enumerate(header)}
    cancelled = {"cancelled", "cancelorder", "invalid"}
    idx_status = idx.get("item-status")
    groups = defaultdict(lambda: {"data": None, "valor": 0.0, "estado": "",
                                  "pagamento": "", "unidades": 0, "itens": defaultdict(int)})
    for row in rows:
        order_id = _cell(row, idx.get("amazon-order-id"))
        if not order_id:
            continue
        if (_cell(row, idx_status) or "").strip().lower() in cancelled:
            continue
        g = groups[order_id]
        qty = _to_int(_cell(row, idx.get("quantity")))
        price = _to_float(_cell(row, idx.get("item-price")))
        g["valor"] += (
            qty * price
            + _to_float(_cell(row, idx.get("item-tax")))
            + _to_float(_cell(row, idx.get("shipping-price")))
            + _to_float(_cell(row, idx.get("shipping-tax")))
            + _to_float(_cell(row, idx.get("gift-wrap-price")))
            + _to_float(_cell(row, idx.get("gift-wrap-tax")))
            - _to_float(_cell(row, idx.get("item-promotion-discount")))
            - _to_float(_cell(row, idx.get("ship-promotion-discount")))
        )
        g["unidades"] += qty
        sku = _cell(row, idx.get("sku")) or ""
        g["itens"][sku] += qty
        raw = (_cell(row, idx.get("purchase-date")) or "").strip()
        try:
            ts = datetime.fromisoformat(raw.replace("Z", "+00:00"))
        except ValueError:
            ts = None
        if ts is not None:
            ts = ts.replace(tzinfo=None)
        if ts is not None and (g["data"] is None or ts < g["data"]):
            g["data"] = ts
        if not g["estado"]:
            g["estado"] = _cell(row, idx.get("ship-state")) or ""
        if not g["pagamento"]:
            g["pagamento"] = _cell(row, idx.get("payment-method-details")) or ""

    sales_rows = []
    for order_id, g in groups.items():
        itens = ", ".join(f"{sku} x{qty}" if sku else f"{qty} un."
                          for sku, qty in sorted(g["itens"].items()))
        sales_rows.append([
            order_id, g["data"], round(g["valor"], 2),
            g["estado"], g["pagamento"], g["unidades"], itens,
        ])
    sales_rows.sort(key=lambda r: r[1] or datetime.min)
    log.info("Vendas detalhadas: %d pedidos", len(sales_rows))
    return sales_rows


def parse_settlement_fees(set_texts):
    """Reduz os settlements a {order-id: {taxa: valor}}.

    Considera só linhas com pedido + item + SKU (deixando de fora taxas de
    armazenagem, assinatura, publicidade etc.) e deduplica por linha, pois um
    mesmo período de settlement pode vir em vários relatórios complementares.
    Valores negativos = custo pago; positivos = devolução/ajuste.
    """
    fees_by_order = {}
    seen = set()
    for text in set_texts or []:
        header, rows = parse_flat_file(text)
        idx = {name: i for i, name in enumerate(header)}
        oid_i = idx.get("order-id")
        item_i = idx.get("order-item-code")
        desc_i = idx.get("amount-description")
        amt_i = idx.get("amount")
        sku_i = idx.get("sku")
        if None in (oid_i, item_i, desc_i, amt_i, sku_i):
            continue
        for row in rows:
            oid = _cell(row, oid_i).strip()
            item = _cell(row, item_i).strip()
            sku = _cell(row, sku_i).strip()
            desc = _cell(row, desc_i).strip()
            raw_amt = _cell(row, amt_i).strip()
            if not (oid and item and sku and desc and raw_amt):
                continue
            sig = (oid, item, desc, raw_amt)
            if sig in seen:
                continue
            seen.add(sig)
            valor = _parse_money(raw_amt)
            if valor == 0:
                continue
            fees_by_order.setdefault(oid, defaultdict(float))[desc] += valor
    fees = {oid: dict(d) for oid, d in fees_by_order.items()}
    n_pedidos = len(fees)
    n_tipos = len({desc for f in fees.values() for desc in f})
    log.info("Taxas do settlement: %d pedidos com taxa (%d tipos)",
             n_pedidos, n_tipos)
    return fees


def append_sales_sheet(wb, period_text, header, rows, fees_by_order=None):
    fees_by_order = fees_by_order or {}
    fee_descs = sorted({d for f in fees_by_order.values() for d in f},
                       key=lambda d: SETTLEMENT_FEE_LABELS.get(d, d))
    n_fee = len(fee_descs)
    ncols = 7 + n_fee + (1 if n_fee else 0)
    ws = wb.create_sheet(title="Todas as Vendas")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    cell = ws.cell(row=1, column=1,
                   value=f"Todas as vendas realizadas — {period_text}")
    cell.font = Font(bold=True, size=12)
    sales_header = ["ID do Pedido", "Data da Compra", "Valor Pago",
                    "Estado", "Tipo de Pagamento", "Unidades", "Itens"]
    fee_labels = [SETTLEMENT_FEE_LABELS.get(d, d) for d in fee_descs]
    if n_fee:
        fee_labels.append("Total Taxas Amazon")
    for col, text in enumerate(sales_header + fee_labels, start=1):
        ws.cell(row=2, column=col, value=text).font = Font(bold=True)
    for order_id, data, valor, estado, pagamento, unidades, itens in build_sales_rows(header, rows):
        row_values = [order_id, data, valor, estado, pagamento, unidades, itens]
        fees = fees_by_order.get(order_id, {})
        for desc in fee_descs:
            row_values.append(round(fees.get(desc, 0.0), 2))
        if n_fee:
            row_values.append(round(sum(fees.get(d, 0.0) for d in fee_descs), 2))
        ws.append(row_values)
        ws.cell(row=ws.max_row, column=3).number_format = "#,##0.00"
        ws.cell(row=ws.max_row, column=2).number_format = "DD/MM/YYYY HH:MM"
        for col_idx in range(7 + 1, ncols + 1):
            ws.cell(row=ws.max_row, column=col_idx).number_format = "#,##0.00"
        if n_fee:
            ws.cell(row=ws.max_row, column=ncols).font = Font(bold=True)
    for col_idx in range(1, ncols + 1):
        width = 14
        for cells in ws.iter_rows(min_col=col_idx, max_col=col_idx):
            for c in cells:
                if c.value is not None and len(str(c.value)) > width:
                    width = min(len(str(c.value)), 45)
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.freeze_panes = "A3"
    return ws


def append_summary_sheet(wb, period_text, summary):
    ws = wb.create_sheet(title="Vendas Mensais")
    ncols = 6
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    cell = ws.cell(row=1, column=1, value=f"Relatório mensal de vendas — {period_text}")
    cell.font = Font(bold=True, size=12)

    ws.cell(row=3, column=1, value="Pedidos (mês)").font = Font(bold=True)
    ws.cell(row=3, column=2, value=summary["pedidos"])
    ws.cell(row=4, column=1, value="Unidades (mês)").font = Font(bold=True)
    ws.cell(row=4, column=2, value=summary["unidades"])
    ws.cell(row=5, column=1, value="Total (mês)").font = Font(bold=True)
    ws.cell(row=5, column=2, value=summary["total"]).number_format = "#,##0.00"

    row = 7
    ws.cell(row=row, column=1, value="Resumo diário").font = Font(bold=True, size=11)
    row += 1
    daily_header = ["Data", "Pedidos", "Unidades", "Total"]
    for col, text in enumerate(daily_header, start=1):
        ws.cell(row=row, column=col, value=text).font = Font(bold=True)
    for data, pedidos, unidades, total in summary["daily"]:
        row += 1
        ws.cell(row=row, column=1, value=data)
        ws.cell(row=row, column=2, value=pedidos)
        ws.cell(row=row, column=3, value=unidades)
        ws.cell(row=row, column=4, value=total).number_format = "#,##0.00"

    row += 2
    ws.cell(row=row, column=1, value="Por produto (ASIN)").font = Font(bold=True, size=11)
    row += 1
    prod_header = ["ASIN", "SKU", "Título", "Unidades", "Pedidos", "Total"]
    for col, text in enumerate(prod_header, start=1):
        ws.cell(row=row, column=col, value=text).font = Font(bold=True)
    for sku, asin, titulo, unidades, pedidos, total in summary["products"]:
        row += 1
        ws.cell(row=row, column=1, value=asin)
        ws.cell(row=row, column=2, value=sku)
        ws.cell(row=row, column=3, value=titulo)
        ws.cell(row=row, column=4, value=unidades)
        ws.cell(row=row, column=5, value=pedidos)
        ws.cell(row=row, column=6, value=total).number_format = "#,##0.00"

    for col_idx in range(1, ncols + 1):
        width = 12
        for cells in ws.iter_rows(min_col=col_idx, max_col=col_idx):
            for c in cells:
                if c.value is not None and len(str(c.value)) > width:
                    width = min(len(str(c.value)), 45)
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.freeze_panes = "A6"
    return ws


INSTRUCTIONS_LINE = [
    ("COMO GERAR O RELATÓRIO MENSAL", ""),
    ("Pasta", "<USERHOME>/Work/amazon-inventory"),
    ("Comando", ".venv/bin/python inventario.py --mensal"),
    ("Atualizar sem rede (cache)",
     ".venv/bin/python inventario.py --update  — reconstrói a planilha dos dados salvos no último --mensal"),
    ("Mesas antigas (além de 30 dias)",
     ".venv/bin/python inventario.py --historico --start AAAA-MM-DD --end AAAA-MM-DD  (via Orders API; também usa cache próprio)"),
    ("Janela padrão", "1º dia do mês atual até hoje (mês/ano no nome do arquivo)"),
    ("Arquivo gerado", "relatorio_mensal_<mês>_<ano>.xlsx"),
    ("Janela personalizada",
     ".venv/bin/python inventario.py --mensal --start AAAA-MM-DD --end AAAA-MM-DD"),
    ("Salvar em outro arquivo", ".venv/bin/python inventario.py --mensal --out caminho/arquivo.xlsx"),
    ("", ""),
    ("OUTROS RELATÓRIOS (--report)", ""),
    ("  listagens", "inventário por SKU: preço, quantidade, status (padrão)"),
    ("  pedidos", "vendas do mês, uma linha por item comprado"),
    ("  myi / reservado", "estoque FBA gerenciado / reservado (papel Pricing)"),
    ("  arquivado", "pedidos antigos/arquivados (até 2 anos); cancela se não houver dados elegíveis"),
    ("  afn / saude", "exigem o papel 'Amazon Fulfillment' (erro 403)"),
    ("Combinar relatórios", ".venv/bin/python inventario.py --report listagens pedidos"),
    ("Janela para 'pedidos'", "--start e --end (YYYY-MM-DD) definem o período"),
    ("", ""),
    ("SIGNIFICADO DAS ABAS", ""),
    ("  Inventario", "listagens ativas do vendedor (SKU, ASIN, preço, estoque, status)"),
    ("  Vendas Mensais", "agregado por SKU: unidades, pedidos, bruto, desconto, líquido + linha TOTAL"),
    ("  Todas as Vendas", "detalhe por pedido + colunas de taxas Amazon (do settlement) e Total Taxas"),
    ("  Instruções", "este guia de uso do script"),
    ("", ""),
    ("TAXAS AMAZON (SETTLEMENT)", ""),
    ("Fonte", "relatório de settlement da Amazon (GET_V2_SETTLEMENT_REPORT_DATA_FLAT_FILE_V2)"),
    ("Significado", "negativo = custo pago pela venda; positivo = devolução/ajuste (ex.: comissão devolvida)"),
    ("Junção", "por ID do pedido — pedidos com 2 SKUs somam as taxas dos itens na linha"),
    ("Limite", "settlements ficam retidos ~90 dias na SP-API; meses antigos podem sair sem colunas de taxa"),
    ("Fora do escopo", "armazenagem, assinatura e publicidade não são taxas por venda e ficam de fora"),
    ("", ""),
    ("PRÓXIMA FUNCIONALIDADE — BAIXAR NOTAS FISCAIS (NF-e)", ""),
    ("Status", "em planejamento — ainda não implementada"),
    ("O que faz", "baixa automática das notas fiscais (NF-e em XML) dos pedidos FBA do período"),
    ("Como", "via Invoices API (InvoicesV20240619, só Brasil): buscar → exportar → baixar ZIP com XMLs"),
    ("Bloqueio", "exige o papel 'Tax Invoicing (Restricted)' no app (avaliação da Amazon + nova autorização)"),
    ("Usável para", "apenas pedidos FBA; MFN/frete próprio não tem NF-e pela SP-API"),
    ("", ""),
    ("OBSERVAÇÕES", ""),
    ("  Cache", "a pasta 'cache/' guarda os dados brutos do último --mensal (usada pelo --update)"),
    ("  Pedidos cancelados", "são excluídos do resumo de vendas"),
    ("  Pedidos com 2 SKUs", "contam 1 vez no total de pedidos"),
    ("  Moeda", "valores em BRL (marketplace BR)"),
]


def append_instructions_sheet(wb):
    ws = wb.create_sheet("Instruções")
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 78
    ws.merge_cells("A1:B1")
    ws["A1"] = "INSTRUÇÕES DE USO — Relatório Mensal Amazon (SP-API)"
    ws["A1"].font = Font(bold=True, size=13)
    row = 3
    for cat, cmd in INSTRUCTIONS_LINE:
        if not cmd:
            ws.merge_cells(start_row=row, start_column=1,
                           end_row=row, end_column=2)
            header = ws.cell(row=row, column=1, value=cat)
            header.font = Font(bold=True, size=11)
        else:
            a = ws.cell(row=row, column=1, value=cat)
            a.font = Font(bold=True)
            a.alignment = Alignment(vertical="top")
            b = ws.cell(row=row, column=2, value=cmd)
            b.alignment = Alignment(wrap_text=True, vertical="top")
        row += 1
    ws.freeze_panes = "A3"
    return ws


def append_sheet(wb, title, header, rows):
    ws = wb.active
    if ws.max_row > 1 or ws.max_column > 1 or ws.title != "Sheet":
        ws = wb.create_sheet(title=title)
    ws.title = title
    ws.append(header)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in rows:
        ws.append(row[:len(header)])
    ws.freeze_panes = "A2"
    for col_idx in range(1, len(header) + 1):
        width = 12
        for row in ws.iter_rows(min_col=col_idx, max_col=col_idx):
            for cell in row:
                if cell.value is not None and len(str(cell.value)) > width:
                    width = min(len(str(cell.value)), 40)
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def save_workbook(wb, out_path):
    wb.save(out_path)
    log.info("Planilha gravada: %s", out_path)


def resolve_window(start_arg, end_arg):
    def parse(value):
        return datetime.strptime(value, "%Y-%m-%d").date() if value else None
    try:
        start = parse(start_arg)
        end = parse(end_arg)
    except ValueError:
        raise RuntimeError("--start/--end devem ser YYYY-MM-DD")
    start = start or date.today().replace(day=1)
    end = end or date.today()
    if start > end:
        raise RuntimeError("--start é depois de --end")
    return start, end


def fetch_mensal_data(private_key, timeout_min, start, end):
    log.info("=== INVENTARIO ===")
    inv_text = fetch_report_text(REPORT_TYPE_DEFAULT, private_key, timeout_min)
    log.info("=== VENDAS (período %s a %s) ===", start, end)
    data_start = f"{start.isoformat()}T00:00:00Z"
    data_end = f"{end.isoformat()}T23:59:59Z"
    ord_text = fetch_report_text(ORDER_TYPE_MONTHLY, private_key, timeout_min,
                                 data_start, data_end)
    log.info("=== TAXAS (SETTLEMENT) ===")
    set_texts = fetch_settlement_texts(private_key)
    return inv_text, ord_text, set_texts


def fetch_settlement_texts(private_key):
    """Baixa os relatórios de settlement disponíveis (taxas por venda).

    A Amazon gera o settlement sozinha (~ a cada 2 semanas) e o relatório fica
    retido por ~90 dias; aqui apenas consultamos e baixamos o que existe.
    """
    marketplace_name = os.environ.get("SP_API_DEFAULT_MARKETPLACE", "BR")
    marketplace = getattr(Marketplaces, marketplace_name)
    reports = Reports(marketplace=marketplace)
    texts = []
    try:
        next_token = None
        while True:
            kwargs = {"reportTypes": [SETTLEMENT_TYPE], "pageSize": 100,
                      "processingStatuses": ["DONE"]}
            if next_token:
                kwargs = {"nextToken": next_token}
            resp = _api_get("getReports", reports.get_reports, **kwargs)
            payload = resp.payload or {}
            items = payload.get("reports") or []
            done = [r for r in items if r.get("reportDocumentId")]
            log.info("Settlement: %d documento(s) disponíveis", len(done))
            for rep in done:
                doc_payload = _api_get(
                    "getReportDocument", reports.get_report_document,
                    rep["reportDocumentId"]).payload
                texts.append(fetch_document(doc_payload, private_key))
                time.sleep(0.25)
            next_token = payload.get("NextToken") or payload.get("nextToken")
            if not next_token or not done:
                break
    finally:
        reports.close()
    if not texts:
        log.info("Nenhum settlement disponível (ou falta o papel Finance, "
                 "ou o período saiu da retenção de ~90 dias).")
    return texts


def cache_path(start, end):
    return CACHE_DIR / f"cache_mensal_{end.strftime('%Y%m')}{CACHE_EXT}"


def save_cache(start, end, inv_text, ord_text, set_texts):
    payload = {
        "start": start.isoformat(),
        "end": end.isoformat(),
        "inventario": inv_text,
        "pedidos": ord_text,
        "settlements": set_texts or [],
    }
    CACHE_DIR.mkdir(exist_ok=True)
    path = cache_path(start, end)
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    log.info("Cache salvo: %s", path)
    return path


def load_cache(start, end):
    path = cache_path(start, end)
    if not path.is_file():
        raise RuntimeError(
            f"Cache não encontrado: {path}\n"
            f"[Dica] rode '--mensal' primeiro para baixar os dados do mês."
        )
    payload = json.loads(path.read_text(encoding="utf-8"))
    cache_start = date.fromisoformat(payload["start"])
    cache_end = date.fromisoformat(payload["end"])
    log.info("Cache lido: %s a %s (%s)", cache_start, cache_end, path)
    return (cache_start, cache_end, payload["inventario"], payload["pedidos"],
            payload.get("settlements", []))


def build_mensal_wb(inv_text, ord_text, start, end, set_texts=None):
    header, rows = parse_flat_file(inv_text)
    oheader, orows = parse_flat_file(ord_text)
    return build_wb_from_rows(start, end, oheader, orows,
                              inv_header=header, inv_rows=rows,
                              settlements=set_texts)


def build_wb_from_rows(start, end, ord_header, ord_rows,
                       inv_header=None, inv_rows=None, settlements=None):
    wb = Workbook()
    if inv_header is not None:
        append_sheet(wb, "Inventario", inv_header, inv_rows)
    summary = summarise_orders(ord_header, ord_rows)
    period_text = f"{start.strftime('%d/%m/%Y')} a {end.strftime('%d/%m/%Y')}"
    append_summary_sheet(wb, period_text, summary)
    fees = parse_settlement_fees(settlements)
    append_sales_sheet(wb, period_text, ord_header, ord_rows,
                       fees_by_order=fees)
    append_instructions_sheet(wb)
    return wb


def _amount(money):
    if isinstance(money, dict):
        return money.get("Amount")
    return money if money is not None else ""


def _order_item_status(order_status, item_status):
    value = item_status or order_status or "Unknown"
    if str(value).strip().lower() in {"cancelled", "cancelorder", "invalid"}:
        return "Cancelled"
    return value


def _api_get(label, fn, *args, **kwargs):
    """Executa chamada à API com retry em erros de limite (QuotaExceeded/SlowDown)."""
    attempt = 0
    while True:
        try:
            return fn(*args, **kwargs)
        except Exception as exc:
            msg = str(exc)
            if any(k in msg for k in
                   ("QuotaExceeded", "SlowDown", "RateLimit", "Throttl", "TooMany")):
                attempt += 1
                if attempt > 12:
                    raise
                wait = min(60, 5 * 2 ** (attempt - 1))
                log.info("%s: limite da API (%s), aguardando %.0fs...",
                         label, msg[:80], wait)
                time.sleep(wait)
            else:
                raise


def fetch_historical_rows(start, end):
    """Baixa pedidos antigos via Orders API (getOrders + getOrderItems)."""
    orders = Orders()
    header = list(HISTORICO_HEADER)
    rows = []
    page = 0
    try:
        next_token = None
        while True:
            if next_token:
                payload = _api_get(
                    "getOrders", orders.get_orders, NextToken=next_token).payload
            else:
                payload = _api_get(
                    "getOrders", orders.get_orders,
                    MarketplaceIds=[MARKETPLACE_ID_BR],
                    CreatedAfter=f"{start.isoformat()}T00:00:00Z",
                    CreatedBefore=f"{end.isoformat()}T23:59:59Z",
                    MaxResultsPerPage=100,
                ).payload
            for order in payload.get("Orders", []):
                oid = order.get("AmazonOrderId")
                items_next = None
                while True:
                    if items_next:
                        items_page = _api_get(
                            "getOrderItems", orders.get_order_items,
                            oid, NextToken=items_next)
                    else:
                        items_page = _api_get(
                            "getOrderItems", orders.get_order_items, oid)
                    items = items_page.payload.get("OrderItems", [])
                    for item in items:
                        row = {name: "" for name in header}
                        row["amazon-order-id"] = oid
                        row["merchant-order-id"] = order.get("SellerOrderId") or ""
                        row["purchase-date"] = order.get("PurchaseDate") or ""
                        row["order-status"] = order.get("OrderStatus") or ""
                        row["fulfillment-channel"] = order.get("FulfillmentChannel") or ""
                        row["sales-channel"] = order.get("SalesChannel") or ""
                        row["ship-service-level"] = order.get("ShipServiceLevel") or ""
                        pay = order.get("PaymentMethodDetails")
                        row["payment-method-details"] = (
                            ", ".join(pay) if isinstance(pay, list) else (pay or ""))
                        row["sku"] = item.get("SellerSKU") or ""
                        row["asin"] = item.get("ASIN") or ""
                        row["product-name"] = item.get("Title") or ""
                        row["order-item-id"] = item.get("OrderItemId") or ""
                        row["quantity"] = item.get("QuantityOrdered") or 0
                        cur = (item.get("ItemPrice") or {}).get("CurrencyCode") or ""
                        row["currency"] = cur
                        row["item-price"] = _amount(item.get("ItemPrice"))
                        row["item-tax"] = _amount(item.get("ItemTax"))
                        row["shipping-price"] = _amount(item.get("ShippingPrice"))
                        row["shipping-tax"] = _amount(item.get("ShippingTax"))
                        row["gift-wrap-price"] = _amount(item.get("GiftWrapPrice"))
                        row["gift-wrap-tax"] = _amount(item.get("GiftWrapTax"))
                        row["item-promotion-discount"] = _amount(
                            item.get("ItemPromotionDiscount"))
                        row["ship-promotion-discount"] = _amount(
                            item.get("ShipPromotionDiscount"))
                        row["item-status"] = _order_item_status(
                            order.get("OrderStatus"), item.get("ItemStatus"))
                        rows.append([row[name] for name in header])
                    time.sleep(0.25)
                    items_next = (items_page.payload or {}).get("NextToken")
                    if not items_next:
                        break
            next_token = payload.get("NextToken")
            page += 1
            if next_token is None:
                break
            if page > HISTORICO_LIMIT_PAGES:
                raise RuntimeError("Paginação do Orders API excedeu o limite")
        log.info("Histórico via Orders API: %d linhas de itens (%d pedidos)",
                 len(rows), len({r[0] for r in rows}))
    finally:
        orders.close()
    return header, rows


def historico_cache_path(start, end):
    return CACHE_DIR / f"cache_historico_{end.strftime('%Y%m')}{CACHE_EXT}"


def save_historico_cache(start, end, header, rows, set_texts=None):
    payload = {"start": start.isoformat(), "end": end.isoformat(),
               "header": header, "rows": rows,
               "settlements": set_texts or []}
    CACHE_DIR.mkdir(exist_ok=True)
    path = historico_cache_path(start, end)
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    log.info("Cache histórico salvo: %s", path)
    return path


def load_historico_cache(start, end):
    path = historico_cache_path(start, end)
    if not path.is_file():
        raise RuntimeError(
            f"Cache histórico não encontrado: {path}\n"
            f"[Dica] rode '--historico --start ... --end ...' primeiro."
        )
    payload = json.loads(path.read_text(encoding="utf-8"))
    cache_start = date.fromisoformat(payload["start"])
    cache_end = date.fromisoformat(payload["end"])
    log.info("Cache histórico lido: %s a %s (%s)", cache_start, cache_end, path)
    return (cache_start, cache_end, payload["header"], payload["rows"],
            payload.get("settlements", []))


def mensal_default_output(end):
    MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho",
             "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
    return Path(f"relatorio_mensal_{MESES[end.month - 1]}_{end.year}.xlsx")


historico_default_output = mensal_default_output


def main():
    parser = argparse.ArgumentParser(description="Coleta o inventário da Amazon (SP-API) para XLSX")
    parser.add_argument("--out", help="Arquivo XLSX de saída")
    parser.add_argument("--mensal", action="store_true",
                        help="Gera relatório mensal (Inventario + Vendas Mensais)")
    parser.add_argument("--update", action="store_true",
                        help="Reconstrói a planilha a partir do cache, sem acessar a Amazon")
    parser.add_argument("--historico", action="store_true",
                        help="Baixa pedidos antigos (meses além de 30 dias) via Orders API")
    parser.add_argument("--report", nargs="+", choices=sorted(REPORT_CHOICES),
                        help="Relatórios a gerar (atalhos): %(choices)s")
    parser.add_argument("--report-type", default=None,
                        help="Report type SP-API bruto (avançado)")
    parser.add_argument("--start", help="Data inicial (YYYY-MM-DD) para relatórios de pedidos")
    parser.add_argument("--end", help="Data final (YYYY-MM-DD) para relatórios de pedidos")
    parser.add_argument("--timeout-min", type=int, default=30,
                        help="Tempo máximo de espera do relatório (min)")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(message)s",
    )

    load_dotenv()
    for name in ("LWA_APP_ID", "LWA_CLIENT_SECRET", "SP_API_REFRESH_TOKEN"):
        require_env(name)

    private_key = None
    if os.environ.get("SP_API_PRIVATE_KEY"):
        private_key = load_private_key(os.environ["SP_API_PRIVATE_KEY"])

    if args.historico:
        start, end = resolve_window(args.start, args.end)
        log.info("Janela de dados: %s a %s", start, end)
        hp = historico_cache_path(start, end)
        if hp.is_file():
            _, _, header, rows, set_texts = load_historico_cache(start, end)
            wb = build_wb_from_rows(start, end, header, rows,
                                    settlements=set_texts)
        else:
            header, rows = fetch_historical_rows(start, end)
            log.info("=== TAXAS (SETTLEMENT) ===")
            set_texts = fetch_settlement_texts(private_key)
            save_historico_cache(start, end, header, rows, set_texts)
            wb = build_wb_from_rows(start, end, header, rows,
                                    settlements=set_texts)
        out_path = args.out or historico_default_output(end)
        save_workbook(wb, out_path)
        print(out_path)
        return

    if args.mensal or args.update:
        start, end = resolve_window(args.start, args.end)
        log.info("Janela de dados: %s a %s", start, end)
        if args.update:
            _, _, inv_text, ord_text, set_texts = load_cache(start, end)
            wb = build_mensal_wb(inv_text, ord_text, start, end,
                                 set_texts=set_texts)
        else:
            inv_text, ord_text, set_texts = fetch_mensal_data(
                private_key, args.timeout_min, start, end
            )
            save_cache(start, end, inv_text, ord_text, set_texts)
            wb = build_mensal_wb(inv_text, ord_text, start, end,
                                 set_texts=set_texts)
        out_path = args.out or mensal_default_output(end)
        save_workbook(wb, out_path)
        print(out_path)
        return

    if args.report:
        report_types = [REPORT_CHOICES[name] for name in args.report]
    elif args.report_type:
        report_types = [args.report_type]
    else:
        report_types = [REPORT_TYPE_DEFAULT]
    needs_range = set(report_types) & DATE_RANGE_REPORTS

    data_start_time = data_end_time = None
    if needs_range:
        start, end = resolve_window(args.start, args.end)
        data_start_time = f"{start.isoformat()}T00:00:00Z"
        data_end_time = f"{end.isoformat()}T23:59:59Z"
        log.info("Janela de dados: %s a %s", start, end)

    wb = Workbook()
    for report_type in report_types:
        log.info("=== %s ===", report_type)
        try:
            text = fetch_report_text(
                report_type, private_key, args.timeout_min,
                data_start_time if report_type in DATE_RANGE_REPORTS else None,
                data_end_time if report_type in DATE_RANGE_REPORTS else None,
            )
        except Exception as exc:  # noqa: BLE001
            if report_type in FBA_REPORT_TYPES and "Unauthorized" in str(exc):
                raise RuntimeError(
                    f"{exc}\n"
                    f"[Dica] {report_type} exige o papel 'Amazon Fulfillment' no app "
                    f"(Seller Central > Integração > Suas aplicações > Gerenciar permissões). "
                    f"Depois de adicionar, reautorize o vendedor e rode de novo."
                ) from exc
            raise
        header, rows = parse_flat_file(text)
        append_sheet(wb, SHEET_TITLES.get(report_type, report_type), header, rows)

    if all(rt in FBA_REPORT_TYPES for rt in report_types):
        base_name = "inventario_fba"
    elif all(rt in DATE_RANGE_REPORTS for rt in report_types):
        base_name = "inventario_pedidos"
    else:
        base_name = "inventario_amazon"
    out_path = args.out or Path(f"{base_name}_%s.xlsx"
                                % datetime.now().strftime("%Y%m%d_%H%M"))
    save_workbook(wb, out_path)
    print(out_path)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        log.error("Erro: %s", exc)
        sys.exit(1)