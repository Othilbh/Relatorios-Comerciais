# pages/4_Metas_Semanais.py
import streamlit as st
from datetime import date, timedelta
import sys, os, tempfile
import pandas as pd
import io
import json
import math
import base64
import requests
sys.path.append(os.path.dirname(os.path.dirname(__file__)))

from utils.parser import extrair_vendas_por_vendedor, extrair_estoque_por_vendedor, VENDEDORES_ATIVOS
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.styles.numbers import FORMAT_NUMBER_COMMA_SEPARATED1
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint

st.set_page_config(page_title="Metas Semanais · OTHIL", page_icon="🎯", layout="wide")
st.title("🎯 Metas Semanais")
st.caption("Relatório diário por vendedor + acompanhamento de metas. Luca excluído automaticamente.")
st.divider()

PERCENTUAIS = {
    "Farley":    0.175,
    "Dora":      0.175,
    "Afanais":   0.250,
    "Roni":      0.250,
    "Reginaldo": 0.225,
    "Juliana":   0.075,
    "Claudia":   0.075,
    "Luciano":   0.075,
}

PRODUTOS_DEFAULT = [
    "Portuguesa Sabor 55/60", "Portuguesa 60/70", "Portuguesa 70/80",
    "Forelle", "Ercoline", "Pera Asiática",
    "Gala Santa Carol", "Gala Azaleia",
    "Fuji Expressa 180", "Fuji Azaleia", "Fuji Hiragami", "Fuji Suprema",
    "Thompson Vitace", "Thompson Seedless",
    "Uva Crimson", "Uva Isis", "Uva Jubilee", "Uva Itália",
    "Maçã Argentina", "Maçã Chilena", "Maçã Pink Lady",
    "Maçã Granny Smith", "Maçã Red Globe",
    "Mamão Havai", "Mamão Formoso",
    "Goiaba", "Melão Amarelo", "Melão Gaia", "Melão Cantaloupe",
    "Tangerina Cumbuca", "Tangerina Ponkan",
    "Ameixa", "Pêssego", "Nectarina",
    "Morango", "Mirtilo", "Abacaxi",
    "Manga Palmer", "Manga Tommy",
    "Laranja", "Limão Tahiti", "Limão Siciliano", "Tomate Roma",
]

MAPA_PRODUTO = {
    "Portuguesa Sabor 55/60": ["PORTUGUESA SABOR", "PORTUGUESA 55", "PORTUGUESA SAB", "02050032"],
    "Portuguesa 60/70": ["PORTUGUESA 60"],
    "Portuguesa 70/80": ["PORTUGUESA 70"],
    "Forelle": ["FORELLE", "020502701"],
    "Ercoline": ["ERCOLINE"],
    "Pera Asiática": ["ASIATICA", "ASIÁTICA"],
    "Gala Santa Carol": ["GALA SANTA", "SANTA CAROL"],
    "Gala Azaleia": ["GALA AZALEIA"],
    "Fuji Expressa 180": ["FUJI EXPRESSA", "FUJI 180"],
    "Fuji Azaleia": ["FUJI AZALEIA", "702145135", "702145165"],
    "Fuji Hiragami": ["HIRAGAMI"],
    "Fuji Suprema": ["FUJI SUPREMA"],
    "Thompson Vitace": ["THOMPSON VITACE"],
    "Thompson Seedless": ["THOMPSON SEEDLESS"],
    "Uva Crimson": ["CRINSON", "CRIMSON"],
    "Uva Isis": ["ISIS"],
    "Uva Jubilee": ["JUBILEE"],
    "Uva Itália": ["ITALIA", "ITÁLIA"],
    "Maçã Argentina": ["ARGENTINA"],
    "Maçã Chilena": ["CHILENA"],
    "Maçã Pink Lady": ["PINK LADY"],
    "Maçã Granny Smith": ["GRAN SMITH", "GRANNY SMITH"],
    "Maçã Red Globe": ["RED GLOBE"],
    "Mamão Havai": ["MAMAO HAVAI", "MAMÃO HAVAI"],
    "Mamão Formoso": ["MAMAO FORMOSO", "MAMÃO FORMOSO"],
    "Goiaba": ["GOIABA", "300200203", "300200208"],
    "Melão Amarelo": ["MELAO AMARELO", "MELÃO AMARELO"],
    "Melão Gaia": ["MELAO GAIA", "MELÃO GAIA", "MELAO GALIA", "MELÃO GALIA", "3102006"],
    "Melão Cantaloupe": ["CANTALOUPE"],
    "Tangerina Cumbuca": ["CUMBUCA", "830100903"],
    "Tangerina Ponkan": ["PONKAN"],
    "Ameixa": ["AMEIXA"],
    "Pêssego": ["PESSEGO", "PÊSSEGO"],
    "Nectarina": ["NECTARINA"],
    "Morango": ["MORANGO"],
    "Mirtilo": ["MIRTILO", "BLUEBERRY"],
    "Abacaxi": ["ABACAXI"],
    "Manga Palmer": ["MANGA PALMER"],
    "Manga Tommy": ["MANGA TOMMY"],
    "Laranja": ["LARANJA"],
    "Limão Tahiti": ["LIMAO TAHITI", "LIMÃO TAHITI"],
    "Limão Siciliano": ["LIMAO SICILIANO", "LIMÃO SICILIANO"],
    "Tomate Roma": ["TOMATE ROMA", "ROMA"],
}

RODAPE = [
    "É DE RESPONSABILIDADE DO VENDEDOR:",
    "AVALIAR DIARIAMENTE A QUALIDADE E ARMAZENAGEM DE CADA PRODUTO DE SUA RESPONSABILIDADE.",
    "CONFERIR O QUE ESTA EM CADA PAVILHÃO",
    "CONFERIR O QUE ESTA NA VENDA FUTURA E ACOMPANHAR DIARIAMENTE",
    "CONFERIR O QUE ESTA ARMAZENADO EM OUTROS FRIGORIFICOS",
    "VENDER ATÉ A ÚLTIMA CAIXA",
    "",
    "DEVOLUCAO SO SE FOR NO MESMO DIA",
    "",
    "MERCADORIAS NO SOL",
    "",
    "CAMINHOES REFRIGERADOS SEMPRE FECHADOS",
]

def mapear_produto(desc):
    desc_u = desc.upper()
    for prod_meta, keywords in MAPA_PRODUTO.items():
        for kw in keywords:
            if kw in desc_u:
                return prod_meta
    return None

# ── GitHub Storage ────────────────────────────────────────────────────────
GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", "")
GITHUB_REPO  = st.secrets.get("GITHUB_REPO", "Othilbh/Relatorios-Comerciais")
GITHUB_FILE_METAS    = "data/metas_semana.json"
GITHUB_FILE_PRODUTOS = "data/produtos_extra.json"

def github_get(filepath):
    if not GITHUB_TOKEN:
        return None, None
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{filepath}"
    r = requests.get(url, headers={"Authorization": f"token {GITHUB_TOKEN}"})
    if r.status_code == 200:
        data = r.json()
        content = json.loads(base64.b64decode(data["content"]).decode("utf-8"))
        return content, data["sha"]
    return None, None

def github_save(filepath, content, sha=None):
    if not GITHUB_TOKEN:
        return False
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{filepath}"
    body = {
        "message": f"Atualiza {filepath}",
        "content": base64.b64encode(json.dumps(content, ensure_ascii=False, indent=2).encode()).decode(),
    }
    if sha:
        body["sha"] = sha
    r = requests.put(url, headers={"Authorization": f"token {GITHUB_TOKEN}"}, json=body)
    return r.status_code in [200, 201]

# ── Carrega dados ─────────────────────────────────────────────────────────
if "produtos_meta" not in st.session_state:
    dados, _ = github_get(GITHUB_FILE_METAS)
    st.session_state.produtos_meta = dados or []
if "produtos_extra" not in st.session_state:
    dados, _ = github_get(GITHUB_FILE_PRODUTOS)
    st.session_state.produtos_extra = dados or []
if "vendido" not in st.session_state:
    st.session_state.vendido = {}
if "estoque" not in st.session_state:
    st.session_state.estoque = {}
if "sem_vendedor" not in st.session_state:
    st.session_state.sem_vendedor = []

PRODUTOS_LISTA = PRODUTOS_DEFAULT + [
    p for p in st.session_state.produtos_extra if p not in PRODUTOS_DEFAULT
]

# ── Estilos Excel ─────────────────────────────────────────────────────────
COR_H   = "1A3A5C"
COR_S   = "2E6DA4"
COR_V   = "C6EFCE"; COR_VF = "276221"
COR_A   = "FFEB9C"; COR_AF = "7D6608"
COR_R   = "FFC7CE"; COR_RF = "9C0006"
COR_CINZA = "D9D9D9"
COR_TOTAL = "BDD7EE"

def _fill(c): return PatternFill("solid", fgColor=c)
def _font(bold=False, color="000000", size=9):
    return Font(bold=bold, color=color, size=size, name="Calibri")
def _border(style="thin", color="BFBFBF"):
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)
def _medium():
    m = Side(style="medium", color="1A3A5C")
    return Border(left=m, right=m, top=m, bottom=m)
def _alinhar(h="center", wrap=False):
    return Alignment(horizontal=h, vertical="center", wrap_text=wrap)

def set_print_a4(ws):
    from openpyxl.worksheet.page import PageMargins
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.75, bottom=0.75)
    ws.print_title_rows = "1:3"

def cor_pct(pct):
    if pct >= 100: return COR_V, COR_VF
    if pct >= 50:  return COR_A, COR_AF
    return COR_R, COR_RF

# ── Aba por Vendedor ──────────────────────────────────────────────────────
def gerar_aba_vendedor(wb, vendedor, data_ref, itens_estoque, metas, vendido_v):
    ws = wb.create_sheet(vendedor[:31])
    ws.sheet_view.showGridLines = False
    set_print_a4(ws)

    # Cabeçalho
    ws.row_dimensions[1].height = 22
    ws.merge_cells("A1:G1")
    c = ws["A1"]
    c.value = f"OTHIL — RELATÓRIO DIÁRIO  |  Vendedor: {vendedor.upper()}  |  Data: {data_ref}"
    c.fill = _fill(COR_H)
    c.font = _font(bold=True, color="FFFFFF", size=11)
    c.alignment = _alinhar("left")
    c.border = _medium()

    # Título estoque
    ws.row_dimensions[2].height = 16
    ws.merge_cells("A2:G2")
    c2 = ws["A2"]
    c2.value = "▌ ESTOQUE — PRODUTOS SOB SUA RESPONSABILIDADE"
    c2.fill = _fill(COR_S)
    c2.font = _font(bold=True, color="FFFFFF", size=10)
    c2.alignment = _alinhar("left")

    # Header estoque
    hdrs = ["Produto", "Complemento", "Dt. Entrada", "Saldo Atual", "Qtde Vendida", "Custo Unit.", "Md Venda"]
    ws.row_dimensions[3].height = 14
    for ci, h in enumerate(hdrs, 1):
        c = ws.cell(row=3, column=ci, value=h)
        c.fill = _fill(COR_CINZA)
        c.font = _font(bold=True, size=9)
        c.alignment = _alinhar("center")
        c.border = _border("thin", "1A3A5C")

    row = 4
    for i, item in enumerate(itens_estoque):
        bg = "FFFFFF" if i % 2 == 0 else "F5F8FF"
        vals = [item["descricao"], item["complemento"], item["data_entrada"],
                item["saldo_atual"], item["qtd_vendida"], item["custo"], item["md_venda"]]
        ws.row_dimensions[row].height = 13
        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=row, column=ci, value=val)
            c.fill = _fill(bg)
            c.font = _font(size=9)
            c.border = _border()
            if ci == 1:
                c.alignment = _alinhar("left")
            elif ci in [6, 7]:
                c.number_format = 'R$ #,##0.00'
                c.alignment = _alinhar("right")
                if val == 0:
                    c.font = _font(color="FF0000", size=9)
            elif ci == 5:
                c.alignment = _alinhar("center")
                if val == 0:
                    c.font = _font(color="FF0000", size=9)
            else:
                c.alignment = _alinhar("center")
        row += 1

    row += 1

    # Título metas
    if metas:
        ws.merge_cells(f"A{row}:G{row}")
        c = ws.cell(row=row, column=1,
            value=f"▌ METAS SEMANAIS — {vendedor.upper()}")
        c.fill = _fill(COR_H)
        c.font = _font(bold=True, color="FFFFFF", size=10)
        c.alignment = _alinhar("left")
        ws.row_dimensions[row].height = 16
        row += 1

        # Header metas
        hdrs_m = ["Produto", "Estoque", "Meta (cx)", "Vendido (cx)", "Falta (cx)", "% Atingido", "Status"]
        for ci, h in enumerate(hdrs_m, 1):
            c = ws.cell(row=row, column=ci, value=h)
            c.fill = _fill(COR_S)
            c.font = _font(bold=True, color="FFFFFF", size=9)
            c.alignment = _alinhar("center")
            c.border = _border("thin", "1A3A5C")
        ws.row_dimensions[row].height = 14
        row += 1

        total_meta = total_vend = 0
        for item in metas:
            produto = item["produto"]
            est = item["estoque"]
            meta = math.ceil(est * PERCENTUAIS.get(vendedor, 0))
            vend = vendido_v.get(produto, 0.0)
            falta = max(meta - vend, 0)
            pct = (vend / meta * 100) if meta > 0 else 0.0
            total_meta += meta
            total_vend += vend
            bg_m, fg_m = cor_pct(pct)
            status = "✅ Atingida" if pct >= 100 else "⚠️ Andamento" if pct >= 50 else "❌ Abaixo"

            vals = [produto, int(est), meta, round(vend,1), round(falta,1), f"{pct:.1f}%", status]
            ws.row_dimensions[row].height = 13
            for ci, val in enumerate(vals, 1):
                c = ws.cell(row=row, column=ci, value=val)
                c.fill = _fill(bg_m)
                c.font = _font(color=fg_m, size=9)
                c.border = _border()
                c.alignment = _alinhar("left" if ci == 1 else "center")
            row += 1

        # Total metas
        pct_t = (total_vend / total_meta * 100) if total_meta > 0 else 0
        bg_t, fg_t = cor_pct(pct_t)
        vals_t = ["TOTAL", "", total_meta, round(total_vend,1),
                  round(max(total_meta-total_vend,0),1), f"{pct_t:.1f}%", ""]
        ws.row_dimensions[row].height = 14
        for ci, val in enumerate(vals_t, 1):
            c = ws.cell(row=row, column=ci, value=val)
            c.fill = _fill(bg_t)
            c.font = _font(bold=True, color=fg_t, size=9)
            c.border = _border("thin", "1A3A5C")
            c.alignment = _alinhar("left" if ci == 1 else "center")
        row += 2

    # Rodapé
    for texto in RODAPE:
        ws.merge_cells(f"A{row}:G{row}")
        c = ws.cell(row=row, column=1, value=texto)
        if texto.startswith("É DE RESPONSABILIDADE"):
            c.font = _font(bold=True, size=10, color="FF0000")
            ws.row_dimensions[row].height = 15
        else:
            c.font = _font(size=9)
            ws.row_dimensions[row].height = 12
        c.alignment = _alinhar("left")
        row += 1

    # Larguras
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 11
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 12

    return ws


# ── Aba Resumo Geral ──────────────────────────────────────────────────────
def gerar_aba_resumo(wb, metas, vendido_todos, data_ref, periodo):
    ws = wb.create_sheet("Resumo Geral", 0)
    ws.sheet_view.showGridLines = False
    set_print_a4(ws)

    # Título
    n_cols = len(VENDEDORES_ATIVOS) + 2
    ws.merge_cells(f"A1:{get_column_letter(n_cols)}1")
    c = ws["A1"]
    c.value = f"OTHIL — RESUMO GERAL DE METAS  |  {periodo}  |  Gerado: {data_ref}"
    c.fill = _fill(COR_H)
    c.font = _font(bold=True, color="FFFFFF", size=12)
    c.alignment = _alinhar("center")
    ws.row_dimensions[1].height = 24

    secoes = [
        ("META (cx)", lambda p, v, e: math.ceil(e * PERCENTUAIS.get(v, 0))),
        ("VENDIDO (cx)", lambda p, v, e: round(vendido_todos.get(v, {}).get(p, 0), 1)),
        ("FALTA (cx)", lambda p, v, e: max(math.ceil(e * PERCENTUAIS.get(v, 0)) - vendido_todos.get(v, {}).get(p, 0), 0)),
        ("% ATINGIDO", lambda p, v, e: round((vendido_todos.get(v, {}).get(p, 0) / math.ceil(e * PERCENTUAIS.get(v, 0)) * 100) if math.ceil(e * PERCENTUAIS.get(v, 0)) > 0 else 0, 1)),
    ]

    row = 2
    for secao_nome, calc in secoes:
        # Título seção
        ws.merge_cells(f"A{row}:{get_column_letter(n_cols)}{row}")
        c = ws.cell(row=row, column=1, value=f"▌ {secao_nome}")
        c.fill = _fill(COR_S)
        c.font = _font(bold=True, color="FFFFFF", size=10)
        c.alignment = _alinhar("left")
        ws.row_dimensions[row].height = 16
        row += 1

        # Header
        ws.cell(row=row, column=1, value="Produto").fill = _fill(COR_CINZA)
        ws.cell(row=row, column=1).font = _font(bold=True, size=9)
        ws.cell(row=row, column=1).border = _border("thin", "1A3A5C")
        ws.cell(row=row, column=1).alignment = _alinhar("left")

        for ci, v in enumerate(VENDEDORES_ATIVOS, 2):
            c = ws.cell(row=row, column=ci, value=v)
            c.fill = _fill(COR_CINZA)
            c.font = _font(bold=True, size=9)
            c.border = _border("thin", "1A3A5C")
            c.alignment = _alinhar("center")

        c_total = ws.cell(row=row, column=n_cols, value="TOTAL")
        c_total.fill = _fill(COR_H)
        c_total.font = _font(bold=True, color="FFFFFF", size=9)
        c_total.border = _border("thin", "1A3A5C")
        c_total.alignment = _alinhar("center")
        ws.row_dimensions[row].height = 14
        row += 1

        total_vendedor = {v: 0.0 for v in VENDEDORES_ATIVOS}

        for i, item in enumerate(metas):
            produto = item["produto"]
            est = item["estoque"]
            bg = "FFFFFF" if i % 2 == 0 else "F5F8FF"

            ws.cell(row=row, column=1, value=produto).alignment = _alinhar("left")
            ws.cell(row=row, column=1).font = _font(size=9)
            ws.cell(row=row, column=1).fill = _fill(bg)
            ws.cell(row=row, column=1).border = _border()

            total_linha = 0
            for ci, v in enumerate(VENDEDORES_ATIVOS, 2):
                val = calc(produto, v, est)
                total_linha += val
                total_vendedor[v] += val

                c = ws.cell(row=row, column=ci, value=val)
                c.font = _font(size=9)
                c.alignment = _alinhar("center")
                c.border = _border()

                if secao_nome == "% ATINGIDO":
                    bg_pct, fg_pct = cor_pct(val)
                    c.fill = _fill(bg_pct)
                    c.font = _font(color=fg_pct, size=9)
                    c.number_format = '0.0"%"'
                else:
                    c.fill = _fill(bg)

            # Total linha
            c_tl = ws.cell(row=row, column=n_cols,
                value=round(total_linha/len(VENDEDORES_ATIVOS), 1) if secao_nome == "% ATINGIDO" else round(total_linha, 1))
            c_tl.fill = _fill(COR_TOTAL)
            c_tl.font = _font(bold=True, size=9)
            c_tl.border = _border("thin", "1A3A5C")
            c_tl.alignment = _alinhar("center")
            if secao_nome == "% ATINGIDO":
                bg_pct, fg_pct = cor_pct(total_linha/len(VENDEDORES_ATIVOS))
                c_tl.fill = _fill(bg_pct)
                c_tl.font = _font(bold=True, color=fg_pct, size=9)
                c_tl.number_format = '0.0"%"'
            ws.row_dimensions[row].height = 13
            row += 1

        # Total coluna
        ws.cell(row=row, column=1, value="TOTAL").fill = _fill(COR_H)
        ws.cell(row=row, column=1).font = _font(bold=True, color="FFFFFF", size=9)
        ws.cell(row=row, column=1).border = _border("thin", "1A3A5C")
        ws.cell(row=row, column=1).alignment = _alinhar("left")

        grand_total = 0
        for ci, v in enumerate(VENDEDORES_ATIVOS, 2):
            val = round(total_vendedor[v]/len(metas), 1) if secao_nome == "% ATINGIDO" else round(total_vendedor[v], 1)
            grand_total += total_vendedor[v]
            c = ws.cell(row=row, column=ci, value=val)
            c.fill = _fill(COR_TOTAL)
            c.font = _font(bold=True, size=9)
            c.border = _border("thin", "1A3A5C")
            c.alignment = _alinhar("center")
            if secao_nome == "% ATINGIDO":
                bg_pct, fg_pct = cor_pct(val)
                c.fill = _fill(bg_pct)
                c.font = _font(bold=True, color=fg_pct, size=9)
                c.number_format = '0.0"%"'

        grand_val = round(grand_total/len(VENDEDORES_ATIVOS)/len(metas), 1) if secao_nome == "% ATINGIDO" else round(grand_total, 1)
        c_gt = ws.cell(row=row, column=n_cols, value=grand_val)
        c_gt.fill = _fill(COR_H)
        c_gt.font = _font(bold=True, color="FFFFFF", size=9)
        c_gt.border = _border("thin", "1A3A5C")
        c_gt.alignment = _alinhar("center")
        if secao_nome == "% ATINGIDO":
            bg_pct, fg_pct = cor_pct(grand_val)
            c_gt.fill = _fill(bg_pct)
            c_gt.font = _font(bold=True, color=fg_pct, size=9)
            c_gt.number_format = '0.0"%"'

        ws.row_dimensions[row].height = 14
        row += 2

    # Larguras
    ws.column_dimensions["A"].width = 28
    for i in range(2, n_cols + 1):
        ws.column_dimensions[get_column_letter(i)].width = 11

    return ws


# ── Aba Dashboard ─────────────────────────────────────────────────────────
def gerar_aba_dashboard(wb, metas, vendido_todos, data_ref, periodo):
    ws = wb.create_sheet("Dashboard", 1)
    ws.sheet_view.showGridLines = False

    # Título
    ws.merge_cells("A1:L1")
    c = ws["A1"]
    c.value = f"📊 DASHBOARD — METAS SEMANAIS OTHIL  |  {periodo}"
    c.fill = _fill(COR_H)
    c.font = _font(bold=True, color="FFFFFF", size=13)
    c.alignment = _alinhar("center")
    ws.row_dimensions[1].height = 28

    # KPIs
    ws.merge_cells("A2:L2")
    ws.cell(row=2, column=1, value="KPIs GERAIS").fill = _fill(COR_S)
    ws.cell(row=2, column=1).font = _font(bold=True, color="FFFFFF", size=10)
    ws.cell(row=2, column=1).alignment = _alinhar("left")
    ws.row_dimensions[2].height = 16

    # Calcula KPIs
    total_meta_geral = sum(
        math.ceil(item["estoque"] * PERCENTUAIS.get(v, 0))
        for item in metas for v in VENDEDORES_ATIVOS
    )
    total_vend_geral = sum(
        vendido_todos.get(v, {}).get(item["produto"], 0)
        for item in metas for v in VENDEDORES_ATIVOS
    )
    pct_geral = (total_vend_geral / total_meta_geral * 100) if total_meta_geral > 0 else 0
    produtos_criticos = sum(
        1 for item in metas
        if sum(vendido_todos.get(v, {}).get(item["produto"], 0) for v in VENDEDORES_ATIVOS) /
           max(sum(math.ceil(item["estoque"] * PERCENTUAIS.get(v, 0)) for v in VENDEDORES_ATIVOS), 1) * 100 < 50
    )

    kpis = [
        ("Meta Total (cx)", total_meta_geral, COR_S, "FFFFFF"),
        ("Vendido Total (cx)", round(total_vend_geral, 1), "2E7D32", "FFFFFF"),
        ("Falta Total (cx)", round(max(total_meta_geral - total_vend_geral, 0), 1), "C62828", "FFFFFF"),
        ("% Atingido Geral", f"{pct_geral:.1f}%", *cor_pct(pct_geral)),
        ("Produtos Críticos (<50%)", produtos_criticos, "E65100", "FFFFFF"),
        ("Produtos na Meta", len(metas) - produtos_criticos, "1B5E20", "FFFFFF"),
    ]

    row = 3
    for col_start, (label, valor, bg, fg) in enumerate(kpis):
        col = col_start * 2 + 1
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col+1)
        ws.merge_cells(start_row=row+1, start_column=col, end_row=row+1, end_column=col+1)
        c_label = ws.cell(row=row, column=col, value=label)
        c_label.fill = _fill(COR_CINZA)
        c_label.font = _font(bold=True, size=9)
        c_label.alignment = _alinhar("center")
        c_label.border = _border("thin", "1A3A5C")
        c_valor = ws.cell(row=row+1, column=col, value=valor)
        c_valor.fill = _fill(bg)
        c_valor.font = _font(bold=True, color=fg, size=14)
        c_valor.alignment = _alinhar("center")
        c_valor.border = _border("medium", "1A3A5C")
        ws.row_dimensions[row].height = 14
        ws.row_dimensions[row+1].height = 28

    row = 6

    # % por vendedor
    ws.merge_cells(f"A{row}:F{row}")
    ws.cell(row=row, column=1, value="% ATINGIDO POR VENDEDOR").fill = _fill(COR_S)
    ws.cell(row=row, column=1).font = _font(bold=True, color="FFFFFF", size=10)
    ws.cell(row=row, column=1).alignment = _alinhar("left")
    ws.row_dimensions[row].height = 16
    row += 1

    hdrs_v = ["Vendedor", "Meta", "Vendido", "Falta", "% Atingido", "Status"]
    for ci, h in enumerate(hdrs_v, 1):
        c = ws.cell(row=row, column=ci, value=h)
        c.fill = _fill(COR_CINZA)
        c.font = _font(bold=True, size=9)
        c.border = _border("thin", "1A3A5C")
        c.alignment = _alinhar("center")
    ws.row_dimensions[row].height = 14
    row += 1

    pct_por_vendedor = []
    for v in VENDEDORES_ATIVOS:
        meta_v = sum(math.ceil(item["estoque"] * PERCENTUAIS.get(v, 0)) for item in metas)
        vend_v = sum(vendido_todos.get(v, {}).get(item["produto"], 0) for item in metas)
        falta_v = max(meta_v - vend_v, 0)
        pct_v = (vend_v / meta_v * 100) if meta_v > 0 else 0
        pct_por_vendedor.append(round(pct_v, 1))
        status = "✅ Atingiu" if pct_v >= 100 else "⚠️ Andamento" if pct_v >= 50 else "❌ Abaixo"
        bg_v, fg_v = cor_pct(pct_v)

        vals = [v, meta_v, round(vend_v,1), round(falta_v,1), f"{pct_v:.1f}%", status]
        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=row, column=ci, value=val)
            c.fill = _fill(bg_v) if ci >= 4 else _fill("FFFFFF" if row%2==0 else "F5F8FF")
            c.font = _font(color=fg_v if ci >= 4 else "000000", size=9)
            c.border = _border()
            c.alignment = _alinhar("left" if ci == 1 else "center")
        ws.row_dimensions[row].height = 13
        row += 1

    # Gráfico % por vendedor
    chart_row = 6
    chart = BarChart()
    chart.type = "col"
    chart.title = "% Atingido por Vendedor"
    chart.y_axis.title = "%"
    chart.style = 10
    chart.width = 14
    chart.height = 10

    data_chart = Reference(ws, min_col=5, min_row=8, max_row=8+len(VENDEDORES_ATIVOS)-1)
    cats = Reference(ws, min_col=1, min_row=9, max_row=8+len(VENDEDORES_ATIVOS)-1)
    chart.add_data(data_chart)
    chart.set_categories(cats)
    chart.series[0].title = None
    ws.add_chart(chart, f"G{chart_row}")

    # Produtos críticos
    row += 2
    ws.merge_cells(f"A{row}:F{row}")
    ws.cell(row=row, column=1, value="🚨 PRODUTOS CRÍTICOS (< 50% ATINGIDO)").fill = _fill("C62828")
    ws.cell(row=row, column=1).font = _font(bold=True, color="FFFFFF", size=10)
    ws.cell(row=row, column=1).alignment = _alinhar("left")
    ws.row_dimensions[row].height = 16
    row += 1

    hdrs_c = ["Produto", "Meta Total", "Vendido Total", "Falta", "% Geral", "Vendedor com Pior %"]
    for ci, h in enumerate(hdrs_c, 1):
        c = ws.cell(row=row, column=ci, value=h)
        c.fill = _fill(COR_CINZA)
        c.font = _font(bold=True, size=9)
        c.border = _border("thin", "1A3A5C")
        c.alignment = _alinhar("center")
    ws.row_dimensions[row].height = 14
    row += 1

    for item in metas:
        produto = item["produto"]
        est = item["estoque"]
        meta_tot = sum(math.ceil(est * PERCENTUAIS.get(v, 0)) for v in VENDEDORES_ATIVOS)
        vend_tot = sum(vendido_todos.get(v, {}).get(produto, 0) for v in VENDEDORES_ATIVOS)
        pct_tot = (vend_tot / meta_tot * 100) if meta_tot > 0 else 0

        if pct_tot < 50:
            pior_v = min(VENDEDORES_ATIVOS,
                key=lambda v: (vendido_todos.get(v,{}).get(produto,0) /
                    max(math.ceil(est*PERCENTUAIS.get(v,0)),1)*100))
            vals = [produto, meta_tot, round(vend_tot,1),
                    round(max(meta_tot-vend_tot,0),1), f"{pct_tot:.1f}%", pior_v]
            for ci, val in enumerate(vals, 1):
                c = ws.cell(row=row, column=ci, value=val)
                c.fill = _fill(COR_R)
                c.font = _font(color=COR_RF, size=9)
                c.border = _border()
                c.alignment = _alinhar("left" if ci == 1 else "center")
            ws.row_dimensions[row].height = 13
            row += 1

    # Larguras
    ws.column_dimensions["A"].width = 28
    for col in ["B","C","D","E","F"]:
        ws.column_dimensions[col].width = 13
    for col in ["G","H","I","J","K","L"]:
        ws.column_dimensions[col].width = 10

    return ws


# ══════════════════════════════════════════════════════════════════════════
# INTERFACE
# ══════════════════════════════════════════════════════════════════════════

st.subheader("1️⃣ Metas da Semana")

col_ini, col_fim = st.columns(2)
with col_ini:
    semana_ini = st.date_input("Início",
        value=date.today() - timedelta(days=date.today().weekday()))
with col_fim:
    semana_fim = st.date_input("Fim", value=semana_ini + timedelta(days=5))

periodo = f"{semana_ini.strftime('%d/%m/%Y')} a {semana_fim.strftime('%d/%m/%Y')}"

st.caption("Busque ou digite um produto para adicionar:")
produtos_ja = [p["produto"] for p in st.session_state.produtos_meta]
produtos_disp = [p for p in PRODUTOS_LISTA if p not in produtos_ja]

col_busca, col_novo, col_est, col_btn = st.columns([2, 2, 1, 1])
with col_busca:
    prod_sel = st.selectbox("Buscar", options=[""] + produtos_disp,
        format_func=lambda x: "🔍 Selecione da lista..." if x == "" else x,
        label_visibility="collapsed")
with col_novo:
    prod_novo = st.text_input("Novo", placeholder="✏️ Ou digite produto novo...",
        label_visibility="collapsed")
with col_est:
    est_input = st.number_input("Estoque", min_value=0.0, step=1.0,
        label_visibility="collapsed", placeholder="Estoque CX")
with col_btn:
    if st.button("➕ Adicionar", use_container_width=True, type="primary"):
        prod_final = prod_novo.strip() if prod_novo.strip() else prod_sel
        if prod_final and prod_final != "":
            if prod_final not in produtos_ja:
                st.session_state.produtos_meta.append({"produto": prod_final, "estoque": est_input})
                if prod_final not in PRODUTOS_LISTA:
                    st.session_state.produtos_extra.append(prod_final)
                    _, sha = github_get(GITHUB_FILE_PRODUTOS)
                    github_save(GITHUB_FILE_PRODUTOS, st.session_state.produtos_extra, sha)
                    st.toast(f"✅ '{prod_final}' salvo na lista permanente!")
                st.rerun()
        else:
            st.warning("Selecione ou digite um produto.")

if st.session_state.produtos_meta:
    rows = []
    for item in st.session_state.produtos_meta:
        est = item["estoque"]
        row = {"Produto": item["produto"], "Estoque CX": int(est)}
        for v in VENDEDORES_ATIVOS:
            row[f"{v} ({int(PERCENTUAIS[v]*100)}%)"] = math.ceil(est * PERCENTUAIS[v])
        rows.append(row)

    df_edit = st.data_editor(
        pd.DataFrame(rows), use_container_width=True, hide_index=True,
        disabled=["Produto"] + [f"{v} ({int(PERCENTUAIS[v]*100)}%)" for v in VENDEDORES_ATIVOS],
        key="editor_metas"
    )
    for i, row in df_edit.iterrows():
        if i < len(st.session_state.produtos_meta):
            st.session_state.produtos_meta[i]["estoque"] = row["Estoque CX"]

    col_salvar, col_apagar, col_info = st.columns([1, 1, 3])
    with col_salvar:
        if st.button("💾 Salvar Metas", type="primary", use_container_width=True):
            _, sha = github_get(GITHUB_FILE_METAS)
            ok = github_save(GITHUB_FILE_METAS, st.session_state.produtos_meta, sha)
            if ok:
                st.success("✅ Metas salvas!")
            else:
                st.error("❌ Erro ao salvar.")
    with col_apagar:
        if st.button("🗑️ Apagar Metas", type="secondary", use_container_width=True):
            _, sha = github_get(GITHUB_FILE_METAS)
            github_save(GITHUB_FILE_METAS, [], sha)
            st.session_state.produtos_meta = []
            st.session_state.vendido = {}
            st.session_state.estoque = {}
            st.rerun()
    with col_info:
        st.caption(f"💡 Semana: **{periodo}** · Salve toda segunda-feira.")

st.divider()

# ── Upload PDFs ───────────────────────────────────────────────────────────
st.subheader("2️⃣ Upload dos PDFs")

col_pdf1, col_pdf2 = st.columns(2)
with col_pdf1:
    st.caption("📊 PDF de Vendas Acumuladas")
    pdf_vendas = st.file_uploader("Vendas", type=["pdf"], label_visibility="collapsed")
with col_pdf2:
    st.caption("📦 PDF de Estoque Físico")
    pdf_estoque = st.file_uploader("Estoque", type=["pdf"], label_visibility="collapsed")

if pdf_vendas:
    with st.spinner("🔍 Lendo vendas..."):
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp.write(pdf_vendas.read()); tmp_path = tmp.name
        vendas_raw = extrair_vendas_por_vendedor(tmp_path)
        os.unlink(tmp_path)
    vendido_c = {v: {} for v in VENDEDORES_ATIVOS}
    for vend, prods in vendas_raw.items():
        if vend not in vendido_c: continue
        for desc, qtd in prods.items():
            pm = mapear_produto(desc)
            if pm:
                vendido_c[vend][pm] = vendido_c[vend].get(pm, 0) + qtd
    st.session_state.vendido = vendido_c
    found = [v for v in vendas_raw if v in VENDEDORES_ATIVOS]
    st.success(f"✅ Vendas: {', '.join(found)}")

if pdf_estoque:
    with st.spinner("🔍 Lendo estoque..."):
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp.write(pdf_estoque.read()); tmp_path = tmp.name
        est_raw = extrair_estoque_por_vendedor(tmp_path)
        os.unlink(tmp_path)
    st.session_state.estoque = {v: est_raw.get(v, []) for v in VENDEDORES_ATIVOS}
    st.session_state.sem_vendedor = est_raw.get("SEM_VENDEDOR", [])
    total = sum(len(v) for v in st.session_state.estoque.values())
    st.success(f"✅ Estoque: {total} produto(s) distribuídos.")
    if st.session_state.sem_vendedor:
        with st.expander(f"⚠️ {len(st.session_state.sem_vendedor)} produto(s) SEM VENDEDOR", expanded=True):
            df_sv = pd.DataFrame(st.session_state.sem_vendedor)[["codigo","descricao","data_entrada","saldo_atual"]]
            df_sv.columns = ["Código","Descrição","Data Entrada","Saldo Atual"]
            st.dataframe(df_sv, use_container_width=True, hide_index=True)

st.divider()

# ── Gerar Relatórios ──────────────────────────────────────────────────────
st.subheader("3️⃣ Gerar Relatórios")

if st.button("📋 Gerar Relatórios Completos", use_container_width=True, type="primary"):
    if not st.session_state.estoque:
        st.warning("Faça upload do PDF de Estoque primeiro.")
    else:
        with st.spinner("⚙️ Gerando Excel..."):
            wb = Workbook()
            wb.remove(wb.active)
            data_ref = date.today().strftime("%d/%m/%Y")

            # Abas de resumo e dashboard primeiro
            if st.session_state.produtos_meta:
                gerar_aba_resumo(wb, st.session_state.produtos_meta,
                    st.session_state.vendido, data_ref, periodo)
                gerar_aba_dashboard(wb, st.session_state.produtos_meta,
                    st.session_state.vendido, data_ref, periodo)

            # Abas por vendedor
            for vendedor in VENDEDORES_ATIVOS:
                gerar_aba_vendedor(
                    wb, vendedor, data_ref,
                    st.session_state.estoque.get(vendedor, []),
                    st.session_state.produtos_meta,
                    st.session_state.vendido.get(vendedor, {})
                )

            buf = io.BytesIO()
            wb.save(buf); buf.seek(0)

        nome = f"Relatorios_{date.today().strftime('%d%m%Y')}.xlsx"
        st.success("✅ Relatórios gerados com sucesso!")
        st.download_button(
            label="⬇️ Baixar Excel Completo",
            data=buf, file_name=nome,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )