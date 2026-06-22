# pages/4_Metas_Semanais.py
import streamlit as st
from datetime import date, timedelta
import sys, os, tempfile
import pandas as pd
import io
import json
import math
sys.path.append(os.path.dirname(os.path.dirname(__file__)))

from utils.parser import extrair_vendas_por_vendedor, extrair_estoque_por_vendedor, VENDEDORES_ATIVOS
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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

def mapear_produto(desc: str) -> str:
    desc_u = desc.upper()
    for prod_meta, keywords in MAPA_PRODUTO.items():
        for kw in keywords:
            if kw in desc_u:
                return prod_meta
    return None

ARQUIVO_PRODUTOS_EXTRA = "/tmp/othil_produtos_extra.json"
ARQUIVO_METAS_SEMANA   = "/tmp/othil_metas_semana.json"

def carregar_json(path):
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return None

def salvar_json(path, dados):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

produtos_extras = carregar_json(ARQUIVO_PRODUTOS_EXTRA) or []
PRODUTOS_LISTA = PRODUTOS_DEFAULT + [p for p in produtos_extras if p not in PRODUTOS_DEFAULT]
metas_salvas = carregar_json(ARQUIVO_METAS_SEMANA) or []

if "produtos_meta" not in st.session_state:
    st.session_state.produtos_meta = metas_salvas
if "vendido" not in st.session_state:
    st.session_state.vendido = {}
if "estoque" not in st.session_state:
    st.session_state.estoque = {}
if "sem_vendedor" not in st.session_state:
    st.session_state.sem_vendedor = []

# ── Cores e estilos Excel ─────────────────────────────────────────────────
COR_H  = "1A3A5C"
COR_S  = "2E6DA4"
COR_V  = "C6EFCE"; COR_VF = "276221"
COR_A  = "FFEB9C"; COR_AF = "7D6608"
COR_R  = "FFC7CE"; COR_RF = "9C0006"
COR_TITULO = "F5A623"

def _fill(c): return PatternFill("solid", fgColor=c)
def _font(bold=False, color="000000", size=10): return Font(bold=bold, color=color, size=size)
def _border():
    t = Side(style="thin", color="BFBFBF")
    return Border(left=t, right=t, top=t, bottom=t)
def _alinhar(h="center"): return Alignment(horizontal=h, vertical="center", wrap_text=True)

def hc(ws, r, c, v, bg=COR_H, fg="FFFFFF", bold=True, size=10, align="center"):
    cell = ws.cell(row=r, column=c, value=v)
    cell.fill = _fill(bg)
    cell.font = _font(bold=bold, color=fg, size=size)
    cell.alignment = _alinhar(align)
    cell.border = _border()
    return cell

def dc(ws, r, c, v, fmt=None, bg=None, bold=False, align="center"):
    cell = ws.cell(row=r, column=c, value=v)
    if fmt: cell.number_format = fmt
    if bg: cell.fill = _fill(bg)
    cell.font = _font(bold=bold)
    cell.alignment = _alinhar(align)
    cell.border = _border()
    return cell

def gerar_aba_vendedor(wb, vendedor, data_ref, itens_estoque, metas_vendedor, vendido_vendedor):
    ws = wb.create_sheet(vendedor[:31])
    ws.sheet_view.showGridLines = False

    # ── Cabeçalho ──────────────────────────────────────────────────────
    ws.merge_cells("A1:G1")
    c = ws["A1"]
    c.value = f"Vendedor: {vendedor.upper()}"
    c.font = _font(bold=True, size=13, color=COR_H)
    c.alignment = _alinhar("left")

    ws.merge_cells("A2:G2")
    c2 = ws["A2"]
    c2.value = f"Data: {data_ref}"
    c2.font = _font(size=10, color="666666")
    c2.alignment = _alinhar("left")

    # ── Tabela de Estoque ───────────────────────────────────────────────
    ws.merge_cells("A4:G4")
    hc(ws, 4, 1, "ESTOQUE — PRODUTOS SOB SUA RESPONSABILIDADE",
       bg=COR_H, size=11, align="left")

    hdrs_est = ["Produto", "Complemento", "Data Entrada", "Saldo Atual", "Qtde Vendida", "Custo Unitário", "Md Venda"]
    for ci, h in enumerate(hdrs_est, 1):
        hc(ws, 5, ci, h, bg=COR_S)

    row = 6
    for item in itens_estoque:
        dc(ws, row, 1, item["descricao"], align="left")
        dc(ws, row, 2, item["complemento"], align="left")
        dc(ws, row, 3, item["data_entrada"])
        dc(ws, row, 4, item["saldo_atual"], fmt='#,##0.0')
        dc(ws, row, 5, item["qtd_vendida"], fmt='#,##0.0')
        dc(ws, row, 6, item["custo"], fmt='R$ #,##0.00')
        dc(ws, row, 7, item["md_venda"], fmt='R$ #,##0.00')
        row += 1

    row += 1  # linha em branco

    # ── Tabela de Metas ─────────────────────────────────────────────────
    if metas_vendedor:
        ws.merge_cells(f"A{row}:G{row}")
        hc(ws, row, 1, f"METAS SEMANAIS — {vendedor.upper()}",
           bg=COR_TITULO, fg="000000", size=11, align="left")
        row += 1

        hdrs_meta = ["Produto", "Estoque CX", "Meta (cx)", "Vendido (cx)", "Falta (cx)", "% Atingido", "Status"]
        for ci, h in enumerate(hdrs_meta, 1):
            hc(ws, row, ci, h, bg=COR_S)
        row += 1

        total_meta = total_vend = 0
        for item in metas_vendedor:
            produto = item["produto"]
            est = item["estoque"]
            meta = math.ceil(est * PERCENTUAIS[vendedor])
            vend = vendido_vendedor.get(produto, 0.0)
            falta = max(meta - vend, 0)
            pct = (vend / meta * 100) if meta > 0 else 0.0
            total_meta += meta
            total_vend += vend

            if pct >= 100: bg_m, fg_m, status = COR_V, COR_VF, "✅ Atingida"
            elif pct >= 50: bg_m, fg_m, status = COR_A, COR_AF, "⚠️ Em andamento"
            else: bg_m, fg_m, status = COR_R, COR_RF, "❌ Abaixo"

            dc(ws, row, 1, produto, align="left")
            dc(ws, row, 2, int(est))
            dc(ws, row, 3, meta)
            dc(ws, row, 4, round(vend, 1))
            dc(ws, row, 5, round(falta, 1))

            c_pct = ws.cell(row=row, column=6, value=round(pct, 1))
            c_pct.fill = _fill(bg_m); c_pct.font = _font(color=fg_m)
            c_pct.alignment = _alinhar(); c_pct.border = _border()
            c_pct.number_format = '0.00"%"'

            c_st = ws.cell(row=row, column=7, value=status)
            c_st.fill = _fill(bg_m); c_st.font = _font(color=fg_m)
            c_st.alignment = _alinhar(); c_st.border = _border()
            row += 1

        # Total
        pct_total = (total_vend / total_meta * 100) if total_meta > 0 else 0
        hc(ws, row, 1, "TOTAL", bg=COR_H, align="left")
        dc(ws, row, 2, "")
        dc(ws, row, 3, total_meta, bold=True)
        dc(ws, row, 4, round(total_vend, 1), bold=True)
        dc(ws, row, 5, round(max(total_meta - total_vend, 0), 1), bold=True)
        c_pt = ws.cell(row=row, column=6, value=round(pct_total, 1))
        c_pt.font = _font(bold=True); c_pt.alignment = _alinhar(); c_pt.border = _border()
        c_pt.number_format = '0.00"%"'
        dc(ws, row, 7, "")
        row += 1

    row += 2  # espaço antes do rodapé

    # ── Rodapé ──────────────────────────────────────────────────────────
    for texto in RODAPE:
        ws.merge_cells(f"A{row}:G{row}")
        c_rod = ws.cell(row=row, column=1, value=texto)
        if texto.startswith("É DE RESPONSABILIDADE"):
            c_rod.font = _font(bold=True, size=11, color="FF0000")
        else:
            c_rod.font = _font(size=10)
        c_rod.alignment = _alinhar("left")
        row += 1

    # ── Larguras ────────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 13
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 15
    ws.column_dimensions["G"].width = 15

    return ws


# ══════════════════════════════════════════════════════════════════════════
# INTERFACE
# ══════════════════════════════════════════════════════════════════════════

st.subheader("1️⃣ Definir Metas da Semana")
col_ini, col_fim = st.columns(2)
with col_ini:
    semana_ini = st.date_input("Início da semana",
        value=date.today() - timedelta(days=date.today().weekday()))
with col_fim:
    semana_fim = st.date_input("Fim da semana",
        value=semana_ini + timedelta(days=5))

st.caption("Busque ou digite um produto para adicionar:")

produtos_ja_adicionados = [p["produto"] for p in st.session_state.produtos_meta]
produtos_disponiveis = [p for p in PRODUTOS_LISTA if p not in produtos_ja_adicionados]

col_busca, col_novo, col_estoque, col_btn = st.columns([2, 2, 1, 1])
with col_busca:
    produto_selecionado = st.selectbox(
        "Buscar", options=[""] + produtos_disponiveis,
        format_func=lambda x: "🔍 Selecione da lista..." if x == "" else x,
        label_visibility="collapsed"
    )
with col_novo:
    produto_novo = st.text_input("Novo", placeholder="✏️ Ou digite um produto novo...",
        label_visibility="collapsed")
with col_estoque:
    estoque_input = st.number_input("Estoque", min_value=0.0, step=1.0,
        label_visibility="collapsed", placeholder="Estoque CX")
with col_btn:
    if st.button("➕ Adicionar", use_container_width=True, type="primary"):
        produto_final = produto_novo.strip() if produto_novo.strip() else produto_selecionado
        if produto_final and produto_final != "":
            if produto_final not in produtos_ja_adicionados:
                st.session_state.produtos_meta.append({"produto": produto_final, "estoque": estoque_input})
                if produto_final not in PRODUTOS_LISTA:
                    produtos_extras.append(produto_final)
                    salvar_json(ARQUIVO_PRODUTOS_EXTRA, produtos_extras)
                    st.toast(f"✅ '{produto_final}' salvo na lista permanente!")
                salvar_json(ARQUIVO_METAS_SEMANA, st.session_state.produtos_meta)
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

    df_editado = st.data_editor(
        pd.DataFrame(rows), use_container_width=True, hide_index=True,
        disabled=["Produto"] + [f"{v} ({int(PERCENTUAIS[v]*100)}%)" for v in VENDEDORES_ATIVOS],
        key="editor_metas"
    )
    for i, row in df_editado.iterrows():
        if i < len(st.session_state.produtos_meta):
            st.session_state.produtos_meta[i]["estoque"] = row["Estoque CX"]
    salvar_json(ARQUIVO_METAS_SEMANA, st.session_state.produtos_meta)

    col_limpar, col_info, _ = st.columns([1, 2, 3])
    with col_limpar:
        if st.button("🗑️ Limpar tudo", type="secondary"):
            st.session_state.produtos_meta = []
            st.session_state.vendido = {}
            st.session_state.estoque = {}
            st.session_state.sem_vendedor = []
            salvar_json(ARQUIVO_METAS_SEMANA, [])
            st.rerun()
    with col_info:
        st.caption("💾 Salvo automaticamente")

st.divider()

# ── Upload PDFs ───────────────────────────────────────────────────────────
st.subheader("2️⃣ Upload dos PDFs")

col_pdf1, col_pdf2 = st.columns(2)
with col_pdf1:
    st.caption("📊 PDF de Vendas Acumuladas (Lucratividade por Vendedor)")
    pdf_vendas = st.file_uploader("Vendas", type=["pdf"], label_visibility="collapsed")
with col_pdf2:
    st.caption("📦 PDF de Estoque Físico (Responsáveis)")
    pdf_estoque = st.file_uploader("Estoque", type=["pdf"], label_visibility="collapsed")

if pdf_vendas:
    with st.spinner("🔍 Lendo PDF de vendas..."):
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp.write(pdf_vendas.read()); tmp_path = tmp.name
        vendas_raw = extrair_vendas_por_vendedor(tmp_path)
        os.unlink(tmp_path)

    vendido_consolidado = {v: {} for v in VENDEDORES_ATIVOS}
    for vendedor, produtos in vendas_raw.items():
        if vendedor not in vendido_consolidado:
            continue
        for desc, qtd in produtos.items():
            prod_meta = mapear_produto(desc)
            if prod_meta:
                vendido_consolidado[vendedor][prod_meta] = vendido_consolidado[vendedor].get(prod_meta, 0) + qtd
    st.session_state.vendido = vendido_consolidado
    found = [v for v in vendas_raw if v in VENDEDORES_ATIVOS]
    st.success(f"✅ Vendas: {len(found)} vendedor(es) — {', '.join(found)}")

if pdf_estoque:
    with st.spinner("🔍 Lendo PDF de estoque..."):
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp.write(pdf_estoque.read()); tmp_path = tmp.name
        estoque_raw = extrair_estoque_por_vendedor(tmp_path)
        os.unlink(tmp_path)

    st.session_state.estoque = {v: estoque_raw.get(v, []) for v in VENDEDORES_ATIVOS}
    st.session_state.sem_vendedor = estoque_raw.get("SEM_VENDEDOR", [])

    total_itens = sum(len(v) for v in st.session_state.estoque.values())
    st.success(f"✅ Estoque: {total_itens} produto(s) distribuídos por vendedor.")

    if st.session_state.sem_vendedor:
        with st.expander(f"⚠️ {len(st.session_state.sem_vendedor)} produto(s) SEM VENDEDOR identificado", expanded=True):
            df_sv = pd.DataFrame(st.session_state.sem_vendedor)[
                ["codigo", "descricao", "data_entrada", "saldo_atual"]
            ]
            df_sv.columns = ["Código", "Descrição", "Data Entrada", "Saldo Atual"]
            st.dataframe(df_sv, use_container_width=True, hide_index=True)

st.divider()

# ── Gerar Relatórios ──────────────────────────────────────────────────────
st.subheader("3️⃣ Gerar Relatórios por Vendedor")

if st.button("📋 Gerar Relatórios", use_container_width=True, type="primary"):
    if not st.session_state.estoque:
        st.warning("Faça upload do PDF de Estoque primeiro.")
    else:
        with st.spinner("⚙️ Gerando Excel..."):
            wb = Workbook()
            wb.remove(wb.active)
            data_ref = date.today().strftime("%d/%m/%Y")

            for vendedor in VENDEDORES_ATIVOS:
                itens_estoque = st.session_state.estoque.get(vendedor, [])
                vendido_v = st.session_state.vendido.get(vendedor, {})
                gerar_aba_vendedor(
                    wb, vendedor, data_ref,
                    itens_estoque,
                    st.session_state.produtos_meta,
                    vendido_v
                )

            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)

        nome = f"Relatorios_Vendedores_{date.today().strftime('%d%m%Y')}.xlsx"
        st.success("✅ Relatórios gerados com sucesso!")
        st.download_button(
            label="⬇️ Baixar Excel com todos os vendedores",
            data=buf,
            file_name=nome,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )