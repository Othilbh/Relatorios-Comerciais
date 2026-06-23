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

# ── GitHub Storage ────────────────────────────────────────────────────────
GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", "")
GITHUB_REPO  = st.secrets.get("GITHUB_REPO", "Othilbh/Relatorios-Comerciais")
GITHUB_FILE_METAS = "data/metas_semana.json"
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

# ── Carrega dados salvos ──────────────────────────────────────────────────
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

# ── Cores Excel ───────────────────────────────────────────────────────────
COR_H  = "1A3A5C"
COR_S  = "2E6DA4"
COR_V  = "C6EFCE"; COR_VF = "276221"
COR_A  = "FFEB9C"; COR_AF = "7D6608"
COR_R  = "FFC7CE"; COR_RF = "9C0006"
COR_TITULO = "F5A623"

def _fill(c): return PatternFill("solid", fgColor=c)
def _font(bold=False, color="000000", size=10): return Font(bold=bold, color=color, size=size, name="Calibri")
def _border_thin():
    t = Side(style="thin", color="000000")
    return Border(left=t, right=t, top=t, bottom=t)
def _border_medium():
    m = Side(style="medium", color="000000")
    return Border(left=m, right=m, top=m, bottom=m)
def _alinhar(h="center", wrap=False):
    return Alignment(horizontal=h, vertical="center", wrap_text=wrap)

def gerar_aba_vendedor(wb, vendedor, data_ref, itens_estoque, metas, vendido_v):
    ws = wb.create_sheet(vendedor[:31])
    ws.sheet_view.showGridLines = False

    # ── Cabeçalho ──────────────────────────────────────────────────────
    ws.merge_cells("A1:B1")
    c1 = ws["A1"]
    c1.value = f"Vendedor :{vendedor.upper()}"
    c1.font = _font(bold=True, size=11)
    c1.border = _border_medium()
    c1.alignment = _alinhar("left")

    ws["A2"] = f"Data {data_ref}"
    ws["A2"].font = _font(size=10)

    # ── Tabela Estoque ──────────────────────────────────────────────────
    hdrs = ["Produto", "Complemento", "Data Entrada", "Saldo Atual",
            "Qtde Vendida", "Custo Unitario", "Md Venda"]

    row = 4
    for ci, h in enumerate(hdrs, 1):
        c = ws.cell(row=row, column=ci, value=h)
        c.fill = _fill("D9D9D9")
        c.font = _font(bold=True, size=10)
        c.alignment = _alinhar("center")
        c.border = _border_thin()

    row = 5
    for item in itens_estoque:
        vals = [
            item["descricao"], item["complemento"], item["data_entrada"],
            item["saldo_atual"], item["qtd_vendida"], item["custo"], item["md_venda"]
        ]
        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=row, column=ci, value=val)
            c.font = _font(size=10)
            c.border = _border_thin()
            if ci == 1:
                c.alignment = _alinhar("left")
            elif ci in [6, 7]:
                c.number_format = 'R$ #,##0.00'
                c.alignment = _alinhar("right")
                # Zera em vermelho
                if val == 0:
                    c.font = _font(color="FF0000", size=10)
            elif ci == 5 and val == 0:
                c.font = _font(color="FF0000", size=10)
                c.alignment = _alinhar("center")
            else:
                c.alignment = _alinhar("center")
        row += 1

    row += 1  # linha em branco

    # ── Tabela Metas ────────────────────────────────────────────────────
    if metas:
        ws.merge_cells(f"A{row}:G{row}")
        c_titulo = ws[f"A{row}"]
        c_titulo.value = f"METAS SEMANAIS — {vendedor.upper()}"
        c_titulo.fill = _fill(COR_H)
        c_titulo.font = _font(bold=True, color="FFFFFF", size=11)
        c_titulo.alignment = _alinhar("center")
        c_titulo.border = _border_medium()
        row += 1

        hdrs_meta = ["Produto", "Meta (cx)", "Vendido (cx)", "Falta (cx)", "% Atingido"]
        larguras_meta = [5, 1, 1, 1, 1]
        for ci, h in enumerate(hdrs_meta, 1):
            c = ws.cell(row=row, column=ci, value=h)
            c.fill = _fill(COR_S)
            c.font = _font(bold=True, color="FFFFFF", size=10)
            c.alignment = _alinhar("center")
            c.border = _border_thin()
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

            if pct >= 100:   bg_m, fg_m = COR_V,  COR_VF
            elif pct >= 50:  bg_m, fg_m = COR_A,  COR_AF
            else:            bg_m, fg_m = COR_R,  COR_RF

            vals_meta = [produto, meta, round(vend, 1), round(falta, 1), f"{pct:.2f}%"]
            for ci, val in enumerate(vals_meta, 1):
                c = ws.cell(row=row, column=ci, value=val)
                c.fill = _fill(bg_m)
                c.font = _font(color=fg_m, size=10)
                c.border = _border_thin()
                c.alignment = _alinhar("left" if ci == 1 else "center")
            row += 1

        # Total
        pct_total = (total_vend / total_meta * 100) if total_meta > 0 else 0
        if pct_total >= 100:   bg_t, fg_t = COR_V,  COR_VF
        elif pct_total >= 50:  bg_t, fg_t = COR_A,  COR_AF
        else:                  bg_t, fg_t = COR_R,  COR_RF

        vals_total = ["TOTAL", total_meta, round(total_vend, 1),
                      round(max(total_meta - total_vend, 0), 1), f"{pct_total:.2f}%"]
        for ci, val in enumerate(vals_total, 1):
            c = ws.cell(row=row, column=ci, value=val)
            c.fill = _fill(bg_t)
            c.font = _font(bold=True, color=fg_t, size=10)
            c.border = _border_thin()
            c.alignment = _alinhar("left" if ci == 1 else "center")
        row += 1

    row += 2

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
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 13
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 16
    ws.column_dimensions["G"].width = 14

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
    semana_fim = st.date_input("Fim",
        value=semana_ini + timedelta(days=5))

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
                st.success("✅ Metas salvas com sucesso!")
            else:
                st.error("❌ Erro ao salvar. Verifique o token.")
    with col_apagar:
        if st.button("🗑️ Apagar Metas", type="secondary", use_container_width=True):
            _, sha = github_get(GITHUB_FILE_METAS)
            github_save(GITHUB_FILE_METAS, [], sha)
            st.session_state.produtos_meta = []
            st.session_state.vendido = {}
            st.session_state.estoque = {}
            st.session_state.sem_vendedor = []
            st.rerun()
    with col_info:
        st.caption("💡 Clique em **Salvar Metas** toda segunda-feira. Na semana seguinte, apague e salve novamente.")

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
st.subheader("3️⃣ Gerar Relatórios por Vendedor")

if st.button("📋 Gerar Relatórios", use_container_width=True, type="primary"):
    if not st.session_state.estoque:
        st.warning("Faça upload do PDF de Estoque primeiro.")
    else:
        with st.spinner("⚙️ Gerando Excel..."):
            wb = Workbook()
            wb.remove(wb.active)
            data_ref = date.today().strftime("%d/%m")

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
        st.success("✅ Relatórios gerados!")
        st.download_button(
            label="⬇️ Baixar Excel — todos os vendedores",
            data=buf, file_name=nome,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )