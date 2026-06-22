# pages/4_Metas_Semanais.py
import streamlit as st
from datetime import date, timedelta
import sys, os, tempfile
import pandas as pd
import io
sys.path.append(os.path.dirname(os.path.dirname(__file__)))

from utils.parser import extrair_dados_pdf, VENDEDORES_ATIVOS
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

st.set_page_config(page_title="Metas Semanais · OTHIL", page_icon="🎯", layout="wide")
st.title("🎯 Metas Semanais")
st.caption("Acompanhe o atingimento semanal por vendedor e produto. Luca excluído automaticamente.")
st.divider()

PRODUTOS_META = [
    "Portuguesa 55/60", "Forelle", "Gala Santa Carol",
    "Fuji Expressa 180", "Fuji Azaleia", "Thompson Vitace",
    "Mamão", "Goiaba", "Melão Amarelo", "Melão Gaia", "Tangerina Cumbuca",
]

MAPA_PRODUTO = {
    "Portuguesa 55/60": ["PORTUGUESA"],
    "Forelle": ["FORELLE"],
    "Gala Santa Carol": ["GALA", "SANTA CAROL"],
    "Fuji Expressa 180": ["FUJI EXPRESSA", "FUJI 180"],
    "Fuji Azaleia": ["FUJI AZALEIA", "AZALEIA"],
    "Thompson Vitace": ["THOMPSON"],
    "Mamão": ["MAMAO", "MAMÃO"],
    "Goiaba": ["GOIABA"],
    "Melão Amarelo": ["MELAO AMARELO", "MELÃO AMARELO"],
    "Melão Gaia": ["MELAO GALIA", "MELÃO GALIA", "GAIA"],
    "Tangerina Cumbuca": ["CUMBUCA", "TANGERINA"],
}

def mapear_produto(desc):
    desc_u = desc.upper()
    for prod_meta, keywords in MAPA_PRODUTO.items():
        for kw in keywords:
            if kw in desc_u:
                return prod_meta
    return None

st.subheader("1️⃣ Definir Metas")
col_ini, col_fim = st.columns(2)
with col_ini:
    semana_ini = st.date_input("Início da semana", value=date.today() - timedelta(days=date.today().weekday()))
with col_fim:
    semana_fim = st.date_input("Fim da semana", value=semana_ini + timedelta(days=5))

st.caption("Preencha as metas (CX) por produto para cada vendedor:")
df_metas = st.data_editor(
    pd.DataFrame({v: [0.0] * len(PRODUTOS_META) for v in VENDEDORES_ATIVOS}, index=PRODUTOS_META),
    use_container_width=True,
    num_rows="fixed",
    key="editor_metas"
)

st.divider()
st.subheader("2️⃣ Upload do PDF de Vendas Acumuladas")
uploaded_pdf = st.file_uploader("📂 PDF de vendas acumuladas", type=["pdf"])

vendido = {v: {p: 0.0 for p in PRODUTOS_META} for v in VENDEDORES_ATIVOS}

if uploaded_pdf:
    with st.spinner("🔍 Lendo PDF..."):
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp.write(uploaded_pdf.read())
            tmp_path = tmp.name
        dados = extrair_dados_pdf(tmp_path)
        os.unlink(tmp_path)

    vendedor_pdf = dados.get("vendedor", "")
    for cli in dados.get("clientes", []):
        for prod in cli.get("produtos", []):
            prod_meta = mapear_produto(prod.get("descricao", ""))
            if prod_meta and vendedor_pdf in vendido:
                vendido[vendedor_pdf][prod_meta] += prod.get("qtd", 0)

    st.success(f"✅ Dados de {vendedor_pdf} carregados.")

st.divider()
st.subheader("3️⃣ Resultado das Metas")

if st.button("📊 Calcular Metas", use_container_width=True, type="primary"):
    def status(pct):
        if pct >= 100: return "✅ Atingida"
        if pct >= 50: return "⚠️ Em andamento"
        return "❌ Abaixo"

    resultados = []
    for v in VENDEDORES_ATIVOS:
        for p in PRODUTOS_META:
            meta = float(df_metas.at[p, v]) if p in df_metas.index and v in df_metas.columns else 0.0
            vend = vendido.get(v, {}).get(p, 0.0)
            pct = (vend / meta * 100) if meta > 0 else 0.0
            resultados.append({
                "Vendedor": v, "Produto": p,
                "Meta CX": round(meta, 1), "Vendido CX": round(vend, 1),
                "Falta CX": round(max(meta - vend, 0), 1),
                "% Atingido": round(pct, 1), "Status": status(pct),
            })

    df_result = pd.DataFrame(resultados)

    for v in VENDEDORES_ATIVOS:
        df_v = df_result[df_result["Vendedor"] == v].drop(columns=["Vendedor"])
        with st.expander(f"📌 {v}", expanded=False):
            st.dataframe(df_v, use_container_width=True, hide_index=True)

    st.divider()

    # Gerar Excel
    thin = Side(style="thin", color="BFBFBF")
    brd = Border(left=thin, right=thin, top=thin, bottom=thin)
    COR_H = "1A3A5C"; COR_S = "2E6DA4"
    COR_V = "C6EFCE"; COR_A = "FFEB9C"; COR_R = "FFC7CE"
    COR_VF = "276221"; COR_AF = "7D6608"; COR_RF = "9C0006"

    def hc(ws, r, c, v, bg=COR_H, fg="FFFFFF"):
        cell = ws.cell(row=r, column=c, value=v)
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.font = Font(bold=True, color=fg, size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = brd
        return cell

    wb = Workbook()
    ws_r = wb.active
    ws_r.title = "Resumo Geral"
    periodo = f"{semana_ini.strftime('%d/%m/%Y')} a {semana_fim.strftime('%d/%m/%Y')}"
    ws_r.merge_cells("A1:G1")
    hc(ws_r, 1, 1, f"METAS SEMANAIS — {periodo}")
    for ci, h in enumerate(["Vendedor","Produto","Meta CX","Vendido CX","Falta CX","% Atingido","Status"], 1):
        hc(ws_r, 2, ci, h, bg=COR_S)

    for ri, row_data in enumerate(df_result.itertuples(index=False), start=3):
        pct = row_data[5]
        bg = COR_V if pct >= 100 else COR_A if pct >= 50 else COR_R
        fg = COR_VF if pct >= 100 else COR_AF if pct >= 50 else COR_RF
        for ci, val in enumerate(row_data, 1):
            c = ws_r.cell(row=ri, column=ci, value=val)
            c.fill = PatternFill("solid", fgColor=bg)
            c.font = Font(color=fg if ci >= 3 else "000000")
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = brd

    for v in VENDEDORES_ATIVOS:
        df_v = df_result[df_result["Vendedor"] == v]
        ws_v = wb.create_sheet(v[:31])
        ws_v.merge_cells("A1:F1")
        hc(ws_v, 1, 1, f"{v.upper()} — {periodo}")
        for ci, h in enumerate(["Produto","Meta CX","Vendido CX","Falta CX","% Atingido","Status"], 1):
            hc(ws_v, 2, ci, h, bg=COR_S)
        for ri, row_data in enumerate(df_v.drop(columns=["Vendedor"]).itertuples(index=False), start=3):
            pct = row_data[4]
            bg = COR_V if pct >= 100 else COR_A if pct >= 50 else COR_R
            fg = COR_VF if pct >= 100 else COR_AF if pct >= 50 else COR_RF
            for ci, val in enumerate(row_data, 1):
                c = ws_v.cell(row=ri, column=ci, value=val)
                c.fill = PatternFill("solid", fgColor=bg)
                c.font = Font(color=fg if ci >= 2 else "000000")
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = brd

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    st.download_button(
        label="⬇️ Baixar Excel de Metas",
        data=buf,
        file_name=f"Metas_{semana_ini.strftime('%d%m')}_{semana_fim.strftime('%d%m%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )