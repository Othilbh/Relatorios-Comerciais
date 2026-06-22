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

PRODUTOS_DEFAULT = [
    "Portuguesa Sabor 55/60",
    "Portuguesa 60/70",
    "Portuguesa 70/80",
    "Forelle",
    "Ercoline",
    "Pera Asiática",
    "Gala Santa Carol",
    "Gala Azaleia",
    "Fuji Expressa 180",
    "Fuji Azaleia",
    "Fuji Hiragami",
    "Fuji Suprema",
    "Thompson Vitace",
    "Thompson Seedless",
    "Uva Crimson",
    "Uva Isis",
    "Uva Jubilee",
    "Uva Itália",
    "Maçã Argentina",
    "Maçã Chilena",
    "Maçã Pink Lady",
    "Maçã Granny Smith",
    "Maçã Red Globe",
    "Mamão Havai",
    "Mamão Formoso",
    "Goiaba",
    "Melão Amarelo",
    "Melão Galia",
    "Melão Cantaloupe",
    "Tangerina Cumbuca",
    "Tangerina Ponkan",
    "Ameixa",
    "Pêssego",
    "Nectarina",
    "Morango",
    "Mirtilo",
    "Abacaxi",
    "Manga Palmer",
    "Manga Tommy",
    "Laranja",
    "Limão Tahiti",
    "Limão Siciliano",
    "Tomate Roma",
]

MAPA_PRODUTO = {
    "Portuguesa Sabor 55/60": ["PORTUGUESA SABOR", "PORTUGUESA 55", "PORTUGUESA SAB"],
    "Portuguesa 60/70": ["PORTUGUESA 60"],
    "Portuguesa 70/80": ["PORTUGUESA 70"],
    "Forelle": ["FORELLE"],
    "Ercoline": ["ERCOLINE"],
    "Pera Asiática": ["ASIATICA", "ASIÁTICA", "PERA ASIAT"],
    "Gala Santa Carol": ["GALA SANTA", "SANTA CAROL"],
    "Gala Azaleia": ["GALA AZALEIA", "GALA AZ"],
    "Fuji Expressa 180": ["FUJI EXPRESSA", "FUJI 180"],
    "Fuji Azaleia": ["FUJI AZALEIA"],
    "Fuji Hiragami": ["HIRAGAMI", "HIRAGAMIS"],
    "Fuji Suprema": ["FUJI SUPREMA"],
    "Thompson Vitace": ["THOMPSON VITACE", "THOMPSON VIT"],
    "Thompson Seedless": ["THOMPSON SEEDLESS", "THOMPSON SEED"],
    "Uva Crimson": ["CRINSON", "CRIMSON"],
    "Uva Isis": ["ISIS"],
    "Uva Jubilee": ["JUBILEE"],
    "Uva Itália": ["ITALIA", "ITÁLIA"],
    "Maçã Argentina": ["ARGENTINA"],
    "Maçã Chilena": ["CHILENA"],
    "Maçã Pink Lady": ["PINK LADY"],
    "Maçã Granny Smith": ["GRAN SMITH", "GRANNY SMITH"],
    "Maçã Red Globe": ["RED GLOBE"],
    "Mamão Havai": ["MAMAO HAVAI", "MAMÃO HAVAI", "MAMAO HAVAÍ", "MAMÃO HAVAÍ"],
    "Mamão Formoso": ["MAMAO FORMOSO", "MAMÃO FORMOSO"],
    "Goiaba": ["GOIABA"],
    "Melão Amarelo": ["MELAO AMARELO", "MELÃO AMARELO"],
    "Melão Galia": ["MELAO GALIA", "MELÃO GALIA", "MELÃO GAÍA", "GALIA"],
    "Melão Cantaloupe": ["CANTALOUPE"],
    "Tangerina Cumbuca": ["CUMBUCA"],
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

st.caption("Preencha as metas (CX) por produto. Adicione linhas para novos produtos se necessário:")

df_metas_default = pd.DataFrame(
    {v: [0.0] * len(PRODUTOS_DEFAULT) for v in VENDEDORES_ATIVOS},
    index=PRODUTOS_DEFAULT
)
df_metas_default.index.name = "Produto"
df_metas_default = df_metas_default.reset_index()

df_metas = st.data_editor(
    df_metas_default,
    use_container_width=True,
    num_rows="dynamic",
    key="editor_metas"
)

st.divider()
st.subheader("2️⃣ Upload do PDF de Vendas Acumuladas")
uploaded_pdf = st.file_uploader("📂 PDF de vendas acumuladas", type=["pdf"])

vendido = {v: {} for v in VENDEDORES_ATIVOS}

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
                vendido[vendedor_pdf][prod_meta] = vendido[vendedor_pdf].get(prod_meta, 0) + prod.get("qtd", 0)

    st.success(f"✅ Dados de {vendedor_pdf} carregados.")

st.divider()
st.subheader("3️⃣ Resultado das Metas")

if st.button("📊 Calcular Metas", use_container_width=True, type="primary"):
    def status(pct):
        if pct >= 100: return "✅ Atingida"
        if pct >= 50: return "⚠️ Em andamento"
        return "❌ Abaixo"

    resultados = []
    for _, row in df_metas.iterrows():
        produto = row.get("Produto", "")
        if not produto:
            continue
        for v in VENDEDORES_ATIVOS:
            meta = float(row.get(v, 0) or 0)
            if meta == 0:
                continue
            vend = vendido.get(v, {}).get(produto, 0.0)
            pct = (vend / meta * 100) if meta > 0 else 0.0
            resultados.append({
                "Vendedor": v, "Produto": produto,
                "Meta CX": round(meta, 1), "Vendido CX": round(vend, 1),
                "Falta CX": round(max(meta - vend, 0), 1),
                "% Atingido": round(pct, 1), "Status": status(pct),
            })

    if not resultados:
        st.warning("Nenhuma meta preenchida. Preencha as metas antes de calcular.")
    else:
        df_result = pd.DataFrame(resultados)

        for v in VENDEDORES_ATIVOS:
            df_v = df_result[df_result["Vendedor"] == v].drop(columns=["Vendedor"])
            if df_v.empty:
                continue
            with st.expander(f"📌 {v}", expanded=False):
                st.dataframe(df_v, use_container_width=True, hide_index=True)

        st.divider()

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
            if df_v.empty:
                continue
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