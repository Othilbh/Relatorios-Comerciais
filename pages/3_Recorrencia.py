# pages/3_Recorrencia.py
import streamlit as st
from datetime import date
import sys, os, tempfile
import pandas as pd
import io
sys.path.append(os.path.dirname(os.path.dirname(__file__)))

from utils.parser import extrair_dados_pdf, validar_totais
from utils.consolidacao import normalizar_cliente, calcular_margem
from utils.categorias import categorizar_produto, ORDEM_COLUNAS, margem_real
from utils.builder import criar_aba_recorrencia
from openpyxl import Workbook

st.set_page_config(page_title="Recorrência · OTHIL", page_icon="🔄", layout="wide")
st.title("🔄 Recorrência")
st.caption("Matriz de produtos por cliente. Upload dos PDFs dos vendedores.")
st.divider()

col1, col2 = st.columns([2, 1])
with col1:
    uploaded_files = st.file_uploader(
        "📂 Selecione os PDFs",
        type=["pdf"],
        accept_multiple_files=True
    )
with col2:
    data_ref = st.date_input("📅 Período de referência", value=date.today())

if uploaded_files:
    todos_clientes = []

    with st.spinner("🔍 Processando PDFs..."):
        for uf in uploaded_files:
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                tmp.write(uf.read())
                tmp_path = tmp.name
            dados = extrair_dados_pdf(tmp_path)
            os.unlink(tmp_path)

            ok, msg = validar_totais(dados)
            vendedor = dados.get("vendedor", uf.name)
            if ok:
                st.success(f"✅ {vendedor}")
            else:
                st.warning(f"⚠️ {vendedor}: {msg}")

            todos_clientes.extend(dados.get("clientes", []))

    if todos_clientes:
        matriz = {}
        for item in todos_clientes:
            nome = normalizar_cliente(item["cliente"])
            if nome not in matriz:
                matriz[nome] = {}
            for prod in item.get("produtos", []):
                cat = categorizar_produto(prod.get("descricao", ""))
                if cat not in matriz[nome]:
                    matriz[nome][cat] = {"qtd": 0.0, "faturamento": 0.0, "custo": 0.0}
                matriz[nome][cat]["qtd"] += prod.get("qtd", 0)
                matriz[nome][cat]["faturamento"] += prod.get("faturamento", 0)
                custo = prod.get("custo_total", prod.get("qtd", 0) * prod.get("custo_unit", 0))
                matriz[nome][cat]["custo"] += custo

        cats_usadas = [c for c in ORDEM_COLUNAS if any(c in m for m in matriz.values())]
        if any("OUTROS" in m for m in matriz.values()):
            cats_usadas.append("OUTROS")

        st.divider()
        st.subheader("📊 Matriz de Recorrência")

        rows = []
        for cliente in sorted(matriz.keys()):
            row = {"Cliente": cliente}
            fat_t = custo_t = 0.0
            for cat in cats_usadas:
                d = matriz[cliente].get(cat)
                if d and d["qtd"] > 0:
                    row[cat] = round(d["qtd"], 1)
                    fat_t += d["faturamento"]
                    custo_t += d["custo"]
                else:
                    row[cat] = ""
            row["Margem Real %"] = round(margem_real(calcular_margem(fat_t, custo_t)), 2)
            rows.append(row)

        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

        st.divider()
        with st.spinner("⚙️ Gerando Excel..."):
            wb = Workbook()
            wb.remove(wb.active)
            criar_aba_recorrencia(wb, todos_clientes)
            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)

        st.download_button(
            label="⬇️ Baixar Excel",
            data=buf,
            file_name=f"Recorrencia_{data_ref.strftime('%d%m%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )