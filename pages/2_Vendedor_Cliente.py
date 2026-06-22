# pages/2_Vendedor_Cliente.py
import streamlit as st
from datetime import date
import sys, os, tempfile
import pandas as pd
sys.path.append(os.path.dirname(os.path.dirname(__file__)))

from utils.parser import extrair_dados_pdf, validar_totais
from utils.builder import gerar_excel
from utils.consolidacao import calcular_margem

st.set_page_config(page_title="Vendedor-Cliente · OTHIL", page_icon="👤", layout="wide")
st.title("👤 Vendedor-Cliente")
st.caption("Faça upload dos PDFs por vendedor. Consolida grupos de clientes automaticamente.")
st.divider()

col1, col2 = st.columns([2, 1])
with col1:
    uploaded_files = st.file_uploader(
        "📂 Selecione os PDFs dos vendedores",
        type=["pdf"],
        accept_multiple_files=True
    )
with col2:
    data_ref = st.date_input("📅 Data de referência", value=date.today())

if uploaded_files:
    todos_dados = []

    with st.spinner(f"🔍 Processando {len(uploaded_files)} PDF(s)..."):
        for uf in uploaded_files:
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                tmp.write(uf.read())
                tmp_path = tmp.name
            dados = extrair_dados_pdf(tmp_path)
            os.unlink(tmp_path)

            ok, msg = validar_totais(dados)
            vendedor = dados.get("vendedor", uf.name)
            if ok:
                st.success(f"✅ {vendedor}: {msg}")
            else:
                st.warning(f"⚠️ {vendedor}: {msg}")

            if dados.get("clientes"):
                todos_dados.append(dados)

    if todos_dados:
        st.divider()
        st.subheader("📊 Resumo Consolidado")

        fat_total = sum(d.get("total_faturamento", 0) for d in todos_dados)
        custo_total = sum(d.get("total_custo", 0) for d in todos_dados)
        margem_geral = calcular_margem(fat_total, custo_total)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Vendedores", len(todos_dados))
        m2.metric("Faturamento Total", f"R$ {fat_total:,.2f}")
        m3.metric("Custo Total", f"R$ {custo_total:,.2f}")
        m4.metric("Margem Geral %", f"{margem_geral:.2f}%")

        rows = []
        for d in todos_dados:
            fat = d.get("total_faturamento", 0)
            custo = d.get("total_custo", 0)
            rows.append({
                "Vendedor": d.get("vendedor", "—"),
                "Clientes": len(d.get("clientes", [])),
                "Faturamento R$": round(fat, 2),
                "Custo R$": round(custo, 2),
                "Margem %": round(calcular_margem(fat, custo), 2),
            })
        df = pd.DataFrame(rows).sort_values("Faturamento R$", ascending=False)
        st.dataframe(df, use_container_width=True, hide_index=True)

        st.divider()
        with st.spinner("⚙️ Gerando Excel..."):
            buf = gerar_excel(todos_dados, data_ref=data_ref.strftime("%d/%m/%Y"))

        st.download_button(
            label="⬇️ Baixar Excel",
            data=buf,
            file_name=f"Vendedor_Cliente_{data_ref.strftime('%d%m%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )