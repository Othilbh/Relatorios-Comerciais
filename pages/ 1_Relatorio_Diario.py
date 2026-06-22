# pages/1_Relatorio_Diario.py

import streamlit as st
from datetime import date
import sys, os
sys.path.append(os.path.dirname(os.path.dirname(__file__)))

from utils.parser import extrair_dados_pdf, validar_totais
from utils.builder import gerar_excel
import tempfile

st.set_page_config(page_title="Relatório Diário · OTHIL", page_icon="📋", layout="wide")

st.title("📋 Relatório Diário")
st.caption("Faça upload do PDF de Lucratividade do dia para gerar o Excel completo.")

st.divider()

col1, col2 = st.columns([2, 1])
with col1:
    uploaded = st.file_uploader("📂 Selecione o PDF do dia", type=["pdf"])
with col2:
    data_ref = st.date_input("📅 Data de referência", value=date.today())

if uploaded:
    with st.spinner("🔍 Lendo PDF..."):
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp.write(uploaded.read())
            tmp_path = tmp.name

        dados = extrair_dados_pdf(tmp_path)
        os.unlink(tmp_path)

    ok, msg_val = validar_totais(dados)
    if ok:
        st.success(msg_val)
    else:
        st.warning(msg_val)

    vendedor = dados.get("vendedor", "—")
    clientes = dados.get("clientes", [])
    total_fat = dados.get("total_faturamento", 0)
    total_custo = dados.get("total_custo", 0)

    st.subheader(f"Vendedor: {vendedor}")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Clientes", len(clientes))
    m2.metric("Faturamento Total", f"R$ {total_fat:,.2f}")
    m3.metric("Custo Total", f"R$ {total_custo:,.2f}")
    margem_geral = (total_fat - total_custo) / total_fat * 100 if total_fat else 0
    m4.metric("Margem %", f"{margem_geral:.2f}%")

    if clientes:
        with st.expander("📊 Ver tabela de clientes"):
            import pandas as pd
            from utils.consolidacao import calcular_margem, normalizar_cliente
            rows = []
            for c in clientes:
                nome = normalizar_cliente(c["cliente"])
                fat = c["faturamento"]
                custo = c["custo"]
                rows.append({
                    "Cliente": nome,
                    "Volume CX": c["volume"],
                    "Faturamento R$": fat,
                    "Custo R$": custo,
                    "Margem %": round(calcular_margem(fat, custo), 2),
                })
            df = pd.DataFrame(rows)
            st.dataframe(df, use_container_width=True, hide_index=True)

    st.divider()
    with st.spinner("⚙️ Gerando Excel..."):
        buf = gerar_excel([dados], data_ref=data_ref.strftime("%d/%m/%Y"))

    nome_arquivo = f"Relatorio_Diario_{data_ref.strftime('%d%m%Y')}.xlsx"
    st.download_button(
        label="⬇️ Baixar Excel",
        data=buf,
        file_name=nome_arquivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )