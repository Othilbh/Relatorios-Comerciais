# app.py — Página inicial do app OTHIL Relatórios Comerciais

import streamlit as st

st.set_page_config(
    page_title="OTHIL — Relatórios Comerciais",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
    <style>
        .main-title {
            font-size: 2.2rem;
            font-weight: 800;
            color: #1A3A5C;
            margin-bottom: 0;
        }
        .sub-title {
            font-size: 1rem;
            color: #666;
            margin-top: 0;
        }
        .card {
            background: #f8f9fa;
            border-left: 4px solid #1A3A5C;
            border-radius: 8px;
            padding: 1rem 1.5rem;
            margin-bottom: 1rem;
        }
        .card h3 { color: #1A3A5C; margin-bottom: 0.3rem; }
        .card p  { color: #444; margin: 0; font-size: 0.92rem; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-title">📊 OTHIL — Relatórios Comerciais</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Sistema automatizado de análise de lucratividade · CEASA Minas · BH</p>', unsafe_allow_html=True)

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    <div class="card">
        <h3>📋 Relatório Diário</h3>
        <p>Faça upload do PDF diário e gere o Excel completo com faturamento, custo e margem por cliente e produto.</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="card">
        <h3>👤 Vendedor-Cliente</h3>
        <p>Upload dos PDFs por vendedor. Gera análise de lucratividade consolidada com grupos de clientes.</p>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="card">
        <h3>🔄 Recorrência</h3>
        <p>Matriz de produtos × clientes mostrando quais categorias cada cliente compra e a margem real.</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="card">
        <h3>🎯 Metas Semanais</h3>
        <p>Acompanhe o atingimento semanal por vendedor e produto com status colorido e % de atingimento.</p>
    </div>
    """, unsafe_allow_html=True)

st.divider()
st.caption("Use o menu lateral para navegar entre os relatórios. · OTHIL Distribuidora de Frutas")