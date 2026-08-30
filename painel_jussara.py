import streamlit as st
import pandas as pd
import os

# Configuração da página do Painel Gerencial
st.set_page_config(
    page_title="Next - Painel Gerencial",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilização visual em tons escuros e verde tecnológico (igual ao print)
st.markdown("""
    <style>
    .main {
        background-color: #0e1117;
    }
    .stMetric {
        background-color: #1a1c23;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #262730;
    }
    .stMetric [data-testid="stMetricValue"] {
        color: #10b981 !important;
    }
    h1, h2, h3 {
        font-family: 'Inter', sans-serif;
        color: #ffffff;
    }
    </style>
""", unsafe_allow_html=True)

# Sistema de Segurança / Senha para a Jussara
st.sidebar.title("🔐 Acesso Restrito")
senha = st.sidebar.text_input("Senha da Gerência:", type="password")

# Defina a senha da Jussara aqui
SENHA_GERENTE = "jussara2026"

if senha != SENHA_GERENTE:
    st.warning("⚠️ Esta área é restrita. Digite a senha gerencial na barra lateral para acessar o painel.")
    st.stop()  # Para a execução aqui e protege os dados dos vendedores

# Cabeçalho do Painel
st.title("📊 Painel Gerencial - Next")
st.markdown("Acompanhamento em tempo real das operações e faturamento.")

# Lendo os dados (aqui ele lê do arquivo Excel local ou da nuvem que os vendedores preenchem)
ARQUIVO_DADOS = "vendas.xlsx"  # Altere para o nome do arquivo que os vendedores usam

if os.path.exists(ARQUIVO_DADOS):
    df = pd.read_excel(ARQUIVO_DADOS)
else:
    # Dados de exemplo caso o arquivo ainda esteja vazio
    df = pd.DataFrame({
        "ID": [1, 2, 3, 4],
        "Produto": ["Produto A", "Produto B", "Produto A", "Produto C"],
        "Quantidade": [10, 5, 8, 20],
        "Valor Total": [1200.0, 750.0, 960.0, 1500.0],
        "Pagamento": ["Pix", "Cartão", "Pix", "Boleto"]
    })

# Métricas Principais (KPIs) no topo
col1, col2, col3 = st.columns(3)

faturamento_total = df["Valor Total"].sum() if "Valor Total" in df.columns else 0
total_vendas = len(df)
ticket_medio = faturamento_total / total_vendas if total_vendas > 0 else 0

with col1:
    st.metric(label="💰 Faturamento Total", value=f"R$ {faturamento_total:,.2f}")
with col2:
    st.metric(label="📦 Total de Vendas", value=total_vendas)
with col3:
    st.metric(label="📈 Ticket Médio", value=f"R$ {ticket_medio:,.2f}")

st.markdown("---")

# Seção de Gráficos e Análises (Mais e Menos Vendidos)
col_graf1, col_graf2 = st.columns(2)

with col_graf1:
    st.subheader("🔥 Produtos Mais Vendidos")
    if "Produto" in df.columns and "Quantidade" in df.columns:
        mais_vendidos = df.groupby("Produto")["Quantidade"].sum().reset_index()
        st.bar_chart(mais_vendidos.set_index("Produto"))
    else:
        st.info("Aguardando dados de produtos...")

with col_graf2:
    st.subheader("💳 Formas de Pagamento")
    if "Pagamento" in df.columns and "Valor Total" in df.columns:
        pagamentos = df.groupby("Pagamento")["Valor Total"].sum().reset_index()
        st.bar_chart(pagamentos.set_index("Pagamento"))
    else:
        st.info("Aguardando dados de pagamento...")

# Tabela Detalhada com os Registros
st.markdown("---")
st.subheader("📋 Detalhamento das Últimas Vendas")
st.dataframe(df, use_container_width=True)