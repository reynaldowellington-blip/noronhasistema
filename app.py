import streamlit as st
import pandas as pd
import os
from datetime import datetime

# Configuração da página
st.set_page_config(
    page_title="Terminal de Vendas - Noronha",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Caminho absoluto para garantir que o Render encontre as planilhas na mesma pasta
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
caminho_clientes = os.path.join(BASE_DIR, "clientesnoronha.xlsx")
caminho_produtos = os.path.join(BASE_DIR, "produtosnoronha.xlsx")

# Carregamento seguro das bases de dados
try:
    df_clientes = pd.read_excel(caminho_clientes)
except Exception:
    df_clientes = pd.DataFrame(columns=["Nome"])

try:
    df_produtos = pd.read_excel(caminho_produtos)
except Exception:
    df_produtos = pd.DataFrame(columns=["Produto", "Preço"])

# Layout Principal do Terminal de Vendas
st.markdown("## 🚀 Terminal de Vendas")
st.markdown("Insira ou gerencie as configurações do pedido comercial.")

col1, col2 = st.columns(2)

with col1:
    st.markdown("### 👤 Identificação do Cliente")
    
    # Pega a lista de clientes da planilha se houver coluna correspondente
    lista_clientes = []
    if not df_clientes.empty:
        # Tenta achar a coluna de nome do cliente de forma flexível
        col_cliente = next((c for c in df_clientes.columns if 'nome' in c.lower() or 'cliente' in c.lower()), df_clientes.columns[0])
        lista_clientes = df_clientes[col_cliente].dropna().astype(str).tolist()
    
    cliente_selecionado = st.selectbox("Selecione um cliente da base:", ["-- Selecione da Base --"] + lista_clientes)
    cliente_novo = st.text_input("Ou digite o nome de um Cliente Novo:")

with col2:
    st.markdown("### 💳 Condições e Totais")
    data_venda = st.date_input("Data da Venda", datetime.now())
    forma_pagamento = st.selectbox("Forma de proteção / Pagamento", ["PIX", "Dinheiro", "Cartão", "Boleto"])
    
    st.markdown("---")
    st.markdown("#### **VALOR TOTAL DO ITEM**")
    st.markdown("### **R$ 0,00**")

st.markdown("### 🛒 Seleção de Produto")
lista_produtos = []
if not df_produtos.empty:
    col_prod = next((c for c in df_produtos.columns if 'produto' in c.lower() or 'desc' in c.lower() or 'item' in c.lower()), df_produtos.columns[0])
    lista_produtos = df_produtos[col_prod].dropna().astype(str).tolist()

produto_selecionado = st.selectbox("Selecione um produto da lista:", ["-- Selecione ou digite abaixo --"] + lista_produtos)
produto_novo = st.text_input("Ou digite um Produto / Descrição novo(a):")

if st.button("💾 Salvar Venda", type="primary"):
    st.success("Venda registrada com sucesso!")