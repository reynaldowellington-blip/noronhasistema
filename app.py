import streamlit as st
import pandas as pd
import os
from datetime import datetime

# Configuração da página com tema moderno e centralizado
st.set_page_config(
    page_title="Next",
    page_icon="⚡",
    layout="wide"

    initial_sidebar_state="expanded"
)

# Estilização CSS personalizada com os valores em verde tecnológico
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
        /* Deixa o número da métrica em verde brilhante */
        .stMetric [data-testid="stMetricValue"] {
            color: #10b981 !important;
        }
        h1, h2, h3 {
            font-family: 'Inter', sans-serif;
            letter-spacing: -0.5px;
        }
        .stButton>button {
            width: 100%;
            border-radius: 8px;
            font-weight: 600;
            background: linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%);
            color: white;
            border: none;
            padding: 0.6rem 1rem;
            transition: all 0.3s ease;
        }
        .stButton>button:hover {
            opacity: 0.9;
            box-shadow: 0 4px 12px rgba(59, 130, 246, 0.4);
        }
    </style>
""", unsafe_allow_html=True)

# 1. Carregar Clientes da planilha específica
ARQUIVO_CLIENTES = "clientesnext.xlsx"
clientes_cadastrados = []
if os.path.exists(ARQUIVO_CLIENTES):
    try:
        df_cli = pd.read_excel(ARQUIVO_CLIENTES)
        for col in df_cli.columns:
            if any(x in str(col).lower() for x in ['cliente', 'nome', 'empresa']):
                clientes_cadastrados = sorted(df_cli[col].dropna().astype(str).str.strip().unique().tolist())
                break
        if not clientes_cadastrados and not df_cli.empty:
            clientes_cadastrados = sorted(df_cli.iloc[:, 0].dropna().astype(str).str.strip().unique().tolist())
    except Exception as e:
        st.error(f"Erro ao ler clientes: {e}")

# 2. Carregar Produtos da planilha específica
ARQUIVO_PRODUTOS = "produtosnext"
produtos_dict = {}
if os.path.exists(ARQUIVO_PRODUTOS):
    try:
        df_prod = pd.read_excel(ARQUIVO_PRODUTOS)
        cols = list(df_prod.columns)
        
        col_nome = next((c for c in cols if any(x in str(c).lower() for x in ['nome', 'produto', 'descrição'])), cols[0])
        col_preco = next((c for c in cols if any(x in str(c).lower() for x in ['preço', 'valor', 'venda'])), cols[1] if len(cols) > 1 else None)
        col_unidade = next((c for c in cols if any(x in str(c).lower() for x in ['unidade', 'und'])), cols[2] if len(cols) > 2 else None)

        for _, row in df_prod.iterrows():
            p_nome = str(row[col_nome]).strip()
            if p_nome and p_nome != 'nan':
                p_preco = 0.0
                p_unidade = "UN"
                
                if col_preco is not None and pd.notna(row[col_preco]):
                    try:
                        val_str = str(row[col_preco]).replace('R$', '').replace('.', '').replace(',', '.').strip()
                        p_preco = float(val_str)
                    except:
                        p_preco = 0.0
                        
                if col_unidade is not None and pd.notna(row[col_unidade]):
                    p_unidade = str(row[col_unidade]).strip()
                
                produtos_dict[p_nome] = {'preco': p_preco, 'unidade': p_unidade}
    except Exception as e:
        st.error(f"Erro ao ler produtos: {e}")

produtos_cadastrados = sorted(list(produtos_dict.keys()))

# --- SIDEBAR TECNOLÓGICA ---
st.sidebar.markdown("### ⚡ NEXT")
st.sidebar.markdown("---")
aba = st.sidebar.radio("Navegação do Sistema", ["🚀 Registrar Venda", "📦 Registrar Compra"])
st.sidebar.markdown("---")
st.sidebar.info("💡 **Dica Tech:** Os dados são sincronizados instantaneamente com os arquivos Excel locais.")

# --- MÓDULO DE VENDAS ---
if aba == "🚀 Registrar Venda":
    st.markdown("## 🚀 Terminal de Vendas")
    st.markdown("Insira ou selecione os parâmetros do pedido comercial.")
    
    col_main1, col_main2 = st.columns([2, 1], gap="large")
    
    with col_main1:
        st.markdown("### 👤 Identificação do Cliente")
        cliente_input = st.selectbox(
            "Selecione um cliente da base:",
            options=["-- Selecione da Base --"] + clientes_cadastrados,
            index=0,
            key="venda_cliente_select"
        )
        
        digitar_novo_cliente = st.text_input("Ou digite o nome de um Cliente Novo:").strip()
        
        if digitar_novo_cliente:
            nome_cliente = digitar_novo_cliente
            st.warning(f"⚠️ **'{nome_cliente}'** será registrado como **(cliente novo)**.")
        elif cliente_input != "-- Selecione da Base --":
            nome_cliente = cliente_input
            st.success(f"✅ Cliente ativo: **{nome_cliente}**")
        else:
            nome_cliente = ""

        st.markdown("---")
        st.markdown("### 🛒 Seleção de Produto")
        produto_input = st.selectbox(
            "Selecione um produto da lista:",
            options=["-- Selecione ou digite abaixo --"] + produtos_cadastrados,
            index=0,
            key="venda_prod_select"
        )
        
        digitar_novo_produto = st.text_input("Ou digite um Produto / Descrição novo(a):").strip()
        
        produto = ""
        preco_base = 0.0
        unidade_base = "UN"
        
        if digitar_novo_produto:
            produto = digitar_novo_produto
        elif produto_input != "-- Selecione ou digite abaixo --":
            produto = produto_input
            if produto in produtos_dict:
                preco_base = produtos_dict[produto]['preco']
                unidade_base = produtos_dict[produto]['unidade']

        col_q1, col_q2 = st.columns(2)
        with col_q1:
            quantidade = st.number_input("Quantidade", min_value=0.01, value=1.0, key="venda_qtd")
        with col_q2:
            opcoes_unidades = ["UN", "Barra", "Metro (m)", "Centímetro (cm)", "KG", "Litro (L)", "Caixa", "Pacote", "Peça"]
            if unidade_base not in opcoes_unidades and unidade_base != 'nan':
                opcoes_unidades.insert(0, unidade_base)
            
            idx_un = opcoes_unidades.index(unidade_base) if unidade_base in opcoes_unidades else 0
            unidade = st.selectbox("Unidade de Medida", options=opcoes_unidades, index=idx_un, key="venda_unidade_select")

        preco_unitario = st.number_input("Preço Unitário (R$)", min_value=0.0, value=float(preco_base), key="venda_preco")

    with col_main2:
        st.markdown("### 💳 Condições & Totais")
        data_venda = st.date_input("Data da Venda")
        
        forma_pgto = st.selectbox("Forma de Pagamento", ["PIX", "Dinheiro", "Cartão de Crédito", "Cartão de Débito", "A Prazo"], key="venda_pgto")
        
        if forma_pgto in ["Cartão de Crédito", "A Prazo"]:
            parcelas = st.selectbox("Parcelas", ["À vista", "2x", "3x", "4x", "5x", "6x", "7x", "8x", "9x", "10x", "11x", "12x"], key="venda_parcelas")
        else:
            parcelas = "À vista"

        total = quantidade * preco_unitario
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.metric(label="VALOR TOTAL DO ITEM", value=f"R$ {total:,.2f}")
        st.caption(f"Qtde: {quantidade} {unidade} | {forma_pgto} ({parcelas})")
        
        st.markdown("<br>", unsafe_allow_html=True)
        btn_adicionar = st.button("💾 Salvar Venda")
        
        if btn_adicionar:
            if not nome_cliente or not produto:
                st.error("Preencha o cliente e o produto antes de continuar.")
            else:
                ARQ_VENDAS = "vendas.xlsx"
                novo_registro = {
                    "Data": str(data_venda),
                    "Cliente": nome_cliente,
                    "Produto": produto,
                    "Quantidade": quantidade,
                    "Unidade": unidade,
                    "Preço Unitário": preco_unitario,
                    "Total": total,
                    "Forma de Pagamento": forma_pgto,
                    "Parcelas": parcelas
                }
                
                try:
                    if os.path.exists(ARQ_VENDAS):
                        df_vendas = pd.read_excel(ARQ_VENDAS)
                        if df_vendas.empty:
                            df_vendas = pd.DataFrame([novo_registro])
                        else:
                            df_vendas = pd.concat([df_vendas, pd.DataFrame([novo_registro])], ignore_index=True)
                    else:
                        df_vendas = pd.DataFrame([novo_registro])
                    
                    df_vendas.to_excel(ARQ_VENDAS, index=False)
                    st.success("✅ Venda salva com sucesso!")
                except Exception as e:
                    st.error(f"Erro ao salvar no Excel: {e}")

# --- MÓDULO DE COMPRAS ---
elif aba == "📦 Registrar Compra":
    st.markdown("## 📦 Terminal de Compras & Suprimentos")
    st.markdown("Registre entradas de produtos e custos com fornecedores.")
    
    col_c1, col_c2 = st.columns([2, 1], gap="large")
    
    with col_c1:
        fornecedor = st.text_input("Nome do Fornecedor / Estabelecimento").strip()
        produto_compra = st.text_input("Produto / Descrição da Compra").strip()
        
        col_cq1, col_cq2 = st.columns(2)
        with col_cq1:
            quantidade_compra = st.number_input("Quantidade", min_value=0.01, value=1.0, key="compra_qtd")
        with col_cq2:
            unidade_compra = st.selectbox("Unidade de Medida", ["UN", "Barra", "Metro (m)", "Centímetro (cm)", "KG", "Litro (L)", "Caixa", "Pacote", "Peça"], key="compra_unidade")

        preco_unitario_compra = st.number_input("Preço Unitário (R$)", min_value=0.0, value=0.0, key="compra_preco_compra")

    with col_c2:
        data_compra = st.date_input("Data da Compra")
        forma_pgto_compra = st.selectbox("Forma de Pagamento", ["PIX", "Dinheiro", "Cartão de Crédito", "Boleto", "A Prazo"], key="compra_pgto")
        
        if forma_pgto_compra in ["Cartão de Crédito", "Boleto", "A Prazo"]:
            parcelas_compra = st.selectbox("Parcelas", ["À vista", "2x", "3x", "4x", "5x", "6x", "7x", "8x", "9x", "10x", "11x", "12x"], key="compra_parcelas")
        else:
            parcelas_compra = "À vista"

        total_compra = quantidade_compra * preco_unitario_compra
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.metric(label="VALOR TOTAL DA COMPRA", value=f"R$ {total_compra:,.2f}")
        st.caption(f"Qtde: {quantidade_compra} {unidade_compra} | {forma_pgto_compra} ({parcelas_compra})")
        
        st.markdown("<br>", unsafe_allow_html=True)
        btn_adicionar_compra = st.button("💾 Salvar Compra")
        
        if btn_adicionar_compra:
            if not fornecedor or not produto_compra:
                st.error("Preencha o fornecedor e o produto.")
            else:
                ARQ_COMPRAS = "compras.xlsx"
                novo_registro_compra = {
                    "Data": str(data_compra),
                    "Fornecedor": fornecedor,
                    "Produto": produto_compra,
                    "Quantidade": quantidade_compra,
                    "Unidade": unidade_compra,
                    "Preço Unitário": preco_unitario_compra,
                    "Total": total_compra,
                    "Forma de Pagamento": forma_pgto_compra,
                    "Parcelas": parcelas_compra
                }
                
                try:
                    if os.path.exists(ARQ_COMPRAS):
                        df_compras = pd.read_excel(ARQ_COMPRAS)
                        if df_compras.empty:
                            df_compras = pd.DataFrame([novo_registro_compra])
                        else:
                            df_compras = pd.concat([df_compras, pd.DataFrame([novo_registro_compra])], ignore_index=True)
                    else:
                        df_compras = pd.DataFrame([novo_registro_compra])
                    
                    df_compras.to_excel(ARQ_COMPRAS, index=False)
                    st.success("✅ Compra salva com sucesso!")
                except Exception as e:
                    st.error(f"Erro ao salvar no Excel: {e}")