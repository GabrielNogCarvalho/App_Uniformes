import streamlit as st
import pandas as pd
from datetime import date
import gspread
import os

# Configuração da página
st.set_page_config(
    page_title="Controle de Estoque - Uniformes",
    page_icon="👷",
    layout="wide"
)

st.title("📦 Controle de Estoque de Uniformes")

# ID da sua planilha no Google Sheets
SPREADSHEET_ID = "1JhRbD0lKigbScd7BaNyMZXJIcG2_QhvQSs1YLCT4VwM"

# Localiza automaticamente o arquivo .json de credenciais na pasta
@st.cache_resource
def conectar_google():
    arquivos_json = [f for f in os.listdir('.') if f.endswith('.json')]
    if not arquivos_json:
        raise FileNotFoundError("Nenhum arquivo .json de credenciais foi encontrado na pasta do projeto.")
    
    # Usa o arquivo json encontrado
    caminho_credencial = arquivos_json[0]
    gc = gspread.service_account(filename=caminho_credencial)
    sh = gc.open_by_key(SPREADSHEET_ID)
    return sh

try:
    planilha = conectar_google()
    aba_estoque = planilha.worksheet("Estoque Atual")
    dados_estoque = aba_estoque.get_all_records()
    df_estoque = pd.DataFrame(dados_estoque)
except Exception as e:
    st.error(f"Erro ao conectar com a planilha: {e}")
    st.info("Verifique se o arquivo .json baixado está salvo dentro da pasta do projeto e se a planilha foi compartilhada com o e-mail da service account.")
    st.stop()

# Layout em duas colunas: Esquerda (Estoque) e Direita (Lançamento)
col_estoque, col_form = st.columns([3, 2], gap="large")

# --- COLUNA 1: VISUALIZAÇÃO DO ESTOQUE ---
with col_estoque:
    st.subheader("📊 Saldo em Estoque")
    
    busca = st.text_input("🔍 Buscar item ou tamanho:", placeholder="Ex: Calça, NR-10, G...")
    
    if busca and not df_estoque.empty:
        df_filtrado = df_estoque[
            df_estoque["ITEM"].astype(str).str.contains(busca, case=False, na=False) |
            df_estoque["TAMANHO"].astype(str).str.contains(busca, case=False, na=False)
        ]
    else:
        df_filtrado = df_estoque

    st.dataframe(
        df_filtrado,
        use_container_width=True,
        hide_index=True
    )

# --- COLUNA 2: FORMULÁRIO DE MOVIMENTAÇÃO ---
with col_form:
    st.subheader("📝 Registrar Movimentação")
    
    lista_itens = sorted(df_estoque["ITEM"].dropna().unique().tolist()) if not df_estoque.empty else []
    
    with st.form("form_movimentacao", clear_on_submit=True):
        data_mov = st.date_input("Data da Movimentação", value=date.today())
        
        tipo_mov = st.radio(
            "Tipo de Movimento:",
            ["SAÍDA", "ENTRADA"],
            horizontal=True
        )
        
        item_selecionado = st.selectbox("Item:", lista_itens)
        tamanho_selecionado = st.selectbox("Tamanho:", ["P", "M", "G", "GG", "XG", "XXG"])
        quantidade = st.number_input("Quantidade:", min_value=1, step=1, value=1)
        
        responsavel = st.text_input("Responsável (Quem retirou / entregou):", placeholder="Ex: João Silva")
        obra = st.text_input("Obra / Destino:", placeholder="Ex: Obra Central, Almoxarifado...")
        
        salvar = st.form_submit_button("💾 Gravar Movimentação", use_container_width=True)
        
        if salvar:
            if not responsavel.strip():
                st.warning("Por favor, preencha o nome do Responsável.")
            elif not obra.strip():
                st.warning("Por favor, preencha a Obra.")
            else:
                with st.spinner("Gravando na planilha..."):
                    try:
                        aba_mov = planilha.worksheet("Movimentações")
                        
                        # Insere uma nova linha na aba de Movimentações
                        nova_linha = [
                            data_mov.strftime("%d/%m/%Y"),
                            tipo_mov,
                            item_selecionado,
                            tamanho_selecionado,
                            int(quantidade),
                            responsavel.strip(),
                            obra.strip()
                        ]
                        
                        aba_mov.append_row(nova_linha, value_input_option="USER_ENTERED")
                        
                        st.success(f"{tipo_mov} de {quantidade}x {item_selecionado} ({tamanho_selecionado}) gravada com sucesso!")
                        st.rerun()
                    except Exception as erro:
                        st.error(f"Erro ao salvar na planilha: {erro}")
