import streamlit as st
import pandas as pd
from datetime import date
import gspread
import os
import io

# Configuracao da pagina (sem emojis)
st.set_page_config(
    page_title="Controle de Estoque - Uniformes",
    layout="wide"
)

# Estilo Personalizado: Tons escuros coordenados com Laranja e Branco
st.markdown("""
<style>
    /* Cabecalho Institucional */
    .header-bar {
        background-color: #0F1E32;
        padding: 22px 28px;
        border-radius: 8px;
        border: 1px solid #243E63;
        border-left: 6px solid #E65100;
        margin-bottom: 24px;
    }
    .header-title {
        color: #FFFFFF;
        font-size: 24px;
        font-weight: 700;
        margin: 0;
        letter-spacing: 0.5px;
    }
    .header-subtitle {
        color: #94A3B8;
        font-size: 13px;
        margin-top: 4px;
        font-weight: 400;
    }

    /* Cartoes de Saldo/Indicadores (Tom escuro destacado) */
    .kpi-card {
        background-color: #172A45;
        border: 1px solid #243E63;
        border-radius: 8px;
        padding: 18px 20px;
        border-top: 4px solid #3B82F6;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.2);
    }
    .kpi-card-alert {
        background-color: #172A45;
        border: 1px solid #243E63;
        border-radius: 8px;
        padding: 18px 20px;
        border-top: 4px solid #E65100;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.2);
    }
    .kpi-label {
        font-size: 12px;
        color: #CBD5E1;
        text-transform: uppercase;
        font-weight: 600;
        letter-spacing: 0.5px;
        margin-bottom: 6px;
    }
    .kpi-value {
        font-size: 28px;
        color: #FFFFFF;
        font-weight: 700;
    }
    .kpi-value-alert {
        font-size: 28px;
        color: #FF7A00;
        font-weight: 700;
    }

    /* Container do Formulario */
    [data-testid="stForm"] {
        background-color: #172A45 !important;
        border: 1px solid #243E63 !important;
        border-radius: 8px !important;
        padding: 24px !important;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.2) !important;
    }

    /* Garante que todos os rotulos e textos fiquem visiveis em branco */
    [data-testid="stForm"] label, 
    [data-testid="stForm"] p, 
    [data-testid="stWidgetLabel"] p,
    div[data-testid="stRadio"] label p {
        color: #FFFFFF !important;
        font-weight: 600 !important;
        font-size: 14px !important;
    }

    /* Campos de Entrada e Selecao */
    div[data-baseweb="input"] {
        background-color: #0F1E32 !important;
        border: 1px solid #243E63 !important;
        border-radius: 6px !important;
    }
    div[data-baseweb="input"] input {
        color: #FFFFFF !important;
    }
    div[data-baseweb="select"] > div {
        background-color: #0F1E32 !important;
        color: #FFFFFF !important;
        border: 1px solid #243E63 !important;
        border-radius: 6px !important;
    }

    /* Titulos das secoes */
    h3 {
        color: #FFFFFF !important;
        font-weight: 600 !important;
    }

    /* Botoes com destaque Laranja e Azul */
    div.stButton > button {
        background-color: #E65100 !important;
        color: #FFFFFF !important;
        border: 1px solid #E65100 !important;
        border-radius: 6px !important;
        font-weight: 700 !important;
        padding: 10px 20px !important;
        transition: all 0.2s ease-in-out !important;
    }
    div.stButton > button:hover {
        background-color: #FF7A00 !important;
        border-color: #FF7A00 !important;
        color: #FFFFFF !important;
    }
    div.stDownloadButton > button {
        background-color: #0F1E32 !important;
        color: #FFFFFF !important;
        border: 1px solid #243E63 !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        padding: 8px 16px !important;
        transition: all 0.2s ease-in-out !important;
    }
    div.stDownloadButton > button:hover {
        background-color: #E65100 !important;
        border-color: #E65100 !important;
        color: #FFFFFF !important;
    }
</style>
""", unsafe_allow_html=True)

# Cabecalho da Aplicacao
st.markdown("""
<div class="header-bar">
    <div class="header-title">SISTEMA DE CONTROLE DE ESTOQUE</div>
    <div class="header-subtitle">Gestão Integrada de Uniformes e Equipamentos de Proteção</div>
</div>
""", unsafe_allow_html=True)

# ID da Planilha no Google Sheets
SPREADSHEET_ID = "1JhRbD0lKigbScd7BaNyMZXJIcG2_QhvQSs1YLCT4VwM"

# Conexao com Google Sheets
@st.cache_resource
def conectar_google():
    if "gcp_service_account" in st.secrets:
        gc = gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
        return gc.open_by_key(SPREADSHEET_ID)
    
    locais = [".", ".streamlit"]
    arquivos_json = []
    for local in locais:
        if os.path.exists(local):
            arquivos_json.extend([os.path.join(local, f) for f in os.listdir(local) if f.endswith(".json")])
    
    if arquivos_json:
        gc = gspread.service_account(filename=arquivos_json[0])
        return gc.open_by_key(SPREADSHEET_ID)
        
    raise FileNotFoundError("Credenciais nao encontradas.")

# Exportacao para Excel em memoria
def gerar_relatorio_excel(df_est, df_mov):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_est.to_excel(writer, sheet_name="Estoque Atual", index=False)
        df_mov.to_excel(writer, sheet_name="Movimentações", index=False)
    return output.getvalue()

try:
    planilha = conectar_google()
    
    aba_estoque = planilha.worksheet("Estoque Atual")
    df_estoque = pd.DataFrame(aba_estoque.get_all_records())
    
    aba_mov = planilha.worksheet("Movimentações")
    df_mov_historico = pd.DataFrame(aba_mov.get_all_records())
except Exception as e:
    st.error(f"Falha de conexão com a base de dados: {e}")
    st.stop()

# --- CARTOES DE INDICADORES (KPIs) ---
if not df_estoque.empty:
    col_estoque_atual = pd.to_numeric(df_estoque.get("ESTOQUE ATUAL", 0), errors="coerce").fillna(0)
    col_entradas = pd.to_numeric(df_estoque.get("ENTRADAS", 0), errors="coerce").fillna(0)
    col_saidas = pd.to_numeric(df_estoque.get("SAÍDAS", 0), errors="coerce").fillna(0)
    
    total_pecas = int(col_estoque_atual.sum())
    total_entradas_acum = int(col_entradas.sum())
    total_saidas_acum = int(col_saidas.sum())
    itens_zerados = int((col_estoque_atual <= 0).sum())

    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    with kpi1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Saldo em Estoque</div>
            <div class="kpi-value">{total_pecas} <span style="font-size:14px; font-weight:400; color:#94A3B8;">peças</span></div>
        </div>
        """, unsafe_allow_html=True)
    with kpi2:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Total de Entradas</div>
            <div class="kpi-value">{total_entradas_acum} <span style="font-size:14px; font-weight:400; color:#94A3B8;">peças</span></div>
        </div>
        """, unsafe_allow_html=True)
    with kpi3:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Total de Saídas</div>
            <div class="kpi-value">{total_saidas_acum} <span style="font-size:14px; font-weight:400; color:#94A3B8;">peças</span></div>
        </div>
        """, unsafe_allow_html=True)
    with kpi4:
        st.markdown(f"""
        <div class="kpi-card-alert">
            <div class="kpi-label">Itens com Saldo Zero</div>
            <div class="kpi-value-alert">{itens_zerados} <span style="font-size:14px; font-weight:400; color:#94A3B8;">itens</span></div>
        </div>
        """, unsafe_allow_html=True)

st.write("") # Espacamento

# --- ESTRUTURA PRINCIPAL ---
col_tabela, col_formulario = st.columns([3, 2], gap="large")

# COLUNA 1: CONSULTA DE ESTOQUE E EXPORTACAO
with col_tabela:
    st.subheader("Posição Atual de Estoque")
    
    col_busca, col_botao_export = st.columns([2, 1])
    with col_busca:
        busca = st.text_input("Filtrar por item ou tamanho:", placeholder="Digite para filtrar...", label_visibility="collapsed")
    with col_botao_export:
        arquivo_excel = gerar_relatorio_excel(df_estoque, df_mov_historico)
        st.download_button(
            label="Exportar para Excel",
            data=arquivo_excel,
            file_name=f"relatorio_estoque_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    if busca and not df_estoque.empty:
        df_exibicao = df_estoque[
            df_estoque["ITEM"].astype(str).str.contains(busca, case=False, na=False) |
            df_estoque["TAMANHO"].astype(str).str.contains(busca, case=False, na=False)
        ]
    else:
        df_exibicao = df_estoque

    st.dataframe(
        df_exibicao,
        use_container_width=True,
        hide_index=True
    )

# COLUNA 2: FORMULARIO DE REGISTRO
with col_formulario:
    st.subheader("Registrar Movimentação")
    
    lista_itens = sorted(df_estoque["ITEM"].dropna().unique().tolist()) if not df_estoque.empty else []
    
    with st.form("form_registro", clear_on_submit=True):
        data_mov = st.date_input("Data da Operação", value=date.today())
        
        tipo_mov = st.radio(
            "Tipo de Movimento",
            ["SAÍDA", "ENTRADA"],
            horizontal=True
        )
        
        item_selecionado = st.selectbox("Item", lista_itens)
        tamanho_selecionado = st.selectbox("Tamanho", ["P", "M", "G", "GG", "XG", "XXG", "EG", "EXG", "G1"])
        quantidade = st.number_input("Quantidade", min_value=1, step=1, value=1)
        
        responsavel = st.text_input("Responsável", placeholder="Nome de quem retirou ou entregou")
        obra = st.text_input("Obra ou Destino", placeholder="Identificação da obra")
        
        btn_gravar = st.form_submit_button("Confirmar Movimentação", use_container_width=True)
        
        if btn_gravar:
            if not responsavel.strip():
                st.warning("Preencha o campo Responsável.")
            elif not obra.strip():
                st.warning("Preencha o campo Obra.")
            else:
                with st.spinner("Registrando dados..."):
                    try:
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
                        st.success(f"{tipo_mov} registrada: {quantidade} unidade(s) de {item_selecionado} ({tamanho_selecionado}).")
                        st.rerun()
                    except Exception as erro:
                        st.error(f"Nao foi possivel salvar: {erro}")
