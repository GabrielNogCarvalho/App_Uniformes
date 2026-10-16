import streamlit as st
import pandas as pd
from datetime import date
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Configuracao da pagina (sem emojis)
st.set_page_config(
    page_title="Sistema de Gestao Integrada",
    layout="wide"
)

# Caminho do arquivo Excel no OneDrive
CAMINHO_EXCEL = "Controle_Estoque_Uniformes.xlsx"

# Configuracoes de Envio de Email (Outlook / Microsoft 365)
EMAIL_CONFIG = {
    "ativar_alertas": True,
    "servidor_smtp": "smtp.office365.com",
    "porta_smtp": 587,
    "remetente": "seu_email@defendi.com.br",
    "senha_app": "sua_senha_de_aplicativo",
    "destinatario": "seu_email@defendi.com.br"
}

# Estilo visual corporativo (Azul escuro, Laranja e Branco)
st.markdown("""
<style>
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
    [data-testid="stForm"] {
        background-color: #172A45 !important;
        border: 1px solid #243E63 !important;
        border-radius: 8px !important;
        padding: 24px !important;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.2) !important;
    }
    [data-testid="stForm"] label, 
    [data-testid="stForm"] p, 
    [data-testid="stWidgetLabel"] p,
    div[data-testid="stRadio"] label p {
        color: #FFFFFF !important;
        font-weight: 600 !important;
        font-size: 14px !important;
    }
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
    h3 {
        color: #FFFFFF !important;
        font-weight: 600 !important;
    }
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

# Funcao de envio de e-mail via Outlook
def disparar_email_reposicao(item, tamanho, saldo_final, minimo_configurado):
    if not EMAIL_CONFIG["ativar_alertas"]:
        return False, "Alertas desativados."
    
    try:
        msg = MIMEMultipart()
        msg["From"] = EMAIL_CONFIG["remetente"]
        msg["To"] = EMAIL_CONFIG["destinatario"]
        msg["Subject"] = f"[ALERTA DE ESTOQUE] Nivel Critico: {item} ({tamanho})"
        
        corpo = f"""
Prezados,

Este e um aviso automatico de necessidade de reposicao de estoque.

Item: {item}
Tamanho: {tamanho}
Saldo Apos Movimentacao: {saldo_final} unidade(s)
Estoque Minimo Definido: {minimo_configurado} unidade(s)

Recomenda-se providenciar pedido de compra ou transferencia para este item.

Sistema de Gestao de Uniformes - Defendi Engenharia
        """
        msg.attach(MIMEText(corpo, "plain"))
        
        with smtplib.SMTP(EMAIL_CONFIG["servidor_smtp"], EMAIL_CONFIG["porta_smtp"]) as server:
            server.starttls()
            server.login(EMAIL_CONFIG["remetente"], EMAIL_CONFIG["senha_app"])
            server.sendmail(EMAIL_CONFIG["remetente"], [EMAIL_CONFIG["destinatario"]], msg.as_string())
        return True, "Email enviado com sucesso."
    except Exception as erro:
        return False, str(erro)

# Identifica os nomes reais das abas no Excel (com ou sem acento)
def obter_nomes_abas(xl_file):
    folhas = xl_file.sheet_names
    aba_est = "Estoque Atual" if "Estoque Atual" in folhas else folhas[0]
    aba_mov = "Movimentações" if "Movimentações" in folhas else ("Movimentacoes" if "Movimentacoes" in folhas else folhas[1])
    aba_km = "Controle de KM" if "Controle de KM" in folhas else folhas[2]
    return aba_est, aba_mov, aba_km

# Leitura segura dos dados diretamente do disco
def carregar_dados():
    if not os.path.exists(CAMINHO_EXCEL):
        st.error(f"Arquivo nao encontrado: {CAMINHO_EXCEL}")
        st.stop()
        
    xl = pd.ExcelFile(CAMINHO_EXCEL)
    aba_est, aba_mov, aba_km = obter_nomes_abas(xl)
    
    df_est = pd.read_excel(xl, sheet_name=aba_est)
    df_m = pd.read_excel(xl, sheet_name=aba_mov)
    df_k = pd.read_excel(xl, sheet_name=aba_km)
    return df_est, df_m, df_k

# Gravacao segura de nova movimentacao de uniforme sem perder dados anteriores
def gravar_movimentacao_segura(nova_linha):
    xl = pd.ExcelFile(CAMINHO_EXCEL)
    aba_est, aba_mov, aba_km = obter_nomes_abas(xl)
    
    # Carrega historico real do disco
    df_est = pd.read_excel(xl, sheet_name=aba_est)
    df_m_existente = pd.read_excel(xl, sheet_name=aba_mov)
    df_k_existente = pd.read_excel(xl, sheet_name=aba_km)
    
    # Anexa a nova movimentacao no final do historico
    df_m_atualizado = pd.concat([df_m_existente, pd.DataFrame([nova_linha])], ignore_index=True)
    
    # Garante a presenca da coluna ESTOQUE MINIMO
    if "ESTOQUE MÍNIMO" not in df_est.columns and "ESTOQUE MINIMO" not in df_est.columns:
        df_est["ESTOQUE MÍNIMO"] = 2
        
    col_mov = "MOVIMENTO (ENTRADA/SAÍDA)" if "MOVIMENTO (ENTRADA/SAÍDA)" in df_m_atualizado.columns else df_m_atualizado.columns[1]

    # Recalcula entradas e saidas mantendo o ESTOQUE INICIAL digitado no inventario
    for idx, row in df_est.iterrows():
        item = row["ITEM"]
        tam = row["TAMANHO"]
        
        filtro = (df_m_atualizado["ITEM"] == item) & (df_m_atualizado["TAMANHO"] == tam)
        ent = df_m_atualizado[filtro & (df_m_atualizado[col_mov].astype(str).str.upper().str.contains("ENTRADA"))]["QUANTIDADE"].sum()
        sai = df_m_atualizado[filtro & (df_m_atualizado[col_mov].astype(str).str.upper().str.contains("SAÍDA|SAIDA"))]["QUANTIDADE"].sum()
        
        ini = row.get("ESTOQUE INICIAL", 0)
        try:
            ini = float(ini) if pd.notnull(ini) else 0.0
        except:
            ini = 0.0
            
        df_est.at[idx, "ENTRADAS"] = int(ent)
        df_est.at[idx, "SAÍDAS"] = int(sai)
        df_est.at[idx, "ESTOQUE ATUAL"] = int(ini + ent - sai)

    # Grava todas as abas preservadas de volta no arquivo
    with pd.ExcelWriter(CAMINHO_EXCEL, engine="openpyxl") as writer:
        df_est.to_excel(writer, sheet_name=aba_est, index=False)
        df_m_atualizado.to_excel(writer, sheet_name=aba_mov, index=False)
        df_k_existente.to_excel(writer, sheet_name=aba_km, index=False)

# Gravacao segura de nova viagem sem mexer nos uniformes
def gravar_viagem_segura(nova_linha_km):
    xl = pd.ExcelFile(CAMINHO_EXCEL)
    aba_est, aba_mov, aba_km = obter_nomes_abas(xl)
    
    df_est_existente = pd.read_excel(xl, sheet_name=aba_est)
    df_m_existente = pd.read_excel(xl, sheet_name=aba_mov)
    df_k_existente = pd.read_excel(xl, sheet_name=aba_km)
    
    df_k_atualizado = pd.concat([df_k_existente, pd.DataFrame([nova_linha_km])], ignore_index=True)
    
    with pd.ExcelWriter(CAMINHO_EXCEL, engine="openpyxl") as writer:
        df_est_existente.to_excel(writer, sheet_name=aba_est, index=False)
        df_m_existente.to_excel(writer, sheet_name=aba_mov, index=False)
        df_k_atualizado.to_excel(writer, sheet_name=aba_km, index=False)

try:
    df_estoque, df_mov, df_km = carregar_dados()
except Exception as e:
    st.error(f"Falha ao abrir a planilha: {e}")
    st.stop()

# --- MENU LATERAL DE NAVEGACAO ---
st.sidebar.markdown("## NAVEGACAO")
modulo_selecionado = st.sidebar.radio(
    "Selecione o Modulo:",
    ["Uniformes e EPIs", "Controle de KM"],
    label_visibility="collapsed"
)

# ==============================================================================
# MODULO 1: UNIFORMES E EPIS
# ==============================================================================
if modulo_selecionado == "Uniformes e EPIs":
    st.markdown("""
    <div class="header-bar">
        <div class="header-title">CONTROLE DE ESTOQUE DE UNIFORMES E EPIS</div>
        <div class="header-subtitle">Gestao Integrada com Arquivo Oficial no OneDrive</div>
    </div>
    """, unsafe_allow_html=True)

    # Tratamento seguro para colunas numericas e estoque minimo
    if "ESTOQUE MÍNIMO" in df_estoque.columns:
        col_minimo = pd.to_numeric(df_estoque["ESTOQUE MÍNIMO"], errors="coerce").fillna(2)
    elif "ESTOQUE MINIMO" in df_estoque.columns:
        col_minimo = pd.to_numeric(df_estoque["ESTOQUE MINIMO"], errors="coerce").fillna(2)
    else:
        col_minimo = pd.Series([2] * len(df_estoque))
        df_estoque["ESTOQUE MÍNIMO"] = 2

    col_estoque_atual = pd.to_numeric(df_estoque["ESTOQUE ATUAL"], errors="coerce").fillna(0) if "ESTOQUE ATUAL" in df_estoque.columns else pd.Series([0] * len(df_estoque))
    col_entradas = pd.to_numeric(df_estoque["ENTRADAS"], errors="coerce").fillna(0) if "ENTRADAS" in df_estoque.columns else pd.Series([0] * len(df_estoque))
    col_saidas = pd.to_numeric(df_estoque["SAÍDAS"], errors="coerce").fillna(0) if "SAÍDAS" in df_estoque.columns else pd.Series([0] * len(df_estoque))
    
    total_pecas = int(col_estoque_atual.sum())
    total_entradas_acum = int(col_entradas.sum())
    total_saidas_acum = int(col_saidas.sum())
    itens_criticos = int((col_estoque_atual <= col_minimo).sum())

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Saldo Total</div>
            <div class="kpi-value">{total_pecas} <span style="font-size:14px; font-weight:400; color:#94A3B8;">pecas</span></div>
        </div>
        """, unsafe_allow_html=True)
    with k2:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Total de Entradas</div>
            <div class="kpi-value">{total_entradas_acum} <span style="font-size:14px; font-weight:400; color:#94A3B8;">pecas</span></div>
        </div>
        """, unsafe_allow_html=True)
    with k3:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Total de Saidas</div>
            <div class="kpi-value">{total_saidas_acum} <span style="font-size:14px; font-weight:400; color:#94A3B8;">pecas</span></div>
        </div>
        """, unsafe_allow_html=True)
    with k4:
        st.markdown(f"""
        <div class="kpi-card-alert">
            <div class="kpi-label">Itens no Limite / Zerados</div>
            <div class="kpi-value-alert">{itens_criticos} <span style="font-size:14px; font-weight:400; color:#94A3B8;">itens</span></div>
        </div>
        """, unsafe_allow_html=True)

    st.write("")

    col_tabela, col_form = st.columns([3, 2], gap="large")

    with col_tabela:
        st.subheader("Posicao Atual de Estoque")
        busca = st.text_input("Filtrar item:", placeholder="Digite para filtrar...", label_visibility="collapsed")

        df_exibir = df_estoque
        if busca and not df_estoque.empty:
            df_exibir = df_estoque[
                df_estoque["ITEM"].astype(str).str.contains(busca, case=False, na=False) |
                df_estoque["TAMANHO"].astype(str).str.contains(busca, case=False, na=False)
            ]

        st.dataframe(df_exibir, use_container_width=True, hide_index=True)

    with col_form:
        st.subheader("Registrar Movimentacao")
        lista_itens = sorted(df_estoque["ITEM"].dropna().unique().tolist()) if not df_estoque.empty else []
        
        with st.form("form_uniforme", clear_on_submit=True):
            data_mov = st.date_input("Data da Operacao", value=date.today())
            tipo_mov = st.radio("Tipo de Movimento", ["SAÍDA", "ENTRADA"], horizontal=True)
            item_sel = st.selectbox("Item", lista_itens)
            tam_sel = st.selectbox("Tamanho", ["P", "M", "G", "GG", "XG", "XXG", "EG", "EGG", "EXG", "EXGG", "G1"])
            qtd = st.number_input("Quantidade", min_value=1, step=1, value=1)
            resp = st.text_input("Responsavel", placeholder="Nome de quem retirou ou entregou")
            obra = st.text_input("Obra ou Destino", placeholder="Identificacao da obra")
            
            btn_salvar = st.form_submit_button("Confirmar Movimentacao", use_container_width=True)
            
            if btn_salvar:
                if not resp.strip():
                    st.warning("Preencha o campo Responsavel.")
                elif not obra.strip():
                    st.warning("Preencha o campo Obra.")
                else:
                    # Identifica saldo atual e estoque minimo do item antes de salvar
                    filtro_item_atual = (df_estoque["ITEM"] == item_sel) & (df_estoque["TAMANHO"] == tam_sel)
                    saldo_atual_item = int(df_estoque.loc[filtro_item_atual, "ESTOQUE ATUAL"].values[0]) if (not df_estoque[filtro_item_atual].empty and "ESTOQUE ATUAL" in df_estoque.columns) else 0
                    
                    col_ref_min = "ESTOQUE MÍNIMO" if "ESTOQUE MÍNIMO" in df_estoque.columns else ("ESTOQUE MINIMO" if "ESTOQUE MINIMO" in df_estoque.columns else None)
                    minimo_item = int(df_estoque.loc[filtro_item_atual, col_ref_min].values[0]) if (col_ref_min and not df_estoque[filtro_item_atual].empty) else 2

                    nova_mov = {
                        "DATA": data_mov.strftime("%d/%m/%Y"),
                        "MOVIMENTO (ENTRADA/SAÍDA)": tipo_mov,
                        "ITEM": item_sel,
                        "TAMANHO": tam_sel,
                        "QUANTIDADE": int(qtd),
                        "RESPONSÁVEL": resp.strip(),
                        "OBRA": obra.strip()
                    }

                    # Grava diretamente no arquivo do disco de forma acumulativa
                    gravar_movimentacao_segura(nova_mov)
                    st.success(f"{tipo_mov} registrada com sucesso no historico!")

                    # Verificacao de alerta de estoque para saidas
                    if tipo_mov == "SAÍDA":
                        saldo_restante = saldo_atual_item - int(qtd)
                        if saldo_restante <= minimo_item:
                            st.warning(f"Atencao: Saldo restante ({saldo_restante} un.) atingiu o estoque minimo ({minimo_item} un.).")
                            enviou, msg_envio = disparar_email_reposicao(item_sel, tam_sel, saldo_restante, minimo_item)
                            if enviou:
                                st.info("Alerta de reposicao enviado por e-mail com sucesso.")
                            else:
                                st.caption(f"Status do e-mail: {msg_envio}")

                    st.rerun()

# ==============================================================================
# MODULO 2: CONTROLE DE KM
# ==============================================================================
elif modulo_selecionado == "Controle de KM":
    st.markdown("""
    <div class="header-bar">
        <div class="header-title">CONTROLE DE QUILOMETRAGEM E FROTAS</div>
        <div class="header-subtitle">Registro de Viagens Salvo no Arquivo do OneDrive</div>
    </div>
    """, unsafe_allow_html=True)

    km_percorrido_col = pd.to_numeric(df_km.get("KM PERCORRIDO", 0), errors="coerce").fillna(0)
    abastecimento_col = pd.to_numeric(df_km.get("VALOR ABASTECIMENTO (R$)", 0), errors="coerce").fillna(0)
    
    total_km = int(km_percorrido_col.sum())
    total_viagens = len(df_km)
    total_abast = float(abastecimento_col.sum())
    media_km = int(total_km / total_viagens) if total_viagens > 0 else 0

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Total Percorrido</div>
            <div class="kpi-value">{total_km} <span style="font-size:14px; font-weight:400; color:#94A3B8;">km</span></div>
        </div>
        """, unsafe_allow_html=True)
    with k2:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Viagens Registradas</div>
            <div class="kpi-value">{total_viagens} <span style="font-size:14px; font-weight:400; color:#94A3B8;">registros</span></div>
        </div>
        """, unsafe_allow_html=True)
    with k3:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Custo Abastecimento</div>
            <div class="kpi-value">R$ {total_abast:.2f}</div>
        </div>
        """, unsafe_allow_html=True)
    with k4:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Media por Viagem</div>
            <div class="kpi-value">{media_km} <span style="font-size:14px; font-weight:400; color:#94A3B8;">km/viagem</span></div>
        </div>
        """, unsafe_allow_html=True)

    st.write("")

    col_tabela_km, col_form_km = st.columns([3, 2], gap="large")

    with col_tabela_km:
        st.subheader("Historico de Deslocamentos")
        busca_km = st.text_input("Filtrar viagens:", placeholder="Digite para filtrar...", label_visibility="collapsed")

        df_exibir_km = df_km
        if busca_km and not df_km.empty:
            df_exibir_km = df_km[
                df_km["VEÍCULO"].astype(str).str.contains(busca_km, case=False, na=False) |
                df_km["MOTORISTA"].astype(str).str.contains(busca_km, case=False, na=False) |
                df_km["OBRA"].astype(str).str.contains(busca_km, case=False, na=False)
            ]

        st.dataframe(df_exibir_km, use_container_width=True, hide_index=True)

    with col_form_km:
        st.subheader("Registrar Nova Viagem")
        
        with st.form("form_km", clear_on_submit=True):
            data_viagem = st.date_input("Data do Deslocamento", value=date.today())
            veiculo = st.text_input("Veiculo e Placa", placeholder="Ex: Strada ABC-1234")
            motorista = st.text_input("Motorista / Responsavel", placeholder="Ex: Carlos Silva")
            destino_obra = st.text_input("Obra ou Destino", placeholder="Ex: Obra Central")
            
            c_kmi, c_kmf = st.columns(2)
            with c_kmi:
                km_inicial = st.number_input("KM Inicial", min_value=0, step=1, value=0)
            with c_kmf:
                km_final = st.number_input("KM Final", min_value=0, step=1, value=0)
            
            motivo_viagem = st.text_input("Motivo do Deslocamento", placeholder="Ex: Visita tecnica")
            valor_abastecimento = st.number_input("Valor de Abastecimento (R$)", min_value=0.0, step=10.0, value=0.0)
            
            btn_salvar_km = st.form_submit_button("Confirmar Viagem", use_container_width=True)
            
            if btn_salvar_km:
                if not veiculo.strip():
                    st.warning("Preencha o campo Veiculo.")
                elif not motorista.strip():
                    st.warning("Preencha o campo Motorista.")
                elif not destino_obra.strip():
                    st.warning("Preencha o campo Obra.")
                elif km_final < km_inicial:
                    st.warning("O KM Final nao pode ser menor que o KM Inicial.")
                else:
                    km_calculado = int(km_final - km_inicial)
                    nova_viagem = {
                        "DATA": data_viagem.strftime("%d/%m/%Y"),
                        "VEÍCULO": veiculo.strip(),
                        "MOTORISTA": motorista.strip(),
                        "OBRA": destino_obra.strip(),
                        "KM INICIAL": int(km_inicial),
                        "KM FINAL": int(km_final),
                        "KM PERCORRIDO": km_calculado,
                        "MOTIVO": motivo_viagem.strip(),
                        "VALOR ABASTECIMENTO (R$)": float(valor_abastecimento)
                    }

                    # Grava diretamente no arquivo do disco de forma acumulativa
                    gravar_viagem_segura(nova_viagem)
                    st.success(f"Viagem gravada com sucesso! Percorrido: {km_calculado} km.")
                    st.rerun()
