import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from io import BytesIO

# Configuração da página
st.set_page_config(page_title="Conversor de XML em dados", layout="wide", page_icon="⚡")

# --- ESTILIZAÇÃO CSS FUTURISTA & GLASSMORPHISM ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700;800&family=Rajdhani:wght@600;700&display=swap');

    /* Fundo Geral */
    .stApp {
        background: radial-gradient(circle at 15% 15%, rgba(0, 240, 255, 0.05) 0%, transparent 40%),
                    radial-gradient(circle at 85% 85%, rgba(138, 43, 226, 0.05) 0%, transparent 40%),
                    #090D14;
        color: #E2E8F0;
        font-family: 'Inter', sans-serif;
    }

    /* Cabeçalho Principal Futurista */
    .cyber-header {
        background: linear-gradient(135deg, rgba(15, 23, 42, 0.8), rgba(30, 41, 59, 0.6));
        border: 1px solid rgba(0, 240, 255, 0.25);
        border-radius: 18px;
        padding: 24px 30px;
        margin-bottom: 24px;
        box-shadow: 0 10px 30px -10px rgba(0, 240, 255, 0.15), inset 0 0 15px rgba(0, 240, 255, 0.05);
        backdrop-filter: blur(12px);
    }
    .cyber-tag {
        font-family: 'Rajdhani', sans-serif;
        color: #00F0FF;
        font-size: 13px;
        font-weight: 700;
        letter-spacing: 3px;
        text-transform: uppercase;
        text-shadow: 0 0 10px rgba(0, 240, 255, 0.6);
    }
    .cyber-title {
        color: #FFFFFF;
        font-size: 32px;
        font-weight: 800;
        margin: 4px 0 6px 0;
        letter-spacing: -0.5px;
    }
    .cyber-desc {
        color: #94A3B8;
        font-size: 14px;
        margin: 0;
    }

    /* Cards Glassmorphism */
    .glass-card {
        background: rgba(15, 23, 42, 0.65);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 14px;
        padding: 20px;
        backdrop-filter: blur(10px);
        margin-bottom: 16px;
    }

    /* Cartões de Indicadores (KPIs) */
    .kpi-container {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 16px;
        margin-bottom: 24px;
    }
    .kpi-box {
        background: linear-gradient(145deg, rgba(15, 23, 42, 0.9), rgba(30, 41, 59, 0.5));
        border: 1px solid rgba(0, 240, 255, 0.2);
        border-radius: 14px;
        padding: 16px 20px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
    }
    .kpi-label {
        font-size: 12px;
        color: #94A3B8;
        text-transform: uppercase;
        letter-spacing: 1px;
        font-weight: 600;
    }
    .kpi-value {
        font-size: 24px;
        font-weight: 800;
        color: #00F0FF;
        margin-top: 4px;
        font-family: 'Rajdhani', sans-serif;
        text-shadow: 0 0 10px rgba(0, 240, 255, 0.3);
    }

    /* Botão Primário Futurista */
    div.stButton > button:first-child {
        background: linear-gradient(135deg, #00F0FF 0%, #3B82F6 50%, #8B5CF6 100%) !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        letter-spacing: 0.5px !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 14px 28px !important;
        box-shadow: 0 0 20px rgba(0, 240, 255, 0.35) !important;
        transition: all 0.3s ease !important;
    }
    div.stButton > button:first-child:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 0 30px rgba(0, 240, 255, 0.6) !important;
    }

    /* Botão de Download */
    div.stDownloadButton > button:first-child {
        background: linear-gradient(135deg, #10B981 0%, #059669 100%) !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 14px 28px !important;
        box-shadow: 0 0 20px rgba(16, 185, 129, 0.35) !important;
    }
    div.stDownloadButton > button:first-child:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 0 30px rgba(16, 185, 129, 0.6) !important;
    }

    /* Área de Upload */
    div[data-testid="stFileUploader"] {
        border: 2px dashed rgba(0, 240, 255, 0.35) !important;
        border-radius: 16px !important;
        background: rgba(15, 23, 42, 0.4) !important;
        padding: 20px !important;
    }

    /* Tabelas */
    div[data-testid="stDataFrame"] {
        border-radius: 14px;
        overflow: hidden;
        border: 1px solid rgba(255, 255, 255, 0.08);
    }
</style>
""", unsafe_allow_html=True)

# --- CONTROLE DE ACESSO POR SENHA ---
SENHA_DE_ACESSO = st.secrets.get("SENHA_ACESSO", "suasenha123")

def verificar_autenticacao():
    if "autenticado" not in st.session_state:
        st.session_state.autenticado = False

    if not st.session_state.autenticado:
        st.markdown("""
        <div style="max-width: 480px; margin: 60px auto 20px auto; background: rgba(15, 23, 42, 0.85); border: 1px solid rgba(0, 240, 255, 0.3); border-radius: 20px; padding: 30px; box-shadow: 0 15px 35px rgba(0,0,0,0.5); backdrop-filter: blur(15px); text-align: center;">
            <div style="font-size: 36px; margin-bottom: 8px;">🔐</div>
            <div style="font-family: 'Rajdhani', sans-serif; color: #00F0FF; font-size: 13px; letter-spacing: 3px; font-weight: 700;">ACESSO SEGURO</div>
            <h2 style="color: #FFF; margin: 6px 0 10px 0; font-size: 24px;">Conversor de XML em dados</h2>
            <p style="color: #94A3B8; font-size: 13px; margin: 0;">Insira a credencial de segurança para acessar o sistema.</p>
        </div>
        """, unsafe_allow_html=True)
        
        col_esq, col_login, col_dir = st.columns(3)
        with col_login:
            senha = st.text_input("Credencial de Acesso", type="password", label_visibility="collapsed", placeholder="Digite sua senha...")
            if st.button("Autenticar e Entrar", use_container_width=True):
                if senha == SENHA_DE_ACESSO:
                    st.session_state.autenticado = True
                    st.rerun()
                else:
                    st.error("Credencial inválida. Acesso negado.")
        return False
    return True

if not verificar_autenticacao():
    st.stop()

# --- HEADER PRINCIPAL ---
st.markdown("""
<div class="cyber-header">
    <div style="display: flex; align-items: center; justify-content: space-between;">
        <div>
            <span class="cyber-tag">⚡ INTELLIGENCE MATRIX V2.0</span>
            <div class="cyber-title">Conversor de XML em dados</div>
            <p class="cyber-desc">Extração automatizada de XMLs fiscais, consolidação de dados e geração de planilhas inteligentes.</p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# Botão de Logout no Menu Lateral
with st.sidebar:
    st.markdown("### 🛡️ Painel de Segurança")
    st.write("Sessão ativa e protegida.")
    if st.button("Encerrar Sessão", use_container_width=True):
        st.session_state.autenticado = False
        st.rerun()

# 1. Upload de Arquivos
st.markdown("#### 📂 1. Importação de Arquivos XML")
arquivos_xml = st.file_uploader(
    "Selecione ou arraste todos os arquivos XML", 
    type=["xml"], 
    accept_multiple_files=True,
    help="Suporta notas fiscais dos modelos NF-e (55), NFC-e (65) e SAT CF-e (59)."
)

# 2. Configurações de Consolidação e Colunas
st.markdown("#### ⚙️ 2. Parâmetros de Consolidação e Colunas")
col_card1, col_card2 = st.columns(2)

with col_card1:
    st.markdown("""
    <div class="glass-card">
        <span style="color: #00F0FF; font-weight: 700; font-size: 14px;">📊 REGRA DE DUPLICADOS</span>
        <p style="color: #94A3B8; font-size: 13px; margin: 4px 0 12px 0;">Como consolidar itens repetidos ao longo dos meses:</p>
    </div>
    """, unsafe_allow_html=True)
    modo_consolidacao = st.radio(
        "Modo de consolidação:",
        options=[
            "Consolidar produtos (Somar quantidades e calcular custo médio)",
            "Manter apenas o preço da última compra (Custo mais recente)",
            "Listar todas as compras (Histórico completo sem agrupar)"
        ],
        index=0,
        label_visibility="collapsed"
    )

with col_card2
