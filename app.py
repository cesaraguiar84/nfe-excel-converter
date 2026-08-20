import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from io import BytesIO

# 1. Configuração da página
st.set_page_config(
    page_title="Conversor de XML em dados",
    layout="wide",
    page_icon="⚡"
)

# 2. Estilo Visual Futurista (Dark Mode + Glassmorphism)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700;800&family=Rajdhani:wght@600;700&display=swap');

    .stApp {
        background: radial-gradient(circle at 15% 15%, rgba(0, 240, 255, 0.05) 0%, transparent 40%),
                    radial-gradient(circle at 85% 85%, rgba(138, 43, 226, 0.05) 0%, transparent 40%),
                    #090D14;
        color: #E2E8F0;
        font-family: 'Inter', sans-serif;
    }

    .cyber-header {
        background: linear-gradient(135deg, rgba(15, 23, 42, 0.8), rgba(30, 41, 59, 0.6));
        border: 1px solid rgba(0, 240, 255, 0.25);
        border-radius: 18px;
        padding: 24px 30px;
        margin-bottom: 24px;
        box-shadow: 0 10px 30px -10px rgba(0, 240, 255, 0.15);
        backdrop-filter: blur(12px);
    }
    .cyber-tag {
        font-family: 'Rajdhani', sans-serif;
        color: #00F0FF;
        font-size: 13px;
        font-weight: 700;
        letter-spacing: 3px;
        text-transform: uppercase;
    }
    .cyber-title {
        color: #FFFFFF;
        font-size: 32px;
        font-weight: 800;
        margin: 4px 0 6px 0;
    }
    .cyber-desc {
        color: #94A3B8;
        font-size: 14px;
        margin: 0;
    }

    .glass-card {
        background: rgba(15, 23, 42, 0.65);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 14px;
        padding: 20px;
        backdrop-filter: blur(10px);
        margin-bottom: 16px;
    }

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
    }

    div.stButton > button:first-child {
        background: linear-gradient(135deg, #00F0FF 0%, #3B82F6 50%, #8B5CF6 100%) !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 14px 28px !important;
    }

    div.stDownloadButton > button:first-child {
        background: linear-gradient(135deg, #10B981 0%, #059669 100%) !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 14px 28px !important;
    }
</style>
""", unsafe_allow_html=True)

# 3. Controle de Acesso por Senha
SENHA_DE_ACESSO = st.secrets.get("SENHA_ACESSO", "suasenha123")

def verificar_autenticacao():
    if "autenticado" not in st.session_state:
        st.session_state.autenticado = False

    if not st.session_state.autenticado:
        st.markdown("""
        <div style="max-width: 480px; margin: 60px auto 20px auto; background: rgba(15, 23, 42, 0.85); border: 1px solid rgba(0, 240, 255, 0.3); border-radius: 20px; padding: 30px; text-align: center;">
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

# 4. Header Principal
st.markdown("""
<div class="cyber-header">
    <div>
        <span class="cyber-tag">⚡ INTELLIGENCE MATRIX V2.0</span>
        <div class="cyber-title">Conversor de XML em dados</div>
        <p class="cyber-desc">Extração automatizada de XMLs fiscais, consolidação analítica e exportação em planilha formatada.</p>
    </div>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 🛡️ Painel de Segurança")
    st.write("Sessão ativa e autenticada.")
    if st.button("Encerrar Sessão", use_container_width=True):
        st.session_state.autenticado = False
        st.rerun()

# 5. Upload de Arquivos
st.markdown("#### 📂 1. Importação de Arquivos XML")
arquivos_xml = st.file_uploader("Selecione ou arraste todos os arquivos XML", type=["xml"], accept_multiple_files=True)

# 6. Configurações de Consolidação e Colunas
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

with col_card2:
    st.markdown("""
    <div class="glass-card">
        <span style="color: #8B5CF6; font-weight: 700; font-size: 14px;">📌 SELEÇÃO DE CAMPOS EXPORTADOS</span>
        <p style="color: #94A3B8; font-size: 13px; margin: 4px 0 8px 0;">Marque as colunas que devem constar no relatório:</p>
    </div>
    """, unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        f_arquivo = st.checkbox("Arquivo", value=True)
        f_data = st.checkbox("Data", value=True)
        f_estab = st.checkbox("Estabelecimento", value=True)
        f_end = st.checkbox("Endereço", value=True)
        f_cod = st.checkbox("Código do Produto", value=True)
    with c2:
        f_ean = st.checkbox("EAN", value=True)
        f_desc = st.checkbox("Descrição do Produto", value=True)
        f_qtd = st.checkbox("Quantidade", value=True)
        f_unit = st.checkbox("Custo Unitário", value=True)
        f_total = st.checkbox("Custo Total", value=True)

colunas_selecionadas = []
if f_arquivo: colunas_selecionadas.append("Arquivo")
if f_data: colunas_selecionadas.append("Data")
if f_estab: colunas_selecionadas.append("Estabelecimento")
if f_end: colunas_selecionadas.append("Endereço")
if f_cod: colunas_selecionadas.append("Código do Produto")
if f_ean: colunas_selecionadas.append("EAN")
if f_desc: colunas_selecionadas.append("Descrição do Produto")
if f_qtd: colunas_selecionadas.append("Quantidade")
if f_unit: colunas_selecionadas.append("Custo Unitário")
if f_total: colunas_selecionadas.append("Custo Total")

# 7. Funções de Processamento de XML
def parse_xml_content(content_bytes, file_name=""):
    try:
        text = content_bytes.decode('utf-8')
    except UnicodeDecodeError:
        try:
            text = content_bytes.decode('iso-8859-1')
        except Exception:
            text = content_bytes.decode('utf-8', errors='ignore')
            
    text = re.sub(r'\sxmlns(:\w+)?="[^"]+"', '', text)
    root = ET.fromstring(text)
    
    inf = root.find('.//infNFe')
    if inf is None:
        inf = root.find('.//infCFe')
    if inf is None:
        inf = root.find('.//infNFCe')
    if inf is None:
        inf = root

    numero_nf = ''
    data_emissao = ''
    data_iso = ''
    ide = inf.find('.//ide')
    if ide is not None:
        nNF_el = ide.find('nNF')
        if nNF_el is not None and nNF_el.text:
            numero_nf = nNF_el.text
            
        dhEmi_el = ide.find('dhEmi') or ide.find('dEmi')
        if dhEmi_el is not None and dhEmi_el.text:
            raw_date = dhEmi_el.text[:10]
            data_iso = raw_date
            parts = raw_date.split('-')
            if len(parts) == 3:
                data_emissao = f"{parts}/{parts}/{parts[0]}"
            else:
                data_emissao = raw_date
        
    nome_fornecedor = ''
    endereco_fornecedor = ''
    emit = inf.find('.//emit')
    if emit is not None:
        xNome = emit.find('xNome')
        nome_fornecedor = xNome.text if xNome is not None and xNome.text else ''
        
        ender = emit.find('enderEmit')
        if ender is not None:
            def get_field(tag):
                el = ender.find(tag)
                return el.text if el is not None and el.text else ''
            
            lgr = get_field('xLgr')
            nro = get_field('nro')
            bairro = get_field('xBairro')
            mun = get_field('xMun')
            uf = get_field('UF')
            endereco_fornecedor = f"{lgr}, {nro} - {bairro}, {mun} - {uf}".strip(" ,-")
            
    dets = inf.findall('.//det')
    itens = []
    for det in dets:
        prod = det.find('prod')
        if prod is None:
            continue
            
        def get_val(tag, default=''):
            el = prod.find(tag)
            return el.text if el is not None and el.text else default
            
        cProd = get_val('cProd')
        cEAN = get_val('cEAN')
        xProd = get_val('xProd')
        qCom_str = get_val('qCom', '0')
        vUnCom_str = get_val('vUnCom', '0')
        vProd_str = get_val('vProd', '0')
        
        try:
            qCom = float(qCom_str)
        except ValueError:
            qCom = 0.0
            
        try:
            vUnCom = float(vUnCom_str)
        except ValueError:
            vUnCom = 0.0
            
        try:
            vProd = float(vProd_str)
        except ValueError:
            vProd = 0.0
            
        itens.append({
            'Arquivo': file_name,
            'Data': data_emissao,
            'Data_ISO': data_iso,
            'Número da NF': numero_nf,
            'Estabelecimento': nome_fornecedor,
            'Endereço': endereco_fornecedor,
            'Código do Produto': cProd,
            'EAN': cEAN,
            'Descrição do Produto': xProd,
            'Quantidade': qCom,
            'Custo Unitário': vUnCom,
            'Custo Total': vProd
        })
        
    return itens

def processar_todos_xmls(arquivos):
    todos_itens = []
    arquivos_com_erro = []
    
    for arquivo in arquivos:
        try:
            conteudo = arquivo.read()
            itens = parse_xml_content(conteudo, file_name=arquivo.name)
            if itens:
                todos_itens.extend(itens)
            else:
                arquivos_com_erro.append(f"{arquivo.name} (nenhum produto encontrado)")
        except Exception as e:
            arquivos_com_erro.append(f"{arquivo.name} (erro: {e})")
            
    return pd.DataFrame(todos_itens), arquivos_com_erro

def aplicar_tratamento_e_colunas(df, modo, colunas_escolhidas):
    if df.empty:
        return df

    if 'Data_ISO' in df.columns:
        df = df.sort_values(by=['Data_ISO', 'Número da NF'])

    group_cols = ['Estabelecimento', 'Endereço', 'Código do Produto', 'EAN', 'Descrição do Produto']

    if modo.startswith("Consolidar produtos"):
        df_agrupado = df.groupby(group_cols, as_index=False).agg(
            Quantidade=('Quantidade', 'sum'),
            Custo_Total=('Custo Total', 'sum'),
            Ultimo_Custo=('Custo Unitário', 'last'),
            Data=('Data', 'last'),
            Arquivo=('Arquivo', lambda x: ', '.join(x.unique()))
        )
        df_agrupado['Custo Unitário'] = (df_agrupado['Custo_Total'] / df_agrupado['Quantidade']).round(2)
        df_agrupado['Custo Total'] = df_agrupado['Custo_Total'].round(2)
        df_final = df_agrupado

    elif modo.startswith("Manter apenas o preço da última compra"):
        df_final = df.drop_duplicates(subset=group_cols, keep='last').copy()

    else:
        df_final = df.copy()

    colunas_validas = [c for c in colunas_escolhidas if c in df_final.columns]
    return df_final[colunas_validas]

def gerar_excel(df, titulo_aba="Notas Fiscais"):
    wb = Workbook()
    ws = wb.active
    ws.title = titulo_aba[:31]
    headers = list(df.columns)
    ws.append(headers)

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    for col_idx in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=col_idx)
        c.fill = header_fill
        c.font = header_font
        c.alignment = align_center
        c.border = border

    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=False), 2):
        fill = PatternFill(start_color="F9F9F9", end_color="F9F9F9", fill_type="solid") if r_idx % 2 == 0 else PatternFill(fill_type=None)
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.fill = fill
            cell.border = border
            header = headers[c_idx - 1]
            
            if 'Quantidade' in header:
                cell.number_format = '#,##0.00'
                cell.alignment = align_center
            elif 'Custo' in header or 'Preço' in header or 'Valor' in header:
                cell.number_format = '"R$ "#,##0.00'
                cell.alignment = align_center
            elif header in ['Código do Produto', 'EAN', 'Arquivo', 'Data']:
                cell.alignment = align_center
            else:
                cell.alignment = align_left

    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except Exception:
                pass
        ws.column_dimensions[column].width = min(max_length + 3, 60)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# 8. Execução
if arquivos_xml:
    if not colunas_selecionadas:
        st.warning("⚠️ Marque ao menos uma coluna acima para gerar o relatório.")
    else:
        if st.button("⚡ Executar Processamento e Gerar Planilha", type="primary", use_container_width=True):
            with st.spinner("Processando dados e estruturando planilha..."):
                df_bruto, erros = processar_todos_xmls(arquivos_xml)
                
                if not df_bruto.empty:
                    df_final = aplicar_tratamento_e_colunas(df_bruto, modo_consolidacao, colunas_selecionadas)
                    
                    total_arquivos = len(arquivos_xml)
                    total_compras = len(df_bruto)
                    total_linhas = len(df_final)
                    valor_total = df_bruto['Custo Total'].sum() if 'Custo Total' in df_bruto.columns else 0.0
                    
                    st.markdown(f"""
                    <div class="kpi-container">
                        <div class="kpi-box">
                            <div class="kpi-label">📁 Arquivos Lidos</div>
                            <div class="kpi-value">{total_arquivos}</div>
                        </div>
                        <div class="kpi-box">
                            <div class="kpi-label">📦 Total de Compras</div>
                            <div class="kpi-value">{total_compras}</div>
                        </div>
                        <div class="kpi-box">
                            <div class="kpi-label">💎 Produtos no Relatório</div>
                            <div class="kpi-value">{total_linhas}</div>
                        </div>
                        <div class="kpi-box">
                            <div class="kpi-label">💰 Valor Consolidado</div>
                            <div class="kpi-value">R$ {valor_total:,.2f}</div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    if erros:
                        with st.expander("⚠️ Alertas de arquivos com inconsistência:"):
                            for err in erros:
                                st.warning(err)
                    
                    excel_data = gerar_excel(df_final, titulo_aba="Relatório NFe")
                    
                    st.download_button(
                        label="📥 Baixar Planilha Excel Formatada (.xlsx)",
                        data=excel_data,
                        file_name="Relatorio_Notas_Fiscais.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                    st.markdown("#### 🔍 Prévia dos Dados Estruturados")
                    st.dataframe(df_final, use_container_width=True)
                else:
                    st.error("Nenhum produto válido foi identificado nos arquivos enviados.")
                    if erros:
                        for err in erros:
                            st.warning(err)
