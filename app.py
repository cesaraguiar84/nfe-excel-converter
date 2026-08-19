import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from io import BytesIO

# Configuração da página
st.set_page_config(page_title="Conversor de NFe para Excel", layout="wide", page_icon="📄")

# --- CONTROLE DE ACESSO POR SENHA ---
SENHA_DE_ACESSO = st.secrets.get("SENHA_ACESSO", "suasenha123")

def verificar_autenticacao():
    if "autenticado" not in st.session_state:
        st.session_state.autenticado = False

    if not st.session_state.autenticado:
        st.title("🔒 Acesso Restrito")
        st.write("Digite a senha de autorização para acessar o conversor.")
        senha = st.text_input("Senha", type="password")
        if st.button("Entrar"):
            if senha == SENHA_DE_ACESSO:
                st.session_state.autenticado = True
                st.rerun()
            else:
                st.error("Senha incorreta. Tente novamente.")
        return False
    return True

if not verificar_autenticacao():
    st.stop()

# --- APLICAÇÃO PRINCIPAL ---
st.title("📄 Conversor de Notas Fiscais (XML) para Excel")
st.write("Selecione ou arraste múltiplos arquivos XML de notas fiscais para consolidar todos os produtos em uma planilha.")

# Botão de logout no menu lateral
with st.sidebar:
    st.write("Painel de Controle")
    if st.button("Sair / Bloquear"):
        st.session_state.autenticado = False
        st.rerun()

arquivos_xml = st.file_uploader(
    "Arraste ou selecione todos os arquivos XML", 
    type=["xml"], 
    accept_multiple_files=True
)

def parse_xml_content(content_bytes, file_name=""):
    # Decodificação robusta para evitar falhas de charset
    try:
        text = content_bytes.decode('utf-8')
    except UnicodeDecodeError:
        try:
            text = content_bytes.decode('iso-8859-1')
        except Exception:
            text = content_bytes.decode('utf-8', errors='ignore')
            
    # Remove declarações de namespace para garantir compatibilidade total
    text = re.sub(r'\sxmlns(:\w+)?="[^"]+"', '', text)
    root = ET.fromstring(text)
    
    # Procura bloco de informações da nota (NFe, NFCe ou CFe SAT)
    inf = root.find('.//infNFe')
    if inf is None:
        inf = root.find('.//infCFe')
    if inf is None:
        inf = root.find('.//infNFCe')
    if inf is None:
        inf = root
        
    # Dados do Fornecedor / Emitente
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
            
    # Extração dos itens
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
            'Nome do Estabelecimento': nome_fornecedor,
            'Endereço': endereco_fornecedor,
            'Código do Produto': cProd,
            'EAN': cEAN,
            'Descrição do Produto': xProd,
            'Quantidade': qCom,
            'Custo Unitário (R$)': vUnCom,
            'Custo Total (R$)': vProd
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

def gerar_excel(df):
    wb = Workbook()
    ws = wb.active
    ws.title = "Notas Fiscais"
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

    # Estiliza apenas a linha 1 (cabeçalho)
    for cell in ws:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center
        cell.border = border

    # Insere dados e formata células
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=False), 2):
        fill = PatternFill(start_color="F9F9F9", end_color="F9F9F9", fill_type="solid") if r_idx % 2 == 0 else PatternFill(fill_type=None)
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.fill = fill
            cell.border = border
            header = headers[c_idx - 1]
            if header == 'Quantidade':
                cell.number_format = '#,##0.00'
                cell.alignment = align_center
            elif header in ['Custo Unitário (R$)', 'Custo Total (R$)']:
                cell.number_format = '"R$ "#,##0.00'
                cell.alignment = align_center
            elif header in ['Código do Produto', 'EAN', 'Arquivo']:
                cell.alignment = align_center
            else:
                cell.alignment = align_left

    # Ajusta largura das colunas
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

if arquivos_xml:
    if st.button("Processar Notas Fiscais", type="primary"):
        with st.spinner("Lendo todos os arquivos XML e montando a planilha..."):
            df_resultados, erros = processar_todos_xmls(arquivos_xml)
            
            if not df_resultados.empty:
                st.success(f"Sucesso! {len(arquivos_xml)} arquivo(s) processado(s) — total de {len(df_resultados)} produtos extraídos.")
                
                if erros:
                    with st.expander("Avisos sobre arquivos não lidos:"):
                        for err in erros:
                            st.warning(err)
                
                # Gera o arquivo Excel (.xlsx)
                excel_data = gerar_excel(df_resultados)
                
                # Botão de download do Excel
                st.download_button(
                    label="📥 Baixar Planilha Excel Formatada (.xlsx)",
                    data=excel_data,
                    file_name="Relatorio_Notas_Fiscais.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
                st.write("### Prévia dos dados extraídos:")
                st.dataframe(df_resultados, use_container_width=True)
            else:
                st.error("Não foi possível extrair produtos dos arquivos enviados.")
                if erros:
                    for err in erros:
                        st.warning(err)
