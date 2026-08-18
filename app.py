import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from io import BytesIO

# Configuração da página
st.set_page_config(page_title="Conversor de NFe para Excel", layout="centered", page_icon="📄")

# --- CONTROLE DE ACESSO POR SENHA ---
# Defina aqui a senha que você passará para os 2 colaboradores
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
st.write("Selecione ou arraste os arquivos XML das notas fiscais e baixe o relatório consolidado.")

# Botão de logout no menu lateral
with st.sidebar:
    st.write("Painel de Controle")
    if st.button("Sair / Bloquear"):
        st.session_state.autenticado = False
        st.rerun()

arquivos_xml = st.file_uploader(
    "Selecione um ou mais arquivos XML", 
    type=["xml"], 
    accept_multiple_files=True
)

def processar_xmls(arquivos):
    rows = []
    ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
    
    for arquivo in arquivos:
        try:
            tree = ET.parse(arquivo)
            root = tree.getroot()
            
            # Busca infNFe com ou sem namespace
            infNFe = root.find('.//nfe:infNFe', ns)
            if infNFe is None:
                infNFe = root.find('.//infNFe')
                ns_used = {}
            else:
                ns_used = ns
            
            if infNFe is None:
                continue
                
            # Dados do Fornecedor / Emitente
            emit = infNFe.find('nfe:emit', ns_used) if ns_used else infNFe.find('emit')
            nome_fornecedor = ''
            endereco_fornecedor = ''
            
            if emit is not None:
                xNome = emit.find('nfe:xNome', ns_used) if ns_used else emit.find('xNome')
                nome_fornecedor = xNome.text if xNome is not None and xNome.text else ''
                
                ender = emit.find('nfe:enderEmit', ns_used) if ns_used else emit.find('enderEmit')
                if ender is not None:
                    def get_field(tag):
                        el = ender.find(f'nfe:{tag}', ns_used) if ns_used else ender.find(tag)
                        return el.text if el is not None and el.text else ''
                    
                    lgr = get_field('xLgr')
                    nro = get_field('nro')
                    bairro = get_field('xBairro')
                    mun = get_field('xMun')
                    uf = get_field('UF')
                    endereco_fornecedor = f"{lgr}, {nro} - {bairro}, {mun} - {uf}".strip(" ,-")
            
            # Dados dos Itens / Produtos
            dets = infNFe.findall('nfe:det', ns_used) if ns_used else infNFe.findall('det')
            for det in dets:
                prod = det.find('nfe:prod', ns_used) if ns_used else det.find('prod')
                if prod is None:
                    continue
                
                def get_prod_val(tag, default=''):
                    el = prod.find(f'nfe:{tag}', ns_used) if ns_used else prod.find(tag)
                    return el.text if el is not None and el.text else default
                
                cProd = get_prod_val('cProd')
                cEAN = get_prod_val('cEAN')
                xProd = get_prod_val('xProd')
                qCom_str = get_prod_val('qCom', '0')
                vUnCom_str = get_prod_val('vUnCom', '0')
                vProd_str = get_prod_val('vProd', '0')
                
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
                
                rows.append({
                    'Nome do Estabelecimento': nome_fornecedor,
                    'Endereço': endereco_fornecedor,
                    'Código do Produto': cProd,
                    'EAN': cEAN,
                    'Descrição do Produto': xProd,
                    'Quantidade': qCom,
                    'Custo Unitário (R$)': vUnCom,
                    'Custo Total (R$)': vProd
                })
        except Exception as e:
            st.error(f"Erro ao processar o arquivo {getattr(arquivo, 'name', 'XML')}: {e}")
            
    return pd.DataFrame(rows)

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

    for cell in ws:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center
        cell.border = border

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
            elif header in ['Código do Produto', 'EAN']:
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

if arquivos_xml:
    if st.button("Processar Notas Fiscais", type="primary"):
        with st.spinner("Extraindo informações dos XMLs e gerando planilha..."):
            df_resultados = processar_xmls(arquivos_xml)
            
            if not df_resultados.empty:
                st.success(f"Processamento concluído! {len(df_resultados)} itens extraídos.")
                st.dataframe(df_resultados.head(10))
                
                excel_data = gerar_excel(df_resultados)
                
                st.download_button(
                    label="📥 Baixar Planilha Formatada (.xlsx)",
                    data=excel_data,
                    file_name="Relatorio_Notas_Fiscais.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Nenhum item válido encontrado nos arquivos XML fornecidos.")