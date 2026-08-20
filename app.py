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

# Upload dos arquivos
arquivos_xml = st.file_uploader(
    "Arraste ou selecione todos os arquivos XML", 
    type=["xml"], 
    accept_multiple_files=True
)

# Opções de agrupamento e duplicados
st.write("### ⚙️ Opções de Relatório")
modo_consolidacao = st.radio(
    "Como deseja tratar itens comprados mais de uma vez no período?",
    options=[
        "Consolidar produtos (Somar quantidades e calcular custo médio)",
        "Manter apenas o preço da última compra (Custo mais recente)",
        "Listar todas as compras (Histórico completo sem agrupar)"
    ],
    index=0
)

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

    # Dados da Nota Fiscal (Data e Número)
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
            
    # Extração dos Itens
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
            'Data da Emissão': data_emissao,
            'Data_ISO': data_iso,
            'Número da NF': numero_nf,
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

def aplicar_tratamento_duplicados(df, modo):
    if df.empty:
        return df

    # Ordena por data (ISO) para garantir que a última compra seja a mais recente
    if 'Data_ISO' in df.columns:
        df = df.sort_values(by=['Data_ISO', 'Número da NF'])

    group_cols = ['Nome do Estabelecimento', 'Código do Produto', 'EAN', 'Descrição do Produto']

    if modo.startswith("Consolidar produtos"):
        df_agrupado = df.groupby(group_cols, as_index=False).agg(
            Quantidade_Total=('Quantidade', 'sum'),
            Custo_Total=('Custo Total (R$)', 'sum'),
            Ultimo_Custo=('Custo Unitário (R$)', 'last'),
            Qtd_Compras=('Quantidade', 'count'),
            Ultima_Compra=('Data da Emissão', 'last')
        )
        df_agrupado['Custo Médio Unitário (R$)'] = (df_agrupado['Custo_Total'] / df_agrupado['Quantidade_Total']).round(2)
        
        df_agrupado = df_agrupado.rename(columns={
            'Quantidade_Total': 'Quantidade Total',
            'Custo_Total': 'Custo Total (R$)',
            'Ultimo_Custo': 'Último Custo Unitário (R$)',
            'Qtd_Compras': 'Qtd. de Compras',
            'Ultima_Compra': 'Última Compra'
        })
        
        colunas_finais = [
            'Nome do Estabelecimento', 'Código do Produto', 'EAN', 'Descrição do Produto',
            'Qtd. de Compras', 'Quantidade Total', 'Custo Médio Unitário (R$)',
            'Último Custo Unitário (R$)', 'Custo Total (R$)', 'Última Compra'
        ]
        return df_agrupado[colunas_finais]

    elif modo.startswith("Manter apenas o preço da última compra"):
        df_ultimo = df.drop_duplicates(subset=group_cols, keep='last').copy()
        colunas_finais = [
            'Data da Emissão', 'Número da NF', 'Nome do Estabelecimento',
            'Código do Produto', 'EAN', 'Descrição do Produto',
            'Quantidade', 'Custo Unitário (R$)', 'Custo Total (R$)', 'Arquivo'
        ]
        colunas_existentes = [c for c in colunas_finais if c in df_ultimo.columns]
        return df_ultimo[colunas_existentes]

    else:
        # Modo detalhado (todas as linhas)
        colunas_finais = [
            'Data da Emissão', 'Número da NF', 'Arquivo', 'Nome do Estabelecimento',
            'Endereço', 'Código do Produto', 'EAN', 'Descrição do Produto',
            'Quantidade', 'Custo Unitário (R$)', 'Custo Total (R$)'
        ]
        colunas_existentes = [c for c in colunas_finais if c in df.columns]
        return df[colunas_existentes]

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

    # 1. Estiliza o cabeçalho célula a célula
    for col_idx in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=col_idx)
        c.fill = header_fill
        c.font = header_font
        c.alignment = align_center
        c.border = border

    # 2. Insere dados e aplica formatos
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
            elif header in ['Qtd. de Compras', 'Número da NF']:
                cell.number_format = '#,##0'
                cell.alignment = align_center
            elif 'Custo' in header or 'Preço' in header or 'Valor' in header:
                cell.number_format = '"R$ "#,##0.00'
                cell.alignment = align_center
            elif header in ['Código do Produto', 'EAN', 'Arquivo', 'Data da Emissão', 'Última Compra']:
                cell.alignment = align_center
            else:
                cell.alignment = align_left

    # 3. Ajusta largura das colunas
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
        with st.spinner("Processando os arquivos XML e gerando a planilha..."):
            df_bruto, erros = processar_todos_xmls(arquivos_xml)
            
            if not df_bruto.empty:
                # Aplica o filtro/consolidação escolhido
                df_final = aplicar_tratamento_duplicados(df_bruto, modo_consolidacao)
                
                total_itens_extraidos = len(df_bruto)
                total_itens_unicos = len(df_final)
                
                st.success(
                    f"Sucesso! {len(arquivos_xml)} arquivo(s) lido(s). "
                    f"Total de {total_itens_extraidos} compras encontradas ➔ {total_itens_unicos} produtos gerados na planilha."
                )
                
                if erros:
                    with st.expander("Avisos sobre arquivos não lidos:"):
                        for err in erros:
                            st.warning(err)
                
                excel_data = gerar_excel(df_final, titulo_aba="Relatório")
                
                st.download_button(
                    label="📥 Baixar Planilha Excel Formatada (.xlsx)",
                    data=excel_data,
                    file_name="Relatorio_Notas_Fiscais.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
                st.write("### Prévia do Relatório Gerado:")
                st.dataframe(df_final, use_container_width=True)
            else:
                st.error("Não foi possível extrair produtos dos arquivos enviados.")
                if erros:
                    for err in erros:
                        st.warning(err)
