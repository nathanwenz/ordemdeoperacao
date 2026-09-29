import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import os
import re
import base64
import pdfplumber
import docx
from datetime import datetime

# ReportLab para geração garantida e nativa de PDF (sem depender de LibreOffice)
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch

# Configuração da página e tema
st.set_page_config(page_title="18º BPM — Gerador de Ordem de Serviço", page_icon="📑", layout="centered")

# Busca flexível da imagem do brasão
def obter_caminho_brasao():
    for nome in ["brasao.png", "brasao.PNG", "Brasao.png", "BRASAO.PNG", "brasao.jpg", "brasao.jpeg"]:
        if os.path.exists(nome):
            return nome
    return None

caminho_brasao = obter_caminho_brasao()

# 🎨 MODO ESCURO COM BRASÃO EM MARCA D'ÁGUA EM TELA CHEIA (LEVEMENTE NO FUNDO DE TODO O SITE)
def gerar_css_app(caminho_img):
    css_base = """
    <style>
        /* Estilização dos Containers e Caixas com leve transparência para ver o fundo */
        div[data-testid="stFileUploader"], div[data-testid="stTextInput"] {
            background-color: rgba(30, 34, 45, 0.88) !important;
            border-radius: 10px;
            padding: 10px;
            border: 1px solid #2E364A;
        }
        /* Botões operacionais */
        .stButton>button {
            background-color: #002060;
            color: #FFFFFF;
            font-weight: bold;
            border-radius: 8px;
            border: 1px solid #1E3A8A;
            width: 100%;
            padding: 10px;
        }
        .stButton>button:hover {
            background-color: #1E40AF;
            border-color: #3B82F6;
        }
        /* Títulos */
        h1, h2, h3 {
            color: #F3F4F6 !important;
        }
    </style>
    """
    
    if caminho_img and os.path.exists(caminho_img):
        try:
            with open(caminho_img, "rb") as f:
                encoded = base64.b64encode(f.read()).decode("utf-8")
            ext = caminho_img.split(".")[-1].lower()
            mime = "image/png" if "png" in ext else "image/jpeg"
            bg_css = f"""
            <style>
                .stApp {{
                    background-color: #0E1117;
                    color: #E0E6ED;
                    background-image: linear-gradient(rgba(14, 17, 23, 0.92), rgba(14, 17, 23, 0.92)), url('data:{mime};base64,{encoded}');
                    background-size: contain;
                    background-repeat: no-repeat;
                    background-position: center center;
                    background-attachment: fixed;
                }}
            </style>
            """
            return bg_css + css_base
        except Exception:
            pass

    return """
    <style>
        .stApp {
            background-color: #0E1117;
            color: #E0E6ED;
        }
    </style>
    """ + css_base

st.markdown(gerar_css_app(caminho_brasao), unsafe_allow_html=True)

# 🏷️ ASSINATURA MOVIDA PARA A ESQUERDA
st.markdown("""
<div style="position: fixed; bottom: 15px; right: 140px; text-align: right; color: #9CA3AF; font-size: 12px; font-family: sans-serif; z-index: 999999; line-height: 1.4; background-color: rgba(14, 17, 23, 0.9); padding: 6px 12px; border-radius: 6px; border: 1px solid #2E364A;">
    Desenvolvido por:<br>
    <strong style="color: #60A5FA; font-size: 13px;">Nathan Wenzel</strong>
</div>
""", unsafe_allow_html=True)

# 🔒 CONFIGURAÇÃO DA SENHA DE ACESSO
SENHA_CORRETA = "18BPM2026"

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

# Tela de Login
if not st.session_state.autenticado:
    if caminho_brasao:
        col1, col2, col3 = st.columns(3)
        with col2:
            st.image(caminho_brasao, width=150)

    st.title("🔒 Acesso Restrito — 18º BPM")
    st.write("Digite a senha de acesso para utilizar o Gerador de Ordens de Serviço (OS).")
    
    senha_input = st.text_input("Senha de acesso:", type="password")
    
    if st.button("Entrar"):
        if senha_input == SENHA_CORRETA:
            st.session_state.autenticado = True
            st.success("Acesso liberado!")
            st.rerun()
        else:
            st.error("Senha incorreta! Verifique e tente novamente.")
    st.stop()

# =========================================================
# FUNÇÕES DE EXTRAÇÃO, LIMPEZA DE ASSINATURA E PARSER DE OO
# =========================================================

MESES = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
    5: "maio", 6: "junho", 7: "julho", 8: "agosto",
    9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
}

def data_atual_extenso():
    now = datetime.now()
    return f"{now.day} de {MESES[now.month]} de {now.year}"

def limpar_assinaturas_eletronicas(texto):
    """Remove linhas e metadados de assinatura eletrônica, hash e e-Protocolo do texto extraído"""
    if not texto:
        return ""
    
    padroes_assinatura = [
        r'Documento\s+assinado\s+digitalmente.*',
        r'Inserido\s+ao\s+protocolo.*',
        r'conforme\s+MP\s+n[º°\.]?\s*2\.?200-2/2001.*',
        r'Validação:.*',
        r'https?://[^\s]*eprotocolo[^\s]*',
        r'www\.[^\s]*eprotocolo[^\s]*',
        r'Código\s+de\s+autenticidade:.*',
        r'Assinado\s+eletronicamente\s+por:.*',
        r'Assinatura\s+Qualificada\s+efetuada\s+por:.*',
        r'Chave\s+de\s+Autenticação:.*',
        r'Para\s+verificar\s+a\s+autenticidade.*',
        r'e-Protocolo\s*\d+.*',
        r'SHA-?256:.*',
        r'Hash\s*:.*'
    ]
    
    linhas = texto.split('\n')
    linhas_limpas = []
    
    for linha in linhas:
        descartar = False
        for p in padroes_assinatura:
            if re.search(p, linha, re.IGNORECASE):
                descartar = True
                break
        if not descartar:
            linhas_limpas.append(linha)
            
    return '\n'.join(linhas_limpas)

def extrair_texto_arquivo(uploaded_file):
    """Extrai todo o texto de um arquivo PDF ou DOCX enviado e remove assinaturas digitais"""
    ext = uploaded_file.name.split(".")[-1].lower()
    text = ""
    
    if ext == "pdf":
        uploaded_file.seek(0)
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text += t + "\n"
        uploaded_file.seek(0)
    elif ext in ["docx", "doc"]:
        uploaded_file.seek(0)
        doc = docx.Document(uploaded_file)
        for p in doc.paragraphs:
            text += p.text + "\n"
        for t in doc.tables:
            for r in t.rows:
                row_str = " | ".join([c.text.strip() for c in r.cells])
                text += row_str + "\n"
        uploaded_file.seek(0)
        
    text_limpo = limpar_assinaturas_eletronicas(text)
    return text_limpo

def extrair_secao(texto, inicio_regex, fim_regex):
    """Auxiliar para extrair bloco de texto entre duas seções usando Regex"""
    pattern = f"(?:{inicio_regex})(.*?)(?=(?:{fim_regex})|$)"
    match = re.search(pattern, texto, re.DOTALL | re.IGNORECASE)
    if match and match.group(1):
        res = match.group(1).strip()
        res = re.sub(r'\n{3,}', '\n\n', res)
        return res
    return ""

def parsear_ordem_operacao(texto):
    """Analisa a Ordem de Operação e extrai os campos estruturados para a OS"""
    dados = {}
    
    # 1. Número da OO
    match_num = re.search(r'ORDEM DE OPERAÇÃO\s*(?:Nº|N°|Nº\.|N°\.|N°\s*|Nº\s*)?(\d+/\d+)', texto, re.IGNORECASE)
    dados['num_oo'] = match_num.group(1) if match_num else "000/2026"

    # 2. Nome da Operação
    match_nome = re.search(r'“([^”]+)”|"([^"]+)"', texto)
    if match_nome:
        dados['nome_op'] = match_nome.group(1) or match_nome.group(2)
    else:
        match_nome2 = re.search(r'OPERAÇÃO\s+([A-Z0-9\s–\-]{4,})', texto)
        dados['nome_op'] = match_nome2.group(0).strip() if match_nome2 else "OPERAÇÃO POLICIAL"

    dados['nome_op'] = dados['nome_op'].replace('\n', ' ').strip().upper()

    # 3. Finalidade
    finalidade = extrair_secao(texto, r'1\.?\s*FINALIDADE', r'\n+2\.?\s*SITUAÇÃO')
    if not finalidade:
        finalidade = extrair_secao(texto, r'FINALIDADE', r'SITUAÇÃO')
    dados['finalidade'] = finalidade if finalidade else "Realizar ações de policiamento ostensivo preventivo e preservação da ordem pública."

    # 4. Situação / Contexto
    situacao = extrair_secao(texto, r'2\.?\s*SITUAÇÃO', r'\n+3\.?\s*MISSÃO')
    dados['situacao'] = situacao if situacao else "Ações de policiamento ostensivo e preventivo para a manutenção da ordem pública."

    # 5. Missão / Execução
    execucao = extrair_secao(texto, r'4\.?\s*EXECUÇÃO|3\.?\s*MISSÃO', r'\n+5\.?\s*ADMINISTRAÇÃO|\n+5\.?\s*LOGÍSTICA')
    if not execucao:
        execucao = extrair_secao(texto, r'EXECUÇÃO', r'ADMINISTRAÇÃO')
    dados['execucao'] = execucao if execucao else "Atuação integrada das equipes operacionais do 18º BPM em conformidade com o planejamento."

    # 6. Logística / Administração
    logistica = extrair_secao(texto, r'5\.?\s*ADMINISTRAÇÃO|5\.?\s*LOGÍSTICA', r'\n+6\.?\s*RELATÓRIOS|\n+6\.?\s*PRESCRIÇÕES')
    dados['logistica'] = logistica if logistica else "Uniforme: Orgânico da OPM (4º RUPM).\nArmamento e equipamento: Orgânico compatível com o serviço.\nTransporte: Viaturas operacionais do 18º BPM."

    # 7. Relatórios e SISGCOP
    relatorios = extrair_secao(texto, r'6\.?\s*RELATÓRIOS', r'\n+7\.?\s*PRESCRIÇÕES|\n+REFERÊNCIAS')
    
    match_sisgcop = re.search(r'(\d{5,6})\s*[\-–]?\s*[\"“]?OPERAÇÃO', texto, re.IGNORECASE)
    if not match_sisgcop:
        match_sisgcop = re.search(r'SISGCOP[^\d]*(\d{5,6})', texto, re.IGNORECASE)
    
    num_sisgcop = match_sisgcop.group(1) if match_sisgcop else ""
    
    if relatorios:
        dados['relatorios'] = relatorios
    else:
        dados['relatorios'] = f"Os resultados obtidos deverão ser lançados no SISGCOP{' sob o código ' + num_sisgcop if num_sisgcop else ''} até o término da operação. Confecção dos Boletins de Ocorrência (BOU) no SADE."

    # 8. Prescrições Diversas
    prescricoes = extrair_secao(texto, r'7\.?\s*PRESCRIÇÕES DIVERSAS|PRESCRIÇÕES DIVERSAS', r'\n+REFERÊNCIAS|\n+DISTRIBUIÇÃO')
    dados['prescricoes'] = prescricoes if prescricoes else "Os policiais militares deverão atuar com bom senso, urbanidade, legalidade e estrito cumprimento do dever legal. Preleção obrigatória antes do início do serviço."

    # 9. Referências Padrão
    dados['referencias'] = f"a. Constituição da República Federativa do Brasil de 1988;\nb. Constituição do Estado do Paraná de 1989;\nc. Lei n.º 22.354/2025 – Lei de Organização Básica da PMPR;\nd. Ordem de Operação nº {dados['num_oo']} – 2º CRPM ({dados['nome_op']});\ne. Determinação do Comandante do 18º BPM."

    return dados

# =========================================================
# GERADOR DO DOCUMENTO WORD (.DOCX)
# =========================================================

def gerar_ordem_servico_docx(fields):
    doc = Document()

    # Margens Amplas e Elegantes
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # Tabela de Cabeçalho Institucional
    table_hdr = doc.add_table(rows=1, cols=2)
    table_hdr.autofit = False
    table_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    r0 = table_hdr.rows
    r0.cells[0].width = Inches(3.5)
    r0.cells[1].width = Inches(3.0)

    # Célula Esquerda (Unidade)
    p_left = r0.cells[0].paragraphs[0]
    p_left.paragraph_format.space_after = Pt(2)
    p_left.paragraph_format.line_spacing = 1.2
    r_l = p_left.add_run("PMPR\n2º CRPM/18º BPM\nP/3")
    r_l.bold = True
    r_l.font.name = "Arial"
    r_l.font.size = Pt(10)

    # Célula Direita (Local, Data, OS)
    p_right = r0.cells[1].paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_right.paragraph_format.space_after = Pt(2)
    p_right.paragraph_format.line_spacing = 1.2
    r_r = p_right.add_run(f"Cornélio Procópio, PR.\nEm {fields.get('data_expedicao', data_atual_extenso())}\nORDEM DE SERVIÇO Nº {fields.get('num_os', '077')}")
    r_r.bold = True
    r_r.font.name = "Arial"
    r_r.font.size = Pt(10)

    # Linha Divisória
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_after = Pt(10)
    p_div.paragraph_format.space_before = Pt(6)
    r_div = p_div.add_run("___________________________________________________________________")
    r_div.bold = True
    r_div.font.size = Pt(9)

    # Título da Operação
    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_titulo.paragraph_format.space_before = Pt(6)
    p_titulo.paragraph_format.space_after = Pt(16)
    r_tit = p_titulo.add_run(f"“{fields.get('nome_operacao', '').strip().upper()}”")
    r_tit.bold = True
    r_tit.font.name = "Arial"
    r_tit.font.size = Pt(12)
    r_tit.font.color.rgb = RGBColor(0, 32, 96)

    # Seções Estruturadas da OS
    secoes = [
        ("1. FINALIDADE", fields.get('finalidade', '')),
        ("2. REFERÊNCIAS", fields.get('referencias', '')),
        ("3. SITUAÇÃO E OBJETIVOS", fields.get('situacao', '')),
        ("4. EXECUÇÃO E MISSÃO", fields.get('execucao', '')),
        ("5. ADMINISTRAÇÃO E LOGÍSTICA", fields.get('logistica', '')),
        ("6. RELATÓRIOS E SISGCOP", fields.get('relatorios', '')),
        ("7. PRESCRIÇÕES DIVERSAS", fields.get('prescricoes', ''))
    ]

    for tit, conteudo in secoes:
        if conteudo and conteudo.strip():
            p_sec = doc.add_paragraph()
            p_sec.paragraph_format.space_before = Pt(12)
            p_sec.paragraph_format.space_after = Pt(4)
            r_sec = p_sec.add_run(tit)
            r_sec.bold = True
            r_sec.font.name = "Arial"
            r_sec.font.size = Pt(11)

            linhas = conteudo.strip().split('\n')
            for linha in linhas:
                if linha.strip():
                    p_cnt = doc.add_paragraph()
                    p_cnt.paragraph_format.space_before = Pt(0)
                    p_cnt.paragraph_format.space_after = Pt(6)
                    p_cnt.paragraph_format.line_spacing = 1.2
                    r_cnt = p_cnt.add_run(linha.strip())
                    r_cnt.font.name = "Arial"
                    r_cnt.font.size = Pt(10)

    # Assinatura do Comandante do 18º BPM
    p_ass = doc.add_paragraph()
    p_ass.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ass.paragraph_format.space_before = Pt(28)
    p_ass.paragraph_format.space_after = Pt(4)
    
    r_ass1 = p_ass.add_run("(Assinado eletronicamente)\n")
    r_ass1.italic = True
    r_ass1.font.name = "Arial"
    r_ass1.font.size = Pt(10)

    r_ass2 = p_ass.add_run(f"{fields.get('nome_comandante', 'Ten.-Cel. QOEM PM Helder de Lima Dantas Junior')},\n")
    r_ass2.bold = True
    r_ass2.font.name = "Arial"
    r_ass2.font.size = Pt(10.5)

    r_ass3 = p_ass.add_run(f"{fields.get('cargo_comandante', 'Comandante do 18º BPM.')}")
    r_ass3.bold = True
    r_ass3.font.name = "Arial"
    r_ass3.font.size = Pt(10)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# =========================================================
# GERADOR NATIVO DE PDF (USANDO REPORTLAB - 100% GARANTIDO)
# =========================================================

def gerar_ordem_servico_pdf(fields):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=0.75*inch,
        leftMargin=0.75*inch,
        topMargin=0.75*inch,
        bottomMargin=0.75*inch
    )
    
    styles = getSampleStyleSheet()
    
    style_hdr_left = ParagraphStyle(
        'HdrLeft',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=13,
        textColor=colors.black
    )
    
    style_hdr_right = ParagraphStyle(
        'HdrRight',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=13,
        alignment=2,
        textColor=colors.black
    )
    
    style_title = ParagraphStyle(
        'OpTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=16,
        alignment=1,
        textColor=colors.HexColor('#002060')
    )
    
    style_sec_title = ParagraphStyle(
        'SecTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=15,
        textColor=colors.black,
        spaceBefore=10,
        spaceAfter=4
    )
    
    style_body = ParagraphStyle(
        'BodyTextCustom',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        leading=14,
        textColor=colors.black,
        spaceAfter=5
    )
    
    style_ass = ParagraphStyle(
        'Assinatura',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        leading=14,
        alignment=1,
        textColor=colors.black
    )

    story = []
    
    # Cabeçalho Tabela
    left_p = Paragraph("PMPR<br/>2º CRPM/18º BPM<br/>P/3", style_hdr_left)
    right_p = Paragraph(f"Cornélio Procópio, PR.<br/>Em {fields.get('data_expedicao', '')}<br/><b>ORDEM DE SERVIÇO Nº {fields.get('num_os', '077')}</b>", style_hdr_right)
    
    tbl_hdr = Table([[left_p, right_p]], colWidths=[3.2*inch, 3.2*inch])
    tbl_hdr.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
    ]))
    
    story.append(tbl_hdr)
    story.append(Spacer(1, 4))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.black, spaceBefore=4, spaceAfter=12))
    
    # Título da Operação
    op_name = fields.get('nome_operacao', '').strip().upper()
    story.append(Paragraph(f"“{op_name}”", style_title))
    story.append(Spacer(1, 10))
    
    # Seções
    secoes = [
        ("1. FINALIDADE", fields.get('finalidade', '')),
        ("2. REFERÊNCIAS", fields.get('referencias', '')),
        ("3. SITUAÇÃO E OBJETIVOS", fields.get('situacao', '')),
        ("4. EXECUÇÃO E MISSÃO", fields.get('execucao', '')),
        ("5. ADMINISTRAÇÃO E LOGÍSTICA", fields.get('logistica', '')),
        ("6. RELATÓRIOS E SISGCOP", fields.get('relatorios', '')),
        ("7. PRESCRIÇÕES DIVERSAS", fields.get('prescricoes', ''))
    ]
    
    for tit, conteudo in secoes:
        if conteudo and conteudo.strip():
            story.append(Paragraph(tit, style_sec_title))
            linhas = conteudo.strip().split('\n')
            for linha in linhas:
                if linha.strip():
                    l_clean = linha.strip().replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                    story.append(Paragraph(l_clean, style_body))
    
    # Assinatura
    story.append(Spacer(1, 20))
    story.append(Paragraph("<i>(Assinado eletronicamente)</i>", style_ass))
    story.append(Paragraph(f"<b>{fields.get('nome_comandante', 'Ten.-Cel. QOEM PM Helder de Lima Dantas Junior')}</b>,", style_ass))
    story.append(Paragraph(f"<b>{fields.get('cargo_comandante', 'Comandante do 18º BPM.')}</b>", style_ass))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# =========================================================
# INTERFACE PRINCIPAL STREAMLIT
# =========================================================

if caminho_brasao:
    col1, col2, col3 = st.columns(3)
    with col2:
        st.image(caminho_brasao, width=150)

st.title("📑 Gerador de Ordem de Serviço (OS)")
st.caption("18º Batalhão de Polícia Militar — PMPR")
st.write("Envie a **Ordem de Operação (OO)** para extração automática e geração da **Ordem de Serviço (OS)**.")

arquivo_oo = st.file_uploader("Envie o arquivo da Ordem de Operação (PDF ou DOCX)", type=["pdf", "docx", "doc"])

if arquivo_oo:
    texto_extraido = extrair_texto_arquivo(arquivo_oo)
    
    if texto_extraido.strip():
        st.success("Ordem de Operação analisada com sucesso!")
        
        parsed_data = parsear_ordem_operacao(texto_extraido)
        
        st.subheader("📝 Informações da Ordem de Serviço")
        st.write("Ajuste os campos essenciais abaixo antes de gerar os documentos:")

        # TELA ENXUTA: APENAS OS CAMPOS ESSENCIAIS
        col_a, col_b = st.columns(2)
        with col_a:
            num_os = st.text_input("Número da Ordem de Serviço (OS)", value="077")
            num_oo = st.text_input("Ordem de Operação de Origem", value=parsed_data.get('num_oo', '000/2026'))
        with col_b:
            data_expedicao = st.text_input("Data de Expedição da OS", value=data_atual_extenso())
            nome_operacao = st.text_input("Nome da Operação", value=parsed_data.get('nome_op', 'OPERAÇÃO POLICIAL'))

        col_c, col_d = st.columns(2)
        with col_c:
            nome_comandante = st.text_input("Comandante / Assinatura", value="Ten.-Cel. QOEM PM Helder de Lima Dantas Junior")
        with col_d:
            cargo_comandante = st.text_input("Cargo / Função", value="Comandante do 18º BPM.")

        st.session_state['parsed_full_data'] = parsed_data

        if st.button("🚀 Gerar Ordem de Serviço (Word e PDF)"):
            fields_final = {
                'num_os': num_os,
                'num_oo': num_oo,
                'data_expedicao': data_expedicao,
                'nome_operacao': nome_operacao,
                'finalidade': parsed_data.get('finalidade', ''),
                'referencias': parsed_data.get('referencias', ''),
                'situacao': parsed_data.get('situacao', ''),
                'execucao': parsed_data.get('execucao', ''),
                'logistica': parsed_data.get('logistica', ''),
                'relatorios': parsed_data.get('relatorios', ''),
                'prescricoes': parsed_data.get('prescricoes', ''),
                'nome_comandante': nome_comandante,
                'cargo_comandante': cargo_comandante
            }
            
            # 1. Gera DOCX em memória (Sem Brasão no relatório Word)
            docx_bytes = gerar_ordem_servico_docx(fields_final)
            st.session_state['generated_docx'] = docx_bytes
            
            # 2. Gera PDF nativo em memória usando ReportLab (100% confiável no Streamlit Cloud)
            pdf_bytes = gerar_ordem_servico_pdf(fields_final)
            st.session_state['generated_pdf'] = pdf_bytes
            
            st.session_state['filename_base'] = f"OS_{num_os.replace('/', '_')}_{nome_operacao.replace(' ', '_')}"

        # Exibe os botões de download
        if 'generated_docx' in st.session_state:
            st.markdown("---")
            st.subheader("📥 Baixar Arquivo Gerado")
            
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.download_button(
                    label="📄 Baixar em Word (.docx)",
                    data=st.session_state['generated_docx'],
                    file_name=f"{st.session_state['filename_base']}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            
            with col_d2:
                st.download_button(
                    label="📕 Baixar em PDF (.pdf)",
                    data=st.session_state['generated_pdf'],
                    file_name=f"{st.session_state['filename_base']}.pdf",
                    mime="application/pdf"
                )
    else:
        st.error("Não foi possível extrair texto do arquivo enviado. Verifique se o PDF ou DOCX contém texto pesquisável.")
