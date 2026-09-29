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
import pdfplumber
import docx
from datetime import datetime

# Configuração da página e tema
st.set_page_config(page_title="18º BPM — Gerador de Ordem de Serviço", page_icon="📑", layout="centered")

# 🎨 MODO ESCURO (CSS Personalizado)
st.markdown("""
<style>
    /* Fundo Escuro */
    .stApp {
        background-color: #0E1117;
        color: #E0E6ED;
    }
    /* Estilização dos Containers e Caixas */
    div[data-testid="stFileUploader"], div[data-testid="stTextInput"], div[data-testid="stTextArea"] {
        background-color: #1E222D;
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
""", unsafe_allow_html=True)

# 🏷️ ASSINATURA MOVIDA PARA A ESQUERDA
st.markdown("""
<div style="position: fixed; bottom: 15px; right: 140px; text-align: right; color: #9CA3AF; font-size: 12px; font-family: sans-serif; z-index: 999999; line-height: 1.4; background-color: rgba(14, 17, 23, 0.9); padding: 6px 12px; border-radius: 6px; border: 1px solid #2E364A;">
    Desenvolvido por:<br>
    <strong style="color: #60A5FA; font-size: 13px;">Nathan Wenzel</strong>
</div>
""", unsafe_allow_html=True)

# 🔒 CONFIGURAÇÃO DA SENHA DE ACESSO
SENHA_CORRETA = "deusa"

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

# Busca flexível da imagem do brasão
def obter_caminho_brasao():
    for nome in ["brasao.png", "brasao.PNG", "Brasao.png", "BRASAO.PNG", "brasao.jpg", "brasao.jpeg"]:
        if os.path.exists(nome):
            return nome
    return None

caminho_brasao = obter_caminho_brasao()

# Tela de Login
if not st.session_state.autenticado:
    if caminho_brasao:
        col1, col2, col3 = st.columns(3)
        with col2:
            st.image(caminho_brasao, width=130)

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
# FUNÇÕES DE EXTRAÇÃO E PARSER DE ORDEM DE OPERAÇÃO (OO)
# =========================================================

MESES = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
    5: "maio", 6: "junho", 7: "julho", 8: "agosto",
    9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
}

def data_atual_extenso():
    now = datetime.now()
    return f"{now.day} de {MESES[now.month]} de {now.year}"

def extrair_texto_arquivo(uploaded_file):
    """Extrai todo o texto de um arquivo PDF ou DOCX enviado"""
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
        
    return text

def extrair_secao(texto, inicio_regex, fim_regex):
    """Auxiliar para extrair bloco de texto entre duas seções usando Regex (Corrigido)"""
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

    # Margens Oficiais
    for section in doc.sections:
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # Brasão
    caminho_img = obter_caminho_brasao()
    if caminho_img:
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_after = Pt(4)
        run_img = p_img.add_run()
        run_img.add_picture(caminho_img, width=Inches(0.85))

    # Tabela de Cabeçalho Institucional
    table_hdr = doc.add_table(rows=1, cols=2)
    table_hdr.autofit = False
    table_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    row = table_hdr.rows[0]
    row.cells[0].width = Inches(3.5)
    row.cells[1].width = Inches(3.0)

    # Célula Esquerda (Unidade)
    p_left = row.cells[0].paragraphs[0]
    p_left.paragraph_format.space_after = Pt(2)
    p_left.paragraph_format.line_spacing = 1.15
    r_l = p_left.add_run("PMPR\n2º CRPM/18º BPM\nP/3")
    r_l.bold = True
    r_l.font.name = "Arial"
    r_l.font.size = Pt(10)

    # Célula Direita (Local, Data, OS)
    p_right = row.cells[1].paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_right.paragraph_format.space_after = Pt(2)
    p_right.paragraph_format.line_spacing = 1.15
    r_r = p_right.add_run(f"Cornélio Procópio, PR.\nEm {fields.get('data_expedicao', data_atual_extenso())}\nORDEM DE SERVIÇO Nº {fields.get('num_os', '077')}")
    r_r.bold = True
    r_r.font.name = "Arial"
    r_r.font.size = Pt(10)

    # Linha Divisória
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_after = Pt(6)
    p_div.paragraph_format.space_before = Pt(4)
    r_div = p_div.add_run("___________________________________________________________________")
    r_div.bold = True
    r_div.font.size = Pt(9)

    # Título da Operação
    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_titulo.paragraph_format.space_after = Pt(12)
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
            p_sec.paragraph_format.space_before = Pt(8)
            p_sec.paragraph_format.space_after = Pt(3)
            r_sec = p_sec.add_run(tit)
            r_sec.bold = True
            r_sec.font.name = "Arial"
            r_sec.font.size = Pt(11)

            linhas = conteudo.strip().split('\n')
            for linha in linhas:
                if linha.strip():
                    p_cnt = doc.add_paragraph()
                    p_cnt.paragraph_format.space_after = Pt(4)
                    p_cnt.paragraph_format.line_spacing = 1.15
                    r_cnt = p_cnt.add_run(linha.strip())
                    r_cnt.font.name = "Arial"
                    r_cnt.font.size = Pt(10)

    # Assinatura do Comandante do 18º BPM
    p_ass = doc.add_paragraph()
    p_ass.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ass.paragraph_format.space_before = Pt(20)
    p_ass.paragraph_format.space_after = Pt(2)
    
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

    # Distribuição
    p_dist = doc.add_paragraph()
    p_dist.paragraph_format.space_before = Pt(14)
    r_dist_lbl = p_dist.add_run("DISTRIBUIÇÃO: ")
    r_dist_lbl.bold = True
    r_dist_lbl.font.size = Pt(9)
    r_dist_cnt = p_dist.add_run(fields.get('distribuicao', 'Cmdo. 2º CRPM; Cmdo. e Subcmdo. 18º BPM; P/1; P/2; P/3; P/4; P/5; PCS; ROTAM; 1ª, 2ª e 3ª Cias.; PRC; K9; Adjunto COPOM e CPU.'))
    r_dist_cnt.font.size = Pt(9)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# =========================================================
# INTERFACE PRINCIPAL STREAMLIT
# =========================================================

if caminho_brasao:
    col1, col2, col3 = st.columns(3)
    with col2:
        st.image(caminho_brasao, width=120)

st.title("📑 Gerador de Ordem de Serviço (OS)")
st.caption("18º Batalhão de Polícia Militar — PMPR")
st.write("Envie o arquivo de **Ordem de Operação (OO)** em **PDF** ou **DOCX** para gerar automaticamente a **Ordem de Serviço (OS)** correspondente.")

arquivo_oo = st.file_uploader("Envie o arquivo da Ordem de Operação (PDF ou DOCX)", type=["pdf", "docx", "doc"])

if arquivo_oo:
    texto_extraido = extrair_texto_arquivo(arquivo_oo)
    
    if texto_extraido.strip():
        st.success("Ordem de Operação carregada e analisada com sucesso!")
        
        parsed_data = parsear_ordem_operacao(texto_extraido)
        
        st.subheader("📝 Edição e Validação das Informações da OS")
        st.write("Verifique ou ajuste abaixo os campos extraídos antes de gerar a Ordem de Serviço final em Word:")

        col_a, col_b = st.columns(2)
        with col_a:
            num_os = st.text_input("Número da Ordem de Serviço (OS)", value="077")
            num_oo = st.text_input("Ordem de Operação de Origem", value=parsed_data.get('num_oo', '000/2026'))
        with col_b:
            data_expedicao = st.text_input("Data de Expedição da OS", value=data_atual_extenso())
            nome_operacao = st.text_input("Nome da Operação", value=parsed_data.get('nome_op', 'OPERAÇÃO POLICIAL'))

        finalidade = st.text_area("1. FINALIDADE", value=parsed_data.get('finalidade', ''), height=100)
        referencias = st.text_area("2. REFERÊNCIAS", value=parsed_data.get('referencias', ''), height=110)
        situacao = st.text_area("3. SITUAÇÃO E OBJETIVOS", value=parsed_data.get('situacao', ''), height=120)
        execucao = st.text_area("4. EXECUÇÃO E MISSÃO", value=parsed_data.get('execucao', ''), height=200)
        logistica = st.text_area("5. ADMINISTRAÇÃO E LOGÍSTICA", value=parsed_data.get('logistica', ''), height=130)
        relatorios = st.text_area("6. RELATÓRIOS E SISGCOP", value=parsed_data.get('relatorios', ''), height=120)
        prescricoes = st.text_area("7. PRESCRIÇÕES DIVERSAS", value=parsed_data.get('prescricoes', ''), height=120)

        col_c, col_d = st.columns(2)
        with col_c:
            nome_comandante = st.text_input("Comandante / Assinatura", value="Ten.-Cel. QOEM PM Helder de Lima Dantas Junior")
        with col_d:
            cargo_comandante = st.text_input("Cargo / Função", value="Comandante do 18º BPM.")

        distribuicao = st.text_input("Distribuição do Documento", value="Cmdo. 2º CRPM; Cmdo. e Subcmdo. 18º BPM; P/1; P/2; P/3; P/4; P/5; PCS; ROTAM; 1ª, 2ª e 3ª Cias.; PRC; K9; Adjunto COPOM e CPU.")

        if st.button("📄 Gerar Ordem de Serviço (.docx)"):
            fields_final = {
                'num_os': num_os,
                'num_oo': num_oo,
                'data_expedicao': data_expedicao,
                'nome_operacao': nome_operacao,
                'finalidade': finalidade,
                'referencias': referencias,
                'situacao': situacao,
                'execucao': execucao,
                'logistica': logistica,
                'relatorios': relatorios,
                'prescricoes': prescricoes,
                'nome_comandante': nome_comandante,
                'cargo_comandante': cargo_comandante,
                'distribuicao': distribuicao
            }
            
            docx_os_bytes = gerar_ordem_servico_docx(fields_final)
            
            st.download_button(
                label="📥 Baixar Ordem de Serviço Pronta (.docx)",
                data=docx_os_bytes,
                file_name=f"OS_{num_os.replace('/', '_')}_{nome_operacao.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.error("Não foi possível extrair texto do arquivo enviado. Verifique se o PDF ou DOCX contém texto pesquisável.")
