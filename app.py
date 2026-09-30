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

# ReportLab para geração garantida e nativa de PDF
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

# 🎨 MODO ESCURO COM BRASÃO EM MARCA D'ÁGUA EM TELA CHEIA
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
SENHA_CORRETA = "deusa"

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

def set_cell_bg(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def limpar_assinaturas_e_ruidos(texto):
    """Remove assinaturas eletrônicas, hash, e-Protocolo, números de páginas e cabeçalhos duplicados"""
    if not texto:
        return ""
    
    padroes_lixo = [
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
        r'Hash\s*:.*',
        r'Demais\s+assinaturas\s+na\s+folha.*',
        r'A\s+autenticidade\s+deste.*',
        r'\d{1,2}\.\d{3}\.\d{3}-\d.*por:.*',
        r'.*por:\s*\d*º?\s*(Sgt|Ten|Cel|Cap|Maj|Subten|Cb|Sd|Aux|PM).*em:\s*\d{2}/\d{2}/\d{4}.*',
        r'.*em:\s*\d{2}/\d{2}/\d{4}\s*\d{2}:\d{2}.*',
        r'Folha\s+\d+[a-z]?\s*de\s*\d+.*',
        r'Página\s+\d+\s+de\s+\d+.*',
        r'Aux\.\s*P/\d+.*em:\s*\d{2}/\d{2}/\d{4}.*',
        r'.*documento\s+pode\s+ser\s+verificad[oa].*',
        r'.*site\s+do\s+eprotocolo.*',
        r'^\s*\d{1,3}[a-z]?\s*$',
        r'^\s*-\s*\d{1,3}\s*-\s*$',
        r'^\s*\d{2}:\d{2}\b.*',
    ]
    
    linhas = texto.split('\n')
    linhas_limpas = []
    
    for linha in linhas:
        l_trim = linha.strip()
        if not l_trim:
            continue
            
        descartar = False
        for p in padroes_lixo:
            if re.search(p, l_trim, re.IGNORECASE):
                descartar = True
                break
                
        if not descartar:
            if '|' not in l_trim:
                l_trim = re.sub(r'\s+\d{1,2}$', '', l_trim)
            linhas_limpas.append(l_trim)
            
    res = '\n'.join(linhas_limpas)
    res = re.sub(r'\n{3,}', '\n\n', res)
    return res.strip()

def formatar_quebras_de_secao(texto):
    if not texto:
        return ""
    linhas = texto.split('\n')
    novas_linhas = []
    for l in linhas:
        if '|' in l:
            novas_linhas.append(l)
        else:
            l = re.sub(r'([^\n])\s*(\d+[ªº]\s*FASE|FASE\s+\d+|FASE\s+[I|V|X]+)', r'\1\n\2', l, flags=re.IGNORECASE)
            l = re.sub(r'([^\n])\s+([a-z0-9]{1,3}[\.\)])\s+', r'\1\n\2 ', l, flags=re.IGNORECASE)
            l = re.sub(r'([^\n])\s+(\d+\.\d+(?:\.\d+)?)\s+', r'\1\n\2 ', l)
            novas_linhas.append(l)
    return '\n'.join(novas_linhas)

def safe_crop_text(page, y0, y1):
    """Realiza o recorte seguro de área no PDF prevenindo erros de coordenadas"""
    if y1 <= y0 + 1:
        return ""
    y0 = max(0, min(y0, page.height - 1))
    y1 = max(y0 + 1, min(y1, page.height))
    if y1 <= y0:
        return ""
    try:
        cropped = page.crop((0, y0, page.width, y1))
        return cropped.extract_text() or ""
    except Exception:
        return ""

def extrair_texto_arquivo(uploaded_file):
    """
    Extrai texto e tabelas de PDF ou DOCX em ordem estrita de leitura
    SEM duplicar o conteúdo das tabelas como texto puro.
    """
    ext = uploaded_file.name.split(".")[-1].lower()
    text = ""
    
    if ext == "pdf":
        uploaded_file.seek(0)
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                tables = page.find_tables()
                if not tables:
                    page_text = page.extract_text() or ""
                    if page_text.strip():
                        text += page_text + "\n\n"
                else:
                    table_bboxes = [t.bbox for t in tables]
                    
                    def not_in_table(obj):
                        obj_x0 = obj.get('x0', 0)
                        obj_top = obj.get('top', 0)
                        obj_x1 = obj.get('x1', 0)
                        obj_bottom = obj.get('bottom', 0)
                        for (tx0, ttop, tx1, tbottom) in table_bboxes:
                            if not (obj_x1 <= tx0 or obj_x0 >= tx1 or obj_bottom <= ttop or obj_top >= tbottom):
                                return False
                        return True

                    filtered_page = page.filter(not_in_table)
                    tables_sorted = sorted(tables, key=lambda t: t.bbox)
                    
                    page_content = []
                    last_top = 0
                    page_height = page.height
                    
                    for t in tables_sorted:
                        tx0, ttop, tx1, tbottom = t.bbox
                        if ttop > last_top + 2:
                            t_above = safe_crop_text(filtered_page, last_top, ttop)
                            if t_above and t_above.strip():
                                page_content.append(t_above.strip())
                        
                        extracted_tbl = t.extract()
                        if extracted_tbl:
                            table_lines = []
                            for row in extracted_tbl:
                                if row and any(c is not None and str(c).strip() != "" for c in row):
                                    clean_row = [str(c).replace('\n', ' ').strip() if c is not None else "" for c in row]
                                    table_lines.append(" | ".join(clean_row))
                            if table_lines:
                                page_content.append("\n".join(table_lines))
                        
                        last_top = max(last_top, tbottom)
                    
                    if last_top < page_height - 2:
                        t_below = safe_crop_text(filtered_page, last_top, page_height)
                        if t_below and t_below.strip():
                            page_content.append(t_below.strip())
                            
                    text += "\n\n".join(page_content) + "\n\n"
        uploaded_file.seek(0)
    elif ext in ["docx", "doc"]:
        uploaded_file.seek(0)
        doc = docx.Document(uploaded_file)
        items = []
        for element in doc.element.body:
            if element.tag.endswith('p'):
                p = docx.text.paragraph.Paragraph(element, doc)
                if p.text.strip():
                    items.append(p.text.strip())
            elif element.tag.endswith('tbl'):
                tbl = docx.table.Table(element, doc)
                table_lines = []
                for row in tbl.rows:
                    row_cells = [c.text.replace('\n', ' ').strip() for c in row.cells]
                    if any(cell for cell in row_cells):
                        table_lines.append(" | ".join(row_cells))
                if table_lines:
                    items.append("\n".join(table_lines))
        text = "\n\n".join(items)
        uploaded_file.seek(0)
        
    text_limpo = limpar_assinaturas_e_ruidos(text)
    return text_limpo

def extrair_secao_flexivel(texto, padrao_inicio, padraos_fim):
    """
    Extrai o conteúdo de uma seção consumindo a linha inteira do título de início,
    buscando flexivelmente até a ocorrência de qualquer um dos padrões em padraos_fim.
    """
    regex_inicio = rf'^[ \t]*\d*\.?\s*{padrao_inicio}[^\n]*\n'
    match_inicio = re.search(regex_inicio, texto, re.IGNORECASE | re.MULTILINE)
    
    if not match_inicio:
        regex_fb = rf'{padrao_inicio}'
        match_fb = re.search(regex_fb, texto, re.IGNORECASE)
        if not match_fb:
            return ""
        pos_inicio = match_fb.end()
        nl = texto.find('\n', pos_inicio)
        pos_conteudo = nl + 1 if nl != -1 else pos_inicio
    else:
        pos_conteudo = match_inicio.end()
        
    conteudo_restante = texto[pos_conteudo:]
    
    regex_fim = rf'^[ \t]*\d*\.?\s*(?:{"|".join(padraos_fim)})[^\n]*'
    match_fim = re.search(regex_fim, conteudo_restante, re.IGNORECASE | re.MULTILINE)
    
    if match_fim:
        conteudo_secao = conteudo_restante[:match_fim.start()]
    else:
        conteudo_secao = conteudo_restante
        
    return conteudo_secao.strip()

def normalizar_subnumeracao_secao(conteudo, sec_num):
    """
    Normaliza rigorosamente a sub-numeração dentro de uma seção para que reflita o número correto da seção sec_num.
    Remove repetições de cabeçalhos de seção, números isolados e re-sequencia sub-itens desalinhados.
    """
    if not conteudo:
        return ""

    linhas = conteudo.strip().split('\n')
    linhas_limpas = []

    padrao_titulo_principal = re.compile(
        r'^\s*(\d*\.?\s*)?(E\s+LOGÍSTICA|FINALIDADE|INFORMAÇÕES\s+GERAIS|SITUAÇÃO|MISSÃO|EXECUÇÃO|ADMINISTRAÇÃO|LOGÍSTICA|RELATÓRIOS|PRESCRIÇÕES\s+DIVERSAS|REFERÊNCIAS)\b.*$',
        re.IGNORECASE
    )

    padrao_numero_isolado = re.compile(r'^\s*\d+\s*[\.\)]?\s*$')

    sub_counter = 1
    seen_subnums = set()

    for linha in linhas:
        l_str = linha.strip()
        if not l_str:
            continue

        if padrao_titulo_principal.match(l_str) and len(l_str) < 70 and not '|' in l_str:
            continue

        if padrao_numero_isolado.match(l_str):
            continue

        match_sub = re.match(r'^(\d+)\.(\d+)(\.\d+)?\s*(.*)', l_str)
        if match_sub:
            prefix_main, prefix_sub, prefix_subsub, rest = match_sub.groups()
            prefix_subsub = prefix_subsub if prefix_subsub else ""

            curr_sub_num = int(prefix_sub)
            sub_key = f"{sec_num}.{curr_sub_num}{prefix_subsub}"

            if int(prefix_main) != sec_num or sub_key in seen_subnums:
                new_sub_str = f"{sec_num}.{sub_counter}{prefix_subsub} {rest}"
                seen_subnums.add(f"{sec_num}.{sub_counter}{prefix_subsub}")
                sub_counter += 1
            else:
                new_sub_str = f"{sec_num}.{curr_sub_num}{prefix_subsub} {rest}"
                seen_subnums.add(sub_key)
                sub_counter = max(sub_counter, curr_sub_num + 1)

            linhas_limpas.append(new_sub_str)
        else:
            linhas_limpas.append(l_str)

    return "\n".join(linhas_limpas).strip()

def parsear_ordem_operacao(texto):
    """Analisa a Ordem de Operação e extrai os 8 campos na sequência exata solicitada"""
    dados = {}
    
    # Número da OO
    match_num = re.search(r'ORDEM DE OPERAÇÃO\s*(?:Nº|N°|Nº\.|N°\.|N°\s*|Nº\s*)?(\d+/\d+)', texto, re.IGNORECASE)
    dados['num_oo'] = match_num.group(1) if match_num else "000/2026"

    # Nome da Operação
    match_nome = re.search(r'“([^”]+)”|"([^"]+)"', texto)
    if match_nome:
        dados['nome_op'] = match_nome.group(1) or match_nome.group(2)
    else:
        match_nome2 = re.search(r'OPERAÇÃO\s+([A-Z0-9\s–\-]{4,})', texto)
        dados['nome_op'] = match_nome2.group(0).strip() if match_nome2 else "OPERAÇÃO POLICIAL"

    dados['nome_op'] = dados['nome_op'].replace('\n', ' ').strip().upper()

    # 1. FINALIDADE
    f = extrair_secao_flexivel(texto, r'FINALIDADE', [r'INFORMAÇÕES\s+GERAIS', r'SITUAÇÃO', r'MISSÃO', r'EXECUÇÃO'])
    f = normalizar_subnumeracao_secao(f, 1)
    dados['finalidade'] = f if f else "Realizar ações de policiamento ostensivo preventivo e preservação da ordem pública."

    # Extração conjunta do bloco entre Finalidade e Missão para separar Informações Gerais e Situação
    bloco_sit_completo = extrair_secao_flexivel(texto, r'SITUAÇÃO|INFORMAÇÕES\s+GERAIS', [r'MISSÃO', r'EXECUÇÃO'])

    linhas_bloco = [l.strip() for l in bloco_sit_completo.split('\n') if l.strip()]
    linhas_ig = []
    linhas_sit = []

    for l in linhas_bloco:
        if re.match(r'^\s*(\d*\.?\s*)?(SITUAÇÃO|INFORMAÇÕES\s+GERAIS)\s*$', l, re.IGNORECASE):
            continue
        if re.match(r'^\s*2\.1\s*INFORMAÇÕES\s+GERAIS\s*$', l, re.IGNORECASE):
            continue

        if re.match(r'^[a-z0-9][\.\)\-]\s+', l, re.IGNORECASE) or 'abrangência' in l.lower() or 'execução regionalizada' in l.lower():
            linhas_ig.append(l)
        elif l.startswith('A Polícia Militar do Paraná') or 'atuação contínua' in l.lower() or 'mandados de prisão' in l.lower():
            linhas_sit.append(l)
        else:
            if linhas_ig and not linhas_sit:
                linhas_ig.append(l)
            else:
                linhas_sit.append(l)

    # 2. INFORMAÇÕES GERAIS
    ig_str = "\n".join(linhas_ig).strip()
    if not ig_str:
        ig_str = f"A Operação “{dados['nome_op']}” possui abrangência na área do 18º BPM, concentrando esforços na preservação da ordem pública, cumprimento de mandados e fiscalização."
    else:
        ig_str = normalizar_subnumeracao_secao(ig_str, 2)
    dados['informacoes_gerais'] = ig_str

    # 3. SITUAÇÃO
    sit_str = "\n".join(linhas_sit).strip()
    if not sit_str:
        sit_str = "A Polícia Militar do Paraná, em sua atuação contínua e ostensiva, desenvolve operações voltadas à manutenção da ordem pública e segurança da comunidade."
    else:
        sit_str = normalizar_subnumeracao_secao(sit_str, 3)
    dados['situacao'] = sit_str

    # 4. MISSÃO (Garantida)
    m = extrair_secao_flexivel(texto, r'MISSÃO', [r'EXECUÇÃO', r'ADMINISTRAÇÃO', r'LOGÍSTICA'])
    if not m or len(m) < 10:
        m = f"O 18º BPM executará o policiamento ostensivo e a preservação da ordem pública na sua circunscrição territorial no âmbito da “{dados['nome_op']}”, visando a prevenção de crimes e a garantia da segurança pública."
    else:
        m = normalizar_subnumeracao_secao(m, 4)
    dados['missao'] = m

    # 5. EXECUÇÃO
    e = extrair_secao_flexivel(texto, r'EXECUÇÃO', [r'ADMINISTRAÇÃO', r'LOGÍSTICA', r'RELATÓRIOS'])
    e = normalizar_subnumeracao_secao(e, 5)
    dados['execucao'] = e if e else "Atuação integrada das equipes operacionais do 18º BPM em conformidade com o planejamento."

    # 6. ADMINISTRAÇÃO E LOGÍSTICA
    l = extrair_secao_flexivel(texto, r'ADMINISTRAÇÃO|LOGÍSTICA', [r'RELATÓRIOS', r'PRESCRIÇÕES'])
    l = normalizar_subnumeracao_secao(l, 6)
    dados['logistica'] = l if l else "Uniforme: Orgânico da OPM (4º RUPM).\nArmamento e equipamento: Orgânico compatível com o serviço.\nTransporte: Viaturas operacionais do 18º BPM."

    # 7. RELATÓRIOS E SISGCOP
    r = extrair_secao_flexivel(texto, r'RELATÓRIOS', [r'PRESCRIÇÕES', r'REFERÊNCIAS'])
    match_sisgcop = re.search(r'(\d{5,6})\s*[\-–]?\s*[\"“]?OPERAÇÃO', texto, re.IGNORECASE)
    if not match_sisgcop:
        match_sisgcop = re.search(r'SISGCOP[^\d]*(\d{5,6})', texto, re.IGNORECASE)
    num_sisgcop = match_sisgcop.group(1) if match_sisgcop else ""
    
    if r:
        r = normalizar_subnumeracao_secao(r, 7)
        dados['relatorios'] = r
    else:
        dados['relatorios'] = f"Os resultados obtidos deverão ser lançados no SISGCOP{' sob o código ' + num_sisgcop if num_sisgcop else ''} até o término da operação. Confecção dos Boletins de Ocorrência (BOU) no SADE."

    # 8. PRESCRIÇÕES DIVERSAS
    p = extrair_secao_flexivel(texto, r'PRESCRIÇÕES\s+DIVERSAS|PRESCRIÇÕES', [r'REFERÊNCIAS', r'DISTRIBUIÇÃO'])
    if p:
        p = normalizar_subnumeracao_secao(p, 8)
        dados['prescricoes'] = p
    else:
        dados['prescricoes'] = "Os policiais militares deverão atuar com bom senso, urbanidade, legalidade e estrito cumprimento do dever legal. Preleção obrigatória antes do início do serviço."

    # REFERÊNCIAS
    ref = extrair_secao_flexivel(texto, r'REFERÊNCIAS', [r'DISTRIBUIÇÃO', r'$'])
    if ref and len(ref.strip()) > 10:
        dados['referencias'] = ref
    else:
        dados['referencias'] = f"a) Constituição da República Federativa do Brasil de 1988;\nb) Constituição do Estado do Paraná de 1989;\nc) Lei n.º 22.354/2025 – Lei de Organização Básica da PMPR;\nd) Ordem de Operação nº {dados['num_oo']} – 2º CRPM ({dados['nome_op']});\ne) Determinação do Comandante do 18º BPM."

    return dados

def eh_titulo_subsecao(linha):
    """Verifica se uma linha é um título ou subtópico que deve ser formatado EM NEGRITO INTEIRO"""
    l = linha.strip()
    if not l:
        return False
    if l.endswith(':') and len(l) < 90:
        return True
    if re.match(r'^\d+(\.\d+)+[\.\)\-]?\s+', l):
        return True
    if re.match(r'^(FASE|ETAPA|GRUPO|ROTA|ZONA)\s+', l, re.IGNORECASE):
        return True
    if re.match(r'^[I|V|X]+[\.\)\-]\s+', l, re.IGNORECASE):
        return True
    if l.isupper() and len(l) < 90 and not '|' in l and not l.startswith('PMPR'):
        return True
    return False

# =========================================================
# GERADOR DO DOCUMENTO WORD (.DOCX)
# =========================================================

def renderizar_conteudo_docx(doc, conteudo):
    conteudo_formatado = formatar_quebras_de_secao(conteudo)
    if not conteudo_formatado or not conteudo_formatado.strip():
        return

    linhas = conteudo_formatado.strip().split('\n')
    i = 0
    while i < len(linhas):
        linha = linhas[i].strip()
        if not linha:
            i += 1
            continue
            
        # Tabela detectada por '|'
        if '|' in linha:
            tabela_linhas = []
            while i < len(linhas) and '|' in linhas[i]:
                cels = [c.strip() for c in linhas[i].split('|')]
                if len(cels) > 1 and cels[0] == "":
                    cels = cels[1:]
                if len(cels) > 1 and cels[-1] == "":
                    cels = cels[:-1]
                if any(c for c in cels):
                    tabela_linhas.append(cels)
                i += 1
                
            if tabela_linhas:
                max_cols = max(len(r) for r in tabela_linhas)
                for r_data in tabela_linhas:
                    while len(r_data) < max_cols:
                        r_data.append("")

                tbl = doc.add_table(rows=len(tabela_linhas), cols=max_cols)
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                tbl.style = 'Table Grid'
                
                for r_idx, row_data in enumerate(tabela_linhas):
                    row_cells = tbl.rows[r_idx].cells
                    for c_idx, cell_value in enumerate(row_data):
                        if c_idx < len(row_cells):
                            cell = row_cells[c_idx]
                            p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
                            p.paragraph_format.space_before = Pt(3)
                            p.paragraph_format.space_after = Pt(3)
                            p.paragraph_format.line_spacing = 1.15
                            
                            r = p.add_run(cell_value)
                            r.font.name = "Arial"
                            
                            if r_idx == 0:
                                set_cell_bg(cell, "002060")
                                r.bold = True
                                r.font.size = Pt(9.5)
                                r.font.color.rgb = RGBColor(255, 255, 255)
                            else:
                                if r_idx % 2 == 1:
                                    set_cell_bg(cell, "F8FAFC")
                                r.font.size = Pt(9.0)
                                r.font.color.rgb = RGBColor(30, 41, 59)
                p_sp = doc.add_paragraph()
                p_sp.paragraph_format.space_after = Pt(4)
            continue
            
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.2

        if eh_titulo_subsecao(linha):
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(4)
            r = p.add_run(linha)
            r.bold = True
            r.font.name = "Arial"
            r.font.size = Pt(10.5)
        else:
            m_item = re.match(r'^([a-z0-9]{1,3}[\.\)\-]|[A-Z][\.\)\-])\s+(.*)', linha, re.IGNORECASE)
            if m_item:
                prefix, rest = m_item.groups()
                p.paragraph_format.space_before = Pt(3)
                p.paragraph_format.space_after = Pt(4)
                
                r_pre = p.add_run(prefix + " ")
                r_pre.bold = True
                r_pre.font.name = "Arial"
                r_pre.font.size = Pt(10)
                
                r_rest = p.add_run(rest)
                r_rest.font.name = "Arial"
                r_rest.font.size = Pt(10)
            else:
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(6)
                r = p.add_run(linha)
                r.font.name = "Arial"
                r.font.size = Pt(10)
        
        i += 1

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
    
    table_hdr.columns[0].width = Inches(3.5)
    table_hdr.columns[1].width = Inches(3.0)
    
    row0 = table_hdr.rows[0]
    cell_left = row0.cells[0]
    cell_right = row0.cells[1]
    
    cell_left.width = Inches(3.5)
    cell_right.width = Inches(3.0)

    # Célula Esquerda (Unidade)
    p_left = cell_left.paragraphs[0]
    p_left.paragraph_format.space_after = Pt(2)
    p_left.paragraph_format.line_spacing = 1.2
    r_l = p_left.add_run("PMPR\n2º CRPM/18º BPM\nP/3")
    r_l.bold = True
    r_l.font.name = "Arial"
    r_l.font.size = Pt(10)

    # Célula Direita (Local, Data, OS)
    p_right = cell_right.paragraphs[0]
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

    # Seções Estruturadas na Sequência Exata Solicitada (1 a 8 + Referências)
    secoes = [
        ("1. FINALIDADE", fields.get('finalidade', '')),
        ("2. INFORMAÇÕES GERAIS", fields.get('informacoes_gerais', '')),
        ("3. SITUAÇÃO", fields.get('situacao', '')),
        ("4. MISSÃO", fields.get('missao', '')),
        ("5. EXECUÇÃO", fields.get('execucao', '')),
        ("6. ADMINISTRAÇÃO E LOGÍSTICA", fields.get('logistica', '')),
        ("7. RELATÓRIOS E SISGCOP", fields.get('relatorios', '')),
        ("8. PRESCRIÇÕES DIVERSAS", fields.get('prescricoes', '')),
        ("REFERÊNCIAS", fields.get('referencias', ''))
    ]

    for tit, conteudo in secoes:
        p_sec = doc.add_paragraph()
        p_sec.paragraph_format.space_before = Pt(14)
        p_sec.paragraph_format.space_after = Pt(4)
        r_sec = p_sec.add_run(tit)
        r_sec.bold = True
        r_sec.font.name = "Arial"
        r_sec.font.size = Pt(11)

        renderizar_conteudo_docx(doc, conteudo)

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
# GERADOR NATIVO DE PDF (USANDO REPORTLAB)
# =========================================================

def renderizar_conteudo_pdf(story, conteudo, style_subnum, style_body, style_table_hdr):
    conteudo_formatado = formatar_quebras_de_secao(conteudo)
    if not conteudo_formatado or not conteudo_formatado.strip():
        return

    linhas = conteudo_formatado.strip().split('\n')
    i = 0
    while i < len(linhas):
        linha = linhas[i].strip()
        if not linha:
            i += 1
            continue
            
        # Tabela no PDF
        if '|' in linha:
            tabela_linhas = []
            while i < len(linhas) and '|' in linhas[i]:
                cels = [c.strip() for c in linhas[i].split('|')]
                if len(cels) > 1 and cels[0] == "":
                    cels = cels[1:]
                if len(cels) > 1 and cels[-1] == "":
                    cels = cels[:-1]
                if any(c for c in cels):
                    tabela_linhas.append(cels)
                i += 1
                
            if tabela_linhas:
                max_cols = max(len(r) for r in tabela_linhas)
                for r_data in tabela_linhas:
                    while len(r_data) < max_cols:
                        r_data.append("")

                col_w = (6.4 * inch) / max_cols
                
                pdf_table_data = []
                for r_idx, r_data in enumerate(tabela_linhas):
                    row_p = []
                    for c_val in r_data:
                        c_clean = c_val.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                        if r_idx == 0:
                            p_cell = Paragraph(f"<b><font color='white'>{c_clean}</font></b>", style_table_hdr)
                        else:
                            p_cell = Paragraph(c_clean, style_body)
                        row_p.append(p_cell)
                    pdf_table_data.append(row_p)
                
                tbl = Table(pdf_table_data, colWidths=[col_w]*max_cols)
                tbl.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#002060')),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#A0AAB5')),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('TOPPADDING', (0,0), (-1,-1), 4),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                    ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')])
                ]))
                story.append(Spacer(1, 4))
                story.append(tbl)
                story.append(Spacer(1, 6))
            continue
            
        l_clean = linha.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

        if eh_titulo_subsecao(linha):
            story.append(Paragraph(f"<b>{l_clean}</b>", style_subnum))
        else:
            m_item = re.match(r'^([a-z0-9]{1,3}[\.\)\-]|[A-Z][\.\)\-])\s+(.*)', linha, re.IGNORECASE)
            if m_item:
                prefix, rest = m_item.groups()
                p_pre = prefix.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                p_rest = rest.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                story.append(Paragraph(f"<b>{p_pre}</b> {p_rest}", style_body))
            else:
                story.append(Paragraph(l_clean, style_body))
            
        i += 1

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
    
    style_hdr_left = ParagraphStyle('HdrLeft', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=13, textColor=colors.black)
    style_hdr_right = ParagraphStyle('HdrRight', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=13, alignment=2, textColor=colors.black)
    style_title = ParagraphStyle('OpTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=12, leading=16, alignment=1, textColor=colors.HexColor('#002060'))
    style_sec_title = ParagraphStyle('SecTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=11, leading=15, textColor=colors.black, spaceBefore=14, spaceAfter=4)
    style_subnum_title = ParagraphStyle('SubNumTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10.5, leading=14, textColor=colors.black, spaceBefore=10, spaceAfter=4)
    style_body = ParagraphStyle('BodyTextCustom', parent=styles['Normal'], fontName='Helvetica', fontSize=10, leading=14, textColor=colors.black, spaceAfter=5)
    style_table_hdr = ParagraphStyle('TableHdrCustom', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9.5, leading=13, textColor=colors.white)
    style_ass = ParagraphStyle('Assinatura', parent=styles['Normal'], fontName='Helvetica', fontSize=10, leading=14, alignment=1, textColor=colors.black)

    story = []
    
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
    
    op_name = fields.get('nome_operacao', '').strip().upper()
    story.append(Paragraph(f"“{op_name}”", style_title))
    story.append(Spacer(1, 10))
    
    secoes = [
        ("1. FINALIDADE", fields.get('finalidade', '')),
        ("2. INFORMAÇÕES GERAIS", fields.get('informacoes_gerais', '')),
        ("3. SITUAÇÃO", fields.get('situacao', '')),
        ("4. MISSÃO", fields.get('missao', '')),
        ("5. EXECUÇÃO", fields.get('execucao', '')),
        ("6. ADMINISTRAÇÃO E LOGÍSTICA", fields.get('logistica', '')),
        ("7. RELATÓRIOS E SISGCOP", fields.get('relatorios', '')),
        ("8. PRESCRIÇÕES DIVERSAS", fields.get('prescricoes', '')),
        ("REFERÊNCIAS", fields.get('referencias', ''))
    ]
    
    for tit, conteudo in secoes:
        story.append(Paragraph(tit, style_sec_title))
        renderizar_conteudo_pdf(story, conteudo, style_subnum_title, style_body, style_table_hdr)
    
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
                'informacoes_gerais': parsed_data.get('informacoes_gerais', ''),
                'situacao': parsed_data.get('situacao', ''),
                'missao': parsed_data.get('missao', ''),
                'execucao': parsed_data.get('execucao', ''),
                'logistica': parsed_data.get('logistica', ''),
                'relatorios': parsed_data.get('relatorios', ''),
                'prescricoes': parsed_data.get('prescricoes', ''),
                'referencias': parsed_data.get('referencias', ''),
                'nome_comandante': nome_comandante,
                'cargo_comandante': cargo_comandante
            }
            
            # 1. Gera DOCX em memória
            docx_bytes = gerar_ordem_servico_docx(fields_final)
            st.session_state['generated_docx'] = docx_bytes
            
            # 2. Gera PDF nativo em memória usando ReportLab
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
