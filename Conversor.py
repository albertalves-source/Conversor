import streamlit as st
import pandas as pd
import io
import warnings
import re
import html

# Tenta importar as bibliotecas específicas para PDF
try:
    import pdfplumber
    import xlsxwriter
    from reportlab.lib.pagesizes import landscape, A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    # Bibliotecas para OCR
    import pytesseract
    import fitz  # PyMuPDF (Substituiu o pdf2image para não precisar do Poppler)
    from PIL import Image
except ImportError:
    st.error("⚠️ Faltam bibliotecas! Certifique-se de ter no seu requirements.txt: pandas, streamlit, pdfplumber, reportlab, openpyxl, xlsxwriter, pytesseract, PyMuPDF, Pillow")
    st.stop()

# Configurações de Página
st.set_page_config(page_title="Conversor Universal | PDF ↔ Excel", layout="centered", page_icon="🔄")
warnings.filterwarnings("ignore")

# --- FUNÇÕES AUXILIARES ---

def parse_pages(page_str, max_pages):
    """Converte uma string '1, 3, 5-7' numa lista de índices de páginas (0-based)."""
    if not page_str.strip():
        return list(range(max_pages))
    
    pages = set()
    for part in page_str.split(','):
        part = part.strip()
        if '-' in part:
            start, end = part.split('-')
            if start.isdigit() and end.isdigit():
                # Streamlit usa 1-based para o utilizador, pdfplumber usa 0-based
                pages.update(range(int(start) - 1, int(end)))
        elif part.isdigit():
            pages.add(int(part) - 1)
    
    # Filtra páginas inválidas
    return sorted([p for p in pages if 0 <= p < max_pages])

# --- FUNÇÕES DE CONVERSÃO ---

def pdf_para_dataframe(file, modo, paginas_str="", **kwargs):
    """Lê um PDF e transforma num DataFrame do Pandas."""
    
    # --- MODO 4: OCR (Reconhecimento Ótico de Caracteres) ---
    if modo == "Imagem/Scan (OCR)":
        try:
            file.seek(0)
            # Usa PyMuPDF (fitz) para abrir o PDF sem precisar do Poppler
            doc = fitz.open(stream=file.read(), filetype="pdf")
            max_pages = len(doc)
            indices_paginas = parse_pages(paginas_str, max_pages)
            
            all_rows = []
            for idx in indices_paginas:
                if idx < max_pages:
                    # Carrega a página e converte para imagem (pixmap) com 200 DPI para boa qualidade
                    page = doc.load_page(idx)
                    pix = page.get_pixmap(dpi=200)
                    
                    # Converte o Pixmap para Imagem PIL
                    mode = "RGBA" if pix.alpha else "RGB"
                    img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
                    
                    # Tenta extrair texto da imagem (assumindo português)
                    text = pytesseract.image_to_string(img, lang='por')
                    for l in text.split('\n'):
                        if l.strip():
                            all_rows.append(re.split(r'\s{2,}', l.strip()))
            
            if not all_rows:
                return pd.DataFrame()
                
            max_cols = max((len(row) for row in all_rows if row), default=0)
            normalized_rows = [row + [""] * (max_cols - len(row)) for row in all_rows if row]
            return pd.DataFrame(normalized_rows)
            
        except Exception as e:
            erro_str = str(e).lower()
            if "tesseract" in erro_str:
                st.error("⚠️ O Tesseract-OCR não foi encontrado! Ele é um software que precisa ser instalado **apenas no Servidor** (ou PC principal) onde a aplicação está alojada.")
            else:
                st.error(f"Erro no OCR. Detalhes: {e}")
            return pd.DataFrame()

    # --- MODOS BASEADOS EM TEXTO NATIVO (PDFPLUMBER) ---
    with pdfplumber.open(file) as pdf:
        max_pages = len(pdf.pages)
        indices_paginas = parse_pages(paginas_str, max_pages)
        
        if not indices_paginas:
            raise ValueError("As páginas selecionadas são inválidas ou excedem o limite do PDF.")

        all_rows = []
        
        # MODO 1: Tabelas com Bordas
        if modo == "Tabelas com Bordas (Padrão)":
            for idx in indices_paginas:
                page = pdf.pages[idx]
                tables = page.extract_tables()
                if tables:
                    for table in tables:
                        for row in table:
                            cleaned_row = [str(cell).replace('\n', ' ').replace('\xa0', ' ').strip() if cell else "" for cell in row]
                            all_rows.append(cleaned_row)
                            
        # MODO 2: Extratos Bancários (Inteligência Contábil)
        elif modo == "Inteligente (Agrupar por Data) - Ideal para Extratos Bancários":
            linhas_processadas = []
            linha_atual = {}
            
            def is_id(token):
                if len(token) >= 6 and re.search(r'\d', token) and (re.search(r'[a-zA-Z]', token) or '-' in token):
                    # Atualizado para aceitar datas sem o ano (DD/MM) e ignorar como ID
                    if not re.search(r'\d{2}[/-]\d{2}(?:[/-]\d{2,4})?', token):
                        return True
                if len(token) > 10 and token.isdigit():
                    return True
                return False

            regex_valor = r'(?:R\$\s*)?-?\d{1,3}(?:\.\d{3})*,\d{2}\b'

            for idx in indices_paginas:
                page = pdf.pages[idx]
                text = page.extract_text()
                if not text: continue
                
                for line in text.split('\n'):
                    line = line.strip()
                    if not line: continue
                    
                    if any(header in line.upper() for header in ["SALDO ANTERIOR", "SALDO FINAL", "EXTRATO", "DATA MOVIMENTO", "PERÍODO", "NOME:", "CONTA:", "DÉBITO", "CRÉDITO"]):
                        continue
                    
                    # Atualizado para aceitar datas com ou sem o ano (ex: 15/08 ou 15/08/2023)
                    match_data = re.search(r'^\s*(\d{2}[/-]\d{2}(?:[/-]\d{2,4})?)\s+(.*)', line)
                    
                    if match_data:
                        if linha_atual:
                            linhas_processadas.append(linha_atual)
                            
                        data, resto = match_data.group(1), match_data.group(2)
                        valores = re.findall(regex_valor, resto)
                        texto_sem_valor = resto
                        for v in valores: texto_sem_valor = texto_sem_valor.replace(v, '').strip()
                            
                        tokens = texto_sem_valor.split()
                        id_transacao, historico = "", texto_sem_valor
                        
                        for token in reversed(tokens[-3:]):
                            if is_id(token):
                                id_transacao = token + id_transacao
                                historico = historico.replace(token, '').strip()
                        
                        linha_atual = {
                            "Data": data, "Histórico": historico.strip(),
                            "ID / Documento": id_transacao.strip(), "Valores_Lista": valores
                        }
                    else:
                        if linha_atual:
                            valores_extra = re.findall(regex_valor, line)
                            if valores_extra:
                                linha_atual["Valores_Lista"].extend(valores_extra)
                                
                            texto_extra = line
                            for v in valores_extra: texto_extra = texto_extra.replace(v, '').strip()
                                
                            tokens_extra = texto_extra.split()
                            id_extra, hist_extra = "", texto_extra
                            
                            for token in reversed(tokens_extra[-3:]):
                                if is_id(token):
                                    id_extra = token + id_extra
                                    hist_extra = hist_extra.replace(token, '').strip()
                                    
                            if hist_extra: linha_atual["Histórico"] += " " + hist_extra.strip()
                            if id_extra: linha_atual["ID / Documento"] += id_extra.strip()
            
            if linha_atual: linhas_processadas.append(linha_atual)
                
            for row in linhas_processadas:
                hist_up = row["Histórico"].upper()
                is_cred = False
                
                if any(x in hist_up for x in ["RECEBID", "DEVOLU", "ESTORNO", "RESSARCIMENTO", "CREDIT", "CRÉDIT", "DEPÓSIT", "DEPOSIT", "PIX RECEBIDO"]):
                    is_cred = True
                elif any(x in hist_up for x in ["ENVIAD", "PAGAMENTO", "PAGTO", "SAQUE", "COMPRA", "DEBITO", "DÉBITO", "TARIFA", "TAXA"]):
                    is_cred = False
                
                vals = row["Valores_Lista"]
                
                if is_cred:
                    row["Débito"] = ""
                    row["Crédito"] = vals[0] if len(vals) > 0 else ""
                else:
                    row["Débito"] = vals[0] if len(vals) > 0 else ""
                    row["Crédito"] = ""
                
                row["Saldo"] = vals[-1] if len(vals) > 1 else ""
                del row["Valores_Lista"]

            return pd.DataFrame(linhas_processadas)
            
        # MODO 3: Regras Personalizadas (Agrupar por Palavra-Chave)
        elif modo == "Personalizado (Palavras-Chave)":
            palavra_chave = kwargs.get('palavra_chave', '').strip().upper()
            if not palavra_chave:
                st.warning("⚠️ Precisa de definir uma palavra-chave no campo de texto para este modo funcionar.")
                return pd.DataFrame()
                
            linhas_processadas = []
            linha_atual = {}
            
            for idx in indices_paginas:
                page = pdf.pages[idx]
                text = page.extract_text()
                if not text: continue
                
                for line in text.split('\n'):
                    line = line.strip()
                    if not line: continue
                    
                    # Se encontrou a palavra-chave, inicia um novo registo
                    if palavra_chave in line.upper():
                        if linha_atual:
                            linhas_processadas.append(linha_atual)
                        # Separa o conteúdo associado à palavra-chave (se houver)
                        partes = re.split(re.escape(palavra_chave), line, flags=re.IGNORECASE)
                        valor_chave = partes[1].strip() if len(partes) > 1 else ""
                        linha_atual = {"Identificador": linha_atual.get("Identificador", f"Registo {len(linhas_processadas)+1}"),
                                       "Palavra-Chave": valor_chave, 
                                       "Detalhes": ""}
                    elif linha_atual:
                        # Tudo o que vem a seguir até à próxima palavra-chave é detalhe
                        linha_atual["Detalhes"] += f" {line}"
            
            if linha_atual:
                linhas_processadas.append(linha_atual)
                
            return pd.DataFrame(linhas_processadas)

        # MODO 5: Texto Quebrado
        elif modo == "Texto Bruto (Separar colunas por espaço)":
            for idx in indices_paginas:
                page = pdf.pages[idx]
                text = page.extract_text()
                if not text: continue
                for l in text.split('\n'):
                    if l.strip():
                        all_rows.append(re.split(r'\s{2,}', l.strip()))
        
        if not all_rows and modo != "Personalizado (Palavras-Chave)":
            return pd.DataFrame()
            
        if all_rows:
            max_cols = max((len(row) for row in all_rows if row), default=0)
            normalized_rows = [row + [""] * (max_cols - len(row)) for row in all_rows if row]
            
            if modo == "Tabelas com Bordas (Padrão)" and len(normalized_rows) > 1:
                df = pd.DataFrame(normalized_rows[1:], columns=normalized_rows[0])
            else:
                df = pd.DataFrame(normalized_rows)
            return df

def dataframe_para_pdf(df, titulo="Relatório Convertido"):
    """Transforma um DataFrame do Pandas num ficheiro PDF formatado."""
    buffer = io.BytesIO()
    
    if df.empty:
        doc = SimpleDocTemplate(buffer, pagesize=landscape(A4))
        styles = getSampleStyleSheet()
        doc.build([Paragraph("O relatório não contém dados.", styles['Title'])])
        return buffer.getvalue()

    largura_util = 782 
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=18)
    elements = []
    
    styles = getSampleStyleSheet()
    titulo_formatado = Paragraph(f"<b>{html.escape(titulo)}</b>", styles['Title'])
    elements.append(titulo_formatado)
    elements.append(Spacer(1, 20))
    
    num_cols = len(df.columns)
    col_widths = []
    
    if num_cols > 0:
        max_lens = []
        for col in df.columns:
            col_data = df[col].dropna().astype(str)
            max_len = max(col_data.apply(len).max(), len(str(col))) if not col_data.empty else len(str(col))
            max_lens.append(min(max(max_len, 5), 80))
        
        total_len = sum(max_lens)
        col_widths = [(l / total_len) * largura_util for l in max_lens]
    
    style_header = ParagraphStyle('HeaderStyle', parent=styles['Normal'], fontSize=9, textColor=colors.whitesmoke, alignment=1, fontName='Helvetica-Bold')
    style_cell = ParagraphStyle('CellStyle', parent=styles['Normal'], fontSize=8, textColor=colors.HexColor("#2C3E50"), alignment=1)
    
    cabecalho = [Paragraph(html.escape(str(c)), style_header) for c in df.columns]
    dados_formatados = [cabecalho]
    
    for _, row in df.iterrows():
        linha_formatada = [Paragraph(html.escape(str(val)) if pd.notna(val) else "", style_cell) for val in row]
        dados_formatados.append(linha_formatada)
    
    t = Table(dados_formatados, colWidths=col_widths, repeatRows=1)
    
    estilo = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2C3E50")), 
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('GRID', (0,0), (-1,-1), 1, colors.HexColor("#BDC3C7")), 
    ])
    
    if len(dados_formatados) > 1:
        estilo.add('BACKGROUND', (0,1), (-1,-1), colors.HexColor("#ECF0F1"))
        
    t.setStyle(estilo)
    elements.append(t)
    doc.build(elements)
    
    return buffer.getvalue()


# --- INTERFACE DO UTILIZADOR ---

st.title("🔄 Conversor Universal")
st.markdown("Converta ficheiros **PDF para Excel** e **Excel para PDF** em segundos.")

# Inicializar variáveis de estado
if 'df_extraido' not in st.session_state:
    st.session_state.df_extraido = None

aba1, aba2 = st.tabs(["📄 PDF ➡️ Excel", "📊 Excel ➡️ PDF"])

# ==========================================
# ABA 1: PDF PARA EXCEL
# ==========================================
with aba1:
    st.subheader("Extrair tabelas de um ficheiro PDF")
    
    col1, col2 = st.columns([2, 1])
    with col1:
        modo_extracao = st.selectbox(
            "Método de Extração:",
            [
                "Tabelas com Bordas (Padrão)",
                "Inteligente (Agrupar por Data) - Ideal para Extratos Bancários",
                "Personalizado (Palavras-Chave)",
                "Imagem/Scan (OCR)",
                "Texto Bruto (Separar colunas por espaço)"
            ]
        )
    with col2:
        paginas_input = st.text_input("Páginas (ex: 1, 3, 5-7)", placeholder="Deixe vazio para todas")

    # Opções dinâmicas mediante o modo selecionado
    kwargs_extracao = {}
    if modo_extracao == "Personalizado (Palavras-Chave)":
        palavra = st.text_input("Palavra-chave para iniciar nova linha (Ex: 'Matrícula:', 'Nome:', 'Fatura:'):")
        kwargs_extracao['palavra_chave'] = palavra
        st.caption("Esta regra vai procurar a palavra-chave indicada e agrupar tudo o que vem a seguir até à próxima ocorrência numa única linha estruturada.")
    elif modo_extracao == "Imagem/Scan (OCR)":
        st.info("ℹ️ **Como funciona o OCR na rede do escritório:** Só precisa de instalar o Tesseract-OCR na máquina central (Servidor) que corre esta aplicação. Os outros funcionários acedem apenas pelo navegador (Chrome/Edge) e não precisam de instalar nada nas suas próprias máquinas!")

    limpar_dados = st.checkbox("Limpar colunas e linhas vazias automaticamente", value=True)
    
    pdf_file = st.file_uploader("Arraste o seu PDF aqui", type=["pdf"])
    
    if pdf_file:
        if st.button("🚀 Processar PDF", type="primary", use_container_width=True):
            with st.spinner("A analisar a estrutura do PDF..."):
                try:
                    df_pdf = pdf_para_dataframe(pdf_file, modo_extracao, paginas_input, **kwargs_extracao)
                    
                    if df_pdf is not None and not df_pdf.empty and limpar_dados:
                        # Substitui strings vazias por NaN e remove linhas/colunas 100% vazias
                        df_pdf.replace(r'^\s*$', pd.NA, regex=True, inplace=True)
                        df_pdf.dropna(how='all', inplace=True)
                        df_pdf.dropna(axis=1, how='all', inplace=True)
                        df_pdf.fillna("", inplace=True)
                        
                    st.session_state.df_extraido = df_pdf
                except Exception as e:
                    st.error(f"Erro ao processar o PDF: {e}")
                    st.session_state.df_extraido = None

        # Mostra os resultados se existirem no state
        if st.session_state.df_extraido is not None:
            df_mostrar = st.session_state.df_extraido
            
            if df_mostrar.empty:
                st.warning("Não foi possível encontrar dados com as definições atuais. Dica: Se o 'Modo Inteligente' falhou com o seu extrato, tente usar o 'Texto Bruto' para ver como o PDF está formatado originalmente.")
            else:
                st.success(f"PDF extraído com sucesso! Encontradas {len(df_mostrar)} linhas.")
                st.dataframe(df_mostrar, use_container_width=True)
                
                out_excel = io.BytesIO()
                with pd.ExcelWriter(out_excel, engine='xlsxwriter') as wr:
                    df_mostrar.to_excel(wr, index=False)
                
                nome_excel = pdf_file.name.replace('.pdf', '.xlsx')
                st.download_button(
                    label="📥 Baixar Excel",
                    data=out_excel.getvalue(),
                    file_name=nome_excel,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

# ==========================================
# ABA 2: EXCEL PARA PDF
# ==========================================
with aba2:
    st.subheader("Transformar grelha em Documento PDF")
    
    excel_file = st.file_uploader("Arraste o seu Excel/CSV aqui", type=["xlsx", "xls", "csv"])
    
    if excel_file:
        with st.spinner("A preparar o documento..."):
            try:
                if excel_file.name.endswith('.csv'):
                    df_excel = pd.read_csv(excel_file)
                else:
                    df_excel = pd.read_excel(excel_file)
                
                st.success("Ficheiro lido com sucesso! Pré-visualização:")
                st.dataframe(df_excel.head(10), use_container_width=True)
                
                nome_base = excel_file.name.rsplit('.', 1)[0]
                pdf_bytes = dataframe_para_pdf(df_excel, titulo=f"Relatório: {nome_base}")
                
                nome_pdf = f"{nome_base}.pdf"
                st.download_button(
                    label="📄 Baixar PDF Formatado",
                    data=pdf_bytes,
                    file_name=nome_pdf,
                    mime="application/pdf",
                    type="primary"
                )
            except Exception as e:
                st.error(f"Erro ao processar o ficheiro: {e}")
