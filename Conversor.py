import streamlit as st
import pandas as pd
import io
import warnings
import re
import html

# Tenta importar as bibliotecas específicas para PDF
try:
    import pdfplumber
    from reportlab.lib.pagesizes import landscape, A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
except ImportError:
    st.error("⚠️ Faltam bibliotecas! Certifique-se de ter no seu requirements.txt: pandas, streamlit, pdfplumber, reportlab, openpyxl")

# Configurações de Página
st.set_page_config(page_title="Conversor Universal | PDF ↔ Excel", layout="centered", page_icon="🔄")
warnings.filterwarnings("ignore")

# --- FUNÇÕES DE CONVERSÃO ---

def pdf_para_dataframe(file, modo):
    """Lê um PDF e transforma num DataFrame do Pandas de acordo com o modo selecionado."""
    with pdfplumber.open(file) as pdf:
        all_rows = []
        
        # MODO 1: Tabelas com linhas desenhadas
        if modo == "Tabelas com Bordas (Padrão)":
            for page in pdf.pages:
                tables = page.extract_tables()
                if tables:
                    for table in tables:
                        for row in table:
                            cleaned_row = [str(cell).replace('\n', ' ').strip() if cell else "" for cell in row]
                            all_rows.append(cleaned_row)
                            
        # MODO 2: Extratos Bancários (Inteligência Contábil de Débito/Crédito Correta)
        elif modo == "Inteligente (Agrupar por Data) - Ideal para Extratos Bancários":
            linhas_processadas = []
            linha_atual = {}
            
            def is_id(token):
                # Identifica códigos de transação/IDs (Alfanuméricos longos)
                if len(token) >= 6 and re.search(r'\d', token) and (re.search(r'[a-zA-Z]', token) or '-' in token):
                    if not re.search(r'\d{2}/\d{2}/\d{2,4}', token):
                        return True
                if len(token) > 10 and token.isdigit():
                    return True
                return False

            # Regex para detetar valores financeiros com ou sem R$
            regex_valor = r'(?:R\$\s*)?-?\d{1,3}(?:\.\d{3})*,\d{2}\b'

            for page in pdf.pages:
                text = page.extract_text()
                if not text: continue
                
                for line in text.split('\n'):
                    line = line.strip()
                    if not line: continue
                    
                    # Ignorar cabeçalhos repetitivos do PDF que poluem a tabela
                    if any(header in line.upper() for header in ["SALDO", "EXTRATO", "DATA MOVIMENTO", "PERÍODO", "NOME:", "CONTA:", "ID", "DÉBITO", "CRÉDITO"]):
                        continue
                    
                    # Se começar com uma Data, é o início de uma nova transação
                    match_data = re.match(r'^(\d{2}[/-]\d{2}[/-]\d{2,4})\s+(.*)', line)
                    if match_data:
                        if linha_atual:
                            linhas_processadas.append(linha_atual)
                            
                        data = match_data.group(1)
                        resto = match_data.group(2)
                        
                        # Extrai Valores Monetários
                        valores = re.findall(regex_valor, resto)
                        texto_sem_valor = resto
                        for v in valores:
                            texto_sem_valor = texto_sem_valor.replace(v, '').strip()
                            
                        # Tenta separar o ID longo do Histórico
                        tokens = texto_sem_valor.split()
                        id_transacao = ""
                        historico = texto_sem_valor
                        
                        for token in reversed(tokens[-3:]):
                            if is_id(token):
                                id_transacao = token + id_transacao
                                historico = historico.replace(token, '').strip()
                        
                        linha_atual = {
                            "Data": data,
                            "Histórico": historico.strip(),
                            "ID / Documento": id_transacao.strip(),
                            "Valores_Lista": valores
                        }
                    else:
                        # Se não tem data, é linha de continuação da transação anterior
                        if linha_atual:
                            valores_extra = re.findall(regex_valor, line)
                            if valores_extra:
                                linha_atual["Valores_Lista"].extend(valores_extra)
                                
                            texto_extra = line
                            for v in valores_extra:
                                texto_extra = texto_extra.replace(v, '').strip()
                                
                            tokens_extra = texto_extra.split()
                            id_extra = ""
                            hist_extra = texto_extra
                            
                            for token in reversed(tokens_extra[-3:]):
                                if is_id(token):
                                    id_extra = token + id_extra
                                    hist_extra = hist_extra.replace(token, '').strip()
                                    
                            if hist_extra:
                                linha_atual["Histórico"] += " " + hist_extra.strip()
                            if id_extra:
                                linha_atual["ID / Documento"] += id_extra.strip()
            
            if linha_atual: # Guarda a última linha
                linhas_processadas.append(linha_atual)
                
            # --- Tratamento Final: Débito vs Crédito (Sistema de Prioridades) ---
            for row in linhas_processadas:
                hist_up = row["Histórico"].upper()
                is_cred = False # Padrão é Débito
                
                # PRIORIDADE 1: Se tiver uma destas palavras, é garantidamente CRÉDITO (Entrada)
                if any(x in hist_up for x in ["RECEBID", "DEVOLU", "DESFAZIMENTO", "ESTORNO", "RESSARCIMENTO", "CREDIT", "CRÉDIT", "DEPÓSIT", "DEPOSIT"]):
                    is_cred = True
                # PRIORIDADE 2: Só se não for crédito explícito, é que verifica se é DÉBITO (Saída)
                elif any(x in hist_up for x in ["ENVIAD", "PAGAMENTO", "PAGTO", "SAQUE", "COMPRA", "DEBITO", "DÉBITO"]):
                    is_cred = False
                
                vals = row["Valores_Lista"]
                
                # Aloca os valores nas colunas corretas
                if is_cred:
                    row["Débito"] = ""
                    row["Crédito"] = vals[0] if len(vals) > 0 else ""
                else:
                    row["Débito"] = vals[0] if len(vals) > 0 else ""
                    row["Crédito"] = ""
                
                # O último valor do array nos extratos é sempre o Saldo Final daquele movimento
                row["Saldo"] = vals[-1] if len(vals) > 1 else ""
                
                # Limpa a chave temporária
                del row["Valores_Lista"]

            df = pd.DataFrame(linhas_processadas)
            return df

        # MODO 3: Texto Quebrado Simples
        elif modo == "Texto Bruto (Separar colunas por espaço)":
            for page in pdf.pages:
                text = page.extract_text()
                if not text: continue
                for l in text.split('\n'):
                    if l.strip():
                        all_rows.append(re.split(r'\s{2,}', l.strip()))
        
        if not all_rows:
            return pd.DataFrame()
            
        max_cols = max((len(row) for row in all_rows if row), default=0)
        normalized_rows = [row + [""] * (max_cols - len(row)) for row in all_rows if row]
        
        if modo == "Tabelas com Bordas (Padrão)" and len(normalized_rows) > 1:
            df = pd.DataFrame(normalized_rows[1:], columns=normalized_rows[0])
        else:
            df = pd.DataFrame(normalized_rows)
            
        return df

def dataframe_para_pdf(df, titulo="Relatório Convertido"):
    """Transforma um DataFrame do Pandas num ficheiro PDF formatado com quebra automática de texto."""
    buffer = io.BytesIO()
    
    if df.empty:
        # Previne erro caso tentem converter um Excel vazio
        doc = SimpleDocTemplate(buffer, pagesize=landscape(A4))
        styles = getSampleStyleSheet()
        doc.build([Paragraph("O relatório não contém dados para converter.", styles['Title'])])
        return buffer.getvalue()

    # A4 em paisagem tem aprox 842 pontos de largura. Com margens de 30, sobram cerca de 782 utilizáveis.
    largura_util = 782 
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=18)
    elements = []
    
    styles = getSampleStyleSheet()
    # Escapa caracteres especiais no título também (ex: &)
    titulo_formatado = Paragraph(f"<b>{html.escape(titulo)}</b>", styles['Title'])
    elements.append(titulo_formatado)
    elements.append(Spacer(1, 20))
    
    num_cols = len(df.columns)
    col_widths = []
    
    if num_cols > 0:
        # Lógica de proteção contra colunas totalmente vazias
        max_lens = []
        for col in df.columns:
            col_data = df[col].dropna().astype(str)
            if not col_data.empty:
                max_len = max(col_data.apply(len).max(), len(str(col)))
            else:
                max_len = len(str(col))
                
            max_lens.append(min(max(max_len, 5), 80))
        
        total_len = sum(max_lens)
        col_widths = [(l / total_len) * largura_util for l in max_lens]
    
    # Cria Estilos Separados para Cabeçalho e Células para não se sobreporem
    style_header = ParagraphStyle(
        'HeaderStyle',
        parent=styles['Normal'],
        fontSize=9,
        textColor=colors.whitesmoke,
        alignment=1, # Centro
        fontName='Helvetica-Bold'
    )
    
    style_cell = ParagraphStyle(
        'CellStyle',
        parent=styles['Normal'],
        fontSize=8,
        textColor=colors.HexColor("#2C3E50"),
        alignment=1 # Centro
    )
    
    cabecalho = [Paragraph(html.escape(str(c)), style_header) for c in df.columns]
    dados_formatados = [cabecalho]
    
    for _, row in df.iterrows():
        # html.escape protege o ReportLab de caracteres perigosos como <, > ou &
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
    
    # Adiciona a cor de linha alternada apenas se existir mais do que a linha de cabeçalho
    if len(dados_formatados) > 1:
        estilo.add('BACKGROUND', (0,1), (-1,-1), colors.HexColor("#ECF0F1"))
        
    t.setStyle(estilo)
    
    elements.append(t)
    doc.build(elements)
    
    return buffer.getvalue()


# --- INTERFACE DO UTILIZADOR ---

st.title("🔄 Conversor Universal")
st.markdown("Converta ficheiros **PDF para Excel** e **Excel para PDF** em segundos.")

aba1, aba2 = st.tabs(["📄 PDF ➡️ Excel", "📊 Excel ➡️ PDF"])

# ==========================================
# ABA 1: PDF PARA EXCEL
# ==========================================
with aba1:
    st.subheader("Extrair tabelas de um ficheiro PDF")
    
    modo_extracao = st.radio(
        "Selecione o Método de Extração:",
        [
            "Inteligente (Agrupar por Data) - Ideal para Extratos Bancários",
            "Tabelas com Bordas (Padrão)",
            "Texto Bruto (Separar colunas por espaço)"
        ]
    )
    
    st.info("💡 **Dica:** Os Extratos Bancários devem ser extraídos pelo modo Inteligente. Se a tabela final ficar estranha, troque de modo e converta de novo.")
    
    pdf_file = st.file_uploader("Arraste o seu PDF aqui", type=["pdf"], key="pdf_up")
    
    if pdf_file:
        with st.spinner("A analisar a estrutura do PDF..."):
            try:
                df_pdf = pdf_para_dataframe(pdf_file, modo_extracao)
                
                if df_pdf.empty:
                    st.warning("Não foi possível encontrar dados neste PDF usando o modo atual.")
                else:
                    st.success("PDF extraído com sucesso!")
                    st.dataframe(df_pdf, use_container_width=True)
                    
                    out_excel = io.BytesIO()
                    with pd.ExcelWriter(out_excel, engine='xlsxwriter') as wr:
                        df_pdf.to_excel(wr, index=False)
                    
                    nome_excel = pdf_file.name.replace('.pdf', '.xlsx')
                    st.download_button(
                        label="📥 Baixar Excel",
                        data=out_excel.getvalue(),
                        file_name=nome_excel,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Erro ao processar o PDF: {e}")

# ==========================================
# ABA 2: EXCEL PARA PDF
# ==========================================
with aba2:
    st.subheader("Transformar grelha em Documento PDF")
    
    excel_file = st.file_uploader("Arraste o seu Excel/CSV aqui", type=["xlsx", "xls", "csv"], key="excel_up")
    
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
                    mime="application/pdf"
                )
            except Exception as e:
                st.error(f"Erro ao processar o ficheiro: {e}")
