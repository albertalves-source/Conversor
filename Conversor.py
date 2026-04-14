import streamlit as st
import pandas as pd
import io
import warnings
import re

# Tenta importar as bibliotecas específicas para PDF
try:
    import pdfplumber
    from reportlab.lib.pagesizes import landscape, A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
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
                            # Limpa quebras de linha dentro das células do PDF
                            cleaned_row = [str(cell).replace('\n', ' ').strip() if cell else "" for cell in row]
                            all_rows.append(cleaned_row)
                            
        # MODO 2: Extratos Bancários (Inteligência Contábil de Débito/Crédito)
        elif modo == "Inteligente (Agrupar por Data) - Ideal para Extratos Bancários":
            linhas_processadas = []
            linha_atual = {}
            
            def is_id(token):
                # Heurística para identificar códigos de transação/IDs (Alfanuméricos longos)
                if len(token) >= 6 and re.search(r'\d', token) and (re.search(r'[a-zA-Z]', token) or '-' in token):
                    # Ignora se por acaso for uma data mal formatada
                    if not re.search(r'\d{2}/\d{2}/\d{2,4}', token):
                        return True
                # IDs puramente numéricos mas muito grandes (ex: código de barras)
                if len(token) > 10 and token.isdigit():
                    return True
                return False

            # Regex para detetar valores financeiros com ou sem R$ (ex: R$ 43,52 ou -1.250,00)
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
                        # Guarda a transação anterior na lista
                        if linha_atual:
                            linhas_processadas.append(linha_atual)
                            
                        data = match_data.group(1)
                        resto = match_data.group(2)
                        
                        # 1. Extrai Valores Monetários
                        valores = re.findall(regex_valor, resto)
                        texto_sem_valor = resto
                        for v in valores:
                            texto_sem_valor = texto_sem_valor.replace(v, '').strip()
                            
                        # 2. Tenta separar o ID longo do Histórico
                        tokens = texto_sem_valor.split()
                        id_transacao = ""
                        historico = texto_sem_valor
                        
                        # Verifica os últimos tokens para ver se são o ID da transação
                        for token in reversed(tokens[-3:]):
                            if is_id(token):
                                id_transacao = token + id_transacao
                                historico = historico.replace(token, '').strip()
                        
                        # 3. Inteligência Contábil Inicial
                        hist_upper = historico.upper()
                        is_credito = False # Padrão
                        if any(x in hist_upper for x in ["RECEBID", "DEVOLU", "DESFAZIMENTO", "ESTORNO", "RESSARCIMENTO", "CREDITO", "CRÉDITO", "DEPÓSITO", "DEPOSITO"]):
                            is_credito = True
                        if any(x in hist_upper for x in ["ENVIAD", "PAGAMENTO", "PAGTO", "SAQUE", "COMPRA", "DEBITO", "DÉBITO"]):
                            is_credito = False
                        
                        linha_atual = {
                            "Data": data,
                            "Histórico": historico.strip(),
                            "ID / Documento": id_transacao.strip(),
                            "Valores_Lista": valores,
                            "Is_Credito": is_credito
                        }
                    else:
                        # Se não tem data, é linha de continuação da transação anterior
                        if linha_atual:
                            # Tenta puxar valores perdidos nesta linha
                            valores_extra = re.findall(regex_valor, line)
                            if valores_extra:
                                linha_atual["Valores_Lista"].extend(valores_extra)
                                
                            texto_extra = line
                            for v in valores_extra:
                                texto_extra = texto_extra.replace(v, '').strip()
                                
                            # Tenta puxar restos de IDs que foram quebrados
                            tokens_extra = texto_extra.split()
                            id_extra = ""
                            hist_extra = texto_extra
                            
                            for token in reversed(tokens_extra[-3:]):
                                if is_id(token):
                                    id_extra = token + id_extra
                                    hist_extra = hist_extra.replace(token, '').strip()
                                    
                            # Cola o que sobrou no sítio certo
                            if hist_extra:
                                linha_atual["Histórico"] += " " + hist_extra.strip()
                            if id_extra:
                                linha_atual["ID / Documento"] += id_extra.strip()
                                
                            # Reavaliação da inteligência contábil com as novas palavras
                            hist_completo_up = linha_atual["Histórico"].upper()
                            if any(x in hist_completo_up for x in ["RECEBID", "DEVOLU", "DESFAZIMENTO", "ESTORNO", "RESSARCIMENTO", "CREDITO", "CRÉDITO", "DEPÓSITO", "DEPOSITO"]):
                                linha_atual["Is_Credito"] = True
                            elif any(x in hist_completo_up for x in ["ENVIAD", "PAGAMENTO", "PAGTO", "SAQUE", "COMPRA", "DEBITO", "DÉBITO"]):
                                linha_atual["Is_Credito"] = False
            
            if linha_atual: # Guarda a última linha do documento
                linhas_processadas.append(linha_atual)
                
            # Tratamento Final de Colunas
            for row in linhas_processadas:
                vals = row["Valores_Lista"]
                is_cred = row["Is_Credito"]
                
                # Regras de Débito, Crédito e Saldo
                if is_cred:
                    row["Débito"] = ""
                    row["Crédito"] = vals[0] if len(vals) > 0 else ""
                else:
                    row["Débito"] = vals[0] if len(vals) > 0 else ""
                    row["Crédito"] = ""
                
                # Se existir um 2º ou 3º valor, normalmente o último é o saldo do dia
                row["Saldo"] = vals[-1] if len(vals) > 1 else ""
                
                # Limpeza das chaves temporárias
                del row["Valores_Lista"]
                del row["Is_Credito"]

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
            
        # Ajusta o tamanho das linhas para o Pandas não dar erro
        max_cols = max((len(row) for row in all_rows if row), default=0)
        normalized_rows = [row + [""] * (max_cols - len(row)) for row in all_rows if row]
        
        # Só assume a primeira linha como cabeçalho se o utilizador escolher Tabelas Padrão
        if modo == "Tabelas com Bordas (Padrão)" and len(normalized_rows) > 1:
            df = pd.DataFrame(normalized_rows[1:], columns=normalized_rows[0])
        else:
            df = pd.DataFrame(normalized_rows)
            
        return df

def dataframe_para_pdf(df, titulo="Relatório Convertido"):
    """Transforma um DataFrame do Pandas num ficheiro PDF formatado."""
    buffer = io.BytesIO()
    # Usa landscape (paisagem) para caberem mais colunas
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=18)
    elements = []
    
    # Adiciona um título ao PDF
    styles = getSampleStyleSheet()
    titulo_formatado = Paragraph(f"<b>{titulo}</b>", styles['Title'])
    elements.append(titulo_formatado)
    elements.append(Spacer(1, 20))
    
    # Prepara os dados (Converte tudo para texto para evitar erros no ReportLab)
    colunas = [str(c) for c in df.columns]
    dados = [colunas] + df.astype(str).values.tolist()
    
    # Cria a Tabela
    t = Table(dados)
    
    # Estilo da Tabela
    estilo = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2C3E50")), # Cor do cabeçalho
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 10),
        ('BOTTOMPADDING', (0,0), (-1,0), 10),
        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor("#ECF0F1")), # Cor das linhas
        ('TEXTCOLOR', (0,1), (-1,-1), colors.HexColor("#2C3E50")),
        ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,1), (-1,-1), 8),
        ('GRID', (0,0), (-1,-1), 1, colors.HexColor("#BDC3C7")), # Linhas da grelha
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ])
    t.setStyle(estilo)
    
    elements.append(t)
    doc.build(elements)
    
    return buffer.getvalue()


# --- INTERFACE DO UTILIZADOR ---

st.title("🔄 Conversor Universal")
st.markdown("Converta ficheiros **PDF para Excel** e **Excel para PDF** em segundos.")

# Cria as duas abas
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
                    
                    # Botão de Download
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
                
                # Gera o PDF
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
