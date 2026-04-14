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
                            
        # MODO 2: Extratos Bancários (Agrupa texto solto na mesma data)
        elif modo == "Inteligente (Agrupar por Data) - Ideal para Extratos Bancários":
            linha_atual = []
            for page in pdf.pages:
                text = page.extract_text()
                if not text: continue
                
                for l in text.split('\n'):
                    if not l.strip(): continue
                    
                    # Se começar com uma Data, é o início de uma nova transação
                    if re.search(r'^\s*\d{2}[/-]\d{2}[/-]\d{2,4}', l):
                        if linha_atual:
                            all_rows.append(linha_atual)
                        # Separa as colunas por grandes espaços
                        linha_atual = re.split(r'\s{2,}', l.strip())
                    else:
                        # Se não tem data, é a continuação do histórico anterior (ou cabeçalho)
                        if linha_atual:
                            partes_extra = re.split(r'\s{2,}', l.strip())
                            if len(partes_extra) == 1:
                                # Cola na 2ª coluna (Histórico) ou 1ª se só houver uma
                                if len(linha_atual) > 1:
                                    linha_atual[1] += " " + partes_extra[0]
                                else:
                                    linha_atual[0] += " " + partes_extra[0]
                            else:
                                # Tenta colar nas colunas respetivas
                                for idx, p in enumerate(partes_extra):
                                    target_idx = min(idx + 1, len(linha_atual) - 1)
                                    linha_atual[target_idx] += " " + p
                        else:
                            # Cabeçalhos iniciais antes das transações
                            all_rows.append(re.split(r'\s{2,}', l.strip()))
            
            if linha_atual: # Guarda a última linha
                all_rows.append(linha_atual)

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
