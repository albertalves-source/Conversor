import streamlit as st
import pandas as pd
import io
import warnings

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

def pdf_para_dataframe(file):
    """Lê um PDF, procura tabelas e transforma num DataFrame do Pandas."""
    with pdfplumber.open(file) as pdf:
        all_rows = []
        encontrou_tabela = False
        
        for page in pdf.pages:
            tables = page.extract_tables()
            if tables:
                encontrou_tabela = True
                for table in tables:
                    for row in table:
                        # Limpa quebras de linha dentro das células do PDF
                        cleaned_row = [str(cell).replace('\n', ' ').strip() if cell else "" for cell in row]
                        all_rows.append(cleaned_row)
            else:
                # Se não achar tabela, tenta ler texto linha a linha (Fallback)
                text = page.extract_text()
                if text and not encontrou_tabela:
                    for line in text.split('\n'):
                        # Separa por espaços múltiplos para simular colunas
                        import re
                        cleaned_row = re.split(r'\s{2,}', line.strip())
                        all_rows.append(cleaned_row)
        
        if not all_rows:
            return pd.DataFrame()
            
        # Ajusta o tamanho das linhas para não dar erro no Pandas
        max_cols = max(len(row) for row in all_rows)
        normalized_rows = [row + [""] * (max_cols - len(row)) for row in all_rows]
        
        # Assume a primeira linha como cabeçalho
        if len(normalized_rows) > 1:
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
    st.info("💡 **Dica:** Este método funciona melhor em PDFs gerados por sistemas (como extratos bancários) que contenham tabelas estruturadas.")
    
    pdf_file = st.file_uploader("Arraste o seu PDF aqui", type=["pdf"], key="pdf_up")
    
    if pdf_file:
        with st.spinner("A analisar a estrutura do PDF..."):
            try:
                df_pdf = pdf_para_dataframe(pdf_file)
                
                if df_pdf.empty:
                    st.warning("Não foi possível encontrar tabelas estruturadas neste PDF.")
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