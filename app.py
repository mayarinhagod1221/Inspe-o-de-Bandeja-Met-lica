import streamlit as st
import pandas as pd
from datetime import datetime
import os

# Google Sheets
import gspread
from google.oauth2.service_account import Credentials

# PDF
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Image

# ================= CONFIG =================
st.set_page_config(page_title="Inspeção Bandeja Metálica", layout="wide")

# ================= GOOGLE SHEETS =================
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, "credenciais.json")
NOME_PLANILHA = "Inspecao Bandeja Metalica"

# ================= VERIFICA CREDENCIAL =================
if not os.path.exists(SERVICE_ACCOUNT_FILE):
    st.error("credenciais.json não encontrado!")
    st.stop()

# ================= FUNÇÃO SHEETS =================
def conectar_sheets():
    credentials = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=SCOPES
    )
    client = gspread.authorize(credentials)
    return client.open(NOME_PLANILHA).sheet1

def salvar_no_sheets(df):
    try:
        aba = conectar_sheets()
        for linha in df.values.tolist():
            aba.append_row(linha)
        return True
    except Exception as e:
        if "200" in str(e):
            return True
        st.error(f"Erro ao salvar: {e}")
        return False

# ================= ESTILO =================
st.markdown("""
<style>
.stApp { background-color: #0e1117; color: white; }
</style>
""", unsafe_allow_html=True)

# ================= HEADER =================
col_logo, col_titulo = st.columns([1, 4])

with col_logo:
    st.image("logo.png", width=120)

with col_titulo:
    st.title("INSPEÇÃO BANDEJA METÁLICA")
    st.caption("Desenvolvido por Programadora Web Mayara Helouise Maia")

col1, col2 = st.columns(2)

with col1:
    st.write("**Data da Edição:** 01/08/2025")
    st.write("**Revisão:** 3")
    st.write("**Setor:** Qualidade Industrial")

with col2:
    st.write("**Autor(a):** Evelyn Silva")
    st.write("**Verificação:** Melina Favaro")
    st.write("**Inspeção:** a cada 60 minutos")

st.divider()

st.subheader("Procedimento de Inspeção")

st.write(
    "**1° PASSO:** Inspecione as medidas das abas maiores (A e B), o comprimento (C) e a medida (D) "
    "com auxílio de uma trena e um paquímetro, seguindo as cotas dos desenhos das ordens de produção "
    "e preencha nos registros abaixo."
)

st.write(
    "**2° PASSO:** Com auxílio de um paquímetro e um esquadro, inspecione as abas menores e registre "
    '"OK" ou "N OK".'
)

st.write(
    "**3° PASSO:** Caso necessário, realize a pré montagem, liberando as medidas sob concessão e "
    "registrando nas observações."
)

st.divider()

# ================= FORM =================
col1, col2, col3, col4 = st.columns(4)
pr = col1.text_input("PR (Máquina)", key="pr")
inspetor = col2.text_input("Inspetor")
op = col3.text_input("O.P.")
data = col4.date_input("Data", datetime.now())

# ================= SESSION =================
if "dados" not in st.session_state:
    st.session_state.dados = []

if "salvo" not in st.session_state:
    st.session_state.salvo = False

# ================= ENTRADA =================
col1, col2, col3, col4, col5, col6 = st.columns(6)

hora = datetime.now().strftime("%H:%M:%S")
col1.text_input("Hora", value=hora, disabled=True)

A = col2.number_input("A")
B = col3.number_input("B")
C = col4.number_input("C")
D = col5.number_input("D")
esquadro = col6.selectbox("Esquadro", ["OK", "Ñ OK"])

observacoes = st.text_area("Observações", key="obs")
# ================= ADICIONAR =================
if st.button("Adicionar"):
    id_registro = datetime.now().strftime("%Y%m%d%H%M%S%f")

    nova_linha = {
    "ID": id_registro,
    "PR": pr,
    "Inspetor": inspetor,
    "OP": op,
    "Data": data.strftime("%d/%m/%Y"),
    "Hora": hora,
    "A": A,
    "B": B,
    "C": C,
    "D": D,
    "Esquadro": esquadro,
    "Observações": observacoes
}
    
    st.session_state.dados.append(nova_linha)
    st.session_state.salvo = False

# ================= DATAFRAME =================
df = pd.DataFrame(st.session_state.dados)

if not df.empty:
    st.dataframe(df, use_container_width=True)


def gerar_pdf():
    try:
        nome = f"inspecao_{op}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
        doc = SimpleDocTemplate(nome)
        styles = getSampleStyleSheet()

        elementos = []

        # LOGO
        logo_path = os.path.join(BASE_DIR, "logo.png")
        if os.path.exists(logo_path):
            logo = Image(logo_path, width=120, height=60)
            logo.hAlign = 'CENTER'
            elementos.append(logo)

        # CONTEÚDO
        elementos.append(Paragraph("INSPEÇÃO BANDEJA METÁLICA", styles['Title']))
        elementos.append(Paragraph(f"Inspetor: {inspetor}", styles['Normal']))
        elementos.append(Paragraph(f"O.P.: {op}", styles['Normal']))
        elementos.append(Paragraph(f"Data: {data.strftime('%d/%m/%Y')}", styles['Normal']))

        tabela = Table([df.columns.tolist()] + df.values.tolist())
        tabela.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))

        elementos.append(tabela)
        elementos.append(Paragraph(f"Observações: {observacoes}", styles['Normal']))

        doc.build(elementos)

        return nome

    except Exception as e:
        st.error(f"Erro ao gerar PDF: {e}")
        return None
    
# ================= BOTÕES =================
colA, colB = st.columns(2)

pdf = None

with colA:
    if st.button("Gerar PDF"):
        if df.empty:
            st.warning("Adicione dados!")
        else:
            pdf = gerar_pdf()

            if pdf and os.path.exists(pdf):
                with open(pdf, "rb") as f:
                    st.download_button("Baixar PDF", f, file_name=pdf)
            else:
                st.error("Erro ao gerar PDF.")

with colB:
    if st.button("Salvar no Google Sheets"):
        if df.empty:
            st.warning("Adicione dados!")
        elif st.session_state.salvo:
            st.info("Dados já foram enviados!")
        else:
            sucesso = salvar_no_sheets(df)

            if sucesso:
                st.success("Dados enviados com sucesso!")
                st.session_state.salvo = True
            else:
                st.error("Erro ao salvar.")