import streamlit as st
import pdfplumber
import pandas as pd
import re
import unicodedata
from fpdf import FPDF
from datetime import datetime
from zoneinfo import ZoneInfo

class DataEngine:
    @staticmethod
    def normalizar(texto):
        if not texto: return ""
        nfkd_form = unicodedata.normalize('NFKD', str(texto).upper().strip())
        return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

    @staticmethod
    def limpar_valor(v):
        if v is None: return 0.0
        v_str = str(v).upper()
        if "USD" in v_str: return 0.0
        v_limpo = re.sub(r'[^\d,\.]', '', v_str)
        if not v_limpo or v_limpo == '.': return 0.0
        if ',' in v_limpo: v_limpo = v_limpo.replace('.', '').replace(',', '.')
        try: return float(v_limpo)
        except: return 0.0

    @classmethod
    def extrair_pdf(cls, arquivo):
        mapa_final = {}
        with pdfplumber.open(arquivo) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        if "R$" in linha:
                            partes = linha.split("R$")
                            nome = re.sub(r'[^A-Z/À-ÿ ]', '', partes[0].upper())
                            nome = re.sub(r'\s+', ' ', nome).strip()
                            valor = cls.limpar_valor(partes[1].strip().split(" ")[0])
                            if len(nome) > 3 and "TOTAL" not in nome:
                                mapa_final[nome] = valor
                tabelas = page.extract_tables()
                for tab in tabelas:
                    df = pd.DataFrame(tab)
                    if not df.empty and any("II" in str(c).upper() for c in df.iloc[0]):
                        headers = [str(h).strip().upper() for h in df.iloc[0]]
                        for _, row in df.iterrows():
                            if "RECOLHER" in str(row[0]).upper():
                                for i, cell in enumerate(row):
                                    if i < len(headers):
                                        n = re.sub(r'[^A-Z ]', '', headers[i]).strip()
                                        if len(n) > 1: mapa_final[n] = cls.limpar_valor(cell)
                                break
        return mapa_final

class relatorioImg(FPDF):
    # Cabecalho da plannilha
    def header(self):
        # Logo
        try:
            self.image('jslogo.png', 14, 10, 48)
        except:
            pass
            
        # Titulo
        self.set_y(14) 
        self.set_font('Helvetica', 'B', 16)
        titulo = 'ESTIMATIVA DE CUSTOS DE IMPORTAÇÃO E EXPORTAÇÃO'
        titulo_pdf = titulo.encode('latin-1', 'replace').decode('latin-1')
        self.cell(0, 12, titulo_pdf, 0, 1, 'C') 
        
        # Registro de data e hora no canto do cabecalho
        self.set_y(27) 
        self.set_font('Helvetica', '', 8)
        data_texto = f'Gerado em: {datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%d/%m/%Y %H:%M (%Z)")}'
        self.cell(0, 5, data_texto, 0, 1, 'R')

        # Ajustes para separar o cabecalho das colunas e colocar uma linha 
        self.set_draw_color(114, 171, 181) 
        self.set_line_width(0.5)
        self.line(10, 34, 287, 34)
        self.ln(20) 

    def criar_layout(self, df_selecionado, fob, frete):
        self.add_page()
        cor_header = (114, 171, 181) # Aqui é o cabecalho de cada coluna
        cor_zebra = (245, 245, 245)
        
        t_imp = df_selecionado[df_selecionado["Grupo"] == "Impostos"]["Valor (R$)"].sum()
        t_tax = df_selecionado[df_selecionado["Grupo"] == "Taxas"]["Valor (R$)"].sum()
        t_des = df_selecionado[df_selecionado["Grupo"] == "Despesas"]["Valor (R$)"].sum()

        larg_col, espaco, x_in, y_in = 65, 6, 10, 45
        grupos = ["Impostos", "Taxas", "Despesas"]
        for idx, g in enumerate(grupos):
            pos_x = x_in + (idx * (larg_col + espaco))
            self.set_xy(pos_x, y_in)
            self.set_fill_color(*cor_header)
            self.set_text_color(255, 255, 255)
            self.set_font('Helvetica', 'B', 10)
            self.cell(larg_col, 8, g.upper(), 0, 1, 'C', 1)
            
            self.set_text_color(0, 0, 0)
            self.set_font('Helvetica', '', 8)
            itens = df_selecionado[df_selecionado["Grupo"] == g]
            for i, (_, row) in enumerate(itens.iterrows()):
                self.set_x(pos_x)
                fill = 1 if i % 2 == 1 else 0
                self.set_fill_color(*cor_zebra)
                nome = DataEngine.normalizar(row['Item'])[:28]
                self.cell(larg_col * 0.65, 7, f" {nome}", 0, 0, 'L', fill)
                self.cell(larg_col * 0.35, 7, f"R$ {row['Valor (R$)']:,.2f} ", 0, 1, 'R', fill)

        x_box, y_box = 230, 45
        self.set_xy(x_box, y_box)
        self.set_fill_color(230, 230, 230)
        self.set_font('Helvetica', 'B', 8)
        self.cell(20, 6, " FOB", 0, 0, 'L', 1)
        self.set_font('Helvetica', '', 8)
        self.cell(35, 6, f"R$ {fob:,.2f}", 0, 1, 'R', 1)
        self.set_x(x_box)
        self.cell(20, 6, " FRETE", 0, 0, 'L', 1)
        self.set_font('Helvetica', '', 8)
        self.cell(35, 6, f"R$ {frete:,.2f}", 0, 1, 'R', 1)
        
        self.ln(8)
        for label, val in [("Total Impostos", t_imp), ("Total Taxas", t_tax), ("Total Despesas", t_des)]:
            self.set_x(x_box); self.set_fill_color(210, 210, 210)
            self.set_font('Helvetica', 'B', 7)
            self.cell(30, 6, f" {label}", 0, 0, 'L', 1)
            self.set_font('Helvetica', '', 8)
            self.cell(25, 6, f"R$ {val:,.2f}", 0, 1, 'R', 1)
            self.ln(2)

        self.ln(4); self.set_x(x_box)
        self.set_fill_color(31, 110, 212); self.set_text_color(255, 255, 255)
        self.set_font('Helvetica', 'B', 10)
        self.cell(30, 10, " TOTAL GERAL:", 0, 0, 'L', 1)
        self.cell(25, 10, f"R$ {(t_imp+t_tax+t_des):,.2f}", 0, 1, 'R', 1)
        return bytes(self.output())

# ================== Aqui é a interface ==================
st.set_page_config(page_title="Estimativas JS Energy", layout="centered")

st.title("PLANILHA DE ESTIMATIVAS")

def sync_editor():
    if "editor_v27" in st.session_state:
        edits = st.session_state["editor_v27"]
        for idx, changes in edits.get("edited_rows", {}).items():
            for col, val in changes.items():
                st.session_state.df_final.at[idx, col] = val
        for row in edits.get("added_rows", []):
            new_row = {"Incluir?": True, "Item": "Novo Item", "Valor (R$)": 0.0, "Grupo": "Despesas"}
            new_row.update(row)
            st.session_state.df_final = pd.concat([st.session_state.df_final, pd.DataFrame([new_row])], ignore_index=True)
        for idx in edits.get("deleted_rows", []):
            st.session_state.df_final = st.session_state.df_final.drop(idx).reset_index(drop=True)

if 'df_final' not in st.session_state:
    st.session_state.df_final = pd.DataFrame(columns=["Incluir?", "Item", "Valor (R$)", "Grupo"])

uploaded_file = st.file_uploader("Suba o PDF original", type=["pdf"])

if uploaded_file:
    if 'pdf_map' not in st.session_state:
        st.session_state.pdf_map = DataEngine.extrair_pdf(uploaded_file)
        prioritarios = ["II", "IPI", "PIS", "COFINS", "ICMS", "FECP", "AFRMM", "TAXA SISCOMEX"]
        lista_inicial = []
        for p in prioritarios:
            val, nome = 0.0, p
            for k, v in st.session_state.pdf_map.items():
                if DataEngine.normalizar(p) in DataEngine.normalizar(k):
                    val, nome = v, k
                    break
            g = "Impostos" if p not in ["AFRMM", "TAXA SISCOMEX"] else "Taxas"
            lista_inicial.append({"Incluir?": True, "Item": nome, "Valor (R$)": val, "Grupo": g})
        st.session_state.df_final = pd.DataFrame(lista_inicial)

    st.subheader("Edição da Estimativa")
    df_atualizado = st.data_editor(
        st.session_state.df_final,
        column_config={
            "Incluir?": st.column_config.CheckboxColumn("Incluir?", default=True, width="small"),
            "Grupo": st.column_config.SelectboxColumn(options=["Impostos", "Taxas", "Despesas"]),
            "Valor (R$)": st.column_config.NumberColumn(format="R$ %.2f", step=0.01)
        },
        num_rows="dynamic", use_container_width=True, hide_index=True, 
        key="editor_v27", on_change=sync_editor
    )
    st.session_state.df_final = df_atualizado

    st.divider()
    col_info, col_busca = st.columns(2)

    with col_info:
        st.subheader("Dados Gerais")
        fob = st.number_input("Valor FOB (BRL)", value=st.session_state.pdf_map.get("FOB", 0.0), format="%.2f")
        frete = st.number_input("Valor Frete (BRL)", value=st.session_state.pdf_map.get("FRETE", 0.0), format="%.2f")

    with col_busca:
        st.subheader("Adicionar do PDF")
        busca = st.text_input("Localizar item...")
        if busca:
            matches = {k: v for k, v in st.session_state.pdf_map.items() if DataEngine.normalizar(busca) in DataEngine.normalizar(k)}
            for n, v in matches.items():
                c_item, c_grp, c_btn = st.columns([2, 1.5, 0.5])
                c_item.write(f"**{n}**")
                grp_escolhido = c_grp.selectbox("Grupo", ["Impostos", "Taxas", "Despesas"], key=f"grp_{n}", index=2, label_visibility="collapsed")
                if c_btn.button("➕", key=f"btn_{n}"):
                    nova_linha = pd.DataFrame([{"Incluir?": True, "Item": n, "Valor (R$)": v, "Grupo": grp_escolhido}])
                    st.session_state.df_final = pd.concat([st.session_state.df_final, nova_linha], ignore_index=True)
                    st.rerun()

    st.divider()
    df_sel = st.session_state.df_final[st.session_state.df_final["Incluir?"] == True]
    total = df_sel["Valor (R$)"].sum()
    
    st.markdown(f"<h3 style='text-align: center;'>Total Selecionado: R$ {total:,.2f}</h3>", unsafe_allow_html=True)
    
    if st.button("Gerar PDF Selecionado", use_container_width=True):
        if df_sel.empty:
            st.warning("Selecione ao menos um item!")
        else:
            pdf_gen = relatorioImg(orientation='L')
            arquivo_bytes = pdf_gen.criar_layout(df_sel, fob, frete)
            
            data_hora = datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%d_%m_%Y_%Hh%M")
            nome_arquivo = f"Estimativa_Energy_{data_hora}.pdf"
            
            st.download_button(
                label=f"Baixar PDF ({nome_arquivo})", 
                data=arquivo_bytes, 
                file_name=nome_arquivo, 
                mime="application/pdf", 
                use_container_width=True
            )
           
