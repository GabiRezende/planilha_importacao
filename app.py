import hashlib
import re
import unicodedata
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import pdfplumber
import streamlit as st
from fpdf import FPDF


class DataEngine:
    CAMPOS_PRIORITARIOS = [
        "FOB",
        "FRETE",
        "II",
        "IPI",
        "PIS",
        "COFINS",
        "ICMS",
        "ICMS REPETRO",
        "FECP",
        "AFRMM",
        "TAXA SISCOMEX",
    ]

    @staticmethod
    def normalizar(texto):
        if not texto:
            return ""
        nfkd_form = unicodedata.normalize("NFKD", str(texto).upper().strip())
        return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

    @staticmethod
    def limpar_valor(v):
        if v is None:
            return 0.0

        v_str = str(v).strip().upper()
        if not v_str:
            return 0.0

        v_limpo = re.sub(r"[^\d,.\-]", "", v_str)
        if not v_limpo or v_limpo in [".", ",", "-", "-.", "-,"]:
            return 0.0

        if "," in v_limpo:
            v_limpo = v_limpo.replace(".", "").replace(",", ".")

        try:
            return float(v_limpo)
        except Exception:
            return 0.0

    @classmethod
    def extrair_primeiro_numero(cls, texto):
        if not texto:
            return None

        texto = str(texto)

        padroes = [
            r"\d{1,3}(?:\.\d{3})*,\d{2}",  # 1.234,56
            r"\d+,\d{2}",                  # 1234,56
            r"\d+\.\d{2}",                 # 1234.56
        ]

        for padrao in padroes:
            m = re.search(padrao, texto)
            if m:
                return cls.limpar_valor(m.group(0))

        return None

    @classmethod
    def canonizar_campo(cls, texto):
        n = cls.normalizar(texto)

        if not n:
            return None

        if "ICMS REPETRO" in n:
            return "ICMS REPETRO"
        if "TAXA SISCOMEX" in n or "SISCOMEX" in n:
            return "TAXA SISCOMEX"
        if "AFRMM" in n:
            return "AFRMM"
        if re.search(r"\bFOB\b", n):
            return "FOB"
        if "FRETE" in n:
            return "FRETE"
        if re.search(r"\bII\b", n):
            return "II"
        if re.search(r"\bIPI\b", n):
            return "IPI"
        if re.search(r"\bPIS\b", n):
            return "PIS"
        if re.search(r"\bCOFINS\b", n):
            return "COFINS"
        if "ICMS" in n:
            return "ICMS"
        if re.search(r"\bFECP\b", n):
            return "FECP"

        return None

    @classmethod
    def extrair_linhas_texto(cls, page):
        try:
            texto = page.extract_text(x_tolerance=2, y_tolerance=3)
        except Exception:
            texto = None

        if texto:
            return [linha.strip() for linha in texto.split("\n") if linha.strip()]
        return []

    @classmethod
    def extrair_linhas_palavras(cls, page):
        try:
            palavras = page.extract_words(use_text_flow=True, keep_blank_chars=False)
        except Exception:
            palavras = []

        if not palavras:
            return []

        linhas = {}
        for p in palavras:
            topo = round(p["top"] / 3) * 3
            linhas.setdefault(topo, []).append(p)

        resultado = []
        for topo in sorted(linhas.keys()):
            grupo = sorted(linhas[topo], key=lambda x: x["x0"])
            linha = " ".join(p["text"] for p in grupo).strip()
            if linha:
                resultado.append(linha)

        return resultado

    @classmethod
    def varrer_linhas(cls, linhas, mapa):
        for linha in linhas:
            linha_norm = cls.normalizar(linha)

            if "R$" in linha:
                partes = linha.split("R$")
                nome = re.sub(r"[^A-Z/À-ÿ ]", "", partes[0].upper())
                nome = re.sub(r"\s+", " ", nome).strip()

                valor = None
                if len(partes) > 1:
                    valor = cls.extrair_primeiro_numero(partes[1])

                if nome and valor is not None and "TOTAL" not in cls.normalizar(nome):
                    mapa[nome] = valor

            for campo in cls.CAMPOS_PRIORITARIOS:
                campo_norm = cls.normalizar(campo)
                if campo_norm in linha_norm:
                    valor = cls.extrair_primeiro_numero(linha)
                    if valor is not None:
                        mapa[campo] = valor

    @classmethod
    def varrer_tabelas(cls, page, mapa):
        configuracoes = [
            {
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "intersection_tolerance": 5,
                "snap_tolerance": 4,
                "join_tolerance": 4,
            },
            {
                "vertical_strategy": "text",
                "horizontal_strategy": "text",
                "text_x_tolerance": 2,
                "text_y_tolerance": 2,
            },
            {
                "vertical_strategy": "lines",
                "horizontal_strategy": "text",
                "text_x_tolerance": 2,
                "text_y_tolerance": 2,
                "intersection_tolerance": 5,
            },
        ]

        for config in configuracoes:
            try:
                tabelas = page.extract_tables(table_settings=config) or []
            except Exception:
                tabelas = []

            for tab in tabelas:
                if not tab:
                    continue

                df = pd.DataFrame(tab).fillna("")

                header_idx = None
                recolher_idx = None

                for i in range(len(df)):
                    row_norm = [cls.normalizar(c) for c in df.iloc[i].tolist()]

                    if any(x in row_norm for x in ["II", "IPI", "PIS", "COFINS", "ICMS", "FECP"]) or \
                       any("AFRMM" in x or "SISCOMEX" in x or "ICMS REPETRO" in x for x in row_norm):
                        header_idx = i

                    if any("RECOLHER" in x for x in row_norm):
                        recolher_idx = i

                if header_idx is not None and recolher_idx is not None:
                    headers = df.iloc[header_idx].tolist()
                    valores = df.iloc[recolher_idx].tolist()

                    for j, h in enumerate(headers):
                        campo = cls.canonizar_campo(h)
                        if campo and j < len(valores):
                            valor = cls.extrair_primeiro_numero(valores[j])
                            if valor is not None:
                                mapa[campo] = valor

                for i in range(len(df)):
                    row = df.iloc[i].tolist()
                    row_texto = " | ".join(str(c) for c in row)
                    row_norm = cls.normalizar(row_texto)

                    for campo in cls.CAMPOS_PRIORITARIOS:
                        if cls.normalizar(campo) in row_norm:
                            numeros = [cls.extrair_primeiro_numero(c) for c in row]
                            numeros_validos = [n for n in numeros if n is not None]
                            if numeros_validos:
                                mapa[campo] = numeros_validos[-1]

    @classmethod
    def extrair_pdf(cls, pdf_bytes):
        mapa_final = {}

        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            for raw_page in pdf.pages:
                try:
                    page = raw_page.dedupe_chars()
                except Exception:
                    page = raw_page

                cls.varrer_linhas(cls.extrair_linhas_texto(page), mapa_final)
                cls.varrer_tabelas(page, mapa_final)

        if cls.resultado_fraco(mapa_final):
            with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
                for raw_page in pdf.pages:
                    try:
                        page = raw_page.dedupe_chars()
                    except Exception:
                        page = raw_page

                    cls.varrer_linhas(cls.extrair_linhas_palavras(page), mapa_final)
                    cls.varrer_tabelas(page, mapa_final)

        return mapa_final

    @classmethod
    def resultado_fraco(cls, mapa):
        if not mapa:
            return True

        encontrados = [
            campo for campo in cls.CAMPOS_PRIORITARIOS
            if mapa.get(campo) not in [None, 0, 0.0]
        ]
        return len(encontrados) < 3

    @classmethod
    def resumo_extracao(cls, mapa):
        encontrados = {
            campo: mapa.get(campo, 0.0)
            for campo in cls.CAMPOS_PRIORITARIOS
            if mapa.get(campo) not in [None, 0, 0.0]
        }
        return {
            "fraco": len(encontrados) < 3,
            "qtd_prioritarios": len(encontrados),
            "prioritarios_encontrados": list(encontrados.keys()),
            "itens_encontrados": encontrados,
        }


class RelatorioImg(FPDF):
    def header(self):
        try:
            self.image("jslogo.png", 14, 10, 48)
        except Exception:
            pass

        self.set_y(14)
        self.set_font("Helvetica", "B", 16)
        titulo = "ESTIMATIVA DE CUSTOS DE IMPORTAÇÃO E EXPORTAÇÃO"
        titulo_pdf = titulo.encode("latin-1", "replace").decode("latin-1")
        self.cell(0, 12, titulo_pdf, 0, 1, "C")

        self.set_y(27)
        self.set_font("Helvetica", "", 8)
        data_texto = f'Gerado em: {datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%d/%m/%Y %H:%M")} (BRT)'
        self.cell(0, 5, data_texto, 0, 1, "R")

        self.set_draw_color(114, 171, 181)
        self.set_line_width(0.5)
        self.line(10, 34, 287, 34)
        self.ln(20)

    def criar_layout(self, df_selecionado, fob, frete):
        self.add_page()

        cor_header = (114, 171, 181)
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
            self.set_font("Helvetica", "B", 10)
            self.cell(larg_col, 8, g.upper(), 0, 1, "C", 1)

            self.set_text_color(0, 0, 0)
            self.set_font("Helvetica", "", 8)
            itens = df_selecionado[df_selecionado["Grupo"] == g]

            for i, (_, row) in enumerate(itens.iterrows()):
                self.set_x(pos_x)
                fill = 1 if i % 2 == 1 else 0
                self.set_fill_color(*cor_zebra)
                nome = DataEngine.normalizar(row["Item"])[:28]
                self.cell(larg_col * 0.65, 7, f" {nome}", 0, 0, "L", fill)
                self.cell(larg_col * 0.35, 7, f'R$ {row["Valor (R$)"]:,.2f} ', 0, 1, "R", fill)

        x_box, y_box = 230, 45
        self.set_xy(x_box, y_box)
        self.set_fill_color(230, 230, 230)
        self.set_font("Helvetica", "B", 8)
        self.cell(20, 6, " FOB", 0, 0, "L", 1)
        self.set_font("Helvetica", "", 8)
        self.cell(35, 6, f"R$ {fob:,.2f}", 0, 1, "R", 1)

        self.set_x(x_box)
        self.set_font("Helvetica", "B", 8)
        self.cell(20, 6, " FRETE", 0, 0, "L", 1)
        self.set_font("Helvetica", "", 8)
        self.cell(35, 6, f"R$ {frete:,.2f}", 0, 1, "R", 1)

        self.ln(8)

        for label, val in [
            ("Total Impostos", t_imp),
            ("Total Taxas", t_tax),
            ("Total Despesas", t_des),
        ]:
            self.set_x(x_box)
            self.set_fill_color(210, 210, 210)
            self.set_font("Helvetica", "B", 7)
            self.cell(30, 6, f" {label}", 0, 0, "L", 1)
            self.set_font("Helvetica", "", 8)
            self.cell(25, 6, f"R$ {val:,.2f}", 0, 1, "R", 1)
            self.ln(2)

        self.ln(4)
        self.set_x(x_box)
        self.set_fill_color(31, 110, 212)
        self.set_text_color(255, 255, 255)
        self.set_font("Helvetica", "B", 10)
        self.cell(30, 10, " TOTAL GERAL:", 0, 0, "L", 1)
        self.cell(25, 10, f"R$ {(t_imp + t_tax + t_des):,.2f}", 0, 1, "R", 1)

        return bytes(self.output())


def build_initial_dataframe(pdf_map):
    prioritarios = ["II", "IPI", "PIS", "COFINS", "ICMS", "FECP", "AFRMM", "TAXA SISCOMEX"]
    lista_inicial = []

    for p in prioritarios:
        if p in pdf_map:
            val = pdf_map[p]
            nome = p
        else:
            val, nome = 0.0, p
            for k, v in pdf_map.items():
                if DataEngine.normalizar(p) == DataEngine.normalizar(k):
                    val, nome = v, k
                    break

        grupo = "Impostos" if p not in ["AFRMM", "TAXA SISCOMEX"] else "Taxas"
        lista_inicial.append(
            {
                "Incluir?": True,
                "Item": nome,
                "Valor (R$)": val,
                "Grupo": grupo,
            }
        )

    return pd.DataFrame(lista_inicial)


def reset_pdf_state(file_hash, file_name, pdf_bytes):
    st.session_state.arquivo_pdf_hash = file_hash
    st.session_state.arquivo_pdf_nome = file_name
    st.session_state.pdf_map = DataEngine.extrair_pdf(pdf_bytes)
    st.session_state.df_final = build_initial_dataframe(st.session_state.pdf_map)
    st.session_state.extracao_info = DataEngine.resumo_extracao(st.session_state.pdf_map)

    if "editor_v27" in st.session_state:
        del st.session_state["editor_v27"]


st.set_page_config(
    page_title="Estimativas JS Energy",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    .stApp {
        background: #0B1220;
    }

    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 1.8rem;
        max-width: 1280px;
    }

    .app-header {
        background: linear-gradient(180deg, #111827 0%, #0F172A 100%);
        border: 1px solid rgba(148, 163, 184, 0.16);
        border-radius: 14px;
        padding: 1rem 1.2rem;
        margin-bottom: 1rem;
    }

    .app-badge {
        display: inline-block;
        padding: 0.28rem 0.7rem;
        border-radius: 999px;
        background: rgba(37, 99, 235, 0.14);
        color: #93C5FD;
        font-size: 0.78rem;
        font-weight: 600;
        letter-spacing: 0.02em;
        margin-bottom: 0.55rem;
    }

    .app-header h1 {
        margin: 0;
        font-size: 1.7rem;
        color: #F8FAFC;
        font-weight: 700;
        line-height: 1.2;
    }

    .app-header p {
        margin: 0.25rem 0 0 0;
        color: #94A3B8;
        font-size: 0.96rem;
    }

    .metric-card {
        background: linear-gradient(180deg, #0F172A 0%, #111827 100%);
        border: 1px solid rgba(59, 130, 246, 0.18);
        border-radius: 14px;
        padding: 1rem 1.2rem;
        margin: 0.3rem 0 1rem 0;
    }

    .metric-label {
        color: #94A3B8;
        font-size: 0.85rem;
        margin-bottom: 0.15rem;
    }

    .metric-value {
        color: #F8FAFC;
        font-size: 1.65rem;
        font-weight: 700;
        line-height: 1.1;
    }

    div[data-testid="stFileUploader"] {
        border-radius: 12px;
    }

    div[data-testid="stDataEditor"] {
        border-radius: 12px;
        overflow: hidden;
        border: 1px solid rgba(148, 163, 184, 0.12);
    }

    div[data-baseweb="input"] > div,
    div[data-baseweb="select"] > div {
        border-radius: 10px;
    }

    div[data-testid="stButton"] > button,
    div[data-testid="stDownloadButton"] > button {
        border-radius: 10px;
        font-weight: 600;
        min-height: 2.8rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="app-header">
        <div class="app-badge">JS Energy • Importação e Exportação</div>
        <h1>Planilha de Estimativas</h1>
        <p>Extração automática de dados do PDF, edição manual e geração do relatório final.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

if "df_final" not in st.session_state:
    st.session_state.df_final = pd.DataFrame(
        columns=["Incluir?", "Item", "Valor (R$)", "Grupo"]
    )

uploaded_file = st.file_uploader(
    "Selecione o PDF original para extrair os dados",
    type=["pdf"],
)

if uploaded_file:
    pdf_bytes = uploaded_file.getvalue()
    file_hash = hashlib.sha1(pdf_bytes).hexdigest()

    if st.session_state.get("arquivo_pdf_hash") != file_hash:
        reset_pdf_state(file_hash, uploaded_file.name, pdf_bytes)

    extracao_info = st.session_state.get("extracao_info", {})

    if extracao_info.get("fraco"):
        st.warning(
            "A extração automática encontrou poucos campos confiáveis neste PDF. "
            "Ainda assim, você pode complementar a estimativa manualmente e buscar itens detectados."
        )

        with st.expander("Ver diagnóstico da extração"):
            st.write("Campos prioritários encontrados:")
            st.write(extracao_info.get("prioritarios_encontrados", []))
            st.write("Mapa bruto extraído:")
            st.json(st.session_state.get("pdf_map", {}))
    else:
        encontrados = extracao_info.get("prioritarios_encontrados", [])
        if encontrados:
            st.caption("Campos detectados automaticamente: " + ", ".join(encontrados))

    col_esq, col_dir = st.columns([1.75, 1], gap="large")

    with col_esq:
        st.subheader("Itens da estimativa")

        df_atualizado = st.data_editor(
            st.session_state.df_final,
            column_config={
                "Incluir?": st.column_config.CheckboxColumn(
                    "Incluir?",
                    default=True,
                    width="small",
                ),
                "Grupo": st.column_config.SelectboxColumn(
                    "Grupo",
                    options=["Impostos", "Taxas", "Despesas"],
                ),
                "Valor (R$)": st.column_config.NumberColumn(
                    "Valor (R$)",
                    format="R$ %.2f",
                    step=0.01,
                ),
            },
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key="editor_v27",
        )

        st.session_state.df_final = df_atualizado

    with col_dir:
        st.subheader("Valores base")

        fob = st.number_input(
            "Valor FOB (BRL)",
            value=float(st.session_state.get("pdf_map", {}).get("FOB", 0.0)),
            format="%.2f",
        )

        frete = st.number_input(
            "Valor do Frete (BRL)",
            value=float(st.session_state.get("pdf_map", {}).get("FRETE", 0.0)),
            format="%.2f",
        )

        st.divider()
        st.subheader("Buscar itens no PDF")

        busca = st.text_input("Localizar item")

        if busca:
            matches = {
                k: v
                for k, v in st.session_state.get("pdf_map", {}).items()
                if DataEngine.normalizar(busca) in DataEngine.normalizar(k)
            }

            if matches:
                for n, v in matches.items():
                    st.markdown(f"**{n}**")
                    grp_escolhido = st.selectbox(
                        f"Grupo para {n}",
                        ["Impostos", "Taxas", "Despesas"],
                        key=f"grp_{n}",
                        index=2,
                    )
                    if st.button(f"Adicionar {n}", key=f"btn_{n}", use_container_width=True):
                        nova_linha = pd.DataFrame(
                            [
                                {
                                    "Incluir?": True,
                                    "Item": n,
                                    "Valor (R$)": float(v),
                                    "Grupo": grp_escolhido,
                                }
                            ]
                        )
                        st.session_state.df_final = pd.concat(
                            [st.session_state.df_final, nova_linha],
                            ignore_index=True,
                        )
                        st.rerun()
            else:
                st.caption("Nenhum item encontrado para esse termo.")

    df_sel = st.session_state.df_final[st.session_state.df_final["Incluir?"] == True]
    total = df_sel["Valor (R$)"].sum()

    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">Total selecionado</div>
            <div class="metric-value">R$ {total:,.2f}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.button("Gerar PDF selecionado", use_container_width=True):
        if df_sel.empty:
            st.warning("Selecione ao menos um item.")
        else:
            pdf_gen = RelatorioImg(orientation="L")
            arquivo_bytes = pdf_gen.criar_layout(df_sel, fob, frete)

            data_hora = datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%d_%m_%Y_%Hh%M")
            nome_arquivo = f"Estimativa_Energy_{data_hora}.pdf"

            st.download_button(
                label=f"Baixar PDF ({nome_arquivo})",
                data=arquivo_bytes,
                file_name=nome_arquivo,
                mime="application/pdf",
                use_container_width=True,
            )
else:
    st.info("Envie um PDF para carregar automaticamente os itens da estimativa.")


