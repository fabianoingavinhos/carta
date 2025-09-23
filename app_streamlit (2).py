
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
app_streamlit_final_v3_persist2.py

Novidades:
- "Sugestões Salvas": ao selecionar uma sugestão, ela é CARREGADA automaticamente,
  a relação de itens aparece abaixo e você pode incluir novos itens e salvar MESCLANDO.
- "Preço de venda" agora é garantido: preco_de_venda = preco_base * fator
  com parsing robusto (vírgula decimal) e fator zerado/NaN substituído pelo fator global.
"""

import os
import io
from datetime import datetime

import streamlit as st
import pandas as pd
from PIL import Image

# --- PDF (ReportLab) ---
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

# --- Excel (openpyxl) ---
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage

# --- Constantes e diretórios ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMAGEM_DIR = os.path.join(BASE_DIR, "imagens")
SUGESTOES_DIR = os.path.join(BASE_DIR, "sugestoes")
CARTA_DIR = os.path.join(BASE_DIR, "CARTA")
LOGO_PADRAO = os.path.join(CARTA_DIR, "logo_inga.png")

TIPO_ORDEM_FIXA = [
    "Espumantes", "Brancos", "Rosés", "Tintos",
    "Frisantes", "Fortificados", "Vinhos de sobremesa", "Licorosos"
]

# ===== Helpers =====
def garantir_pastas():
    for p in (IMAGEM_DIR, SUGESTOES_DIR, CARTA_DIR):
        os.makedirs(p, exist_ok=True)

def parse_money_series(s, default=0.0):
    """Converte série textual com possível separador de milhar '.' e decimal ',' em float."""
    s = s.astype(str).str.replace("\u00A0", "", regex=False).str.strip()
    s = s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(default)

def to_float_series(s, default=0.0):
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").fillna(default)
    # tenta parse "1.234,56" também
    try:
        return parse_money_series(s, default=default)
    except Exception:
        return pd.to_numeric(s, errors="coerce").fillna(default)

def ler_excel_vinhos(caminho="vinhos1.xls"):
    _, ext = os.path.splitext(caminho.lower())
    engine = None
    if ext == ".xls":
        engine = "xlrd"
    elif ext in (".xlsx", ".xlsm"):
        engine = "openpyxl"
    try:
        df = pd.read_excel(caminho, engine=engine)
    except ImportError:
        st.error("Para ler .xls instale xlrd>=2.0.1, ou converta para .xlsx (openpyxl).")
        raise
    except Exception:
        df = pd.read_excel(caminho)
    df.columns = [c.strip().lower() for c in df.columns]
    if "idx" not in df.columns or df["idx"].isna().all():
        df = df.reset_index(drop=False).rename(columns={"index": "idx"})
    # normaliza tipos
    df["idx"] = pd.to_numeric(df["idx"], errors="coerce").fillna(-1).astype(int)

    # preços e fator: aceitar vírgula
    for col in ["preco38","preco39","preco1","preco2","preco15","preco55","preco63","preco_base","fator","preco_de_venda"]:
        if col not in df.columns:
            df[col] = 0.0
        else:
            df[col] = to_float_series(df[col], default=0.0)

    # textos
    for col in ["cod","descricao","pais","regiao","tipo","uva1","uva2","uva3","amadurecimento","vinicola","corpo","visual","olfato","gustativo","premiacoes"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].astype(str)
    return df

def get_imagem_file(cod: str):
    caminho_win = os.path.join(r"C:/carta/imagens", f"{cod}.png")
    if os.path.exists(caminho_win):
        return caminho_win
    for ext in ['.png', '.jpg', '.jpeg', '.PNG', '.JPG', '.JPEG']:
        img_path = os.path.join(IMAGEM_DIR, f"{cod}{ext}")
        if os.path.exists(img_path):
            return os.path.abspath(img_path)
    try:
        for fname in os.listdir(IMAGEM_DIR):
            if fname.startswith(str(cod)):
                return os.path.abspath(os.path.join(IMAGEM_DIR, fname))
    except Exception:
        pass
    return None

def atualiza_coluna_preco_base(df: pd.DataFrame, flag: str, fator_global: float):
    # define preco_base pela flag escolhida
    base = df[flag] if flag in df.columns else df.get("preco1", 0.0)
    df["preco_base"] = to_float_series(base, default=0.0)
    # fator: se NaN/<=0, usa fator_global
    if "fator" not in df.columns:
        df["fator"] = fator_global
    df["fator"] = to_float_series(df["fator"], default=fator_global)
    df["fator"] = df["fator"].apply(lambda x: fator_global if pd.isna(x) or x <= 0 else x)
    # preco_de_venda = preco_base * fator
    df["preco_de_venda"] = (df["preco_base"].astype(float) * df["fator"].astype(float)).astype(float)
    return df

def ordenar_para_saida(df):
    def normaliza_tipo(t):
        t = str(t).strip().lower()
        if "espum" in t: return "Espumantes"
        if "branc" in t: return "Brancos"
        if "ros" in t: return "Rosés"
        if "tint" in t: return "Tintos"
        if "fris" in t: return "Frisantes"
        if "forti" in t: return "Fortificados"
        if "sobrem" in t: return "Vinhos de sobremesa"
        if "licor" in t: return "Licorosos"
        return t.title()
    tipos_norm = df.get("tipo", pd.Series([""]*len(df))).astype(str).map(normaliza_tipo)
    ordem_map = {t: i for i, t in enumerate(TIPO_ORDEM_FIXA)}
    ordem = tipos_norm.map(lambda x: ordem_map.get(x, 999))
    df2 = df.copy()
    df2["__tipo_ordem"] = ordem
    cols_exist = [c for c in ["__tipo_ordem","pais","descricao"] if c in df2.columns]
    return df2.sort_values(cols_exist).drop(columns=["__tipo_ordem"], errors="ignore")

def add_pdf_footer(c, contagem, total_rotulos, fator_geral):
    from reportlab.lib.pagesizes import A4
    width, height = A4
    y_rodape = 35
    now = datetime.now().strftime("%d/%m/%Y %H:%M")
    c.setLineWidth(0.4)
    c.line(30, y_rodape+32, width-30, y_rodape+32)
    c.setFont("Helvetica", 5)
    c.drawString(32, y_rodape+20, f"Gerado em: {now}")
    try:
        fator_str = f"{float(fator_geral):.2f}"
    except Exception:
        fator_str = str(fator_geral)
    c.setFont("Helvetica-Bold", 6)
    c.drawString(32, y_rodape+7,
        f"Brancos: {contagem.get('Brancos',0)} | Tintos: {contagem.get('Tintos',0)} | "
        f"Rosés: {contagem.get('Rosés',0)} | Espumantes: {contagem.get('Espumantes',0)} | "
        f"Total: {int(total_rotulos)} | Fator: {fator_str}")
    c.setFont("Helvetica", 5)
    c.drawString(32, y_rodape-5, "Ingá Distribuidora Ltda | CNPJ 05.390.477/0002-25 Rod BR 232, KM 18,5 - S/N- Manassu - CEP 54130-340 Jaboatão")
    c.setFont("Helvetica-Bold", 6)
    c.drawString(width-190, y_rodape-5, "b2b.ingavinhos.com.br")

def gerar_pdf(df, titulo, cliente, inserir_foto, logo_cliente_bytes=None):
    from reportlab.lib.pagesizes import A4
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    if logo_cliente_bytes:
        try:
            c.drawImage(ImageReader(io.BytesIO(logo_cliente_bytes)), 40, height-60, width=120, height=40, mask='auto')
        except Exception:
            pass
    if os.path.exists(LOGO_PADRAO):
        try:
            c.drawImage(LOGO_PADRAO, width-80, height-40, width=48, height=24, mask='auto')
        except Exception:
            pass

    x_texto = 90
    y = height - 40
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(width/2, y, titulo)
    y -= 20
    if cliente:
        c.setFont("Helvetica", 10)
        c.drawCentredString(width/2, y, f"Cliente: {cliente}")
        y -= 20

    ordem_geral = 1
    contagem = {'Brancos':0, 'Tintos':0, 'Rosés':0, 'Espumantes':0, 'outros':0}
    df_sorted = ordenar_para_saida(df)

    for tipo in df_sorted['tipo'].fillna("").astype(str).unique():
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x_texto, y, str(tipo).upper()); y -= 14
        for pais in df_sorted[df_sorted['tipo']==tipo]['pais'].dropna().unique():
            c.setFont("Helvetica-Bold", 8)
            c.drawString(x_texto, y, str(pais).upper()); y -= 12
            grupo = df_sorted[(df_sorted['tipo']==tipo) & (df_sorted['pais']==pais)]
            for _, row in grupo.iterrows():
                t = str(tipo).lower()
                if "branc" in t: contagem["Brancos"] += 1
                elif "tint" in t: contagem["Tintos"] += 1
                elif "ros" in t: contagem["Rosés"] += 1
                elif "espum" in t: contagem["Espumantes"] += 1
                else: contagem["outros"] += 1

                c.setFont("Helvetica", 6)
                try: codtxt = f"{int(row['cod'])}"
                except Exception: codtxt = str(row.get('cod',""))
                c.drawString(x_texto, y, f"{ordem_geral:02d} ({codtxt})")
                c.setFont("Helvetica-Bold", 7)
                c.drawString(x_texto+55, y, str(row['descricao']))
                uvas = [str(row.get(f"uva{i}", "")).strip() for i in range(1,4)]
                uvas = [u for u in uvas if u and u.lower() != "nan"]
                regiao_str = f"{row.get('pais','')} | {row.get('regiao','')}"
                if uvas: regiao_str += f" | {', '.join(uvas)}"
                c.setFont("Helvetica", 5); c.drawString(x_texto+55, y-10, regiao_str)

                amad = str(row.get("amadurecimento", ""))
                if amad and amad.lower() != "nan":
                    c.setFont("Helvetica", 7); c.drawString(220, y-7, "🛢️")

                c.setFont("Helvetica", 5)
                try: c.drawRightString(width-120, y, f"(R$ {float(row['preco_base']):.2f})")
                except Exception: c.drawRightString(width-120, y, "(R$ -)")
                c.setFont("Helvetica-Bold", 7)
                try: c.drawRightString(width-40, y, f"R$ {float(row['preco_de_venda']):.2f}")
                except Exception: c.drawRightString(width-40, y, "R$ -")

                if inserir_foto:
                    imgfile = get_imagem_file(str(row.get('cod','')))
                    if imgfile:
                        try:
                            c.drawImage(imgfile, x_texto+340, y-2, width=40, height=30, mask='auto'); y -= 28
                        except Exception: y -= 20
                    else:
                        y -= 20
                else:
                    y -= 20

                ordem_geral += 1

                if y < 100:
                    add_pdf_footer(c, contagem, ordem_geral-1, fator_geral=df.get('fator', pd.Series([0])).median())
                    c.showPage()
                    y = height - 40
                    if logo_cliente_bytes:
                        try: c.drawImage(ImageReader(io.BytesIO(logo_cliente_bytes)), 40, height-60, width=120, height=40, mask='auto')
                        except Exception: pass
                    if os.path.exists(LOGO_PADRAO):
                        try: c.drawImage(LOGO_PADRAO, width-80, height-40, width=48, height=24, mask='auto')
                        except Exception: pass
                    c.setFont("Helvetica-Bold", 16); c.drawCentredString(width/2, y, titulo); y -= 20
                    if cliente: c.setFont("Helvetica", 10); c.drawCentredString(width/2, y, f"Cliente: {cliente}"); y -= 20

    add_pdf_footer(c, contagem, ordem_geral-1, fator_geral=df.get('fator', pd.Series([0])).median())
    c.save(); buffer.seek(0)
    return buffer

def exportar_excel_like_pdf(df, inserir_foto=True):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sugestão"
    row_num = 1; ordem_geral = 1
    df_sorted = ordenar_para_saida(df)
    for tipo in df_sorted['tipo'].fillna("").unique():
        ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=8)
        cell = ws.cell(row=row_num, column=1, value=str(tipo).upper()); cell.font = Font(bold=True, size=18); row_num += 1
        for pais in df_sorted[df_sorted['tipo'] == tipo]['pais'].dropna().unique():
            ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=8)
            cell = ws.cell(row=row_num, column=1, value=str(pais).upper()); cell.font = Font(bold=True, size=14); row_num += 1
            grupo = df_sorted[(df_sorted['tipo'] == tipo) & (df_sorted['pais'] == pais)]
            for _, row in grupo.iterrows():
                ws.cell(row=row_num, column=1, value=f"{ordem_geral:02d} ({int(row['cod']) if str(row['cod']).isdigit() else ''})").font = Font(size=11)
                ws.cell(row=row_num, column=2, value=str(row['descricao'])).font = Font(bold=True, size=12)
                if inserir_foto:
                    imgfile = get_imagem_file(str(row.get('cod','')))
                    if imgfile and os.path.exists(imgfile):
                        try:
                            img = XLImage(imgfile); img.width, img.height = 32, 24; ws.add_image(img, f"C{row_num}")
                        except Exception: pass
                try:
                    base_val = float(row['preco_base']); pv_val = float(row['preco_de_venda'])
                    base_str = f"(R$ {base_val:.2f})"; pv_str = f"R$ {pv_val:.2f}"
                except Exception:
                    base_str = "(R$ -)"; pv_str = "R$ -"
                ws.cell(row=row_num, column=7, value=base_str).alignment = Alignment(horizontal='right'); ws.cell(row=row_num, column=7).font = Font(size=10)
                ws.cell(row=row_num, column=8, value=pv_str).font = Font(bold=True, size=13); ws.cell(row=row_num, column=8).alignment = Alignment(horizontal='right')
                uvas = [str(row.get(f"uva{i}", "")).strip() for i in range(1,4)]; uvas = [u for u in uvas if u and u.lower() != "nan"]
                regiao_str = f"{row.get('pais','')} | {row.get('regiao','')}"; 
                if uvas: regiao_str += f" | {', '.join(uvas)}"
                ws.cell(row=row_num+1, column=2, value=regiao_str).font = Font(size=10)
                amad = str(row.get("amadurecimento", ""))
                if amad and amad.lower() != "nan":
                    ws.cell(row=row_num+1, column=3, value="🛢️").font = Font(size=10)
                row_num += 2; ordem_geral += 1
    stream = io.BytesIO(); wb.save(stream); stream.seek(0); return stream

# ===================== APP =====================
def main():
    st.set_page_config(page_title="Sugestão de Carta de Vinhos", layout="wide")
    garantir_pastas()

    # Estado
    if "selected_idxs" not in st.session_state:
        st.session_state.selected_idxs = set()
    if "prev_view_state" not in st.session_state:
        st.session_state.prev_view_state = {}
    if "manual_fat" not in st.session_state:
        st.session_state.manual_fat = {}
    if "manual_preco_venda" not in st.session_state:
        st.session_state.manual_preco_venda = {}
    if "cadastrados" not in st.session_state:
        st.session_state.cadastrados = []

    st.markdown("### Sugestão de Carta de Vinhos")

    with st.container():
        c1, c2, c3, c4, c5, c6, c7, c8 = st.columns([1.4,1.2,1,1,1.6,0.9,1.2,1.6])
        with c1:
            cliente = st.text_input("Nome do Cliente", value="", placeholder="(opcional)", key="cliente_nome")
        with c2:
            logo_cliente = st.file_uploader("Carregar logo (cliente)", type=["png","jpg","jpeg"], key="logo_cliente")
            logo_bytes = logo_cliente.read() if logo_cliente else None
        with c3:
            inserir_foto = st.checkbox("Inserir foto no PDF/Excel", value=True, key="chk_foto")
        with c4:
            preco_flag = st.selectbox("Tabela de preço",
                                      ["preco1", "preco2", "preco15", "preco38", "preco39", "preco55", "preco63"],
                                      index=0, key="preco_flag")
        with c5:
            termo_global = st.text_input("Buscar", value="", key="termo_global")
        with c6:
            fator_global = st.number_input("Fator", min_value=0.0, value=2.0, step=0.1, key="fator_global_input")
        with c7:
            resetar = st.button("Resetar/Mostrar Todos", key="btn_resetar")
        with c8:
            caminho_planilha = st.text_input("Arquivo de dados", value="vinhos1.xls",
                                             help="Caminho do arquivo XLS/XLSX (ex.: vinhos1.xls)",
                                             key="caminho_planilha")

    # Carrega DF base
    df = ler_excel_vinhos(caminho_planilha)
    df = atualiza_coluna_preco_base(df, preco_flag, fator_global=float(fator_global))

    # Integra itens cadastrados (sessão)
    if st.session_state.cadastrados:
        cad_df = pd.DataFrame(st.session_state.cadastrados)
        for col in df.columns:
            if col not in cad_df.columns:
                cad_df[col] = None
        cad_df["idx"] = pd.to_numeric(cad_df["idx"], errors="coerce").fillna(-1).astype(int)
        df = pd.concat([df, cad_df[df.columns]], ignore_index=True)

    # Sidebar de filtros
    st.sidebar.header("Filtros")
    pais_opc = [""] + sorted([p for p in df["pais"].dropna().astype(str).unique().tolist() if p])
    tipo_opc = [""] + sorted([t for t in df["tipo"].dropna().astype(str).unique().tolist() if t])
    desc_opc = [""] + sorted([d for d in df["descricao"].dropna().astype(str).unique().tolist() if d])
    regiao_opc = [""] + sorted([r for r in df["regiao"].dropna().astype(str).unique().tolist() if r])
    cod_opc = [""] + sorted([str(c) for c in df["cod"].dropna().astype(str).unique().tolist()])

    filt_pais = st.sidebar.selectbox("País", pais_opc, index=0, key="filt_pais")
    filt_tipo = st.sidebar.selectbox("Tipo", tipo_opc, index=0, key="filt_tipo")
    filt_desc = st.sidebar.selectbox("Descrição", desc_opc, index=0, key="filt_desc")
    filt_regiao = st.sidebar.selectbox("Região", regiao_opc, index=0, key="filt_regiao")
    filt_cod = st.sidebar.selectbox("Código", cod_opc, index=0, key="filt_cod")

    colp1, colp2 = st.sidebar.columns(2)
    with colp1:
        preco_min = st.number_input("Preço mín (base)", min_value=0.0, value=0.0, step=1.0, key="preco_min")
    with colp2:
        preco_max = st.number_input("Preço máx (base)", min_value=0.0, value=0.0, step=1.0, help="0 = sem limite", key="preco_max")

    # Aplicar filtros (VIEW)
    df_filtrado = df.copy()
    if termo_global.strip():
        term = termo_global.strip().lower()
        mask = df_filtrado.apply(lambda row: term in " ".join(str(v).lower() for v in row.values), axis=1)
        df_filtrado = df_filtrado[mask]
    if filt_pais:
        df_filtrado = df_filtrado[df_filtrado["pais"] == filt_pais]
    if filt_tipo:
        df_filtrado = df_filtrado[df_filtrado["tipo"] == filt_tipo]
    if filt_desc:
        df_filtrado = df_filtrado[df_filtrado["descricao"] == filt_desc]
    if filt_regiao:
        df_filtrado = df_filtrado[df_filtrado["regiao"] == filt_regiao]
    if filt_cod:
        df_filtrado = df_filtrado[df_filtrado["cod"].astype(str) == filt_cod]
    if preco_min:
        df_filtrado = df_filtrado[df_filtrado["preco_base"].fillna(0) >= float(preco_min)]
    if preco_max and preco_max > 0:
        df_filtrado = df_filtrado[df_filtrado["preco_base"].fillna(0) <= float(preco_max)]

    if resetar:
        df_filtrado = df.copy()

    # Contagem por tipo + status seleção
    contagem = {'Brancos': 0, 'Tintos': 0, 'Rosés': 0, 'Espumantes': 0, 'outros': 0}
    for t, n in df_filtrado.groupby('tipo').size().items():
        t_low = str(t).lower()
        if "branc" in t_low: contagem['Brancos'] += int(n)
        elif "tint" in t_low: contagem['Tintos'] += int(n)
        elif "ros" in t_low: contagem['Rosés'] += int(n)
        elif "espum" in t_low: contagem['Espumantes'] += int(n)
        else: contagem['outros'] += int(n)
    total = len(df_filtrado)
    selecionados = len(st.session_state.selected_idxs)
    st.caption(f"Brancos: {contagem.get('Brancos', 0)} | Tintos: {contagem.get('Tintos', 0)} | Rosés: {contagem.get('Rosés', 0)} | Espumantes: {contagem.get('Espumantes', 0)} | Total: {total} | Selecionados: {selecionados} | Fator: {float(fator_global):.2f}")

    # === Grade com seleção ===
    view_df = df_filtrado.copy()

    # --- Normalização robusta + remoção de colunas duplicadas ---
    if not isinstance(view_df, pd.DataFrame):
        view_df = pd.DataFrame(view_df)
    try:
        view_df = view_df.loc[:, ~view_df.columns.duplicated()].copy()
    except Exception:
        pass
    if "idx" not in view_df.columns:
        view_df = view_df.reset_index(drop=False).rename(columns={"index": "idx"})
    _idx_col = view_df["idx"]
    if isinstance(_idx_col, pd.DataFrame):
        _idx_col = _idx_col.iloc[:, 0]
    view_df["idx"] = pd.to_numeric(_idx_col, errors="coerce").fillna(-1).astype(int)
    if "cod" in view_df.columns:
        _cod_col = view_df["cod"]
        if isinstance(_cod_col, pd.DataFrame):
            _cod_col = _cod_col.iloc[:, 0]
        view_df["cod"] = _cod_col.astype(str)
    else:
        view_df["cod"] = ""
    for _c in ["preco_base", "preco_de_venda", "fator"]:
        if _c in view_df.columns:
            _col = view_df[_c]
            if isinstance(_col, pd.DataFrame):
                _col = _col.iloc[:, 0]
            view_df[_c] = to_float_series(_col, default=0.0)
        else:
            view_df[_c] = 0.0

    view_df["selecionado"] = view_df["idx"].apply(lambda i: i in st.session_state.selected_idxs)
    view_df["foto"] = view_df["cod"].apply(lambda c: "●" if get_imagem_file(str(c)) else "")

    edited = st.data_editor(
        view_df[["selecionado","foto","cod","descricao","pais","regiao","preco_base","preco_de_venda","fator","idx"]],
        hide_index=True,
        column_config={
            "selecionado": st.column_config.CheckboxColumn("SELECIONADO"),
            "foto": st.column_config.TextColumn("FOTO"),
            "cod": st.column_config.TextColumn("COD"),
            "descricao": st.column_config.TextColumn("DESCRICAO"),
            "pais": st.column_config.TextColumn("PAIS"),
            "regiao": st.column_config.TextColumn("REGIAO"),
            "preco_base": st.column_config.NumberColumn("PRECO_BASE", format="R$ %.2f", step=0.01),
            "preco_de_venda": st.column_config.NumberColumn("PRECO_VENDA", format="R$ %.2f", step=0.01),
            "fator": st.column_config.NumberColumn("FATOR", format="%.2f", step=0.1),
            "idx": st.column_config.NumberColumn("IDX", help="Identificador interno"),
        },
        use_container_width=True,
        num_rows="dynamic",
        key="editor_main",
    )

    # --- Persistência incremental das seleções ---
    curr_state = {}
    if isinstance(edited, pd.DataFrame) and not edited.empty:
        for _, row in edited.iterrows():
            try:
                idx_i = int(row["idx"])
            except Exception:
                continue
            sel = bool(row.get("selecionado", False))
            curr_state[idx_i] = sel

    prev_state = st.session_state.get("prev_view_state", {})
    global_sel = set(st.session_state.selected_idxs)

    to_add = {i for i, s in curr_state.items() if s and prev_state.get(i) is not True}
    to_remove = {i for i, s in curr_state.items() if (prev_state.get(i) is True) and not s}

    global_sel |= to_add
    global_sel -= to_remove

    st.session_state.selected_idxs = global_sel
    st.session_state.prev_view_state = curr_state

    # Ajustes manuais (aplicados no DF base, por idx) + recomputa preco_de_venda
    if isinstance(edited, pd.DataFrame) and not edited.empty:
        for _, r in edited.iterrows():
            try:
                idx = int(r["idx"])
            except Exception:
                continue
            if pd.notnull(r.get("fator")):
                st.session_state.manual_fat[idx] = float(r["fator"])
            if pd.notnull(r.get("preco_de_venda")):
                st.session_state.manual_preco_venda[idx] = float(r["preco_de_venda"])

    for idx, fat in st.session_state.manual_fat.items():
        df.loc[df["idx"]==idx, "fator"] = float(fat)
    # caso fator <=0, usa fator_global
    df["fator"] = to_float_series(df["fator"], default=float(fator_global))
    df["fator"] = df["fator"].apply(lambda x: float(fator_global) if pd.isna(x) or x <= 0 else float(x))

    df["preco_base"] = to_float_series(df["preco_base"], default=0.0)
    df["preco_de_venda"] = (df["preco_base"].astype(float) * df["fator"].astype(float)).astype(float)

    for idx, pv in st.session_state.manual_preco_venda.items():
        df.loc[df["idx"]==idx, "preco_de_venda"] = float(pv)

    # Botões de ação + salvar sugestão
    cA, cB, cC, cD, cE, cF = st.columns([1,1.2,1.2,1.2,1.6,1.2])
    with cA:
        ver_preview = st.button("Visualizar Sugestão", key="btn_preview")
    with cB:
        ver_marcados = st.button("Visualizar Itens Marcados", key="btn_marcados")
    with cC:
        gerar_pdf_btn = st.button("Gerar PDF", key="btn_pdf")
    with cD:
        exportar_excel_btn = st.button("Exportar para Excel", key="btn_excel")
    with cE:
        nome_sugestao = st.text_input("Nome da sugestão", value="", key="nome_sugestao_input")
    with cF:
        salvar_sugestao_btn = st.button("Salvar Sugestão (mesclar se existir)", key="btn_salvar")

    if ver_preview:
        if not st.session_state.selected_idxs:
            st.info("Nenhum item selecionado.")
        else:
            st.subheader("Pré-visualização da Sugestão")
            df_sel = df[df["idx"].isin(st.session_state.selected_idxs)].copy()
            df_sel = ordenar_para_saida(df_sel)
            preview_lines = []
            preview_lines.append("Sugestão Carta de Vinhos")
            if cliente:
                preview_lines.append(f"Cliente: {cliente}")
            preview_lines.append("="*70)
            ordem_geral = 1
            for tipo in df_sel['tipo'].dropna().unique():
                preview_lines.append(f"\n{str(tipo).upper()}")
                for pais in df_sel[df_sel['tipo']==tipo]['pais'].dropna().unique():
                    preview_lines.append(f"  {str(pais).upper()}")
                    grupo = df_sel[(df_sel['tipo']==tipo) & (df_sel['pais']==pais)]
                    for _, row in grupo.iterrows():
                        desc = row['descricao']
                        try:
                            preco = f"R$ {float(row['preco_base']):.2f}"
                            pvenda = f"R$ {float(row['preco_de_venda']):.2f}"
                        except Exception:
                            preco = "R$ -"; pvenda = "R$ -"
                        try: cod = int(row['cod']) if str(row['cod']).isdigit() else ""
                        except Exception: cod = str(row.get('cod',""))
                        regiao = row.get('regiao',"")
                        preview_lines.append(f"    {ordem_geral:02d} ({cod}) {desc}")
                        uvas = [str(row.get(f"uva{i}", "")).strip() for i in range(1,4)]
                        uvas_str = ", ".join([u for u in uvas if u and u.lower()!='nan'])
                        linha2 = f"      {row.get('pais','')} | {regiao}"
                        if uvas_str: linha2 += f" | {uvas_str}"
                        preview_lines.append(linha2)
                        preview_lines.append(f"      ({preco})  {pvenda}")
                        if inserir_foto and get_imagem_file(str(row.get('cod',''))):
                            preview_lines.append("      [COM FOTO]")
                        ordem_geral += 1
            preview_lines.append("\n" + "="*70)
            now = datetime.now().strftime("%d/%m/%Y %H:%M")
            preview_lines.append(f"Gerado em: {now}")
            st.code("\n".join(preview_lines))

    if ver_marcados:
        if not st.session_state.selected_idxs:
            st.info("Nenhum item selecionado.")
        else:
            st.subheader("Itens Marcados")
            df_sel = df[df["idx"].isin(st.session_state.selected_idxs)].copy()
            df_sel = df_sel[["cod","descricao","pais","regiao","preco_base","preco_de_venda","fator"]].sort_values(["pais","descricao"])
            st.dataframe(df_sel, use_container_width=True)

    if gerar_pdf_btn:
        if not st.session_state.selected_idxs:
            st.warning("Selecione ao menos um vinho.")
        else:
            df_sel = df[df["idx"].isin(st.session_state.selected_idxs)].copy()
            df_sel = ordenar_para_saida(df_sel)
            pdf_buffer = gerar_pdf(df_sel, "Sugestão Carta de Vinhos", cliente, inserir_foto, logo_bytes)
            st.download_button("Baixar PDF", data=pdf_buffer, file_name="sugestao_carta_vinhos.pdf", mime="application/pdf", key="dl_pdf")

    if exportar_excel_btn:
        if not st.session_state.selected_idxs:
            st.warning("Selecione ao menos um vinho.")
        else:
            df_sel = df[df["idx"].isin(st.session_state.selected_idxs)].copy()
            df_sel = ordenar_para_saida(df_sel)
            xlsx = exportar_excel_like_pdf(df_sel, inserir_foto=inserir_foto)
            st.download_button("Baixar Excel", data=xlsx, file_name="sugestao_carta_vinhos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_xlsx")

    if salvar_sugestao_btn:
        garantir_pastas()
        nome = nome_sugestao.strip()
        if not nome:
            st.warning("Informe um nome para a sugestão antes de salvar.")
        elif not st.session_state.selected_idxs:
            st.info("Selecione produtos para salvar.")
        else:
            path = os.path.join(SUGESTOES_DIR, f"{nome}.txt")
            new_set = set(st.session_state.selected_idxs)
            if os.path.exists(path):
                try:
                    with open(path) as f:
                        old = [int(x) for x in f.read().strip().split(",") if x]
                    new_set |= set(old)
                except Exception:
                    pass
            try:
                with open(path, "w") as f:
                    f.write(",".join(map(str, sorted(list(new_set)))))
                st.success(f"Sugestão '{nome}' salva (mesclada) em {path}.")
            except Exception as e:
                st.error(f"Erro ao salvar: {e}")

    # Abas
    st.markdown("---")
    tab1, tab2 = st.tabs(["Sugestões Salvas", "Cadastro de Vinhos"])

    with tab1:
        garantir_pastas()
        arquivos = [f for f in os.listdir(SUGESTOES_DIR) if f.endswith(".txt")]
        sel = st.selectbox("Abrir sugestão", [""] + [a[:-4] for a in arquivos], key="sel_sugestao")

        # Ao selecionar, carregar automaticamente e mostrar a RELAÇÃO abaixo
        sugestao_indices = []
        if sel:
            path = os.path.join(SUGESTOES_DIR, f"{sel}.txt")
            if os.path.exists(path):
                try:
                    with open(path) as f:
                        sugestao_indices = [int(x) for x in f.read().strip().split(",") if x]
                    # Carrega a sugestão (substitui seleção atual)
                    st.session_state.selected_idxs = set(sugestao_indices)
                    st.info(f"Sugestão '{sel}' carregada: {len(sugestao_indices)} itens.")
                except Exception as e:
                    st.error(f"Erro ao carregar '{sel}': {e}")

        # Relação da sugestão (abaixo da seleção)
        if sugestao_indices:
            st.subheader("Relação da Sugestão")
            df_rel = df[df["idx"].isin(sugestao_indices)].copy()
            if not df_rel.empty:
                df_rel = df_rel[["cod","descricao","pais","regiao","preco_base","fator","preco_de_venda"]].sort_values(["pais","descricao"])
                st.dataframe(df_rel, use_container_width=True, height=min(500, 50 + 28*len(df_rel)))
            else:
                st.caption("Nenhum item encontrado no DF atual para esses índices.")

        colx, coly, colz = st.columns([1,1,1])
        with colx:
            if st.button("Excluir sugestão selecionada", key="btn_excluir_sug"):
                if sel:
                    try:
                        os.remove(os.path.join(SUGESTOES_DIR, f"{sel}.txt"))
                        st.success(f"Sugestão '{sel}' excluída.")
                        st.experimental_rerun()
                    except Exception as e:
                        st.error(f"Erro ao excluir: {e}")
                else:
                    st.info("Selecione uma sugestão na lista.")
        with coly:
            if st.button("Salvar alterações nesta sugestão (mesclar)", key="btn_merge_sug"):
                if sel:
                    path = os.path.join(SUGESTOES_DIR, f"{sel}.txt")
                    try:
                        old = []
                        if os.path.exists(path):
                            with open(path) as f:
                                old = [int(x) for x in f.read().strip().split(",") if x]
                        new_set = set(old) | set(st.session_state.selected_idxs)
                        with open(path, "w") as f:
                            f.write(",".join(map(str, sorted(list(new_set)))))
                        st.success(f"Sugestão '{sel}' atualizada (itens mesclados).")
                    except Exception as e:
                        st.error(f"Erro ao salvar: {e}")
                else:
                    st.info("Selecione uma sugestão na lista.")
        with colz:
            if st.button("Limpar seleção atual", key="btn_limpar_sel"):
                st.session_state.selected_idxs = set()
                st.experimental_rerun()

    with tab2:
        st.caption("Cadastrar novo produto (entra apenas na sessão atual; salve no seu Excel depois, se quiser persistir).")
        c1b, c2b, c3b, c4b, c5b, c6b, c7b = st.columns([1,2,1,1,1,1,1.2])
        with c1b:
            new_cod = st.text_input("Código", key="cad_cod")
        with c2b:
            new_desc = st.text_input("Descrição", key="cad_desc")
        with c3b:
            new_preco = st.number_input("Preço", min_value=0.0, value=0.0, step=0.01, key="cad_preco")
        with c4b:
            new_fat = st.number_input("Fator", min_value=0.0, value=float(fator_global), step=0.1, key="cad_fator")
        with c5b:
            new_pv = st.number_input("Preço Venda", min_value=0.0, value=0.0, step=0.01, key="cad_pv")
        with c6b:
            new_pais = st.text_input("País", key="cad_pais")
        with c7b:
            new_regiao = st.text_input("Região", key="cad_regiao")

        if st.button("Cadastrar", key="btn_cadastrar"):
            try:
                cod_int = int(float(new_cod)) if new_cod else None
                pv_calc = new_pv if new_pv > 0 else new_preco * new_fat
                # idx único baseado no maior existente
                idx_next = 0
                if "idx" in df.columns and not df["idx"].isna().all():
                    try:
                        idx_next = int(pd.to_numeric(df["idx"], errors="coerce").max()) + 1
                    except Exception:
                        idx_next = len(df) + 1
                novo = {
                    "idx": idx_next,
                    "cod": cod_int if cod_int is not None else "",
                    "descricao": new_desc,
                    "preco_base": float(new_preco),
                    "fator": float(new_fat),
                    "preco_de_venda": float(pv_calc),
                    "pais": new_pais,
                    "regiao": new_regiao,
                    "tipo": "",
                }
                st.session_state.cadastrados.append(novo)
                st.success("Produto cadastrado na sessão atual. Ele já aparece na grade após o recarregamento.")
                st.experimental_rerun()
            except Exception as e:
                st.error(f"Erro ao cadastrar: {e}")

if __name__ == "__main__":
    main()
