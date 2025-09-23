
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
app_streamlit.py
Conversão da aplicação Tkinter para Streamlit mantendo:
- Filtros (global, país, tipo, região, código, preço min/máx)
- Escolha de tabela de preços (preco1, preco2, preco15, preco38, preco39, preco55, preco63)
- Fator global e ajustes manuais por item (fator e preço de venda)
- Seleção de itens (marcar/desmarcar; marcar tudo/desmarcar tudo dos filtrados)
- Geração de PDF (layout com logo Ingá no topo direito + logo do cliente opcional + fotos opcionais)
- Exportação Excel com layout semelhante ao PDF
- Salvar/Carregar sugestões (lista de índices selecionados) em /sugestoes/*.txt

Pastas esperadas (iguais ao Tkinter):
- imagens/           (fotos dos produtos, nomeadas por código: 407.png, 123.jpg etc)
- CARTA/logo_inga.png
- sugestoes/

Observações:
- Se for ler .xls você precisa ter xlrd instalado (pip install xlrd>=2.0.1)
- Para .xlsx é usado openpyxl
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

CAMPOS_NOVOS = [
    "cod", "descricao", "visual", "olfato", "gustativo", "premiacoes", "amadurecimento",
    "regiao", "pais", "vinicola", "corpo", "tipo",
    "uva1", "uva2", "uva3",
    "preco38", "preco39", "preco1", "preco2", "preco15", "preco55", "preco63"
]

# --- Utilitários ---
def garantir_pastas():
    for p in (IMAGEM_DIR, SUGESTOES_DIR, CARTA_DIR):
        if not os.path.isdir(p):
            try:
                os.makedirs(p, exist_ok=True)
            except Exception:
                pass

def ler_excel_vinhos(caminho="vinhos1.xls"):
    # Aceita .xls (xlrd) e .xlsx (openpyxl)
    _, ext = os.path.splitext(caminho.lower())
    engine = None
    if ext == ".xls":
        engine = "xlrd"
    elif ext in (".xlsx", ".xlsm"):
        engine = "openpyxl"
    try:
        df = pd.read_excel(caminho, engine=engine)
    except Exception:
        # Tenta sem engine como fallback
        df = pd.read_excel(caminho)
    df.columns = [c.strip().lower() for c in df.columns]
    for col in CAMPOS_NOVOS:
        if col not in df.columns:
            df[col] = ""
    for col in ["preco38", "preco39", "preco1", "preco2", "preco15", "preco55", "preco63"]:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
    # Índice original preservado para seleção
    df = df.reset_index(drop=False).rename(columns={"index": "idx"})
    return df

def get_imagem_file(cod: str):
    # Compatibilidade com C:\carta\imagens\cod.png
    caminho_win = os.path.join(r"C:/carta/imagens", f"{cod}.png")
    if os.path.exists(caminho_win):
        return caminho_win
    # Varre imagens/ por extensões comuns
    for ext in ['.png', '.jpg', '.jpeg', '.PNG', '.JPG', '.JPEG']:
        img_path = os.path.join(IMAGEM_DIR, f"{cod}{ext}")
        if os.path.exists(img_path):
            return os.path.abspath(img_path)
    # Fallback: qualquer arquivo que comece com o código
    try:
        for fname in os.listdir(IMAGEM_DIR):
            if fname.startswith(str(cod)):
                return os.path.abspath(os.path.join(IMAGEM_DIR, fname))
    except Exception:
        pass
    return None

def atualiza_coluna_preco_base(df: pd.DataFrame, flag: str):
    if flag not in df.columns:
        base = df.get("preco1", 0.0)
    else:
        base = df[flag].fillna(0.0)
    df["preco_base"] = pd.to_numeric(base, errors="coerce").fillna(0.0)
    if "fator" not in df.columns:
        df["fator"] = 2.0
    df["fator"] = pd.to_numeric(df["fator"], errors="coerce").fillna(2.0)
    df["preco_de_venda"] = df["preco_base"] * df["fator"]
    return df

def ordenar_para_saida(df):
    # Mantém o comportamento do Tkinter: ordena por tipo, pais, descricao.
    # Caso queira uma ordem fixa posterior, ajuste aqui.
    cols_exist = [c for c in ["tipo","pais","descricao"] if c in df.columns]
    if cols_exist:
        return df.sort_values(cols_exist)
    return df

def gerar_pdf(df, titulo, cliente, inserir_foto, logo_cliente_bytes=None):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    # Logo do cliente (se houver)
    if logo_cliente_bytes:
        try:
            c.drawImage(ImageReader(io.BytesIO(logo_cliente_bytes)), 40, height-60, width=120, height=40, mask='auto')
        except Exception:
            pass

    # Logo Ingá SEMPRE topo direito (se existir)
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
    tipo_map = {'branco':'Brancos', 'tinto':'Tintos', 'rose':'Rosés', 'rosé':'Rosés', 'espumante':'Espumantes'}

    for tipo in df['tipo'].dropna().unique():
        tipo_label = next((lbl for k,lbl in tipo_map.items() if k in str(tipo).lower()), 'outros')
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x_texto, y, str(tipo).upper())
        y -= 14
        for pais in df[df['tipo']==tipo]['pais'].dropna().unique():
            c.setFont("Helvetica-Bold", 8)
            c.drawString(x_texto, y, str(pais).upper())
            y -= 12
            grupo = df[(df['tipo'] == tipo) & (df['pais'] == pais)]
            for _, row in grupo.iterrows():
                contagem[tipo_label] = contagem.get(tipo_label,0) + 1
                c.setFont("Helvetica", 6)
                c.drawString(x_texto, y, f"{ordem_geral:02d} ({int(row['cod']) if pd.notnull(row['cod']) else ''})")
                c.setFont("Helvetica-Bold", 7)
                c.drawString(x_texto+55, y, str(row['descricao']))
                # linha inferior com regiao/uva
                uvas = [str(row.get(f"uva{i}", "")).strip() for i in range(1,4)]
                uvas = [u for u in uvas if u and u.lower() != "nan"]
                regiao_str = f"{row.get('pais','')} | {row.get('regiao','')}"
                if uvas:
                    regiao_str += f" | {', '.join(uvas)}"
                c.setFont("Helvetica", 5)
                c.drawString(x_texto+55, y-10, regiao_str)

                amad = str(row.get("amadurecimento", ""))
                if amad and amad.lower() != "nan":
                    c.setFont("Helvetica", 7)
                    c.drawString(220, y-7, "🛢️")

                # preços
                c.setFont("Helvetica", 5)
                c.drawRightString(width-120, y, f"(R$ {row['preco_base']:.2f})")
                c.setFont("Helvetica-Bold", 7)
                c.drawRightString(width-40, y, f"R$ {row['preco_de_venda']:.2f}")

                # imagem (opcional)
                if inserir_foto:
                    imgfile = get_imagem_file(str(row.get('cod','')))
                    if imgfile:
                        try:
                            c.drawImage(imgfile, x_texto+340, y-2, width=40, height=30, mask='auto')
                            y -= 28
                        except Exception:
                            y -= 20
                    else:
                        y -= 20
                else:
                    y -= 20

                ordem_geral += 1

                # quebra de página
                if y < 100:
                    add_pdf_footer(c, contagem, ordem_geral-1, fator_geral=df.get('fator', pd.Series([0])).median())
                    c.showPage()
                    y = height - 40
                    # cabeçalhos
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
                    c.setFont("Helvetica-Bold", 16)
                    c.drawCentredString(width/2, y, titulo)
                    y -= 20
                    if cliente:
                        c.setFont("Helvetica", 10)
                        c.drawCentredString(width/2, y, f"Cliente: {cliente}")
                        y -= 20

    add_pdf_footer(c, contagem, ordem_geral-1, fator_geral=df.get('fator', pd.Series([0])).median())
    c.save()
    buffer.seek(0)
    return buffer

def add_pdf_footer(c, contagem, total_rotulos, fator_geral):
    width, height = A4
    y_rodape = 35
    now = datetime.now().strftime("%d/%m/%Y %H:%M")
    c.setLineWidth(0.4)
    c.line(30, y_rodape+32, width-30, y_rodape+32)
    c.setFont("Helvetica", 5)
    c.drawString(32, y_rodape+20, f"Gerado em: {now}")

    # garante formatação correta do fator
    if isinstance(fator_geral, (int, float)):
        fator_str = f"{fator_geral:.2f}"
    else:
        fator_str = str(fator_geral)

    c.setFont("Helvetica-Bold", 6)
    c.drawString(
        32,
        y_rodape+7,
        f"Brancos: {contagem.get('Brancos',0)} | Tintos: {contagem.get('Tintos',0)} | "
        f"Rosés: {contagem.get('Rosés',0)} | Espumantes: {contagem.get('Espumantes',0)} | "
        f"Total: {int(total_rotulos)} | Fator: {fator_str}"
    )
    c.setFont("Helvetica", 5)
    c.drawString(32, y_rodape-5, "Ingá Distribuidora Ltda | CNPJ 05.390.477/0002-25 Rod BR 232, KM 18,5 - S/N- Manassu - CEP 54130-340 Jaboatão")
    c.setFont("Helvetica-Bold", 6)
    c.drawString(width-190, y_rodape-5, "b2b.ingavinhos.com.br")


def exportar_excel_like_pdf(df, inserir_foto=True):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sugestão"

    row_num = 1
    ordem_geral = 1

    for tipo in df['tipo'].dropna().unique():
        ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=8)
        cell = ws.cell(row=row_num, column=1, value=str(tipo).upper())
        cell.font = Font(bold=True, size=18)
        row_num += 1

        for pais in df[df['tipo'] == tipo]['pais'].dropna().unique():
            ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=8)
            cell = ws.cell(row=row_num, column=1, value=str(pais).upper())
            cell.font = Font(bold=True, size=14)
            row_num += 1

            grupo = df[(df['tipo'] == tipo) & (df['pais'] == pais)]
            for _, row in grupo.iterrows():
                ws.cell(row=row_num, column=1, value=f"{ordem_geral:02d} ({int(row['cod']) if pd.notnull(row['cod']) else ''})").font = Font(size=11)
                ws.cell(row=row_num, column=2, value=str(row['descricao'])).font = Font(bold=True, size=12)

                if inserir_foto:
                    imgfile = get_imagem_file(str(row.get('cod','')))
                    if imgfile and os.path.exists(imgfile):
                        try:
                            img = XLImage(imgfile)
                            img.width, img.height = 32, 24
                            cell_ref = f"C{row_num}"
                            ws.add_image(img, cell_ref)
                        except Exception:
                            pass

                ws.cell(row=row_num, column=7, value=f"(R$ {row['preco_base']:.2f})").alignment = Alignment(horizontal='right')
                ws.cell(row=row_num, column=7).font = Font(size=10)
                ws.cell(row=row_num, column=8, value=f"R$ {row['preco_de_venda']:.2f}").font = Font(bold=True, size=13)
                ws.cell(row=row_num, column=8).alignment = Alignment(horizontal='right')

                uvas = [str(row.get(f"uva{i}", "")).strip() for i in range(1,4)]
                uvas = [u for u in uvas if u and u.lower() != "nan"]
                regiao_str = f"{row.get('pais','')} | {row.get('regiao','')}"
                if uvas:
                    regiao_str += f" | {', '.join(uvas)}"
                ws.cell(row=row_num+1, column=2, value=regiao_str).font = Font(size=10)
                amad = str(row.get("amadurecimento", ""))
                if amad and amad.lower() != "nan":
                    ws.cell(row=row_num+1, column=3, value="🛢️").font = Font(size=10)

                row_num += 2
                ordem_geral += 1

    ws.column_dimensions[get_column_letter(1)].width = 13
    ws.column_dimensions[get_column_letter(2)].width = 45
    ws.column_dimensions[get_column_letter(3)].width = 8
    ws.column_dimensions[get_column_letter(7)].width = 16
    ws.column_dimensions[get_column_letter(8)].width = 16

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream

# ===================== APP =====================
def main():
    st.set_page_config(page_title="Sugestão de Carta de Vinhos", layout="wide")
    garantir_pastas()

    # Barra de título semelhante ao Tkinter
    st.markdown("### Sugestão de Carta de Vinhos")

    # Top bar (nome cliente, logo, inserir foto, tabela de preço, busca global, fator, reset)
    with st.container():
        c1, c2, c3, c4, c5, c6, c7, c8 = st.columns([1.4,1.2,1,1,1.6,0.7,0.9,1.2])
        with c1:
            cliente = st.text_input("Nome do Cliente", value="", placeholder="(opcional)")
        with c2:
            logo_cliente = st.file_uploader("Carregar logo (cliente)", type=["png","jpg","jpeg"])
            logo_bytes = logo_cliente.read() if logo_cliente else None
        with c3:
            inserir_foto = st.checkbox("Inserir foto no PDF/Excel", value=True)
        with c4:
            preco_flag = st.selectbox("Tabela de preço", ["preco1", "preco2", "preco15", "preco38", "preco39", "preco55", "preco63"], index=0, help="Escolha a coluna base de preço")
        with c5:
            termo_global = st.text_input("Buscar", value="")
        with c6:
            fator_global = st.number_input("Fator", min_value=0.0, value=2.0, step=0.1)
        with c7:
            resetar = st.button("Resetar/Mostrar Todos")
        with c8:
            caminho_planilha = st.text_input("Arquivo de dados", value="vinhos1.xls", help="Caminho do arquivo XLS/XLSX (ex.: vinhos1.xls)")

    # Carregamento dos dados
    df = ler_excel_vinhos(caminho_planilha)
    df = atualiza_coluna_preco_base(df, preco_flag)

    # Aplicação do fator global
    df["fator"] = fator_global
    df["preco_de_venda"] = df["preco_base"] * df["fator"]

    # Sidebar de filtros (espelhando o LabelFrame de filtros no Tkinter)
    st.sidebar.header("Filtros")
    pais_opc = [""] + sorted([p for p in df["pais"].dropna().astype(str).unique().tolist() if p])
    tipo_opc = [""] + sorted([t for t in df["tipo"].dropna().astype(str).unique().tolist() if t])
    desc_opc = [""] + sorted([d for d in df["descricao"].dropna().astype(str).unique().tolist() if d])
    regiao_opc = [""] + sorted([r for r in df["regiao"].dropna().astype(str).unique().tolist() if r])
    cod_opc = [""] + sorted([str(c) for c in df["cod"].dropna().astype(str).unique().tolist()])

    filt_pais = st.sidebar.selectbox("País", pais_opc, index=0)
    filt_tipo = st.sidebar.selectbox("Tipo", tipo_opc, index=0)
    filt_desc = st.sidebar.selectbox("Descrição", desc_opc, index=0)
    filt_regiao = st.sidebar.selectbox("Região", regiao_opc, index=0)
    filt_cod = st.sidebar.selectbox("Código", cod_opc, index=0)
    col1, col2 = st.sidebar.columns(2)
    with col1:
        preco_min = st.number_input("Preço mín (base)", min_value=0.0, value=0.0, step=1.0)
    with col2:
        preco_max = st.number_input("Preço máx (base)", min_value=0.0, value=0.0, step=1.0, help="0 = sem limite")

    # Estado de seleção
    if "selected_idxs" not in st.session_state:
        st.session_state.selected_idxs = set()
    if "manual_fat" not in st.session_state:
        st.session_state.manual_fat = {}  # idx -> fator
    if "manual_preco_venda" not in st.session_state:
        st.session_state.manual_preco_venda = {}  # idx -> pv

    # Aplicar filtros
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
        df_filtrado = df_filtrado[df_filtrado["preco_base"] >= float(preco_min)]
    if preco_max and preco_max > 0:
        df_filtrado = df_filtrado[df_filtrado["preco_base"] <= float(preco_max)]

    if resetar:
        st.session_state.selected_idxs = set()
        df_filtrado = df.copy()

    # Contagem por tipo (igual ao label no Tkinter)
    tipo_map = {'branco': 'Brancos', 'tinto': 'Tintos', 'rose': 'Rosés', 'rosé': 'Rosés', 'espumante': 'Espumantes'}
    contagem = {'Brancos': 0, 'Tintos': 0, 'Rosés': 0, 'Espumantes': 0, 'outros': 0}
    for t in df_filtrado['tipo'].dropna().unique():
        lbl = next((lbl for k, lbl in tipo_map.items() if k in str(t).lower()), 'outros')
        contagem[lbl] = len(df_filtrado[df_filtrado['tipo'] == t])
    total = len(df_filtrado)
    selecionados = len(st.session_state.selected_idxs)
    st.caption(f"Brancos: {contagem.get('Brancos', 0)} | Tintos: {contagem.get('Tintos', 0)} | Rosés: {contagem.get('Rosés', 0)} | Espumantes: {contagem.get('Espumantes', 0)} | Total: {total} | Selecionados: {selecionados} | Fator: {fator_global:.2f}")

    # Tabela editável semelhante ao Treeview, com seleção
    # Construímos uma view com colunas chave e flags de seleção/foto
    view = df_filtrado.copy()
    view["selecionado"] = view["idx"].apply(lambda i: i in st.session_state.selected_idxs)
    view["foto"] = view["cod"].apply(lambda c: "●" if get_imagem_file(str(c)) else "")
    view_display_cols = ["selecionado", "foto", "cod", "descricao", "pais", "regiao", "preco_base", "preco_de_venda", "fator", "idx"]
    view = view_display_cols[:8] + ["fator"] # ensure order

    editable_cols = ["selecionado", "preco_de_venda", "fator"]
    edited = st.data_editor(
        df_filtrado.assign(
            selecionado=lambda d: d["idx"].apply(lambda i: i in st.session_state.selected_idxs),
            foto=lambda d: d["cod"].apply(lambda c: "●" if get_imagem_file(str(c)) else ""),
        )[["selecionado","foto","cod","descricao","pais","regiao","preco_base","preco_de_venda","fator","idx"]],
        hide_index=True,
        column_config={
            "selecionado": st.column_config.CheckboxColumn("SELECIONADO"),
            "foto": st.column_config.TextColumn("FOTO"),
            "cod": st.column_config.NumberColumn("COD"),
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

    # Sincroniza seleção/edições com o estado
    # Atualiza seleções
    new_selected = set(edited.loc[edited["selecionado"]==True, "idx"].astype(int).tolist()) if not edited.empty else set()
    st.session_state.selected_idxs -= set([i for i in st.session_state.selected_idxs if i not in new_selected])
    st.session_state.selected_idxs |= new_selected

    # Atualiza ajustes manuais
    for _, r in edited.iterrows():
        idx = int(r["idx"])
        if pd.notnull(r.get("fator")):
            st.session_state.manual_fat[idx] = float(r["fator"])
        if pd.notnull(r.get("preco_de_venda")):
            st.session_state.manual_preco_venda[idx] = float(r["preco_de_venda"])

    # Aplica ajustes no df base (apenas nos índices visíveis/gerais)
    for idx, fat in st.session_state.manual_fat.items():
        df.loc[df["idx"]==idx, "fator"] = fat
    df["preco_de_venda"] = df["preco_base"] * df["fator"]
    for idx, pv in st.session_state.manual_preco_venda.items():
        df.loc[df["idx"]==idx, "preco_de_venda"] = pv

    # Botões como no Tkinter
    cA, cB, cC, cD, cE = st.columns([1,1.2,1.2,1.2,1.2])
    with cA:
        ver_preview = st.button("Visualizar Sugestão")
    with cB:
        ver_marcados = st.button("Visualizar Itens Marcados")
    with cC:
        gerar_pdf_btn = st.button("Gerar PDF")
    with cD:
        exportar_excel_btn = st.button("Exportar para Excel")
    with cE:
        salvar_sugestao_btn = st.button("Salvar Sugestão")

    # Funções dos botões
    if ver_preview:
        if not st.session_state.selected_idxs:
            st.info("Nenhum item selecionado.")
        else:
            st.subheader("Pré-visualização da Sugestão")
            df_sel = df[df["idx"].isin(st.session_state.selected_idxs)].copy()
            df_sel = ordenar_para_saida(df_sel)
            # Gera preview textual
            preview_lines = []
            preview_lines.append("Sugestão Carta de Vinhos")
            if cliente:
                preview_lines.append(f"Cliente: {cliente}")
            preview_lines.append("="*70)
            ordem_geral = 1
            tipo_map = {'branco':'Brancos', 'tinto':'Tintos', 'rose':'Rosés', 'rosé':'Rosés', 'espumante':'Espumantes'}
            for tipo in df_sel['tipo'].dropna().unique():
                preview_lines.append(f"\n{str(tipo).upper()}")
                for pais in df_sel[df_sel['tipo']==tipo]['pais'].dropna().unique():
                    preview_lines.append(f"  {str(pais).upper()}")
                    grupo = df_sel[(df_sel['tipo']==tipo) & (df_sel['pais']==pais)]
                    for _, row in grupo.iterrows():
                        desc = row['descricao']
                        preco = f"R$ {row['preco_base']:.2f}"
                        pvenda = f"R$ {row['preco_de_venda']:.2f}"
                        regiao = row['regiao']
                        cod = int(row['cod']) if pd.notnull(row['cod']) else ""
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
            df_sel = df_sel[["cod","descricao","pais","regiao","preco_base","preco_de_venda"]].sort_values(["pais","descricao"])
            st.dataframe(df_sel, use_container_width=True)

    if gerar_pdf_btn:
        if not st.session_state.selected_idxs:
            st.warning("Selecione ao menos um vinho.")
        else:
            df_sel = df[df["idx"].isin(st.session_state.selected_idxs)].copy()
            df_sel = ordenar_para_saida(df_sel)
            pdf_buffer = gerar_pdf(df_sel, "Sugestão Carta de Vinhos", cliente, inserir_foto, logo_bytes)
            st.download_button("Baixar PDF", data=pdf_buffer, file_name="sugestao_carta_vinhos.pdf", mime="application/pdf")

    if exportar_excel_btn:
        if not st.session_state.selected_idxs:
            st.warning("Selecione ao menos um vinho.")
        else:
            df_sel = df[df["idx"].isin(st.session_state.selected_idxs)].copy()
            df_sel = ordenar_para_saida(df_sel)
            xlsx = exportar_excel_like_pdf(df_sel, inserir_foto=inserir_foto)
            st.download_button("Baixar Excel", data=xlsx, file_name="sugestao_carta_vinhos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    if salvar_sugestao_btn:
        if not st.session_state.selected_idxs:
            st.info("Selecione produtos para salvar.")
        else:
            garantir_pastas()
            nome = st.text_input("Nome da sugestão para salvar:", value="", key="nome_sugestao")
            if nome:
                path = os.path.join(SUGESTOES_DIR, f"{nome}.txt")
                try:
                    with open(path, "w") as f:
                        f.write(",".join(map(str, sorted(list(st.session_state.selected_idxs)))))
                    st.success(f"Sugestão '{nome}' salva em {path}.")
                except Exception as e:
                    st.error(f"Erro ao salvar: {e}")

    # Aba Sugestões Salvas (lista e ações) + Cadastro
    st.markdown("---")
    tab1, tab2 = st.tabs(["Sugestões Salvas", "Cadastro de Vinhos"])

    with tab1:
        garantir_pastas()
        # Lista de sugestões existentes
        arquivos = [f for f in os.listdir(SUGESTOES_DIR) if f.endswith(".txt")]
        sel = st.selectbox("Abrir sugestão", [""] + [a[:-4] for a in arquivos])
        colx, coly, colz = st.columns([1,1,1])
        with colx:
            if st.button("Excluir sugestão selecionada") and sel:
                try:
                    os.remove(os.path.join(SUGESTOES_DIR, f"{sel}.txt"))
                    st.success(f"Sugestão '{sel}' excluída. Recarregue a página para atualizar a lista.")
                except Exception as e:
                    st.error(f"Erro ao excluir: {e}")
        with coly:
            if st.button("Editar itens (carregar na grade)") and sel:
                path = os.path.join(SUGESTOES_DIR, f"{sel}.txt")
                if os.path.exists(path):
                    try:
                        with open(path) as f:
                            indices = [int(x) for x in f.read().strip().split(",") if x]
                        st.session_state.selected_idxs = set(indices)
                        st.success("Itens carregados. Role até a grade principal para ver/editar.")
                    except Exception as e:
                        st.error(f"Erro ao carregar: {e}")
        with colz:
            if st.button("Adicionar todos os produtos (mostrar todos)"):
                st.session_state.selected_idxs = set()

    with tab2:
        st.caption("Cadastrar novo produto (entra apenas na sessão atual; salve no seu Excel depois, se quiser persistir).")
        c1, c2, c3, c4, c5, c6, c7 = st.columns([1,2,1,1,1,1,1.2])
        with c1: new_cod = st.text_input("Código")
        with c2: new_desc = st.text_input("Descrição")
        with c3: new_preco = st.number_input("Preço" min_value=0.0, value=0.0, step=0.01, , key="cadastro_preco")
        with c4: new_fat = st.number_input("Fator", min_value=0.0, value=float(fator_global), key="cadastro_fator", step=0.1)
        with c5: new_pv = st.number_input("Preço Venda" min_value=0.0, value=0.0, step=0.01, , key="cadastro_preco_venda")
        with c6: new_pais = st.text_input("País")
        with c7: new_regiao = st.text_input("Região")
        if st.button("Cadastrar"):
            try:
                cod_int = int(float(new_cod)) if new_cod else None
                pv_calc = new_pv if new_pv > 0 else new_preco * new_fat
                novo = {
                    "idx": int(df["idx"].max()+1) if len(df)>0 else 0,
                    "cod": cod_int,
                    "descricao": new_desc,
                    "preco_base": new_preco,
                    "fator": new_fat,
                    "preco_de_venda": pv_calc,
                    "pais": new_pais,
                    "regiao": new_regiao,
                    "tipo": "",
                }
                # Atualiza apenas sessão atual (não salva no Excel automaticamente)
                # Caso deseje persistir, exporte/mescle externamente no seu arquivo de dados.
                st.session_state.setdefault("cadastrados", [])
                st.session_state["cadastrados"].append(novo)
                st.success("Produto cadastrado na sessão atual.")
            except Exception as e:
                st.error(f"Erro ao cadastrar: {e}")

        # Visualização dos cadastrados nesta sessão
        if st.session_state.get("cadastrados"):
            st.dataframe(pd.DataFrame(st.session_state["cadastrados"]), use_container_width=True)

if __name__ == "__main__":
    import pandas as pd
    main()
