
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
app_streamlit_final_v3_persist5.py

Melhorias adicionais:
- País e Região agora usam o mesmo padrão do Tipo:
  - st.multiselect (rápido, com busca)
  - expander com checkboxes por grupo
  - seleção incremental via session_state (Adicionar/limpar)
- Mantém filtros/ordenações avançadas, persistência da seleção de itens,
  mesclagem de sugestões e cálculo robusto do preço de venda.
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
    s = s.astype(str).str.replace("\u00A0", "", regex=False).str.strip()
    s = s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(default)

def to_float_series(s, default=0.0):
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").fillna(default)
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
    df["idx"] = pd.to_numeric(df["idx"], errors="coerce").fillna(-1).astype(int)

    for col in ["preco38","preco39","preco1","preco2","preco15","preco55","preco63","preco_base","fator","preco_de_venda"]:
        if col not in df.columns:
            df[col] = 0.0
        else:
            df[col] = to_float_series(df[col], default=0.0)

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
    base = df[flag] if flag in df.columns else df.get("preco1", 0.0)
    df["preco_base"] = to_float_series(base, default=0.0)
    if "fator" not in df.columns:
        df["fator"] = fator_global
    df["fator"] = to_float_series(df["fator"], default=fator_global)
    df["fator"] = df["fator"].apply(lambda x: fator_global if pd.isna(x) or x <= 0 else x)
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

# ============== Filtros Avançados e Ordenação ==============
def init_rule_states():
    if "filter_rules" not in st.session_state:
        st.session_state.filter_rules = []
    if "sort_rules" not in st.session_state:
        st.session_state.sort_rules = []
    # Tipos
    if "filt_tipos" not in st.session_state:
        st.session_state.filt_tipos = []
    if "filt_tipos_expander" not in st.session_state:
        st.session_state.filt_tipos_expander = {}
    # Países
    if "filt_paises" not in st.session_state:
        st.session_state.filt_paises = []
    if "filt_paises_expander" not in st.session_state:
        st.session_state.filt_paises_expander = {}
    # Regiões
    if "filt_regioes" not in st.session_state:
        st.session_state.filt_regioes = []
    if "filt_regioes_expander" not in st.session_state:
        st.session_state.filt_regioes_expander = {}

def add_filter_rule(col, op, val):
    if col and op:
        st.session_state.filter_rules.append({"col": col, "op": op, "val": val})

def remove_filter_rule(idx):
    if 0 <= idx < len(st.session_state.filter_rules):
        st.session_state.filter_rules.pop(idx)

def clear_filter_rules():
    st.session_state.filter_rules = []

def add_sort_rule(col, direction):
    if col and direction:
        st.session_state.sort_rules.append({"col": col, "dir": direction})

def remove_sort_rule(idx):
    if 0 <= idx < len(st.session_state.sort_rules):
        st.session_state.sort_rules.pop(idx)

def clear_sort_rules():
    st.session_state.sort_rules = []

def apply_filter_rules(df):
    if not st.session_state.filter_rules:
        return df
    mask = pd.Series(True, index=df.index)
    for rule in st.session_state.filter_rules:
        col = rule["col"]
        op = rule["op"]
        val = rule["val"]
        if col not in df.columns:
            continue
        series = df[col]
        as_num = pd.to_numeric(series, errors="coerce")
        val_num = pd.to_numeric(pd.Series([val]), errors="coerce").iloc[0]

        if op in ("contém", "não contém", "=", "<>"):
            series_str = series.astype(str).str.lower()
            val_str = str(val).lower()
            if op == "contém":
                cond = series_str.str.contains(val_str, na=False)
            elif op == "não contém":
                cond = ~series_str.str.contains(val_str, na=False)
            elif op == "=":
                cond = series_str.fillna("") == val_str
            elif op == "<>":
                cond = series_str.fillna("") != val_str
        else:
            if op == ">":
                cond = as_num > val_num
            elif op == "<":
                cond = as_num < val_num
            elif op == ">=":
                cond = as_num >= val_num
            elif op == "<=":
                cond = as_num <= val_num
            else:
                cond = pd.Series(True, index=df.index)

        mask &= cond.fillna(False)
    return df[mask]

def apply_sort_rules(df):
    if not st.session_state.sort_rules:
        return df
    cols = [r["col"] for r in st.session_state.sort_rules if r["col"] in df.columns]
    if not cols:
        return df
    asc = [True if r["dir"] == "asc" else False for r in st.session_state.sort_rules if r["col"] in df.columns]
    sort_df = df.copy()
    for c in cols:
        try:
            sort_df[c] = pd.to_numeric(sort_df[c], errors="ignore")
        except Exception:
            pass
    return sort_df.sort_values(by=cols, ascending=asc)

# ===================== APP =====================
def main():
    st.set_page_config(page_title="Sugestão de Carta de Vinhos", layout="wide")
    garantir_pastas()
    init_rule_states()

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
            termo_global = st.text_input("Buscar (contém em qualquer coluna)", value="", key="termo_global")
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

    # Sidebar: Filtros rápidos
    st.sidebar.header("Filtros rápidos")

    def options_from(col):
        if col not in df.columns: return []
        return sorted([x for x in df[col].dropna().astype(str).unique().tolist() if x])

    # ====== TIPO (multiselect, expander, incremental) ======
    tipo_opc = options_from("tipo")
    mult_tipos = st.sidebar.multiselect("Tipos (multi)", tipo_opc, default=st.session_state.get("filt_tipos", []), key="ms_tipos")
    st.session_state.filt_tipos = mult_tipos

    with st.sidebar.expander("Selecionar tipos (checkbox por grupo)"):
        sel_map = st.session_state.get("filt_tipos_expander", {})
        new_sel_map = {}
        for tp in tipo_opc:
            checked = (tp in st.session_state.filt_tipos) if tp not in sel_map else sel_map.get(tp, False)
            new_sel_map[tp] = st.checkbox(tp, value=checked, key=f"chk_tipo_{tp}")
        st.session_state.filt_tipos_expander = new_sel_map

        if st.button("Aplicar seleção de tipos"):
            nova_lista = [tp for tp, val in st.session_state.filt_tipos_expander.items() if val]
            st.session_state.filt_tipos = nova_lista
            st.session_state.ms_tipos = nova_lista
            st.experimental_rerun()

    with st.sidebar.expander("Adicionar tipo (incremental)"):
        escolha_add = st.selectbox("Adicionar tipo:", [""] + tipo_opc, key="add_tipo_select")
        if st.button("Adicionar tipo"):
            if escolha_add and escolha_add not in st.session_state.filt_tipos:
                st.session_state.filt_tipos.append(escolha_add)
                st.session_state.ms_tipos = list(st.session_state.filt_tipos)
                st.success(f"'{escolha_add}' adicionado ao filtro.")
                st.experimental_rerun()
        if st.button("Limpar tipos"):
            st.session_state.filt_tipos = []
            st.session_state.ms_tipos = []
            st.session_state.filt_tipos_expander = {}
            st.info("Filtros de Tipo limpos.")
            st.experimental_rerun()

    # ====== PAÍS (multiselect, expander, incremental) ======
    pais_opc = options_from("pais")
    mult_paises = st.sidebar.multiselect("País (multi)", pais_opc, default=st.session_state.get("filt_paises", []), key="ms_paises")
    st.session_state.filt_paises = mult_paises

    with st.sidebar.expander("Selecionar países (checkbox por grupo)"):
        sel_map_p = st.session_state.get("filt_paises_expander", {})
        new_sel_map_p = {}
        for pv in pais_opc:
            checked = (pv in st.session_state.filt_paises) if pv not in sel_map_p else sel_map_p.get(pv, False)
            new_sel_map_p[pv] = st.checkbox(pv, value=checked, key=f"chk_pais_{pv}")
        st.session_state.filt_paises_expander = new_sel_map_p

        if st.button("Aplicar seleção de países"):
            nova_lista = [pv for pv, val in st.session_state.filt_paises_expander.items() if val]
            st.session_state.filt_paises = nova_lista
            st.session_state.ms_paises = nova_lista
            st.experimental_rerun()

    with st.sidebar.expander("Adicionar país (incremental)"):
        escolha_p = st.selectbox("Adicionar país:", [""] + pais_opc, key="add_pais_select")
        if st.button("Adicionar país"):
            if escolha_p and escolha_p not in st.session_state.filt_paises:
                st.session_state.filt_paises.append(escolha_p)
                st.session_state.ms_paises = list(st.session_state.filt_paises)
                st.success(f"'{escolha_p}' adicionado ao filtro.")
                st.experimental_rerun()
        if st.button("Limpar países"):
            st.session_state.filt_paises = []
            st.session_state.ms_paises = []
            st.session_state.filt_paises_expander = {}
            st.info("Filtros de País limpos.")
            st.experimental_rerun()

    # ====== REGIÃO (multiselect, expander, incremental) ======
    regiao_opc = options_from("regiao")
    mult_regioes = st.sidebar.multiselect("Região (multi)", regiao_opc, default=st.session_state.get("filt_regioes", []), key="ms_regioes")
    st.session_state.filt_regioes = mult_regioes

    with st.sidebar.expander("Selecionar regiões (checkbox por grupo)"):
        sel_map_r = st.session_state.get("filt_regioes_expander", {})
        new_sel_map_r = {}
        for rv in regiao_opc:
            checked = (rv in st.session_state.filt_regioes) if rv not in sel_map_r else sel_map_r.get(rv, False)
            new_sel_map_r[rv] = st.checkbox(rv, value=checked, key=f"chk_regiao_{rv}")
        st.session_state.filt_regioes_expander = new_sel_map_r

        if st.button("Aplicar seleção de regiões"):
            nova_lista = [rv for rv, val in st.session_state.filt_regioes_expander.items() if val]
            st.session_state.filt_regioes = nova_lista
            st.session_state.ms_regioes = nova_lista
            st.experimental_rerun()

    with st.sidebar.expander("Adicionar região (incremental)"):
        escolha_r = st.selectbox("Adicionar região:", [""] + regiao_opc, key="add_regiao_select")
        if st.button("Adicionar região"):
            if escolha_r and escolha_r not in st.session_state.filt_regioes:
                st.session_state.filt_regioes.append(escolha_r)
                st.session_state.ms_regioes = list(st.session_state.filt_regioes)
                st.success(f"'{escolha_r}' adicionado ao filtro.")
                st.experimental_rerun()
        if st.button("Limpar regiões"):
            st.session_state.filt_regioes = []
            st.session_state.ms_regioes = []
            st.session_state.filt_regioes_expander = {}
            st.info("Filtros de Região limpos.")
            st.experimental_rerun()

    # ====== Descrição e Código (simples) ======
    desc_opc = [""] + options_from("descricao")
    cod_opc = [""] + options_from("cod")
    filt_desc = st.sidebar.selectbox("Descrição", desc_opc, index=0, key="filt_desc")
    filt_cod  = st.sidebar.selectbox("Código", cod_opc, index=0, key="filt_cod")

    colp1, colp2 = st.sidebar.columns(2)
    with colp1:
        preco_min = st.number_input("Preço mín (base)", min_value=0.0, value=0.0, step=1.0, key="preco_min")
    with colp2:
        preco_max = st.number_input("Preço máx (base)", min_value=0.0, value=0.0, step=1.0, help="0 = sem limite", key="preco_max")

    # Filtros Avançados
    with st.sidebar.expander("Filtros avançados (todas as colunas)", expanded=False):
        cols = df.columns.tolist()
        fc1, fc2, fc3 = st.columns([1.1,0.9,1.2])
        with fc1:
            col_sel = st.selectbox("Coluna", cols, key="adv_col")
        with fc2:
            ops = ["=", "<>", "contém", "não contém", ">", "<", ">=", "<="]
            op_sel = st.selectbox("Operador", ops, index=2, key="adv_op")
        with fc3:
            val_sel = st.text_input("Valor", value="", key="adv_val")

        add_r, clear_r = st.columns([1,1])
        with add_r:
            if st.button("Adicionar regra", key="btn_add_rule"):
                add_filter_rule(col_sel, op_sel, val_sel)
        with clear_r:
            if st.button("Limpar regras", key="btn_clear_rules"):
                clear_filter_rules()

        if st.session_state.filter_rules:
            st.caption("Regras ativas:")
            for i, r in enumerate(st.session_state.filter_rules):
                c1, c2, c3, c4 = st.columns([1.2,0.8,1.2,0.6])
                c1.write(f"**{r['col']}**")
                c2.write(r["op"])
                c3.write(str(r["val"]))
                if c4.button("Remover", key=f"rm_rule_{i}"):
                    remove_filter_rule(i)
                    st.experimental_rerun()

    # Ordenação
    with st.sidebar.expander("Ordenação (todas as colunas)", expanded=False):
        cols = df.columns.tolist()
        sc1, sc2 = st.columns([1.2,0.8])
        with sc1:
            sort_col = st.selectbox("Coluna", cols, key="sort_col")
        with sc2:
            sort_dir = st.selectbox("Direção", ["asc", "desc"], index=0, key="sort_dir")
        sbtn1, sbtn2 = st.columns([1,1])
        with sbtn1:
            if st.button("Adicionar ordenação", key="btn_add_sort"):
                add_sort_rule(sort_col, sort_dir)
        with sbtn2:
            if st.button("Limpar ordenações", key="btn_clear_sort"):
                clear_sort_rules()

        if st.session_state.sort_rules:
            st.caption("Ordenações ativas:")
            for i, r in enumerate(st.session_state.sort_rules):
                c1, c2, c3 = st.columns([1.2,0.8,0.6])
                c1.write(f"**{r['col']}**")
                c2.write("ascendente" if r["dir"]=="asc" else "descendente")
                if c3.button("Remover", key=f"rm_sort_{i}"):
                    remove_sort_rule(i)
                    st.experimental_rerun()

    # Aplicar filtros (VIEW)
    df_filtrado = df.copy()
    if termo_global.strip():
        term = termo_global.strip().lower()
        mask = df_filtrado.apply(lambda row: term in " ".join(str(v).lower() for v in row.values), axis=1)
        df_filtrado = df_filtrado[mask]

    # TIPOS
    if st.session_state.filt_tipos:
        df_filtrado = df_filtrado[df_filtrado["tipo"].astype(str).isin(st.session_state.filt_tipos)]
    # PAÍSES
    if st.session_state.filt_paises:
        df_filtrado = df_filtrado[df_filtrado["pais"].astype(str).isin(st.session_state.filt_paises)]
    # REGIÕES
    if st.session_state.filt_regioes:
        df_filtrado = df_filtrado[df_filtrado["regiao"].astype(str).isin(st.session_state.filt_regioes)]

    if filt_desc:
        df_filtrado = df_filtrado[df_filtrado["descricao"] == filt_desc]
    if filt_cod:
        df_filtrado = df_filtrado[df_filtrado["cod"].astype(str) == filt_cod]
    if preco_min:
        df_filtrado = df_filtrado[df_filtrado["preco_base"].fillna(0) >= float(preco_min)]
    if preco_max and preco_max > 0:
        df_filtrado = df_filtrado[df_filtrado["preco_base"].fillna(0) <= float(preco_max)]

    # Filtros/Ordenações avançadas
    df_filtrado = apply_filter_rules(df_filtrado)
    df_filtrado = apply_sort_rules(df_filtrado)

    if resetar:
        df_filtrado = df.copy()
        clear_filter_rules()
        clear_sort_rules()
        # limpa listas & mapas de seleção múltipla
        for k in ["filt_tipos","filt_paises","filt_regioes"]:
            st.session_state[k] = []
        for k in ["ms_tipos","ms_paises","ms_regioes"]:
            st.session_state[k] = []
        for k in ["filt_tipos_expander","filt_paises_expander","filt_regioes_expander"]:
            st.session_state[k] = {}
        st.experimental_rerun()

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

    # Ajustes manuais + recomputa preco_de_venda
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
        if sel:
            path = os.path.join(SUGESTOES_DIR, f"{sel}.txt")
            if os.path.exists(path):
                try:
                    with open(path) as f:
                        sugestao_indices = [int(x) for x in f.read().strip().split(",") if x]
                    st.session_state.selected_idxs = set(sugestao_indices)
                    st.info(f"Sugestão '{sel}' carregada: {len(sugestao_indices)} itens.")
                except Exception as e:
                    st.error(f"Erro ao carregar '{sel}': {e}")

        # Relação da sugestão (abaixo)
        if sel:
            st.subheader("Relação da Sugestão")
            df_rel = df[df["idx"].isin(st.session_state.selected_idxs)].copy()
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
                    garantir_pastas()
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
