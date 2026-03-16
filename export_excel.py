"""
export_excel.py
---------------
Gera o Excel de tabulação MesoIDR a partir dos dados do site Flask.

Uso no app.py:
    from export_excel import gerar_excel, orm_para_dict

    @app.route("/exportar_excel")
    def exportar_excel():
        from flask import send_file
        import os
        todos = FormularioSoja.query.order_by(FormularioSoja.id).all()
        registros = [orm_para_dict(r) for r in todos]
        filepath = os.path.join(app.config.get("UPLOAD_FOLDER", "/tmp"),
                                "MesoIDR_Export.xlsx")
        gerar_excel(registros, filepath)
        return send_file(filepath, as_attachment=True,
                         download_name="MesoIDR_Exportacao.xlsx",
                         mimetype="application/vnd.openxmlformats-"
                                  "officedocument.spreadsheetml.sheet")
"""

from __future__ import annotations
from collections import defaultdict, Counter
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Constantes de domínio
# ---------------------------------------------------------------------------
REGIOES_IDR = [
    "Noroeste", "Norte", "Oeste", "Sudoeste",
    "Centro Sul", "Centro", "Metropolitana e Litoral",
]

PRAGAS = [
    "Lagarta da soja (Anticarsia gemmatalis)",
    "Lagarta das vagens (Spodoptera spp.)",
    "Lagarta falsa medideira (Chrysodeixis includens)",
    "Lagartas do grupo Heliothinae",
    "Percevejo barriga verde (Dichelops spp.)",
    "Percevejo marrom (Euschistus heros)",
    "Percevejo verde (Nezara viridula)",
    "Percevejo verde pequeno (Piezodorus guildinii)",
    "Broca dos ponteiros (Crocidosema aporema)",
    "Mosca Branca",
    "Outros insetos praga",
    "Tamandua da soja (Sternechus subsignatus)",
    "Tripes",
    "Vaquinhas (Diabrotica/ Cerotoma/ Colapsis)",
    "Acaros",
]

DOENCAS_FUNGICAS = [
    "Antracnose (Colletotrichum truncatum)",
    "Cancro da haste (Diaporthe spp.)",
    "Ferrugem asiatica (Phakopsora pachyrhizi)",
    "Mancha alvo (Corynespora cassiicola)",
    "Mancha de cercospora (Cercospora kikuchii)",
    "Mancha olho-de-ra (Cercospora sojina)",
    "Mancha parda (Septoria glycines)",
    "Mela ou requeima (Rhizoctonia solani)",
    "Mofo branco (Sclerotinia sclerotiorum)",
    "Mildio (Peronospora manshurica)",
    "Oidio (Microsphaera diffusa)",
    "Outras Doencas Fungicas",
]
DOENCAS_BACT = [
    "Crestamento bacteriano (Pseudomonas savastanoi pv. glycinea)",
    "Fogo selvagem (Pseudomonas syringae pv. tabaci)",
    "Pustula bacteriana (Xanthomonas axonopodis pv. glycines)",
    "Mancha bacteriana marrom (Curtobacterium flaccumfaciens pv. flaccumfaciens)",
]
DOENCAS = DOENCAS_FUNGICAS + DOENCAS_BACT

N_PULV = 7

_LAGARTAS = [p for p in PRAGAS if "Lagarta" in p]
_PERCEVEJOS = [p for p in PRAGAS if "Percevejo" in p]

# ---------------------------------------------------------------------------
# Paleta de cores
# ---------------------------------------------------------------------------
_AZUL_ESC  = "1F4E79"
_AZUL_MED  = "2E75B6"
_AZUL_CLA  = "BDD7EE"
_LINHA_PAR = "DEEAF1"
_BRANCO    = "FFFFFF"


# ---------------------------------------------------------------------------
# Helpers de estilo
# ---------------------------------------------------------------------------
def _brd():
    s = Side(border_style="thin", color="B8CCE4")
    return Border(left=s, right=s, top=s, bottom=s)

def _fill(hex_c):
    return PatternFill("solid", fgColor=hex_c)

def _hfont(bold=True, color=_BRANCO, size=9):
    return Font(name="Arial", bold=bold, color=color, size=size)

def _dfont(bold=False, size=9):
    return Font(name="Arial", bold=bold, size=size)

def _ctr():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def _lft():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

def _hdr(ws, row, col, text, bg=_AZUL_ESC, end_col=None):
    c = ws.cell(row=row, column=col, value=text)
    c.font = _hfont(color=_BRANCO if bg in (_AZUL_ESC, _AZUL_MED) else _AZUL_ESC)
    c.fill = _fill(bg)
    c.alignment = _ctr()
    c.border = _brd()
    if end_col and end_col > col:
        ws.merge_cells(start_row=row, start_column=col,
                       end_row=row, end_column=end_col)
    return c

def _dat(ws, row, col, val, bg=None, bold=False, fmt=None):
    c = ws.cell(row=row, column=col, value=val)
    c.font = _dfont(bold=bold)
    c.alignment = _lft()
    c.border = _brd()
    if bg:
        c.fill = _fill(bg)
    if fmt:
        c.number_format = fmt
    return c

def _auto_width(ws, mn=8, mx=35):
    for col_cells in ws.columns:
        ltr = get_column_letter(col_cells[0].column)
        w = max((len(str(c.value or "")) for c in col_cells), default=0)
        ws.column_dimensions[ltr].width = min(mx, max(mn, w * 1.1))


# ---------------------------------------------------------------------------
# Conversão ORM -> dict
# ---------------------------------------------------------------------------
def orm_para_dict(r) -> dict:
    """
    Converte FormularioSoja (com Pulverizacoes) para dict usado por gerar_excel().

    Modelo Pulverizacao:
      .tipo           -> 'dessecacao' | '1' | '2' ... '7'
      .data           -> string 'YYYY-MM-DD'
      .classe_produto -> ex. "Inseticida, Fungicida"
      .alvo           -> ex. "Lagarta da soja (Anticarsia gemmatalis), Ferrugem..."
    """
    # Indexa pulverizacoes por tipo
    pulvs: dict[str, Any] = {}
    for p in (r.pulverizacoes or []):
        pulvs[str(p.tipo).strip()] = p

    def _split(text: str) -> list[str]:
        if not text:
            return []
        return [x.strip() for x in text.replace("\n", ",").split(",") if x.strip()]

    def _dae(n: int):
        obj = pulvs.get(str(n))
        if not obj or not obj.data or not r.data_emergencia:
            return None
        try:
            from datetime import datetime
            dp = datetime.strptime(obj.data[:10], "%Y-%m-%d").date()
            de = datetime.strptime(r.data_emergencia[:10], "%Y-%m-%d").date()
            return (dp - de).days
        except Exception:
            return None

    d: dict[str, Any] = {}

    # Identificacao
    d["N"]             = r.id
    d["Meso_IDR"]      = r.meso_idr
    d["Regiao"]        = r.regiao
    d["Municipio"]     = r.municipio
    d["Area_Soja"]     = r.area_soja
    d["Cultivar"]      = r.cultivar
    d["Bt"]            = r.bt
    d["Produtividade"] = r.produtividade_media
    d["Dt_Plantio"]    = r.data_plantio
    d["Adversidade"]   = r.qual_adversidade if r.houve_adversidade == "SIM" else None
    d["Sinistro"]      = r.houve_adversidade

    # Monitoramento
    d["Conhec_MID"]  = r.conhecimento_mid
    d["Utiliza_MID"] = r.utiliza_mid
    d["Conhec_MIP"]  = r.conhecimento_mip
    d["Utiliza_MIP"] = r.utiliza_mip

    # Herbicidas
    herbs = [
        ("Herbicida", r.herbicida_dessecacao_alvo,  r.herbicida_dessecacao_aplicacoes),
        ("Herbicida", r.herbicida_pre_alvo,          r.herbicida_pre_aplicacoes),
        ("Herbicida", r.herbicida_pos_alvo,          r.herbicida_pos_aplicacoes),
    ]
    for i, (cl, alv, nap) in enumerate(herbs, start=1):
        d[f"Herb_Cl{i}"]  = cl if alv else None
        d[f"Herb_Alv{i}"] = alv
        d[f"Herb_Nap{i}"] = nap

    # Dessecacao
    dess = pulvs.get("dessecacao")
    d["Dess_Sim"]  = "SIM" if dess else "NAO"
    d["Dess_Dt"]   = dess.data if dess else None
    d["Dess_Cl"]   = dess.classe_produto if dess else None
    dess_alvos = _split(dess.alvo) if dess else []
    for i in range(1, 4):
        d[f"Dess_Alv{i}"] = dess_alvos[i - 1] if i <= len(dess_alvos) else None

    # Pulverizacoes 1-7
    for n in range(1, N_PULV + 1):
        obj = pulvs.get(str(n))
        alvos   = _split(obj.alvo)           if obj else []
        classes = _split(obj.classe_produto) if obj else []

        d[f"P{n}_DAE"]  = _dae(n)
        d[f"P{n}_Data"] = obj.data if obj else None

        for k in range(1, 6):
            d[f"P{n}_Cl{k}"]  = classes[k - 1] if k <= len(classes) else None
            d[f"P{n}_Alv{k}"] = alvos[k - 1]   if k <= len(alvos)   else None

        al = [a.lower() for a in alvos]
        d[f"P{n}_pragas"]  = {p: any(p.lower() in x or x in p.lower() for x in al)
                               for p in PRAGAS}
        d[f"P{n}_doencas"] = {dc: any(dc.lower() in x or x in dc.lower() for x in al)
                               for dc in DOENCAS}

    # Outras
    d["Tto_Semente"] = r.tratamento_sementes
    d["SAL_CB"]      = r.sal_mistura
    d["Ctrl_Biol"]   = r.controle_biologico

    # FBN
    d["Inoc_Usa"]   = r.inoculacao_sementes
    d["Inoc_Forma"] = r.forma_inoculacao
    d["Coinoc"]     = r.coinoculacao
    d["CoMo_Usa"]   = r.co_mo
    d["CoMo_Forma"] = r.co_mo_aplicacao

    return d


# ---------------------------------------------------------------------------
# Definicao de colunas do BD
# ---------------------------------------------------------------------------
_ID_COLS = [
    ("N",            "N° Questionario",             "num"),
    ("Meso_IDR",     "Mesorregiao IDR",             "txt"),
    ("Regiao",       "Regiao",                      "txt"),
    ("Municipio",    "Municipio",                   "txt"),
    ("Area_Soja",    "Area com Soja (ha)",           "num"),
    ("Cultivar",     "Cultivar",                    "txt"),
    ("Bt",           "Bt",                          "txt"),
    ("Produtividade","Produtividade Media (sc/ha)",  "num"),
    ("Dt_Plantio",   "Data Plantio",                "txt"),
    ("Adversidade",  "Adversidade",                 "txt"),
    ("Sinistro",     "Sinistro",                    "txt"),
]
_MID_COLS = [
    ("Conhec_MID", "Conhec. MID", "txt"),
    ("Utiliza_MID","Utiliza MID", "txt"),
    ("Conhec_MIP", "Conhec. MIP", "txt"),
    ("Utiliza_MIP","Utiliza MIP", "txt"),
]
_HERB_COLS = []
for i in range(1, 4):
    _HERB_COLS += [
        (f"Herb_Cl{i}",  f"Classe Produto {i}", "txt"),
        (f"Herb_Alv{i}", f"Alvo {i}",            "txt"),
        (f"Herb_Nap{i}", f"N Aplicacoes {i}",    "num"),
    ]
_DESS_COLS = [
    ("Dess_Sim",  "Pulverizou", "txt"),
    ("Dess_Dt",   "Data",       "txt"),
    ("Dess_Cl",   "Classe",     "txt"),
    ("Dess_Alv1", "Alvo 1",     "txt"),
    ("Dess_Alv2", "Alvo 2",     "txt"),
    ("Dess_Alv3", "Alvo 3",     "txt"),
]

def _pulv_cols(n):
    cols = [
        (f"P{n}_DAE",  f"{n}a Pulv - DAE",   "num"),
        (f"P{n}_Data", f"{n}a Pulv - Data",  "txt"),
    ]
    for k in range(1, 4):
        cols += [
            (f"P{n}_Cl{k}",  f"Classe {k}", "txt"),
            (f"P{n}_Alv{k}", f"Alvo {k}",   "txt"),
        ]
    return cols

_OUTRAS_COLS = [
    ("Tto_Semente", "Trat. Semente",   "txt"),
    ("SAL_CB",      "SAL+Inseticida",  "txt"),
    ("Ctrl_Biol",   "Ctrl. Biologico", "txt"),
]
_INOC_COLS = [
    ("Inoc_Usa",   "Inoculacao",   "txt"),
    ("Inoc_Forma", "Forma Inoc.",  "txt"),
    ("Coinoc",     "Coinoculacao", "txt"),
    ("CoMo_Usa",   "Co+Mo",        "txt"),
    ("CoMo_Forma", "Forma Co+Mo",  "txt"),
]

GRUPOS_BD = [
    ("IDENTIFICACAO",              _ID_COLS),
    ("MONITORAMENTO MIP/MID",      _MID_COLS),
    ("3. PLANTAS INVASORAS",       _HERB_COLS),
    ("4.0 DESSECACAO",             _DESS_COLS),
    *[(f"4.{n} - {n}a PULVERIZACAO", _pulv_cols(n)) for n in range(1, N_PULV + 1)],
    ("5. OUTRAS INFORMACOES",      _OUTRAS_COLS),
    ("6. FBN / INOCULACAO",        _INOC_COLS),
]

ALL_COLS: list[tuple] = []
for _, cols in GRUPOS_BD:
    ALL_COLS.extend(cols)

_CI: dict[str, int] = {c[0]: i + 1 for i, c in enumerate(ALL_COLS)}


# ---------------------------------------------------------------------------
# Aba BD
# ---------------------------------------------------------------------------
def _build_bd(wb, registros):
    ws = wb.create_sheet("BD")
    ws.sheet_view.showGridLines = False
    nc = len(ALL_COLS)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
    c = ws.cell(1, 1, "PLANILHA TABULACAO DADOS - APLICACAO DEFENSIVOS PARA "
                       "CONTROLE PRAGAS E DOENCAS - PR - SAFRA 2024/25")
    c.font = Font(name="Arial", bold=True, size=11, color=_BRANCO)
    c.fill = _fill(_AZUL_ESC); c.alignment = _ctr()
    ws.row_dimensions[1].height = 28

    col = 1
    for label, cols in GRUPOS_BD:
        _hdr(ws, 2, col, label, _AZUL_MED, col + len(cols) - 1)
        col += len(cols)

    ws.row_dimensions[3].height = 40
    for i, (key, label, _) in enumerate(ALL_COLS, start=1):
        c = _hdr(ws, 3, i, label, _AZUL_CLA)
        c.font = _hfont(color=_AZUL_ESC, size=8)

    for ri, reg in enumerate(registros, start=4):
        bg = _LINHA_PAR if ri % 2 == 0 else None
        for key, _, fmt in ALL_COLS:
            val = reg.get(key)
            nf = "#,##0.00" if fmt == "num" else None
            _dat(ws, ri, _CI[key], val, bg=bg, fmt=nf)

    _auto_width(ws)
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:{get_column_letter(nc)}3"


# ---------------------------------------------------------------------------
# Aba Medias_Geral
# ---------------------------------------------------------------------------
_ITENS_MEDIAS = [
    ("N Questionarios Aplicados",          "count",    None),
    ("Area Total Soja (ha)",               "sum",      "Area_Soja"),
    ("Area Media Cultivada (ha)",          "avg",      "Area_Soja"),
    ("Produtividade Media (sc/ha)",        "avg",      "Produtividade"),
    ("N com Sinistro",                     "count_if", ("Sinistro", "SIM")),
    ("% com Sinistro",                     "pct",      ("Sinistro", "SIM")),
    ("Conhecimento MID (%)",               "pct",      ("Conhec_MID",  "SIM")),
    ("Utiliza MID (%)",                    "pct",      ("Utiliza_MID", "SIM")),
    ("Conhecimento MIP (%)",               "pct",      ("Conhec_MIP",  "SIM")),
    ("Utiliza MIP (%)",                    "pct",      ("Utiliza_MIP", "SIM")),
    ("N com Aplicacao PRAGAS",             "count_ap", "pragas"),
    ("% com Aplicacao PRAGAS",             "pct_ap",   "pragas"),
    ("N aplicacoes PRAGAS - Total",        "sum_ap",   "pragas"),
    ("N aplicacoes PRAGAS - Aplicantes",   "avg_ap",   "pragas"),
    ("DAE Medio 1a Aplicacao PRAGAS",      "dae_1",    "pragas"),
    ("N com Aplicacao LAGARTAS",           "count_ap", "lagartas"),
    ("% com Aplicacao LAGARTAS",           "pct_ap",   "lagartas"),
    ("N medio aplicacoes LAGARTAS",        "avg_ap",   "lagartas"),
    ("N com Aplicacao PERCEVEJOS",         "count_ap", "percevejos"),
    ("% com Aplicacao PERCEVEJOS",         "pct_ap",   "percevejos"),
    ("N medio aplicacoes PERCEVEJOS",      "avg_ap",   "percevejos"),
    ("N com Aplicacao DOENCAS",            "count_ap", "doencas"),
    ("% com Aplicacao DOENCAS",            "pct_ap",   "doencas"),
    ("N medio aplicacoes DOENCAS",         "avg_ap",   "doencas"),
    ("DAE Medio 1a Aplicacao FERRUGEM",    "dae_1",    "ferrugem"),
    ("DAE Medio 1a Aplicacao DEMAIS DOE.", "dae_1",    "demais_doencas"),
    ("% Trat. Semente",                    "pct",      ("Tto_Semente", "SIM")),
    ("% SAL+Inseticida",                   "pct",      ("SAL_CB",      "SIM")),
    ("% Controle Biologico",               "pct",      ("Ctrl_Biol",   "SIM")),
    ("% Usa Inoculacao",                   "pct",      ("Inoc_Usa",    "SIM")),
    ("% Coinoculacao",                     "pct",      ("Coinoc",      "SIM")),
    ("% Co+Mo",                            "pct",      ("CoMo_Usa",    "SIM")),
]


def _n_aplic(reg, grupo):
    total = 0
    for n in range(1, N_PULV + 1):
        pm = reg.get(f"P{n}_pragas", {})
        dm = reg.get(f"P{n}_doencas", {})
        if grupo == "pragas":
            hit = any(pm.values()) or any(dm.values())
        elif grupo == "lagartas":
            hit = any(pm.get(p, False) for p in _LAGARTAS)
        elif grupo == "percevejos":
            hit = any(pm.get(p, False) for p in _PERCEVEJOS)
        elif grupo == "doencas":
            hit = any(dm.values())
        elif grupo == "ferrugem":
            hit = dm.get("Ferrugem asiatica (Phakopsora pachyrhizi)", False)
        elif grupo == "demais_doencas":
            hit = any(v for k, v in dm.items() if "Ferrugem" not in k and v)
        elif grupo == "op_acaros":
            hit = any(pm.get(p, False) for p in PRAGAS
                      if "Lagarta" not in p and "Percevejo" not in p)
        else:
            hit = False
        if hit:
            total += 1
    return total


def _dae_1(reg, grupo):
    for n in range(1, N_PULV + 1):
        pm = reg.get(f"P{n}_pragas", {})
        dm = reg.get(f"P{n}_doencas", {})
        if grupo == "pragas":
            hit = any(pm.values()) or any(dm.values())
        elif grupo == "ferrugem":
            hit = dm.get("Ferrugem asiatica (Phakopsora pachyrhizi)", False)
        elif grupo == "demais_doencas":
            hit = any(v for k, v in dm.items() if "Ferrugem" not in k and v)
        else:
            hit = False
        if hit:
            return reg.get(f"P{n}_DAE")
    return None


def _calc(regs, tipo, param):
    n = len(regs)
    if n == 0:
        return 0
    if tipo == "count":
        return n
    if tipo == "sum":
        return round(sum(r.get(param) or 0 for r in regs), 2)
    if tipo == "avg":
        v = [r.get(param) for r in regs if r.get(param) is not None]
        return round(sum(v) / len(v), 2) if v else 0
    if tipo == "count_if":
        k, val = param; return sum(1 for r in regs if r.get(k) == val)
    if tipo == "pct":
        k, val = param
        return round(sum(1 for r in regs if r.get(k) == val) / n, 4)
    if tipo == "count_ap":
        return sum(1 for r in regs if _n_aplic(r, param) > 0)
    if tipo == "pct_ap":
        return round(sum(1 for r in regs if _n_aplic(r, param) > 0) / n, 4)
    if tipo == "sum_ap":
        return sum(_n_aplic(r, param) for r in regs)
    if tipo == "avg_ap":
        ap = [_n_aplic(r, param) for r in regs if _n_aplic(r, param) > 0]
        return round(sum(ap) / len(ap), 2) if ap else 0
    if tipo == "dae_1":
        ds = [_dae_1(r, param) for r in regs]
        ds = [d for d in ds if d is not None]
        return round(sum(ds) / len(ds), 1) if ds else 0
    return 0


def _build_medias_geral(wb, registros):
    ws = wb.create_sheet("Medias_Geral")
    ws.sheet_view.showGridLines = False

    sub = ["Total", "Cultivares Bt", "Cultivares Nao Bt"]
    n_reg = len(REGIOES_IDR) + 1
    nc = 1 + n_reg * 3

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
    c = ws.cell(1, 1, "MEDIAS GERAIS - SAFRA 2024/25")
    c.font = Font(name="Arial", bold=True, size=11, color=_BRANCO)
    c.fill = _fill(_AZUL_ESC); c.alignment = _ctr()

    _hdr(ws, 2, 1, "Item", _AZUL_ESC)
    col = 2
    for reg in ["Parana"] + REGIOES_IDR:
        _hdr(ws, 2, col, reg, _AZUL_MED, col + 2); col += 3

    ws.cell(3, 1).border = _brd()
    col = 2
    for _ in range(n_reg):
        for s in sub:
            c = _hdr(ws, 3, col, s, _AZUL_CLA)
            c.font = _hfont(color=_AZUL_ESC, size=8); col += 1

    grp_r: dict[str, list] = defaultdict(list)
    for reg in registros:
        grp_r[reg.get("Meso_IDR", "")].append(reg)

    bt_filt = {
        "Total":           lambda r: True,
        "Cultivares Bt":   lambda r: r.get("Bt") == "SIM",
        "Cultivares Nao Bt": lambda r: r.get("Bt") == "NAO",
    }
    pct_tipos = {"pct", "pct_ap"}

    for ri, (label, tipo, param) in enumerate(_ITENS_MEDIAS, start=4):
        bg = _LINHA_PAR if ri % 2 == 0 else None
        _dat(ws, ri, 1, label, bg=bg, bold=True)
        col = 2
        for regiao in ["Parana"] + REGIOES_IDR:
            base = registros if regiao == "Parana" else grp_r.get(regiao, [])
            for s in sub:
                grp = [r for r in base if bt_filt[s](r)]
                val = _calc(grp, tipo, param)
                fmt = "0.0%" if tipo in pct_tipos else None
                _dat(ws, ri, col, val, bg=bg, fmt=fmt); col += 1

    _auto_width(ws, mn=12, mx=22)
    ws.freeze_panes = "B4"


# ---------------------------------------------------------------------------
# Aba Contagem_Pragas
# ---------------------------------------------------------------------------
_PR_EXTRAS = ["Total Lagartas", "Total Percevejos", "Total Outras Pragas",
              "Total Acaros", "Total Outras+Acaros", "TOTAL PRAGAS"]

def _build_contagem_pragas(wb, registros):
    ws = wb.create_sheet("Contagem_Pragas")
    ws.sheet_view.showGridLines = False

    id_h = ["N", "Meso_IDR", "Regiao", "Municipio", "Area (ha)",
            "Cultivar", "Bt", "Produtividade", "Data Plantio",
            "Adversidade", "Sinistro"]
    id_k = ["N", "Meso_IDR", "Regiao", "Municipio", "Area_Soja",
            "Cultivar", "Bt", "Produtividade", "Dt_Plantio",
            "Adversidade", "Sinistro"]
    ni = len(id_h)
    cp = 1 + len(PRAGAS) + len(_PR_EXTRAS)
    nc = ni + N_PULV * cp

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
    c = ws.cell(1, 1, "CONTAGEM DE PRAGAS POR PULVERIZACAO - SAFRA 2024/25")
    c.font = Font(name="Arial", bold=True, size=11, color=_BRANCO)
    c.fill = _fill(_AZUL_ESC); c.alignment = _ctr()

    _hdr(ws, 2, 1, "IDENTIFICACAO", _AZUL_ESC, ni)
    col = ni + 1
    for n in range(1, N_PULV + 1):
        _hdr(ws, 2, col, f"{n}a PULVERIZACAO", _AZUL_MED, col + cp - 1); col += cp

    for i, h in enumerate(id_h, start=1):
        c = _hdr(ws, 3, i, h, _AZUL_CLA); c.font = _hfont(color=_AZUL_ESC, size=8)

    col = ni + 1
    for n in range(1, N_PULV + 1):
        c = _hdr(ws, 3, col, "DAE", _AZUL_CLA); c.font = _hfont(color=_AZUL_ESC, size=8); col += 1
        for p in PRAGAS:
            c = _hdr(ws, 3, col, p, _AZUL_CLA); c.font = _hfont(color=_AZUL_ESC, size=8); col += 1
        for e in _PR_EXTRAS:
            _hdr(ws, 3, col, e, _AZUL_MED); col += 1

    for ri, reg in enumerate(registros, start=4):
        bg = _LINHA_PAR if ri % 2 == 0 else None
        for ci, key in enumerate(id_k, start=1):
            _dat(ws, ri, ci, reg.get(key), bg=bg)
        col = ni + 1
        for n in range(1, N_PULV + 1):
            _dat(ws, ri, col, reg.get(f"P{n}_DAE"), bg=bg, fmt="#,##0"); col += 1
            pm = reg.get(f"P{n}_pragas", {})
            ps = col
            for p in PRAGAS:
                _dat(ws, ri, col, 1 if pm.get(p) else 0, bg=bg); col += 1
            L = get_column_letter
            lag = [ps + i for i, p in enumerate(PRAGAS) if "Lagarta" in p]
            pev = [ps + i for i, p in enumerate(PRAGAS) if "Percevejo" in p]
            op  = [ps + i for i, p in enumerate(PRAGAS)
                   if "Lagarta" not in p and "Percevejo" not in p and "Acaro" not in p]
            ac  = [ps + i for i, p in enumerate(PRAGAS) if "Acaro" in p]
            tc = []
            for idx_list in [lag, pev, op, ac]:
                f = "=" + "+".join(f"{L(c)}{ri}" for c in idx_list)
                _dat(ws, ri, col, f, bg=bg); tc.append(col); col += 1
            _dat(ws, ri, col, f"={L(tc[2])}{ri}+{L(tc[3])}{ri}", bg=bg); col += 1
            _dat(ws, ri, col,
                 f"=SUM({L(ps)}{ri}:{L(ps+len(PRAGAS)-1)}{ri})", bg=bg); col += 1

    _auto_width(ws, mn=5, mx=20)
    ws.freeze_panes = f"{get_column_letter(ni+1)}4"


# ---------------------------------------------------------------------------
# Aba Contagem_Doencas
# ---------------------------------------------------------------------------
_DC_EXTRAS = ["Total Ferrugem", "Total Mancha Alvo", "Total Oidio",
              "Total Demais Fungicas", "Total Bacterianas", "TOTAL DOENCAS"]

def _build_contagem_doencas(wb, registros):
    ws = wb.create_sheet("Contagem_Doencas")
    ws.sheet_view.showGridLines = False

    id_h = ["N", "Meso_IDR", "Regiao", "Municipio", "Area (ha)",
            "Cultivar", "Bt", "Produtividade", "Data Plantio",
            "Adversidade", "Sinistro"]
    id_k = ["N", "Meso_IDR", "Regiao", "Municipio", "Area_Soja",
            "Cultivar", "Bt", "Produtividade", "Dt_Plantio",
            "Adversidade", "Sinistro"]
    ni = len(id_h)
    cp = 1 + len(DOENCAS) + len(_DC_EXTRAS)
    nc = ni + N_PULV * cp

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
    c = ws.cell(1, 1, "CONTAGEM DE DOENCAS POR PULVERIZACAO - SAFRA 2024/25")
    c.font = Font(name="Arial", bold=True, size=11, color=_BRANCO)
    c.fill = _fill(_AZUL_ESC); c.alignment = _ctr()

    _hdr(ws, 2, 1, "IDENTIFICACAO", _AZUL_ESC, ni)
    col = ni + 1
    for n in range(1, N_PULV + 1):
        _hdr(ws, 2, col, f"{n}a PULVERIZACAO", _AZUL_MED, col + cp - 1); col += cp

    for i, h in enumerate(id_h, start=1):
        c = _hdr(ws, 3, i, h, _AZUL_CLA); c.font = _hfont(color=_AZUL_ESC, size=8)

    col = ni + 1
    for n in range(1, N_PULV + 1):
        c = _hdr(ws, 3, col, "DAE", _AZUL_CLA); c.font = _hfont(color=_AZUL_ESC, size=8); col += 1
        for d in DOENCAS:
            c = _hdr(ws, 3, col, d, _AZUL_CLA); c.font = _hfont(color=_AZUL_ESC, size=8); col += 1
        for e in _DC_EXTRAS:
            _hdr(ws, 3, col, e, _AZUL_MED); col += 1

    for ri, reg in enumerate(registros, start=4):
        bg = _LINHA_PAR if ri % 2 == 0 else None
        for ci, key in enumerate(id_k, start=1):
            _dat(ws, ri, ci, reg.get(key), bg=bg)
        col = ni + 1
        for n in range(1, N_PULV + 1):
            _dat(ws, ri, col, reg.get(f"P{n}_DAE"), bg=bg); col += 1
            dm = reg.get(f"P{n}_doencas", {})
            ds = col
            for d in DOENCAS:
                _dat(ws, ri, col, 1 if dm.get(d) else 0, bg=bg); col += 1
            L = get_column_letter
            i_f = ds + DOENCAS.index("Ferrugem asiatica (Phakopsora pachyrhizi)")
            i_a = ds + DOENCAS.index("Mancha alvo (Corynespora cassiicola)")
            i_o = ds + DOENCAS.index("Oidio (Microsphaera diffusa)")
            i_b0 = ds + len(DOENCAS_FUNGICAS)
            i_b1 = ds + len(DOENCAS) - 1
            _dat(ws, ri, col, f"={L(i_f)}{ri}", bg=bg); col += 1
            _dat(ws, ri, col, f"={L(i_a)}{ri}", bg=bg); col += 1
            _dat(ws, ri, col, f"={L(i_o)}{ri}", bg=bg); col += 1
            dem = (f"=SUM({L(ds)}{ri}:{L(ds+len(DOENCAS_FUNGICAS)-1)}{ri})"
                   f"-{L(i_f)}{ri}-{L(i_a)}{ri}-{L(i_o)}{ri}")
            _dat(ws, ri, col, dem, bg=bg); col += 1
            _dat(ws, ri, col, f"=SUM({L(i_b0)}{ri}:{L(i_b1)}{ri})", bg=bg); col += 1
            _dat(ws, ri, col, f"=SUM({L(ds)}{ri}:{L(ds+len(DOENCAS)-1)}{ri})", bg=bg); col += 1

    _auto_width(ws, mn=5, mx=20)
    ws.freeze_panes = f"{get_column_letter(ni+1)}4"


# ---------------------------------------------------------------------------
# Aba FBN
# ---------------------------------------------------------------------------
_FORMAS = [
    ("Caixa da Plantadeira", "Caixa"),
    ("Industrial",           "TIS"),
    ("Misturador de semente","Misturador"),
    ("Betoneira",            "Betoneira"),
    ("Lona",                 "Lona"),
    ("No sulco",             "Sulco"),
]

def _build_fbn(wb, registros):
    ws = wb.create_sheet("FBN")
    ws.sheet_view.showGridLines = False

    nc = 3 + len(_FORMAS) + 2

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
    c = ws.cell(1, 1, "FBN - INOCULACAO - SAFRA 2024/25")
    c.font = Font(name="Arial", bold=True, size=11, color=_BRANCO)
    c.fill = _fill(_AZUL_ESC); c.alignment = _ctr()

    fi0 = 4; fi1 = fi0 + len(_FORMAS) - 1
    _hdr(ws, 2, 1, "Mesorregiao", _AZUL_MED)
    _hdr(ws, 2, 2, "N Produtores", _AZUL_MED)
    _hdr(ws, 2, 3, "Inoculacao (%)", _AZUL_MED)
    _hdr(ws, 2, fi0, "Formas de Inoculacao (%)", _AZUL_MED, fi1)
    _hdr(ws, 2, fi1 + 1, "Coinoculacao (%)", _AZUL_MED)
    _hdr(ws, 2, fi1 + 2, "Co+Mo (%)", _AZUL_MED)

    for col, (_, lbl) in enumerate(_FORMAS, start=fi0):
        c = _hdr(ws, 3, col, lbl, _AZUL_CLA)
        c.font = _hfont(color=_AZUL_ESC, size=8)
    for col in [1, 2, 3, fi1 + 1, fi1 + 2]:
        ws.cell(3, col).border = _brd()

    grp: dict[str, list] = defaultdict(list)
    for reg in registros:
        grp[reg.get("Meso_IDR", "")].append(reg)

    def pct(lst, k, v):
        n = len(lst)
        return round(sum(1 for r in lst if r.get(k) == v) / n, 4) if n else 0

    def pct_f(lst, fv):
        usam = [r for r in lst if r.get("Inoc_Usa") == "SIM"]
        n = len(usam)
        return round(sum(1 for r in usam if r.get("Inoc_Forma") == fv) / n, 4) if n else 0

    row = 4
    for meso in REGIOES_IDR:
        g = grp.get(meso, [])
        bg = _LINHA_PAR if row % 2 == 0 else None
        vals = ([meso, len(g), pct(g, "Inoc_Usa", "SIM")]
                + [pct_f(g, f) for f, _ in _FORMAS]
                + [pct(g, "Coinoc", "SIM"), pct(g, "CoMo_Usa", "SIM")])
        for ci, v in enumerate(vals, start=1):
            _dat(ws, row, ci, v, bg=bg, fmt="0.0%" if ci >= 3 else None)
        row += 1

    vals = (["Media Parana*", len(registros), pct(registros, "Inoc_Usa", "SIM")]
            + [pct_f(registros, f) for f, _ in _FORMAS]
            + [pct(registros, "Coinoc", "SIM"), pct(registros, "CoMo_Usa", "SIM")])
    for ci, v in enumerate(vals, start=1):
        _dat(ws, row, ci, v, bg=_AZUL_CLA, bold=True,
             fmt="0.0%" if ci >= 3 else None)

    _auto_width(ws, mn=10, mx=25)


# ---------------------------------------------------------------------------
# Aba Tto_Sal_CB
# ---------------------------------------------------------------------------
def _build_tto_sal_cb(wb, registros):
    ws = wb.create_sheet("Tto_Sal_CB")
    ws.sheet_view.showGridLines = False

    hdrs = ["N", "Meso_IDR", "Regiao", "Municipio", "Area (ha)", "Cultivar",
            "Bt", "Produtividade", "Data Plantio", "Adversidade", "Sinistro",
            "Trat. Semente", "SAL+Inseticida", "Ctrl. Biologico",
            "Conhec. MID", "Utiliza MID", "Conhec. MIP", "Utiliza MIP",
            "Inoculacao", "Forma Inoc.", "Coinoculacao", "Co+Mo", "Forma Co+Mo"]
    keys = ["N", "Meso_IDR", "Regiao", "Municipio", "Area_Soja", "Cultivar",
            "Bt", "Produtividade", "Dt_Plantio", "Adversidade", "Sinistro",
            "Tto_Semente", "SAL_CB", "Ctrl_Biol",
            "Conhec_MID", "Utiliza_MID", "Conhec_MIP", "Utiliza_MIP",
            "Inoc_Usa", "Inoc_Forma", "Coinoc", "CoMo_Usa", "CoMo_Forma"]

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(hdrs))
    c = ws.cell(1, 1, "TRAT. SEMENTE / SAL / CTRL. BIOLOGICO / INOCULACAO - SAFRA 2024/25")
    c.font = Font(name="Arial", bold=True, size=11, color=_BRANCO)
    c.fill = _fill(_AZUL_ESC); c.alignment = _ctr()

    for ci, h in enumerate(hdrs, start=1):
        c = _hdr(ws, 2, ci, h, _AZUL_MED); c.font = _hfont(size=8)

    for ri, reg in enumerate(registros, start=3):
        bg = _LINHA_PAR if ri % 2 == 0 else None
        for ci, key in enumerate(keys, start=1):
            _dat(ws, ri, ci, reg.get(key), bg=bg)

    _auto_width(ws, mn=8, mx=25)
    ws.freeze_panes = "A3"


# ---------------------------------------------------------------------------
# Abas de frequencia: Lagartas / Percevejos / OP+Acaros / Doencas
# ---------------------------------------------------------------------------
def _build_freq(wb, sheet_name, titulo, registros, grupo):
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False

    freq_vals = list(range(N_PULV + 1))
    n_sub = len(freq_vals) + 2
    nc = 1 + n_sub * 3

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
    c = ws.cell(1, 1, f"FREQUENCIA DE APLICACOES - {titulo.upper()} - SAFRA 2024/25")
    c.font = Font(name="Arial", bold=True, size=11, color=_BRANCO)
    c.fill = _fill(_AZUL_ESC); c.alignment = _ctr()

    grp_m: dict[str, list] = defaultdict(list)
    for reg in registros:
        grp_m[reg.get("Meso_IDR", "")].append(reg)

    cur = 2
    for regiao in ["Parana"] + REGIOES_IDR:
        g   = registros if regiao == "Parana" else grp_m.get(regiao, [])
        bt  = [r for r in g if r.get("Bt") == "SIM"]
        nbt = [r for r in g if r.get("Bt") == "NAO"]
        nt  = len(g)

        ws.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=nc)
        c = ws.cell(cur, 1, regiao)
        c.font = Font(name="Arial", bold=True, size=10, color=_BRANCO)
        c.fill = _fill(_AZUL_MED); c.alignment = _ctr(); cur += 1

        _hdr(ws, cur, 1, "Frequencia / Tecnologia", _AZUL_ESC)
        col = 2
        for fv in freq_vals:
            _hdr(ws, cur, col, str(fv), _AZUL_CLA, col + 2); col += 3
        _hdr(ws, cur, col, "Total c/ Aplicacao", _AZUL_CLA, col + 2); col += 3
        _hdr(ws, cur, col, "Total Levantamentos", _AZUL_CLA, col + 2); cur += 1

        ws.cell(cur, 1).border = _brd()
        col = 2
        for _ in range(n_sub):
            for s in ["Bt", "Nao Bt", "Total"]:
                c = _hdr(ws, cur, col, s, _AZUL_CLA)
                c.font = _hfont(color=_AZUL_ESC, size=8); col += 1
        cur += 1

        cnt_bt  = Counter(_n_aplic(r, grupo) for r in bt)
        cnt_nbt = Counter(_n_aplic(r, grupo) for r in nbt)
        cnt_tot = Counter(_n_aplic(r, grupo) for r in g)

        for lbl, is_pct, cnt_fn in [
            ("N Levantamentos", False,
             lambda fv: (cnt_bt[fv], cnt_nbt[fv], cnt_tot[fv])),
            ("% Imoveis", True,
             lambda fv: (
                 round(cnt_bt[fv]  / len(bt),  4) if bt  else 0,
                 round(cnt_nbt[fv] / len(nbt), 4) if nbt else 0,
                 round(cnt_tot[fv] / nt,       4) if g   else 0,
             )),
        ]:
            _dat(ws, cur, 1, lbl)
            col = 2
            for fv in freq_vals:
                for v in cnt_fn(fv):
                    _dat(ws, cur, col, v, fmt="0.0%" if is_pct else None); col += 1
            _dat(ws, cur, col, sum(1 for r in bt  if _n_aplic(r, grupo) > 0)); col += 1
            _dat(ws, cur, col, sum(1 for r in nbt if _n_aplic(r, grupo) > 0)); col += 1
            _dat(ws, cur, col, sum(1 for r in g   if _n_aplic(r, grupo) > 0)); col += 1
            _dat(ws, cur, col, len(bt)); col += 1
            _dat(ws, cur, col, len(nbt)); col += 1
            _dat(ws, cur, col, nt)
            cur += 1
        cur += 1

    _auto_width(ws, mn=6, mx=16)


# ---------------------------------------------------------------------------
# Funcao principal
# ---------------------------------------------------------------------------
def gerar_excel(registros: list[dict],
                filepath: str = "MesoIDR_Export.xlsx") -> str:
    wb = Workbook()
    wb.remove(wb.active)

    _build_bd(wb, registros)
    _build_medias_geral(wb, registros)
    _build_contagem_pragas(wb, registros)
    _build_contagem_doencas(wb, registros)
    _build_tto_sal_cb(wb, registros)
    _build_fbn(wb, registros)
    _build_freq(wb, "Lagartas",   "Lagartas",               registros, "lagartas")
    _build_freq(wb, "Percevejos", "Percevejos",              registros, "percevejos")
    _build_freq(wb, "OP_Acaros",  "Outras Pragas + Acaros",  registros, "op_acaros")
    _build_freq(wb, "Doencas",    "Doencas",                 registros, "doencas")

    wb.save(filepath)
    return filepath


# ---------------------------------------------------------------------------
# Teste local
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import random
    random.seed(7)

    MESOS = REGIOES_IDR
    MUNICS = {
        "Noroeste": ["Campo Mourao", "Umuarama", "Cianorte"],
        "Norte": ["Londrina", "Maringa", "Cornelio Procopio"],
        "Oeste": ["Cascavel", "Toledo", "Foz do Iguacu"],
        "Sudoeste": ["Pato Branco", "Francisco Beltrao"],
        "Centro Sul": ["Guarapuava", "Irati"],
        "Centro": ["Ponta Grossa", "Castro"],
        "Metropolitana e Litoral": ["Curitiba", "Paranagua"],
    }
    CULTIVARES = [" 50I52 RSF IPRO", " 5400 IPRO", " 5644 IPRO", " 6039 IPRO"]
    FORMAS_INOC = [f for f, _ in _FORMAS]

    def _fake(n):
        d: dict[str, Any] = {}
        if random.random() < max(0.1, 0.85 - n * 0.09):
            ap = random.sample(PRAGAS, random.randint(0, 3))
            ad = random.sample(DOENCAS, random.randint(0, 2))
            cls = []
            if ap: cls.append("Inseticida")
            if ad: cls.append("Fungicida")
            d[f"P{n}_DAE"]  = random.randint(15, 90)
            d[f"P{n}_Data"] = f"2024-{random.randint(10,12):02d}-{random.randint(1,28):02d}"
            al = ap + ad
            for k in range(1, 6):
                d[f"P{n}_Cl{k}"]  = cls[k-1] if k <= len(cls) else None
                d[f"P{n}_Alv{k}"] = al[k-1]  if k <= len(al)  else None
            d[f"P{n}_pragas"]  = {p: p in ap for p in PRAGAS}
            d[f"P{n}_doencas"] = {dc: dc in ad for dc in DOENCAS}
        else:
            d[f"P{n}_DAE"] = d[f"P{n}_Data"] = None
            for k in range(1, 6):
                d[f"P{n}_Cl{k}"] = d[f"P{n}_Alv{k}"] = None
            d[f"P{n}_pragas"]  = {p: False for p in PRAGAS}
            d[f"P{n}_doencas"] = {dc: False for dc in DOENCAS}
        return d

    registros = []
    for i in range(1, 61):
        meso  = random.choice(MESOS)
        munic = random.choice(MUNICS[meso])
        reg: dict[str, Any] = {
            "N": i, "Meso_IDR": meso, "Regiao": meso, "Municipio": munic,
            "Area_Soja": round(random.uniform(50, 900), 1),
            "Cultivar": random.choice(CULTIVARES),
            "Bt": random.choice(["SIM", "NAO"]),
            "Produtividade": round(random.uniform(40, 85), 1),
            "Dt_Plantio": f"2024-10-{random.randint(1,28):02d}",
            "Adversidade": random.choice([None, "Seca", "Granizo"]),
            "Sinistro": random.choice(["SIM", "NAO"]),
            "Conhec_MID": random.choice(["SIM", "NAO"]),
            "Utiliza_MID": random.choice(["SIM", "NAO"]),
            "Conhec_MIP": random.choice(["SIM", "NAO"]),
            "Utiliza_MIP": random.choice(["SIM", "NAO"]),
            "Herb_Cl1": "Herbicida",
            "Herb_Alv1": random.choice(["Folhas largas", "Folhas estreitas"]),
            "Herb_Nap1": random.randint(1, 2),
            "Herb_Cl2": None, "Herb_Alv2": None, "Herb_Nap2": None,
            "Herb_Cl3": None, "Herb_Alv3": None, "Herb_Nap3": None,
            "Dess_Sim": "SIM", "Dess_Dt": "2024-09-20",
            "Dess_Cl": "Herbicida", "Dess_Alv1": "Folhas largas",
            "Dess_Alv2": None, "Dess_Alv3": None,
            "Tto_Semente": random.choice(["SIM", "NAO"]),
            "SAL_CB": random.choice(["SIM", "NAO"]),
            "Ctrl_Biol": random.choice(["SIM", "NAO"]),
            "Inoc_Usa": random.choice(["SIM", "NAO"]),
            "Inoc_Forma": random.choice(FORMAS_INOC),
            "Coinoc": random.choice(["SIM", "NAO"]),
            "CoMo_Usa": random.choice(["SIM", "NAO"]),
            "CoMo_Forma": random.choice(["Via semente", "Foliar", None]),
        }
        for n in range(1, N_PULV + 1):
            reg.update(_fake(n))
        registros.append(reg)

    out = gerar_excel(registros, "/home/claude/MesoIDR_Export.xlsx")
    print(f"Gerado: {out}  ({len(registros)} registros)")
