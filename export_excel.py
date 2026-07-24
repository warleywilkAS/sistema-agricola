"""
export_excel.py
---------------
Gera o Excel de exportacao no formato EXATO do arquivo "Modelo_1.xlsx":

    Aba 1 "BD"        -> dados de referencia gerais (listas fixas:
                          doencas, pragas, cultivares, regionais/municipios,
                          plantas invasoras, classes de produto, etc.)
    Aba 2 "Total_Pr"  -> tabulacao com as 129 colunas, no mesmo layout
                          de cabecalho do modelo, preenchida com os
                          dados reais cadastrados no site.

Nao ha formatacao (cor, negrito, largura, congelar paineis etc.) - apenas
o conteudo, exatamente como solicitado.

Uso no app.py (sem mudancas):

    from export_excel import gerar_excel, orm_para_dict

    @app.route("/exportar_excel")
    def exportar_excel():
        todos = FormularioSoja.query.order_by(FormularioSoja.id).all()
        registros = [orm_para_dict(r) for r in todos]
        filepath = os.path.join("/tmp", "MesoIDR_Export.xlsx")
        gerar_excel(registros, filepath)
        return send_file(filepath, as_attachment=True,
                         download_name="MesoIDR_Exportacao.xlsx",
                         mimetype="application/vnd.openxmlformats-"
                                  "officedocument.spreadsheetml.sheet")

IMPORTANTE: o arquivo "bd_dados.json" (dados fixos da aba BD) precisa estar
na MESMA PASTA deste arquivo export_excel.py.
"""

from __future__ import annotations

import json
import os
from typing import Any

from openpyxl import Workbook

# ---------------------------------------------------------------------------
# Constantes de dominio (usadas apenas para reconhecer "alvo" texto -> praga/
# doenca ao montar o dicionario de cada registro)
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
# Conversao ORM -> dict
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
    d["N"]               = r.id
    d["Numero_Produtor"] = r.numero_produtor
    d["Meso_IDR"]        = r.meso_idr
    d["Regiao"]          = r.regiao
    d["Municipio"]       = r.municipio
    d["Area_Soja"]       = r.area_soja
    d["Cultivar"]        = r.cultivar
    d["Bt"]              = r.bt
    d["Produtividade"]   = r.produtividade_media
    d["Dt_Plantio"]      = r.data_plantio
    d["Adversidade"]     = r.qual_adversidade if r.houve_adversidade == "SIM" else None
    d["Sinistro"]        = r.houve_adversidade

    # Monitoramento
    d["Conhec_MID"]  = r.conhecimento_mid
    d["Utiliza_MID"] = r.utiliza_mid
    d["Conhec_MIP"]  = r.conhecimento_mip
    d["Utiliza_MIP"] = r.utiliza_mip

    # Plantas invasoras (ate 3 categorias hoje: dessecacao / pre / pos;
    # a 4a coluna do modelo fica em branco pois o site nao coleta esse dado)
    herbs = [
        ("Herbicida", r.herbicida_dessecacao_alvo, r.herbicida_dessecacao_aplicacoes),
        ("Herbicida", r.herbicida_pre_alvo,         r.herbicida_pre_aplicacoes),
        ("Herbicida", r.herbicida_pos_alvo,         r.herbicida_pos_aplicacoes),
    ]
    for i, (cl, alv, nap) in enumerate(herbs, start=1):
        d[f"Herb_Cl{i}"]  = cl if alv else None
        d[f"Herb_Alv{i}"] = alv
        d[f"Herb_Nap{i}"] = nap

    # Dessecacao (evento especifico, com data e ate 3 alvos)
    dess = pulvs.get("dessecacao")
    d["Dess_Sim"] = "SIM" if dess else "NAO"
    d["Dess_Dt"]  = dess.data if dess else None
    d["Dess_Cl"]  = dess.classe_produto if dess else None
    dess_alvos = _split(dess.alvo) if dess else []
    for i in range(1, 4):
        d[f"Dess_Alv{i}"] = dess_alvos[i - 1] if i <= len(dess_alvos) else None

    # Pulverizacoes 1-7 (ate 5 classes/alvos por aplicacao)
    for n in range(1, N_PULV + 1):
        obj = pulvs.get(str(n))
        alvos   = _split(obj.alvo)           if obj else []
        classes = _split(obj.classe_produto) if obj else []

        d[f"P{n}_DAE"]  = _dae(n)
        d[f"P{n}_Data"] = obj.data if obj else None

        for k in range(1, 6):
            d[f"P{n}_Cl{k}"]  = classes[k - 1] if k <= len(classes) else None
            d[f"P{n}_Alv{k}"] = alvos[k - 1]   if k <= len(alvos)   else None

    # Outras
    d["Tto_Semente"] = r.tratamento_sementes
    d["SAL_CB"]      = r.sal_mistura
    d["Ctrl_Biol"]   = r.controle_biologico

    # FBN / Inoculacao
    d["Inoc_Usa"]   = r.inoculacao_sementes
    d["Inoc_Forma"] = r.forma_inoculacao
    d["Coinoc"]     = r.coinoculacao
    d["CoMo_Usa"]   = r.co_mo
    d["CoMo_Forma"] = r.co_mo_aplicacao

    return d


# ---------------------------------------------------------------------------
# Aba "BD" - dados de referencia gerais (identico ao Modelo_1.xlsx)
# ---------------------------------------------------------------------------
def _carregar_bd_dados() -> list[list]:
    """Le o arquivo bd_dados.json (deve estar ao lado deste .py)."""
    caminho = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bd_dados.json")
    with open(caminho, "r", encoding="utf-8") as f:
        return json.load(f)


def _build_BD(wb: Workbook):
    ws = wb.create_sheet("BD")
    linhas = _carregar_bd_dados()
    for ri, linha in enumerate(linhas, start=1):
        for ci, val in enumerate(linha, start=1):
            if val is not None:
                ws.cell(row=ri, column=ci, value=val)


# ---------------------------------------------------------------------------
# Layout da aba "Total_Pr" (129 colunas), identico ao Modelo_1.xlsx
# Cada item: (chave_no_dict, texto_do_cabecalho, formato)
# formato: "num" -> numero, "txt"/None -> texto
# ---------------------------------------------------------------------------
def _pulv_group_cols(n: int, texto_dae: str):
    cols = [
        (f"P{n}_DAE",  texto_dae, "num"),
        (f"P{n}_Data", "Data",    "txt"),
    ]
    for k in range(1, 6):
        cols += [
            (f"P{n}_Cl{k}",  "Classe do Produto", "txt"),
            (f"P{n}_Alv{k}", "Alvo",              "txt"),
        ]
    return cols


_TEXTO_DAE = {
    1: "1ª Pulverização   (DAE)",
    2: "2ª Pulverização (DAE)",
    3: "3ª Pulverização   (DAE)",
    4: "4ª Pulverização  (DAE)",
    5: "5ª Pulverização  (DAE)",
    6: "6ª Pulverização  (DAE)",
    7: "7ª Pulverização (DAE)",
}

_TITULO_TOTAL_PR = (
    "PLANILHA TABULAÇÃO DADOS QUESTIONÁRIOS APLICAÇÃO DEFENSIVOS PARA "
    "CONTROLE PRAGAS E DOENÇAS_PR_SAFRA 19_20_V1"
)

# Colunas 1-4: sem "chave" de grupo (label fica direto na linha 3, sem
# subcabecalho na linha 4), exatamente como no modelo.
_COLS_INICIAIS = [
    (None,               None,     None),   # col 1 - sem cabecalho (numero sequencial)
    ("_TABELA",          "Tabela", "txt"),   # col 2
    ("Numero_Produtor",  "N° P",   "txt"),   # col 3
    (None,               "Ordem ", None),    # col 4 - sem subcabecalho (numero sequencial)
]

_ID_COLS = [
    ("Meso_IDR",      "Meso_IDR",                    "txt"),
    ("Regiao",        "Região",                      "txt"),
    ("Municipio",     "Município",                   "txt"),
    ("Area_Soja",     "Área com  Soja (ha)",          "num"),
    ("Cultivar",      "Cultivar",                     "txt"),
    ("Bt",            "Bt",                           "txt"),
    ("Produtividade", "Produtividade Média (sc/ha)",  "num"),
    ("Dt_Plantio",    "Data Plantio",                 "txt"),
    ("Adversidade",   "Adversidade",                  "txt"),
    ("Sinistro",      "Sinistro",                     "txt"),
]

_MID_COLS = [
    ("Conhec_MID",  "Conhec. MID",  "txt"),
    ("Utiliza_MID", "Utiliza MID",  "txt"),
    ("Conhec_MIP",  "Conhec. MIP",  "txt"),
    ("Utiliza_MIP", "Utiliza MIP",  "txt"),
]

# 4 blocos de (Classe do Produto / Alvo / N° Aplicações)
_HERB_COLS = []
for _i in range(1, 5):
    _HERB_COLS += [
        (f"Herb_Cl{_i}",  "Classe do Produto", "txt"),
        (f"Herb_Alv{_i}", "Alvo",              "txt"),
        (f"Herb_Nap{_i}", "N° Aplicações",     "num"),
    ]

_DESS_COLS = [
    ("Dess_Sim",  " Pulverização na Dessecação  ", "txt"),
    ("Dess_Dt",   "Data",                          "txt"),
    ("Dess_Cl",   "Classe do Produto",             "txt"),
    ("Dess_Alv1", "Alvo_1",                        "txt"),
    ("Dess_Alv2", "Alvo_2",                        "txt"),
    ("Dess_Alv3", "Alvo_3",                        "txt"),
]

_OUTRAS_COLS = [
    ("Tto_Semente", "Tratamento de Semente",           "txt"),
    ("SAL_CB",      "Utilização de SAL + Inseticida",  "txt"),
    ("Ctrl_Biol",   "Utilizou Controle Biológico",     "txt"),
]

_INOC_COLS = [
    ("Inoc_Usa",   "Utiliza Inoculação", "txt"),
    ("Inoc_Forma", "Forma Inoculação",   "txt"),
    ("Coinoc",     "Coinoculação",       "txt"),
    ("CoMo_Usa",   "Utiliza Co e Mo",    "txt"),
    ("CoMo_Forma", "Forma Co e Mo",      "txt"),
]

# grupos = (texto_do_grupo_ou_None, lista_de_colunas)
GRUPOS_TOTAL_PR = [
    (None,                                                     _COLS_INICIAIS),
    (None,                                                     _ID_COLS),
    ("CONHECIMENTO MONITORAMENTO",                             _MID_COLS),
    ("3_Informação Plantas Invasoras",                         _HERB_COLS),
    ("4.0_INFORMAÇÃO _PULVERIZAÇÃO DESSECAÇÃO",                _DESS_COLS),
    ("4.1_INFORMAÇÃO _PRIMEIRA PULVERIZAÇÃO APÓS EMERGÊNCIA",  _pulv_group_cols(1, _TEXTO_DAE[1])),
    ("4.2_INFORMAÇÃO _SEGUNDA PULVERIZAÇÃO APÓS EMERGÊNCIA",   _pulv_group_cols(2, _TEXTO_DAE[2])),
    ("4.3_INFORMAÇÃO _TERCEIRA PULVERIZAÇÃO APÓS EMERGÊNCIA",  _pulv_group_cols(3, _TEXTO_DAE[3])),
    ("4.4_INFORMAÇÃO _QUARTA PULVERIZAÇÃO APÓS EMERGÊNCIA",    _pulv_group_cols(4, _TEXTO_DAE[4])),
    ("4.5_INFORMAÇÃO _QUINTA PULVERIZAÇÃO APÓS EMERGÊNCIA",    _pulv_group_cols(5, _TEXTO_DAE[5])),
    ("4.6_INFORMAÇÃO _SEXTA PULVERIZAÇÃO APÓS EMERGÊNCIA",     _pulv_group_cols(6, _TEXTO_DAE[6])),
    ("4.7_INFORMAÇÃO _SÉTIMA PULVERIZAÇÃO APÓS EMERGÊNCIA",    _pulv_group_cols(7, _TEXTO_DAE[7])),
    ("5.OUTRAS INFORMAÇÕES",                                   _OUTRAS_COLS),
    ("6.INOCULAÇÃO",                                           _INOC_COLS),
]

ALL_COLS: list[tuple] = []
for _, cols in GRUPOS_TOTAL_PR:
    ALL_COLS.extend(cols)

# coluna 129 (DY) fica em branco, igual ao modelo
ALL_COLS.append((None, None, None))

_CI: dict[str, int] = {}
for _idx, (key, _label, _fmt) in enumerate(ALL_COLS, start=1):
    if key:
        _CI[key] = _idx


def _build_total_pr(wb: Workbook, registros: list[dict]):
    ws = wb.create_sheet("Total_Pr")
    nc = len(ALL_COLS)

    # Linha 1: titulo (somente na coluna D, igual ao modelo)
    ws.cell(row=1, column=4, value=_TITULO_TOTAL_PR)

    # Linha 3: cabecalhos de grupo (so na primeira coluna de cada grupo)
    # Linha 4: subcabecalhos (uma celula por coluna)
    col = 1
    for grupo_texto, cols in GRUPOS_TOTAL_PR:
        if grupo_texto is not None:
            ws.cell(row=3, column=col, value=grupo_texto)
        for key, label, _fmt in cols:
            if label is not None:
                # colunas iniciais (Tabela/N°P) tem o texto na propria linha 3
                if grupo_texto is None and cols is _COLS_INICIAIS:
                    ws.cell(row=3, column=col, value=label)
                else:
                    ws.cell(row=4, column=col, value=label)
            col += 1

    # Linha 5 fica em branco (igual ao modelo); dados comecam na linha 6
    linha_inicial = 6
    for i, reg in enumerate(registros):
        ri = linha_inicial + i
        ws.cell(row=ri, column=1, value=i + 1)      # col 1 - sequencial
        ws.cell(row=ri, column=2, value="TB1.")      # col 2 - Tabela
        ws.cell(row=ri, column=4, value=i + 1)       # col 4 - Ordem
        for key, _label, _fmt in ALL_COLS:
            if not key or key in ("_TABELA",):
                continue
            ci = _CI.get(key)
            if ci:
                ws.cell(row=ri, column=ci, value=reg.get(key))


# ---------------------------------------------------------------------------
# Funcao principal
# ---------------------------------------------------------------------------
def gerar_excel(registros: list[dict], filepath: str = "MesoIDR_Export.xlsx") -> str:
    wb = Workbook()
    wb.remove(wb.active)

    _build_BD(wb)
    _build_total_pr(wb, registros)

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

    def _fake_reg(i):
        meso = random.choice(MESOS)
        munic = random.choice(MUNICS[meso])
        reg: dict[str, Any] = {
            "N": i, "Numero_Produtor": f"{i:04d}",
            "Meso_IDR": meso, "Regiao": meso, "Municipio": munic,
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
            "Herb_Cl1": "Herbicida", "Herb_Alv1": "Folhas largas", "Herb_Nap1": 1,
            "Dess_Sim": "SIM", "Dess_Dt": "2024-09-20",
            "Dess_Cl": "Herbicida", "Dess_Alv1": "Folhas largas",
            "Tto_Semente": random.choice(["SIM", "NAO"]),
            "SAL_CB": random.choice(["SIM", "NAO"]),
            "Ctrl_Biol": random.choice(["SIM", "NAO"]),
            "Inoc_Usa": random.choice(["SIM", "NAO"]),
            "Inoc_Forma": "Via semente",
            "Coinoc": random.choice(["SIM", "NAO"]),
            "CoMo_Usa": random.choice(["SIM", "NAO"]),
            "CoMo_Forma": random.choice(["Via semente", "Foliar", None]),
        }
        for n in range(1, N_PULV + 1):
            reg[f"P{n}_DAE"] = random.randint(15, 90)
            reg[f"P{n}_Data"] = f"2024-{random.randint(10,12):02d}-{random.randint(1,28):02d}"
            reg[f"P{n}_Cl1"] = "Inseticida"
            reg[f"P{n}_Alv1"] = random.choice(PRAGAS)
        return reg

    registros = [_fake_reg(i) for i in range(1, 21)]
    out = gerar_excel(registros, "/home/claude/out/MesoIDR_Export_teste.xlsx")
    print(f"Gerado: {out} ({len(registros)} registros)")
    out = gerar_excel(registros, "/home/claude/MesoIDR_Export.xlsx")
    print(f"Gerado: {out}  ({len(registros)} registros)")
