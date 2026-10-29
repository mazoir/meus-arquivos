#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
===============================================================================
 build_municipios_ibge.py
===============================================================================
 OBJETIVO
   Gerar a planilha Excel "Municipios_Brasil_IBGE.xlsx" com TODOS os municípios
   brasileiros (5.571, DTB 2025) e seus distritos (10.751), contendo:
     - Código IBGE (7 e 6 dígitos), nome do município, UF, Estado, Grande Região
     - Região Geográfica Imediata e Intermediária (código + nome)
     - Distritos (lista no grão município + aba no grão distrito)
     - Latitude/Longitude da sede municipal (ponto oficial IBGE), com fonte,
       status e observação de auditoria por linha
     - Centroide geométrico e ponto interno do polígono municipal (Malha 2025)
     - Coordenada da sede distrital, quando houver ponto oficial IBGE

 QUANDO USAR
   - Para (re)gerar a base de referência territorial usada em BI/geomarketing.
   - Sempre que o IBGE publicar nova DTB / Localidades / Malha: troque os
     arquivos em --fontes, ajuste o bloco CONFIG (FONTES, ESPERADO) e rode.

 ENTRADAS  (diretório --fontes; layout esperado no dicionário FONTES abaixo)
   dtb2025/RELATORIO_DTB_BRASIL_2025_MUNICIPIOS.ods   (IBGE, original)
   dtb2025/RELATORIO_DTB_BRASIL_2025_DISTRITOS.csv    (IBGE, convertido p/ CSV)
   api_localidades/*.json, *.parquet                  (snapshots da API IBGE)
   localidades_2022/...                               (IBGE Localidades 2022)
   localidades_2010/BR_Localidades_2010_v1.csv        (IBGE Localidades 2010)
   malha_2025/BR_Municipios_2025_simplificada.topojson.gz (IBGE Malha 2025)
   referencia_independente/kelvins_municipios.csv     (só para checagem cruzada)

 SAÍDAS  (diretório --saida)
   Municipios_Brasil_IBGE.xlsx   (abas: LEIA-ME, Municipios, Distritos,
                                  Regioes_Imediatas, Regioes_Intermediarias,
                                  UFs, Resumo_UF, Auditoria_Coordenadas,
                                  Validacao, Fontes)
   build_log.txt                 (log completo da execução)
   Se alguma validação crítica falhar, o arquivo é gravado como
   Municipios_Brasil_IBGE_FALHOU.xlsx (para diagnóstico), o entregável anterior
   NÃO é sobrescrito e o script termina com exit code 1.

 PRÉ-REQUISITOS
   Python >= 3.10 | pandas >= 2.0 | openpyxl >= 3.1 | shapely >= 2.0 | numpy
   pyarrow (opcional: checagem cruzada com o parquet da API)
   LibreOffice (opcional: grava os valores das fórmulas; sem ele o Excel
   recalcula ao abrir)
   pip install pandas openpyxl shapely numpy pyarrow

 EXEMPLO DE EXECUÇÃO
   python build_municipios_ibge.py --fontes ./fontes --saida ./saida

   Exemplo de linha de saída (aba Municipios, colunas principais):
   3106200 | 310620 | Belo Horizonte | Belo Horizonte - MG | MG | 31 |
   Minas Gerais | Sudeste | 3101 | Belo Horizonte | 310001 | Belo Horizonte |
   -19.9237 | -43.9363 | IBGE Localidades 2022 (Sede Municipal) | OK | Sim |
   3 | 310620005 | Belo Horizonte | Belo Horizonte; Barreiro; Venda Nova | ...

 PERFORMANCE / BOAS PRÁTICAS
   - Tudo em memória (≈16 mil linhas): execução típica < 1 min, dominada pelo
     cálculo de polylabel/point-in-polygon com shapely (geometrias preparadas
     + STRtree para busca espacial) e pelo recálculo no LibreOffice.
   - Nenhum código baixado é executado: arquivos ODS/KML/TopoJSON são lidos
     com parsers próprios (zipfile + ElementTree / regex / json).
   - O entregável só é substituído se TODAS as validações críticas passarem.
===============================================================================
"""

from __future__ import annotations

import argparse
import datetime as dt
import difflib
import gzip
import hashlib
import json
import logging
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd
import shapely
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from shapely.geometry import MultiPolygon, Point, Polygon
from shapely.ops import nearest_points, polylabel
from shapely.validation import make_valid

# =============================================================================
# 1. CONFIG
# =============================================================================
ARQUIVO_SAIDA = "Municipios_Brasil_IBGE.xlsx"
ARQUIVO_FALHA = "Municipios_Brasil_IBGE_FALHOU.xlsx"

# Valores oficiais esperados (IBGE). Mudam a cada edição da DTB.
ESPERADO = {
    "municipios": 5571,      # DTB 2025 (inclui DF e Fernando de Noronha)
    "distritos": 10751,      # DTB 2025
    "ufs": 27,
    "reg_imediatas": 510,
    "reg_intermediarias": 133,
}

# --- Saneamento da coordenada da sede municipal (km) ---
LIMIAR_DIVERGENCIA_KM = 5.0          # 2022 é tido como erro se divergir > X km do IBGE 2010 E da base independente
LIMIAR_DIVERGENCIA_CAPITAL_KM = 10.0 # capitais: área urbana extensa, pontos representativos divergem mais
RAIO_LOCALIDADE_KM = 1.5             # evidência: ponto 2022 "grudado" em outra localidade (povoado, vila...)
RAIO_ISOLADO_KM = 10.0               # evidência: ponto 2022 sem nenhuma localidade 2010 num raio de X km
FAIXA_REVISAO_KM = (3.0, 5.0)        # divergência moderada: anotada e auditada, ponto mantido
CONCORDANCIA_REF_KM = 2.5            # IBGE 2010 e base independente "concordam" se a menos de X km
RAIO_VILA_PROXIMA_KM = 0.5           # sede a < X km da sede de OUTRO distrito do município: anotar
MAX_MINUTO_PARA_2010_KM = 3.0        # ponto 2022 arredondado ao minuto: usar 2010 se a menos de X km
TOLERANCIA_POLIGONO_KM = 1.0         # malha simplificada: até ~1 km fora é ruído de borda
LIMIAR_AUDITORIA_KM = 5.0            # linhas com qualquer divergência > X km vão p/ auditoria

# --- Coordenada da sede distrital ---
LIMIAR_VILAS_DIVERGENTES_KM = 3.0    # Vila 2022 x Vila 2010 divergentes
RAIO_HOMONIMO_KM = 1.0               # localidade 2022 homônima perto do ponto 2010 confirma o ponto 2010
LIMIAR_NOME_PARECIDO = 0.6           # difflib ratio: >= é variação de grafia; < é outra localidade

# --- Centroide ---
DISTANCIA_PARTES_CENTROIDE_KM = 50.0  # ilhas oceânicas (ex.: Trindade/Martim Vaz em Vitória) fora do centroide

# Caixa envolvente do Brasil (graus; inclui ilhas oceânicas)
BBOX_BRASIL = {"lat_min": -34.0, "lat_max": 5.5, "lon_min": -74.2, "lon_max": -28.5}

GRANDES_REGIOES = {1: "Norte", 2: "Nordeste", 3: "Sudeste", 4: "Sul", 5: "Centro-Oeste"}

# Correções de caracteres corrompidos nos nomes oficiais (o próprio IBGE publica 9 nomes de
# regiões com '¿' no lugar do travessão; a API de Localidades replica o defeito).
NORMALIZACAO_TEXTO = {" ¿ ": " - ", "¿": "-"}

# Curadoria manual documentada (evidência externa verificada em 29/10/2026). Só altera STATUS e
# acrescenta NOTA; a coordenada continua sendo um ponto oficial do IBGE. 'ref' = coordenada
# publicada por terceiros, usada apenas para medir a distância (calculada pelo pipeline).
CURADORIA = {
    "1100205": {  # Porto Velho/RO
        "status": "Corrigida - a conferir",
        "ref": (-8.761944, -63.903889, "Prefeitura de Porto Velho (8°45'43\"S 63°54'14\"W, centro)"),
        "nota": "O ponto IBGE 2010 fica na Zona Leste da cidade; para análises que exijam o centro, use a "
                "coordenada publicada pela Prefeitura.",
    },
    "1100072": {  # Corumbiara/RO
        "status": "Corrigida",
        "nota": "A base independente aponta para o Povoado Alto Guarajus; o ponto IBGE 2010 é a área urbana "
                "(CIDADE) da sede.",
    },
    "3541000": {  # Praia Grande/SP
        "status": "Corrigida",
        "ref": (-24.0058, -46.4028, "geodatos.net / Cidade Brasil (centro)"),
        "nota": "Referência externa de centro confirma o ponto IBGE 2010.",
    },
    "2107506": {  # Paço do Lumiar/MA
        "status": "Corrigida",
        "ref": (-2.53194, -44.10778, "Wikipedia (sede, Praça N. Sra. da Luz)"),
        "nota": "O ponto IBGE 2022 fica no Maiobão; a sede histórica e a prefeitura ficam junto ao ponto IBGE 2010.",
    },
    "1300706": {  # Boca do Acre/AM
        "status": "Corrigida - a conferir",
        "nota": "Fontes divergem ~7 km; o município tem mais de um núcleo urbano.",
    },
    "2515401": {  # São Vicente do Seridó/PB
        "nota": "IBGE 2022 põe a sede no distrito 10 (São Vicente do Seridó, ~8,7 km ao sul); o Censo 2010 e a "
                "DTB 2025 (distrito-sede 05 = Seridó) mantêm a sede em Seridó.",
    },
    "4322350": {  # União da Serra/RS
        "status": "A conferir",
        "nota": "Na DTB 2025 o distrito-sede (05) chama-se 'São Luiz' e o distrito 20 'União da Serra'; no IBGE "
                "2010 eram 05 'União da Serra' e 20 'Oeste'. O ponto IBGE 2022 fica junto à sede do distrito 20, "
                "~3,2 km do ponto IBGE 2010 da sede.",
    },
    "3550308": {  # São Paulo/SP
        "ref": (-23.5503, -46.6339, "Praça da Sé (marco zero)"),
        "nota": "Ponto representativo da área urbana definido pelo IBGE (2010 e 2022 concordam); não é o marco zero.",
    },
}

# Catálogo de fontes: caminho relativo a --fontes + metadados exibidos na aba Fontes.
FONTES = {
    "dtb_municipios": {
        "arquivo": "dtb2025/RELATORIO_DTB_BRASIL_2025_MUNICIPIOS.ods",
        "produto": "IBGE — Divisão Territorial Brasileira (DTB) 2025 — Relatório de Municípios",
        "safra": "Data-base 31/12/2025 (divulgada pelo IBGE em 30/03/2026)",
        "uso": "PRIMÁRIA: código/nome do município, UF, Região Imediata e Intermediária",
        "origem": "Arquivo original do IBGE (DTB_2025.zip), espelhado em github.com/murilo-mg/rastro-de-territorio "
                  "(referencias/ibge/dtb2025/original.ods)",
        "oficial": "https://www.ibge.gov.br/geociencias/organizacao-do-territorio/estrutura-territorial/"
                   "23701-divisao-territorial-brasileira.html",
    },
    "dtb_distritos": {
        "arquivo": "dtb2025/RELATORIO_DTB_BRASIL_2025_DISTRITOS.csv",
        "produto": "IBGE — DTB 2025 — Relatório de Distritos (RELATORIO_DTB_BRASIL_2025_DISTRITOS.xls convertido para CSV)",
        "safra": "Data-base 31/12/2025",
        "uso": "PRIMÁRIA: código e nome dos 10.751 distritos",
        "origem": "github.com/norteiadev/norteia-brave @24a365ae (data/ibge/ibge_distritos.csv)",
        "oficial": "https://www.ibge.gov.br/geociencias/organizacao-do-territorio/estrutura-territorial/"
                   "23701-divisao-territorial-brasileira.html",
    },
    "api_municipios": {
        "arquivo": "api_localidades/municipalities.json",
        "produto": "IBGE — API de Localidades v1 (/localidades/municipios) — snapshot",
        "safra": "Capturado em 11/09/2026",
        "uso": "CHECAGEM CRUZADA de municípios/regiões; sigla da UF e capital",
        "origem": "npm @br-geo-kit/ibge@1.0.0 (data/2026/tables)",
        "oficial": "https://servicodados.ibge.gov.br/api/docs/localidades",
    },
    "api_regioes": {
        "arquivo": "api_localidades/regions.json",
        "produto": "IBGE — API de Localidades v1 (regiões imediatas/intermediárias) — snapshot",
        "safra": "Capturado em 11/09/2026",
        "uso": "CHECAGEM CRUZADA dos nomes das regiões",
        "origem": "npm @br-geo-kit/ibge@1.0.0",
        "oficial": "https://servicodados.ibge.gov.br/api/docs/localidades",
    },
    "api_ufs": {
        "arquivo": "api_localidades/states.json",
        "produto": "IBGE — API de Localidades v1 (/localidades/estados) — snapshot",
        "safra": "Capturado em 11/09/2026",
        "uso": "Sigla da UF, Grande Região e código da capital",
        "origem": "npm @br-geo-kit/ibge@1.0.0",
        "oficial": "https://servicodados.ibge.gov.br/api/docs/localidades",
    },
    "api_distritos": {
        "arquivo": "api_localidades/distritos_api_ibge_2026-10.parquet",
        "produto": "IBGE — API de Localidades v1 (/localidades/distritos) — snapshot",
        "safra": "Capturado em ~08/10/2026",
        "uso": "CHECAGEM CRUZADA da lista de distritos (código a código, nome a nome)",
        "origem": "github.com/domjerse/duckdb @7a539c4d (uf_municipio.parquet)",
        "oficial": "https://servicodados.ibge.gov.br/api/docs/localidades",
    },
    "sedes_2022_a": {
        "arquivo": "localidades_2022/sedes_municipais_2022_A.csv",
        "produto": "IBGE — Localidades do Brasil 2022 (Censo 2022) — pontos 'Sede Municipal'/'Capital'",
        "safra": "Base Censo 2022 (arquivo IBGE publicado em nov/2025)",
        "uso": "PRIMÁRIA: latitude/longitude da sede municipal",
        "origem": "github.com/luandarod/luandarodrigues @4b5a148d "
                  "(projects/ubs-healthcare-mapping/data/reference/ibge_municipal_seats_2022.csv)",
        "oficial": "https://geoftp.ibge.gov.br/organizacao_do_territorio/estrutura_territorial/localidades/"
                   "Localidades_do_Brasil/2022/Localidades_Brasil_shp.zip",
    },
    "sedes_2022_b": {
        "arquivo": "localidades_2022/sedes_municipais_2022_B.csv",
        "produto": "IBGE — Localidades do Brasil 2022 — sedes municipais (extração independente, por UF)",
        "safra": "Base Censo 2022",
        "uso": "CHECAGEM CRUZADA da extração A (consistência de extração)",
        "origem": "github.com/brasil-em-camadas/brasil-em-camadas.github.io @2e9cbdbd (scripts/dados/sedes_municipais.csv)",
        "oficial": "https://geoftp.ibge.gov.br/organizacao_do_territorio/estrutura_territorial/localidades/"
                   "Localidades_do_Brasil/2022/Localidades_UFs_shp.zip",
    },
    "loc_2022_go": {
        "arquivo": "localidades_2022/vilas/GO",
        "produto": "IBGE — Localidades do Brasil 2022 — KML por município (GO), todas as categorias",
        "safra": "Base Censo 2022",
        "uso": "Sede distrital ('Vila') e checagem independente das sedes municipais (GO)",
        "origem": "github.com/johnerick-py/map_select @69477491 (goias_kml/)",
        "oficial": "https://geoftp.ibge.gov.br/organizacao_do_territorio/estrutura_territorial/localidades/",
    },
    "loc_2022_sc": {
        "arquivo": "localidades_2022/vilas/SC",
        "produto": "IBGE — Localidades do Brasil 2022 — KML por município (SC), todas as categorias",
        "safra": "Base Censo 2022",
        "uso": "Sede distrital ('Vila') e checagem independente das sedes municipais (SC)",
        "origem": "github.com/FelipeFDN/Projeto-ML-Seguranca @f5501208 (data/IBGE/SC/)",
        "oficial": "https://geoftp.ibge.gov.br/organizacao_do_territorio/estrutura_territorial/localidades/",
    },
    "loc_2022_norte": {
        "arquivo": "localidades_2022/vilas/AM_AP_PA_RR",
        "produto": "IBGE — Localidades do Brasil 2022 — GeoJSON (AM, AP, PA, RR), todas as categorias",
        "safra": "Base Censo 2022",
        "uso": "Sede distrital ('Vila') e checagem independente das sedes municipais (AM, AP, PA, RR)",
        "origem": "github.com/Charles-Igeotecnologia/infovias @d9f9e88d (localidades_por_uf/)",
        "oficial": "https://geoftp.ibge.gov.br/organizacao_do_territorio/estrutura_territorial/localidades/",
    },
    "localidades_2010": {
        "arquivo": "localidades_2010/BR_Localidades_2010_v1.csv",
        "produto": "IBGE — Cadastro de Localidades Selecionadas 2010 (BR_Localidades_2010_v1)",
        "safra": "Base Censo 2010",
        "uso": "FALLBACK de sede municipal (pontos 2022 com erro), sede distrital e evidência de localidades",
        "origem": "github.com/GusFurtado/dab_assets @5c1ef16d (data/coordenadas.csv)",
        "oficial": "https://geoftp.ibge.gov.br/organizacao_do_territorio/estrutura_territorial/localidades/"
                   "cadastro_de_localidades_selecionadas_2010/",
    },
    "malha_2025": {
        "arquivo": "malha_2025/BR_Municipios_2025_simplificada.topojson.gz",
        "produto": "IBGE — Malha Municipal 2025 (BR_Municipios_2025), simplificada (Visvalingam 5%, topologia preservada)",
        "safra": "Malha 2025",
        "uso": "Centroide, ponto interno e teste 'ponto dentro do limite municipal'",
        "origem": "CRAN brazilmaps 1.0.0 (github.com/cran/brazilmaps @a945a374, inst/maps/municipality/City_2025.topojson.gz)",
        "oficial": "https://geoftp.ibge.gov.br/organizacao_do_territorio/malhas_territoriais/malhas_municipais/"
                   "municipio_2025/Brasil/BR_Municipios_2025.zip",
    },
    "kelvins": {
        "arquivo": "referencia_independente/kelvins_municipios.csv",
        "produto": "Base comunitária kelvins/municipios-brasileiros (NÃO oficial; fonte das coordenadas não documentada)",
        "safra": "Commit 503e2f70 (10/12/2025)",
        "uso": "SOMENTE CHECAGEM CRUZADA independente (nunca usada como valor final)",
        "origem": "github.com/kelvins/municipios-brasileiros (csv/municipios.csv)",
        "oficial": "-",
    },
}

ROTULO_FONTE = {
    "2022": "IBGE Localidades 2022 (Sede Municipal)",
    "2010": "IBGE Localidades 2010 (Cidade)",
    "vila2022": "IBGE Localidades 2022 (Vila)",
    "vila2010": "IBGE Localidades 2010 (Vila)",
    "hom2022": "IBGE Localidades 2022 (localidade homônima)",
    "hom2010": "IBGE Localidades 2010 (localidade homônima)",
    "nenhuma": "Sem ponto oficial de sede distrital",
}
SUFIXO_NOME_DIFERENTE = " — localidade com outro nome na fonte, a conferir"
CATEGORIA_2010 = {"AUI": "Aglomerado Urbano Isolado", "POVOADO": "Povoado", "NÚCLEO": "Núcleo", "LUGAREJO": "Lugarejo",
                  "ALDEIA INDÍGENA": "Aldeia Indígena", "PROJETO DE ASSENTAMENTO": "Projeto de Assentamento"}

log = logging.getLogger("build")


# =============================================================================
# 2. FUNÇÕES — utilitários
# =============================================================================
def configurar_log(saida: Path) -> None:
    fmt = logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s", "%H:%M:%S")
    log.setLevel(logging.INFO)
    log.handlers.clear()
    for h in (logging.StreamHandler(sys.stdout), logging.FileHandler(saida / "build_log.txt", "w", "utf-8")):
        h.setFormatter(fmt)
        log.addHandler(h)


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    if path.is_dir():  # diretório: hash dos hashes dos arquivos, em ordem
        for f in sorted(p for p in path.rglob("*") if p.is_file()):
            h.update(f.name.encode())
            h.update(sha256(f).encode())
        return h.hexdigest()
    with open(path, "rb") as fh:
        for bloco in iter(lambda: fh.read(1 << 20), b""):
            h.update(bloco)
    return h.hexdigest()


def norm_nome(s) -> str:
    """Normaliza nomes para comparação (sem acento, minúsculo, só [a-z0-9])."""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]", "", s)


def parecido(a, b) -> float:
    return difflib.SequenceMatcher(None, norm_nome(a), norm_nome(b)).ratio()


def haversine_km(lat1, lon1, lat2, lon2):
    """Distância geodésica aproximada (esfera R=6371,0088 km). Aceita escalares ou arrays."""
    lat1, lon1, lat2, lon2 = map(np.radians, (lat1, lon1, lat2, lon2))
    a = np.sin((lat2 - lat1) / 2) ** 2 + np.cos(lat1) * np.cos(lat2) * np.sin((lon2 - lon1) / 2) ** 2
    return 2 * 6371.0088 * np.arcsin(np.sqrt(a))


def cod_str(x) -> str:
    """Normaliza códigos que podem vir como float ('150010710.0') ou int."""
    s = str(x).strip()
    return s[:-2] if s.endswith(".0") else s


def br_int(n) -> str:
    return f"{int(n):,}".replace(",", ".")


def br_num(x, casas=1) -> str:
    return f"{x:.{casas}f}".replace(".", ",")


def br_pct(p, casas=1) -> str:
    return f"{p * 100:.{casas}f}".replace(".", ",") + "%"


def normalizar_texto(s):
    if not isinstance(s, str):
        return s
    for a, b in NORMALIZACAO_TEXTO.items():
        s = s.replace(a, b)
    return s


def dv_ibge_ok(cod7: str) -> bool:
    """Dígito verificador do código de município (pesos 1,2,1,2,1,2; soma dos dígitos dos produtos)."""
    s = sum(sum(int(c) for c in str(int(d) * p)) for d, p in zip(cod7[:6], (1, 2, 1, 2, 1, 2)))
    return (10 - s % 10) % 10 == int(cod7[6])


def arredondado_ao_minuto(x: float) -> bool:
    """True se o valor publicado com 4 casas é um minuto de arco inteiro, arredondado OU truncado
    (ex.: -5,6666 = 5°40'). Calibrado nas 5.571 sedes 2022: marca só Apodi/RN e Catuji/MG."""
    return abs(x - round(x * 60) / 60) <= 1e-4


# =============================================================================
# 3. FUNÇÕES — leitores de fontes (parsers puros; nada é executado)
# =============================================================================
def ler_ods(path: Path) -> list[list[str]]:
    """Lê a 1ª planilha de um .ods (stdlib), expandindo células repetidas."""
    ns = {"table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
          "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0"}
    T = "{%s}" % ns["table"]
    raiz = ET.fromstring(zipfile.ZipFile(path).read("content.xml"))
    tabela = raiz.findall(".//table:table", ns)[0]
    linhas = []
    for tr in tabela.iter(T + "table-row"):
        vals = []
        for tc in tr:
            if tc.tag not in (T + "table-cell", T + "covered-table-cell"):
                continue
            rep = int(tc.get(T + "number-columns-repeated", "1"))
            txt = "\n".join("".join(p.itertext()) for p in tc.findall("text:p", ns))
            vals.extend([txt] * min(rep, 64))  # repetição longa = células vazias de cauda
        while vals and vals[-1] == "":
            vals.pop()
        if vals:
            linhas.append(vals)
    return linhas


def contar_espacos_irregulares(serie: pd.Series) -> int:
    s = serie.dropna().astype(str)
    return int((s.ne(s.str.strip()) | s.str.contains(r"\s{2,}| ", regex=True)).sum())


def carregar_dtb_municipios(path: Path) -> tuple[pd.DataFrame, int]:
    linhas = ler_ods(path)
    data_base = next((l[0] for l in linhas[:6] if "DATA BASE" in l[0].upper()), "")
    i_hdr = next(i for i, l in enumerate(linhas) if "Código Município Completo" in l)
    hdr = linhas[i_hdr]
    dados = [l + [""] * (len(hdr) - len(l)) for l in linhas[i_hdr + 1:] if len(l) >= len(hdr) - 1]
    df = pd.DataFrame(dados, columns=hdr)
    df = df.rename(columns={
        "UF": "cod_uf", "Nome_UF": "estado",
        "Região Geográfica Intermediária": "cod_rgint", "Nome Região Geográfica Intermediária": "nome_rgint",
        "Região Geográfica Imediata": "cod_rgi", "Nome Região Geográfica Imediata": "nome_rgi",
        "Código Município Completo": "cod_mun", "Nome_Município": "municipio",
    })[["cod_uf", "estado", "cod_rgint", "nome_rgint", "cod_rgi", "nome_rgi", "cod_mun", "municipio"]]
    df = df[df["cod_mun"].str.fullmatch(r"\d{7}")].copy()
    irregulares = sum(contar_espacos_irregulares(df[c]) for c in ("estado", "nome_rgint", "nome_rgi", "municipio"))
    for c in df.columns:
        df[c] = df[c].str.strip()
    log.info("DTB 2025 municípios: %d linhas (%s)", len(df), data_base)
    return df, irregulares


def carregar_dtb_distritos(path: Path) -> tuple[pd.DataFrame, int]:
    df = pd.read_csv(path, dtype=str, encoding="utf-8", keep_default_na=False)
    df = df.rename(columns={"distrito_code": "cod_distrito", "nome": "distrito", "ibge_code": "cod_mun",
                            "municipio_nome": "municipio_dtb_dist", "uf": "uf_dtb_dist"})
    irregulares = contar_espacos_irregulares(df["distrito"])
    for c in df.columns:
        df[c] = df[c].str.strip()
    log.info("DTB 2025 distritos: %d linhas", len(df))
    return df, irregulares


def carregar_api(fontes: Path):
    mun = pd.DataFrame(json.load(open(fontes / FONTES["api_municipios"]["arquivo"], encoding="utf-8")))
    reg = pd.DataFrame(json.load(open(fontes / FONTES["api_regioes"]["arquivo"], encoding="utf-8")))
    ufs = pd.DataFrame(json.load(open(fontes / FONTES["api_ufs"]["arquivo"], encoding="utf-8")))
    log.info("API IBGE (snapshot): %d municípios, %d regiões, %d UFs", len(mun), len(reg), len(ufs))
    return mun, reg, ufs


def carregar_api_distritos(path: Path) -> pd.DataFrame | None:
    try:
        df = pd.read_parquet(path)
    except Exception as exc:  # pyarrow ausente: checagem é opcional
        log.warning("Parquet da API de distritos não lido (%s) — checagem cruzada pulada", exc)
        return None
    df["cod_distrito"] = df["codigo_distrito"].astype("int64").astype(str)
    log.info("API IBGE distritos (snapshot): %d linhas", len(df))
    return df


def carregar_sedes_2022(fontes: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    a = pd.read_csv(fontes / FONTES["sedes_2022_a"]["arquivo"], dtype={"ibge_municipio_7": str})
    a = a.rename(columns={"ibge_municipio_7": "cod_mun", "seat_latitude": "lat_2022", "seat_longitude": "lon_2022",
                          "seat_category": "cat_2022", "seat_subcategory": "subcat_2022"})
    b = pd.read_csv(fontes / FONTES["sedes_2022_b"]["arquivo"], dtype={"codigo_ibge": str})
    b = b.rename(columns={"codigo_ibge": "cod_mun", "latitude": "lat_b", "longitude": "lon_b"})
    log.info("Sedes 2022: extração A=%d, extração B=%d", len(a), len(b))
    return a, b


def carregar_localidades_2010(path: Path) -> pd.DataFrame:
    """Todas as localidades IBGE 2010 (CIDADE, VILA, POVOADO, AUI, ...)."""
    df = pd.read_csv(path, sep=";", dtype=str, encoding="utf-8")
    out = pd.DataFrame({
        "cod_mun": df["CD_GEOCODMU"].map(cod_str),
        "cod_distrito": df["CD_GEOCODDS"].map(cod_str),
        "categoria": df["NM_CATEGORIA"].str.strip(),
        "nome": df["NM_LOCALIDADE"].str.strip(),
        "tipo": df["TIPO"].str.strip(),
        "lat": df["LAT"].astype(float),
        "lon": df["LONG"].astype(float),
    })
    log.info("Localidades 2010: %d registros | %s", len(out), dict(Counter(out["categoria"]).most_common(4)))
    return out


_RE_PLACEMARK = re.compile(r"<Placemark.*?</Placemark>", re.S)
_RE_SIMPLEDATA = re.compile(r'<SimpleData name="([^"]+)">([^<]*)</SimpleData>')
_RE_COORD = re.compile(r"<Point>.*?<coordinates>([^<]*)</coordinates>", re.S)


def carregar_localidades_2022_parciais(fontes: Path) -> pd.DataFrame:
    """Todas as localidades IBGE 2022 das UFs com espelho por município (GO, SC, AM, AP, PA, RR).
    Usa a GEOMETRIA (precisão total), não os atributos arredondados a 4 casas."""
    regs = []
    for chave in ("loc_2022_go", "loc_2022_sc"):
        for kml in sorted((fontes / FONTES[chave]["arquivo"]).glob("*.kml")):
            for pm in _RE_PLACEMARK.findall(kml.read_text(encoding="utf-8")):
                d = dict(_RE_SIMPLEDATA.findall(pm))
                m = _RE_COORD.search(pm)
                if "CT_LOCALIDADE" not in d or not m:
                    continue
                lon, lat = (float(v) for v in m.group(1).strip().split(",")[:2])
                regs.append((cod_str(d.get("CD_MUN", "")), cod_str(d.get("CD_LOCALIDADE", "")), d["CT_LOCALIDADE"],
                             d.get("SCT_LOCALIDADE", ""), d.get("NM_LOCALIDADE", ""), lat, lon))
    for gj in sorted((fontes / FONTES["loc_2022_norte"]["arquivo"]).glob("*.geojson")):
        for f in json.load(open(gj, encoding="utf-8"))["features"]:
            p = f["properties"]
            if not f.get("geometry") or "CT_LOCALIDADE" not in p:
                continue
            lon, lat = f["geometry"]["coordinates"][:2]
            regs.append((cod_str(p.get("CD_MUN", "")), cod_str(p.get("CD_LOCALIDADE", "")), p["CT_LOCALIDADE"],
                         p.get("SCT_LOCALIDADE") or "", p.get("NM_LOCALIDADE", ""), float(lat), float(lon)))
    df = pd.DataFrame(regs, columns=["cod_mun", "cod_localidade", "categoria", "subcategoria", "nome", "lat", "lon"])
    log.info("Localidades 2022 (GO, SC, AM, AP, PA, RR): %d registros | %d municípios | Vilas=%d", len(df),
             df["cod_mun"].nunique(), int((df["categoria"] == "Vila").sum()))
    return df


def carregar_kelvins(path: Path) -> pd.DataFrame:
    k = pd.read_csv(path, dtype={"codigo_ibge": str})
    return k.rename(columns={"codigo_ibge": "cod_mun", "latitude": "lat_k", "longitude": "lon_k"})[
        ["cod_mun", "lat_k", "lon_k"]]


# ---------------- malha (TopoJSON) ----------------
def _decodificar_arcos(arcos, transform):
    sx, sy = transform["scale"]
    tx, ty = transform["translate"]
    saida = []
    for arco in arcos:
        x = y = 0
        pts = []
        for p in arco:  # delta-encoding (TopoJSON quantizado)
            x += p[0]
            y += p[1]
            pts.append((x * sx + tx, y * sy + ty))
        saida.append(pts)
    return saida


def _anel(arcos, refs):
    pts = []
    for r in refs:
        seq = arcos[~r][::-1] if r < 0 else arcos[r]
        pts.extend(seq[1:] if pts else seq)
    return pts


def carregar_malha(path: Path) -> dict[str, shapely.Geometry]:
    topo = json.load(gzip.open(path))
    arcos = _decodificar_arcos(topo["arcs"], topo["transform"])
    obj = topo["objects"][next(iter(topo["objects"]))]
    geoms = {}
    for g in obj["geometries"]:
        cod = str(g["properties"]["municipality_code"])
        if g["type"] == "Polygon":
            aneis = [_anel(arcos, r) for r in g["arcs"]]
            geom = Polygon(aneis[0], aneis[1:])
        else:
            polys = []
            for p in g["arcs"]:
                aneis = [_anel(arcos, r) for r in p]
                polys.append(Polygon(aneis[0], aneis[1:]))
            geom = MultiPolygon(polys)
        if not geom.is_valid:
            geom = make_valid(geom)
            if geom.geom_type == "GeometryCollection":
                partes = [q for q in geom.geoms if q.geom_type in ("Polygon", "MultiPolygon")]
                geom = shapely.union_all(partes)
        geoms[cod] = geom
    log.info("Malha 2025: %d polígonos", len(geoms))
    return geoms


# =============================================================================
# 4. FUNÇÕES — regras de negócio
# =============================================================================
R_TERRA = 6371.0088


def _laea(lon0, lat0):
    """Projeção azimutal equivalente de Lambert local (área/centroide corretos)."""
    l0, p0 = math.radians(lon0), math.radians(lat0)

    def fwd(c):
        lam = np.radians(c[:, 0]) - l0
        phi = np.radians(c[:, 1])
        k = np.sqrt(2 / (1 + math.sin(p0) * np.sin(phi) + math.cos(p0) * np.cos(phi) * np.cos(lam)))
        return np.column_stack([R_TERRA * k * np.cos(phi) * np.sin(lam),
                                R_TERRA * k * (math.cos(p0) * np.sin(phi) - math.sin(p0) * np.cos(phi) * np.cos(lam))])

    def inv(x, y):
        rho = math.hypot(x, y)
        if rho == 0:
            return lon0, lat0
        c = 2 * math.asin(rho / (2 * R_TERRA))
        phi = math.asin(math.cos(c) * math.sin(p0) + y * math.sin(c) * math.cos(p0) / rho)
        lam = l0 + math.atan2(x * math.sin(c), rho * math.cos(p0) * math.cos(c) - y * math.sin(p0) * math.sin(c))
        return math.degrees(lam), math.degrees(phi)

    return fwd, inv


def derivar_pontos_malha(geoms: dict) -> tuple[pd.DataFrame, list]:
    """Centroide (equivalente de área) e ponto interno (polylabel, garantido dentro do polígono).
    Partes a mais de DISTANCIA_PARTES_CENTROIDE_KM da maior parte (ilhas oceânicas) ficam fora do centroide."""
    regs, com_ilhas = [], []
    for cod, g in geoms.items():
        maior_geo = max(g.geoms, key=lambda q: q.area) if g.geom_type == "MultiPolygon" else g
        c0 = maior_geo.centroid
        fwd, inv = _laea(c0.x, c0.y)
        gp = shapely.transform(g, fwd)
        partes = list(gp.geoms) if gp.geom_type == "MultiPolygon" else [gp]
        maior = max(partes, key=lambda q: q.area)
        perto = [q for q in partes if q.distance(maior) <= DISTANCIA_PARTES_CENTROIDE_KM]
        if len(perto) < len(partes):
            com_ilhas.append(cod)
        base = shapely.union_all(perto) if len(perto) > 1 else maior
        cx, cy = inv(base.centroid.x, base.centroid.y)
        pl = polylabel(maior, tolerance=0.01)  # 10 m
        px, py = inv(pl.x, pl.y)
        regs.append({"cod_mun": cod, "lat_centroide": round(cy, 6), "lon_centroide": round(cx, 6),
                     "centroide_dentro": bool(g.covers(Point(cx, cy))),
                     "lat_interno": round(py, 6), "lon_interno": round(px, 6),
                     "interno_dentro": bool(g.covers(Point(px, py)))})
    return pd.DataFrame(regs), com_ilhas


@dataclass
class TestePoligono:
    dentro: bool
    dist_km: float          # 0 se dentro; senão, distância até o próprio polígono
    cai_em: str             # código do município onde o ponto cai ('' se nenhum)

    def aceito(self) -> bool:
        return self.dentro or (not math.isnan(self.dist_km) and self.dist_km <= TOLERANCIA_POLIGONO_KM)


class IndiceEspacial:
    def __init__(self, geoms: dict):
        self.codigos = list(geoms)
        self.geoms = [geoms[c] for c in self.codigos]
        for g in self.geoms:
            shapely.prepare(g)
        self.arvore = shapely.STRtree(self.geoms)
        self.por_cod = geoms

    def testar(self, cod: str, lat: float, lon: float) -> TestePoligono:
        pt = Point(lon, lat)
        g = self.por_cod.get(cod)
        if g is not None and g.covers(pt):
            return TestePoligono(True, 0.0, cod)
        cand = self.arvore.query(pt, predicate="intersects")
        cai_em = self.codigos[cand[0]] if len(cand) else ""
        dist = float("nan")
        if g is not None:
            q = nearest_points(g, pt)[0]
            dist = float(haversine_km(lat, lon, q.y, q.x))
        return TestePoligono(False, dist, cai_em)


class IndiceLocalidades:
    """Busca da localidade IBGE 2010 mais próxima de um ponto (evidência para o saneamento)."""

    def __init__(self, loc: pd.DataFrame):
        self.loc = loc.reset_index(drop=True)
        self.lat = self.loc["lat"].to_numpy()
        self.lon = self.loc["lon"].to_numpy()

    def mais_proxima(self, lat, lon, excluir_mask=None):
        d = haversine_km(lat, lon, self.lat, self.lon)
        if excluir_mask is not None:
            d = np.where(excluir_mask, np.inf, d)
        i = int(np.argmin(d))
        return self.loc.iloc[i], float(d[i])


def escolher_coordenada_sede(row, limiar: float) -> tuple[float, float, str, str, str]:
    """Regra de saneamento (documentada na aba LEIA-ME):
    - Base: IBGE Localidades 2022 (ponto oficial mais recente).
    - Se o ponto 2022 divergir > limiar do IBGE 2010 E da base independente (duas fontes
      independentes concordam que 2022 está fora), usa IBGE 2010.
    - Se só existir UMA das duas referências e ela divergir > limiar, mantém 2022 e marca 'A conferir'.
    Retorna (lat, lon, chave_fonte, status, observação)."""
    d_22_10, d_22_k = row["km_2022_2010"], row["km_2022_kelvins"]
    if pd.isna(row["lat_2022"]):
        if pd.notna(row["lat_2010"]):
            return row["lat_2010"], row["lon_2010"], "2010", "Corrigida", "Sem ponto 2022; usado IBGE 2010"
        return np.nan, np.nan, "", "Sem coordenada", "Nenhuma fonte oficial"
    tem10, temk = pd.notna(d_22_10), pd.notna(d_22_k)
    if tem10 and temk and d_22_10 > limiar and d_22_k > limiar:
        obs = (f"Ponto IBGE 2022 ({row['lat_2022']:.4f}; {row['lon_2022']:.4f}) divergia {br_num(d_22_10)} km do "
               f"IBGE 2010 e {br_num(d_22_k)} km da base independente; substituído pelo IBGE 2010")
        return row["lat_2010"], row["lon_2010"], "2010", "Corrigida", obs
    if tem10 != temk:
        unica = d_22_10 if tem10 else d_22_k
        nome = "IBGE 2010" if tem10 else "a base independente"
        if unica > limiar:
            return (row["lat_2022"], row["lon_2022"], "2022", "A conferir",
                    f"Só há uma referência ({nome}) e ela diverge {br_num(unica)} km do ponto IBGE 2022")
    return row["lat_2022"], row["lon_2022"], "2022", "OK", ""


def juntar_obs(*partes) -> str:
    return "; ".join(p for p in partes if isinstance(p, str) and p)


# =============================================================================
# 5. FUNÇÕES — validações (testes)
# =============================================================================
class Validador:
    def __init__(self):
        self.resultados = []  # (grupo, verificação, esperado, obtido, ok, crítico)

    def checar(self, grupo, nome, esperado, obtido, ok=None, critico=True):
        ok = (esperado == obtido) if ok is None else bool(ok)
        self.resultados.append((grupo, nome, esperado, obtido, ok, critico))
        (log.info if ok else (log.error if critico else log.warning))(
            "[%s] %s | esperado=%s | obtido=%s | %s", grupo, nome, esperado, obtido, "OK" if ok else ("FALHA" if critico else "ALERTA"))
        return ok

    def falhas_criticas(self):
        return [r for r in self.resultados if r[5] and not r[4]]


# =============================================================================
# 6. FUNÇÕES — escrita do Excel e pós-processamento
# =============================================================================
FMT_COORD = "0.0000##"
FMT_INT = "0"
COR_CABECALHO = "1F4E78"
FONTE_DADOS = Font(name="Arial", size=10)
FONTE_CAB = Font(name="Arial", bold=True, color="FFFFFF", size=10)
AUTOR_COMENTARIO = "Pipeline IBGE"


def letra(df: pd.DataFrame, coluna: str) -> str:
    return get_column_letter(list(df.columns).index(coluna) + 1)


def _valor_celula(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        return float(v)
    return v


def escrever_tabela(ws, df: pd.DataFrame, nome_tabela: str, larguras: dict, formatos: dict,
                    comentarios: dict | None = None, freeze: str = "A2", estilo="TableStyleMedium2"):
    ws.append(list(df.columns))
    for linha in df.itertuples(index=False, name=None):
        ws.append([_valor_celula(v) for v in linha])
    n_lin, n_col = len(df) + 1, len(df.columns)
    tab = Table(displayName=nome_tabela, ref=f"A1:{get_column_letter(n_col)}{n_lin}")
    tab.tableStyleInfo = TableStyleInfo(name=estilo, showRowStripes=True)
    ws.add_table(tab)
    for j, col in enumerate(df.columns, start=1):
        ws.column_dimensions[get_column_letter(j)].width = larguras.get(col, 14)
        hc = ws.cell(row=1, column=j)
        hc.font = FONTE_CAB
        hc.fill = PatternFill("solid", fgColor=COR_CABECALHO)
        hc.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        if comentarios and col in comentarios:
            hc.comment = Comment(comentarios[col], AUTOR_COMENTARIO, width=320, height=140)
        fmt = formatos.get(col)
        for (c,) in ws.iter_rows(min_row=2, max_row=n_lin, min_col=j, max_col=j):
            c.font = FONTE_DADOS
            if fmt:
                c.number_format = fmt
    ws.row_dimensions[1].height = 42
    ws.freeze_panes = freeze
    return n_lin


def escrever_texto(ws, linhas: list[tuple], largura_a=28, largura_b=120):
    """Linhas no formato (texto_col_A, texto_col_B, estilo) com estilo em {'titulo','secao','normal'}."""
    ws.column_dimensions["A"].width = largura_a
    ws.column_dimensions["B"].width = largura_b
    for i, (a, b, estilo) in enumerate(linhas, start=1):
        ca, cb = ws.cell(row=i, column=1, value=a), ws.cell(row=i, column=2, value=b)
        for c in (ca, cb):
            c.alignment = Alignment(wrap_text=True, vertical="top")
        if estilo == "titulo":
            ca.font = Font(name="Arial", bold=True, size=16, color=COR_CABECALHO)
        elif estilo == "secao":
            for c in (ca, cb):
                c.font = Font(name="Arial", bold=True, size=11, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor=COR_CABECALHO)
        else:
            ca.font = Font(name="Arial", bold=True, size=10)
            cb.font = FONTE_DADOS
    ws.sheet_view.showGridLines = False


def recalcular_com_libreoffice(xlsx: Path) -> bool:
    """Recalcula as fórmulas (grava valores em cache) via LibreOffice headless, se instalado.
    Sem isso o Excel recalcula sozinho ao abrir (fullCalcOnLoad), mas leitores como pandas/Power Query
    veriam as células de fórmula vazias."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        log.warning("LibreOffice não encontrado: fórmulas serão calculadas pelo Excel ao abrir o arquivo")
        return False
    with tempfile.TemporaryDirectory() as tmp:
        saida = Path(tmp, "out")
        cmd = [soffice, "--headless", f"-env:UserInstallation={Path(tmp, 'perfil').as_uri()}", "--calc",
               "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(saida), str(xlsx)]
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=900, check=False)
        gerado = saida / xlsx.name
        if proc.returncode != 0 or not gerado.exists():
            log.warning("Recalculo via LibreOffice falhou (rc=%s): %s", proc.returncode, proc.stderr[-500:])
            return False
        shutil.move(str(gerado), str(xlsx))
    log.info("Fórmulas recalculadas via LibreOffice")
    return True


def pos_processar_xml(xlsx: Path, estilo: str = "TableStyleMedium2") -> dict:
    """Ajustes que o LibreOffice desfaz ao regravar: estilo das tabelas (<tableStyleInfo>),
    recálculo ao abrir (fullCalcOnLoad) e autor dos comentários."""
    tag = (f'<tableStyleInfo name="{estilo}" showFirstColumn="0" showLastColumn="0" '
           f'showRowStripes="1" showColumnStripes="0"/></table>').encode()
    tmp = xlsx.with_name(xlsx.stem + "_tmp.xlsx")
    cont = Counter()
    with zipfile.ZipFile(xlsx) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            dados = zin.read(item.filename)
            if re.fullmatch(r"xl/tables/table\d+\.xml", item.filename) and b"tableStyleInfo" not in dados:
                dados = dados.replace(b"</table>", tag)
                cont["tabelas"] += 1
            elif item.filename == "xl/workbook.xml" and b"fullCalcOnLoad" not in dados:
                dados, n = re.subn(rb"<calcPr\b", b'<calcPr fullCalcOnLoad="1"', dados, count=1)
                cont["calcPr"] += n
            elif re.fullmatch(r"xl/comments\d+\.xml", item.filename):
                dados = dados.replace(b"<author>Unknown Author</author>", f"<author>{AUTOR_COMENTARIO}</author>".encode())
                cont["comentarios"] += 1
            zout.writestr(item, dados)
    tmp.replace(xlsx)
    return dict(cont)


def checar_formulas(xlsx: Path, aba_validacao: str, col_resultado: int) -> tuple[int, list, list]:
    """Após recálculo: (nº de fórmulas, células com erro/sem valor, checagens 'FALHA' na aba de validação)."""
    erros = ("#N/A", "#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!")
    wb_f, wb_v = load_workbook(xlsx), load_workbook(xlsx, data_only=True)
    n_form, ruins = 0, []
    for ws in wb_f.worksheets:
        wv = wb_v[ws.title]
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    n_form += 1
                    v = wv[c.coordinate].value
                    if v is None or (isinstance(v, str) and v in erros):
                        ruins.append(f"{ws.title}!{c.coordinate}={v}")
    falhas = [f"{aba_validacao}!{get_column_letter(col_resultado)}{r}"
              for r in range(2, wb_v[aba_validacao].max_row + 1)
              if wb_v[aba_validacao].cell(row=r, column=col_resultado).value == "FALHA"]
    return n_form, ruins, falhas


# =============================================================================
# 7. MAIN
# =============================================================================
def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--fontes", required=True, type=Path, help="diretório com as fontes (layout em FONTES)")
    ap.add_argument("--saida", required=True, type=Path, help="diretório de saída")
    ap.add_argument("--sem-recalculo", action="store_true",
                    help="não recalcular fórmulas via LibreOffice (o Excel recalcula ao abrir)")
    args = ap.parse_args(argv)
    args.saida.mkdir(parents=True, exist_ok=True)
    configurar_log(args.saida)
    F = args.fontes
    V = Validador()
    gerado_em = dt.datetime.now().strftime("%d/%m/%Y %H:%M")
    log.info("Início do build | fontes=%s | saída=%s", F, args.saida)

    # ---------- 7.1 Carga ----------
    dtb, irreg_mun = carregar_dtb_municipios(F / FONTES["dtb_municipios"]["arquivo"])
    dist, irreg_dist = carregar_dtb_distritos(F / FONTES["dtb_distritos"]["arquivo"])
    api_mun, api_reg, api_ufs = carregar_api(F)
    api_dist = carregar_api_distritos(F / FONTES["api_distritos"]["arquivo"])
    s22a, s22b = carregar_sedes_2022(F)
    loc10 = carregar_localidades_2010(F / FONTES["localidades_2010"]["arquivo"])
    loc22 = carregar_localidades_2022_parciais(F)
    kel = carregar_kelvins(F / FONTES["kelvins"]["arquivo"])
    malha = carregar_malha(F / FONTES["malha_2025"]["arquivo"])

    c10 = (loc10[loc10["categoria"] == "CIDADE"].drop_duplicates("cod_mun")
           .rename(columns={"lat": "lat_2010", "lon": "lon_2010", "cod_distrito": "cod_distrito_2010"}))
    v10 = loc10[loc10["categoria"] == "VILA"]
    n_dup10 = int(v10["cod_distrito"].duplicated().sum())
    v10 = v10.drop_duplicates("cod_distrito").set_index("cod_distrito")
    v22 = loc22[(loc22["categoria"] == "Vila") & loc22["cod_localidade"].str.fullmatch(r"\d{9}")]
    n_dup22 = int(v22["cod_localidade"].duplicated().sum())
    v22 = v22.drop_duplicates("cod_localidade").set_index("cod_localidade")
    if n_dup10 or n_dup22:
        log.warning("Vilas duplicadas por distrito (mantida a 1ª): 2010=%d, 2022=%d", n_dup10, n_dup22)
    ufs_loc22 = set(loc22["cod_mun"].str[:2])

    # ---------- 7.2 Normalização de texto (defeito de codificação no próprio IBGE) ----------
    nomes_corrigidos = {}
    for col in ("nome_rgi", "nome_rgint"):
        for v in dtb[col].unique():
            nv = normalizar_texto(v)
            if nv != v:
                nomes_corrigidos[v] = nv
        dtb[col] = dtb[col].map(normalizar_texto)
    api_reg["name"] = api_reg["name"].map(normalizar_texto)
    log.info("Nomes de região normalizados ('¿' -> '-'): %d — %s", len(nomes_corrigidos), list(nomes_corrigidos.values()))

    # ---------- 7.3 Municípios + hierarquia ----------
    ufs = api_ufs.rename(columns={"code": "cod_uf", "uf": "uf", "name": "estado_api", "regionName": "regiao",
                                  "capital": "cod_capital"})[["cod_uf", "uf", "estado_api", "regiao", "cod_capital"]]
    mun = dtb.merge(ufs, on="cod_uf", how="left", validate="many_to_one")
    mun["cod_regiao"] = mun["cod_uf"].str[0].astype(int)

    V.checar("Contagem", "Municípios na DTB 2025", ESPERADO["municipios"], len(mun))
    V.checar("Contagem", "Códigos de município únicos", len(mun), int(mun["cod_mun"].nunique()))
    V.checar("Contagem", "UFs", ESPERADO["ufs"], int(mun["cod_uf"].nunique()))
    V.checar("Contagem", "Regiões Imediatas", ESPERADO["reg_imediatas"], int(mun["cod_rgi"].nunique()))
    V.checar("Contagem", "Regiões Intermediárias", ESPERADO["reg_intermediarias"], int(mun["cod_rgint"].nunique()))
    V.checar("Hierarquia", "Código do município começa com o código da UF", 0,
             int((mun["cod_mun"].str[:2] != mun["cod_uf"]).sum()))
    V.checar("Hierarquia", "Código da Reg. Imediata começa com o código da UF", 0,
             int((mun["cod_rgi"].str[:2] != mun["cod_uf"]).sum()))
    V.checar("Hierarquia", "Código da Reg. Intermediária começa com o código da UF", 0,
             int((mun["cod_rgint"].str[:2] != mun["cod_uf"]).sum()))
    V.checar("Hierarquia", "Cada Reg. Imediata pertence a uma única Reg. Intermediária", 0,
             int((mun.groupby("cod_rgi")["cod_rgint"].nunique() > 1).sum()))
    V.checar("Hierarquia", "Grande Região (API) coerente com 1º dígito da UF", 0,
             int((mun["cod_regiao"].map(GRANDES_REGIOES) != mun["regiao"]).sum()))
    V.checar("Hierarquia", "Nome da UF (DTB) = nome da UF (API)", 0, int((mun["estado"] != mun["estado_api"]).sum()))
    V.checar("Texto", "Nomes (DTB municípios) com espaços irregulares na fonte", 0, irreg_mun)
    dv_excecoes = sorted(c for c in mun["cod_mun"] if not dv_ibge_ok(c))
    V.checar("Hierarquia", "Códigos oficiais que não satisfazem o dígito verificador (exceções conhecidas do IBGE)",
             "informativo", len(dv_excecoes), ok=True, critico=False)

    api_mun = api_mun.rename(columns={"code": "cod_mun"})
    rn = {(k, c): n for k, c, n in zip(api_reg["kind"], api_reg["code"], api_reg["name"])}
    x = mun.merge(api_mun, on="cod_mun", how="outer", indicator=True)
    V.checar("DTB x API", "Municípios só em uma das fontes", 0, int((x["_merge"] != "both").sum()))
    x = x[x["_merge"] == "both"]
    V.checar("DTB x API", "Nome do município divergente", 0, int((x["municipio"] != x["name"]).sum()))
    V.checar("DTB x API", "UF divergente", 0, int((x["uf_x"] != x["uf_y"]).sum()))
    V.checar("DTB x API", "Código Reg. Imediata divergente", 0, int((x["cod_rgi"] != x["immediate"]).sum()))
    V.checar("DTB x API", "Código Reg. Intermediária divergente", 0, int((x["cod_rgint"] != x["intermediate"]).sum()))
    V.checar("DTB x API", "Nome Reg. Imediata divergente", 0,
             int((x["nome_rgi"] != [rn.get(("immediate", c)) for c in x["cod_rgi"]]).sum()))
    V.checar("DTB x API", "Nome Reg. Intermediária divergente", 0,
             int((x["nome_rgint"] != [rn.get(("intermediate", c)) for c in x["cod_rgint"]]).sum()))
    capitais_api = set(api_mun.loc[api_mun["capital"], "cod_mun"])
    mun["capital"] = np.where(mun["cod_mun"] == mun["cod_capital"], "Sim", "Não")
    V.checar("Hierarquia", "Capitais (26 estaduais + Brasília)", 27, int((mun["capital"] == "Sim").sum()))
    V.checar("DTB x API", "Capitais divergentes da flag da API", 0,
             len(set(mun.loc[mun.capital == "Sim", "cod_mun"]) ^ capitais_api))
    nome_mun = dict(zip(mun["cod_mun"], mun["municipio"]))
    uf_mun = dict(zip(mun["cod_mun"], mun["uf"]))
    eh_capital = dict(zip(mun["cod_mun"], mun["capital"] == "Sim"))

    # ---------- 7.4 Distritos ----------
    dist["sufixo"] = dist["cod_distrito"].str[7:].astype(int)
    V.checar("Distritos", "Distritos na DTB 2025", ESPERADO["distritos"], len(dist))
    V.checar("Distritos", "Códigos de distrito únicos", len(dist), int(dist["cod_distrito"].nunique()))
    V.checar("Distritos", "Código do distrito com 9 dígitos", 0, int((~dist["cod_distrito"].str.fullmatch(r"\d{9}")).sum()))
    V.checar("Distritos", "Código do distrito começa com o código do município", 0,
             int((dist["cod_distrito"].str[:7] != dist["cod_mun"]).sum()))
    V.checar("Distritos", "Municípios sem distrito", 0, len(set(mun["cod_mun"]) - set(dist["cod_mun"])))
    V.checar("Distritos", "Distritos com município inexistente", 0, len(set(dist["cod_mun"]) - set(mun["cod_mun"])))
    V.checar("Distritos", "Nome do município no relatório de distritos = relatório de municípios", 0,
             int((dist["municipio_dtb_dist"] != dist["cod_mun"].map(nome_mun)).sum()))
    V.checar("Texto", "Nomes de distrito com espaços irregulares na fonte", 0, irreg_dist)
    if api_dist is not None:
        a = dist.merge(api_dist[["cod_distrito", "distrito_bairro", "municipio", "uf"]], on="cod_distrito",
                       how="outer", indicator=True)
        V.checar("DTB x API", "Distritos só em uma das fontes (API out/2026)", 0, int((a["_merge"] != "both").sum()))
        a = a[a["_merge"] == "both"]
        V.checar("DTB x API", "Nome de distrito divergente (API out/2026)", 0, int((a["distrito"] != a["distrito_bairro"]).sum()))
        V.checar("DTB x API", "Município do distrito divergente (API out/2026)", 0,
                 int((a["municipio_dtb_dist"] != a["municipio"]).sum()))
        V.checar("DTB x API", "UF do distrito divergente ou nula na API (out/2026)", 0,
                 int((a["uf_dtb_dist"] != a["uf"]).sum()), critico=False)

    # Distrito-sede: convenção IBGE = sufixo 05. Exceção: São Paulo (distritos 01..96, sem sede codificada).
    sem_sede_codificada = set(dist.loc[dist["sufixo"] < 5, "cod_mun"])
    dist["distrito_sede"] = np.where(dist["cod_mun"].isin(sem_sede_codificada), "Não se aplica",
                                     np.where(dist["sufixo"] == 5, "Sim", "Não"))
    log.info("Municípios sem distrito-sede codificado (sufixos 01-04 em uso): %s", sorted(sem_sede_codificada))
    V.checar("Distritos", "Municípios com exatamente 1 distrito-sede (sufixo 05), exceto sem sede codificada",
             len(mun) - len(sem_sede_codificada),
             int(dist[dist["distrito_sede"] == "Sim"].groupby("cod_mun").size().eq(1).sum()))
    c10_suf = c10[~c10["cod_mun"].isin(sem_sede_codificada)]
    V.checar("Distritos", f"Sedes IBGE 2010 (Cidade) no distrito de sufixo 05 ({br_int(len(c10_suf))} verificadas)",
             0, int((c10_suf["cod_distrito_2010"].str[7:] != "05").sum()))
    sede_nome = dist[dist["distrito_sede"] == "Sim"].set_index("cod_mun")["distrito"]
    sede_nome_dif = sorted(c for c, n in sede_nome.items() if norm_nome(n) != norm_nome(nome_mun[c]))
    V.checar("Distritos", "Distrito-sede com nome diferente do município (verificar renomeações)", "informativo",
             ", ".join(f"{c} {nome_mun[c]}/{uf_mun[c]} (sede '{sede_nome[c]}')" for c in sede_nome_dif) or "nenhum",
             ok=True, critico=False)

    # ---------- 7.5 Coordenadas da sede municipal ----------
    s = mun[["cod_mun", "municipio", "uf"]].merge(s22a[["cod_mun", "lat_2022", "lon_2022", "cat_2022"]],
                                                    on="cod_mun", how="left")
    s = s.merge(s22b[["cod_mun", "lat_b", "lon_b"]], on="cod_mun", how="left")
    s = s.merge(c10[["cod_mun", "lat_2010", "lon_2010"]], on="cod_mun", how="left")
    s = s.merge(kel, on="cod_mun", how="left")
    V.checar("Coordenadas", "Municípios com ponto IBGE 2022", len(mun), int(s["lat_2022"].notna().sum()))
    V.checar("Coordenadas", "Extração 2022 B cobre todos os códigos da extração A", 0, int(s["lat_b"].isna().sum()))
    d_ab = haversine_km(s["lat_2022"], s["lon_2022"], s["lat_b"], s["lon_b"])
    V.checar("Coordenadas", "Consistência de extração 2022 A x B: maior diferença (m)", "< 1",
             round(float(np.nanmax(d_ab)) * 1000, 2), ok=np.nanmax(d_ab) * 1000 < 1)
    # Checagem independente: pontos 'Sede Municipal' das bases 2022 por município (geometria, precisão total)
    c22 = loc22[(loc22["categoria"] == "Cidade") & loc22["subcategoria"].isin(["Sede Municipal", "Capital Federal"])]
    c22 = c22.drop_duplicates("cod_mun").merge(s22a[["cod_mun", "lat_2022", "lon_2022"]], on="cod_mun")
    d_ind = haversine_km(c22["lat"], c22["lon"], c22["lat_2022"], c22["lon_2022"]) * 1000
    V.checar("Coordenadas", f"Sedes 2022 x bases 2022 por município ({br_int(len(c22))} sedes, GO/SC/AM/AP/PA/RR): "
             "maior diferença (m)", "< 20", round(float(d_ind.max()), 1), ok=float(d_ind.max()) < 20)

    s["km_2022_2010"] = haversine_km(s["lat_2022"], s["lon_2022"], s["lat_2010"], s["lon_2010"])
    s["km_2022_kelvins"] = haversine_km(s["lat_2022"], s["lon_2022"], s["lat_k"], s["lon_k"])
    s["km_2010_kelvins"] = haversine_km(s["lat_2010"], s["lon_2010"], s["lat_k"], s["lon_k"])
    s["limiar"] = [LIMIAR_DIVERGENCIA_CAPITAL_KM if eh_capital[c] else LIMIAR_DIVERGENCIA_KM for c in s["cod_mun"]]
    escolha = pd.DataFrame([escolher_coordenada_sede(r, r["limiar"]) for _, r in s.iterrows()],
                           columns=["lat", "lon", "fonte_chave", "status_coord", "obs_coord"], index=s.index)
    s = pd.concat([s, escolha], axis=1)

    # Evidência objetiva para cada ponto 2022 rejeitado: grudado em outra localidade ou isolado
    ilocal = IndiceLocalidades(loc10)
    for i in s.index[s["fonte_chave"] == "2010"]:
        if pd.isna(s.at[i, "lat_2022"]):
            continue
        cod = s.at[i, "cod_mun"]
        excl = ((loc10["cod_mun"] == cod) & (loc10["categoria"] == "CIDADE")).to_numpy()
        viz, d = ilocal.mais_proxima(s.at[i, "lat_2022"], s.at[i, "lon_2022"], excl)
        if d <= RAIO_LOCALIDADE_KM:
            ev = (f"o ponto 2022 coincide com {CATEGORIA_2010.get(viz['categoria'], viz['categoria'].title())} '{viz['nome'].title()}' "
                  f"(IBGE 2010, a {br_num(d, 2)} km)")
        elif d > RAIO_ISOLADO_KM:
            ev = f"o ponto 2022 está isolado (localidade IBGE 2010 mais próxima a {br_num(d)} km)"
        else:
            ev = ""
            s.at[i, "status_coord"] = "Corrigida - a conferir"
        if pd.notna(s.at[i, "km_2010_kelvins"]) and s.at[i, "km_2010_kelvins"] > LIMIAR_DIVERGENCIA_KM:
            s.at[i, "status_coord"] = "Corrigida - a conferir"
            ev = juntar_obs(ev, f"IBGE 2010 x base independente: {br_num(s.at[i, 'km_2010_kelvins'])} km")
        s.at[i, "obs_coord"] = juntar_obs(s.at[i, "obs_coord"], ev)

    # Ponto 2022 arredondado ao minuto de arco (~1-2 km): preferir o ponto 2010, de precisão total
    for i in s.index[(s["fonte_chave"] == "2022") & s["lat_2022"].notna()]:
        if (arredondado_ao_minuto(s.at[i, "lat_2022"]) and arredondado_ao_minuto(s.at[i, "lon_2022"])
                and pd.notna(s.at[i, "km_2022_2010"]) and s.at[i, "km_2022_2010"] <= MAX_MINUTO_PARA_2010_KM):
            s.loc[i, ["lat", "lon"]] = [s.at[i, "lat_2010"], s.at[i, "lon_2010"]]
            s.at[i, "fonte_chave"] = "2010"
            s.at[i, "status_coord"] = "Corrigida"
            s.at[i, "obs_coord"] = juntar_obs(s.at[i, "obs_coord"], (
                f"Ponto IBGE 2022 ({s.at[i, 'lat_2022']:.4f}; {s.at[i, 'lon_2022']:.4f}) arredondado ao minuto de arco; "
                f"usado o ponto IBGE 2010, de precisão total, a {br_num(s.at[i, 'km_2022_2010'], 2)} km"))

    # Divergência moderada (3-5 km) com IBGE 2010 e base independente concordando entre si: anotar
    lo, hi = FAIXA_REVISAO_KM
    mod = ((s["fonte_chave"] == "2022") & s["km_2022_2010"].between(lo, hi) & s["km_2022_kelvins"].between(lo, hi)
           & (s["km_2010_kelvins"] <= CONCORDANCIA_REF_KM))
    for i in s.index[mod]:
        s.at[i, "obs_coord"] = juntar_obs(s.at[i, "obs_coord"], (
            f"Divergência moderada: IBGE 2010 ({s.at[i, 'lat_2010']:.4f}; {s.at[i, 'lon_2010']:.4f}) e base "
            f"independente concordam entre si e ficam a {br_num(s.at[i, 'km_2022_2010'])} km do ponto IBGE 2022 mantido"))
    s["divergencia_moderada"] = mod
    s["lat"] = s["lat"].astype(float).round(6)
    s["lon"] = s["lon"].astype(float).round(6)
    s["fonte_coord"] = s["fonte_chave"].map(ROTULO_FONTE)
    noronha = s["cat_2022"].astype(str).str.contains("Fernando de Noronha", na=False)
    s.loc[noronha & (s["fonte_chave"] == "2022"), "fonte_coord"] = "IBGE Localidades 2022 (Distrito Estadual)"

    # Curadoria documentada (status + nota + distância a referência externa calculada aqui)
    s["curadoria"] = ""
    for cod, cur in CURADORIA.items():
        idx_c = s.index[s["cod_mun"] == cod]
        if not len(idx_c):
            log.warning("Curadoria para código inexistente: %s", cod)
            continue
        i = idx_c[0]
        partes = [cur.get("nota", "")]
        if "ref" in cur:
            rlat, rlon, rnome = cur["ref"]
            d = float(haversine_km(s.at[i, "lat"], s.at[i, "lon"], rlat, rlon))
            partes.append(f"Referência externa — {rnome}: ({br_num(rlat, 4)}; {br_num(rlon, 4)}), a {br_num(d)} km "
                          "do ponto adotado")
        if "status" in cur:
            s.at[i, "status_coord"] = cur["status"]
        s.at[i, "obs_coord"] = juntar_obs(s.at[i, "obs_coord"], *partes)
        s.at[i, "curadoria"] = "Sim"

    # Teste ponto-no-polígono (Malha 2025)
    idx = IndiceEspacial(malha)
    tp = [idx.testar(c, la, lo) for c, la, lo in zip(s["cod_mun"], s["lat"], s["lon"])]
    s["dentro_poligono"] = [t.dentro for t in tp]
    s["dist_poligono_km"] = [t.dist_km for t in tp]
    fora_tol = [i for i, t in zip(s.index, tp) if not t.aceito()]
    for i in fora_tol:
        st = s.at[i, "status_coord"]
        s.at[i, "status_coord"] = {"OK": "A conferir", "Corrigida": "Corrigida - a conferir"}.get(st, st)
        s.at[i, "obs_coord"] = juntar_obs(s.at[i, "obs_coord"],
                                          f"Ponto {br_num(s.at[i, 'dist_poligono_km'])} km fora do limite municipal (Malha 2025)")
    s["sede_no_limite"] = np.where(s["dentro_poligono"], "Sim",
                                   np.where(s["dist_poligono_km"] <= TOLERANCIA_POLIGONO_KM, "Na borda (< 1 km)", "Não"))
    log.info("Status das coordenadas: %s", dict(Counter(s["status_coord"])))
    V.checar("Coordenadas", "Municípios sem coordenada final", 0, int(s["lat"].isna().sum()))
    V.checar("Coordenadas", "Coordenadas fora da caixa envolvente do Brasil", 0, int((
        ~s["lat"].between(BBOX_BRASIL["lat_min"], BBOX_BRASIL["lat_max"])
        | ~s["lon"].between(BBOX_BRASIL["lon_min"], BBOX_BRASIL["lon_max"])).sum()))
    V.checar("Coordenadas", f"Sede fora do próprio polígono por mais de {br_num(TOLERANCIA_POLIGONO_KM, 0)} km (Malha 2025)",
             0, len(fora_tol), critico=False)
    V.checar("Coordenadas", "Sede dentro do próprio polígono (estrito, malha simplificada)", "≥ 99,5%",
             br_pct(s["dentro_poligono"].mean(), 2), ok=s["dentro_poligono"].mean() >= 0.995, critico=False)
    V.checar("Coordenadas", "Pontos IBGE 2022 substituídos pelo IBGE 2010", "informativo",
             int((s["fonte_chave"] == "2010").sum()), ok=True, critico=False)
    V.checar("Coordenadas", "Pontos mantidos com divergência moderada (3-5 km) anotada", "informativo",
             int(mod.sum()), ok=True, critico=False)
    d_final_k = haversine_km(s["lat"], s["lon"], s["lat_k"], s["lon_k"])
    V.checar("Coordenadas", "Mediana da distância final x base independente (km)", "< 1",
             round(float(np.nanmedian(d_final_k)), 3), ok=np.nanmedian(d_final_k) < 1, critico=False)

    # ---------- 7.6 Pontos derivados da malha ----------
    pm, com_ilhas = derivar_pontos_malha(malha)
    log.info("Municípios com partes > %.0f km (ilhas) fora do centroide: %s", DISTANCIA_PARTES_CENTROIDE_KM,
             [f"{c} {nome_mun.get(c)}" for c in com_ilhas])
    V.checar("Malha", "Municípios com polígono na Malha 2025", len(mun), int(pm["cod_mun"].isin(mun["cod_mun"]).sum()))
    V.checar("Malha", "Ponto interno dentro do polígono", len(pm), int(pm["interno_dentro"].sum()))
    V.checar("Malha", "Centroide dentro do polígono (informativo: formas côncavas ficam fora)", "informativo",
             int(pm["centroide_dentro"].sum()), ok=True, critico=False)
    pm_s = pm.merge(s[["cod_mun", "lat", "lon"]], on="cod_mun")
    d_ci = haversine_km(pm_s["lat_centroide"], pm_s["lon_centroide"], pm_s["lat_interno"], pm_s["lon_interno"])
    d_sc = haversine_km(pm_s["lat"], pm_s["lon"], pm_s["lat_centroide"], pm_s["lon_centroide"])
    stats_malha = {"ci_med": float(np.median(d_ci)), "ci_p25": float(np.percentile(d_ci, 25)),
                   "ci_p75": float(np.percentile(d_ci, 75)), "sc_med": float(np.median(d_sc))}

    # ---------- 7.7 Coordenadas dos distritos ----------
    coord_mun = s.set_index("cod_mun")[["lat", "lon", "fonte_coord"]]
    dist = dist.merge(coord_mun.rename(columns={"lat": "lat_mun", "lon": "lon_mun", "fonte_coord": "fonte_mun"}),
                      left_on="cod_mun", right_index=True, how="left")
    # Localidades homônimas (não Cidade/Vila) por município, para distritos sem ponto 'Vila'
    def homonimos(df, col_cat, excluir):
        h = df[~df[col_cat].isin(excluir)].copy()
        h["chave"] = h["cod_mun"] + "|" + h["nome"].map(norm_nome)
        unicos = h[~h["chave"].duplicated(keep=False)]
        return unicos.set_index("chave")
    hom22 = homonimos(loc22, "categoria", {"Cidade", "Vila"})
    hom10 = homonimos(loc10, "categoria", {"CIDADE", "VILA"})
    loc22_nome = loc22.assign(n=loc22["nome"].map(norm_nome))

    eh_sede = dist["distrito_sede"] == "Sim"
    escolhas, rejeitados, n_vila_corr = [], [], 0
    for cod_d, cod_m, nome_d, sede_flag in zip(dist["cod_distrito"], dist["cod_mun"], dist["distrito"], eh_sede):
        if sede_flag:
            escolhas.append((np.nan, np.nan, None, None, "", ""))
            continue
        chave_h = cod_m + "|" + norm_nome(nome_d)
        cands = []
        if cod_d in v22.index:
            r = v22.loc[cod_d]
            cands.append(("vila2022", r["lat"], r["lon"], r["nome"], ""))
        if cod_d in v10.index:
            r = v10.loc[cod_d]
            cands.append(("vila2010", r["lat"], r["lon"], r["nome"], ""))
        if chave_h in hom22.index:
            r = hom22.loc[chave_h]
            cands.append(("hom2022", r["lat"], r["lon"], r["nome"], r["categoria"]))
        if chave_h in hom10.index:
            r = hom10.loc[chave_h]
            cands.append(("hom2010", r["lat"], r["lon"], r["nome"], CATEGORIA_2010.get(r["categoria"], r["categoria"].title())))
        obs = []
        # Vila 2022 x Vila 2010 muito distantes: a 2010 vence se uma localidade 2022 homônima a confirma
        if len(cands) >= 2 and cands[0][0] == "vila2022" and cands[1][0] == "vila2010":
            d_vv = float(haversine_km(cands[0][1], cands[0][2], cands[1][1], cands[1][2]))
            if d_vv > LIMIAR_VILAS_DIVERGENTES_KM:
                l22 = loc22_nome[(loc22_nome["cod_mun"] == cod_m) & (loc22_nome["categoria"] != "Vila")
                                 & (loc22_nome["n"] == norm_nome(nome_d))]
                d_hom = haversine_km(l22["lat"], l22["lon"], cands[1][1], cands[1][2]) if len(l22) else np.array([])
                if len(d_hom) and float(np.min(d_hom)) <= RAIO_HOMONIMO_KM:
                    j = int(np.argmin(d_hom))
                    obs.append(f"Vila IBGE 2022 ({cands[0][1]:.4f}; {cands[0][2]:.4f}) a {br_num(d_vv)} km da Vila 2010; "
                               f"a localidade 2022 homônima ({l22.iloc[j]['categoria']}) confirma o ponto 2010")
                    cands[0], cands[1] = cands[1], cands[0]
                    n_vila_corr += 1
                else:
                    obs.append(f"Vila IBGE 2022 e Vila IBGE 2010 ({cands[1][1]:.4f}; {cands[1][2]:.4f}) divergem "
                               f"{br_num(d_vv)} km; mantido o ponto 2022")
        achou = None
        for chave, la, lo_, nome_f, cat in cands:
            t = idx.testar(cod_m, la, lo_)
            if t.aceito():
                achou = (la, lo_, chave, nome_f, cat)
                break
            rejeitados.append((cod_d, nome_d, chave, la, lo_))
            obs.append(f"Ponto {ROTULO_FONTE[chave]} ({la:.4f}; {lo_:.4f}) descartado: caía {br_num(t.dist_km)} km "
                       f"fora do município (em {nome_mun.get(t.cai_em, 'área sem município')})")
        if achou:
            escolhas.append((achou[0], achou[1], achou[2], achou[3], juntar_obs(*obs), achou[4]))
        else:
            escolhas.append((np.nan, np.nan, "nenhuma", None, juntar_obs(*obs), ""))
    esc = pd.DataFrame(escolhas, columns=["lat", "lon", "fonte", "nome_fonte", "obs_dist", "cat_hom"], index=dist.index)
    dist = pd.concat([dist, esc], axis=1)
    log.info("Distritos: pontos reprovados no teste de polígono = %d %s | Vila 2010 preferida à 2022 = %d",
             len(rejeitados), [r[:3] for r in rejeitados], n_vila_corr)
    dist.loc[eh_sede, ["lat", "lon"]] = dist.loc[eh_sede, ["lat_mun", "lon_mun"]].values
    dist["fonte_dist"] = np.where(eh_sede, dist["fonte_mun"] + " — sede municipal", dist["fonte"].map(ROTULO_FONTE))
    hom = dist["fonte"].isin(["hom2022", "hom2010"])
    dist.loc[hom, "fonte_dist"] = [f"IBGE Localidades {f[-4:]} (localidade homônima: {c})"
                                   for f, c in zip(dist.loc[hom, "fonte"], dist.loc[hom, "cat_hom"])]
    # Nome do ponto x nome do distrito: grafia (anota) ou outra localidade (anota + qualifica a fonte)
    tem_vila = dist["fonte"].isin(["vila2022", "vila2010"]) & dist["lat"].notna()
    for i in dist.index[tem_vila]:
        nf, nd = dist.at[i, "nome_fonte"], dist.at[i, "distrito"]
        if norm_nome(nf) == norm_nome(nd):
            continue
        if parecido(nf, nd) >= LIMIAR_NOME_PARECIDO:
            dist.at[i, "obs_dist"] = juntar_obs(dist.at[i, "obs_dist"], f"Grafia na fonte do ponto: {str(nf).title()}")
        else:
            dist.at[i, "obs_dist"] = juntar_obs(dist.at[i, "obs_dist"], f"Nome do ponto na fonte: {str(nf).title()}")
            dist.at[i, "fonte_dist"] = dist.at[i, "fonte_dist"] + SUFIXO_NOME_DIFERENTE
    # Motivo explícito para cada distrito sem coordenada
    codigos10 = set(loc10["cod_distrito"])
    sem = dist["lat"].isna()
    motivos = Counter()
    for i in dist.index[sem]:
        cod_d, cod_m = dist.at[i, "cod_distrito"], dist.at[i, "cod_mun"]
        if cod_m in sem_sede_codificada:
            m = "Distrito urbano: o IBGE não publica ponto de sede para estes distritos"
        elif dist.at[i, "obs_dist"]:
            m = "Ponto disponível reprovado no teste de limite municipal"
        elif cod_m[:2] in ufs_loc22:
            m = "Sem ponto 'Vila' no IBGE Localidades 2022"
        elif cod_d in codigos10:
            m = "Distrito já existia em 2010, mas sem ponto 'Vila' no cadastro IBGE 2010"
        else:
            m = "Distrito criado após 2010; ponto IBGE 2022 desta UF não disponível no espelho usado"
        motivos[m] += 1
        dist.at[i, "obs_dist"] = juntar_obs(m, dist.at[i, "obs_dist"])
    dist["lat"] = dist["lat"].astype(float).round(6)
    dist["lon"] = dist["lon"].astype(float).round(6)
    log.info("Distritos sem coordenada por motivo: %s", dict(motivos))

    # Sede municipal colada à sede de OUTRO distrito do mesmo município: anotar no município
    sede_proxima_vila = {}
    nao_sede_pt = dist[(~eh_sede) & dist["lat"].notna()]
    d_ns = haversine_km(nao_sede_pt["lat"], nao_sede_pt["lon"], nao_sede_pt["lat_mun"], nao_sede_pt["lon_mun"])
    for (_, r), d in zip(nao_sede_pt.iterrows(), d_ns):
        if d <= RAIO_VILA_PROXIMA_KM and r["cod_mun"] not in sem_sede_codificada:
            sede_proxima_vila[r["cod_mun"]] = (r["distrito"], float(d))
    for cod, (nd, d) in sede_proxima_vila.items():
        i = s.index[s["cod_mun"] == cod][0]
        s.at[i, "obs_coord"] = juntar_obs(s.at[i, "obs_coord"], (
            f"O ponto da sede fica a {br_num(d, 2)} km da sede do distrito '{nd}' (não-sede)"))
    log.info("Sedes municipais a < %.1f km da sede de outro distrito: %s", RAIO_VILA_PROXIMA_KM,
             {c: v[0] for c, v in sede_proxima_vila.items()})
    cobertura = dist["lat"].notna().mean()
    V.checar("Distritos", "Distritos com coordenada de sede distrital", "informativo",
             f"{br_int(dist['lat'].notna().sum())} de {br_int(len(dist))} ({br_pct(cobertura)})", ok=True, critico=False)
    V.checar("Distritos", "Coordenadas de distrito obtidas por localidade homônima", "informativo",
             int(hom.sum()), ok=True, critico=False)
    V.checar("Distritos", "Coordenadas de distrito de localidade com outro nome (a conferir)", "informativo",
             int(dist["fonte_dist"].astype(str).str.endswith(SUFIXO_NOME_DIFERENTE).sum()), ok=True, critico=False)

    # ---------- 7.8 Montagem das tabelas finais ----------
    agg = (dist.sort_values("cod_distrito").groupby("cod_mun")
           .agg(qtd=("cod_distrito", "size"), lista=("distrito", lambda v: "; ".join(v))))
    sede = dist[dist["distrito_sede"] == "Sim"].set_index("cod_mun")
    M = mun.merge(s[["cod_mun", "lat", "lon", "fonte_coord", "status_coord", "obs_coord", "sede_no_limite"]], on="cod_mun")
    M = M.merge(pm, on="cod_mun", how="left").merge(agg, left_on="cod_mun", right_index=True, how="left")
    M["qtd"] = M["qtd"].fillna(0)
    M["cod_sede"] = M["cod_mun"].map(sede["cod_distrito"])
    M["nome_sede"] = M["cod_mun"].map(sede["distrito"])
    M = M.sort_values("cod_mun").reset_index(drop=True)
    tab_mun = pd.DataFrame({
        "Código IBGE": M["cod_mun"].astype(int),
        "Código IBGE (6 dígitos)": M["cod_mun"].str[:6].astype(int),
        "Município": M["municipio"],
        "Município - UF": M["municipio"] + " - " + M["uf"],
        "UF": M["uf"],
        "Código UF": M["cod_uf"].astype(int),
        "Estado": M["estado"],
        "Região": M["regiao"],
        "Código Região Intermediária": M["cod_rgint"].astype(int),
        "Região Geográfica Intermediária": M["nome_rgint"],
        "Código Região Imediata": M["cod_rgi"].astype(int),
        "Região Geográfica Imediata": M["nome_rgi"],
        "Latitude": M["lat"],
        "Longitude": M["lon"],
        "Fonte da Coordenada": M["fonte_coord"],
        "Status da Coordenada": M["status_coord"],
        "Capital (Sim/Não)": M["capital"],
        "Qtd. Distritos": M["qtd"].astype(int),
        "Código Distrito-Sede": M["cod_sede"].map(lambda v: int(v) if pd.notna(v) else None),
        "Distrito-Sede": M["nome_sede"],
        "Distritos": M["lista"],
        "Sede dentro do Limite Municipal": M["sede_no_limite"],
        "Observação da Coordenada": M["obs_coord"].replace("", None),
        "Latitude Centroide": M["lat_centroide"],
        "Longitude Centroide": M["lon_centroide"],
        "Centroide dentro do Município": np.where(M["centroide_dentro"], "Sim", "Não"),
        "Latitude Ponto Interno": M["lat_interno"],
        "Longitude Ponto Interno": M["lon_interno"],
    })

    D = dist.merge(mun[["cod_mun", "municipio", "uf", "cod_uf", "estado", "regiao", "cod_rgint", "nome_rgint",
                        "cod_rgi", "nome_rgi"]], on="cod_mun", how="left").sort_values("cod_distrito")
    tab_dist = pd.DataFrame({
        "Código Distrito": D["cod_distrito"].astype(int),
        "Distrito": D["distrito"],
        "Distrito-Sede": D["distrito_sede"],
        "Código IBGE": D["cod_mun"].astype(int),
        "Município": D["municipio"],
        "Município - UF": D["municipio"] + " - " + D["uf"],
        "UF": D["uf"],
        "Código UF": D["cod_uf"].astype(int),
        "Estado": D["estado"],
        "Região": D["regiao"],
        "Código Região Intermediária": D["cod_rgint"].astype(int),
        "Região Geográfica Intermediária": D["nome_rgint"],
        "Código Região Imediata": D["cod_rgi"].astype(int),
        "Região Geográfica Imediata": D["nome_rgi"],
        "Latitude Município": D["lat_mun"],
        "Longitude Município": D["lon_mun"],
        "Latitude Distrito": D["lat"],
        "Longitude Distrito": D["lon"],
        "Fonte Coordenada Distrito": D["fonte_dist"],
        "Observação": D["obs_dist"].replace("", None),
    }).reset_index(drop=True)

    tab_rgi = (M.groupby(["cod_rgi", "nome_rgi", "cod_rgint", "nome_rgint", "uf", "cod_uf", "estado"], as_index=False)
               .agg(n=("cod_mun", "size")).sort_values("cod_rgi"))
    tab_rgi = pd.DataFrame({"Código Região Imediata": tab_rgi["cod_rgi"].astype(int),
                            "Região Geográfica Imediata": tab_rgi["nome_rgi"],
                            "Região Imediata - UF": tab_rgi["nome_rgi"] + " - " + tab_rgi["uf"],
                            "Código Região Intermediária": tab_rgi["cod_rgint"].astype(int),
                            "Região Geográfica Intermediária": tab_rgi["nome_rgint"],
                            "UF": tab_rgi["uf"], "Código UF": tab_rgi["cod_uf"].astype(int),
                            "Qtd. Municípios": tab_rgi["n"].astype(int)})
    tab_rgint = (M.groupby(["cod_rgint", "nome_rgint", "uf", "cod_uf", "estado"], as_index=False)
                 .agg(n=("cod_mun", "size"), nrgi=("cod_rgi", "nunique")).sort_values("cod_rgint"))
    tab_rgint = pd.DataFrame({"Código Região Intermediária": tab_rgint["cod_rgint"].astype(int),
                              "Região Geográfica Intermediária": tab_rgint["nome_rgint"],
                              "UF": tab_rgint["uf"], "Código UF": tab_rgint["cod_uf"].astype(int),
                              "Qtd. Regiões Imediatas": tab_rgint["nrgi"].astype(int),
                              "Qtd. Municípios": tab_rgint["n"].astype(int)})
    cap = M[M["capital"] == "Sim"].set_index("cod_uf")["municipio"]
    tab_uf = M[["cod_uf", "uf", "estado", "cod_regiao", "regiao"]].drop_duplicates().sort_values("cod_uf")
    tab_uf = pd.DataFrame({"Código UF": tab_uf["cod_uf"].astype(int), "UF": tab_uf["uf"], "Estado": tab_uf["estado"],
                           "Código Região": tab_uf["cod_regiao"].astype(int), "Região": tab_uf["regiao"],
                           "Município Capital": tab_uf["cod_uf"].map(cap)})
    V.checar("Dimensões", "Linhas em Regioes_Imediatas", ESPERADO["reg_imediatas"], len(tab_rgi))
    V.checar("Dimensões", "Linhas em Regioes_Intermediarias", ESPERADO["reg_intermediarias"], len(tab_rgint))
    V.checar("Dimensões", "Linhas em UFs", ESPERADO["ufs"], len(tab_uf))

    # Auditoria: divergência > 5 km entre fontes, status ≠ OK, sede na borda, divergência moderada ou curadoria
    s_aud = s.copy()
    s_aud["motivo"] = ""
    regras = [
        ((s_aud[["km_2022_2010", "km_2022_kelvins", "km_2010_kelvins"]] > LIMIAR_AUDITORIA_KM).any(axis=1),
         f"Divergência > {br_num(LIMIAR_AUDITORIA_KM, 0)} km entre fontes"),
        (s_aud["divergencia_moderada"], "Divergência moderada (3-5 km)"),
        (s_aud["status_coord"] != "OK", "Status diferente de OK"),
        (s_aud["sede_no_limite"] != "Sim", "Sede na borda/fora do limite"),
        (s_aud["cod_mun"].isin(sede_proxima_vila), "Sede junto à sede de outro distrito"),
        (s_aud["curadoria"] == "Sim", "Curadoria documentada"),
    ]
    for mask, rot in regras:
        s_aud.loc[mask, "motivo"] = s_aud.loc[mask, "motivo"].map(lambda v, r=rot: juntar_obs(v, r))
    aud = s_aud[s_aud["motivo"] != ""].copy()

    def decisao(r):
        if r["fonte_chave"] == "2010":
            return "Usado IBGE 2010 (ponto 2022 rejeitado — ver observação)"
        tem10, temk = pd.notna(r["km_2022_2010"]), pd.notna(r["km_2022_kelvins"])
        if r["status_coord"] == "A conferir":
            return "Mantido IBGE 2022, a conferir (ver observação)"
        if r["divergencia_moderada"]:
            return "Mantido IBGE 2022 (divergência moderada anotada)"
        if tem10 and temk and r["km_2022_kelvins"] > LIMIAR_AUDITORIA_KM and r["km_2022_2010"] <= LIMIAR_AUDITORIA_KM:
            return f"Mantido IBGE 2022 (IBGE 2010 a {br_num(r['km_2022_2010'])} km confirma; base independente diverge)"
        if tem10 and temk and r["km_2022_2010"] > LIMIAR_AUDITORIA_KM and r["km_2022_kelvins"] <= LIMIAR_AUDITORIA_KM:
            return (f"Mantido IBGE 2022 (base independente a {br_num(r['km_2022_kelvins'])} km confirma; "
                    "ponto 2010 desatualizado)")
        if tem10 and temk and min(r["km_2022_2010"], r["km_2022_kelvins"]) > LIMIAR_AUDITORIA_KM:
            return "Mantido IBGE 2022 (capital: divergência dentro do limiar de 10 km para áreas urbanas extensas)"
        if not tem10 or not temk:
            return "Mantido IBGE 2022 (uma das referências não existe)"
        return "Mantido IBGE 2022"

    aud["decisao"] = aud.apply(decisao, axis=1)
    ordem = {"Corrigida - a conferir": 0, "A conferir": 1, "Corrigida": 2, "OK": 3}
    aud = aud.assign(_o=aud["status_coord"].map(ordem).fillna(9)).sort_values(["_o", "km_2022_2010"], ascending=[True, False])
    tab_aud = pd.DataFrame({
        "Código IBGE": aud["cod_mun"].astype(int), "Município": aud["municipio"], "UF": aud["uf"],
        "Status da Coordenada": aud["status_coord"], "Decisão": aud["decisao"], "Motivo da Inclusão": aud["motivo"],
        "Latitude Final": aud["lat"], "Longitude Final": aud["lon"],
        "Lat. IBGE 2022": aud["lat_2022"], "Lon. IBGE 2022": aud["lon_2022"],
        "Lat. IBGE 2010": aud["lat_2010"].round(6), "Lon. IBGE 2010": aud["lon_2010"].round(6),
        "Lat. Base Independente": aud["lat_k"], "Lon. Base Independente": aud["lon_k"],
        "Dist. 2022 x 2010 (km)": aud["km_2022_2010"].round(2),
        "Dist. 2022 x Independente (km)": aud["km_2022_kelvins"].round(2),
        "Dist. 2010 x Independente (km)": aud["km_2010_kelvins"].round(2),
        "Sede dentro do Limite Municipal": aud["sede_no_limite"],
        "Observação": aud["obs_coord"].replace("", None),
    })
    n_aud_div = int(regras[0][0].sum())
    log.info("Auditoria de coordenadas: %d municípios listados", len(tab_aud))

    # ---------- 7.9 Validações finais de integridade das tabelas ----------
    V.checar("Tabelas", "Soma de 'Qtd. Distritos' = linhas da aba Distritos", len(tab_dist), int(tab_mun["Qtd. Distritos"].sum()))
    V.checar("Tabelas", "Campos-chave vazios (aba Municipios)", 0,
             int(tab_mun[["Código IBGE", "Município", "UF", "Código Região Imediata", "Código Região Intermediária",
                          "Latitude", "Longitude"]].isna().sum().sum()))
    V.checar("Tabelas", "Campos-chave vazios (aba Distritos)", 0,
             int(tab_dist[["Código Distrito", "Distrito", "Código IBGE", "Município", "UF",
                           "Latitude Município", "Longitude Município"]].isna().sum().sum()))
    ruins = 0
    for t in (tab_mun, tab_dist, tab_rgi, tab_rgint, tab_uf):
        for c in t.columns:
            if t[c].dtype == object:
                ruins += int(t[c].astype(str).str.contains("[¿�\x00-\x08\x0b\x0c\x0e-\x1f]", regex=True).sum())
    V.checar("Texto", "Células de texto com caractere corrompido ('¿', '�') ou de controle", 0, ruins)

    # ---------- 7.10 Escrita do Excel ----------
    wb = Workbook()
    wb._named_styles["Normal"].font = Font(name="Arial", size=10)
    ws_leia = wb.active
    ws_leia.title = "LEIA-ME"
    ws_m = wb.create_sheet("Municipios")
    ws_d = wb.create_sheet("Distritos")
    ws_rgi = wb.create_sheet("Regioes_Imediatas")
    ws_rgint = wb.create_sheet("Regioes_Intermediarias")
    ws_uf = wb.create_sheet("UFs")
    ws_res = wb.create_sheet("Resumo_UF")
    ws_aud = wb.create_sheet("Auditoria_Coordenadas")
    ws_val = wb.create_sheet("Validacao")
    ws_fon = wb.create_sheet("Fontes")

    W_COD, W_NOME, W_REG = 12, 30, 30
    larg_m = {"Código IBGE": W_COD, "Código IBGE (6 dígitos)": W_COD, "Município": W_NOME, "Município - UF": 34,
              "UF": 6, "Código UF": 8, "Estado": 20, "Região": 13, "Código Região Intermediária": 13,
              "Região Geográfica Intermediária": W_REG, "Código Região Imediata": 12,
              "Região Geográfica Imediata": W_REG, "Latitude": 12, "Longitude": 12, "Fonte da Coordenada": 36,
              "Status da Coordenada": 22, "Capital (Sim/Não)": 10, "Qtd. Distritos": 10, "Código Distrito-Sede": 13,
              "Distrito-Sede": 26, "Distritos": 45, "Sede dentro do Limite Municipal": 16,
              "Observação da Coordenada": 70, "Latitude Centroide": 12, "Longitude Centroide": 12,
              "Centroide dentro do Município": 13, "Latitude Ponto Interno": 12, "Longitude Ponto Interno": 12}
    fmt_m = {c: FMT_INT for c in ("Código IBGE", "Código IBGE (6 dígitos)", "Código UF", "Código Região Intermediária",
                                  "Código Região Imediata", "Qtd. Distritos", "Código Distrito-Sede")}
    fmt_m.update({c: FMT_COORD for c in tab_mun.columns if c.startswith(("Latitude", "Longitude"))})
    dv_txt = ", ".join(dv_excecoes)
    com_m = {
        "Código IBGE": "Código IBGE de 7 dígitos (com dígito verificador). Chave para relacionamentos. "
                       f"Atenção: {len(dv_excecoes)} códigos oficiais não satisfazem o algoritmo do DV ({dv_txt}); "
                       "não os rejeite em validações de ETL.",
        "Código IBGE (6 dígitos)": "Código sem o dígito verificador (padrão DATASUS/SIH/SIM).",
        "Município - UF": "Rótulo único para filtros e mapas: há nomes de município repetidos em UFs diferentes.",
        "Distritos": "Todos os distritos do município (DTB 2025), em ordem de código; o distrito-sede (sufixo 05) vem primeiro.",
        "Código Distrito-Sede": "Distrito de sufixo 05 (convenção IBGE). Vazio em São Paulo/SP, cujos 96 distritos são "
                                "numerados 01–96 e não têm sede codificada. Coluna anulável: leia como inteiro anulável (Int64).",
        "Latitude": "Latitude (graus decimais, SIRGAS 2000) do ponto oficial da SEDE MUNICIPAL (IBGE). Ver 'Fonte da Coordenada'.",
        "Longitude": "Longitude (graus decimais, SIRGAS 2000) do ponto oficial da SEDE MUNICIPAL (IBGE).",
        "Status da Coordenada": "OK = ponto IBGE 2022 mantido. Corrigida = ponto 2022 trocado pelo IBGE 2010 (evidência na "
                                "observação). 'Corrigida - a conferir' / 'A conferir' = fontes ainda divergem; ver observação.",
        "Sede dentro do Limite Municipal": "Teste ponto-no-polígono contra a Malha Municipal IBGE 2025 (simplificada).",
        "Latitude Centroide": "Centroide geométrico (equivalente de área) do polígono municipal (Malha IBGE 2025), sem ilhas "
                              "oceânicas distantes. Pode cair fora do município em formas côncavas.",
        "Latitude Ponto Interno": "Ponto mais interno do polígono (polo de inacessibilidade). Sempre dentro do município; "
                                  "ideal para rótulos de mapa.",
    }
    escrever_tabela(ws_m, tab_mun, "tbl_Municipios", larg_m, fmt_m, com_m, freeze="D2")

    larg_d = {"Código Distrito": 13, "Distrito": 30, "Distrito-Sede": 13, "Código IBGE": W_COD, "Município": W_NOME,
              "Município - UF": 34, "UF": 6, "Código UF": 8, "Estado": 20, "Região": 13,
              "Código Região Intermediária": 13, "Região Geográfica Intermediária": W_REG, "Código Região Imediata": 12,
              "Região Geográfica Imediata": W_REG, "Latitude Município": 12, "Longitude Município": 12,
              "Latitude Distrito": 12, "Longitude Distrito": 12, "Fonte Coordenada Distrito": 52, "Observação": 70}
    fmt_d = {c: FMT_INT for c in ("Código Distrito", "Código IBGE", "Código UF", "Código Região Intermediária",
                                  "Código Região Imediata")}
    fmt_d.update({c: FMT_COORD for c in tab_dist.columns if c.startswith(("Latitude", "Longitude"))})
    com_d = {
        "Código Distrito": "Código IBGE do distrito (9 dígitos) = código do município (7) + sufixo (2).",
        "Código IBGE": "Código IBGE do MUNICÍPIO (7 dígitos) — chave para a aba Municipios.",
        "Distrito-Sede": "Sim = distrito-sede (sufixo 05). 'Não se aplica' = São Paulo/SP (distritos 01–96, sem sede codificada).",
        "Latitude Município": "Coordenada da sede municipal (mesma da aba Municipios).",
        "Latitude Distrito": "Coordenada da sede do distrito (ponto IBGE 'Vila' ou localidade homônima). No distrito-sede = "
                             "sede municipal. Vazio quando não há ponto oficial (motivo na coluna Observação).",
    }
    escrever_tabela(ws_d, tab_dist, "tbl_Distritos", larg_d, fmt_d, com_d, freeze="C2")

    escrever_tabela(ws_rgi, tab_rgi, "tbl_Regioes_Imediatas",
                    {"Código Região Imediata": 13, "Região Geográfica Imediata": W_REG, "Região Imediata - UF": 34,
                     "Código Região Intermediária": 13, "Região Geográfica Intermediária": W_REG, "UF": 6,
                     "Código UF": 8, "Qtd. Municípios": 11},
                    {"Código Região Imediata": FMT_INT, "Código Região Intermediária": FMT_INT, "Código UF": FMT_INT})
    escrever_tabela(ws_rgint, tab_rgint, "tbl_Regioes_Intermediarias",
                    {"Código Região Intermediária": 13, "Região Geográfica Intermediária": W_REG, "UF": 6,
                     "Código UF": 8, "Qtd. Regiões Imediatas": 12, "Qtd. Municípios": 11},
                    {"Código Região Intermediária": FMT_INT, "Código UF": FMT_INT})
    escrever_tabela(ws_uf, tab_uf, "tbl_UFs", {"Código UF": 9, "UF": 6, "Estado": 22, "Código Região": 9,
                                               "Região": 14, "Município Capital": 22},
                    {"Código UF": FMT_INT, "Código Região": FMT_INT})

    # Resumo_UF — fórmulas vivas; letras de coluna derivadas dos cabeçalhos (sem posições fixas)
    n_m, n_d = len(tab_mun) + 1, len(tab_dist) + 1
    n_rgi, n_rgint = len(tab_rgi) + 1, len(tab_rgint) + 1
    cM_uf, cM_fonte, cM_qtd = letra(tab_mun, "UF"), letra(tab_mun, "Fonte da Coordenada"), letra(tab_mun, "Qtd. Distritos")
    cM_lat, cM_lon = letra(tab_mun, "Latitude"), letra(tab_mun, "Longitude")
    cD_uf, cD_sede = letra(tab_dist, "UF"), letra(tab_dist, "Distrito-Sede")
    cRI_uf, cRT_uf = letra(tab_rgi, "UF"), letra(tab_rgint, "UF")
    hdr = ["UF", "Estado", "Região", "Municípios", "Distritos", "Regiões Imediatas", "Regiões Intermediárias",
           "Coord. corrigidas (IBGE 2010)", "Média distritos/município"]
    ws_res.append(hdr)
    for i, r in enumerate(tab_uf.itertuples(index=False), start=2):
        ws_res.append([
            r.UF, r.Estado, r.Região,
            f"=COUNTIF(Municipios!${cM_uf}$2:${cM_uf}${n_m},$A{i})",
            f"=COUNTIF(Distritos!${cD_uf}$2:${cD_uf}${n_d},$A{i})",
            f"=COUNTIF(Regioes_Imediatas!${cRI_uf}$2:${cRI_uf}${n_rgi},$A{i})",
            f"=COUNTIF(Regioes_Intermediarias!${cRT_uf}$2:${cRT_uf}${n_rgint},$A{i})",
            f'=COUNTIFS(Municipios!${cM_uf}$2:${cM_uf}${n_m},$A{i},Municipios!${cM_fonte}$2:${cM_fonte}${n_m},"IBGE Localidades 2010*")',
            f"=IF(D{i}=0,0,E{i}/D{i})",
        ])
    ult = len(tab_uf) + 1
    tot = ult + 2  # uma linha em branco entre a tabela e o total
    ws_res.cell(row=tot, column=1, value="TOTAL BRASIL")
    for j, c in enumerate("DEFGH", start=4):
        ws_res.cell(row=tot, column=j, value=f"=SUM({c}2:{c}{ult})")
    ws_res.cell(row=tot, column=9, value=f"=IF(D{tot}=0,0,E{tot}/D{tot})")
    tab = Table(displayName="tbl_Resumo_UF", ref=f"A1:I{ult}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws_res.add_table(tab)
    for j, w in enumerate([8, 22, 14, 12, 11, 12, 14, 14, 13], start=1):
        ws_res.column_dimensions[get_column_letter(j)].width = w
        c = ws_res.cell(row=1, column=j)
        c.font = FONTE_CAB
        c.fill = PatternFill("solid", fgColor=COR_CABECALHO)
        c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    for rr in list(range(2, ult + 1)) + [tot]:
        for j in range(1, 10):
            cel = ws_res.cell(row=rr, column=j)
            cel.font = Font(name="Arial", size=10, bold=(rr == tot))
            if j >= 4:
                cel.number_format = "0.00" if j == 9 else "#,##0"
            if rr == tot:
                cel.fill = PatternFill("solid", fgColor="D9E1F2")
    ws_res.row_dimensions[1].height = 42
    ws_res.freeze_panes = "A2"
    ws_res.cell(row=tot + 2, column=1, value="Todas as colunas numéricas são fórmulas (COUNTIF/COUNTIFS/SUM) sobre as abas "
                                              "de dados. No Power BI, importe a tabela tbl_Resumo_UF (sem a linha de total).").font = \
        Font(name="Arial", italic=True, size=9)

    escrever_tabela(ws_aud, tab_aud, "tbl_Auditoria_Coordenadas",
                    {"Código IBGE": W_COD, "Município": 28, "UF": 6, "Status da Coordenada": 22, "Decisão": 60,
                     "Motivo da Inclusão": 40, "Sede dentro do Limite Municipal": 16, "Observação": 90},
                    {"Código IBGE": FMT_INT, **{c: FMT_COORD for c in tab_aud.columns if c.startswith(("Lat", "Lon"))},
                     **{c: "0.00" for c in tab_aud.columns if c.startswith("Dist.")}}, freeze="D2")

    # Validacao — checagens do pipeline + checagens com fórmula viva, na MESMA tabela
    val = [{"Origem": "Pipeline Python", "Grupo": g, "Verificação": n, "Esperado": e, "Obtido": o,
            "Resultado": "OK" if ok else ("FALHA" if cr else "ALERTA"), "Crítica": "Sim" if cr else "Não"}
           for g, n, e, o, ok, cr in V.resultados]
    formulas = [
        ("Linhas na aba Municipios", ESPERADO["municipios"], f"=COUNTA(Municipios!A2:A{n_m})"),
        ("Linhas na aba Distritos", ESPERADO["distritos"], f"=COUNTA(Distritos!A2:A{n_d})"),
        ("Soma de Qtd. Distritos (aba Municipios)", ESPERADO["distritos"], f"=SUM(Municipios!{cM_qtd}2:{cM_qtd}{n_m})"),
        ("Regiões Imediatas", ESPERADO["reg_imediatas"], f"=COUNTA(Regioes_Imediatas!A2:A{n_rgi})"),
        ("Regiões Intermediárias", ESPERADO["reg_intermediarias"], f"=COUNTA(Regioes_Intermediarias!A2:A{n_rgint})"),
        ("UFs", ESPERADO["ufs"], "=COUNTA(UFs!A:A)-1"),
        ("Municípios sem latitude", 0, f"=COUNTBLANK(Municipios!{cM_lat}2:{cM_lat}{n_m})"),
        ("Municípios sem longitude", 0, f"=COUNTBLANK(Municipios!{cM_lon}2:{cM_lon}{n_m})"),
        ("Latitude fora do Brasil (< -34 ou > 5,5)", 0,
         f"=SUMPRODUCT(--(Municipios!{cM_lat}2:{cM_lat}{n_m}<-34))+SUMPRODUCT(--(Municipios!{cM_lat}2:{cM_lat}{n_m}>5.5))"),
        ("Longitude fora do Brasil (< -74,2 ou > -28,5)", 0,
         f"=SUMPRODUCT(--(Municipios!{cM_lon}2:{cM_lon}{n_m}<-74.2))+SUMPRODUCT(--(Municipios!{cM_lon}2:{cM_lon}{n_m}>-28.5))"),
        ("Distrito-sede marcado 'Sim'", ESPERADO["municipios"] - len(sem_sede_codificada),
         f'=COUNTIF(Distritos!{cD_sede}2:{cD_sede}{n_d},"Sim")'),
    ]
    n_py = len(val)
    for k, (nome, esp, f) in enumerate(formulas):
        lin = n_py + 2 + k
        val.append({"Origem": "Fórmula Excel", "Grupo": "Excel", "Verificação": nome, "Esperado": esp, "Obtido": f,
                    "Resultado": f'=IF(E{lin}=D{lin},"OK","FALHA")', "Crítica": "Sim"})
    val_df = pd.DataFrame(val)
    escrever_tabela(ws_val, val_df, "tbl_Validacao",
                    {"Origem": 16, "Grupo": 13, "Verificação": 80, "Esperado": 16, "Obtido": 40, "Resultado": 11,
                     "Crítica": 8}, {}, freeze="A2")
    for (c,) in ws_val.iter_rows(min_row=2, min_col=3, max_col=3):
        c.alignment = Alignment(wrap_text=True, vertical="top")
    col_resultado = list(val_df.columns).index("Resultado") + 1

    fontes_df = pd.DataFrame([{
        "Chave": k, "Produto": v["produto"], "Safra / Data-base": v["safra"], "Uso no pipeline": v["uso"],
        "Origem do arquivo usado": v["origem"], "Endereço oficial (IBGE)": v["oficial"],
        "Arquivo": v["arquivo"], "SHA-256": sha256(F / v["arquivo"]),
    } for k, v in FONTES.items()])
    escrever_tabela(ws_fon, fontes_df, "tbl_Fontes",
                    {"Chave": 16, "Produto": 60, "Safra / Data-base": 30, "Uso no pipeline": 50,
                     "Origem do arquivo usado": 60, "Endereço oficial (IBGE)": 60, "Arquivo": 45, "SHA-256": 66}, {})
    for row in ws_fon.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")

    # LEIA-ME — todos os números vêm dos dados
    n_corr = int((s["fonte_chave"] == "2010").sum())
    n_conf = int(s["status_coord"].str.contains("conferir").sum())
    rej = s[(s["fonte_chave"] == "2010") & s["km_2022_2010"].notna()]
    maior = rej.loc[rej["km_2022_2010"].idxmax()] if len(rej) else None
    exemplo = (f" (ex.: {maior['municipio']}/{maior['uf']}, cujo ponto 2022 estava {br_num(maior['km_2022_2010'], 0)} km "
               "deslocado)") if maior is not None else ""
    sp_dist = int((dist["cod_mun"].isin(sem_sede_codificada)).sum())
    cov_d = int(dist["lat"].notna().sum())
    bloco_motivos = "; ".join(f"{br_int(n)} — {m.lower()}" for m, n in motivos.most_common())
    n_mod = int(mod.sum())
    n_div = n_aud_div
    nomes_txt = "; ".join(f"'{a}' → '{b}'" for a, b in nomes_corrigidos.items())
    txt_sp = ""
    if "3550308" in CURADORIA and "ref" in CURADORIA["3550308"]:
        i_sp = s.index[s["cod_mun"] == "3550308"][0]
        d_sp = float(haversine_km(s.at[i_sp, "lat"], s.at[i_sp, "lon"], *CURADORIA["3550308"]["ref"][:2]))
        txt_sp = f" (ex.: em São Paulo o ponto fica a {br_num(d_sp)} km da Praça da Sé)"
    ex_rej = next((r for r in rejeitados if r[2] == "vila2022"), None)
    txt_rej = (f" Isso reprovou {len(rejeitados)} ponto(s) — p.ex. a Vila 2022 de {ex_rej[1]} "
               f"({br_num(ex_rej[3], 4)}; {br_num(ex_rej[4], 4)}), que caía fora do município." if ex_rej else "")
    cur_capitais = [nome_mun[c] for c in CURADORIA if eh_capital.get(c)]
    texto = [
        ("Municípios do Brasil — IBGE", None, "titulo"),
        ("Gerado em", f"{gerado_em} por build_municipios_ibge.py (pipeline reprodutível; fontes e hashes na aba Fontes).", "normal"),
        ("Cobertura", f"{br_int(len(tab_mun))} municípios (DTB 2025, que inclui Brasília/DF e o Distrito Estadual de Fernando "
                      f"de Noronha/PE, codificado pelo IBGE como município) · {br_int(len(tab_dist))} distritos · "
                      f"{len(tab_rgi)} Regiões Geográficas Imediatas · {len(tab_rgint)} Intermediárias · {len(tab_uf)} UFs.", "normal"),
        ("", "", "normal"),
        ("ABAS", "", "secao"),
        ("Municipios", "1 linha por município: códigos, UF, regiões imediata e intermediária, coordenada da sede (com fonte, "
                       "status e observação), distritos, centroide e ponto interno do polígono.", "normal"),
        ("Distritos", "1 linha por distrito (tabela plana): distrito + município + UF + regiões + coordenada do município e "
                      "da sede distrital.", "normal"),
        ("Regioes_Imediatas / Regioes_Intermediarias / UFs", "Dimensões para modelo estrela (Power BI).", "normal"),
        ("Resumo_UF", "Contagens por UF com fórmulas vivas (COUNTIF/COUNTIFS/SUM).", "normal"),
        ("Auditoria_Coordenadas", f"{br_int(len(tab_aud))} municípios: {br_int(n_div)} com divergência > "
                                  f"{br_num(LIMIAR_AUDITORIA_KM, 0)} km entre fontes, {br_int(n_mod)} com divergência moderada "
                                  "(3-5 km) e os demais por status ≠ OK, sede na borda do limite, sede junto à de outro "
                                  "distrito ou curadoria — com os valores de cada fonte e a decisão.", "normal"),
        ("Validacao", f"{len(V.resultados)} checagens automáticas do pipeline + {len(formulas)} checagens com fórmula viva.", "normal"),
        ("Fontes", "Produto IBGE, safra, origem do arquivo, endereço oficial e SHA-256 de cada insumo.", "normal"),
        ("", "", "normal"),
        ("COORDENADAS — O QUE SIGNIFICAM", "", "secao"),
        ("Latitude / Longitude", "Ponto oficial da SEDE MUNICIPAL (área urbana da sede) do IBGE — Localidades do Brasil 2022 "
                                 "(categoria 'Cidade' / 'Sede Municipal'). Graus decimais, datum SIRGAS 2000 (EPSG:4674). "
                                 "Município é uma área; este é o ponto de referência oficial da cidade-sede, padrão para mapas "
                                 "de pontos, distâncias e geocodificação. Não é necessariamente a prefeitura nem o marco zero.", "normal"),
        ("Precisão", "O IBGE publica os atributos de 2022 com 4 casas decimais (≈ 11 m). Os pontos IBGE 2010 têm precisão "
                     "total (gravados com 6 casas, ≈ 0,1 m).", "normal"),
        ("Regra de correção", f"O ponto IBGE 2022 é substituído pelo IBGE 2010 quando diverge > {br_num(LIMIAR_DIVERGENCIA_KM, 0)} "
                              f"km (> {br_num(LIMIAR_DIVERGENCIA_CAPITAL_KM, 0)} km nas capitais, de área urbana extensa) de "
                              "DUAS fontes independentes ao mesmo tempo (IBGE 2010 e a base comunitária kelvins). A "
                              "observação registra a evidência: o ponto 2022 coincide com outra localidade (povoado, vila) "
                              "ou está isolado. Também usa o 2010 quando o ponto 2022 foi arredondado ao minuto de arco. "
                              f"Resultado: {n_corr} municípios corrigidos{exemplo}; {n_conf} seguem marcados 'a conferir'. "
                              f"{n_mod} pontos com divergência moderada (3-5 km) foram mantidos e anotados.", "normal"),
        ("Curadoria", f"{len(CURADORIA)} municípios têm nota de curadoria documentada (status e/ou referência externa, com a "
                      "distância calculada pelo pipeline); a coordenada continua sendo um ponto oficial do IBGE. "
                      f"Inclui as capitais: {', '.join(cur_capitais)}.", "normal"),
        ("Validação geométrica", "Todas as sedes foram testadas contra o polígono do próprio município (Malha Municipal IBGE "
                                 "2025). 'Na borda' = até 1 km fora, efeito da simplificação da malha.", "normal"),
        ("Centroide x Ponto Interno", f"Centroide = centro de massa do polígono (pode cair fora em municípios côncavos; ilhas "
                                      f"oceânicas a > {br_num(DISTANCIA_PARTES_CENTROIDE_KM, 0)} km ficam de fora, ex.: "
                                      f"Trindade/Martim Vaz em Vitória). Ponto interno = polo de inacessibilidade, sempre dentro. "
                                      f"Entre si diferem em mediana {br_num(stats_malha['ci_med'])} km (p25 "
                                      f"{br_num(stats_malha['ci_p25'])}; p75 {br_num(stats_malha['ci_p75'])}); a sede fica em "
                                      f"mediana a {br_num(stats_malha['sc_med'])} km do centroide.", "normal"),
        ("Grandes cidades", "Em metrópoles o ponto IBGE da sede é um ponto representativo da área urbana e pode ficar a alguns "
                            f"km do centro histórico{txt_sp}. Para análises intramunicipais, use distritos ou setores "
                            "censitários.", "normal"),
        ("", "", "normal"),
        ("DISTRITOS", "", "secao"),
        ("Distrito-sede", f"Convenção do IBGE: o distrito-sede recebe o sufixo 05 — validado: {br_int(len(c10_suf))} das "
                          f"{br_int(len(c10))} sedes IBGE 2010 estão no distrito 05; a exceção é São Paulo/SP, com "
                          f"{sp_dist} distritos numerados 01–{sp_dist} sem sede codificada ('Não se aplica'). Em "
                          f"{len(sede_nome_dif)} municípios o distrito-sede tem nome diferente do município (renomeações): "
                          + ", ".join(f"{nome_mun[c]}/{uf_mun[c]} (sede '{sede_nome[c]}')" for c in sede_nome_dif) + ".", "normal"),
        ("Coordenada do distrito", "Distrito-sede = sede municipal. Demais distritos, nesta ordem: ponto 'Vila' (sede "
                                   "distrital) do IBGE Localidades 2022 (GO, SC, AM, AP, PA, RR) → 'Vila' do IBGE 2010 → "
                                   "localidade IBGE homônima e única no município (povoado, núcleo etc.). Cada ponto só é aceito "
                                   f"se cair dentro do próprio município.{txt_rej} Cobertura: {br_int(cov_d)} de {br_int(len(dist))} "
                                   f"distritos ({br_pct(cov_d / len(dist))}). Sem coordenada: {bloco_motivos}. O motivo está "
                                   "na coluna Observação de cada linha.", "normal"),
        ("", "", "normal"),
        ("NOMES OFICIAIS", "", "secao"),
        ("Correção de codificação", f"O IBGE publica {len(nomes_corrigidos)} nomes de regiões com '¿' no lugar do travessão "
                                    f"(defeito replicado pela API). Normalizados para ' - ': {nomes_txt}.", "normal"),
        ("Grafias das regiões", "Os nomes das regiões seguem a grafia oficial da DTB, que às vezes difere da grafia atual do "
                                "município homônimo (ex.: 'Juíz de Fora', 'Açu', 'Itapagé'). Relacione sempre por código.", "normal"),
        ("Dígito verificador", f"{len(dv_excecoes)} códigos oficiais não satisfazem o algoritmo do DV ({dv_txt}). São válidos.", "normal"),
        ("", "", "normal"),
        ("USO E ATUALIZAÇÃO", "", "secao"),
        ("Power BI — modelo", "Importe tbl_Municipios, tbl_Distritos, tbl_Regioes_Imediatas, tbl_Regioes_Intermediarias e tbl_UFs "
                              "e desligue a detecção automática de relacionamentos. Relacione (N:1): Distritos[Código IBGE] → "
                              "Municipios[Código IBGE]; Municipios[Código Região Imediata] → Regioes_Imediatas; "
                              "Regioes_Imediatas[Código Região Intermediária] → Regioes_Intermediarias; "
                              "Regioes_Intermediarias[Código UF] → UFs. Marque Latitude/Longitude com a categoria de dados "
                              "correspondente, colunas de código como 'Não resumir' e use 'Município - UF' em filtros e "
                              "rótulos (há nomes repetidos).", "normal"),
        ("Como atualizar", "Baixe as novas edições oficiais do IBGE (DTB, Localidades, Malha Municipal — endereços na aba Fontes), "
                           "substitua os arquivos no diretório de fontes e rode: python build_municipios_ibge.py --fontes <dir> "
                           "--saida <dir>. Se qualquer validação crítica falhar, o resultado vai para "
                           f"{ARQUIVO_FALHA} e o arquivo anterior não é sobrescrito.", "normal"),
    ]
    escrever_texto(ws_leia, texto)
    wb.active = wb.sheetnames.index("Municipios")
    for w in wb.worksheets:
        w.sheet_view.tabSelected = (w.title == "Municipios")
    wb.calculation.fullCalcOnLoad = True  # Excel recalcula tudo ao abrir

    # ---------- 7.11 Gravação atômica + pós-processamento ----------
    falhas = V.falhas_criticas()
    tmp = args.saida / (".tmp_" + ARQUIVO_SAIDA)
    wb.save(tmp)
    if not args.sem_recalculo and recalcular_com_libreoffice(tmp):
        log.info("Pós-processamento XML: %s", pos_processar_xml(tmp))
        n_form, ruins_f, falhas_f = checar_formulas(tmp, "Validacao", col_resultado)
        log.info("Fórmulas: %d | sem valor/erro: %d %s | checagens 'FALHA' no Excel: %s", n_form, len(ruins_f),
                 ruins_f[:10], falhas_f)
        if ruins_f:
            falhas.append(("Excel", "Fórmulas com erro após recálculo", 0, len(ruins_f), False, True))
        if falhas_f:
            falhas.append(("Excel", "Checagens com fórmula retornaram FALHA", 0, len(falhas_f), False, True))
    if falhas:
        destino = args.saida / ARQUIVO_FALHA
        os.replace(tmp, destino)
        for f in falhas:
            log.error("VALIDAÇÃO CRÍTICA FALHOU: %s — %s (esperado %s, obtido %s)", f[0], f[1], f[2], f[3])
        log.error("Arquivo de diagnóstico: %s (o entregável anterior foi preservado)", destino)
        return 1
    destino = args.saida / ARQUIVO_SAIDA
    os.replace(tmp, destino)
    log.info("Excel salvo: %s", destino)
    log.info("Build concluído sem falhas críticas (%d checagens + %d fórmulas de validação).",
             len(V.resultados), len(formulas))
    return 0


if __name__ == "__main__":
    sys.exit(main())
