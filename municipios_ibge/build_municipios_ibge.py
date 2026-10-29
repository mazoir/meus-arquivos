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
     - Latitude/Longitude da sede municipal (ponto oficial IBGE), com fonte e
       status de auditoria por linha
     - Centroide geométrico e ponto interno do polígono municipal (Malha 2025)
     - Coordenada da sede distrital (vila), quando houver ponto oficial IBGE

 QUANDO USAR
   - Para (re)gerar a base de referência territorial usada em BI/geomarketing.
   - Sempre que o IBGE publicar nova DTB / Localidades / Malha: troque os
     arquivos em --fontes, ajuste o bloco CONFIG (FONTES) e rode novamente.

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

 PRÉ-REQUISITOS
   Python >= 3.10 | pandas >= 2.0 | openpyxl >= 3.1 | shapely >= 2.0 |
   numpy | pyarrow (opcional: só para a checagem cruzada com o parquet da API)
   pip install pandas openpyxl shapely numpy pyarrow

 EXEMPLO DE EXECUÇÃO
   python build_municipios_ibge.py --fontes ./fontes --saida ./saida

   Exemplo de linha de saída (aba Municipios):
   3106200 | 310620 | Belo Horizonte | MG | 31 | Minas Gerais | Sudeste |
   3101 | Belo Horizonte | 310001 | Belo Horizonte | Sim | 2 | 310620005 |
   Belo Horizonte; Venda Nova | -19.9208 | -43.9378 | IBGE Localidades 2022 ...

 PERFORMANCE / BOAS PRÁTICAS
   - Tudo em memória (≈16 mil linhas): execução típica < 2 min, dominada pelo
     cálculo de polylabel/point-in-polygon com shapely (geometrias preparadas
     + STRtree para busca espacial).
   - Nenhum código baixado é executado: arquivos ODS/KML/TopoJSON são lidos
     com parsers próprios (zipfile + ElementTree / regex / json).
   - O build ABORTA (exit 1) se qualquer validação crítica falhar.
===============================================================================
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import gzip
import hashlib
import json
import logging
import math
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd
import shapely
from openpyxl import Workbook
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

# Valores oficiais esperados (IBGE). Mudam a cada edição da DTB.
ESPERADO = {
    "municipios": 5571,      # DTB 2025 (inclui DF e Fernando de Noronha)
    "distritos": 10751,      # DTB 2025
    "ufs": 27,
    "reg_imediatas": 510,
    "reg_intermediarias": 133,
}

# Regra de saneamento das coordenadas de sede (km).
LIMIAR_DIVERGENCIA_KM = 5.0    # IBGE 2022 é tido como erro se divergir > 5 km
                               # do IBGE 2010 E da base independente ao mesmo tempo
TOLERANCIA_POLIGONO_KM = 1.0   # malha simplificada: até ~1 km fora é ruído de borda
LIMIAR_AUDITORIA_KM = 5.0      # linhas com qualquer divergência > 5 km vão p/ auditoria

# Caixa envolvente do Brasil (graus; inclui ilhas oceânicas)
BBOX_BRASIL = {"lat_min": -34.0, "lat_max": 5.5, "lon_min": -74.2, "lon_max": -28.5}

GRANDES_REGIOES = {1: "Norte", 2: "Nordeste", 3: "Sudeste", 4: "Sul", 5: "Centro-Oeste"}

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
        "uso": "CHECAGEM CRUZADA da extração A (devem ser idênticas)",
        "origem": "github.com/brasil-em-camadas/brasil-em-camadas.github.io @2e9cbdbd (scripts/dados/sedes_municipais.csv)",
        "oficial": "https://geoftp.ibge.gov.br/organizacao_do_territorio/estrutura_territorial/localidades/"
                   "Localidades_do_Brasil/2022/Localidades_UFs_shp.zip",
    },
    "vilas_2022_go": {
        "arquivo": "localidades_2022/vilas/GO",
        "produto": "IBGE — Localidades do Brasil 2022 — KML por município (GO), pontos 'Vila' (sede distrital)",
        "safra": "Base Censo 2022",
        "uso": "Coordenada da sede distrital (GO)",
        "origem": "github.com/johnerick-py/map_select @69477491 (goias_kml/)",
        "oficial": "https://geoftp.ibge.gov.br/organizacao_do_territorio/estrutura_territorial/localidades/",
    },
    "vilas_2022_sc": {
        "arquivo": "localidades_2022/vilas/SC",
        "produto": "IBGE — Localidades do Brasil 2022 — KML por município (SC), pontos 'Vila' (sede distrital)",
        "safra": "Base Censo 2022",
        "uso": "Coordenada da sede distrital (SC)",
        "origem": "github.com/FelipeFDN/Projeto-ML-Seguranca @f5501208 (data/IBGE/SC/)",
        "oficial": "https://geoftp.ibge.gov.br/organizacao_do_territorio/estrutura_territorial/localidades/",
    },
    "vilas_2022_norte": {
        "arquivo": "localidades_2022/vilas/AM_AP_PA_RR",
        "produto": "IBGE — Localidades do Brasil 2022 — GeoJSON (AM, AP, PA, RR), pontos 'Vila' (sede distrital)",
        "safra": "Base Censo 2022",
        "uso": "Coordenada da sede distrital (AM, AP, PA, RR)",
        "origem": "github.com/Charles-Igeotecnologia/infovias @d9f9e88d (localidades_por_uf/)",
        "oficial": "https://geoftp.ibge.gov.br/organizacao_do_territorio/estrutura_territorial/localidades/",
    },
    "localidades_2010": {
        "arquivo": "localidades_2010/BR_Localidades_2010_v1.csv",
        "produto": "IBGE — Cadastro de Localidades Selecionadas 2010 (BR_Localidades_2010_v1)",
        "safra": "Base Censo 2010",
        "uso": "FALLBACK de sede municipal (pontos 2022 com erro) e sede distrital (demais UFs)",
        "origem": "github.com/GusFurtado/dab_assets @5c1ef16d (data/coordenadas.csv)",
        "oficial": "https://geoftp.ibge.gov.br/organizacao_do_territorio/estrutura_territorial/localidades/"
                   "cadastro_de_localidades_selecionadas_2010/",
    },
    "malha_2025": {
        "arquivo": "malha_2025/BR_Municipios_2025_simplificada.topojson.gz",
        "produto": "IBGE — Malha Municipal 2025 (BR_Municipios_2025), simplificada (Visvalingam 5%, topologia preservada)",
        "safra": "Malha 2025",
        "uso": "Centroide, ponto interno e teste 'sede dentro do limite municipal'",
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
    "nenhuma": "Sem ponto oficial de sede distrital",
}

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


def norm_nome(s: str) -> str:
    """Normaliza nomes para comparação (sem acento, minúsculo, só [a-z0-9])."""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]", "", s)


def haversine_km(lat1, lon1, lat2, lon2):
    """Distância geodésica aproximada (esfera R=6371,0088 km). Aceita escalares ou arrays."""
    lat1, lon1, lat2, lon2 = map(np.radians, (lat1, lon1, lat2, lon2))
    a = np.sin((lat2 - lat1) / 2) ** 2 + np.cos(lat1) * np.cos(lat2) * np.sin((lon2 - lon1) / 2) ** 2
    return 2 * 6371.0088 * np.arcsin(np.sqrt(a))


def cod_str(x) -> str:
    """Normaliza códigos que podem vir como float ('150010710.0') ou int."""
    s = str(x).strip()
    return s[:-2] if s.endswith(".0") else s


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


def carregar_dtb_municipios(path: Path) -> pd.DataFrame:
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
    for c in df.columns:
        df[c] = df[c].str.strip()
    log.info("DTB 2025 municípios: %d linhas (%s)", len(df), data_base)
    return df


def carregar_dtb_distritos(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path, dtype=str, encoding="utf-8")
    df = df.rename(columns={"distrito_code": "cod_distrito", "nome": "distrito", "ibge_code": "cod_mun",
                            "municipio_nome": "municipio_dtb_dist", "uf": "uf_dtb_dist"})
    for c in df.columns:
        df[c] = df[c].str.strip()
    log.info("DTB 2025 distritos: %d linhas", len(df))
    return df


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


def carregar_localidades_2010(path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    df = pd.read_csv(path, sep=";", dtype=str, encoding="utf-8")
    for c in ("CD_GEOCODMU", "CD_GEOCODDS"):
        df[c] = df[c].map(cod_str)
    df["LAT"] = df["LAT"].astype(float)
    df["LONG"] = df["LONG"].astype(float)
    cid = df[df["NM_CATEGORIA"] == "CIDADE"].drop_duplicates("CD_GEOCODMU")
    cid = cid.rename(columns={"CD_GEOCODMU": "cod_mun", "LAT": "lat_2010", "LONG": "lon_2010",
                              "CD_GEOCODDS": "cod_distrito_2010"})[["cod_mun", "lat_2010", "lon_2010", "cod_distrito_2010"]]
    vil = df[df["NM_CATEGORIA"] == "VILA"]
    dup = vil["CD_GEOCODDS"].duplicated().sum()
    if dup:
        log.warning("Localidades 2010: %d vilas duplicadas por distrito — mantida a 1ª ocorrência", dup)
    vil = vil.drop_duplicates("CD_GEOCODDS").rename(columns={
        "CD_GEOCODDS": "cod_distrito", "LAT": "lat", "LONG": "lon", "NM_LOCALIDADE": "nome_fonte"})
    vil = vil[["cod_distrito", "lat", "lon", "nome_fonte"]].assign(fonte="vila2010")
    log.info("Localidades 2010: %d cidades, %d vilas", len(cid), len(vil))
    return cid, vil


_RE_PLACEMARK = re.compile(r"<Placemark.*?</Placemark>", re.S)
_RE_SIMPLEDATA = re.compile(r'<SimpleData name="([^"]+)">([^<]*)</SimpleData>')
_RE_COORD = re.compile(r"<Point>.*?<coordinates>([^<]*)</coordinates>", re.S)


def carregar_vilas_2022(fontes: Path) -> pd.DataFrame:
    """Pontos 'Vila' (sede distrital) do IBGE Localidades 2022 — UFs com espelho disponível.
    Usa a GEOMETRIA (precisão total), não os atributos arredondados a 4 casas."""
    regs = []
    for chave in ("vilas_2022_go", "vilas_2022_sc"):
        for kml in sorted((fontes / FONTES[chave]["arquivo"]).glob("*.kml")):
            texto = kml.read_text(encoding="utf-8")
            for pm in _RE_PLACEMARK.findall(texto):
                d = dict(_RE_SIMPLEDATA.findall(pm))
                if d.get("CT_LOCALIDADE") != "Vila":
                    continue
                m = _RE_COORD.search(pm)
                if not m:
                    continue
                lon, lat = (float(v) for v in m.group(1).strip().split(",")[:2])
                regs.append((cod_str(d["CD_LOCALIDADE"]), lat, lon, d.get("NM_LOCALIDADE", "")))
    for gj in sorted((fontes / FONTES["vilas_2022_norte"]["arquivo"]).glob("*.geojson")):
        for f in json.load(open(gj, encoding="utf-8"))["features"]:
            p = f["properties"]
            if p.get("CT_LOCALIDADE") != "Vila" or not f.get("geometry"):
                continue
            lon, lat = f["geometry"]["coordinates"][:2]
            regs.append((cod_str(p["CD_LOCALIDADE"]), float(lat), float(lon), p.get("NM_LOCALIDADE", "")))
    df = pd.DataFrame(regs, columns=["cod_distrito", "lat", "lon", "nome_fonte"])
    dup = df["cod_distrito"].duplicated().sum()
    if dup:
        log.warning("Vilas 2022: %d códigos duplicados — mantida a 1ª ocorrência", dup)
    df = df.drop_duplicates("cod_distrito").assign(fonte="vila2022")
    log.info("Vilas 2022 (GO, SC, AM, AP, PA, RR): %d pontos | por UF: %s", len(df),
             dict(Counter(df["cod_distrito"].str[:2])))
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


def derivar_pontos_malha(geoms: dict) -> pd.DataFrame:
    """Centroide (equivalente de área) e ponto interno (polylabel, garantido dentro do polígono)."""
    regs = []
    for cod, g in geoms.items():
        c0 = g.centroid
        fwd, inv = _laea(c0.x, c0.y)
        gp = shapely.transform(g, fwd)
        cx, cy = inv(gp.centroid.x, gp.centroid.y)
        maior = max(gp.geoms, key=lambda q: q.area) if gp.geom_type == "MultiPolygon" else gp
        pl = polylabel(maior, tolerance=0.01)  # 10 m
        px, py = inv(pl.x, pl.y)
        regs.append({"cod_mun": cod, "lat_centroide": round(cy, 6), "lon_centroide": round(cx, 6),
                     "centroide_dentro": bool(g.covers(Point(cx, cy))),
                     "lat_interno": round(py, 6), "lon_interno": round(px, 6),
                     "interno_dentro": bool(g.covers(Point(px, py)))})
    return pd.DataFrame(regs)


@dataclass
class TestePoligono:
    dentro: bool
    dist_km: float          # 0 se dentro; senão, distância até o próprio polígono
    cai_em: str             # código do município onde o ponto cai ('' se nenhum)


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


def escolher_coordenada_sede(row) -> tuple[float, float, str, str, str]:
    """Regra de saneamento (documentada na aba LEIA-ME):
    - Base: IBGE Localidades 2022 (ponto oficial mais recente).
    - Se o ponto 2022 divergir > LIMIAR km do IBGE 2010 E da base independente
      (2 fontes independentes concordam que 2022 está fora), usa IBGE 2010.
    - Se, após a troca, 2010 e a base independente ainda divergirem > LIMIAR km,
      marca 'A conferir'.
    Retorna (lat, lon, chave_fonte, status, observação)."""
    d_22_10, d_22_k, d_10_k = row["km_2022_2010"], row["km_2022_kelvins"], row["km_2010_kelvins"]
    if pd.isna(row["lat_2022"]):
        if pd.notna(row["lat_2010"]):
            return row["lat_2010"], row["lon_2010"], "2010", "Corrigida", "Sem ponto 2022; usado IBGE 2010"
        return np.nan, np.nan, "", "Sem coordenada", "Nenhuma fonte oficial"
    if (pd.notna(d_22_10) and pd.notna(d_22_k)
            and d_22_10 > LIMIAR_DIVERGENCIA_KM and d_22_k > LIMIAR_DIVERGENCIA_KM):
        obs = (f"Ponto IBGE 2022 ({row['lat_2022']:.4f}, {row['lon_2022']:.4f}) divergia {d_22_10:.1f} km do "
               f"IBGE 2010 e {d_22_k:.1f} km da base independente; substituído pelo IBGE 2010")
        status = "Corrigida"
        if pd.notna(d_10_k) and d_10_k > LIMIAR_DIVERGENCIA_KM:
            status = "Corrigida - a conferir"
            obs += f" (IBGE 2010 x base independente: {d_10_k:.1f} km)"
        return row["lat_2010"], row["lon_2010"], "2010", status, obs
    return row["lat_2022"], row["lon_2022"], "2022", "OK", ""


# =============================================================================
# 5. FUNÇÕES — validações (testes)
# =============================================================================
class Validador:
    def __init__(self):
        self.resultados = []  # (grupo, verificação, esperado, obtido, ok, crítico)

    def checar(self, grupo, nome, esperado, obtido, ok=None, critico=True):
        ok = (esperado == obtido) if ok is None else bool(ok)
        self.resultados.append((grupo, nome, str(esperado), str(obtido), ok, critico))
        (log.info if ok else (log.error if critico else log.warning))(
            "[%s] %s | esperado=%s | obtido=%s | %s", grupo, nome, esperado, obtido, "OK" if ok else "FALHA")
        return ok

    def falhas_criticas(self):
        return [r for r in self.resultados if r[5] and not r[4]]


# =============================================================================
# 6. FUNÇÕES — escrita do Excel
# =============================================================================
FMT_COORD = "0.0000##"
FMT_INT = "0"
FMT_KM = "#,##0.0"
COR_CABECALHO = "1F4E78"


def escrever_tabela(ws, df: pd.DataFrame, nome_tabela: str, larguras: dict, formatos: dict,
                    comentarios: dict | None = None, freeze: str = "A2", estilo="TableStyleMedium2"):
    ws.append(list(df.columns))
    for linha in df.itertuples(index=False, name=None):
        ws.append([None if (isinstance(v, float) and math.isnan(v)) else v for v in linha])
    n_lin, n_col = len(df) + 1, len(df.columns)
    ref = f"A1:{get_column_letter(n_col)}{n_lin}"
    tab = Table(displayName=nome_tabela, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(name=estilo, showRowStripes=True)
    ws.add_table(tab)
    for j, col in enumerate(df.columns, start=1):
        letra = get_column_letter(j)
        ws.column_dimensions[letra].width = larguras.get(col, 14)
        hc = ws.cell(row=1, column=j)
        hc.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        hc.fill = PatternFill("solid", fgColor=COR_CABECALHO)
        hc.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        if comentarios and col in comentarios:
            hc.comment = Comment(comentarios[col], "Pipeline IBGE")
        fmt = formatos.get(col)
        if fmt:
            for (c,) in ws.iter_rows(min_row=2, max_row=n_lin, min_col=j, max_col=j):
                c.number_format = fmt
    ws.row_dimensions[1].height = 42
    ws.freeze_panes = freeze
    return n_lin


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


def reaplicar_estilo_tabelas(xlsx: Path, estilo: str = "TableStyleMedium2") -> int:
    """O LibreOffice descarta <tableStyleInfo> ao regravar; reinserimos direto no XML de cada tabela
    (posição válida no schema: após <tableColumns>, antes de </table>)."""
    tag = (f'<tableStyleInfo name="{estilo}" showFirstColumn="0" showLastColumn="0" '
           f'showRowStripes="1" showColumnStripes="0"/></table>').encode()
    tmp = xlsx.with_name(xlsx.stem + "_tmp.xlsx")
    n = 0
    with zipfile.ZipFile(xlsx) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            dados = zin.read(item.filename)
            if re.fullmatch(r"xl/tables/table\d+\.xml", item.filename) and b"tableStyleInfo" not in dados:
                dados = dados.replace(b"</table>", tag)
                n += 1
            zout.writestr(item, dados)
    tmp.replace(xlsx)
    return n


def contar_erros_formula(xlsx: Path) -> tuple[int, int, list]:
    """Após recálculo: (nº de fórmulas, nº de células com erro, amostra de erros)."""
    from openpyxl import load_workbook
    erros = ("#N/A", "#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!")
    wb_f = load_workbook(xlsx)
    wb_v = load_workbook(xlsx, data_only=True)
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
    return n_form, len(ruins), ruins[:20]


def escrever_texto(ws, linhas: list[tuple], largura_a=26, largura_b=120):
    """Linhas no formato (texto_col_A, texto_col_B, estilo) com estilo em {'titulo','secao','normal','nota'}."""
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
        elif estilo == "nota":
            ca.font = Font(name="Arial", bold=True, size=10)
            cb.font = Font(name="Arial", size=10, italic=True)
        else:
            ca.font = Font(name="Arial", bold=True, size=10)
            cb.font = Font(name="Arial", size=10)
    ws.sheet_view.showGridLines = False


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
    dtb = carregar_dtb_municipios(F / FONTES["dtb_municipios"]["arquivo"])
    dist = carregar_dtb_distritos(F / FONTES["dtb_distritos"]["arquivo"])
    api_mun, api_reg, api_ufs = carregar_api(F)
    api_dist = carregar_api_distritos(F / FONTES["api_distritos"]["arquivo"])
    s22a, s22b = carregar_sedes_2022(F)
    c10, v10 = carregar_localidades_2010(F / FONTES["localidades_2010"]["arquivo"])
    v22 = carregar_vilas_2022(F)
    kel = carregar_kelvins(F / FONTES["kelvins"]["arquivo"])
    malha = carregar_malha(F / FONTES["malha_2025"]["arquivo"])

    # ---------- 7.2 Municípios + hierarquia ----------
    ufs = api_ufs.rename(columns={"code": "cod_uf", "uf": "uf", "name": "estado_api", "regionName": "regiao",
                                  "capital": "cod_capital"})[["cod_uf", "uf", "estado_api", "regiao", "cod_capital"]]
    mun = dtb.merge(ufs, on="cod_uf", how="left", validate="many_to_one")
    mun["cod_regiao"] = mun["cod_uf"].str[0].astype(int)

    V.checar("Contagem", "Municípios na DTB 2025", ESPERADO["municipios"], len(mun))
    V.checar("Contagem", "Códigos de município únicos", len(mun), mun["cod_mun"].nunique())
    V.checar("Contagem", "UFs", ESPERADO["ufs"], mun["cod_uf"].nunique())
    V.checar("Contagem", "Regiões Imediatas", ESPERADO["reg_imediatas"], mun["cod_rgi"].nunique())
    V.checar("Contagem", "Regiões Intermediárias", ESPERADO["reg_intermediarias"], mun["cod_rgint"].nunique())
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

    # Checagem cruzada DTB 2025 x API IBGE (snapshot set/2026)
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
    V.checar("DTB x API", "Capitais iguais à flag da API", set(), set(mun.loc[mun.capital == "Sim", "cod_mun"]) ^ capitais_api)

    # ---------- 7.3 Distritos ----------
    dist["sufixo"] = dist["cod_distrito"].str[7:].astype(int)
    V.checar("Distritos", "Distritos na DTB 2025", ESPERADO["distritos"], len(dist))
    V.checar("Distritos", "Códigos de distrito únicos", len(dist), dist["cod_distrito"].nunique())
    V.checar("Distritos", "Código do distrito com 9 dígitos", 0, int((~dist["cod_distrito"].str.fullmatch(r"\d{9}")).sum()))
    V.checar("Distritos", "Código do distrito começa com o código do município", 0,
             int((dist["cod_distrito"].str[:7] != dist["cod_mun"]).sum()))
    V.checar("Distritos", "Municípios sem distrito", 0, len(set(mun["cod_mun"]) - set(dist["cod_mun"])))
    V.checar("Distritos", "Distritos com município inexistente", 0, len(set(dist["cod_mun"]) - set(mun["cod_mun"])))
    nome_mun = dict(zip(mun["cod_mun"], mun["municipio"]))
    V.checar("Distritos", "Nome do município no relatório de distritos = relatório de municípios", 0,
             int((dist["municipio_dtb_dist"] != dist["cod_mun"].map(nome_mun)).sum()))
    V.checar("Distritos", "Nomes de distrito com espaços sobrando", 0,
             int((dist["distrito"] != dist["distrito"].str.strip()).sum()))
    if api_dist is not None:
        a = dist.merge(api_dist[["cod_distrito", "distrito_bairro"]], on="cod_distrito", how="outer", indicator=True)
        V.checar("DTB x API", "Distritos só em uma das fontes (API out/2026)", 0, int((a["_merge"] != "both").sum()))
        a = a[a["_merge"] == "both"]
        V.checar("DTB x API", "Nome de distrito divergente (API out/2026)", 0, int((a["distrito"] != a["distrito_bairro"]).sum()))

    # Distrito-sede: convenção IBGE = sufixo 05. Exceção: São Paulo (distritos 01..96, sem sede codificada).
    sem_sede_codificada = set(dist.loc[dist["sufixo"] < 5, "cod_mun"])
    dist["distrito_sede"] = np.where(dist["cod_mun"].isin(sem_sede_codificada), "Não se aplica",
                                     np.where(dist["sufixo"] == 5, "Sim", "Não"))
    log.info("Municípios sem distrito-sede codificado (sufixos 01-04 em uso): %s", sorted(sem_sede_codificada))
    V.checar("Distritos", "Municípios com exatamente 1 distrito-sede (sufixo 05), exceto sem sede codificada",
             len(mun) - len(sem_sede_codificada),
             int(dist[dist["distrito_sede"] == "Sim"].groupby("cod_mun").size().eq(1).sum()))
    c10_suf = c10[~c10["cod_mun"].isin(sem_sede_codificada)]
    V.checar("Distritos", "Sede IBGE 2010 (Cidade) cai no distrito de sufixo 05", 0,
             int((c10_suf["cod_distrito_2010"].str[7:] != "05").sum()))

    # ---------- 7.4 Coordenadas da sede municipal ----------
    s = mun[["cod_mun", "municipio", "uf"]].merge(s22a[["cod_mun", "lat_2022", "lon_2022", "cat_2022"]],
                                                    on="cod_mun", how="left")
    s = s.merge(s22b[["cod_mun", "lat_b", "lon_b"]], on="cod_mun", how="left")
    s = s.merge(c10[["cod_mun", "lat_2010", "lon_2010"]], on="cod_mun", how="left")
    s = s.merge(kel, on="cod_mun", how="left")
    V.checar("Coordenadas", "Municípios com ponto IBGE 2022", len(mun), int(s["lat_2022"].notna().sum()))
    d_ab = haversine_km(s["lat_2022"], s["lon_2022"], s["lat_b"], s["lon_b"])
    V.checar("Coordenadas", "Extrações 2022 A x B: maior diferença (m)", "< 1",
             f"{np.nanmax(d_ab) * 1000:.2f}", ok=np.nanmax(d_ab) * 1000 < 1)
    s["km_2022_2010"] = haversine_km(s["lat_2022"], s["lon_2022"], s["lat_2010"], s["lon_2010"])
    s["km_2022_kelvins"] = haversine_km(s["lat_2022"], s["lon_2022"], s["lat_k"], s["lon_k"])
    s["km_2010_kelvins"] = haversine_km(s["lat_2010"], s["lon_2010"], s["lat_k"], s["lon_k"])
    escolha = s.apply(escolher_coordenada_sede, axis=1, result_type="expand")
    escolha.columns = ["lat", "lon", "fonte_chave", "status_coord", "obs_coord"]
    s = pd.concat([s, escolha], axis=1)
    s["lat"] = s["lat"].astype(float).round(6)
    s["lon"] = s["lon"].astype(float).round(6)
    s["fonte_coord"] = s["fonte_chave"].map(ROTULO_FONTE)
    noronha = s["cat_2022"].astype(str).str.contains("Fernando de Noronha", na=False)
    s.loc[noronha & (s["fonte_chave"] == "2022"), "fonte_coord"] = "IBGE Localidades 2022 (Distrito Estadual)"

    idx = IndiceEspacial(malha)
    tp = [idx.testar(c, la, lo) for c, la, lo in zip(s["cod_mun"], s["lat"], s["lon"])]
    s["dentro_poligono"] = [t.dentro for t in tp]
    s["dist_poligono_km"] = [t.dist_km for t in tp]
    s["cai_em"] = [t.cai_em for t in tp]
    fora_tol = s[(~s["dentro_poligono"]) & (s["dist_poligono_km"] > TOLERANCIA_POLIGONO_KM)]
    s.loc[fora_tol.index, "status_coord"] = s.loc[fora_tol.index, "status_coord"].replace("OK", "A conferir")
    s["sede_no_limite"] = np.where(s["dentro_poligono"], "Sim",
                                   np.where(s["dist_poligono_km"] <= TOLERANCIA_POLIGONO_KM, "Na borda (< 1 km)", "Não"))
    log.info("Coordenadas: %s", dict(Counter(s["status_coord"])))
    V.checar("Coordenadas", "Municípios sem coordenada final", 0, int(s["lat"].isna().sum()))
    V.checar("Coordenadas", "Coordenadas fora da caixa envolvente do Brasil", 0, int((
        ~s["lat"].between(BBOX_BRASIL["lat_min"], BBOX_BRASIL["lat_max"])
        | ~s["lon"].between(BBOX_BRASIL["lon_min"], BBOX_BRASIL["lon_max"])).sum()))
    V.checar("Coordenadas", f"Sede fora do próprio polígono por mais de {TOLERANCIA_POLIGONO_KM:.0f} km (Malha 2025)",
             0, len(fora_tol))
    V.checar("Coordenadas", "Sede dentro do próprio polígono (estrito, malha simplificada)", "≥ 99,5%",
             f"{s['dentro_poligono'].mean():.2%}", ok=s["dentro_poligono"].mean() >= 0.995, critico=False)
    V.checar("Coordenadas", "Pontos IBGE 2022 substituídos pelo IBGE 2010 (regra 2-de-3)", "informativo",
             int((s["fonte_chave"] == "2010").sum()), ok=True, critico=False)
    d_final_k = haversine_km(s["lat"], s["lon"], s["lat_k"], s["lon_k"])
    V.checar("Coordenadas", "Mediana da distância final x base independente (km)", "< 1",
             f"{np.nanmedian(d_final_k):.3f}", ok=np.nanmedian(d_final_k) < 1, critico=False)

    # ---------- 7.5 Pontos derivados da malha ----------
    pm = derivar_pontos_malha(malha)
    V.checar("Malha", "Municípios com polígono na Malha 2025", len(mun), int(pm["cod_mun"].isin(mun["cod_mun"]).sum()))
    V.checar("Malha", "Ponto interno dentro do polígono", len(pm), int(pm["interno_dentro"].sum()))
    V.checar("Malha", "Centroide dentro do polígono (informativo: formas côncavas ficam fora)", "informativo",
             int(pm["centroide_dentro"].sum()), ok=True, critico=False)

    # ---------- 7.6 Coordenadas dos distritos ----------
    coord_mun = s.set_index("cod_mun")[["lat", "lon", "fonte_coord"]]
    dist = dist.merge(coord_mun.rename(columns={"lat": "lat_mun", "lon": "lon_mun", "fonte_coord": "fonte_mun"}),
                      left_on="cod_mun", right_index=True, how="left")
    # Cascata por distrito não-sede: Vila 2022 -> Vila 2010 -> vazio. Cada candidato só é aceito se
    # cair no polígono do próprio município (tolerância de borda). Pega erros do próprio IBGE, p.ex.
    # Marcílio Dias/Canoinhas-SC, cuja Vila 2022 tem latitude -2,1355 em vez de -26,1355.
    cand = {"vila2022": v22.set_index("cod_distrito"), "vila2010": v10.set_index("cod_distrito")}
    eh_sede = dist["distrito_sede"] == "Sim"
    escolhas, rejeitados = [], 0
    for cod_d, cod_m, sede_flag in zip(dist["cod_distrito"], dist["cod_mun"], eh_sede):
        if sede_flag:
            escolhas.append((np.nan, np.nan, None, None, ""))
            continue
        obs = []
        achou = None
        for chave in ("vila2022", "vila2010"):
            if cod_d not in cand[chave].index:
                continue
            r = cand[chave].loc[cod_d]
            t = idx.testar(cod_m, r["lat"], r["lon"])
            if t.dentro or t.dist_km <= TOLERANCIA_POLIGONO_KM:
                achou = (r["lat"], r["lon"], chave, r["nome_fonte"])
                break
            rejeitados += 1
            obs.append(f"Ponto {ROTULO_FONTE[chave]} ({r['lat']:.4f}, {r['lon']:.4f}) descartado: caía "
                       f"{t.dist_km:.1f} km fora do município (em {nome_mun.get(t.cai_em, 'área sem município')})")
            log.info("   vila descartada: %s (%s) %.1f km fora; cai em %s", cod_d, chave, t.dist_km, t.cai_em)
        if achou:
            escolhas.append((achou[0], achou[1], achou[2], achou[3], "; ".join(obs)))
        else:
            escolhas.append((np.nan, np.nan, "nenhuma", None, "; ".join(obs)))
    esc = pd.DataFrame(escolhas, columns=["lat", "lon", "fonte", "nome_fonte", "obs_dist"], index=dist.index)
    dist = pd.concat([dist, esc], axis=1)
    log.info("Vilas reprovadas no teste de polígono (fallback aplicado): %d", rejeitados)
    dist.loc[eh_sede, ["lat", "lon"]] = dist.loc[eh_sede, ["lat_mun", "lon_mun"]].values
    dist["fonte_dist"] = np.where(eh_sede, dist["fonte_mun"] + " — sede municipal", dist["fonte"].map(ROTULO_FONTE))
    # Nome divergente entre a fonte do ponto e a DTB 2025 (renomeações) — só informativo
    tem_vila = dist["nome_fonte"].notna() & (~eh_sede) & dist["lat"].notna()
    div = tem_vila & (dist["nome_fonte"].map(norm_nome) != dist["distrito"].map(norm_nome))
    nota_nome = "Nome do ponto na fonte: " + dist.loc[div, "nome_fonte"].astype(str).str.title()
    dist.loc[div, "obs_dist"] = np.where(dist.loc[div, "obs_dist"] != "",
                                         dist.loc[div, "obs_dist"] + "; " + nota_nome, nota_nome)
    dist["lat"] = dist["lat"].astype(float).round(6)
    dist["lon"] = dist["lon"].astype(float).round(6)
    cobertura = dist["lat"].notna().mean()
    V.checar("Distritos", "Distritos com coordenada de sede distrital", "informativo",
             f"{int(dist['lat'].notna().sum())} de {len(dist)} ({cobertura:.1%})", ok=True, critico=False)
    V.checar("Distritos", "Vilas com nome divergente da DTB 2025 (renomeações; ponto mantido)", "informativo",
             int(div.sum()), ok=True, critico=False)

    # ---------- 7.7 Montagem das tabelas finais ----------
    agg = (dist.sort_values("cod_distrito").groupby("cod_mun")
           .agg(qtd=("cod_distrito", "size"), lista=("distrito", lambda v: "; ".join(v))))
    sede = dist[dist["distrito_sede"] == "Sim"].set_index("cod_mun")["cod_distrito"]
    M = mun.merge(s[["cod_mun", "lat", "lon", "fonte_coord", "status_coord", "obs_coord", "sede_no_limite"]], on="cod_mun")
    M = M.merge(pm, on="cod_mun", how="left").merge(agg, left_on="cod_mun", right_index=True, how="left")
    M["cod_sede"] = M["cod_mun"].map(sede)
    M = M.sort_values("cod_mun").reset_index(drop=True)
    tab_mun = pd.DataFrame({
        "Código IBGE": M["cod_mun"].astype(int),
        "Código IBGE (6 dígitos)": M["cod_mun"].str[:6].astype(int),
        "Município": M["municipio"],
        "UF": M["uf"],
        "Código UF": M["cod_uf"].astype(int),
        "Estado": M["estado"],
        "Região": M["regiao"],
        "Código Região Intermediária": M["cod_rgint"].astype(int),
        "Região Geográfica Intermediária": M["nome_rgint"],
        "Código Região Imediata": M["cod_rgi"].astype(int),
        "Região Geográfica Imediata": M["nome_rgi"],
        "Capital": M["capital"],
        "Qtd. Distritos": M["qtd"].astype(int),
        "Código Distrito-Sede": M["cod_sede"].map(lambda v: int(v) if pd.notna(v) else None),
        "Distritos": M["lista"],
        "Latitude": M["lat"],
        "Longitude": M["lon"],
        "Fonte da Coordenada": M["fonte_coord"],
        "Status da Coordenada": M["status_coord"],
        "Sede dentro do Limite Municipal": M["sede_no_limite"],
        "Observação da Coordenada": M["obs_coord"].replace("", None),
        "Latitude Centroide": M["lat_centroide"],
        "Longitude Centroide": M["lon_centroide"],
        "Centroide dentro do Município": np.where(M["centroide_dentro"], "Sim", "Não"),
        "Latitude Ponto Interno": M["lat_interno"],
        "Longitude Ponto Interno": M["lon_interno"],
    })
    sp = tab_mun["Código Distrito-Sede"].isna()
    log.info("Municípios sem distrito-sede codificado na tabela: %s", tab_mun.loc[sp, "Município"].tolist())

    D = dist.merge(mun[["cod_mun", "municipio", "uf", "cod_uf", "estado", "regiao", "cod_rgint", "nome_rgint",
                        "cod_rgi", "nome_rgi"]], on="cod_mun", how="left").sort_values("cod_distrito")
    tab_dist = pd.DataFrame({
        "Código Distrito": D["cod_distrito"].astype(int),
        "Distrito": D["distrito"],
        "Distrito-Sede": D["distrito_sede"],
        "Código IBGE": D["cod_mun"].astype(int),
        "Município": D["municipio"],
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
    tab_uf = (M[["cod_uf", "uf", "estado", "cod_regiao", "regiao"]].drop_duplicates().sort_values("cod_uf"))
    tab_uf = pd.DataFrame({"Código UF": tab_uf["cod_uf"].astype(int), "UF": tab_uf["uf"], "Estado": tab_uf["estado"],
                           "Código Região": tab_uf["cod_regiao"].astype(int), "Região": tab_uf["regiao"],
                           "Capital": tab_uf["cod_uf"].map(cap)})
    V.checar("Dimensões", "Linhas em Regioes_Imediatas", ESPERADO["reg_imediatas"], len(tab_rgi))
    V.checar("Dimensões", "Linhas em Regioes_Intermediarias", ESPERADO["reg_intermediarias"], len(tab_rgint))
    V.checar("Dimensões", "Linhas em UFs", ESPERADO["ufs"], len(tab_uf))

    # Auditoria: municípios com qualquer divergência relevante entre fontes, ou status ≠ OK
    s_aud = s.merge(mun[["cod_mun"]], on="cod_mun")
    cond = ((s_aud[["km_2022_2010", "km_2022_kelvins", "km_2010_kelvins"]] > LIMIAR_AUDITORIA_KM).any(axis=1)
            | (s_aud["status_coord"] != "OK") | (s_aud["sede_no_limite"] != "Sim"))
    aud = s_aud[cond].copy()

    def decisao(r):
        if r["fonte_chave"] == "2010":
            return "Usado IBGE 2010 (ponto 2022 divergente de 2 fontes)"
        if r["sede_no_limite"] != "Sim":
            return "Mantido IBGE 2022 (ponto na borda do limite; efeito da malha simplificada)"
        if r["km_2022_kelvins"] > LIMIAR_AUDITORIA_KM and (pd.isna(r["km_2022_2010"]) or r["km_2022_2010"] <= LIMIAR_AUDITORIA_KM):
            return "Mantido IBGE 2022 (IBGE 2010 confirma; base independente é que diverge)"
        if r["km_2022_2010"] > LIMIAR_AUDITORIA_KM:
            return "Mantido IBGE 2022 (base independente confirma 2022; ponto 2010 desatualizado)"
        return "Mantido IBGE 2022"

    aud["decisao"] = aud.apply(decisao, axis=1)
    aud = aud.sort_values(["status_coord", "km_2022_2010"], ascending=[True, False])
    tab_aud = pd.DataFrame({
        "Código IBGE": aud["cod_mun"].astype(int), "Município": aud["municipio"], "UF": aud["uf"],
        "Status da Coordenada": aud["status_coord"], "Decisão": aud["decisao"],
        "Latitude Final": aud["lat"], "Longitude Final": aud["lon"],
        "Lat. IBGE 2022": aud["lat_2022"], "Lon. IBGE 2022": aud["lon_2022"],
        "Lat. IBGE 2010": aud["lat_2010"].round(6), "Lon. IBGE 2010": aud["lon_2010"].round(6),
        "Lat. Base Independente": aud["lat_k"], "Lon. Base Independente": aud["lon_k"],
        "Dist. 2022 x 2010 (km)": aud["km_2022_2010"].round(2),
        "Dist. 2022 x Independente (km)": aud["km_2022_kelvins"].round(2),
        "Dist. 2010 x Independente (km)": aud["km_2010_kelvins"].round(2),
        "Sede dentro do Limite Municipal": aud["sede_no_limite"],
    })
    log.info("Auditoria de coordenadas: %d municípios listados", len(tab_aud))

    # ---------- 7.8 Validações finais de integridade das tabelas ----------
    V.checar("Tabelas", "Soma de 'Qtd. Distritos' = linhas da aba Distritos", len(tab_dist), int(tab_mun["Qtd. Distritos"].sum()))
    V.checar("Tabelas", "Campos-chave vazios (aba Municipios)", 0,
             int(tab_mun[["Código IBGE", "Município", "UF", "Código Região Imediata", "Código Região Intermediária",
                          "Latitude", "Longitude"]].isna().sum().sum()))
    V.checar("Tabelas", "Campos-chave vazios (aba Distritos)", 0,
             int(tab_dist[["Código Distrito", "Distrito", "Código IBGE", "Município", "UF",
                           "Latitude Município", "Longitude Município"]].isna().sum().sum()))
    falhas = V.falhas_criticas()

    # ---------- 7.9 Escrita do Excel ----------
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
    larg_m = {"Código IBGE": W_COD, "Código IBGE (6 dígitos)": W_COD, "Município": W_NOME, "UF": 6, "Código UF": 8,
              "Estado": 20, "Região": 13, "Código Região Intermediária": 13, "Região Geográfica Intermediária": W_REG,
              "Código Região Imediata": 12, "Região Geográfica Imediata": W_REG, "Capital": 9, "Qtd. Distritos": 10,
              "Código Distrito-Sede": 13, "Distritos": 45, "Latitude": 12, "Longitude": 12,
              "Fonte da Coordenada": 36, "Status da Coordenada": 22, "Sede dentro do Limite Municipal": 16,
              "Observação da Coordenada": 60, "Latitude Centroide": 12, "Longitude Centroide": 12,
              "Centroide dentro do Município": 13, "Latitude Ponto Interno": 12, "Longitude Ponto Interno": 12}
    fmt_m = {"Código IBGE": FMT_INT, "Código IBGE (6 dígitos)": FMT_INT, "Código UF": FMT_INT,
             "Código Região Intermediária": FMT_INT, "Código Região Imediata": FMT_INT, "Qtd. Distritos": FMT_INT,
             "Código Distrito-Sede": FMT_INT, "Latitude": FMT_COORD, "Longitude": FMT_COORD,
             "Latitude Centroide": FMT_COORD, "Longitude Centroide": FMT_COORD,
             "Latitude Ponto Interno": FMT_COORD, "Longitude Ponto Interno": FMT_COORD}
    com_m = {
        "Código IBGE": "Código IBGE de 7 dígitos (com dígito verificador). Chave para relacionamentos.",
        "Código IBGE (6 dígitos)": "Código sem o dígito verificador (padrão DATASUS/SIH/SIM).",
        "Distritos": "Todos os distritos do município (DTB 2025), em ordem de código; o distrito-sede (sufixo 05) vem primeiro.",
        "Código Distrito-Sede": "Distrito de sufixo 05 (convenção IBGE). Vazio em São Paulo/SP, cujos 96 distritos são numerados 01–96 e não têm sede codificada.",
        "Latitude": "Latitude (graus decimais, SIRGAS 2000) do ponto oficial da SEDE MUNICIPAL (IBGE). Ver 'Fonte da Coordenada'.",
        "Longitude": "Longitude (graus decimais, SIRGAS 2000) do ponto oficial da SEDE MUNICIPAL (IBGE).",
        "Status da Coordenada": "OK = ponto IBGE 2022 mantido. Corrigida = ponto 2022 trocado pelo IBGE 2010 (regra 2-de-3). 'a conferir' = fontes ainda divergem > 5 km.",
        "Sede dentro do Limite Municipal": "Teste ponto-no-polígono contra a Malha Municipal IBGE 2025 (simplificada).",
        "Latitude Centroide": "Centroide geométrico (equivalente de área) do polígono municipal (Malha IBGE 2025). Pode cair fora do município em formas côncavas.",
        "Latitude Ponto Interno": "Ponto mais interno do polígono (polo de inacessibilidade). Sempre dentro do município; ideal para rótulos de mapa.",
    }
    escrever_tabela(ws_m, tab_mun, "tbl_Municipios", larg_m, fmt_m, com_m, freeze="D2")

    larg_d = {"Código Distrito": 13, "Distrito": 30, "Distrito-Sede": 13, "Código IBGE": W_COD, "Município": W_NOME,
              "UF": 6, "Código UF": 8, "Estado": 20, "Região": 13, "Código Região Intermediária": 13,
              "Região Geográfica Intermediária": W_REG, "Código Região Imediata": 12,
              "Região Geográfica Imediata": W_REG, "Latitude Município": 12, "Longitude Município": 12,
              "Latitude Distrito": 12, "Longitude Distrito": 12, "Fonte Coordenada Distrito": 48, "Observação": 55}
    fmt_d = {"Código Distrito": FMT_INT, "Código IBGE": FMT_INT, "Código UF": FMT_INT,
             "Código Região Intermediária": FMT_INT, "Código Região Imediata": FMT_INT,
             "Latitude Município": FMT_COORD, "Longitude Município": FMT_COORD,
             "Latitude Distrito": FMT_COORD, "Longitude Distrito": FMT_COORD}
    com_d = {
        "Código Distrito": "Código IBGE do distrito (9 dígitos) = código do município (7) + sufixo (2).",
        "Distrito-Sede": "Sim = distrito-sede (sufixo 05). 'Não se aplica' = São Paulo/SP (distritos 01–96, sem sede codificada).",
        "Latitude Município": "Coordenada da sede municipal (mesma da aba Municipios).",
        "Latitude Distrito": "Coordenada da sede do distrito (ponto 'Vila' do IBGE). No distrito-sede = sede municipal. Vazio quando não há ponto oficial.",
    }
    escrever_tabela(ws_d, tab_dist, "tbl_Distritos", larg_d, fmt_d, com_d, freeze="C2")

    escrever_tabela(ws_rgi, tab_rgi, "tbl_Regioes_Imediatas",
                    {"Código Região Imediata": 13, "Região Geográfica Imediata": W_REG, "Código Região Intermediária": 13,
                     "Região Geográfica Intermediária": W_REG, "UF": 6, "Código UF": 8, "Qtd. Municípios": 11},
                    {"Código Região Imediata": FMT_INT, "Código Região Intermediária": FMT_INT, "Código UF": FMT_INT})
    escrever_tabela(ws_rgint, tab_rgint, "tbl_Regioes_Intermediarias",
                    {"Código Região Intermediária": 13, "Região Geográfica Intermediária": W_REG, "UF": 6,
                     "Código UF": 8, "Qtd. Regiões Imediatas": 12, "Qtd. Municípios": 11},
                    {"Código Região Intermediária": FMT_INT, "Código UF": FMT_INT})
    escrever_tabela(ws_uf, tab_uf, "tbl_UFs", {"Código UF": 9, "UF": 6, "Estado": 22, "Código Região": 9,
                                               "Região": 14, "Capital": 22}, {"Código UF": FMT_INT, "Código Região": FMT_INT})

    # Resumo_UF — fórmulas vivas sobre as abas de dados
    n_m, n_d = len(tab_mun) + 1, len(tab_dist) + 1
    n_rgi, n_rgint = len(tab_rgi) + 1, len(tab_rgint) + 1
    hdr = ["UF", "Estado", "Região", "Municípios", "Distritos", "Regiões Imediatas", "Regiões Intermediárias",
           "Coord. corrigidas (IBGE 2010)", "Média distritos/município"]
    ws_res.append(hdr)
    for i, r in enumerate(tab_uf.itertuples(index=False), start=2):
        ws_res.append([
            r.UF, r.Estado, r.Região,
            f"=COUNTIF(Municipios!$D$2:$D${n_m},$A{i})",
            f"=COUNTIF(Distritos!$F$2:$F${n_d},$A{i})",
            f"=COUNTIF(Regioes_Imediatas!$E$2:$E${n_rgi},$A{i})",
            f"=COUNTIF(Regioes_Intermediarias!$C$2:$C${n_rgint},$A{i})",
            f'=COUNTIFS(Municipios!$D$2:$D${n_m},$A{i},Municipios!$R$2:$R${n_m},"IBGE Localidades 2010*")',
            f"=IF(D{i}=0,0,E{i}/D{i})",
        ])
    tot = len(tab_uf) + 2
    ws_res.append(["TOTAL BRASIL", None, None] + [f"=SUM({c}2:{c}{tot - 1})" for c in "DEFGH"] + [f"=IF(D{tot}=0,0,E{tot}/D{tot})"])
    tab = Table(displayName="tbl_Resumo_UF", ref=f"A1:I{tot - 1}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws_res.add_table(tab)
    for j, w in enumerate([8, 22, 14, 12, 11, 12, 14, 14, 13], start=1):
        ws_res.column_dimensions[get_column_letter(j)].width = w
        c = ws_res.cell(row=1, column=j)
        c.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor=COR_CABECALHO)
        c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    for col in "DEFGH":
        for rr in range(2, tot + 1):
            ws_res[f"{col}{rr}"].number_format = "#,##0"
    for rr in range(2, tot + 1):
        ws_res[f"I{rr}"].number_format = "0.00"
    for j in range(1, 10):
        ws_res.cell(row=tot, column=j).font = Font(name="Arial", bold=True, size=10)
        ws_res.cell(row=tot, column=j).fill = PatternFill("solid", fgColor="D9E1F2")
    ws_res.row_dimensions[1].height = 42
    ws_res.freeze_panes = "A2"
    ws_res.cell(row=tot + 2, column=1, value="Todas as colunas numéricas são fórmulas (COUNTIF/COUNTIFS/SUM) sobre as abas "
                                              "Municipios, Distritos, Regioes_Imediatas e Regioes_Intermediarias.").font = \
        Font(name="Arial", italic=True, size=9)

    escrever_tabela(ws_aud, tab_aud, "tbl_Auditoria_Coordenadas",
                    {"Código IBGE": W_COD, "Município": 28, "UF": 6, "Status da Coordenada": 22, "Decisão": 58,
                     "Sede dentro do Limite Municipal": 16},
                    {"Código IBGE": FMT_INT, **{c: FMT_COORD for c in tab_aud.columns if c.startswith(("Lat", "Lon"))},
                     **{c: "0.00" for c in tab_aud.columns if c.startswith("Dist.")}}, freeze="D2")

    # Validacao — checagens do pipeline + checagens com fórmula viva no próprio Excel
    val_df = pd.DataFrame(V.resultados, columns=["Grupo", "Verificação", "Esperado", "Obtido", "OK", "Crítica"])
    val_df["Resultado"] = np.where(val_df["OK"], "OK", np.where(val_df["Crítica"], "FALHA", "ALERTA"))
    val_df["Crítica"] = np.where(val_df["Crítica"], "Sim", "Não")
    val_df = val_df[["Grupo", "Verificação", "Esperado", "Obtido", "Resultado", "Crítica"]]
    val_df.insert(0, "Origem", "Pipeline Python")
    escrever_tabela(ws_val, val_df, "tbl_Validacao",
                    {"Origem": 16, "Grupo": 14, "Verificação": 70, "Esperado": 18, "Obtido": 26, "Resultado": 11, "Crítica": 8},
                    {}, freeze="A2")
    lin = len(val_df) + 3
    ws_val.cell(row=lin, column=1, value="Checagens recalculadas pelo próprio Excel (fórmulas vivas)").font = \
        Font(name="Arial", bold=True, size=11, color=COR_CABECALHO)
    formulas = [
        ("Linhas na aba Municipios", ESPERADO["municipios"], f"=COUNTA(Municipios!A2:A{n_m})"),
        ("Linhas na aba Distritos", ESPERADO["distritos"], f"=COUNTA(Distritos!A2:A{n_d})"),
        ("Soma de Qtd. Distritos (aba Municipios)", ESPERADO["distritos"], f"=SUM(Municipios!M2:M{n_m})"),
        ("Regiões Imediatas", ESPERADO["reg_imediatas"], f"=COUNTA(Regioes_Imediatas!A2:A{n_rgi})"),
        ("Regiões Intermediárias", ESPERADO["reg_intermediarias"], f"=COUNTA(Regioes_Intermediarias!A2:A{n_rgint})"),
        ("UFs", ESPERADO["ufs"], f"=COUNTA(UFs!A2:A{ESPERADO['ufs'] + 1})"),
        ("Municípios sem latitude", 0, f"=COUNTBLANK(Municipios!P2:P{n_m})"),
        ("Municípios sem longitude", 0, f"=COUNTBLANK(Municipios!Q2:Q{n_m})"),
        ("Latitude fora do Brasil (< -34 ou > 5,5)", 0,
         f'=COUNTIF(Municipios!P2:P{n_m},"<-34")+COUNTIF(Municipios!P2:P{n_m},">5.5")'),
        ("Longitude fora do Brasil (< -74,2 ou > -28,5)", 0,
         f'=COUNTIF(Municipios!Q2:Q{n_m},"<-74.2")+COUNTIF(Municipios!Q2:Q{n_m},">-28.5")'),
        ("Distrito-sede marcado 'Sim'", ESPERADO["municipios"] - len(sem_sede_codificada),
         f'=COUNTIF(Distritos!C2:C{n_d},"Sim")'),
    ]
    ws_val.cell(row=lin + 1, column=2, value="Verificação")
    ws_val.cell(row=lin + 1, column=3, value="Esperado")
    ws_val.cell(row=lin + 1, column=4, value="Obtido (fórmula)")
    ws_val.cell(row=lin + 1, column=5, value="Resultado")
    for j in range(2, 6):
        c = ws_val.cell(row=lin + 1, column=j)
        c.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor=COR_CABECALHO)
    for k, (nome, esp, f) in enumerate(formulas, start=lin + 2):
        ws_val.cell(row=k, column=1, value="Fórmula Excel")
        ws_val.cell(row=k, column=2, value=nome)
        ws_val.cell(row=k, column=3, value=esp)
        ws_val.cell(row=k, column=4, value=f)
        ws_val.cell(row=k, column=5, value=f'=IF(D{k}=C{k},"OK","FALHA")')

    # Fontes — catálogo com hash
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

    # LEIA-ME
    n_corr = int((s["fonte_chave"] == "2010").sum())
    n_conf = int(s["status_coord"].str.contains("conferir").sum())
    cov_d = dist["lat"].notna().sum()
    texto = [
        ("Municípios do Brasil — IBGE", None, "titulo"),
        ("Gerado em", f"{gerado_em} por build_municipios_ibge.py (pipeline reprodutível; fontes e hashes na aba Fontes).", "normal"),
        ("Cobertura", f"{len(tab_mun):,} municípios (DTB 2025 — inclui Brasília/DF e o Distrito Estadual de Fernando de "
                      f"Noronha/PE, que o IBGE codifica como município) · {len(tab_dist):,} distritos · "
                      f"{len(tab_rgi)} Regiões Geográficas Imediatas · {len(tab_rgint)} Intermediárias · 27 UFs.".replace(",", "."), "normal"),
        ("", "", "normal"),
        ("ABAS", "", "secao"),
        ("Municipios", "1 linha por município. Códigos, UF, regiões imediata/intermediária, distritos, coordenada da sede "
                       "(com fonte e status), centroide e ponto interno do polígono.", "normal"),
        ("Distritos", "1 linha por distrito (tabela plana): distrito + município + UF + regiões + coordenada do município e "
                      "da sede distrital.", "normal"),
        ("Regioes_Imediatas / Regioes_Intermediarias / UFs", "Dimensões para modelo estrela (Power BI): relacione pelas colunas de código.", "normal"),
        ("Resumo_UF", "Contagens por UF com fórmulas vivas (COUNTIF/COUNTIFS/SUM).", "normal"),
        ("Auditoria_Coordenadas", f"{len(tab_aud)} municípios em que alguma fonte diverge > {LIMIAR_AUDITORIA_KM:.0f} km, "
                                  "com os valores de cada fonte e a decisão tomada.", "normal"),
        ("Validacao", f"{len(val_df)} checagens automáticas do pipeline + checagens com fórmula viva no Excel.", "normal"),
        ("Fontes", "Produto IBGE, safra, origem do arquivo, endereço oficial e SHA-256 de cada insumo.", "normal"),
        ("", "", "normal"),
        ("COORDENADAS — O QUE SIGNIFICAM", "", "secao"),
        ("Latitude / Longitude", "Ponto oficial da SEDE MUNICIPAL (área urbana da sede) do IBGE — Localidades do Brasil 2022 "
                                 "(categoria 'Cidade'/'Sede Municipal'). Graus decimais, datum SIRGAS 2000 (EPSG:4674). "
                                 "Município é uma área; este é o ponto de referência oficial da cidade-sede, padrão para "
                                 "mapas de pontos, cálculo de distâncias e geocodificação.", "normal"),
        ("Precisão", "O IBGE publica os atributos com 4 casas decimais (≈ 11 m). Nos 14 pontos corrigidos, o IBGE 2010 traz "
                     "precisão total (gravado com 6 casas, ≈ 0,1 m).", "normal"),
        ("Regra de correção", f"O ponto IBGE 2022 é substituído pelo IBGE 2010 somente quando diverge > "
                              f"{LIMIAR_DIVERGENCIA_KM:.0f} km de DUAS fontes independentes ao mesmo tempo (IBGE 2010 e a base "
                              f"comunitária kelvins). Resultado: {n_corr} municípios corrigidos (ex.: Manicoré/AM, cujo ponto 2022 "
                              f"estava 237 km deslocado, numa vila). {n_conf} seguem marcados 'a conferir'. Detalhes na aba "
                              "Auditoria_Coordenadas.", "normal"),
        ("Validação geométrica", "Todas as sedes foram testadas contra o polígono do próprio município (Malha Municipal IBGE "
                                 "2025). Coluna 'Sede dentro do Limite Municipal'. 'Na borda' = até 1 km fora, efeito da "
                                 "simplificação da malha.", "normal"),
        ("Centroide x Ponto Interno", "Centroide = centro de massa do polígono (pode cair fora em municípios côncavos). Ponto "
                                      "interno = polo de inacessibilidade, sempre dentro (melhor para rótulos). Ambos "
                                      "calculados sobre a Malha 2025 simplificada (diferença típica < 0,1 km).", "normal"),
        ("Grandes cidades", "Em metrópoles o ponto IBGE da sede é um ponto representativo da área urbana e pode ficar a "
                            "alguns km do centro histórico (ex.: São Paulo fica ≈ 6 km a leste da Sé). Para análises "
                            "intramunicipais, use os distritos ou dados de setor censitário.", "normal"),
        ("", "", "normal"),
        ("DISTRITOS", "", "secao"),
        ("Distrito-sede", "Convenção do IBGE: o distrito-sede recebe o sufixo 05 (validado contra as 5.564 sedes do IBGE 2010). "
                          "Exceção: São Paulo/SP tem 96 distritos numerados 01–96 sem sede codificada → 'Não se aplica'.", "normal"),
        ("Coordenada do distrito", f"Distrito-sede = sede municipal. Demais distritos = ponto 'Vila' (sede distrital) do IBGE "
                                   f"Localidades 2022 (GO, SC, AM, AP, PA, RR) ou, nas demais UFs, IBGE Localidades 2010. "
                                   f"Cobertura: {cov_d:,} de {len(dist):,} distritos ({cov_d / len(dist):.1%}).".replace(",", ".") +
                                   " Distritos criados após 2010 fora dessas UFs e os distritos urbanos de São Paulo ficam "
                                   "sem coordenada (não há ponto oficial), em vez de receberem um valor aproximado.", "normal"),
        ("", "", "normal"),
        ("ATUALIZAÇÃO", "", "secao"),
        ("Como atualizar", "Baixe as novas edições oficiais do IBGE (DTB, Localidades, Malha Municipal — endereços na aba "
                           "Fontes), substitua os arquivos no diretório de fontes e rode: python build_municipios_ibge.py "
                           "--fontes <dir> --saida <dir>. O build aborta se qualquer validação crítica falhar.", "normal"),
        ("Uso no Power BI", "Importe as tabelas tbl_Municipios, tbl_Distritos, tbl_Regioes_Imediatas, "
                            "tbl_Regioes_Intermediarias e tbl_UFs. Relacione por 'Código IBGE', 'Código Região Imediata', "
                            "'Código Região Intermediária' e 'Código UF' (1:N). Lat/Long já em graus decimais: marque a "
                            "categoria de dados como Latitude/Longitude.", "normal"),
    ]
    escrever_texto(ws_leia, texto)
    wb.active = wb.sheetnames.index("Municipios")
    for w in wb.worksheets:
        w.sheet_view.tabSelected = (w.title == "Municipios")

    wb.calculation.fullCalcOnLoad = True  # Excel recalcula tudo ao abrir
    destino = args.saida / ARQUIVO_SAIDA
    wb.save(destino)
    log.info("Excel salvo: %s", destino)
    if not args.sem_recalculo and recalcular_com_libreoffice(destino):
        log.info("Estilo reaplicado em %d tabelas", reaplicar_estilo_tabelas(destino))
        n_form, n_err, amostra = contar_erros_formula(destino)
        log.info("Fórmulas: %d | células com erro/sem valor: %d %s", n_form, n_err, amostra)
        if n_err:
            falhas.append(("Excel", "Fórmulas com erro após recálculo", 0, n_err, False, True))
    if falhas:
        for f in falhas:
            log.error("VALIDAÇÃO CRÍTICA FALHOU: %s — %s (esperado %s, obtido %s)", f[0], f[1], f[2], f[3])
        return 1
    log.info("Build concluído sem falhas críticas (%d checagens).", len(V.resultados))
    return 0


if __name__ == "__main__":
    sys.exit(main())
