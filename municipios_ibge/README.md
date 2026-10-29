# Municípios do Brasil — IBGE (DTB 2025)

`Municipios_Brasil_IBGE.xlsx`: base territorial de referência com os **5.571 municípios** e os **10.751 distritos** do Brasil.

| Aba | Conteúdo |
|---|---|
| `Municipios` | 1 linha por município: código IBGE (7 e 6 dígitos), município, UF, estado, região, regiões geográficas **imediata** e **intermediária**, **latitude/longitude da sede** (com fonte, status e observação), distritos, centroide e ponto interno do polígono |
| `Distritos` | 1 linha por distrito (tabela plana) com município, UF, regiões e coordenadas do município e da sede distrital |
| `Regioes_Imediatas`, `Regioes_Intermediarias`, `UFs` | Dimensões para modelo estrela (Power BI) |
| `Resumo_UF` | Contagens por UF com fórmulas vivas |
| `Auditoria_Coordenadas` | Municípios em que as fontes divergem, com os valores de cada fonte e a decisão tomada |
| `Validacao` | 61 checagens do pipeline + 11 checagens com fórmula |
| `Fontes` | Produto IBGE, safra, origem e SHA-256 de cada insumo |
| `LEIA-ME` | Metodologia, regras e dicionário |

## Fontes (produtos oficiais do IBGE)

- **DTB 2025** (data-base 31/12/2025): municípios, regiões imediatas e intermediárias, distritos. Bate 100% com a API de Localidades do IBGE (snapshots de set/out 2026).
- **Localidades do Brasil 2022**: ponto oficial da sede municipal. Há 15 pontos saneados com o IBGE Localidades 2010, cada um com a evidência registrada na planilha.
- **Malha Municipal 2025**: teste de ponto-no-polígono, centroide e ponto interno.

## Como regenerar

```bash
pip install pandas openpyxl shapely numpy pyarrow
python build_municipios_ibge.py --fontes ./fontes --saida ./saida
```

O diretório de fontes segue o layout do dicionário `FONTES` no script, onde também estão os endereços oficiais do IBGE. Se uma validação crítica falhar, o script grava `Municipios_Brasil_IBGE_FALHOU.xlsx`, preserva o entregável anterior e sai com código 1.
