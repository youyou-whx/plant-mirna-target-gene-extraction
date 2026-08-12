# Plant miRNA–Target Gene Relationship Extraction Pipeline

A modular Python pipeline for extracting miRNA–target gene pairs from plant literature abstracts, with LLM-assisted verification, species-aware miRNA normalization, gene ID resolution, and a Streamlit web interface.

## Overview

```
┌─────────────────────┐     ┌──────────────────────┐     ┌──────────────────────┐
│  extract_mirna_     │───▶│     llm_review.py     │───▶│  normalize_names.py  │
│  genes.py           │     │  (LLM verification)   │     │  (miRNA normalization│
│  regex extraction   │     │  + manual review      │     │  + dedup + gene IDs) │
└─────────┬───────────┘     └──────────┬───────────┘     └──────────┬───────────┘
          │                            │                            │
          ▼                            ▼                            ▼
  miRNA-Target_Gene_       miRNA-Target_Gene_            miRNA-Target_Gene_
  Pairs.xlsx               Pairs_Reviewed.xlsx           Pairs_Reviewed_Final.xlsx
                                  │
                                  ▼
                      ┌──────────────────────┐
                      │  auto_update_lists.  │
                      │  py                  │
                      │  feedback → blacklist│
                      │  / whitelist rules   │
                      └──────────────────────┘
```

## Features

- **Case-insensitive miRNA detection**: catches `miR156`, `MIR156`, `Mir156`, and legacy forms like `TaMIR5062-5A`, plus `.2` variants (`miR444b.2`)
- **Five-tier gene extraction**: database IDs → species-prefixed genes → bare uppercase+digits → mixed-case genes (`HsfA1`) → gene family names
- **Per-pair sentence selection**: each miRNA–gene pair gets a sentence containing both entities when possible, prioritized by relation signal strength
- **All-occurrence tracking**: tracks every mention of a miRNA or gene in an abstract, avoiding missed co-occurrences from position deduplication
- **Cross-match detection**: identifies crowded sentences with multiple miRNA/gene candidates and flags ambiguous pairings
- **30 plant species** supported out of the box — see [Supported Species](#supported-species)
- **Automatic species detection**: title-first matching with weighted voting across miRNA prefixes, gene prefixes, and text keywords
- **LLM batch review**: sends 5 entries per API call with checkpoint resume; supports DeepSeek, OpenAI, and other OpenAI-compatible providers
- **miRNA normalization**: `OsmiR156` → `osa-miR156`, strips `pre-`/`pri-` precursors, species-aware prefix assignment
- **Gene ID resolution**: maps gene symbols to database IDs — RAP-DB and MSU via [ricedata.cn](https://www.ricedata.cn/gene/) for rice, NCBI Entrez Gene for all other species
- **Manual review support**: Weak association entries are left for human judgment; mark `Yes` in the Manual Keep column to retain
- **Self-improving**: `auto_update_lists.py` feeds LLM review results back into the extraction blacklist/whitelist

## Requirements

- Python 3.8+
- An API key from a supported LLM provider
- Dependencies installed automatically on first run: `openpyxl`, `requests`

## Quick Start

```powershell
$env:LLM_API_KEY = "sk-xxxxxxxx"

# Extract miRNA–gene pairs
python extract_mirna_genes.py abstracts.txt

# LLM review (or extract + auto-review in one step)
python llm_review.py
# Or: python extract_mirna_genes.py abstracts.txt --auto-review

# Normalize miRNA names, deduplicate, resolve gene IDs
python normalize_names.py
```

## Scripts

### 1. `extract_mirna_genes.py` — Regex Extraction

Scans each abstract for miRNAs and genes using tiered regex patterns, pairs co-occurring entities, and outputs an Excel file with 10 columns including detected species.

```powershell
python extract_mirna_genes.py <abstracts.txt> [--auto-review]
```

### 2. `llm_review.py` — LLM Verification

Sends uncertain pairs to an LLM for four-way classification: Confirmed, Weak association, Excluded (non-target), Excluded (non-gene). The output Excel includes a **Manual Keep** column — Confirmed pairs are pre-filled with `Yes`; Weak association rows are highlighted in yellow for manual review.

```powershell
python llm_review.py
```

### 3. `normalize_names.py` — Normalization, Deduplication & Gene ID Resolution

Standardizes miRNA names, merges duplicate pairs across studies, and resolves gene symbols to database IDs:

- **miRNA normalization**: `OsmiR156` → `osa-miR156`, strips precursors, species-aware
- **Filtering**: retains Confirmed and manually-kept entries; drops Excluded and unmarked Weak pairs
- **Deduplication**: `(species, normalized miRNA, gene)` key — prevents cross-species collisions
- **Gene ID resolution**: queries [ricedata.cn](https://www.ricedata.cn/gene/) for rice genes (RAP, MSU, NCBI Gene ID), NCBI E-utils for all other species

Output: `Normalized-Deduped` sheet with columns `RAP_ID`, `MSU_ID`, `NCBI_GeneID`; `Statistics` sheet aggregating results from all three pipeline stages.

```powershell
python normalize_names.py [input.xlsx]
```

### 4. `auto_update_lists.py` — Rule Feedback Loop

Collects non-gene and confirmed-gene candidates from LLM review, uses batch LLM judgment to update the extraction blacklist/whitelist in `extract_mirna_genes.py`.

```powershell
python auto_update_lists.py       # preview
python auto_update_lists.py --apply
```

## Web Interface

`app.py` provides a Streamlit-based web UI: upload abstracts → extract pairs → LLM review → normalize & resolve gene IDs → download results. Supports multiple LLM providers via the sidebar.

```powershell
streamlit run app.py
```

## Supported Species

| Category | Species |
|---|---|
| **Cereals** | rice, wheat, maize, barley, sorghum |
| **Model plants** | Arabidopsis, *A. lyrata* |
| **Legumes** | soybean, common bean, alfalfa, chickpea, lotus |
| **Solanaceae** | tomato, potato, tobacco, pepper |
| **Fruits** | peach, apple, grape, orange, melon/cucumber, strawberry, banana, papaya |
| **Fiber / oil** | cotton, rapeseed/canola, turnip |
| **Trees** | poplar, pine, cassava |

## Species Detection

Title-first matching with weighted voting fallback:

| Signal | Weight | Example |
|---|---|---|
| Title match | Immediate | "Small RNA profiling in peach fruit..." → `peach` |
| Database ID | 10× | `LOC_Os01g05600` → `rice` |
| miRNA standard prefix | 8× | `tae-miR156` → `wheat` |
| Gene species prefix | 5× | `ZmGRF8` → `maize` |
| Text keyword frequency | 1× per occurrence | Highest-count species wins |

## Gene ID Resolution

Gene symbols are mapped to canonical database IDs to enable downstream analysis (GO enrichment, KEGG pathways, protein networks):

| Species | Source | IDs Retrieved |
|---|---|---|
| Rice | [ricedata.cn](https://www.ricedata.cn/gene/) | RAP-DB, MSU, NCBI Gene |
| All others | NCBI E-utils | NCBI Gene (Entrez) |

The resolution runs automatically during normalization. Unresolved genes are left blank — no IDs are fabricated.

## File Naming

| Stage | Output |
|---|---|
| Extraction | `miRNA-Target_Gene_Pairs.xlsx` |
| LLM Review | `miRNA-Target_Gene_Pairs_Reviewed.xlsx` |
| Normalization | `miRNA-Target_Gene_Pairs_Reviewed_Final.xlsx` |

## License

This project is intended for academic research use.
