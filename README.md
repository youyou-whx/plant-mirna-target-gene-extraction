# miRNA–Target Gene Relationship Extraction Pipeline

A modular Python pipeline for extracting miRNA–target gene pairs from plant literature abstracts, with LLM-assisted verification, miRNA name normalization, and multi-species support.

## Overview

This pipeline processes a batch of PubMed abstracts and produces a curated, deduplicated list of miRNA–target gene relationships. It combines **regex-based extraction** (fast, high recall) with **LLM-based review** (high precision) and **miRNA normalization** (cross-study deduplication), then **auto-updates** its own extraction rules based on review feedback.

```
┌─────────────────────┐     ┌──────────────────────┐     ┌──────────────────────┐
│  extract_mirna_     │───▶│     llm_review.py    │───▶│  normalize_names.py  │
│  genes.py           │     │  (DeepSeek API)      │     │ (dedup + standardize)│
│  regex extraction   │     │  verify/filter pairs │     │                      │
└─────────┬───────────┘     └──────────┬───────────┘     └──────────┬───────────┘
          │                            │                            │
          ▼                            │                            ▼
miRNA-Target_Gene_Pairs.xlsx           ▼    miRNA-Target_Gene_Pairs_Reviewed_Final.xlsx
                       miRNA-Target_Gene_Pairs_Reviewed.xlsx            
                                       │
                                       ▼
                            ┌────────────────────────┐
                            │  auto_update_lists.py  │
                            │  feedback → blacklist  │
                            │  / whitelist rules     │
                            └────────────────────────┘
```

## Features

- **Case-insensitive miRNA detection**: catches `miR156`, `MIR156`, `Mir156`, `mir156`, and legacy forms like `TaMIR5062-5A`
- **Four-tier gene extraction**: database IDs (MSU/RAP) → species-prefixed genes → bare uppercase+digits → gene family names
- **Cross-match detection**: identifies crowded sentences with multiple miRNA/gene candidates and flags ambiguous pairings
- **Sentence-level proximity**: pairs miRNA and gene based on sentence co-occurrence, not arbitrary character windows
- **30 plant species** supported out of the box — see [Supported Species](#supported-species)
- **Automatic species detection**: title-first matching with weighted voting across miRNA prefixes, gene prefixes, and text keywords; bare miRNA names are left unmodified when species is unknown
- **LLM batch review**: sends 5 entries per API call for ~3× speed improvement; supports checkpoint resume on interruption
- **miRNA normalization**: `OsmiR156` → `osa-miR156`, strips `pre-`/`pri-` precursors, species-aware prefix assignment; gene names are preserved as-is
- **Self-improving**: `auto_update_lists.py` feeds LLM review results back into the extraction blacklist/whitelist
- **Fully relative paths**: all scripts locate sibling files automatically — no hardcoded directories

## Requirements

- Python 3.8+
- An API key from a supported LLM provider (DeepSeek, OpenAI, Groq, Together AI, OpenRouter, SiliconFlow, Moonshot, Zhipu, or any OpenAI-compatible endpoint)
- Dependencies installed automatically on first run:
  - `openpyxl` — Excel read/write
  - `requests` — HTTP client for LLM API

## Quick Start

```powershell
# 1. Set your API key (PowerShell)
$env:LLM_API_KEY = "sk-xxxxxxxx"

# 2. Extract miRNA–gene pairs from abstracts
python extract_mirna_genes.py abstracts.txt

# 3. Review uncertain pairs with LLM
python llm_review.py

# 4. Normalize miRNA names and deduplicate
python normalize_names.py

# Or: extract + auto-review in one command
python extract_mirna_genes.py abstracts.txt --auto-review
```

## Scripts

### 1. `extract_mirna_genes.py` — Regex Extraction

Scans each abstract for miRNAs and genes using tiered regex patterns, pairs co-occurring entities, and outputs an Excel file.

| Tier | Pattern | Example | Confidence |
|---|---|---|---|
| miRNA (std) | `xxx-miR###` | `osa-miR156a-5p`, `tae-MIR398` | — |
| miRNA (legacy) | `OsmiR###` | `OsmiR156`, `TaMIR5062-5A` | — |
| miRNA (bare) | `miR###` | `miR156`, `MIR398` | — |
| Gene Tier 1 | Database IDs | `LOC_Os01g05600`, `Os01g0100100` | Confirmed |
| Gene Tier 2 | Species prefix + digits | `OsSPL14`, `TaNAC1` | High confidence |
| Gene Tier 3 | Bare uppercase + digits | `WRKY45`, `NAC1` | Unverified |
| Gene Tier 4 | Gene family names | `ARF`, `NAC`, `MYB` | Unverified |

**Output columns**: `#`, `miRNA`, `Target Gene`, `Relation Confidence`, `Relation Type`, `Source Sentence`, `Article Title`, `PMID`, `DOI`, `Species`

```powershell
python extract_mirna_genes.py <abstracts.txt> [output.xlsx] [--auto-review]
```

### 2. `llm_review.py` — LLM Verification

Sends uncertain pairs (confidence = "Unverified" or relation = "associated") to an LLM for judgment. Supports DeepSeek, OpenAI, Groq, Together AI, OpenRouter, SiliconFlow, Moonshot, Zhipu, and any OpenAI-compatible endpoint. Each entry receives a four-way classification:

| Result | Meaning |
|---|---|
| Confirmed | Real gene + validated targeting relationship |
| Weak association | Possibly related but evidence is indirect |
| Excluded (non-target) | Real gene but no targeting relationship |
| Excluded (non-gene) | Not a gene name at all |

The LLM prompt includes the detected species for context-aware judgment. Batch processing (5 entries per API call) and checkpoint resume on interruption are supported.

The output Excel includes a **Manual Keep** column (col 13). Confirmed entries are pre-filled with `Yes`; Weak association / Excluded rows are left empty with a yellow highlight. To retain a low-confidence pair, fill in `Yes` (or `y` / `keep` / `1`) before running normalization.

```powershell
python llm_review.py
```

### 3. `normalize_names.py` — miRNA Normalization & Deduplication

Standardizes miRNA names across studies and merges duplicate pairs:

- **miRNA**: `OsmiR156` → `osa-miR156`, strips `pre-`/`pri-` precursors, adds species prefix from detected species
- **Gene**: preserved in original extracted form (no normalization applied)
- **Unknown species**: bare miRNA names (`miR156`) are kept as-is rather than guessed
- **Filtering**: only Confirmed and manually-kept entries are retained; Weak association / Excluded are dropped unless the user marks `Yes` in the Manual Keep column
- **Dedup key**: `(species, normalized miRNA, gene)` — prevents cross-species collisions

Output three sheets: `Normalized-Deduped`, `Normalized-Detail`, `Statistics`

```powershell
python normalize_names.py [input.xlsx]
```

### 4. `auto_update_lists.py` — Rule Feedback Loop

Collects non-gene candidates (LLM-excluded names) and confirmed-gene candidates (originally "Unverified" but LLM-confirmed), uses batch LLM judgment to decide which to add to the extraction blacklist/whitelist in `extract_mirna_genes.py`.

**Safety rules enforced**:
- Known plant gene family names (NAC, WRKY, MYB, TCP, SPL, AGO, ARF, etc.) are **never** blacklisted
- Whitelist entries are **never** added to blacklist

```powershell
# Preview (dry run)
python auto_update_lists.py

# Apply changes
python auto_update_lists.py --apply
```

## Supported Species

The pipeline detects **30 plant species** across major crop and model plant categories:

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

The pipeline detects the study organism for each abstract using title-first matching with weighted voting fallback:

| Signal | Weight | Example |
|---|---|---|
| **Title** | Exact match → immediate | "Small RNA profiling in peach fruit..." → `peach` |
| Database ID | 10× | `LOC_Os01g05600` → `rice` |
| miRNA standard prefix | 8× | `tae-miR156` → `wheat` |
| Gene species prefix | 5× | `ZmGRF8` → `maize` |
| Text keyword frequency | 1× per occurrence | "Arabidopsis" (2) vs "peach" (5) → `peach` wins |

Title matching short-circuits the process: if exactly one species is named in the title, it is accepted immediately. When zero or multiple species appear in the title, weighted voting across the full abstract determines the result. If no signal is found, the species column shows `unknown` and bare miRNA names are left unmodified during normalization.

## Input Format

The input file should contain PubMed abstracts with the following structure:

```
1. Journal Name Year Month Day.
   Article Title
   Author names (affiliations)
   Abstract body text.
   DOI: 10.xxxx/xxxxx
   PMID: 12345678
```

Each record begins with a line matching `N. JournalName`. Records are split automatically.

## File Naming Convention

| Stage | Output File |
|---|---|
| Extraction | `miRNA-Target_Gene_Pairs.xlsx` |
| LLM Review | `miRNA-Target_Gene_Pairs_Reviewed.xlsx` |
| Normalization | `miRNA-Target_Gene_Pairs_Final.xlsx` |

All files are written to the script directory (same folder as the `.py` files).

## Configuration

Each script has a configuration block at the top. Key settings:

| Setting | Default | Description |
|---|---|---|
| `LLM_API_URL` | Provider-specific | Set automatically in the web app; override in CLI scripts |
| `LLM_MODEL` | Provider-specific | Set in the web app sidebar or CLI config |
| `ENTRIES_PER_BATCH` | 5 | Entries per API call (speed vs. reliability) |
| `BATCH_SIZE` | 60 | Names per blacklist/whitelist judgment batch |
| `LLM_MAX_RETRIES` | 3 | Retries on API failure |
| `BATCH_SAVE_EVERY` | 100 | Save progress checkpoint every N entries |

## License

This project is intended for academic research use.
