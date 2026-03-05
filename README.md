# Total Asia – Supplier ↔ ERP Matching & Stock Check

## Purpose
Match supplier availability data (**Supplier**) into the internal ERP item master (**ERP**) and generate a **Stock Check** file for import/checking.

The output **Stock Check**:
- Keeps **the exact row order of ERP**
- Uses only the required columns:
  `Index, Product, Suppl category, Description, Size, Pack, Store, Price, Qty`
- Fills `Price` and `Qty` from Supplier where a match is confirmed
- Optionally fills missing `Qty` as `0` (supplier out of stock)

## File Naming
- **Supplier.xlsx** (formerly A)
- **ERP.xlsx** (formerly B)
- **Stock Check.xlsx** (formerly C) – generated output

## Required Columns

### Supplier.xlsx
| Column | Meaning |
|---|---|
| CODE | Supplier item code (may change over time) |
| DESCRIPTION | Supplier product name (English used for matching; Chinese ignored) |
| PACK SIZE | Package/size string (parsed for size matching) |
| QTY | Supplier available quantity |
| PRICE | Supplier price |

### ERP.xlsx
| Column | Meaning |
|---|---|
| Product | ERP product ID (used as stable ID in output) |
| Suppl category | Supplier category/code field in ERP |
| Description | ERP description (English used for matching; Chinese ignored) |
| Size | ERP size |
| Pack | ERP pack |
| Store | ERP store |

## Product ID Handling (Leading Zeros)
Excel often stores Product as a number (e.g., `401`) while displaying `000401` via formatting.  
This project **standardizes Product as a string and pads it to a fixed width** (default `6` digits).

Example:
- `401` -> `000401`

You can configure the padding width via `MatchConfig(product_pad_width=6)`.

## Matching Logic
Matching is performed per Supplier row against ERP items.

### Text rules for matching descriptions
- Case-insensitive
- Chinese is ignored
- Only English alphanumeric tokens are used

### Size parsing rules (Supplier PACK SIZE)
- If contains `x` (or `×`): take the segment after the last `x`, then extract number+unit  
  Examples:
  - `2x5L` -> `5L`
  - `12x540g bag` -> `540g`
- Else: take the first number+unit  
  Example:
  - `20kg bag` -> `20kg`

Units supported: `g, kg, ml, L, pcs` (case-insensitive)

### Confirmation Rules (in order)
Let:
1) `Supplier.CODE` numeric equals `ERP.Suppl category` numeric  
2) Description common English tokens meet threshold (primary: >= 3; relaxed: >= 2)  
3) Parsed Supplier size equals ERP.Size (normalized)

Rules:
1. **Rule 1 (1+2)**: If (1) AND (2 with >=3 tokens), confirm match.
2. **Rule 2 (2+3)**: Else if (2 with >=3 tokens) AND (3), confirm match.
3. **Rule 3 (fallback, relaxed description)**: Else, reduce description requirement to **>=2 tokens**, but **must satisfy (1+2+3)** simultaneously.

Unmatched Supplier rows are exported to `Unmatched.xlsx`.

## Environment Setup (Conda)
```bash
conda create -n TA python=3.11 -y
conda activate TA

conda install -y pandas numpy openpyxl jupyter ipykernel
pip install rapidfuzz

# Only needed if you still have .xls files:
conda install -y xlrd

python -m ipykernel install --user --name TA --display-name "Python (TA)"
conda env export -n TA --no-builds > TA_env.yml# Supplier_StockCheck
# Supplier-CompanyStockMatcher
# Supplier-CompanyStockMatcher
