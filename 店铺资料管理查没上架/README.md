# Product Code Gap Checker

This program finds product codes that exist in `商品资料` but **never appear** in `店铺商品资料`.

It also filters out:
- codes starting with `FH`
- codes containing Chinese characters

## Project structure

- `src/find_unlisted_product_codes.py`: main script
- `data/input/`: input Excel files (`.xlsx`)
- `data/output/`: generated result file

## Required input files

Put these two files in `data/input/`:
- `店铺商品资料*.xlsx` (contains column `原始商品编码`)
- `商品资料*.xlsx` (contains columns `商品编码`, `创建时间`)

By default, the script picks the latest file matching each prefix.

## Setup

From project root:

```bash
cd "/Users/churongguo/Documents/Duoyi/店铺资料管理查没上架"
python3 -m venv .venv
source .venv/bin/activate
pip install openpyxl
```

## Run

### Default run (auto-pick latest input files)

```bash
source .venv/bin/activate
python src/find_unlisted_product_codes.py
```

### Run with explicit files

```bash
source .venv/bin/activate
python src/find_unlisted_product_codes.py \
  --shop-file "data/input/店铺商品资料_xxx.xlsx" \
  --catalog-file "data/input/商品资料_xxx.xlsx"
```

## Output

Default output file:

- `data/output/店铺未上架商品编码.xlsx`

Sheets:
- `完全未出现明细`: columns `商品编码`, `创建时间` (sorted by created time, newest first)
- `统计`: total count

## Optional arguments

```bash
python src/find_unlisted_product_codes.py --help
```

Useful options:
- `--output`: custom output path
- `--shop-sheet`: sheet name for store file
- `--catalog-sheet`: sheet name for catalog file
- `--shop-sku-col`: default `原始商品编码`
- `--catalog-sku-col`: default `商品编码`
- `--time-col`: default `创建时间`
