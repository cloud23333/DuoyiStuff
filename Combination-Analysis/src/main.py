from datetime import datetime
from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_DATA_DIR = PROJECT_ROOT / "data" / "input"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "data" / "output"
DEFAULT_INPUT_FILE_PATH = DEFAULT_DATA_DIR / "模拟-i输入-讲解.xlsx"
DEFAULT_PRODUCT_DB_PATH = DEFAULT_DATA_DIR / "商品资料.xlsx"
DEFAULT_COMBO_DB_PATH = DEFAULT_DATA_DIR / "组合资料.xlsx"
PROCESS_COLUMNS = [
    "单品是否存在",
    "组合是否存在",
    "倍数前",
    "倍数后",
]


class ExcelDataService:

    def __init__(self):
        self.product_db_path = DEFAULT_PRODUCT_DB_PATH
        self.combo_db_path = DEFAULT_COMBO_DB_PATH
        self.product_df = None
        self.combo_df = None

    @staticmethod
    def _parse_multiplier_codes(code_series):
        source_codes = code_series.astype(str)
        code_parts = source_codes.str.extract(r"^(.*)\*(\d+)$")
        has_multiplier = code_parts[1].notna()
        return (
            source_codes,
            has_multiplier,
            code_parts[0].fillna(""),
            code_parts[1].fillna(""),
        )

    def _load_database(self, file_path):
        if not file_path.exists():
            return None
        return pd.read_excel(file_path)

    def load_databases(self):
        self.product_df = self._load_database(self.product_db_path)
        self.combo_df = self._load_database(self.combo_db_path)

    def load_input_file(self, file_path):
        return pd.read_excel(file_path)

    def process_data(self, input_df):
        df = input_df.copy().drop_duplicates(subset=["原始商品编码"])
        source_codes, has_multiplier, base_codes, multiplier_codes = (
            self._parse_multiplier_codes(df["原始商品编码"])
        )
        df = df[has_multiplier].copy()

        for column in PROCESS_COLUMNS:
            df[column] = ""

        source_codes = source_codes[has_multiplier]
        product_lookup = self.product_df.copy()
        product_lookup["商品编码"] = product_lookup["商品编码"].astype(str)
        product_lookup = product_lookup.drop_duplicates(
            subset=["商品编码"],
            keep="first",
        ).set_index("商品编码")

        product_codes = product_lookup.index
        combo_codes = self.combo_df["组合商品编码"].astype(str)

        product_exists = source_codes.isin(product_codes)
        combo_exists = source_codes.isin(combo_codes)

        df["单品是否存在"] = product_exists.map({True: "存在", False: ""})
        df["组合是否存在"] = combo_exists.map({True: "存在", False: ""})

        missing_mask = ~product_exists & ~combo_exists
        df = df[missing_mask].copy()
        if df.empty:
            return df

        df["倍数前"] = base_codes.loc[df.index]
        df["倍数后"] = multiplier_codes.loc[df.index]
        return df

    @staticmethod
    def build_export_df(processed_df):
        return pd.DataFrame(
            {
                "组合商品编码": processed_df["原始商品编码"],
                "商品编码": processed_df["倍数前"],
                "数量": processed_df["倍数后"],
            }
        )


def resolve_output_path():
    return DEFAULT_OUTPUT_DIR / f"组合装数据{datetime.now().strftime('%m%d')}.xlsx"


def run():
    input_path = DEFAULT_INPUT_FILE_PATH
    service = ExcelDataService()

    service.load_databases()
    input_df = service.load_input_file(input_path)
    processed_df = service.process_data(input_df)
    export_df = service.build_export_df(processed_df)
    output_path = resolve_output_path()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    export_df.to_excel(output_path, index=False)


def main():
    run()


if __name__ == "__main__":
    main()
