import pandas as pd


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = (
        df.columns.astype(str)  # 防止列名是 None / 数字
        .str.replace(r"\s+", "", regex=True)  # ✅ 去掉所有空格、换行、Tab
        .str.lower()  # ✅ 统一小写
    )
    return df
