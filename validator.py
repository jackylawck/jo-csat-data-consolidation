# validator.py
import logging
import pandas as pd


class SurveyDataValidationError(Exception):
    pass


def validate_raw_schema(df: pd.DataFrame) -> bool:
    """防禦性校驗：確認上傳檔案是否具備問卷必要題目關鍵字"""
    if df.empty:
        raise SurveyDataValidationError(
            "上傳檔案為空，請檢查 Microsoft Forms 匯出檔！"
        )

    cols_str = " ".join([str(c) for c in df.columns])
    required_keywords = [
        "姓名",
        "公司名稱",
        "施工前的準備",
        "圖紙質量",
        "安環人員",
        "員工之工作態度",
    ]

    missing_keys = [kw for kw in required_keywords if kw not in cols_str]
    if missing_keys:
        error_msg = f"RAW 檔案格式不符！缺少必要欄位關鍵字：{', '.join(missing_keys)}"
        logging.error(error_msg)
        raise SurveyDataValidationError(error_msg)

    return True
