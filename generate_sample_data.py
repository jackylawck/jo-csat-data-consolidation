# generate_sample_data.py
import os
import numpy as np
import pandas as pd


def generate_mock_raw():
    os.makedirs("sample_data", exist_ok=True)
    np.random.seed(42)
    n = 12

    companies = [
        "Alpha Corp",
        "Alpha Ltd",
        "Beta Group",
        "Beta HK",
        "Gamma Co",
        "Delta Inc",
    ] * 2

    data = {
        "Id": list(range(1, n + 1)),
        "開始時間": ["2025-01-01 10:00:00"] * n,
        "完成時間": ["2025-01-01 10:15:00"] * n,
        "電子郵件": ["anonymous"] * n,
        "名稱": [np.nan] * n,
        "姓名 (Name) : ": [f"User_{chr(65+i)}" for i in range(n)],
        "公司名稱 (Company Name) : ": companies,
        "職位 (Position) :": ["Project Manager", "Engineer", "QS", "Safety"]
        * 3,
        "聯絡電話 /電郵 (Telephone number /email) : ": [
            f"1234-567{i} / demo_{i}@example.com" for i in range(n)
        ],
    }

    questions = [
        "施工前的準備 ",
        "圖紙質量-質素及準時性 ",
        "產品質量及監管 ",
        "手工質量及監管 ",
        "地盤管理對工程之熟悉度 ",
        "準時及準確地齊備相關文件,如證書、施工方案、報告等 ",
        "工程能妥善及按時交付 ",
        "就工程質量方面,請對本公司作為 貴司之承判商於市場上的競爭力提出意見 ",
        "若有單項評分低於6分,或對上述範疇有其他意見,敬希告知 ",
        "安環人員對安全及環保法例法規/地盤內部守則之熟悉度 ",
        "安環人員對安全及環保的培訓/推廣 ",
        "安環人員對安全及環保監管 ",
        "工人行為符合安全及環保法例法規/地盤內部守則 ",
        "工人對安全及環保之意識 ",
        "就安全及環保方面,請對本公司作為 貴司之承判商於市場上的競爭力 提出意見 ",
        "若有單項評分低於6分,或對上述範疇有其他意見,敬希告知 .1",
        "員工之工作態度 ",
        "創新技術及設備支援,如BIM ",
        "工程協調性及配合度 ",
        "企業傳承 ",
        "能否就過往狀況、培訓,汲取經驗,達至持續改善 ",
        "就服務質素方面,請對本公司作為 貴司之承判商於市場上的競爭力提出意見 ",
        "若有單項評分低於6分,或對上述範疇有其他意見,敬希告知 .2",
        "我們致力積極培養人才及參與各類社會公益事業/活動,從而弘揚東淦 樂於回饋社會的精神文化,同時也與我們的合作夥伴共同策劃活動及宣 傳策略 , 期盼能收到閣下對此的意見",
        "我們明白誠信乃商業往來之基礎,故期望聆聽閣下對本公司誠信表現的 評價",
        "若閣下有其他寶貴意見及建議,懇祈賜覆",
    ]

    for q in questions:
        if any(kw in q for kw in ["意見", "評價", "文化"]):
            data[q] = ["(Sample feedback comment for open source demo)"] * n
        else:
            data[q] = np.random.randint(6, 11, size=n)

    df = pd.DataFrame(data)
    df.to_excel("sample_data/RAW_SAMPLE.xlsx", index=False)
    print("✅ 已生成開源測試檔案：sample_data/RAW_SAMPLE.xlsx")


if __name__ == "__main__":
    generate_mock_raw()
