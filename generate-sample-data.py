import pandas as pd
import random
from pathlib import Path

Path("data").mkdir(exist_ok=True)

branches = ["مشهد","شیراز","اصفهان","تهران"]
products = {900_000 :"کیبورد" , 450_000 : "موس" , 35_000_000 : "لپ تاپ" , 8_000_000 : "مانیتور" ,1_500_000: "هدفون"}

for branch in branches:
    rows = []
    for _ in range(50):
        product = random.choice(list(products))
        qty = random.randint(1, 5)
        rows.append({
            "تاریخ": f"1405/0{random.randint(1, 6)}/{random.randint(10, 28)}",
            "شعبه": branch,
            "محصول": products[product],
            "تعداد": qty,
            "قیمت واحد": product ,
        })
    pd.DataFrame(rows).to_excel(f"data/sales_{branch}.xlsx", index=False)

print("فایل نمونه ساخته شد.")