from pathlib import Path
import sys

import pandas as pd

# --------------------
DATA_FOLDER = "data"
OUTPUT_FILE = "sales_report.xlsx"
REQUIRED_COLUMNS = ["تاریخ", "شعبه", "محصول", "تعداد", "قیمت واحد"]
# -------------------------------------------------------


def to_number(series: pd.Series) -> pd.Series:
    """تبدیل متن به عدد؛ ارقام فارسی و جداکننده هزارگان را هم می‌فهمد."""
    persian = "۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩"
    english = "01234567890123456789"
    table = str.maketrans(persian, english)
    cleaned = (series.astype(str)
                     .str.translate(table)
                     .str.replace(",", "", regex=False)
                     .str.replace("٬", "", regex=False)
                     .str.strip())
    return pd.to_numeric(cleaned, errors="coerce")


def load_files(folder: str) -> pd.DataFrame:
    """همه فایل‌های اکسل پوشه را می‌خواند و ادغام می‌کند."""
    files = sorted(Path(folder).glob("*.xlsx"))
    if not files:
        sys.exit(f"خطا: هیچ فایل اکسلی در پوشه '{folder}' پیدا نشد.")

    frames = []
    for f in files:
        part = pd.read_excel(f)
        missing = [c for c in REQUIRED_COLUMNS if c not in part.columns]
        if missing:
            print(f"هشدار: فایل {f.name} ستون‌های {missing} را ندارد و نادیده گرفته شد.")
            continue
        frames.append(part[REQUIRED_COLUMNS])

    if not frames:
        sys.exit("خطا: هیچ فایل معتبری برای ادغام وجود ندارد.")

    print(f"{len(frames)} فایل خوانده شد.")
    return pd.concat(frames, ignore_index=True)


def clean(df: pd.DataFrame) -> pd.DataFrame:
    """حذف ردیف‌های نامعتبر و تکراری."""
    before = len(df)

    df["تعداد"] = to_number(df["تعداد"])
    df["قیمت واحد"] = to_number(df["قیمت واحد"])

    df = df.dropna(subset=REQUIRED_COLUMNS)
    df = df[(df["تعداد"] > 0) & (df["قیمت واحد"] > 0)]
    df = df.drop_duplicates()

    print(f"{before - len(df)} ردیف نامعتبر یا تکراری حذف شد. {len(df)} ردیف باقی ماند.")
    return df.reset_index(drop=True)


def summarize(df: pd.DataFrame):
    """ساخت جدول‌های خلاصه."""
    df = df.copy()
    df["مبلغ کل"] = df["تعداد"] * df["قیمت واحد"]

    by_branch = (df.groupby("شعبه")["مبلغ کل"]
                   .sum()
                   .sort_values(ascending=False)
                   .reset_index())

    by_product = (df.groupby("محصول")
                    .agg(تعداد_فروش=("تعداد", "sum"),
                         مبلغ_کل=("مبلغ کل", "sum"))
                    .sort_values("مبلغ_کل", ascending=False)
                    .reset_index())

    return df, by_branch, by_product


def save_report(df, by_branch, by_product, output: str) -> None:
    """ذخیره در اکسل با چند شیت و فرمت‌دهی."""
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        sheets = {
            "داده‌های ادغام‌شده": df,
            "فروش شعب": by_branch,
            "فروش محصولات": by_product,
        }
        for name, data in sheets.items():
            data.to_excel(writer, sheet_name=name, index=False)
            ws = writer.sheets[name]

            # عرض ستون‌ها بر اساس طول محتوا
            for col_cells in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0
                            for c in col_cells)
                ws.column_dimensions[col_cells[0].column_letter].width = width + 4

            # جداکننده هزارگان برای ستون‌های عددی
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = "#,##0"

            ws.freeze_panes = "A2"   # ثابت ماندن سطر عنوان

    print(f"گزارش در '{output}' ذخیره شد.")


def main():
    df = load_files(DATA_FOLDER)
    df = clean(df)
    df, by_branch, by_product = summarize(df)

    print("\nفروش هر شعبه:")
    print(by_branch.to_string(index=False))

    save_report(df, by_branch, by_product, OUTPUT_FILE)


if __name__ == "__main__":
    main()