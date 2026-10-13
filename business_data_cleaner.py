
from pathlib import Path
import pandas as pd


# ============================================================
# 1. CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
INPUT_FILE = BASE_DIR / "data" / "messy_sales.xlsx"
OUTPUT_DIR = BASE_DIR / "outputs"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

CLEANED_FILE = OUTPUT_DIR / "cleaned_sales_report.xlsx"


# ============================================================
# 2. LOAD AND VALIDATE DATA
# ============================================================

def load_data(file_path):
    if not file_path.exists():
        raise FileNotFoundError(
            f"Input file not found: {file_path}\n"
            "Place messy_sales.xlsx inside the data folder."
        )

    df = pd.read_excel(file_path)

    required_columns = [
        "Customer",
        "Product",
        "Date",
        "Quantity",
        "Price",
    ]

    missing_columns = [
        column
        for column in required_columns
        if column not in df.columns
    ]

    if missing_columns:
        raise ValueError(
            f"Missing required columns: {missing_columns}"
        )

    if df.empty:
        raise ValueError("The input Excel file contains no rows.")

    return df


# ============================================================
# 3. CLEAN AND STANDARDIZE DATA
# ============================================================

def clean_data(df):
    df = df.copy()

    original_rows = len(df)

    # Preserve information about missing values.
    df["Missing_Customer"] = df["Customer"].isna()
    df["Missing_Product"] = df["Product"].isna()
    df["Missing_Date"] = df["Date"].isna()
    df["Missing_Quantity"] = df["Quantity"].isna()
    df["Missing_Price"] = df["Price"].isna()

    # Standardize text values.
    for column in ["Customer", "Product"]:
        df[column] = (
            df[column]
            .astype("string")
            .str.strip()
            .str.lower()
        )

        # Treat empty strings as missing values.
        df[column] = df[column].replace("", pd.NA)

    # Parse dates and flag values that cannot be interpreted.
    df["Date"] = pd.to_datetime(
        df["Date"],
        errors="coerce",
        dayfirst=True,
    )

    # Identify dates that may need manual review.
    df["Suspicious_Date"] = (
        (df["Date"].dt.month == 12)
        & (df["Date"].dt.year == 2024)
    ).fillna(False)

    # Convert numeric columns safely.
    for column in ["Quantity", "Price"]:
        df[column] = pd.to_numeric(
            df[column],
            errors="coerce",
        )

    # Detect invalid quantities before any imputation.
    df["Invalid_Quantity"] = (
        df["Quantity"].notna()
        & (df["Quantity"] <= 0)
    )

    # Replace missing customer names with a clear placeholder.
    df["Customer"] = df["Customer"].fillna("Unknown")

    # Impute missing prices using the first available
    # price for that product. Track every imputed value.
    df["Price_Imputed"] = (
        df["Price"].isna()
        & df["Product"].notna()
    )

    product_prices = (
        df.groupby("Product", dropna=True)["Price"]
        .transform("first")
    )

    df["Price"] = df["Price"].fillna(product_prices)

    # Remove exact duplicates after text normalization.
    rows_before_deduplication = len(df)
    df = df.drop_duplicates().copy()

    removed_duplicates = (
        rows_before_deduplication - len(df)
    )

    # Missing or invalid quantities are NOT automatically
    # changed to 1. They remain visible for review.
    df["Invalid_Quantity"] = (
        df["Quantity"].isna()
        | (df["Quantity"] <= 0)
    )

    # A sale is valid for revenue reporting only when
    # quantity is positive and price is non-negative.
    df["Valid_Sale"] = (
        df["Quantity"].notna()
        & (df["Quantity"] > 0)
        & df["Price"].notna()
        & (df["Price"] >= 0)
    )

    # Calculate sales values.
    df["Total_Sales"] = df["Quantity"] * df["Price"]

    df["Valid_Revenue"] = df["Total_Sales"].where(
        df["Valid_Sale"],
        0,
    )

    # Record summary information.
    cleaning_stats = {
        "Original Rows": original_rows,
        "Rows After Cleaning": len(df),
        "Duplicates Removed": removed_duplicates,
        "Missing Dates": int(df["Date"].isna().sum()),
        "Invalid Quantities": int(
            df["Invalid_Quantity"].sum()
        ),
        "Missing Prices After Imputation": int(
            df["Price"].isna().sum()
        ),
        "Prices Imputed": int(df["Price_Imputed"].sum()),
        "Suspicious Dates": int(
            df["Suspicious_Date"].sum()
        ),
    }

    return df, cleaning_stats


# ============================================================
# 4. GENERATE SALES SUMMARIES
# ============================================================

def generate_reports(df, cleaning_stats):
    total_revenue = df["Valid_Revenue"].sum()

    product_summary = (
        df.groupby("Product", dropna=False)
        .agg(
            Total_Quantity=("Quantity", "sum"),
            Total_Revenue=("Valid_Revenue", "sum"),
            Number_of_Records=("Product", "size"),
        )
        .sort_values("Total_Revenue", ascending=False)
    )

    customer_summary = (
        df.groupby("Customer", dropna=False)
        .agg(
            Total_Revenue=("Valid_Revenue", "sum"),
            Number_of_Records=("Customer", "size"),
        )
        .sort_values("Total_Revenue", ascending=False)
    )

    # Exclude missing product/customer categories from
    # the best-product and best-customer selections.
    valid_products = product_summary.dropna(
        subset=["Total_Revenue"]
    )

    valid_customers = customer_summary.dropna(
        subset=["Total_Revenue"]
    )

    best_product = (
        valid_products.index[0]
        if not valid_products.empty
        else "N/A"
    )

    best_customer = (
        valid_customers.index[0]
        if not valid_customers.empty
        else "N/A"
    )

    # Record remaining missing values.
    missing_values = (
        df.isna()
        .sum()
        .rename("Missing_Count")
        .to_frame()
    )

    # Keep records needing attention in a separate sheet.
    review_mask = (
        df["Missing_Date"]
        | df["Missing_Product"]
        | df["Missing_Quantity"]
        | df["Invalid_Quantity"]
        | df["Price"].isna()
        | df["Suspicious_Date"]
    )

    review_records = df.loc[review_mask].copy()

    summary = pd.DataFrame(
        {
            "Metric": [
                "Total Valid Revenue",
                "Best Product",
                "Best Product Revenue",
                "Best Customer",
                "Best Customer Revenue",
                "Invalid Quantity Records",
                "Unknown Customer Revenue",
            ],
            "Value": [
                total_revenue,
                best_product,
                (
                    product_summary.loc[
                        best_product, "Total_Revenue"
                    ]
                    if best_product != "N/A"
                    else 0
                ),
                best_customer,
                (
                    customer_summary.loc[
                        best_customer, "Total_Revenue"
                    ]
                    if best_customer != "N/A"
                    else 0
                ),
                int(df["Invalid_Quantity"].sum()),
                df.loc[
                    df["Customer"] == "Unknown",
                    "Valid_Revenue",
                ].sum(),
            ],
        }
    )

    cleaning_summary = pd.DataFrame(
        list(cleaning_stats.items()),
        columns=["Metric", "Value"],
    )

    return {
        "Sales Summary": summary,
        "Cleaned Sales": df,
        "Product Summary": product_summary,
        "Customer Summary": customer_summary,
        "Cleaning Stats": cleaning_summary,
        "Missing Values": missing_values,
        "Review Required": review_records,
    }


# ============================================================
# 5. EXPORT EXCEL REPORT
# ============================================================

def export_reports(reports, output_file):
    with pd.ExcelWriter(
        output_file,
        engine="openpyxl",
    ) as writer:
        for sheet_name, report_df in reports.items():
            report_df.to_excel(
                writer,
                sheet_name=sheet_name,
                index=True if sheet_name in [
                    "Product Summary",
                    "Customer Summary",
                ] else False,
            )

    print(f"\nExcel report saved successfully:\n{output_file}")


# ============================================================
# 6. MAIN PROGRAM
# ============================================================

def main():
    try:
        print("Loading sales data...")

        original_df = load_data(INPUT_FILE)

        print(f"Original rows: {len(original_df)}")

        cleaned_df, cleaning_stats = clean_data(
            original_df
        )

        print("\nGenerating reports...")

        reports = generate_reports(
            cleaned_df,
            cleaning_stats,
        )

        export_reports(
            reports,
            CLEANED_FILE,
        )

        print("\n========== SALES REPORT ==========")

        print(
            reports["Sales Summary"].to_string(
                index=False
            )
        )

        print("\n========== CLEANING STATISTICS ==========")

        print(
            reports["Cleaning Stats"].to_string(
                index=False
            )
        )

        print("\n========================================")

    except (FileNotFoundError, ValueError) as error:
        print(f"\nInput or validation error: {error}")

    except PermissionError:
        print(
            "\nPermission error: close the output Excel file "
            "if it is already open, then try again."
        )

    except Exception as error:
        print(
            f"\nUnexpected error: {error}\n"
            "Check the input data and traceback if needed."
        )
        raise


if __name__ == "__main__":
    main()
