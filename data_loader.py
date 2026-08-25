import pandas as pd
import numpy as np
import os


# ============================================================
# 1. LOAD EXCEL FILE
# ============================================================

file_path = r"FRP_master_integrated_dataset.xlsx"

output_file = r"FRP_preprocessed_dataset.xlsx"


data = pd.read_excel(
    file_path,
    sheet_name=None
)


print("Excel file loaded successfully.")

print("Available sheets:")

print(list(data.keys()))


# ============================================================
# 2. CREATE STORAGE FOR CLEANED DATA
# ============================================================

cleaned_data = {}


# ============================================================
# 3. PROCESS EACH SHEET
# ============================================================

for sheet_name, df in data.items():

    print("\nProcessing sheet:")

    print(sheet_name)


    # ========================================================
    # 3.1 REMOVE COMPLETELY EMPTY ROWS
    # ========================================================

    df = df.dropna(
        axis=0,
        how="all"
    )


    # ========================================================
    # 3.2 REMOVE COMPLETELY EMPTY COLUMNS
    # ========================================================

    df = df.dropna(
        axis=1,
        how="all"
    )


    # ========================================================
    # 3.3 REMOVE DUPLICATE RECORDS
    # ========================================================

    original_records = len(df)


    df = df.drop_duplicates(
        keep="first"
    )


    duplicate_records = (
        original_records -
        len(df)
    )


    # ========================================================
    # 3.4 IDENTIFY NUMERIC COLUMNS
    # ========================================================

    numeric_columns = df.select_dtypes(
        include=np.number
    ).columns.tolist()


    # ========================================================
    # 3.5 IDENTIFY CATEGORICAL COLUMNS
    # ========================================================

    categorical_columns = df.select_dtypes(
        exclude=np.number
    ).columns.tolist()


    # ========================================================
    # 3.6 MISSING VALUE IMPUTATION
    # ========================================================

    # Numeric columns → Median

    for column in numeric_columns:

        if df[column].isna().sum() > 0:

            median_value = df[column].median()

            df[column] = df[column].fillna(
                median_value
            )


    # Categorical columns → Mode

    for column in categorical_columns:

        if df[column].isna().sum() > 0:

            mode_value = df[column].mode()

            if len(mode_value) > 0:

                df[column] = df[column].fillna(
                    mode_value.iloc[0]
                )


    # ========================================================
    # 3.7 IQR OUTLIER DETECTION
    # ========================================================

    outlier_count = 0


    for column in numeric_columns:

        Q1 = df[column].quantile(
            0.25
        )

        Q3 = df[column].quantile(
            0.75
        )

        IQR = Q3 - Q1


        lower_limit = Q1 - (
            1.5 * IQR
        )

        upper_limit = Q3 + (
            1.5 * IQR
        )


        outliers = (
            (df[column] < lower_limit) |
            (df[column] > upper_limit)
        )


        outlier_count += outliers.sum()


    # ========================================================
    # 3.8 STORE PREPROCESSED DATA
    # ========================================================

    cleaned_data[sheet_name] = df


    # ========================================================
    # 3.9 DISPLAY INFORMATION
    # ========================================================

    print("Original records:")

    print(original_records)


    print("Duplicate records removed:")

    print(duplicate_records)


    print("Remaining records:")

    print(len(df))


    print("Numeric features:")

    print(len(numeric_columns))


    print("Categorical features:")

    print(len(categorical_columns))


    print("IQR outlier values identified:")

    print(outlier_count)


# ============================================================
# 4. SAVE PREPROCESSED DATA
# ============================================================

with pd.ExcelWriter(
    output_file,
    engine="openpyxl"
) as writer:

    for sheet_name, df in cleaned_data.items():

        df.to_excel(
            writer,
            sheet_name=sheet_name[:31],
            index=False
        )


# ============================================================
# 5. FINAL INFORMATION
# ============================================================

print("\nData preprocessing completed.")

print("Preprocessed Excel file saved:")

print(output_file)


print("Number of sheets:")

print(len(cleaned_data))