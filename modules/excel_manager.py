from io import BytesIO
import pandas as pd

def create_excel_download(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Facturas")
    return output.getvalue()

def save_to_database(df):
    return True
