import pandas as pd
from io import BytesIO


def process_csv_files(files):

    def read_csv_smart(file):

        encodings = ["utf-8", "latin1", "iso-8859-1"]

        for enc in encodings:
            try:
                df = pd.read_csv(file, encoding=enc, sep=None, engine="python")
                return df
            except:
                file.seek(0)

        return None


    def validate_df(df):

        issues = []

        if df.empty:
            issues.append("Arquivo vazio")

        if df.isnull().sum().sum() > 0:
            issues.append("Possui valores nulos")

        return issues


    all_dataframes = {}
    report = []

    for file in files:

        df = read_csv_smart(file)

        if df is None:
            continue

        issues = validate_df(df)

        report.append({
            "arquivo": file.name,
            "linhas": len(df),
            "problemas": ", ".join(issues) if issues else "OK"
        })

        sheet_name = file.name[:30]
        all_dataframes[sheet_name] = df


    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        for sheet, df in all_dataframes.items():
            df.to_excel(writer, sheet_name=sheet, index=False)

        pd.DataFrame(report).to_excel(writer, sheet_name="Relatorio", index=False)

    return output.getvalue()