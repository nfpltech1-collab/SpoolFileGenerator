import pandas as pd
from datetime import datetime


def load_excel_data(excel_path, is_spare, selected_date=None):
    if not excel_path:
        return None

    if excel_path.lower().endswith('.csv'):
        if is_spare:
            df = pd.read_csv(excel_path)
        else:
            df = pd.read_csv(excel_path, skiprows=2)
    else:
        xls = pd.ExcelFile(excel_path)
        sheet_names = xls.sheet_names

        search_date = selected_date if selected_date else datetime.now()

        def normalize(name):
            return str(name).strip().upper()

        date_tokens = {
            search_date.strftime('%d-%m-%Y'),
            search_date.strftime('%d/%m/%Y'),
            search_date.strftime('%d-%b-%Y'),
            search_date.strftime('%d-%b-%y'),
            search_date.strftime('%d-%m-%y'),
        }

        candidates = []
        for name in sheet_names:
            upper_name = normalize(name)
            is_match = any(token.upper() in upper_name for token in date_tokens)

            if is_spare:
                if is_match:
                    candidates.append(name)
            else:
                if is_match and 'RPDC' not in upper_name:
                    candidates.append(name)

        dfs = []
        for sheet_name in candidates:
            if is_spare:
                sheet_df = pd.read_excel(excel_path, sheet_name=sheet_name)
            else:
                sheet_df = pd.read_excel(excel_path, sheet_name=sheet_name, skiprows=2)
            dfs.append(sheet_df)

        if not dfs:
            if is_spare:
                for name in sheet_names:
                    if 'RPDC' in normalize(name):
                        dfs.append(pd.read_excel(excel_path, sheet_name=name))
                        break
            if not dfs and sheet_names:
                if is_spare:
                    dfs.append(pd.read_excel(excel_path, sheet_name=sheet_names[0]))
                else:
                    dfs.append(pd.read_excel(excel_path, sheet_name=sheet_names[0], skiprows=2))

        if dfs:
            df = pd.concat(dfs, ignore_index=True)
        else:
            raise ValueError("No sheets found to load")

    df.columns = df.columns.str.strip().str.upper()
    return df
