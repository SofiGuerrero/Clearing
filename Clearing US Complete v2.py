"""
SAP FAGLL03 - Descarga y pivot de GL Account Line Items ISP US
Requiere: pip install pywin32 openpyxl pandas
"""

import sys
import os
import calendar
import pandas as pd
import openpyxl
import win32com.client
from datetime import date
from clearing_utils import write_pivot_sheet, write_docs_sheet, write_resumen_sheet

# ─── CONFIGURACION ────────────────────────────────────────────────────────────
SCRIPT_DIR = os.path.join(os.path.expanduser("~"), "OneDrive - SAP SE", "Varios SAP - Daily", "JE-Sofi", "Clearing")
INPUT_FILE   = SCRIPT_DIR + r"\Clearing US Input.xlsx"
OUTPUT_FILE  = SCRIPT_DIR + r"\Clearing US Pivot.xlsx"
# ──────────────────────────────────────────────────────────────────────────────


def get_last_day_of_month():
    today = date.today()
    last = calendar.monthrange(today.year, today.month)[1]
    return date(today.year, today.month, last).strftime("%Y%m%d")


def get_isp_connection(sap_app):
    for i in range(sap_app.Children.Count):
        connection = sap_app.Children(i)
        if connection.Children.Count > 0:
            system_id = connection.Children(0).PassportSystemid
            if system_id[:3] == "ISP":
                return connection
    return None


def download_report():
    s_last_day = get_last_day_of_month()

    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
    except Exception:
        print("ERROR: No se pudo conectar al SAP GUI. Asegurate de que SAP GUI esté abierto.")
        sys.exit(1)

    sap_app = sap_gui_auto.GetScriptingEngine

    sap_con = get_isp_connection(sap_app)
    if sap_con is None:
        print("ERROR: Por favor abrí la conexión ISP en SAP GUI.")
        sys.exit(1)

    session = sap_con.Children(0)

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "FAGLL03"
    session.findById("wnd[0]").sendVKey(0)

    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").text = "/GL ACC SA"
    session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    session.findById("wnd[0]/usr/ctxtPA_STIDA").setFocus()
    session.findById("wnd[0]/usr/ctxtPA_STIDA").caretPosition = 7
    session.findById("wnd[0]").sendVKey(4)

    shell = session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell")
    shell.focusDate = s_last_day
    shell.firstVisibleDate = s_last_day
    shell.selectionInterval = f"{s_last_day},{s_last_day}"

    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    session.findById("wnd[0]/usr/lbl[17,14]").setFocus()
    session.findById("wnd[0]/usr/lbl[17,14]").caretPosition = 0
    session.findById("wnd[0]").sendVKey(16)
    session.findById("wnd[1]/tbar[0]/btn[20]").press()

    session.findById("wnd[1]/usr/ctxtDY_PATH").text = SCRIPT_DIR
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Clearing US Input.xlsx"
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    print(f"Reporte FAGLL03 descargado correctamente. Fecha clave: {s_last_day}")


def apply_account_rules(df):
    frames = []

    for acc, group in df.groupby("Account"):

        if acc == 172501:
            mask = (
                group["Text"].str.contains(r"ADP TAX", case=False, na=False) |
                group["Text"].str.contains(r"CIT EMPLOYEE", case=False, na=False)
            )
            filtered = group[mask]
            excluded = len(group) - len(filtered)
            if excluded:
                print(f"  172501: excluidos {excluded} rows (no son ADP TAX ni CIT EMPLOYEE)")

        elif acc == 176000:
            mask = (
                group["Text"].str.contains(r"RECLASS", case=False, na=False) |
                group["Text"].str.contains(r"WAGES & SALARIES TO BE PAID", case=False, na=False)
            )
            filtered = group[mask]
            excluded = len(group) - len(filtered)
            if excluded:
                print(f"  176000: excluidos {excluded} rows (no son RECLASS ni WAGES & SALARIES)")

        elif acc == 179000:
            def is_valid_assignment(val):
                if pd.isna(val) or str(val).strip() == "":
                    return True
                s = str(val).strip()
                if s.upper() == "PAYROLL":
                    return True
                return s.replace("0", "").isdigit() or s.isdigit()

            mask = group["Assignment"].apply(is_valid_assignment)
            filtered = group[mask]
            excluded = len(group) - len(filtered)
            if excluded:
                print(f"  179000: excluidos {excluded} rows (Assignment no es PAYROLL, numérico ni blanco)")

        else:
            filtered = group

        frames.append(filtered)

    return pd.concat(frames, ignore_index=True) if frames else df.iloc[0:0]


def build_pivot():
    print(f"Leyendo archivo: {INPUT_FILE}")
    df = pd.read_excel(INPUT_FILE)
    print(f"  {len(df):,} registros cargados.")

    before = len(df)
    df = df[df["Year/Month"] >= "2024/01"]
    print(f"  Filtro Year/Month >= 2024/01: excluidos {before - len(df):,} rows. Quedan {len(df):,}.")

    df = apply_account_rules(df)
    print(f"  {len(df):,} registros tras aplicar reglas por cuenta.")

    pivot = df.pivot_table(
        index=["Account", "Company Code"],
        values="Amount in Local Currency",
        aggfunc="sum"
    ).reset_index()

    pivot.columns = ["G/L Account", "Company Code", "Balance"]
    pivot["Balance"] = pivot["Balance"].round(2)
    pivot = pivot.sort_values(["G/L Account", "Company Code"]).reset_index(drop=True)

    zero_count = (pivot["Balance"] == 0).sum()
    print(f"  Cuentas con balance 0: {zero_count}")
    print(f"  Total combinaciones cuenta/sociedad: {len(pivot)}")

    wb = openpyxl.Workbook()
    write_pivot_sheet(wb, pivot)
    write_docs_sheet(wb, df, pivot)
    write_resumen_sheet(wb, pivot)

    wb.save(OUTPUT_FILE)
    print(f"\nArchivo generado: {OUTPUT_FILE}")
    print(f"  Verde (CLEAR)      : {zero_count} cuentas con balance 0")
    print(f"  Rojo  (OPEN ITEMS) : {len(pivot) - zero_count} cuentas con balance pendiente")


if __name__ == "__main__":
    download_report()
    build_pivot()
