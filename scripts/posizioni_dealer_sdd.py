"""
Nome: Posizioni Dealer Sdd
Descrizione: Raggruppa ed elabora i flussi e le scadenze SDD (Sepa Direct Debit) per i dealer a partire dall'export contabile SAP.
Categoria: Utility

Istruzioni:
■ FILE DI INPUT RICHIESTI:
  1. Export Excel SAP:
     • Struttura: Il foglio deve chiamarsi "Esportazione SAPUI5" e contenere 19 colonne (formato standard SAP) o 18 colonne (formato anonimizzato senza Nome del cliente).
     • Colonne attese: Cliente, Nome del cliente, Data doc. contabile, Documento contabile, Riferimento, Testo posizione, Attribuzione, Tipo doc. contabile, Importo div. soc., Scad. importo netto, Netta scad. simb., Stato pareggio, Banca interna, Modalit pagamento, Cond. pagamento, Data di pareggio, Doc. cont. pareggio, Pagamento, Partita pareggiata.

■ MODALITÀ D'USO:
  1. Carica il file Excel delle posizioni dealer.
  2. Selezionalo nel dropdown.
  3. Scegli la modalità ('all' per tutti i dealer, o 'single' per indicare un singolo dealer specifico tramite codice K).
  4. Premi Esegui. Il file Excel finale formattato e suddiviso apparirà tra i File Generati.
"""


# --- Latest Version Code Cell ---
# @title
# ============================================================
# ELABORAZIONE POSIZIONI SDD DA EXPORT SAP - VERSIONE COLAB
# Output 1: Excel con tutti i dealer separati e ordinati
# Output 2: Excel con un solo dealer scelto tramite codice K
# ============================================================

import pandas as pd
import numpy as np
import io
import warnings
from collections import Counter
from datetime import datetime
from google.colab import files
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# ============================================================
# CONFIG
# ============================================================

MESI = {
    1: "Gennaio", 2: "Febbraio", 3: "Marzo", 4: "Aprile",
    5: "Maggio", 6: "Giugno", 7: "Luglio", 8: "Agosto",
    9: "Settembre", 10: "Ottobre", 11: "Novembre", 12: "Dicembre"
}

OUTPUT_COLUMNS = [
    "Data doc. contabile",
    "Documento contabile",
    "Riferimento",
    "Testo posizione",
    "Attribuzione",
    "Tipo doc. contabile",
    "Importo (div. soc.)",
    "Scad. importo netto",
    "Doc. cont. pareggio"
]

SAP_EXPECTED_COLUMNS = [
    "Cliente", "Nome del cliente", "Data doc. contabile", "Documento contabile",
    "Riferimento", "Testo posizione", "Attribuzione", "Tipo doc. contabile",
    "Importo div. soc.", "Scad. importo netto", "Netta scad. simb.",
    "Stato pareggio", "Banca interna", "Modalit pagamento", "Cond. pagamento",
    "Data di pareggio", "Doc. cont. pareggio", "Pagamento", "Partita pareggiata"
]

TIPO_DOC_LABEL = {
    "CD": "Cessione del Credito",
    "W1": "Fattura",
    "W4": "Nota di Credito",
    "W7": "Nota di Debito",
    "ZD": "Fattura",
    "DZ": "Incasso",
    "DA": "Insoluto",
    "ZP": "Pagamento ns./vs favore",
    "WP": "Acconto Cliente",
    "AB": "Rettifica",
    "SU": ""
}

# soglia: primo giorno di N mesi prima rispetto a oggi
N_MESI_DEFAULT = 3

# ============================================================
# FUNZIONI BASE
# ============================================================

def clean(v):
    if pd.isna(v):
        return ""
    s = str(v).strip()
    return "" if s in ("nan", "None", "NaT", "null") else s

def to_num(v):
    try:
        if pd.isna(v) or clean(v) == "":
            return np.nan
        return float(v)
    except:
        return np.nan

def to_int_doc(v):
    try:
        if pd.isna(v) or clean(v) == "":
            return None
        return int(round(float(v)))
    except:
        try:
            return int(str(v).strip())
        except:
            return None

def fmt_doc(v):
    i = to_int_doc(v)
    return "" if i is None else str(i)

def to_ts(v):
    try:
        return pd.to_datetime(v, errors="coerce")
    except:
        return pd.NaT

def fmt_date(v):
    try:
        if pd.isna(v):
            return ""
        return pd.to_datetime(v).strftime("%d/%m/%Y")
    except:
        return ""

def blank_row():
    return {c: "" for c in OUTPUT_COLUMNS}

def first_day_n_months_before(today=None, n_months=3):
    if today is None:
        today = pd.Timestamp.today().normalize()
    else:
        today = pd.Timestamp(today).normalize()
    first_current_month = pd.Timestamp(year=today.year, month=today.month, day=1)
    return first_current_month - pd.DateOffset(months=n_months)

def safe_sheet_name(name, used_names):
    base = str(name)[:31]
    if base not in used_names:
        used_names.add(base)
        return base
    i = 1
    while True:
        candidate = f"{base[:28]}_{i}"
        if candidate not in used_names:
            used_names.add(candidate)
            return candidate
        i += 1

# ============================================================
# NORMALIZZAZIONE COLONNE
# ============================================================

def normalize_columns(df):
    cols = list(df.columns)

    if len(cols) == 19:
        df = df.iloc[:, :19].copy()
        df.columns = SAP_EXPECTED_COLUMNS
        return df

    if len(cols) == 18:
        # File anonimizzato senza "Nome del cliente"
        expected_18 = [
            "Cliente",
            "Data doc. contabile",
            "Documento contabile",
            "Riferimento",
            "Testo posizione",
            "Attribuzione",
            "Tipo doc. contabile",
            "Importo div. soc.",
            "Scad. importo netto",
            "Netta scad. simb.",
            "Stato pareggio",
            "Banca interna",
            "Modalit pagamento",
            "Cond. pagamento",
            "Data di pareggio",
            "Doc. cont. pareggio",
            "Pagamento",
            "Partita pareggiata"
        ]
        df = df.iloc[:, :18].copy()
        df.columns = expected_18
        df["Nome del cliente"] = ""
        return df[
            [
                "Cliente", "Nome del cliente", "Data doc. contabile", "Documento contabile",
                "Riferimento", "Testo posizione", "Attribuzione", "Tipo doc. contabile",
                "Importo div. soc.", "Scad. importo netto", "Netta scad. simb.",
                "Stato pareggio", "Banca interna", "Modalit pagamento", "Cond. pagamento",
                "Data di pareggio", "Doc. cont. pareggio", "Pagamento", "Partita pareggiata"
            ]
        ]

    raise ValueError(f"Il foglio ha {len(cols)} colonne, attese 18 o 19.")

def prepare_dataframe(df_raw):
    df = normalize_columns(df_raw)

    date_cols = ["Data doc. contabile", "Scad. importo netto", "Data di pareggio"]
    for c in date_cols:
        df[c] = pd.to_datetime(df[c], errors="coerce")

    num_cols = ["Documento contabile", "Importo div. soc.", "Doc. cont. pareggio"]
    for c in num_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    for c in df.columns:
        if c not in date_cols + num_cols:
            df[c] = df[c].apply(clean)

    return df

# ============================================================
# TRASFORMAZIONI CAMPI
# ============================================================

def trasforma_riferimento(row):
    tipo = clean(row["Tipo doc. contabile"])
    rif = clean(row["Riferimento"])
    testo = clean(row["Testo posizione"])

    if tipo == "CD":
        return f"IMEI {rif}" if rif else ""

    if tipo == "W1":
        if rif.startswith("O"):
            return f"Ordine {rif[1:]}"
        if rif.startswith("L"):
            return f"Lok-Me {rif[1:]}"
        if rif.startswith("A"):
            return f"Attivazione {rif[1:]}"
        return rif

    if tipo == "ZD":
        return rif

    if tipo == "W4":
        return ""

    if tipo == "W7":
        return testo if testo else ""

    if tipo == "DZ":
        return rif

    if tipo == "ZP":
        return ""

    if tipo == "WP":
        return rif

    if tipo == "DA":
        return rif

    if tipo == "AB":
        return rif

    if tipo == "SU":
        return ""

    return rif

def trasforma_testo(row):
    tipo = clean(row["Tipo doc. contabile"])
    testo = clean(row["Testo posizione"])

    if tipo == "CD":
        return testo
    if tipo == "W7":
        return testo
    return ""

def tipo_doc_label(row):
    tipo = clean(row["Tipo doc. contabile"])
    data_pareggio = row["Data di pareggio"]
    importo = to_num(row["Importo div. soc."])

    if tipo == "ZP":
        if pd.notna(data_pareggio):
            mese = MESI.get(data_pareggio.month, "")
            anno = data_pareggio.year
            if pd.notna(importo) and importo > 0:
                return f"Pagamento vs. favore {mese} {anno}"
            return f"Pagamento ns. favore {mese} {anno}"
        return "Pagamento ns./vs favore"

    return TIPO_DOC_LABEL.get(tipo, tipo)

def su_special_label(row, in_group=False):
    testo = clean(row["Testo posizione"])
    doc_num = fmt_doc(row["Documento contabile"])
    importo = to_num(row["Importo div. soc."])

    if "Credito al" in testo:
        return testo
    if "Debito al" in testo:
        return testo

    if in_group:
        return f"Pareggio partite doc {doc_num}"

    if pd.notna(importo):
        if importo < 0:
            return f"Credito generico doc {doc_num}"
        else:
            return f"Debito generico doc {doc_num}"

    return f"Pareggio partite doc {doc_num}"

def output_row_from_source(row, in_group=False):
    tipo = clean(row["Tipo doc. contabile"])

    tipo_label = tipo_doc_label(row)
    if tipo == "SU":
        tipo_label = su_special_label(row, in_group=in_group)

    return {
        "Data doc. contabile": fmt_date(row["Data doc. contabile"]),
        "Documento contabile": fmt_doc(row["Documento contabile"]),
        "Riferimento": trasforma_riferimento(row),
        "Testo posizione": trasforma_testo(row),
        "Attribuzione": clean(row["Attribuzione"]),
        "Tipo doc. contabile": tipo_label,
        "Importo (div. soc.)": round(float(row["Importo div. soc."]), 2) if pd.notna(row["Importo div. soc."]) else "",
        "Scad. importo netto": fmt_date(row["Scad. importo netto"]),
        "Doc. cont. pareggio": fmt_doc(row["Doc. cont. pareggio"])
    }

# ============================================================
# PULIZIA STORNI GENERICI
# Regola prudente: stesso doc contabile, stesso doc pareggio, stesso tipo, importi opposti
# ============================================================

def remove_generic_storni(df):
    work = df.copy()
    work["_doc"] = work["Documento contabile"].apply(to_int_doc)
    work["_docpar"] = work["Doc. cont. pareggio"].apply(to_int_doc)
    work["_imp"] = work["Importo div. soc."].apply(to_num)

    to_remove = set()

    grouped = work.groupby(["_doc", "_docpar", "Tipo doc. contabile"], dropna=False)
    for _, g in grouped:
        idxs = list(g.index)
        used = set()
        for i in range(len(idxs)):
            if idxs[i] in used:
                continue
            imp_i = work.loc[idxs[i], "_imp"]
            if pd.isna(imp_i):
                continue
            for j in range(i + 1, len(idxs)):
                if idxs[j] in used:
                    continue
                imp_j = work.loc[idxs[j], "_imp"]
                if pd.isna(imp_j):
                    continue
                if abs(imp_i + imp_j) < 0.01:
                    used.add(idxs[i])
                    used.add(idxs[j])
                    to_remove.add(idxs[i])
                    to_remove.add(idxs[j])
                    break

    return df.drop(index=list(to_remove)).copy()

# ============================================================
# ANALISI SU SUL DATASET COMPLETO
# ============================================================

def analyze_su_pairs(df_full):
    su = df_full[df_full["Tipo doc. contabile"] == "SU"].copy()
    su["_doc"] = su["Documento contabile"].apply(to_int_doc)
    su["_docpar"] = su["Doc. cont. pareggio"].apply(to_int_doc)
    su["_imp"] = su["Importo div. soc."].apply(to_num)

    valid_indices = set()
    warnings_list = []

    for doc_num, g in su.groupby("_doc"):
        if doc_num is None:
            continue

        if len(g) != 2:
            warnings_list.append(f"SU anomalo doc {doc_num}: trovate {len(g)} righe, attese 2")
            continue

        idxs = list(g.index)
        r1 = g.loc[idxs[0]]
        r2 = g.loc[idxs[1]]

        imp_sum = (0 if pd.isna(r1["_imp"]) else r1["_imp"]) + (0 if pd.isna(r2["_imp"]) else r2["_imp"])
        if abs(imp_sum) >= 0.01:
            warnings_list.append(f"SU doc {doc_num}: somma importi != 0 ({imp_sum:.2f})")

        valid_indices.update(idxs)

    return valid_indices, warnings_list

# ============================================================
# ORDINAMENTI
# ============================================================

def group_month(rows_df):
    vals = rows_df["Scad. importo netto"].dropna()
    if len(vals) == 0:
        return (9999, 99)
    counts = Counter((d.year, d.month) for d in vals)
    return counts.most_common(1)[0][0]

def group_sort_key(rows_df):
    types = set(rows_df["Tipo doc. contabile"].astype(str).tolist())
    if "ZP" in types:
        return 3
    if "DZ" in types:
        return 2
    if "WP" in types:
        return 1
    return 0

def sort_group_rows(rows_df):
    non_su = rows_df[rows_df["Tipo doc. contabile"] != "SU"].copy()
    su = rows_df[rows_df["Tipo doc. contabile"] == "SU"].copy()

    non_su = non_su.sort_values(
        by=["Data doc. contabile", "Importo div. soc."],
        ascending=[False, False],
        na_position="last"
    )

    # SU in coda, positivi prima dei negativi
    su["_sign_order"] = su["Importo div. soc."].apply(lambda x: 0 if pd.notna(x) and x >= 0 else 1)
    su = su.sort_values(
        by=["_sign_order", "Importo div. soc."],
        ascending=[True, False],
        na_position="last"
    ).drop(columns=["_sign_order"], errors="ignore")

    return pd.concat([non_su, su], axis=0)

# ============================================================
# COSTRUZIONE OUTPUT DEALER
# ============================================================

def build_dealer_output(df_dealer_full, n_mesi=N_MESI_DEFAULT):
    # analisi SU sul dataset completo dealer
    valid_su_indices, su_warnings = analyze_su_pairs(df_dealer_full)

    # filtro validità SU: se SU e non parte di coppia valida, lo escludiamo
    df_work = df_dealer_full.copy()
    mask_invalid_su = (df_work["Tipo doc. contabile"] == "SU") & (~df_work.index.isin(valid_su_indices))
    df_work = df_work[~mask_invalid_su].copy()

    # rimozione storni generici
    df_work = remove_generic_storni(df_work)

    # filtro temporale SOLO dopo analisi SU
    soglia = first_day_n_months_before(n_months=n_mesi)
    df_show = df_work[df_work["Scad. importo netto"] >= soglia].copy()

    # gruppi / aperte
    docpar_numeric = df_show["Doc. cont. pareggio"].apply(to_int_doc)
    stato_pareggio = df_show["Stato pareggio"].apply(lambda x: clean(x))

    # SU credito aperto da mettere in testa
    su_credito_open = df_show[
        (df_show["Tipo doc. contabile"] == "SU") &
        (df_show["Doc. cont. pareggio"].isna()) &
        (df_show["Testo posizione"].astype(str).str.contains("Credito al", case=False, na=False))
    ].copy()

    # aperte ordinarie
    open_positions = df_show[
        (
            (stato_pareggio == "2") |
            (df_show["Doc. cont. pareggio"].isna())
        ) &
        ~(
            (df_show["Tipo doc. contabile"] == "SU") &
            (df_show["Doc. cont. pareggio"].isna()) &
            (df_show["Testo posizione"].astype(str).str.contains("Credito al", case=False, na=False))
        )
    ].copy()

    grouped_positions = df_show[
        ~df_show.index.isin(open_positions.index) &
        ~df_show.index.isin(su_credito_open.index)
    ].copy()

    # costruzione gruppi doc_par
    grouped_positions["_docpar"] = grouped_positions["Doc. cont. pareggio"].apply(to_int_doc)
    groups = {}
    for dp, g in grouped_positions.groupby("_docpar"):
        if dp is None or pd.isna(dp):
            continue
        groups[dp] = g.copy()

    # mapping gruppi per mese
    months = {}
    for dp, g in groups.items():
        ym = group_month(g)
        months.setdefault(ym, []).append(dp)

    for ym in months:
        months[ym] = sorted(months[ym], key=lambda dp: (group_sort_key(groups[dp]), dp))

    output_rows = []
    output_types = []

    # 1) SU "Credito al..." aperto in testa
    if len(su_credito_open) > 0:
        su_credito_open = su_credito_open.sort_values(
            by=["Data doc. contabile"], ascending=[False], na_position="last"
        )
        for _, r in su_credito_open.iterrows():
            output_rows.append(output_row_from_source(r, in_group=False))
            output_types.append("su_open")
        output_rows.append(blank_row())
        output_types.append("blank")

    # 2) gruppi per mese
    for ym in sorted(months.keys()):
        y, m = ym
        if ym == (9999, 99):
            continue

        hdr = blank_row()
        hdr["Tipo doc. contabile"] = f"{MESI.get(m, '?')} {y}"
        output_rows.append(hdr)
        output_types.append("month_hdr")

        month_total = 0.0

        for dp in months[ym]:
            grp = sort_group_rows(groups[dp])

            for _, r in grp.iterrows():
                out = output_row_from_source(r, in_group=True)
                output_rows.append(out)

                tipo = clean(r["Tipo doc. contabile"])
                if tipo == "SU":
                    output_types.append("su")
                elif tipo == "ZP":
                    output_types.append("zp")
                elif tipo == "DZ":
                    output_types.append("dz")
                elif tipo == "WP":
                    output_types.append("wp")
                else:
                    output_types.append("data")

                if pd.notna(r["Importo div. soc."]):
                    month_total += float(r["Importo div. soc."])

            output_rows.append(blank_row())
            output_types.append("sep")

        tot = blank_row()
        tot["Tipo doc. contabile"] = f"Totale {MESI.get(m, '?')} {y}"
        tot["Importo (div. soc.)"] = round(month_total, 2)
        output_rows.append(tot)
        output_types.append("month_tot")

        output_rows.append(blank_row())
        output_types.append("blank")

    # 3) posizioni aperte
    if len(open_positions) > 0:
        hdr = blank_row()
        hdr["Tipo doc. contabile"] = "POSIZIONI APERTE"
        output_rows.append(hdr)
        output_types.append("open_hdr")

        open_positions["_ym"] = open_positions["Scad. importo netto"].apply(
            lambda d: (d.year, d.month) if pd.notna(d) else (9999, 99)
        )

        for ym, g in sorted(open_positions.groupby("_ym"), key=lambda x: x[0]):
            y, m = ym
            if ym != (9999, 99):
                sub = blank_row()
                sub["Tipo doc. contabile"] = f"Scadenza {MESI.get(m, '?')} {y}"
                output_rows.append(sub)
                output_types.append("open_month")

            g = g.sort_values(by=["Data doc. contabile"], ascending=[False], na_position="last")

            for _, r in g.iterrows():
                out = output_row_from_source(r, in_group=False)
                output_rows.append(out)
                if clean(r["Tipo doc. contabile"]) == "SU":
                    output_types.append("su")
                else:
                    output_types.append("open")

        tot_open = round(open_positions["Importo div. soc."].fillna(0).sum(), 2)
        tot = blank_row()
        tot["Tipo doc. contabile"] = "Totale Posizioni Aperte"
        tot["Importo (div. soc.)"] = tot_open
        output_rows.append(tot)
        output_types.append("open_tot")

    output_df = pd.DataFrame(output_rows, columns=OUTPUT_COLUMNS)
    return output_df, output_types, su_warnings, soglia

# ============================================================
# WRITER EXCEL
# ============================================================

STYLE = {
    "data":       ("FFFFFF","000000",False,15),
    "su":         ("FFF2CC","806000",True,15),
    "su_open":    ("FFD966","000000",True,18),
    "zp":         ("E2EFDA","006100",True,15),
    "dz":         ("D9EAD3","000000",False,15),
    "wp":         ("D9E1F2","000000",False,15),
    "month_hdr":  ("2E5D9E","FFFFFF",True,22),
    "month_tot":  ("BDD7EE","1F4E79",True,18),
    "open_hdr":   ("8B4513","FFFFFF",True,22),
    "open_month": ("D2691E","FFFFFF",True,18),
    "open":       ("FFF5EE","8B4513",False,15),
    "open_tot":   ("F4A460","000000",True,18),
    "blank":      ("FFFFFF","000000",False,4),
    "sep":        ("F5F5F5","000000",False,9),
}

def style_worksheet(ws, row_types):
    widths = [14, 18, 28, 30, 18, 34, 16, 16, 18]

    for ci, cell in enumerate(ws[1], 1):
        cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        cell.fill = PatternFill("solid", fgColor="1F4E79")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(ci)].width = widths[ci - 1]
    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"

    side = Side(style="thin", color="D0D0D0")
    border = Border(top=side, bottom=side, left=side, right=side)

    for excel_row_idx, rtype in enumerate(row_types, start=2):
        fill_c, font_c, bold, ht = STYLE.get(rtype, STYLE["data"])

        # evidenzia SU positivi in gruppo
        if rtype == "su":
            val = ws.cell(excel_row_idx, 7).value
            try:
                if val is not None and float(val) > 0:
                    fill_c = "FCE4D6"
                    font_c = "C00000"
            except:
                pass

        for ci in range(1, len(OUTPUT_COLUMNS) + 1):
            cell = ws.cell(excel_row_idx, ci)
            cell.fill = PatternFill("solid", fgColor=fill_c)
            cell.font = Font(name="Arial", size=10 if ht >= 15 else 6, bold=bold, color=font_c)
            cell.border = border

            if ci == 7 and isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0.00;[Red]-#,##0.00'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        ws.row_dimensions[excel_row_idx].height = ht

def add_df_to_sheet(ws, df_out):
    ws.append(OUTPUT_COLUMNS)
    for _, row in df_out.iterrows():
        ws.append([row.get(c, "") for c in OUTPUT_COLUMNS])

def create_workbook_all_dealers(results_dict, meta_info=None):
    wb = Workbook()
    first = True
    used_names = set()

    # sheet indice
    if first:
        ws_index = wb.active
        ws_index.title = "Indice"
        first = False

    ws_index.append(["Dealer", "Nome dealer", "Righe output", "Note"])
    for c in ws_index[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E79")
        c.alignment = Alignment(horizontal="center")

    for dealer_code, payload in results_dict.items():
        dealer_name = payload["dealer_name"]
        df_out = payload["df_out"]
        notes = "; ".join(payload["warnings"][:3]) if payload["warnings"] else ""
        ws_index.append([dealer_code, dealer_name, len(df_out), notes])

    for col, width in zip(["A","B","C","D"], [16, 36, 14, 70]):
        ws_index.column_dimensions[col].width = width

    # sheet per dealer
    for dealer_code, payload in results_dict.items():
        dealer_name = payload["dealer_name"]
        df_out = payload["df_out"]
        row_types = payload["row_types"]

        sheet_name = safe_sheet_name(f"{dealer_code}", used_names)
        ws = wb.create_sheet(title=sheet_name)
        add_df_to_sheet(ws, df_out)
        style_worksheet(ws, row_types)

    return wb

def create_workbook_single_dealer(dealer_code, dealer_name, df_out, row_types, warnings_list=None):
    wb = Workbook()
    ws = wb.active
    ws.title = safe_sheet_name(f"{dealer_code}", set())

    add_df_to_sheet(ws, df_out)
    style_worksheet(ws, row_types)

    if warnings_list:
        ws_note = wb.create_sheet(title="Note")
        ws_note.append(["Messaggi"])
        ws_note["A1"].font = Font(bold=True, color="FFFFFF")
        ws_note["A1"].fill = PatternFill("solid", fgColor="1F4E79")
        for w in warnings_list:
            ws_note.append([w])
        ws_note.column_dimensions["A"].width = 120

    return wb

# ============================================================
# PROCESSORI PRINCIPALI
# ============================================================

def process_all_dealers(df_full, n_mesi=N_MESI_DEFAULT):
    results = {}

    for dealer_code, g in df_full.groupby("Cliente", dropna=False):
        dealer_code = clean(dealer_code)
        if dealer_code == "":
            continue

        dealer_name = clean(g["Nome del cliente"].iloc[0]) if len(g) > 0 else ""
        df_out, row_types, warnings_list, soglia = build_dealer_output(g.copy(), n_mesi=n_mesi)

        results[dealer_code] = {
            "dealer_name": dealer_name,
            "df_out": df_out,
            "row_types": row_types,
            "warnings": warnings_list,
            "soglia": soglia
        }

    return results

def process_single_dealer(df_full, dealer_code, n_mesi=N_MESI_DEFAULT):
    dealer_code = clean(dealer_code).upper()

    g = df_full[df_full["Cliente"].astype(str).str.upper() == dealer_code].copy()
    if len(g) == 0:
        raise ValueError(f"Dealer {dealer_code} non trovato nel file.")

    dealer_name = clean(g["Nome del cliente"].iloc[0])
    df_out, row_types, warnings_list, soglia = build_dealer_output(g, n_mesi=n_mesi)

    return dealer_name, df_out, row_types, warnings_list, soglia

# ============================================================
# MAIN COLAB
# ============================================================

print("=" * 80)
print("ELABORAZIONE POSIZIONI SDD / EXPORT SAP - MULTI DEALER + SINGLE DEALER")
print("=" * 80)
print("\nCarica il file Excel 'Posizioni_SDD....xlsx'")

uploaded = files.upload()
filename = list(uploaded.keys())[0]

# lettura foglio
excel_bytes = io.BytesIO(uploaded[filename])
raw = pd.read_excel(excel_bytes, sheet_name="Esportazione SAPUI5")
df_full = prepare_dataframe(raw)

print(f"\nFile caricato: {filename}")
print(f"Righe totali lette: {len(df_full)}")
print(f"Dealer unici trovati: {df_full['Cliente'].nunique()}")

# parametri utente
print("\n--- PARAMETRI ---")
mode = input("Scegli modalità ('all' = tutti i dealer, 'single' = un dealer): ").strip().lower()
n_mesi_input = input(f"Numero mesi per soglia temporale [default {N_MESI_DEFAULT}]: ").strip()

if n_mesi_input == "":
    n_mesi = N_MESI_DEFAULT
else:
    n_mesi = int(n_mesi_input)

# MODALITA 1 - TUTTI I DEALER
if mode == "all":
    print("\nElaborazione di tutti i dealer in corso...")
    results = process_all_dealers(df_full, n_mesi=n_mesi)

    out_name = filename.rsplit(".", 1)[0] + "_ALL_DEALERS.xlsx"
    wb = create_workbook_all_dealers(results)
    wb.save(out_name)

    print(f"\nCompletato. File creato: {out_name}")
    print(f"Dealer elaborati: {len(results)}")
    files.download(out_name)

# MODALITA 2 - SINGOLO DEALER
elif mode == "single":
    dealer_code = input("Inserisci codice dealer (es. K11725): ").strip().upper()

    print(f"\nElaborazione dealer {dealer_code} in corso...")
    dealer_name, df_out, row_types, warnings_list, soglia = process_single_dealer(
        df_full, dealer_code=dealer_code, n_mesi=n_mesi
    )

    out_name = filename.rsplit(".", 1)[0] + f"_{dealer_code}.xlsx"
    wb = create_workbook_single_dealer(dealer_code, dealer_name, df_out, row_types, warnings_list)
    wb.save(out_name)

    print(f"\nCompletato. File creato: {out_name}")
    print(f"Dealer: {dealer_code} - {dealer_name}")
    print(f"Soglia temporale applicata: {soglia.strftime('%d/%m/%Y')}")
    if warnings_list:
        print("\nWarning SU:")
        for w in warnings_list[:10]:
            print("-", w)

    files.download(out_name)

else:
    raise ValueError("Modalità non valida. Usa 'all' oppure 'single'.")
