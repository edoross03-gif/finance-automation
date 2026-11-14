# @title
"""
Nome: Riconciliazione Hipay vs SAP
Descrizione: Riconcilia i flussi di pagamento Hipay (Captured/Settled) con le registrazioni contabili di SAP.
Categoria: Riconciliazione

Istruzioni:
■ FILE DI INPUT RICHIESTI:
  1. Export CSV Hipay:
     • Colonne attese: Una colonna contenente "ordine" (ID ordine), una colonna contenente "stato" (Stato, con valori Captured o Settled), Authorized, Op. di pagamento, Ordinato, Operazione di pagamento.
  2. Export Excel SAP:
     • Colonne attese: Testo pos. doc.cont. (contenente il codice transazione), Data registrazione, Importo in div. soc.

■ MODALITÀ D'USO:
  1. Carica il CSV Hipay e l'Excel SAP.
  2. Seleziona i rispettivi file nei dropdown.
  3. Clicca su Esegui per incrociare le date e gli importi abbinando le transazioni e individuando le anomalie.
"""
import pandas as pd
import re
import os
import sys
from datetime import datetime
from google.colab import files # Import files for file upload

# =============================================================================
# 1. CONFIGURAZIONE – modifica questi path per Colab
# =============================================================================
HIPAY_CSV_NAME  = "OGNI_ACCOUNT_20260506-122603_297379_56.csv"
SAP_XLSX_NAME   = "Partite singole (5).xlsx"
OUTPUT_XLSX = os.environ.get("OUT_EXCEL", "Riconciliazione_risultato.xlsx")

# =============================================================================
# 2. LETTURA E PULIZIA HIPAY
# =============================================================================
def load_hipay(path: str) -> pd.DataFrame:
    """Legge l'export CSV di Hipay e restituisce solo Captured / Settled
    con il codice transazione estratto."""

    df = pd.read_csv(path, sep=";", encoding="utf-8-sig")

    # Rinomina per comodità (gestisce l'apostrofo curvo nell'header)
    col_id    = [c for c in df.columns if "ordine" in c.lower()][0]
    col_stato = [c for c in df.columns if "stato" in c.lower()][0]

    # Filtra solo Captured e Settled
    df = df[df[col_stato].isin(["Captured", "Settled"])].copy()

    # Pulisci colonne importo (formato europeo: "1.234,56 €" → float)
    for col in ["Authorized", "Op. di pagamento", "Ordinato"]:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace("\xa0", "", regex=False)   # non-breaking space
                .str.replace("€", "", regex=False)
                .str.replace(" ", "", regex=False)
                .str.replace(".", "", regex=False)       # migliaia
                .str.replace(",", ".", regex=False)      # decimale
            )
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Estrai codice transazione: LEFT(7) per P-xxxxx, LEFT(6) per ordini
    def extract_code(raw_id: str) -> str:
        if raw_id.upper().startswith("P-"):
            return raw_id[:7]
        else:
            return raw_id[:6]

    df["codice_trans"] = df[col_id].apply(extract_code).str.upper()

    # Parsing date
    col_op = "Operazione di pagamento"
    if col_op in df.columns:
        df["data_operazione"] = pd.to_datetime(
            df[col_op], format="%d/%m/%Y", errors="coerce"
        )

    return df


# =============================================================================
# 3. LETTURA E PULIZIA SAP
# =============================================================================
def load_sap(path: str) -> pd.DataFrame:
    """Legge l'export XLSX di SAP, esclude settlement e AMEX,
    e restituisce il codice transazione estratto."""

    df = pd.read_excel(path, engine="openpyxl")

    col_testo = "Testo pos. doc.cont."
    col_data  = "Data registrazione"
    col_imp   = "Importo in div. soc."

    # Escludi righe non rilevanti (AMEX, Settlement, righe vuote)
    mask_exclude = (
        df[col_testo].isna()
        | df[col_testo].str.upper().str.contains("AMEX", na=False)
        | df[col_testo].str.lower().str.contains("settlement", na=False)
    )
    df = df[~mask_exclude].copy()

    # Estrai codice transazione dal testo
    def extract_code_sap(testo: str) -> str:
        t = testo.strip()

        # Rimuovi prefissi standard
        if t.lower().startswith("saldo riga pagamento "):
            t = t[len("saldo riga pagamento "):]
        elif t.lower().startswith("saldo ordine "):
            t = t[len("saldo ordine "):]

        # LEFT(7) per righe pagamento, LEFT(6) per ordini
        if t.upper().startswith("P-"):
            return t[:7]
        else:
            return t[:6]

    df["codice_trans"] = df[col_testo].apply(extract_code_sap).str.upper()

    # Assicurati che la data sia datetime
    df[col_data] = pd.to_datetime(df[col_data], errors="coerce")

    return df


# =============================================================================
# 4. RICONCILIAZIONE
# =============================================================================
def reconcile(hipay: pd.DataFrame, sap: pd.DataFrame) -> pd.DataFrame:
    """Esegue il lookup di ogni transazione Hipay nel dataset SAP.
    Restituisce il dataframe Hipay arricchito con le info di match."""

    col_id    = [c for c in hipay.columns if "ordine" in c.lower()][0]
    col_stato = [c for c in hipay.columns if "stato" in c.lower()][0]

    # Costruisci dizionario SAP: codice → (data_registrazione, importo, testo)
    sap_lookup = {}
    for _, r in sap.iterrows():
        code = r["codice_trans"]
        sap_lookup.setdefault(code, []).append({
            "data_reg": r["Data registrazione"],
            "importo":  r["Importo in div. soc."],
            "testo":    r["Testo pos. doc.cont."],
        })

    results = []
    for _, h in hipay.iterrows():
        code = h["codice_trans"]
        match_list = sap_lookup.get(code)

        if match_list:
            m = match_list[0]  # primo match
            data_sap = m["data_reg"]
            importo_sap = m["importo"]

            # Check data
            data_hipay = h.get("data_operazione")
            date_ok = ""
            if pd.notna(data_hipay) and pd.notna(data_sap):
                date_ok = "OK" if data_hipay == data_sap else "DATA DIVERSA"

            # Check importo
            importo_hipay = h.get("Authorized")
            importo_sap = m["importo"]
            importo_ok = ""
            if pd.notna(importo_hipay) and pd.notna(importo_sap):
                importo_ok = (
                    "OK"
                    if abs(float(importo_hipay) - float(importo_sap)) < 0.01
                    else "IMPORTO DIVERSO"
                )

            results.append({
                "codice_trans":    code,
                "id_ordine_hipay": h[col_id],
                "stato":           h[col_stato],
                "importo_hipay":   importo_hipay,
                "data_hipay":      data_hipay,
                "su_SAP":          "SI",
                "data_SAP":        data_sap,
                "importo_SAP":     importo_sap,
                "check_data":      date_ok,
                "check_importo":   importo_ok,
            })
        else:
            results.append({
                "codice_trans":    code,
                "id_ordine_hipay": h[col_id],
                "stato":           h[col_stato],
                "importo_hipay":   h.get("Authorized"),
                "data_hipay":      h.get("data_operazione"),
                "su_SAP":          "NON C'È",
                "data_SAP":        None,
                "importo_SAP":     None,
                "check_data":      "",
                "check_importo":   "",
            })

    return pd.DataFrame(results)


# =============================================================================
# 5. REVERSE CHECK: transazioni SAP non presenti su Hipay
# =============================================================================
def reverse_check(hipay: pd.DataFrame, sap: pd.DataFrame) -> pd.DataFrame:
    """Trova transazioni SAP che non hanno corrispondenza su Hipay."""
    hipay_codes = set(hipay["codice_trans"].str.upper())
    missing = sap[~sap["codice_trans"].str.upper().isin(hipay_codes)].copy()
    return missing


# =============================================================================
# 6. MAIN
# =============================================================================
def main():
    print("=" * 60)
    print("  RICONCILIAZIONE HIPAY vs SAP")
    print("=" * 60)

    # Carica dati
    print("\n⬆️ Carica il file Hipay CSV...")
    uploaded_hipay = files.upload()
    if not uploaded_hipay:
        print("Nessun file Hipay CSV caricato. Annullamento.")
        return
    hipay_csv_path = list(uploaded_hipay.keys())[0]
    print(f"File Hipay CSV caricato: {hipay_csv_path}")
    hipay = load_hipay(hipay_csv_path)
    print(f"   Transazioni Captured/Settled: {len(hipay)}")

    print("\n⬆️ Carica il file SAP XLSX...")
    uploaded_sap = files.upload()
    if not uploaded_sap:
        print("Nessun file SAP XLSX caricato. Annullamento.")
        return
    sap_xlsx_path = list(uploaded_sap.keys())[0]
    print(f"File SAP XLSX caricato: {sap_xlsx_path}")
    sap = load_sap(sap_xlsx_path)
    print(f"   Movimentazioni (esclusi settlement/AMEX): {len(sap)}")

    # Riconciliazione
    print("\n🔍 Riconciliazione in corso...")
    result = reconcile(hipay, sap)

    # Riepilogo
    trovate   = result[result["su_SAP"] == "SI"]
    mancanti  = result[result["su_SAP"] == "NON C'È"]
    data_div  = trovate[trovate["check_data"] == "DATA DIVERSA"]
    imp_div   = trovate[trovate["check_importo"] == "IMPORTO DIVERSO"]

    print(f"\n{'─' * 60}")
    print(f"  RISULTATI")
    print(f"{'─' * 60}")
    print(f"  Transazioni Hipay totali:       {len(result)}")
    print(f"  ✅ Trovate su SAP:               {len(trovate)}")
    print(f"  ❌ NON trovate su SAP:            {len(mancanti)}")
    if len(data_div) > 0:
        print(f"  ⚠️  Con data diversa:             {len(data_div)}")
    if len(imp_div) > 0:
        print(f"  ⚠️  Con importo diverso:           {len(imp_div)}")

    # Dettaglio mancanti
    if len(mancanti) > 0:
        print(f"\n{'─' * 60}")
        print("  TRANSAZIONI HIPAY NON TROVATE SU SAP:")
        print(f"{'─' * 60}")
        for _, r in mancanti.iterrows():
            importo_val = r['importo_hipay']
            try:
                importo_val = float(importo_val)
                importo_str = f"€ {importo_val:>10,.2f}"
            except Exception:
                importo_str = f"{importo_val}"
            print(f"  {r['codice_trans']:>7}  |  {r['stato']:<10}  |  "
                  f"{importo_str}  |  {r['data_hipay']}")

    # Dettaglio discrepanze data
    if len(data_div) > 0:
        print(f"\n{'─' * 60}")
        print("  DISCREPANZE DI DATA:")
        print(f"{'─' * 60}")
        for _, r in data_div.iterrows():
            dh = r['data_hipay'].strftime('%d/%m/%Y') if pd.notna(r['data_hipay']) else '?'
            ds = r['data_SAP'].strftime('%d/%m/%Y') if pd.notna(r['data_SAP']) else '?'
            print(f"  {r['codice_trans']:>7}  |  Hipay: {dh}  |  SAP: {ds}")

    # Dettaglio discrepanze importo
    if len(imp_div) > 0:
        print(f"\n{'─' * 60}")
        print("  DISCREPANZE DI IMPORTO:")
        print(f"{'─' * 60}")
        for _, r in imp_div.iterrows():
            importo_h_val = r['importo_hipay']
            importo_s_val = r['importo_SAP']
            try:
                importo_h_val = float(importo_h_val)
                importo_h_str = f"€ {importo_h_val:>10,.2f}"
            except Exception:
                importo_h_str = f"{importo_h_val}"
            try:
                importo_s_val = float(importo_s_val)
                importo_s_str = f"€ {importo_s_val:>10,.2f}"
            except Exception:
                importo_s_str = f"{importo_s_val}"
            print(f"  {r['codice_trans']:>7}  |  Hipay: {importo_h_str}  |  SAP: {importo_s_str}")

    # Reverse check
    rev = reverse_check(hipay, sap)
    if len(rev) > 0:
        print(f"\n{'─' * 60}")
        print(f"  TRANSAZIONI SAP NON TROVATE SU HIPAY: {len(rev)}")
        print(f"{'─' * 60}")
        for _, r in rev.iterrows():
            d = r['Data registrazione'].strftime('%d/%m/%Y') if pd.notna(r['Data registrazione']) else '?'
            print(f"  {r['codice_trans']:>7}  |  € {r['Importo in div. soc.']:>10,.2f}"
                  f"  |  {d}  |  {r['Testo pos. doc.cont.']}")

    # Salva Excel
    print(f"\n💾 Salvataggio risultato → {OUTPUT_XLSX}")
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        result.to_excel(writer, sheet_name="Riconciliazione", index=False)
        sap.to_excel(writer, sheet_name="SAP_pulito", index=False)
        if len(rev) > 0:
            rev.to_excel(writer, sheet_name="SAP_non_in_Hipay", index=False)

    # Download per compatibilità con la dashboard
    files.download(OUTPUT_XLSX)
    print("\n✅ Fatto!")


if __name__ == "__main__":
    main()