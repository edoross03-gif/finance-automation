# @title Script di Riconciliazione SAP vs SDI (Agenzia delle Entrate)
"""
Nome: Riconciliazione SAP vs SDI
Descrizione: Riscontra i documenti contabili SAP con i tracciati dell'Agenzia delle Entrate (SDI) per rilevare fatture non pervenute.
Categoria: SDI & Fatturazione

Istruzioni:
■ FILE DI INPUT RICHIESTI:
  1. CSV SDI (Agenzia delle Entrate):
     • Colonna attesa: "Numero fattura / Documento" (codici puliti dagli apostrofi).
  2. Excel SAP:
     • Colonna attesa: "Documento contabile".

■ MODALITÀ D'USO:
  1. Carica il CSV SDI e l'Excel SAP.
  2. Selezionali nei rispettivi dropdown.
  3. Clicca su Esegui per rilevare quali fatture emesse su SAP non risultano pervenute allo SDI.
"""
import pandas as pd
import os
import sys
from google.colab import files

# Nome del file Excel di output
OUTPUT_XLSX = os.environ.get("OUT_EXCEL", "Riconciliazione_SAP_SDI.xlsx")

def load_sdi_data(path: str) -> pd.DataFrame:
    """Carica il file CSV dell'Agenzia delle Entrate e pulisce il numero fattura."""
    # Lettura del CSV con separatore punto e virgola
    df = pd.read_csv(path, sep=";", encoding="utf-8-sig")

    col_sdi = "Numero fattura / Documento"
    if col_sdi not in df.columns:
        raise ValueError(f"Colonna '{col_sdi}' non trovata nel file SDI. Verifica l'intestazione.")

    # Pulizia: rimozione degli apostrofi ed eventuali spazi bianchi ai margini
    df[col_sdi] = df[col_sdi].astype(str).str.replace("'", "", regex=False).str.strip()
    return df

def load_sap_data(path: str) -> pd.DataFrame:
    """Carica il file Excel di SAP."""
    df = pd.read_excel(path, engine="openpyxl")

    col_sap = "Documento contabile"
    if col_sap not in df.columns:
        raise ValueError(f"Colonna '{col_sap}' non trovata nel file SAP. Verifica l'intestazione.")

    # Normalizzazione a stringa per il confronto uniforme
    df[col_sap] = df[col_sap].astype(str).str.strip()
    return df

def main():
    print("=" * 60)
    print("  VERIFICA PASSAGGIO FATTURE: SAP vs AGENZIA DELLE ENTRATE")
    print("=" * 60)

    # 1. Caricamento file SDI
    print("\n⬆️ Carica il file CSV dell'Agenzia delle Entrate (SDI)...")
    uploaded_sdi = files.upload()
    if not uploaded_sdi:
        print("❌ Nessun file SDI caricato. Processo interrotto.")
        return
    sdi_path = list(uploaded_sdi.keys())[0]

    try:
        df_sdi = load_sdi_data(sdi_path)
        print(f"📊 File SDI caricato con successo. Righe totali: {len(df_sdi)}")
    except Exception as e:
        print(f"❌ Errore durante la lettura del file SDI: {e}")
        return

    # 2. Caricamento file SAP
    print("\n⬆️ Carica il file Excel di SAP (XLSX)...")
    uploaded_sap = files.upload()
    if not uploaded_sap:
        print("❌ Nessun file SAP caricato. Processo interrotto.")
        return
    sap_path = list(uploaded_sap.keys())[0]

    try:
        df_sap = load_sap_data(sap_path)
        print(f"📊 File SAP caricato con successo. Righe totali: {len(df_sap)}")
    except Exception as e:
        print(f"❌ Errore durante la lettura del file SAP: {e}")
        return

    # 3. Riconciliazione (Equivalente del CONTA.SE / COUNTIF)
    print("\n🔍 Analisi e riscontro dei documenti contabili...")

    col_sdi = "Numero fattura / Documento"
    col_sap = "Documento contabile"

    # Calcola le frequenze dei codici nel file SDI
    sdi_counts = df_sdi[col_sdi].value_counts()

    # Applica il CONTA.SE sul dataframe SAP
    df_sap["Conteggio su SDI"] = df_sap[col_sap].map(sdi_counts).fillna(0).astype(int)

    # Crea la colonna di Esito per identificare i mancanti
    df_sap["Esito Check"] = df_sap["Conteggio su SDI"].apply(lambda x: "OK" if x > 0 else "MANCANTE SU SDI")

    # Estrae il sottoinsieme delle sole fatture scoperte (presenti su SAP ma non su SDI)
    df_mancanti = df_sap[df_sap["Esito Check"] == "MANCANTE SU SDI"].copy()

    # 4. Output Statistiche a video
    tot_sap = len(df_sap)
    tot_ok = len(df_sap[df_sap["Esito Check"] == "OK"])
    tot_mancanti = len(df_mancanti)

    print(f"\n{'─' * 60}")
    print(f"  RISULTATI DEL CHECK")
    print(f"{'─' * 60}")
    print(f"  Fatture totali estratte da SAP:   {tot_sap}")
    print(f"  ✅ Consegnate/Presenti su SDI:    {tot_ok}")
    print(f"  ❌ NON TROVATE su SDI:            {tot_mancanti}")

    if tot_mancanti > 0:
        print(f"⚠️  Attenzione! Ci sono {tot_mancanti} fatture su SAP non pervenute allo SDI.")
    else:
        print("\n🎉 Ottimo! Tutte le fatture SAP risultano registrate sul file dello SDI.")

    # 5. Salvataggio del report Excel multifoglio
    print(f"\n💾 Generazione del file di output in corso: {OUTPUT_XLSX}")
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        df_sap.to_excel(writer, sheet_name="Report_Totale_SAP", index=False)
        df_mancanti.to_excel(writer, sheet_name="Solo_Mancanti_SDI", index=False)
        df_sdi.to_excel(writer, sheet_name="Dati_SDI_Puliti", index=False)

    # Abilita il download automatico tramite la dashboard
    files.download(OUTPUT_XLSX)
    print("\n✅ Fatto!")

if __name__ == "__main__":
    main()
