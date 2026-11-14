# @title 2. CONTRO-CHECK FATTURE MANCANTI (Solo Nuovo File SDI)
"""
Nome: Controcheck Fatture SDI
Descrizione: Controlla se le fatture precedentemente contrassegnate come "mancanti" compaiono nel nuovo tracciato SDI.
Categoria: SDI & Fatturazione

Istruzioni:
■ FILE DI INPUT RICHIESTI:
  1. Report Riconciliazione Precedente (Excel):
     • Struttura: Deve contenere la scheda denominata "Solo_Mancanti_SDI" con la colonna "Documento contabile".
  2. Nuovo CSV SDI aggiornato:
     • Colonna attesa: "Numero fattura / Documento".

■ MODALITÀ D'USO:
  1. Carica il report del precedente controllo (Riconciliazione_SAP_SDI.xlsx).
  2. Carica il CSV SDI aggiornato ad oggi.
  3. Clicca su Esegui per verificare se le fatture mancanti sono state registrate.
"""
import pandas as pd
import os
import sys
from google.colab import files

col_sap = "Documento contabile"
col_sdi = "Numero fattura / Documento"
fatture_da_controllare = []

print("=" * 60)
print("  CONTRO-CHECK DIFFERITO FATTURE DA EMETTERE")
print("=" * 60)

# 1. GESTIONE MEMORIA / CARICAMENTO DATI PRECEDENTI
# Nota: nell'esecuzione tramite dashboard locale, la sessione non è condivisa,
# quindi carichiamo sempre il file Excel del controllo precedente.
print("ℹ️ Esecuzione in modalità server: caricamento del file precedente necessario.")
print("⬆️ Carica il file Excel di output generato nel precedente controllo ('Riconciliazione_SAP_SDI.xlsx'):")
uploaded_prev = files.upload()

if uploaded_prev:
    prev_path = list(uploaded_prev.keys())[0]
    try:
        # Legge solo la scheda delle fatture che erano risultate mancanti
        df_prev = pd.read_excel(prev_path, sheet_name="Solo_Mancanti_SDI", engine="openpyxl")
        fatture_da_controllare = df_prev[col_sap].astype(str).str.strip().tolist()
        print(f"✅ Caricate con successo {len(fatture_da_controllare)} fatture da monitorare.")
    except Exception as e:
        print(f"❌ Errore nel caricamento del file Excel precedente: {e}")
else:
    print("❌ Operazione annullata. Nessun file precedente caricato.")

# 2. CARICAMENTO NUOVO FILE SDI E CONFRONTO
if fatture_da_controllare:
    print("\n⬆️ Ora carica il NUOVO file CSV dell'Agenzia delle Entrate (SDI) aggiornato ad oggi:")
    uploaded_sdi_new = files.upload()

    if uploaded_sdi_new:
        sdi_new_path = list(uploaded_sdi_new.keys())[0]
        try:
            # Carica e pulisce il nuovo tracciato SDI dagli apostrofi
            df_sdi_new = pd.read_csv(sdi_new_path, sep=";", encoding="utf-8-sig")
            df_sdi_new[col_sdi] = df_sdi_new[col_sdi].astype(str).str.replace("'", "", regex=False).str.strip()

            # Trasforma in un set per una ricerca istantanea ad alte prestazioni
            fatture_sdi_presenti = set(df_sdi_new[col_sdi].tolist())

            # Verifica quali delle vecchie fatture NON sono ancora presenti nel nuovo file SDI
            ancora_mancanti = [f for f in fatture_da_controllare if f not in fatture_sdi_presenti]

            # 3. OUTPUT DEI RISULTATI
            print(f"\n{'─' * 60}")
            print(f"  ESITO DEL VERIFICA A DISTANZA DI GIORNI")
            print(f"{'─' * 60}")

            if not ancora_mancanti:
                print("🎉 Tutto a posto sono state emesse")
            else:
                # Trasforma la lista delle fatture ancora scoperte in una stringa separata da virgole
                stringa_fatture = ", ".join(ancora_mancanti)
                print(f"❌ Mancano ancora da emettere fatture: {stringa_fatture}")

        except Exception as e:
            print(f"❌ Errore durante l'analisi del nuovo file SDI: {e}")
    else:
        print("❌ Operazione annullata. Nuovo file SDI non caricato.")
else:
    print("\n❌ Impossibile procedere: nessuna lista di fatture arretrate trovata o caricata.")
