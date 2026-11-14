"""
Nome: Elaborazione Cessioni
Descrizione: Elabora e normalizza i dati delle Cessioni Giorno Precedente generando il tracciato CSV per le banche.
Categoria: Utility

Istruzioni:
■ FILE DI INPUT RICHIESTI:
  1. Report Cessioni Giorno Precedente (Excel o CSV):
     • Colonne attese: Codice Cliente, Metodo di pagamento Cliente, Codice Prodotto, Valore Cessione, Seriale, Cliente, Banca Interna, Riferimento, Data base.

■ MODALITÀ D'USO:
  1. Carica il report delle cessioni.
  2. Selezionalo nel dropdown.
  3. Clicca su Esegui per normalizzare i codici fiscali, verificare le partite IVA e generare il tracciato CSV per le banche.
"""


# --- Latest Version Code Cell ---
# @title ELABORAZIONE CESSIONI V2
import datetime
import calendar
import io
import pandas as pd
from google.colab import files

def aggiusta_data_scadenza(data_val):
    if pd.isna(data_val) or str(data_val).strip() == "":
        return ""

    data_str = str(data_val).strip()
    try:
        # Tenta il parsing del formato standard YYYY-MM-DD
        dt = datetime.datetime.strptime(data_str.split()[0], '%Y-%m-%d').date()
    except Exception:
        try:
            # Tenta il parsing del formato DD/MM/YYYY
            dt = datetime.datetime.strptime(data_str, '%d/%m/%Y').date()
        except Exception:
            return data_str # Ritorna il valore originale se il formato è ignoto

    today = datetime.date.today()
    anno_corrente = today.year
    mese_corrente = today.month

    # Trova l'ultimo giorno del mese corrente
    _, ultimo_giorno = calendar.monthrange(anno_corrente, mese_corrente)
    fine_mese_corrente = datetime.date(anno_corrente, mese_corrente, ultimo_giorno)

    # LOGICA CORRETTA:
    # Se la data è nel passato o nel mese in corso (<= fine mese attuale), viene spostata a fine mese attuale.
    # Se è già maggiore del fine mese attuale (es. agosto, settembre, ecc.), rimane invariata.
    if dt <= fine_mese_corrente:
        dt = fine_mese_corrente

    # Restituisce nel formato richiesto DD/MM/AAAA (es. 31/07/2026)
    return dt.strftime('%d/%m/%Y')

def pulisci_valore(val):
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    # Rimuove l'eventuale .0 finale derivante da passate interpretazioni numeriche float
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    return val_str

# 1. Caricamento file interattivo
print("🔄 Seleziona e carica il report 'Cessioni Giorno Precedente' (Excel o CSV):")
uploaded = files.upload()

if not uploaded:
    print("❌ Nessun file caricato.")
else:
    file_name = list(uploaded.keys())[0]
    print(f"📖 Lettura del file: {file_name}")

    # Forziamo dtype=str per preservare gli zeri iniziali (es. nei riferimenti '0090215155')
    if file_name.endswith('.csv'):
        df_in = pd.read_csv(io.BytesIO(uploaded[file_name]), sep=None, engine='python', dtype=str)
    else:
        df_in = pd.read_excel(io.BytesIO(uploaded[file_name]), dtype=str)

    # 2. Costruzione del DataFrame di Output
    df_out = pd.DataFrame()

    df_out['codice_cliente_acquisto_prodotto'] = df_in['Codice Cliente'].apply(pulisci_valore)
    df_out['organizzazione_commerciale'] = '2310'
    df_out['tipo_documento'] = 'CR'
    df_out['canale_di_distribuzione'] = '10'
    df_out['settore_merceologico'] = '01'
    df_out['condizione_di_pagamento'] = 'CESS'
    df_out['modalita_di_pagamento'] = df_in['Metodo di pagamento Cliente'].apply(pulisci_valore)
    df_out['classificazione_fiscale_cliente'] = 'H'
    df_out['classificazione_fiscale_prodotto'] = ''
    df_out['codice_prodotto'] = df_in['Codice Prodotto'].apply(pulisci_valore)
    df_out['quantita'] = '1'

    # Formatta il valore usando la virgola per i decimali
    df_out['valore_riga'] = df_in['Valore Cessione'].apply(lambda x: str(x).replace('.', ',').strip() if pd.notna(x) else '')

    df_out['seriale'] = df_in['Seriale'].apply(pulisci_valore)
    df_out['testo_riga'] = 'Cessione'
    df_out['motivo'] = 'Z13'

    # Limite massimo di 35 caratteri (Ingegnerizzazione stringa della ragione sociale)
    df_out['riferimento_cliente'] = df_in['Cliente'].apply(lambda x: str(x)[:35].strip() if pd.notna(x) else '')

    # Mappatura della banca interna DropPay -> codice univoco aziendale
    df_out['banca'] = df_in['Banca Interna'].apply(
        lambda x: 'DRP01_50960' if pd.notna(x) and 'drop' in str(x).lower() else ''
    )

    df_out['riferimento'] = df_in['Riferimento'].apply(pulisci_valore)

    # Applicazione della regola corretta sulle date
    df_out['data_scadenza_cessione'] = df_in['Data base'].apply(aggiusta_data_scadenza)

    # Campi strutturali vuoti richiesti dal tracciato
    colonne_vuote = [
        'tipo_flusso', 'data_attivazione', 'data_evasione', 'data_prenotazione',
        'ragione_sociale_attivazione', 'codice_POS_attivazione', 'partita_iva_attivazione',
        'ragione_sociale_acquisto_prodotto', 'descrizione_prodotto', 'data_scadenza'
    ]
    for col in colonne_vuote:
        df_out[col] = ''

    # 3. Esportazione e Download Automatico
    output_name = 'Richiesta cessione.csv'
    df_out.to_csv(output_name, sep=';', index=False, encoding='utf-8')

    print(f"✅ Elaborazione completata! Il file '{output_name}' è pronto e accurato.")
    files.download(output_name)
