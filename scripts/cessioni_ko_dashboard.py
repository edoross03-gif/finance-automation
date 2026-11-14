"""
Nome: Cessioni Ko Dashboard
Descrizione: Genera una dashboard Excel ad alto impatto con grafici nativi e statistiche sulle cessioni di credito fallite (KO).
Categoria: Utility

Istruzioni:
■ FILE DI INPUT RICHIESTI:
  1. Report Cessioni KO (Excel o CSV):
     • Colonne attese: Dealer, Causale Anomalia, Agente, Importo (formato numero o valuta).

■ MODALITÀ D'USO:
  1. Carica il file Excel o CSV delle Cessioni KO.
  2. Selezionalo nel dropdown.
  3. Clicca su Esegui per consolidare i dati ed elaborare grafici e classifiche senza filtri distruttivi.
"""


# --- Latest Version Code Cell ---
#@title DASHBOARD CESSIONI KO V3

# Installa xlsxwriter nel caso non fosse presente nell'ambiente Colab
# !pip install xlsxwriter -q

import pandas as pd
import io
import os
from google.colab import files

# ==========================================
# 1. CARICAMENTO E CONSOLIDAMENTO DEI FILE
# ==========================================
print("👇 CLICCA SUL BOTTONE SOTTO PER SELEZIONARE I FILE CSV/EXCEL (ANCHE PIÙ DI UNO):")
uploaded = files.upload()

dfs = []
for filename, content in uploaded.items():
    print(f"🔄 Lettura ed elaborazione di: {filename}...")

    # Rilevamento automatico del formato (CSV o Excel reale)
    if filename.lower().endswith('.csv'):
        # sep=None e engine='python' indovinano automaticamente il delimitatore
        df = pd.read_csv(io.BytesIO(content), sep=None, engine='python')
    else:
        df = pd.read_excel(io.BytesIO(content))
    dfs.append(df)

if not dfs:
    raise ValueError("❌ Nessun file caricato! Riavvia la cella e seleziona i file.")

# Fusione di tutti i file in un unico database
df_master = pd.concat(dfs, ignore_index=True)

# --- PULIZIA E CONVERSIONE DATI ---
# (Rimozione duplicati disabilitata come da richiesta dell'utente per conservare l'integrità dei dati originari)

# Funzione robusta per pulire e convertire l'importo (gestisce il formato italiano: 1.229,91 o 139,91)
def clean_importo(val):
    if pd.isna(val):
        return 0.0
    val_str = str(val).strip()
    if not val_str:
        return 0.0
    # Se ci sono sia punti che virgole (es: 1.229,91)
    if '.' in val_str and ',' in val_str:
        val_str = val_str.replace('.', '').replace(',', '.')
    # Se c'è solo la virgola (es: 139,91 o 1229,91)
    elif ',' in val_str:
        val_str = val_str.replace(',', '.')
    try:
        return float(val_str)
    except ValueError:
        return 0.0

if 'Importo' in df_master.columns:
    df_master['Importo'] = df_master['Importo'].apply(clean_importo)
else:
    df_master['Importo'] = 0.0
    print("⚠️ Colonna 'Importo' non trovata! Creata colonna fittizia con valore 0.0")

# Funzione per pulire stringhe numeriche (es. IMEI e CF/PI) rimuovendo .0 e notazione scientifica
def clean_numeric_string(val):
    if pd.isna(val):
        return ""
    if isinstance(val, (int, float)):
        if pd.isna(val):
            return ""
        if val == int(val):
            return str(int(val))
        else:
            return f"{val:.0f}"
    val_str = str(val).strip()
    if val_str.lower() == 'nan':
        return ""
    if val_str.endswith('.0'):
        prefix = val_str[:-2]
        if prefix.isdigit():
            val_str = prefix
    return val_str

# Standardizzazione testi (rimozione spazi extra e tutto in MAIUSCOLO per raggruppamenti perfetti)
for col in ['Dealer', 'Causale Anomalia', 'Agente']:
    if col in df_master.columns:
        df_master[col] = df_master[col].astype(str).str.strip().str.upper()

# Trova i nomi reali delle colonne (case-insensitive) per pulizia e formattazione
cf_pi_col = next((c for c in df_master.columns if c.lower() in ['cf/pi', 'cf_pi', 'codice fiscale', 'partita iva']), 'CF/PI')
imei_col = next((c for c in df_master.columns if 'imei' in c.lower()), 'IMEI/Seriale')

# Pulizia preventiva per CF/PI e IMEI/Seriale
for col in [cf_pi_col, imei_col]:
    if col in df_master.columns:
        df_master[col] = df_master[col].apply(clean_numeric_string)

# Ordinamento di df_master per mantenere gli agenti e i dealer contigui (necessario per il foglio gerarchico)
sort_cols = []
if 'Agente' in df_master.columns:
    sort_cols.append('Agente')
if 'Dealer' in df_master.columns:
    sort_cols.append('Dealer')

if sort_cols:
    df_master = df_master.sort_values(by=sort_cols).reset_index(drop=True)

# ==========================================
# 2. CALCOLO METRICHE E AGGREGRAZIONI
# ==========================================
tot_ko = len(df_master)
tot_dealer = df_master['Dealer'].nunique() if 'Dealer' in df_master.columns else 0
tot_agenti = df_master['Agente'].nunique() if 'Agente' in df_master.columns else 0
tot_importo = df_master['Importo'].sum()
medio_importo = df_master['Importo'].mean() if tot_ko > 0 else 0

# Generazione classifiche basate sull'IMPATTO ECONOMICO (Importo)
if 'Dealer' in df_master.columns:
    top_dealer = df_master.groupby('Dealer').agg(
        Volume=('Dealer', 'count'),
        Importo=('Importo', 'sum')
    ).sort_values(by='Importo', ascending=False).head(10).reset_index()
else:
    top_dealer = pd.DataFrame(columns=['Dealer', 'Volume', 'Importo'])

if 'Causale Anomalia' in df_master.columns:
    top_anomalie = df_master.groupby('Causale Anomalia').agg(
        Frequenza=('Causale Anomalia', 'count'),
        Importo=('Importo', 'sum')
    ).sort_values(by='Importo', ascending=False).head(10).reset_index()
else:
    top_anomalie = pd.DataFrame(columns=['Causale Anomalia', 'Frequenza', 'Importo'])

if 'Agente' in df_master.columns:
    top_agenti = df_master.groupby('Agente').agg(
        Volume=('Agente', 'count'),
        Importo=('Importo', 'sum')
    ).sort_values(by='Importo', ascending=False).head(10).reset_index()
else:
    top_agenti = pd.DataFrame(columns=['Agente', 'Volume', 'Importo'])

# ==========================================
# 3. GENERAZIONE EXCEL CON DESIGN AVANZATO (V3)
# ==========================================
output_filename = "Dashboard_Cessioni_Credito_Advanced_V3.xlsx"
writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')
workbook = writer.book

# --- DEFINIZIONE DEGLI STILI GRAFICI ---
font_family = 'Segoe UI'

# Stile Banner Principale
title_format = workbook.add_format({
    'font_name': font_family, 'font_size': 18, 'bold': True, 'font_color': '#FFFFFF',
    'bg_color': '#1F4E78', 'align': 'center', 'valign': 'vcenter'
})

# Stili per le schede KPI
kpi_title_format = workbook.add_format({
    'font_name': font_family, 'font_size': 9, 'bold': True, 'font_color': '#5A6A85',
    'bg_color': '#F4F6F9', 'align': 'center', 'valign': 'vcenter',
    'top': 1, 'left': 1, 'right': 1, 'top_color': '#D0D5DD', 'left_color': '#D0D5DD', 'right_color': '#D0D5DD'
})
kpi_val_format = workbook.add_format({
    'font_name': font_family, 'font_size': 18, 'bold': True, 'font_color': '#1F4E78',
    'bg_color': '#F4F6F9', 'align': 'center', 'valign': 'vcenter',
    'bottom': 1, 'left': 1, 'right': 1, 'bottom_color': '#D0D5DD', 'left_color': '#D0D5DD', 'right_color': '#D0D5DD',
    'num_format': '#,##0'
})
kpi_currency_format = workbook.add_format({
    'font_name': font_family, 'font_size': 18, 'bold': True, 'font_color': '#C0392B', # Rosso scuro per evidenziare la perdita economica
    'bg_color': '#F4F6F9', 'align': 'center', 'valign': 'vcenter',
    'bottom': 1, 'left': 1, 'right': 1, 'bottom_color': '#D0D5DD', 'left_color': '#D0D5DD', 'right_color': '#D0D5DD',
    'num_format': '€ #,##0.00'
})

# Stili per le Tabelle
header_format = workbook.add_format({
    'font_name': font_family, 'font_size': 11, 'bold': True, 'font_color': '#FFFFFF',
    'bg_color': '#2C3E50', 'align': 'left', 'valign': 'vcenter', 'bottom': 2, 'bottom_color': '#1A252F'
})
cell_format = workbook.add_format({
    'font_name': font_family, 'font_size': 10, 'font_color': '#333333',
    'align': 'left', 'valign': 'vcenter', 'bottom': 1, 'bottom_color': '#EAEAEA'
})
num_cell_format = workbook.add_format({
    'font_name': font_family, 'font_size': 10, 'font_color': '#333333', 'bold': True,
    'align': 'right', 'valign': 'vcenter', 'bottom': 1, 'bottom_color': '#EAEAEA', 'num_format': '#,##0'
})
currency_format = workbook.add_format({
    'font_name': font_family, 'font_size': 10, 'font_color': '#333333', 'bold': True,
    'align': 'right', 'valign': 'vcenter', 'bottom': 1, 'bottom_color': '#EAEAEA', 'num_format': '€ #,##0.00'
})

# Nuovi formati per Partite IVA e IMEI senza virgole e senza decimali
pi_num_format = workbook.add_format({
    'font_name': font_family, 'font_size': 10, 'font_color': '#333333', 'bold': True,
    'align': 'right', 'valign': 'vcenter', 'bottom': 1, 'bottom_color': '#EAEAEA', 'num_format': '00000000000'
})
plain_num_format = workbook.add_format({
    'font_name': font_family, 'font_size': 10, 'font_color': '#333333', 'bold': True,
    'align': 'right', 'valign': 'vcenter', 'bottom': 1, 'bottom_color': '#EAEAEA', 'num_format': '0'
})

# Stili per la struttura gerarchica ad albero
agent_tree_format = workbook.add_format({
    'font_name': font_family, 'font_size': 11, 'bold': True, 'font_color': '#1F4E78',
    'bg_color': '#D9E1F2', 'align': 'left', 'valign': 'vcenter', 'bottom': 1, 'bottom_color': '#B4C6E7'
})
agent_tree_val_format = workbook.add_format({
    'font_name': font_family, 'font_size': 11, 'bold': True, 'font_color': '#1F4E78',
    'bg_color': '#D9E1F2', 'align': 'right', 'valign': 'vcenter', 'bottom': 1, 'bottom_color': '#B4C6E7', 'num_format': '#,##0'
})
agent_tree_currency_format = workbook.add_format({
    'font_name': font_family, 'font_size': 11, 'bold': True, 'font_color': '#1F4E78',
    'bg_color': '#D9E1F2', 'align': 'right', 'valign': 'vcenter', 'bottom': 1, 'bottom_color': '#B4C6E7', 'num_format': '€ #,##0.00'
})

dealer_tree_format = workbook.add_format({
    'font_name': font_family, 'font_size': 10, 'bold': True, 'font_color': '#333333',
    'bg_color': '#F2F2F2', 'align': 'left', 'valign': 'vcenter', 'bottom': 1, 'bottom_color': '#D9D9D9'
})
dealer_tree_val_format = workbook.add_format({
    'font_name': font_family, 'font_size': 10, 'bold': True, 'font_color': '#333333',
    'bg_color': '#F2F2F2', 'align': 'right', 'valign': 'vcenter', 'bottom': 1, 'bottom_color': '#D9D9D9', 'num_format': '#,##0'
})
dealer_tree_currency_format = workbook.add_format({
    'font_name': font_family, 'font_size': 10, 'bold': True, 'font_color': '#333333',
    'bg_color': '#F2F2F2', 'align': 'right', 'valign': 'vcenter', 'bottom': 1, 'bottom_color': '#D9D9D9', 'num_format': '€ #,##0.00'
})

# ------------------------------------------
# TAB 1: DASHBOARD VISIVA (Senza griglie Excel)
# ------------------------------------------
ws_dash = workbook.add_worksheet('Dashboard')
ws_dash.hide_gridlines(2)  # Nasconde la griglia per un look pulito

# Titolo ad alto impatto (allargato a colonna T per coprire la nuova larghezza della dashboard)
ws_dash.merge_range('B2:T3', 'PERFORMANCE & ANALISI FINANZIARIA: CESSIONI CREDITO KO', title_format)
ws_dash.set_row(1, 20)
ws_dash.set_row(2, 20)

# Creazione dei blocchi KPI (Ora include anche metriche sugli importi!)
ws_dash.merge_range('B5:D5', 'VOLUME TOTALE KO (PRATICHE)', kpi_title_format)
ws_dash.merge_range('B6:D7', tot_ko, kpi_val_format)

ws_dash.merge_range('F5:H5', 'VALORE TOTALE KO (€)', kpi_title_format)
ws_dash.merge_range('F6:H7', tot_importo, kpi_currency_format)

ws_dash.merge_range('J5:L5', 'VALORE MEDIO KO (€)', kpi_title_format)
ws_dash.merge_range('J6:L7', medio_importo, kpi_currency_format)

ws_dash.merge_range('N5:P5', 'TOTAL DEALER COINVOLTI', kpi_title_format)
ws_dash.merge_range('N6:P7', tot_dealer, kpi_val_format)

ws_dash.merge_range('R5:T5', 'TOTAL AGENTI IMPATTATI', kpi_title_format)
ws_dash.merge_range('R6:T7', tot_agenti, kpi_val_format)

ws_dash.set_row(4, 15)
ws_dash.set_row(5, 18)
ws_dash.set_row(6, 18)

# Larghezza colonne fisse per allineamento elementi dashboard (allineata fino alla colonna T)
for col_letter, width in [
    ('A', 3), ('B', 15), ('C', 15), ('D', 15), ('E', 4),
    ('F', 15), ('G', 15), ('H', 15), ('I', 5), ('J', 15),
    ('K', 15), ('L', 15), ('M', 4), ('N', 15), ('O', 15),
    ('P', 15), ('Q', 4), ('R', 15), ('S', 15), ('T', 15)
]:
    ws_dash.set_column(f'{col_letter}:{col_letter}', width)

# ------------------------------------------
# TAB 2: DATI AGGREGATI (Supporto per i Grafici)
# ------------------------------------------
ws_data_aggr = workbook.add_worksheet('Analisi_Aggregata')
ws_data_aggr.hide_gridlines(2)

# Popolamento tabelle di riepilogo con l'aggiunta delle colonne importo
ws_data_aggr.write('A1', 'Dealer', header_format)
ws_data_aggr.write('B1', 'Volume KO', header_format)
ws_data_aggr.write('C1', 'Importo KO', header_format)
for idx, row in top_dealer.iterrows():
    ws_data_aggr.write(idx+1, 0, row['Dealer'], cell_format)
    ws_data_aggr.write(idx+1, 1, row['Volume'], num_cell_format)
    ws_data_aggr.write(idx+1, 2, row['Importo'], currency_format)

ws_data_aggr.write('E1', 'Causale Anomalia', header_format)
ws_data_aggr.write('F1', 'Frequenza', header_format)
ws_data_aggr.write('G1', 'Importo KO', header_format)
for idx, row in top_anomalie.iterrows():
    ws_data_aggr.write(idx+1, 4, row['Causale Anomalia'], cell_format)
    ws_data_aggr.write(idx+1, 5, row['Frequenza'], num_cell_format)
    ws_data_aggr.write(idx+1, 6, row['Importo'], currency_format)

ws_data_aggr.write('I1', 'Agente', header_format)
ws_data_aggr.write('J1', 'Volume KO', header_format)
ws_data_aggr.write('K1', 'Importo KO', header_format)
for idx, row in top_agenti.iterrows():
    ws_data_aggr.write(idx+1, 8, row['Agente'], cell_format)
    ws_data_aggr.write(idx+1, 9, row['Volume'], num_cell_format)
    ws_data_aggr.write(idx+1, 10, row['Importo'], currency_format)

# Dimensionamento colonne tecniche dell'analisi aggregata
for i, width in enumerate([35, 12, 16, 5, 50, 12, 16, 5, 25, 12, 16]):
    ws_data_aggr.set_column(i, i, width)

# ------------------------------------------
# COSTRUZIONE DEI GRAFICI NATIVI EXCEL (V3 orientata al Valore)
# ------------------------------------------

# 1. Grafico a Barre Orizzontali (Top Dealer per Valore)
chart_dealer = workbook.add_chart({'type': 'bar'})
chart_dealer.add_series({
    'name': '=Analisi_Aggregata!$C$1',
    'categories': '=Analisi_Aggregata!$A$2:$A$11',
    'values': '=Analisi_Aggregata!$C$2:$C$11',
    'fill': {'color': '#1F4E78'},
    'border': {'none': True}
})
chart_dealer.set_title({'name': 'Top 10 Dealer per Impatto Economico KO', 'name_font': {'name': font_family, 'size': 11, 'bold': True, 'color': '#2C3E50'}})
chart_dealer.set_legend({'none': True})
chart_dealer.set_x_axis({
    'major_gridlines': {'visible': True, 'line': {'color': '#EAEAEA', 'dash_type': 'dash'}},
    'num_format': '€ #,##0'
})
chart_dealer.set_y_axis({'reverse': True, 'line': {'none': True}})
chart_dealer.set_size({'width': 540, 'height': 320})
ws_dash.insert_chart('B10', chart_dealer)

# 2. Grafico a Ciambella (Top 5 Cause di Scarto per Valore)
chart_causali = workbook.add_chart({'type': 'doughnut'})
chart_causali.add_series({
    'name': 'Valore Anomalia',
    'categories': '=Analisi_Aggregata!$E$2:$E$6',
    'values': '=Analisi_Aggregata!$G$2:$G$6',
    'points': [
        {'fill': {'color': '#1F4E78'}},
        {'fill': {'color': '#2E75B6'}},
        {'fill': {'color': '#BDD7EE'}},
        {'fill': {'color': '#A6A6A6'}},
        {'fill': {'color': '#D9D9D9'}},
    ]
})
chart_causali.set_title({'name': 'Top 5 Cause per Valore Perdite KO', 'name_font': {'name': font_family, 'size': 11, 'bold': True, 'color': '#2C3E50'}})
chart_causali.set_legend({'position': 'right', 'font': {'name': font_family, 'size': 9}})
chart_causali.set_size({'width': 580, 'height': 320})
ws_dash.insert_chart('L10', chart_causali)

# 3. Grafico a Istogramma Verticale (Top Agenti per Valore)
chart_agenti = workbook.add_chart({'type': 'column'})
chart_agenti.add_series({
    'name': '=Analisi_Aggregata!$K$1',
    'categories': '=Analisi_Aggregata!$I$2:$I$11',
    'values': '=Analisi_Aggregata!$K$2:$K$11',
    'fill': {'color': '#2E75B6'},
    'border': {'none': True}
})
chart_agenti.set_title({'name': 'Classifica Top 10 Agenti (Impatto Economico KO)', 'name_font': {'name': font_family, 'size': 11, 'bold': True, 'color': '#2C3E50'}})
chart_agenti.set_legend({'none': True})
chart_agenti.set_y_axis({
    'major_gridlines': {'visible': True, 'line': {'color': '#EAEAEA', 'dash_type': 'dash'}},
    'num_format': '€ #,##0'
})
chart_agenti.set_size({'width': 1160, 'height': 300})
ws_dash.insert_chart('B27', chart_agenti)

# ------------------------------------------
# TAB 3: DETTAGLIO GERARCHICO AD ALBERO COLLASSABILE
# ------------------------------------------
ws_tree = workbook.add_worksheet('Dettaglio_Gerarchico')
ws_tree.hide_gridlines(2)
ws_tree.outline_settings(visible=True, symbols_below=False, symbols_right=False, auto_style=True)

# Scrittura intestazioni
headers_tree = [
    'Struttura (Agente / Rivenditore)',
    'Data Ordine',
    'Cliente',
    'CF/PI',
    'Causale Anomalia',
    'IMEI/Seriale',
    'Importo KO'
]
for col_idx, header in enumerate(headers_tree):
    ws_tree.write(0, col_idx, header, header_format)

def format_cliente(row):
    nome = str(row.get('Nome/ Rag. Sociale', '')).strip()
    cognome = str(row.get('Cognome', '')).strip()
    if nome.lower() in ['nan', '']:
        nome = ""
    if cognome.lower() in ['nan', '']:
        cognome = ""
    return f"{nome} {cognome}".strip()

# Popolamento gerarchico
agenti_unici = sorted(df_master['Agente'].unique()) if 'Agente' in df_master.columns else []
row_idx_tree = 1

for ag in agenti_unici:
    df_ag = df_master[df_master['Agente'] == ag]
    ag_vol = len(df_ag)
    ag_imp = df_ag['Importo'].sum()
    
    ws_tree.write(row_idx_tree, 0, f"AGENTE: {ag}", agent_tree_format)
    ws_tree.write(row_idx_tree, 1, "", agent_tree_format)
    ws_tree.write(row_idx_tree, 2, "", agent_tree_format)
    ws_tree.write(row_idx_tree, 3, "", agent_tree_format)
    ws_tree.write(row_idx_tree, 4, "", agent_tree_format)
    ws_tree.write(row_idx_tree, 5, f"Volume: {ag_vol}", agent_tree_val_format)
    ws_tree.write(row_idx_tree, 6, ag_imp, agent_tree_currency_format)
    ws_tree.set_row(row_idx_tree, 20, None, {'level': 0})
    row_idx_tree += 1
    
    dealers_ag = sorted(df_ag['Dealer'].unique()) if 'Dealer' in df_ag.columns else []
    for dl in dealers_ag:
        df_dl = df_ag[df_ag['Dealer'] == dl]
        dl_vol = len(df_dl)
        dl_imp = df_dl['Importo'].sum()
        
        ws_tree.write(row_idx_tree, 0, f"   RIVENDITORE: {dl}", dealer_tree_format)
        ws_tree.write(row_idx_tree, 1, "", dealer_tree_format)
        ws_tree.write(row_idx_tree, 2, "", dealer_tree_format)
        ws_tree.write(row_idx_tree, 3, "", dealer_tree_format)
        ws_tree.write(row_idx_tree, 4, "", dealer_tree_format)
        ws_tree.write(row_idx_tree, 5, f"Volume: {dl_vol}", dealer_tree_val_format)
        ws_tree.write(row_idx_tree, 6, dl_imp, dealer_tree_currency_format)
        ws_tree.set_row(row_idx_tree, 18, None, {'level': 1, 'hidden': True})
        row_idx_tree += 1
        
        for _, detail_row in df_dl.iterrows():
            ws_tree.write(row_idx_tree, 0, "", cell_format)
            
            data_ord = str(detail_row.get('Data Ordine', '')).strip()
            if data_ord.lower() in ['nan', '']:
                data_ord = ""
            ws_tree.write(row_idx_tree, 1, data_ord, cell_format)
            
            cliente = format_cliente(detail_row)
            ws_tree.write(row_idx_tree, 2, cliente, cell_format)
            
            cf_pi_val = detail_row.get(cf_pi_col, '')
            if pd.isna(cf_pi_val) or cf_pi_val == "":
                ws_tree.write(row_idx_tree, 3, "", cell_format)
            elif str(cf_pi_val).isdigit():
                ws_tree.write_number(row_idx_tree, 3, int(cf_pi_val), pi_num_format)
            else:
                ws_tree.write(row_idx_tree, 3, str(cf_pi_val), cell_format)
                
            causale = str(detail_row.get('Causale Anomalia', '')).strip()
            if causale.lower() in ['nan', '']:
                causale = ""
            ws_tree.write(row_idx_tree, 4, causale, cell_format)
            
            imei_val = detail_row.get(imei_col, '')
            if pd.isna(imei_val) or imei_val == "":
                ws_tree.write(row_idx_tree, 5, "", cell_format)
            elif str(imei_val).isdigit():
                ws_tree.write_number(row_idx_tree, 5, int(imei_val), plain_num_format)
            else:
                ws_tree.write(row_idx_tree, 5, str(imei_val), cell_format)
                
            importo_val = detail_row.get('Importo', 0.0)
            ws_tree.write(row_idx_tree, 6, importo_val, currency_format)
            
            ws_tree.set_row(row_idx_tree, 17, None, {'level': 2, 'hidden': True})
            row_idx_tree += 1

for col_idx, width in enumerate([40, 15, 30, 16, 50, 18, 16]):
    ws_tree.set_column(col_idx, col_idx, width)

# ------------------------------------------
# TAB 4: DATI CONSOLIDATI MASTER (Con formattazione delle colonne numeriche)
# ------------------------------------------
df_master.to_excel(writer, sheet_name='Dati_Consolidati', index=False)
ws_master = writer.sheets['Dati_Consolidati']

# Formattazione Intestazioni Master
for col_num, header in enumerate(df_master.columns):
    ws_master.write(0, col_num, header, header_format)

# Formattazione righe del Database Master (Auto-fit intelligente e applicazione degli stili)
for col_num, col_name in enumerate(df_master.columns):
    max_len = max(df_master[col_name].astype(str).str.len().max(), len(col_name)) + 4
    ws_master.set_column(col_num, col_num, min(max_len, 45))

    # Identificatori colonne
    is_importo_col = (col_name.lower() == 'importo')
    is_cf_pi_col = (col_name == cf_pi_col)
    is_imei_col = (col_name == imei_col)

    for row_num in range(1, len(df_master) + 1):
        val = df_master.iloc[row_num-1, col_num]
        if pd.isna(val) or val == "":
            ws_master.write(row_num, col_num, "", cell_format)
        elif is_importo_col:
            ws_master.write(row_num, col_num, val, currency_format)
        elif is_cf_pi_col:
            val_str = str(val).strip()
            if val_str.isdigit():
                ws_master.write_number(row_num, col_num, int(val_str), pi_num_format)
            else:
                ws_master.write(row_num, col_num, val_str, cell_format)
        elif is_imei_col:
            val_str = str(val).strip()
            if val_str.isdigit():
                ws_master.write_number(row_num, col_num, int(val_str), plain_num_format)
            else:
                ws_master.write(row_num, col_num, val_str, cell_format)
        elif isinstance(val, (int, float)):
            ws_master.write(row_num, col_num, val, num_cell_format)
        else:
            ws_master.write(row_num, col_num, str(val), cell_format)

# Salvataggio e download
writer.close()
print("\n🎉 REPORT ELABORATO CON SUCCESSO!")
files.download(output_filename)
