"""
Nome: Elaborazione Aging
Descrizione: Calcola l'invecchiamento dei crediti SAP incrociando i saldi con il bilancio di verifica.
Categoria: Crediti

Istruzioni:
■ FILE DI INPUT RICHIESTI:
  1. File delle POSIZIONI (es. Posizioni_-_2026-07-15.xlsx)
     • Colonne attese: Cliente, Nome del cliente, Data doc. contabile, Tipo doc. contabile, Importo (div. soc.), Scad. importo netto.
  2. File del BILANCIO (es. Bilancio_Giugno2.xlsx)
     • Struttura: Prima colonna con codici conto (es. YCOA00121000000), terza colonna con il saldo del periodo.

■ MODALITÀ D'USO:
  1. Carica entrambi i file nella sezione di sinistra.
  2. Seleziona i file corrispondenti nei dropdown del modulo.
  3. Imposta la data di riferimento per il calcolo dell'aging.
  4. Clicca su Esegui per avviare l'elaborazione.
"""
# @title ELABORAZIONE AGING ⏳
import pandas as pd
import json
import datetime
import os
import sys
from pathlib import Path
from google.colab import files
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ==========================================
# 1. CARICAMENTO AUTOMATICO DEI FILE
# ==========================================
print("Carica il file delle POSIZIONI (es. Posizioni__48_.xlsx):")
caricato_sap = files.upload()
nome_file_sap = list(caricato_sap.keys())[0]
sap = pd.read_excel(nome_file_sap)

# Pulizia Scad. importo netto
sap['Scad. importo netto'] = sap['Scad. importo netto'].apply(
    lambda x: pd.NaT if isinstance(x, datetime.time) else x
)
sap['Scad. importo netto'] = pd.to_datetime(sap['Scad. importo netto'], errors='coerce')

print("\nCarica il file del BILANCIO (es. Bilancio_fino_ad_aprile.xlsx):")
caricato_bil = files.upload()
nome_file_bil = list(caricato_bil.keys())[0]
bil = pd.read_excel(nome_file_bil)
print("\nFile caricati!")

# Richiesta dinamica della data di riferimento (con supporto a variabile d'ambiente per automazione)
env_ref_date = os.environ.get("REF_DATE")
if env_ref_date:
    try:
        REF_DATE = pd.to_datetime(env_ref_date)
        print(f"✅ Data impostata da variabile d'ambiente: {REF_DATE.strftime('%d/%m/%Y')}")
        print("Avvio l'elaborazione...\n")
    except Exception as e:
        print(f"❌ Errore conversione data da ambiente '{env_ref_date}': {e}. Passaggio a richiesta interattiva.")
        env_ref_date = None

if not env_ref_date:
    while True:
        data_input = input("\n📅 Inserisci la Data di Riferimento per l'Aging (formato AAAA-MM-GG, es. 2026-06-30): ")
        try:
            # Tenta di convertire l'input in una data valida
            REF_DATE = pd.to_datetime(data_input)
            print(f"✅ Data impostata correttamente: {REF_DATE.strftime('%d/%m/%Y')}")
            print("Avvio l'elaborazione...\n")
            break # Esce dal ciclo se la data è corretta
        except Exception:
            print("❌ Errore: Formato non valido. Assicurati di usare l'anno a 4 cifre col trattino (es. 2026-06-30). Riprova.")

# Pulizia Data doc. contabile + calcolo doc oltre 365 gg
sap['Data doc. contabile'] = sap['Data doc. contabile'].apply(
    lambda x: pd.NaT if isinstance(x, datetime.time) else x
)
sap['Data doc. contabile'] = pd.to_datetime(sap['Data doc. contabile'], errors='coerce')

mask_oltre365 = (REF_DATE - sap['Data doc. contabile']).dt.days > 365
doc_oltre365 = float(sap[mask_oltre365]['Importo (div. soc.)'].sum())
n_doc_oltre365 = int(mask_oltre365.sum())
print(f"Documenti emessi oltre 365 gg: {n_doc_oltre365} righe, totale € {doc_oltre365:,.2f}")

# ==========================================
# 2. ELABORAZIONE DATI
# ==========================================
STANDARD_TYPES = ['CD','DZ','SU','W1','W4','W7','WP','ZD','ZV']
BUCKETS_ORDER  = ['0-30','31-60','61-90','91-120','121-180','181-360','>360','A scadere']

def get_bucket(due_date, ref):
    if pd.isna(due_date): return None
    if isinstance(due_date, datetime.time): return None
    if due_date > ref: return 'A scadere'
    m = (ref.year - due_date.year)*12 + (ref.month - due_date.month)
    if m == 0: return '0-30'
    if m == 1: return '31-60'
    if m == 2: return '61-90'
    if m == 3: return '91-120'
    if m in (4,5): return '121-180'
    if 6 <= m <= 11: return '181-360'
    return '>360'

sap['Bucket'] = sap['Scad. importo netto'].apply(lambda d: get_bucket(d, REF_DATE))

MESI = {1:'Gennaio',2:'Febbraio',3:'Marzo',4:'Aprile',5:'Maggio',6:'Giugno',
        7:'Luglio',8:'Agosto',9:'Settembre',10:'Ottobre',11:'Novembre',12:'Dicembre'}
def m_off(ref, k):
    y, m = ref.year, ref.month - k
    while m <= 0: m += 12; y -= 1
    return MESI[m], str(y)[-2:]

def lbl(b):
    if b=='0-30':    n,y=m_off(REF_DATE,0);  return f"0-30 gg  ({n} {y})"
    if b=='31-60':   n,y=m_off(REF_DATE,1);  return f"31-60 gg  ({n} {y})"
    if b=='61-90':   n,y=m_off(REF_DATE,2);  return f"61-90 gg  ({n} {y})"
    if b=='91-120':  n,y=m_off(REF_DATE,3);  return f"91-120 gg  ({n} {y})"
    if b=='121-180':
        n1,y1=m_off(REF_DATE,4); n2,y2=m_off(REF_DATE,5)
        return f"121-180 gg  ({n2} {y2} / {n1} {y1})"
    if b=='181-360':
        n1,y1=m_off(REF_DATE,6); n2,y2=m_off(REF_DATE,11)
        return f"181-360 gg  ({n2} {y2} → {n1} {y1})"
    if b=='>360':    n,y=m_off(REF_DATE,12); return f">360 gg  (ante {n} {y})"
    if b=='A scadere':
        n,y=m_off(REF_DATE,-1); return f"A scadere  (da {n} {y})"

# Pivot aging principale
main = sap[sap['Tipo doc. contabile'].isin(STANDARD_TYPES)]
M = main.pivot_table(index='Bucket', columns='Tipo doc. contabile',
                     values='Importo (div. soc.)', aggfunc='sum', fill_value=0)
M = M.reindex(BUCKETS_ORDER, fill_value=0).reindex(columns=STANDARD_TYPES, fill_value=0)
M['Totale'] = M.sum(axis=1)

# Pivot insoluti DA/AB
ins = sap[sap['Tipo doc. contabile'].isin(['DA','AB'])]
I = ins.pivot_table(index='Bucket', columns='Tipo doc. contabile',
                    values='Importo (div. soc.)', aggfunc='sum', fill_value=0)
I = I.reindex(BUCKETS_ORDER, fill_value=0)
for c in ['DA','AB']:
    if c not in I.columns: I[c] = 0
I = I[['DA','AB']]
I['Totale'] = I.sum(axis=1)

# Valori bilancio
def bil_val(prefix):
    m = bil.iloc[:,0].astype(str).str.startswith(prefix)
    return float(bil.loc[m].iloc[0,2]) if m.any() else 0.0

cred_interno = bil_val('YCOA00121000000')
cred_estero  = bil_val('YCOA00121200000')
fondo_gen    = bil_val('YCOA00121030000')
fondo_spec   = bil_val('YCOA00121030100')
fondo_legali = bil_val('YCOA00121040000')

cw_tot  = float(sap[sap['Tipo doc. contabile']=='CW']['Importo (div. soc.)'].sum())
da_tot  = float(sap[sap['Tipo doc. contabile']=='DA']['Importo (div. soc.)'].sum())
ab_tot  = float(sap[sap['Tipo doc. contabile']=='AB']['Importo (div. soc.)'].sum())
floa_nc = float(sap[sap['Cliente'].astype(str)=='1000535']['Importo (div. soc.)'].sum())
floa_ft = float(sap[sap['Cliente'].astype(str)=='1000530']['Importo (div. soc.)'].sum())
saldo_part = float(sap['Importo (div. soc.)'].sum())

# ==========================================
# 3. COSTRUZIONE EXCEL
# ==========================================
NAVY      = '1E293B'
SLATE     = '475569'
GRAY_BG   = 'F8FAFC'
GRAY_HD   = 'E2E8F0'
TXT_DARK  = '0F172A'
TXT_LIGHT = 'FFFFFF'
GREEN     = '059669'
RED       = 'DC2626'
AMBER     = 'F59E0B'

RAMP = {
    '0-30':'FEF9C3','31-60':'FEF08A','61-90':'FDBA74','91-120':'FB923C',
    '121-180':'F97316','181-360':'EA580C','>360':'B91C1C','A scadere':'D1FAE5',
}
RAMP_TEXT = {
    '0-30':TXT_DARK,'31-60':TXT_DARK,'61-90':TXT_DARK,'91-120':TXT_DARK,
    '121-180':TXT_DARK,'181-360':TXT_LIGHT,'>360':TXT_LIGHT,'A scadere':TXT_DARK
}
pale = {
    '0-30':'FEFCE8','31-60':'FEF9C3','61-90':'FFEDD5','91-120':'FED7AA',
    '121-180':'FFD8B5','181-360':'FFCC9E','>360':'FCA5A5','A scadere':'ECFDF5'
}

FONT_NAME = 'Calibri'
def font(**kw): return Font(name=FONT_NAME, **kw)
thin   = Side(style='thin',   color='CBD5E1')
medium = Side(style='medium', color='475569')
border_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
def fill(hex_color): return PatternFill('solid', start_color=hex_color, end_color=hex_color)
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT   = Alignment(horizontal='left',   vertical='center', wrap_text=True, indent=1)
RIGHT  = Alignment(horizontal='right',  vertical='center', indent=1)
NUM_FMT        = '#,##0.00;[Red](#,##0.00);"–"'
NUM_FMT_NO_DEC = '#,##0;[Red](#,##0);"–"'

wb = Workbook()
ws = wb.active
ws.title = 'Dashboard'
ws.sheet_view.showGridLines = False
ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
ws.page_setup.paperSize   = ws.PAPERSIZE_A3
ws.page_setup.fitToWidth  = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True

ws.column_dimensions['A'].width = 2
ws.column_dimensions['B'].width = 36
ws.column_dimensions['C'].width = 13
for col in ['D','E','F','G','H','I','J','K','L']: ws.column_dimensions[col].width = 13
ws.column_dimensions['M'].width = 16

# Intestazione
ws.row_dimensions[1].height = 8
ws.merge_cells('B2:M2')
c = ws['B2']
c.value = 'AGING ANALISI CREDITI CLIENTI'
c.font = font(size=20, bold=True, color=TXT_LIGHT); c.fill = fill(NAVY)
c.alignment = Alignment(horizontal='left', vertical='center', indent=2)
ws.row_dimensions[2].height = 38

ws.merge_cells('B3:M3')
c = ws['B3']
c.value = f"Data di riferimento: {REF_DATE.strftime('%d/%m/%Y')}   ·   Fonte: Esportazione SAP + Bilancio CoGe"
c.font = font(size=10, italic=True, color=TXT_LIGHT); c.fill = fill(SLATE)
c.alignment = Alignment(horizontal='left', vertical='center', indent=2)
ws.row_dimensions[3].height = 22
ws.row_dimensions[4].height = 14

# KPI totali
totale_aging_main  = float(M['Totale'].sum())
totale_scaduto     = float(M.loc[[b for b in BUCKETS_ORDER if b!='A scadere'], 'Totale'].sum())
totale_a_scadere   = float(M.loc['A scadere', 'Totale'])
crediti_120        = float(M.loc[['121-180','181-360','>360'], 'Totale'].sum())
crediti_180        = float(M.loc[['181-360','>360'], 'Totale'].sum())
crediti_sofferenza = crediti_180
fondo_tot          = abs(fondo_gen + fondo_legali) + abs(fondo_spec)
delta_fondo        = fondo_tot - crediti_sofferenza
totale_insoluti    = float(I['Totale'].sum())

def kpi_tile(cell_range, label, value, accent_color):
    ws.merge_cells(cell_range)
    cell = ws[cell_range.split(':')[0]]
    cell.value = f"{label}\n€ {value:,.0f}".replace(',', '.')
    cell.font = font(size=11, bold=True, color=TXT_LIGHT)
    cell.fill = fill(accent_color)
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

kpi_tile('B5:D6', 'TOTALE CREDITI',              totale_aging_main,  NAVY)
kpi_tile('E5:F6', 'TOTALE SCADUTO',              totale_scaduto,     AMBER)
kpi_tile('G5:H6', 'CREDITI > 120 gg',            crediti_120,        'EA580C')
kpi_tile('I5:K6', 'CREDITI SOFFERENZA (>180gg)', crediti_180,        'B91C1C')
kpi_tile('L5:M6', 'INSOLUTI (DA+AB)',             totale_insoluti,    '7C3AED')
ws.row_dimensions[5].height = 22; ws.row_dimensions[6].height = 22; ws.row_dimensions[7].height = 14

# --- Tabella aging ---
HDR_ROW = 8
ws.merge_cells(f'B{HDR_ROW}:M{HDR_ROW}')
c = ws[f'B{HDR_ROW}']; c.value = 'AGING PER FASCIA E TIPO DOCUMENTO'
c.font = font(size=12, bold=True, color=TXT_LIGHT); c.fill = fill(SLATE)
c.alignment = Alignment(horizontal='left', vertical='center', indent=2)
ws.row_dimensions[HDR_ROW].height = 26

HDR2 = HDR_ROW + 1
for col in range(2, 14):
    ws.cell(row=HDR2, column=col).fill = fill(NAVY)
    ws.cell(row=HDR2, column=col).border = border_thin
ws.cell(row=HDR2, column=2, value='Fascia di scaduto').font  = font(size=10, bold=True, color=TXT_LIGHT)
ws.cell(row=HDR2, column=2).alignment = LEFT
for i, h in enumerate(STANDARD_TYPES):
    cell = ws.cell(row=HDR2, column=3+i, value=h)
    cell.font = font(size=10, bold=True, color=TXT_LIGHT); cell.alignment = CENTER
ws.cell(row=HDR2, column=12, value='TOTALE').font = font(size=10, bold=True, color=TXT_LIGHT)
ws.cell(row=HDR2, column=12).alignment = CENTER
ws.row_dimensions[HDR2].height = 28

DATA_START    = HDR2 + 1
scaduto_buckets = ['0-30','31-60','61-90','91-120','121-180','181-360','>360']

for ri, b in enumerate(scaduto_buckets):
    r = DATA_START + ri
    ws.row_dimensions[r].height = 24
    c = ws.cell(row=r, column=2, value=lbl(b))
    c.fill = fill(RAMP[b]); c.font = font(size=10, bold=True, color=RAMP_TEXT[b]); c.alignment = LEFT; c.border = border_thin
    for ci, dt in enumerate(STANDARD_TYPES):
        val = float(M.loc[b, dt])
        cell = ws.cell(row=r, column=3+ci, value=val if val != 0 else None)
        cell.number_format = NUM_FMT_NO_DEC; cell.font = font(size=10); cell.alignment = RIGHT
        cell.border = border_thin; cell.fill = fill(pale[b])
    tot_val = float(M.loc[b, 'Totale'])
    cell = ws.cell(row=r, column=12, value=tot_val)
    cell.number_format = NUM_FMT_NO_DEC; cell.font = font(size=10, bold=True, color=RAMP_TEXT[b])
    cell.alignment = RIGHT; cell.fill = fill(RAMP[b]); cell.border = border_thin

r_ts = DATA_START + 7
ws.row_dimensions[r_ts].height = 26
c = ws.cell(row=r_ts, column=2, value='TOTALE SCADUTO')
c.fill = fill(NAVY); c.font = font(size=11, bold=True, color=TXT_LIGHT); c.alignment = LEFT; c.border = border_thin
for ci, dt in enumerate(STANDARD_TYPES):
    val = float(M.loc[scaduto_buckets, dt].sum())
    cell = ws.cell(row=r_ts, column=3+ci, value=val if val != 0 else None)
    cell.number_format = NUM_FMT_NO_DEC; cell.font = font(size=10, bold=True, color=TXT_LIGHT)
    cell.alignment = RIGHT; cell.fill = fill(NAVY); cell.border = border_thin
val = float(M.loc[scaduto_buckets, 'Totale'].sum())
cell = ws.cell(row=r_ts, column=12, value=val)
cell.number_format = NUM_FMT_NO_DEC; cell.font = font(size=11, bold=True, color=TXT_LIGHT)
cell.alignment = RIGHT; cell.fill = fill(NAVY); cell.border = border_thin

r_as = DATA_START + 8
ws.row_dimensions[r_as].height = 24
b = 'A scadere'
c = ws.cell(row=r_as, column=2, value=lbl(b))
c.fill = fill(RAMP[b]); c.font = font(size=10, bold=True, color=RAMP_TEXT[b]); c.alignment = LEFT; c.border = border_thin
for ci, dt in enumerate(STANDARD_TYPES):
    val = float(M.loc[b, dt])
    cell = ws.cell(row=r_as, column=3+ci, value=val if val != 0 else None)
    cell.number_format = NUM_FMT_NO_DEC; cell.font = font(size=10); cell.alignment = RIGHT
    cell.border = border_thin; cell.fill = fill(pale[b])
tot_val = float(M.loc[b, 'Totale'])
cell = ws.cell(row=r_as, column=12, value=tot_val)
cell.number_format = NUM_FMT_NO_DEC; cell.font = font(size=10, bold=True, color=RAMP_TEXT[b])
cell.alignment = RIGHT; cell.fill = fill(RAMP[b]); cell.border = border_thin

r_tc = DATA_START + 9
ws.row_dimensions[r_tc].height = 28
c = ws.cell(row=r_tc, column=2, value='TOTALE CREDITI (aging principale)')
c.fill = fill('0F172A'); c.font = font(size=11, bold=True, color=TXT_LIGHT); c.alignment = LEFT; c.border = border_thin
for ci, dt in enumerate(STANDARD_TYPES):
    val = float(M[dt].sum())
    cell = ws.cell(row=r_tc, column=3+ci, value=val if val != 0 else None)
    cell.number_format = NUM_FMT_NO_DEC; cell.font = font(size=10, bold=True, color=TXT_LIGHT)
    cell.alignment = RIGHT; cell.fill = fill('0F172A'); cell.border = border_thin
val = float(M['Totale'].sum())
cell = ws.cell(row=r_tc, column=12, value=val)
cell.number_format = NUM_FMT_NO_DEC; cell.font = font(size=12, bold=True, color=TXT_LIGHT)
cell.alignment = RIGHT; cell.fill = fill('0F172A'); cell.border = border_thin

# --- Insoluti ---
INS_START = r_tc + 2
ws.merge_cells(f'B{INS_START}:M{INS_START}')
c = ws[f'B{INS_START}']
c.value = 'AGING INSOLUTI  (Sezione DA + AB — partite riemerse non incassate)'
c.font = font(size=12, bold=True, color=TXT_LIGHT); c.fill = fill('7C3AED')
c.alignment = Alignment(horizontal='left', vertical='center', indent=2)
ws.row_dimensions[INS_START].height = 26

INS_HDR = INS_START + 1
ws.cell(row=INS_HDR, column=2, value='Fascia di scaduto').font = font(size=10, bold=True, color=TXT_LIGHT)
ws.cell(row=INS_HDR, column=2).fill = fill(NAVY); ws.cell(row=INS_HDR, column=2).alignment = LEFT; ws.cell(row=INS_HDR, column=2).border = border_thin
for ci, h in enumerate(['DA','AB','Totale Insoluti']):
    cell = ws.cell(row=INS_HDR, column=3+ci, value=h)
    cell.font = font(size=10, bold=True, color=TXT_LIGHT); cell.fill = fill(NAVY); cell.alignment = CENTER; cell.border = border_thin
for col in range(6, 13):
    ws.cell(row=INS_HDR, column=col).fill = fill(NAVY); ws.cell(row=INS_HDR, column=col).border = border_thin
ws.row_dimensions[INS_HDR].height = 26

INS_DATA = INS_HDR + 1
ins_buckets_with_values = [b for b in BUCKETS_ORDER if abs(I.loc[b].sum()) > 0.005]
for ri, b in enumerate(ins_buckets_with_values):
    r = INS_DATA + ri
    ws.row_dimensions[r].height = 22
    c = ws.cell(row=r, column=2, value=lbl(b))
    c.fill = fill(RAMP[b]); c.font = font(size=10, bold=True, color=RAMP_TEXT[b]); c.alignment = LEFT; c.border = border_thin
    for ci, dt in enumerate(['DA','AB']):
        val = float(I.loc[b, dt])
        cell = ws.cell(row=r, column=3+ci, value=val if val != 0 else None)
        cell.number_format = NUM_FMT_NO_DEC; cell.font = font(size=10); cell.alignment = RIGHT
        cell.border = border_thin; cell.fill = fill(pale[b])
    val = float(I.loc[b, 'Totale'])
    cell = ws.cell(row=r, column=5, value=val if val != 0 else None)
    cell.number_format = NUM_FMT_NO_DEC; cell.font = font(size=10, bold=True, color=RAMP_TEXT[b])
    cell.alignment = RIGHT; cell.fill = fill(RAMP[b]); cell.border = border_thin
    for col in range(6, 13):
        ws.cell(row=r, column=col).fill = fill(pale[b]); ws.cell(row=r, column=col).border = border_thin

r = INS_DATA + len(ins_buckets_with_values)
ws.row_dimensions[r].height = 26
c = ws.cell(row=r, column=2, value='TOTALE INSOLUTI')
c.fill = fill('7C3AED'); c.font = font(size=11, bold=True, color=TXT_LIGHT); c.alignment = LEFT; c.border = border_thin
for ci, dt in enumerate(['DA','AB']):
    val = float(I[dt].sum())
    cell = ws.cell(row=r, column=3+ci, value=val if val != 0 else None)
    cell.number_format = NUM_FMT_NO_DEC; cell.font = font(size=10, bold=True, color=TXT_LIGHT)
    cell.alignment = RIGHT; cell.fill = fill('7C3AED'); cell.border = border_thin
val = float(I['Totale'].sum())
cell = ws.cell(row=r, column=5, value=val)
cell.number_format = NUM_FMT_NO_DEC; cell.font = font(size=11, bold=True, color=TXT_LIGHT)
cell.alignment = RIGHT; cell.fill = fill('7C3AED'); cell.border = border_thin
for col in range(6, 13):
    ws.cell(row=r, column=col).fill = fill('7C3AED'); ws.cell(row=r, column=col).border = border_thin
INS_END = r

# --- Riconciliazione ---
REC_START = INS_END + 2
ws.merge_cells(f'B{REC_START}:M{REC_START}')
c = ws[f'B{REC_START}']
c.value = 'RICONCILIAZIONE PARTITARIO ↔ BILANCIO'
c.font = font(size=12, bold=True, color=TXT_LIGHT); c.fill = fill(SLATE)
c.alignment = Alignment(horizontal='left', vertical='center', indent=2)
ws.row_dimensions[REC_START].height = 26

REC_HDR = REC_START + 1
ws.merge_cells(f'B{REC_HDR}:F{REC_HDR}')
c = ws[f'B{REC_HDR}']; c.value = 'Partitario SAP'
c.font = font(size=11, bold=True, color=TXT_LIGHT); c.fill = fill(NAVY)
c.alignment = Alignment(horizontal='center', vertical='center'); c.border = border_thin
ws.merge_cells(f'H{REC_HDR}:L{REC_HDR}')
c = ws[f'H{REC_HDR}']; c.value = 'Bilancio CO.GE.'
c.font = font(size=11, bold=True, color=TXT_LIGHT); c.fill = fill(NAVY)
c.alignment = Alignment(horizontal='center', vertical='center'); c.border = border_thin
ws.row_dimensions[REC_HDR].height = 24

left_items = [
    ('Aging principale (CD, DZ, SU, W1, W4, W7, WP, ZD, ZV)', totale_aging_main, False),
    ('  Cessioni Wind3 / WINDFITT1 (tipo doc CW)',              cw_tot,            False),
    ('  Insoluti (tipo doc DA + AB)',                           da_tot + ab_tot,   False),
    ('  Note di credito Floa (cli. 1000535)',                   floa_nc,           False),
    ('  Fatture Floa (cli. 1000530)',                           floa_ft,           False),
    ('TOTALE PARTITARIO',                                       saldo_part,        True),
]
right_items = [
    ('Crediti interno (c. 12100000)',            cred_interno,               False),
    ('Crediti estero  (c. 12120000)',            cred_estero,                False),
    ('Crediti gruppo  (c. 12202000)',            0,                          False),
    ('Crediti conti occasionali (c. 12100100)', 0,                          False),
    ('',                                         None,                       False),
    ('TOTALE CREDITI BILANCIO',                  cred_interno + cred_estero, True),
]

for i in range(max(len(left_items), len(right_items))):
    r = REC_HDR + 1 + i
    ws.row_dimensions[r].height = 22
    if i < len(left_items):
        label, val, is_total = left_items[i]
        ws.merge_cells(f'B{r}:E{r}')
        c = ws[f'B{r}']; c.value = label
        c.font = font(size=10, bold=is_total, color=TXT_LIGHT if is_total else TXT_DARK)
        c.fill = fill(NAVY if is_total else (GRAY_HD if i%2==0 else GRAY_BG)); c.alignment = LEFT; c.border = border_thin
        cell = ws[f'F{r}']; cell.value = val if val != 0 or is_total else None
        cell.number_format = NUM_FMT_NO_DEC
        cell.font = font(size=10, bold=is_total, color=TXT_LIGHT if is_total else TXT_DARK)
        cell.fill = fill(NAVY if is_total else (GRAY_HD if i%2==0 else GRAY_BG)); cell.alignment = RIGHT; cell.border = border_thin
    if i < len(right_items):
        label, val, is_total = right_items[i]
        ws.merge_cells(f'H{r}:K{r}')
        c = ws[f'H{r}']; c.value = label
        c.font = font(size=10, bold=is_total, color=TXT_LIGHT if is_total else TXT_DARK)
        c.fill = fill(NAVY if is_total else (GRAY_HD if i%2==0 else GRAY_BG)); c.alignment = LEFT; c.border = border_thin
        cell = ws[f'L{r}']; cell.value = val if val is not None and (val != 0 or is_total) else None
        cell.number_format = NUM_FMT_NO_DEC
        cell.font = font(size=10, bold=is_total, color=TXT_LIGHT if is_total else TXT_DARK)
        cell.fill = fill(NAVY if is_total else (GRAY_HD if i%2==0 else GRAY_BG)); cell.alignment = RIGHT; cell.border = border_thin

DELTA_ROW = REC_HDR + 1 + max(len(left_items), len(right_items)) + 1
ws.row_dimensions[DELTA_ROW].height = 30
delta_val = saldo_part - (cred_interno + cred_estero)
delta_ok  = abs(delta_val) < 1
ws.merge_cells(f'B{DELTA_ROW}:E{DELTA_ROW}')
c = ws[f'B{DELTA_ROW}']; c.value = 'DELTA  (Partitario − Bilancio)'
c.font = font(size=11, bold=True, color=TXT_LIGHT); c.fill = fill(GREEN if delta_ok else RED); c.alignment = LEFT; c.border = border_thin
cell = ws[f'F{DELTA_ROW}']; cell.value = delta_val; cell.number_format = NUM_FMT
cell.font = font(size=11, bold=True, color=TXT_LIGHT); cell.fill = fill(GREEN if delta_ok else RED); cell.alignment = RIGHT; cell.border = border_thin
ws.merge_cells(f'H{DELTA_ROW}:L{DELTA_ROW}')
c = ws[f'H{DELTA_ROW}']; c.value = '✓ QUADRATO' if delta_ok else '✗ DA VERIFICARE'
c.font = font(size=12, bold=True, color=TXT_LIGHT); c.fill = fill(GREEN if delta_ok else RED)
c.alignment = Alignment(horizontal='center', vertical='center'); c.border = border_thin

# --- Fondo svalutazione ---
FONDO_START = DELTA_ROW + 2
ws.merge_cells(f'B{FONDO_START}:M{FONDO_START}')
c = ws[f'B{FONDO_START}']; c.value = 'FONDO SVALUTAZIONE & CREDITI A RISCHIO'
c.font = font(size=12, bold=True, color=TXT_LIGHT); c.fill = fill(SLATE)
c.alignment = Alignment(horizontal='left', vertical='center', indent=2)
ws.row_dimensions[FONDO_START].height = 26

FH = FONDO_START + 1
ws.merge_cells(f'B{FH}:F{FH}')
c = ws[f'B{FH}']; c.value = 'Crediti a rischio'
c.font = font(size=11, bold=True, color=TXT_LIGHT); c.fill = fill(NAVY); c.alignment = Alignment(horizontal='center'); c.border = border_thin
ws.merge_cells(f'H{FH}:L{FH}')
c = ws[f'H{FH}']; c.value = 'Fondi disponibili'
c.font = font(size=11, bold=True, color=TXT_LIGHT); c.fill = fill(NAVY); c.alignment = Alignment(horizontal='center'); c.border = border_thin
ws.row_dimensions[FH].height = 24

wind_non_ric  = 4203.58
rischio_items = [
    ('Crediti > 120 gg',                          crediti_120,                                    False),
    ('Crediti > 180 gg',                          crediti_180,                                    False),
    ('Cessioni WindTre non riconosciute (manuale)', wind_non_ric,                                 False),
    ('CREDITI IN SOFFERENZA',                     crediti_sofferenza,                             True),
    ('Crediti non coperti da fondo specifico',    max(0, crediti_sofferenza - abs(fondo_spec)),   True),
]
fondo_items = [
    ('Fondo svalutazione generico (12103000)',    abs(fondo_gen),    False),
    ('Fondo svalut. procedure legali (12104000)', abs(fondo_legali), False),
    ('Fondo svalutazione specifico (12103010)',   abs(fondo_spec),   False),
    ('',                                          None,              False),
    ('TOTALE FONDI',                              fondo_tot,         True),
]

for i in range(max(len(rischio_items), len(fondo_items))):
    r = FH + 1 + i
    ws.row_dimensions[r].height = 22
    if i < len(rischio_items):
        label, val, is_total = rischio_items[i]
        ws.merge_cells(f'B{r}:E{r}')
        c = ws[f'B{r}']; c.value = label
        c.font = font(size=10, bold=is_total, color=TXT_LIGHT if is_total else TXT_DARK)
        c.fill = fill('B91C1C' if is_total else (GRAY_HD if i%2==0 else GRAY_BG)); c.alignment = LEFT; c.border = border_thin
        cell = ws[f'F{r}']; cell.value = val; cell.number_format = NUM_FMT_NO_DEC
        cell.font = font(size=10, bold=is_total, color=TXT_LIGHT if is_total else TXT_DARK)
        cell.fill = fill('B91C1C' if is_total else (GRAY_HD if i%2==0 else GRAY_BG)); cell.alignment = RIGHT; cell.border = border_thin
    if i < len(fondo_items):
        label, val, is_total = fondo_items[i]
        ws.merge_cells(f'H{r}:K{r}')
        c = ws[f'H{r}']; c.value = label
        c.font = font(size=10, bold=is_total, color=TXT_LIGHT if is_total else TXT_DARK)
        c.fill = fill(GREEN if is_total else (GRAY_HD if i%2==0 else GRAY_BG)); c.alignment = LEFT; c.border = border_thin
        cell = ws[f'L{r}']; cell.value = val; cell.number_format = NUM_FMT_NO_DEC
        cell.font = font(size=10, bold=is_total, color=TXT_LIGHT if is_total else TXT_DARK)
        cell.fill = fill(GREEN if is_total else (GRAY_HD if i%2==0 else GRAY_BG)); cell.alignment = RIGHT; cell.border = border_thin

DF = FH + 1 + max(len(rischio_items), len(fondo_items)) + 1
ws.row_dimensions[DF].height = 30
delta_fondo_val = fondo_tot - crediti_sofferenza
ws.merge_cells(f'B{DF}:K{DF}')
c = ws[f'B{DF}']
status = '✓ FONDI SUFFICIENTI' if delta_fondo_val >= 0 else '⚠ FONDI INCAPIENTI — ACCANTONAMENTO NECESSARIO'
c.value = 'DELTA FONDO  (Fondi − Sofferenza):  ' + status
c.font = font(size=11, bold=True, color=TXT_LIGHT); c.fill = fill(GREEN if delta_fondo_val >= 0 else RED); c.alignment = LEFT; c.border = border_thin
cell = ws[f'L{DF}']; cell.value = delta_fondo_val; cell.number_format = NUM_FMT
cell.font = font(size=12, bold=True, color=TXT_LIGHT); cell.fill = fill(GREEN if delta_fondo_val >= 0 else RED); cell.alignment = RIGHT; cell.border = border_thin

# --- Sezione: Documenti emessi oltre 365 gg ---
OLTRE365_START = DF + 2
ws.merge_cells(f'B{OLTRE365_START}:M{OLTRE365_START}')
c = ws[f'B{OLTRE365_START}']
c.value = 'DOCUMENTI EMESSI OLTRE 365 GG  (per data di emissione — tutto il partitario)'
c.font = font(size=12, bold=True, color=TXT_LIGHT); c.fill = fill(SLATE)
c.alignment = Alignment(horizontal='left', vertical='center', indent=2)
ws.row_dimensions[OLTRE365_START].height = 26

r = OLTRE365_START + 1
ws.row_dimensions[r].height = 28
ws.merge_cells(f'B{r}:K{r}')
c = ws[f'B{r}']
c.value = (f'Importo totale posizioni aperte con data emissione > 365 gg  '
           f'(inclusi CW / DA / AB)  —  n° righe: {n_doc_oltre365}')
c.font = font(size=10, color=TXT_DARK); c.fill = fill(GRAY_HD); c.alignment = LEFT; c.border = border_thin
cell = ws[f'L{r}']
cell.value = doc_oltre365; cell.number_format = NUM_FMT_NO_DEC
cell.font = font(size=12, bold=True, color=RED); cell.fill = fill(GRAY_HD); cell.alignment = RIGHT; cell.border = border_thin

# Note metodologiche
NOTE_ROW = r + 2
ws.merge_cells(f'B{NOTE_ROW}:M{NOTE_ROW}')
c = ws[f'B{NOTE_ROW}']
c.value = ("Note metodologiche: Buckets calcolate sul campo «Scad. importo netto». "
           "Le righe CW e DA/AB sono escluse dall'aging principale e gestite a parte. "
           "La sezione «Emessi >365 gg» usa invece il campo «Data doc. contabile» su tutto il partitario.")
c.font = font(size=8, italic=True, color=SLATE)
c.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
ws.row_dimensions[NOTE_ROW].height = 36

# ============ SHEET 2: Dettaglio pivot ============
ws2 = wb.create_sheet('Dettaglio Pivot')
ws2.sheet_view.showGridLines = False

pivot_full = sap[sap['Tipo doc. contabile'].isin(STANDARD_TYPES)].copy()
pivot_full['Anno'] = pivot_full['Scad. importo netto'].dt.year
pivot_full['Mese'] = pivot_full['Scad. importo netto'].dt.month
pv = pivot_full.pivot_table(
    index=['Anno','Mese'], columns='Tipo doc. contabile',
    values='Importo (div. soc.)', aggfunc='sum', fill_value=0
).reindex(columns=STANDARD_TYPES, fill_value=0)
pv['Totale'] = pv.sum(axis=1)

ws2.column_dimensions['A'].width = 2; ws2.column_dimensions['B'].width = 10; ws2.column_dimensions['C'].width = 12
for col in 'DEFGHIJKL': ws2.column_dimensions[col].width = 12
ws2.column_dimensions['M'].width = 14

ws2.merge_cells('B2:M2')
c = ws2['B2']; c.value = 'DETTAGLIO PIVOT  —  Importo per Anno · Mese × Tipo doc.'
c.font = font(size=14, bold=True, color=TXT_LIGHT); c.fill = fill(NAVY)
c.alignment = Alignment(horizontal='left', vertical='center', indent=2)
ws2.row_dimensions[2].height = 32

ws2['B4'].value = 'Anno'; ws2['C4'].value = 'Mese'
for ci, h in enumerate(STANDARD_TYPES + ['Totale']):
    cell = ws2.cell(row=4, column=4+ci, value=h)
    cell.font = font(size=10, bold=True, color=TXT_LIGHT); cell.fill = fill(NAVY); cell.alignment = CENTER; cell.border = border_thin
for col in ['B','C']:
    ws2[f'{col}4'].font = font(size=10, bold=True, color=TXT_LIGHT)
    ws2[f'{col}4'].fill = fill(NAVY); ws2[f'{col}4'].alignment = CENTER; ws2[f'{col}4'].border = border_thin
ws2.row_dimensions[4].height = 24

r = 5
prev_year = None
MESI_SHORT = {1:'gen',2:'feb',3:'mar',4:'apr',5:'mag',6:'giu',
              7:'lug',8:'ago',9:'set',10:'ott',11:'nov',12:'dic'}

for (year, month), row_data in pv.iterrows():
    bg = 'FFFFFF' if (r % 2 == 1) else 'F8FAFC'
    yr_display = year if year != prev_year else ''
    ws2.cell(row=r, column=2, value=yr_display).font = font(size=10, bold=True)
    ws2.cell(row=r, column=2).fill = fill(bg); ws2.cell(row=r, column=2).alignment = CENTER; ws2.cell(row=r, column=2).border = border_thin
    ws2.cell(row=r, column=3, value=MESI_SHORT[month]).font = font(size=10)
    ws2.cell(row=r, column=3).fill = fill(bg); ws2.cell(row=r, column=3).alignment = CENTER; ws2.cell(row=r, column=3).border = border_thin
    for ci, dt in enumerate(STANDARD_TYPES):
        val = float(row_data[dt])
        cell = ws2.cell(row=r, column=4+ci, value=val if val != 0 else None)
        cell.font = font(size=10); cell.number_format = NUM_FMT_NO_DEC; cell.alignment = RIGHT; cell.border = border_thin; cell.fill = fill(bg)
    tot_val = float(row_data['Totale'])
    cell = ws2.cell(row=r, column=13, value=tot_val if tot_val != 0 else None)
    cell.font = font(size=10, bold=True); cell.number_format = NUM_FMT_NO_DEC; cell.alignment = RIGHT; cell.border = border_thin; cell.fill = fill(bg)
    ws2.row_dimensions[r].height = 20
    prev_year = year; r += 1

ws2.row_dimensions[r].height = 26
for col in range(2, 14):
    ws2.cell(row=r, column=col).fill = fill(NAVY); ws2.cell(row=r, column=col).border = border_thin
ws2.cell(row=r, column=2, value='TOTALE').font = font(size=10, bold=True, color=TXT_LIGHT)
ws2.cell(row=r, column=2).alignment = CENTER
for ci, dt in enumerate(STANDARD_TYPES):
    val = float(pv[dt].sum())
    cell = ws2.cell(row=r, column=4+ci, value=val if val != 0 else None)
    cell.number_format = NUM_FMT_NO_DEC; cell.font = font(size=10, bold=True, color=TXT_LIGHT); cell.alignment = RIGHT
grand_total = float(pv['Totale'].sum())
cell = ws2.cell(row=r, column=13, value=grand_total)
cell.number_format = NUM_FMT_NO_DEC; cell.font = font(size=11, bold=True, color=TXT_LIGHT); cell.alignment = RIGHT

# Salva Excel
nome_excel = os.environ.get("OUT_EXCEL", 'Aging_Analisi_Crediti_Redesigned.xlsx')
wb.save(nome_excel)
print(f"File Excel '{nome_excel}' generato con successo.")

# ==========================================
# 4. DIZIONARIO DATA + DASHBOARD HTML
# ==========================================
try:
    top_clients_df = (
        sap[sap['Bucket'].isin(['121-180','181-360','>360'])]
        .groupby(['Cliente', 'Nome del cliente'])['Importo (div. soc.)']
        .sum().reset_index()
        .sort_values('Importo (div. soc.)', ascending=False)
        .head(10)
    )
    top_overdue_list = [
        {"cliente": str(r['Cliente']), "nome": str(r['Nome del cliente']), "totale": float(r['Importo (div. soc.)'])}
        for _, r in top_clients_df.iterrows()
    ]
except Exception as e:
    print(f"Nota: top clienti non calcolabili ({e}).")
    top_overdue_list = [{"cliente":"N/D","nome":"Vedi dettaglio Excel","totale":0.0}]

DATA = {
    "ref_date": REF_DATE.strftime('%Y-%m-%d'),
    "kpis": {
        "totale_crediti_aging": totale_aging_main,
        "totale_scaduto":       totale_scaduto,
        "crediti_120":          crediti_120,
        "crediti_180":          crediti_180,
        "insoluti_total":       totale_insoluti,
        "doc_oltre365":         doc_oltre365,
        "n_doc_oltre365":       n_doc_oltre365,
    },
    "buckets":       BUCKETS_ORDER,
    "doc_types":     STANDARD_TYPES,
    "main_aging":    {b: {dt: float(M.loc[b, dt]) for dt in STANDARD_TYPES} for b in BUCKETS_ORDER},
    "insoluti_aging":{b: {"DA": float(I.loc[b,"DA"]), "AB": float(I.loc[b,"AB"])} for b in BUCKETS_ORDER},
    "top_overdue":   top_overdue_list,
    "reconciliation":{
        "totale_matrice_sap":   totale_aging_main,
        "cessioni_wind_cw":     cw_tot,
        "insoluti_da_ab":       da_tot + ab_tot,
        "sospesi":              floa_nc + floa_ft,
        "totale_lordo_sap":     saldo_part,
        "conto_130100_italia":  cred_interno,
        "conto_130300_estero":  cred_estero,
        "totale_bilancio_coge": cred_interno + cred_estero,
        "delta_quadratura":     delta_val,
    },
    "fondo":{
        "scaduto_critico_180":         crediti_180,
        "insoluti_totali":             totale_insoluti,
        "partite_legali":              abs(fondo_legali),
        "massa_a_rischio_totale":      crediti_sofferenza,
        "fondo_svalutazione_coge":     fondo_tot,
        "totale_coperture_disponibili":fondo_tot,
        "delta_copertura":             delta_fondo_val,
    }
}

DATA_JS = json.dumps(DATA, ensure_ascii=False)

HTML = r"""<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="UTF-8" />
<title>Aging Crediti — Dashboard</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;600&display=swap" rel="stylesheet" />
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
:root{--bg-0:#0a0e1a;--bg-1:#111827;--bg-2:#1a2236;--bg-3:#243049;--border:#2a3651;--border-strong:#3b4a6b;--text:#f8fafc;--text-muted:#94a3b8;--text-dim:#64748b;--accent:#60a5fa;--success:#10b981;--warning:#f59e0b;--danger:#ef4444;}
*{box-spacing:border-box;margin:0;padding:0;}
body{background:radial-gradient(1200px 600px at 10% -10%,rgba(59,130,246,.12),transparent 60%),var(--bg-0);color:var(--text);font-family:'Inter',sans-serif;padding:32px 40px;min-height:100vh;}
.container{max-width:1480px;margin:0 auto;}
.hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:28px;}
.hdr h1{font-size:28px;font-weight:800;letter-spacing:-.02em;}
.badge{padding:8px 16px;border-radius:999px;background:rgba(59,130,246,.15);border:1px solid var(--border-strong);font-size:13px;font-family:'JetBrains Mono',monospace;}
.kpi-row{display:grid;grid-template-columns:repeat(6,1fr);gap:14px;margin-bottom:32px;}
.kpi{background:linear-gradient(180deg,var(--bg-2),var(--bg-1));border:1px solid var(--border);border-radius:16px;padding:20px;position:relative;}
.kpi{position:relative;}
.kpi::before{content:"";position:absolute;top:0;left:0;right:0;height:3px;border-radius:16px 16px 0 0;background:var(--accent);}
.kpi.warn::before{background:var(--warning);}
.kpi.danger::before{background:var(--danger);}
.kpi.muted::before{background:var(--text-dim);}
.kpi-label{font-size:11px;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:.05em;}
.kpi-value{font-size:22px;font-weight:700;margin-top:10px;font-family:'JetBrains Mono',monospace;}
.kpi-sub{font-size:11px;color:var(--text-dim);margin-top:4px;}
.stitle{font-size:11px;font-weight:700;color:var(--text-muted);text-transform:uppercase;letter-spacing:.08em;margin-bottom:14px;display:flex;align-items:center;gap:10px;}
.stitle::after{content:"";flex:1;height:1px;background:linear-gradient(90deg,var(--border),transparent);}
.card{background:linear-gradient(180deg,var(--bg-2),var(--bg-1));border:1px solid var(--border);border-radius:16px;padding:24px;}
.section{margin-bottom:32px;}
.aging-table{width:100%;border-collapse:separate;border-spacing:0;font-size:13px;}
.aging-table th{background:var(--bg-3);color:var(--text-muted);font-weight:600;padding:12px;text-align:right;font-size:11px;text-transform:uppercase;letter-spacing:.05em;}
.aging-table th:first-child{text-align:left;padding-left:18px;}
.aging-table td{padding:11px 12px;text-align:right;font-family:'JetBrains Mono',monospace;border-bottom:1px solid var(--border);font-size:12px;}
.aging-table td:first-child{text-align:left;padding-left:18px;font-family:'Inter',sans-serif;font-weight:600;font-size:13px;}
.row-sub td{background:rgba(59,130,246,.08);font-weight:700;}
.row-tot td{background:var(--bg-3);font-weight:800;font-size:14px;}
.g2{display:grid;grid-template-columns:1.2fr 1fr;gap:20px;margin-bottom:32px;}
.chart-wrap{position:relative;height:280px;}
.top-table{width:100%;font-size:13px;border-collapse:collapse;}
.top-table td{padding:10px 6px;border-bottom:1px solid var(--border);}
.ins-grid{display:grid;grid-template-columns:repeat(8,1fr);gap:8px;}
.ins-card{background:var(--bg-1);border:1px solid var(--border);border-radius:12px;padding:12px;text-align:center;}
.rec-grid{display:grid;grid-template-columns:1fr auto 1fr;gap:20px;}
.rec-item{display:flex;justify-content:space-between;padding:10px 0;border-bottom:1px dashed var(--border);font-size:13px;}
.rec-item.total{font-weight:700;border-top:1px solid var(--border-strong);color:var(--accent);}
.dbadge{margin-top:14px;padding:8px 16px;border-radius:999px;font-size:12px;font-weight:700;text-align:center;color:#fff;}
.g3{display:grid;grid-template-columns:repeat(3,1fr);gap:20px;}
.frow{display:flex;justify-content:space-between;padding:10px 0;border-bottom:1px dashed var(--border);font-size:13px;}
.frow.total{font-weight:700;color:var(--warning);}
.gauge-bar{height:10px;background:var(--bg-3);border-radius:5px;margin-top:12px;overflow:hidden;}
.gauge-fill{height:100%;width:0%;transition:width .5s;border-radius:5px;}
.oltre-banner{background:linear-gradient(135deg,rgba(220,38,38,.15),rgba(220,38,38,.05));border:1px solid rgba(220,38,38,.3);border-radius:16px;padding:24px 32px;display:flex;align-items:center;justify-content:space-between;}
.oltre-val{font-size:36px;font-weight:800;font-family:'JetBrains Mono',monospace;color:#ef4444;}
.oltre-sub{font-size:12px;color:var(--text-dim);margin-top:4px;}
</style>
</head>
<body>
<div class="container">
  <div class="hdr">
    <h1>Dashboard Invecchiamento Crediti</h1>
    <div class="badge" id="ref-date">—</div>
  </div>
  <div class="kpi-row" id="kpi-row"></div>
  <div class="section">
    <div class="stitle">Matrice Standard (Fascia Scaduto × Tipo Doc)</div>
    <div class="card" style="padding:0;overflow:hidden;"><table class="aging-table" id="main-table"></table></div>
  </div>
  <div class="g2">
    <div class="card"><div class="stitle">Distribuzione per fascia</div><div class="chart-wrap"><canvas id="aging-chart"></canvas></div></div>
    <div class="card"><div class="stitle">Top clienti scaduti (&gt; 120 gg)</div><table class="top-table" id="top-overdue"></table></div>
  </div>
  <div class="section">
    <div class="stitle">Aging Insoluti (DA + AB)</div>
    <div class="card"><div class="ins-grid" id="insoluti-grid"></div></div>
  </div>
  <div class="section">
    <div class="stitle">Documenti emessi oltre 365 gg (data emissione — tutto il partitario)</div>
    <div class="oltre-banner">
      <div>
        <div style="font-size:13px;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px;">Importo totale</div>
        <div class="oltre-val" id="oltre-val">—</div>
        <div class="oltre-sub" id="oltre-sub">—</div>
      </div>
      <div style="text-align:right;color:var(--text-dim);font-size:12px;max-width:320px;">
        Posizioni aperte (inclusi CW, DA, AB) la cui <strong style="color:var(--text-muted)">data di emissione</strong> supera i 365 giorni rispetto alla data di riferimento.
      </div>
    </div>
  </div>
  <div class="section">
    <div class="stitle">Quadratura Co.Ge vs Partitario</div>
    <div class="rec-grid">
      <div class="card" id="partitario-side"></div>
      <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;">
        <div class="dbadge" id="delta-badge">—</div>
        <div id="delta-value" style="font-family:'JetBrains Mono',monospace;margin-top:8px;font-size:13px;"></div>
      </div>
      <div class="card" id="bilancio-side"></div>
    </div>
  </div>
  <div class="section">
    <div class="stitle">Capacità di copertura fondo</div>
    <div class="g3">
      <div class="card" id="rischio-card"></div>
      <div class="card" id="fondi-card"></div>
      <div class="card">
        <div class="stitle">Stato fondo</div>
        <div id="df-value" style="font-size:28px;font-weight:800;font-family:'JetBrains Mono';margin-top:10px;">—</div>
        <div class="gauge-bar"><div class="gauge-fill" id="gauge-fill"></div></div>
        <div id="gauge-status" style="margin-top:10px;font-size:12px;font-weight:600;"></div>
      </div>
    </div>
  </div>
</div>
<script>
const DATA = __DATA__;
const fmt  = n => (n===0||n==null) ? "—" : n.toLocaleString('it-IT',{maximumFractionDigits:0});
const fmtE = n => "€ " + fmt(n);

document.getElementById('ref-date').textContent = "Riferimento: " + DATA.ref_date;
const K = DATA.kpis;
document.getElementById('kpi-row').innerHTML = `
  <div class="kpi"><div class="kpi-label">Totale Crediti</div><div class="kpi-value">${fmtE(K.totale_crediti_aging)}</div></div>
  <div class="kpi warn"><div class="kpi-label">Scaduto Totale</div><div class="kpi-value">${fmtE(K.totale_scaduto)}</div></div>
  <div class="kpi warn"><div class="kpi-label">&gt; 120 gg</div><div class="kpi-value">${fmtE(K.crediti_120)}</div></div>
  <div class="kpi danger"><div class="kpi-label">Sofferenza &gt;180gg</div><div class="kpi-value">${fmtE(K.crediti_180)}</div></div>
  <div class="kpi"><div class="kpi-label">Insoluti DA+AB</div><div class="kpi-value">${fmtE(K.insoluti_total)}</div></div>
  <div class="kpi muted"><div class="kpi-label">Emessi &gt;365gg</div><div class="kpi-value" style="color:#ef4444">${fmtE(K.doc_oltre365)}</div><div class="kpi-sub">${(K.n_doc_oltre365||0).toLocaleString('it-IT')} doc</div></div>`;

// Main table
let html = '<thead><tr><th>Fascia</th>' + DATA.doc_types.map(dt=>`<th>${dt}</th>`).join('') + '<th>Totale</th></tr></thead><tbody>';
const scad = DATA.buckets.filter(b=>b!=='A scadere');
scad.forEach(b=>{
  html+=`<tr><td>${b} gg</td>`;
  let tot=0;
  DATA.doc_types.forEach(dt=>{const v=DATA.main_aging[b][dt];tot+=v;html+=`<td>${fmt(v)}</td>`;});
  html+=`<td style="font-weight:700">${fmt(tot)}</td></tr>`;
});
html+=`<tr class="row-sub"><td>TOTALE SCADUTO</td>`;
let sTot=0;
DATA.doc_types.forEach(dt=>{let s=0;scad.forEach(b=>s+=DATA.main_aging[b][dt]);sTot+=s;html+=`<td>${fmt(s)}</td>`;});
html+=`<td>${fmt(sTot)}</td></tr>`;
html+=`<tr><td>A Scadere</td>`;
let asTot=0;
DATA.doc_types.forEach(dt=>{const v=DATA.main_aging['A scadere'][dt];asTot+=v;html+=`<td>${fmt(v)}</td>`;});
html+=`<td style="font-weight:700">${fmt(asTot)}</td></tr>`;
html+=`<tr class="row-tot"><td>TOTALE CREDITI</td>`;
let gTot=0;
DATA.doc_types.forEach(dt=>{let t=0;DATA.buckets.forEach(b=>t+=DATA.main_aging[b][dt]);gTot+=t;html+=`<td>${fmt(t)}</td>`;});
html+=`<td>${fmt(gTot)}</td></tr></tbody>`;
document.getElementById('main-table').innerHTML=html;

new Chart(document.getElementById('aging-chart'),{type:'bar',data:{labels:DATA.buckets,datasets:[{data:DATA.buckets.map(b=>DATA.doc_types.reduce((s,dt)=>s+DATA.main_aging[b][dt],0)),backgroundColor:DATA.buckets.map(b=>b==='A scadere'?'#10b981':'#3b82f6'),borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{grid:{display:false}},y:{grid:{color:'rgba(255,255,255,0.05)'}}}}});

let topH='';
DATA.top_overdue.forEach(c=>topH+=`<tr><td><b>${c.cliente}</b></td><td style="color:var(--text-muted)">${c.nome.slice(0,28)}</td><td style="text-align:right;font-family:'JetBrains Mono';font-weight:600;color:#f59e0b">${fmtE(c.totale)}</td></tr>`);
document.getElementById('top-overdue').innerHTML=topH||'<tr><td colspan="3" style="color:var(--text-dim);text-align:center;padding:20px">Nessun dato</td></tr>';

let insH='';
DATA.buckets.forEach(b=>{const r=DATA.insoluti_aging[b];insH+=`<div class="ins-card"><div style="font-size:11px;color:var(--text-muted);margin-bottom:6px">${b}</div><div style="font-weight:700;font-family:'JetBrains Mono';font-size:13px">${fmt(r.DA+r.AB)}</div><div style="font-size:10px;color:var(--text-dim);margin-top:4px">DA:${fmt(r.DA)}</div></div>`;});
document.getElementById('insoluti-grid').innerHTML=insH;

document.getElementById('oltre-val').textContent=fmtE(K.doc_oltre365||0);
document.getElementById('oltre-sub').textContent=`${(K.n_doc_oltre365||0).toLocaleString('it-IT')} documenti · emissione antecedente al ${new Date(new Date(DATA.ref_date)-365*86400000).toLocaleDateString('it-IT')}`;

const R=DATA.reconciliation;
document.getElementById('partitario-side').innerHTML=`<div class="stitle">Partitario SAP</div><div class="rec-item"><span>Aging principale</span><span>${fmtE(R.totale_matrice_sap)}</span></div><div class="rec-item"><span>Cessioni CW</span><span>${fmtE(R.cessioni_wind_cw)}</span></div><div class="rec-item"><span>Insoluti DA+AB</span><span>${fmtE(R.insoluti_da_ab)}</span></div><div class="rec-item"><span>Floa / Sospesi</span><span>${fmtE(R.sospesi)}</span></div><div class="rec-item total"><span>Totale Lordo SAP</span><span>${fmtE(R.totale_lordo_sap)}</span></div>`;
document.getElementById('bilancio-side').innerHTML=`<div class="stitle">Bilancio Co.Ge</div><div class="rec-item"><span>12100000 — Crediti interno</span><span>${fmtE(R.conto_130100_italia)}</span></div><div class="rec-item"><span>12120000 — Crediti estero</span><span>${fmtE(R.conto_130300_estero)}</span></div><div class="rec-item total"><span>Totale Bilancio</span><span>${fmtE(R.totale_bilancio_coge)}</span></div>`;
document.getElementById('delta-value').textContent="Δ "+(R.delta_quadratura||0).toLocaleString('it-IT',{minimumFractionDigits:2,maximumFractionDigits:2});
const bD=document.getElementById('delta-badge');
if(Math.abs(R.delta_quadratura||0)<5){bD.textContent="✓ Quadrato";bD.style.background="#10b981";}else{bD.textContent="⚠ Disallineato";bD.style.background="#ef4444";}

const F=DATA.fondo;
document.getElementById('rischio-card').innerHTML=`<div class="stitle">Massa a rischio</div><div class="frow"><span>Sofferenza &gt;180gg</span><span>${fmtE(F.scaduto_critico_180)}</span></div><div class="frow"><span>Insoluti attivi</span><span>${fmtE(F.insoluti_totali)}</span></div><div class="frow total"><span>Totale</span><span>${fmtE(F.massa_a_rischio_totale)}</span></div>`;
document.getElementById('fondi-card').innerHTML=`<div class="stitle">Presidi di bilancio</div><div class="frow"><span>Fondo accantonato</span><span>${fmtE(F.fondo_svalutazione_coge)}</span></div><div class="frow"><span>Partite legali</span><span>${fmtE(F.partite_legali)}</span></div><div class="frow total" style="color:var(--accent)"><span>Totale coperture</span><span>${fmtE(F.totale_coperture_disponibili)}</span></div>`;
document.getElementById('df-value').textContent=fmtE(F.delta_copertura);
const pct=F.massa_a_rischio_totale>0?Math.min(Math.max(F.totale_coperture_disponibili/F.massa_a_rischio_totale*100,0),100):100;
document.getElementById('gauge-fill').style.cssText=`width:${pct}%;background:${F.delta_copertura>=0?'#10b981':'#ef4444'}`;
const gs=document.getElementById('gauge-status');
gs.textContent=F.delta_copertura>=0?"✓ Fondo capiente":"⚠ Fondo incapiente — accantonamento necessario";
gs.style.color=F.delta_copertura>=0?'#10b981':'#ef4444';
</script>
</body>
</html>
"""

html_final = HTML.replace("const DATA = __DATA__;", f"const DATA = {DATA_JS};")
nome_dashboard = os.environ.get("OUT_DASH", 'Aging_Dashboard.html')
with open(nome_dashboard, 'w', encoding='utf-8') as f:
    f.write(html_final)
print(f"Dashboard '{nome_dashboard}' generata con successo.")

# ==========================================
# 5. DOWNLOAD
# ==========================================
print("\nDownload avviato...")
files.download(nome_excel)
files.download(nome_dashboard)