"""
Nome: Rielaborazione Estratto Conto
Descrizione: Rielabora l'export SAP ordinando le partite aperte/chiuse per data/mese con quadratura automatica del saldo dealer e link alle coppie SU.
Categoria: Estratto Conto

Istruzioni:
■ FILE DI INPUT RICHIESTI:
  1. Export Excel SAP:
     • Struttura: Il foglio deve chiamarsi "Esportazione SAPUI5" e avere 19 colonne posizionate nell'ordine standard SAP:
       Col 0: Cliente, Col 1: Nome del cliente, Col 2: Data doc. contabile, Col 3: Documento contabile, Col 4: Riferimento, Col 5: Testo posizione, Col 6: Attribuzione, Col 7: Tipo doc. contabile, Col 8: Importo (div. soc.), Col 9: Scad. importo netto, Col 15: Data di pareggio, Col 16: Doc. cont. pareggio.

■ MODALITÀ D'USO:
  1. Carica l'Excel di export SAP.
  2. Selezionalo nel dropdown.
  3. Clicca su Esegui. Lo script filtrerà le righe, analizzerà le coppie SU e caricherà l'estratto conto finale pulito in 'File Generati' per il download.
"""


# --- Latest Version Code Cell ---
# @title ELABORAZIONE ESTRATTO V5.0

"""
Elaborazione Export SAP → Estratto Conto Dealer Riconciliato
Google Colab version — V5.0
Migliorie da Minuta V2:
  1. Sezione Posizioni Aperte (blu) in alto, Posizioni Chiuse (mattone) in basso,
     ciascuna con la propria struttura mensile interna
  2. Secondo foglio "Dati" nello stesso file: formato piatto, filtrabile, senza styling
  3. Link cliccabili tra le due righe di ogni coppia SU (self <-> carried)
  4. Colore SOLO sul font (non sulle celle), basato sul segno dell'importo
     (rosso = Uscita Dealer, verde = Entrata Dealer); grassetto per SU e ZP
  5. ZP: logica di etichettatura invariata, colore/font secondo il punto 4
  6. Colonna Attribuzione rimossa dall'output; "Ass. Iva" (quando presente in
     Attribuzione) spostato nel Testo posizione
"""

import pandas as pd, numpy as np, warnings, io
from collections import Counter
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from google.colab import files
warnings.filterwarnings('ignore')

MESI = {1:'Gennaio',2:'Febbraio',3:'Marzo',4:'Aprile',5:'Maggio',
        6:'Giugno',7:'Luglio',8:'Agosto',9:'Settembre',10:'Ottobre',
        11:'Novembre',12:'Dicembre'}

# Attribuzione RIMOSSA dall'output (punto 6)
OCOLS = ['Data doc. contabile','Documento contabile','Riferimento','Testo posizione',
         'Tipo doc. contabile','Importo (div. soc.)',
         'Scad. importo netto','Doc. cont. pareggio']

TIPO_DOC_LABEL = {
    'CD': 'Cessione del Credito',
    'W1': 'Fattura',
    'W4': 'Nota di Credito',
    'W7': 'Nota di Debito',
    'ZD': 'Fattura',
    'DZ': 'Incasso',
    'ZP': 'Pagamento ns. favore',
    'WP': 'Acconto @ Kolme',       # rinominato da "Acconto Cliente"
    'DA': 'Insoluto',
    'SU': 'Saldo di apertura',
    'SA': 'Movimento Drop'
}

# ═════════════════════════════════════════════════════════════════════════════
# FUNZIONI DI SUPPORTO E TRASFORMAZIONI
# ═════════════════════════════════════════════════════════════════════════════

def dp_int(v):
    try:
        if pd.isna(v): return None
        return int(round(float(v)))
    except: return None

def clean(v):
    if pd.isna(v): return ''
    s = str(v).strip()
    return '' if s in ('nan','None','NaT') else s

def fdate(v):
    try: return v.strftime('%d/%m/%Y') if pd.notna(v) else ''
    except: return ''

def fdoc(v):
    try: return str(int(float(v))) if pd.notna(v) else ''
    except: return clean(v)

def blank():
    return {c:'' for c in OCOLS}

def is_ass_iva(attrib):
    """Rileva 'Ass.Iva' / 'Ass. Iva' nella colonna Attribuzione, tollerando
    spazi e punteggiatura variabile."""
    s = clean(attrib).lower().replace(' ', '').replace('.', '')
    return s == 'assiva'

def trasforma_rif(tipo_doc, rif_grezzo, testo_pos):
    tipo = str(tipo_doc).strip()
    rif = clean(rif_grezzo)
    if tipo == 'CD':
        return f"IMEI {rif}" if rif else ""
    if tipo == 'W1':
        if rif.startswith('O'): return f"Ordine {rif[1:]}"
        if rif.startswith('L'): return f"Lok-Me {rif[1:]}"
        if rif.startswith('A'): return f"Attivazione {rif[1:]}"
        return rif
    if tipo == 'ZD':
        return rif
    if tipo in ('W4','ZP','WP'):
        return ""
    if tipo == 'W7':
        return clean(testo_pos) if clean(testo_pos) else ""
    if tipo == 'DZ':
        return rif
    return rif

def trasforma_testo(tipo_doc, testo, attrib):
    """Testo posizione standard, con override per Ass. Iva (punto 6)."""
    if is_ass_iva(attrib):
        return "Ass. Iva"
    tipo = str(tipo_doc).strip()
    t = clean(testo)
    if tipo == 'CD':
        return t
    if tipo == 'W7':
        return t
    return ""

def su_label(testo_pos, doc_num, importo, in_pareggio=False):
    t = clean(testo_pos)
    if 'Credito al' in t:
        return t
    if 'Debito al' in t:
        return t
    doc_str = fdoc(doc_num) if doc_num else '?'
    if in_pareggio:
        return f"Pareggio partite (doc: {doc_str})"
    else:
        if importo < 0:
            return f"Credito generico (doc: {doc_str})"
        else:
            return f"Debito generico (doc: {doc_str})"

def tipo_doc_label(tipo_doc, data_pareggio=None, importo=0):
    tipo = str(tipo_doc).strip()
    if tipo == 'ZP':
        if pd.notna(data_pareggio):
            m = data_pareggio.month
            y = data_pareggio.year
            if importo > 0:
                return f"Pagamento vs. favore {MESI.get(m,'')} {y}"
            else:
                return f"Pagamento ns. favore {MESI.get(m,'')} {y}"
        return TIPO_DOC_LABEL.get(tipo, tipo)
    if tipo == 'DZ':
        return "Incasso" if importo >= 0 else "Pagamento"
    return TIPO_DOC_LABEL.get(tipo, tipo)

def to_row(r, indent=0, tipo_override=None, in_pareggio=False):
    tipo = str(r[7]).strip()
    rif = trasforma_rif(tipo, r[4], r[5])
    testo = trasforma_testo(tipo, r[5], r[6])
    if indent > 0 and testo:
        testo = "  " * indent + testo
    lbl = tipo_override if tipo_override else tipo_doc_label(
        tipo,
        r[15] if pd.notna(r[15]) else None,
        float(r[8]) if pd.notna(r[8]) else 0
    )
    if tipo == 'SU' and not tipo_override:
        lbl = su_label(r[5], r[3], float(r[8]) if pd.notna(r[8]) else 0, in_pareggio)
    return {
        'Data doc. contabile': fdate(r[2]),
        'Documento contabile': fdoc(r[3]),
        'Riferimento': rif,
        'Testo posizione': testo,
        'Tipo doc. contabile': lbl,
        'Importo (div. soc.)': float(r[8]) if pd.notna(r[8]) else 0.0,
        'Scad. importo netto': fdate(r[9]),
        'Doc. cont. pareggio': fdoc(r[16])
    }

def group_month(rows):
    yms = [(r[9].year, r[9].month) for r in rows if pd.notna(r[9])]
    return Counter(yms).most_common(1)[0][0] if yms else (9999,99)

def group_sort_key(rows):
    types = set(r[7] for r in rows)
    if 'ZP' in types: return 3
    if 'DZ' in types: return 2
    return 1

def sort_group(rows):
    non_su = sorted([r for r in rows if r[7]!='SU'],
                     key=lambda r: (-(r[2].timestamp() if pd.notna(r[2]) else 0), -float(r[8] or 0)))
    su_recs = sorted([r for r in rows if r[7]=='SU'],
                      key=lambda r: -float(r[8] or 0))
    return non_su + su_recs

# ═════════════════════════════════════════════════════════════════════════════
# RIMOZIONE STORNI GENERICI (stesso doc_par, stesso tipo, importi opposti)
# ═════════════════════════════════════════════════════════════════════════════

def rimuovi_storni_generici(df):
    """Rimuove coppie storno: stesso doc_par (gruppo di pareggio), stesso
    Tipo doc. contabile, importi esattamente opposti. Applicata a tutti i
    tipi, SU compreso."""
    to_remove = set()
    grouped = df.groupby(df[16].apply(dp_int))
    for dp, group in grouped:
        if dp is None:
            continue
        indices = group.index.tolist()
        for i in range(len(indices)):
            if indices[i] in to_remove:
                continue
            r1 = df.loc[indices[i]]
            tipo1 = r1[7]
            imp1 = float(r1[8]) if pd.notna(r1[8]) else 0.0
            for j in range(i+1, len(indices)):
                if indices[j] in to_remove:
                    continue
                r2 = df.loc[indices[j]]
                tipo2 = r2[7]
                imp2 = float(r2[8]) if pd.notna(r2[8]) else 0.0
                if tipo1 == tipo2 and abs(imp1 + imp2) < 0.01:
                    to_remove.add(indices[i])
                    to_remove.add(indices[j])
                    break
    if to_remove:
        print(f"   Rimossi {len(to_remove)} record di storni generici")
    return df.drop(index=to_remove)

# ═════════════════════════════════════════════════════════════════════════════
# CARICAMENTO FILE
# ═════════════════════════════════════════════════════════════════════════════

print("="*70)
print("  ELABORAZIONE ESTRATTO CONTO DEALER DA EXPORT SAP  —  V5.0")
print("="*70)
print("\n📂 Seleziona il file Excel di export SAP...")

uploaded = files.upload()
filename = list(uploaded.keys())[0]
print(f"\n📄 File caricato: {filename}")

raw = pd.read_excel(io.BytesIO(uploaded[filename]), sheet_name='Esportazione SAPUI5')
raw.columns = list(range(raw.shape[1]))
for c in (2,9,15):
    raw[c] = pd.to_datetime(raw[c], errors='coerce')

dealer = f"{clean(raw.iloc[0,0])} - {clean(raw.iloc[0,1])}"
print(f"📋 Dealer: {dealer}")
print(f"   Righe totali: {len(raw)}")

# ═════════════════════════════════════════════════════════════════════════════
# FILTRO TEMPORALE
# ═════════════════════════════════════════════════════════════════════════════

SOGLIA = pd.Timestamp('2025-01-01')   # personalizzabile
print(f"\n📅 Soglia scadenza: >= {SOGLIA.strftime('%d/%m/%Y')}")

df = raw[raw[9] >= SOGLIA].copy()
print(f"   Filtro applicato: {len(raw)} → {len(df)} records")

# ═════════════════════════════════════════════════════════════════════════════
# SU ANALYSIS — IDENTIFICAZIONE PER DOC_PAR (per etichette e collegamenti)
# ═════════════════════════════════════════════════════════════════════════════

print("\n🔄 Analisi SU...")

su_all = raw[raw[7]=='SU'].copy()
su_info = {}
su_docs_valid = set()

for doc_num in su_all[3].unique():
    pair = su_all[su_all[3]==doc_num]
    if len(pair) != 2:
        print(f"   ⚠️  SU anomalo doc {int(doc_num)}: {len(pair)} records (attesi 2)")
        continue
    r1, r2 = pair.iloc[0], pair.iloc[1]
    dp1, dp2 = dp_int(r1[16]), dp_int(r2[16])
    dn = int(doc_num)

    # NOTA: NON escludere quando dp1==dp2. Il caso "entrambi i record puntano
    # allo stesso doc_par" è il pattern NORMALE di un SU auto-bilanciato che
    # chiude le partite all'interno del proprio stesso gruppo (dp1==dp2==dn).
    # Non è uno storno tecnico: è la maggioranza dei casi in questo dataset.

    r1_self = (dp1 is not None and dp1 == dn)
    r2_self = (dp2 is not None and dp2 == dn)

    if r1_self and r2_self:
        # Caso normale: SU pareggio partite, entrambi i record nello stesso
        # gruppo (dp1==dp2==dn). Nessun "carry" esterno: la coppia si chiude
        # interamente dentro il proprio gruppo.
        if r1[8] >= 0:
            su_info[dn] = {'self_imp':r1[8],'carried_imp':r2[8],'self_dp':dp1,'carried_dp':dp2,
                            'self_idx':r1.name,'carried_idx':r2.name}
        else:
            su_info[dn] = {'self_imp':r2[8],'carried_imp':r1[8],'self_dp':dp2,'carried_dp':dp1,
                            'self_idx':r2.name,'carried_idx':r1.name}
        su_docs_valid.add(dn)
    elif r1_self and not r2_self:
        su_info[dn] = {'self_imp':r1[8],'carried_imp':r2[8],'self_dp':dp1,'carried_dp':dp2,
                        'self_idx':r1.name,'carried_idx':r2.name}
        su_docs_valid.add(dn)
    elif r2_self and not r1_self:
        su_info[dn] = {'self_imp':r2[8],'carried_imp':r1[8],'self_dp':dp2,'carried_dp':dp1,
                        'self_idx':r2.name,'carried_idx':r1.name}
        su_docs_valid.add(dn)
    elif not r1_self and not r2_self:
        if dp1 is None:
            su_info[dn] = {'self_imp':r2[8],'carried_imp':r1[8],'self_dp':dp2,'carried_dp':dp1,
                            'self_idx':r2.name,'carried_idx':r1.name}
        else:
            su_info[dn] = {'self_imp':r1[8],'carried_imp':r2[8],'self_dp':dp1,'carried_dp':dp2,
                            'self_idx':r1.name,'carried_idx':r2.name}
        su_docs_valid.add(dn)

    if dn in su_info and su_info[dn]['self_imp'] < 0:
        print(f"   ⚡ SU DEBITO doc {dn}: self={su_info[dn]['self_imp']:+.2f}, "
              f"carried={su_info[dn]['carried_imp']:+.2f}")

df = df[~((df[7]=='SU') & (~df[3].isin(su_docs_valid)))].copy()
print(f"   Coppie SU identificate: {len(su_info)}")

# ═════════════════════════════════════════════════════════════════════════════
# RIMOZIONE STORNI GENERICI (ora inclusi gli SU)
# ═════════════════════════════════════════════════════════════════════════════
# Regola: in uno stesso gruppo di pareggio, se due righe hanno lo stesso Tipo
# doc. contabile e importi esattamente opposti, vengono rimosse entrambe.
# Applicata a tutti i tipi, SU compreso: cattura sia le coppie SU
# auto-bilanciate (stesso documento, si chiudono nel proprio gruppo) sia
# eventuali storni tecnici incrociati fra documenti SU diversi che si
# annullano nello stesso gruppo. I carry nidificati non vengono toccati
# perché le due gambe di un documento "portato" vivono in gruppi diversi e
# quindi non trovano un opposto nello stesso gruppo.

print("\n🔄 Rimozione storni generici (inclusi SU)...")
df = rimuovi_storni_generici(df)
print(f"   Records dopo pulizia storni: {len(df)}")

print(f"   Coppie SU valide: {len(su_info)}")

# L'indice originale di `raw` è preservato lungo tutto il flusso (nessun
# reset_index) in modo che self_idx/carried_idx in su_info coincidano con
# l'indice di riga usato in `df` — necessario per i collegamenti ipertestuali
# tra le due righe di ogni coppia SU.

# ═════════════════════════════════════════════════════════════════════════════
# GROUPING PER DOC_PAR
# ═════════════════════════════════════════════════════════════════════════════

print("\n🔄 Raggruppamento per Doc. cont. pareggio...")

groups = {}
open_recs = []
for _, r in df.iterrows():
    dp = dp_int(r[16])
    if dp is None:
        open_recs.append(r)
    else:
        groups.setdefault(dp, []).append(r)

credito_al = [r for r in open_recs
              if r[7]=='SU' and pd.notna(r[5]) and 'Credito al' in str(r[5])]
credito_idx = set(id(r) for r in credito_al)
pure_open = [r for r in open_recs if id(r) not in credito_idx]

print(f"   Gruppi per doc_par: {len(groups)}")
print(f"   Posizioni aperte: {len(pure_open)}")
print(f"   SU 'Credito al' aperto: {len(credito_al)}")

# ═════════════════════════════════════════════════════════════════════════════
# ORGANIZZAZIONE PER MESE (sezione CHIUSE)
# ═════════════════════════════════════════════════════════════════════════════

months = {}
for dp, rows in groups.items():
    ym = group_month(rows)
    months.setdefault(ym, []).append(dp)
for ym in months:
    months[ym].sort(key=lambda dp: (group_sort_key(groups[dp]), -dp))

# ═════════════════════════════════════════════════════════════════════════════
# ORGANIZZAZIONE PER MESE (sezione APERTE)
# ═════════════════════════════════════════════════════════════════════════════

all_open = list(credito_al) + list(pure_open)
open_by_month = {}
for r in all_open:
    ym = (r[9].year, r[9].month) if pd.notna(r[9]) else (9999,99)
    open_by_month.setdefault(ym, []).append(r)

def sort_open_month(rows):
    """SU (credito al) prima, poi il resto per data doc. decrescente."""
    su_recs = [r for r in rows if r[7]=='SU']
    other = sorted([r for r in rows if r[7]!='SU'],
                    key=lambda r: -(r[2].timestamp() if pd.notna(r[2]) else 0))
    return su_recs + other

# ═════════════════════════════════════════════════════════════════════════════
# RENDER OUTPUT — ogni riga è (tipo_riga, dizionario_colonne, riga_raw_o_None)
# la riga_raw serve solo per il collegamento ipertestuale dei record SU
# ═════════════════════════════════════════════════════════════════════════════

print("\n🔄 Generazione output...")

output = []

# ── SEZIONE 1: POSIZIONI APERTE (in alto) ──
hdr = blank(); hdr['Tipo doc. contabile'] = 'POSIZIONI APERTE'
output.append(('sezione_aperte', hdr, None))

tot_aperte = 0.0
for ym in sorted(open_by_month.keys()):
    yr, mo = ym
    month_total = 0.0
    if ym != (9999,99):
        sub = blank(); sub['Tipo doc. contabile'] = f'Scadenza {MESI.get(mo,"?")} {yr}'
        output.append(('open_month', sub, None))
    for r in sort_open_month(open_by_month[ym]):
        val = float(r[8]) if pd.notna(r[8]) else 0.0
        if r[7] == 'SU':
            in_par = False  # SU in sezione aperte = sempre posizione aperta, non pareggio
            output.append(('su_open' if r[5] and 'Credito al' in str(r[5]) else 'su',
                            to_row(r, in_pareggio=in_par), r))
        else:
            output.append(('open', to_row(r), None))
        tot_aperte += val
        month_total += val
    
    if ym != (9999,99):
        tot_m = blank()
        tot_m['Tipo doc. contabile'] = f'Totale Scadenza {MESI.get(mo,"?")} {yr}'
        tot_m['Importo (div. soc.)'] = round(month_total, 2)
        output.append(('month_tot', tot_m, None))
        output.append(('blank', blank(), None))

tot = blank(); tot['Tipo doc. contabile'] = 'Totale Posizioni Aperte'
tot['Importo (div. soc.)'] = round(tot_aperte, 2)
output.append(('open_tot', tot, None))
output.append(('blank', blank(), None))
output.append(('blank', blank(), None))

# ── SEZIONE 2: POSIZIONI CHIUSE (in basso) ──
hdr = blank(); hdr['Tipo doc. contabile'] = 'POSIZIONI CHIUSE'
output.append(('sezione_chiuse', hdr, None))

tot_chiuse = 0.0
for ym in sorted(months.keys(), reverse=True):
    yr, mo = ym
    if ym == (9999,99): continue
    hdr = blank()
    hdr['Tipo doc. contabile'] = f'{MESI.get(mo,"?")} {yr}'
    output.append(('month_hdr', hdr, None))

    month_total = 0.0
    for dp in months[ym]:
        for r in sort_group(groups[dp]):
            tipo = r[7]
            indent = 1
            in_pareggio = (tipo == 'SU')
            if tipo == 'SU':
                output.append(('su', to_row(r, indent=indent, in_pareggio=in_pareggio), r))
            elif tipo == 'ZP':
                label = tipo_doc_label('ZP', r[15] if pd.notna(r[15]) else r[9],
                                       float(r[8]) if pd.notna(r[8]) else 0)
                output.append(('zp', to_row(r, indent=indent, tipo_override=label), None))
            elif tipo == 'DZ':
                output.append(('dz', to_row(r, indent=indent), None))
            elif tipo == 'WP':
                output.append(('wp', to_row(r, indent=indent), None))
            else:
                output.append(('data', to_row(r, indent=indent), None))
            month_total += float(r[8]) if pd.notna(r[8]) else 0
        output.append(('sep', blank(), None))

    tot_m = blank()
    tot_m['Tipo doc. contabile'] = f'Totale {MESI.get(mo,"?")} {yr}'
    tot_m['Importo (div. soc.)'] = round(month_total, 2)
    output.append(('month_tot', tot_m, None))
    output.append(('blank', blank(), None))
    tot_chiuse += month_total

tot_c = blank(); tot_c['Tipo doc. contabile'] = 'Totale Posizioni Chiuse'
tot_c['Importo (div. soc.)'] = round(tot_chiuse, 2)
output.append(('chiuse_tot', tot_c, None))

# ═════════════════════════════════════════════════════════════════════════════
# VERIFICA
# ═════════════════════════════════════════════════════════════════════════════

DATA_TYPES = {'data','su','su_open','zp','dz','wp','open'}
total = sum(rd['Importo (div. soc.)'] for rt, rd, _ in output
            if rt in DATA_TYPES and isinstance(rd.get('Importo (div. soc.)'), (int,float)))
print(f"\n   Output: {len(output)} righe")
print(f"   Somma dati: {total:,.2f}")

# ═════════════════════════════════════════════════════════════════════════════
# EXCEL — FOGLIO 1: ESTRATTO FORMATTATO
# ═════════════════════════════════════════════════════════════════════════════

print("\n💾 Scrittura Excel...")

wb = Workbook()
ws = wb.active
ws.title = f'Estratto {dealer[:20]}'
ws.append(OCOLS)
CWS = [12,14,26,30,34,13,12,14]     # 8 colonne (Attribuzione rimossa)
for ci, c in enumerate(ws[1], 1):
    c.font = Font(name='Arial', bold=True, color='FFFFFF', size=10)
    c.fill = PatternFill('solid', fgColor='1F4E79')
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.column_dimensions[get_column_letter(ci)].width = CWS[ci-1]
ws.row_dimensions[1].height = 28
ws.freeze_panes = 'A2'

# Palette per sezione: banner (header sezione) / sub (header mese) /
# tot (righe di totale) / data (righe dati, tinta leggera) — punto 1
PALETTE = {
    'aperte': {'banner':'2E75B6', 'sub':'BDD7EE', 'tot':'9DC3E6', 'data':'DCE6F1'},
    'chiuse': {'banner':'A0522D', 'sub':'E8C9AE', 'tot':'D9A679', 'data':'F2E0D5'},
}
FLOW_BOLD_TYPES = {'su','su_open','zp'}   # grassetto per SU e ZP (fine mese)
FLOW_COLS = {5, 6}   # Tipo doc. contabile, Importo (div. soc.) — punto 4

side = Side(style='thin', color='D9D9D9')
brd = Border(top=side, bottom=side, left=side, right=side)

# excel_row_by_idx: per i link ipertestuali dei record SU (punto 3)
excel_row_by_idx = {}
current_section = 'aperte'   # sezione attiva durante lo scorrimento

for rtype, rd, raw_row in output:
    ws.append([rd.get(c,'') for c in OCOLS])
    r = ws.max_row

    if raw_row is not None:
        excel_row_by_idx[raw_row.name] = r

    if rtype == 'sezione_aperte': current_section = 'aperte'
    elif rtype == 'sezione_chiuse': current_section = 'chiuse'
    pal = PALETTE[current_section]

    # ── righe strutturali (banner sezione, header mese, totali) ──
    if rtype in ('sezione_aperte','sezione_chiuse'):
        fill_c, font_c, ht = pal['banner'], 'FFFFFF', 24
    elif rtype in ('open_month','month_hdr'):
        fill_c, font_c, ht = pal['sub'], '1F1F1F', 18
    elif rtype in ('month_tot','open_tot','chiuse_tot'):
        fill_c, font_c, ht = pal['tot'], '1F1F1F', 18
    elif rtype in ('blank','sep'):
        fill_c, font_c, ht = 'FFFFFF', '000000', (4 if rtype=='blank' else 6)
    else:
        fill_c, font_c, ht = None, None, None   # righe dati, gestite sotto

    if fill_c is not None:
        bold = rtype not in ('blank','sep')
        for ci, cell in enumerate(ws[r], 1):
            cell.fill = PatternFill('solid', fgColor=fill_c)
            cell.border = brd
            cell.font = Font(name='Arial', size=10 if ht>=15 else 6, bold=bold, color=font_c)
            if ci == 6 and isinstance(cell.value, (int,float)):
                cell.number_format = '#,##0.00;-#,##0.00'
                cell.alignment = Alignment(horizontal='right')
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center')
        ws.row_dimensions[r].height = ht
        continue

    # ── righe DATI: sfondo tinta leggera di sezione (punto 1); colore
    # verde/rosso SOLO su Tipo doc. contabile e Importo (punto 4) ──
    bold = rtype in FLOW_BOLD_TYPES

    for ci, cell in enumerate(ws[r], 1):
        cell.fill = PatternFill('solid', fgColor=pal['data'])
        cell.border = brd
        cell.font = Font(name='Arial', size=10, bold=bold, color='000000')
        if ci == 6 and isinstance(cell.value, (int,float)):
            cell.number_format = '#,##0.00;-#,##0.00'
            cell.alignment = Alignment(horizontal='right')
        else:
            cell.alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[r].height = 15

# ── Collegamenti ipertestuali tra le due righe di ogni coppia SU (punto 3) ──
print("\n🔗 Inserimento collegamenti SU...")
link_count = 0
for dn, info in su_info.items():
    idx_self, idx_carried = info['self_idx'], info['carried_idx']
    row_self = excel_row_by_idx.get(idx_self)
    row_carried = excel_row_by_idx.get(idx_carried)
    if row_self and row_carried:
        cell_self = ws.cell(row=row_self, column=2)      # Documento contabile
        cell_carried = ws.cell(row=row_carried, column=2)
        cell_self.hyperlink = f"#'{ws.title}'!A{row_carried}"
        cell_carried.hyperlink = f"#'{ws.title}'!A{row_self}"
        cell_self.font = Font(name='Arial', size=10, bold=True, color='0563C1', underline='single')
        cell_carried.font = Font(name='Arial', size=10, bold=True, color='0563C1', underline='single')
        link_count += 1
print(f"   Collegamenti creati: {link_count} coppie")

ws.auto_filter.ref = f"A1:{get_column_letter(len(OCOLS))}{ws.max_row}"

# ═════════════════════════════════════════════════════════════════════════════
# EXCEL — FOGLIO 2: DATI (formato piatto, filtrabile) — punto 2
# ═════════════════════════════════════════════════════════════════════════════

ws2 = wb.create_sheet('Dati')
ws2.append(OCOLS)
for ci, c in enumerate(ws2[1], 1):
    c.font = Font(name='Arial', bold=True, size=10)
    ws2.column_dimensions[get_column_letter(ci)].width = CWS[ci-1]

for rtype, rd, _ in output:
    if rtype in DATA_TYPES:
        # rimuove eventuale indentazione testuale usata nel foglio formattato
        row_vals = [rd.get(c,'') for c in OCOLS]
        row_vals[3] = str(row_vals[3]).strip() if row_vals[3] else ''  # Testo posizione
        ws2.append(row_vals)
        r2 = ws2.max_row
        cell6 = ws2.cell(row=r2, column=6)
        if isinstance(cell6.value, (int,float)):
            cell6.number_format = '#,##0.00;-#,##0.00'

ws2.freeze_panes = 'A2'
ws2.auto_filter.ref = f"A1:{get_column_letter(len(OCOLS))}{ws2.max_row}"

# ═════════════════════════════════════════════════════════════════════════════
# SALVATAGGIO
# ═════════════════════════════════════════════════════════════════════════════

nome_out = filename.rsplit('.',1)[0] + '_Estratto.xlsx'
wb.save(nome_out)

print(f"\n✅ ELABORAZIONE COMPLETATA")
print(f"   File: {nome_out}")
print(f"   Dealer: {dealer}")
print(f"   Records elaborati: {len(df)}")
print(f"   Somma totale: {total:,.2f}")
print(f"   Fogli: 'Estratto {dealer[:20]}' (formattato) + 'Dati' (piatto)")

files.download(nome_out)
