"""
Nome: Creazione Dashboard Scadenze Automatico
Descrizione: Crea la dashboard interattiva HTML del Collection Rate e dell'esposizione del fido partendo da un foglio Excel 'DATI'.
Categoria: Utility

Istruzioni:
■ FILE DI INPUT RICHIESTI:
  1. Excel Scadenze (es. DATI_LUGLIO2_2.xlsx):
     • Struttura: Il foglio deve chiamarsi "DATI".
     • Colonne attese: Codice cliente, Nome del cliente, Importo, Saldato (booleano), Agente, Area Geografica, Fido assegnato, Fatturato ultimi 3 mesi e in corso, Condizione di pagamento.
     *Nota: Se 'Importo' e 'Saldato' sono invertite nel file Excel, lo script le rileva e le inverte automaticamente in memoria.*

■ MODALITÀ D'USO:
  1. Carica il file Excel con la scheda DATI.
  2. Selezionalo nel dropdown.
  3. Clicca su Esegui per creare la dashboard interattiva HTML del Collection Rate e dell'esposizione del fido.
"""


# --- Latest Version Code Cell ---
# @title 📊 DASHBOARD INSOLUTI KOLME — Generatore Automatico v3 (light version)

# ════════════════════════════════════════════════════════════════
#  DASHBOARD INSOLUTI — KOLME
#  Google Colab notebook
#
#  Istruzioni:
#  1. Runtime → Run All
#  2. Carica il file Excel contenente il foglio "DATI"
#  3. Lo script genera la dashboard HTML e la scarica
#
#  Colonne attese nel foglio DATI:
#  - Codice cliente, Importo, Saldato, Area Geografica,
#    Fido assegnato, Fatturato ultimi 3 mesi e in corso,
#    Condizione di pagamento
#  - Agente e Nome del cliente: opzionali (se assenti, omessi)
# ════════════════════════════════════════════════════════════════

# %% UPLOAD FILE ─────────────────────────────────────────────────

from google.colab import files
import pandas as pd
import io

print("📂 Carica il file Excel con il foglio 'DATI'...")
uploaded = files.upload()

filename = list(uploaded.keys())[0]
df = pd.read_excel(io.BytesIO(uploaded[filename]), sheet_name='DATI')

print(f"\n✓ File caricato: {filename}")
print(f"✓ Posizioni totali: {len(df)}")

# --- SELF-HEALING & NORMALIZZAZIONE DATI ---
if 'Importo' in df.columns and 'Saldato' in df.columns:
    # Se 'Importo' contiene solo valori booleani e 'Saldato' è numerico con importi reali, sono scambiate
    importo_is_bool = df['Importo'].dtype == 'bool' or df['Importo'].isin([True, False, 1, 0, None, pd.NA]).all()
    saldato_is_numeric = pd.api.types.is_numeric_dtype(df['Saldato']) and df['Saldato'].max() > 1.0
    
    if importo_is_bool and saldato_is_numeric:
        print("⚠️ [Correzione automatica] Le colonne 'Importo' e 'Saldato' risultano invertite nell'Excel. Scambio in corso...")
        df.rename(columns={'Importo': 'Saldato_temp', 'Saldato': 'Importo'}, inplace=True)
        df.rename(columns={'Saldato_temp': 'Saldato'}, inplace=True)

# Assicurati che 'Saldato' sia di tipo booleano per evitare errori con l'operatore di inversione (~)
if 'Saldato' in df.columns:
    if df['Saldato'].dtype == 'object':
        df['Saldato'] = df['Saldato'].astype(str).str.strip().str.upper().isin(['SI', 'YES', '1', 'VERO', 'TRUE'])
    else:
        df['Saldato'] = df['Saldato'].astype(bool)

# %% CONFIGURAZIONE PERIODO ──────────────────────────────────────

PERIODO = "Maggio 2026"    # ← Modifica qui il periodo di riferimento
SDD_COND = "SDD 30 gg Data Fattura Fine Mese"  # ← Label SDD esatta

# %% CALCOLO KPI ─────────────────────────────────────────────────

# Colonne opzionali
HAS_AGENTE = 'Agente' in df.columns
HAS_NOME   = 'Nome del cliente' in df.columns

df['is_SDD'] = df['Condizione di pagamento'].str.contains('SDD', na=False)
df['ratio_ins_fido'] = df.apply(
    lambda r: r['Importo'] / r['Fido assegnato'] if r['Fido assegnato'] > 0 else None,
    axis=1
)

# KPI Portfolio & Cash Flow
tot_scaduto   = df['Importo'].sum()
tot_insoluto  = df.loc[~df['Saldato'], 'Importo'].sum()
tot_incassato = df.loc[df['Saldato'],  'Importo'].sum()
n_posizioni   = len(df)
n_saldate     = int(df['Saldato'].sum())
n_non_saldate = int((~df['Saldato']).sum())

# ── COMPOSIZIONE INSOLUTO (nuova logica) ──────────────────────
# SDD su totale scaduto (tutte le posizioni, saldo + insoluto)
sdd_tot_importo     = df.loc[df['Condizione di pagamento'] == SDD_COND, 'Importo'].sum()
non_sdd_tot_importo = df.loc[df['Condizione di pagamento'] != SDD_COND, 'Importo'].sum()
pct_sdd_scaduto     = sdd_tot_importo     / tot_scaduto if tot_scaduto > 0 else 0
pct_non_sdd_scaduto = non_sdd_tot_importo / tot_scaduto if tot_scaduto > 0 else 0

# Collection Rate
cr_num = n_saldate / n_posizioni    if n_posizioni   > 0 else 0
cr_val = tot_incassato / tot_scaduto if tot_scaduto  > 0 else 0

# Fido & Fatturato medio
fido_med_saldati     = df.loc[ df['Saldato'], 'Fido assegnato'].mean()
fido_med_non_saldati = df.loc[~df['Saldato'], 'Fido assegnato'].mean()
fatt_med_saldati     = df.loc[ df['Saldato'], 'Fatturato ultimi 3 mesi e in corso'].mean()
fatt_med_non_saldati = df.loc[~df['Saldato'], 'Fatturato ultimi 3 mesi e in corso'].mean()

scost_fido = (fido_med_saldati - fido_med_non_saldati) / fido_med_non_saldati if fido_med_non_saldati > 0 else 0
scost_fatt = (fatt_med_saldati - fatt_med_non_saldati) / fatt_med_non_saldati if fatt_med_non_saldati > 0 else 0

# Ticket medio
ticket_medio            = df['Importo'].mean()
ticket_medio_non_saldato = df.loc[~df['Saldato'], 'Importo'].mean()

# Top 5 debitori (non saldati) — sempre con Codice cliente
top_cols = ['Codice cliente', 'Importo', 'Fido assegnato', 'Fatturato ultimi 3 mesi e in corso']
if HAS_AGENTE: top_cols.insert(1, 'Agente')
top_debitori = df.loc[~df['Saldato']].nlargest(5, 'Importo')[top_cols]

# Collection Rate per Area
cr_per_area = []
for area in sorted(df['Area Geografica'].unique()):
    sub = df[df['Area Geografica'] == area]
    pos = len(sub)
    sal = int(sub['Saldato'].sum())
    non_sal = int((~sub['Saldato']).sum())
    imp_tot = sub['Importo'].sum()
    imp_sal = sub.loc[sub['Saldato'], 'Importo'].sum()
    cr_per_area.append({
        'area': area, 'pos': pos, 'sal': sal, 'non_sal': non_sal,
        'cr_n': sal / pos if pos > 0 else 0,
        'cr_v': imp_sal / imp_tot if imp_tot > 0 else 0,
        'imp_non_sal': sub.loc[~sub['Saldato'], 'Importo'].sum()
    })

# Collection Rate per Condizione di Pagamento
cr_per_cond = []
for cond in df['Condizione di pagamento'].unique():
    sub = df[df['Condizione di pagamento'] == cond]
    pos = len(sub)
    sal = int(sub['Saldato'].sum())
    imp_tot = sub['Importo'].sum()
    imp_sal = sub.loc[sub['Saldato'], 'Importo'].sum()
    cr_per_cond.append({
        'cond': cond, 'pos': pos, 'sal': sal, 'non_sal': pos - sal,
        'cr_n': sal / pos if pos > 0 else 0,
        'cr_v': imp_sal / imp_tot if imp_tot > 0 else 0,
    })

print(f"\n✓ Periodo:              {PERIODO}")
print(f"✓ Totale scaduto:       €{tot_scaduto:,.0f}")
print(f"✓ Insoluto residuo:     €{tot_insoluto:,.0f}")
print(f"✓ Collection Rate (val.): {cr_val:.1%}")
print(f"✓ SDD / tot. scaduto:   {pct_sdd_scaduto:.1%}")
print(f"✓ Non-SDD / tot. scad.: {pct_non_sdd_scaduto:.1%}")

# %% HELPER DI FORMATTAZIONE ─────────────────────────────────────

def fmt_euro(v, dec=0):
    if v is None or pd.isna(v): return "—"
    s = f"{v:,.{dec}f}"
    return "€ " + s.replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_pct(v, dec=1):
    if v is None or pd.isna(v): return "—"
    return f"{v*100:.{dec}f}%".replace(".", ",")

# %% GENERAZIONE DASHBOARD HTML (VERSIONE LIGHT) ─────────────────

try:
    from IPython.display import HTML, display
    HAS_IPYTHON = True
except ImportError:
    HAS_IPYTHON = False

def build_dashboard():
    # Donut Collection Rate
    cr_pct = cr_val * 100
    radius = 70
    circumference = 2 * 3.14159 * radius
    dash_offset = circumference - (circumference * cr_val)

    # Altezze barre Fido
    max_fido = max(fido_med_saldati, fido_med_non_saldati)
    fido_sal_h     = (fido_med_saldati     / max_fido) * 100 if max_fido > 0 else 0
    fido_non_sal_h = (fido_med_non_saldati / max_fido) * 100 if max_fido > 0 else 0

    # Altezze barre Fatturato
    max_fatt = max(fatt_med_saldati, fatt_med_non_saldati)
    fatt_sal_h     = (fatt_med_saldati     / max_fatt) * 100 if max_fatt > 0 else 0
    fatt_non_sal_h = (fatt_med_non_saldati / max_fatt) * 100 if max_fatt > 0 else 0

    # ── TOP 5: Codice cliente; Agente se disponibile ─────────
    top_rows = ""
    for i, (_, row) in enumerate(top_debitori.iterrows(), 1):
        ratio = row['Importo'] / row['Fido assegnato'] if row['Fido assegnato'] > 0 else None
        ratio_str   = fmt_pct(ratio) if ratio is not None else "—"
        ratio_color = "#d32f2f" if ratio and ratio > 0.5 else "#e67e22"
        agente_cell = f'<td style="padding:14px 12px;color:#0b1e3d;font-size:12px">{row["Agente"]}</td>' if HAS_AGENTE else ''
        top_rows += f'''
        <tr style="border-bottom:1px solid #e9ecef">
          <td style="padding:14px 12px;color:#0b1e3d;font-weight:700;font-size:13px;text-align:center">#{i}</td>
          <td style="padding:14px 12px;color:#1e3a5f;font-weight:600;font-size:13px;font-family:'JetBrains Mono',monospace">{row['Codice cliente']}</td>
          {agente_cell}
          <td style="padding:14px 12px;color:#0b1e3d;font-weight:700;font-size:14px;text-align:right;font-family:'JetBrains Mono',monospace">{fmt_euro(row['Importo'])}</td>
          <td style="padding:14px 12px;color:#5a6e8c;font-size:12px;text-align:right;font-family:'JetBrains Mono',monospace">{fmt_euro(row['Fido assegnato'])}</td>
          <td style="padding:14px 12px;color:#5a6e8c;font-size:12px;text-align:right;font-family:'JetBrains Mono',monospace">{fmt_euro(row['Fatturato ultimi 3 mesi e in corso'])}</td>
          <td style="padding:14px 12px;text-align:right">
            <span style="background:{ratio_color}15;color:{ratio_color};padding:4px 10px;border-radius:6px;font-size:11px;font-weight:700;font-family:'JetBrains Mono',monospace">{ratio_str}</span>
          </td>
        </tr>'''

    agente_th = '<th style="padding:12px;text-align:left;font-size:10px;color:#0b1e3d;font-weight:700;text-transform:uppercase;letter-spacing:1px">Agente</th>' if HAS_AGENTE else ''

    # ── AREA ROWS ─────────────────────────────────────────────
    area_rows = ""
    for r in cr_per_area:
        cr_color = "#2e7d32" if r['cr_v'] >= 0.85 else ("#e67e22" if r['cr_v'] >= 0.70 else "#c62828")
        area_rows += f'''
        <tr style="border-bottom:1px solid #e9ecef">
          <td style="padding:12px;color:#0b1e3d;font-weight:600;font-size:13px">Area {r['area']}</td>
          <td style="padding:12px;text-align:right;color:#5a6e8c;font-size:12px;font-family:'JetBrains Mono',monospace">{r['pos']}</td>
          <td style="padding:12px;text-align:right;color:#2e7d32;font-size:12px;font-weight:600;font-family:'JetBrains Mono',monospace">{r['sal']}</td>
          <td style="padding:12px;text-align:right;color:#c62828;font-size:12px;font-weight:600;font-family:'JetBrains Mono',monospace">{r['non_sal']}</td>
          <td style="padding:12px;text-align:right;color:#0b1e3d;font-weight:700;font-size:13px;font-family:'JetBrains Mono',monospace">{fmt_pct(r['cr_n'])}</td>
          <td style="padding:12px;text-align:right">
            <span style="background:{cr_color}15;color:{cr_color};padding:4px 10px;border-radius:6px;font-size:11px;font-weight:700;font-family:'JetBrains Mono',monospace">{fmt_pct(r['cr_v'])}</span>
          </td>
          <td style="padding:12px;text-align:right;color:#c62828;font-size:12px;font-weight:600;font-family:'JetBrains Mono',monospace">{fmt_euro(r['imp_non_sal'])}</td>
        </tr>'''

    # ── CONDIZIONE ROWS ───────────────────────────────────────
    cond_rows = ""
    for r in cr_per_cond:
        cr_color  = "#2e7d32" if r['cr_v'] >= 0.85 else ("#e67e22" if r['cr_v'] >= 0.70 else "#c62828")
        cond_short = r['cond'].replace(' Data Fattura Fine Mese', '')
        cond_rows += f'''
        <tr style="border-bottom:1px solid #e9ecef">
          <td style="padding:12px;color:#0b1e3d;font-weight:600;font-size:13px">{cond_short}</td>
          <td style="padding:12px;text-align:right;color:#5a6e8c;font-size:12px;font-family:'JetBrains Mono',monospace">{r['pos']}</td>
          <td style="padding:12px;text-align:right;color:#2e7d32;font-size:12px;font-weight:600;font-family:'JetBrains Mono',monospace">{r['sal']}</td>
          <td style="padding:12px;text-align:right;color:#c62828;font-size:12px;font-weight:600;font-family:'JetBrains Mono',monospace">{r['non_sal']}</td>
          <td style="padding:12px;text-align:right">
            <span style="background:{cr_color}15;color:{cr_color};padding:4px 10px;border-radius:6px;font-size:11px;font-weight:700;font-family:'JetBrains Mono',monospace">{fmt_pct(r['cr_v'])}</span>
          </td>
        </tr>'''

    html = f'''
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600;700&display=swap" rel="stylesheet">

<div style="font-family:'Space Grotesk',-apple-system,BlinkMacSystemFont,sans-serif;max-width:1280px;margin:0 auto;background:#ffffff;border-radius:20px;overflow:hidden;box-shadow:0 10px 40px rgba(0,60,130,0.08),0 0 0 1px rgba(0,80,160,0.08);position:relative">

  <!-- SFONDI DECORATIVI LEGGERI -->
  <div style="position:absolute;top:-50px;right:-50px;width:250px;height:250px;background:radial-gradient(circle,rgba(0,140,220,0.06) 0%,transparent 70%);pointer-events:none"></div>
  <div style="position:absolute;bottom:-80px;left:-40px;width:300px;height:300px;background:radial-gradient(circle,rgba(0,100,200,0.04) 0%,transparent 70%);pointer-events:none"></div>

  <!-- HEADER -->
  <div style="padding:32px 40px;border-bottom:1px solid #e2e8f0;position:relative;z-index:1;background:#f8fafd">
    <div style="display:flex;justify-content:space-between;align-items:flex-start">
      <div>
        <div style="display:inline-block;padding:4px 12px;background:linear-gradient(90deg,#004c99 0%,#0077cc 100%);border-radius:6px;font-size:10px;font-weight:700;color:#ffffff;letter-spacing:2px;text-transform:uppercase;margin-bottom:12px">CONFIDENTIAL · KOLME GROUP</div>
        <div style="font-size:32px;font-weight:700;color:#0b1e3d;letter-spacing:-0.5px;line-height:1">DASHBOARD INSOLUTI</div>
        <div style="font-size:14px;color:#2a4b7c;margin-top:8px;font-weight:500">Periodo di competenza: {PERIODO}</div>
      </div>
      <div style="text-align:right">
        <div style="font-size:11px;color:#5a6e8c;text-transform:uppercase;letter-spacing:2px;font-weight:600">Collection Rate</div>
        <div style="font-size:42px;font-weight:700;background:linear-gradient(90deg,#0077cc 0%,#00a3e0 100%);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;font-family:'JetBrains Mono',monospace;line-height:1;margin-top:4px">{fmt_pct(cr_val)}</div>
        <div style="font-size:11px;color:#5a6e8c;margin-top:4px">sul valore · {fmt_pct(cr_num)} sul numero</div>
      </div>
    </div>
  </div>

  <!-- KPI PORTFOLIO -->
  <div style="padding:28px 40px;position:relative;z-index:1">
    <div style="font-size:11px;color:#1e3a5f;text-transform:uppercase;letter-spacing:3px;font-weight:700;margin-bottom:16px">◆ Portfolio &amp; Cash Flow</div>
    <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:16px">
      <div style="background:#fff5f5;border:1px solid #fecaca;border-radius:14px;padding:20px">
        <div style="font-size:10px;color:#b91c1c;text-transform:uppercase;letter-spacing:1.5px;font-weight:600;margin-bottom:8px">Totale Insoluto</div>
        <div style="font-size:26px;font-weight:700;color:#0b1e3d;font-family:'JetBrains Mono',monospace">{fmt_euro(tot_insoluto)}</div>
        <div style="font-size:11px;color:#c62828;margin-top:4px;font-weight:500">{n_non_saldate} posizioni aperte</div>
      </div>
      <div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:14px;padding:20px">
        <div style="font-size:10px;color:#166534;text-transform:uppercase;letter-spacing:1.5px;font-weight:600;margin-bottom:8px">Totale Incassato</div>
        <div style="font-size:26px;font-weight:700;color:#0b1e3d;font-family:'JetBrains Mono',monospace">{fmt_euro(tot_incassato)}</div>
        <div style="font-size:11px;color:#15803d;margin-top:4px;font-weight:500">{n_saldate} posizioni saldate</div>
      </div>
      <div style="background:#eff6ff;border:1px solid #bfdbfe;border-radius:14px;padding:20px">
        <div style="font-size:10px;color:#1e3a8a;text-transform:uppercase;letter-spacing:1.5px;font-weight:600;margin-bottom:8px">Totale Scaduto</div>
        <div style="font-size:26px;font-weight:700;color:#0b1e3d;font-family:'JetBrains Mono',monospace">{fmt_euro(tot_scaduto)}</div>
        <div style="font-size:11px;color:#2563eb;margin-top:4px;font-weight:500">{n_posizioni} posizioni totali</div>
      </div>
      <div style="background:#f5f3ff;border:1px solid #ddd6fe;border-radius:14px;padding:20px">
        <div style="font-size:10px;color:#5b21b6;text-transform:uppercase;letter-spacing:1.5px;font-weight:600;margin-bottom:8px">Ticket Medio</div>
        <div style="font-size:26px;font-weight:700;color:#0b1e3d;font-family:'JetBrains Mono',monospace">{fmt_euro(ticket_medio)}</div>
        <div style="font-size:11px;color:#7c3aed;margin-top:4px;font-weight:500">Non saldato: {fmt_euro(ticket_medio_non_saldato)}</div>
      </div>
    </div>
  </div>

  <!-- KPI RISK MANAGEMENT -->
  <div style="padding:0 40px 28px;position:relative;z-index:1">
    <div style="font-size:11px;color:#1e3a5f;text-transform:uppercase;letter-spacing:3px;font-weight:700;margin-bottom:16px">◆ KPI &amp; Risk Management</div>
    <div style="display:grid;grid-template-columns:1.2fr 1fr 1fr;gap:16px">

      <!-- Donut -->
      <div style="background:#ffffff;border:1px solid #e2e8f0;border-radius:14px;padding:24px;display:flex;align-items:center;gap:24px">
        <svg width="180" height="180" viewBox="0 0 180 180">
          <defs>
            <linearGradient id="grad1" x1="0%" y1="0%" x2="100%" y2="100%">
              <stop offset="0%" style="stop-color:#15803d;stop-opacity:1"/>
              <stop offset="100%" style="stop-color:#0077cc;stop-opacity:1"/>
            </linearGradient>
          </defs>
          <circle cx="90" cy="90" r="{radius}" fill="none" stroke="#e2e8f0" stroke-width="16"/>
          <circle cx="90" cy="90" r="{radius}" fill="none" stroke="url(#grad1)" stroke-width="16"
            stroke-dasharray="{circumference:.2f}" stroke-dashoffset="{dash_offset:.2f}"
            transform="rotate(-90 90 90)" stroke-linecap="round"/>
          <text x="90" y="88" text-anchor="middle" fill="#0b1e3d" font-size="28" font-weight="700" font-family="JetBrains Mono">{cr_pct:.1f}%</text>
          <text x="90" y="108" text-anchor="middle" fill="#5a6e8c" font-size="10" font-weight="600" letter-spacing="2">RECUPERATO</text>
        </svg>
        <div>
          <div style="font-size:11px;color:#1e3a5f;text-transform:uppercase;letter-spacing:1.5px;font-weight:600;margin-bottom:10px">Collection Rate (Val.)</div>
          <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px">
            <div style="width:10px;height:10px;background:linear-gradient(90deg,#15803d,#0077cc);border-radius:2px"></div>
            <div style="font-size:13px;color:#334155;font-weight:500">Incassato</div>
            <div style="font-size:13px;color:#0b1e3d;font-weight:700;font-family:'JetBrains Mono',monospace;margin-left:auto">{fmt_euro(tot_incassato)}</div>
          </div>
          <div style="display:flex;align-items:center;gap:8px">
            <div style="width:10px;height:10px;background:#c62828;border-radius:2px"></div>
            <div style="font-size:13px;color:#334155;font-weight:500">Residuo</div>
            <div style="font-size:13px;color:#0b1e3d;font-weight:700;font-family:'JetBrains Mono',monospace;margin-left:auto">{fmt_euro(tot_insoluto)}</div>
          </div>
        </div>
      </div>

      <!-- Fido Medio -->
      <div style="background:#ffffff;border:1px solid #e2e8f0;border-radius:14px;padding:20px">
        <div style="font-size:11px;color:#1e3a5f;text-transform:uppercase;letter-spacing:1.5px;font-weight:600;margin-bottom:4px">Fido Medio</div>
        <div style="font-size:9px;color:#5a6e8c;margin-bottom:16px;text-transform:uppercase;letter-spacing:1px">Saldati vs Non Saldati</div>
        <div style="display:flex;align-items:flex-end;justify-content:space-around;height:120px;gap:20px;padding:0 12px">
          <div style="flex:1;display:flex;flex-direction:column;align-items:center;gap:6px">
            <div style="font-size:11px;color:#0b1e3d;font-weight:700;font-family:'JetBrains Mono',monospace">{fmt_euro(fido_med_saldati)}</div>
            <div style="width:100%;height:{fido_sal_h:.1f}px;background:linear-gradient(180deg,#15803d 0%,#22c55e 100%);border-radius:6px 6px 0 0;"></div>
            <div style="font-size:10px;color:#15803d;font-weight:600;text-transform:uppercase;letter-spacing:1px">Saldati</div>
          </div>
          <div style="flex:1;display:flex;flex-direction:column;align-items:center;gap:6px">
            <div style="font-size:11px;color:#0b1e3d;font-weight:700;font-family:'JetBrains Mono',monospace">{fmt_euro(fido_med_non_saldati)}</div>
            <div style="width:100%;height:{fido_non_sal_h:.1f}px;background:linear-gradient(180deg,#c62828 0%,#ef4444 100%);border-radius:6px 6px 0 0;"></div>
            <div style="font-size:10px;color:#c62828;font-weight:600;text-transform:uppercase;letter-spacing:1px">Non Sald.</div>
          </div>
        </div>
        <div style="margin-top:16px;padding-top:12px;border-top:1px solid #e2e8f0;text-align:center">
          <span style="font-size:10px;color:#5a6e8c;text-transform:uppercase;letter-spacing:1px">Scostamento</span>
          <span style="font-size:14px;color:#0077cc;font-weight:700;margin-left:8px;font-family:'JetBrains Mono',monospace">+{fmt_pct(scost_fido)}</span>
        </div>
      </div>

      <!-- Fatturato Medio -->
      <div style="background:#ffffff;border:1px solid #e2e8f0;border-radius:14px;padding:20px">
        <div style="font-size:11px;color:#1e3a5f;text-transform:uppercase;letter-spacing:1.5px;font-weight:600;margin-bottom:4px">Fatturato Medio 3M</div>
        <div style="font-size:9px;color:#5a6e8c;margin-bottom:16px;text-transform:uppercase;letter-spacing:1px">Saldati vs Non Saldati</div>
        <div style="display:flex;align-items:flex-end;justify-content:space-around;height:120px;gap:20px;padding:0 12px">
          <div style="flex:1;display:flex;flex-direction:column;align-items:center;gap:6px">
            <div style="font-size:11px;color:#0b1e3d;font-weight:700;font-family:'JetBrains Mono',monospace">{fmt_euro(fatt_med_saldati)}</div>
            <div style="width:100%;height:{fatt_sal_h:.1f}px;background:linear-gradient(180deg,#15803d 0%,#22c55e 100%);border-radius:6px 6px 0 0;"></div>
            <div style="font-size:10px;color:#15803d;font-weight:600;text-transform:uppercase;letter-spacing:1px">Saldati</div>
          </div>
          <div style="flex:1;display:flex;flex-direction:column;align-items:center;gap:6px">
            <div style="font-size:11px;color:#0b1e3d;font-weight:700;font-family:'JetBrains Mono',monospace">{fmt_euro(fatt_med_non_saldati)}</div>
            <div style="width:100%;height:{fatt_non_sal_h:.1f}px;background:linear-gradient(180deg,#c62828 0%,#ef4444 100%);border-radius:6px 6px 0 0;"></div>
            <div style="font-size:10px;color:#c62828;font-weight:600;text-transform:uppercase;letter-spacing:1px">Non Sald.</div>
          </div>
        </div>
        <div style="margin-top:16px;padding-top:12px;border-top:1px solid #e2e8f0;text-align:center">
          <span style="font-size:10px;color:#5a6e8c;text-transform:uppercase;letter-spacing:1px">Scostamento</span>
          <span style="font-size:14px;color:#0077cc;font-weight:700;margin-left:8px;font-family:'JetBrains Mono',monospace">+{fmt_pct(scost_fatt)}</span>
        </div>
      </div>
    </div>
  </div>

  <!-- COMPOSIZIONE PORTAFOGLIO -->
  <div style="padding:0 40px 28px;position:relative;z-index:1">
    <div style="font-size:11px;color:#1e3a5f;text-transform:uppercase;letter-spacing:3px;font-weight:700;margin-bottom:16px">◆ Composizione Portafoglio</div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">

      <div style="background:#ffffff;border:1px solid #e2e8f0;border-radius:14px;padding:20px">
        <div style="font-size:11px;color:#1e3a5f;text-transform:uppercase;letter-spacing:1.5px;font-weight:600;margin-bottom:4px">SDD su Totale Scaduto</div>
        <div style="font-size:10px;color:#5a6e8c;margin-bottom:12px">Tutte le posizioni SDD / Totale scaduto</div>
        <div style="display:flex;align-items:baseline;gap:12px">
          <div style="font-size:36px;font-weight:700;color:#0077cc;font-family:'JetBrains Mono',monospace;letter-spacing:-1px">{fmt_pct(pct_sdd_scaduto)}</div>
          <div style="font-size:12px;color:#5a6e8c">{fmt_euro(sdd_tot_importo)} su {fmt_euro(tot_scaduto)}</div>
        </div>
        <div style="margin-top:12px;height:8px;background:#e2e8f0;border-radius:4px;overflow:hidden">
          <div style="width:{pct_sdd_scaduto*100:.1f}%;height:100%;background:linear-gradient(90deg,#0077cc,#00a3e0);border-radius:4px;"></div>
        </div>
      </div>

      <div style="background:#ffffff;border:1px solid #e2e8f0;border-radius:14px;padding:20px">
        <div style="font-size:11px;color:#1e3a5f;text-transform:uppercase;letter-spacing:1.5px;font-weight:600;margin-bottom:4px">Non-SDD su Totale Scaduto</div>
        <div style="font-size:10px;color:#5a6e8c;margin-bottom:12px">Tutte le posizioni non-SDD / Totale scaduto</div>
        <div style="display:flex;align-items:baseline;gap:12px">
          <div style="font-size:36px;font-weight:700;color:#e67e22;font-family:'JetBrains Mono',monospace;letter-spacing:-1px">{fmt_pct(pct_non_sdd_scaduto)}</div>
          <div style="font-size:12px;color:#5a6e8c">{fmt_euro(non_sdd_tot_importo)} su {fmt_euro(tot_scaduto)}</div>
        </div>
        <div style="margin-top:12px;height:8px;background:#e2e8f0;border-radius:4px;overflow:hidden">
          <div style="width:{pct_non_sdd_scaduto*100:.1f}%;height:100%;background:linear-gradient(90deg,#e67e22,#f59e0b);border-radius:4px;"></div>
        </div>
      </div>
    </div>
  </div>

  <!-- TOP 5 DEBITORI -->
  <div style="padding:0 40px 28px;position:relative;z-index:1">
    <div style="font-size:11px;color:#1e3a5f;text-transform:uppercase;letter-spacing:3px;font-weight:700;margin-bottom:16px">◆ Top 5 Debitori · Non Saldati</div>
    <div style="background:#ffffff;border:1px solid #e2e8f0;border-radius:14px;overflow:hidden">
      <table style="width:100%;border-collapse:collapse">
        <thead>
          <tr style="background:#f1f5f9;border-bottom:2px solid #cbd5e1">
            <th style="padding:12px;text-align:center;font-size:10px;color:#0b1e3d;font-weight:700;text-transform:uppercase;letter-spacing:1px">#</th>
            <th style="padding:12px;text-align:left;font-size:10px;color:#0b1e3d;font-weight:700;text-transform:uppercase;letter-spacing:1px">Codice</th>
            {agente_th}
            <th style="padding:12px;text-align:right;font-size:10px;color:#0b1e3d;font-weight:700;text-transform:uppercase;letter-spacing:1px">Importo</th>
            <th style="padding:12px;text-align:right;font-size:10px;color:#0b1e3d;font-weight:700;text-transform:uppercase;letter-spacing:1px">Fido</th>
            <th style="padding:12px;text-align:right;font-size:10px;color:#0b1e3d;font-weight:700;text-transform:uppercase;letter-spacing:1px">Fatt. 3M</th>
            <th style="padding:12px;text-align:right;font-size:10px;color:#0b1e3d;font-weight:700;text-transform:uppercase;letter-spacing:1px">Ins/Fido</th>
          </tr>
        </thead>
        <tbody>{top_rows}</tbody>
      </table>
    </div>
  </div>

  <!-- ANALISI AREA + CONDIZIONE -->
  <div style="padding:0 40px 28px;position:relative;z-index:1;display:grid;grid-template-columns:1.3fr 1fr;gap:16px">
    <div>
      <div style="font-size:11px;color:#1e3a5f;text-transform:uppercase;letter-spacing:3px;font-weight:700;margin-bottom:16px">◆ Collection Rate per Area</div>
      <div style="background:#ffffff;border:1px solid #e2e8f0;border-radius:14px;overflow:hidden">
        <table style="width:100%;border-collapse:collapse">
          <thead>
            <tr style="background:#f1f5f9;border-bottom:2px solid #cbd5e1">
              <th style="padding:10px;text-align:left;font-size:10px;color:#0b1e3d;font-weight:700;text-transform:uppercase;letter-spacing:1px">Area</th>
              <th style="padding:10px;text-align:right;font-size:10px;color:#0b1e3d;font-weight:700;text-transform:uppercase">Pos.</th>
              <th style="padding:10px;text-align:right;font-size:10px;color:#0b1e3d;font-weight:700;text-transform:uppercase">Sal.</th>
              <th style="padding:10px;text-align:right;font-size:10px;color:#0b1e3d;font-weight:700;text-transform:uppercase">N.Sal.</th>
              <th style="padding:10px;text-align:right;font-size:10px;color:#0b1e3d;font-weight:700;text-transform:uppercase">CR Num</th>
              <th style="padding:10px;text-align:right;font-size:10px;color:#0b1e3d;font-weight:700;text-transform:uppercase">CR Val</th>
              <th style="padding:10px;text-align:right;font-size:10px;color:#0b1e3d;font-weight:700;text-transform:uppercase">Imp.N.Sal</th>
            </tr>
          </thead>
          <tbody>{area_rows}</tbody>
        </table>
      </div>
    </div>
    <div>
      <div style="font-size:11px;color:#1e3a5f;text-transform:uppercase;letter-spacing:3px;font-weight:700;margin-bottom:16px">◆ Per Condizione Pagamento</div>
      <div style="background:#ffffff;border:1px solid #e2e8f0;border-radius:14px;overflow:hidden">
        <table style="width:100%;border-collapse:collapse">
          <thead>
            <tr style="background:#f1f5f9;border-bottom:2px solid #cbd5e1">
              <th style="padding:10px;text-align:left;font-size:10px;color:#0b1e3d;font-weight:700;text-transform:uppercase">Cond.</th>
              <th style="padding:10px;text-align:right;font-size:10px;color:#0b1e3d;font-weight:700;text-transform:uppercase">Pos.</th>
              <th style="padding:10px;text-align:right;font-size:10px;color:#0b1e3d;font-weight:700;text-transform:uppercase">Sal.</th>
              <th style="padding:10px;text-align:right;font-size:10px;color:#0b1e3d;font-weight:700;text-transform:uppercase">N.Sal.</th>
              <th style="padding:10px;text-align:right;font-size:10px;color:#0b1e3d;font-weight:700;text-transform:uppercase">CR Val</th>
            </tr>
          </thead>
          <tbody>{cond_rows}</tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- FOOTER -->
  <div style="padding:20px 40px;background:#f8fafd;border-top:1px solid #e2e8f0;text-align:center;position:relative;z-index:1">
    <span style="font-size:10px;color:#5a6e8c;letter-spacing:2px;text-transform:uppercase;font-weight:600">Dashboard generata automaticamente · Kolme Group · Sales Administration</span>
  </div>
</div>'''
    return html

dashboard_html = build_dashboard()
if HAS_IPYTHON:
    display(HTML(dashboard_html))

# %% EXPORT FILE (aggiornato con sfondo bianco) ─────────────────

full_html = f'''<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="UTF-8">
<title>Dashboard Insoluti · {PERIODO}</title>
<style>
  body {{ background:#f8fafc; padding:40px 20px; margin:0; min-height:100vh; }}
  * {{ box-sizing:border-box; }}
</style>
</head>
<body>
{dashboard_html}
</body>
</html>'''

output_filename = f"dashboard_insoluti_{PERIODO.lower().replace(' ', '_')}.html"
with open(output_filename, "w", encoding="utf-8") as f:
    f.write(full_html)

print(f"\n✓ Dashboard salvata: {output_filename}")
print("📥 Download in corso...")
files.download(output_filename)
