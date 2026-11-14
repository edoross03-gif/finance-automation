"""
Nome: Dso Report Automatico
Descrizione: Analizza il bilancio SAP e calcola il Days Sales Outstanding (DSO) e Days Payable Outstanding (DPO) su tre livelli di profondità.
Categoria: Utility

Istruzioni:
■ FILE DI INPUT RICHIESTI:
  1. File Excel del Bilancio SAP:
     • Struttura: Deve essere scaricato da "bilancio multidimensionale" selezionando la gerarchia esplosa da "esporta come" e nascondendo le voci di bilancio (lasciando solo i codici dei conti).
     • Colonne attese: Una colonna contenente i codici SAP (con pattern 2310/ o YCOA) e una colonna del saldo denominata "Saldo periodo" (o prima colonna numerica valida).

■ MODALITÀ D'USO:
  1. Carica il bilancio SAP in formato Excel.
  2. Compila i campi Nome Periodo, Data Fine (AAAA-MM-GG) e Saldo Clienti CEE.
  3. Clicca su Esegui per calcolare DSO/DPO e generare il report Excel e HTML.
"""


# --- Latest Version Code Cell ---
# @title REPORT DSO AUTOMATICO V2 💹

# ════════════════════════════════════════════════════════════════
#  DSO / DPO REPORT — KOLME FY26
#  Google Colab notebook
#
#  Istruzioni:
#  1. Esegui lo script (Runtime → Run All)
#  2. Carica fino a 3 file di bilancio SAP
#  3. Inserisci nome, data fine e Saldo Clienti per ciascun periodo
#  4. Per saltare i periodi rimanenti, scrivi 'skip' o annulla l'upload
# ════════════════════════════════════════════════════════════════

# %% CONFIGURAZIONE ─────────────────────────────────────────────

DATA_INIZIO = "2026-05-01"

ACCOUNT_MAP = {
    "conto1":     "2310/038",
    "conto2":    "YCOA00121000000A",
    "conto3":   "2310/0171",
    "conto4":   ["YCOA00211000000A","2310/0180","2310/0170","2310/0175","2310/0176"],
    "conto5":  "2310/087",
    "conto6":     "2310/090",
    "conto7":   "2310/0112",
    "conto8":    ["2310/091","2310/092","2310/093"],
    "conto9":  ["YCOA00420000000A","YCOA00423000000A","YCOA00426000000A"],
    "conto10":         "YCOA00121030000A",
    "conto11":    "YCOA00121030100A",
    "conto12":       "YCOA00121040000A",
}

# %% CARICAMENTO PERIODI ────────────────────────────────────────

from google.colab import files

def _parse_num(s):
    """Parser robusto: accetta 6065846.55 oppure 6.065.846,55 oppure 6065846,55"""
    s = str(s).strip()
    if '.' in s and ',' in s:
        return float(s.replace('.', '').replace(',', '.'))
    elif ',' in s:
        return float(s.replace(',', '.'))
    return float(s)

# Controllo parametri da riga di comando per modalità non interattiva (da dashboard)
import argparse
import os
import shutil

parser = argparse.ArgumentParser(description="Report DSO Automatico")
parser.add_argument("--bilancio", type=str, help="Path del file di bilancio SAP")
parser.add_argument("--periodo", type=str, help="Nome del periodo (es. 'Marzo 2026')")
parser.add_argument("--data-fine", type=str, help="Data fine (formato YYYY-MM-DD)")
parser.add_argument("--saldo-clienti", type=str, help="Valore del Saldo Clienti CEE")

args_parsed, _ = parser.parse_known_args()

PERIODI = []

if args_parsed.bilancio and args_parsed.periodo and args_parsed.data_fine and args_parsed.saldo_clienti:
    filename = args_parsed.bilancio
    if os.path.exists(filename):
        local_name = os.path.basename(filename)
        if filename != local_name:
            shutil.copy(filename, local_name)
        filename = local_name
        
    nome = args_parsed.periodo
    data_fine = args_parsed.data_fine
    saldo = _parse_num(args_parsed.saldo_clienti)
    
    PERIODI.append({
        "nome":          nome,
        "data_fine":     data_fine,
        "export":        filename,
        "saldo_clienti": saldo,
    })
    print(f"✅ Esecuzione non interattiva configurata: {nome}  |  fine: {data_fine}  |  saldo: €{saldo:,.2f}")

if not PERIODI:
    print("📊 CARICAMENTO PERIODI (max 3)\n")
    print("Per ogni periodo verrà chiesto di:")
    print("  1. Caricare il file di bilancio SAP")
    print("  2. Inserire nome, data fine, valore del Saldo Clienti")
    print("\nPer terminare prima dei 3 periodi: scrivi 'skip' o annulla l'upload.")
    
    for i in range(1, 4):
        print(f"\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
        print(f"  PERIODO {i} di 3")
        print(f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")

        if i > 1:
            risposta = input(f"\n  ➕ Aggiungere il periodo {i}? [Invio = sì / 'skip' = termina]: ").strip().lower()
            if risposta in ('skip', 's', 'no', 'n'):
                print(f"  ⏭️  Caricamento completato a {len(PERIODI)} periodi")
                break

        print(f"\n  📁 Carica il file di bilancio (annulla per saltare):")
        uploaded = files.upload()
        if not uploaded:
            print(f"  ⏭️  Nessun file caricato — caricamento terminato a {len(PERIODI)} periodi")
            break

        filename = list(uploaded.keys())[-1]
        print(f"\n  ✓ File caricato: {filename}")

        nome      = input(f"  📝 Nome del periodo (es. 'Marzo 2026'): ").strip()
        data_fine = input(f"  📅 Data fine (formato YYYY-MM-DD): ").strip()
        saldo_str = input(f"  💰 Saldo Clienti (es. 6065846.55 o 6.065.846,55): ").strip()
        saldo     = _parse_num(saldo_str)

        PERIODI.append({
            "nome":          nome,
            "data_fine":     data_fine,
            "export":        filename,
            "saldo_clienti": saldo,
        })
        print(f"\n  ✓ Periodo {i} configurato: {nome}  |  fine: {data_fine}  |  saldo: €{saldo:,.2f}")

print(f"\n{'═'*55}")
print(f"  {len(PERIODI)} periodi caricati, pronti per l'elaborazione")
print(f"{'═'*55}")

# %% PARSER + CALCULATOR ────────────────────────────────────────

import pandas as pd
from datetime import date

def load_sap(filepath):
    """
    Carica un export SAP rilevando automaticamente la struttura del file.
    Funziona con qualsiasi numero di colonne purché ci siano:
      - una colonna con i codici SAP (pattern 2310/ o YCOA)
      - una colonna con il saldo del periodo
    """
    df = pd.read_excel(filepath, header=0)

    # 1. Colonna codici: cerca quella con più match per 2310/ o YCOA
    code_col = None
    for col in df.columns:
        if df[col].astype(str).str.contains(r"2310/|YCOA", regex=True).sum() > 5:
            code_col = col
            break
    if code_col is None:
        raise ValueError(
            f"Nessuna colonna con codici SAP trovata in {filepath}.\n"
            f"Colonne disponibili: {list(df.columns)}"
        )

    # 2. Colonna saldo: priorità a quella chiamata 'Saldo periodo'
    #    (escludendo 'confronto' e 'differenza'); fallback: prima colonna numerica
    saldo_col = None
    for col in df.columns:
        if col == code_col: continue
        name = str(col).lower()
        if 'saldo' in name and 'confronto' not in name and 'differenza' not in name:
            saldo_col = col
            break
    if saldo_col is None:
        for col in df.columns:
            if col == code_col: continue
            if pd.to_numeric(df[col], errors='coerce').notna().sum() > 10:
                saldo_col = col
                break
    if saldo_col is None:
        raise ValueError(f"Nessuna colonna del saldo trovata in {filepath}.")

    df['_codice'] = df[code_col].astype(str).str.strip()
    df['_saldo']  = pd.to_numeric(df[saldo_col], errors='coerce').fillna(0.0)
    return dict(zip(df['_codice'], df['_saldo']))

def v(lookup, code):
    if isinstance(code, list):
        return sum(lookup.get(c, 0.0) for c in code)
    return lookup.get(code, 0.0)

def dso(crediti, ricavi, gg):
    return (crediti / ricavi * gg) if ricavi else 0

def pct_var(curr, prev):
    return ((curr - prev) / abs(prev)) if prev else None

def process(export_path, saldo_clienti_val, giorni, prev=None):
    s  = load_sap(export_path)
    am = ACCOUNT_MAP

    cr_g = abs(v(s, am["conto1"]))
    cr_i = abs(v(s, am["conto2"]))
    cr_s = abs(saldo_clienti_val)              # ← input manuale

    rev_g = abs(v(s, am["conto5"]))
    rev_i = abs(v(s, am["conto7"]))
    rev_s = abs(v(s, am["conto9"]))

    deb_g = abs(v(s, am["conto3"]))
    deb_i = deb_g
    deb_s = abs(v(s, am["conto4"]))

    cos_g = abs(v(s, am["conto6"]))
    cos_i = abs(v(s, am["conto8"]))
    cos_s = cos_i

    r = {
        "giorni": giorni,
        "cr_g": cr_g, "cr_i": cr_i, "cr_s": cr_s,
        "rev_g": rev_g, "rev_i": rev_i, "rev_s": rev_s,
        "deb_g": deb_g, "deb_i": deb_i, "deb_s": deb_s,
        "cos_g": cos_g, "cos_i": cos_i, "cos_s": cos_s,
        "dso_g": dso(cr_g, rev_g, giorni),
        "dso_i": dso(cr_i, rev_i, giorni),
        "dso_s": dso(cr_s, rev_s, giorni),
        "dpo_g": dso(deb_g, cos_g, giorni),
        "dpo_i": dso(deb_i, cos_i, giorni),
        "dpo_s": dso(deb_s, cos_s, giorni),
        "f_base":  abs(v(s, am["conto10"])),
        "f_spec":  abs(v(s, am["conto11"])),
        "f_leg":   abs(v(s, am["conto12"])),
    }
    r["d_g"] = r["dso_g"] - r["dpo_g"]
    r["d_i"] = r["dso_i"] - r["dpo_i"]
    r["d_s"] = r["dso_s"] - r["dpo_s"]

    if prev:
        for k in ["dso_g","dso_i","dso_s","dpo_g","dpo_i","dpo_s","d_g","d_i","d_s","f_base","f_spec","f_leg"]:
            r[f"{k}_%"] = pct_var(r[k], prev[k])
    return r

# %% ESECUZIONE ─────────────────────────────────────────────────

d0 = date.fromisoformat(DATA_INIZIO)
risultati = []

for p in PERIODI:
    gg = (date.fromisoformat(p["data_fine"]) - d0).days + 1
    m  = process(p["export"], p["saldo_clienti"], gg, risultati[-1] if risultati else None)
    m["nome"] = p["nome"]
    risultati.append(m)
    print(f"✓ {p['nome']} ({gg} gg) — DSO g:{m['dso_g']:.2f}  i:{m['dso_i']:.2f}  s:{m['dso_s']:.2f}")

print(f"\n{'═'*50}\n  {len(risultati)} periodi elaborati con successo\n{'═'*50}")

# %% DASHBOARD HTML ─────────────────────────────────────────────

try:
    from IPython.display import HTML, display
    HAS_IPYTHON = True
except ImportError:
    HAS_IPYTHON = False

def _f(val, dec=0):
    s = f"{val:,.{dec}f}"
    return s.replace(",", "\u2009").replace(".", ",")

def _m(val):
    if abs(val) >= 1e6: return f"{val/1e6:,.2f}M".replace(",", "\u2009").replace(".", ",")
    if abs(val) >= 1e3: return f"{val/1e3:,.1f}K".replace(",", "\u2009").replace(".", ",")
    return _f(val, 0)

def _d(val):
    return f"{val:.2f}".replace(".", ",")

def _badge(val):
    if val is None: return ""
    good = val < 0
    c  = "#059669" if good else "#dc2626"
    bg = "#ecfdf5" if good else "#fef2f2"
    arr = "▼" if val < 0 else "▲"
    pct = f"{val*100:+.1f}%".replace(".", ",")
    return f'<span style="font-size:10px;padding:2px 6px;border-radius:6px;background:{bg};color:{c};font-weight:600;white-space:nowrap">{arr}\u2009{pct}</span>'

def _row(r, lev, label, color, has_var):
    cr = r[f"cr_{lev}"]; rev = r[f"rev_{lev}"]; deb = r[f"deb_{lev}"]; cos = r[f"cos_{lev}"]
    d_ = r[f"dso_{lev}"]; p_ = r[f"dpo_{lev}"]; dl = r[f"d_{lev}"]
    f_key = {"g":"f_base","i":"f_spec","s":"f_leg"}[lev]; fv = r[f_key]
    dv  = _badge(r.get(f"dso_{lev}_%")) if has_var else ""
    pv  = _badge(r.get(f"dpo_{lev}_%")) if has_var else ""
    dlv = _badge(r.get(f"d_{lev}_%"))   if has_var else ""
    fvb = _badge(r.get(f"{f_key}_%"))   if has_var else ""
    return f'''<tr style="border-bottom:1px solid #f1f5f9">
<td style="padding:12px 16px;text-align:left;white-space:nowrap"><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:{color};margin-right:8px;vertical-align:middle"></span><span style="font-weight:600;color:#1e293b;font-size:13px">{label}</span></td>
<td style="padding:12px 8px;text-align:right;font-size:12px;color:#64748b">{_m(cr)}</td>
<td style="padding:12px 8px;text-align:right;font-size:12px;color:#64748b">{_m(rev)}</td>
<td style="padding:12px 8px;text-align:right;font-weight:700;font-size:15px;color:#0f172a">{_d(d_)} {dv}</td>
<td style="padding:12px 8px;text-align:right;font-size:12px;color:#64748b">{_m(deb)}</td>
<td style="padding:12px 8px;text-align:right;font-size:12px;color:#64748b">{_m(cos)}</td>
<td style="padding:12px 8px;text-align:right;font-weight:700;font-size:15px;color:#0f172a">{_d(p_)} {pv}</td>
<td style="padding:12px 8px;text-align:right;font-weight:700;font-size:15px;color:{"#059669" if dl > 0 else "#dc2626"}">{_d(dl)} {dlv}</td>
<td style="padding:12px 8px;text-align:right;font-size:12px;color:#64748b">{_m(fv)} {fvb}</td>
</tr>'''

def dashboard(res):
    last = res[-1]
    cards = ""
    for k, lab, sub, col in [("dso_g","DSO","Global","#2563eb"),("dpo_g","DPO","Global","#0d9488"),("d_g","DELTA","Global","#d97706")]:
        val = last[k]
        badge = _badge(last.get(f"{k}_%")) if f"{k}_%" in last else ""
        cards += f'''<div style="flex:1;background:white;padding:20px;border-radius:12px;border-left:4px solid {col};box-shadow:0 1px 3px rgba(0,0,0,0.06)">
<div style="font-size:11px;color:#94a3b8;text-transform:uppercase;letter-spacing:1px;font-weight:600">{lab}</div>
<div style="font-size:32px;font-weight:800;color:#0f172a;margin:4px 0;font-family:'DM Mono',monospace">{_d(val)}</div>
<div style="font-size:12px;color:#64748b">{sub} · {last["giorni"]} gg {badge}</div>
</div>'''

    periods = ""
    for i, r in enumerate(res):
        has_var = i > 0
        periods += f'''<div style="margin-top:20px">
<div style="display:flex;align-items:baseline;gap:12px;padding:12px 0;border-bottom:2px solid #e2e8f0">
<span style="font-size:15px;font-weight:700;color:#1e293b">{r["nome"]}</span>
<span style="font-size:12px;color:#94a3b8;font-weight:500">{r["giorni"]} giorni</span>
</div>
<table style="width:100%;border-collapse:collapse;margin-top:4px">
<tr style="border-bottom:1px solid #e2e8f0">
<th style="text-align:left;padding:8px 16px;font-size:10px;color:#94a3b8;font-weight:600;text-transform:uppercase;letter-spacing:.8px">Livello</th>
<th style="text-align:right;padding:8px;font-size:10px;color:#94a3b8;font-weight:600;text-transform:uppercase">Crediti</th>
<th style="text-align:right;padding:8px;font-size:10px;color:#94a3b8;font-weight:600;text-transform:uppercase">Ricavi</th>
<th style="text-align:right;padding:8px;font-size:10px;color:#94a3b8;font-weight:600;text-transform:uppercase">DSO</th>
<th style="text-align:right;padding:8px;font-size:10px;color:#94a3b8;font-weight:600;text-transform:uppercase">Debiti</th>
<th style="text-align:right;padding:8px;font-size:10px;color:#94a3b8;font-weight:600;text-transform:uppercase">Costi</th>
<th style="text-align:right;padding:8px;font-size:10px;color:#94a3b8;font-weight:600;text-transform:uppercase">DPO</th>
<th style="text-align:right;padding:8px;font-size:10px;color:#94a3b8;font-weight:600;text-transform:uppercase">Δ</th>
<th style="text-align:right;padding:8px;font-size:10px;color:#94a3b8;font-weight:600;text-transform:uppercase">F.Sval</th>
</tr>
{_row(r, "g", "GLOBAL",    "#2563eb", has_var)}
{_row(r, "i", "INTERNO",   "#0d9488", has_var)}
{_row(r, "s", "SPECIFICO", "#7c3aed", has_var)}
</table>
</div>'''

    return f'''<div style="font-family:'DM Sans',-apple-system,BlinkMacSystemFont,sans-serif;max-width:1100px;margin:0 auto;background:#f8fafc;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.08)">
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<div style="background:linear-gradient(135deg,#0f172a 0%,#1e3a5f 100%);padding:28px 32px;color:white">
<div style="display:flex;justify-content:space-between;align-items:center">
<div>
<div style="font-size:22px;font-weight:800;letter-spacing:-0.5px">DSO / DPO Report</div>
<div style="font-size:13px;opacity:0.6;margin-top:4px">Cash conversion cycle analysis</div>
</div>
<div style="text-align:right">
<div style="font-size:14px;font-weight:600">Kolme</div>
<div style="font-size:12px;opacity:0.6">Fiscal Year 2026</div>
</div>
</div>
</div>
<div style="padding:20px 24px;background:#f1f5f9;display:flex;gap:16px">
{cards}
</div>
<div style="padding:8px 24px 28px">
{periods}
</div>
<div style="padding:12px 24px;background:#f1f5f9;border-top:1px solid #e2e8f0;text-align:center">
<span style="font-size:11px;color:#94a3b8">Report generato automaticamente · Dati SAP Kolme Bilancio CEE</span>
</div>
</div>'''

if HAS_IPYTHON:
    display(HTML(dashboard(risultati)))

# %% EXPORT HTML ────────────────────────────────────────────────

from google.colab import files

with open("dso_report.html", "w", encoding="utf-8") as f:
    f.write("<!DOCTYPE html><html><head><meta charset='utf-8'></head><body style='background:#e2e8f0;padding:40px'>")
    f.write(dashboard(risultati))
    f.write("</body></html>")

files.download("dso_report.html")

# %% EXPORT EXCEL ───────────────────────────────────────────────

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "DSO Report"

LEVELS = [("GLOBAL","g","4472C4"),("INTERNO","i","70AD47"),("SPECIFICO","s","9933FF")]
HDR    = ["Livello","Crediti","Ricavi","DSO","Debiti","Costi","DPO","Δ","F.Sval"]

row = 1
for r in risultati:
    ws.cell(row, 1, f"{r['nome']} — {r['giorni']} giorni")
    ws.cell(row, 1).font = Font(bold=True, size=13, color="FFFFFF")
    ws.cell(row, 1).fill = PatternFill("solid", fgColor="1E3A5F")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
    row += 1

    for ci, h in enumerate(HDR, 1):
        c = ws.cell(row, ci, h)
        c.font = Font(bold=True, size=10, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2D3F50")
        c.alignment = Alignment(horizontal="center")
    row += 1

    for label, lev, color in LEVELS:
        f_key = {"g":"f_base","i":"f_spec","s":"f_leg"}[lev]
        vals = [label,
                round(r[f"cr_{lev}"]/1e6, 2),
                round(r[f"rev_{lev}"]/1e6, 2),
                round(r[f"dso_{lev}"], 2),
                round(r[f"deb_{lev}"]/1e6, 2),
                round(r[f"cos_{lev}"]/1e6, 2),
                round(r[f"dpo_{lev}"], 2),
                round(r[f"d_{lev}"], 2),
                round(r[f_key]/1e3, 1)]
        for ci, val in enumerate(vals, 1):
            c = ws.cell(row, ci, val)
            c.alignment = Alignment(horizontal="center" if ci > 1 else "left")
            if ci in (4, 7, 8):
                c.font = Font(bold=True)
        ws.cell(row, 1).font = Font(bold=True, color=color)
        row += 1
    row += 1

for i, w in enumerate([16,10,10,10,10,10,10,10,10], 1):
    ws.column_dimensions[get_column_letter(i)].width = w

wb.save("dso_report.xlsx")
files.download("dso_report.xlsx")
