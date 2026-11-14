#!/usr/bin/env python3
"""
Nome: Generatore Progressivo di Numeri
Descrizione: Uno script di esempio che simula un'operazione complessa o di lunga durata, utile per testare lo streaming in tempo reale dell'output.
Categoria: Utility di Test
"""

import time
import sys

print("🚀 Inizio dell'elaborazione complessa...")
sys.stdout.flush()

for i in range(1, 11):
    time.sleep(1)
    percentuale = i * 10
    print(f"📊 Avanzamento: {percentuale}% completato (elaborazione blocco {i}/10)...")
    # Facciamo il flush esplicito per assicurarci che l'output venga trasmesso subito
    sys.stdout.flush()

print("🔍 Controllo integrità dati in corso...")
time.sleep(1.5)
sys.stdout.flush()

print("✨ Operazione completata con successo!")
sys.stdout.flush()
