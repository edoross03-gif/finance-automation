"""
Nome: Convertitore Markdown -> Word
Categoria: Documenti
Descrizione: Converte un file Markdown (.md) in un documento Word (.docx) rifinito, con indice automatico e stili di codice avanzati.
Istruzioni:
1. Carica il file .md dal pannello di sinistra.
2. Seleziona il file .md dal menu a tendina.
3. Clicca su Esegui per avviare la conversione.
Il file .docx verrà generato nella cartella uploads/ e sarà scaricabile direttamente.
"""
import os
import sys
import subprocess
import argparse

def setup_dependencies():
    """Controlla e installa automaticamente le dipendenze richieste."""
    try:
        import pypandoc
    except ImportError:
        print("Installazione della libreria 'pypandoc' in corso...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pypandoc"])
        
    import pypandoc
    try:
        pypandoc.get_pandoc_version()
    except OSError:
        print("Motore Pandoc non trovato nel sistema. Download automatico in corso (potrebbe richiedere qualche secondo)...")
        pypandoc.download_pandoc()

# Verifica e installa le dipendenze prima di caricare la GUI
print("Verifica librerie necessarie...")
setup_dependencies()
import pypandoc

def run_conversion(input_file, output_file=None):
    if not output_file:
        output_file = os.path.splitext(input_file)[0] + ".docx"
        
    print(f"\nConversione in corso:\nDa: {input_file}\nA:  {output_file}\n")
    try:
        # Usa pypandoc per convertire. 
        # Aggiungo un indice (toc) e uno stile per eventuali blocchi di codice.
        pypandoc.convert_file(
            input_file, 
            'docx', 
            outputfile=output_file,
            extra_args=['--highlight-style=tango', '--toc']
        )
        print("✅ Successo! Il file è pronto.")
        return True, output_file
    except Exception as e:
        error_msg = f"Si è verificato un errore:\n{str(e)}"
        print(f"❌ {error_msg}")
        return False, error_msg

def main():
    parser = argparse.ArgumentParser(description="Convertitore Markdown -> DOCX")
    parser.add_argument("-i", "--input", help="Percorso del file .md di input")
    parser.add_argument("-o", "--output", help="Percorso del file .docx di output")
    args = parser.parse_args()

    if args.input:
        # Modalità a riga di comando (es. chiamata dal sito web o terminale)
        success, res = run_conversion(args.input, args.output)
        if not success:
            sys.exit(1)
    else:
        # Modalità Grafica (GUI) - Doppia modalità
        import tkinter as tk
        from tkinter import filedialog, messagebox
        
        root = tk.Tk()
        root.withdraw() # Nasconde la finestra principale vuota
        
        # Forza la finestra in primo piano (specifico per Mac)
        root.call('wm', 'attributes', '.', '-topmost', True)
        
        print("\n=== Convertitore Markdown -> DOCX ===")
        print("Si aprirà una finestra di dialogo. Seleziona il file .md da convertire.")
        
        input_file = filedialog.askopenfilename(
            title="Seleziona il file Markdown",
            filetypes=[("Markdown files", "*.md"), ("Testo", "*.txt"), ("Tutti i file", "*.*")]
        )
        
        if not input_file:
            print("Nessun file selezionato. Operazione annullata.")
            return
            
        success, res = run_conversion(input_file)
        if success:
            # Mostra un popup di successo
            messagebox.showinfo(
                "Successo", 
                f"File convertito correttamente!\n\nSalvato in:\n{res}"
            )
            # Su Mac, prova ad aprire il file Word automaticamente al termine
            if sys.platform == "darwin":
                os.system(f'open "{res}"')
        else:
            messagebox.showerror("Errore di Conversione", res)

if __name__ == "__main__":
    main()

