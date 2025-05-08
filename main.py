import os
import sys
import json
import importlib.util
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from PIL import Image, ImageTk  # type: ignore
import pandas as pd
from preventivo_class import Preventivo
import moduli.certificato_modulo_26 as certificato_modulo_26
import moduli.modulo_b1 as modulo_b1
import moduli.modulo_b2 as modulo_b2
import moduli.modulo_posizioni as modulo_posizioni
import moduli.modulo_telaio as modulo_telaio
import moduli.modulo_scansioni as modulo_scansioni

class GestionaleApp(tk.Tk):
    """Applicazione gestionale IFG SRL."""
    def __init__(self):
        super().__init__()
        # Configurazione della finestra principale
        self.title("Gestionale IFG SRL")
        self.geometry("1200x700")
        self.minsize(800, 600)
        # Imposta la finestra a schermo intero
        self.state('zoomed')
        # Inizializza il preventivo corrente
        self.preventivo_corrente = Preventivo()
        self.percorso_preventivo_corrente = None
        # Carica i dataframes
        self.dataframes = {}
        self._carica_dataframes()
        # Imposta l'icona dell'applicazione
        try:
            if os.path.exists("risorse/logo.png"):
                logo_img = Image.open("risorse/logo.png")
                logo_img = logo_img.resize((32, 32), Image.LANCZOS)
                logo_tk = ImageTk.PhotoImage(logo_img)
                self.iconphoto(True, logo_tk)
        except Exception as e:
            print(f"Errore nel caricamento dell'icona: {str(e)}")
        # Crea le directory necessarie se non esistono
        for folder in ("data", "moduli", "preventivi"):
            if not os.path.exists(folder):
                os.makedirs(folder)
        # Inizializza le variabili per i moduli
        self.moduli = {}
        self.moduli_disponibili = []
        # Inizializza il log
        self.log_entries = []
        # Crea il menu principale
        self._crea_menu()
        # Crea il frame principale
        self.main_frame = ttk.Frame(self)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        # Crea il notebook per i tab dei moduli
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill="both", expand=True)
        # Barra di stato
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w")
        self.status_bar.pack(side="bottom", fill="x")
        self.status_var.set(f"Pronto - {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        # Crea il tab di benvenuto
        self._crea_tab_benvenuto()
        # --- INTEGRAZIONE CERTIFICAZIONE CE ---
        self.certificazione_ce_tab = None  # Tab non caricato all'avvio
        # Aggiungi "Certificazione CE" al menu Moduli
        self.moduli_menu.add_command(
            label="Certificazione CE",
            command=self._carica_certificazione_ce_tab
        )
        # Configura l'evento di chiusura
        self.protocol("WM_DELETE_WINDOW", self._on_closing)
        # Aggiungi log di avvio
        self._aggiungi_log("Applicazione avviata")
        # Aggiorna l'ora ogni secondo
        self.update_time()
        self.after(1000, self._update_clock)

    def _update_clock(self):
        """Aggiorna l'orologio ogni secondo."""
        self.update_time()
        self.after_id = self.after(1000, self._update_clock)

    def _carica_dataframes(self):
        """Carica i dataframe da utilizzare in tutti i moduli."""
        try:
            # Verifica se il file dataframe.py esiste
            if os.path.exists("data/dataframe.py"):
                # Aggiungi la directory data al path di Python
                sys.path.insert(0, os.path.abspath("data"))
                # Importa il modulo dataframe
                import dataframe
                # Carica i dataframe specifici
                dataframe_names = [
                    'colore_infissi',
                    'scavalco_cerniere',
                    'telaio_prolungato',
                    'elementi',
                    'modello_grata_combinato',
                    'scontistica',
                    'listino'
                ]
                for df_name in dataframe_names:
                    if hasattr(dataframe, df_name):
                        self.dataframes[df_name] = getattr(dataframe, df_name)
                        # print(f"Dataframe '{df_name}' caricato con successo")
                    else:
                        # print(f"Dataframe '{df_name}' non trovato nel modulo dataframe.py")
                        # Crea un dataframe vuoto come fallback
                        self.dataframes[df_name] = pd.DataFrame()
            else:
                # print("File dataframe.py non trovato. Creazione di dataframe vuoti.")
                pass
        except Exception as e:
            print(f"Errore nel caricamento dei dataframe: {e}")

    def _nuovo_preventivo(self):
        """Crea un nuovo preventivo caricando i moduli necessari."""
        self.preventivo_corrente = Preventivo()
        self.percorso_preventivo_corrente = None
        self.moduli = {}
        # Rimuovi tutti i tab tranne quello di benvenuto (indice 0)
        while len(self.notebook.tabs()) > 1:
            self.notebook.forget(self.notebook.tabs()[1])
        self._carica_modulo("modulo_b1")
        self._carica_modulo("modulo_b2")
        self._carica_modulo("modulo_posizioni")
        self._carica_modulo("modulo_telaio")
        self.status_var.set("Nuovo preventivo creato")
        
    def _apri_preventivo(self):
        """
        Apre un preventivo esistente selezionato dall'utente.

        Questa funzione permette di gestire l'apertura e il caricamento di un preventivo salvato in precedenza.

        Funzionalità principali:
        - Apre una finestra di dialogo per selezionare un file JSON di preventivo
        - Carica il file JSON selezionato nella directory "preventivi"
        - Crea un nuovo oggetto Preventivo e popola i suoi dati dal file caricato
        - Resetta l'interfaccia utente rimuovendo tutti i tab esistenti tranne quello di benvenuto
        - Ricarica i moduli principali (B1, B2, Posizioni, Telaio) con i dati del preventivo caricato
        - Aggiorna la barra di stato con il nome del file caricato
        - Registra l'operazione di caricamento nel log

        Gestisce eventuali errori durante il caricamento, mostrando un messaggio di errore all'utente.

        Returns:
            None: Carica il preventivo selezionato o mostra un messaggio di errore
        """
        file_path = filedialog.askopenfilename(
            title="Apri Preventivo",
            filetypes=[("File Preventivo", "*.json")],
            initialdir="preventivi"
        )
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    dati_preventivo = json.load(f)
                self.preventivo_corrente = Preventivo()
                self.preventivo_corrente.from_dict(dati_preventivo)
                self.percorso_preventivo_corrente = file_path
                self.moduli = {}
                # Rimuovi tutti i tab tranne quello di benvenuto (indice 0)
                while len(self.notebook.tabs()) > 1:
                    self.notebook.forget(self.notebook.tabs()[1])
                self._carica_modulo("modulo_b1")
                self._carica_modulo("modulo_b2")
                self._carica_modulo("modulo_posizioni")
                self._carica_modulo("modulo_telaio")
                self.status_var.set(f"Preventivo caricato: {os.path.basename(file_path)}")
                self._aggiungi_log(f"preventivo caricato {os.path.basename(file_path)}")
            except Exception as e:
                messagebox.showerror("Errore", f"Impossibile caricare il preventivo: {str(e)}")
                self._aggiungi_log(f"Errore nel caricamento del preventivo: {str(e)}")

    def _salva_preventivo(self):
        """Salva il preventivo corrente."""
        if not hasattr(self, 'preventivo_corrente') or not self.preventivo_corrente:
            messagebox.showinfo("Informazione", "Nessun preventivo da salvare.")
            return
        # Salva i dati da tutti i moduli nel preventivo
        self.save_preventivo()
        # Costruisci nome file secondo le specifiche richieste
        # 1. Nessuno spazio dopo N.
        # 2. Spazi normali tra parole, nessun underscore
        # 3. Se riferimento_cliente è presente, aggiungi [RIF. <riferimento_cliente>]
        import re
        def sanitize_component(s):
            # Rimuovi solo caratteri non validi per Windows, lascia spazi, elimina virgolette doppie
            s = s.replace('"', '')
            return re.sub(r'[^a-zA-Z0-9 \-]', '', s)
        def get_val(dic, *keys):
            for k in keys:
                v = dic.get(k, '').strip()
                if v:
                    return v
            return ''
        dati_b2 = getattr(self.preventivo_corrente, 'dati_b2', {})
        dati_b1 = getattr(self.preventivo_corrente, 'dati_b1', {})
        numero_protocollo = sanitize_component(
            get_val(dati_b2, 'numero_protocollo', 'Numero_protocollo') or get_val(dati_b1, 'numero_protocollo', 'Numero_protocollo')
        )
        nome_cliente = sanitize_component(
            get_val(dati_b2, 'nome_cliente', 'Nome_cliente') or get_val(dati_b1, 'nome_cliente', 'Nome_cliente')
        )
        riferimento_cliente = sanitize_component(
            get_val(dati_b2, 'rif_cliente') or get_val(dati_b1, 'rif_cliente')
        )
        print(f"[DEBUG] numero_protocollo='{numero_protocollo}', nome_cliente='{nome_cliente}', riferimento_cliente='{riferimento_cliente}'")
        import logging
        logging.debug(f"[DEBUG] INPUTS: numero_protocollo='{numero_protocollo}', nome_cliente='{nome_cliente}', riferimento_cliente='{riferimento_cliente}'")
        nome_file = f"N.{numero_protocollo} {nome_cliente}".strip()
        # Aggiungi la parte RIF. solo se riferimento_cliente è valorizzato (non vuoto dopo strip)
        if riferimento_cliente.strip():
            nome_file += f" [RIF. {riferimento_cliente}]"
        nome_file = nome_file.strip() + ".json"
        percorso_dir = os.path.join(os.getcwd(), "preventivi")
        if not os.path.exists(percorso_dir):
            os.makedirs(percorso_dir)
        percorso_file = os.path.join(percorso_dir, nome_file)
        print(f"[DEBUG] Salvataggio preventivo in: {percorso_file}")
        logging.debug(f"[DEBUG] Salvataggio preventivo in: {percorso_file}")
        try:
            with open(percorso_file, 'w', encoding='utf-8') as f:
                json.dump(self.preventivo_corrente.to_dict(), f, ensure_ascii=False, indent=4)
            self.preventivo_corrente.modificato = False
            self.percorso_preventivo_corrente = percorso_file
            self._aggiungi_log(f"preventivo salvato {nome_file}")
            messagebox.showinfo("Preventivo Salvato", f"Preventivo salvato con successo come {nome_file}.")
            return True
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile salvare il preventivo: {str(e)}")
            self._aggiungi_log(f"Errore nel salvataggio del preventivo: {str(e)}")
            return False

    def _salva_preventivo_come(self):
        """Salva il preventivo corrente con un nuovo nome."""
        file_path = filedialog.asksaveasfilename(
            title="Salva Preventivo Come",
            defaultextension=".json",
            filetypes=[("File Preventivo", "*.json")],
            initialdir="preventivi"
        )
        if file_path:
            try:
                self.save_preventivo()
                self.preventivo_corrente.file_salvataggio = file_path
                self.preventivo_corrente.auto_save()
                self.percorso_preventivo_corrente = file_path
                self.status_var.set(f"Preventivo salvato: {os.path.basename(file_path)}")
                return True
            except Exception as e:
                messagebox.showerror("Errore", f"Impossibile salvare il preventivo: {str(e)}")
                return False
        return False

    def save_preventivo(self):
        """Salva i dati da tutti i moduli nel preventivo."""
        # Modulo B1
        if "modulo_b1" in self.moduli:
            dati_b1 = self.moduli["modulo_b1"].get_data()
            self.preventivo_corrente.dati_b1 = dati_b1
            self.preventivo_corrente.cliente = dati_b1.get("nome_cliente", "")
            self.preventivo_corrente.email = dati_b1.get("email", "")
        # Modulo B2
        if "modulo_b2" in self.moduli:
            dati_b2 = self.moduli["modulo_b2"].get_data()
            self.preventivo_corrente.dati_b2 = dati_b2
        # Modulo posizioni
        if "modulo_posizioni" in self.moduli:
            self.preventivo_corrente.posizioni = self.moduli["modulo_posizioni"].get_all_posizioni()
        # Modulo telaio (se necessario, aggiungi qui la logica per salvare dati specifici)
        if "modulo_telaio" in self.moduli:
            if hasattr(self.moduli["modulo_telaio"], "get_data"):
                self.preventivo_corrente.dati_telaio = self.moduli["modulo_telaio"].get_data()
        # --- PATCH: Forza salvataggio dati modulo_b1 e modulo_b2 nel preventivo ---
        if hasattr(self.preventivo_corrente, 'dati_b1') and self.preventivo_corrente.dati_b1:
            self.preventivo_corrente.dati_b1_salvati = self.preventivo_corrente.dati_b1
        if hasattr(self.preventivo_corrente, 'dati_b2') and self.preventivo_corrente.dati_b2:
            self.preventivo_corrente.dati_b2_salvati = self.preventivo_corrente.dati_b2
        # --- FINE PATCH ---
        self.preventivo_corrente.modificato = False
        # Log dell'evento di modifica
        if self.percorso_preventivo_corrente:
            self._aggiungi_log(f"preventivo modificato {os.path.basename(self.percorso_preventivo_corrente)}")

    def _carica_modulo(self, nome_modulo):
        """Carica un modulo e lo aggiunge come tab nel notebook. Solo modulo_posizioni.py è accettato per le posizioni."""
        # Protezione: solo modulo_posizioni.py è accettato
        if nome_modulo.startswith("modulo_posizioni") and nome_modulo != "modulo_posizioni":
            raise ImportError(f"È consentito solo 'modulo_posizioni.py' come modulo delle posizioni. Trovato: {nome_modulo}")
        if nome_modulo == "modulo_b1":
            modulo = modulo_b1.ModuloB1Frame(self.notebook, preventivo=self.preventivo_corrente, app=self)
            self.moduli[nome_modulo] = modulo
            self.notebook.add(modulo, text="Anagrafica Clienti")
        elif nome_modulo == "modulo_b2":
            modulo = modulo_b2.ModuloB2Frame(self.notebook, preventivo=self.preventivo_corrente, app=self)
            self.moduli[nome_modulo] = modulo
            self.notebook.add(modulo, text="Dati Generali")
        elif nome_modulo == "modulo_posizioni":
            modulo = modulo_posizioni.PosizioniFrame(self.notebook, preventivo=self.preventivo_corrente, app=self)
            self.moduli[nome_modulo] = modulo
            self.notebook.add(modulo, text="Inserimento posizioni")
            if "modulo_b2" in self.moduli:
                modulo.collega_modulo_b2(self.moduli["modulo_b2"])
        elif nome_modulo == "modulo_telaio":
            modulo = modulo_telaio.ModuloTelaio(self.notebook, preventivo=self.preventivo_corrente, app=self)
            self.moduli[nome_modulo] = modulo
            self.notebook.add(modulo, text="Telaio")
        elif nome_modulo == "modulo_scansioni":
            modulo = modulo_scansioni.ModuloScansioni(self.notebook, preventivo=self.preventivo_corrente, app=self)
            self.moduli[nome_modulo] = modulo
            self.notebook.add(modulo, text="Scanner e WhatsApp")
        else:
            modulo = None
            try:
                modulo_path = f"moduli.{nome_modulo}"
                spec = importlib.util.find_spec(modulo_path)
                if spec is not None:
                    imported_modulo = importlib.util.module_from_spec(spec)
                    sys.modules[modulo_path] = imported_modulo
                    spec.loader.exec_module(imported_modulo)
                    if hasattr(imported_modulo, "Frame"):
                        modulo = imported_modulo.Frame(self.notebook, preventivo=self.preventivo_corrente, app=self)
                        self.moduli[nome_modulo] = modulo
                        self.notebook.add(modulo, text=nome_modulo)
            except Exception as e:
                print(f"Errore nel caricamento del modulo {nome_modulo}: {e}")
        return modulo

    def _carica_certificazione_ce_tab(self):
        """Carica la scheda Certificazione CE solo quando richiesto."""
        if self.certificazione_ce_tab is None or self.certificazione_ce_tab not in self.notebook.tabs():
            self.certificazione_ce_tab = certificato_modulo_26.get_certificazione_ce_tab(self.notebook)
            self.notebook.add(self.certificazione_ce_tab, text="Certificazione CE")
        self.notebook.select(self.certificazione_ce_tab)

    def _mostra_impostazioni(self):
        """Mostra la finestra delle impostazioni."""
        settings_window = tk.Toplevel(self)
        settings_window.title("Impostazioni")
        settings_window.geometry("500x400")
        settings_window.grab_set()
        
        # Frame principale
        main_frame = ttk.Frame(settings_window, padding=10)
        main_frame.pack(fill="both", expand=True)
        
        # Notebook per le diverse categorie di impostazioni
        settings_notebook = ttk.Notebook(main_frame)
        settings_notebook.pack(fill="both", expand=True)
        
        # Tab 
        general_frame = ttk.Frame(settings_notebook, padding=10)
        settings_notebook.add(general_frame, text="Generale")
        
        # Impostazioni generali
        ttk.Label(general_frame, text="Tema:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        tema_var = tk.StringVar(value="Chiaro")
        tema_combo = ttk.Combobox(general_frame, textvariable=tema_var, values=["Chiaro", "Scuro"], state="readonly")
        tema_combo.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        ttk.Label(general_frame, text="Lingua:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        lingua_var = tk.StringVar(value="Italiano")
        lingua_combo = ttk.Combobox(general_frame, textvariable=lingua_var, values=["Italiano", "English"], state="readonly")
        lingua_combo.grid(row=1, column=1, padx=5, pady=5, sticky="w")
         
        # Tab Database
        db_frame = ttk.Frame(settings_notebook, padding=10)
        settings_notebook.add(db_frame, text="Database")
        
        ttk.Label(db_frame, text="Percorso Database:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        db_path_var = tk.StringVar(value="./data/")
        db_path_entry = ttk.Entry(db_frame, textvariable=db_path_var, width=30)
        db_path_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Button(db_frame, text="Sfoglia...", command=lambda: None).grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(db_frame, text="Backup automatico:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        backup_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(db_frame, variable=backup_var).grid(row=1, column=1, padx=5, pady=5, sticky="w")
        
        # Tab Utente
        user_frame = ttk.Frame(settings_notebook, padding=10)
        settings_notebook.add(user_frame, text="Utente")
        
        ttk.Label(user_frame, text="Nome:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        nome_var = tk.StringVar()
        ttk.Entry(user_frame, textvariable=nome_var, width=30).grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        ttk.Label(user_frame, text="Email:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        email_var = tk.StringVar()
        ttk.Entry(user_frame, textvariable=email_var, width=30).grid(row=1, column=1, padx=5, pady=5, sticky="w")
        
        # Pulsanti di azione
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=10)
        
        ttk.Button(button_frame, text="Salva", command=lambda: settings_window.destroy()).pack(side="right", padx=5)
        ttk.Button(button_frame, text="Annulla", command=lambda: settings_window.destroy()).pack(side="right", padx=5)
        
        self._aggiungi_log("Aperta finestra impostazioni")

    def _mostra_guida(self):
        """Mostra la guida dell'applicazione."""
        help_window = tk.Toplevel(self)
        help_window.title("Guida")
        help_window.geometry("600x500")
        help_window.grab_set()
        
        # Frame principale
        main_frame = ttk.Frame(help_window, padding=10)
        main_frame.pack(fill="both", expand=True)
        
        # Titolo
        ttk.Label(
            main_frame, 
            text="Guida all'utilizzo del Gestionale",
            font=("Helvetica", 16, "bold")
        ).pack(pady=(0, 10))
        
        # Contenuto della guida
        help_text = tk.Text(main_frame, wrap="word", height=20)
        help_text.pack(fill="both", expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(help_text, orient="vertical", command=help_text.yview)
        scrollbar.pack(side="right", fill="y")
        help_text.configure(yscrollcommand=scrollbar.set)
        
        # Contenuto della guida
        guida_contenuto = """
        Benvenuto nella guida del Gestionale IFG SRL!
        
        Questa applicazione ti permette di gestire i preventivi per i tuoi clienti.
        
        Come iniziare:
        1. Crea un nuovo preventivo dal menu File -> Nuovo Preventivo
        2. Compila i dati anagrafici del cliente nel modulo B1
        3. Inserisci i dettagli tecnici nel modulo B2
        4. Aggiungi gli elementi nel modulo B3
        5. Visualizza il riepilogo nel modulo B4
        6. Gestisci le posizioni nel modulo POSIZIONI
        7. Salva o stampa il preventivo
        
        Per ulteriori informazioni, contatta il supporto tecnico.
        """
        
        help_text.insert("1.0", guida_contenuto)
        help_text.config(state="disabled")
        
        # Pulsante di chiusura
        ttk.Button(main_frame, text="Chiudi", command=help_window.destroy).pack(pady=10)
        
        self._aggiungi_log("Aperta finestra guida")

    def _mostra_info(self):
        """Mostra informazioni sull'applicazione."""
        info_window = tk.Toplevel(self)
        info_window.title("Informazioni")
        info_window.geometry("400x300")
        info_window.grab_set()
        
        # Frame principale
        main_frame = ttk.Frame(info_window, padding=10)
        main_frame.pack(fill="both", expand=True)
        
        # Logo
        try:
            if os.path.exists("risorse/logo.png"):
                # Carica l'immagine
                logo_img = Image.open("risorse/logo.png")
                # Ridimensiona l'immagine se necessario
                logo_img = logo_img.resize((100, 50), Image.LANCZOS)
                # Converti l'immagine in formato Tkinter
                logo_tk = ImageTk.PhotoImage(logo_img)
                # Crea un'etichetta per visualizzare l'immagine
                logo_label = ttk.Label(main_frame, image=logo_tk)
                logo_label.image = logo_tk  # Mantieni un riferimento per evitare il garbage collection
                logo_label.pack(side="left", padx=(0, 20))
            else:
                print("Logo non trovato. Assicurati che il file 'logo.png' sia presente nella cartella 'risorse'.")
        except Exception as e:
            print(f"Impossibile caricare il logo: {str(e)}")
        
        # Informazioni
        ttk.Label(
            main_frame, 
            text="Gestionale IFG SRL",
            font=("Helvetica", 14, "bold")
        ).pack(side="left", pady=20)
        
        ttk.Label(main_frame, text="Versione: 1.0.0").pack(pady=2)
        ttk.Label(main_frame, text=" 2025 IFG SRL").pack(pady=2)
        ttk.Label(main_frame, text="Tutti i diritti riservati").pack(pady=2)
        
        ttk.Separator(main_frame, orient="horizontal").pack(fill="x", pady=10)
        
        ttk.Label(
            main_frame, 
            text="Sviluppato da: Ing. Giuseppe Pisani",
            font=("Helvetica", 10)
        ).pack(pady=2)
        
        ttk.Label(
            main_frame, 
            text="Per assistenza: commercialeifg@infissidisicurezza.it",
            font=("Helvetica", 10)
        ).pack(pady=2)
        
        # Pulsante di chiusura
        ttk.Button(main_frame, text="Chiudi", command=info_window.destroy).pack(pady=5)
        
        self._aggiungi_log("Aperta finestra informazioni")

    def _on_closing(self):
        """Gestisce l'evento di chiusura dell'applicazione: chiede sempre di salvare il preventivo come file JSON, includendo i dati di tutti i moduli principali."""
        risposta = messagebox.askyesnocancel("Salva Preventivo", "Vuoi salvare il preventivo prima di uscire?")
        if risposta is None:
            return  # Annulla la chiusura
        if risposta:
            # Salva in automatico come fa il pulsante "Salva preventivo" del frame Inserimento nuova riga
            result = self._salva_preventivo()
            if not result:
                messagebox.showerror("Errore salvataggio", "Impossibile salvare il preventivo.")
        # Pulizia del preventivo.py (come prima)
        try:
            path_preventivo_py = os.path.join("risorse", "preventivo.py")
            if os.path.exists(path_preventivo_py):
                with open(path_preventivo_py, "r", encoding="utf-8") as f:
                    lines = f.readlines()
                with open(path_preventivo_py, "w", encoding="utf-8") as f:
                    for line in lines:
                        if line.strip().startswith("dati_posizioni"):
                            f.write("dati_posizioni: list[dict] = []\n")
                        else:
                            f.write(line)
        except Exception as e:
            print(f"Errore nella pulizia di preventivo.py: {e}")
        if hasattr(self, 'after_id'):
            self.after_cancel(self.after_id)
        self.destroy()

    def carica_preventivo_da_file(self):
        """Carica un preventivo da file JSON e aggiorna le posizioni correnti."""
        file_path = filedialog.askopenfilename(
            defaultextension=".json",
            filetypes=[("Preventivo files", "*.json"), ("All files", "*.*")],
            initialdir="preventivi"
        )
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    dati_preventivo = json.load(f)
                self.preventivo_corrente = Preventivo()
                self.preventivo_corrente.from_dict(dati_preventivo)
                self._aggiungi_log(f"Preventivo caricato da {file_path}")
                # Aggiorna la GUI se necessario (es. Treeview)
                # self.aggiorna_posizioni_treeview()
            except Exception as e:
                messagebox.showerror("Errore caricamento", f"Impossibile caricare il preventivo: {e}")

    def update_time(self):
        """Aggiorna l'ora corrente."""
        current_time = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        self.time_var.set(f"Data e ora: {current_time}")
        
    def _aggiungi_log(self, messaggio):
        """Aggiunge un'entrata al registro di log."""
        timestamp = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        log_entry = f"[{timestamp}] {messaggio}\n"
        self.log_entries.append(log_entry)
        
        # Aggiorna il widget di testo se esiste
        if hasattr(self, 'log_text'):
            self.log_text.config(state="normal")
            self.log_text.insert("end", log_entry)
            self.log_text.see("end")  # Scorre automaticamente alla fine
            self.log_text.config(state="disabled")

    def _crea_menu(self):
        """Crea il menu principale dell'applicazione"""
        self.menu_bar = tk.Menu(self)
        
        # Menu File
        file_menu = tk.Menu(self.menu_bar, tearoff=0)
        file_menu.add_command(label="Nuovo Preventivo", command=self._nuovo_preventivo)
        file_menu.add_command(label="Apri Preventivo", command=self._apri_preventivo)
        file_menu.add_command(label="Salva Preventivo", command=self._salva_preventivo)
        file_menu.add_command(label="Salva Preventivo Come...", command=self._salva_preventivo_come)
        file_menu.add_separator()
        file_menu.add_command(label="Esci", command=self._on_closing)
        self.menu_bar.add_cascade(label="File", menu=file_menu)
        
        # Menu Moduli
        self.moduli_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Moduli", menu=self.moduli_menu)
        
        # Aggiungi il modulo scansioni al menu
        self.moduli_menu.add_command(
            label="Scanner e WhatsApp",
            command=lambda: self._carica_modulo("modulo_scansioni")
        )
        
        # Menu Aiuto
        help_menu = tk.Menu(self.menu_bar, tearoff=0)
        help_menu.add_command(label="Guida", command=self._mostra_guida)
        help_menu.add_command(label="Informazioni", command=self._mostra_info)
        self.menu_bar.add_cascade(label="Aiuto", menu=help_menu)
        
        self.config(menu=self.menu_bar)

    def _crea_tab_benvenuto(self):
        """Crea il tab di benvenuto."""
        welcome_frame = ttk.Frame(self.notebook)
        self.notebook.add(welcome_frame, text="Home")
        
        # Header con logo
        header_frame = ttk.Frame(welcome_frame)
        header_frame.pack(pady=(20, 10), fill="x")
        
        # Carica e visualizza il logo
        try:
            if os.path.exists("risorse/logo.png"):
                # Carica l'immagine
                logo_img = Image.open("risorse/logo.png")
                # Ridimensiona l'immagine se necessario
                logo_img = logo_img.resize((150, 75), Image.LANCZOS)
                # Converti l'immagine in formato Tkinter
                logo_tk = ImageTk.PhotoImage(logo_img)
                # Crea un'etichetta per visualizzare l'immagine
                logo_label = ttk.Label(header_frame, image=logo_tk)
                logo_label.image = logo_tk  # Mantieni un riferimento per evitare il garbage collection
                logo_label.pack(side="left", padx=(0, 20))
            else:
                print("Logo non trovato. Assicurati che il file 'logo.png' sia presente nella cartella 'risorse'.")
        except Exception as e:
            print(f"Impossibile caricare il logo: {str(e)}")
        
        # Titolo dell'applicazione
        ttk.Label(
            header_frame, 
            text="Benvenuto nel Gestionale Aziendale",
            font=("Helvetica", 16, "bold")
        ).pack(side="left", pady=20)
        
        # Frame per il registro di log
        log_frame = ttk.LabelFrame(welcome_frame, text="Registro Attività")
        log_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Crea un widget Text per il log con scrollbar
        self.log_text = tk.Text(log_frame, wrap="word", height=10)
        self.log_text.pack(side="left", fill="both", expand=True)
        
        log_scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        log_scrollbar.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        # Informazioni di sistema
        info_frame = ttk.LabelFrame(welcome_frame, text="Informazioni di Sistema")
        info_frame.pack(fill="x", padx=20, pady=10)
        
        # Aggiorna l'ora in tempo reale
        self.time_var = tk.StringVar()
        self.update_time()
        
        ttk.Label(info_frame, textvariable=self.time_var).pack(anchor="w", padx=10, pady=5)
        ttk.Label(info_frame, text=f"Versione: 1.0.0").pack(anchor="w", padx=10, pady=5)
        ttk.Label(info_frame, text=f"Python: {sys.version.split()[0]}").pack(anchor="w", padx=10, pady=5)

    def _esporta_pdf(self):
        """Esporta il preventivo in formato PDF."""
        messagebox.showinfo("Funzione non disponibile", "La funzione di esportazione in PDF non è ancora implementata.")

if __name__ == "__main__":
    # --- PATCH: sicurezza import preventivo ---
    try:
        from preventivo import dati_posizioni
        if not isinstance(dati_posizioni, list):
            dati_posizioni = []
    except Exception:
        dati_posizioni = []
    # -----------------------------------------
    app = GestionaleApp()
    app.mainloop()
