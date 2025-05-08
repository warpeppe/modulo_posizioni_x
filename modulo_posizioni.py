import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os
import openpyxl
from data.elementi import elementi
from data.dataframe import listino, modello_grata_combinato, controtelaio
import logging
import re

# Inizializza logging debug su file
logging.basicConfig(
    filename="debug_sconti.log",
    level=logging.DEBUG,
    format="%(asctime)s %(levelname)s %(message)s"
)

# Variabili globali per i valori predefiniti
colore_infissi_generale = "STANDARD RAL"  # Valore predefinito per il colore
traverso_inferiore_generale = "ANTA A GIRO"  # Valore predefinito per anta a giro posizione
maniglia_ribassata_generale = "NO"  # Valore predefinito per maniglia ribassata


class PosizioniFrame(ttk.Frame):
    def __init__(self, parent=None, preventivo=None, app=None, *args, **kwargs):
        """
        Inizializza il modulo posizioni.
        :param parent: il widget padre in cui inserire il modulo
        :param preventivo: riferimento all'oggetto Preventivo centrale
        :param app: riferimento all'app principale
        """
        super().__init__(parent, *args, **kwargs)
        self.parent = parent if parent else self
        self.preventivo = preventivo
        self.app = app
        
        # Assegno direttamente il DataFrame "elementi" importato
        self.dataframe = elementi
        
        # Ottieni i valori unici dalla colonna "TIPOLOGIA COMPLETA DI APERTURA"
        self.serramento_values = sorted(self.dataframe["TIPOLOGIA COMPLETA DI APERTURA"].unique())
        
        # Carica i valori per il campo "Tipo telaio" dal file Excel
        self.telaio_values = self.load_telaio_values()
        
        # Definisci le colonne
        self.colonne = [
            "Pos.", "nr. pezzi", "Serramento", "Modello", "Modello grata combinato", 
            "Colore", "L (mm)", "H (mm)", "Tipo telaio", "BUNK", "Dmcp / Scp", 
            "Defender", "Dist.", "Tipo dist.", "L (mm) dist.", "H (mm) dist.", 
            "Tipologia controtelaio", "Anta a giro posizione", "M.rib.",
            "N.ANTE", "AP.", "TIP.", "DESCR.TIP", "N.CERN.", "XLAV", "MULT.",
            "COSTO MLT", "MIN.1", "MIN.2", "MIN.3", "MIN.4", "MIN.5",
            "P.SERR.", "N.STAFF.",
            # Nuove colonne dopo N.STAFF.
            "Sconto 1", "Sconto 2", "Sconto 3", "Sconto in decimali", "Dicitura sconto",
            "Prezzo_listino",
            # AGGIUNTA
            "MqR", "MIR", "Tabella_minimi", "Unita_di_misura", "Min_fatt_pz", "Mq_fatt_pz",
            "Mq_totali_fatt", "Ml_totali_fatt",
            # Nuove colonne calcolate automaticamente
            "Costo_scontato_Mq", "Prezzo_listino_unitario", "Costo_serramento_listino_posizione", "Costo_serramento_scontato_posizione",
            # Nuovi campi per distanziali/imbotti
            "Distanziali/Imbotti", "Tipo distanziali/imbotti", "Dicitura distanziale/imbotte", 
            "Ml Distanziali", "Ml Imbotti", "Colore dist/imb", "Tip. dist/imb",
            # Nuove colonne per distanziali/imbotti richieste
            "Ml Distanziali Standard Ral", "Ml Distanziali Effetto legno", "Ml Distanziali grezzo",
            "Ml imbotti Standard Ral", "Ml imbotti Effetto legno", "Ml imbotti grezzo",
            "Costo al Ml", "Costo List Dist", "Costo List Imb", "Somma Cost listino dist + imbotte",
            "Costo scontato somma dist/imb posizione", "N. distanziali a 3 lati", "Ml totali Dist/imb",
            # Nuove colonne per controtelai
            "Costo al ml controtelaio singolo", "Verifica controtelaio", "Tipologia ml/nr. Pezzi",
            "Fattore moltiplicatore ml/nr. Pezzi", "Costo listino per posizione controtelaio",
            "Costo scontato controtelaio posizione", "N. Controtelai singoli", "ML Controtelaio singolo",
            "Costo Controtelaio singolo", "N. Controtelai doppi", "ML Controtelaio doppio",
            "Costo Controtelaio doppio", "N. Controtelaio termico TIP A",
            "Costo listino controtelaio termico TIP A", "N. Controtelaio termico TIP B",
            "Costo listino controtelaio termico TIP B"
        ]
        
        # Configura il layout principale
        self.pack_propagate(False)  # Impedisce al frame di ridimensionarsi in base ai contenuti
        
        # ---------------------------
        # AREA DI INSERIMENTO
        # ---------------------------
        insert_frame = tk.LabelFrame(self, text="Inserimento nuova riga", padx=10, pady=10)
        insert_frame.pack(fill="x", padx=10, pady=10)
        
        # Pulsante per aggiungere la riga
        add_button_frame = tk.Frame(insert_frame)
        add_button_frame.pack(fill="x", pady=5)
        
        self.add_button = tk.Button(add_button_frame, text="Aggiungi Riga", command=self.aggiungi_riga, bg="green", fg="white", padx=10, pady=5)
        self.add_button.pack(side="left", padx=5)

        # --- Pulsante Salva Preventivo ---
        self.save_button = tk.Button(add_button_frame, text="Salva preventivo", command=self._salva_preventivo_menu, bg="blue", fg="white", padx=10, pady=5)
        self.save_button.pack(side="left", padx=5)
        
        # Creo il menu contestuale
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="Modifica riga", command=self.modifica_riga_selezionata)
        self.context_menu.add_command(label="Duplica riga", command=self.duplica_riga_selezionata, foreground="blue")
        self.context_menu.add_command(label="Elimina riga", command=self.elimina_riga_selezionata, foreground="red")
        
        # Prima riga: Pos, nr. pezzi, Serramento (ravvicinati al massimo)
        row1_frame = tk.Frame(insert_frame)
        row1_frame.pack(fill="x", pady=5)
        
        # Pos.
        pos_frame = tk.Frame(row1_frame)
        pos_frame.pack(side="left", padx=(0, 2))
        
        tk.Label(pos_frame, text="Pos.:").pack(anchor="w")
        self.pos_label = tk.Label(pos_frame, text="1", width=5, borderwidth=1, relief="solid")
        self.pos_label.pack(anchor="w", pady=2)
        
        # nr. pezzi
        nr_pezzi_frame = tk.Frame(row1_frame)
        nr_pezzi_frame.pack(side="left", padx=(0, 2))
        
        tk.Label(nr_pezzi_frame, text="nr. pz.:").pack(anchor="w")
        self.nr_pezzi_entry = tk.Entry(nr_pezzi_frame, width=5)
        self.nr_pezzi_entry.pack(anchor="w", pady=2)
        self.nr_pezzi_entry.bind('<Return>', lambda e: self.serramento_combobox.focus())
        
        # Serramento
        serramento_frame = tk.Frame(row1_frame)
        serramento_frame.pack(side="left", padx=(0, 2))
        
        tk.Label(serramento_frame, text="Serramento:").pack(anchor="w")
        self.serramento_combobox = ttk.Combobox(serramento_frame, width=25, values=self.serramento_values)
        self.serramento_combobox.pack(anchor="w", pady=2)
        self.serramento_combobox.bind("<<ComboboxSelected>>", self.aggiorna_campi_aggiuntivi)
        
        # Modello
        modello_frame = tk.Frame(row1_frame)
        modello_frame.pack(side="left", padx=(0, 2))
        
        tk.Label(modello_frame, text="Modello:").pack(anchor="w")
        self.modello_combobox = ttk.Combobox(modello_frame, width=15)
        self.modello_combobox.pack(anchor="w", pady=2)
        self.modello_combobox.bind("<<ComboboxSelected>>", self.on_modello_selected)
        
        # Modello grata combinato (inizialmente nascosto)
        self.modello_grata_frame = tk.Frame(row1_frame)
        self.modello_grata_frame.pack(side="left", padx=(0, 2))
        self.modello_grata_frame.pack_forget()  # Nascosto inizialmente
        
        tk.Label(self.modello_grata_frame, text="Modello grata combinato:").pack(anchor="w")
        self.modello_grata_combobox = ttk.Combobox(self.modello_grata_frame, width=15)
        self.modello_grata_combobox.pack(anchor="w", pady=2)
        self.modello_grata_combobox.set("DA DEFINIRE")  # Valore predefinito
        
        # Colore
        colore_frame = tk.Frame(row1_frame)
        colore_frame.pack(side="left", padx=(0, 2))
        
        tk.Label(colore_frame, text="Colore:").pack(anchor="w")
        self.colore_combobox = ttk.Combobox(colore_frame, width=15, values=["STANDARD RAL", "EFFETTO LEGNO", "GREZZO", "EXTRA MAZZETTA"])
        self.colore_combobox.pack(anchor="w", pady=2)
        self.colore_combobox.set(colore_infissi_generale)  # Valore predefinito
        
        # L (mm)
        l_mm_frame = tk.Frame(row1_frame)
        l_mm_frame.pack(side="left", padx=(0, 2))
        
        tk.Label(l_mm_frame, text="L (mm):").pack(anchor="w")
        self.l_mm_entry = tk.Entry(l_mm_frame, width=6)
        self.l_mm_entry.pack(anchor="w", pady=2)
        
        # H (mm)
        h_mm_frame = tk.Frame(row1_frame)
        h_mm_frame.pack(side="left", padx=(0, 2))
        
        tk.Label(h_mm_frame, text="H (mm):").pack(anchor="w")
        self.h_mm_entry = tk.Entry(h_mm_frame, width=6)
        self.h_mm_entry.pack(anchor="w", pady=2)
        
        # Seconda riga per altri campi
        row2_frame = tk.Frame(insert_frame)
        row2_frame.pack(fill="x", pady=5)
        
        # Tipo telaio
        tipo_telaio_frame = tk.Frame(row2_frame)
        tipo_telaio_frame.pack(side="left", padx=(0, 2))
        
        tk.Label(tipo_telaio_frame, text="Tipo telaio:").pack(anchor="w")
        self.tipo_telaio_combobox = ttk.Combobox(tipo_telaio_frame, width=15, values=self.telaio_values)
        self.tipo_telaio_combobox.pack(anchor="w", pady=2)
        # Valore predefinito da impostare in base al valore del campo TELAIO del Modulo "Nuovo preventivo - Dati Generali"
        
        # BUNK
        bunk_frame = tk.Frame(row2_frame)
        bunk_frame.pack(side="left", padx=(0, 2))
        
        tk.Label(bunk_frame, text="BUNK:").pack(anchor="w")
        self.bunk_combobox = ttk.Combobox(bunk_frame, width=5, values=["SI", "NO"])
        self.bunk_combobox.pack(anchor="w", pady=2)
        self.bunk_combobox.set("NO")  # Valore predefinito
        
        # Dmcp / Scp
        dmcp_frame = tk.Frame(row2_frame)
        dmcp_frame.pack(side="left", padx=(0, 2))
        
        tk.Label(dmcp_frame, text="Dmcp / Scp:").pack(anchor="w")
        self.dmcp_combobox = ttk.Combobox(dmcp_frame, width=30, values=[
            "DOPPIA MANIGLIA E CILINDRO PASSANTE", 
            "SOLO CILINDRO PASSANTE", 
            "SOLO MEZZO-CILINDRO INTERNO", 
            "SOLO MEZZO-CILINDRO ESTERNO", 
            "MANIGLIA E MEZZO-CILINDRO ESTERNO", 
            "SENZA MEZZO CILINDRO E MANIGLIA"
        ])
        self.dmcp_combobox.pack(anchor="w", pady=2)
        self.dmcp_combobox.bind("<<ComboboxSelected>>", self.on_dmcp_selected)
        
        # Defender
        defender_frame = tk.Frame(row2_frame)
        defender_frame.pack(side="left", padx=(0, 2))
        
        tk.Label(defender_frame, text="Defender:").pack(anchor="w")
        self.defender_combobox = ttk.Combobox(defender_frame, width=5, values=["SI", "NO"])
        self.defender_combobox.pack(anchor="w", pady=2)
        self.defender_combobox.set("NO")  # Valore predefinito
        
        # Dist.
        dist_frame = tk.Frame(row2_frame)
        dist_frame.pack(side="left", padx=(0, 2))
        
        tk.Label(dist_frame, text="Dist.:").pack(anchor="w")
        self.dist_combobox = ttk.Combobox(dist_frame, width=5, values=["SI", "NO"])
        self.dist_combobox.pack(anchor="w", pady=2)
        self.dist_combobox.set("NO")  # Valore predefinito
        self.dist_combobox.bind("<<ComboboxSelected>>", self.on_dist_selected)
        
        # Tipo dist. (inizialmente nascosto)
        self.tipo_dist_frame = tk.Frame(row2_frame)
        self.tipo_dist_frame.pack(side="left", padx=(0, 2))
        self.tipo_dist_frame.pack_forget()  # Nascosto inizialmente
        
        tk.Label(self.tipo_dist_frame, text="Tipo dist.:").pack(anchor="w")
        self.tipo_dist_combobox = ttk.Combobox(self.tipo_dist_frame, width=10, values=["SALDATO", "IMBOTTE"])
        self.tipo_dist_combobox.pack(anchor="w", pady=2)
        self.tipo_dist_combobox.bind("<<ComboboxSelected>>", self.on_tipo_dist_selected)
        
        # L (mm) distanziale (inizialmente nascosto)
        self.l_dist_frame = tk.Frame(row2_frame)
        self.l_dist_frame.pack(side="left", padx=(0, 2))
        self.l_dist_frame.pack_forget()  # Nascosto inizialmente
        
        tk.Label(self.l_dist_frame, text="L (mm) dist.:").pack(anchor="w")
        self.l_dist_entry = tk.Entry(self.l_dist_frame, width=3)
        self.l_dist_entry.pack(anchor="w", pady=2)
        
        # H (mm) distanziale (inizialmente nascosto)
        self.h_dist_frame = tk.Frame(row2_frame)
        self.h_dist_frame.pack(side="left", padx=(0, 2))
        self.h_dist_frame.pack_forget()  # Nascosto inizialmente
        
        tk.Label(self.h_dist_frame, text="H (mm) dist.:").pack(anchor="w")
        self.h_dist_entry = tk.Entry(self.h_dist_frame, width=3)
        self.h_dist_entry.pack(anchor="w", pady=2)
        
        # Terza riga per altri campi
        row3_frame = tk.Frame(insert_frame)
        row3_frame.pack(fill="x", pady=5)
        
        # Tipologia controtelaio
        tipo_controtelaio_frame = tk.Frame(row3_frame)
        tipo_controtelaio_frame.pack(side="left", padx=(0, 2))
        
        tk.Label(tipo_controtelaio_frame, text="Tipologia controtelaio:").pack(anchor="w")
        self.tipo_controtelaio_combobox = ttk.Combobox(tipo_controtelaio_frame, width=15, values=[
            "C. SINGOLO", "C. DOPPIO", "C. TERMICO TIP A", "C. TERMICO TIP B"
        ])
        self.tipo_controtelaio_combobox.pack(anchor="w", pady=2)
        
        # Anta a giro posizione
        anta_giro_frame = tk.Frame(row3_frame)
        anta_giro_frame.pack(side="left", padx=(0, 2))
        
        tk.Label(anta_giro_frame, text="Anta a giro posizione:").pack(anchor="w")
        self.anta_giro_combobox = ttk.Combobox(anta_giro_frame, width=25, values=[
            "ANTA A GIRO", "LIBERO A MARMO", "ANTA A GIRO CON BATTUTA A MARMO"
        ])
        self.anta_giro_combobox.pack(anchor="w", pady=2)
        self.anta_giro_combobox.set(traverso_inferiore_generale)  # Valore predefinito
        
        # M.rib.
        mrib_frame = tk.Frame(row3_frame)
        mrib_frame.pack(side="left", padx=(0, 2))
        
        tk.Label(mrib_frame, text="M.rib.:").pack(anchor="w")
        self.mrib_combobox = ttk.Combobox(mrib_frame, width=5, values=["SI", "NO"])
        self.mrib_combobox.pack(anchor="w", pady=2)
        self.mrib_combobox.set(maniglia_ribassata_generale)  # Valore predefinito
        
        # Quarta riga: campi aggiuntivi (NUMERO ANTE, APERTURA, ecc.)
        # Poiché abbiamo nascosto tutti i campi non modificabili, questi frame non verranno utilizzati
        # ma li creiamo comunque per mantenere la compatibilità con il resto del codice
        row4_frame = tk.Frame(insert_frame)
        
        # Frame per le etichette
        labels_frame = tk.Frame(row4_frame)
        
        # Frame per i valori
        values_frame = tk.Frame(row4_frame)
        
        # Non facciamo il pack dei frame perché non ci saranno elementi visibili
        
        # Definisco i campi con le relative etichette
        # Campi da nascondere (non verranno visualizzati ma saranno comunque creati come attributi)
        # Nascondiamo tutti i campi non modificabili dall'utente
        campi_nascosti = [
            "N.ANTE", "AP.", "TIP.", "DESCR.TIP", "N.CERN.", "XLAV", "MULT.", "COSTO MLT",
            "MIN.1", "MIN.2", "MIN.3", "MIN.4", "MIN.5", "P.SERR.", "N.STAFF.",
            "Sconto 1", "Sconto 2", "Sconto 3", "Sconto in decimali", "Dicitura sconto",
            "Prezzo_listino", "MqR", "MIR", "Tabella_minimi", "Unita_di_misura", 
            "Min_fatt_pz", "Mq_fatt_pz", "Mq_totali_fatt", "Ml_totali_fatt",
            "Costo_scontato_Mq", "Prezzo_listino_unitario", 
            "Costo_serramento_listino_posizione", "Costo_serramento_scontato_posizione",
            "Distanziali/Imbotti", "Tipo distanziali/imbotti", "Dicitura distanziale/imbotte", 
            "Ml Distanziali", "Ml Imbotti", "Colore dist/imb", "Tip. dist/imb"
        ]
        
        # Tutti i campi definiti (sia visibili che nascosti)
        field_defs = [
            ("N.ANTE", "numero_ante_label"),
            ("AP.", "apertura_label"),
            ("TIP.", "tipologia_label"),
            ("DESCR.TIP", "descrizione_tipologia_label"),
            ("N.CERN.", "n_cerniere_label"),
            ("XLAV", "extra_lavorazione_label"),
            ("MULT.", "multiplo_label"),
            ("COSTO MLT", "costo_multiplo_label"),
            ("MIN.1", "minimi_1_label"),
            ("MIN.2", "minimi_2_label"),
            ("MIN.3", "minimi_3_label"),
            ("MIN.4", "minimi_4_label"),
            ("MIN.5", "minimi_5_label"),
            ("P.SERR.", "presenza_serratura_label"),
            ("N.STAFF.", "numero_staffette_label"),
            ("Sconto 1", "sconto_1_label"),
            ("Sconto 2", "sconto_2_label"),
            ("Sconto 3", "sconto_3_label"),
            ("Sconto in decimali", "sconto_decimali_label"),
            ("Dicitura sconto", "dicitura_sconto_label"),
            ("Prezzo_listino", "prezzo_listino_label"),
            ("MqR", "mqr_label"),
            ("MIR", "mir_label"),
            ("Tabella_minimi", "tabella_minimi_label"),
            ("Unita_di_misura", "unita_di_misura_label"),
            ("Min_fatt_pz", "min_fatt_pz_label"),
            ("Mq_fatt_pz", "mq_fatt_pz_label"),
            ("Mq_totali_fatt", "mq_totali_fatt_label"),
            ("Ml_totali_fatt", "ml_totali_fatt_label"),
            # Nuove colonne calcolate automaticamente
            ("Costo_scontato_Mq", "costo_scontato_mq_label"),
            ("Prezzo_listino_unitario", "prezzo_listino_unitario_label"),
            ("Costo_serramento_listino_posizione", "costo_serramento_listino_posizione_label"),
            ("Costo_serramento_scontato_posizione", "costo_serramento_scontato_posizione_label"),
            # Nuovi campi per distanziali/imbotti
            ("Distanziali/Imbotti", "distanziali_imbotti_label"),
            ("Tipo distanziali/imbotti", "tipo_distanziali_imbotti_label"),
            ("Dicitura distanziale/imbotte", "dicitura_distanziale_imbotte_label"),
            ("Ml Distanziali", "ml_distanziali_label"),
            ("Ml Imbotti", "ml_imbotti_label"),
            ("Colore dist/imb", "colore_dist_imb_label"),
            ("Tip. dist/imb", "tip_dist_imb_label")
        ]
        
        # Creo le etichette e i campi
        for label_text, attr_name in field_defs:
            # Creo il campo non modificabile (sempre, anche se nascosto)
            label = tk.Label(values_frame, text="", width=8, borderwidth=1, relief="solid")
            
            # Lo assegno come attributo della classe (sempre)
            setattr(self, attr_name, label)
            
            # Visualizzo solo i campi che non sono nella lista dei nascosti
            if label_text not in campi_nascosti:
                # Creo l'etichetta solo per i campi visibili
                tk.Label(labels_frame, text=label_text, width=8).pack(side="left", padx=(0, 2))
                # Visualizzo il campo
                label.pack(side="left", padx=(0, 2))
        
        # ---------------------------
        # AREA DI VISUALIZZAZIONE
        # ---------------------------
        view_frame = tk.LabelFrame(self, text="Righe inserite", padx=10, pady=10)
        view_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Crea un frame per il Treeview e le scrollbar
        tree_frame = tk.Frame(view_frame)
        tree_frame.pack(fill="both", expand=True)
        
        # Crea le scrollbar
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        
        # Crea il Treeview
        self.tree = ttk.Treeview(tree_frame, columns=self.colonne, show="headings",
                                yscrollcommand=vsb.set)
        
        # Configura le scrollbar
        vsb.config(command=self.tree.yview)
        
        # Posiziona le scrollbar e il Treeview
        vsb.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)
        
        # Configura le intestazioni e le colonne
        # Definisci colonne che possono essere più strette
        colonne_strette = ["Pos.", "nr. pezzi", "L (mm)", "H (mm)", "Defender", "Dist.", "M.rib."]
        
        # Configura lo stile per le intestazioni
        style = ttk.Style()
        style.configure("Treeview.Heading", background="#FFF8E1", foreground="#b36b00")
        
        for col in self.colonne:
            self.tree.heading(col, text=col)
            # Imposta larghezza e minwidth diverse per colonne specifiche
            if col in colonne_strette:
                self.tree.column(col, width=22, anchor="center", minwidth=22, stretch=False)
            else:
                self.tree.column(col, width=90, anchor="center", minwidth=60, stretch=False)
        
        # Binding per il doppio click per modificare e per il tasto destro
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Button-3>", self.show_context_menu)
        
        # Frame per i pulsanti sotto il Treeview
        button_frame = tk.Frame(view_frame)
        button_frame.pack(fill="x", pady=5)
        
        # Inizializza i valori delle combobox
        self.update_combobox_values()
        
        # Aggiunta la logica per aggiornare dinamicamente la variabile globale colore_infissi_generale e la combobox 'Colore' in base alla selezione del campo 'Colore Infissi' da ModuloB2Frame
        self.aggiorna_colore_infissi_generale()
        
        # Dopo aver creato l'interfaccia, carica i dati dal preventivo se presenti
        self.aggiorna_da_preventivo()
        
        # --- PATCH: contatore posizione ---
        self.pos_counter = 1  # Contatore posizione, parte da 1
        # --- FINE PATCH ---
        
        self.bind_treeview_column_resize()
        
        # --- INIZIO INTESTAZIONE PERSONALIZZATA CONTROTELAI ---
        print("[DEBUG] Inizio creazione header di gruppo Controtelaio")
        
        try:
            # Indici delle colonne controtelai
            col_start = self.colonne.index("Costo al ml controtelaio singolo")
            col_end = self.colonne.index("Costo listino controtelaio termico TIP B")
            print(f"[DEBUG] Colonne trovate: start={col_start}, end={col_end}")
            print(f"[DEBUG] Nome colonna start: {self.colonne[col_start]}")
            print(f"[DEBUG] Nome colonna end: {self.colonne[col_end]}")
        except ValueError as e:
            print(f"[DEBUG] Errore nel trovare le colonne: {e}")
            print(f"[DEBUG] Colonne disponibili: {self.colonne}")
            return
        
        # --- CONFIGURAZIONE HEADER PERSONALIZZATI ---
        # Crea il frame per gli header
        self.header_frame = ttk.Frame(self)
        self.header_frame.pack(side="top", fill="x", before=self.tree)
        
        # Crea il canvas per gli header
        self.header_canvas = tk.Canvas(self.header_frame, height=30, bg="white", highlightthickness=0)
        self.header_canvas.pack(side="top", fill="x")
        
        # Definizione dei gruppi di colonne con i colori
        self.group_headers = {
            "Righe inserite": {
                "start_col": 1,
                "end_col": 19,
                "bg_color": "#F5F5F5",
                "text_color": "#000000"
            },
            "Calcoli": {
                "start_col": 20,
                "end_col": 52,
                "bg_color": "#E0E0E0",
                "text_color": "#000000"
            },
            "Distanziale/Imbotte": {
                "start_col": 53,
                "end_col": 72,
                "bg_color": "#C8E6C9",
                "text_color": "#000000"
            },
            "Controtelai": {
                "start_col": 73,
                "end_col": 88,
                "bg_color": "#FFD580",
                "text_color": "#000000"
            }
        }
        
        # Configura la scrollbar orizzontale del treeview
        self.tree_scrollbar = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree_scrollbar.pack(side="bottom", fill="x")
        
        # Configura la sincronizzazione dello scroll
        def on_treeview_scroll(*args):
            self.draw_custom_headers()
            self.tree_scrollbar.set(*args)
        
        # Associa l'evento di scroll del treeview
        self.tree.configure(xscrollcommand=on_treeview_scroll)
        
        # Aggiorna gli header quando il treeview viene ridimensionato o scrollato
        self.tree.bind('<Configure>', lambda e: self.draw_custom_headers())
        self.tree.bind('<B2-Motion>', lambda e: self.draw_custom_headers())  # Scrolling con il mouse
        self.tree.bind('<Button-4>', lambda e: self.draw_custom_headers())   # Scrolling con la rotella
        self.tree.bind('<Button-5>', lambda e: self.draw_custom_headers())   # Scrolling con la rotella
        self.tree.bind('<Left>', lambda e: self.draw_custom_headers())       # Tasto freccia sinistra
        self.tree.bind('<Right>', lambda e: self.draw_custom_headers())      # Tasto freccia destra
        
        # Aggiungi binding per il ridimensionamento delle colonne
        def on_column_resize(event):
            self.draw_custom_headers()
        
        # Associa l'evento di ridimensionamento a tutte le colonne
        for col in self.colonne:
            self.tree.heading(col, command=lambda c=col: None)  # Rimuovi il comando di ordinamento
            self.tree.column(col, width=100)  # Imposta una larghezza iniziale
            self.tree.column(col, stretch=True)  # Permetti lo stretching
            self.tree.bind(f'<Configure>', on_column_resize, add='+')
        
        print("[DEBUG] Binding eventi impostato")
        
        # Disegna gli header iniziali dopo un breve ritardo per assicurarsi che il treeview sia configurato
        self.after(100, self.draw_custom_headers)
        print("[DEBUG] Header iniziali programmati")
        # --- FINE INTESTAZIONE PERSONALIZZATA CONTROTELAI ---
        # Colora le intestazioni delle colonne controtelai
        style = ttk.Style()
        style.theme_use('default')
        for i, col in enumerate(self.colonne):
            if col_start <= i <= col_end:
                style.configure(f"Treeview.Heading.{col}", background="#FFF8E1", foreground="#b36b00")
                self.tree.heading(col, text=col)
        # Colora le celle delle colonne controtelai (solo visualizzazione, non editing)
        def tag_controtelai():
            for item in self.tree.get_children():
                for i in range(col_start, col_end+1):
                    self.tree.tag_configure('controtelai', background="#FFF8E1")
                    self.tree.item(item, tags=('controtelai',))
        self.tree.bind('<<TreeviewSelect>>', lambda e: tag_controtelai())
        tag_controtelai()
        
        # Rimuovi la scrollbar orizzontale dal frame "Righe inserite"
        for widget in self.winfo_children():
            if isinstance(widget, ttk.Frame):
                for child in widget.winfo_children():
                    if isinstance(child, ttk.Scrollbar) and child.cget('orient') == 'horizontal':
                        child.pack_forget()
        
        # Colora le celle delle colonne in base al gruppo di appartenenza
        # Definisci i gruppi e i loro colori
        group_column_ranges = {
            'righe_inserite': (self.group_headers['Righe inserite']['start_col']-1, self.group_headers['Righe inserite']['end_col']-1),
            'calcoli': (self.group_headers['Calcoli']['start_col']-1, self.group_headers['Calcoli']['end_col']-1),
            'distanziali': (self.group_headers['Distanziale/Imbotte']['start_col']-1, self.group_headers['Distanziale/Imbotte']['end_col']-1),
            'controtelai': (self.group_headers['Controtelai']['start_col']-1, self.group_headers['Controtelai']['end_col']-1),
        }
        group_colors = {
            'righe_inserite': self.group_headers['Righe inserite']['bg_color'],
            'calcoli': self.group_headers['Calcoli']['bg_color'],
            'distanziali': self.group_headers['Distanziale/Imbotte']['bg_color'],
            'controtelai': self.group_headers['Controtelai']['bg_color'],
        }
        # Configura i tag per ogni gruppo
        for group, color in group_colors.items():
            self.tree.tag_configure(group, background=color)
        
        def color_group_cells():
            for item in self.tree.get_children():
                values = self.tree.item(item, 'values')
                tags = []
                # Applica i tag ai gruppi di colonne
                for group, (start, end) in group_column_ranges.items():
                    # Se almeno una colonna del gruppo ha un valore, applica il tag
                    if any(values[i] for i in range(start, end+1) if i < len(values)):
                        tags.append(group)
                self.tree.item(item, tags=tags)
        self.tree.bind('<<TreeviewSelect>>', lambda e: color_group_cells())
        color_group_cells()
        
    def load_telaio_values(self):
        """Carica i valori per il campo 'Tipo telaio' dal file Excel."""
        try:
            # Percorso del file Excel
            excel_path = os.path.join("data", "database_gestionale.xlsx")
            
            # Carica il foglio "telaio"
            wb = openpyxl.load_workbook(excel_path)
            sheet = wb["telaio"]
            
            # Estrai i valori dalla colonna A (dalla riga 2 alla 130)
            values = []
            for row in range(2, 131):
                cell_value = sheet.cell(row=row, column=1).value
                if cell_value:
                    values.append(cell_value)
            
            return values
        except Exception as e:
            print(f"Errore nel caricamento dei valori del telaio: {e}")
            return []

    def update_combobox_values(self):
        """Aggiorna i valori delle combobox con i dati dai DataFrame."""
        # Serramento
        if self.dataframe is not None:
            # Preleva il nome della prima colonna, qualunque esso sia
            first_col = self.dataframe.columns[0]
            # Assegna alla combobox i valori univoci della prima colonna
            self.serramento_combobox['values'] = list(self.dataframe[first_col].unique())
        else:
            self.serramento_combobox['values'] = []
        
        # Modello
        try:
            if 'listino' in globals() and isinstance(listino, pd.DataFrame) and "MODELLO" in listino.columns:
                self.modello_combobox['values'] = list(listino["MODELLO"].unique())
            else:
                self.modello_combobox['values'] = []
        except Exception as e:
            print(f"Errore nell'aggiornamento dei valori del modello: {e}")
            self.modello_combobox['values'] = []
        
        # Modello grata combinato
        try:
            if 'modello_grata_combinato' in globals() and isinstance(modello_grata_combinato, pd.DataFrame):
                # Assumo che il DataFrame abbia una colonna con i modelli
                first_col = modello_grata_combinato.columns[0]
                self.modello_grata_combobox['values'] = list(modello_grata_combinato[first_col].unique())
            else:
                self.modello_grata_combobox['values'] = []
        except Exception as e:
            print(f"Errore nell'aggiornamento dei valori del modello grata combinato: {e}")
            self.modello_grata_combobox['values'] = []

    def on_modello_selected(self, event=None):
        """Gestisce la selezione del modello e mostra/nasconde il campo 'Modello grata combinato'."""
        modello_selezionato = self.modello_combobox.get()
        modelli_duo = [
            "DUO", "DUO MILLENNIUM", "DUO ALLUMINIO", "DUO REVOLUTION", 
            "DUO GRETHA", "DUO GRETHA ALLUMINIO", "DUO BLIND", 
            "DUO BLIND ORIENTABILE", "DUO BLIND ALLUMINIO", "DUO BLIND SLIM"
        ]
        
        if modello_selezionato in modelli_duo:
            # Mostra il campo "Modello grata combinato"
            self.modello_grata_frame.pack(side="left", padx=(0, 2))
        else:
            # Nascondi il campo "Modello grata combinato"
            self.modello_grata_frame.pack_forget()

    def on_dmcp_selected(self, event=None):
        """Gestisce la selezione del campo 'Dmcp / Scp' e imposta il valore del campo 'Defender'."""
        dmcp_selezionato = self.dmcp_combobox.get()
        
        if dmcp_selezionato in ["DOPPIA MANIGLIA E CILINDRO PASSANTE", "SOLO CILINDRO PASSANTE"]:
            self.defender_combobox.set("SI")
        else:
            self.defender_combobox.set("NO")

    def on_dist_selected(self, event=None):
        """Gestisce la selezione del campo 'Dist.' e mostra/nasconde il campo 'Tipo dist.'."""
        dist_selezionato = self.dist_combobox.get()
        
        # Aggiorna il campo "Distanziali/Imbotti" in base al valore di "Dist."
        # Se "Dist." è vuoto o "NO", "Distanziali/Imbotti" deve essere "0", altrimenti "1"
        distanziali_imbotti = "0"
        if dist_selezionato and dist_selezionato != "NO":
            distanziali_imbotti = "1"
        self.distanziali_imbotti_label.config(text=distanziali_imbotti)
        
        if dist_selezionato == "SI":
            # Mostra il campo "Tipo dist."
            self.tipo_dist_frame.pack(side="left", padx=(0, 2))
        else:
            # Nascondi il campo "Tipo dist." e i campi correlati
            self.tipo_dist_frame.pack_forget()
            self.l_dist_frame.pack_forget()
            self.h_dist_frame.pack_forget()
            
            # Resetta i valori
            self.tipo_dist_combobox.set("")
            self.l_dist_entry.delete(0, tk.END)
            self.h_dist_entry.delete(0, tk.END)

    def on_tipo_dist_selected(self, event=None):
        """Gestisce la selezione del campo 'Tipo dist.' e mostra/nasconde i campi correlati."""
        tipo_dist_selezionato = self.tipo_dist_combobox.get()
        
        # Calcola il valore di "Tipo distanziali/imbotti" in base al valore di "Tipo dist."
        # =SE([@[Tipo dist.]]="IMBOTTE";2;SE([@[Tipo dist.]]="SALDATO";3;1))
        tipo_distanziali_imbotti = "1"  # Valore predefinito
        if tipo_dist_selezionato == "IMBOTTE":
            tipo_distanziali_imbotti = "2"
        elif tipo_dist_selezionato == "SALDATO":
            tipo_distanziali_imbotti = "3"
        
        # Aggiorna il campo "Tipo distanziali/imbotti"
        self.tipo_distanziali_imbotti_label.config(text=tipo_distanziali_imbotti)
        
        # Calcola il valore di "Dicitura distanziale/imbotte" in base al valore di "Tipo distanziali/imbotti"
        # =SE([@[Tipo distanziali/imbotti]]=3;"DS - SALDATO";SE([@[Tipo distanziali/imbotti]]=2;"I - IMBOTTE";""))
        dicitura_distanziale_imbotte = ""
        if tipo_distanziali_imbotti == "3":
            dicitura_distanziale_imbotte = "DS - SALDATO"
        elif tipo_distanziali_imbotti == "2":
            dicitura_distanziale_imbotte = "I - IMBOTTE"
        
        # Aggiorna il campo "Dicitura distanziale/imbotte"
        self.dicitura_distanziale_imbotte_label.config(text=dicitura_distanziale_imbotte)
        
        if tipo_dist_selezionato:
            # Mostra i campi "L (mm) distanziale" e "H (mm) distanziale"
            self.l_dist_frame.pack(side="left", padx=(0, 2))
            self.h_dist_frame.pack(side="left", padx=(0, 2))
        else:
            # Nascondi i campi
            self.l_dist_frame.pack_forget()
            self.h_dist_frame.pack_forget()
            
            # Resetta i valori
            self.l_dist_entry.delete(0, tk.END)
            self.h_dist_entry.delete(0, tk.END)

    def aggiorna_campi_aggiuntivi(self, event=None):
        """
        Aggiorna i campi aggiuntivi in base al valore selezionato nella combobox "Serramento".
        """
        serramento_selezionato = self.serramento_combobox.get()
        if serramento_selezionato and self.dataframe is not None:
            # Trova la riga corrispondente nel DataFrame
            riga = self.dataframe[self.dataframe.iloc[:, 0] == serramento_selezionato]
            if not riga.empty:
                # Estrai i valori dalle colonne del DataFrame
                numero_ante = riga["NUMERO ANTE"].values[0] if "NUMERO ANTE" in riga else ""
                apertura = riga["APERTURA"].values[0] if "APERTURA" in riga else ""
                tipologia = riga["TIPOLOGIA"].values[0] if "TIPOLOGIA" in riga else ""
                descrizione_tipologia = riga["DESCRIZIONE TIPOLOGIA"].values[0] if "DESCRIZIONE TIPOLOGIA" in riga else ""
                n_cerniere = riga["N. CERNIERE"].values[0] if "N. CERNIERE" in riga else ""
                extra_lavorazione = riga["EXTRA LAVORAZIONE"].values[0] if "EXTRA LAVORAZIONE" in riga else ""
                multiplo = riga["MULTIPLO"].values[0] if "MULTIPLO" in riga else ""
                costo_multiplo = riga["COSTO MULTIPLO"].values[0] if "COSTO MULTIPLO" in riga else ""
                minimi_1 = riga["MINIMI 1 (protezione singola)"].values[0] if "MINIMI 1 (protezione singola)" in riga else ""
                minimi_2 = riga["MINIMI 2 (snodo)"].values[0] if "MINIMI 2 (snodo)" in riga else ""
                minimi_3 = riga["MINIMI 3 (combinati)"].values[0] if "MINIMI 3 (combinati)" in riga else ""
                minimi_4 = riga["MINIMI 4"].values[0] if "MINIMI 4" in riga else ""
                minimi_5 = riga["MINIMI 5"].values[0] if "MINIMI 5" in riga else ""
                presenza_serratura = riga["Presenza serratura"].values[0] if "Presenza serratura" in riga else ""
                numero_staffette = riga["Numero staffette"].values[0] if "Numero staffette" in riga else ""
                
                # Popola i campi aggiuntivi con i valori estratti
                self.numero_ante_label.config(text=str(numero_ante))
                self.apertura_label.config(text=str(apertura))
                self.tipologia_label.config(text=str(tipologia))
                self.descrizione_tipologia_label.config(text=str(descrizione_tipologia))
                self.n_cerniere_label.config(text=str(n_cerniere))
                self.extra_lavorazione_label.config(text=str(extra_lavorazione))
                self.multiplo_label.config(text=str(multiplo))
                self.costo_multiplo_label.config(text=str(costo_multiplo))
                self.minimi_1_label.config(text=str(minimi_1))
                self.minimi_2_label.config(text=str(minimi_2))
                self.minimi_3_label.config(text=str(minimi_3))
                self.minimi_4_label.config(text=str(minimi_4))
                self.minimi_5_label.config(text=str(minimi_5))
                self.presenza_serratura_label.config(text=str(presenza_serratura))
                self.numero_staffette_label.config(text=str(numero_staffette))
            else:
                # Pulisci i campi se non trovi la riga
                self.pulisci_campi_aggiuntivi()
        else:
            # Pulisci i campi se la combobox è vuota
            self.pulisci_campi_aggiuntivi()

    def pulisci_campi_aggiuntivi(self):
        """
        Pulisce i campi aggiuntivi.
        """
        self.numero_ante_label.config(text="")
        self.apertura_label.config(text="")
        self.tipologia_label.config(text="")
        self.descrizione_tipologia_label.config(text="")
        self.n_cerniere_label.config(text="")
        self.extra_lavorazione_label.config(text="")
        self.multiplo_label.config(text="")
        self.costo_multiplo_label.config(text="")
        self.minimi_1_label.config(text="")
        self.minimi_2_label.config(text="")
        self.minimi_3_label.config(text="")
        self.minimi_4_label.config(text="")
        self.minimi_5_label.config(text="")
        self.presenza_serratura_label.config(text="")
        self.numero_staffette_label.config(text="")
        self.sconto_1_label.config(text="")
        self.sconto_2_label.config(text="")
        self.sconto_3_label.config(text="")
        self.sconto_decimali_label.config(text="")
        self.dicitura_sconto_label.config(text="")
        self.prezzo_listino_label.config(text="")
        self.mqr_label.config(text="")
        self.mir_label.config(text="")
        self.tabella_minimi_label.config(text="")
        self.unita_di_misura_label.config(text="")
        self.min_fatt_pz_label.config(text="")
        self.mq_fatt_pz_label.config(text="")
        self.mq_totali_fatt_label.config(text="")
        self.ml_totali_fatt_label.config(text="")
        self.costo_scontato_mq_label.config(text="")
        self.prezzo_listino_unitario_label.config(text="")
        self.costo_serramento_listino_posizione_label.config(text="")
        self.costo_serramento_scontato_posizione_label.config(text="")
        # Pulizia dei nuovi campi per distanziali/imbotti
        self.distanziali_imbotti_label.config(text="")
        self.tipo_distanziali_imbotti_label.config(text="")
        self.dicitura_distanziale_imbotte_label.config(text="")
        self.ml_distanziali_label.config(text="")
        self.ml_imbotti_label.config(text="")
        self.colore_dist_imb_label.config(text="")
        self.tip_dist_imb_label.config(text="")

    def aggiungi_riga(self):
        """
        Recupera i dati inseriti e li aggiunge al Treeview.
        """
        # Verifica che i campi obbligatori siano compilati
        nr_pezzi = self.nr_pezzi_entry.get()
        serramento = self.serramento_combobox.get()
        
        if not nr_pezzi or not serramento:
            messagebox.showerror("Errore", "I campi 'nr. pezzi' e 'Serramento' sono obbligatori.")
            return
        
        # Il numero di posizione è sempre il prossimo valore libero
        pos = self.pos_counter
        
        # Recupera i valori dei campi aggiuntivi
        modello = self.modello_combobox.get()
        modello_grata_combinato = self.modello_grata_combobox.get()
        colore = self.colore_combobox.get()
        l_mm = self.l_mm_entry.get()
        h_mm = self.h_mm_entry.get()
        tipo_telaio = self.tipo_telaio_combobox.get()
        bunk = self.bunk_combobox.get()
        dmcp = self.dmcp_combobox.get()
        defender = self.defender_combobox.get()
        dist = self.dist_combobox.get()
        tipo_dist = self.tipo_dist_combobox.get()
        l_dist = self.l_dist_entry.get()
        h_dist = self.h_dist_entry.get()
        tipo_controtelaio = self.tipo_controtelaio_combobox.get()
        anta_giro = self.anta_giro_combobox.get()
        mrib = self.mrib_combobox.get()
        
        # Recupera i valori dei campi aggiuntivi dal serramento
        numero_ante = self.numero_ante_label.cget("text")
        apertura = self.apertura_label.cget("text")
        tipologia = self.tipologia_label.cget("text")
        descrizione_tipologia = self.descrizione_tipologia_label.cget("text")
        n_cerniere = self.n_cerniere_label.cget("text")
        extra_lavorazione = self.extra_lavorazione_label.cget("text")
        multiplo = self.multiplo_label.cget("text")
        costo_multiplo = self.costo_multiplo_label.cget("text")
        minimi_1 = self.minimi_1_label.cget("text")
        minimi_2 = self.minimi_2_label.cget("text")
        minimi_3 = self.minimi_3_label.cget("text")
        minimi_4 = self.minimi_4_label.cget("text")
        minimi_5 = self.minimi_5_label.cget("text")
        presenza_serratura = self.presenza_serratura_label.cget("text")
        numero_staffette = self.numero_staffette_label.cget("text")
        
        # Calcolo MqR e MIR
        try:
            l_val = float(l_mm.replace(",", ".")) if l_mm else 0.0
            h_val = float(h_mm.replace(",", ".")) if h_mm else 0.0
        except Exception:
            l_val = 0.0
            h_val = 0.0
        mqr = f"{round(l_val * h_val * 0.000001, 2):.2f}".replace('.', ',')
        mir = f"{round((l_val + 2 * h_val) * 0.001, 2):.2f}".replace('.', ',')
        
        # Calcolo Tabella_minimi
        tabella_minimi = ""
        try:
            if modello:
                minimi_row = listino.loc[listino["MODELLO"] == modello]
                if not minimi_row.empty:
                    # Prova a recuperare la colonna corretta tra le possibili varianti
                    possible_minimi_cols = [
                        "Minimi",
                        "MINIMI 1 (protezione singola)",
                        "MINIMI 1",
                        "MINIMI",
                        "Minimi 1",
                        "MINIMI_1",
                    ]
                    found = False
                    for col in possible_minimi_cols:
                        if col in minimi_row.columns:
                            tabella_minimi = str(minimi_row.iloc[0][col])
                            found = True
                            break
                    if not found:
                        # Debug: colonna non trovata
                        print(f"[DEBUG] Nessuna colonna minimi trovata tra {possible_minimi_cols} per modello {modello}")
                        tabella_minimi = ""
                else:
                    print(f"[DEBUG] Nessuna riga in listino per modello: {modello}")
        except Exception as e:
            print(f"[DEBUG] Errore durante il recupero di Tabella_minimi: {e}")
            tabella_minimi = ""
        
        # Calcolo dinamico Unita_di_misura dal listino
        unita_di_misura = ""
        try:
            if modello:
                riga = listino[listino['MODELLO'] == modello]
                if not riga.empty:
                    # Sostituisci la chiave 'Unità misura' o simili con 'Unita_di_misura' senza accento
                    for col in riga.columns:
                        if col.lower().replace('à','a').replace(' ','_') == 'unita_di_misura':
                            unita_di_misura = riga.iloc[0][col]
                            break
        except Exception as e:
            unita_di_misura = f"Errore: {e}"
        
        # Calcolo Prezzo_listino
        prezzo_listino = ""
        try:
            if modello and colore:
                row = listino[listino['MODELLO'] == modello]
                if not row.empty and colore in listino.columns:
                    prezzo = row.iloc[0][colore]
                    if pd.notnull(prezzo):
                        try:
                            prezzo_float = float(prezzo)
                            prezzo_listino = f"€ {prezzo_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                        except Exception:
                            prezzo_listino = f"€ {prezzo},00"
        except Exception:
            pass
        
        # --- POPOLAMENTO AUTOMATICO CAMPI SCONTO E CALCOLATI ---
        # Recupera i valori degli sconti dal preventivo se disponibili
        sconto_1 = sconto_2 = sconto_3 = sconto_decimali = dicitura_sconto = ""
        if self.preventivo and hasattr(self.preventivo, 'dati_b1'):
            dati_b1 = self.preventivo.dati_b1
            def get_sconto_value(dati_b1, *keys):
                for k in keys:
                    if k in dati_b1:
                        return dati_b1[k]
                return ""
            sconto_1 = get_sconto_value(dati_b1, 'Sconto 1', 'Sconto1')
            sconto_2 = get_sconto_value(dati_b1, 'Sconto 2', 'Sconto2')
            sconto_3 = get_sconto_value(dati_b1, 'Sconto 3', 'Sconto3')
            sconto_decimali = get_sconto_value(dati_b1, 'Sconto in decimali', 'Sconto_in_decimali')
            dicitura_sconto = get_sconto_value(dati_b1, 'Dicitura sconto', 'Dicitura_sconto')
        # Calcolo Min_fatt_pz
        min_fatt_pz = ""
        try:
            tabella_val = tabella_minimi.strip()
            tab_min = int(tabella_val) if tabella_val.isdigit() else None
            minimi = [minimi_1, minimi_2, minimi_3, minimi_4, minimi_5]
            if tab_min and 1 <= tab_min <= 5:
                min_fatt_pz = minimi[tab_min-1]
        except Exception:
            min_fatt_pz = ""
        # Calcolo Mq_fatt_pz
        try:
            mqr_float = float(mqr.replace(",", ".")) if mqr else 0.0
            min_fatt_pz_float = float(min_fatt_pz.replace(",", ".")) if min_fatt_pz else 0.0
            mq_fatt_pz = max(mqr_float, min_fatt_pz_float)
            mq_fatt_pz_str = f"{mq_fatt_pz:.2f}".replace('.', ',')
        except Exception:
            mq_fatt_pz_str = ""
        # Calcolo Mq_totali_fatt
        try:
            mq_totali_fatt = mq_fatt_pz * float(nr_pezzi.replace(",", ".")) if mq_fatt_pz_str and nr_pezzi else 0.0
            mq_totali_fatt_str = f"{mq_totali_fatt:.2f}".replace('.', ',')
        except Exception:
            mq_totali_fatt_str = ""
        # Calcolo Ml_totali_fatt
        try:
            mir_float = float(mir.replace(",", ".")) if mir else 0.0
            ml_totali_fatt = mir_float * float(nr_pezzi.replace(",", ".")) if mir and nr_pezzi else 0.0
            ml_totali_fatt_str = f"{ml_totali_fatt:.2f}".replace('.', ',')
        except Exception:
            ml_totali_fatt_str = ""
        # Calcolo dei campi di costo
        try:
            # Estrai il valore numerico dal prezzo listino
            prezzo_listino_val = prezzo_listino
            if isinstance(prezzo_listino_val, str):
                # Prendi solo la parte numerica (gestione "€" e formati)
                import re
                match = re.search(r"([\d.,]+)", prezzo_listino_val)
                if match:
                    prezzo_listino_val = match.group(1).replace(".", "").replace(",", ".")
                else:
                    prezzo_listino_val = "0"
            try:
                prezzo_listino_float = float(prezzo_listino_val)
            except Exception:
                prezzo_listino_float = 0.0
                
            # Gestione sconto in decimali
            sconto_decimali_val = sconto_decimali.strip().replace("%", "") if sconto_decimali else "0"
            sconto_decimali_val = sconto_decimali_val.replace(",", ".")
            try:
                sconto_decimali_float = float(sconto_decimali_val)
                if sconto_decimali_float > 1:
                    sconto_decimali_float = sconto_decimali_float / 100.0
            except Exception:
                sconto_decimali_float = 0.0
                
            # Conversione mq_fatt_pz
            mq_fatt_pz_val = mq_fatt_pz_str.replace(",", ".") if isinstance(mq_fatt_pz_str, str) else "0"
            try:
                mq_fatt_pz_float = float(mq_fatt_pz_val)
            except Exception:
                mq_fatt_pz_float = 0.0
                
            # Conversione mq_totali_fatt
            mq_totali_fatt_val = mq_totali_fatt_str.replace(",", ".") if isinstance(mq_totali_fatt_str, str) else "0"
            try:
                mq_totali_fatt_float = float(mq_totali_fatt_val)
            except Exception:
                mq_totali_fatt_float = 0.0

            # Calcolo dei valori
            costo_scontato_mq_float = prezzo_listino_float * (1 - sconto_decimali_float)
            prezzo_listino_unitario_float = prezzo_listino_float * mq_fatt_pz_float
            costo_serramento_listino_posizione_float = prezzo_listino_float * mq_totali_fatt_float
            costo_serramento_scontato_posizione_float = costo_scontato_mq_float * mq_totali_fatt_float

            # Formattazione in euro
            def euro(val):
                return f"€ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if val else ""
                
            costo_scontato_mq = euro(costo_scontato_mq_float)
            prezzo_listino_unitario = euro(prezzo_listino_unitario_float)
            costo_serramento_listino_posizione = euro(costo_serramento_listino_posizione_float)
            costo_serramento_scontato_posizione = euro(costo_serramento_scontato_posizione_float)
            
        except Exception as e:
            print(f"[DEBUG] Errore nel calcolo dei campi di costo: {e}")
            costo_scontato_mq = ""
            prezzo_listino_unitario = ""
            costo_serramento_listino_posizione = ""
            costo_serramento_scontato_posizione = ""
            
        # Calcolo dei nuovi campi per distanziali/imbotti
        # Campo "Distanziali/Imbotti" = SE(O([Dist.]=="",[Dist.]=="NO"),0,1)
        distanziali_imbotti = "0"
        if dist and dist != "NO":
            distanziali_imbotti = "1"
            
        # Calcolo del campo "Tipo distanziali/imbotti" in base al valore di "Tipo dist."
        # =SE([@[Tipo dist.]]="IMBOTTE";2;SE([@[Tipo dist.]]="SALDATO";3;1))
        tipo_distanziali_imbotti = "1"  # Valore predefinito
        if tipo_dist == "IMBOTTE":
            tipo_distanziali_imbotti = "2"
        elif tipo_dist == "SALDATO":
            tipo_distanziali_imbotti = "3"
            
        # Calcolo del campo "Dicitura distanziale/imbotte" in base al valore di "Tipo distanziali/imbotti"
        # =SE([@[Tipo distanziali/imbotti]]=3;"DS - SALDATO";SE([@[Tipo distanziali/imbotti]]=2;"I - IMBOTTE";""))
        dicitura_distanziale_imbotte = ""
        if tipo_distanziali_imbotti == "3":
            dicitura_distanziale_imbotte = "DS - SALDATO"
        elif tipo_distanziali_imbotti == "2":
            dicitura_distanziale_imbotte = "I - IMBOTTE"
            
        # Altri campi per distanziali/imbotti (inizialmente vuoti)
        ml_distanziali = ""
        ml_imbotti = ""
        # --- LOGICA COLORE DIST/IMB ---
        colore_dist_imb = ""
        if colore == "STANDARD RAL":
            colore_dist_imb = "1"
        elif colore == "EFFETTO LEGNO":
            colore_dist_imb = "2"
        elif colore == "GREZZO":
            colore_dist_imb = "3"
        elif colore == "EXTRA MAZZETTA":
            colore_dist_imb = "4"
        # --- FINE LOGICA COLORE DIST/IMB ---
        tip_dist_imb = ""
        
        # Calcolo del campo "N. distanziali a 3 lati" in base al valore di "Distanziali/Imbotti" e "nr. pezzi"
        # =SE([@[Distanziali/Imbotti]]=1;[@[nr. pezzi]]*3;"")
        n_distanziali_a_3_lati = ""
        if distanziali_imbotti == "1":
            try:
                n_pezzi_val = float(nr_pezzi.replace(",", ".")) if nr_pezzi else 0
                n_distanziali_a_3_lati = str(int(n_pezzi_val * 3))
            except (ValueError, TypeError):
                n_distanziali_a_3_lati = ""
                
        # Campi aggiuntivi per distanziali/imbotti (inizialmente vuoti)
        ml_distanziali_standard_ral = ""
        ml_distanziali_effetto_legno = ""
        ml_distanziali_grezzo = ""
        ml_imbotti_standard_ral = ""
        ml_imbotti_effetto_legno = ""
        ml_imbotti_grezzo = ""
        costo_al_ml = ""
        costo_list_dist = ""
        costo_list_imb = ""
        somma_cost_listino_dist_imbotte = ""
        costo_scontato_somma_dist_imb_posizione = ""
        ml_totali_dist_imb = ""

        # Calcolo del costo al ml controtelaio singolo
        costo_ml_controtelaio_singolo = ""
        try:
            if tipo_controtelaio:
                row = controtelaio[controtelaio["CONTROTELAIO"] == tipo_controtelaio]
                if not row.empty:
                    costo_ml_controtelaio_singolo = row.iloc[0]["COSTO"]
        except Exception as e:
            print(f"[DEBUG] Errore nel calcolo costo_ml_controtelaio_singolo: {e}")

        # --- AGGIORNAMENTO AUTOMATICO NUOVI CAMPI CONTROTELAI ---
        # Verifica controtelaio
        if tipo_controtelaio and tipo_controtelaio != "0":
            verifica_controtelaio = "1"
        else:
            verifica_controtelaio = ""
        # Tipologia ml/nr. Pezzi
        tipologia_ml_nr_pezzi = ""
        try:
            if tipo_controtelaio:
                # Assegna il valore 1 se il controtelaio è singolo o doppio
                if tipo_controtelaio in ["C. SINGOLO", "C. DOPPIO"]:
                    tipologia_ml_nr_pezzi = "1"
                # Assegna il valore 2 se il controtelaio è termico TIP A o TIP B
                elif tipo_controtelaio in ["C. TERMICO TIP A", "C. TERMICO TIP B"]:
                    tipologia_ml_nr_pezzi = "2"
                # Altrimenti lascia vuoto
                else:
                    tipologia_ml_nr_pezzi = ""
            else:
                # Se il campo "Tipologia controtelaio" è vuoto, lascia vuoto anche questo campo
                tipologia_ml_nr_pezzi = ""
        except Exception as e:
            print(f"[DEBUG] Errore nell'aggiornamento di tipologia_ml_nr_pezzi: {e}")

        # Fattore moltiplicatore ml/nr. Pezzi
        fattore_moltiplicatore = ""
        try:
            if verifica_controtelaio == "1" and tipologia_ml_nr_pezzi == "1":
                fattore_moltiplicatore = str(ml_totali_fatt)
            else:
                fattore_moltiplicatore = str(nr_pezzi)
        except Exception as e:
            print(f"[DEBUG] Errore nel calcolo del fattore moltiplicatore: {e}")
            fattore_moltiplicatore = ""

        # Costo listino per posizione controtelaio
        costo_listino_posizione = ""
        try:
            if costo_ml_controtelaio_singolo and fattore_moltiplicatore:
                # Rimuovi il simbolo dell'euro e converte in float
                costo_ml = float(str(costo_ml_controtelaio_singolo).replace("€", "").replace(".", "").replace(",", ".").strip())
                fattore = float(str(fattore_moltiplicatore).replace(",", ".").strip())
                # Calcola il costo totale
                costo_totale = costo_ml * fattore
                # Formatta il risultato in euro
                costo_listino_posizione = f"€ {costo_totale:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except Exception as e:
            print(f"[DEBUG] Errore nel calcolo del costo listino per posizione controtelaio: {e}")
            costo_listino_posizione = ""

        # Costo scontato controtelaio posizione
        costo_scontato_posizione = ""
        try:
            if costo_listino_posizione and sconto_decimali:
                # Rimuovi il simbolo dell'euro e converte in float
                costo_listino = float(str(costo_listino_posizione).replace("€", "").replace(".", "").replace(",", ".").strip())
                # Rimuovi il simbolo % e converte in float
                sconto = float(str(sconto_decimali).replace("%", "").replace(",", ".").strip()) / 100
                # Calcola il costo scontato
                costo_scontato = costo_listino * (1 - sconto)
                # Formatta il risultato in euro
                costo_scontato_posizione = f"€ {costo_scontato:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except Exception as e:
            print(f"[DEBUG] Errore nel calcolo del costo scontato controtelaio posizione: {e}")
            costo_scontato_posizione = ""

        # N. Controtelai singoli
        n_controtelai_singoli = "0"
        try:
            if tipo_controtelaio == "C. SINGOLO":
                n_controtelai_singoli = str(nr_pezzi)
        except Exception as e:
            print(f"[DEBUG] Errore nel calcolo del numero controtelai singoli: {e}")
            n_controtelai_singoli = "0"

        # ML Controtelaio singolo
        ml_controtelaio_singolo = mir if tipo_controtelaio == "C. SINGOLO" else ""
        # Costo Controtelaio singolo
        costo_controtelaio_singolo = costo_ml_controtelaio_singolo if tipo_controtelaio == "C. SINGOLO" else ""
        # N. Controtelai doppi
        n_controtelai_doppi = "1" if tipo_controtelaio == "C. DOPPIO" else "0"
        # ML Controtelaio doppio
        ml_controtelaio_doppio = mir if tipo_controtelaio == "C. DOPPIO" else ""
        # Costo Controtelaio doppio
        costo_controtelaio_doppio = costo_ml_controtelaio_singolo if tipo_controtelaio == "C. DOPPIO" else ""
        # N. Controtelaio termico TIP A
        n_controtelaio_termico_a = "1" if tipo_controtelaio == "C. TERMICO TIP A" else "0"
        # Costo listino controtelaio termico TIP A
        costo_listino_termico_a = costo_ml_controtelaio_singolo if tipo_controtelaio == "C. TERMICO TIP A" else ""
        # N. Controtelaio termico TIP B
        n_controtelaio_termico_b = "1" if tipo_controtelaio == "C. TERMICO TIP B" else "0"
        # Costo listino controtelaio termico TIP B
        costo_listino_termico_b = costo_ml_controtelaio_singolo if tipo_controtelaio == "C. TERMICO TIP B" else ""

        # Inserisco i valori dei nuovi campi nella lista values
        values = [
            pos, nr_pezzi, serramento, modello, modello_grata_combinato, colore, l_mm, h_mm, tipo_telaio, bunk, dmcp, defender, dist, tipo_dist, l_dist, h_dist,
            tipo_controtelaio, anta_giro, mrib, numero_ante, apertura, tipologia, descrizione_tipologia, n_cerniere, extra_lavorazione, multiplo, costo_multiplo, minimi_1, minimi_2, minimi_3, minimi_4, minimi_5, presenza_serratura, numero_staffette,
            sconto_1, sconto_2, sconto_3, sconto_decimali, dicitura_sconto,  # <-- qui i campi sconto
            prezzo_listino,
            mqr, mir, tabella_minimi, unita_di_misura, min_fatt_pz, mq_fatt_pz_str, mq_totali_fatt_str, ml_totali_fatt_str,
            costo_scontato_mq, prezzo_listino_unitario, costo_serramento_listino_posizione, costo_serramento_scontato_posizione,
            distanziali_imbotti, tipo_distanziali_imbotti, dicitura_distanziale_imbotte, ml_distanziali, ml_imbotti, colore_dist_imb, tip_dist_imb,
            # Nuove colonne per distanziali/imbotti richieste
            ml_distanziali_standard_ral, ml_distanziali_effetto_legno, ml_distanziali_grezzo,
            ml_imbotti_standard_ral, ml_imbotti_effetto_legno, ml_imbotti_grezzo,
            costo_al_ml, costo_list_dist, costo_list_imb, somma_cost_listino_dist_imbotte,
            costo_scontato_somma_dist_imb_posizione, n_distanziali_a_3_lati, ml_totali_dist_imb,
            # Nuove colonne per controtelai
            costo_ml_controtelaio_singolo, verifica_controtelaio, tipologia_ml_nr_pezzi, fattore_moltiplicatore, costo_listino_posizione, costo_scontato_posizione,
            n_controtelai_singoli, ml_controtelaio_singolo, costo_controtelaio_singolo, n_controtelai_doppi, ml_controtelaio_doppio, costo_controtelaio_doppio,
            n_controtelaio_termico_a, costo_listino_termico_a, n_controtelaio_termico_b, costo_listino_termico_b
        ]
        
        # --- DEBUG LOG: valori inseriti in aggiungi_riga ---
        print("[DEBUG aggiungi_riga] values:", values)
        print(f"[DEBUG aggiungi_riga] Sconti: S1={sconto_1}, S2={sconto_2}, S3={sconto_3}, Sdec={sconto_decimali}, Dic={dicitura_sconto}")
        
        # Inserisci la riga nel Treeview
        self.tree.insert("", "end", values=values)
        
        # Aggiorna la numerazione delle posizioni (sempre progressiva)
        for i, item in enumerate(self.tree.get_children(), start=1):
            vals = list(self.tree.item(item, "values"))
            vals[0] = i
            self.tree.item(item, values=vals)
        
        # Aggiorna il contatore posizione e il campo Pos. nel frame di inserimento
        self.pos_counter = len(self.tree.get_children()) + 1
        self.pos_label.config(text=str(self.pos_counter))
        
        # Salvataggio dei dati nel file preventivo.py (append)
        try:
            with open("preventivo.py", "a") as f:
                f.write(f"# Riga inserita: pos={pos}, nr_pezzi='{nr_pezzi}', serramento='{serramento}', modello='{modello}', modello_grata_combinato='{modello_grata_combinato}', colore='{colore}', l_mm='{l_mm}', h_mm='{h_mm}', tipo_telaio='{tipo_telaio}', bunk='{bunk}', dmcp='{dmcp}', defender='{defender}', dist='{dist}', tipo_dist='{tipo_dist}', l_dist='{l_dist}', h_dist='{h_dist}', tipo_controtelaio='{tipo_controtelaio}', anta_giro='{anta_giro}', mrib='{mrib}', numero_ante='{numero_ante}', apertura='{apertura}', tipologia='{tipologia}', descrizione_tipologia='{descrizione_tipologia}', n_cerniere='{n_cerniere}', extra_lavorazione='{extra_lavorazione}', multiplo='{multiplo}', costo_multiplo='{costo_multiplo}', minimi_1='{minimi_1}', minimi_2='{minimi_2}', minimi_3='{minimi_3}', minimi_4='{minimi_4}', minimi_5='{minimi_5}', presenza_serratura='{presenza_serratura}', numero_staffette='{numero_staffette}'\n")
                f.write(f"dati_posizioni.append({{ 'pos': {pos}, 'nr_pezzi': '{nr_pezzi}', 'serramento': '{serramento}', 'modello': '{modello}', 'modello_grata_combinato': '{modello_grata_combinato}', 'colore': '{colore}', 'l_mm': '{l_mm}', 'h_mm': '{h_mm}', 'tipo_telaio': '{tipo_telaio}', 'bunk': '{bunk}', 'dmcp': '{dmcp}', 'defender': '{defender}', 'dist': '{dist}', 'tipo_dist': '{tipo_dist}', 'l_dist': '{l_dist}', 'h_dist': '{h_dist}', 'tipo_controtelaio': '{tipo_controtelaio}', 'anta_giro': '{anta_giro}', 'mrib': '{mrib}', 'NUMERO ANTE': '{numero_ante}', 'APERTURA': '{apertura}', 'TIPOLOGIA': '{tipologia}', 'DESCRIZIONE TIPOLOGIA': '{descrizione_tipologia}', 'N. CERNIERE': '{n_cerniere}', 'EXTRA LAVORAZIONE': '{extra_lavorazione}', 'MULTIPLO': '{multiplo}', 'COSTO MULTIPLO': '{costo_multiplo}', 'MINIMI 1': '{minimi_1}', 'MINIMI 2': '{minimi_2}', 'MINIMI 3': '{minimi_3}', 'MINIMI 4': '{minimi_4}', 'MINIMI 5': '{minimi_5}', 'Presenza serratura': '{presenza_serratura}', 'Numero staffette': '{numero_staffette}', 'Sconto 1': '{sconto_1}', 'Sconto 2': '{sconto_2}', 'Sconto 3': '{sconto_3}', 'Sconto in decimali': '{sconto_decimali}', 'Dicitura sconto': '{dicitura_sconto}', 'Prezzo_listino': '{prezzo_listino}', 'MqR': '{mqr}', 'MIR': '{mir}', 'Tabella_minimi': '{tabella_minimi}', 'Unita_di_misura': '{unita_di_misura}', 'Min_fatt_pz': '{min_fatt_pz}', 'Mq_fatt_pz': '{mq_fatt_pz_str}', 'Mq_totali_fatt': '{mq_totali_fatt_str}', 'Ml_totali_fatt': '{ml_totali_fatt_str}', 'Costo_scontato_Mq': '{values[-4]}', 'Prezzo_listino_unitario': '{values[-3]}', 'Costo_serramento_listino_posizione': '{values[-2]}', 'Costo_serramento_scontato_posizione': '{values[-1]}'}})\n")
        except Exception as e:
            print("Errore durante il salvataggio nel file preventivo.py:", e)
        
        # Pulisci i campi di input
        self.nr_pezzi_entry.delete(0, tk.END)
        self.serramento_combobox.set("")
        self.modello_combobox.set("")
        self.modello_grata_combobox.set("DA DEFINIRE")
        self.modello_grata_frame.pack_forget()  # Nascondi il campo
        self.colore_combobox.set(colore_infissi_generale)  # Ripristina il valore predefinito
        self.l_mm_entry.delete(0, tk.END)
        self.h_mm_entry.delete(0, tk.END)
        self.tipo_telaio_combobox.set("")
        self.bunk_combobox.set("NO")  # Ripristina il valore predefinito
        self.dmcp_combobox.set("")
        self.defender_combobox.set("NO")  # Ripristina il valore predefinito
        self.dist_combobox.set("NO")  # Ripristina il valore predefinito
        self.tipo_dist_combobox.set("")
        self.tipo_dist_frame.pack_forget()  # Nascondi il campo
        self.l_dist_entry.delete(0, tk.END)
        self.l_dist_frame.pack_forget()  # Nascondi il campo
        self.h_dist_entry.delete(0, tk.END)
        self.h_dist_frame.pack_forget()  # Nascondi il campo
        self.tipo_controtelaio_combobox.set("")
        self.anta_giro_combobox.set(traverso_inferiore_generale)  # Ripristina il valore predefinito
        self.mrib_combobox.set(maniglia_ribassata_generale)  # Ripristina il valore predefinito
        
        # Pulisci i campi aggiuntivi
        self.pulisci_campi_aggiuntivi()
        
        # Aggiorna l'oggetto preventivo
        self.salva_in_preventivo()
        # Aggiorna tutti i campi controtelaio per tutte le righe
        self.aggiorna_tutti_campi_controtelaio_treeview()

    def elimina_riga_selezionata(self):
        """
        Elimina la riga selezionata dal Treeview e aggiorna le posizioni delle righe successive.
        """
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Attenzione", "Nessuna riga selezionata.")
            return
        
        # Chiedi conferma prima di eliminare
        if messagebox.askyesno("Conferma", "Sei sicuro di voler eliminare la riga selezionata?"):
            # Ottieni tutte le righe del Treeview
            all_items = self.tree.get_children()
            
            # Trova l'indice della riga selezionata
            selected_index = all_items.index(selected_item[0])
            
            # Elimina la riga selezionata
            self.tree.delete(selected_item)
            
            # Rinumera tutte le righe in modo sequenziale
            for i, item in enumerate(self.tree.get_children(), start=1):
                values = list(self.tree.item(item)["values"])
                values[0] = i  # Aggiorna il campo Pos.
                self.tree.item(item, values=values)
            
            # Aggiorna il contatore della posizione e il campo Pos. nel frame di inserimento
            self.aggiorna_contatore_posizione()
            
            # Aggiorna l'oggetto preventivo
            self.salva_in_preventivo()

    def modifica_riga_selezionata(self):
        """
        Apre una finestra di dialogo per modificare la riga selezionata.
        """
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Attenzione", "Nessuna riga selezionata.")
            return
        
        # Prendi il primo elemento selezionato
        item = selected_items[0]
        values = self.tree.item(item, "values")
        
        # Apri la finestra di dialogo per la modifica
        self.show_edit_dialog(item, values)

    def show_edit_dialog(self, item, values):
        """
        Mostra una finestra di dialogo per modificare i valori di una riga.
        
        :param item: ID dell'elemento nel Treeview
        :param values: valori attuali della riga
        """
        dialog = tk.Toplevel(self)
        dialog.title("Modifica Riga")
        dialog.grab_set()  # Rendi la finestra modale
        
        # Imposta la larghezza della finestra a circa la metà dello schermo
        screen_width = dialog.winfo_screenwidth()
        screen_height = dialog.winfo_screenheight()
        dialog_width = int(screen_width * 0.7)  # 70% della larghezza dello schermo
        dialog_height = int(screen_height * 0.8)  # 80% dell'altezza dello schermo
        dialog.geometry(f"{dialog_width}x{dialog_height}")
        
        # Crea un frame principale che conterrà i campi a sinistra e l'immagine a destra
        main_frame = tk.Frame(dialog)
        main_frame.pack(fill="both", expand=True)
        
        # Frame per i campi a sinistra (circa metà della larghezza)
        fields_frame = tk.Frame(main_frame, width=int(dialog_width * 0.5))
        fields_frame.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        fields_frame.pack_propagate(False)  # Impedisce al frame di ridimensionarsi
        
        # Frame per l'immagine a destra
        image_frame = tk.LabelFrame(main_frame, text="Immagine Serramento", width=int(dialog_width * 0.5))
        image_frame.pack(side="right", fill="both", expand=True, padx=5, pady=5)
        image_frame.pack_propagate(False)  # Impedisce al frame di ridimensionarsi
        
        # Label per l'immagine
        self.image_label = tk.Label(image_frame)
        self.image_label.pack(fill="both", expand=True, padx=5, pady=5)
        
        entries = {}
        row = 0
        
        # Crea i campi di input per i valori modificabili
        for i, col in enumerate(self.colonne[:19]):  # Solo i primi 19 campi sono modificabili
            if col in ["Pos.", "Mq_fatt_pz", "Mq_totali_fatt", "Ml_totali_fatt"]:
                continue
                
            frame = tk.Frame(fields_frame)
            frame.pack(fill="x", padx=5, pady=2)
            
            # Etichetta del campo (larghezza fissa)
            label = tk.Label(frame, text=f"{col}:", width=15, anchor="w")
            label.pack(side="left")
            
            if col in ["Serramento", "Modello", "Modello grata combinato", "Colore", "Tipo telaio",
                      "BUNK", "Dmcp / Scp", "Defender", "Dist.", "Tipo dist.",
                      "Tipologia controtelaio", "Anta a giro posizione", "M.rib."]:
                # Usa Combobox per i campi con valori predefiniti
                entry = ttk.Combobox(frame, width=20)  # Ridotta la larghezza
                entry.pack(side="left", fill="x", expand=True)
                
                # Popola i valori del combobox in base al campo
                if col == "Serramento" and self.dataframe is not None:
                    entry['values'] = sorted(list(self.dataframe["TIPOLOGIA COMPLETA DI APERTURA"].unique()))
                elif col == "Modello" and 'listino' in globals():
                    entry['values'] = sorted(list(listino["MODELLO"].unique()))
                elif col == "Modello grata combinato" and 'modello_grata_combinato' in globals():
                    entry['values'] = sorted(list(modello_grata_combinato.iloc[:, 0].unique()))
                elif col == "Colore":
                    entry['values'] = ["STANDARD RAL", "EFFETTO LEGNO", "GREZZO"]
                    if i < len(values) and values[i]:
                        entry.set(values[i])
                elif col == "Tipo telaio":
                    entry['values'] = sorted(self.telaio_values)
                elif col == "BUNK":
                    entry['values'] = ["SI", "NO"]
                elif col == "Dmcp / Scp":
                    entry['values'] = sorted([
                        "DOPPIA MANIGLIA E CILINDRO PASSANTE", 
                        "SOLO CILINDRO PASSANTE", 
                        "SOLO MEZZO-CILINDRO INTERNO", 
                        "SOLO MEZZO-CILINDRO ESTERNO", 
                        "MANIGLIA E MEZZO-CILINDRO ESTERNO", 
                        "SENZA MEZZO CILINDRO E MANIGLIA"
                    ])
                elif col == "Defender":
                    entry['values'] = ["SI", "NO"]
                elif col == "Dist.":
                    entry['values'] = ["SI", "NO"]
                elif col == "Tipo dist.":
                    entry['values'] = ["SALDATO", "IMBOTTE"]
                elif col == "Tipologia controtelaio":
                    entry['values'] = sorted(["C. SINGOLO", "C. DOPPIO", "C. TERMICO TIP A", "C. TERMICO TIP B"])
                elif col == "Anta a giro posizione":
                    entry['values'] = sorted(["ANTA A GIRO", "LIBERO A MARMO", "ANTA A GIRO CON BATTUTA A MARMO"])
                elif col == "M.rib.":
                    entry['values'] = ["SI", "NO"]
                
                if i < len(values) and values[i]:
                    entry.set(values[i])
                # --- AGGIUNTA: aggiorna campi aggiuntivi live quando cambia Serramento ---
                if col == "Serramento":
                    def on_serramento_selected(event=None, entries=entries, item=item):
                        serramento = entries["Serramento"].get()
                        print(f"[DEBUG] Cambio serramento in dialog: {serramento!r}")
                        # Trova la riga corrispondente nel DataFrame
                        row = self.dataframe[self.dataframe["TIPOLOGIA COMPLETA DI APERTURA"] == serramento]
                        print(f"[DEBUG] Colonne disponibili nella riga: {row.columns}")
                        minimi_map = {
                            "MIN.1": ["MINIMI 1 (protezione singola)", "MINIMI 1", "Minimi 1", "MINIMI_1"],
                            "MIN.2": ["MINIMI 2 (snodo)", "MINIMI 2", "Minimi 2", "MINIMI_2"],
                            "MIN.3": ["MINIMI 3 (combinati)", "MINIMI 3", "Minimi 3", "MINIMI_3"],
                            "MIN.4": ["MINIMI 4", "Minimi 4", "MINIMI_4"],
                            "MIN.5": ["MINIMI 5", "Minimi 5", "MINIMI_5"],
                        }
                        if not row.empty:
                            for min_col, possibili in minimi_map.items():
                                valore = ""
                                for col_df in possibili:
                                    if col_df in row.columns:
                                        valore = row.iloc[0][col_df]
                                        print(f"[DEBUG] Mapping {col_df} -> {min_col}: {valore}")
                                        break
                                if min_col in entries:
                                    entries[min_col].config(text=str(valore))
                        self.load_serramento_image(serramento)
                    entry.bind("<<ComboboxSelected>>", on_serramento_selected)
                    # AGGIUNGI: aggiorna subito i campi anche al caricamento della dialog
                    entry.event_generate("<<ComboboxSelected>>")
                elif col == "Colore":
                    def on_colore_selected(event=None, entries=entries):
                        colore = entries["Colore"].get()
                        colore_dist_imb = ""
                        if colore == "STANDARD RAL":
                            colore_dist_imb = "1"
                        elif colore == "EFFETTO LEGNO":
                            colore_dist_imb = "2"
                        elif colore == "GREZZO":
                            colore_dist_imb = "3"
                        elif colore == "EXTRA MAZZETTA":
                            colore_dist_imb = "4"
                        if "Colore dist/imb" in entries:
                            entries["Colore dist/imb"].config(text=colore_dist_imb)
                    entry.bind("<<ComboboxSelected>>", on_colore_selected)
                    # Aggiorna subito il campo anche al caricamento della dialog
                    entry.event_generate("<<ComboboxSelected>>")
            else:
                # Usa Entry per i campi di testo
                entry = tk.Entry(frame, width=20)  # Ridotta la larghezza
                entry.pack(side="left", fill="x", expand=True)
                if i < len(values) and values[i]:
                    entry.insert(0, values[i])
            
            entries[col] = entry
        
        # Aggiungi i campi non modificabili (ma non li visualizziamo)
        non_editable_fields = [
            ("N.ANTE", ""),
            ("AP.", ""),
            ("TIP.", ""),
            ("DESCR.TIP", ""),
            ("N.CERN.", ""),
            ("XLAV", ""),
            ("MULT.", ""),
            ("COSTO MLT", ""),
            ("MIN.1", ""),
            ("MIN.2", ""),
            ("MIN.3", ""),
            ("MIN.4", ""),
            ("MIN.5", ""),
            ("P.SERR.", ""),
            ("N.STAFF.", ""),
            ("Sconto 1", ""),
            ("Sconto 2", ""),
            ("Sconto 3", ""),
            ("Sconto in decimali", ""),
            ("Dicitura sconto", ""),
            ("Prezzo_listino", ""),
            ("MqR", ""),
            ("MIR", ""),
            ("Tabella_minimi", ""),
            ("Unita_di_misura", ""),
            ("Min_fatt_pz", ""),
            ("Mq_fatt_pz", ""),
            ("Mq_totali_fatt", ""),
            ("Ml_totali_fatt", ""),
            # Nuove colonne calcolate automaticamente
            ("Costo_scontato_Mq", ""),
            ("Prezzo_listino_unitario", ""),
            ("Costo_serramento_listino_posizione", ""),
            ("Costo_serramento_scontato_posizione", ""),
            # Nuovi campi per distanziali/imbotti
            ("Distanziali/Imbotti", ""),
            ("Tipo distanziali/imbotti", ""),
            ("Dicitura distanziale/imbotte", ""),
            ("Ml Distanziali", ""),
            ("Ml Imbotti", ""),
            ("Colore dist/imb", ""),
            ("Tip. dist/imb", "")
        ]
        
        # Creiamo i campi non modificabili ma li nascondiamo (sono necessari per il funzionamento)
        hidden_frame = tk.Frame(dialog)
        hidden_frame.pack_forget()  # Nascondi il frame
        
        # Crea i campi nascosti per i valori calcolati
        for field_name, default_value in non_editable_fields:
            label = tk.Label(hidden_frame, text=default_value)
            entries[field_name] = label
            
            # Se il campo esiste nei valori correnti, imposta il valore
            if field_name in self.colonne:
                idx = self.colonne.index(field_name)
                if idx < len(values) and values[idx]:
                    label.config(text=values[idx])
        
        # Pulsanti Salva modifica e Annulla
        button_frame = tk.Frame(fields_frame)
        button_frame.pack(fill="x", pady=10)
        
        tk.Button(button_frame, text="Salva modifica", command=lambda: self.save_edited_row(dialog, item, entries), 
                 bg="green", fg="white").pack(side="left", padx=5)
        tk.Button(button_frame, text="Annulla", command=dialog.destroy).pack(side="left")
        
        # Binding per gli eventi
        if "Modello" in entries:
            entries["Modello"].bind("<<ComboboxSelected>>", lambda e: self.on_edit_modello_selected(entries))
        if "Dmcp / Scp" in entries:
            entries["Dmcp / Scp"].bind("<<ComboboxSelected>>", lambda e: self.on_edit_dmcp_selected(entries))
        if "Dist." in entries:
            entries["Dist."].bind("<<ComboboxSelected>>", lambda e: self.on_edit_dist_selected(entries))
        if "Tipo dist." in entries:
            entries["Tipo dist."].bind("<<ComboboxSelected>>", lambda e: self.on_edit_tipo_dist_selected(entries))
    
    def load_serramento_image(self, serramento_name):
        """
        Carica l'immagine del serramento dalla cartella risorse/serramenti.
        
        :param serramento_name: Nome del serramento
        """
        try:
            # Normalizza il nome del serramento per il nome del file
            # Rimuovi spazi e caratteri speciali, converti in minuscolo
            import re
            normalized_name = re.sub(r'[^a-zA-Z0-9]', '_', serramento_name.lower())
            
            # Percorsi possibili per l'immagine
            possible_extensions = ['.png', '.jpg', '.jpeg', '.gif']
            image_path = None
            
            for ext in possible_extensions:
                path = os.path.join("risorse", "serramenti", f"{normalized_name}{ext}")
                if os.path.exists(path):
                    image_path = path
                    break
            
            # Se non troviamo un'immagine con il nome normalizzato, cerchiamo qualsiasi immagine che contenga parte del nome
            if not image_path:
                for filename in os.listdir(os.path.join("risorse", "serramenti")):
                    if any(filename.lower().endswith(ext) for ext in possible_extensions):
                        # Verifica se il nome del file contiene parte del nome del serramento
                        name_parts = normalized_name.split('_')
                        for part in name_parts:
                            if len(part) > 3 and part in filename.lower():  # Solo parti significative (più di 3 caratteri)
                                image_path = os.path.join("risorse", "serramenti", filename)
                                break
                        if image_path:
                            break
            
            # Se abbiamo trovato un'immagine, la carichiamo
            if image_path and os.path.exists(image_path):
                # Carica l'immagine con PIL
                img = Image.open(image_path)
                
                # Ridimensiona l'immagine mantenendo le proporzioni
                max_width = 400
                max_height = 400
                width, height = img.size
                
                # Calcola le nuove dimensioni mantenendo le proporzioni
                if width > height:
                    new_width = max_width
                    new_height = int(height * (max_width / width))
                else:
                    new_height = max_height
                    new_width = int(width * (max_height / height))
                
                img = img.resize((new_width, new_height), Image.LANCZOS)
                
                # Converti in formato Tkinter
                photo = ImageTk.PhotoImage(img)
                
                # Aggiorna l'etichetta con l'immagine
                self.image_label.config(image=photo)
                self.image_label.image = photo  # Mantieni un riferimento per evitare il garbage collection
                
                # Aggiorna il testo dell'etichetta
                self.image_label.config(text="")
            else:
                # Se non troviamo l'immagine, mostriamo un messaggio
                self.image_label.config(image="")
                self.image_label.config(text=f"Immagine non trovata per:\n{serramento_name}")
        
        except Exception as e:
            print(f"Errore nel caricamento dell'immagine: {e}")
            self.image_label.config(image="")
            self.image_label.config(text=f"Errore nel caricamento dell'immagine:\n{str(e)}")
        
    def on_edit_modello_selected(self, entries):
        """Gestisce la selezione del modello nella finestra di modifica."""
        modello_selezionato = entries["Modello"].get()
        modelli_duo = [
            "DUO", "DUO MILLENNIUM", "DUO ALLUMINIO", "DUO REVOLUTION", 
            "DUO GRETHA", "DUO GRETHA ALLUMINIO", "DUO BLIND", 
            "DUO BLIND ORIENTABILE", "DUO BLIND ALLUMINIO", "DUO BLIND SLIM"
        ]
        
        # Trova il frame contenitore del campo "Modello grata combinato"
        for widget in entries["Modello grata combinato"].master.master.winfo_children():
            if isinstance(widget, ttk.Frame) and "Modello grata combinato" in [child.cget("text").replace(":", "") for child in widget.winfo_children() if hasattr(child, "cget")]:
                if modello_selezionato in modelli_duo:
                    widget.pack(fill="x", pady=5)
                else:
                    widget.pack_forget()
                break

    def on_edit_dmcp_selected(self, entries):
        """Gestisce la selezione del campo 'Dmcp / Scp' nella finestra di modifica."""
        dmcp_selezionato = entries["Dmcp / Scp"].get()
        
        if dmcp_selezionato in ["DOPPIA MANIGLIA E CILINDRO PASSANTE", "SOLO CILINDRO PASSANTE"]:
            entries["Defender"].set("SI")
        else:
            entries["Defender"].set("NO")

    def on_edit_dist_selected(self, entries):
        """Gestisce la selezione del campo 'Dist.' nella finestra di modifica."""
        dist_selezionato = entries["Dist."].get()
        
        # Aggiorna il campo "Distanziali/Imbotti" in base al valore di "Dist."
        # Se "Dist." è vuoto o "NO", "Distanziali/Imbotti" deve essere "0", altrimenti "1"
        distanziali_imbotti = "0"
        if dist_selezionato and dist_selezionato != "NO":
            distanziali_imbotti = "1"
        if "Distanziali/Imbotti" in entries:
            entries["Distanziali/Imbotti"].config(text=distanziali_imbotti)
        
        # Trova il frame contenitore del campo "Tipo dist."
        for widget in entries["Tipo dist."].master.master.winfo_children():
            if isinstance(widget, ttk.Frame) and "Tipo dist." in [child.cget("text").replace(":", "") for child in widget.winfo_children() if hasattr(child, "cget")]:
                if dist_selezionato == "SI":
                    widget.pack(fill="x", pady=5)
                else:
                    widget.pack_forget()
                    # Reset the value of Tipo dist. when Dist. is NO
                    entries["Tipo dist."].set("")
                    # Trigger the Tipo dist. event handler to update related fields
                    self.on_edit_tipo_dist_selected(entries)
                    
                    # Nascondi anche i campi correlati
                    for w in entries["L (mm) dist."].master.master.winfo_children():
                        if isinstance(w, ttk.Frame) and "L (mm) dist." in [c.cget("text").replace(":", "") for c in w.winfo_children() if hasattr(c, "cget")]:
                            w.pack_forget()
                    
                    for w in entries["H (mm) dist."].master.master.winfo_children():
                        if isinstance(w, ttk.Frame) and "H (mm) dist." in [c.cget("text").replace(":", "") for c in w.winfo_children() if hasattr(c, "cget")]:
                            w.pack_forget()
                break

    def on_edit_tipo_dist_selected(self, entries):
        """Gestisce la selezione del campo 'Tipo dist.' nella finestra di modifica."""
        tipo_dist_selezionato = entries["Tipo dist."].get()
        
        # Calcola il valore di "Tipo distanziali/imbotti" in base al valore di "Tipo dist."
        # =SE([@[Tipo dist.]]="IMBOTTE";2;SE([@[Tipo dist.]]="SALDATO";3;1))
        tipo_distanziali_imbotti = "1"  # Valore predefinito
        if tipo_dist_selezionato == "IMBOTTE":
            tipo_distanziali_imbotti = "2"
        elif tipo_dist_selezionato == "SALDATO":
            tipo_distanziali_imbotti = "3"
        
        # Aggiorna il campo "Tipo distanziali/imbotti"
        if "Tipo distanziali/imbotti" in entries:
            entries["Tipo distanziali/imbotti"].config(text=tipo_distanziali_imbotti)
        
        # Calcola il valore di "Dicitura distanziale/imbotte" in base al valore di "Tipo distanziali/imbotti"
        # =SE([@[Tipo distanziali/imbotti]]=3;"DS - SALDATO";SE([@[Tipo distanziali/imbotti]]=2;"I - IMBOTTE";""))
        dicitura_distanziale_imbotte = ""
        if tipo_distanziali_imbotti == "3":
            dicitura_distanziale_imbotte = "DS - SALDATO"
        elif tipo_distanziali_imbotti == "2":
            dicitura_distanziale_imbotte = "I - IMBOTTE"
        
        # Aggiorna il campo "Dicitura distanziale/imbotte"
        if "Dicitura distanziale/imbotte" in entries:
            entries["Dicitura distanziale/imbotte"].config(text=dicitura_distanziale_imbotte)
        
        # Trova i frame contenitori dei campi "L (mm) distanziale" e "H (mm) distanziale"
        for widget in entries["L (mm) dist."].master.master.winfo_children():
            if isinstance(widget, ttk.Frame) and "L (mm) dist." in [child.cget("text").replace(":", "") for child in widget.winfo_children() if hasattr(child, "cget")]:
                if tipo_dist_selezionato:
                    widget.pack(fill="x", pady=5)
                else:
                    widget.pack_forget()
                break
        
        for widget in entries["H (mm) dist."].master.master.winfo_children():
            if isinstance(widget, ttk.Frame) and "H (mm) dist." in [child.cget("text").replace(":", "") for child in widget.winfo_children() if hasattr(child, "cget")]:
                if tipo_dist_selezionato:
                    widget.pack(fill="x", pady=5)
                else:
                    widget.pack_forget()
                break

    def update_edit_dialog_fields(self, item, entries):
        """
        Aggiorna i campi aggiuntivi nella finestra di modifica in base al serramento selezionato.
        """
        try:
            values = self.tree.item(item, "values")
            serramento = values[2] if len(values) > 2 else ""
            
            # Verifica se il DataFrame è stato caricato correttamente
            if self.dataframe is None or self.dataframe.empty:
                print("DataFrame non caricato o vuoto")
                return
                
            # Trova il nome corretto della colonna del serramento
            serramento_column = None
            for col in self.dataframe.columns:
                if "SERRAMENTO" in str(col).upper():
                    serramento_column = col
                    break
            
            if serramento_column is None:
                print("Colonna SERRAMENTO non trovata nel DataFrame")
                return
            
            # Trova la riga corrispondente nel DataFrame
            row = self.dataframe[self.dataframe[serramento_column] == serramento]
            if row.empty:
                print(f"Nessuna riga trovata per il serramento: {serramento}")
                return
            
            # Mappa dei nomi delle colonne del DataFrame ai nomi abbreviati
            column_mapping = {
                "NUMERO ANTE": "N.ANTE",
                "APERTURA": "AP.",
                "TIPOLOGIA": "TIP.",
                "DESCRIZIONE TIPOLOGIA": "DESCR.TIP",
                "N. CERNIERE": "N.CERN.",
                "EXTRA LAVORAZIONE": "XLAV",
                "MULTIPLO": "MULT.",
                "COSTO MULTIPLO": "COSTO MLT",
                "MINIMI 1": "MIN.1",
                "MINIMI 2": "MIN.2",
                "MINIMI 3": "MIN.3",
                "MINIMI 4": "MIN.4",
                "MINIMI 5": "MIN.5",
                "Presenza serratura": "P.SERR.",
                "Numero staffette": "N.STAFF.",
                "Unità_di_misura": "Unita_di_misura"
            }
            
            # Aggiorna i campi con i nuovi valori
            for df_col, abbrev in column_mapping.items():
                try:
                    if df_col in row.columns:
                        value = row[df_col].iloc[0]
                        if abbrev in entries:
                            entries[abbrev].config(text=str(value))
                except Exception as e:
                    print(f"Errore nell'aggiornamento del campo {df_col}: {e}")
                    
        except Exception as e:
            print(f"Errore in update_edit_dialog_fields: {e}")

    def save_edited_row(self, dialog, item, entries):
        """
        Salva i valori modificati nella riga selezionata.
        
        :param dialog: finestra di dialogo
        :param item: ID dell'elemento nel Treeview
        :param entries: dizionario dei widget di input
        """
        try:
            # Ottieni i valori correnti della riga
            current_values = list(self.tree.item(item, "values"))
            
            # Crea una nuova lista per i valori aggiornati
            new_values = []
            
            # Itera attraverso tutte le colonne
            for i, col in enumerate(self.colonne):
                if col == "Pos.":
                    # Mantieni il valore della posizione invariato
                    new_values.append(current_values[i] if i < len(current_values) else "")
                elif col in entries:
                    # Ottieni il valore dal widget appropriato
                    widget = entries[col]
                    if isinstance(widget, (ttk.Combobox, tk.Entry)):
                        value = widget.get()
                    elif isinstance(widget, tk.Label):
                        value = widget.cget("text")
                    else:
                        value = ""
                    new_values.append(value)
                else:
                    # Se la colonna non ha un widget corrispondente, mantieni il valore corrente
                    new_values.append(current_values[i] if i < len(current_values) else "")
            
            # --- FIX: Ripristina MIN.1, MIN.2, MIN.3 se vuoti ---
            for min_col in ["MIN.1", "MIN.2", "MIN.3"]:
                idx = self.colonne.index(min_col)
                if idx < len(new_values) and (not new_values[idx] or new_values[idx] == ""):
                    new_values[idx] = current_values[idx] if idx < len(current_values) else ""
            
            # --- AGGIORNAMENTO AUTOMATICO MqR e MIR SE MODIFICATI L (mm) o H (mm) ---
            try:
                l_index = self.colonne.index("L (mm)")
                h_index = self.colonne.index("H (mm)")
                mqr_index = self.colonne.index("MqR")
                mir_index = self.colonne.index("MIR")
                tabella_minimi_index = self.colonne.index("Tabella_minimi")
                l_mm_str = new_values[l_index] if l_index < len(new_values) else ""
                h_mm_str = new_values[h_index] if h_index < len(new_values) else ""
                modello_index = self.colonne.index("Modello")
                modello_val = new_values[modello_index] if modello_index < len(new_values) else ""
                l_val = float(str(l_mm_str).replace(",", ".")) if l_mm_str else 0.0
                h_val = float(str(h_mm_str).replace(",", ".")) if h_mm_str else 0.0
                mqr = f"{round(l_val * h_val * 0.000001, 2):.2f}".replace('.', ',')
                mir = f"{round((l_val + 2 * h_val) * 0.001, 2):.2f}".replace('.', ',')
                # Aggiorna i valori calcolati
                if mqr_index < len(new_values):
                    new_values[mqr_index] = mqr
                if mir_index < len(new_values):
                    new_values[mir_index] = mir
                # Aggiorna Tabella_minimi
                tabella_minimi = ""
                try:
                    if modello_val:
                        minimi_row = listino.loc[listino["MODELLO"] == modello_val]
                        if not minimi_row.empty:
                            # Prova a recuperare la colonna corretta tra le possibili varianti
                            possible_minimi_cols = [
                                "Minimi",
                                "MINIMI 1 (protezione singola)",
                                "MINIMI 1",
                                "MINIMI",
                                "Minimi 1",
                                "MINIMI_1",
                            ]
                            found = False
                            for col in possible_minimi_cols:
                                if col in minimi_row.columns:
                                    tabella_minimi = str(minimi_row.iloc[0][col])
                                    found = True
                                    break
                            if not found:
                                # Debug: colonna non trovata
                                print(f"[DEBUG] Nessuna colonna minimi trovata tra {possible_minimi_cols} per modello {modello_val}")
                                tabella_minimi = ""
                        else:
                            print(f"[DEBUG] Nessuna riga in listino per modello: {modello_val}")
                except Exception as e:
                    print(f"[DEBUG] Errore durante il recupero di Tabella_minimi: {e}")
                    tabella_minimi = ""
                if tabella_minimi_index < len(new_values):
                    new_values[tabella_minimi_index] = tabella_minimi
            except Exception as e:
                print(f"[DEBUG] Errore nell'aggiornamento di MqR e MIR: {e}")

            # --- AGGIORNAMENTO AUTOMATICO COSTO AL ML CONTROTELAIO SINGOLO ---
            try:
                tipo_controtelaio_index = self.colonne.index("Tipologia controtelaio")
                costo_ml_controtelaio_singolo_index = self.colonne.index("Costo al ml controtelaio singolo")
                tipo_controtelaio_val = new_values[tipo_controtelaio_index] if tipo_controtelaio_index < len(new_values) else ""
                costo_ml_controtelaio_singolo = ""
                if tipo_controtelaio_val:
                    row = controtelaio[controtelaio["CONTROTELAIO"] == tipo_controtelaio_val]
                    if not row.empty:
                        costo_ml_controtelaio_singolo = row.iloc[0]["COSTO"]
                if costo_ml_controtelaio_singolo_index < len(new_values):
                    new_values[costo_ml_controtelaio_singolo_index] = costo_ml_controtelaio_singolo
            except Exception as e:
                print(f"[DEBUG] Errore nell'aggiornamento di costo_ml_controtelaio_singolo: {e}")

            # --- AGGIORNAMENTO AUTOMATICO VERIFICA CONTROTELAIO ---
            try:
                verifica_controtelaio_index = self.colonne.index("Verifica controtelaio")
                # Nuova logica: 1 solo se Tipologia controtelaio non è vuota e diversa da 0
                if tipo_controtelaio_val and tipo_controtelaio_val != "0":
                    verifica_controtelaio = "1"
                else:
                    verifica_controtelaio = ""
                if verifica_controtelaio_index < len(new_values):
                    new_values[verifica_controtelaio_index] = verifica_controtelaio
            except Exception as e:
                print(f"[DEBUG] Errore nell'aggiornamento di verifica_controtelaio: {e}")

            # --- AGGIORNAMENTO AUTOMATICO TIPOLOGIA ML/NR. PEZZI ---
            try:
                tipologia_ml_nr_pezzi_index = self.colonne.index("Tipologia ml/nr. Pezzi")
                tipologia_ml_nr_pezzi = ""
                if tipo_controtelaio_val:
                    row = controtelaio[controtelaio["CONTROTELAIO"] == tipo_controtelaio_val]
                    if not row.empty:
                        tipologia_ml_nr_pezzi = str(row.iloc[0]["Ml / nr. Pezzi"])
                if tipologia_ml_nr_pezzi_index < len(new_values):
                    new_values[tipologia_ml_nr_pezzi_index] = tipologia_ml_nr_pezzi
            except Exception as e:
                print(f"[DEBUG] Errore nell'aggiornamento di tipologia_ml_nr_pezzi: {e}")

            # --- AGGIORNAMENTO AUTOMATICO FATTORE MOLTIPLICATORE ML/NR. PEZZI ---
            try:
                fattore_moltiplicatore_index = self.colonne.index("Fattore moltiplicatore ml/nr. Pezzi")
                verifica_controtelaio_index = self.colonne.index("Verifica controtelaio")
                tipologia_ml_nr_pezzi_index = self.colonne.index("Tipologia ml/nr. Pezzi")
                ml_totali_fatt_index = self.colonne.index("Ml_totali_fatt")
                nr_pezzi_index = self.colonne.index("nr. pezzi")

                verifica_controtelaio_val = new_values[verifica_controtelaio_index] if verifica_controtelaio_index < len(new_values) else ""
                tipologia_ml_nr_pezzi_val = new_values[tipologia_ml_nr_pezzi_index] if tipologia_ml_nr_pezzi_index < len(new_values) else ""
                ml_totali_fatt_val = new_values[ml_totali_fatt_index] if ml_totali_fatt_index < len(new_values) else ""
                nr_pezzi_val = new_values[nr_pezzi_index] if nr_pezzi_index < len(new_values) else ""

                # Converti i valori in stringhe e rimuovi eventuali spazi
                verifica_controtelaio_val = str(verifica_controtelaio_val).strip()
                tipologia_ml_nr_pezzi_val = str(tipologia_ml_nr_pezzi_val).strip()
                ml_totali_fatt_val = str(ml_totali_fatt_val).strip()
                nr_pezzi_val = str(nr_pezzi_val).strip()

                # Debug per verificare i valori prima del calcolo
                print(f"[DEBUG] Verifica controtelaio (raw): {verifica_controtelaio_val}")
                print(f"[DEBUG] Tipologia ml/nr. Pezzi (raw): {tipologia_ml_nr_pezzi_val}")
                print(f"[DEBUG] Ml_totali_fatt (raw): {ml_totali_fatt_val}")
                print(f"[DEBUG] nr. pezzi (raw): {nr_pezzi_val}")

                # Implementazione della formula richiesta
                if verifica_controtelaio_val == "1" and tipologia_ml_nr_pezzi_val == "1":
                    # Se entrambe le condizioni sono vere, usa Ml_totali_fatt
                    if ml_totali_fatt_val and ml_totali_fatt_val != "0":
                        try:
                            # Prova a convertire in float per verificare se è un numero valido
                            float(ml_totali_fatt_val)
                            fattore_moltiplicatore = ml_totali_fatt_val
                            print(f"[DEBUG] Condizioni soddisfatte: usando Ml_totali_fatt = {ml_totali_fatt_val}")
                        except ValueError:
                            fattore_moltiplicatore = nr_pezzi_val
                            print(f"[DEBUG] Ml_totali_fatt non è un numero valido: usando nr. pezzi = {nr_pezzi_val}")
                    else:
                        fattore_moltiplicatore = nr_pezzi_val
                        print(f"[DEBUG] Ml_totali_fatt vuoto o zero: usando nr. pezzi = {nr_pezzi_val}")
                else:
                    fattore_moltiplicatore = nr_pezzi_val
                    print(f"[DEBUG] Condizioni non soddisfatte: usando nr. pezzi = {nr_pezzi_val}")

                # Aggiorna il valore nel dizionario
                if fattore_moltiplicatore_index < len(new_values):
                    new_values[fattore_moltiplicatore_index] = fattore_moltiplicatore

                # Debug per verificare il risultato finale
                print(f"[DEBUG] Fattore moltiplicatore risultante: {fattore_moltiplicatore}")

            except Exception as e:
                print(f"[DEBUG] Errore nell'aggiornamento di fattore_moltiplicatore: {e}")
                import traceback
                print(f"[DEBUG] Stack trace: {traceback.format_exc()}")

            # --- AGGIORNAMENTO AUTOMATICO COSTO LISTINO PER POSIZIONE CONTROTELAIO ---
            try:
                costo_listino_posizione_index = self.colonne.index("Costo listino per posizione controtelaio")
                costo_listino_posizione = costo_ml_controtelaio_singolo
                if costo_listino_posizione_index < len(new_values):
                    new_values[costo_listino_posizione_index] = costo_listino_posizione
            except Exception as e:
                print(f"[DEBUG] Errore nell'aggiornamento di costo_listino_posizione: {e}")

            # --- AGGIORNAMENTO AUTOMATICO COSTO SCONTATO CONTROTELAIO POSIZIONE ---
            try:
                costo_scontato_posizione_index = self.colonne.index("Costo scontato controtelaio posizione")
                costo_scontato_posizione = costo_ml_controtelaio_singolo
                if costo_scontato_posizione_index < len(new_values):
                    new_values[costo_scontato_posizione_index] = costo_scontato_posizione
            except Exception as e:
                print(f"[DEBUG] Errore nell'aggiornamento di costo_scontato_posizione: {e}")

            # --- AGGIORNAMENTO AUTOMATICO N. CONTROTELAI SINGOLI ---
            try:
                n_controtelai_singoli_index = self.colonne.index("N. Controtelai singoli")
                n_controtelai_singoli = "0"
                try:
                    nr_pezzi_val = values[self.colonne.index("nr. pezzi")] if "nr. pezzi" in self.colonne else ""
                    if tipo_controtelaio == "C. SINGOLO":
                        n_controtelai_singoli = str(nr_pezzi_val)
                except Exception as e:
                    print(f"[DEBUG] Errore nel calcolo del numero controtelai singoli: {e}")
                    n_controtelai_singoli = "0"
                if n_controtelai_singoli_index < len(new_values):
                    new_values[n_controtelai_singoli_index] = n_controtelai_singoli
            except Exception as e:
                print(f"[DEBUG] Errore nell'aggiornamento di n_controtelai_singoli: {e}")

            # --- AGGIORNAMENTO AUTOMATICO ML CONTROTELAIO SINGOLO ---
            try:
                ml_controtelaio_singolo_index = self.colonne.index("ML Controtelaio singolo")
                ml_controtelaio_singolo = mir if tipo_controtelaio == "C. SINGOLO" else ""
                if ml_controtelaio_singolo_index < len(new_values):
                    new_values[ml_controtelaio_singolo_index] = ml_controtelaio_singolo
            except Exception as e:
                print(f"[DEBUG] Errore nell'aggiornamento di ml_controtelaio_singolo: {e}")

            # --- AGGIORNAMENTO AUTOMATICO COSTO CONTROTELAIO SINGOLO ---
            try:
                costo_controtelaio_singolo_index = self.colonne.index("Costo Controtelaio singolo")
                costo_controtelaio_singolo = costo_ml_controtelaio_singolo if tipo_controtelaio == "C. SINGOLO" else ""
                if costo_controtelaio_singolo_index < len(new_values):
                    new_values[costo_controtelaio_singolo_index] = costo_controtelaio_singolo
            except Exception as e:
                print(f"[DEBUG] Errore nell'aggiornamento di costo_controtelaio_singolo: {e}")

            # --- AGGIORNAMENTO AUTOMATICO N. CONTROTELAI DOPPI ---
            try:
                n_controtelai_doppi_index = self.colonne.index("N. Controtelai doppi")
                n_controtelai_doppi = "1" if tipo_controtelaio == "C. DOPPIO" else "0"
                if n_controtelai_doppi_index < len(new_values):
                    new_values[n_controtelai_doppi_index] = n_controtelai_doppi
            except Exception as e:
                print(f"[DEBUG] Errore nell'aggiornamento di n_controtelai_doppi: {e}")

            # --- AGGIORNAMENTO AUTOMATICO ML CONTROTELAIO DOPPIO ---
            try:
                ml_controtelaio_doppio_index = self.colonne.index("ML Controtelaio doppio")
                ml_controtelaio_doppio = mir if tipo_controtelaio == "C. DOPPIO" else ""
                if ml_controtelaio_doppio_index < len(new_values):
                    new_values[ml_controtelaio_doppio_index] = ml_controtelaio_doppio
            except Exception as e:
                print(f"[DEBUG] Errore nell'aggiornamento di ml_controtelaio_doppio: {e}")

            # --- AGGIORNAMENTO AUTOMATICO COSTO CONTROTELAIO DOPPIO ---
            try:
                costo_controtelaio_doppio_index = self.colonne.index("Costo Controtelaio doppio")
                costo_controtelaio_doppio = costo_ml_controtelaio_singolo if tipo_controtelaio == "C. DOPPIO" else ""
                if costo_controtelaio_doppio_index < len(new_values):
                    new_values[costo_controtelaio_doppio_index] = costo_controtelaio_doppio
            except Exception as e:
                print(f"[DEBUG] Errore nell'aggiornamento di costo_controtelaio_doppio: {e}")

            # --- AGGIORNAMENTO AUTOMATICO N. CONTROTELAIO TERMICO TIP A ---
            try:
                n_controtelaio_termico_a_index = self.colonne.index("N. Controtelaio termico TIP A")
                n_controtelaio_termico_a = "1" if tipo_controtelaio == "C. TERMICO TIP A" else "0"
                if n_controtelaio_termico_a_index < len(new_values):
                    new_values[n_controtelaio_termico_a_index] = n_controtelaio_termico_a
            except Exception as e:
                print(f"[DEBUG] Errore nell'aggiornamento di n_controtelaio_termico_a: {e}")

            # --- AGGIORNAMENTO AUTOMATICO COSTO LISTINO CONTROTELAIO TERMICO TIP A ---
            try:
                costo_listino_termico_a_index = self.colonne.index("Costo listino controtelaio termico TIP A")
                costo_listino_termico_a = costo_ml_controtelaio_singolo if tipo_controtelaio == "C. TERMICO TIP A" else ""
                if costo_listino_termico_a_index < len(new_values):
                    new_values[costo_listino_termico_a_index] = costo_listino_termico_a
            except Exception as e:
                print(f"[DEBUG] Errore nell'aggiornamento di costo_listino_termico_a: {e}")

            # --- AGGIORNAMENTO AUTOMATICO N. CONTROTELAIO TERMICO TIP B ---
            try:
                n_controtelaio_termico_b_index = self.colonne.index("N. Controtelaio termico TIP B")
                n_controtelaio_termico_b = "1" if tipo_controtelaio == "C. TERMICO TIP B" else "0"
                if n_controtelaio_termico_b_index < len(new_values):
                    new_values[n_controtelaio_termico_b_index] = n_controtelaio_termico_b
            except Exception as e:
                print(f"[DEBUG] Errore nell'aggiornamento di n_controtelaio_termico_b: {e}")

            # --- AGGIORNAMENTO AUTOMATICO COSTO LISTINO CONTROTELAIO TERMICO TIP B ---
            try:
                costo_listino_termico_b_index = self.colonne.index("Costo listino controtelaio termico TIP B")
                costo_listino_termico_b = costo_ml_controtelaio_singolo if tipo_controtelaio == "C. TERMICO TIP B" else ""
                if costo_listino_termico_b_index < len(new_values):
                    new_values[costo_listino_termico_b_index] = costo_listino_termico_b
            except Exception as e:
                print(f"[DEBUG] Errore nell'aggiornamento di costo_listino_termico_b: {e}")

            # --- FINE AGGIORNAMENTO AUTOMATICO CONTROTELAI ---
            
            # --- AGGIORNAMENTO AUTOMATICO SCONTI ---
            try:
                s1_index = self.colonne.index("Sconto 1")
                s2_index = self.colonne.index("Sconto 2")
                s3_index = self.colonne.index("Sconto 3")
                sdec_index = self.colonne.index("Sconto in decimali")
                ds_index = self.colonne.index("Dicitura sconto")
                if self.preventivo and hasattr(self.preventivo, 'dati_b1'):
                    dati_b1 = self.preventivo.dati_b1
                    def get_sconto_value(dati_b1, *keys):
                        for k in keys:
                            if k in dati_b1:
                                return dati_b1[k]
                        return ""
                    new_values[s1_index] = get_sconto_value(dati_b1, 'Sconto 1', 'Sconto1')
                    new_values[s2_index] = get_sconto_value(dati_b1, 'Sconto 2', 'Sconto2')
                    new_values[s3_index] = get_sconto_value(dati_b1, 'Sconto 3', 'Sconto3')
                    new_values[sdec_index] = get_sconto_value(dati_b1, 'Sconto in decimali', 'Sconto_in_decimali')
                    new_values[ds_index] = get_sconto_value(dati_b1, 'Dicitura sconto', 'Dicitura_sconto')
            except Exception as e:
                print(f"[DEBUG] Errore aggiornamento automatico sconti: {e}")
            # --- FINE AGGIORNAMENTO AUTOMATICO SCONTI ---
            
            # --- AGGIORNAMENTO AUTOMATICO PREZZO_LISTINO ---
            try:
                modello_index = self.colonne.index("Modello")
                colore_index = self.colonne.index("Colore")
                prezzo_listino_index = self.colonne.index("Prezzo_listino")
                modello = new_values[modello_index] if modello_index < len(new_values) else ""
                colore = new_values[colore_index] if colore_index < len(new_values) else ""
                prezzo_listino = ""
                if modello and colore:
                    row = listino[listino['MODELLO'] == modello]
                    if not row.empty and colore in listino.columns:
                        prezzo = row.iloc[0][colore]
                        if pd.notnull(prezzo):
                            try:
                                prezzo_float = float(prezzo)
                                prezzo_listino = f"€ {prezzo_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                            except Exception:
                                prezzo_listino = f"€ {prezzo},00"
                new_values[prezzo_listino_index] = prezzo_listino
            except Exception as e:
                print(f"[DEBUG] Errore aggiornamento automatico Prezzo_listino: {e}")
            # --- FINE AGGIORNAMENTO AUTOMATICO PREZZO_LISTINO ---
            
            # --- AGGIORNAMENTO AUTOMATICO MIN_FATT_PZ ---
            try:
                tabella_minimi_index = self.colonne.index("Tabella_minimi")
                min1_index = self.colonne.index("MIN.1")
                min2_index = self.colonne.index("MIN.2")
                min3_index = self.colonne.index("MIN.3")
                min4_index = self.colonne.index("MIN.4")
                min5_index = self.colonne.index("MIN.5")
                min_fatt_pz_index = self.colonne.index("Min_fatt_pz")
                tabella_val = str(new_values[tabella_minimi_index]).strip()
                print(f"[DEBUG] tabella_minimi: '{tabella_val}'")
                try:
                    tab_min = int(tabella_val)
                except Exception:
                    tab_min = None
                min_val = ""
                print(f"[DEBUG] MIN.1: {new_values[min1_index]}, MIN.2: {new_values[min2_index]}, MIN.3: {new_values[min3_index]}, MIN.4: {new_values[min4_index]}, MIN.5: {new_values[min5_index]}")
                if tab_min == 1:
                    min_val = new_values[min1_index]
                elif tab_min == 2:
                    min_val = new_values[min2_index]
                elif tab_min == 3:
                    min_val = new_values[min3_index]
                elif tab_min == 4:
                    min_val = new_values[min4_index]
                elif tab_min == 5:
                    min_val = new_values[min5_index]
                print(f"[DEBUG] min_fatt_pz calcolato: {min_val}")
                new_values[min_fatt_pz_index] = min_val
            except Exception:
                new_values[min_fatt_pz_index] = ""
            # --- FINE AGGIORNAMENTO AUTOMATICO MIN_FATT_PZ ---
            
            # --- AGGIORNAMENTO AUTOMATICO Mq_fatt_pz ---
            try:
                mqr_index = self.colonne.index("MqR")
                min_fatt_pz_index = self.colonne.index("Min_fatt_pz")
                mq_fatt_pz_index = self.colonne.index("Mq_fatt_pz")
                mqr_val = new_values[mqr_index] if mqr_index < len(new_values) else ""
                min_fatt_pz_val = new_values[min_fatt_pz_index] if min_fatt_pz_index < len(new_values) else ""
                mqr_float = float(mqr_val.replace(",", ".")) if mqr_val else 0.0
                min_fatt_pz_float = float(min_fatt_pz_val.replace(",", ".")) if min_fatt_pz_val else 0.0
                mq_fatt_pz = max(mqr_float, min_fatt_pz_float)
                new_values[mq_fatt_pz_index] = f"{mq_fatt_pz:.2f}".replace('.', ',')
            except Exception:
                new_values[mq_fatt_pz_index] = ""
            # --- FINE AGGIORNAMENTO AUTOMATICO Mq_fatt_pz ---
            
            # --- AGGIORNAMENTO AUTOMATICO Mq_totali_fatt ---
            try:
                mq_fatt_pz_index = self.colonne.index("Mq_fatt_pz")
                nr_pezzi_index = self.colonne.index("nr. pezzi")
                mq_totali_fatt_index = self.colonne.index("Mq_totali_fatt")
                mq_fatt_pz_val = new_values[mq_fatt_pz_index] if mq_fatt_pz_index < len(new_values) else ""
                nr_pezzi_val = new_values[nr_pezzi_index] if nr_pezzi_index < len(new_values) else ""
                mq_fatt_pz_float = float(mq_fatt_pz_val.replace(",", ".")) if mq_fatt_pz_val else 0.0
                nr_pezzi_float = float(nr_pezzi_val.replace(",", ".")) if nr_pezzi_val else 0.0
                mq_totali_fatt = mq_fatt_pz_float * nr_pezzi_float
                new_values[mq_totali_fatt_index] = f"{mq_totali_fatt:.2f}".replace('.', ',')
            except Exception:
                new_values[mq_totali_fatt_index] = ""
            # --- FINE AGGIORNAMENTO AUTOMATICO Mq_totali_fatt ---
            
            # --- AGGIORNAMENTO AUTOMATICO Ml_totali_fatt ---
            try:
                mir_index = self.colonne.index("MIR")
                nr_pezzi_index = self.colonne.index("nr. pezzi")
                ml_totali_fatt_index = self.colonne.index("Ml_totali_fatt")
                mir_val = new_values[mir_index] if mir_index < len(new_values) else ""
                nr_pezzi_val = new_values[nr_pezzi_index] if nr_pezzi_index < len(new_values) else ""
                mir_float = float(mir_val.replace(",", ".")) if mir_val else 0.0
                nr_pezzi_float = float(nr_pezzi_val.replace(",", ".")) if nr_pezzi_val else 0.0
                ml_totali_fatt = mir_float * nr_pezzi_float
                new_values[ml_totali_fatt_index] = f"{ml_totali_fatt:.2f}".replace('.', ',')
            except Exception:
                new_values[ml_totali_fatt_index] = ""
            # --- FINE AGGIORNAMENTO AUTOMATICO Ml_totali_fatt ---
            
            # --- AGGIORNAMENTO AUTOMATICO COLONNE CALCOLATE ---
            try:
                prezzo_listino_index = self.colonne.index("Prezzo_listino")
                sconto_decimali_index = self.colonne.index("Sconto in decimali")
                mq_fatt_pz_index = self.colonne.index("Mq_fatt_pz")
                mq_totali_fatt_index = self.colonne.index("Mq_totali_fatt")
                costo_scontato_mq_index = self.colonne.index("Costo_scontato_Mq")
                prezzo_listino_unitario_index = self.colonne.index("Prezzo_listino_unitario")
                costo_serramento_listino_posizione_index = self.colonne.index("Costo_serramento_listino_posizione")
                costo_serramento_scontato_posizione_index = self.colonne.index("Costo_serramento_scontato_posizione")

                prezzo_listino_val = new_values[prezzo_listino_index] if prezzo_listino_index < len(new_values) else ""
                if isinstance(prezzo_listino_val, str):
                    # Prendi solo la parte numerica (gestione "€" e formati)
                    import re
                    match = re.search(r"([\d.,]+)", prezzo_listino_val)
                    if match:
                        prezzo_listino_val = match.group(1).replace(".", "").replace(",", ".")
                    else:
                        prezzo_listino_val = "0"
                try:
                    prezzo_listino_float = float(prezzo_listino_val)
                except Exception:
                    prezzo_listino_float = 0.0
                sconto_decimali_val = new_values[sconto_decimali_index] if sconto_decimali_index < len(new_values) else "0"
                sconto_decimali_val = sconto_decimali_val.strip().replace("%", "")
                sconto_decimali_val = sconto_decimali_val.replace(",", ".")
                try:
                    sconto_decimali_float = float(sconto_decimali_val)
                    if sconto_decimali_float > 1:
                        sconto_decimali_float = sconto_decimali_float / 100.0
                except Exception:
                    sconto_decimali_float = 0.0
                logging.debug("SCONTO_DECIMALI ORIGINALE: '%s' -> float: %s", new_values[sconto_decimali_index], sconto_decimali_float)
                mq_fatt_pz_val = new_values[mq_fatt_pz_index] if mq_fatt_pz_index < len(new_values) else ""
                mq_fatt_pz_val = mq_fatt_pz_val.replace(",", ".") if isinstance(mq_fatt_pz_val, str) else "0"
                try:
                    mq_fatt_pz_float = float(mq_fatt_pz_val)
                except Exception:
                    mq_fatt_pz_float = 0.0
                mq_totali_fatt_val = new_values[mq_totali_fatt_index] if mq_totali_fatt_index < len(new_values) else ""
                mq_totali_fatt_val = mq_totali_fatt_val.replace(",", ".") if isinstance(mq_totali_fatt_val, str) else "0"
                try:
                    mq_totali_fatt_float = float(mq_totali_fatt_val)
                except Exception:
                    mq_totali_fatt_float = 0.0

                costo_scontato_mq = prezzo_listino_float * (1 - sconto_decimali_float)
                prezzo_listino_unitario = prezzo_listino_float * mq_fatt_pz_float
                costo_serramento_listino_posizione = prezzo_listino_float * mq_totali_fatt_float
                costo_serramento_scontato_posizione = costo_scontato_mq * mq_totali_fatt_float

                def euro(val):
                    return f"€ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if val else ""
                if costo_scontato_mq_index < len(new_values):
                    new_values[costo_scontato_mq_index] = euro(costo_scontato_mq)
                if prezzo_listino_unitario_index < len(new_values):
                    new_values[prezzo_listino_unitario_index] = euro(prezzo_listino_unitario)
                if costo_serramento_listino_posizione_index < len(new_values):
                    new_values[costo_serramento_listino_posizione_index] = euro(costo_serramento_listino_posizione)
                if costo_serramento_scontato_posizione_index < len(new_values):
                    new_values[costo_serramento_scontato_posizione_index] = euro(costo_serramento_scontato_posizione)
            except Exception as e:
                print(f"[DEBUG] Errore colonne calcolate: {e}")
                if 'costo_scontato_mq_index' in locals() and costo_scontato_mq_index < len(new_values):
                    new_values[costo_scontato_mq_index] = ""
                if 'prezzo_listino_unitario_index' in locals() and prezzo_listino_unitario_index < len(new_values):
                    new_values[prezzo_listino_unitario_index] = ""
                if 'costo_serramento_listino_posizione_index' in locals() and costo_serramento_listino_posizione_index < len(new_values):
                    new_values[costo_serramento_listino_posizione_index] = ""
                if 'costo_serramento_scontato_posizione_index' in locals() and costo_serramento_scontato_posizione_index < len(new_values):
                    new_values[costo_serramento_scontato_posizione_index] = ""
            
            # --- AGGIORNAMENTO AUTOMATICO CAMPI DISTANZIALI/IMBOTTI ---
            try:
                # Indici dei campi
                dist_index = self.colonne.index("Dist.")
                tipo_dist_index = self.colonne.index("Tipo dist.")
                distanziali_imbotti_index = self.colonne.index("Distanziali/Imbotti")
                tipo_distanziali_imbotti_index = self.colonne.index("Tipo distanziali/imbotti")
                dicitura_distanziale_imbotte_index = self.colonne.index("Dicitura distanziale/imbotte")
                colore_index = self.colonne.index("Colore")
                colore_dist_imb_index = self.colonne.index("Colore dist/imb")

                # Calcolo "Distanziali/Imbotti" = SE(O([Dist.]=="",[Dist.]=="NO"),0,1)
                dist_val = new_values[dist_index] if dist_index < len(new_values) else ""
                distanziali_imbotti = "0"
                if dist_val and dist_val != "NO":
                    distanziali_imbotti = "1"
                if distanziali_imbotti_index < len(new_values):
                    new_values[distanziali_imbotti_index] = distanziali_imbotti
                
                # Calcolo "Tipo distanziali/imbotti" = SE([@[Tipo dist.]]="IMBOTTE";2;SE([@[Tipo dist.]]="SALDATO";3;1))
                tipo_dist_val = new_values[tipo_dist_index] if tipo_dist_index < len(new_values) else ""
                tipo_distanziali_imbotti = "1"  # Valore predefinito
                if tipo_dist_val == "IMBOTTE":
                    tipo_distanziali_imbotti = "2"
                elif tipo_dist_val == "SALDATO":
                    tipo_distanziali_imbotti = "3"
                if tipo_distanziali_imbotti_index < len(new_values):
                    new_values[tipo_distanziali_imbotti_index] = tipo_distanziali_imbotti
            except Exception as e:
                print(f"[DEBUG] Errore nell'aggiornamento dei campi distanziali/imbotti: {e}")
            
            # --- AGGIORNAMENTO AUTOMATICO N. DISTANZIALI A 3 LATI ---
            try:
                n_distanziali_a_3_lati_index = self.colonne.index("N. distanziali a 3 lati")
                nr_pezzi_index = self.colonne.index("nr. pezzi")
                dist_index = self.colonne.index("Dist.")
                
                nr_pezzi_val = new_values[nr_pezzi_index] if nr_pezzi_index < len(new_values) else ""
                dist_val = new_values[dist_index] if dist_index < len(new_values) else ""
                
                n_distanziali_a_3_lati = ""
                # Calcolo "N. distanziali a 3 lati" = SE([@[Distanziali/Imbotti]]=1;[@[nr. pezzi]]*3;"")
                # Dove Distanziali/Imbotti = 1 se Dist. è valorizzato e diverso da "NO"
                if dist_val and dist_val != "NO":
                    try:
                        n_pezzi_val = float(str(nr_pezzi_val).replace(",", ".")) if nr_pezzi_val else 0
                        n_distanziali_a_3_lati = str(int(n_pezzi_val * 3))
                    except (ValueError, TypeError):
                        n_distanziali_a_3_lati = ""
                
                if n_distanziali_a_3_lati_index < len(new_values):
                    new_values[n_distanziali_a_3_lati_index] = n_distanziali_a_3_lati
                    print(f"[DEBUG] N. distanziali a 3 lati aggiornato a: {n_distanziali_a_3_lati}")
            except Exception as e:
                print(f"[DEBUG] Errore nell'aggiornamento di N. distanziali a 3 lati: {e}")
                
            # --- AGGIORNAMENTO AUTOMATICO DICITURA DISTANZIALE/IMBOTTE ---
            try:
                dicitura_distanziale_imbotte_index = self.colonne.index("Dicitura distanziale/imbotte")
                tipo_distanziali_imbotti_index = self.colonne.index("Tipo distanziali/imbotti")
                
                tipo_distanziali_imbotti = new_values[tipo_distanziali_imbotti_index] if tipo_distanziali_imbotti_index < len(new_values) else ""
                
                # Calcolo "Dicitura distanziale/imbotte" = SE([@[Tipo distanziali/imbotti]]=3;"DS - SALDATO";SE([@[Tipo distanziali/imbotti]]=2;"I - IMBOTTE";""))
                dicitura_distanziale_imbotte = ""
                if tipo_distanziali_imbotti == "3":
                    dicitura_distanziale_imbotte = "DS - SALDATO"
                elif tipo_distanziali_imbotti == "2":
                    dicitura_distanziale_imbotte = "I - IMBOTTE"
                
                if dicitura_distanziale_imbotte_index < len(new_values):
                    new_values[dicitura_distanziale_imbotte_index] = dicitura_distanziale_imbotte
            except Exception as e:
                print(f"[DEBUG] Errore nell'aggiornamento di Dicitura distanziale/imbotte: {e}")
                
            # --- LOGICA COLORE DIST/IMB anche in modifica ---
            try:
                colore_index = self.colonne.index("Colore")
                colore_dist_imb_index = self.colonne.index("Colore dist/imb")
                
                colore_val = new_values[colore_index] if colore_index < len(new_values) else ""
                colore_dist_imb = ""
                if colore_val == "STANDARD RAL":
                    colore_dist_imb = "1"
                elif colore_val == "EFFETTO LEGNO":
                    colore_dist_imb = "2"
                elif colore_val == "GREZZO":
                    colore_dist_imb = "3"
                elif colore_val == "EXTRA MAZZETTA":
                    colore_dist_imb = "4"
                if colore_dist_imb_index < len(new_values):
                    new_values[colore_dist_imb_index] = colore_dist_imb
            except Exception as e:
                print(f"[DEBUG] Errore nell'aggiornamento di Colore dist/imb: {e}")
            # --- FINE LOGICA COLORE DIST/IMB ---
            
            # --- AGGIORNAMENTO AUTOMATICO CAMPI DA N.ANTE A DICITURA SCONTO E UNITA_DI_MISURA ---
            try:
                # N.ANTE, AP., TIP., DESCR.TIP, N.CERN., XLAV, MULT., COSTO MLT, MIN.1-5, P.SERR., N.STAFF.
                # (già gestiti da entries, ma se vuoti e disponibili nel DataFrame, aggiorna)
                modello_index = self.colonne.index("Modello")
                serramento_index = self.colonne.index("Serramento")
                modello_val = new_values[modello_index] if modello_index < len(new_values) else ""
                serramento_val = new_values[serramento_index] if serramento_index < len(new_values) else ""
                if self.dataframe is not None and serramento_val:
                    riga = self.dataframe[self.dataframe.iloc[:, 0] == serramento_val]
                    if not riga.empty:
                        mapping = {
                            "NUMERO ANTE": "N.ANTE",
                            "APERTURA": "AP.",
                            "TIPOLOGIA": "TIP.",
                            "DESCRIZIONE TIPOLOGIA": "DESCR.TIP",
                            "N. CERNIERE": "N.CERN.",
                            "EXTRA LAVORAZIONE": "XLAV",
                            "MULTIPLO": "MULT.",
                            "COSTO MULTIPLO": "COSTO MLT",
                            "MINIMI 1": "MIN.1",
                            "MINIMI 2": "MIN.2",
                            "MINIMI 3": "MIN.3",
                            "MINIMI 4": "MIN.4",
                            "MINIMI 5": "MIN.5",
                            "Presenza serratura": "P.SERR.",
                            "Numero staffette": "N.STAFF."
                        }
                        for df_col, col_name in mapping.items():
                            idx = self.colonne.index(col_name)
                            if idx < len(new_values) and (not new_values[idx] or new_values[idx] == ""):
                                try:
                                    value = riga[df_col].iloc[0]
                                    print(f"[DEBUG] Mapping {df_col} -> {col_name}: {value}")
                                    new_values[idx] = str(value)
                                except Exception:
                                    pass
                # Unita_di_misura
                unita_idx = self.colonne.index("Unita_di_misura")
                unita_di_misura = ""
                try:
                    if modello_val:
                        riga_listino = listino[listino['MODELLO'] == modello_val]
                        if not riga_listino.empty:
                            for col in riga_listino.columns:
                                if col.lower().replace('à','a').replace(' ','_') == 'unita_di_misura':
                                    unita_di_misura = riga_listino.iloc[0][col]
                                    break
                except Exception as e:
                    unita_di_misura = f"Errore: {e}"
                if unita_idx < len(new_values):
                    new_values[unita_idx] = unita_di_misura
            except Exception as e:
                print(f"[DEBUG] Errore aggiornamento automatico campi da N.ANTE a Unita_di_misura: {e}")
            # --- FINE AGGIORNAMENTO ---
            
            # --- DEBUG LOG: valori prima di aggiornare la riga in save_edited_row ---
            print("[DEBUG save_edited_row] new_values prima dell'update:", new_values)
            
            # Aggiorna la riga nel Treeview
            self.tree.item(item, values=new_values)
            
            # Chiudi la finestra di dialogo
            dialog.destroy()
            
            # Aggiorna l'oggetto preventivo
            self.salva_in_preventivo()
            # Aggiorna tutti i campi controtelaio per tutte le righe
            self.aggiorna_tutti_campi_controtelaio_treeview()
            
        except Exception as e:
            messagebox.showerror("Errore", f"Si è verificato un errore durante il salvataggio: {str(e)}")

    def on_double_click(self, event):
        """
        Gestisce il doppio click su una riga del Treeview.
        """
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell":
            self.modifica_riga_selezionata()

    def show_context_menu(self, event):
        """
        Mostra il menu contestuale quando si fa click con il tasto destro su una riga.
        """
        # Seleziona la riga cliccata
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            # Mostra il menu contestuale nella posizione del click
            self.context_menu.post(event.x_root, event.y_root)

    def load_data_from_preventivo(self):
        """
        Carica i dati dal file preventivo.py se esiste.
        """
        try:
            # Verifica se il file preventivo.py esiste
            if os.path.exists("preventivo.py"):
                # Importa il modulo preventivo
                import importlib.util
                spec = importlib.util.spec_from_file_location("preventivo", "preventivo.py")
                preventivo = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(preventivo)
                
                # Verifica se la variabile dati_posizioni esiste nel modulo
                if hasattr(preventivo, "dati_posizioni"):
                    # Trova il valore massimo di pos per impostare il contatore
                    max_pos = 0
                    
                    # Carica i dati nel Treeview
                    for dati in preventivo.dati_posizioni:
                        values = [
                            dati.get('pos', ''),
                            dati.get('nr_pezzi', ''),
                            dati.get('serramento', ''),
                            dati.get('modello', ''),
                            dati.get('modello_grata_combinato', ''),
                            dati.get('colore', ''),
                            dati.get('l_mm', ''),
                            dati.get('h_mm', ''),
                            dati.get('tipo_telaio', ''),
                            dati.get('bunk', ''),
                            dati.get('dmcp', ''),
                            dati.get('defender', ''),
                            dati.get('dist', ''),
                            dati.get('tipo_dist', ''),
                            dati.get('l_dist', ''),
                            dati.get('h_dist', ''),
                            dati.get('tipo_controtelaio', ''),
                            dati.get('anta_giro', ''),
                            dati.get('mrib', ''),
                            dati.get('NUMERO ANTE', ''),
                            dati.get('APERTURA', ''),
                            dati.get('TIPOLOGIA', ''),
                            dati.get('DESCRIZIONE TIPOLOGIA', ''),
                            dati.get('N. CERNIERE', ''),
                            dati.get('EXTRA LAVORAZIONE', ''),
                            dati.get('MULTIPLO', ''),
                            dati.get('COSTO MULTIPLO', ''),
                            dati.get('MINIMI 1', ''),
                            dati.get('MINIMI 2', ''),
                            dati.get('MINIMI 3', ''),
                            dati.get('MINIMI 4', ''),
                            dati.get('MINIMI 5', ''),
                            dati.get('Presenza serratura', ''),
                            dati.get('Numero staffette', ''),
                            dati.get('Sconto 1', ''),
                            dati.get('Sconto 2', ''),
                            dati.get('Sconto 3', ''),
                            dati.get('Sconto in decimali', ''),
                            dati.get('Dicitura sconto', ''),
                            dati.get('Prezzo_listino', ''),
                            dati.get('MqR', ''),
                            dati.get('MIR', ''),
                            dati.get('Tabella_minimi', ''),
                            dati.get('Unita_di_misura', ''),
                            dati.get('Min_fatt_pz', ''),
                            dati.get('Mq_fatt_pz', ''),
                            dati.get('Mq_totali_fatt', ''),
                            dati.get('Ml_totali_fatt', ''),
                            dati.get('Costo_scontato_Mq', ''),
                            dati.get('Prezzo_listino_unitario', ''),
                            dati.get('Costo_serramento_listino_posizione', ''),
                            dati.get('Costo_serramento_scontato_posizione', '')
                        ]
                        
                        # Assicura che la lista abbia esattamente len(self.colonne) elementi
                        if len(values) < len(self.colonne):
                            values.extend([""] * (len(self.colonne) - len(values)))
                        elif len(values) > len(self.colonne):
                            values = values[:len(self.colonne)]
                        self.tree.insert("", "end", values=values)
                        
                        # Aggiorna il valore massimo di pos
                        pos = dati.get('pos', 0)
                        if isinstance(pos, (int, float)) and pos > max_pos:
                            max_pos = pos
                    
                    # Imposta il contatore della posizione
                    self.pos_counter = max_pos + 1
                    self.pos_label.config(text=str(self.pos_counter))
        except Exception as e:
            print(f"Errore durante il caricamento dei dati dal file preventivo.py: {e}")
            
    def aggiorna_da_preventivo(self):
        """Aggiorna la tabella delle posizioni leggendo i dati dal preventivo centrale."""
        if hasattr(self, 'tree'):
            self.tree.delete(*self.tree.get_children())
            if self.preventivo and hasattr(self.preventivo, 'posizioni') and self.preventivo.posizioni:
                for posizione in self.preventivo.posizioni:
                    values = [posizione.get(col, "") for col in self.colonne]
                    if len(values) < len(self.colonne):
                        values.extend([""] * (len(self.colonne) - len(values)))
                    elif len(values) > len(self.colonne):
                        values = values[:len(self.colonne)]
                    self.tree.insert("", "end", values=values)
                self.aggiorna_contatore_posizione()

    def get_all_posizioni(self):
        """Restituisce tutte le posizioni attualmente inserite nella tabella come lista di dict."""
        posizioni = []
        if hasattr(self, 'tree'):
            for item in self.tree.get_children():
                valori = self.tree.item(item, 'values')
                posizione_dict = {col: valori[idx] if idx < len(valori) else "" for idx, col in enumerate(self.colonne)}
                posizioni.append(posizione_dict)
        return posizioni

    def salva_in_preventivo(self):
        """Aggiorna l'oggetto preventivo con i dati attuali della tabella."""
        if self.preventivo:
            self.preventivo.posizioni = self.get_all_posizioni()
            if self.app:
                self.app.preventivo_corrente.modificato = True

    # (Opzionale) Da chiamare dopo ogni modifica alla tabella per mantenere sincronizzato il preventivo
    def on_tabella_modificata(self, event=None):
        self.salva_in_preventivo()

    # Sostituisci la vecchia load_data_from_preventivo con una chiamata a aggiorna_da_preventivo
    def load_data_from_preventivo(self):
        self.aggiorna_da_preventivo()

    # (Eventuali altri metodi già presenti restano invariati)
    
    def _salva_preventivo_menu(self):
        """Salva il preventivo corrente tramite l'app principale."""
        if self.app and hasattr(self.app, '_salva_preventivo'):
            try:
                result = self.app._salva_preventivo()
                if result:
                    messagebox.showinfo("Salvataggio riuscito", "Preventivo salvato correttamente.")
                else:
                    messagebox.showerror("Errore", "Errore durante il salvataggio del preventivo.")
            except Exception as e:
                messagebox.showerror("Errore", f"Errore durante il salvataggio del preventivo: {str(e)}")
        else:
            messagebox.showerror("Errore", "Impossibile accedere alla funzione di salvataggio dell'app principale.")
        
    def aggiorna_colore_infissi_generale(self, colore_infissi=None):
        """
        Aggiorna la variabile globale colore_infissi_generale e la combobox 'Colore' in base
        al valore selezionato in 'Colore Infissi' nel modulo dati generali (ModuloB2Frame).
        Se non viene passato nessun valore, mantiene il valore attuale.
        """
        global colore_infissi_generale
        if colore_infissi is not None:
            colore_infissi_generale = colore_infissi
            # Aggiorna la combobox solo se il valore è tra quelli previsti
            if colore_infissi in self.colore_combobox['values']:
                self.colore_combobox.set(colore_infissi)
            else:
                self.colore_combobox.set('STANDARD RAL')
        else:
            # Se non viene passato nulla, imposta il valore globale nella combobox
            self.colore_combobox.set(colore_infissi_generale)

    def collega_modulo_b2(self, modulo_b2):
        """
        Collega il modulo_b2 (ModuloB2Frame) e imposta il callback per la selezione del colore infissi.
        """
        self.modulo_b2 = modulo_b2
        # Collegamento diretto: quando cambia il colore infissi in ModuloB2Frame, aggiorna anche qui
        if hasattr(modulo_b2, 'colore_infissi_cb'):
            modulo_b2.colore_infissi_cb.bind('<<ComboboxSelected>>', self._on_colore_infissi_generale_changed)
            # Imposta subito il colore iniziale
            self.aggiorna_colore_infissi_generale(modulo_b2.colore_infissi_cb.get())

    def _on_colore_infissi_generale_changed(self, event=None):
        """
        Callback che aggiorna il colore generale quando cambia la selezione nel modulo B2.
        """
        if hasattr(self, 'modulo_b2') and hasattr(self.modulo_b2, 'colore_infissi_cb'):
            colore = self.modulo_b2.colore_infissi_cb.get()
            self.aggiorna_colore_infissi_generale(colore)

    def duplica_riga_selezionata(self):
        """
        Duplica la riga attualmente selezionata nel Treeview, inserendo la copia subito dopo la riga madre e rinumerando tutte le righe per mantenere la coerenza della colonna Pos.
        """
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Nessuna selezione", "Seleziona una riga da duplicare.")
            return
        
        item = selected[0]
        values = list(self.tree.item(item, 'values'))
        all_items = self.tree.get_children()
        idx = all_items.index(item)
        # Inserisci la riga duplicata subito dopo la riga madre
        self.tree.insert('', idx + 1, values=values)
        # Rinumera tutte le righe in modo sequenziale
        for i, item_id in enumerate(self.tree.get_children(), start=1):
            vals = list(self.tree.item(item_id, "values"))
            vals[0] = i
            self.tree.item(item_id, values=vals)
        self.aggiorna_contatore_posizione()
        self.salva_in_preventivo()

    def aggiorna_contatore_posizione(self):
        """Aggiorna il contatore posizione in base al massimo valore di Pos. presente nel Treeview."""
        max_pos = 0
        for item in self.tree.get_children():
            try:
                pos = int(self.tree.item(item)["values"][0])
                if pos > max_pos:
                    max_pos = pos
            except Exception:
                continue
        self.pos_counter = max_pos + 1
        self.pos_label.config(text=str(self.pos_counter))

    def update_pos_label(self):
        """Aggiorna il campo Pos. ogni volta che cambia self.pos_counter."""
        self.pos_label.config(text=str(self.pos_counter))

    def aggiorna_tutti_gli_sconti_treeview(self):
        """
        Aggiorna tutti i campi sconto di tutte le righe del Treeview con i valori attuali di self.preventivo.dati_b1.
        """
        def format_percent(val):
            try:
                v = str(val).replace('%', '').replace(',', '.').strip()
                if v == '' or v.lower() == 'nan':
                    return ''
                v = float(v)
                return f"{v:.2f} %".replace('.', ',')
            except Exception:
                return str(val)
        def format_dicitura_sconto(dicitura):
            def repl(match):
                try:
                    v = float(match.group(1).replace(',', '.'))
                    return f"{v:.2f} %".replace('.', ',')
                except Exception:
                    return match.group()
            # Sostituisce ogni numero (intero o decimale) non già seguito da % con la formattazione corretta
            return re.sub(r'\b\d+[\.,]?\d*\b(?!\s*%)', repl, str(dicitura))
        if not (self.preventivo and hasattr(self.preventivo, 'dati_b1')):
            print("[DEBUG aggiorna_tutti_gli_sconti_treeview] Nessun dati_b1 presente nel preventivo.")
            return
        dati_b1 = self.preventivo.dati_b1
        def get_sconto_value(dati_b1, *keys):
            for k in keys:
                if k in dati_b1:
                    return dati_b1[k]
            return ""
        sconto_1 = format_percent(get_sconto_value(dati_b1, 'Sconto 1', 'Sconto1'))
        sconto_2 = format_percent(get_sconto_value(dati_b1, 'Sconto 2', 'Sconto2'))
        sconto_3 = format_percent(get_sconto_value(dati_b1, 'Sconto 3', 'Sconto3'))
        sconto_decimali = format_percent(get_sconto_value(dati_b1, 'Sconto in decimali', 'Sconto_in_decimali'))
        dicitura_sconto = format_dicitura_sconto(get_sconto_value(dati_b1, 'Dicitura sconto', 'Dicitura_sconto'))
        print(f"[DEBUG aggiorna_tutti_gli_sconti_treeview] Nuovi valori: S1={sconto_1}, S2={sconto_2}, S3={sconto_3}, Sdec={sconto_decimali}, Dic={dicitura_sconto}")
        for item in self.tree.get_children():
            values = list(self.tree.item(item, "values"))
            try:
                s1_index = self.colonne.index("Sconto 1")
                s2_index = self.colonne.index("Sconto 2")
                s3_index = self.colonne.index("Sconto 3")
                sdec_index = self.colonne.index("Sconto in decimali")
                ds_index = self.colonne.index("Dicitura sconto")
                values[s1_index] = sconto_1
                values[s2_index] = sconto_2
                values[s3_index] = sconto_3
                values[sdec_index] = sconto_decimali
                values[ds_index] = dicitura_sconto
                self.tree.item(item, values=values)
            except Exception as e:
                print(f"[DEBUG aggiorna_tutti_gli_sconti_treeview] Errore aggiornando item {item}: {e}")
        print("[DEBUG aggiorna_tutti_gli_sconti_treeview] Aggiornamento completato.")
        self.restore_treeview_column_widths()

    def save_treeview_column_widths(self):
        """Salva le larghezze delle colonne del Treeview in un attributo."""
        if not hasattr(self, '_treeview_col_widths'):
            self._treeview_col_widths = {}
        for col in self.tree['columns']:
            self._treeview_col_widths[col] = self.tree.column(col)['width']

    def restore_treeview_column_widths(self):
        """Ripristina le larghezze delle colonne del Treeview se sono state salvate."""
        if hasattr(self, '_treeview_col_widths'):
            # Definisci colonne che possono essere più strette
            colonne_strette = ["Pos.", "nr. pezzi", "L (mm)", "H (mm)", "Defender", "Dist.", "M.rib."]
            
            for col, width in self._treeview_col_widths.items():
                try:
                    # Imposta la larghezza minima in base al tipo di colonna
                    if col in colonne_strette:
                        min_width = 22  # Minimo per colonne strette
                    else:
                        min_width = 60  # Minimo per colonne normali
                        
                    self.tree.column(col, width=width, minwidth=min_width, stretch=False)
                except Exception:
                    pass

    def bind_treeview_column_resize(self):
        """Associa il salvataggio delle larghezze colonne dopo ogni ridimensionamento."""
        def on_release(event):
            # Salva le larghezze delle colonne
            self.save_treeview_column_widths()
            
            # Definisci colonne che possono essere più strette
            colonne_strette = ["Pos.", "nr. pezzi", "L (mm)", "H (mm)", "Defender", "Dist.", "M.rib."]
            
            # Assicura che tutte le colonne mantengano la loro larghezza minima appropriata
            for col in self.tree['columns']:
                current_width = self.tree.column(col, 'width')
                
                # Imposta la larghezza minima in base al tipo di colonna
                if col in colonne_strette:
                    min_width = 22  # Minimo per colonne strette
                else:
                    min_width = 60  # Minimo per colonne normali
                
                # Mantieni la larghezza corrente, ma assicurati che rispetti il minimo
                self.tree.column(col, width=current_width, minwidth=min_width, stretch=False)
                
        self.tree.bind('<ButtonRelease-1>', on_release)

    def aggiorna_tutti_campi_controtelaio_treeview(self):
        """Aggiorna tutti i campi relativi al controtelaio nel treeview."""
        for item in self.tree.get_children():
            try:
                values = self.tree.item(item)['values']
                if not values:
                    continue

                # Ottieni i valori necessari
                tipo_controtelaio = values[self.colonne.index("Tipologia controtelaio")] if "Tipologia controtelaio" in self.colonne else ""
                nr_pezzi = values[self.colonne.index("nr. pezzi")] if "nr. pezzi" in self.colonne else ""
                mir = values[self.colonne.index("Ml_totali_fatt")] if "Ml_totali_fatt" in self.colonne else ""
                costo_ml_controtelaio_singolo = values[self.colonne.index("Costo al ml controtelaio singolo")] if "Costo al ml controtelaio singolo" in self.colonne else ""

                def set_val(col, val):
                    try:
                        col_index = self.colonne.index(col)
                        if col_index < len(values):
                            values[col_index] = val
                    except Exception as e:
                        print(f"[DEBUG] Errore nell'aggiornamento di {col}: {e}")

                # Calcolo Tipologia ml/nr. Pezzi
                tipologia_ml_nr_pezzi = ""
                try:
                    if tipo_controtelaio:
                        row = controtelaio[controtelaio["CONTROTELAIO"] == tipo_controtelaio]
                        if not row.empty:
                            tipologia_ml_nr_pezzi = str(row.iloc[0]["Ml / nr. Pezzi"])
                except Exception as e:
                    print(f"[DEBUG] Errore nel calcolo della tipologia ml/nr. pezzi: {e}")
                    tipologia_ml_nr_pezzi = ""

                set_val("Tipologia ml/nr. Pezzi", tipologia_ml_nr_pezzi)

                # Calcolo Fattore moltiplicatore ml/nr. Pezzi
                fattore_moltiplicatore = "1"
                try:
                    if tipo_controtelaio and tipo_controtelaio != "0":
                        if tipologia_ml_nr_pezzi == "1":
                            fattore_moltiplicatore = "1"
                        elif tipologia_ml_nr_pezzi == "2":
                            try:
                                ml_totali = float(mir.replace(",", ".")) if mir else 0
                                nr_pezzi_val = float(nr_pezzi.replace(",", ".")) if nr_pezzi else 0
                                if nr_pezzi_val > 0:
                                    fattore_moltiplicatore = str(round(ml_totali / nr_pezzi_val, 2))
                                else:
                                    fattore_moltiplicatore = "1"
                            except ValueError:
                                print("[DEBUG] Ml_totali_fatt non è un numero valido: usando nr. pezzi = 1")
                                fattore_moltiplicatore = "1"
                    else:
                        fattore_moltiplicatore = "1"
                except Exception as e:
                    print(f"[DEBUG] Errore nel calcolo del fattore moltiplicatore: {e}")
                    fattore_moltiplicatore = "1"

                set_val("Fattore moltiplicatore ml/nr. Pezzi", fattore_moltiplicatore)

                # Calcolo N. Controtelai singoli
                n_controtelai_singoli = "0"
                try:
                    if tipo_controtelaio == "C. SINGOLO":
                        n_controtelai_singoli = str(nr_pezzi)
                except Exception as e:
                    print(f"[DEBUG] Errore nel calcolo del numero controtelai singoli: {e}")
                    n_controtelai_singoli = "0"

                set_val("N. Controtelai singoli", n_controtelai_singoli)

                # ML Controtelaio singolo
                ml_controtelaio_singolo = "0"
                try:
                    if tipo_controtelaio == "C. SINGOLO":
                        ml_controtelaio_singolo = mir if mir else "0"
                except Exception as e:
                    print(f"[DEBUG] Errore nell'aggiornamento di ml_controtelaio_singolo: {e}")
                    ml_controtelaio_singolo = "0"

                set_val("ML Controtelaio singolo", ml_controtelaio_singolo)

                # Costo Controtelaio singolo
                costo_controtelaio_singolo = "0"
                try:
                    if tipo_controtelaio == "C. SINGOLO":
                        try:
                            # Converti i valori in float, gestendo la formattazione italiana
                            costo_ml = float(str(costo_ml_controtelaio_singolo).replace("€", "").replace(".", "").replace(",", ".").strip()) if costo_ml_controtelaio_singolo else 0
                            ml = float(str(ml_controtelaio_singolo).replace(",", ".").strip()) if ml_controtelaio_singolo else 0
                            # Calcola il costo totale
                            costo_totale = costo_ml * ml
                            # Formatta il risultato in euro
                            costo_controtelaio_singolo = f"€ {costo_totale:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                        except ValueError as ve:
                            print(f"[DEBUG] Errore nella conversione dei valori numerici: {ve}")
                            costo_controtelaio_singolo = "€ 0,00"
                except Exception as e:
                    print(f"[DEBUG] Errore nell'aggiornamento di costo_controtelaio_singolo: {e}")
                    costo_controtelaio_singolo = "€ 0,00"

                set_val("Costo Controtelaio singolo", costo_controtelaio_singolo)

                # N. Controtelai doppi
                n_controtelai_doppi = "0"
                try:
                    if tipo_controtelaio == "C. DOPPIO":
                        n_controtelai_doppi = str(nr_pezzi)
                except Exception as e:
                    print(f"[DEBUG] Errore nell'aggiornamento di n_controtelai_doppi: {e}")
                    n_controtelai_doppi = "0"

                set_val("N. Controtelai doppi", n_controtelai_doppi)

                # ML Controtelaio doppio
                ml_controtelaio_doppio = "0"
                try:
                    if tipo_controtelaio == "C. DOPPIO":
                        ml_controtelaio_doppio = mir if mir else "0"
                except Exception as e:
                    print(f"[DEBUG] Errore nell'aggiornamento di ml_controtelaio_doppio: {e}")
                    ml_controtelaio_doppio = "0"

                set_val("ML Controtelaio doppio", ml_controtelaio_doppio)

                # Costo Controtelaio doppio
                costo_controtelaio_doppio = ""
                try:
                    if tipo_controtelaio == "C. DOPPIO":
                        costo_controtelaio_doppio = costo_ml_controtelaio_singolo
                except Exception as e:
                    print(f"[DEBUG] Errore nell'aggiornamento di costo_controtelaio_doppio: {e}")

                set_val("Costo Controtelaio doppio", costo_controtelaio_doppio)

                # N. Controtelaio termico TIP A
                n_controtelaio_termico_a = "0"
                try:
                    if tipo_controtelaio == "C. TERMICO TIP A":
                        n_controtelaio_termico_a = "1"
                except Exception as e:
                    print(f"[DEBUG] Errore nell'aggiornamento di n_controtelaio_termico_a: {e}")

                set_val("N. Controtelaio termico TIP A", n_controtelaio_termico_a)

                # Costo listino controtelaio termico TIP A
                costo_listino_termico_a = ""
                try:
                    if tipo_controtelaio == "C. TERMICO TIP A":
                        costo_listino_termico_a = costo_ml_controtelaio_singolo
                except Exception as e:
                    print(f"[DEBUG] Errore nell'aggiornamento di costo_listino_termico_a: {e}")

                set_val("Costo listino controtelaio termico TIP A", costo_listino_termico_a)

                # N. Controtelaio termico TIP B
                n_controtelaio_termico_b = "0"
                try:
                    if tipo_controtelaio == "C. TERMICO TIP B":
                        n_controtelaio_termico_b = "1"
                except Exception as e:
                    print(f"[DEBUG] Errore nell'aggiornamento di n_controtelaio_termico_b: {e}")

                set_val("N. Controtelaio termico TIP B", n_controtelaio_termico_b)

                # Costo listino controtelaio termico TIP B
                costo_listino_termico_b = ""
                try:
                    if tipo_controtelaio == "C. TERMICO TIP B":
                        costo_listino_termico_b = costo_ml_controtelaio_singolo
                except Exception as e:
                    print(f"[DEBUG] Errore nell'aggiornamento di costo_listino_termico_b: {e}")

                set_val("Costo listino controtelaio termico TIP B", costo_listino_termico_b)

                # Aggiorna i valori nel treeview
                self.tree.item(item, values=values)

            except Exception as e:
                print(f"[DEBUG] Errore nell'aggiornamento dei campi controtelaio: {e}")

    def on_header_scroll(self, *args):
        """Gestisce lo scrolling degli header sincronizzandolo con il treeview."""
        self.tree.xview(*args)
        self.draw_custom_headers()

    def draw_custom_headers(self):
        """Disegna gli header personalizzati per i gruppi di colonne."""
        try:
            # Ottieni le dimensioni del treeview
            treeview_width = self.tree.winfo_width()
            if treeview_width <= 0:
                treeview_width = 1455  # Larghezza di default
            
            # Calcola la larghezza totale delle colonne
            total_width = 0
            column_widths = []
            for col in self.tree['columns']:
                width = self.tree.column(col)['width']
                column_widths.append(width)
                total_width += width
            
            # Imposta la larghezza del canvas uguale a quella del treeview
            self.header_canvas.configure(width=treeview_width)
            
            # Calcola l'offset di scrolling dal treeview
            xview = self.tree.xview()
            scroll_offset = int(xview[0] * total_width)
            
            # Pulisci gli header esistenti
            self.header_canvas.delete("all")
            
            # Disegna gli header per ogni gruppo
            for group_name, group_info in self.group_headers.items():
                # Calcola le coordinate x basate sulle larghezze effettive delle colonne
                x_start = sum(column_widths[:group_info['start_col']-1])
                x_end = sum(column_widths[:group_info['end_col']])
                
                # Applica l'offset di scrolling
                x_start -= scroll_offset
                x_end -= scroll_offset
                
                # Disegna l'header solo se è visibile
                if x_end > 0 and x_start < treeview_width:
                    # Crea il rettangolo dell'header
                    self.header_canvas.create_rectangle(
                        x_start, 0, x_end, 30,
                        fill=group_info['bg_color'],
                        outline="#999999",
                        tags=("header", f"header_{group_name}")
                    )
                    
                    # Aggiungi il testo
                    self.header_canvas.create_text(
                        (x_start + x_end) / 2, 15,
                        text=group_name,
                        fill=group_info['text_color'],
                        font=("Arial", 10, "bold"),
                        tags=("header", f"text_{group_name}")
                    )
            
            # Aggiorna il canvas
            self.header_canvas.update_idletasks()
            
        except Exception as e:
            print(f"Errore nel disegno degli header: {str(e)}")
            import traceback
            traceback.print_exc()









