import mailbox
import os
import sys
from email.header import decode_header, make_header
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from threading import Thread
import tkinter.font as tkfont
import subprocess
import platform

def sanitize_filename(filename):
    """
    Pulisce una stringa per renderla un nome di file valido.
    Rimuove i caratteri illegali e limita la lunghezza.
    """
    if not isinstance(filename, str):
        filename = str(filename)

    # Caratteri non validi (AGGIUNTO ':')
    invalid_chars = '\\/:*?"<>|'
    for char in invalid_chars:
        filename = filename.replace(char, '_')
        
    # Rimuove caratteri di controllo (es. newline)
    filename = "".join(c for c in filename if c.isprintable())
    
    # Rimuovi spazi bianchi all'inizio o alla fine
    filename = filename.strip()
    
    # Limita la lunghezza per evitare errori "File name too long"
    # Teniamoci molto corti per sicurezza
    return filename[:150] # <-- Limite ridotto a 150

def decode_mail_header(header_str):
    """
    Decodifica correttamente l'oggetto o altri header delle email.
    """
    if header_str is None:
        return "Senza Oggetto"
    
    try:
        header_obj = make_header(decode_header(header_str))
        return str(header_obj)
    except Exception as e:
        print(f"  [!] Errore nella decodifica dell'header: {e}")
        if isinstance(header_str, bytes):
            return header_str.decode('latin-1', errors='ignore')
        return str(header_str)

def convert_mbox_to_eml(mbox_path, output_dir, progress_callback=None, status_callback=None):
    """
    Converte un file MBOX in file EML individuali.
    """
    if not os.path.exists(mbox_path):
        if status_callback:
            status_callback(f"ERRORE: Il file MBOX '{mbox_path}' non è stato trovato.", "error")
        return False

    os.makedirs(output_dir, exist_ok=True)
    if status_callback:
        status_callback(f"I file EML verranno salvati in: '{os.path.abspath(output_dir)}'", "info")

    try:
        mbox = mailbox.mbox(mbox_path)
    except Exception as e:
        if status_callback:
            status_callback(f"ERRORE: Impossibile aprire il file '{mbox_path}'. È un file MBOX valido? Dettagli: {e}", "error")
        return False

    count = 0
    total = len(mbox)
    errors = 0
    
    if status_callback:
        status_callback(f"Inizio elaborazione di {total} messaggi...", "info")

    for i, message in enumerate(mbox):
        # Estrai e decodifica Oggetto
        subject = decode_mail_header(message['subject'])

        # Pulisci Oggetto per creare un nome file sicuro e CORTO
        subject_clean = sanitize_filename(subject)
        
        # Crea un nome file unico e corto
        filename = f"{i:05d}_{subject_clean}.eml"
        
        output_path = os.path.join(output_dir, filename)

        # Scrivi il messaggio nel file .eml
        try:
            with open(output_path, 'wb') as f:
                f.write(message.as_bytes())
            count += 1
            
            if progress_callback:
                progress_callback(i + 1, total)

        except OSError as e:
            print(f"ERRORE: Impossibile scrivere il file per il messaggio {i}. Motivo: {e}")
            print(f"  -> Nome file problematico (tentato): {filename}")
            errors += 1
        except Exception as e:
            print(f"ERRORE sconosciuto sul messaggio {i}: {e}")
            errors += 1

    if status_callback:
        status_callback(f"Conversione completata! File creati: {count}, Errori: {errors}", "success")
    
    return True

class ModernMBOXConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("MBOX to EML Converter")
        self.root.geometry("700x550")  # Altezza ridotta poiché i log sono collassati
        self.root.configure(bg='#f0f0f0')
        
        # Imposta l'icona
        self.set_icon()
        
        # Variabili
        self.mbox_file = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.is_processing = False
        self.log_visible = False  # Stato iniziale: log nascosti
        
        self.setup_ui()
        
    def set_icon(self):
        """Imposta l'icona dell'applicazione"""
        try:
            # Prova diversi percorsi per trovare l'icona
            icon_paths = [
                "logo.ico",  # Nella cartella corrente
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.ico"),  # Cartella dello script
                os.path.join(os.path.dirname(sys.executable), "logo.ico"),  # Cartella dell'eseguibile
            ]
            
            # Se l'applicazione è frozen (es. eseguibile PyInstaller)
            if getattr(sys, 'frozen', False):
                base_path = sys._MEIPASS
                icon_paths.insert(0, os.path.join(base_path, "logo.ico"))
            
            icon_found = False
            for icon_path in icon_paths:
                if os.path.exists(icon_path):
                    self.root.iconbitmap(icon_path)
                    icon_found = True
                    break
            
            if not icon_found:
                print("Icona logo.ico non trovata. Verrà utilizzata l'icona predefinita.")
                
        except Exception as e:
            print(f"Impossibile caricare l'icona: {e}. Verrà utilizzata l'icona predefinita.")
        
    def setup_ui(self):
        # Titolo principale
        title_font = tkfont.Font(family="Segoe UI", size=16, weight="bold")
        title_label = tk.Label(self.root, text="MBOX to EML Converter", 
                              font=title_font, bg='#f0f0f0', fg='#2c3e50')
        title_label.pack(pady=20)
        
        # Frame principale
        main_frame = tk.Frame(self.root, bg='#f0f0f0')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=10)
        
        # Sezione file MBOX
        mbox_frame = tk.Frame(main_frame, bg='#f0f0f0')
        mbox_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(mbox_frame, text="File MBOX:", 
                font=('Segoe UI', 10, 'bold'), bg='#f0f0f0').pack(anchor='w')
        
        mbox_input_frame = tk.Frame(mbox_frame, bg='#f0f0f0')
        mbox_input_frame.pack(fill=tk.X, pady=5)
        
        tk.Entry(mbox_input_frame, textvariable=self.mbox_file, 
                font=('Segoe UI', 9), width=50).pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        tk.Button(mbox_input_frame, text="Sfoglia", command=self.browse_mbox,
                 font=('Segoe UI', 9), bg='#3498db', fg='white',
                 relief='flat', padx=15).pack(side=tk.RIGHT, padx=(10, 0))
        
        # Sezione cartella output
        output_frame = tk.Frame(main_frame, bg='#f0f0f0')
        output_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(output_frame, text="Cartella di destinazione:", 
                font=('Segoe UI', 10, 'bold'), bg='#f0f0f0').pack(anchor='w')
        
        output_input_frame = tk.Frame(output_frame, bg='#f0f0f0')
        output_input_frame.pack(fill=tk.X, pady=5)
        
        tk.Entry(output_input_frame, textvariable=self.output_dir, 
                font=('Segoe UI', 9), width=50).pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        tk.Button(output_input_frame, text="Sfoglia", command=self.browse_output,
                 font=('Segoe UI', 9), bg='#3498db', fg='white',
                 relief='flat', padx=15).pack(side=tk.RIGHT, padx=(10, 0))
        
        # Progress bar
        self.progress_frame = tk.Frame(main_frame, bg='#f0f0f0')
        self.progress_frame.pack(fill=tk.X, pady=20)
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, mode='determinate')
        self.progress_bar.pack(fill=tk.X)
        
        self.progress_label = tk.Label(self.progress_frame, text="", 
                                      font=('Segoe UI', 9), bg='#f0f0f0')
        self.progress_label.pack(pady=5)
        
        # Area di log - con pulsante per espandere/contrarre
        log_section_frame = tk.Frame(main_frame, bg='#f0f0f0')
        log_section_frame.pack(fill=tk.X, pady=10)
        
        # Pulsante per mostrare/nascondere i log
        self.log_toggle_btn = tk.Button(log_section_frame, text="Mostra Log", command=self.toggle_log,
                                       font=('Segoe UI', 9), bg='#95a5a6', fg='white',
                                       relief='flat', padx=10)
        self.log_toggle_btn.pack(anchor='w')
        
        # Frame che conterrà il log (inizialmente nascosto)
        self.log_frame = tk.Frame(main_frame, bg='#f0f0f0')
        # Non packato inizialmente, verrà mostrato quando richiesto
        
        self.log_text = tk.Text(self.log_frame, height=8, font=('Consolas', 9),
                               bg='#2c3e50', fg='#ecf0f1', relief='flat')
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Scrollbar per il log
        scrollbar = tk.Scrollbar(self.log_text)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.log_text.yview)
        
        # Pulsanti di controllo
        button_frame = tk.Frame(main_frame, bg='#f0f0f0')
        button_frame.pack(fill=tk.X, pady=20)
        
        self.convert_btn = tk.Button(button_frame, text="Converti", command=self.start_conversion,
                                   font=('Segoe UI', 11, 'bold'), bg='#27ae60', fg='white',
                                   relief='flat', padx=30, pady=10)
        self.convert_btn.pack(side=tk.RIGHT, padx=(10, 0))
        
        tk.Button(button_frame, text="Pulisci Log", command=self.clear_log,
                 font=('Segoe UI', 10), bg='#e74c3c', fg='white',
                 relief='flat', padx=20).pack(side=tk.RIGHT)
        
        tk.Button(button_frame, text="Apri Cartella Output", command=self.open_output_dir,
                 font=('Segoe UI', 10), bg='#9b59b6', fg='white',
                 relief='flat', padx=20).pack(side=tk.LEFT)
        
        # Footer
        footer_frame = tk.Frame(main_frame, bg='#f0f0f0')
        footer_frame.pack(fill=tk.X, pady=10)
        
        footer_label = tk.Label(footer_frame, text="v1.0 Powered by Guido Ballarini", 
                               font=('Segoe UI', 8), bg='#f0f0f0', fg='#7f8c8d')
        footer_label.pack(side=tk.RIGHT)
    
    def toggle_log(self):
        """Mostra o nasconde l'area dei log"""
        if self.log_visible:
            self.log_frame.pack_forget()
            self.log_toggle_btn.config(text="Mostra Log")
            self.log_visible = False
            # Riduci l'altezza della finestra quando i log sono nascosti
            self.root.geometry("700x550")
        else:
            self.log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
            self.log_toggle_btn.config(text="Nascondi Log")
            self.log_visible = True
            # Aumenta l'altezza della finestra quando i log sono visibili
            self.root.geometry("700x650")
    
    def browse_mbox(self):
        file_path = filedialog.askopenfilename(
            title="Seleziona file MBOX",
            filetypes=[("MBOX files", "*.mbox"), ("All files", "*.*")]
        )
        if file_path:
            self.mbox_file.set(file_path)
    
    def browse_output(self):
        dir_path = filedialog.askdirectory(title="Seleziona cartella di destinazione")
        if dir_path:
            self.output_dir.set(dir_path)
    
    def log_message(self, message, message_type="info"):
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        
        # Colori in base al tipo di messaggio
        if message_type == "error":
            self.log_text.tag_add("error", "end-2l", "end-1l")
            self.log_text.tag_config("error", foreground="#e74c3c")
        elif message_type == "success":
            self.log_text.tag_add("success", "end-2l", "end-1l")
            self.log_text.tag_config("success", foreground="#27ae60")
        elif message_type == "warning":
            self.log_text.tag_add("warning", "end-2l", "end-1l")
            self.log_text.tag_config("warning", foreground="#f39c12")
    
    def clear_log(self):
        self.log_text.delete(1.0, tk.END)
    
    def update_progress(self, current, total):
        percentage = (current / total) * 100
        self.progress_bar['value'] = percentage
        self.progress_label.config(text=f"Progresso: {current}/{total} ({percentage:.1f}%)")
    
    def status_callback(self, message, message_type="info"):
        self.root.after(0, self.log_message, message, message_type)
    
    def start_conversion(self):
        if self.is_processing:
            return
        
        if not self.mbox_file.get():
            messagebox.showerror("Errore", "Seleziona un file MBOX!")
            return
        
        if not self.output_dir.get():
            messagebox.showerror("Errore", "Seleziona una cartella di destinazione!")
            return
        
        self.is_processing = True
        self.convert_btn.config(state=tk.DISABLED, text="Conversione in corso...")
        self.progress_bar['value'] = 0
        self.progress_label.config(text="")
        
        # Esegui la conversione in un thread separato
        thread = Thread(target=self.run_conversion)
        thread.daemon = True
        thread.start()
    
    def run_conversion(self):
        try:
            success = convert_mbox_to_eml(
                self.mbox_file.get(),
                self.output_dir.get(),
                progress_callback=self.update_progress,
                status_callback=self.status_callback
            )
        except Exception as e:
            self.status_callback(f"Errore durante la conversione: {str(e)}", "error")
            success = False
        
        # Ripristina l'interfaccia
        self.root.after(0, self.conversion_finished, success)
    
    def conversion_finished(self, success):
        self.is_processing = False
        self.convert_btn.config(state=tk.NORMAL, text="Converti")
        
        if success:
            messagebox.showinfo("Successo", "Conversione completata con successo!")
        else:
            messagebox.showerror("Errore", "Si è verificato un errore durante la conversione.")
    
    def open_output_dir(self):
        output_dir = self.output_dir.get()
        if not output_dir:
            messagebox.showwarning("Attenzione", "Nessuna cartella di output specificata!")
            return
            
        if os.path.exists(output_dir):
            try:
                # Metodo cross-platform per aprire la cartella
                if platform.system() == "Windows":
                    os.startfile(output_dir)
                elif platform.system() == "Darwin":  # macOS
                    subprocess.Popen(["open", output_dir])
                else:  # Linux e altri
                    subprocess.Popen(["xdg-open", output_dir])
            except Exception as e:
                messagebox.showerror("Errore", f"Impossibile aprire la cartella: {str(e)}")
        else:
            messagebox.showwarning("Attenzione", "La cartella di output non esiste!")

if __name__ == "__main__":
    root = tk.Tk()
    app = ModernMBOXConverter(root)
    root.mainloop()
