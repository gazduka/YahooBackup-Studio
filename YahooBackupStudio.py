#!/usr/bin/env python3
"""YahooBackup Studio: GUI for folder-scoped Yahoo EML backup, EML->MSG and Outlook import."""
import csv
import hashlib
import json
import os
import queue
import re
import subprocess
import sys
import threading
import time
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

APP_DIR = Path(__file__).resolve().parent
BASE_DIR = Path.home() / 'YahooBackup'
IMAP_HOST = 'imap.mail.yahoo.com'
LOG_MAX_LINES = 2000


def safe_folder(text):
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', text).rstrip('. ')
    return s or '_folder'


def safe_file(text):
    return safe_folder(text)


def pst_key(pst_path, folder_name, src_dir):
    v = '|'.join((str(Path(pst_path).resolve()).lower(), folder_name.lower(), str(Path(src_dir).resolve()).lower()))
    return hashlib.sha256(v.encode('utf-8')).hexdigest()[:16]


class Studio(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('YahooBackup Studio  |  Yahoo → EML → MSG → Outlook')
        self.geometry('1000x760')
        self.minsize(880, 620)
        self.configure(bg='#f5f7fb')
        self.events = queue.Queue()
        self.stop_requested = threading.Event()
        self.worker = None
        self.process = None
        self.current_stopfile = None
        self.folders = []
        self.email = tk.StringVar()
        self.password = tk.StringVar()
        self.folder_var = tk.StringVar()
        self.base = tk.StringVar(value=str(BASE_DIR))
        self.eml_var = tk.StringVar()
        self.msg_var = tk.StringVar()
        self.pst_var = tk.StringVar()
        self.dest_name = tk.StringVar(value='inviate 2019 - 2021')
        self.use_old_log = tk.BooleanVar(value=False)
        self.max_import = tk.StringVar(value="10")
        self.status = tk.StringVar(value='Pronto. Seleziona una funzione.')
        self.progress = tk.DoubleVar(value=0)
        self.setup_ui()
        self.after(80, self.poll_events)
        self.protocol('WM_DELETE_WINDOW', self.on_close)

    def setup_ui(self):
        st = ttk.Style(self)
        try: st.theme_use('clam')
        except tk.TclError: pass
        st.configure('TNotebook.Tab', padding=(18, 10), font=('Segoe UI', 10))
        st.configure('TButton', padding=(11, 7), font=('Segoe UI', 10))
        st.configure('TLabel', font=('Segoe UI', 10))
        st.configure('TEntry', padding=5)
        st.configure('TProgressbar', thickness=13)

        top = tk.Frame(self, bg='#15345c', padx=20, pady=17)
        top.pack(fill='x')
        tk.Label(top, text='YahooBackup Studio', fg='white', bg='#15345c', font=('Segoe UI', 22, 'bold')).pack(anchor='w')
        tk.Label(top, text='Tre fasi indipendenti • una cartella alla volta • originali EML sempre conservati',
                 fg='#d4e5fb', bg='#15345c', font=('Segoe UI', 10)).pack(anchor='w')

        rootframe = tk.Frame(self, bg='#f5f7fb', padx=16, pady=12)
        rootframe.pack(fill='both', expand=True)
        self.notebook = ttk.Notebook(rootframe)
        self.notebook.pack(fill='both', expand=True)
        t1, t2, t3 = (ttk.Frame(self.notebook, padding=18), ttk.Frame(self.notebook, padding=18), ttk.Frame(self.notebook, padding=18))
        self.notebook.add(t1, text='1   Yahoo → EML')
        self.notebook.add(t2, text='2   EML → MSG')
        self.notebook.add(t3, text='3   MSG → Outlook')

        self.label(t1, 'Account Yahoo (password per app; non viene memorizzata)')
        r = ttk.Frame(t1); r.pack(fill='x', pady=(2, 10))
        ttk.Label(r, text='Email:').pack(side='left')
        ttk.Entry(r, textvariable=self.email, width=33).pack(side='left', padx=(7, 18))
        ttk.Label(r, text='Password per app:').pack(side='left')
        ttk.Entry(r, textvariable=self.password, show='•', width=27).pack(side='left', padx=7)
        ttk.Button(t1, text='1. Leggi cartelle Yahoo', command=self.list_folders).pack(anchor='w', pady=(0, 8))
        r = ttk.Frame(t1); r.pack(fill='x', pady=3)
        ttk.Label(r, text='Cartella Yahoo:').pack(side='left')
        self.folder_combo = ttk.Combobox(r, textvariable=self.folder_var, state='readonly', width=43)
        self.folder_combo.pack(side='left', padx=8)
        ttk.Button(t1, text='2. Scarica EML della cartella selezionata', command=self.download_selected).pack(anchor='w', pady=12)
        self.label(t1, 'I file finiscono sotto YahooBackup\\archivio_yahoo\\<cartella>\\<UIDVALIDITY>. La ripresa salta i file già integri.')

        self.label(t2, 'Cartella contenente gli EML (quella numerica UIDVALIDITY)')
        self.path_selector(t2, self.eml_var, folder=True)
        self.label(t2, 'Destinazione MSG (separata dagli EML)')
        self.path_selector(t2, self.msg_var, folder=True)
        ttk.Button(t2, text='Converti EML → MSG', command=self.convert_eml).pack(anchor='w', pady=12)
        self.label(t2, 'Aspose.Email FOSS; conversione locale, controllo file non vuoti, ripresa e log degli errori.')

        self.label(t3, 'Origine: cartella contenente i MSG')
        self.path_selector(t3, self.msg_var, folder=True)
        self.label(t3, 'PST locale di Outlook (deve essere già aperto in Outlook classico)')
        self.path_selector(t3, self.pst_var, folder=False)
        self.label(t3, 'Nome della cartella di destinazione nella radice del PST')
        ttk.Entry(t3, textvariable=self.dest_name, width=60).pack(anchor='w', fill='x', pady=(0, 8))
        r = ttk.Frame(t3); r.pack(anchor='w', pady=(0, 5))
        ttk.Label(r, text='Importa massimo N email (10 = prova, 0 = tutte):').pack(side='left')
        ttk.Entry(r, textvariable=self.max_import, width=8).pack(side='left', padx=8)
        ttk.Checkbutton(t3, text='Considera anche il vecchio importati_outlook.log (solo se sono proprio le stesse email già importate nella stessa cartella)',
                        variable=self.use_old_log).pack(anchor='w', pady=3)
        ttk.Button(t3, text='Importa in Outlook (riprende automaticamente)', command=self.import_outlook).pack(anchor='w', pady=12)
        self.label(t3, 'Controllo PST obbligatorio; nessun accesso alla cartella IMAP Yahoo. Importa un vero messaggio IPM.Note per volta.')

        bottom = ttk.Frame(rootframe)
        bottom.pack(fill='both', expand=False, pady=(12, 0))
        r = ttk.Frame(bottom); r.pack(fill='x')
        ttk.Label(r, textvariable=self.status).pack(side='left', fill='x', expand=True)
        self.stop_btn = ttk.Button(r, text='Interrompi dopo il messaggio corrente', command=self.request_stop, state='disabled')
        self.stop_btn.pack(side='right')
        ttk.Progressbar(bottom, variable=self.progress, maximum=100).pack(fill='x', pady=(9, 8))
        logbox = ttk.Frame(bottom)
        logbox.pack(fill='both', expand=True)
        self.log = tk.Text(logbox, height=12, state='disabled', wrap='word', font=('Consolas', 9), bg='#102030', fg='#e1edf8')
        scrollbar = ttk.Scrollbar(logbox, orient='vertical', command=self.log.yview)
        self.log.configure(yscrollcommand=scrollbar.set)
        self.log.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        self.append_log('Pronto. Tutte le operazioni sono locali, salvo la fase Yahoo IMAP.')

    def label(self, parent, txt):
        ttk.Label(parent, text=txt, wraplength=890).pack(anchor='w', pady=(8, 4))

    def path_selector(self, parent, var, folder):
        r = ttk.Frame(parent); r.pack(fill='x', pady=(0, 8))
        ttk.Entry(r, textvariable=var).pack(side='left', fill='x', expand=True)
        def pick():
            selected = filedialog.askdirectory(initialdir=str(BASE_DIR)) if folder else filedialog.askopenfilename(
                title='Seleziona il PST locale', filetypes=[('Outlook PST', '*.pst'), ('Tutti i file', '*.*')])
            if selected: var.set(selected)
        ttk.Button(r, text='Sfoglia…', command=pick).pack(side='left', padx=(8, 0))

    def append_log(self, message):
        self.log.configure(state='normal')
        self.log.insert('end', str(message).rstrip() + '\n')
        if int(self.log.index('end-1c').split('.')[0]) > LOG_MAX_LINES:
            self.log.delete('1.0', '500.0')
        self.log.see('end')
        self.log.configure(state='disabled')

    def post(self, name, *args): self.events.put((name, args))

    def poll_events(self):
        try:
            while True:
                kind, args = self.events.get_nowait()
                if kind == 'log': self.append_log(*args)
                elif kind == 'progress': self.progress.set(args[0]); self.status.set(args[1])
                elif kind == 'folders':
                    self.folders = args[0]
                    self.folder_combo['values'] = self.folders
                    if self.folders: self.folder_combo.current(0)
                elif kind == 'download_path':
                    self.eml_var.set(args[0]); self.msg_var.set(args[1]); self.dest_name.set(args[2])
                elif kind == 'done':
                    self.worker = None
                    self.stop_btn.configure(state='disabled')
                    if args[0]:
                        self.status.set('Operazione terminata.')
                    else:
                        self.status.set('Operazione interrotta o con errore. Consulta il log.')
                    if args[1]: messagebox.showinfo('YahooBackup Studio', args[1])
                elif kind == 'error': messagebox.showerror('Errore', args[0])
        except queue.Empty: pass
        self.after(80, self.poll_events)

    def spawn(self, task):
        if self.worker and self.worker.is_alive():
            messagebox.showwarning('Operazione attiva', 'Una fase è già in esecuzione. Non avviarne un’altra in parallelo.')
            return
        self.stop_requested.clear()
        self.progress.set(0)
        self.stop_btn.configure(state='normal')
        def run():
            try: task()
            except Exception as exc:
                self.post('log', 'ERRORE: ' + str(exc))
                self.post('error', str(exc))
                self.post('done', False, '')
        self.worker = threading.Thread(target=run, daemon=True)
        self.worker.start()

    def request_stop(self):
        self.stop_requested.set()
        if self.current_stopfile:
            try: Path(self.current_stopfile).write_text('stop\n', encoding='utf-8')
            except OSError: pass
        self.post('log', 'Interruzione richiesta: verrà completato il messaggio corrente.')

    def list_folders(self):
        email, password = self.email.get().strip(), self.password.get()
        if not email or not password:
            messagebox.showerror('Dati mancanti', 'Inserisci email Yahoo e password per app.'); return
        def task():
            from imapclient import IMAPClient
            with IMAPClient(IMAP_HOST, port=993, ssl=True, timeout=60) as server:
                server.login(email, password)
                items = sorted((name for flags, delimiter, name in server.list_folders()), key=str.casefold)
                self.post('folders', items)
                self.post('log', f'Cartelle Yahoo trovate: {len(items)}')
                server.logout()
            self.post('done', True, 'Elenco cartelle letto. Seleziona la cartella e premi Scarica EML.')
        self.spawn(task)

    def download_selected(self):
        folder = self.folder_var.get()
        email, password = self.email.get().strip(), self.password.get()
        if not folder or not email or not password:
            messagebox.showerror('Dati mancanti', 'Leggi le cartelle e selezionane una; inserisci le credenziali Yahoo.'); return
        base = Path(self.base.get())
        def task(): self.download_worker(email, password, folder, base)
        self.spawn(task)

    def download_worker(self, email, password, folder, base):
        from imapclient import IMAPClient
        def connect():
            srv = IMAPClient(IMAP_HOST, port=993, ssl=True, timeout=120)
            srv.login(email, password)
            srv.select_folder(folder, readonly=True)
            return srv
        server = connect()
        try:
            uidvalidity = server.folder_status(folder, ['UIDVALIDITY'])[b'UIDVALIDITY']
            dest = base / 'archivio_yahoo' / safe_folder(folder) / str(uidvalidity)
            dest.mkdir(parents=True, exist_ok=True)
            msgdest = base / 'archivio_msg' / safe_folder(folder)
            self.post('download_path', str(dest), str(msgdest), folder)
            uids = server.search(['ALL'])
            total = len(uids)
            self.post('log', f'Yahoo: {folder} | messaggi: {total} | UIDVALIDITY: {uidvalidity}')
            self.post('log', f'Destinazione: {dest}')
            saved = skipped = errors = 0
            fail_log = dest / 'download_errors.log'
            for i, uid in enumerate(uids, 1):
                if self.stop_requested.is_set(): break
                output = dest / f'uid-{uid}.eml'
                if output.is_file() and output.stat().st_size > 0:
                    # Existing EML: it was previously saved atomically; keep it.
                    skipped += 1
                else:
                    success = False
                    for attempt in range(1, 5):
                        if self.stop_requested.is_set(): break
                        try:
                            info = server.fetch([uid], ['BODY.PEEK[]', 'RFC822.SIZE'])
                            msg = info.get(uid)
                            if not msg:
                                raise ValueError('UID non restituito dal server')
                            data = msg.get(b'BODY[]') or msg.get(b'BODY.PEEK[]')
                            expected = msg.get(b'RFC822.SIZE')
                            if not isinstance(data, bytes) or not data:
                                raise ValueError('Messaggio mancante/vuoto')
                            if expected is not None and len(data) != expected:
                                raise ValueError(f'Dimensione IMAP differente: {len(data)} != {expected}')
                            tmp = dest / f'uid-{uid}.eml.partial'
                            with tmp.open('wb') as out:
                                out.write(data)
                                out.flush()
                                os.fsync(out.fileno())
                            os.replace(tmp, output)
                            saved += 1
                            success = True
                            break
                        except Exception as exc:
                            self.post('log', f'UID {uid}, tentativo {attempt}/4: {exc}')
                            try: server.logout()
                            except Exception: pass
                            time.sleep(min(3 * attempt, 12))
                            try: server = connect()
                            except Exception as err: self.post('log', f'Riconnessione: {err}')
                    if not success and not self.stop_requested.is_set():
                        errors += 1
                        with fail_log.open('a', encoding='utf-8') as f: f.write(f'{uid}\n')
                if i % 25 == 0 or i == total:
                    s = f'[{i}/{total}] nuovi={saved} presenti={skipped} errori={errors}'
                    self.post('log', s)
                    self.post('progress', 100 * i / max(total, 1), s)
            self.post('log', f'FINE download: nuovi={saved}, presenti={skipped}, errori={errors}')
            self.post('done', errors == 0 and not self.stop_requested.is_set(), 'Download terminato. Passa alla scheda 2.' if not errors and not self.stop_requested.is_set() else '')
        finally:
            try: server.logout()
            except Exception: pass

    def convert_eml(self):
        src, dst = Path(self.eml_var.get()), Path(self.msg_var.get())
        if not src.is_dir():
            messagebox.showerror('Origine assente', 'Seleziona una cartella reale contenente i file .eml.'); return
        if not list(src.glob('*.eml')):
            messagebox.showerror('EML assenti', 'La cartella deve contenere direttamente i file .eml (apri la cartella numerica UIDVALIDITY).'); return
        if src.resolve() == dst.resolve():
            messagebox.showerror('Percorsi uguali', 'Origine e destinazione devono essere distinte.'); return
        def task(): self.convert_worker(src, dst)
        self.spawn(task)

    def convert_worker(self, src, dst):
        from email import policy
        from email.parser import BytesParser
        from aspose.email_foss import msg
        dst.mkdir(parents=True, exist_ok=True)
        files = sorted(src.glob('*.eml'))
        done = skipped = errors = 0
        fail_log = dst / 'conversion_errors.log'
        self.post('log', f'Conversione EML → MSG: {len(files)} file. Destinazione {dst}')
        for i, eml in enumerate(files, 1):
            if self.stop_requested.is_set(): break
            output = dst / (eml.stem + '.msg')
            temp = dst / (eml.stem + '.msg.tmp')
            if output.exists() and output.stat().st_size > 0:
                skipped += 1
            else:
                try:
                    with eml.open('rb') as f: parsed = BytesParser(policy=policy.default).parse(f)
                    converted = msg.MapiMessage.from_email_message(parsed)
                    converted.save(str(temp))
                    if not temp.is_file() or temp.stat().st_size <= 0: raise ValueError('MSG vuoto')
                    temp.replace(output)
                    done += 1
                    del converted, parsed
                except Exception as exc:
                    errors += 1
                    try: temp.unlink(missing_ok=True)
                    except OSError: pass
                    with fail_log.open('a', encoding='utf-8') as f: f.write(f'{eml.name}: {exc}\n')
                    self.post('log', f'ERRORE {eml.name}: {exc}')
            if i % 25 == 0 or i == len(files):
                text = f'[{i}/{len(files)}] nuovi={done} presenti={skipped} errori={errors}'
                self.post('log', text)
                self.post('progress', 100*i/max(len(files), 1), text)
        self.post('log', f'FINE conversione: nuovi={done} presenti={skipped} errori={errors}')
        self.post('done', errors == 0 and not self.stop_requested.is_set(), 'Conversione completata. Passa alla scheda 3.' if not errors and not self.stop_requested.is_set() else '')

    def import_outlook(self):
        src = Path(self.msg_var.get())
        pst = Path(self.pst_var.get())
        destname = self.dest_name.get().strip()
        if not src.is_dir() or not list(src.glob('*.msg')):
            messagebox.showerror('MSG mancanti', 'Seleziona una cartella contenente file .msg.'); return
        if not pst.is_file() or pst.suffix.lower() != '.pst':
            messagebox.showerror('PST mancante', 'Seleziona un file PST locale esistente, già aperto in Outlook classico.'); return
        if not destname or '/' in destname or '\\' in destname:
            messagebox.showerror('Nome non valido', 'Inserisci un nome di cartella Outlook valido.'); return
        try:
            limit = int(self.max_import.get().strip())
            if limit < 0: raise ValueError()
        except ValueError:
            messagebox.showerror('Valore non valido', 'Inserisci 10 per una prova, oppure 0 per tutte le email.'); return
        if not messagebox.askyesno('Conferma destinazione',
            f'SORGENTE: {src}\n\nPST LOCALE: {pst}\n\nCARTELLA: {destname}\n\n'
            'Importare i MSG come vere email, senza toccare gli EML?'):
            return
        def task(): self.import_worker(src, pst, destname, self.use_old_log.get(), limit)
        self.spawn(task)

    def import_worker(self, src, pst, destname, use_legacy, limit):
        key = pst_key(pst, destname, src)
        ledger = src / f'importati_{key}.log'
        legacy = src / 'importati_outlook.log'
        stopfile = src / f'stop_{key}.flag'
        try: stopfile.unlink(missing_ok=True)
        except OSError: pass
        self.current_stopfile = stopfile
        ps = APP_DIR / 'importa_outlook.ps1'
        if not ps.is_file(): raise FileNotFoundError(f'Script Outlook mancante: {ps}')
        shell = (os.environ.get('SystemRoot', 'C:\\Windows') + '\\System32\\WindowsPowerShell\\v1.0\\powershell.exe')
        cmd = [shell, '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(ps),
               '-Source', str(src), '-PstPath', str(pst), '-FolderName', destname,
               '-Ledger', str(ledger), '-StopFile', str(stopfile), '-MaxItems', str(limit)]
        if use_legacy and legacy.exists(): cmd.extend(['-LegacyLedger', str(legacy)])
        self.post('log', f'Importazione: {src} → {pst} / {destname}')
        process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                   text=True, encoding='utf-8', errors='replace', bufsize=1,
                                   creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
        self.process = process
        try:
            for line in process.stdout:
                line = line.strip()
                if not line: continue
                if line.startswith('PROGRESS|'):
                    parts = line.split('|', 5)
                    if len(parts) >= 5:
                        try: self.post('progress', float(parts[1]), f'Importati={parts[2]} | saltati={parts[3]} | errori={parts[4]}')
                        except ValueError: pass
                else: self.post('log', line)
            rc = process.wait()
            self.post('log', f'Processo Outlook terminato, codice {rc}')
            self.post('done', rc == 0 and not self.stop_requested.is_set(), 'Importazione conclusa. Verifica il numero di email nella cartella Outlook.' if rc == 0 and not self.stop_requested.is_set() else '')
        finally:
            self.process = None
            self.current_stopfile = None
            try: stopfile.unlink(missing_ok=True)
            except OSError: pass

    def on_close(self):
        if self.worker and self.worker.is_alive():
            messagebox.showwarning('Operazione attiva', 'Premi «Interrompi dopo il messaggio corrente» e attendi la fine prima di chiudere.')
            return
        self.destroy()


if __name__ == '__main__':
    Studio().mainloop()
