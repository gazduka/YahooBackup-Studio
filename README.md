# YahooBackup Studio

**Backup Yahoo Mail, conversione EML → MSG e importazione in Outlook classico — gratuitamente, una cartella alla volta.**

*Free Windows application to archive Yahoo Mail as EML, convert messages to MSG, and import them into a local Outlook PST. See the [English quick guide](#english-quick-guide) below.*

> **Progetto indipendente:** non affiliato, sponsorizzato o approvato da Yahoo o Microsoft. Non è un servizio cloud: i messaggi vengono elaborati sul computer dell'utente.

## Perché nasce

Chi conserva anni di email su Yahoo può incontrare difficoltà nel sincronizzarle con Outlook: archivi molto grandi, connessioni IMAP che si interrompono, messaggi con allegati pesanti e sincronizzazioni che ricominciano dopo un errore.

Durante un backup reale abbiamo osservato risposte come:

```text
[UNAVAILABLE] UID FETCH Server error - Please try again later
command: FETCH => socket error: EOF
command: FETCH => IMAP4rev1 Server logging out
```

Gli errori si sono ripresentati **periodicamente, spesso dopo alcune decine o un centinaio di messaggi**, senza però dimostrare una soglia fissa. Possono dipendere da timeout, gestione delle sessioni, limiti temporanei o altre condizioni del server. **Non è dimostrato che la causa sia la password per app**, né che tutte le caselle Yahoo seguano lo stesso schema.

**YahooBackup Studio non rimuove le limitazioni di Yahoo e non impedisce le disconnessioni.** Le gestisce: in caso di errore temporaneo tenta di ristabilire la connessione e di ritentare il singolo messaggio, fino a **quattro tentativi**. Ogni email completata resta salvata su disco. Se l'app viene chiusa o Windows interrompe il processo, è possibile avviare nuovamente la stessa cartella senza riscaricare i file EML già presenti e non vuoti.

### Risultati delle prove

- **2.715 email** scaricate dalla cartella Yahoo `inviate 2019 - 2021`: **0 errori definitivi** nel log finale, nonostante numerose disconnessioni temporanee.
- Gli stessi **2.715 EML sono stati convertiti in MSG con 0 errori** durante la prova.
- È stata verificata l'importazione in Outlook di messaggi come veri elementi **`IPM.Note`**, con oggetto, date e allegati; una prova da **10 messaggi** ha restituito **10/10 `IPM.Note`**.
- Questi risultati si riferiscono alle prove effettuate, **non costituiscono garanzia** di compatibilità con tutti i messaggi o con tutte le versioni di Outlook.

## Come funziona

```text
Yahoo Mail via IMAP/SSL
         |
         v
[1] Download di UNA cartella -> file .eml originali
         |
         v
[2] Conversione locale      -> file .msg compatibili con Outlook
         |
         v
[3] Importazione tramite Outlook COM -> cartella in un PST LOCALE già esistente
```

Le tre fasi sono **indipendenti**: si può fermare il lavoro dopo il backup EML e riprendere la conversione o l'importazione in un altro momento. I file EML originali non vengono cancellati dall'app.

### 1 — Yahoo → EML

1. Inserire l'indirizzo Yahoo e una **password per app** generata nelle impostazioni dell'account Yahoo.
2. Premere **Leggi cartelle Yahoo**.
3. Scegliere **una sola cartella** dall'elenco.
4. Premere **Scarica EML della cartella selezionata**.
5. Seguire il contatore e la barra di avanzamento; il registro mostra anche errori temporanei e ritentativi.

Il download usa `imap.mail.yahoo.com` su porta **993 con SSL**, seleziona la cartella in modalità di sola lettura e scarica i messaggi **uno alla volta**, senza scaricare automaticamente l'intera casella o la Posta in arrivo.

Per ogni messaggio, la dimensione ricevuta viene confrontata con quella dichiarata da IMAP quando disponibile. Il file viene prima scritto con estensione `.partial` e rinominato `.eml` solo dopo il salvataggio.

**Ripresa:** al successivo avvio, l'app interroga nuovamente gli UID della cartella e salta i corrispondenti `.eml` già presenti e non vuoti. Non si limita a partire dall'ultimo numero: recupera anche eventuali messaggi mancanti fra quelli già scaricati. I file esistenti non vengono ricontrollati integralmente contro il server a ogni esecuzione.

### 2 — EML → MSG

1. Aprire la scheda **EML → MSG**.
2. Selezionare la cartella numerica `UIDVALIDITY` che contiene direttamente gli `.eml`.
3. Scegliere una **cartella di destinazione MSG separata**.
4. Premere **Converti EML → MSG**.

La conversione è locale, utilizza **Aspose.Email FOSS** e non richiede una nuova connessione a Yahoo. I `.msg` già presenti e non vuoti vengono saltati alla ripresa. L'app registra gli eventuali errori di conversione.

### 3 — MSG → Outlook classico

1. Aprire **Outlook classico** e montare il proprio **file PST locale**, se non è già aperto.
2. Nell'app, aprire **MSG → Outlook** e indicare la cartella dei `.msg`.
3. Selezionare **esattamente il PST locale** di destinazione e specificare il nome della cartella Outlook. L'app può crearla nella radice del PST se non esiste.
4. Per la prima prova lasciare il limite a **10** messaggi; verificare in Outlook oggetti, mittenti, date, corpi e allegati.
5. Quando la prova è soddisfacente, impostare il limite a **0** per importare i rimanenti.

L'importazione usa l'interfaccia **COM/MAPI di Outlook**, non modifica direttamente i blocchi interni del PST e non deve trasferire i messaggi all'account Yahoo IMAP. Ciascun messaggio viene aperto dal `.msg` e spostato nell'archivio locale come elemento Outlook.

**Registro di ripresa:** l'app tiene un elenco dei file già importati, distinto per PST, cartella di destinazione e cartella sorgente. Attenzione: se il programma si interrompe **dopo lo spostamento del messaggio ma prima dell'aggiornamento del registro**, un nuovo tentativo può creare un duplicato. Non avviare due importazioni contemporaneamente sulla stessa destinazione.

## Requisiti

- Windows 10/11.
- **Python 3.13** consigliato per gli script di avvio inclusi; l'app usa Tkinter, normalmente distribuito con Python per Windows.
- Pacchetti Python `IMAPClient` e `aspose-email-foss`.
- Connessione Internet per la fase Yahoo → EML.
- **Outlook classico desktop** con un profilo configurato e un PST locale già disponibile, **solo per la fase 3**. Il nuovo Outlook non offre la stessa automazione COM.
- Spazio su disco sufficiente per conservare EML, MSG e copia locale importata in Outlook.

### Installazione

Scaricare i file del progetto, **estrarre lo ZIP completamente** e mantenere insieme:

```text
YahooBackupStudio/
  YahooBackupStudio.py
  importa_outlook.ps1
  Avvia_YahooBackupStudio.cmd
  Installa_dipendenze.cmd
  LEGGIMI.txt
```

1. Installare Python per Windows, se non già presente.
2. Se il programma segnala librerie mancanti, eseguire `Installa_dipendenze.cmd`, oppure:

   ```powershell
   py -3.13 -m pip install IMAPClient aspose-email-foss
   ```

3. Avviare `Avvia_YahooBackupStudio.cmd` con doppio clic.

**Nota di portabilità (versione 1.0):** i file `.cmd` inclusi cercano per prima cosa Python 3.13 nel percorso standard dell'utente Windows; quello di installazione presume tale percorso. Se Python è installato altrove, adattare i `.cmd` o avviare direttamente `YahooBackupStudio.py` con il proprio interprete Python. Non è un `.exe` standalone.

**Importante:** se l'app è stata avviata da una finestra console, chiudere la console può terminarla. Utilizzare il pulsante di interruzione dell'app per fermare ordinatamente un'operazione.

## Dove vengono salvati i file

Per impostazione predefinita la cartella base è `%USERPROFILE%\YahooBackup`, modificabile nell'interfaccia quando previsto.

```text
YahooBackup/
  archivio_yahoo/
    inviate 2019 - 2021/
      <UIDVALIDITY>/
        uid-1.eml
        uid-2.eml
        ...
  archivio_msg/
    inviate 2019 - 2021/
      uid-1.msg
      uid-2.msg
      ...
```

`UIDVALIDITY` identifica una specifica generazione della cartella IMAP; è usato per evitare di confondere UID appartenenti a generazioni differenti.

## Sicurezza e limiti

- **Non pubblicare mai** password per app, messaggi EML/MSG personali, PST, file di registro con dati privati o schermate che rivelino credenziali.
- La password per app viene richiesta nell'interfaccia; il programma non intende salvarla nei file di configurazione. Non condividere il PC o i suoi processi con utenti non fidati durante l'esecuzione.
- **Il backup non è sincronizzazione bidirezionale:** non è progettato per cancellare o modificare i messaggi su Yahoo.
- **Il recupero degli errori non garantisce il successo di ogni download.** Controllare sempre il conteggio finale e gli eventuali log di errore.
- File EML già presenti e non vuoti vengono saltati. Per verifiche approfondite d'integrità servono controlli ulteriori.
- Prima di importazioni consistenti, controllare dimensione del PST, spazio libero e presenza di eventuali duplicati.
- Gli allegati sono conservati nei messaggi convertiti quando supportati; verificare con un campione reale, in particolare email multipart, formati insoliti e allegati molto grandi.
- Il progetto è fornito **senza garanzia**, con codice sorgente disponibile per verifica e modifiche.

## English quick guide

**YahooBackup Studio** is a free Windows GUI for backing up **one Yahoo Mail folder at a time** to `.eml`, converting `.eml` to `.msg`, and importing `.msg` into an existing local **classic Outlook PST** through Outlook COM.

### Why it helps with Yahoo IMAP disconnects

Large Yahoo IMAP downloads can fail intermittently with `UID FETCH Server error`, `socket error: EOF`, or `Server logging out`. In one real-world test, disconnects appeared repeatedly after roughly 75–145 messages, with no evidence of a universal fixed threshold. The downloader **does not bypass server restrictions**: it retries an individual message up to four times, reconnects after temporary failures, saves completed EML files individually and skips existing nonempty EML files on subsequent runs. A 2,715-message folder was successfully backed up with **zero final errors** in that test.

### Getting started

1. Extract the entire project ZIP; do not run scripts directly from a compressed archive.
2. Install Python 3.13 and the `IMAPClient` and `aspose-email-foss` packages.
3. Launch `Avvia_YahooBackupStudio.cmd` (or run the Python script directly).
4. In **Yahoo → EML**, enter your Yahoo address and app password, select a folder, and download.
5. In **EML → MSG**, choose the downloaded UIDVALIDITY directory and a separate output directory.
6. In **MSG → Outlook**, select your existing **local PST** and target folder; test with 10 messages before importing the rest.

**Requirements:** Windows 10/11, Python, classic desktop Outlook for phase 3 only, sufficient disk space. Keep original EML files. Never upload private mailbox archives or credentials to a public repository.

## Licenza e contributi

Il progetto è distribuito sotto **licenza MIT** (vedere `LICENSE`). Issue, miglioramenti e segnalazioni di errori sono benvenuti: nei report pubblici **rimuovere dati personali, indirizzi email e contenuti dei messaggi**.

---

*Yahoo and Outlook are trademarks of their respective owners. This project is independent and unofficial.*
