# simplewright
a simple software for writing thats ignore all useless feature of modern text editor
UniversalWriter - Multi-Page Advanced Text Editor

UniversalWriter (conosciuto internamente come simplewright) è un editor di testo multi-pagina avanzato sviluppato in Python utilizzando l'interfaccia grafica Tkinter. L'applicazione unisce la semplicità di un blocco note a potenti funzionalità di desktop publishing, inclusa la gestione nativa dei layout dei fogli (A4, A3, A2), l'importazione avanzata di documenti Microsoft Word (.docx), la lettura di PDF con motore OCR Tesseract integrato per scansionare fogli grafici non editabili, e un sistema integrato di stampa/esportazione.
1 Funzionalità Principali

    Gestione Multi-Pagina Reale: Possibilità di aggiungere, scambiare, navigare ed eliminare singoli fogli all'interno dello stesso documento tramite una comoda barra laterale.

    Layout di Pagina Professionale: Supporto ai formati standard (A4, A3, A2) sia in orientamento verticale che orizzontale, con righelli interattivi per regolare al volo i margini sinistro, destro, superiore e inferiore.

    Importazione Estesa & OCR:

        Lettura fluida di file di testo semplici (.txt) e documenti Word (.docx).

        Estrazione del testo da documenti PDF. Se una pagina contiene un'immagine o una scansione non editabile, interviene automaticamente il motore Tesseract OCR in lingua italiana per recuperare il testo leggibile.

    Formattazione Ricca: Modifica della famiglia del font, della dimensione, stili applicabili al testo selezionato (Grassetto, Corsivo, Sottolineato) e giustificazione (Sinistra, Centro, Destra).

    Tavolozza Colori Dinamica: Selezione avanzata del colore del testo e dell'evidenziatore tramite una griglia popup di campioni, con memorizzazione degli slot usati di recente.

    Zoom Dinamico: Slider integrato per scalare visivamente il foglio sulla scrivania digitale dal 50% al 200%.

    Esportazione & Stampa PDF: Generazione di report grafici professionali accodando i fogli tramite la libreria ReportLab.

    Salvataggio della Configurazione: Rilevamento automatico del sistema operativo e salvataggio delle preferenze utente (come la Dark Mode del desktop) in cartelle dedicate (AppData/LocalLow su Windows, .simplewright su Linux).

2 Installazione su Windows

Il programma viene distribuito per Windows sotto forma di un unico file eseguibile standalone (.exe) compilato tramite PyInstaller.
Download & Avvio rapido

    Scarica il file simplewright.exe dalla sezione Releases di questo repository GitHub.

    Fai doppio clic sul file per avviare l'applicazione. Non è richiesta alcuna installazione o dipendenza software esterna.

3 Installazione su Linux (Debian / Ubuntu e derivate)

Su Linux l'applicazione viene installata nativamente a livello di sistema come pacchetto .deb, includendo l'icona e l'integrazione nel menu delle applicazioni sotto la categoria Ufficio.
Requisiti di sistema preliminari

Il motore OCR richiede i componenti di sistema per la lettura dei caratteri. Assicurati che il tuo sistema abbia tesseract-ocr installato:
Bash

sudo apt update
sudo apt install tesseract-ocr tesseract-ocr-ita

Metodo 1: Installazione Standard del pacchetto .deb

Scarica il file simplewright_1.0-1.deb dalle Releases di GitHub, apri il terminale nella cartella di download e digita:
Bash

sudo apt install ./simplewright_1.0-1.deb

Nota: L'uso di apt install (rispetto a dpkg -i) garantisce che l'ambiente installi in automatico le dipendenze base mancanti (come python3-tk).
Metodo 2: Installazione Rapida "Stile Hacker" (Via Terminale)

Se non vuoi scaricare manualmente il file dal browser, puoi copiare e incollare questo singolo comando nel terminale per fare in modo che si connetta alle API di GitHub, scarichi l'ultima release, la installi e ripulisca la cartella:
Bash

curl -s https://api.github.com/repos/TUO-USERNAME/NOME-REPO/releases/latest | grep "browser_download_url.*deb" | cut -d '"' -f 4 | wget -qi - && sudo apt install ./simplewright_1.0-1.deb && rm simplewright_1.0-1.deb

4 Disinstallazione su Linux

Se desideri rimuovere completamente il programma dal tuo computer Linux, puoi utilizzare i gestori di pacchetti standard di sistema.

Rimuovi l'applicazione preservando i file di configurazione locali:
Bash

sudo apt remove simplewright

Rimuovi completamente l'applicazione, i file di configurazione globali e i collegamenti di sistema (Consigliato):
Bash

sudo apt purge simplewright

Dopo la rimozione, esegui questo comando per aggiornare la cache del menu delle applicazioni ed eliminare eventuali icone residue:
Bash

sudo update-desktop-database

5 Scorciatoie da Tastiera Nazionali (Shortcuts)

L'applicazione supporta le seguenti scorciatoie rapide per velocizzare i flussi di lavoro:

    Ctrl + N: Crea un nuovo documento

    Ctrl + O: Apri un documento esistente (.txt, .docx, .pdf, .odt)

    Ctrl + S: Salva il file corrente

    Ctrl + P: Invia il documento alla stampante o esportalo in PDF

    Ctrl + A: Seleziona tutto il testo presente nella pagina

    Ctrl + Z: Annulla l'ultima modifica (Undo)

    Ctrl + Alt + X: Ripristina l'azione annullata (Redo)

    Ctrl + Q: Chiudi ed esci dall'applicazione
