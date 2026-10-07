Certo. Ecco il contenuto pronto da salvare come **`README.md`**:

:::writing{variant="document" id="58321" title="README.md"}

# Simplewright

**Un editor di testo e visualizzatore di documenti multi-pagina, flessibile e potente. Niente distrazioni, solo quello che serve.**

---

## Cos'è Simplewright?

Simplewright è un editor di testo e visualizzatore di documenti sviluppato in Python con una filosofia chiara: fornire un ambiente di scrittura e lettura completo, supportando un'ampia varietà di formati senza appesantire l'esperienza d'uso.

Perfetto per chi vuole:

- Scrivere documenti mantenendo il controllo completo sul layout e la formattazione
- Lavorare su più pagine o slide in un unico ambiente
- Aprire ed esportare documenti Word, PDF, RTF, presentazioni e e-book
- Modificare o aggiungere caselle di testo ricche e immagini sul foglio

---

## Funzionalità principali

### Gestione Multi-Pagina & Layout

- **Aggiunta, riordino e rinominazione**: Gestisci e rinomina le pagine direttamente dalla barra laterale intuitiva.
- **Paginazione automatica intelligente**: Il testo fluisce automaticamente alla pagina successiva quando supera i margini inferiori.
- **Formati e orientamento**: Supporto standard per **A4, A3, A2** con orientamento sia Verticale che Orizzontale.
- **Margini e righelli interattivi**: Regolazione dinamica dei margini superiore, inferiore, sinistro e destro.
- **Zoom fluido**: Controllo dello zoom dal 50% al 200% con supporto per lo scroll orizzontale e verticale.

### Importazione & Lettura Avanzata

- **Documenti Word moderni e legacy**: Lettura nativa di file `.docx` e supporto avanzato per i vecchi file `.doc` (Word 97-2003).
- **Documenti PDF**: Rendering visuale pagina per pagina con mantenimento del layout originale.
- **Presentazioni PowerPoint (****`.pptx`****)**: Caricamento di slide con sfondi, immagini e testo riposizionabile.
- **E-book (****`.epub`****) e OpenDocument (****`.odt`****,****`.ods`****)**: Estrazione pulita dei capitoli e lettura diretta dei dati.
- **Rich Text (****`.rtf`****)**: Parsing di stili, font, colori e allineamenti direttamente dai tag RTF.
- **Markdown (****`.md`****) e Testo (****`.txt`****)**: Formattazione visiva automatica per Markdown e rilevamento intelligente dell'encoding (UTF-8, CP1252, Latin-1).

### Formattazione & Elementi Grafici

- **Gestione Font**: Scelta della famiglia di caratteri di sistema e dimensione del testo.
- **Stili di testo**: Grassetto, Corsivo, Sottolineato e Allineamento (Sinistra, Centro, Destra).
- **Colori e Evidenziatore**: Tavolozza dei colori recenti e supporto per lo sfondo del testo.
- **Caselle di testo ricche (Rich Text Box)**: Inserisci e posiziona liberamente caselle di testo mobili con formattazione mista nello stesso blocco.
- **Immagini e GIF animate**: Inserimento, posizionamento e ridimensionamento proporzionale delle immagini, con riproduzione continua per le GIF.
- **Link automatici**: Riconoscimento degli URL e apertura diretta nel browser tramite `Ctrl + Clic`.

### Esportazione & Stampa

- **Esportazione PDF**: Generazione di documenti PDF ad alta qualità tramite ReportLab o struttura vettoriale nativa.
- **Esportazione Word & RTF**: Salvataggio in formato `.docx` ed `.rtf` con preservazione degli stili e della formattazione.
- **Esportazione Testo**: Salvataggio in formato `.txt` pulito o `.md`.
- **Stampa di sistema**: Invio diretto del lavoro alla stampante predefinita.

---

## Formati Supportati

| Formato | Lettura | Esportazione |
| --- | --- | --- |
| **PDF** (`.pdf`) | Sì (Visuale / PyMuPDF) | Sì |
| **Word** (`.docx`, `.doc`) | Sì (Nativo / Converter) | Sì (`.docx`) |
| **Rich Text** (`.rtf`) | Sì (Con formattazione) | Sì |
| **OpenDocument** (`.odt`, `.ods`) | Sì | Sì (via conversione) |
| **Presentazioni** (`.pptx`) | Sì | Sì (via PDF) |
| **E-book** (`.epub`) | Sì | Sì (via TXT/PDF) |
| **Markdown** (`.md`) | Sì (Stili visivi) | Sì |
| **Testo semplice** (`.txt`) | Sì | Sì |

---

## Installazione

### Windows

1. Scarica l'installer `Simplewright_Setup.exe` dalla sezione [Releases su GitHub](<https://github.com/Traphael01/simplewright/releases>).
2. Esegui la procedura guidata di installazione ed avvia l'applicazione dal collegamento sul Desktop o nel Menu Start.

### Linux (Debian/Ubuntu e derivate)

#### Metodo 1: Pacchetto `.deb` (Consigliato)

Scarica il file `.deb` dalle Releases ed eseguilo da terminale:

```
sudo apt install ./simplewright_2.0-1.deb
```

#### Metodo 2: Installazione Automatica (One-liner)

```
curl -s https://api.github.com/repos/Traphael01/simplewright/releases/latest | grep "browser_download_url.*deb" | cut -d '"' -f 4 | wget -qi - && sudo apt install ./simplewright*.deb && rm simplewright*.deb
```

---

## Disinstallazione (Linux)

### Rimuovi l'applicazione mantenendo la configurazione

```
sudo apt remove simplewright
```

### Rimuovi completamente l'applicazione e le preferenze

```
sudo apt purge simplewright
sudo update-desktop-database
```

---

## Nota importante sulla Stampa

Se la stampa diretta dall'applicazione non dovesse avviarsi o dovesse riscontrare problemi di comunicazione con la stampante di sistema:

1. Salva prima il documento come file **PDF** (`Ctrl + Shift + S` oppure dal menu _File \> Salva con nome..._).
2. Apri ed esegui la stampa del file PDF generato utilizzando un'applicazione dedicata (es. _Adobe Acrobat_, _Acrobat Reader_, _Microsoft Edge_ o il visualizzatore PDF predefinito del sistema operativo).

---

## Scorciatoie da Tastiera

| Azione | Scorciatoia |
| --- | --- |
| Nuovo documento | `Ctrl + N` |
| Apri documento | `Ctrl + O` |
| Salva | `Ctrl + S` |
| Salva con nome | `Ctrl + Shift + S` |
| Stampa / Esporta | `Ctrl + P` |
| Seleziona tutto | `Ctrl + A` |
| Annulla (Undo) | `Ctrl + Z` |
| Ripeti (Redo) | `Ctrl + Alt + X` |
| Esci | `Ctrl + Q` |

---

## Tech Stack

- **Linguaggio**: Python 3
- **GUI**: Tkinter (built-in)
- **PDF & Vector Engine**: PyMuPDF (`pymupdf`), ReportLab
- **Documenti & Presentazioni**: `python-docx`, `python-pptx`
- **Gestione Immagini**: Pillow (`PIL`)
- **Parsing OLE2**: `olefile`

---

## Licenza

Distribuito sotto licenza **GPL 2.0**.

---

## Feedback & Bug Report

Hai riscontrato un problema o vuoi proporre una nuova funzionalità?

→ Apri una [Issue su GitHub](<https://github.com/Traphael01/simplewright/issues>) :::

## Filosofia

SimpleWright segue la filosofia Unix: **fai bene una cosa**.

Non troverai:
- Animazioni inutili
- Cloud sync forzato
- Pubblicità
- Tracciamento

Troverai:
- Stabilità
- Velocità
- Semplicità
- Controllo totale

---

## Licenza

[ GPL 2.0 ]

---

## Feedback & Bug Report

Trovato un bug? Vuoi suggerire una feature?
→ Apri un [Issue su GitHub](https://github.com/Traphael01/simplewright/issues)

---

**SimpleWright**: *Scrivi senza distrazioni.*
