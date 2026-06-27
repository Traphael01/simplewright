# SimpleWright

**Un editor di testo minimalista e potente. Niente fronzoli, solo quello che serve.**

---

## Cos'è SimpleWright?

SimpleWright è un editor di testo multi-pagina sviluppato in Python con una filosofia semplice: fornisce le funzionalità essenziali di un word processor professionale senza distrazioni inutili.

Perfetto per chi vuole:
- Scrivere documenti senza compromessi
- Gestire più pagine in un unico file
- Lavorare con PDF e documenti Word
- Esportare in PDF con qualità professionale

---

## Funzionalità

### Gestione Multi-Pagina
- Aggiungi, riordina, elimina pagine con facilità
- Navigazione tramite barra laterale intuitiva
- Ogni pagina mantenuta indipendentemente nel documento

### Layout Professionale
- Formati standard supportati: **A4, A3, A2**
- Orientamento verticale e orizzontale
- Righelli interattivi per margini (sinistro, destro, superiore, inferiore)

### Importazione Intelligente
- **File di testo** (.txt) - importazione diretta
- **Documenti Word** (.docx) - lettura fluida con preservazione del testo
- **PDF** - estrazione testo + OCR integrato
  - Se il PDF contiene scansioni, **Tesseract OCR** (italiano) le legge automaticamente
  - Nessun tool esterno da configurare

### Formattazione
- **Font**: cambio famiglia e dimensione
- **Stili**: Grassetto, Corsivo, Sottolineato
- **Allineamento**: Sinistra, Centro, Destra
- **Colori**: tavolozza dinamica con memoria degli ultimi colori usati

### Zoom Dinamico
- Slider dal **50% al 200%** per adattare la visualizzazione alle tue preferenze

### Esportazione & Stampa
- Stampa diretta su stampante
- Esportazione PDF con qualità grafica professionale (ReportLab)

### Configurazione Cross-Platform
- Salvataggio automatico delle preferenze
- **Windows**: `AppData/LocalLow/simplewright/`
- **Linux**: `~/.simplewright/`

---

## Installazione

### Windows

1. Scarica `simplewright.exe` dalle [Releases](https://github.com/Traphael01/simplewright/releases)
2. Fai doppio clic per avviare
3. **Niente dipendenze, niente configurazione**

### Linux (Debian/Ubuntu e derivate)

#### Prerequisiti: Installa OCR
```bash
sudo apt update
sudo apt install tesseract-ocr tesseract-ocr-ita
```

#### Metodo 1: Pacchetto .deb (Consigliato)
```bash
sudo apt install ./simplewright_1.0-1.deb
```

#### Metodo 2: Installazione Automatica (One-liner)
```bash
curl -s https://api.github.com/repos/TUO-USERNAME/REPO/releases/latest | grep "browser_download_url.*deb" | cut -d '"' -f 4 | wget -qi - && sudo apt install ./simplewright_1.0-1.deb && rm simplewright_1.0-1.deb
```
#BIG DISCLAMER
---
 INSTALL Tesseract OCR FROM https://tesseractocr.org/#install
---

## Disinstallazione (Linux)

### Rimuovi mantenendo configurazione
```bash
sudo apt remove simplewright
```

### Rimuovi completamente tutto
```bash
sudo apt purge simplewright
sudo update-desktop-database
```

---

## Scorciatoie Tastiera

| Azione | Shortcut |
|--------|----------|
| Nuovo documento | `Ctrl + N` |
| Apri documento | `Ctrl + O` |
| Salva | `Ctrl + S` |
| Stampa / Esporta PDF | `Ctrl + P` |
| Seleziona tutto | `Ctrl + A` |
| Annulla (Undo) | `Ctrl + Z` |
| Ripeti (Redo) | `Ctrl + Alt + X` |
| Esci | `Ctrl + Q` |

---

## Formati Supportati

**Lettura:**
- `.txt` - Testo semplice
- `.docx` - Microsoft Word
- `.pdf` - Portable Document Format (con OCR per scansioni)
- `.odt` - OpenDocument Text

**Esportazione:**
- `.txt` - Testo semplice
- `.pdf` - Alta qualità grafica

---

## Tech Stack

- **Linguaggio**: Python 3
- **GUI**: Tkinter (built-in, zero dipendenze per Windows)
- **OCR**: Tesseract
- **PDF**: ReportLab (esportazione), PyPDF2/pdfplumber (lettura)

---

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

[Specifica la tua licenza qui - MIT, GPL, etc.]

---

## Feedback & Bug Report

Trovato un bug? Vuoi suggerire una feature?
→ Apri un [Issue su GitHub](https://github.com/Traphael01/simplewright/issues)

---

**SimpleWright**: *Scrivi senza distrazioni.*
