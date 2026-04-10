# Excel English Bridge 🚀

**Excel English Bridge** è un componente aggiuntivo per Microsoft Excel sviluppato in C# utilizzando la libreria Excel-DNA. Il suo scopo è permettere agli utenti di tutto il mondo di utilizzare le funzioni di Excel con la loro nomenclatura inglese originale (es. `EN_VLOOKUP` invece di `CERCA.VERT` o `BUSCARV`) su **qualsiasi installazione di Excel non in lingua inglese**.

## 🤖 Sviluppato con Gemini
Questo progetto rappresenta un caso studio di sviluppo assistito dall'intelligenza artificiale. L'intero codice sorgente, la gestione del bridge tramite Reflection e la risoluzione dei problemi di compatibilità sono stati realizzati grazie alla collaborazione con **Gemini**.

## ✨ Caratteristiche
- **Prefisso Univoco:** Usa `EN_` davanti a qualsiasi funzione inglese per attivarla, indipendentemente dalla lingua del tuo sistema.
- **Dinamico:** Riconosce automaticamente le funzioni disponibili sul tuo PC in base alla tua versione di Excel.
- **Robusto:** Gestisce correttamente celle vuote, testi e riferimenti circolari come le funzioni native.
- **Supporto Moderno:** Pieno supporto alle matrici dinamiche e alle funzioni di Office 365.

## 🛠️ Installazione
1. Scarica il file `.xll` adatto alla tua versione di Excel (32 o 64 bit) dalla sezione [Releases](https://github.com/FerdinandoBattillocchi/Excel-English-Bridge/releases).
2. Trascina il file all'interno di una sessione di Excel aperta o aggiungilo in modo permanente tramite `File > Opzioni > Componenti aggiuntivi`.
3. Inizia a scrivere formule come `=EN_SUM(A1:A10)`.
   
# Excel English Bridge 🚀

**Excel English Bridge** is a Microsoft Excel add-in developed in C# using the Excel-DNA library. Its goal is to allow users worldwide to use Excel functions with their original English nomenclature (e.g., `EN_VLOOKUP` instead of `CERCA.VERT` or `BUSCARV`) on **any non-English installation of Excel**.

## 🤖 Developed with Gemini
This project represents a case study of AI-assisted development. The entire source code, the bridge management via Reflection, and the resolution of compatibility issues were achieved through collaboration with **Gemini**.

## ✨ Features
- **Universal Prefix:** Use `EN_` before any English function to trigger it, regardless of your system's language.
- **Dynamic:** Automatically recognizes the functions available on your PC based on your specific Excel version.
- **Robust:** Correctly handles empty cells, text, and circular references just like native functions.
- **Modern Support:** Full support for dynamic arrays and Office 365 functions.

## 🛠️ Installation
1. Download the appropriate `.xll` file for your Excel version (32 or 64-bit) from the [Releases](https://github.com/FerdinandoBattillocchi/Excel-English-Bridge/releases) section.
2. Drag and drop the file into an open Excel session, or add it permanently via `File > Options > Add-ins`.
3. Start typing formulas like `=EN_SUM(A1:A10)`.
