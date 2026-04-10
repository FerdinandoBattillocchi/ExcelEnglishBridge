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
1. Scarica il file `.xll` adatto alla tua versione di Excel (32 o 64 bit) dalla sezione [Releases](https://github.com/tuo-username/Excel-English-Bridge/releases).
2. Trascina il file all'interno di una sessione di Excel aperta o aggiungilo in modo permanente tramite `File > Opzioni > Componenti aggiuntivi`.
3. Inizia a scrivere formule come `=EN_SUM(A1:A10)`.
