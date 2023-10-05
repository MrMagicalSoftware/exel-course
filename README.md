# exel-course



EXCEL intermedio/avanzato
Programma:
Funzione logiche 
Funzioni di data
 Gestione dei file e stampe
• Importazione e esportazione di file in/da altri formati
• Le funzioni di testo (stringa.estrai, sinistra, trova, concatena)
• Le funzioni di ricerca
• Ordinamento semplice e personalizzato
• Inserimento di grafici
• Operazioni con i Nomi di Zona
• Progettazione e costruzione di un database in Excel
• Applicazione dei criteri di convalida
• Funzioni avanzate logiche e di database
• Funzioni avanzate di ricerca informazioni 
• Ordinamenti semplici e a chiave multipla
• Selezione mediante i filtri (semplici ed avanzati)
• Uso dei Subtotali
• Analisi dati con le Tabelle Pivot
• Grafici di tabelle Pivot
• Power Pivot: analisi di business intelligence
• Importare dati esterni con Power Query
• Funzionalità Scenari per confrontare ed analizzare i dati
• Strumento Risolutore per risolvere problemi complessi
• Consolidamento dei dati (Consolida)
• Proteggere fogli e cartelle 
• Nascondere le formule
• Registrare macro per automatizzare operazioni ripetitive
• Cenni al linguaggio VBA per modificare una macro registrata in precedenza



## Funzione logiche 



Excel offre una varietà di funzioni logiche che puoi utilizzare per eseguire operazioni basate su condizioni logiche. Di seguito sono elencate alcune delle funzioni logiche più comuni in Excel:

### 1. **IF (SE)**
La funzione IF (SE) è una delle funzioni logiche più utilizzate in Excel. Si usa per eseguire un'operazione se una condizione è vera e un'altra operazione se la condizione è falsa.

Esempio:
```excel
=IF(A1>10, "Maggiore di 10", "Minore o uguale a 10")
```

### 2. **AND (E)**
La funzione AND (E) restituisce TRUE se tutte le condizioni specificate sono vere e FALSE se anche una sola condizione è falsa.

Esempio:
```excel
=AND(A1>10, B1<20)
```

### 3. **OR (O)**
La funzione OR (O) restituisce TRUE se almeno una delle condizioni specificate è vera e FALSE se tutte le condizioni sono false.

Esempio:
```excel
=OR(A1>10, B1>10)
```

### 4. **NOT (NON)**
La funzione NOT (NON) restituisce TRUE se la condizione specificata è falsa e FALSE se la condizione è vera.

Esempio:
```excel
=NOT(A1>10)
```

### 5. **IFERROR (SEERRORE)**
La funzione IFERROR (SEERRORE) restituisce un valore specificato se la formula contiene un errore e il risultato della formula se non c'è alcun errore.

Esempio:
```excel
=IFERROR(A1/B1, "Errore: divisione per zero")
```

### 6. **XOR (XOR)**
La funzione XOR (XOR) restituisce TRUE se un numero dispari di condizioni specificate è vera e FALSE se un numero pari di condizioni è vera.

Esempio:
```excel
=XOR(A1>10, B1<20, C1=0)
```

Queste sono solo alcune delle funzioni logiche disponibili in Excel. Puoi combinare queste funzioni e utilizzarle insieme per creare formule complesse basate su condizioni logiche.











