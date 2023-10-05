# exel-course


<pre>
• EXCEL intermedio/avanzato<br>
• Programma: 
• Funzione logiche 
• Funzioni di data
• Gestione dei file e stampe
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
</pre>


## Funzione logiche 



Excel offre una varietà di funzioni logiche che puoi utilizzare per eseguire operazioni basate su condizioni logiche. Di seguito sono elencate alcune delle funzioni logiche più comuni in Excel:
Puoi combinare queste funzioni e utilizzarle insieme per creare formule complesse basate su condizioni logiche.


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
__________________________________________________________


## ESERCIZI CON IF 



### Esercizio 1:
Supponiamo di avere un foglio di calcolo con i voti degli studenti nella colonna A. Se un voto è maggiore o uguale a 60, vuoi assegnare "Pass" nella colonna B, altrimenti "Non passato".

**Formula:**
```excel
=IF(A1>=60, "Pass", "Non passato")
```

### Esercizio 2:
Hai un elenco di temperature nella colonna A. Vuoi classificare le temperature come "Caldo" se sono superiori a 30 gradi Celsius e come "Freddo" altrimenti.

**Formula:**
```excel
=IF(A1>30, "Caldo", "Freddo")
```

### Esercizio 3:
Hai un elenco di numeri nella colonna A. Vuoi determinare se ciascun numero è positivo, negativo o zero.

**Formula:**
```excel
=IF(A1>0, "Positivo", IF(A1<0, "Negativo", "Zero"))
```

### Esercizio 4:
Supponiamo che tu stia calcolando lo sconto per un elenco di prodotti nella colonna A. Se il prezzo del prodotto è superiore a 100, vuoi applicare uno sconto del 10%, altrimenti uno sconto del 5%.

**Formula:**
```excel
=IF(A1>100, A1*0.9, A1*0.95)
```

### Esercizio 5:
Hai un elenco di età nella colonna A. Vuoi determinare se ciascuna persona è un bambino (età inferiore a 18 anni) o un adulto (età uguale o superiore a 18 anni).

**Formula:**
```excel
=IF(A1<18, "Bambino", "Adulto")
```




### Esercizio 6:
Hai un elenco di punteggi degli studenti nella colonna A. Se il punteggio è maggiore o uguale a 90, assegna "Eccellente", se è tra 70 e 89 assegna "Buono", altrimenti assegna "Da migliorare".

**Formula:**
```excel
=IF(A1>=90, "Eccellente", IF(A1>=70, "Buono", "Da migliorare"))
```

### Esercizio 7:
Hai un elenco di numeri nella colonna A. Se il numero è pari, restituisci "Pari", altrimenti restituisci "Dispari".

**Formula:**
```excel
=IF(MOD(A1,2)=0, "Pari", "Dispari")
```

### Esercizio 8:
Supponiamo che tu abbia un elenco di prodotti nella colonna A e un elenco di quantità nella colonna B. Vuoi calcolare l'importo totale. Se la quantità è superiore a 10, applica uno sconto del 10%, altrimenti non applicare alcuno sconto.

**Formula:**
```excel
=IF(B1>10, A1*B1*0.9, A1*B1)
```

### Esercizio 9:
Hai un elenco di età nella colonna A. Vuoi determinare se ciascuna persona è un bambino (età inferiore a 13 anni), un adolescente (età tra 13 e 19 anni) o un adulto (età uguale o superiore a 20 anni).

**Formula:**
```excel
=IF(A1<13, "Bambino", IF(A1<20, "Adolescente", "Adulto"))
```

### Esercizio 10:
Supponiamo che tu abbia un elenco di voti negli esami di matematica e scienze nella colonna A e B rispettivamente. Vuoi assegnare "Promosso" solo se ha superato entrambi gli esami, altrimenti "Non promosso".

**Formula:**
```excel
=IF(AND(A1>=50, B1>=50), "Promosso", "Non promosso")
```




______________________________________________________________________________________


### esercizi con and


### Esercizio 1:
Hai un elenco di voti degli studenti nella colonna A e i loro punteggi di partecipazione nella colonna B. Vuoi assegnare "Promosso" solo se il voto è maggiore o uguale a 60 e il punteggio di partecipazione è maggiore di 80.

**Formula:**
```excel
=IF(AND(A1>=60, B1>80), "Promosso", "Non promosso")
```

### Esercizio 2:
Hai un elenco di età degli studenti nella colonna A e vuoi determinare se ciascuno di essi è un adolescente (età tra 13 e 19 anni) e se frequenta una scuola media.

**Formula:**
```excel
=IF(AND(A1>=13, A1<=19, B1="Scuola Media"), "Adolescente in Scuola Media", "Altro")
```

### Esercizio 3:
Hai un elenco di temperature nella colonna A e vuoi verificare se ciascuna temperatura è superiore a 0 gradi Celsius e inferiore a 100 gradi Celsius.

**Formula:**
```excel
=IF(AND(A1>0, A1<100), "Valida", "Non valida")
```

### Esercizio 4:
Hai un elenco di orari nella colonna A e vuoi determinare se ciascun orario è tra le 9 del mattino e le 5 del pomeriggio.

**Formula:**
```excel
=IF(AND(A1>=9:00, A1<=17:00), "Orario lavorativo", "Fuori orario")
```

### Esercizio 5:
Hai un elenco di numeri nella colonna A, B e C. Vuoi assegnare "Tutti positivi" solo se tutti e tre i numeri sono positivi.

**Formula:**
```excel
=IF(AND(A1>0, B1>0, C1>0), "Tutti positivi", "Non tutti positivi")
```

### Esercizio 6:
Hai un elenco di valori nella colonna A. Vuoi verificare se ciascun valore è una stringa di testo e contiene la parola "Excel".

**Formula:**
```excel
=IF(AND(ISTEXT(A1), FIND("Excel", A1) > 0), "Contiene Excel", "Non contiene Excel")
```

### Esercizio 7:
Hai un elenco di dati nella colonna A e vuoi verificare se ciascun dato è un numero intero positivo.

**Formula:**
```excel
=IF(AND(ISNUMBER(A1), A1=INT(A1), A1>0), "Numero intero positivo", "Non è un numero intero positivo")
```

### Esercizio 8:
Hai un elenco di orari nella colonna A. Vuoi verificare se ciascun orario è uguale o successivo a un'ora specifica, ad esempio le 14:30.

**Formula:**
```excel
=IF(AND(A1>=14:30), "Uguale o successivo a 14:30", "Antecedente a 14:30")
```

### Esercizio 9:
Hai un elenco di temperature nella colonna A e vuoi verificare se ciascuna temperatura è superiore a una soglia specifica, ad esempio 25 gradi Celsius.

**Formula:**
```excel
=IF(AND(A1>25), "Superiore a 25 gradi", "Non superiore a 25 gradi")
```

### Esercizio 10:
Hai un elenco di numeri nella colonna A, B e C. Vuoi assegnare "Tutti pari" solo se tutti e tre i numeri sono pari.

**Formula:**
```excel
=IF(AND(MOD(A1,2)=0, MOD(B1,2)=0, MOD(C1,2)=0), "Tutti pari", "Non tutti pari")
```

Puoi utilizzare queste formule nelle tue celle Excel per eseguire le verifiche logiche descritte negli esercizi. Assicurati di adattare le formule in base alla disposizione specifica dei dati nel tuo foglio di calcolo.


_____________________________________________________________________



## Esercizi con OR EXEL 


### Esercizio 1:
Hai un elenco di voti nella colonna A. Vuoi assegnare "Approvato" se il voto è maggiore o uguale a 50 o se il punteggio di partecipazione nella colonna B è maggiore di 80.

**Formula:**
```excel
=IF(OR(A1>=50, B1>80), "Approvato", "Non approvato")
```

### Esercizio 2:
Hai un elenco di numeri nella colonna A. Vuoi determinare se ciascun numero è multiplo di 3 o di 5.

**Formula:**
```excel
=IF(OR(MOD(A1,3)=0, MOD(A1,5)=0), "Multiplo di 3 o 5", "Non multiplo di 3 o 5")
```

### Esercizio 3:
Hai un elenco di nomi nella colonna A. Vuoi verificare se ciascun nome è "Alice" o "Bob".

**Formula:**
```excel
=IF(OR(A1="Alice", A1="Bob"), "Nome valido", "Nome non valido")
```

### Esercizio 4:
Hai un elenco di età nella colonna A. Vuoi determinare se ciascuna persona è un bambino (età inferiore a 13 anni) o un anziano (età uguale o superiore a 65 anni).

**Formula:**
```excel
=IF(OR(A1<13, A1>=65), "Bambino o Anziano", "Non Bambino o Anziano")
```

### Esercizio 5:
Hai un elenco di colori nella colonna A. Vuoi verificare se ciascun colore è rosso, verde o blu.

**Formula:**
```excel
=IF(OR(A1="Rosso", A1="Verde", A1="Blu"), "Colore valido", "Colore non valido")
```

### Esercizio 6:
Hai un elenco di temperature nella colonna A. Vuoi determinare se ciascuna temperatura è inferiore a 0 gradi Celsius o superiore a 30 gradi Celsius.

**Formula:**
```excel
=IF(OR(A1<0, A1>30), "Estremo", "Nella norma")
```

### Esercizio 7:
Hai un elenco di numeri nella colonna A. Vuoi verificare se ciascun numero è negativo o superiore a 100.

**Formula:**
```excel
=IF(OR(A1<0, A1>100), "Negativo o Maggiore di 100", "Positivo e Minore o Uguale a 100")
```

### Esercizio 8:
Hai un elenco di date nella colonna A. Vuoi determinare se ciascuna data è di un giorno festivo (ad esempio, Natale o Capodanno).

**Formula:**
```excel
=IF(OR(MONTH(A1)=12, DAY(A1)=1, MONTH(A1)=1, DAY(A1)=1), "Giorno festivo", "Non giorno festivo")
```

### Esercizio 9:
Hai un elenco di numeri nella colonna A. Vuoi verificare se ciascun numero è un quadrato perfetto o un cubo perfetto.

**Formula:**
```excel
=IF(OR(SQRT(A1)=INT(SQRT(A1)), A1^(1/3)=INT(A1^(1/3))), "Quadrato o Cubo perfetto", "Non Quadrato o Cubo perfetto")
```

### Esercizio 10:
Hai un elenco di temperature nella colonna A. Vuoi determinare se ciascuna temperatura è inferiore a 0 gradi Celsius o superiore a 25 gradi Celsius.

**Formula:**
```excel
=IF(OR(A1<0, A1>25), "Estremo", "Nella norma")
```

Questi esercizi ti permetteranno di praticare l'utilizzo della funzione OR in diverse situazioni. Puoi adattare le formule in base alle tue esigenze specifiche e sperimentare con altre condizioni logiche per ampliare le tue competenze in Excel.

_____________________________________________________________


## FUNZIONE NOT IN EXEL  ESERCITAZIONE 


### Esercizio 1:
Hai un elenco di voti nella colonna A. Vuoi assegnare "Promosso" se il voto è superiore a 50, altrimenti "Non promosso".

**Formula:**
```excel
=IF(NOT(A1>50), "Non promosso", "Promosso")
```

### Esercizio 2:
Hai un elenco di età nella colonna A. Vuoi determinare se ciascuna persona non è un bambino (età maggiore o uguale a 13 anni).

**Formula:**
```excel
=IF(NOT(A1<13), "Non bambino", "Bambino")
```

### Esercizio 3:
Hai un elenco di valori nella colonna A. Vuoi verificare se ciascun valore è diverso da zero.

**Formula:**
```excel
=IF(NOT(A1=0), "Diverso da zero", "Zero")
```

### Esercizio 4:
Hai un elenco di stringhe di testo nella colonna A. Vuoi verificare se ciascuna stringa non è vuota.

**Formula:**
```excel
=IF(NOT(A1=""), "Non vuota", "Vuota")
```

### Esercizio 5:
Hai un elenco di numeri nella colonna A. Vuoi verificare se ciascun numero è negativo.

**Formula:**
```excel
=IF(NOT(A1<0), "Non negativo", "Negativo")
```

### Esercizio 6:
Hai un elenco di booleani (VERO o FALSO) nella colonna A. Vuoi ottenere il valore opposto (es. da VERO a FALSO e viceversa).

**Formula:**
```excel
=NOT(A1)
```

### Esercizio 7:
Hai un elenco di valori nella colonna A. Vuoi verificare se ciascun valore è un numero intero.

**Formula:**
```excel
=IF(NOT(ISNUMBER(A1)), "Non è un numero", "È un numero")
```

### Esercizio 8:
Hai un elenco di date nella colonna A. Vuoi verificare se ciascuna data è superiore alla data odierna.

**Formula:**
```excel
=IF(NOT(A1>TODAY()), "Data passata", "Data futura o odierna")
```

### Esercizio 9:
Hai un elenco di valori nella colonna A. Vuoi verificare se ciascun valore è un numero intero positivo.

**Formula:**
```excel
=IF(NOT(INT(A1)=A1, A1>0), "Non è un numero intero positivo", "È un numero intero positivo")
```

### Esercizio 10:
Hai un elenco di stringhe di testo nella colonna A. Vuoi verificare se ciascuna stringa contiene la parola "Excel".

**Formula:**
```excel
=IF(NOT(ISNUMBER(FIND("Excel", A1))), "Non contiene Excel", "Contiene Excel")
```


________________________________________________________________________________

## Esercizi funzione IFERROR (SE.ERRORE) in Excel:

### Esercizio 1:
Hai un elenco di calcoli nella colonna A. Vuoi visualizzare "Errore di calcolo" se c'è un errore, altrimenti il risultato del calcolo.

**Formula:**
```excel
=IFERROR(A1, "Errore di calcolo")
```

### Esercizio 2:
Hai un elenco di numeri nella colonna A. Vuoi calcolare il reciproco di ciascun numero e visualizzare "Errore" se il numero è zero.

**Formula:**
```excel
=IFERROR(1/A1, "Errore")
```

### Esercizio 3:
Hai un elenco di valori nella colonna A. Vuoi visualizzare "Maggiore di 10" se il valore è maggiore di 10, altrimenti "Minore o uguale a 10".

**Formula:**
```excel
=IFERROR(IF(A1>10, "Maggiore di 10", "Minore o uguale a 10"), "Errore")
```

### Esercizio 4:
Hai un elenco di prezzi nella colonna A e un elenco di quantità nella colonna B. Vuoi calcolare l'importo totale e visualizzare "Errore" se uno dei valori è errato.

**Formula:**
```excel
=IFERROR(A1*B1, "Errore")
```

### Esercizio 5:
Hai un elenco di date nella colonna A. Vuoi visualizzare "Data valida" se la data è valida, altrimenti "Data non valida".

**Formula:**
```excel
=IFERROR(IF(DATE(YEAR(A1), MONTH(A1), DAY(A1))=A1, "Data valida", "Data non valida"), "Errore")
```

### Esercizio 6:
Hai un elenco di stringhe di testo nella colonna A. Vuoi visualizzare "Lunghezza valida" se la lunghezza della stringa è inferiore a 10 caratteri, altrimenti "Lunghezza non valida".

**Formula:**
```excel
=IFERROR(IF(LEN(A1)<10, "Lunghezza valida", "Lunghezza non valida"), "Errore")
```

### Esercizio 7:
Hai un elenco di numeri nella colonna A. Vuoi visualizzare "Numero positivo" se il numero è positivo, altrimenti "Numero non positivo".

**Formula:**
```excel
=IFERROR(IF(A1>0, "Numero positivo", "Numero non positivo"), "Errore")
```

### Esercizio 8:
Hai un elenco di valori nella colonna A. Vuoi visualizzare "Pari" se il valore è pari, altrimenti "Dispari".

**Formula:**
```excel
=IFERROR(IF(MOD(A1,2)=0, "Pari", "Dispari"), "Errore")
```

### Esercizio 9:
Hai un elenco di percentuali nella colonna A. Vuoi visualizzare "Valido" se la percentuale è compresa tra 0 e 100, altrimenti "Non valido".

**Formula:**
```excel
=IFERROR(IF(AND(A1>=0, A1<=100), "Valido", "Non valido"), "Errore")
```

### Esercizio 10:
Hai un elenco di codici nella colonna A. Vuoi visualizzare "Formato valido" se il codice segue un formato specifico, altrimenti "Formato non valido".

**Formula:**
```excel
=IFERROR(IF(REGEXMATCH(A1, "^[A-Z]{3}-\d{3}$"), "Formato valido", "Formato non valido"), "Errore")
```

Questi esercizi ti consentiranno di praticare l'utilizzo della funzione IFERROR in diverse situazioni. Puoi adattare le formule in base alle tue esigenze specifiche e sperimentare con altre condizioni per ampliare le tue competenze in Excel.

































