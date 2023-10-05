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


___________________________________________


## Esercizi Misti Exel :



### Esercizio 1:
Hai un elenco di età nella colonna A e un elenco di punteggi nella colonna B. Vuoi assegnare "Ammesso" solo se l'età è maggiore o uguale a 18 e il punteggio è maggiore di 70.

**Formula:**
```excel
=IF(AND(A1>=18, B1>70), "Ammesso", "Non Ammesso")
```

### Esercizio 2:
Hai un elenco di voti nella colonna A. Vuoi calcolare la media dei voti solo se tutti i voti sono superiori a 50.

**Formula:**
```excel
=IF(AND(A1>50, B1>50, C1>50, D1>50, E1>50), (A1+B1+C1+D1+E1)/5, "Almeno un voto inferiore a 50")
```

### Esercizio 3:
Hai un elenco di numeri nella colonna A. Vuoi assegnare "Numero positivo" se il numero è positivo, "Zero" se il numero è zero, e "Numero negativo" se il numero è negativo.

**Formula:**
```excel
=IF(A1>0, "Numero positivo", IF(A1=0, "Zero", "Numero negativo"))
```

### Esercizio 4:
Hai un elenco di temperature nella colonna A. Vuoi assegnare "Caldo" se la temperatura è superiore a 30 gradi Celsius o "Freddo" se è inferiore a 10 gradi Celsius.

**Formula:**
```excel
=IF(OR(A1>30, A1<10), IF(A1>30, "Caldo", "Freddo"), "Temperatura moderata")
```

### Esercizio 5:
Hai un elenco di valori numerici nella colonna A. Vuoi assegnare "Pari" se il numero è pari e "Dispari" se è dispari. Se il valore non è un numero, visualizza "Non è un numero".

**Formula:**
```excel
=IF(ISNUMBER(A1), IF(MOD(A1,2)=0, "Pari", "Dispari"), "Non è un numero")
```

_________________


TIPS :


Per verificare se tutti i dati in una colonna sono pari in Excel, puoi utilizzare la funzione `MOD` insieme alla funzione `SUMPRODUCT`. Ecco come farlo:

Supponiamo che i dati che desideri verificare siano nella colonna A da A1 ad A1000. La formula per verificare se tutti i dati nella colonna A sono pari sarebbe la seguente:

```excel
=IF(SUMPRODUCT(MOD(A1:A1000, 2))=0, "Tutti pari", "Non tutti pari")
```

Questa formula utilizza `MOD(A1:A1000, 2)` per ottenere il resto della divisione di ogni numero nella colonna A per 2. Se tutti i numeri sono pari, la somma di questi resti sarà 0. La funzione `SUMPRODUCT` somma questi resti, e se il risultato è 0, la formula restituirà "Tutti pari", altrimenti "Non tutti pari". Puoi adattare l'intervallo `A1:A1000` nella formula in base alla tua esigenza specifica.


_______________________

 ## RANGE IN EXEL 


In Excel, puoi specificare un intervallo (range) utilizzando la notazione A1:B10, dove A1 rappresenta la cella in alto a sinistra del tuo intervallo e B10 rappresenta la cella in basso a destra. Questo intervallo include tutte le celle dalla A1 alla B10.

Ecco alcuni modi comuni per specificare un intervallo in Excel:

### Intervallo singolo:
- **A1**: Rappresenta una singola cella nella colonna A e nella riga 1.
- **A1:B10**: Rappresenta tutte le celle nel rettangolo dalla A1 alla B10.
- **C**: Rappresenta l'intera colonna C.
- **2**: Rappresenta l'intera riga 2.

### Intervallo combinato:
- **A1, B3, C5**: Rappresenta tre celle separate: A1, B3 e C5.
- **A1:B2, C3:D4**: Rappresenta due intervalli distinti: A1:B2 e C3:D4.

### Intervallo dinamico:
- **A1:A**: Rappresenta l'intera colonna A a partire dalla cella A1.
- **1:1**: Rappresenta l'intera riga 1 a partire dalla colonna A.

### Utilizzo di nomi:
Puoi assegnare un nome a un intervallo per riferirti più facilmente ad esso. Ad esempio, assegnando il nome "MioIntervallo" all'intervallo A1:B10, puoi fare riferimento a questo intervallo utilizzando il nome "MioIntervallo" nelle formule anziché l'indirizzo delle celle.

### Utilizzo di formule e funzioni:
Gli intervalli possono essere specificati anche all'interno di formule e funzioni. Ad esempio, puoi sommare tutti i valori in un intervallo utilizzando la formula `=SUM(A1:B10)`.

Puoi inserire un intervallo direttamente nelle caselle di input delle formule o nelle finestre di dialogo delle formule di Excel. 




__________________


## Funzioni di data :


In Excel, ci sono diverse funzioni di data che puoi utilizzare per eseguire operazioni su dati di data e ora. 

il formato e la lingua del tuo Excel possono influenzare la rappresentazione delle date e le funzioni disponibili. Puoi consultare la documentazione di Excel o il menu di aiuto di Excel per ulteriori dettagli e opzioni specifiche relative alle funzioni di data nel tuo ambiente Excel specifico.


Ecco un elenco delle principali funzioni di data in Excel:

### Funzioni di Data di Base:
1. **TODAY()**: Restituisce la data odierna.
2. **NOW()**: Restituisce la data e l'orario correnti.
3. **DATE(anno, mese, giorno)**: Restituisce una data in base agli argomenti specificati.
4. **TIME(ora, minuto, secondo)**: Restituisce un orario in base agli argomenti specificati.

### Estrazione di Componenti dalla Data:
5. **YEAR(data)**: Restituisce l'anno dalla data specificata.
6. **MONTH(data)**: Restituisce il mese dalla data specificata (da 1 a 12).
7. **DAY(data)**: Restituisce il giorno del mese dalla data specificata.
8. **HOUR(ora)**: Restituisce l'ora dalla data o dall'orario specificato.
9. **MINUTE(ora)**: Restituisce i minuti dalla data o dall'orario specificato.
10. **SECOND(ora)**: Restituisce i secondi dalla data o dall'orario specificato.

### Operazioni su Date:
11. **DATEVALUE(testo)**: Converte una data in formato testo in un valore numerico della data.
12. **TIMEVALUE(testo)**: Converte un orario in formato testo in un valore numerico dell'orario.
13. **DATEDIF(data_iniziale, data_finale, "unità")**: Calcola la differenza tra due date in base all'unità specificata ("y" per anni, "m" per mesi, "d" per giorni, ecc.).

### Formattazione di Date e Orari:
14. **TEXT(data, "formato")**: Converte una data o un orario in formato testo utilizzando il formato specificato.
15. **DAYNAME(data)**: Restituisce il nome del giorno dalla data specificata.
16. **MONTHNAME(data)**: Restituisce il nome del mese dalla data specificata.

### Manipolazione Avanzata:
17. **EDATE(data, numero_mesi)**: Restituisce la data che si trova a un certo numero di mesi prima o dopo la data specificata.
18. **EOMONTH(data, numero_mesi)**: Restituisce l'ultimo giorno del mese, un numero specificato di mesi prima o dopo la data specificata.
19. **NETWORKDAYS(data_iniziale, data_finale, [festivi])**: Restituisce il numero di giorni lavorativi tra due date, escludendo i giorni festivi specificati.


_____________________________________


## Esercitazione :


### 1. Esercizio: Data Odierna
**Descrizione:** Restituisci la data odierna.
**Formula:** `=TODAY()`
**Soluzione:** 01/10/2023 (formato data: gg/mm/aaaa)

### 2. Esercizio: Data e Orario Correnti
**Descrizione:** Restituisci la data e l'orario correnti.
**Formula:** `=NOW()`
**Soluzione:** 01/10/2023 14:30:00 (formato data: gg/mm/aaaa hh:mm:ss)

### 3. Esercizio: Anno dalla Data
**Descrizione:** Estrai l'anno dalla data "15/10/2023".
**Formula:** `=YEAR(DATE(2023, 10, 15))`
**Soluzione:** 2023

### 4. Esercizio: Mese dalla Data
**Descrizione:** Estrai il mese dalla data "15/10/2023".
**Formula:** `=MONTH(DATE(2023, 10, 15))`
**Soluzione:** 10

### 5. Esercizio: Giorno dalla Data
**Descrizione:** Estrai il giorno dalla data "15/10/2023".
**Formula:** `=DAY(DATE(2023, 10, 15))`
**Soluzione:** 15

### 6. Esercizio: Data in Testo
**Descrizione:** Converte la data "15/10/2023" in formato testo.
**Formula:** `=TEXT(DATE(2023, 10, 15), "dd/mm/yyyy")`
**Soluzione:** 15/10/2023

### 7. Esercizio: Ultimo Giorno del Mese
**Descrizione:** Calcola l'ultimo giorno del mese per "15/10/2023".
**Formula:** `=EOMONTH(DATE(2023, 10, 15), 0)`
**Soluzione:** 31/10/2023

### 8. Esercizio: Aggiungi Mesi a una Data
**Descrizione:** Aggiungi 3 mesi alla data "15/10/2023".
**Formula:** `=EDATE(DATE(2023, 10, 15), 3)`
**Soluzione:** 15/01/2024

### 9. Esercizio: Differenza in Giorni tra Date
**Descrizione:** Calcola la differenza in giorni tra "15/10/2023" e "01/01/2023".
**Formula:** `=DATEDIF(DATE(2023, 1, 1), DATE(2023, 10, 15), "d")`
**Soluzione:** 287

### 10. Esercizio: Giorni Lavorativi tra Date
**Descrizione:** Calcola il numero di giorni lavorativi tra "01/10/2023" e "15/10/2023".
**Formula:** `=NETWORKDAYS(DATE(2023, 10, 1), DATE(2023, 10, 15))`
**Soluzione:** 11

### 11. Esercizio: Data in Testo Personalizzata
**Descrizione:** Restituisci la data "15/10/2023" come "15 ottobre 2023".
**Formula:** `=DAY(DATE(2023, 10, 15)) & " " & TEXT(DATE(2023, 10, 15), "mmmm yyyy")`
**Soluzione:** 15 ottobre 2023

### 12. Esercizio: Data Mese Giorno
**Descrizione:** Restituisci la data "15/10/2023" nel formato "MM/GG/AAAA".
**Formula:** `=TEXT(DATE(2023, 10, 15), "mm/dd/yyyy")`
**Soluzione:** 10/15/2023

### 13. Esercizio: Data Ultimo Giorno del Mese Successivo
**Descrizione:** Calcola l'ultimo giorno del mese successivo a "15/10/2023".
**Formula:** `=EOMONTH(DATE(2023, 10, 15), 1)`
**Soluzione:** 30/11/2023

### 14. Esercizio: Data e Orario Personalizzato
**Descrizione:** Restituisci la data e l'orario correnti nel formato "GG/MM/AAAA HH:MM AM/PM".
**Formula:** `=TEXT(NOW(), "dd/mm/yyyy hh:mm AM/PM")`
**Soluzione:** 01/10/2023 02:30 PM

### 15. Esercizio: Mese Successivo alla Data
**Descrizione:** Calcola il mese successivo a "15/10/2023".
**

Formula:** `=TEXT(DATE(2023, 10, 15) + 30, "mmmm")`
**Soluzione:** novembre

### 16. Esercizio: Data Ultimo Giorno dell'Anno Corrente
**Descrizione:** Calcola l'ultimo giorno dell'anno corrente.
**Formula:** `=EOMONTH(DATE(TODAY(), 1, 1), 11)`
**Soluzione:** 31/12/2023

### 17. Esercizio: Data del Primo Giorno del Mese Corrente
**Descrizione:** Calcola il primo giorno del mese corrente.
**Formula:** `=DATE(YEAR(TODAY()), MONTH(TODAY()), 1)`
**Soluzione:** 01/10/2023

### 18. Esercizio: Giorni Lavorativi tra Due Date Personalizzate
**Descrizione:** Calcola il numero di giorni lavorativi tra "01/01/2023" e "15/10/2023" escludendo i mercoledì.
**Formula:** `=NETWORKDAYS.INTL(DATE(2023, 1, 1), DATE(2023, 10, 15), 1111110)`
**Soluzione:** 199

### 19. Esercizio: Aggiungi 15 Giorni alla Data di Nascita
**Descrizione:** Aggiungi 15 giorni alla tua data di nascita.
**Formula:** `=DATE(YEAR("15/10/1990"), MONTH("15/10/1990"), DAY("15/10/1990") + 15)`
**Soluzione:** 30/10/1990

### 20. Esercizio: Aggiungi 2 Ore all'Orario Corrente
**Descrizione:** Aggiungi 2 ore all'orario corrente.
**Formula:** `=NOW() + TIME(2, 0, 0)`
**Soluzione:** 01/10/2023 16:30:00 (formato data: gg/mm/aaaa hh:mm:ss)

### 21. Esercizio: Nome del Giorno dalla Data Futura
**Descrizione:** Restituisci il nome del giorno per "15/11/2023".
**Formula:** `=TEXT(DATE(2023, 11, 15), "dddd")`
**Soluzione:** giovedì

### 22. Esercizio: Tronca l'Orario Corrente alle Ore
**Descrizione:** Tronca l'orario corrente alle ore.
**Formula:** `=TRUNC(NOW())`
**Soluzione:** 01/10/2023 00:00:00 (formato data: gg/mm/aaaa hh:mm:ss)

### 23. Esercizio: Tronca l'Orario Corrente alle Mezze Ore
**Descrizione:** Tronca l'orario corrente alle mezze ore.
**Formula:** `=MROUND(NOW(), TIME(0, 30, 0))`
**Soluzione:** 01/10/2023 14:30:00 (formato data: gg/mm/aaaa hh:mm:ss)

### 24. Esercizio: Calcola l'Era
**Descrizione:** Restituisci l'era per "15/10/2023" (A.C. o D.C.).
**Formula:** `=IF(YEAR(DATE(2023, 10, 15)) <= 0, "A.C.", "D.C.")`
**Soluzione:** D.C.

### 25. Esercizio: Calcola l'Età
**Descrizione:** Calcola l'età a partire dalla data di nascita "15/10/1990".
**Formula:** `=DATEDIF(DATE(1990, 10, 15), TODAY(), "Y")`
**Soluzione:** 32 anni



___________________________________________________


## 3 - Gestione dei file e stampe


NOTA : 
Ricorda che queste opzioni possono variare leggermente a seconda della versione specifica di Excel che stai utilizzando. Assicurati di esplorare attentamente il menu "File" e le opzioni di stampa nel tuo programma Excel per sfruttare appieno le funzionalità di gestione dei file e di stampa.



La gestione dei file e le stampe in Excel sono importanti aspetti che ti consentono di salvare, organizzare e condividere i tuoi documenti, nonché di stampare i dati per scopi diversi. Ecco alcune informazioni utili su come gestire i file e le stampe in Excel:

### Gestione dei File:
1. **Aprire un File Esistente:**
   - Vai su **File > Apri** e seleziona il file che desideri aprire.

2. **Salvare un File:**
   - Vai su **File > Salva o Salva con nome** per salvare il tuo file. Puoi anche premere **Ctrl + S** per salvare rapidamente le modifiche.

3. **Salvare una Copia del File:**
   - Vai su **File > Salva con nome** e scegli "Copia" per salvare una copia del file attuale.

4. **Chiudere un File:**
   - Vai su **File > Chiudi** per chiudere il file corrente.

5. **Creare un Nuovo File:**
   - Vai su **File > Nuovo** per creare un nuovo file. Puoi scegliere tra modelli predefiniti o iniziare con un foglio di lavoro vuoto.

6. **Gestione delle Versioni:**
   - Excel mantiene automaticamente diverse versioni del tuo file. Puoi accedere a queste versioni andando su **File > Informazioni > Versioni precedenti**.

### Stampa in Excel:
1. **Anteprima di Stampa:**
   - Vai su **File > Stampa** per vedere un'anteprima di come apparirà il foglio quando verrà stampato. Puoi modificare impostazioni come l'orientamento e la dimensione della carta.

2. **Impostazioni di Stampa:**
   - Puoi specificare l'area di stampa selezionando l'area desiderata e andando su **Layout di Pagina > Area di Stampa > Imposta Area di Stampa**.

3. **Stampa Rapida:**
   - Puoi premere **Ctrl + P** per avviare immediatamente il processo di stampa.

4. **Esportare in PDF:**
   - Vai su **File > Salva con nome** e scegli "PDF" come formato per esportare il foglio di lavoro come file PDF.

5. **Impostazioni Avanzate di Stampa:**
   - In **Layout di Pagina** puoi trovare varie opzioni come intestazioni e piè di pagina, l'orientamento della pagina, le dimensioni della carta e molto altro.

6. **Stampa di Determinate Parti del Foglio:**
   - Puoi stampare solo parti specifiche del tuo foglio selezionando l'area che desideri stampare prima di andare su **File > Stampa**.

7. **Stampa in Serie:**
   - Se hai più fogli di lavoro che desideri stampare in serie, puoi farlo utilizzando la funzione di stampa in serie di Excel. Vai su **File > Stampa > Stampa in Serie**.

______________


## 4  -- Importazione e esportazione di file in/da altri formati



Excel offre diverse opzioni per importare e esportare file in/da altri formati, consentendoti di lavorare con dati provenienti da diverse fonti o di condividere i tuoi dati in formati diversi. Ecco come puoi importare e esportare file in/da altri formati in Excel:

### Importazione di Dati:

1. **Importare da un File di Testo:**
   - Vai su **File > Apri**, seleziona il file di testo desiderato e segui l'importazione guidata. Puoi specificare delimitatori personalizzati come tabulazioni o virgole.

2. **Importare da un File CSV:**
   - Simile all'importazione da un file di testo, ma specifico per i file CSV (Comma-Separated Values). Vai su **File > Apri**, seleziona il file CSV e segui l'importazione guidata.

3. **Importare da un Database:**
   - Puoi importare dati da database come SQL Server, Access e altri. Vai su **Dati > Da Altre Origini > Da SQL Server/Access**, quindi inserisci le informazioni di connessione e importa i dati.

4. **Importare da un Sito Web:**
   - Vai su **Dati > Da Altre Origini > Da Sito Web**, inserisci l'URL del sito web da cui desideri importare i dati e segui l'assistente per l'importazione guidata.

### Esportazione di Dati:

1. **Esportare in un File di Testo o CSV:**
   - Seleziona l'area di dati che desideri esportare, vai su **File > Salva con nome**, scegli il percorso in cui desideri salvare il file, seleziona "Testo delimitato" o "CSV" come tipo di file e segui l'assistente di esportazione.

2. **Esportare in PDF:**
   - Seleziona l'area di dati che desideri esportare, vai su **File > Salva con nome**, seleziona la posizione in cui desideri salvare il PDF, scegli "PDF" come tipo di file e segui l'assistente di esportazione.

3. **Esportare in Formato Excel Più Vecchio:**
   - Vai su **File > Salva con nome**, seleziona la posizione in cui desideri salvare il file, scegli "Libro di Excel 97-2003 (*.xls)" come tipo di file e segui l'assistente di esportazione.

4. **Esportare in Formato PDF/A:**
   - Seleziona l'area di dati, vai su **File > Salva con nome**, scegli la posizione in cui desideri salvare il PDF/A, seleziona "PDF/A" come tipo di file e segui l'assistente di esportazione.

5. **Esportare in un Sito Web o in un Servizio Cloud:**
   - Alcuni servizi cloud come OneDrive e SharePoint consentono di esportare direttamente i tuoi fogli di lavoro in modo che siano accessibili online da qualsiasi dispositivo.
































