# VBA 


VBA è un linguaggio di programmazione integrato in Microsoft Excel e in altre applicazioni di Microsoft Office. 
Consente agli utenti di automatizzare compiti ripetitivi, creare funzioni personalizzate e sviluppare applicazioni più complesse all'interno di Excel.

Con VBA, si possono creare macro, che sono sequenze di istruzioni che eseguono operazioni specifiche. 
Ad esempio, si può automatizzare la formattazione di fogli di lavoro, l'importazione di dati da altre fonti, o la creazione di report personalizzati. 


https://www.tutorialspoint.com/vba/vba_variables.htm

# SETUP 

![Screenshot 2025-04-16 alle 14 55 22](https://github.com/user-attachments/assets/c181a2f9-77f2-4da5-b071-649cc9e74a32)


![Screenshot 2025-04-16 alle 14 55 38](https://github.com/user-attachments/assets/97a9843c-3e5a-4b43-bb20-be1f39e7cf0d)


![Screenshot 2025-04-16 alle 14 55 48](https://github.com/user-attachments/assets/047f17e8-2779-4e05-833b-b509950b79ff)


![Screenshot 2025-04-16 alle 15 17 53](https://github.com/user-attachments/assets/811e4af9-44fa-496b-8d97-1bcf2659bad4)



Per gli utenti windows :


<img width="1152" alt="Screenshot 2025-04-16 alle 15 26 19" src="https://github.com/user-attachments/assets/84997321-e093-4c33-9b5f-feba32018167" />


<img width="709" alt="Screenshot 2025-04-16 alle 15 25 46" src="https://github.com/user-attachments/assets/a78111ff-7d3b-4bcf-8618-e63e52ad733a" />


<img width="718" alt="Screenshot 2025-04-16 alle 15 25 32" src="https://github.com/user-attachments/assets/3adbffc6-a330-4e54-9f15-1a0efc61b3d2" />



COME ATTIVARE LE MACRO IN EXCEL
✅ 1. Attiva la scheda "Sviluppo"
Se non è già visibile:
Vai su File > Opzioni
Clicca su Personalizzazione barra multifunzione
A destra, spunta la voce “Sviluppo” (Developer)
Clicca su OK
Ora vedrai la scheda "Sviluppo" nella barra in alto di Excel.



# TIPI DI DATI :


In VBA (Visual Basic for Applications), i tipi di dati sono fondamentali per definire la natura delle variabili e come verranno utilizzate nel codice. 


_______________________


1. **Integer**: Utilizzato per memorizzare numeri interi compresi tra -32.768 e 32.767. È utile per contatori e operazioni matematiche semplici.

2. **Long**: Simile a Integer, ma può contenere numeri interi più grandi, da -2.147.483.648 a 2.147.483.647.

3. **Single**: Utilizzato per memorizzare numeri in virgola mobile a precisione singola. È utile per valori decimali, ma con una precisione limitata.

4. **Double**: Utilizzato per memorizzare numeri in virgola mobile a precisione doppia. È più preciso di Single e può gestire numeri molto grandi o molto piccoli.

5. **Currency**: Utilizzato per memorizzare valori monetari. Ha una precisione fissa di quattro decimali e può gestire numeri fino a 15 cifre.

6. **String**: Utilizzato per memorizzare sequenze di caratteri, come parole o frasi. Può contenere fino a circa 2 miliardi di caratteri.

7. **Boolean**: Utilizzato per memorizzare valori logici, che possono essere True (vero) o False (falso).

8. **Date**: Utilizzato per memorizzare date e orari. Può gestire date da gennaio 1, 1753, a dicembre 31, 9999.

9. **Variant**: Un tipo di dato speciale che può contenere qualsiasi tipo di dato, inclusi numeri, stringhe, date e array. È flessibile, ma può essere meno efficiente in termini di prestazioni.

10. **Object**: Utilizzato per riferirsi a oggetti, come fogli di lavoro, celle o altre entità in Excel.

11. **Array**: Un tipo di dato che può contenere più valori dello stesso tipo. Gli array possono essere unidimensionali o multidimensionali.





### 1. **Integer**
```vba
Dim contatore As Integer
contatore = 10
MsgBox "Il valore del contatore è: " & contatore
```

### 2. **Long**
```vba
Dim grandeNumero As Long
grandeNumero = 1234567890
MsgBox "Il grande numero è: " & grandeNumero
```

### 3. **Single**
```vba
Dim valoreDecimale As Single
valoreDecimale = 3.14
MsgBox "Il valore decimale è: " & valoreDecimale
```

### 4. **Double**
```vba
Dim numeroPreciso As Double
numeroPreciso = 3.14159265358979
MsgBox "Il numero preciso è: " & numeroPreciso
```

### 5. **Currency**
```vba
Dim prezzo As Currency
prezzo = 19.99
MsgBox "Il prezzo è: " & prezzo & " €"
```

### 6. **String**
```vba
Dim nome As String
nome = "Mario Rossi"
MsgBox "Il nome è: " & nome
```

### 7. **Boolean**
```vba
Dim attivo As Boolean
attivo = True
If attivo Then
    MsgBox "L'utente è attivo."
Else
    MsgBox "L'utente non è attivo."
End If
```

### 8. **Date**
```vba
Dim dataAttuale As Date
dataAttuale = Now
MsgBox "La data e l'ora attuale sono: " & dataAttuale
```

### 9. **Variant**
```vba
Dim variabile As Variant
variabile = "Testo"
MsgBox "La variabile contiene: " & variabile
variabile = 123
MsgBox "Ora la variabile contiene: " & variabile
```

### 10. **Object**
```vba
Dim foglio As Worksheet
Set foglio = ThisWorkbook.Sheets("Foglio1")
MsgBox "Il nome del foglio è: " & foglio.Name
```

### 11. **Array**
```vba
Dim numeri(1 To 5) As Integer
numeri(1) = 10
numeri(2) = 20
numeri(3) = 30
numeri(4) = 40
numeri(5) = 50

Dim i As Integer
Dim risultato As String
risultato = "I numeri sono: "
For i = 1 To 5
    risultato = risultato & numeri(i) & " "
Next i
MsgBox risultato
```





_______________________




# ISTRUZIONI DI SELEZIONE








# ISTRUZIONI DI CICLI 


_______



