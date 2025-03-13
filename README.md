Uno script Python che prende in input un file Excel con indirizzi e li confronta con un indirizzo di riferimento.
Gli output includono: distanza a piedi, tempo di guida e una mappa Google per la distanza in auto.
Se non siete sviluppatori scaricate direttamente la cartella .exe ed avviate il file relative_distance ivi contenuto.
Istruzioni:

        "1. Inserisci la tua API Key di Google Maps."
        "2. Seleziona il file Excel contenente gli indirizzi in una colonna chiamata 'indirizzo'."
        "3. Inserisci l'indirizzo di riferimento (formato ideale: via, numero civico, cap, città)."
        "4. Specifica la cartella di output e il nome del file Excel dei risultati."
        "5. Clicca su 'Avvia Calcolo'."
    
        "Note:Il programma ha bisogno di almeno via, numero civico e cap per funzionare adeguatamente"
        "Su excel si possono fondere le colonne di indirizzo e cap con la funzione =A1 & ", " & B1 " dove A1 e B1 sono chiaramente le colonne di indirizzo e cap
        "Sarà poi importante copiare i risultati come valori su una nuova colonna 'indirizzo'"
        "Il programmma crea anche un file apikey.json per salvare la key google"
