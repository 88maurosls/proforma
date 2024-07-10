import pandas as pd
import streamlit as st

st.title('Trasformatore di Dati Excel')

# Carica il file Excel
uploaded_file = st.file_uploader("Scegli un file Excel", type="xlsx")

if uploaded_file is not None:
    try:
        # Prova a leggere il file Excel con openpyxl
        uploaded_data = pd.read_excel(uploaded_file, skiprows=20, engine='openpyxl')
        st.write("Dati caricati con successo. Ecco un'anteprima:")
        st.write(uploaded_data.head(10))
    except Exception as e:
        st.error(f"Errore durante il caricamento del file: {e}")
        st.stop()
    
    # Mostra le colonne del DataFrame
    st.write("Colonne del DataFrame:")
    st.write(uploaded_data.columns)
    
    # Mostra le righe che contengono 'Prezzo' per il debug
    st.write("Righe contenenti 'Prezzo':")
    prezzo_rows = uploaded_data[uploaded_data.apply(lambda row: row.astype(str).str.contains('Prezzo', case=False, na=False).any(), axis=1)]
    st.write(prezzo_rows)
    
    # Funzione per estrarre i dati in modo sicuro dal DataFrame
    def safe_extract(df, condition, column):
        try:
            df['DETTAGLI RIGA ARTICOLO'] = df['DETTAGLI RIGA ARTICOLO'].astype(str).str.strip().str.lower()
            condition = condition.lower().strip()
            value = df[df['DETTAGLI RIGA ARTICOLO'].str.contains(condition, case=False, na=False)][column].values[0]
            return value
        except IndexError:
            st.error(f"Non è stato possibile trovare il valore per la condizione: '{condition}' nella colonna: '{column}'")
            return None
    
    # Estrai le informazioni rilevanti dai dati caricati
    articolo = safe_extract(uploaded_data, 'Modello/Colore:', 'Unnamed: 1')
    descrizione = safe_extract(uploaded_data, 'Nome del modello:', 'Unnamed: 1')
    categoria = safe_extract(uploaded_data, 'Tipo di prodotto:', 'Unnamed: 1')
    colore = safe_extract(uploaded_data, 'Descrizione colore:', 'Unnamed: 1')
    qta = safe_extract(uploaded_data, 'Riga articolo:', 'Unnamed: 1')
    
    # Estrai "Prezzo all'ingrosso" usando 'Unnamed: 7'
    try:
        prezzo_row = uploaded_data[uploaded_data['Unnamed: 6'].str.contains("Prezzo all'ingrosso", case=False, na=False)]
        prezzo = prezzo_row['Unnamed: 7'].values[0]
    except IndexError:
        st.error("Non è stato possibile trovare il valore per 'Prezzo all'ingrosso' nella colonna prevista.")
        prezzo = None
    
    # Estrai "Codice a Barre (UPC)" e "Misura"
    try:
        sizes_barcodes = uploaded_data[(uploaded_data['DETTAGLI RIGA ARTICOLO'].str.contains(r'^\d+(\.\d+)?$', na=False))]  # Regex per corrispondere ai valori numerici (misure)
        sizes = sizes_barcodes['DETTAGLI RIGA ARTICOLO'].values.tolist()
        barcodes = sizes_barcodes['Unnamed: 1'].values.tolist()
        st.write("Taglie e Codici a Barre Estratti:")
        st.write(sizes_barcodes)
    except IndexError:
        st.error("Non è stato possibile trovare 'Misura' o 'Codice a Barre (UPC)' nelle colonne previste.")
        sizes = []
        barcodes = []
    
    # Informazioni di debug aggiuntive
    st.write("Valori Estratti:")
    st.write(f"Articolo: {articolo}")
    st.write(f"Descrizione: {descrizione}")
    st.write(f"Categoria: {categoria}")
    st.write(f"Colore: {colore}")
    st.write(f"QTA: {qta}")
    st.write(f"Prezzo: {prezzo}")
    st.write(f"Taglie: {sizes}")
    st.write(f"Codici a Barre: {barcodes}")
    
    if qta is not None:
        qta = int(qta)
    if prezzo is not None:
        prezzo = float(prezzo.replace('€', '').replace(',', '.').strip())
    
    if None in [articolo, descrizione, categoria, colore, qta, prezzo] or not sizes or not barcodes:
        st.error("Alcuni dati richiesti sono mancanti. Si prega di controllare il file di input.")
        st.stop()
    
    # Crea il DataFrame finale di output
    output_data = pd.DataFrame({
        'ARTICOLO': [articolo] * len(sizes),
        'DESCRIZIONE': [descrizione] * len(sizes),
        'CATEGORIA': [categoria] * len(sizes),
        'COLORE': [colore] * len(sizes),
        'TAGLIA': sizes,
        'BARCODE': barcodes,
        'SPEC_MATERIALE': [None] * len(sizes),  # Placeholder, dato non disponibile
        'MADEIN': [None] * len(sizes),  # Placeholder, dato non disponibile
        'ID_ORDINE': [None] * len(sizes),  # Placeholder, dato non disponibile
        'QTA': [qta] * len(sizes),
        'PREZZO+-IVA': [prezzo] * len(sizes),
        'PREZZO_NETTO': [None] * len(sizes),  # Placeholder, dato non disponibile
        '%': [None] * len(sizes),  # Placeholder, dato non disponibile
        'IMPORTO': [None] * len(sizes),  # Placeholder, dato non disponibile
        'HSCODE': [None] * len(sizes)  # Placeholder, dato non disponibile
    })
    
    # Mostra i dati trasformati
    st.write("Dati Trasformati:")
    st.write(output_data)
    
    # Salva i dati trasformati in un nuovo file Excel
    output_file = "dati_trasformati.xlsx"
    output_data.to_excel(output_file, index=False)
    
    # Fornisci un link per scaricare il nuovo file Excel
    st.download_button(
        label="Scarica i Dati Trasformati",
        data=open(output_file, "rb").read(),
        file_name=output_file,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
