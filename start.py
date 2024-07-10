import pandas as pd
import streamlit as st

st.title('Excel Data Transformer')

# Upload the Excel file
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    # Load the Excel file starting from row 21 (skip the first 20 rows)
    uploaded_data = pd.read_excel(uploaded_file, skiprows=20, engine='openpyxl')
    
    # Extract relevant information from the uploaded data
    articolo = uploaded_data.loc[uploaded_data['DETTAGLI RIGA ARTICOLO'] == 'Modello/Colore:', 'Unnamed: 1'].values[0]
    descrizione = uploaded_data.loc[uploaded_data['DETTAGLI RIGA ARTICOLO'] == 'Nome del modello:', 'Unnamed: 1'].values[0]
    categoria = uploaded_data.loc[uploaded_data['DETTAGLI RIGA ARTICOLO'] == 'Tipo di prodotto:', 'Unnamed: 1'].values[0]
    colore = uploaded_data.loc[uploaded_data['DETTAGLI RIGA ARTICOLO'] == 'Descrizione colore:', 'Unnamed: 1'].values[0]
    qta = int(uploaded_data.loc[uploaded_data['DETTAGLI RIGA ARTICOLO'] == 'Riga articolo:', 'Unnamed: 1'].values[0])
    prezzo = uploaded_data.loc[uploaded_data['DETTAGLI RIGA ARTICOLO'] == "Prezzo all'ingrosso", 'Unnamed: 1'].values[0]
    
    # Convert price from string to float
    prezzo = float(prezzo.replace('€', '').replace(',', '.').strip())
    
    # Create the final output DataFrame
    output_data = pd.DataFrame({
        'ARTICOLO': [articolo],
        'DESCRIZIONE': [descrizione],
        'CATEGORIA': [categoria],
        'COLORE': [colore],
        'TAGLIA': [None],  # Placeholder, as information is not available
        'BARCODE': [None],  # Placeholder, as information is not available
        'SPEC_MATERIALE': [None],  # Placeholder, as information is not available
        'MADEIN': [None],  # Placeholder, as information is not available
        'ID_ORDINE': [None],  # Placeholder, as information is not available
        'QTA': [qta],
        'PREZZO+-IVA': [prezzo],
        'PREZZO_NETTO': [None],  # Placeholder, as information is not available
        '%': [None],  # Placeholder, as information is not available
        'IMPORTO': [None],  # Placeholder, as information is not available
        'HSCODE': [None]  # Placeholder, as information is not available
    })
    
    # Display the transformed data
    st.write("Transformed Data:")
    st.write(output_data)
    
    # Save the transformed data to a new Excel file
    output_file = "transformed_data.xlsx"
    output_data.to_excel(output_file, index=False)
    
    # Provide a download link for the new Excel file
    st.download_button(
        label="Download Transformed Data",
        data=open(output_file, "rb").read(),
        file_name=output_file,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
