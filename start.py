import pandas as pd
import streamlit as st

st.title('Excel Data Transformer')

# Upload the Excel file
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    # Load the Excel file starting from row 21 (skip the first 20 rows)
    try:
        uploaded_data = pd.read_excel(uploaded_file, skiprows=20, engine='openpyxl')
        st.write("Data loaded successfully. Here's a preview:")
        st.write(uploaded_data.head(10))
    except Exception as e:
        st.error(f"Error loading the file: {e}")
        st.stop()
    
    # Display the DataFrame columns
    st.write("DataFrame Columns:")
    st.write(uploaded_data.columns)
    
    # Display rows around the expected data to help with debugging
    st.write("Rows containing 'Prezzo':")
    prezzo_rows = uploaded_data[uploaded_data.apply(lambda row: row.astype(str).str.contains('Prezzo', case=False, na=False).any(), axis=1)]
    st.write(prezzo_rows)
    
    # Function to safely extract data from the DataFrame
    def safe_extract(df, condition, column):
        try:
            df['DETTAGLI RIGA ARTICOLO'] = df['DETTAGLI RIGA ARTICOLO'].astype(str).str.strip()
            value = df[df['DETTAGLI RIGA ARTICOLO'].str.contains(condition, case=False, na=False)][column].values[0]
            return value
        except IndexError:
            st.error(f"Could not find value for condition: '{condition}' in column: '{column}'")
            return None
    
    # Extract relevant information from the uploaded data
    articolo = safe_extract(uploaded_data, 'Modello/Colore:', 'Unnamed: 1')
    descrizione = safe_extract(uploaded_data, 'Nome del modello:', 'Unnamed: 1')
    categoria = safe_extract(uploaded_data, 'Tipo di prodotto:', 'Unnamed: 1')
    colore = safe_extract(uploaded_data, 'Descrizione colore:', 'Unnamed: 1')
    qta = safe_extract(uploaded_data, 'Riga articolo:', 'Unnamed: 1')
    prezzo = safe_extract(uploaded_data, "Prezzo all'ingrosso", 'Unnamed: 7')  # Corrected column
    
    if qta is not None:
        qta = int(qta)
    if prezzo is not None:
        prezzo = float(prezzo.replace('€', '').replace(',', '.').strip())
    
    if None in [articolo, descrizione, categoria, colore, qta, prezzo]:
        st.error("Some required data is missing. Please check the input file.")
        st.stop()
    
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
