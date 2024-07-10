import streamlit as st
import pandas as pd

st.title('Excel Data Extractor')

# Upload the Excel file
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    try:
        # Read the Excel file
        data = pd.read_excel(uploaded_file, engine='openpyxl')
        
        # Display the first few rows of the data
        st.write("Here are the first few rows of your Excel file:")
        st.write(data.head())
        
        # Placeholder for further data processing and extraction logic
        # For example, extract specific columns or perform some transformation
        
        # Save the extracted data to a new Excel file
        extracted_data = data  # Modify this to your specific extracted data
        output_file = "extracted_data.xlsx"
        extracted_data.to_excel(output_file, index=False)
        
        # Provide a download link for the new Excel file
        st.download_button(
            label="Download Extracted Data",
            data=open(output_file, "rb").read(),
            file_name=output_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")

