import streamlit as st
import docx
import pandas as pd
import io
import os

st.set_page_config(page_title="Word to Excel Converter", page_icon="📝", layout="wide")

# Custom CSS for a more premium look
st.markdown("""
    <style>
    .main {
        background-color: #f8f9fa;
    }
    .stButton>button {
        width: 100%;
        border-radius: 5px;
        height: 3em;
        background-color: #007bff;
        color: white;
    }
    .stDownloadButton>button {
        width: 100%;
        border-radius: 5px;
        height: 3em;
        background-color: #28a745;
        color: white;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("📝 Spec Table Extractor")
st.markdown("Upload a Word document to extract tables into an Excel file.")

uploaded_file = st.file_uploader("Choose a Word file", type="docx")

if uploaded_file is not None:
    try:
        doc = docx.Document(uploaded_file)
        
        if not doc.tables:
            st.error("No tables found in the uploaded document.")
        else:
            st.success(f"Found {len(doc.tables)} table(s)!")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for i, table in enumerate(doc.tables):
                    data = []
                    for row in table.rows:
                        row_data = [cell.text.strip() for cell in row.cells]
                        data.append(row_data)
                    
                    df = pd.DataFrame(data)
                    
                    if not df.empty:
                        df.columns = df.iloc[0]
                        df = df[1:]
                    
                    sheet_name = f"Table_{i+1}"
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    with st.expander(f"Preview Table {i+1}"):
                        st.dataframe(df)
            
            # Finalize the data
            xlsx_data = output.getvalue()
            
            # Determine download filename
            base_name = os.path.splitext(uploaded_file.name)[0]
            download_filename = f"{base_name}.xlsx"
            
            st.info(f"Ready to download: **{download_filename}**")

            st.download_button(
                label="📥 Download Excel File",
                data=xlsx_data,
                file_name=download_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"An error occurred: {e}")

st.divider()
st.info("Tip: You can drag and drop your file directly into the box above.")
