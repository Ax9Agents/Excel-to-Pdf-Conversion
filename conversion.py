import streamlit as st
import os
import tempfile
import zipfile
from io import BytesIO
from win32com import client
import pythoncom

def convert_excel_to_pdf_bytes(excel_bytes, excel_filename):
    """
    Convert Excel file bytes to PDF bytes.
    
    Args:
        excel_bytes: Bytes of the Excel file
        excel_filename: Original filename of the Excel file
    
    Returns:
        PDF file as bytes
    """
    # Initialize COM
    pythoncom.CoInitialize()
    
    try:
        # Create temporary files
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(excel_filename)[1]) as tmp_excel:
            tmp_excel.write(excel_bytes)
            excel_path = tmp_excel.name
        
        pdf_path = excel_path.replace(os.path.splitext(excel_filename)[1], '.pdf')
        
        # Open Microsoft Excel
        excel = client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # Open and convert
        workbook = excel.Workbooks.Open(excel_path)
        workbook.ExportAsFixedFormat(0, pdf_path)
        workbook.Close(False)
        excel.Quit()
        
        # Read PDF bytes
        with open(pdf_path, 'rb') as pdf_file:
            pdf_bytes = pdf_file.read()
        
        # Cleanup
        os.unlink(excel_path)
        os.unlink(pdf_path)
        
        return pdf_bytes
        
    except Exception as e:
        raise e
    
    finally:
        pythoncom.CoUninitialize()

def create_zip_file(pdf_files_dict):
    """
    Create a ZIP file containing multiple PDFs.
    
    Args:
        pdf_files_dict: Dictionary with {filename: pdf_bytes}
    
    Returns:
        ZIP file as bytes
    """
    zip_buffer = BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for filename, pdf_bytes in pdf_files_dict.items():
            zip_file.writestr(filename, pdf_bytes)
    
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

def main():
    st.set_page_config(
        page_title="Excel to PDF Converter",
        page_icon="📄",
        layout="centered"
    )
    
    st.title("📄 Excel to PDF Converter")
    st.markdown("Upload one or multiple Excel files and convert them to PDF format.")
    
    # File uploader
    uploaded_files = st.file_uploader(
        "Choose Excel file(s)",
        type=['xlsx', 'xls', 'xlsm'],
        accept_multiple_files=True,
        help="Upload one or more Excel files to convert to PDF"
    )
    
    if uploaded_files:
        st.success(f"✓ {len(uploaded_files)} file(s) uploaded successfully!")
        
        # Display uploaded files
        st.subheader("Uploaded Files:")
        for file in uploaded_files:
            st.write(f"• {file.name} ({file.size / 1024:.2f} KB)")
        
        # Convert button
        if st.button("🔄 Convert to PDF", type="primary", use_container_width=True):
            pdf_files = {}
            
            # Progress bar
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Convert each file
            for idx, uploaded_file in enumerate(uploaded_files):
                try:
                    status_text.text(f"Converting {uploaded_file.name}...")
                    
                    # Convert to PDF
                    pdf_bytes = convert_excel_to_pdf_bytes(
                        uploaded_file.getvalue(),
                        uploaded_file.name
                    )
                    
                    # Store with PDF extension
                    pdf_filename = os.path.splitext(uploaded_file.name)[0] + '.pdf'
                    pdf_files[pdf_filename] = pdf_bytes
                    
                    # Update progress
                    progress_bar.progress((idx + 1) / len(uploaded_files))
                    
                except Exception as e:
                    st.error(f"✗ Error converting {uploaded_file.name}: {str(e)}")
            
            status_text.text("Conversion complete!")
            progress_bar.empty()
            
            if pdf_files:
                st.success(f"✓ Successfully converted {len(pdf_files)} file(s)!")
                
                st.subheader("Download Options:")
                
                # Single file download
                if len(pdf_files) == 1:
                    filename, pdf_bytes = list(pdf_files.items())[0]
                    st.download_button(
                        label="⬇️ Download PDF",
                        data=pdf_bytes,
                        file_name=filename,
                        mime="application/pdf",
                        use_container_width=True
                    )
                
                # Multiple files - individual and ZIP download
                else:
                    # Individual download buttons
                    col1, col2 = st.columns([1, 1])
                    
                    with col1:
                        st.markdown("**Individual Downloads:**")
                        for filename, pdf_bytes in pdf_files.items():
                            st.download_button(
                                label=f"⬇️ {filename}",
                                data=pdf_bytes,
                                file_name=filename,
                                mime="application/pdf",
                                key=filename
                            )
                    
                    with col2:
                        st.markdown("**Bulk Download:**")
                        # Create ZIP file
                        zip_bytes = create_zip_file(pdf_files)
                        st.download_button(
                            label="📦 Download All as ZIP",
                            data=zip_bytes,
                            file_name="converted_pdfs.zip",
                            mime="application/zip",
                            type="primary"
                        )
    
    else:
        st.info("👆 Upload Excel files to get started")
    
    # Footer
    st.markdown("---")
    st.markdown(
        """
        <div style='text-align: center; color: gray;'>
            <small>Supports .xlsx, .xls, and .xlsm formats</small>
        </div>
        """,
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
