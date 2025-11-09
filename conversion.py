import streamlit as st
import os
import tempfile
import zipfile
from io import BytesIO
from win32com import client
import pythoncom


@st.cache_data(show_spinner=False)
def convert_excel_to_pdf_bytes(_excel_bytes, excel_filename):
    """
    Convert Excel file bytes to PDF bytes with caching.
    
    Args:
        _excel_bytes: Bytes of the Excel file (underscore prefix to skip hashing)
        excel_filename: Original filename of the Excel file
    
    Returns:
        PDF file as bytes
    """
    # Initialize COM
    pythoncom.CoInitialize()
    
    try:
        # Create temporary files
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(excel_filename)[1]) as tmp_excel:
            tmp_excel.write(_excel_bytes)
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
    """Create a ZIP file containing multiple PDFs."""
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
    
    # Initialize session state for converted PDFs
    if 'pdf_files' not in st.session_state:
        st.session_state.pdf_files = {}
    if 'last_upload_count' not in st.session_state:
        st.session_state.last_upload_count = 0
    
    # File uploader
    uploaded_files = st.file_uploader(
        "Choose Excel file(s)",
        type=['xlsx', 'xls', 'xlsm'],
        accept_multiple_files=True,
        help="Upload one or more Excel files to convert to PDF",
        key='file_uploader'
    )
    
    # Reset PDFs if files are removed or changed
    if uploaded_files:
        current_count = len(uploaded_files)
        if current_count != st.session_state.last_upload_count:
            st.session_state.pdf_files = {}
            st.session_state.last_upload_count = current_count
    else:
        st.session_state.pdf_files = {}
        st.session_state.last_upload_count = 0
    
    if uploaded_files:
        st.success(f"✓ {len(uploaded_files)} file(s) uploaded successfully!")
        
        # Display uploaded files
        with st.expander("📁 Uploaded Files", expanded=True):
            for file in uploaded_files:
                st.write(f"• **{file.name}** ({file.size / 1024:.2f} KB)")
        
        # Convert button - only show if not already converted
        if not st.session_state.pdf_files:
            if st.button("🔄 Convert to PDF", type="primary", use_container_width=True):
                # Progress bar
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                # Convert each file
                for idx, uploaded_file in enumerate(uploaded_files):
                    try:
                        status_text.text(f"Converting {uploaded_file.name}...")
                        
                        # Convert to PDF with caching
                        pdf_bytes = convert_excel_to_pdf_bytes(
                            uploaded_file.getvalue(),
                            uploaded_file.name
                        )
                        
                        # Store with PDF extension in session state
                        pdf_filename = os.path.splitext(uploaded_file.name)[0] + '.pdf'
                        st.session_state.pdf_files[pdf_filename] = pdf_bytes
                        
                        # Update progress
                        progress_bar.progress((idx + 1) / len(uploaded_files))
                        
                    except Exception as e:
                        st.error(f"✗ Error converting {uploaded_file.name}: {str(e)}")
                
                status_text.text("✅ Conversion complete!")
                progress_bar.empty()
                st.rerun()  # Refresh to show download buttons
        
        # Show download options if conversion is done
        if st.session_state.pdf_files:
            st.success(f"✓ Successfully converted {len(st.session_state.pdf_files)} file(s)!")
            
            st.subheader("📥 Download Options:")
            
            # Single file download
            if len(st.session_state.pdf_files) == 1:
                filename, pdf_bytes = list(st.session_state.pdf_files.items())[0]
                st.download_button(
                    label=f"⬇️ Download {filename}",
                    data=pdf_bytes,
                    file_name=filename,
                    mime="application/pdf",
                    use_container_width=True
                )
            
            # Multiple files - individual and ZIP download
            else:
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("**📄 Individual Downloads:**")
                    for filename, pdf_bytes in st.session_state.pdf_files.items():
                        st.download_button(
                            label=f"⬇️ {filename}",
                            data=pdf_bytes,
                            file_name=filename,
                            mime="application/pdf",
                            key=f"download_{filename}"
                        )
                
                with col2:
                    st.markdown("**📦 Bulk Download:**")
                    zip_bytes = create_zip_file(st.session_state.pdf_files)
                    st.download_button(
                        label="📦 Download All as ZIP",
                        data=zip_bytes,
                        file_name="converted_pdfs.zip",
                        mime="application/zip",
                        type="primary"
                    )
            
            # Button to convert new files
            if st.button("🔄 Convert New Files", use_container_width=True):
                st.session_state.pdf_files = {}
                st.rerun()
    
    else:
        st.info("👆 Upload Excel files to get started")
    
    # Footer
    st.markdown("---")
    st.markdown(
        """
        <div style='text-align: center; color: gray;'>
            <small>✨ Optimized with caching • Windows only (win32com)</small>
        </div>
        """,
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
