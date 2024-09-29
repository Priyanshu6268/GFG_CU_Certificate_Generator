import streamlit as st
import cv2
import pandas as pd
import zipfile
import os
import base64

# Function to convert to Pascal Case
def to_pascal_case(s):
    return ' '.join(word.capitalize() for word in s.split())

# Function to generate certificates
def generate_certificates(template_path, names_uids):
    output_files = []
    for index, (name, uid) in enumerate(names_uids):
        formatted_name = to_pascal_case(name)
        formatted_uid = uid.upper()  # Convert UID to uppercase
        text = f"{formatted_name} ({formatted_uid})"
        
        template = cv2.imread(template_path)
        cv2.putText(template, text, (254, 770), cv2.FONT_HERSHEY_SCRIPT_SIMPLEX, 2, (0,255,0), 2, cv2.LINE_AA)
        output_path = f'Certificate_{formatted_name}.jpg'
        cv2.imwrite(output_path, template)
        output_files.append(output_path)  # Keep track of generated files
        print(f'Processing Certificate {index + 1}/{len(names_uids)}')
    print('All Certificates are ready')
    return output_files

# Function to create a ZIP file
def create_zip_file(file_list, zip_filename='certificates.zip'):
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        for file in file_list:
            zipf.write(file)
    return zip_filename

# Streamlit interface
st.set_page_config(page_title="Certificate Generator")

# Add custom CSS for logo positioning
# Function to load and encode the logo
def load_image(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

# Add custom CSS for logo positioning
st.markdown(
    """
    <style>
    .logo {
        position: absolute;
        top: -20px;
        left: -200px;
        width: 150px; /* Adjust size as needed */
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Load and display the logo
logo_path = "GFG CU logo.png"  # Ensure this file is in the same directory
logo_base64 = load_image(logo_path)
st.markdown(f'<img class="logo" src="data:image/png;base64,{logo_base64}" alt="GFG CU Logo">', unsafe_allow_html=True)

st.title("GFG CU Certificate Generator")

# File upload for the certificate template
template_file = st.file_uploader("Upload Certificate Template ", type=["jpg", "jpeg","png"])
if template_file is not None:
    # Display the uploaded template
    st.image(template_file, caption="Uploaded Certificate Template", use_column_width=True)
    
    # File upload for the Excel file
    excel_file = st.file_uploader("Upload Excel File with Names and UIDs", type=["xlsx"])
    if excel_file is not None:
        # Read the Excel file
        df = pd.read_excel(excel_file)
        
        # Check if required columns are present
        if 'Name' in df.columns and 'UID' in df.columns:
            names_uids = df[['Name', 'UID']].values.tolist()  # Create a list of tuples (Name, UID)

            # Process certificates
            if st.button("Generate Certificates"):
                # Save the template file to a temporary location
                template_path = 'temp_template.jpg'
                with open(template_path, "wb") as f:
                    f.write(template_file.getbuffer())
                
                # Generate the certificates
                generated_files = generate_certificates(template_path, names_uids)
                
                # Create a ZIP file
                zip_filename = create_zip_file(generated_files)
                
                # Provide download link for the ZIP file
                with open(zip_filename, 'rb') as f:
                    st.download_button(
                        label="Download All Certificates as ZIP",
                        data=f,
                        file_name=zip_filename,
                        mime='application/zip'
                    )
                
                st.success("All Certificates are generated and ready for download!")

                # Clean up: remove generated files if needed
                for file in generated_files:
                    os.remove(file)  # Optionally remove the files after zipping
                os.remove(template_path)  # Clean up the temporary template file
        else:
            st.error("Excel file must contain 'Name' and 'UID' columns.")
# Footer
# Footer
st.markdown("<hr>", unsafe_allow_html=True)
st.markdown(
    'Designed and developed with ❤️ by [**Priyanshu Kumar Saw**](https://linktr.ee/priyanshu_kumar_saw)', 
    unsafe_allow_html=True
)