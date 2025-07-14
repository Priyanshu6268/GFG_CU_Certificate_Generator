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
def generate_certificates(template_path, names_uids, include_uid=True):
    output_files = []
    for index, item in enumerate(names_uids):
        name = item[0]
        uid = item[1] if include_uid and len(item) > 1 else ""

        formatted_name = to_pascal_case(name)
        formatted_uid = uid.upper() if uid else ""
        text = f"{formatted_name} ({formatted_uid})" if formatted_uid else formatted_name

        template = cv2.imread(template_path)
        cv2.putText(template, text, (254, 770), cv2.FONT_HERSHEY_SCRIPT_SIMPLEX, 2, (0, 255, 0), 2, cv2.LINE_AA)
        output_path = f'Certificate_{formatted_name}.jpg'
        cv2.imwrite(output_path, template)
        output_files.append(output_path)
        print(f'Processing Certificate {index + 1}/{len(names_uids)}')
    print('All Certificates are ready')
    return output_files

# Function to create a ZIP file
def create_zip_file(file_list, zip_filename='certificates.zip'):
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        for file in file_list:
            zipf.write(file)
    return zip_filename

# Function to load and encode the logo
def load_image(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

# Streamlit interface
st.set_page_config(page_title="Certificate Generator")

# Custom CSS for logo
st.markdown(
    """
    <style>
    .logo {
        position: absolute;
        top: -20px;
        left: -200px;
        width: 150px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Load and display the logo
logo_path = "GFG CU logo.png"  # Ensure this file is in the same directory
if os.path.exists(logo_path):
    logo_base64 = load_image(logo_path)
    st.markdown(f'<img class="logo" src="data:image/png;base64,{logo_base64}" alt="GFG CU Logo">', unsafe_allow_html=True)

st.title("GFG CU Certificate Generator")

# Upload certificate template
template_file = st.file_uploader("Upload Certificate Template ", type=["jpg", "jpeg", "png"])
if template_file is not None:
    st.image(template_file, caption="Uploaded Certificate Template", use_column_width=True)

    # Upload Excel file
    excel_file = st.file_uploader("Upload Excel File with Names (and optional UIDs)", type=["xlsx"])
    if excel_file is not None:
        df = pd.read_excel(excel_file)

        # Check for 'Name' column
        if 'Name' not in df.columns:
            st.error("Excel file must contain at least a 'Name' column.")
        else:
            # Determine if 'UID' is present
            include_uid = 'UID' in df.columns
            names_uids = df[['Name', 'UID']].values.tolist() if include_uid else df[['Name']].values.tolist()

            if st.button("Generate Certificates"):
                # Save uploaded template temporarily
                template_path = 'temp_template.jpg'
                with open(template_path, "wb") as f:
                    f.write(template_file.getbuffer())

                # Generate and zip certificates
                generated_files = generate_certificates(template_path, names_uids, include_uid=include_uid)
                zip_filename = create_zip_file(generated_files)

                # Download button
                with open(zip_filename, 'rb') as f:
                    st.download_button(
                        label="Download All Certificates as ZIP",
                        data=f,
                        file_name=zip_filename,
                        mime='application/zip'
                    )

                st.success("All Certificates are generated and ready for download!")

                # Clean up
                for file in generated_files:
                    os.remove(file)
                os.remove(template_path)

# Footer
st.markdown("<hr>", unsafe_allow_html=True)
st.markdown(
    'Designed and developed with ❤️ by [**Priyanshu Kumar Saw**](https://linktr.ee/priyanshu_kumar_saw)', 
    unsafe_allow_html=True
)
