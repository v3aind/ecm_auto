import pandas as pd
import streamlit as st
from io import BytesIO

# Function to process data and create Excel file
def create_excel(input_df):
    # Get unique project names
    project_names = input_df["Project"].unique()

    # Create a dictionary to store output Excel files for each project
    output_files = {}

    for project_name in project_names:
        # Create the 'ProjectName' sheet for the current project
        project_df = input_df[input_df["Project"] == project_name][["Project"]].drop_duplicates()
        project_df.columns = ["ProjectName"]
        # Assign ProjectID from the original "Project" column
        project_df["ProjectID"] = project_df["ProjectName"]  
        project_df["Desc"] = project_df["ProjectName"]
        project_df = project_df[["ProjectID", "ProjectName", "Desc"]]

        # Create the 'PO' sheet for the current project
        po_df = input_df[input_df["Project"] == project_name][
            ["POReference", "POID", "POName", "PODesc", "Validity", "Keyword"]
        ]
        po_df.columns = ["PO Reference", "POID", "PO Name", "PO Desc", "Validity", "Keyword"]
        po_df["Claim AKTIFFI"] = ""
        po_df["Claim Attack"] = ""
        po_df["Claim Default"] = ""
        po_df["Grace Period"] = ""

        # Instead of saving to a file directly, save to a BytesIO object
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine="xlsxwriter")
        project_df.to_excel(writer, sheet_name="ProjectName", index=False)
        po_df.to_excel(writer, sheet_name="PO", index=False)
        characteristic_df = pd.DataFrame()
        characteristic_df.to_excel(writer, sheet_name="Characteristic", index=False)
        writer.close()  # Save to the BytesIO object

        # Store the BytesIO object in the output_files dictionary
        output_files[project_name] = output

    return output_files

# Streamlit app
st.title("ECM Automation input Generator")

# File uploader
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file is not None:
    # Read the uploaded file into a DataFrame
    input_df = pd.read_excel(uploaded_file)

    # Process the data and create Excel files
    output_files = create_excel(input_df)

    # Display download buttons for each output file
    for project_name, output_file in output_files.items():
        st.download_button(
            label=f"Download ECM_Structure_{project_name}.xlsx",
            data=output_file.getvalue(),
            file_name=f"ECM_Structure_{project_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
