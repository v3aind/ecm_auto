import streamlit as st
import pandas as pd

st.title("ECM Structure Generator")

# File Upload
uploaded_file = st.file_uploader("Upload 'Roaming_SC_Completion.xlsx'", type=["xlsx"])

if uploaded_file is not None:
    input_df = pd.read_excel(uploaded_file)
    project_names = input_df["Project"].unique()

    for project_name in project_names:
        # ... (code to create project_df and po_df remains the same) ...

        # Create the output file name
        output_file_name = f"ECM_Structure_{project_name}.xlsx"

        # Create an ExcelWriter object in memory
        import io
        buffer = io.BytesIO()  
        writer = pd.ExcelWriter(buffer, engine="xlsxwriter")

        # Write the DataFrames to the respective sheets
        project_df.to_excel(writer, sheet_name="ProjectName", index=False)
        po_df.to_excel(writer, sheet_name="PO", index=False)

        # Create and add the "Characteristic" sheet
        characteristic_df = pd.DataFrame()
        characteristic_df.to_excel(writer, sheet_name="Characteristic", index=False)

        # Save the Excel file to the buffer
        writer.save() 

        # Download button
        st.download_button(
            label=f"Download ECM_Structure_{project_name}.xlsx",
            data=buffer.getvalue(),
            file_name=output_file_name,
            mime="application/vnd.ms-excel"
        )

    st.success("All output files generated. Click the buttons to download.")
else:
    st.info("Please upload the 'Roaming_SC_Completion.xlsx' file to begin.")
