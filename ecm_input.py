import streamlit as st
import pandas as pd
import io

st.title("ECM Structure Generator")

# File Upload
uploaded_file = st.file_uploader("Upload 'Roaming_SC_Completion.xlsx'", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Read the uploaded Excel file
        input_df = pd.read_excel(uploaded_file)
        
        # Ensure the required column exists
        if "Project" not in input_df.columns:
            st.error("The uploaded file must contain a 'Project' column.")
        else:
            # Get unique project names
            project_names = input_df["Project"].unique()
            
            for project_name in project_names:
                # Filter the data for the current project
                project_df = input_df[input_df["Project"] == project_name]
                
                # Example placeholder for `po_df`
                # Replace this with the actual logic for creating `po_df`
                po_df = pd.DataFrame({
                    "PO Number": ["12345", "67890"],
                    "PO Description": ["Description A", "Description B"]
                })
                
                # Create the output file name
                output_file_name = f"ECM_Structure_{project_name}.xlsx"
                
                # Create an in-memory buffer for the Excel file
                buffer = io.BytesIO()
                
                # Create the ExcelWriter object
                with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                    # Write the DataFrames to their respective sheets
                    project_df.to_excel(writer, sheet_name="ProjectName", index=False)
                    po_df.to_excel(writer, sheet_name="PO", index=False)
                    
                    # Create and add the "Characteristic" sheet (empty for now)
                    characteristic_df = pd.DataFrame()
                    characteristic_df.to_excel(writer, sheet_name="Characteristic", index=False)
                    
                    # Save the writer contents to the buffer
                    writer.save()
                
                # Add a download button for the generated file
                st.download_button(
                    label=f"Download ECM_Structure_{project_name}.xlsx",
                    data=buffer.getvalue(),
                    file_name=output_file_name,
                    mime="application/vnd.ms-excel"
                )
            
            st.success("All output files generated. Click the buttons to download.")
    except Exception as e:
        st.error(f"An error occurred: {e}")
else:
    st.info("Please upload the 'Roaming_SC_Completion.xlsx' file to begin.")
