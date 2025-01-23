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
            
	# Iterate through each project name
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
                    writer.close()
                
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
