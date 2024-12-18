import streamlit as st
import pandas as pd
import re
import io

st.title("Excel Data Processor")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names

        output_buffer = io.BytesIO()  # Use BytesIO for in-memory Excel file

        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            for sheet_name in sheet_names:
                try:
                    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

                    if " Name" not in df.columns:
                        st.warning(f"Skipping sheet '{sheet_name}' as it doesn't contain ' Name' column.")
                        continue

                    data_to_save = []

                    for index, row in df.iterrows():
                        matches = re.findall(r"=(.*?),", str(row[' Name']))
                        if matches:
                            element_name = matches[0]
                            area = row.get('Area', 0.0)
                            volume = row.get('Volume', 0.0)
                            data_to_save.append({'Element Name': element_name, 'Area': area, 'Volume': volume, 'count': 1.0})
                        else:
                            st.warning(f"No match found in row {index} of sheet {sheet_name}")

                    if data_to_save:
                        dftosave = pd.DataFrame(data_to_save)
                        element_sums = dftosave.groupby('Element Name').sum(numeric_only=True).reset_index()
                        element_sums.to_excel(writer, sheet_name=sheet_name, index=False)
                    else:
                        st.warning(f"No data to save for sheet {sheet_name}")

                except Exception as inner_e:
                    st.error(f"An error occurred while processing sheet '{sheet_name}': {inner_e}")

        st.success("Processing complete!")

        # Download button
        st.download_button(
            label="Download Processed Excel",
            data=output_buffer.getvalue(),
            file_name="processed_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")