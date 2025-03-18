import streamlit as st
import pandas as pd
from io import BytesIO
import os

st.title("Excel Processor App")
st.write("Upload an Excel file and set the custom element query parameter.")

# File uploader widget for Excel files
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

# Text input for custom parameter name for the element query column
element_query_param = st.text_input("Element Query Parameter", value="Code métré")

if uploaded_file is not None:
    try:
        # Read the uploaded Excel file into a DataFrame
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
    else:
        # Check for the required column 'Hoeveelheid'
        if 'Hoeveelheid' not in df.columns:
            st.error("Error: 'Hoeveelheid' column not found in the Excel file.")
        else:
            # Filter the DataFrame: keep rows where 'Hoeveelheid' is not 0 or NaN
            filtered_df = df[df['Hoeveelheid'].notna() & (df['Hoeveelheid'] != 0)].copy()

            # Process 'Hours' column: if 0 or NaN, set to 0.1 and create 'Daily Output'
            if 'Hours' in filtered_df.columns:
                filtered_df.loc[(filtered_df['Hours'].isna()) | (filtered_df['Hours'] == 0), 'Hours'] = 0.1
                filtered_df['Daily Output'] = (8 * filtered_df['Hoeveelheid']) / filtered_df['Hours']
            else:
                st.error("Error: 'Hours' column not found in the Excel file.")

            # Map 'Meeteenheid' to 'Quantity Type'
            if 'Meeteenheid' in filtered_df.columns:
                quantity_mapping = {
                    'm3': 'Volume',
                    'm²': 'Area',
                    'm2': 'Area',
                    'st': 'Numeric',
                    's': 'Time',
                    'min': 'Time',
                    'h': 'Time',
                    'd': 'Time',
                    'kg': 'Mass',
                    'g': 'Mass',
                    't': 'Mass',
                    'cm': 'Length',
                    'mm': 'Length',
                    'km': 'Length',
                    'm': 'Length',
                    'cm²': 'Area',
                    'dm²': 'Area',
                    'km²': 'Area',
                    'cm³': 'Volume',
                    'dm³': 'Volume',
                    'km³': 'Volume',
                    'rad': 'Angle',
                    'deg': 'Angle',
                    'grad': 'Angle'
                }
                filtered_df['Quantity Type'] = (
                    filtered_df['Meeteenheid']
                    .str.strip()
                    .str.lower()
                    .map(quantity_mapping)
                    .fillna('Numeric')
                )
            else:
                st.error("Error: 'Meeteenheid' column not found in the Excel file.")

            # Add constant columns
            filtered_df['Classification Level'] = 1
            filtered_df['Outline Level'] = 1.0

            # Create 'Elemental Query' column using the custom parameter name
            if 'ID Klant' in filtered_df.columns:
                filtered_df['Elemental Query'] = filtered_df['ID Klant'].apply(
                    lambda value: f"['{element_query_param}'] = '{value}'"
                )
            else:
                st.error("Error: 'ID Klant' column not found in the Excel file.")

            # Display the processed DataFrame
            st.subheader("Processed Data Preview")
            st.dataframe(filtered_df)

            # Convert the processed DataFrame to an Excel file in memory using openpyxl
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                filtered_df.to_excel(writer, index=False, sheet_name='Filtered Data')
            processed_data = output.getvalue()

            # Create a new file name with " Processed" suffix based on the uploaded file name
            original_file_name = uploaded_file.name
            base_name, ext = os.path.splitext(original_file_name)
            new_file_name = f"{base_name} Processed{ext}"

            # Provide a download button for the processed Excel file
            st.download_button(
                label="Download Processed Excel File",
                data=processed_data,
                file_name=new_file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
