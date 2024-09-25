import os
import pandas as pd
import streamlit as st

# Custom CSS for background and text color
st.markdown("""
    <style>
    .stApp {
        background-color: #ADD8E6;  /* Light blue background */
    }
    .title {
        color: #003366;  /* Dark blue color for the title */
        font-family: 'Arial Black', sans-serif;
        text-align: center;
        margin-bottom: 20px;
    }
    .selectbox, .text_input, .button {
        margin-bottom: 20px;
    }
    .warning {
        color: #FF4500; /* Orange color for warnings */
    }
    </style>
    """, unsafe_allow_html=True)

def select_folder_and_search_document():
    st.markdown("<h1 class='title'>Document Details Search Interface</h1>", unsafe_allow_html=True)

    # Folder path input
    folder_path = st.text_input('Folder Path:', '')

    # Dropdown for selecting the column to search
    selected_column = st.selectbox(
        'Select Column to Search:',
        ['DocNo', 'RegistrationDate', 'SellerParty',
         'PurchaserParty', 'PropertyDescription', 'DateOfExecution'],
        index=4  # Default to 'PropertyDescription'
    )

    # Text input for the search value
    search_value = st.text_input('Enter Search Value:', '')

    # Button to trigger the search
    if st.button('Submit'):
        if not os.path.exists(folder_path):
            st.error("The specified folder does not exist. Please try again.")
            return

        # Define the two sets of expected columns
        expected_columns_set_1 = [
            'srocode', 'internaldocumentnumber', 'docno', 'docname', 'registrationdate',
            'sroname', 'micrno', 'bank_type', 'party_code', 'sellerparty',
            'purchaserparty', 'propertydescription', 'areaname', 'consideration_amt',
            'marketvalue', 'dateofexecution', 'stampdutypaid', 'registrationfees', 'status'
        ]

        expected_columns_set_2 = [
            'SROCode', 'InternalDocumentNumber', 'DocNo', 'DocName', 'RegistrationDate',
            'SROName', 'SellerParty', 'PurchaserParty', 'PropertyDescription', 'AreaName',
            'consideration_amt', 'MarketValue', 'DateOfExecution', 'StampDutyPaid',
            'RegistrationFees', 'status', 'micrno', 'party_code', 'bank_type'
        ]

        # Initialize an empty list to store dataframes
        dataframes = []

        # Load all Excel files into a list of DataFrames
        for file_name in os.listdir(folder_path):
            if file_name.endswith('.xls') or file_name.endswith('.xlsx'):
                if file_name.startswith('~$'):
                    st.warning(f"Skipping temporary file: {file_name}", icon="⚠️")
                    continue  # Skip temporary files

                file_path = os.path.join(folder_path, file_name)
                try:
                    df = pd.read_excel(file_path)

                    # Check if the DataFrame matches either set of expected columns
                    if all(col in df.columns for col in expected_columns_set_1):
                        dataframes.append(df[expected_columns_set_1])
                    elif all(col in df.columns for col in expected_columns_set_2):
                        dataframes.append(df[expected_columns_set_2])
                    else:
                        st.warning(f"File {file_name} does not match the expected column sets and will be skipped.", icon="⚠️")
                except PermissionError:
                    st.warning(f"Permission denied for file {file_name}. Ensure it is not open in another application.", icon="⚠️")
                except Exception as e:
                    st.warning(f"An error occurred with file {file_name}: {e}", icon="⚠️")

        # Initialize an empty list to store results
        results = []

        # Perform case-insensitive column name matching
        for df in dataframes:
            # Make all column names lowercase for easy matching
            df.columns = df.columns.str.lower()
            selected_column_lower = selected_column.lower()

            # Check if the selected column is present
            if selected_column_lower in df.columns:
                result = df[df[selected_column_lower].astype(str).str.contains(search_value, case=False, na=False, regex=False)]
                if not result.empty:
                    results.append(result)
            else:
                st.warning(f"Selected column '{selected_column}' not found in one or more files.", icon="⚠️")

        # Concatenate all found results and return
        if results:
            final_result = pd.concat(results).reset_index(drop=True)
            st.write(f"Results for {selected_column} containing '{search_value}':")
            st.dataframe(final_result)  # Display the result in a tabular format
        else:
            st.write(f"No results found for {selected_column} containing '{search_value}'.")

# Run the function in the Streamlit app
if __name__ == "__main__":
    select_folder_and_search_document()


# python -m streamlit run App.py