import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Cathode-Anode Matching App with Unused Anode Output")

uploaded_file = st.file_uploader("Upload your FC_DIRECTORY.xlsx file", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        anodes = df.iloc[:, [4, 5]].copy()
        cathodes = df.iloc[:, [7, 8]].copy()
        anodes.columns = ['Anode_Name', 'Anode_Capacity']
        cathodes.columns = ['Cathode_Name', 'Cathode_Capacity']

        anodes['Anode_Capacity'] = pd.to_numeric(anodes['Anode_Capacity'], errors='coerce')
        cathodes['Cathode_Capacity'] = pd.to_numeric(cathodes['Cathode_Capacity'], errors='coerce')
        anodes.dropna(inplace=True)
        cathodes.dropna(inplace=True)
        anodes.drop_duplicates(inplace=True)
        cathodes.drop_duplicates(inplace=True)

        used_anodes = set()
        results = []

        for _, cathode in cathodes.iterrows():
            c_name = cathode['Cathode_Name']
            c_cap = cathode['Cathode_Capacity']

            if c_cap == 0:
                results.append({
                    'Cathode_Name': c_name,
                    'Cathode_Capacity': c_cap,
                    'Anode_Name': None,
                    'Anode_Capacity': None,
                    'NP_Ratio': None,
                    'Match_Type': 'INVALID CATHODE'
                })
                continue

            in_range = None
            close_match = None
            last_resort_match = None

            for _, anode in anodes.iterrows():
                a_name = anode['Anode_Name']
                a_cap = anode['Anode_Capacity']

                if a_name in used_anodes or pd.isna(a_cap):
                    continue

                np_ratio = a_cap / c_cap

                if 1.075 <= round(np_ratio, 3) <= 1.124:
                    in_range = {
                        'Cathode_Name': c_name,
                        'Cathode_Capacity': c_cap,
                        'Anode_Name': a_name,
                        'Anode_Capacity': a_cap,
                        'NP_Ratio': round(np_ratio, 2),
                        'Match_Type': 'IN RANGE'
                    }
                    break

                elif 1.125 <= np_ratio <= 1.134 and close_match is None:
                    close_match = {
                        'Cathode_Name': c_name,
                        'Cathode_Capacity': c_cap,
                        'Anode_Name': a_name,
                        'Anode_Capacity': a_cap,
                        'NP_Ratio': round(np_ratio, 3),
                        'Match_Type': 'CLOSE'
                    }
                elif 1.065 <= np_ratio <= 1.074 and last_resort_match is None:
                    last_resort_match = {
                        'Cathode_Name': c_name,
                        'Cathode_Capacity': c_cap,
                        'Anode_Name': a_name,
                        'Anode_Capacity': a_cap,
                        'NP_Ratio': round(np_ratio, 2),
                        'Match_Type': 'LAST RESORT'
                    }

            if in_range:
                results.append(in_range)
                used_anodes.add(in_range['Anode_Name'])
            elif close_match:
                results.append(close_match)
                used_anodes.add(close_match['Anode_Name'])
            elif last_resort_match:
                results.append(last_resort_match)
                used_anodes.add(last_resort_match['Anode_Name'])
            else:
                results.append({
                    'Cathode_Name': c_name,
                    'Cathode_Capacity': c_cap,
                    'Anode_Name': None,
                    'Anode_Capacity': None,
                    'NP_Ratio': None,
                    'Match_Type': 'NO MATCH FOUND'
                })

        result_df = pd.DataFrame(results)
        st.success("Matching completed!")
        st.subheader("Matching Results")
        st.dataframe(result_df)

        # Download matching result
        csv = result_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="Download Matching Results CSV",
            data=csv,
            file_name='matching_results.csv',
            mime='text/csv'
        )

        # Prepare unused anode file:
        used_anode_names = set(result_df['Anode_Name'].dropna())
        unused_anodes = anodes[~anodes['Anode_Name'].isin(used_anode_names)].reset_index(drop=True)

        # Build same format as FC_DIRECTORY but include headers
        num_rows = len(unused_anodes)
        total_columns = 10  # assuming FC_DIRECTORY originally had 10 columns

        # Create dataframe with empty strings
        fc_directory_format = pd.DataFrame('', index=range(num_rows), columns=range(total_columns))

        # Insert anode name and capacity into columns 4 and 5 (index 3 and 4)
        fc_directory_format.iloc[:, 4] = unused_anodes['Anode_Name']
        fc_directory_format.iloc[:, 5] = unused_anodes['Anode_Capacity']

        # Now create headers
        headers = [''] * total_columns
        headers[4] = 'Anode_Name'
        headers[5] = 'Anode_Capacity'
        headers [7] = 'Cathode_Name'
        headers [8] = 'Cathode_Capacity'
        fc_directory_format.columns = headers

        # Export unused anodes as Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            fc_directory_format.to_excel(writer, index=False, sheet_name='Unused Anodes')
        output.seek(0)

        st.subheader("Unused Anodes File (With Headers)")
        st.download_button(
            label="Download Unused Anodes Excel",
            data=output,
            file_name='unused_anodes.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")