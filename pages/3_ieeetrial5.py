import streamlit as st
import pandas as pd
import numpy as np


# App Title
st.set_page_config(page_title="IEEE Yield Analysis!!!", page_icon=":bar_chart:", layout="wide")
st.title("IEEE Yield Analysis")


# Defect distribution Table
st.subheader("Defect distribution per board")
defects = list(range(16))  # Defect distribution per board from 0 to 15
no_of_boards = [
    300, 0, 150, 75, 37.5, 30, 24, 19.2, 15.36, 12.288,
    9.8304, 7.86432, 6.291456, 5.0331648, 4.02653184, 2
]


# Creating the DataFrame
defects_df = pd.DataFrame({
    'No. of Defects': defects,
    'No. Of Boards': no_of_boards
})


defects_tr_df = defects_df.transpose()
st.dataframe(defects_tr_df)  # Display the transposed defects table


st.subheader("Alpha Estimation")


# First row: col1 to col4 under defects_tr_df
col1, col2, col3, col4 = st.columns([1, 1, 1, 1])


with col1:
    # Calculating the average defects per board
    mean_μ = sum(no_of_boards) / len(no_of_boards)
    st.text_input("Average Number of Boards", mean_μ)


with col2:
    # Calculating the std deviation defects per board
    std_deviation_σ = pd.Series(no_of_boards).std()
    st.text_input("Standard Deviation", std_deviation_σ)


with col3:
    # Input for Alpha (α) Assumed
    alpha_α_assumed = st.text_input("Alpha (α) Assumed", value="0.4")


with col4:
    # Calculating Alpha (α)
    alpha_α = (mean_μ ** 2) / ((std_deviation_σ ** 2) - mean_μ)
    st.text_input("Alpha (α)", alpha_α)


# File uploader for Excel/CSV/XLSM files
uploaded_file = st.file_uploader("Choose an Excel/CSV/XLSM file", type=["xlsx", "csv", "xlsm"])


if uploaded_file:
    # Load data from the uploaded file
    @st.cache_data
    def load_data(file):
        if file.name.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file, sheet_name=None)
        return df


    data = load_data(uploaded_file)


    # Initialize session state to store edited data for each sheet
    if 'edited_sheets' not in st.session_state:
        st.session_state.edited_sheets = {}


    # Second row: col5 to col10 under data = load_data(uploaded_file)
    col5, col6, col7, col8 = st.columns([1, 1, 1, 1])
    col9, col10, col11, col12 = st.columns([1, 1, 1, 1])
    col14 = st.columns([1])[0]

    if isinstance(data, dict):
        with col5:
            sheet_name = st.selectbox("Select the sheet", data.keys())


        # Check if the sheet has been edited before; if so, load the edited version
        if sheet_name in st.session_state.edited_sheets:
            st.session_state.df = st.session_state.edited_sheets[sheet_name]
        else:
            selected_data = data[sheet_name]
            st.session_state.df = pd.DataFrame(selected_data)  # Load original data from file


        # 1. Provide an option to select the product development stage
        with col6:            
            stages = ['MK0', 'MK1', 'MK2', 'MK3', 'X1']  # Add more stages if needed
            selected_stage = st.selectbox("Select the product development stage", stages)


        # Display data in a table
        st.subheader("Data Table")
        edited_data = st.data_editor(st.session_state.df)


        # 2. Provide an option to input "Test Efficiency %"
        with col7:
            test_efficiency = st.text_input("Enter Test Efficiency %", value="")
        if test_efficiency:
            try:
                test_efficiency_value = float(test_efficiency) / 100  # Convert percentage to decimal
                
                # Update existing rows for "Test Efficiency %"
                test_efficiency_row_idx = edited_data[edited_data['Data Points'] == "Test Efficiency %"].index

                # Insert "Test Efficiency %" value in the same row under the selected stage
                if not test_efficiency_row_idx.empty:
                    edited_data.at[test_efficiency_row_idx[0], selected_stage] = test_efficiency_value
                else:
                    new_row_te = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                    new_row_te.at[0, 'Data Points'] = "Test Efficiency %"
                    new_row_te.at[0, selected_stage] = test_efficiency_value
                    st.session_state.df = pd.concat([st.session_state.df, new_row_te], ignore_index=True)

                st.session_state.edited_sheets[sheet_name] = st.session_state.df  # Update session state with new rows

            except ValueError:
                st.error("Please enter a valid number for Test Efficiency %")

        # 2. Provide an option to input "No. Solder Joints (N)"
        with col8:
            solder_joint_value = st.text_input("No. Solder Joints (N)", value="")
        if test_efficiency:
            try:
                # Update existing rows for "No. Solder Joints (N)"
                solder_joint_value_row_idx = edited_data[edited_data['Data Points'] == "No. Solder Joints (N)"].index

                # Insert "No. Solder Joints (N)" value in the same row under the selected stage
                if not solder_joint_value_row_idx.empty:
                    edited_data.at[solder_joint_value_row_idx[0], selected_stage] = solder_joint_value
                else:
                    new_row_solder_joint = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                    new_row_solder_joint.at[0, 'Data Points'] = "No. Solder Joints (N)"
                    new_row_solder_joint.at[0, selected_stage] = solder_joint_value
                    test_efficiency_row_idx = edited_data[edited_data['Data Points'] == "Test Efficiency %"].index[0]
                    edited_data = pd.concat([edited_data.iloc[:test_efficiency_row_idx+1], new_row_solder_joint, edited_data.iloc[test_efficiency_row_idx+1:]]).reset_index(drop=True)

                st.session_state.edited_sheets[sheet_name] = st.session_state.df  # Update session state with new rows

            except ValueError:
                st.error("Please enter a valid number for Test Efficiency %")

        with col9:
            # Input for "No. Component"
            no_component_value = st.text_input("No. Component", value="")
            if no_component_value:
                try:
                    # Update existing rows for "No. Component"
                    no_component_row_idx = edited_data[edited_data['Data Points'] == "No. Component"].index


                    # Insert "No. Component" value in the same row under the selected stage
                    if not no_component_row_idx.empty:
                        edited_data.at[no_component_row_idx[0], selected_stage] = no_component_value
                    else:
                        new_row_no_component = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                        new_row_no_component.at[0, 'Data Points'] = "No. Component"
                        new_row_no_component.at[0, selected_stage] = no_component_value
                        solder_joint_value_row_idx = edited_data[edited_data['Data Points'] == "No. Solder Joints (N)"].index
                        edited_data = pd.concat([edited_data.iloc[:solder_joint_value_row_idx+1], new_row_no_component, edited_data.iloc[solder_joint_value_row_idx+1:]]).reset_index(drop=True)


                    st.session_state.edited_sheets[sheet_name] = st.session_state.df  # Update session state with new rows


                except ValueError:
                    st.error("Please enter a valid number for No. Component")


        with col10:
            # Input for "No. Placement"
            no_placement_value = st.text_input("No. Placement", value="")
            if no_placement_value:
                try:
                    # Update existing rows for "No. Placement"
                    no_placement_row_idx = edited_data[edited_data['Data Points'] == "No. Placement"].index


                    # Insert "No. Placement" value in the same row under the selected stage
                    if not no_placement_row_idx.empty:
                        edited_data.at[no_placement_row_idx[0], selected_stage] = no_placement_value
                    else:
                        new_row_no_placement = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                        new_row_no_placement.at[0, 'Data Points'] = "No. Placement"
                        new_row_no_placement.at[0, selected_stage] = no_placement_value
                        no_component_row_idx = edited_data[edited_data['Data Points'] == "No. Component"].index
                        edited_data = pd.concat([edited_data.iloc[:no_component_row_idx+1], new_row_no_placement, edited_data.iloc[no_component_row_idx+1:]]).reset_index(drop=True)


                    st.session_state.edited_sheets[sheet_name] = st.session_state.df  # Update session state with new rows


                except ValueError:
                    st.error("Please enter a valid number for No. Placement")

        with col11:
            defect_rate = st.text_input("Defect Rate per Solder Joint (DR)", value="")
        if defect_rate:
            try:
                defect_rate_value = float(defect_rate) / 1000000  # Covert Defects per million
               
                # Update existing rows for "Test Test Effectiveness(TE) %"
                defect_rate_row_idx = edited_data[edited_data['Data Points'] == "Defect Rate per Solder Joint (DR)"].index


                # Insert "Test Effectiveness(TE) %" value in the same row under the selected stage
                if not defect_rate_row_idx.empty:
                    edited_data.at[defect_rate_row_idx[0], selected_stage] = defect_rate_value
                else:
                    new_row_dr = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                    new_row_dr.at[0, 'Data Points'] = "Defect Rate per Solder Joint (DR)"
                    new_row_dr.at[0, selected_stage] = defect_rate_value
                    st.session_state.df = pd.concat([st.session_state.df, new_row_dr], ignore_index=True)


                st.session_state.edited_sheets[sheet_name] = st.session_state.df  # Update session state with new rows


            except ValueError:
                st.error("Please enter a valid number for Defect Rate per Solder Joint (DR)")

        with col14:
        # Overall yield_Soldering Calculations
        # Button to trigger calculation of Pfi, Pfo, D, Ync, Ycl and Overall yield
            if st.button("Analyse Pfi, Pfo, D, Ync, Ycl and Overall yield"):
                # Ensure both defect_rate_value and solder_joint_value are valid numbers
                try:
                    defect_rate_value = float(defect_rate) / 1000000  # Convert to defect rate per million
                    solder_joint_value = float(solder_joint_value)    # Convert solder_joint_value to float

                    # Calculate pfi_value using the given formula
                    pfi_value = 1 - (1 - defect_rate_value) ** solder_joint_value   # Formula: 1-(1-(500/1000000))^F7

                    # Update or insert "Pfi" value in the DataFrame for the selected stage
                    pfi_value_row_idx = edited_data[edited_data['Data Points'] == "Pfi"].index

                    if not pfi_value_row_idx.empty:
                        # If "Pfi" already exists, update its value
                        edited_data.at[pfi_value_row_idx[0], selected_stage] = pfi_value
                    else:
                        # If "Pfi" doesn't exist, create a new row for it
                        new_row_pfi = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                        new_row_pfi.at[0, 'Data Points'] = "Pfi"
                        new_row_pfi.at[0, selected_stage] = pfi_value
                        edited_data = pd.concat([edited_data, new_row_pfi], ignore_index=True)

                    # Update session state with the modified DataFrame
                    st.session_state.edited_sheets[sheet_name] = edited_data

                    # Optionally show a success message after calculation
                    # st.success(f"Pfi value calculated and updated: {pfi_value}")

                except ValueError:
                    # Show error if the inputs are not valid numbers
                    st.error("Please enter valid numbers for Defect Rate and Solder Joints")

                #"Pfo" calculation
                try:
                    test_efficiency_value = float(test_efficiency) / 100  # Convert percentage to decimal
                    pfi_value = float(pfi_value)    # Convert pfi_value to float

                    # Calculate Pfo using the given formula
                    pfo_value = 1 - (1 - pfi_value) ** (1 - test_efficiency_value)   # Formula: =1-((1-K7)^(1-$F$3))

                    # Update or insert "Pfo" value in the DataFrame for the selected stage
                    pfo_value_row_idx = edited_data[edited_data['Data Points'] == "Pfo"].index

                    if not pfo_value_row_idx.empty:
                        # If "Pfo" already exists, update its value
                        edited_data.at[pfo_value_row_idx[0], selected_stage] = pfo_value
                    else:
                        # If "Pfo" doesn't exist, create a new row for it
                        new_row_pfo = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                        new_row_pfo.at[0, 'Data Points'] = "Pfo"
                        new_row_pfo.at[0, selected_stage] = pfo_value
                        edited_data = pd.concat([edited_data, new_row_pfo], ignore_index=True)

                    # Update session state with the modified DataFrame
                    st.session_state.edited_sheets[sheet_name] = edited_data

                    # Optionally show a success message after calculation
                    # st.success(f"Pfo value calculated and updated: {pfo_value}")

                except ValueError:
                    # Show error if the inputs are not valid numbers
                    st.error("Please enter valid numbers for Pfi and Test Efficiency")

                #"D" calculation
                try:
                    pfi_value = float(pfi_value)    # Convert pfi_value to float
                    pfo_value = float(pfo_value)    # Convert pfo_value to float
                    test_efficiency_value = float(test_efficiency) / 100  # Convert percentage to decimal

                    # Calculate D using the given formula
                    avg_defects_per_brd_value = (pfi_value + pfo_value) * test_efficiency_value   # Formula: =SUM(K7:L7)*$F$3

                    # Update or insert "D" value in the DataFrame for the selected stage
                    avg_defects_per_brd_value_row_idx = edited_data[edited_data['Data Points'] == "D"].index

                    if not avg_defects_per_brd_value_row_idx.empty:
                        # If "Pfo" already exists, update its value
                        edited_data.at[avg_defects_per_brd_value_row_idx[0], selected_stage] = avg_defects_per_brd_value
                    else:
                        # If "D" doesn't exist, create a new row for it
                        new_row_avg_d = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                        new_row_avg_d.at[0, 'Data Points'] = "D"
                        new_row_avg_d.at[0, selected_stage] = avg_defects_per_brd_value
                        edited_data = pd.concat([edited_data, new_row_avg_d], ignore_index=True)

                    # Update session state with the modified DataFrame
                    st.session_state.edited_sheets[sheet_name] = edited_data

                    # Optionally show a success message after calculation
                    # st.success(f"Avg D value calculated and updated: {avg_defects_per_brd_value}")

                except ValueError:
                    # Show error if the inputs are not valid numbers
                    st.error("Please enter valid numbers for Pfi, Pfo and Test Efficiency")

                #"Poisson Yield (Ync)" calculation
                try:
                    avg_defects_per_brd_value = float(avg_defects_per_brd_value)    # Convert avg_defects_per_brd_value to float
                    # Calculate Poisson Yield (Ync) using the given formula
                    non_clustered_yield_value = np.exp(-avg_defects_per_brd_value)   # Formula: =EXP(-M7)

                    # Update or insert "Poisson Yield (Ync)" value in the DataFrame for the selected stage
                    non_clustered_yield_value_row_idx = edited_data[edited_data['Data Points'] == "Poisson Yield (Ync)"].index

                    if not non_clustered_yield_value_row_idx.empty:
                        # If "Poisson Yield (Ync)" already exists, update its value
                        edited_data.at[non_clustered_yield_value_row_idx[0], selected_stage] = non_clustered_yield_value
                    else:
                        # If "Poisson Yield (Ync)" doesn't exist, create a new row for it
                        new_row_ync = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                        new_row_ync.at[0, 'Data Points'] = "Poisson Yield (Ync)"
                        new_row_ync.at[0, selected_stage] = non_clustered_yield_value
                        edited_data = pd.concat([edited_data, new_row_ync], ignore_index=True)

                    # Update session state with the modified DataFrame
                    st.session_state.edited_sheets[sheet_name] = edited_data

                    # Optionally show a success message after calculation
                    # st.success(f"Ync value calculated and updated: {non_clustered_yield_value}")

                except ValueError:
                    # Show error if the inputs are not valid numbers
                    st.error("Please enter valid numbers for Avg D")

                #"Poisson Yield (Ycl)" calculation
                try:
                    avg_defects_per_brd_value = float(avg_defects_per_brd_value)    # Convert avg_defects_per_brd_value to float
                    alpha_α_assumed = float(alpha_α_assumed)  # Convert alpha_α_assumed to float

                    # Calculate Poisson Yield (Ycl) using the given formula
                    clustered_yield_value = (1 + (avg_defects_per_brd_value / alpha_α_assumed)) ** (-alpha_α_assumed)   # Formula: =(1+(M7/$C$27))^-$C$27
                    # clustered_yield_value = (1 + (avg_defects_per_brd_value / 0.4)) ** (-0.4)   # Formula: =(1+(M7/$C$27))^-$C$27
                    # Update or insert "Poisson Yield (Ycl)" value in the DataFrame for the selected stage
                    clustered_yield_value_row_idx = edited_data[edited_data['Data Points'] == "Clustered Yield (Ycl)"].index

                    if not clustered_yield_value_row_idx.empty:
                        # If "Clustered Yield (Ycl)" already exists, update its value
                        edited_data.at[clustered_yield_value_row_idx[0], selected_stage] = clustered_yield_value
                    else:
                        # If "Clustered Yield (Ycl)" doesn't exist, create a new row for it
                        new_row_ycl = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                        new_row_ycl.at[0, 'Data Points'] = "Clustered Yield (Ycl)"
                        new_row_ycl.at[0, selected_stage] = clustered_yield_value
                        edited_data = pd.concat([edited_data, new_row_ycl], ignore_index=True)

                    # Update session state with the modified DataFrame
                    st.session_state.edited_sheets[sheet_name] = edited_data

                    # Optionally show a success message after calculation
                    # st.success(f"Ycl value calculated and updated: {clustered_yield_value}")

                except ValueError:
                    # Show error if the inputs are not valid numbers
                    st.error("Please enter valid numbers for Avg D & mean_μ")

                #"Overall yield_Soldering" calculation
                try:
                    non_clustered_yield_value = float(non_clustered_yield_value)   # Convert non_clustered_yield_value to float
                    clustered_yield_value = float(clustered_yield_value)   # Convert clustered_yield_value to float

                    # Calculate Overall yield_Soldering using the given formula
                    overall_yield_s_value = non_clustered_yield_value * clustered_yield_value  # Formula: =H7*I7

                    # Update or insert "Overall yield_Soldering" value in the DataFrame for the selected stage
                    overall_yield_s_value_row_idx = edited_data[edited_data['Data Points'] == "Overall yield_Soldering"].index

                    if not overall_yield_s_value_row_idx.empty:
                        # If "Overall yield_Soldering" already exists, update its value
                        edited_data.at[overall_yield_s_value_row_idx[0], selected_stage] = overall_yield_s_value
                    else:
                        # If "Overall yield_Soldering" doesn't exist, create a new row for it
                        new_row_oy = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                        new_row_oy.at[0, 'Data Points'] = "Overall yield_Soldering"
                        new_row_oy.at[0, selected_stage] = overall_yield_s_value
                        edited_data = pd.concat([edited_data, new_row_oy], ignore_index=True)

                    # Update session state with the modified DataFrame
                    st.session_state.edited_sheets[sheet_name] = edited_data

                    # Optionally show a success message after calculation
                    # st.success(f"Overall yield_Soldering value calculated and updated: {overall_yield_s_value}")

                except ValueError:
                    # Show error if the inputs are not valid numbers
                    st.error("Please enter valid numbers for Ync & Ycl")
# *****************************************************************************************************************
# *****************************************************************************************************************
# *****************************************************************************************************************
# *****************************************************************************************************************
# *****************************************************************************************************************
# *****************************************************************************************************************

        # Overall yield_Component Calculations
        # Button to trigger calculation of Pfi, Pfo, D, Ync, Ycl and Overall yield            
                try:
                    defect_rate_value = float(defect_rate) / 1000000  # Convert to defect rate per million
                    no_placement_value = float(no_placement_value)    # Convert solder_joint_value to float

                    # Calculate pfi_value using the given formula
                    pfi_value = 1 - (1 - defect_rate_value) ** no_placement_value   # Formula: =1-(1-(500/1000000))^F25

                    # Update or insert "Pfi" value in the DataFrame for the selected stage
                    pfi_value_row_idx = edited_data[edited_data['Data Points'] == "Pfi"].index

                    if not pfi_value_row_idx.empty:
                        # If "Pfi" already exists, update its value
                        edited_data.at[pfi_value_row_idx[0], selected_stage] = pfi_value
                    else:
                        # If "Pfi" doesn't exist, create a new row for it
                        new_row_pfi = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                        new_row_pfi.at[0, 'Data Points'] = "Pfi"
                        new_row_pfi.at[0, selected_stage] = pfi_value
                        edited_data = pd.concat([edited_data, new_row_pfi], ignore_index=True)

                    # Update session state with the modified DataFrame
                    st.session_state.edited_sheets[sheet_name] = edited_data

                    # Optionally show a success message after calculation
                    # st.success(f"Pfi value calculated and updated: {pfi_value}")

                except ValueError:
                    # Show error if the inputs are not valid numbers
                    st.error("Please enter valid numbers for Defect Rate and Solder Joints")

                #"Pfo" calculation
                try:
                    test_efficiency_value = float(test_efficiency) / 100  # Convert percentage to decimal
                    pfi_value = float(pfi_value)    # Convert pfi_value to float

                    # Calculate Pfo using the given formula
                    pfo_value = 1 - (1 - pfi_value) ** (1 - test_efficiency_value)   # Formula: =1-((1-K25)^(1-$F$3))

                    # Update or insert "Pfo" value in the DataFrame for the selected stage
                    pfo_value_row_idx = edited_data[edited_data['Data Points'] == "Pfo"].index

                    if not pfo_value_row_idx.empty:
                        # If "Pfo" already exists, update its value
                        edited_data.at[pfo_value_row_idx[0], selected_stage] = pfo_value
                    else:
                        # If "Pfo" doesn't exist, create a new row for it
                        new_row_pfo = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                        new_row_pfo.at[0, 'Data Points'] = "Pfo"
                        new_row_pfo.at[0, selected_stage] = pfo_value
                        edited_data = pd.concat([edited_data, new_row_pfo], ignore_index=True)

                    # Update session state with the modified DataFrame
                    st.session_state.edited_sheets[sheet_name] = edited_data

                    # Optionally show a success message after calculation
                    # st.success(f"Pfo value calculated and updated: {pfo_value}")

                except ValueError:
                    # Show error if the inputs are not valid numbers
                    st.error("Please enter valid numbers for Pfi and Test Efficiency")

                #"D" calculation
                try:
                    pfi_value = float(pfi_value)    # Convert pfi_value to float
                    pfo_value = float(pfo_value)    # Convert pfo_value to float
                    test_efficiency_value = float(test_efficiency) / 100  # Convert percentage to decimal

                    # Calculate D using the given formula
                    avg_defects_per_brd_value = (pfi_value + pfo_value) * test_efficiency_value   # Formula: =SUM(K25:L25)*$F$3

                    # Update or insert "D" value in the DataFrame for the selected stage
                    avg_defects_per_brd_value_row_idx = edited_data[edited_data['Data Points'] == "D"].index

                    if not avg_defects_per_brd_value_row_idx.empty:
                        # If "Pfo" already exists, update its value
                        edited_data.at[avg_defects_per_brd_value_row_idx[0], selected_stage] = avg_defects_per_brd_value
                    else:
                        # If "D" doesn't exist, create a new row for it
                        new_row_avg_d = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                        new_row_avg_d.at[0, 'Data Points'] = "D"
                        new_row_avg_d.at[0, selected_stage] = avg_defects_per_brd_value
                        edited_data = pd.concat([edited_data, new_row_avg_d], ignore_index=True)

                    # Update session state with the modified DataFrame
                    st.session_state.edited_sheets[sheet_name] = edited_data

                    # Optionally show a success message after calculation
                    # st.success(f"Avg D value calculated and updated: {avg_defects_per_brd_value}")

                except ValueError:
                    # Show error if the inputs are not valid numbers
                    st.error("Please enter valid numbers for Pfi, Pfo and Test Efficiency")

                #"Poisson Yield (Ync)" calculation
                try:
                    avg_defects_per_brd_value = float(avg_defects_per_brd_value)    # Convert avg_defects_per_brd_value to float
                    # Calculate Poisson Yield (Ync) using the given formula
                    non_clustered_yield_value = np.exp(-avg_defects_per_brd_value)   # Formula: =EXP(-M25)

                    # Update or insert "Poisson Yield (Ync)" value in the DataFrame for the selected stage
                    non_clustered_yield_value_row_idx = edited_data[edited_data['Data Points'] == "Poisson Yield (Ync)"].index

                    if not non_clustered_yield_value_row_idx.empty:
                        # If "Poisson Yield (Ync)" already exists, update its value
                        edited_data.at[non_clustered_yield_value_row_idx[0], selected_stage] = non_clustered_yield_value
                    else:
                        # If "Poisson Yield (Ync)" doesn't exist, create a new row for it
                        new_row_ync = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                        new_row_ync.at[0, 'Data Points'] = "Poisson Yield (Ync)"
                        new_row_ync.at[0, selected_stage] = non_clustered_yield_value
                        edited_data = pd.concat([edited_data, new_row_ync], ignore_index=True)

                    # Update session state with the modified DataFrame
                    st.session_state.edited_sheets[sheet_name] = edited_data

                    # Optionally show a success message after calculation
                    # st.success(f"Ync value calculated and updated: {non_clustered_yield_value}")

                except ValueError:
                    # Show error if the inputs are not valid numbers
                    st.error("Please enter valid numbers for Avg D")

                #"Poisson Yield (Ycl)" calculation
                try:
                    avg_defects_per_brd_value = float(avg_defects_per_brd_value)    # Convert avg_defects_per_brd_value to float
                    alpha_α_assumed = float(alpha_α_assumed)  # Convert alpha_α_assumed to float

                    # Calculate Poisson Yield (Ycl) using the given formula
                    clustered_yield_value = (1 + (avg_defects_per_brd_value / alpha_α_assumed)) ** (-alpha_α_assumed)   # Formula: =(1+(M25/$C$27))^-$C$27
                    
                    # Update or insert "Poisson Yield (Ycl)" value in the DataFrame for the selected stage
                    clustered_yield_value_row_idx = edited_data[edited_data['Data Points'] == "Clustered Yield (Ycl)"].index

                    if not clustered_yield_value_row_idx.empty:
                        # If "Clustered Yield (Ycl)" already exists, update its value
                        edited_data.at[clustered_yield_value_row_idx[0], selected_stage] = clustered_yield_value
                    else:
                        # If "Clustered Yield (Ycl)" doesn't exist, create a new row for it
                        new_row_ycl = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                        new_row_ycl.at[0, 'Data Points'] = "Clustered Yield (Ycl)"
                        new_row_ycl.at[0, selected_stage] = clustered_yield_value
                        edited_data = pd.concat([edited_data, new_row_ycl], ignore_index=True)

                    # Update session state with the modified DataFrame
                    st.session_state.edited_sheets[sheet_name] = edited_data

                    # Optionally show a success message after calculation
                    # st.success(f"Ycl value calculated and updated: {clustered_yield_value}")

                except ValueError:
                    # Show error if the inputs are not valid numbers
                    st.error("Please enter valid numbers for Avg D & mean_μ")

                #"Overall Overall yield_Placement" calculation
                try:
                    non_clustered_yield_value = float(non_clustered_yield_value)   # Convert non_clustered_yield_value to float
                    clustered_yield_value = float(clustered_yield_value)   # Convert clustered_yield_value to float

                    # Calculate Overall yield_Placement using the given formula
                    overall_yield_p_value = non_clustered_yield_value * clustered_yield_value  # Formula: =H25*I25

                    # Update or insert "Overall yield_Placement" value in the DataFrame for the selected stage
                    overall_yield_p_value_row_idx = edited_data[edited_data['Data Points'] == "Overall yield_Placement"].index

                    if not overall_yield_p_value_row_idx.empty:
                        # If "Overall yield_Placement" already exists, update its value
                        edited_data.at[overall_yield_p_value_row_idx[0], selected_stage] = overall_yield_p_value
                    else:
                        # If "Overall yield_Placement" doesn't exist, create a new row for it
                        new_row_oy = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                        new_row_oy.at[0, 'Data Points'] = "Overall yield_Placement"
                        new_row_oy.at[0, selected_stage] = overall_yield_p_value
                        edited_data = pd.concat([edited_data, new_row_oy], ignore_index=True)

                    # Update session state with the modified DataFrame
                    st.session_state.edited_sheets[sheet_name] = edited_data

                    # Optionally show a success message after calculation
                    # st.success(f"Overall yield_Component value calculated and updated: {overall_yield_p_value}")

                except ValueError:
                    # Show error if the inputs are not valid numbers
                    st.error("Please enter valid numbers for Ync & Ycl")

# *****************************************************************************************************************
# *****************************************************************************************************************
# *****************************************************************************************************************
# *****************************************************************************************************************
# *****************************************************************************************************************
# *****************************************************************************************************************

        # Overall yield_Placement Calculations
        # Button to trigger calculation of Pfi, Pfo, D, Ync, Ycl and Overall yield            
                try:
                    defect_rate_value = float(defect_rate) / 1000000  # Convert to defect rate per million
                    no_component_value = float(no_component_value)    # Convert solder_joint_value to float

                    # Calculate pfi_value using the given formula
                    pfi_value = 1 - (1 - defect_rate_value) ** no_component_value   # Formula: =1-(1-(500/1000000))^F16

                    # Update or insert "Pfi" value in the DataFrame for the selected stage
                    pfi_value_row_idx = edited_data[edited_data['Data Points'] == "Pfi"].index

                    if not pfi_value_row_idx.empty:
                        # If "Pfi" already exists, update its value
                        edited_data.at[pfi_value_row_idx[0], selected_stage] = pfi_value
                    else:
                        # If "Pfi" doesn't exist, create a new row for it
                        new_row_pfi = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                        new_row_pfi.at[0, 'Data Points'] = "Pfi"
                        new_row_pfi.at[0, selected_stage] = pfi_value
                        edited_data = pd.concat([edited_data, new_row_pfi], ignore_index=True)

                    # Update session state with the modified DataFrame
                    st.session_state.edited_sheets[sheet_name] = edited_data

                    # Optionally show a success message after calculation
                    # st.success(f"Pfi value calculated and updated: {pfi_value}")

                except ValueError:
                    # Show error if the inputs are not valid numbers
                    st.error("Please enter valid numbers for Defect Rate and Solder Joints")

                #"Pfo" calculation
                try:
                    test_efficiency_value = float(test_efficiency) / 100  # Convert percentage to decimal
                    pfi_value = float(pfi_value)    # Convert pfi_value to float

                    # Calculate Pfo using the given formula
                    pfo_value = 1 - (1 - pfi_value) ** (1 - test_efficiency_value)   # Formula: =1-((1-K16)^(1-$F$3))

                    # Update or insert "Pfo" value in the DataFrame for the selected stage
                    pfo_value_row_idx = edited_data[edited_data['Data Points'] == "Pfo"].index

                    if not pfo_value_row_idx.empty:
                        # If "Pfo" already exists, update its value
                        edited_data.at[pfo_value_row_idx[0], selected_stage] = pfo_value
                    else:
                        # If "Pfo" doesn't exist, create a new row for it
                        new_row_pfo = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                        new_row_pfo.at[0, 'Data Points'] = "Pfo"
                        new_row_pfo.at[0, selected_stage] = pfo_value
                        edited_data = pd.concat([edited_data, new_row_pfo], ignore_index=True)

                    # Update session state with the modified DataFrame
                    st.session_state.edited_sheets[sheet_name] = edited_data

                    # Optionally show a success message after calculation
                    # st.success(f"Pfo value calculated and updated: {pfo_value}")

                except ValueError:
                    # Show error if the inputs are not valid numbers
                    st.error("Please enter valid numbers for Pfi and Test Efficiency")

                #"D" calculation
                try:
                    pfi_value = float(pfi_value)    # Convert pfi_value to float
                    pfo_value = float(pfo_value)    # Convert pfo_value to float
                    test_efficiency_value = float(test_efficiency) / 100  # Convert percentage to decimal

                    # Calculate D using the given formula
                    avg_defects_per_brd_value = (pfi_value + pfo_value) * test_efficiency_value   # Formula: =SUM(K7:L7)*$F$3

                    # Update or insert "D" value in the DataFrame for the selected stage
                    avg_defects_per_brd_value_row_idx = edited_data[edited_data['Data Points'] == "D"].index

                    if not avg_defects_per_brd_value_row_idx.empty:
                        # If "Pfo" already exists, update its value
                        edited_data.at[avg_defects_per_brd_value_row_idx[0], selected_stage] = avg_defects_per_brd_value
                    else:
                        # If "D" doesn't exist, create a new row for it
                        new_row_avg_d = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                        new_row_avg_d.at[0, 'Data Points'] = "D"
                        new_row_avg_d.at[0, selected_stage] = avg_defects_per_brd_value
                        edited_data = pd.concat([edited_data, new_row_avg_d], ignore_index=True)

                    # Update session state with the modified DataFrame
                    st.session_state.edited_sheets[sheet_name] = edited_data

                    # Optionally show a success message after calculation
                    # st.success(f"Avg D value calculated and updated: {avg_defects_per_brd_value}")

                except ValueError:
                    # Show error if the inputs are not valid numbers
                    st.error("Please enter valid numbers for Pfi, Pfo and Test Efficiency")

                #"Poisson Yield (Ync)" calculation
                try:
                    avg_defects_per_brd_value = float(avg_defects_per_brd_value)    # Convert avg_defects_per_brd_value to float
                    # Calculate Poisson Yield (Ync) using the given formula
                    non_clustered_yield_value = np.exp(-avg_defects_per_brd_value)   # Formula: =EXP(-M16)

                    # Update or insert "Poisson Yield (Ync)" value in the DataFrame for the selected stage
                    non_clustered_yield_value_row_idx = edited_data[edited_data['Data Points'] == "Poisson Yield (Ync)"].index

                    if not non_clustered_yield_value_row_idx.empty:
                        # If "Poisson Yield (Ync)" already exists, update its value
                        edited_data.at[non_clustered_yield_value_row_idx[0], selected_stage] = non_clustered_yield_value
                    else:
                        # If "Poisson Yield (Ync)" doesn't exist, create a new row for it
                        new_row_ync = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                        new_row_ync.at[0, 'Data Points'] = "Poisson Yield (Ync)"
                        new_row_ync.at[0, selected_stage] = non_clustered_yield_value
                        edited_data = pd.concat([edited_data, new_row_ync], ignore_index=True)

                    # Update session state with the modified DataFrame
                    st.session_state.edited_sheets[sheet_name] = edited_data

                    # Optionally show a success message after calculation
                    # st.success(f"Ync value calculated and updated: {non_clustered_yield_value}")

                except ValueError:
                    # Show error if the inputs are not valid numbers
                    st.error("Please enter valid numbers for Avg D")

                #"Poisson Yield (Ycl)" calculation
                try:
                    avg_defects_per_brd_value = float(avg_defects_per_brd_value)    # Convert avg_defects_per_brd_value to float
                    alpha_α_assumed = float(alpha_α_assumed)  # Convert alpha_α_assumed to float

                    # Calculate Poisson Yield (Ycl) using the given formula
                    clustered_yield_value = (1 + (avg_defects_per_brd_value / alpha_α_assumed)) ** (-alpha_α_assumed)   # Formula: =(1+(M16/$C$27))^-$C$27
                    
                    # Update or insert "Poisson Yield (Ycl)" value in the DataFrame for the selected stage
                    clustered_yield_value_row_idx = edited_data[edited_data['Data Points'] == "Clustered Yield (Ycl)"].index

                    if not clustered_yield_value_row_idx.empty:
                        # If "Clustered Yield (Ycl)" already exists, update its value
                        edited_data.at[clustered_yield_value_row_idx[0], selected_stage] = clustered_yield_value
                    else:
                        # If "Clustered Yield (Ycl)" doesn't exist, create a new row for it
                        new_row_ycl = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                        new_row_ycl.at[0, 'Data Points'] = "Clustered Yield (Ycl)"
                        new_row_ycl.at[0, selected_stage] = clustered_yield_value
                        edited_data = pd.concat([edited_data, new_row_ycl], ignore_index=True)

                    # Update session state with the modified DataFrame
                    st.session_state.edited_sheets[sheet_name] = edited_data

                    # Optionally show a success message after calculation
                    # st.success(f"Ycl value calculated and updated: {clustered_yield_value}")

                except ValueError:
                    # Show error if the inputs are not valid numbers
                    st.error("Please enter valid numbers for Avg D & mean_μ")

                #"Overall Overall yield_Component" calculation
                try:
                    non_clustered_yield_value = float(non_clustered_yield_value)   # Convert non_clustered_yield_value to float
                    clustered_yield_value = float(clustered_yield_value)   # Convert clustered_yield_value to float

                    # Calculate Overall yield_Component using the given formula
                    overall_yield_c_value = non_clustered_yield_value * clustered_yield_value  # Formula: =H7*I7

                    # Update or insert "Overall yield_Component" value in the DataFrame for the selected stage
                    overall_yield_c_value_row_idx = edited_data[edited_data['Data Points'] == "Overall yield_Component"].index

                    if not overall_yield_c_value_row_idx.empty:
                        # If "Overall yield_Component" already exists, update its value
                        edited_data.at[overall_yield_c_value_row_idx[0], selected_stage] = overall_yield_c_value
                    else:
                        # If "Overall yield_Component" doesn't exist, create a new row for it
                        new_row_oy = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                        new_row_oy.at[0, 'Data Points'] = "Overall yield_Component"
                        new_row_oy.at[0, selected_stage] = overall_yield_c_value
                        edited_data = pd.concat([edited_data, new_row_oy], ignore_index=True)

                    # Update session state with the modified DataFrame
                    st.session_state.edited_sheets[sheet_name] = edited_data

                    # Optionally show a success message after calculation
                    st.success("Calculation for (Analyse Pfi, Pfo, D, Ync, Ycl and Overall yield) done successfully!!!")

                except ValueError:
                    # Show error if the inputs are not valid numbers
                    st.error("Please enter valid numbers for Ync & Ycl")


    # Third row: col11 to col14 under col5 to col10
    col13, col15, col16 = st.columns([1, 1, 1])




    with col13:
        # Save the edited table
        if st.button("Save Edited Table"):
            st.session_state.edited_sheets[sheet_name] = edited_data


            with pd.ExcelWriter(uploaded_file.name, engine="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
                edited_data.to_excel(writer, sheet_name=sheet_name, index=False)
            st.success("Table saved successfully!")
            st.rerun()


# Revision History  Date: 19-Oct-2024
# (updated the code for Overall yield_Component & Overall yield_Placement)
# Pending 2 error worning msg has to be resolved.
