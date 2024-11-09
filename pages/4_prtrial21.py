import streamlit as st
import pandas as pd
import numpy as np
import math
import os

# App Title
st.set_page_config(page_title="Process Yield Analysis!!!", page_icon=":bar_chart:", layout="wide")
st.title("Process Yield Analysis")

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

    # Assuming 'data' is a dictionary of DataFrames loaded from an Excel file
    col5, col6, col7, col8, col9 = st.columns([1, 1, 1, 1, 1])
    
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
            stages = ['MK0', 'MK1', 'MK2', 'MK3', 'X1', 'X1.1', 'X1.2']  # Add more stages if needed
            selected_stage = st.selectbox("Select the product development stage", stages)

        # Display data in a table
        st.subheader("Data Table")
        edited_data = st.data_editor(st.session_state.df)

        # 2. Actual first pass yield input
        with col7:
            actual_fpy = st.text_input("Enter Actual First Pass Yield (%)", value="")

        if actual_fpy:
            try:
                actual_fpy_value = float(actual_fpy) / 100  # Convert percentage to decimal
                log_afpy = -math.log10(actual_fpy_value)  # Calculate -Log of AFPY
                log_afpy_str = f"{log_afpy:.6f}"

                # Update existing rows for "Actual first pass yield" and "-Log of AFPY"
                afpy_row_idx = edited_data[edited_data['Package'] == "Actual first pass yield"].index
                log_afpy_row_idx = edited_data[edited_data['Package'] == "-Log of AFPY"].index

                # Insert "Actual first pass yield" value in the same row under the selected stage
                if not afpy_row_idx.empty:
                    edited_data.at[afpy_row_idx[0], selected_stage] = actual_fpy_value
                else:
                    new_row_fpy = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                    new_row_fpy.at[0, 'Package'] = "Actual first pass yield"
                    new_row_fpy.at[0, selected_stage] = actual_fpy_value
                    st.session_state.df = pd.concat([st.session_state.df, new_row_fpy], ignore_index=True)

                # Insert "-Log of AFPY" value in the same row under the selected stage
                if not log_afpy_row_idx.empty:
                    edited_data.at[log_afpy_row_idx[0], selected_stage] = log_afpy
                else:
                    new_row_log = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                    new_row_log.at[0, 'Package'] = "-Log of AFPY"
                    new_row_log.at[0, selected_stage] = log_afpy
                    st.session_state.df = pd.concat([st.session_state.df, new_row_log], ignore_index=True)

                st.session_state.edited_sheets[sheet_name] = st.session_state.df  # Update session state with new rows

            except ValueError:
                st.error("Please enter a valid number for Actual First Pass Yield")

        # 3. Calculate SUMPRODUCT for the selected stage and its corresponding fault probability
        fault_probability_column = f"{selected_stage} - Fault Probability"

        if selected_stage in edited_data.columns and fault_probability_column in edited_data.columns:
            try:
                # Calculate sum product for the selected stage and its corresponding fault probability
                sumproduct = (edited_data[selected_stage] * edited_data[fault_probability_column]).sum()
                # st.write(f"**Sum Product for {selected_stage}:** {sumproduct:.6f}")

                # Find the row where the "Package" is labeled as "Sum Product"
                sumproduct_row_idx = edited_data[edited_data['Package'] == "Sum Product"].index

                # Update the "Sum Product" row after the "-Log of AFPY" row
                if not sumproduct_row_idx.empty:
                    edited_data.at[sumproduct_row_idx[0], selected_stage] = sumproduct
                else:
                    # Insert "Sum Product" after the "-Log of AFPY" row
                    log_afpy_row_idx = edited_data[edited_data['Package'] == "-Log of AFPY"].index[0]
                    new_row_sumproduct = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                    new_row_sumproduct.at[0, 'Package'] = "Sum Product"
                    new_row_sumproduct.at[0, selected_stage] = sumproduct
                    edited_data = pd.concat([edited_data.iloc[:log_afpy_row_idx+1], new_row_sumproduct, edited_data.iloc[log_afpy_row_idx+1:]]).reset_index(drop=True)

                # 4. Calculate Estimated yield from current p value
                estimated_yield = np.exp(-(sumproduct))
                # st.write(f"**Estimated Yield from Current p Value for {selected_stage}:** {estimated_yield:.6f}")

                # Insert Estimated Yield row
                estimated_yield_row_idx = edited_data[edited_data['Package'] == "Estimated Yield from Current p Value"].index
                if not estimated_yield_row_idx.empty:
                    edited_data.at[estimated_yield_row_idx[0], selected_stage] = estimated_yield
                else:
                    new_row_estimated_yield = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                    new_row_estimated_yield.at[0, 'Package'] = "Estimated Yield from Current p Value"
                    new_row_estimated_yield.at[0, selected_stage] = estimated_yield
                    sumproduct_row_idx = edited_data[edited_data['Package'] == "Sum Product"].index[0]
                    edited_data = pd.concat([edited_data.iloc[:sumproduct_row_idx+1], new_row_estimated_yield, edited_data.iloc[sumproduct_row_idx+1:]]).reset_index(drop=True)

                # 5. Calculate Squared Error
                squared_error = np.exp(-(estimated_yield))
                # st.write(f"**Squared Error for {selected_stage}:** {squared_error:.6f}")

                # Insert Squared Error row
                squared_error_row_idx = edited_data[edited_data['Package'] == "Squared Error"].index
                if not squared_error_row_idx.empty:
                    edited_data.at[squared_error_row_idx[0], selected_stage] = squared_error
                else:
                    new_row_squared_error = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                    new_row_squared_error.at[0, 'Package'] = "Squared Error"
                    new_row_squared_error.at[0, selected_stage] = squared_error
                    estimated_yield_row_idx = edited_data[edited_data['Package'] == "Estimated Yield from Current p Value"].index[0]
                    edited_data = pd.concat([edited_data.iloc[:estimated_yield_row_idx+1], new_row_squared_error, edited_data.iloc[estimated_yield_row_idx+1:]]).reset_index(drop=True)

                # 6. Calculate z (which is what we have to minimize)
                z_value = np.exp(-(squared_error))
                # st.write(f"**z (which is what we have to minimize) for {selected_stage}:** {z_value:.6f}")

                # Insert z row
                z_row_idx = edited_data[edited_data['Package'] == "z (which is what we have to minimize)"].index
                if not z_row_idx.empty:
                    edited_data.at[z_row_idx[0], selected_stage] = z_value
                else:
                    new_row_z = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                    new_row_z.at[0, 'Package'] = "z (which is what we have to minimize)"
                    new_row_z.at[0, selected_stage] = z_value
                    squared_error_row_idx = edited_data[edited_data['Package'] == "Squared Error"].index[0]
                    edited_data = pd.concat([edited_data.iloc[:squared_error_row_idx+1], new_row_z, edited_data.iloc[squared_error_row_idx+1:]]).reset_index(drop=True)

                st.session_state.edited_sheets[sheet_name] = edited_data  # Update session state with the calculations
            except Exception as e:
                st.error(f"Error calculating SUMPRODUCT and related values: {e}")

        # Enter Solder Joint
        with col8:
            solder_joint_value = st.text_input("Enter No Of Solder Joint", value="")
        if solder_joint_value:
            try:
                solder_joint_value = float(solder_joint_value)
                end_index = edited_data[edited_data['Package'] == "End Of Input Table"].index[0]
                no_of_opportunities_to_failure = ((edited_data.loc[:end_index - 1, selected_stage].sum()) * 2) + solder_joint_value
                no_of_opportunities_to_failure_str = f"{no_of_opportunities_to_failure:.6f}"
                
                # Find the row where the "Package" is labeled as "Solder Joint" & "Number of opportunities to failure at the point n, On"
                solder_joint_row_idx = edited_data[edited_data['Package'] == "Solder Joint"].index
                no_of_opportunities_to_failure_idx = edited_data[edited_data['Package'] == "Number of opportunities to failure at the point n, On"].index

                # Insert "Solder Joint" value in the same row under the selected stage at the end
                if not solder_joint_row_idx.empty:
                    edited_data.at[solder_joint_row_idx[0], selected_stage] = solder_joint_value
                else:
                    new_row_solder_joint = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                    new_row_solder_joint.at[0, 'Package'] = "Solder Joint"
                    new_row_solder_joint.at[0, selected_stage] = solder_joint_value
                    z_row_idx = edited_data[edited_data['Package'] == "z (which is what we have to minimize)"].index[0]
                    edited_data = pd.concat([edited_data.iloc[:z_row_idx+1], new_row_solder_joint, edited_data.iloc[z_row_idx+1:]]).reset_index(drop=True)

                # Insert "Number of opportunities to failure at the point n, On" value in the same row under the selected stage
                if not no_of_opportunities_to_failure_idx.empty:
                    edited_data.at[no_of_opportunities_to_failure_idx[0], selected_stage] = no_of_opportunities_to_failure
                else:
                    new_row_opp = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                    new_row_opp.at[0, 'Package'] = "Number of opportunities to failure at the point n, On"
                    new_row_opp.at[0, selected_stage] = no_of_opportunities_to_failure
                    st.session_state.df = pd.concat([st.session_state.df, new_row_opp], ignore_index=True)

                st.session_state.edited_sheets[sheet_name] = st.session_state.df  # Update session state with new row
            except Exception as e:
                st.error(f"Error calculating SUMPRODUCT and related values: {e}")


        # Enter Test coverage
        with col9:
            test_coverage_value = st.text_input("Enter Test coverage %, Ct", value="")
        if test_coverage_value:
            try:
                # Ensure valid conversion to float
                test_coverage_value = float(test_coverage_value)                

                # Calculate estimated DPMO
                estimated_dpmo = (-math.log(estimated_yield) / (no_of_opportunities_to_failure * test_coverage_value)) * (10**6)
                estimated_dpmo_str = f"{no_of_opportunities_to_failure:.6f}"

                # Find the row for "Enter Test coverage %, Ct" and "Estimated DPMO"
                test_coverage_row_idx = edited_data[edited_data['Package'] == "Enter Test coverage %, Ct"].index
                estimated_dpmo_row_idx = edited_data[edited_data['Package'] == "Estimated DPMO"].index

                # Insert "Enter Test coverage %, Ct" value in the same row
                if not test_coverage_row_idx.empty:
                    edited_data.at[test_coverage_row_idx[0], selected_stage] = test_coverage_value
                else:
                    new_row_test_coverage = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                    new_row_test_coverage.at[0, 'Package'] = "Enter Test coverage %, Ct"
                    new_row_test_coverage.at[0, selected_stage] = test_coverage_value
                    no_of_opportunities_to_failure_idx = edited_data[edited_data['Package'] == "Number of opportunities to failure at the point n, On"].index[0]
                    edited_data = pd.concat([edited_data.iloc[:no_of_opportunities_to_failure_idx+1], new_row_test_coverage, edited_data.iloc[no_of_opportunities_to_failure_idx+1:]]).reset_index(drop=True)

                # Insert "Estimated DPMO" value in the same row
                if not estimated_dpmo_row_idx.empty:
                    edited_data.at[estimated_dpmo_row_idx[0], selected_stage] = estimated_dpmo
                else:
                    new_row_estimated_dpmo = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
                    new_row_estimated_dpmo.at[0, 'Package'] = "Estimated DPMO"
                    new_row_estimated_dpmo.at[0, selected_stage] = estimated_dpmo
                    st.session_state.df = pd.concat([st.session_state.df, new_row_estimated_dpmo], ignore_index=True)

                st.session_state.edited_sheets[sheet_name] = st.session_state.df  # Update session state with new row

            except ValueError:
                st.error("Please enter a valid number for Test Coverage")

        # Select the position to add the new row
        row_position_options = ["Above", "Below"]
        selected_row = st.selectbox("Select row to insert new row above or below", edited_data.index)
        insert_position = st.selectbox("Insert row", row_position_options)

        # Create buttons side by side for adding new rows and saving the table
        col1, col2 = st.columns([1, 1])

        with col1:
            # Add new row at the specified position
            if st.button("Add New Row"):
                new_row = pd.DataFrame({col: [np.nan] for col in edited_data.columns})

                # Insert new row above or below the selected row
                if insert_position == "Above":
                    st.session_state.df = pd.concat([edited_data.iloc[:selected_row], new_row, edited_data.iloc[selected_row:]]).reset_index(drop=True)
                elif insert_position == "Below":
                    st.session_state.df = pd.concat([edited_data.iloc[:selected_row + 1], new_row, edited_data.iloc[selected_row + 1:]]).reset_index(drop=True)

                st.session_state.edited_sheets[sheet_name] = st.session_state.df  # Update session state with new row
                st.rerun()

        # Save the edited table
        with col2:
            if st.button("Save Edited Table"):
                st.session_state.edited_sheets[sheet_name] = edited_data
                with pd.ExcelWriter(uploaded_file.name, engine="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
                    edited_data.to_excel(writer, sheet_name=sheet_name, index=False)
                st.success("Table saved successfully!")
                st.rerun()

        # Create buttons side by side for removing rows and saving removed rows
        col3, col4 = st.columns([1, 1])

        with col3:
            row_to_delete = st.selectbox("Select row to delete", st.session_state.df.index)
            if st.button("Remove Row"):
                if 'removed_rows' not in st.session_state:
                    st.session_state.removed_rows = pd.DataFrame()
                st.session_state.removed_rows = pd.concat([st.session_state.removed_rows, st.session_state.df.loc[[row_to_delete]]])
                st.session_state.df = st.session_state.df.drop(row_to_delete).reset_index(drop=True)
                st.session_state.edited_sheets[sheet_name] = st.session_state.df  # Update session state with removed row
                st.rerun()

        with col4:
            if st.button("Save Removed Rows"):
                if 'removed_rows' in st.session_state and not st.session_state.removed_rows.empty:
                    excel_file = pd.read_excel(uploaded_file.name, sheet_name=None)
                    excel_file[sheet_name] = st.session_state.df
                    with pd.ExcelWriter(uploaded_file.name, engine="openpyxl", mode="w") as writer:
                        for sheet_name, sheet_data in excel_file.items():
                            sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)
                    st.success("Removed rows saved successfully!")
                    st.rerun()
