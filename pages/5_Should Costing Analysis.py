import streamlit as st
import pandas as pd
import openpyxl
import numpy as np
from openpyxl.utils.dataframe import dataframe_to_rows
import matplotlib
import os
import tempfile
import io
import plotly.graph_objects as go
import uuid  # Add this import at the top of your script
import math
import plotly.express as px

# Set the page layout to wide
st.set_page_config(layout="wide")

# Title of the app
st.title(":bar_chart: Should Costing Analysis")

# Sidebar configuration
st.sidebar.header("Category")

# Create an expander for "Offers"
with st.sidebar.expander("Analysis"):
    # Create checkboxes for each offer
    new_analysis = st.checkbox("New")
    existing_analysis = st.checkbox("Existing")

# Display selected analysis
if new_analysis:
    st.subheader("New Analysis")

    # File uploader for the first Excel file (simulation_db.xlsx, sheet 'Process_CT')
    uploaded_file_simulation_db = st.file_uploader("Upload the simulation_db.xlsx file", type=["xlsx"])
    
    # Load data if the file is uploaded
    if uploaded_file_simulation_db:        
        try:
            df = pd.read_excel(uploaded_file_simulation_db, sheet_name='Process_CT')
            # st.write("File successfully read. Preview below:")
            # st.dataframe(df.head())
        except Exception as e:
            st.error(f"Error while reading the file: {e}")
            
        # Extract the required values
        shift_hr_day = df.at[0, 'Shift Hr/day']
        days_week = df.at[0, 'Days/Week']
        weeks_year = df.at[0, 'Weeks/Year']
        hr_year_shift = shift_hr_day * days_week * weeks_year
        overall_labor_efficiency = df.at[0, 'Overall Labor Efficiency']
        
        # Hide the dataframe
        st.write("")

        # Create text inputs for each value
        col1, col2, col3, vol_col1, vol_col2, vol_col3 = st.columns(6)

        with col1:
            shift_hr_day_input = st.text_input('Shift Hr/day', value=shift_hr_day, disabled=True)

        with col2:
            days_week_input = st.text_input('Days/Week', value=days_week, disabled=True)

        with col3:
            overall_labor_efficiency_input = st.text_input('Overall Labor Efficiency', value=overall_labor_efficiency, disabled=True)

        # Load the specific sheet from simulation_db.xlsx for 'NRE'
        df2 = pd.read_excel(uploaded_file_simulation_db, sheet_name='NRE')

        st.write("-------------------")

        # Create an empty DataFrame with the defined columns
        initial_df = pd.DataFrame(columns=['Item', 'Unit Price (₹)', 'Life Cycle (Boards)', 'Qty for LCV', "Extended Price (₹)"])

        # Initialize session state variables
        if 'data' not in st.session_state:
            st.session_state['data'] = initial_df

        if 'filtered_data' not in st.session_state:
            st.session_state['filtered_data'] = initial_df

        # Initialize dropdown values if not set
        if 'item' not in st.session_state:
            st.session_state['item'] = ''

        if 'unit_price' not in st.session_state:
            st.session_state['unit_price'] = ''

        if 'life_cycle_boards' not in st.session_state:
            st.session_state['life_cycle_boards'] = ''

        if 'qty_for_lcv' not in st.session_state:
            st.session_state['qty_for_lcv'] = ''

        if 'ext_price' not in st.session_state:
            st.session_state['ext_price'] = ''

        if 'reset_selectbox' not in st.session_state:
            st.session_state['reset_selectbox'] = 0

        # Define the Product Volume from the 'Process_CT' sheet
        # vol_col1, vol_col2, vol_col3 = st.columns(3)

        # Text inputs for Annual Volume and Product Life
        with vol_col1:
            annual_volume = st.text_input('Annual Volume', value="", disabled=False)

        with vol_col2:
            product_life = st.text_input('Product Life', value="", disabled=False)

        # Safely convert inputs to float, defaulting to 0 if conversion fails
        try:
            annual_volume = float(annual_volume) if annual_volume else 0.0
        except ValueError:
            annual_volume = 0.0
            st.warning("Invalid input for 'Annual Volume'. Please enter a number.")

        try:
            product_life = float(product_life) if product_life else 0.0
        except ValueError:
            product_life = 0.0
            st.warning("Invalid input for 'Product Life'. Please enter a number.")

        # Perform the calculation for Annual Volume
        product_volume = annual_volume * product_life

        # Display results
        with vol_col3:
            st.text_input('Product Volume', value=product_volume, disabled=True)
              
        # Display the headings
        header_cols = st.columns(5)
        header_cols[0].markdown("<h6 style='text-align: center;'>Item</h6>", unsafe_allow_html=True)
        header_cols[1].markdown("<h6 style='text-align: center;'>Unit Price (₹)</h6>", unsafe_allow_html=True)
        header_cols[2].markdown("<h6 style='text-align: center;'>Life Cycle (Boards)</h6>", unsafe_allow_html=True)
        header_cols[3].markdown("<h6 style='text-align: center;'>Qty for LCV</h6>", unsafe_allow_html=True)
        header_cols[4].markdown("<h6 style='text-align: center;'>Extended Price (₹)</h6>", unsafe_allow_html=True)
        
        # Function to display a row
        def display_row():
            row_cols = st.columns(5)
            
            # Select boxes to select the “Item”
            item = row_cols[0].selectbox('', [''] + list(df2['Item'].unique()), key=f'item_{st.session_state.reset_selectbox}')
            unit_price = df2[df2['Item'] == item]['Unit Price (₹)'].values[0] if item else ''
            life_cycle_boards = df2[df2['Item'] == item]['Life Cycle (Boards)'].values[0] if item else ''
            
            # Apply the formula for Qty for LCV
            qty_for_lcv = 1 * (max(product_volume, life_cycle_boards) / life_cycle_boards) if life_cycle_boards else ''
            ext_price = unit_price * qty_for_lcv if unit_price and qty_for_lcv else ''

            with row_cols[1]:
                unit_price_input = st.text_input('', value=unit_price, key=f'unit_price_{st.session_state.reset_selectbox}')

            with row_cols[2]:
                life_cycle_boards_input = st.text_input('', value=life_cycle_boards, key=f'life_cycle_boards_{st.session_state.reset_selectbox}')

            with row_cols[3]:
                qty_for_lcv_input = st.text_input('', value=qty_for_lcv, key=f'qty_for_lcv_{st.session_state.reset_selectbox}')

            with row_cols[4]:
                ext_price_input = st.text_input('', value=ext_price, key=f'ext_price_{st.session_state.reset_selectbox}')

        # Display the row
        display_row()


        # Add Save, Clear, and Delete buttons
        save_col, clear_col, delete_col3, delete_col4 = st.columns(4)
        with save_col:
            if st.button('Save'):
                # Save the current selection to session state data
                item = st.session_state[f'item_{st.session_state.reset_selectbox}']
                unit_price = st.session_state[f'unit_price_{st.session_state.reset_selectbox}']
                life_cycle_boards = st.session_state[f'life_cycle_boards_{st.session_state.reset_selectbox}']
                qty_for_lcv = st.session_state[f'qty_for_lcv_{st.session_state.reset_selectbox}']
                ext_price = st.session_state[f'ext_price_{st.session_state.reset_selectbox}']

                if item:
                    new_row = {
                        'Item': item,
                        'Unit Price (₹)': unit_price,                        
                        'Life Cycle (Boards)': life_cycle_boards,
                        'Qty for LCV': qty_for_lcv,
                        'Extended Price (₹)': ext_price
                    }
                    if not st.session_state['filtered_data']['Item'].eq(item).any():
                        st.session_state['filtered_data'] = pd.concat([st.session_state['filtered_data'], pd.DataFrame([new_row])], ignore_index=True)
                        st.success("Record added successfully. Select Your Next Side & Stage")
                    else:
                        st.warning("Record Already Exists in the Table")

        with clear_col:
            if st.button('Clear'):
                # Increment the key to reset the select boxes
                st.session_state['reset_selectbox'] += 1

        # Display the updated dataframe with a header
        st.markdown("## NRE Mapping")
        st.dataframe(st.session_state['filtered_data'], use_container_width=True)

        totalcost_col1, toolmaintenance_col2, totalextendedprice_col3, nreperunit_col4 = st.columns(4)

        with totalcost_col1:
            # Convert the 'Extended Price (₹)' column to numeric
            st.session_state['filtered_data']['Extended Price (₹)'] = pd.to_numeric(st.session_state['filtered_data']['Extended Price (₹)'], errors='coerce')            
            # Calculate the total cost
            total_cost = st.session_state['filtered_data']['Extended Price (₹)'].sum()
            total_cost_value = float(total_cost)
            st.text_input('Total Cost (₹)', value=total_cost, disabled=True)

        with toolmaintenance_col2:
            tool_maintenance_rate = st.text_input('Tool Maintenance Rate (%)', value="", disabled=False)
            tool_maintenance_rate_value = float(tool_maintenance_rate) / 100 if tool_maintenance_rate else 0.0

        with totalextendedprice_col3:
            tool_maintenance_cost = total_cost_value * tool_maintenance_rate_value
            total_extended_price = total_cost_value + tool_maintenance_cost
            st.text_input('Total Extended Price (₹)', value=total_extended_price, disabled=True)

        with nreperunit_col4:
            nre_per_unit = total_extended_price / product_volume if product_volume else 0
            st.text_input('NRE per Unit (₹)', value=nre_per_unit, disabled=True)

        # Provide inputs for file name, sheet name, and path
        st.markdown("### Save Data to Excel")

        # Provide inputs for file name and sheet name only (without path)
        file_name = st.text_input("Enter the Excel file name (with .xlsx extension):")
        sheet_name = st.text_input("Enter the sheet name:")

        # Add a button to save the entire DataFrame
        if st.button("Save DataFrame to Excel"):
            with tempfile.TemporaryDirectory() as tmpdirname:
                full_path = os.path.join(tmpdirname, file_name)

                # Write data to Excel file
                with pd.ExcelWriter(full_path, engine='openpyxl') as writer:
                    # Prepare the DataFrame for saving
                    final_df = st.session_state['filtered_data'].copy()
                    
                    # Add the additional fields in the first row
                    for col, value in zip(
                        ['Annual Volume', 'Product Life', 'Product Volume', 'Total Cost (₹)', 'Tool Maintenance Rate (%)', 
                        'Extended Price (₹)', 'NRE Per Unit ($)'],
                        [annual_volume, product_life, product_volume, total_cost, tool_maintenance_rate_value, 
                        total_extended_price, nre_per_unit]):
                        final_df.at[0, col] = value
                    final_df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Load the saved file and create a download link
                with open(full_path, "rb") as f:
                    st.download_button(
                        label="Download Excel file",
                        data=f,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        st.write("-------------------")

        # Load the specific sheet from simulation_db.xlsx for 'MMR-EMS'
        df3 = pd.read_excel(uploaded_file_simulation_db, sheet_name='MMR-EMS')
        df4 = pd.read_excel(uploaded_file_simulation_db, sheet_name='Assumptions')

        # File uploader for Excel/CSV/XLSM files
        uploaded_file = st.file_uploader("Choose Process Mapping Excel/CSV/XLSM file", type=["xlsx", "csv", "xlsm"])

        if uploaded_file:
            # Load data from the uploaded file
            @st.cache_data
            def load_data(file):
                if file.name.endswith('.csv'):
                    df = pd.read_csv(file)
                else:
                    df = pd.read_excel(file, sheet_name=None)
                return df

            df6 = load_data(uploaded_file)  # Load the Process Mapping data

            # Initialize session state to store edited data for each sheet
            if 'edited_sheets' not in st.session_state:
                st.session_state.edited_sheets = {}

            processmapping_col1, processmapping_col2 = st.columns(2)

            if isinstance(df6, dict):
                with processmapping_col1:
                    # Allow user to select a sheet
                    sheet_name = st.selectbox("Select the sheet", df6.keys())

                if sheet_name in df6:
                    selected_data = df6[sheet_name]
                    st.session_state.df = pd.DataFrame(selected_data)  # Load original data from the selected sheet

                    st.subheader("Data Table")

                # Ensure 'MMR' exists in df3
                if 'MMR' not in df3.columns:
                    st.error("'MMR' column not found in 'MMR-EMS' sheet. Please check the input file.")
                else:
                    # Merge the data, keeping df5 (selected sheet) as left and matching with df3
                    edited_data = st.session_state.df.merge(
                        df3, left_on='Stage', right_on='Process Name', how='inner'
                    )

                    # Display only matching rows after the merge
                    if edited_data.empty:
                        st.warning("No matching content found between the Process Mapping and MMR-EMS datasets.")
                    else:
                        # Fill NaN values with 0
                        edited_data.fillna(0, inplace=True)

                    # Fill NaN values in 'MMR' with 0
                    edited_data['MMR'] = edited_data['MMR'].fillna(0)

                    # Add necessary columns with default values
                    edited_data['VA MC Cost'] = np.nan
                    edited_data['Batch Set up Cost'] = np.nan
                    edited_data['Labour cost/Hr'] = np.nan

                    # Calculate VA MC Cost
                    edited_data['VA MC Cost'] = edited_data['Process Cycle Time'] * edited_data['MMR']
                    edited_data['VA MC Cost'] = edited_data['VA MC Cost'].fillna(0)

                    try:
                        labour_cost_hr = df4.loc[0, 'Labour cost/Hr']
                        idl_cost_hr = df4.loc[0, 'Idl Cost/Hr']
                        months_per_year = 12
                        batch_qty = annual_volume / months_per_year if annual_volume > 0 else 1  # Avoid division by zero

                        # Calculate Batch Set up Cost
                        edited_data['Batch Set up Cost'] = (
                            (((edited_data['Batch Set up Time'] * labour_cost_hr) / 3600) * 1.15) / batch_qty
                        ) * edited_data['FTE for Batch Set up']
                        
                        # Calculate Labour Cost
                        edited_data['Labour cost/Hr'] = (
                            ((((edited_data['Process Cycle Time'] * labour_cost_hr) / 3600) * 1.15) * edited_data['DL FTE']) + 
                            ((((edited_data['Process Cycle Time'] * idl_cost_hr) / 3600) * 1.15) * edited_data['IDL FTE'])
                        )

                        # Fill NaN values with 0
                        edited_data['Batch Set up Cost'] = edited_data['Batch Set up Cost'].fillna(0)
                        edited_data['Labour cost/Hr'] = edited_data['Labour cost/Hr'].fillna(0)

                        # Update session state
                        st.session_state.edited_sheets[sheet_name] = edited_data

                    except KeyError as e:
                        st.error(f"Error in calculation: Missing column {e}. Please check the input data.")
                # Display the updated DataFrame
                st.data_editor(edited_data, key=f"data_editor_{sheet_name}_updated")

                st.header("Consumable Costing")

                # Create one row with 4 columns for headings
                rtv_col, solder_top_col, solder_bar_col = st.columns(3)

                # RTV Glue
                with rtv_col:
                    st.subheader("RTV Glue")
                    rtv_1_col, rtv_2_col = st.columns(2)
                    with rtv_1_col:
                    # Input Fields
                        glue_wt_per_board = st.text_input('RTV Wt/Brd Est', value="", key="glue_wt_per_board")
                        rtv_glue_cost = st.text_input('RTV Cost/ml', value="0.052", key="rtv_glue_cost", disabled=True)
                    
                    with rtv_2_col:
                        wastage_percentage_per_board = st.text_input('RTV Wastage %', value="", key="wastage_percentage_per_board")
                        specific_gravity_of_solder = st.text_input('RTV Solder SG', value="1.09", key="specific_gravity_of_solder", disabled=True)

                    # Initialize the output variables
                    wt_per_board_incl_wastage = 0.0
                    rtv_cost_per_board = 0.0

                    try:
                        # Safely convert inputs to float
                        glue_wt_per_board = float(glue_wt_per_board) if glue_wt_per_board else 0.0
                        wastage_percentage_per_board = float(wastage_percentage_per_board) if wastage_percentage_per_board else 0.0
                        rtv_glue_cost = float(rtv_glue_cost) if rtv_glue_cost else 0.0
                        specific_gravity_of_solder = float(specific_gravity_of_solder) if specific_gravity_of_solder else 0.0

                        # Calculate wt_per_board_incl_wastage
                        wt_per_board_incl_wastage = (
                            (glue_wt_per_board * specific_gravity_of_solder) * (1 + (wastage_percentage_per_board / 100))
                        )

                        # Calculate rtv_cost_per_board
                        rtv_cost_per_board = wt_per_board_incl_wastage * rtv_glue_cost

                    except ValueError:
                        st.error("Please enter valid numeric values for all inputs.")

                    # Display results
                    with rtv_1_col:
                        st.text_input('RTV Wt/Brd', value=f"{wt_per_board_incl_wastage:.2f}", key="wt_per_board_incl_wastage", disabled=True)
                    # with rtv_2_col:
                    st.text_input('RTV Cost Per Board', value=f"{rtv_cost_per_board:.2f}", key="rtv_cost_per_board", disabled=True)

                # Solder Paste - Top
                with solder_top_col:
                    st.subheader("Solder Paste - Top & Bot")
                    # Board dimensions (Board Length and Board Width in the same line)
                    board_dim_col1, board_dim_col2 = st.columns(2)
                    with board_dim_col1:
                        board_length = st.text_input('Board Length(mm)', value="", key="board_length")
                    with board_dim_col2:
                        # st.write("Solder Paste - Bot")
                        board_width = st.text_input('Board Width(mm)', value="", key="board_width")

                    sp_sg_col, sp_cost_col = st.columns(2)
                    with sp_sg_col:
                        paste_specific_gravity = st.text_input('Solder Paste SG (g/cc)', value="7.31", key="paste_specific_gravity", disabled=True)
                    with sp_cost_col:
                        cost_of_solder_paste = st.text_input('Solder Paste Cost($/g)', value="0.065", key="cost_of_solder_paste", disabled=True)

                    sp_top_col1, sp_bot_col2 = st.columns(2)
                    with sp_top_col1:
                        top_weight_estimate_percentage = st.text_input('Top Wt Estimate %', value="", key="top_weight_estimate_percentage")
                        top_sp_wastage_percentage = st.text_input('Top Wastage %', value="", key="top_sp_wastage_percentage")
                        solder_paste_thickness = st.text_input('Top SP Thick(mm)', value="", key="solder_paste_thickness")

                    try:
                        # Safely convert inputs to float
                        board_length = float(board_length) if board_length else 0.0
                        board_width = float(board_width) if board_width else 0.0
                        top_weight_estimate_percentage = float(top_weight_estimate_percentage) if top_weight_estimate_percentage else 0.0
                        top_sp_wastage_percentage = float(top_sp_wastage_percentage) if top_sp_wastage_percentage else 0.0
                        paste_specific_gravity = float(paste_specific_gravity) if paste_specific_gravity else 7.31
                        cost_of_solder_paste = float(cost_of_solder_paste) if cost_of_solder_paste else 0.065
                        solder_paste_thickness = float(solder_paste_thickness) if solder_paste_thickness else 0.0

                        # Calculate weight_of_solder_paste_for_100percentage_wt
                        weight_of_solder_paste_for_100percentage_wt_value = (
                            (board_length * board_width * solder_paste_thickness * paste_specific_gravity) / 1000
                        )

                        # Convert percentages to fractions
                        top_weight_estimate_percentage /= 100
                        top_sp_wastage_percentage /= 100

                        # Calculate top_solder_paste_weight_estimate
                        top_weight_of_solder_paste_for_wt_estimate_value = (
                            weight_of_solder_paste_for_100percentage_wt_value
                            * top_weight_estimate_percentage
                            * (1 + top_sp_wastage_percentage)
                        )

                        # Calculate top_side_cost_per_board
                        top_side_cost_per_board_value = top_weight_of_solder_paste_for_wt_estimate_value * cost_of_solder_paste

                    except ValueError:
                        st.error("Please enter valid numeric values for all inputs.")
                    with sp_top_col1:
                        # Display results
                        st.text_input('Top SP Wt (100%)(g)', value=f"{weight_of_solder_paste_for_100percentage_wt_value:.2f}", key="weight_of_solder_paste_for_100percentage_wt", disabled=True)
                        st.text_input('Top SP Wt Estimate(g)', value=f"{top_weight_of_solder_paste_for_wt_estimate_value:.2f}", key="top_solder_paste_weight_estimate", disabled=True)
                        st.text_input('Top SP Cost/Brd($)', value=f"{top_side_cost_per_board_value:.2f}", key="top_side_cost_per_board", disabled=True)                # Solder Paste - Bottom


                # Solder Paste - Bottom
                # with solder_bottom_col:
                with solder_top_col:
                    # st.subheader("Solder Paste - Bottom")
                    # Input Fields
                    with sp_bot_col2:
                        bot_weight_estimate_percentage = st.text_input('Bot Wt Estimate %', value="", key="bot_weight_estimate_percentage")
                        bot_sp_wastage_percentage = st.text_input('Bot Wastage %', value="", key="bot_sp_wastage_percentage")
                        bot_solder_paste_thickness = st.text_input('Bot SP Thick(mm)', value="", key="bot_solder_paste_thickness")

                    try:
                        # Safely convert inputs to float
                        bot_weight_estimate_percentage = float(bot_weight_estimate_percentage) if bot_weight_estimate_percentage else 0.0
                        bot_sp_wastage_percentage = float(bot_sp_wastage_percentage) if bot_sp_wastage_percentage else 0.0
                        bot_solder_paste_thickness = float(bot_solder_paste_thickness) if bot_solder_paste_thickness else 0.0

                        # Calculate weight_of_solder_paste_for_100percentage_wt
                        bot_weight_of_solder_paste_for_100percentage_wt_value = (
                            (board_length * board_width * bot_solder_paste_thickness * paste_specific_gravity) / 1000
                        )

                        # Convert percentages to fractions
                        bot_weight_estimate_percentage /= 100
                        bot_sp_wastage_percentage /= 100

                        # Compute bot_weight_of_solder_paste_for_wt_estimate_value
                        bot_weight_of_solder_paste_for_wt_estimate_value = (
                            bot_weight_of_solder_paste_for_100percentage_wt_value
                            * bot_weight_estimate_percentage
                            * (1 + bot_sp_wastage_percentage)
                        )

                        # Compute bot_side_cost_per_board
                        bot_side_cost_per_board_value = bot_weight_of_solder_paste_for_wt_estimate_value * cost_of_solder_paste

                    except ValueError:
                        st.error("Please enter valid numeric values for all inputs.")
                    with sp_bot_col2:
                        # Display results
                        st.text_input('Bot SP Wt (100%)(g)', value=f"{bot_weight_of_solder_paste_for_100percentage_wt_value:.2f}", key="bot_weight_of_solder_paste_for_100percentage_wt_value", disabled=True)
                        st.text_input('Bot SP Wt Estimate(g)', value=f"{bot_weight_of_solder_paste_for_wt_estimate_value:.2f}", key="bot_solder_paste_weight_estimate", disabled=True)
                        st.text_input('Bot SP Cost/Brd($)', value=f"{bot_side_cost_per_board_value:.2f}", key="bot_side_cost_per_board", disabled=True)

                # Flux Wave Soldering
                with rtv_col:
                    st.subheader("Flux Wave Soldering")
                    rtv_3_col, rtv_4_col = st.columns(2)

                    # Input Fields
                    with rtv_3_col:
                        flux_wastage_percentage = st.text_input('Flux Wastage %', value="", key="flux_wastage_percentage")
                    with rtv_4_col:
                        flux_cost = st.text_input('Flux Cost($/ml)', value="0.0055", key="flux_cost", disabled=True)

                    try:
                        # Safely convert inputs to float
                        flux_wastage_percentage = float(flux_wastage_percentage) if flux_wastage_percentage else 0.0
                        flux_cost = float(flux_cost) if flux_cost else 0.0055

                        # Compute flux_board_area
                        flux_board_area_value = board_length * board_width

                        # Convert percentage to fraction
                        flux_wastage_percentage /= 100

                        # Compute flux_spread_area
                        flux_spread_area_value = ((flux_board_area_value / 100) * 0.1) * (1 + flux_wastage_percentage)

                    except ValueError:
                        st.error("Please enter valid numeric values for all inputs.")

                    # Display results
                    with rtv_3_col:
                        st.text_input('Flux Area/Brd(mm^2)', value=f"{flux_board_area_value:.2f}", key="flux_board_area", disabled=True)
                    with rtv_4_col:
                        st.text_input('Flux Spray Area(mm^2)', value=f"{flux_spread_area_value:.2f}", key="flux_spread_area", disabled=True)
                    flux_cost_per_board_value = flux_spread_area_value * flux_cost
                    flux_cost_per_board = st.text_input('Flux Cost Per Board($)', value=flux_cost_per_board_value, key="flux_cost_per_board", disabled=True)

                # Solder bar
                with solder_bar_col:
                    circumferential_col, barrel1_col, barrel2_col = st.columns(3)
                    with circumferential_col:
                        st.subheader("Solder Bar")
                        # Input Fields
                        outer_dia_of_pad = st.text_input('Pad OD (mm)', value="", key="outer_dia_of_pad", disabled=False)
                        inner_dia_of_pad = st.text_input('Pad ID (mm)', value="", key="inner_dia_of_pad", disabled=False)
                        no_of_solder_joints = st.text_input('Solder Joints', value="", key="no_of_solder_joints", disabled=False)
                        thickness_of_solder = st.text_input('Solder Thick (mm)', value="0.6", key="thickness_of_solder", disabled=True)

                        try:
                            # Safely convert inputs to float
                            outer_dia_of_pad = float(outer_dia_of_pad) if outer_dia_of_pad else 0.0
                            inner_dia_of_pad = float(inner_dia_of_pad) if inner_dia_of_pad else 0.0
                            no_of_solder_joints = float(no_of_solder_joints) if no_of_solder_joints else 0.0
                            thickness_of_solder = float(thickness_of_solder) if thickness_of_solder else 0.0

                            # Calculate the area of the annular ring
                            area_of_ring = math.pi * ((outer_dia_of_pad/2)**2 - (inner_dia_of_pad/2)**2)

                            # Calculate the volume of solder per joint - Circumferential Fill
                            volume_of_solder_per_joint = area_of_ring * thickness_of_solder
                            weight_of_Solder_per_joint = (volume_of_solder_per_joint/1000) * paste_specific_gravity
                            weight_of_Solder_per_board = weight_of_Solder_per_joint * no_of_solder_joints

                        except ValueError:
                            st.error("Please enter valid numeric values for all inputs.")

                        # Display the calculated value - Circumferential Fill
                        st.text_input('Solder Vol(mm^3)', value=f"{volume_of_solder_per_joint:.2f}", key="volume_of_solder_per_joint", disabled=True) 
                        st.text_input('Solder Wt/Joint(g)', value=f"{weight_of_Solder_per_joint:.2f}", key="weight_of_Solder_per_joint", disabled=True) 
                        st.text_input('Solder Wt/Brd(g)', value=f"{weight_of_Solder_per_board:.2f}", key="weight_of_Solder_per_board", disabled=True) 

                    with barrel1_col:
                        st.subheader("")
                        # Input Fields                  
                        barrel_dia = st.text_input('Barrel Dia(mm)', value="", key="barrel_dia", disabled=False)
                        board_thick = st.text_input('Board Thick(mm)', value="", key="board_thick", disabled=False)
                        barrel_joints = st.text_input('Barrel Joints', value="", key="barrel_joints", disabled=False)
                        barrel_solder_thick = st.text_input('Barrel Solder Thick(mm)', value="", key="barrel_solder_thick", disabled=False)

                    solder_bar_cost_value = 0.024

                    try:
                        # Safely convert inputs to float
                        barrel_dia = float(barrel_dia) if barrel_dia else 0.0
                        board_thick = float(board_thick) if board_thick else 0.0
                        barrel_joints = float(barrel_joints) if barrel_joints else 0.0
                        barrel_solder_thick = float(barrel_solder_thick) if barrel_solder_thick else 0.0
                        solder_bar_cost_value = float(solder_bar_cost_value) if solder_bar_cost_value else 0.0

                        # Calculate the volume of solder per joint - Barrel Fill

                        # Calculate the Barrel Solder Vol
                        barrel_solder_vol = (
                            (math.pi * (barrel_dia ** 2) / 4) -
                            (math.pi * ((barrel_dia - 2 * barrel_solder_thick) ** 2) / 4)
                        ) * board_thick

                        # Calculation of Barrel Solder Weight per Joint
                        barrel_solder_wt_per_joint = (barrel_solder_vol / 1000) * specific_gravity_of_solder  # Convert mm³ to cm³

                        # Calculation of Barrel Solder Weight per Board
                        barrel_solder_wt_per_board = barrel_solder_wt_per_joint * barrel_joints

                        circumferential_plus_barrel_fill_solder_wt = weight_of_Solder_per_board + barrel_solder_wt_per_board
                        solderbar_cost_per_brd = circumferential_plus_barrel_fill_solder_wt * solder_bar_cost_value

                    except ValueError:
                        st.error("Please enter valid numeric values for all inputs.")

                    with barrel1_col:
                        # Display the calculated value - Barrel Fill
                        st.text_input('Barrel Solder Vol(mm^3)', value=f"{barrel_solder_vol:.2f}", key="barrel_solder_vol", disabled=True) 
                        st.text_input('Barrel Solder Wt/Joint(g)', value=f"{barrel_solder_wt_per_joint:.4f}", key="barrel_solder_wt_per_joint", disabled=True) 
                        st.text_input('Barrel Solder Wt/Brd(g)', value=f"{barrel_solder_wt_per_board:.2f}", key="barrel_solder_wt_per_board", disabled=True) 
                    with barrel2_col:  
                        st.subheader("")                    
                        st.text_input('Total Solder Wt(g)', value=f"{circumferential_plus_barrel_fill_solder_wt:.2f}", key="circumferential_plus_barrel_fill_solder_wt", disabled=True) 
                        solder_bar_cost = st.text_input('Solder Bar Cost($/g)', value=solder_bar_cost_value, key="solder_bar_cost", disabled=True)
                        st.text_input('Solder Bar Cost/Brd($)', value=f"{solderbar_cost_per_brd:.2f}", key="solderbar_cost_per_brd", disabled=True) 

                    cost_consumables_value = (rtv_cost_per_board + top_side_cost_per_board_value + bot_side_cost_per_board_value + flux_cost_per_board_value + solderbar_cost_per_brd)


                st.header("RM & Conversion Cost Summary")                
                # Create one row with 4 columns for headings
                input_cost_col, ohpandother_percentage_col, ohpandother_cost_col, placeholder2_col = st.columns(4)
                
                with input_cost_col:
                    st.subheader("Input Cost")
                    cost_pcb = st.text_input('PCB ($)', value="", key="cost_pcb")
                    cost_electronics_components = st.text_input('Electronics Component ($)', value="", key="cost_electronics_components")
                    cost_mech_components = st.text_input('Mechanical Component ($)', value="", key="cost_mech_components")
                    cost_nre = st.text_input('NRE ($)', value=nre_per_unit, key="cost_nre", disabled=True)
                    # cost_consumables_value = (rtv_cost_per_board + top_side_cost_per_board_value + bot_side_cost_per_board_value + flux_cost_per_board_value + solderbar_cost_per_brd)
                    cost_consumables = st.text_input('Consumables ($)', value=cost_consumables_value, key="cost_consumables", disabled=True)
                    # Safely convert inputs to float
                    try:
                        cost_pcb = float(cost_pcb) if cost_pcb else 0.0
                        cost_electronics_components = float(cost_electronics_components) if cost_electronics_components else 0.0
                        cost_mech_components = float(cost_mech_components) if cost_mech_components else 0.0
                        cost_nre = float(cost_nre) if cost_nre else 0.0
                        cost_consumables = float(cost_consumables) if cost_consumables else 0.0
                    except ValueError:
                        st.error("Please enter valid numeric values for all cost fields.")
                        cost_pcb = 0.0
                        cost_electronics_components = 0.0
                        cost_mech_components = 0.0
                        cost_nre = 0.0
                        cost_consumables = 0.0


                with ohpandother_percentage_col:
                    st.subheader("OHP% Model Vs. Ann. Volume", )
                    # Define the predefined percentages as a dictionary
                    percentages = {
                        "<100K": {"MOH %": 1, "FOH %": 12.5, "Profit on RM %": 1.5, "Profit on VA %": 8, "R&D %": 1, "Warranty %": 1, "SG&A %": 3},
                        ">100K": {"MOH %": 1, "FOH %": 12.5, "Profit on RM %": 1, "Profit on VA %": 8, "R&D %": 2, "Warranty %": 1, "SG&A %": 2},
                        "5K/10K": {"MOH %": 1, "FOH %": 20, "Profit on RM %": 1, "Profit on VA %": 8, "R&D %": 3, "Warranty %": 1, "SG&A %": 4},
                    }

                    # Annual volume selection dropdown
                    annual_volume = st.selectbox("Select Annual Volume", options=["<100K", ">100K", "5K/10K"])

                    # Get the percentages for the selected annual volume
                    selected_percentages = percentages.get(annual_volume, {})

                    # Display the percentages dynamically in a single column
                    st.subheader(f"Percentages for {annual_volume}")

                    # Use st.columns to arrange the inputs in rows
                    percentage_labels_grouped = [
                        ["MOH %", "FOH %"],
                        ["Profit on RM %", "Profit on VA %"],
                        ["R&D %","Warranty %","SG&A %"],
                    ]
                    # Display percentage inputs dynamically in rows
                    percentage_values = {}
                    for row_labels in percentage_labels_grouped:
                        cols = st.columns(len(row_labels))
                        for i, label in enumerate(row_labels):
                            # Fetch the value for this label from selected_percentages, default to ""
                            value = selected_percentages.get(label, "")
                            # Show the value in a disabled text input
                            percentage_values[label] = cols[i].text_input(label, value=str(value), disabled=True)

                    # Aggregate dummy data
                    total_factory_overheads_batchsetup = sum(edited_data["Batch Set up Cost"])
                    total_factory_overheads_vamachine = sum(edited_data["VA MC Cost"])
                    total_factory_overheads_labour = sum(edited_data["Labour cost/Hr"])

                    # Convert percentages to fractions and compute costs
                    pcb_comp_mech_cost = cost_pcb + cost_electronics_components + cost_mech_components

                    moh_cost_value = pcb_comp_mech_cost * (selected_percentages["MOH %"] / 100)
                    foh_cost_value = (total_factory_overheads_batchsetup + total_factory_overheads_vamachine + total_factory_overheads_labour) * (selected_percentages["FOH %"] / 100)
                    profit_on_rm_cost_value = pcb_comp_mech_cost * (selected_percentages["Profit on RM %"] / 100)
                    profit_on_va_cost_value = (total_factory_overheads_batchsetup + total_factory_overheads_vamachine + total_factory_overheads_labour) * (selected_percentages["Profit on VA %"] / 100)

                    total_material_cost_value = pcb_comp_mech_cost + nre_per_unit
                    total_manufacturing_cost_value = total_factory_overheads_batchsetup + total_factory_overheads_vamachine + total_factory_overheads_labour
                    total_ohp_cost_value = moh_cost_value + foh_cost_value + profit_on_rm_cost_value + profit_on_va_cost_value

                    r_n_d_cost_value = (total_material_cost_value + total_manufacturing_cost_value) * (selected_percentages["R&D %"] / 100)
                    warranty_cost_value = (total_material_cost_value + total_manufacturing_cost_value) * (selected_percentages["Warranty %"] / 100)
                    sg_and_a_cost_value = (total_material_cost_value + total_manufacturing_cost_value) * (selected_percentages["SG&A %"] / 100)

                with ohpandother_cost_col:
                    st.subheader("Cost Computation")
                    ohpandother_cost_col1, ohpandother_cost_col2 = st.columns(2)

                    with ohpandother_cost_col1:
                        st.text_input("MOH ($)", value=f"{moh_cost_value:.2f}", disabled=True)
                        st.text_input("Profit on RM ($)", value=f"{profit_on_rm_cost_value:.2f}", disabled=True)
                        st.text_input("Material Cost ($)", value=f"{total_material_cost_value:.2f}", disabled=True)
                        st.text_input("OH&P ($)", value=f"{total_ohp_cost_value:.2f}", disabled=True)
                        st.text_input("Warranty ($)", value=f"{warranty_cost_value:.2f}", disabled=True)
                    with ohpandother_cost_col2:
                        st.text_input("FOH ($)", value=f"{foh_cost_value:.2f}", disabled=True)
                        st.text_input("Profit on VA ($)", value=f"{profit_on_va_cost_value:.2f}", disabled=True)
                        st.text_input("Manufacturing Cost ($)", value=f"{total_manufacturing_cost_value:.2f}", disabled=True)
                        st.text_input("R&D ($)", value=f"{r_n_d_cost_value:.2f}", disabled=True)
                        st.text_input("SG&A ($)", value=f"{sg_and_a_cost_value:.2f}", disabled=True)

                with placeholder2_col:
                    st.subheader("Cost Summary")
                    grand_total_cost_value = ((total_material_cost_value + total_manufacturing_cost_value) + moh_cost_value + 
                                                    foh_cost_value + profit_on_rm_cost_value + profit_on_va_cost_value +
                                                    r_n_d_cost_value + warranty_cost_value + sg_and_a_cost_value )
                    st.text_input('Total Cost ($)', value=grand_total_cost_value, disabled=True)

                    rm_cost_value = total_material_cost_value
                    st.text_input('RM Cost ($)', value=rm_cost_value, disabled=True)

                    conversion_cost_value = grand_total_cost_value - total_material_cost_value
                    st.text_input('Conversion Cost ($)', value=conversion_cost_value, disabled=True)

                    if st.button("Save Consumable, RM & Conversion Costing Details"):
                        if sheet_name in st.session_state.edited_sheets:
                            # Retrieve the current edited_data
                            current_data = st.session_state.edited_sheets[sheet_name].copy()

                            # Add new columns if they don't exist
                            columns_to_add = [
                                'RTV Wt/Brd Est','RTV Wastage %','RTV Cost/ml','RTV Solder SG','Wt per Board (Incl Wastage %)','RTV Cost Per Board', #RTV Glue section
                                "Board Length(mm)","Board Width(mm)","Top Wt Estimate %","Top Wastage %","Solder Paste SG (g/cc)","Solder Paste Cost($/g)","Top SP Thick(mm)","Top SP Wt (100%)(g)", "Top SP Wt Estimate(g)","Top SP Cost/Brd($)", #Solder Paste - Top section
                                "Bot Wt Estimate %","Bot Wastage %","Bot SP Thick(mm)","Bot SP Wt (100%)(g)","Bot SP Wt Estimate(g)","Bot SP Cost/Brd($)",  #Solder Paste - Bottom section
                                "Flux Wastage %","Flux Cost($/ml)","Flux Area/Brd(mm^2)","Flux Spray Area(mm^2)","Flux Cost Per Board($)", #Flux Wave Soldering section
                                "Pad OD (mm)","Pad ID (mm)","Solder Joints","Solder Thick (mm)","Solder Vol(mm^3)","Solder Wt/Joint(g)","Solder Wt/Brd(g)", # Circumferential Fill Solder Bar
                                "Barrel Dia(mm)","Board Thick(mm)","Barrel Joints","Barrel Solder Thick(mm)","Solder Bar Cost($/g)","Barrel Solder Vol(mm^3)","Barrel Solder Wt/Joint(g)","Barrel Solder Wt/Brd(g)","Solder Bar Cost($/g)","Total Solder Wt(g)","Solder Bar Cost/Brd($)", # Barrel Fill Solder Bar
                                "PCB ($)","Electronics Component ($)","Mechanical Component ($)","NRE ($)","Consumables ($)", #Input Cost section
                                "Select Annual Volume","MOH %","FOH %","Profit on RM %","Profit on VA %","R&D %","Warranty %","SG&A %", #OHP% Model Vs. Ann. Volume section
                                "MOH ($)","Profit on RM ($)","FOH ($)","Profit on VA ($)","Material Cost ($)","Manufacturing Cost ($)","OH&P ($)","R&D ($)","Warranty ($)","SG&A ($)", #Cost Computation section
                                "Total Cost ($)","RM Cost ($)","Conversion Cost ($)" # Cost Summary section
                            ]
                            for column in columns_to_add:
                                if column not in current_data.columns:
                                    current_data[column] = np.nan

                            # Assign values to the first row of respective columns
                            current_data.loc[0, 'RTV Wt/Brd Est'] = glue_wt_per_board
                            current_data.loc[0, 'RTV Wastage %'] = wastage_percentage_per_board
                            current_data.loc[0, 'RTV Cost/ml'] = rtv_glue_cost
                            current_data.loc[0, 'RTV Solder SGRTV Solder SG'] = specific_gravity_of_solder
                            current_data.loc[0, 'Wt per Board (Incl Wastage %)'] = wt_per_board_incl_wastage
                            current_data.loc[0, 'RTV Cost Per Board'] = rtv_cost_per_board
                            current_data.loc[0, "Board Length(mm)"] = board_length  
                            current_data.loc[0, "Board Width(mm)"] = board_width  
                            current_data.loc[0, "Top Wt Estimate %"] = top_weight_estimate_percentage
                            current_data.loc[0, "Top Wastage %"] = top_sp_wastage_percentage
                            current_data.loc[0, "Solder Paste SG (g/cc)"] = paste_specific_gravity
                            current_data.loc[0, "Solder Paste Cost($/g)"] = cost_of_solder_paste
                            current_data.loc[0, "Top SP Thick(mm)"] = solder_paste_thickness
                            current_data.loc[0, "Top SP Wt (100%)(g)"] = weight_of_solder_paste_for_100percentage_wt_value
                            current_data.loc[0, "Top SP Wt Estimate(g)"] = top_weight_of_solder_paste_for_wt_estimate_value
                            current_data.loc[0, "Top SP Cost/Brd($)"] = top_side_cost_per_board_value
                            current_data.loc[0, "Bot Wt Estimate %"] = bot_weight_estimate_percentage
                            current_data.loc[0, "Bot Wastage %"] = bot_sp_wastage_percentage
                            current_data.loc[0, "Bot SP Thick(mm)"] = bot_solder_paste_thickness
                            current_data.loc[0, "Bot SP Wt (100%)(g)"] = bot_weight_of_solder_paste_for_100percentage_wt_value
                            current_data.loc[0, "Bot SP Wt Estimate(g)"] = bot_weight_of_solder_paste_for_wt_estimate_value
                            current_data.loc[0, "Bot SP Cost/Brd($)"] = bot_side_cost_per_board_value
                            current_data.loc[0, "Flux Wastage %"] = flux_wastage_percentage
                            current_data.loc[0, "Flux Cost($/ml)"] = flux_cost
                            current_data.loc[0, "Flux Area/Brd(mm^2)"] = flux_board_area_value
                            current_data.loc[0, "Flux Spray Area(mm^2)"] = flux_spread_area_value
                            current_data.loc[0, "Flux Cost Per Board($)"] = flux_cost_per_board
                            current_data.loc[0, "Pad OD (mm)"] = outer_dia_of_pad
                            current_data.loc[0, "Pad ID (mm)"] = inner_dia_of_pad
                            current_data.loc[0, "Solder Joints"] = no_of_solder_joints
                            current_data.loc[0, "Solder Thick (mm)"] = thickness_of_solder
                            current_data.loc[0, "Solder Vol(mm^3)"] = volume_of_solder_per_joint
                            current_data.loc[0, "Solder Wt/Joint(g)"] = weight_of_Solder_per_joint
                            current_data.loc[0, "Solder Wt/Brd(g)"] = weight_of_Solder_per_board
                            current_data.loc[0, "Barrel Dia(mm)"] = barrel_dia
                            current_data.loc[0, "Board Thick(mm)"] = board_thick
                            current_data.loc[0, "Barrel Joints"] = barrel_joints
                            current_data.loc[0, "Barrel Solder Thick(mm)"] = barrel_solder_thick
                            current_data.loc[0, "Solder Bar Cost($/g)"] = solder_bar_cost_value
                            current_data.loc[0, "Barrel Solder Vol(mm^3)"] = barrel_solder_vol
                            current_data.loc[0, "Barrel Solder Wt/Joint(g)"] = barrel_solder_wt_per_joint
                            current_data.loc[0, "Barrel Solder Wt/Brd(g)"] = barrel_solder_wt_per_board
                            current_data.loc[0, "Solder Bar Cost($/g)"] = solder_bar_cost
                            current_data.loc[0, "Total Solder Wt(g)"] = circumferential_plus_barrel_fill_solder_wt
                            current_data.loc[0, "Solder Bar Cost/Brd($)"] = solderbar_cost_per_brd
                            current_data.loc[0, "PCB ($)"] = cost_pcb
                            current_data.loc[0, "Electronics Component ($)"] = cost_electronics_components
                            current_data.loc[0, "Mechanical Component ($)"] = cost_mech_components
                            current_data.loc[0, "NRE ($)"] = cost_nre
                            current_data.loc[0, "Consumables ($)"] = cost_consumables
                            current_data.loc[0, "Select Annual Volume"] = annual_volume

                            # Update the values in `edited_data` (current_data) for these headers
                            for label, value in percentage_values.items():
                                try:
                                    # Convert value to float and update the data
                                    current_data.loc[0, label] = float(value)
                                except ValueError:
                                    st.warning(f"Invalid value for {label}. Please check the input.")

                            current_data.loc[0, "MOH ($)"] = moh_cost_value
                            current_data.loc[0, "FOH ($)"] = foh_cost_value
                            current_data.loc[0, "Profit on RM ($)"] = profit_on_rm_cost_value
                            current_data.loc[0, "Profit on VA ($)"] = profit_on_va_cost_value
                            current_data.loc[0, "Material Cost ($)"] = total_material_cost_value
                            current_data.loc[0, "Manufacturing Cost ($)"] = total_manufacturing_cost_value
                            current_data.loc[0, "OH&P ($)"] = total_ohp_cost_value
                            current_data.loc[0, "R&D ($)"] = r_n_d_cost_value
                            current_data.loc[0, "Warranty ($)"] = warranty_cost_value
                            current_data.loc[0, "SG&A ($)"] = sg_and_a_cost_value
                            current_data.loc[0, "Total Cost ($)"] = grand_total_cost_value
                            current_data.loc[0, "RM Cost ($)"] = rm_cost_value
                            current_data.loc[0, "Conversion Cost ($)"] = conversion_cost_value

                        if sheet_name:  # Check if the sheet_name is not empty
                            # Update the session state
                            st.session_state.edited_sheets[sheet_name] = current_data
                            st.success("Consumable, RM & Conversion Costing Details saved successfully.")
                        else:
                            st.error("No data available to save Consumable, RM & Conversion Costing Details.")
                
                # Update the data editor with the latest data
                edited_data2 = st.session_state.edited_sheets.get(sheet_name, pd.DataFrame())

                # Generate a unique key for the data editor widget
                unique_key = f"data_editor_{sheet_name}_{uuid.uuid4()}"

                # Display the data editor
                edited_data2 = st.data_editor(edited_data2, key=unique_key)

                # Option to download the data using BytesIO
                if not edited_data2.empty:
                    # Serialize the edited_data2 DataFrame into an Excel file
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        edited_data2.to_excel(writer, sheet_name=sheet_name, index=False)
                    output.seek(0)  # Reset the buffer position to the beginning

                    # Provide download button for the serialized data
                    st.download_button(
                        label="Download Excel file",
                        data=output,
                        file_name=f"{sheet_name}_Costing_Details.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("No data available for download.")


                
                st.write("-------------------")


                # Data for the pie chart
                labels = ['RM Cost', 'Conversion Cost']
                values = [rm_cost_value, conversion_cost_value]

                # Create the pie chart
                fig = go.Figure(data=[go.Pie(labels=labels, values=values, hole=0.4)])

                # Add title and adjust layout
                fig.update_layout(
                    title_text='RM Cost vs Conversion Cost % Visualization',
                    height=600,  # Adjust height as needed
                    width=800,   # Adjust width as needed
                    margin=dict(t=50, b=50, l=50, r=50)  # Adjust margins as needed
                )

                # Streamlit layout to display the pie chart
                graph_col1, graph_col2 = st.columns(2)

                with graph_col2:
                    st.plotly_chart(fig, use_container_width=True)
                
                # Horizontal Bar Chart 
                data = {
                    "Cost Component": [
                        "Material Cost ($)",
                        "Manufacturing Cost ($)",
                        "OH&P ($)",
                        "R&D ($)",
                        "Warranty ($)",
                        "SG&A ($)",
                        "Total Cost ($)"
                    ],
                    "Amount": [
                        total_material_cost_value,
                        total_manufacturing_cost_value,
                        total_ohp_cost_value,
                        r_n_d_cost_value,
                        warranty_cost_value,
                        sg_and_a_cost_value,
                        grand_total_cost_value
                    ]
                }

                # Convert to DataFrame
                df = pd.DataFrame(data)

                # Horizontal Bar Chart using Plotly Express
                fig = px.bar(
                    df,
                    y="Cost Component",
                    x="Amount",
                    orientation="h",
                    title="Cost Breakdown",
                    color="Cost Component",  # Optional: adds color for each component
                    text="Amount"  # Show values on bars
                )

                fig.update_traces(texttemplate='%{text:.2f}', textposition='outside')
                fig.update_layout(xaxis_title="Cost Amount ($)", 
                                    height=600,  # Adjust height as needed
                                    width=800,   # Adjust width as needed
                                    margin=dict(t=50, b=50, l=50, r=50),  # Adjust margins as needed
                                  yaxis_title="Cost Component")

                # Display Chart in Streamlit
                with graph_col1:
                    st.plotly_chart(fig)


# Example to show how the existing_analysis would be implemented
if existing_analysis:
    st.subheader("Existing Analysis")

    upload_nre_file_col, upload_costing_file_col = st.columns(2)

    with upload_nre_file_col:
        # File uploader for Excel/CSV/XLSM files
        uploaded_file = st.file_uploader("Choose NRE Costing Excel/CSV/XLSM file", type=["xlsx", "csv", "xlsm"])

        if uploaded_file:
            # Load data from the uploaded file
            @st.cache_data
            def load_data(file):
                if file.name.endswith('.csv'):
                    df = pd.read_csv(file)
                else:
                    df = pd.read_excel(file, sheet_name=None)
                return df

            df7 = load_data(uploaded_file)  # Load the Process Mapping data

            # Initialize session state to store edited data for each sheet
            if 'edited_sheets' not in st.session_state:
                st.session_state.edited_sheets = {}

            processmapping_col1, _ = st.columns(2)

            if isinstance(df7, dict):
                with processmapping_col1:
                    # Allow user to select a sheet
                    sheet_name = st.selectbox("Select the relavant NRE sheet", df7.keys())

                if sheet_name in df7:
                    nre_selected_data = df7[sheet_name]
                    st.session_state.df = pd.DataFrame(nre_selected_data)  # Load original data from the selected sheet

        st.subheader("NRE Costing Data Table")
        # Display the updated DataFrame
        st.data_editor(nre_selected_data, key=f"data_editor_{sheet_name}_nre_updated")

    with upload_costing_file_col:
        # File uploader for Excel/CSV/XLSM files
        uploaded_file = st.file_uploader("Choose Should Costing Excel/CSV/XLSM file", type=["xlsx", "csv", "xlsm"])

        if uploaded_file:
            # Load data from the uploaded file
            @st.cache_data
            def load_data(file):
                if file.name.endswith('.csv'):
                    df = pd.read_csv(file)
                else:
                    df = pd.read_excel(file, sheet_name=None)
                return df

            df8 = load_data(uploaded_file)  # Load the Process Mapping data

            # Initialize session state to store edited data for each sheet
            if 'edited_sheets' not in st.session_state:
                st.session_state.edited_sheets = {}

            processmapping_col2, _ = st.columns(2)

            if isinstance(df8, dict):
                with processmapping_col2:
                    # Allow user to select a sheet
                    sheet_name = st.selectbox("Select the relavant Should Costing sheet", df8.keys())

                if sheet_name in df8:
                    selected_data = df8[sheet_name]
                    st.session_state.df = pd.DataFrame(selected_data)  # Load original data from the selected sheet

        st.subheader("Should Costing Data Table")
        # Display the updated DataFrame
        st.data_editor(selected_data, key=f"data_editor_{sheet_name}_shc_updated")
    st.session_state.edited_sheets[sheet_name] = selected_data

    st.header("Consumable Costing")
    

    # Create one row with 4 columns for headings
    rtv_col, solder_top_col, solder_bar_col = st.columns(3)

    # RTV Glue
    with rtv_col:
        st.subheader("RTV Glue")
        rtv_1_col, rtv_2_col = st.columns(2)
        with rtv_1_col:
        # Input Fields
            glue_wt_per_board = st.text_input('RTV Wt/Brd Est', value="", key="glue_wt_per_board")
            rtv_glue_cost = st.text_input('RTV Cost/ml', value="0.052", key="rtv_glue_cost", disabled=True)
        
        with rtv_2_col:
            wastage_percentage_per_board = st.text_input('RTV Wastage %', value="", key="wastage_percentage_per_board")
            specific_gravity_of_solder = st.text_input('RTV Solder SG', value="1.09", key="specific_gravity_of_solder", disabled=True)

        # Initialize the output variables
        wt_per_board_incl_wastage = 0.0
        rtv_cost_per_board = 0.0

        try:
            # Safely convert inputs to float
            glue_wt_per_board = float(glue_wt_per_board) if glue_wt_per_board else 0.0
            wastage_percentage_per_board = float(wastage_percentage_per_board) if wastage_percentage_per_board else 0.0
            rtv_glue_cost = float(rtv_glue_cost) if rtv_glue_cost else 0.0
            specific_gravity_of_solder = float(specific_gravity_of_solder) if specific_gravity_of_solder else 0.0

            # Calculate wt_per_board_incl_wastage
            wt_per_board_incl_wastage = (
                (glue_wt_per_board * specific_gravity_of_solder) * (1 + (wastage_percentage_per_board / 100))
            )

            # Calculate rtv_cost_per_board
            rtv_cost_per_board = wt_per_board_incl_wastage * rtv_glue_cost

        except ValueError:
            st.error("Please enter valid numeric values for all inputs.")

        # Display results
        with rtv_1_col:
            st.text_input('RTV Wt/Brd', value=f"{wt_per_board_incl_wastage:.2f}", key="wt_per_board_incl_wastage", disabled=True)
        # with rtv_2_col:
        st.text_input('RTV Cost Per Board', value=f"{rtv_cost_per_board:.2f}", key="rtv_cost_per_board", disabled=True)

    # Solder Paste - Top
    with solder_top_col:
        st.subheader("Solder Paste - Top & Bot")
        # Board dimensions (Board Length and Board Width in the same line)
        board_dim_col1, board_dim_col2 = st.columns(2)
        with board_dim_col1:
            board_length = st.text_input('Board Length(mm)', value="", key="board_length")
        with board_dim_col2:
            # st.write("Solder Paste - Bot")
            board_width = st.text_input('Board Width(mm)', value="", key="board_width")

        sp_sg_col, sp_cost_col = st.columns(2)
        with sp_sg_col:
            paste_specific_gravity = st.text_input('Solder Paste SG (g/cc)', value="7.31", key="paste_specific_gravity", disabled=True)
        with sp_cost_col:
            cost_of_solder_paste = st.text_input('Solder Paste Cost($/g)', value="0.065", key="cost_of_solder_paste", disabled=True)

        sp_top_col1, sp_bot_col2 = st.columns(2)
        with sp_top_col1:
            top_weight_estimate_percentage = st.text_input('Top Wt Estimate %', value="", key="top_weight_estimate_percentage")
            top_sp_wastage_percentage = st.text_input('Top Wastage %', value="", key="top_sp_wastage_percentage")
            solder_paste_thickness = st.text_input('Top SP Thick(mm)', value="", key="solder_paste_thickness")

        try:
            # Safely convert inputs to float
            board_length = float(board_length) if board_length else 0.0
            board_width = float(board_width) if board_width else 0.0
            top_weight_estimate_percentage = float(top_weight_estimate_percentage) if top_weight_estimate_percentage else 0.0
            top_sp_wastage_percentage = float(top_sp_wastage_percentage) if top_sp_wastage_percentage else 0.0
            paste_specific_gravity = float(paste_specific_gravity) if paste_specific_gravity else 7.31
            cost_of_solder_paste = float(cost_of_solder_paste) if cost_of_solder_paste else 0.065
            solder_paste_thickness = float(solder_paste_thickness) if solder_paste_thickness else 0.0

            # Calculate weight_of_solder_paste_for_100percentage_wt
            weight_of_solder_paste_for_100percentage_wt_value = (
                (board_length * board_width * solder_paste_thickness * paste_specific_gravity) / 1000
            )

            # Convert percentages to fractions
            top_weight_estimate_percentage /= 100
            top_sp_wastage_percentage /= 100

            # Calculate top_solder_paste_weight_estimate
            top_weight_of_solder_paste_for_wt_estimate_value = (
                weight_of_solder_paste_for_100percentage_wt_value
                * top_weight_estimate_percentage
                * (1 + top_sp_wastage_percentage)
            )

            # Calculate top_side_cost_per_board
            top_side_cost_per_board_value = top_weight_of_solder_paste_for_wt_estimate_value * cost_of_solder_paste

        except ValueError:
            st.error("Please enter valid numeric values for all inputs.")
        with sp_top_col1:
            # Display results
            st.text_input('Top SP Wt (100%)(g)', value=f"{weight_of_solder_paste_for_100percentage_wt_value:.2f}", key="weight_of_solder_paste_for_100percentage_wt", disabled=True)
            st.text_input('Top SP Wt Estimate(g)', value=f"{top_weight_of_solder_paste_for_wt_estimate_value:.2f}", key="top_solder_paste_weight_estimate", disabled=True)
            st.text_input('Top SP Cost/Brd($)', value=f"{top_side_cost_per_board_value:.2f}", key="top_side_cost_per_board", disabled=True)                # Solder Paste - Bottom


    # Solder Paste - Bottom
    # with solder_bottom_col:
    with solder_top_col:
        # st.subheader("Solder Paste - Bottom")
        # Input Fields
        with sp_bot_col2:
            bot_weight_estimate_percentage = st.text_input('Bot Wt Estimate %', value="", key="bot_weight_estimate_percentage")
            bot_sp_wastage_percentage = st.text_input('Bot Wastage %', value="", key="bot_sp_wastage_percentage")
            bot_solder_paste_thickness = st.text_input('Bot SP Thick(mm)', value="", key="bot_solder_paste_thickness")

        try:
            # Safely convert inputs to float
            bot_weight_estimate_percentage = float(bot_weight_estimate_percentage) if bot_weight_estimate_percentage else 0.0
            bot_sp_wastage_percentage = float(bot_sp_wastage_percentage) if bot_sp_wastage_percentage else 0.0
            bot_solder_paste_thickness = float(bot_solder_paste_thickness) if bot_solder_paste_thickness else 0.0

            # Calculate weight_of_solder_paste_for_100percentage_wt
            bot_weight_of_solder_paste_for_100percentage_wt_value = (
                (board_length * board_width * bot_solder_paste_thickness * paste_specific_gravity) / 1000
            )

            # Convert percentages to fractions
            bot_weight_estimate_percentage /= 100
            bot_sp_wastage_percentage /= 100

            # Compute bot_weight_of_solder_paste_for_wt_estimate_value
            bot_weight_of_solder_paste_for_wt_estimate_value = (
                bot_weight_of_solder_paste_for_100percentage_wt_value
                * bot_weight_estimate_percentage
                * (1 + bot_sp_wastage_percentage)
            )

            # Compute bot_side_cost_per_board
            bot_side_cost_per_board_value = bot_weight_of_solder_paste_for_wt_estimate_value * cost_of_solder_paste

        except ValueError:
            st.error("Please enter valid numeric values for all inputs.")
        with sp_bot_col2:
            # Display results
            st.text_input('Bot SP Wt (100%)(g)', value=f"{bot_weight_of_solder_paste_for_100percentage_wt_value:.2f}", key="bot_weight_of_solder_paste_for_100percentage_wt_value", disabled=True)
            st.text_input('Bot SP Wt Estimate(g)', value=f"{bot_weight_of_solder_paste_for_wt_estimate_value:.2f}", key="bot_solder_paste_weight_estimate", disabled=True)
            st.text_input('Bot SP Cost/Brd($)', value=f"{bot_side_cost_per_board_value:.2f}", key="bot_side_cost_per_board", disabled=True)

    # Flux Wave Soldering
    with rtv_col:
        st.subheader("Flux Wave Soldering")
        rtv_3_col, rtv_4_col = st.columns(2)

        # Input Fields
        with rtv_3_col:
            flux_wastage_percentage = st.text_input('Flux Wastage %', value="", key="flux_wastage_percentage")
        with rtv_4_col:
            flux_cost = st.text_input('Flux Cost($/ml)', value="0.0055", key="flux_cost", disabled=True)

        try:
            # Safely convert inputs to float
            flux_wastage_percentage = float(flux_wastage_percentage) if flux_wastage_percentage else 0.0
            flux_cost = float(flux_cost) if flux_cost else 0.0055

            # Compute flux_board_area
            flux_board_area_value = board_length * board_width

            # Convert percentage to fraction
            flux_wastage_percentage /= 100

            # Compute flux_spread_area
            flux_spread_area_value = ((flux_board_area_value / 100) * 0.1) * (1 + flux_wastage_percentage)

        except ValueError:
            st.error("Please enter valid numeric values for all inputs.")

        # Display results
        with rtv_3_col:
            st.text_input('Flux Area/Brd(mm^2)', value=f"{flux_board_area_value:.2f}", key="flux_board_area", disabled=True)
        with rtv_4_col:
            st.text_input('Flux Spray Area(mm^2)', value=f"{flux_spread_area_value:.2f}", key="flux_spread_area", disabled=True)
        flux_cost_per_board_value = flux_spread_area_value * flux_cost
        flux_cost_per_board = st.text_input('Flux Cost Per Board($)', value=flux_cost_per_board_value, key="flux_cost_per_board", disabled=True)

    # Solder bar
    with solder_bar_col:
        circumferential_col, barrel1_col, barrel2_col = st.columns(3)
        with circumferential_col:
            st.subheader("Solder Bar")
            # Input Fields
            outer_dia_of_pad = st.text_input('Pad OD (mm)', value="", key="outer_dia_of_pad", disabled=False)
            inner_dia_of_pad = st.text_input('Pad ID (mm)', value="", key="inner_dia_of_pad", disabled=False)
            no_of_solder_joints = st.text_input('Solder Joints', value="", key="no_of_solder_joints", disabled=False)
            thickness_of_solder = st.text_input('Solder Thick (mm)', value="0.6", key="thickness_of_solder", disabled=True)

            try:
                # Safely convert inputs to float
                outer_dia_of_pad = float(outer_dia_of_pad) if outer_dia_of_pad else 0.0
                inner_dia_of_pad = float(inner_dia_of_pad) if inner_dia_of_pad else 0.0
                no_of_solder_joints = float(no_of_solder_joints) if no_of_solder_joints else 0.0
                thickness_of_solder = float(thickness_of_solder) if thickness_of_solder else 0.0

                # Calculate the area of the annular ring
                area_of_ring = math.pi * ((outer_dia_of_pad/2)**2 - (inner_dia_of_pad/2)**2)

                # Calculate the volume of solder per joint - Circumferential Fill
                volume_of_solder_per_joint = area_of_ring * thickness_of_solder
                weight_of_Solder_per_joint = (volume_of_solder_per_joint/1000) * paste_specific_gravity
                weight_of_Solder_per_board = weight_of_Solder_per_joint * no_of_solder_joints

            except ValueError:
                st.error("Please enter valid numeric values for all inputs.")

            # Display the calculated value - Circumferential Fill
            st.text_input('Solder Vol(mm^3)', value=f"{volume_of_solder_per_joint:.2f}", key="volume_of_solder_per_joint", disabled=True) 
            st.text_input('Solder Wt/Joint(g)', value=f"{weight_of_Solder_per_joint:.2f}", key="weight_of_Solder_per_joint", disabled=True) 
            st.text_input('Solder Wt/Brd(g)', value=f"{weight_of_Solder_per_board:.2f}", key="weight_of_Solder_per_board", disabled=True) 

        with barrel1_col:
            st.subheader("")
            # Input Fields                  
            barrel_dia = st.text_input('Barrel Dia(mm)', value="", key="barrel_dia", disabled=False)
            board_thick = st.text_input('Board Thick(mm)', value="", key="board_thick", disabled=False)
            barrel_joints = st.text_input('Barrel Joints', value="", key="barrel_joints", disabled=False)
            barrel_solder_thick = st.text_input('Barrel Solder Thick(mm)', value="", key="barrel_solder_thick", disabled=False)

        solder_bar_cost_value = 0.024

        try:
            # Safely convert inputs to float
            barrel_dia = float(barrel_dia) if barrel_dia else 0.0
            board_thick = float(board_thick) if board_thick else 0.0
            barrel_joints = float(barrel_joints) if barrel_joints else 0.0
            barrel_solder_thick = float(barrel_solder_thick) if barrel_solder_thick else 0.0
            solder_bar_cost_value = float(solder_bar_cost_value) if solder_bar_cost_value else 0.0

            # Calculate the volume of solder per joint - Barrel Fill

            # Calculate the Barrel Solder Vol
            barrel_solder_vol = (
                (math.pi * (barrel_dia ** 2) / 4) -
                (math.pi * ((barrel_dia - 2 * barrel_solder_thick) ** 2) / 4)
            ) * board_thick

            # Calculation of Barrel Solder Weight per Joint
            barrel_solder_wt_per_joint = (barrel_solder_vol / 1000) * specific_gravity_of_solder  # Convert mm³ to cm³

            # Calculation of Barrel Solder Weight per Board
            barrel_solder_wt_per_board = barrel_solder_wt_per_joint * barrel_joints

            circumferential_plus_barrel_fill_solder_wt = weight_of_Solder_per_board + barrel_solder_wt_per_board
            solderbar_cost_per_brd = circumferential_plus_barrel_fill_solder_wt * solder_bar_cost_value

        except ValueError:
            st.error("Please enter valid numeric values for all inputs.")

        with barrel1_col:
            # Display the calculated value - Barrel Fill
            st.text_input('Barrel Solder Vol(mm^3)', value=f"{barrel_solder_vol:.2f}", key="barrel_solder_vol", disabled=True) 
            st.text_input('Barrel Solder Wt/Joint(g)', value=f"{barrel_solder_wt_per_joint:.4f}", key="barrel_solder_wt_per_joint", disabled=True) 
            st.text_input('Barrel Solder Wt/Brd(g)', value=f"{barrel_solder_wt_per_board:.2f}", key="barrel_solder_wt_per_board", disabled=True) 
        with barrel2_col:  
            st.subheader("")                    
            st.text_input('Total Solder Wt(g)', value=f"{circumferential_plus_barrel_fill_solder_wt:.2f}", key="circumferential_plus_barrel_fill_solder_wt", disabled=True) 
            solder_bar_cost = st.text_input('Solder Bar Cost($/g)', value=solder_bar_cost_value, key="solder_bar_cost", disabled=True)
            st.text_input('Solder Bar Cost/Brd($)', value=f"{solderbar_cost_per_brd:.2f}", key="solderbar_cost_per_brd", disabled=True) 

        nre_per_unit = nre_selected_data.at[0, "NRE Per Unit ($)"]


    st.header("RM & Conversion Cost Summary")                
    # Create one row with 4 columns for headings
    input_cost_col, ohpandother_percentage_col, ohpandother_cost_col, placeholder2_col = st.columns(4)
    
    with input_cost_col:
        st.subheader("Input Cost")
        cost_pcb = st.text_input('PCB ($)', value="", key="cost_pcb")
        cost_electronics_components = st.text_input('Electronics Component ($)', value="", key="cost_electronics_components")
        cost_mech_components = st.text_input('Mechanical Component ($)', value="", key="cost_mech_components")
        cost_nre = st.text_input('NRE ($)', value=nre_per_unit, key="cost_nre", disabled=True)
        cost_consumables_value = (rtv_cost_per_board + top_side_cost_per_board_value + bot_side_cost_per_board_value + flux_cost_per_board_value)
        cost_consumables = st.text_input('Consumables ($)', value=cost_consumables_value, key="cost_consumables", disabled=True)
        # Safely convert inputs to float
        try:
            cost_pcb = float(cost_pcb) if cost_pcb else 0.0
            cost_electronics_components = float(cost_electronics_components) if cost_electronics_components else 0.0
            cost_mech_components = float(cost_mech_components) if cost_mech_components else 0.0
            cost_nre = float(cost_nre) if cost_nre else 0.0
            cost_consumables = float(cost_consumables) if cost_consumables else 0.0
        except ValueError:
            st.error("Please enter valid numeric values for all cost fields.")
            cost_pcb = 0.0
            cost_electronics_components = 0.0
            cost_mech_components = 0.0
            cost_nre = 0.0
            cost_consumables = 0.0


    with ohpandother_percentage_col:
        st.subheader("OHP% Model Vs. Ann. Volume", )
        # Define the predefined percentages as a dictionary
        percentages = {
            "<100K": {"MOH %": 1, "FOH %": 12.5, "Profit on RM %": 1.5, "Profit on VA %": 8, "R&D %": 1, "Warranty %": 1, "SG&A %": 3},
            ">100K": {"MOH %": 1, "FOH %": 12.5, "Profit on RM %": 1, "Profit on VA %": 8, "R&D %": 2, "Warranty %": 1, "SG&A %": 2},
            "5K/10K": {"MOH %": 1, "FOH %": 20, "Profit on RM %": 1, "Profit on VA %": 8, "R&D %": 3, "Warranty %": 1, "SG&A %": 4},
        }

        # Annual volume selection dropdown
        annual_volume = st.selectbox("Select Annual Volume", options=["<100K", ">100K", "5K/10K"])

        # Get the percentages for the selected annual volume
        selected_percentages = percentages.get(annual_volume, {})

        # Display the percentages dynamically in a single column
        st.subheader(f"Percentages for {annual_volume}")

        # Use st.columns to arrange the inputs in rows
        percentage_labels_grouped = [
            ["MOH %", "FOH %"],
            ["Profit on RM %", "Profit on VA %"],
            ["R&D %","Warranty %","SG&A %"],
        ]
        # Display percentage inputs dynamically in rows
        percentage_values = {}
        for row_labels in percentage_labels_grouped:
            cols = st.columns(len(row_labels))
            for i, label in enumerate(row_labels):
                # Fetch the value for this label from selected_percentages, default to ""
                value = selected_percentages.get(label, "")
                # Show the value in a disabled text input
                percentage_values[label] = cols[i].text_input(label, value=str(value), disabled=True)

        # Aggregate dummy data
        total_factory_overheads_batchsetup = sum(selected_data["Batch Set up Cost"])
        total_factory_overheads_vamachine = sum(selected_data["VA MC Cost"])
        total_factory_overheads_labour = sum(selected_data["Labour cost/Hr"])

        # Convert percentages to fractions and compute costs
        pcb_comp_mech_cost = cost_pcb + cost_electronics_components + cost_mech_components

        moh_cost_value = pcb_comp_mech_cost * (selected_percentages["MOH %"] / 100)
        foh_cost_value = (total_factory_overheads_batchsetup + total_factory_overheads_vamachine + total_factory_overheads_labour) * (selected_percentages["FOH %"] / 100)
        profit_on_rm_cost_value = pcb_comp_mech_cost * (selected_percentages["Profit on RM %"] / 100)
        profit_on_va_cost_value = (total_factory_overheads_batchsetup + total_factory_overheads_vamachine + total_factory_overheads_labour) * (selected_percentages["Profit on VA %"] / 100)

        total_material_cost_value = pcb_comp_mech_cost + nre_per_unit
        total_manufacturing_cost_value = total_factory_overheads_batchsetup + total_factory_overheads_vamachine + total_factory_overheads_labour
        total_ohp_cost_value = moh_cost_value + foh_cost_value + profit_on_rm_cost_value + profit_on_va_cost_value

        r_n_d_cost_value = (total_material_cost_value + total_manufacturing_cost_value) * (selected_percentages["R&D %"] / 100)
        warranty_cost_value = (total_material_cost_value + total_manufacturing_cost_value) * (selected_percentages["Warranty %"] / 100)
        sg_and_a_cost_value = (total_material_cost_value + total_manufacturing_cost_value) * (selected_percentages["SG&A %"] / 100)

    with ohpandother_cost_col:
        st.subheader("Cost Computation")
        ohpandother_cost_col1, ohpandother_cost_col2 = st.columns(2)

        with ohpandother_cost_col1:
            st.text_input("MOH ($)", value=f"{moh_cost_value:.2f}", disabled=True)
            st.text_input("Profit on RM ($)", value=f"{profit_on_rm_cost_value:.2f}", disabled=True)
            st.text_input("Material Cost ($)", value=f"{total_material_cost_value:.2f}", disabled=True)
            st.text_input("OH&P ($)", value=f"{total_ohp_cost_value:.2f}", disabled=True)
            st.text_input("Warranty ($)", value=f"{warranty_cost_value:.2f}", disabled=True)
        with ohpandother_cost_col2:
            st.text_input("FOH ($)", value=f"{foh_cost_value:.2f}", disabled=True)
            st.text_input("Profit on VA ($)", value=f"{profit_on_va_cost_value:.2f}", disabled=True)
            st.text_input("Manufacturing Cost ($)", value=f"{total_manufacturing_cost_value:.2f}", disabled=True)
            st.text_input("R&D ($)", value=f"{r_n_d_cost_value:.2f}", disabled=True)
            st.text_input("SG&A ($)", value=f"{sg_and_a_cost_value:.2f}", disabled=True)

    with placeholder2_col:
        st.subheader("Cost Summary")
        grand_total_cost_value = ((total_material_cost_value + total_manufacturing_cost_value) + moh_cost_value + 
                                        foh_cost_value + profit_on_rm_cost_value + profit_on_va_cost_value +
                                        r_n_d_cost_value + warranty_cost_value + sg_and_a_cost_value )
        st.text_input('Total Cost ($)', value=grand_total_cost_value, disabled=True)

        rm_cost_value = total_material_cost_value
        st.text_input('RM Cost ($)', value=rm_cost_value, disabled=True)

        conversion_cost_value = grand_total_cost_value - total_material_cost_value
        st.text_input('Conversion Cost ($)', value=conversion_cost_value, disabled=True)

        if st.button("Save Consumable, RM & Conversion Costing Details"):
            if sheet_name in st.session_state.edited_sheets:
                # Retrieve the current edited_data
                current_data = st.session_state.edited_sheets[sheet_name].copy()

                # Add new columns if they don't exist
                columns_to_add = [
                    'RTV Wt/Brd Est','RTV Wastage %','RTV Cost/ml','RTV Solder SG','Wt per Board (Incl Wastage %)','RTV Cost Per Board', #RTV Glue section
                    "Board Length(mm)","Board Width(mm)","Top Wt Estimate %","Top Wastage %","Solder Paste SG (g/cc)","Solder Paste Cost($/g)","Top SP Thick(mm)","Top SP Wt (100%)(g)", "Top SP Wt Estimate(g)","Top SP Cost/Brd($)", #Solder Paste - Top section
                    "Bot Wt Estimate %","Bot Wastage %","Bot SP Thick(mm)","Bot SP Wt (100%)(g)","Bot SP Wt Estimate(g)","Bot SP Cost/Brd($)",  #Solder Paste - Bottom section
                    "Flux Wastage %","Flux Cost($/ml)","Flux Area/Brd(mm^2)","Flux Spray Area(mm^2)","Flux Cost Per Board($)", #Flux Wave Soldering section
                    "Pad OD (mm)","Pad ID (mm)","Solder Joints","Solder Thick (mm)","Solder Vol(mm^3)","Solder Wt/Joint(g)","Solder Wt/Brd(g)", # Circumferential Fill Solder Bar
                    "Barrel Dia(mm)","Board Thick(mm)","Barrel Joints","Barrel Solder Thick(mm)","Solder Bar Cost($/g)","Barrel Solder Vol(mm^3)","Barrel Solder Wt/Joint(g)","Barrel Solder Wt/Brd(g)","Solder Bar Cost($/g)","Total Solder Wt(g)","Solder Bar Cost/Brd($)", # Barrel Fill Solder Bar
                    "PCB ($)","Electronics Component ($)","Mechanical Component ($)","NRE ($)","Consumables ($)", #Input Cost section
                    "Select Annual Volume","MOH %","FOH %","Profit on RM %","Profit on VA %","R&D %","Warranty %","SG&A %", #OHP% Model Vs. Ann. Volume section
                    "MOH ($)","Profit on RM ($)","FOH ($)","Profit on VA ($)","Material Cost ($)","Manufacturing Cost ($)","OH&P ($)","R&D ($)","Warranty ($)","SG&A ($)", #Cost Computation section
                    "Total Cost ($)","RM Cost ($)","Conversion Cost ($)" # Cost Summary section
                ]
                for column in columns_to_add:
                    if column not in current_data.columns:
                        current_data[column] = np.nan

                # Assign values to the first row of respective columns
                current_data.loc[0, 'RTV Wt/Brd Est'] = glue_wt_per_board
                current_data.loc[0, 'RTV Wastage %'] = wastage_percentage_per_board
                current_data.loc[0, 'RTV Cost/ml'] = rtv_glue_cost
                current_data.loc[0, 'RTV Solder SGRTV Solder SG'] = specific_gravity_of_solder
                current_data.loc[0, 'Wt per Board (Incl Wastage %)'] = wt_per_board_incl_wastage
                current_data.loc[0, 'RTV Cost Per Board'] = rtv_cost_per_board
                current_data.loc[0, "Board Length(mm)"] = board_length  
                current_data.loc[0, "Board Width(mm)"] = board_width  
                current_data.loc[0, "Top Wt Estimate %"] = top_weight_estimate_percentage
                current_data.loc[0, "Top Wastage %"] = top_sp_wastage_percentage
                current_data.loc[0, "Solder Paste SG (g/cc)"] = paste_specific_gravity
                current_data.loc[0, "Solder Paste Cost($/g)"] = cost_of_solder_paste
                current_data.loc[0, "Top SP Thick(mm)"] = solder_paste_thickness
                current_data.loc[0, "Top SP Wt (100%)(g)"] = weight_of_solder_paste_for_100percentage_wt_value
                current_data.loc[0, "Top SP Wt Estimate(g)"] = top_weight_of_solder_paste_for_wt_estimate_value
                current_data.loc[0, "Top SP Cost/Brd($)"] = top_side_cost_per_board_value
                current_data.loc[0, "Bot Wt Estimate %"] = bot_weight_estimate_percentage
                current_data.loc[0, "Bot Wastage %"] = bot_sp_wastage_percentage
                current_data.loc[0, "Bot SP Thick(mm)"] = bot_solder_paste_thickness
                current_data.loc[0, "Bot SP Wt (100%)(g)"] = bot_weight_of_solder_paste_for_100percentage_wt_value
                current_data.loc[0, "Bot SP Wt Estimate(g)"] = bot_weight_of_solder_paste_for_wt_estimate_value
                current_data.loc[0, "Bot SP Cost/Brd($)"] = bot_side_cost_per_board_value
                current_data.loc[0, "Flux Wastage %"] = flux_wastage_percentage
                current_data.loc[0, "Flux Cost($/ml)"] = flux_cost
                current_data.loc[0, "Flux Area/Brd(mm^2)"] = flux_board_area_value
                current_data.loc[0, "Flux Spray Area(mm^2)"] = flux_spread_area_value
                current_data.loc[0, "Flux Cost Per Board($)"] = flux_cost_per_board
                current_data.loc[0, "Pad OD (mm)"] = outer_dia_of_pad
                current_data.loc[0, "Pad ID (mm)"] = inner_dia_of_pad
                current_data.loc[0, "Solder Joints"] = no_of_solder_joints
                current_data.loc[0, "Solder Thick (mm)"] = thickness_of_solder
                current_data.loc[0, "Solder Vol(mm^3)"] = volume_of_solder_per_joint
                current_data.loc[0, "Solder Wt/Joint(g)"] = weight_of_Solder_per_joint
                current_data.loc[0, "Solder Wt/Brd(g)"] = weight_of_Solder_per_board
                current_data.loc[0, "Barrel Dia(mm)"] = barrel_dia
                current_data.loc[0, "Board Thick(mm)"] = board_thick
                current_data.loc[0, "Barrel Joints"] = barrel_joints
                current_data.loc[0, "Barrel Solder Thick(mm)"] = barrel_solder_thick
                current_data.loc[0, "Solder Bar Cost($/g)"] = solder_bar_cost_value
                current_data.loc[0, "Barrel Solder Vol(mm^3)"] = barrel_solder_vol
                current_data.loc[0, "Barrel Solder Wt/Joint(g)"] = barrel_solder_wt_per_joint
                current_data.loc[0, "Barrel Solder Wt/Brd(g)"] = barrel_solder_wt_per_board
                current_data.loc[0, "Solder Bar Cost($/g)"] = solder_bar_cost
                current_data.loc[0, "Total Solder Wt(g)"] = circumferential_plus_barrel_fill_solder_wt
                current_data.loc[0, "Solder Bar Cost/Brd($)"] = solderbar_cost_per_brd
                current_data.loc[0, "PCB ($)"] = cost_pcb
                current_data.loc[0, "Electronics Component ($)"] = cost_electronics_components
                current_data.loc[0, "Mechanical Component ($)"] = cost_mech_components
                current_data.loc[0, "NRE ($)"] = cost_nre
                current_data.loc[0, "Consumables ($)"] = cost_consumables
                current_data.loc[0, "Select Annual Volume"] = annual_volume

                # Update the values in `edited_data` (current_data) for these headers
                for label, value in percentage_values.items():
                    try:
                        # Convert value to float and update the data
                        current_data.loc[0, label] = float(value)
                    except ValueError:
                        st.warning(f"Invalid value for {label}. Please check the input.")

                current_data.loc[0, "MOH ($)"] = moh_cost_value
                current_data.loc[0, "FOH ($)"] = foh_cost_value
                current_data.loc[0, "Profit on RM ($)"] = profit_on_rm_cost_value
                current_data.loc[0, "Profit on VA ($)"] = profit_on_va_cost_value
                current_data.loc[0, "Material Cost ($)"] = total_material_cost_value
                current_data.loc[0, "Manufacturing Cost ($)"] = total_manufacturing_cost_value
                current_data.loc[0, "OH&P ($)"] = total_ohp_cost_value
                current_data.loc[0, "R&D ($)"] = r_n_d_cost_value
                current_data.loc[0, "Warranty ($)"] = warranty_cost_value
                current_data.loc[0, "SG&A ($)"] = sg_and_a_cost_value
                current_data.loc[0, "Total Cost ($)"] = grand_total_cost_value
                current_data.loc[0, "RM Cost ($)"] = rm_cost_value
                current_data.loc[0, "Conversion Cost ($)"] = conversion_cost_value

            if sheet_name:  # Check if the sheet_name is not empty
                # Update the session state
                st.session_state.edited_sheets[sheet_name] = current_data
                st.success("Consumable, RM & Conversion Costing Details saved successfully.")
            else:
                st.error("No data available to save Consumable, RM & Conversion Costing Details.")
    
    # Update the data editor with the latest data
    edited_data2 = st.session_state.edited_sheets.get(sheet_name, pd.DataFrame())

    # Generate a unique key for the data editor widget
    unique_key = f"data_editor_{sheet_name}_{uuid.uuid4()}"

    # Display the data editor
    edited_data2 = st.data_editor(edited_data2, key=unique_key)

    # Option to download the data using BytesIO
    if not edited_data2.empty:
        # Serialize the edited_data2 DataFrame into an Excel file
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            edited_data2.to_excel(writer, sheet_name=sheet_name, index=False)
        output.seek(0)  # Reset the buffer position to the beginning

        # Provide download button for the serialized data
        st.download_button(
            label="Download Excel file",
            data=output,
            file_name=f"{sheet_name}_Costing_Details.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No data available for download.")


    
    st.write("-------------------")


    # Data for the pie chart
    rm_cost_value_from_edited_data2 = edited_data2.at[0, "RM Cost ($)"] # To take the contents from edited_data2 
    conversion_cost_value_from_edited_data2 = edited_data2.at[0, "Conversion Cost ($)"]
    labels = ["RM Cost ($)", "Conversion Cost ($)"]
    values = [rm_cost_value_from_edited_data2, conversion_cost_value_from_edited_data2]

    # Create the pie chart
    fig = go.Figure(data=[go.Pie(labels=labels, values=values, hole=0.4)])

    # Add title and adjust layout
    fig.update_layout(
        title_text='RM vs Conversion Cost %',
        height=600,  # Adjust height as needed
        width=800,   # Adjust width as needed
        margin=dict(t=50, b=50, l=50, r=50)  # Adjust margins as needed
    )

    # Streamlit layout to display the pie chart
    graph_col1, graph_col2 = st.columns(2)

    with graph_col2:
        st.plotly_chart(fig, use_container_width=True)

    # Horizontal Bar Chart 
    total_material_cost_value_edited_data2 = edited_data2.at[0, "Material Cost ($)"] # To take the contents from edited_data2 
    total_manufacturing_cost_value_edited_data2 = edited_data2.at[0, "Manufacturing Cost ($)"]
    total_ohp_cost_value_edited_data2 = edited_data2.at[0, "OH&P ($)"]
    r_n_d_cost_value_edited_data2 = edited_data2.at[0, "R&D ($)"]
    warranty_cost_value_edited_data2 = edited_data2.at[0, "Warranty ($)"]
    sg_and_a_cost_value_edited_data2 = edited_data2.at[0, "SG&A ($)"]
    grand_total_cost_value_edited_data2 = edited_data2.at[0, "Total Cost ($)"]

    data = {
        "Cost Component": [
            "Material Cost ($)",
            "Manufacturing Cost ($)",
            "OH&P ($)",
            "R&D ($)",
            "Warranty ($)",
            "SG&A ($)",
            "Total Cost ($)"
        ],
        "Amount": [
            total_material_cost_value_edited_data2,
            total_manufacturing_cost_value_edited_data2,
            total_ohp_cost_value_edited_data2,
            r_n_d_cost_value_edited_data2,
            warranty_cost_value_edited_data2,
            sg_and_a_cost_value_edited_data2,
            grand_total_cost_value_edited_data2
        ]
    }

    # Convert to DataFrame
    df = pd.DataFrame(data)


    # Horizontal Bar Chart using Plotly Express
    fig = px.bar(
        df,
        y="Cost Component",
        x="Amount",
        orientation="h",
        title="Cost Breakdown",
        color="Cost Component",  # Optional: adds color for each component
        text="Amount"  # Show values on bars
    )

    fig.update_traces(texttemplate='%{text:.2f}', textposition='outside')
    fig.update_layout(xaxis_title="Cost Amount ($)", 
                        height=600,  # Adjust height as needed
                        width=800,   # Adjust width as needed
                        margin=dict(t=50, b=50, l=50, r=50),  # Adjust margins as needed
                        yaxis_title="Cost Component")

    # Display Chart in Streamlit
    with graph_col1:
        st.plotly_chart(fig)

#revision history 26-12-2024 - the merging issues resolved
