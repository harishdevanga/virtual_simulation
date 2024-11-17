import streamlit as st

st.set_page_config(
    page_title = "Multipage app",
    page_icon= ":wave:",
    layout="wide"
)

st.title("Main Page")
st.sidebar.success("Select a page above.")

import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
import openpyxl
import plotly.graph_objects as go


# App Title
# st.set_page_config(page_title="FPY Prediction!!!", page_icon=":bar_chart:", layout="wide")
st.subheader(":point_right: FPY Prediction Summary")


# Step 1: Upload the Excel file 
# Dropdown to open and minimize file uploader
with st.expander("Upload Your <Process Mapping> File Here", expanded=True):  # Collapsible dropdown for file uploader
    uploaded_file = st.file_uploader("Choose Your <Process Mapping> File", type=["xlsx", "csv", "xlsm"])

if uploaded_file:
    # Load data from the uploaded file
    @st.cache_data
    def load_data(file):
        if file.name.endswith('.csv'):
            df = pd.read_csv(file)
            return {"csv_data": df}  # Return a dictionary with the data
        else:
            df = pd.read_excel(file, sheet_name=None)  # Load all sheets as a dictionary
            return df

    xls_summary = load_data(uploaded_file)

    # Filter sheet names that start with 'MK', 'X', or 'SOP'
    sheets_to_read = [sheet for sheet in xls_summary.keys() if sheet.startswith(('MK', 'X', 'SOP'))]

    # Initialize an empty list to store dataframes
    all_data = []

    # Step 2: Loop through the sheets and process data
    for sheet_name in sheets_to_read:
        # Read the data from the sheet
        df_table = xls_summary[sheet_name]
        
        # Extract relevant values
        annual_vol_nos = df_table.at[0, 'Annual Volume']
        batch_qty = annual_vol_nos / 12
        end_to_end_cycle_time = df_table.at[0, 'Total Cycle Time, sec'] / 60  # Convert seconds to minutes
        available_time_per_day = 24 * 60 - (3 * 45)  # Convert hours to minutes and subtract the time
        desired_takt_time = available_time_per_day / (batch_qty / 24)
        desired_trough_put = ((batch_qty / 24) / available_time_per_day) * 60
        cycle_time = df_table.at[0, 'Max Overall PCBA CT'] / 60
        actual_trough_put = available_time_per_day / cycle_time

        # Create a dictionary for the current sheet's data
        processct_data = {
            'Metric': ['End to End Cycle Time', 'Desired Takt Time', 'Desired Trough Put(UPH)', 'Cycle Time', 'Actual Trough Put', 'Annual Volume', 'Batch Qty'],
            'UOM': ['Minute', 'Minute', 'Units/hour', 'Minute', 'Units/shift', 'nos', 'nos'],
            sheet_name: [end_to_end_cycle_time, desired_takt_time, desired_trough_put, cycle_time, actual_trough_put, annual_vol_nos, batch_qty]  # Values specific to the sheet
        }

        # Convert the dictionary to a pandas DataFrame
        df_processct = pd.DataFrame(processct_data)
        all_data.append(df_processct)  # Add the current dataframe to the list

    # Step 3: Combine all dataframes into one based on 'Metric' and 'UOM'
    final_df = all_data[0]  # Start with the first dataframe
    for df in all_data[1:]:
        final_df = pd.merge(final_df, df, on=['Metric', 'UOM'], how='outer')  # Merge based on 'Metric' and 'UOM'

    # Display the final combined DataFrame
    st.write("### PCB Assembly Process Cycle Time Summary")
    st.write(final_df)

# Step 3: DFM Details - Allow user input for percentages

st.subheader(":point_right: DFx Summary")

# Function to color the FPY values based on thresholds
def color_fpy(val):
    try:
        # Try to convert the value to a float for comparison
        val = float(val)
    except (ValueError, TypeError):
        return ""  # Return empty string if conversion fails (for non-numeric or NaN)

    # Apply color based on FPY thresholds
    color = "red" if val < 0.85 else "#00ff00" if val >= 0.98 else "yellow"
    return f'background-color: {color}; color: black'

with st.expander("Upload your <DFx Summary> File Here", expanded=True):  # Collapsible dropdown for file uploader
    # Step 1: Upload the Excel file and load the sheet named "dfx_analysis"
    uploaded_file = st.file_uploader("Choose your <DFx Summary> File", type=["xlsx"])

if uploaded_file:
    # Load the specific sheet
    df_dfm = pd.read_excel(uploaded_file, sheet_name="dfx_analysis")

    # Step 1: Ensure that numeric columns (MK0 to SOP) are converted to numeric
    columns_to_convert = ['MK0', 'MK1', 'MK2', 'MK3', 'X1', 'X1.1', 'X1.2', 'SOP']
    df_dfm[columns_to_convert] = df_dfm[columns_to_convert].apply(pd.to_numeric, errors='coerce')

    # Step 2: Apply the color formatting function to the relevant columns (MK0 to SOP)
    st.write(df_dfm.style.applymap(color_fpy, subset=columns_to_convert))


# Step 4: Process Yield Prediction - Allow user input for percentages

st.subheader(":point_right: Process Yield Prediction")

# Dropdown to open and minimize file uploader
with st.expander("Upload Your <PFMEA> File Here", expanded=True):  # Collapsible dropdown for file uploader
    uploaded_file = st.file_uploader("Choose Your <PFMEA> File", type=["xlsx", "csv", "xlsm"])

if uploaded_file:
    # Load data from the uploaded file
    @st.cache_data
    def load_data(file):
        if file.name.endswith('.csv'):
            df = pd.read_csv(file)
            return {"csv_data": df}  # Return a dictionary with the data
        else:
            df = pd.read_excel(file, sheet_name=None)  # Load all sheets as a dictionary
            return df

    xls = load_data(uploaded_file)

    # Filter sheet names that start with 'MK', 'X', or 'SOP'
    sheets_to_read = [sheet for sheet in xls.keys() if sheet.startswith(('MK', 'X', 'SOP'))]

    # Create the number of columns based on the number of sheets (up to 5 in this example)
    columns = st.columns(min(len(sheets_to_read), 5))  # Limits columns to 5 if more sheets are present

    # Step 2: Loop through the sheets and generate a graph for each
    for i, sheet_name in enumerate(sheets_to_read):
        # Read the data from the sheet
        df_yp = xls[sheet_name]
        
        # Calculate the OMI product
        if 'Process step/Input' in df_yp.columns and 'OMI' in df_yp.columns:
            group_omi = df_yp.groupby('Process step/Input')[['OMI']].mean()
            omi_product = group_omi['OMI'].prod()
            
            # Step 3: Generate the gauge graph
            fig = go.Figure(go.Indicator(
                mode="gauge+number",
                value=omi_product * 100,
                gauge={'axis': {'range': [0, 100]}, 'bar': {'color': "darkblue"}},
                title={'text': f"Process FPY_ {sheet_name}"},
                number={'suffix': "%", 'valueformat': ".2f"}  # Format value as percentage with two decimals
            ))
            
            # Step 4: Display each chart in its respective column
            with columns[i % 5]:  # % 5 to wrap around if there are more than 5 sheets
                st.plotly_chart(fig)

# Step 5: End Item DPMO for PCBA (IPC7912) - Allow user input for percentages

st.subheader(":point_right: End Item DPMO for PCBA (IPC7912)")
# Dropdown to open and minimize file uploader
with st.expander("Upload <End Item DPMO for PCBA (IPC7912)> File Here", expanded=True):  # Collapsible dropdown for file uploader
    # Step 1: Upload the Excel file and load the sheet named "ipc_analysis"
    uploaded_file = st.file_uploader("Choose your <End Item DPMO for PCBA (IPC7912)> File", type=["xlsx"])

if uploaded_file:
    # Load the specific sheet
    df = pd.read_excel(uploaded_file, sheet_name="ipc_analysis")

    # Step 2: Filter the columns that start with 'MK', 'X', or 'SOP'
    relevant_columns = [col for col in df.columns if col.startswith(('MK', 'X', 'SOP'))]

    # Step 3: Filter rows where 'Package' column contains 'Estimated Yield from Current p Value' and 'Estimated DPMO'
    filtered_df = df[df['Package'].isin(["Estimated Yield from Current p Value", "Estimated DPMO"])]

    # Step 4: Prepare the final table with filtered rows and relevant columns
    final_table = filtered_df[['Package'] + relevant_columns].set_index('Package')

    # Step 5: Format the 'Estimated Yield from Current p Value' row values as percentages
    if "Estimated Yield from Current p Value" in final_table.index:
        final_table.loc["Estimated Yield from Current p Value"] = final_table.loc["Estimated Yield from Current p Value"].astype(float) * 100

    # Step 6: Display the final table with percentages
    st.write("Filtered Data:")
    st.dataframe(final_table.style.format({"Estimated Yield from Current p Value": "{:.2f}%"}))

    # Step 7: Create a combination graph
    yield_values = final_table.loc["Estimated Yield from Current p Value"]
    dpmo_values = final_table.loc["Estimated DPMO"]

    # Create the figure
    fig = go.Figure()

    # Add bar chart for DPMO
    fig.add_trace(go.Bar(
        x=relevant_columns,
        y=dpmo_values,
        name='Estimated DPMO',
        marker_color='indianred',
        yaxis='y'
    ))

    # Add line chart for Estimated Yield in percentages
    fig.add_trace(go.Scatter(
        x=relevant_columns,
        y=yield_values,
        name='Estimated Yield (%)',
        mode='lines+markers',
        marker_color='blue',
        yaxis='y2'
    ))

    # Update the layout to have dual y-axes
    fig.update_layout(
        title="Estimated Yield (%) and DPMO Comparison",
        xaxis_title="Columns (MK, X, SOP)",
        yaxis_title="DPMO",
        yaxis2=dict(
            title="Estimated Yield (%)",
            overlaying="y",
            side="right"
        ),
        legend=dict(x=0.1, y=1.1, orientation="h"),
        height=600
    )

    # Step 8: Display the graph below the table
    st.write("Graph Representation:")
    st.plotly_chart(fig)

st.subheader(":point_right: Yield in PCBA (IEEE ref.)")

# Dropdown to open and minimize file uploader
with st.expander("Upload Your <Yield in PCBA (IEEE ref.)> File Here", expanded=True):
    # Step 1: Upload the Excel file
    uploaded_file = st.file_uploader("Choose Your <Yield in PCBA (IEEE ref.)> File", type=["xlsx"])
    if uploaded_file is not None:
        display_data = pd.read_excel(uploaded_file, sheet_name=0)
        display_data    

if uploaded_file is not None:
    # Read the uploaded Excel file into a DataFrame
    try:
        edited_data = pd.read_excel(uploaded_file, sheet_name=0)  # Assuming the first sheet contains data

        st.write("-----------------------------------")
        st.subheader("Yield vs. Solder Defects Analysis")
        col_graph1, col_graph1_1 = st.columns(2)
        with col_graph1:
            # Extracting the columns (board names)
            board_names = [col for col in edited_data.columns if col != "Data Points"]

            # Extract solder joint counts dynamically based on the edited_data
            solder_joint_row = edited_data.loc[edited_data["Data Points"] == "No. Solder Joints (N)"].squeeze()
            solder_joints = {board: solder_joint_row[board] for board in board_names}

            # Defect rate scaling values (logarithmic range for better visualization)
            defect_rate_scaling = np.logspace(-1, 2, 50)  # From 0.1 to 100 in logarithmic scale

            # Define function to calculate yield based on defect rate scaling
            def calculate_yield(solder_joints, defect_rate_scaling, defect_rate_per_joint):
                overall_defect_rate = defect_rate_scaling * defect_rate_per_joint
                yield_percentage = 100 * np.exp(-overall_defect_rate * solder_joints)
                return yield_percentage

            # Get defect rate per joint from the edited_data DataFrame
            defect_rate_per_joint_row = edited_data.loc[edited_data["Data Points"] == "Defect Rate per Solder Joint (DR)"].squeeze()
            defect_rates = {board: defect_rate_per_joint_row[board] for board in board_names}

            # Calculate yields for each board
            yield_data = []
            for board in board_names:
                yield_percentage = calculate_yield(solder_joints[board], defect_rate_scaling, defect_rates[board])
                yield_data.append(pd.DataFrame({
                    "Defect Rate Scaling": defect_rate_scaling,
                    "Yield (%)": yield_percentage,
                    "Board": board
                }))

            # Concatenate all yield data into a single DataFrame
            df = pd.concat(yield_data, ignore_index=True)

            # Plotting the yield vs defect rate scaling
            fig = px.line(
                df, x="Defect Rate Scaling", y="Yield (%)", color="Board",
                log_x=True,  # Log scale for defect rate scaling
                title="Yield vs. Solder Defect Rate (Full Scale)"
            )
            fig.update_layout(xaxis_title="Solder Defect Rate Scaling", yaxis_title="Yield (%)")
            st.plotly_chart(fig)

            # Filter the DataFrame to focus on scaling factors <= 10
            detail_df = df[df["Defect Rate Scaling"] <= 10]

            # Create the line chart using Plotly Express
            fig2 = px.line(
                detail_df,
                x="Defect Rate Scaling",
                y="Yield (%)",
                color="Board",  # Differentiates lines by board
                title="Yield vs. Solder Defect Rate Scaling (Detail View)",
                labels={
                    "Defect Rate Scaling": "Solder Defect Rate Scaling (<= 10)",
                    "Yield (%) Change": "Yield (%) Change (%)"
                },
                template="plotly_white"
            )
        with col_graph1_1:
            # Display the chart in the second column
            st.plotly_chart(fig2, use_container_width=True)

        st.write("-----------------------------------")

        st.subheader("Cost vs. Solder Defects Analysis")         
        col_graph3, col_graph3_3 = st.columns(2)
        with col_graph3:
            # Extracting the columns (board names)
            board_names = [col for col in edited_data.columns if col != "Data Points"]

            # Extract solder joint counts dynamically based on the edited_data
            solder_joint_row = edited_data.loc[edited_data["Data Points"] == "No. Solder Joints (N)"].squeeze()
            defect_rate_row = edited_data.loc[edited_data["Data Points"] == "Defect Rate per Solder Joint (DR)"].squeeze()

            scaling_factors = np.linspace(0.1, 100, 50)  # Scaling factors from 0.1 to 100

            # Initialize a DataFrame to store results for all boards
            plot_df = pd.DataFrame()

            for board in board_names:
                # Extract defect rate and solder joint values for the current board
                solder_joints = solder_joint_row[board]
                defect_rate = defect_rate_row[board]

                # Calculate costs and cost percentages
                costs = [defect_rate * solder_joints * scale for scale in scaling_factors]
                cost_percentages = [scale * 100 for scale in scaling_factors]

                # Create a temporary DataFrame for the current board
                temp_df = pd.DataFrame({
                    "Solder Defect Rate Scaling": scaling_factors,
                    "Cost % Change": costs,
                    "Board": board
                })

                # Append to the main DataFrame
                plot_df = pd.concat([plot_df, temp_df], ignore_index=True)
            
            # Graph 1: Cost vs Solder Def Rate (Full Scale)
            fig1 = px.line(
                plot_df,
                x="Solder Defect Rate Scaling",
                y="Cost % Change",
                color="Board",
                title="Cost vs Solder Defect Rate (Full Scale)",
                labels={"Solder Defect Rate Scaling": "Solder Defect Rate Scaling", "Cost % Change": "Cost % Change"},
                template="plotly_white",
                log_x=True
            )

            st.plotly_chart(fig1, use_container_width=True)


        with col_graph3_3:
            # Filter the DataFrame to focus on scaling factors <= 10
            detail_df = plot_df[plot_df["Solder Defect Rate Scaling"] <= 10]

            # Create the line chart using Plotly Express
            fig2 = px.line(
                detail_df,
                x="Solder Defect Rate Scaling",
                y="Cost % Change",
                color="Board",  # Differentiates lines by board
                title="Cost vs Solder Defect Rate (Detail View)",
                labels={
                    "Solder Defect Rate Scaling": "Solder Defect Rate Scaling (<= 10)",
                    "Cost % Change": "Cost % Change"
                },
                template="plotly_white"
            )

            # Display the chart in the second column
            st.plotly_chart(fig2, use_container_width=True)

        st.write("-----------------------------------")

    except Exception as e:
        st.error(f"Error reading the file: {e}")
else:
    st.warning("Please upload a valid Excel file.")


# Logic Sheet to select similar to super market select box, then apply pivot
# , then show the MKx data in a table summary similar to DFx summary
# recent code ==> Homepage1_1.py
