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
with st.expander("Upload Your <Yield in PCBA (IEEE ref.)> File Here", expanded=True):  # Collapsible dropdown for file uploader
    # Step 1: Upload the Excel file and load the sheet named "ieee_analysis"
    uploaded_file = st.file_uploader("Choose Your <Yield in PCBA (IEEE ref.)> File", type=["xlsx"])

if uploaded_file:
    # Load the specific sheet
    df = pd.read_excel(uploaded_file, sheet_name="ieee_analysis")

    # Step 2: Filter the columns that start with 'MK', 'X', or 'SOP'
    relevant_columns = [col for col in df.columns if col.startswith(('MK', 'X', 'SOP'))]

    # Step 3: Filter rows where 'Data Points' column contains "Overall yield_Soldering", "Overall yield_Component", "Overall yield_Placement"
    filtered_df = df[df['Data Points'].isin(["Overall yield_Soldering", "Overall yield_Component", "Overall yield_Placement"])]

    # Step 4: Prepare the final table with filtered rows and relevant columns
    final_table = filtered_df[['Data Points'] + relevant_columns].set_index('Data Points')

    # Step 5: Format the 'Yield Value' row values as percentages
    final_table.loc["Overall yield_Soldering"] = final_table.loc["Overall yield_Soldering"].astype(float) * 100
    final_table.loc["Overall yield_Component"] = final_table.loc["Overall yield_Component"].astype(float) * 100
    final_table.loc["Overall yield_Placement"] = final_table.loc["Overall yield_Placement"].astype(float) * 100

    # Step 6: Display the final table with percentages
    st.write("Filtered Data:")
    st.dataframe(final_table.style.format({
        "Overall yield_Soldering": "{:.2f}%",
        "Overall yield_Component": "{:.2f}%",
        "Overall yield_Placement": "{:.2f}%"
    }))

    # Step 7: Create a combination graph (e.g., grouped bar chart)
    yield_s_values = final_table.loc["Overall yield_Soldering"]
    yield_c_values = final_table.loc["Overall yield_Component"]
    yield_p_values = final_table.loc["Overall yield_Placement"]

    # Create the figure for a grouped bar chart
    fig = go.Figure()

    # Add bar traces for each yield type
    fig.add_trace(go.Bar(
        x=relevant_columns,  # MK, X, and SOP columns
        y=yield_s_values,
        name='Overall yield_Soldering',
        marker_color='blue'
    ))

    fig.add_trace(go.Bar(
        x=relevant_columns,
        y=yield_c_values,
        name='Overall yield_Component',
        marker_color='green'
    ))

    fig.add_trace(go.Bar(
        x=relevant_columns,
        y=yield_p_values,
        name='Overall yield_Placement',
        marker_color='red'
    ))

    # Update the layout to be a grouped bar chart
    fig.update_layout(
        title="Comparison of Yields (Soldering, Component, Placement)",
        xaxis_title="Columns (MK, X, SOP)",
        yaxis_title="Yield (%)",
        barmode='group',  # This ensures the bars are grouped side by side
        legend=dict(x=0.1, y=1.1, orientation="h"),
        height=600
    )

    # Step 8: Display the graph below the table
    st.write("Graph Representation:")
    st.plotly_chart(fig)

# Logic Sheet to select similar to super market select box, then apply pivot
# , then show the MKx data in a table summary similar to DFx summary
# recent code ==> Homepage1.py


