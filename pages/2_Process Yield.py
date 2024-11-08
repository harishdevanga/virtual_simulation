import streamlit as st
import pandas as pd
import plotly.express as px
from sklearn.linear_model import LinearRegression
from sklearn.model_selection import train_test_split
from sklearn.impute import SimpleImputer
import numpy as np

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
    if isinstance(data, dict):
        sheet_name = st.selectbox("Select the sheet", data.keys())

        # Check if the sheet has been edited before; if so, load the edited version
        if sheet_name in st.session_state.edited_sheets:
            st.session_state.df = st.session_state.edited_sheets[sheet_name]
        else:
            selected_data = data[sheet_name]
            st.session_state.df = pd.DataFrame(selected_data)  # Load original data from file

    # Display data in a table
    st.subheader("Data Table")
    edited_data = st.data_editor(st.session_state.df)

    # Create buttons side by side for adding new rows and saving the table
    col1, col2 = st.columns([1, 1])

    # Create a separate save button for general edits in the table
    with col1:
        if st.button("Save Edited Table"):
            # Save the edited data in the session state
            st.session_state.edited_sheets[sheet_name] = edited_data
            
            # Save all changes made in the table to the file
            with pd.ExcelWriter(uploaded_file.name, engine="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
                edited_data.to_excel(writer, sheet_name=sheet_name, index=False)
            st.success("Table saved successfully!")
            st.rerun()
    
    with col2:
        if st.button("Add New Row"):
            # Create a new row with NaN values
            new_row = pd.DataFrame({col: [np.nan] for col in edited_data.columns})
            # Append the new row to the DataFrame
            st.session_state.df = pd.concat([st.session_state.df, new_row], ignore_index=True)
            st.session_state.edited_sheets[sheet_name] = st.session_state.df  # Update session state with new row
            st.rerun()  # Rerun the app to update the data editor with the new row

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
                # Load the entire Excel file into a dictionary of DataFrames
                excel_file = pd.read_excel(uploaded_file.name, sheet_name=None)
                # Access the specific sheet and update it
                excel_file[sheet_name] = st.session_state.df
                # Save all sheets back to the Excel file
                with pd.ExcelWriter(uploaded_file.name, engine="openpyxl", mode="w") as writer:
                    for sheet_name, sheet_data in excel_file.items():
                        sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)
                st.success("Removed rows saved successfully!")
                st.rerun()

    # OCC Ranking Table
    OCC_Ranking_Table = {
        'OCC_Ranking': [10, 9, 8, 7, 6, 5, 4, 3, 2, 1],
        'OCC_Incidents_Per_Item': [0.1, 0.05, 0.02, 0.01, 0.002, 0.0005, 0.0001, 0.00001, 0.000001, 0]
    }
    occ_ranking_df = pd.DataFrame(OCC_Ranking_Table)

    # Compute OMI based on OCC_Ranking
    edited_data = edited_data.merge(occ_ranking_df, left_on='OCC', right_on='OCC_Ranking', how='left')
    edited_data['OMI'] = 1 - edited_data['OCC_Incidents_Per_Item']

    # Scatter plot
    st.subheader("Scatter Plot of OCC vs OMI")
    fig = px.scatter(edited_data, x='Process step/Input', y='OMI', text='OCC',
                     title='Scatter Plot of OCC vs OMI')
    st.plotly_chart(fig)

    # Subheading for Prediction Model
    st.subheader("Prediction Model")

    # Input fields for OCC and Value Lookup
    st.subheader("Input Data for Prediction")
    occ_input = st.number_input("Enter OCC value")

    # Prepare data for the model
    # Training on a single feature (OCC)
    X = edited_data[['OCC']]
    y = edited_data['OMI']

    # Impute missing values with the mean
    imputer = SimpleImputer(strategy='mean')
    X_imputed = imputer.fit_transform(X)

    # Split data into train and test sets
    X_train, X_test, y_train, y_test = train_test_split(X_imputed, y, test_size=0.2, random_state=42)

    # Train the model
    model = LinearRegression()
    model.fit(X_train, y_train)

    cl1, cl2 = st.columns(2)

    with cl1:
        # Predict and display the results
        st.subheader("Actual vs. Prediction Table")
        predictions = model.predict(X_test)
        results = pd.DataFrame({'Actual': y_test, 'Predicted': predictions})
        st.write(results)

    with cl2:
        # Display DataFrame as a table
        st.subheader("OCC Ranking Table")
        st.table(occ_ranking_df)

    # Predict OMI using the model for input data
    # Predict using a single feature (OCC)
    predicted_omi = model.predict(np.array([[occ_input]]))
    st.write(f"Predicted OMI for input data: {predicted_omi[0]}")

    # Calculate CP and CPK
    st.subheader("Process Capability Statistics")

    def calculate_cp(data):
        USL = data['OMI'].max()
        LSL = data['OMI'].min()
        sigma = data['OMI'].std()
        cp = (USL - LSL) / (6 * sigma)
        return cp

    def calculate_cpk(data):
        USL = data['OMI'].max()
        LSL = data['OMI'].min()
        mean = data['OMI'].mean()
        sigma = data['OMI'].std()
        cpu = (USL - mean) / (3 * sigma)
        cpl = (mean - LSL) / (3 * sigma)
        cpk = min(cpu, cpl)
        return cpk

    cp = calculate_cp(edited_data)
    cpk = calculate_cpk(edited_data)
    st.write(f"CP (Process Capability): {cp}")
    st.write(f"CPK (Process Capability Index): {cpk}")

    st.success("Predictive model built and predictions displayed successfully!")

