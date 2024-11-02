import pandas as pd
import streamlit as st

# Set the title of the web app
st.title("Meeting Attendance Analysis")

# File upload functionality with a unique key
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"], key="unique_file_uploader")

if uploaded_file is not None:
    # Load the Excel file and list available sheet names
    excel_file = pd.ExcelFile(uploaded_file)
    sheet_names = excel_file.sheet_names

    # Display a select box to choose the sheet
    sheet_name = st.selectbox("Select the sheet containing attendance data", sheet_names)

    # Load the selected sheet
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, skiprows=3)

    # Rename columns for ease of use
    df.columns = ['Name', 'First Join', 'Last Leave', 'In-Meeting Duration', 'Email', 'Participant ID', 'Role']

    # Convert 'First Join' and 'Last Leave' columns to datetime format
    df['First Join'] = pd.to_datetime(df['First Join'], errors='coerce')
    df['Last Leave'] = pd.to_datetime(df['Last Leave'], errors='coerce')

    # Calculate the duration based on 'First Join' and 'Last Leave'
    df['Calculated Duration'] = df['Last Leave'] - df['First Join']

    # Group by participant name to sum total duration
    duracao_por_participante = df.groupby('Name')['Calculated Duration'].sum().reset_index()

    # Format the duration to show hours, minutes, and seconds
    duracao_por_participante['Calculated Duration'] = duracao_por_participante['Calculated Duration'].apply(
        lambda x: str(x).split('.')[0]  # Remove milliseconds for precise time display
    )

    # Display the results in Streamlit
    st.write("Total participation time for each participant (HH:MM:SS):")
    st.dataframe(duracao_por_participante)


