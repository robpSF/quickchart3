import streamlit as st
import pandas as pd
import io
import matplotlib.pyplot as plt

# Function to process data and create the table and exceptions
def process_data(file):
    data = pd.read_excel(file, sheet_name='opportunities')

    # Extract relevant columns
    relevant_data = data[['Contact Name', 'Opportunity Name', 'Milestone', 'GBP Value', 'Close Date','Owner','Updated','Owner']]
    
    # Convert 'Expected_Renewal' to datetime
    relevant_data['Close Date'] = pd.to_datetime(relevant_data['Close Date'], errors='coerce')
    
    # Filter out rows where Expected_Renewal is NaT
    relevant_data = relevant_data.dropna(subset=['Close Date'])
    
    # Create a new column with year-month format
    relevant_data['YearMonth'] = relevant_data['Close Date'].dt.strftime('%Y-%m')
    
    # Create a pivot table with year-month as columns and names as rows, including 'Licence'
    pivot_table = relevant_data.pivot_table(index=['Contact Name', 'Opportunity Name'], columns='YearMonth', values='GBP Value', aggfunc='sum', fill_value=0).reset_index()
    
    # Additional data columns to be merged
    additional_columns = data[['Contact Name', 'Milestone', 'Owner', 'Updated','Owner']].drop_duplicates()
    
    # Merge the additional columns into the pivot table
    merged_data = pd.merge(pivot_table, additional_columns, on='Contact Name', how='left')
    
    # Reorder columns to move 'LicenceChange' and 'RenewalStatus' next to 'Name'
    cols = ['Contact Name'] + [col for col in merged_data.columns if col not in ['Contact Name', 'Opportunity Name', 'Milestone','Status','Updated']]
    merged_data = merged_data[cols]
    
    return merged_data

# Streamlit app
st.title('Forecast')

# File uploader
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file is not None:
    merged_data = process_data(uploaded_file)
    
    st.header("Forecast")
    st.dataframe(merged_data)
    
   
    # Create a bar chart for the new data
    st.header("Bar Chart: Forecast by Year-Month")
    merged_data_melted = merged_data.melt(id_vars=['Name', 'Opportunity Name', 'Status'], var_name='YearMonth', value_name='GBP Value')
    merged_data_melted = merged_data_melted[merged_data_melted['YearMonth'].str.match(r'\d{4}-\d{2}')]

    pivot_table_sum = merged_data_melted.groupby('YearMonth')['GBP Value'].sum().reset_index()
    pivot_table_sum.columns = ['YearMonth', 'GBP Value']
    
    fig, ax = plt.subplots()
    ax.bar(pivot_table_sum['YearMonth'], pivot_table_sum['GBP Value'])
    ax.set_xlabel('Year-Month')
    ax.set_ylabel('Forcast Revenue')
    ax.set_title('Forecast Revenue by Year-Month')
    st.pyplot(fig)
    
    # Function to convert dataframe to Excel and provide a download link
    def to_excel(df):
        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, index=True, sheet_name='Sheet1')
        writer.close()
        processed_data = output.getvalue()
        return processed_data
    
    # Provide download links
    merged_data_excel = to_excel(merged_data)
    
    st.download_button(label="Download Relevant Data as Excel", data=merged_data_excel, file_name='forecast_data.xlsx')
