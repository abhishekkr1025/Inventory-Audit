import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
# Set the title of the Streamlit app
st.title('Top 70% Consuming Items Finder')

# Description
st.write("""
This app allows you to upload a sales data Excel file and identifies the top 70% highly consuming items based on quantity.
Please ensure your file has columns named 'Item' and 'Quantity'.
""")

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    st.write("File uploaded successfully.")
    try:
        # Read the Excel file
        df = pd.read_excel(uploaded_file)
        st.write("File read successfully.")
        
        # Display the uploaded dataframe
        st.write("Uploaded Sales Data:")
        st.dataframe(df)
        
        # Check if the necessary columns exist
        if 'Item' in df.columns and 'Quantity' in df.columns:
            st.write("Required columns found in the data.")
            
            # Aggregate the data by item and calculate the sum of quantities
            df_agg = df.groupby('Item', as_index=False)['Quantity'].sum()
            
            # Sort the dataframe by quantity in descending order
            df_sorted = df_agg.sort_values(by='Quantity', ascending=False)
            
            # Calculate the cumulative sum of quantities
            df_sorted['CumulativeQuantity'] = df_sorted['Quantity'].cumsum()
            
            # Calculate the total quantity
            total_quantity = df_sorted['Quantity'].sum()
            
            # Determine the threshold for the top 70%
            threshold = total_quantity * 0.7
            
            # Find the top 70% consuming items
            df_sorted['CumulativePercentage'] = df_sorted['CumulativeQuantity'] / total_quantity
            top_70_percent_items = df_sorted[df_sorted['CumulativePercentage'] <= 0.7]
            
            # Display the top 70% consuming items
            st.write("Top 70% Consuming Items:")
            top_70_percent_items['Available Quantity'] = None
            for i, row in top_70_percent_items.iterrows():
                available_qty = st.text_input(f"Available Quantity for {row['Item']}", key=row['Item'])
                top_70_percent_items.at[i, 'Available Quantity'] = available_qty
            # Display the dataframe with the available quantities
            st.write("Top 10 Most Selling Items with Available Quantities:")
            st.dataframe(top_70_percent_items[['Item', 'Quantity', 'Available Quantity']])
            
            # Display a bar chart for the top 70% consuming items
            st.bar_chart(top_70_percent_items.set_index('Item')['Quantity'])

            # Function to convert the dataframe to an Excel file in memory
            def to_excel(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                processed_data = output.getvalue()
                return processed_data

            # Filter the dataframe to include only the 'Item' and 'Available Quantity' columns
            save_df = top_70_percent_items[['Item', 'Available Quantity']]

            # Convert the filtered dataframe to Excel format
            df_xlsx = to_excel(save_df)
            
            # Add a download button to save the Excel file
            st.download_button(
                label="Save to Excel",
                data=df_xlsx,
                file_name='top_70_percent_items_with_available_qty.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            st.error("The uploaded file does not contain the required columns 'Item' and 'Quantity'.")
    except Exception as e:
        st.error(f"An error occurred: {e}")
else:
    st.info("Please upload an Excel file to proceed.")
