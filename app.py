import streamlit as st
import pandas as pd
from io import BytesIO

# Set the title of the Streamlit app
st.title('Items Finder')

# Description
st.write("""
This app allows you to upload a sales data Excel file and identifies the most selling items based on total quantity sold.
You can also manually input the available quantity for each item and save the result as a new Excel file.
Please ensure your file has columns named 'Item', 'Quantity', 'Category', and 'Purchase Qty in last 6 months'.
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
        if 'Item' in df.columns and 'Quantity' in df.columns and 'Category' in df.columns and 'Purchase Qty in last 6 months' in df.columns:
            st.write("Required columns found in the data.")
            
            # Get unique categories
            categories = df['Category'].unique()
            
            # Select category
            selected_category = st.selectbox("Select a Category", categories)
            
            # Filter data based on the selected category
            filtered_df = df[df['Category'] == selected_category]
            
            # Exclude items with zero quantity in 'Purchase Qty in last 6 months'
            filtered_df = filtered_df[filtered_df['Purchase Qty in last 6 months'] > 0]
            
            # Aggregate the data by item and calculate the sum of quantities
            df_agg = filtered_df.groupby('Item', as_index=False)['Quantity'].sum()
            
            # Sort the dataframe by quantity in descending order
            df_sorted = df_agg.sort_values(by='Quantity', ascending=False)
            
            # Add a column for available quantities
            df_sorted['Available Quantity'] = None
            
            for i, row in df_sorted.iterrows():
                available_qty = st.text_input(f"Available Quantity for {row['Item']}", key=row['Item'])
                df_sorted.at[i, 'Available Quantity'] = available_qty
            
            # Display the dataframe with the available quantities
            st.write("Most Selling Items with Available Quantities:")
            st.dataframe(df_sorted[['Item', 'Quantity', 'Available Quantity']])
            
            # Display a bar chart for the most selling items
            st.bar_chart(df_sorted.set_index('Item')['Quantity'])

            # Function to convert the dataframe to an Excel file in memory
            def to_excel(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                processed_data = output.getvalue()
                return processed_data

            # Filter the dataframe to include only the 'Item' and 'Available Quantity' columns
            save_df = df_sorted[['Item', 'Available Quantity']]

            # Convert the filtered dataframe to Excel format
            df_xlsx = to_excel(save_df)
            
            # Add a download button to save the Excel file
            st.download_button(
                label="Save to Excel",
                data=df_xlsx,
                file_name='items_with_available_qty.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            st.error("The uploaded file does not contain the required columns 'Item', 'Quantity', 'Category', and 'Purchase Qty in last 6 months'.")
    except Exception as e:
        st.error(f"An error occurred: {e}")
else:
    st.info("Please upload an Excel file to proceed.")