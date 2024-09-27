# -*- coding: utf-8 -*-
"""
Created on Fri Sep 27 13:18:07 2024

@author: User
"""

import streamlit as st
import pandas as pd
from io import BytesIO

# Step 1: Read the .CSV file for fuel transaction
st.sidebar.title("Fuel Efficiency Calculator")
uploaded_file = st.sidebar.file_uploader("Upload your main Excel file (fuel transactions)", type="xlsx")

# Step 2: Upload the file that contains VehicleRegistrationNo and associated fuel efficiency factors
uploaded_factors = st.sidebar.file_uploader("Upload your Excel file with fuel efficiency factors", type="xlsx")

if uploaded_file is not None and uploaded_factors is not None:
    # Load the fuel transactions data
    df = pd.read_excel(uploaded_file)
    
    # Load the fuel efficiency factors data
    efficiency_factors_df = pd.read_excel(uploaded_factors)

    # Merge the two dataframes on VehicleRegistrationNo
    df = pd.merge(df, efficiency_factors_df, on="VehicleRegistrationNo", how="left")

    # Remove the timestamp from TransactionDate, keeping only the date
    df['TransactionDate'] = pd.to_datetime(df['TransactionDate']).dt.strftime('%d-%m-%Y')

    # Step 3: Group by VehicleRegistrationNo and sort by TransactionDate
    df.sort_values(by=["VehicleRegistrationNo", "TransactionDate"], inplace=True)

    # Drop the specified columns
    columns_to_drop = [
        'CreationDate', 'CreationTime', 'AccountType', 'AccountNo', 'AccountName',
        'PetrolStationCode', 'CashierCode', 'DriverCode', 'VehicleCode', 'ItemCode',
        'GPSCoordinatelatitude', 'GPSCoordinateLongitude', 'Bypass Types', 
        'SalesInvoiceNo', 'AppMode'
    ]
    df.drop(columns=columns_to_drop, inplace=True, errors='ignore')

    # Create a dropdown list for selecting a specific VehicleRegistrationNo or 'All'
    vehicle_options = ['All'] + df['VehicleRegistrationNo'].unique().tolist()
    selected_vehicle = st.sidebar.selectbox("Select VehicleRegistrationNo", vehicle_options)

    # Filter the dataframe based on the selected VehicleRegistrationNo, or show all vehicles if 'All' is selected
    if selected_vehicle != 'All':
        df = df[df['VehicleRegistrationNo'] == selected_vehicle]

    # Initialize columns for new data
    df['Initial Odometer'] = None
    df['Final Odometer'] = None
    df['Distance'] = None
    df['Rolling Quantity'] = None
    df['Fuel Efficiency'] = None
    df['Fuel Usage'] = None
    df['Usage Type'] = None

    # Step 4 to 12: Process each transaction for the selected VehicleRegistrationNo or all vehicles
    for vehicle, vehicle_df in df.groupby("VehicleRegistrationNo"):
        initial_odometer = None
        initial_quantity = 0
        rolling_quantity = 0
        final_odometer = None

        # Get the specific efficiency factor for this vehicle
        vehicle_efficiency_factor = vehicle_df['FuelEfficiencyFactor'].iloc[0]  # Assuming the column name in the second file is FuelEfficiencyFactor
        
        # Sort transactions for each vehicle by TransactionDate
        vehicle_df = vehicle_df.sort_values(by="TransactionDate")

        for index, row in vehicle_df.iterrows():
            if row['Capacity'] == 'Y':
                if initial_odometer is not None and row['Odometer'] > initial_odometer:
                    # Calculate Distance and Fuel Efficiency
                    final_odometer = row['Odometer']
                    distance = final_odometer - initial_odometer

                    if distance > 0:  # Ensure valid distance
                        fuel_efficiency = distance / rolling_quantity

                        # Calculate Fuel Usage using the specific vehicle efficiency factor
                        fuel_usage = (distance / vehicle_efficiency_factor) - rolling_quantity
                        usage_type = "Saving" if fuel_usage >= 0 else "Excessive Use"

                        # Update the DataFrame with calculated values
                        df.at[index, 'Initial Odometer'] = initial_odometer
                        df.at[index, 'Final Odometer'] = final_odometer
                        df.at[index, 'Distance'] = distance
                        df.at[index, 'Rolling Quantity'] = rolling_quantity
                        df.at[index, 'Fuel Efficiency'] = fuel_efficiency
                        df.at[index, 'Fuel Usage'] = fuel_usage
                        df.at[index, 'Usage Type'] = usage_type

                # Reset Initial Odometer and Initial Quantity after a valid calculation
                initial_odometer = row['Odometer']
                initial_quantity = row['Quantity']
                rolling_quantity = initial_quantity
            else:
                # Accumulate Quantity for transactions where Capacity = 'N'
                rolling_quantity += row['Quantity']

    # Display the modified DataFrame in the main window
    if selected_vehicle == 'All':
        st.write("Fuel Efficiency Results for All Vehicles")
    else:
        st.write(f"Fuel Efficiency Results for {selected_vehicle}")

    st.dataframe(df)

    # Option to download the modified DataFrame as an Excel file
    def to_excel(df):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        writer.save()
        processed_data = output.getvalue()
        return processed_data

    excel_data = to_excel(df)

    st.download_button(
        label="Download Output File as Excel",
        data=excel_data,
        file_name=f'fuel_efficiency_results_{selected_vehicle}.xlsx' if selected_vehicle != 'All' else 'fuel_efficiency_results_all_vehicles.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )

else:
    st.sidebar.warning("Please upload both the main Excel file and the efficiency factors file.")
