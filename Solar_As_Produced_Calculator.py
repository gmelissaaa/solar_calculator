import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

# Set page configuration to wide layout
st.set_page_config(layout="wide")

# Step 1: Upload Excel File
st.title('Solar As Produced Calculator')

uploaded_file = st.file_uploader('Upload Load Profile File', type=['xlsx'])

if uploaded_file:
    # Read the Excel file and extract necessary columns
    df = pd.read_excel(uploaded_file, usecols=['supply period', 'datetime', 'hour', 'wesm', 'kWh'])
    
    # Display the uploaded data
    with st.expander('Extracted Load Profile Data'):
        st.dataframe(df, use_container_width=True)

    # Step 2: Editable Solar Guarantee Percentage Table
    default_solar_guarantee = {
        'Hour': list(range(1, 25)),
        'Solar Guarantee (%)': [0, 0, 0, 0, 0, 0, 0, 29, 57, 79, 93, 100, 100, 93, 79, 57, 21, 0, 0, 0, 0, 0, 0, 0]
    }
    solar_guarantee_df = pd.DataFrame(default_solar_guarantee)

    # Transpose the DataFrame for horizontal display
    solar_guarantee_df_t = solar_guarantee_df.set_index('Hour').T

    st.write('Edit Solar Guarantee Percentage Table:')
    edited_df_t = st.data_editor(solar_guarantee_df_t, num_rows="fixed", use_container_width=True)

    # Transpose back for processing
    edited_df = edited_df_t.T.reset_index().rename(columns={'index': 'Hour'})
    edited_df['Hour'] = edited_df['Hour'].astype(int)

    # Validate Solar Guarantee Percentages
    if (edited_df['Solar Guarantee (%)'] < 0).any() or (edited_df['Solar Guarantee (%)'] > 100).any():
        st.warning('Solar Guarantee Percentages should be between 0% and 100%. Values outside this range will be adjusted.')
        # Clip values to be within range
        edited_df['Solar Guarantee (%)'] = edited_df['Solar Guarantee (%)'].clip(lower=0, upper=100)
        st.stop()

    # Merge the solar guarantee data with the load profile data
    df = df.merge(edited_df, left_on='hour', right_on='Hour', how='left')

    # Step 3: Input Data
    st.write('Input Data:')

    # Create columns for Solar Rate and Line Rental
    col1, col2 = st.columns(2)
    
    with col1:
        solar_rate = st.number_input('Solar Rate (PHP/KWH)', min_value=0.0, value=0.0, step=0.1)
    with col2:
        line_rental = st.number_input('Line Rental (PHP/KWH)', min_value=0.0, value=0.0, step=0.1)

    # Admin Fee on the second line
    admin_fee = st.number_input('Admin Fee (PHP/KWH)', min_value=0.0, value=0.0, step=0.1)

    if st.button('Calculate Charges'):
        # Step 4: Calculate Charges
        df['Solar Consumption (kWh)'] = df['kWh'] * df['Solar Guarantee (%)'] / 100
        df['Non Solar Consumption (kWh)'] = df['kWh'] - df['Solar Consumption (kWh)']
        df['Solar Charge (pHp)'] = df['Solar Consumption (kWh)'] * (solar_rate + line_rental)
        df['Non Solar Charge (pHp)'] = df['Non Solar Consumption (kWh)'] * (df['wesm'] + admin_fee)
        df['Total Charge (pHp)'] = df['Solar Charge (pHp)'] + df['Non Solar Charge (pHp)']

        # Ensure 'datetime' column is in correct format
        df['datetime'] = pd.to_datetime(df['datetime'])

        # Step 5: Create Custom Supply Period
        def get_supply_period(date):
            if date.day >= 26:
                period_start = date.replace(day=26)
                period_end = (period_start + pd.DateOffset(months=1)).replace(day=25)
            else:
                period_end = date.replace(day=25)
                period_start = (period_end - pd.DateOffset(months=1)).replace(day=26)
            return period_end.strftime('%b-%y'), period_end

        df['Supply Period'], df['Supply Period End'] = zip(*df['datetime'].apply(get_supply_period))

        # Ensure DataFrames are sorted and correctly indexed
        with st.expander('Detailed Charges'):
            detailed_charges_df = df[['supply period', 'datetime', 'hour', 'wesm', 'kWh', 'Solar Consumption (kWh)', 'Non Solar Consumption (kWh)', 'Solar Charge (pHp)', 'Non Solar Charge (pHp)', 'Total Charge (pHp)']]
            detailed_charges_df = detailed_charges_df.sort_values(by='datetime').reset_index(drop=True)
            st.dataframe(detailed_charges_df, use_container_width=True)

        # Create Pivot Table
        pivot_df = df.pivot_table(
            index=['Supply Period'],
            values=['kWh', 'Solar Consumption (kWh)', 'Solar Charge (pHp)', 'Non Solar Consumption (kWh)', 'Non Solar Charge (pHp)', 'Total Charge (pHp)'],
            aggfunc=np.sum
        ).reset_index()

        # Calculate Effective Rate (pHp/kWh) for each supply period
        pivot_df['Effective Rate (pHp/kWh)'] = pivot_df['Total Charge (pHp)'] / pivot_df['kWh']

        # Convert 'Supply Period' to datetime for sorting
        pivot_df['Supply Period Date'] = pd.to_datetime(pivot_df['Supply Period'], format='%b-%y')

        # Sort by Supply Period Date
        pivot_df = pivot_df.sort_values(by='Supply Period Date').reset_index(drop=True)

        # Drop the temporary 'Supply Period Date' column
        pivot_df.drop(columns=['Supply Period Date'], inplace=True)

        # Create Grand Total
        grand_total = pivot_df[['kWh', 'Solar Consumption (kWh)', 'Non Solar Consumption (kWh)', 'Solar Charge (pHp)', 'Non Solar Charge (pHp)', 'Total Charge (pHp)']].sum()
        grand_total['Supply Period'] = 'Grand Total'
        grand_total['Effective Rate (pHp/kWh)'] = grand_total['Total Charge (pHp)'] / grand_total['kWh']

        # Append Grand Total
        pivot_df = pd.concat([pivot_df, grand_total.to_frame().T], ignore_index=True)

        # Reorder columns
        pivot_df = pivot_df[['Supply Period', 'kWh', 'Solar Consumption (kWh)', 'Non Solar Consumption (kWh)', 'Solar Charge (pHp)', 'Non Solar Charge (pHp)', 'Total Charge (pHp)', 'Effective Rate (pHp/kWh)']]

        with st.expander('Summary Charges by Supply Period'):
            st.dataframe(pivot_df, use_container_width=True)

        # Step 6: Create Download Button
        with pd.ExcelWriter("result.xlsx", engine='xlsxwriter') as output:
            # Write Detailed Charges
            if not detailed_charges_df.empty:
                detailed_charges_df.to_excel(output, sheet_name='Detailed Charges', index=False)
            else:
                st.error("Detailed Charges DataFrame is empty. Cannot generate the Excel file.")

            # Write Summary Charges
            pivot_df.to_excel(output, sheet_name='Summary Charges', index=False)

            # Write Solar Guarantee Percentage Table as Reference Data
            edited_df.rename(columns={'Solar Guarantee (%)': 'Solar Guarantee (%)'}, inplace=True)  # Rename column for clarity
            edited_df.to_excel(output, sheet_name='Reference Data', index=False)
        
        # Provide download button
        with open("result.xlsx", "rb") as file:
            st.download_button(
                label="Download",
                data=file,
                file_name='result.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

        # Prepare data for the area chart
        hourly_consumption = detailed_charges_df.groupby('hour')[['kWh', 'Solar Consumption (kWh)']].sum().reset_index()
        hourly_consumption['Non Solar Consumption (kWh)'] = hourly_consumption['kWh'] - hourly_consumption['Solar Consumption (kWh)']

        # Plot stacked area chart with Solar Consumption in Orange
        fig = px.area(hourly_consumption, x='hour', y=['Solar Consumption (kWh)', 'Non Solar Consumption (kWh)'],
                      labels={'value': 'kWh', 'hour': 'Hour'},
                      title='Solar As Produced',
                      color_discrete_map={'Solar Consumption (kWh)': 'orange'})  # Set color for Solar Consumption

        st.plotly_chart(fig)
