import streamlit as st
import csv

# Set the title of the app
st.title('Comma Label Generator')

# Create layout and input fields
layout = st.empty()

# Number of groups
num_groups = st.number_input('Enter number of groups:', min_value=1, step=1)

# Number of samples
num_samples = st.number_input('Enter number of samples:', min_value=1, step=1)

# Number of procedures
num_procedures = st.number_input('Enter number of procedures:', min_value=1, step=1)

# Blade length selection
blade_length = st.selectbox('Select blade length:', ['6"', '9"', '12"', 'Custom'])

if blade_length == 'Custom':
    custom_spaces = st.number_input('Enter custom spaces:', min_value=0, step=1)
else:
    custom_spaces = None

# Run button
run_button = st.button('Run')

# Check if the Run button is clicked
if run_button:
    # Create the output table
    output_table = []

    # Loop through each group
    for group_num in range(1, num_groups + 1):
        # Loop through each procedure
        for p in range(1, num_procedures + 1):
            # Loop through each sample in the group
            for s in range(1, num_samples + 1):
                # Generate the label
                label = f"{group_num}-{p:02}{s:02}"
                if blade_length == 'Custom':
                    label += ' ' * custom_spaces + f"{group_num}-{p:02}{s:02}"
                else:
                    label += ' ' * (int(blade_length[:-1]) - 2) + f"{group_num}-{p:02}{s:02}"

                # Add the label to the output table
                output_table.append([label])

    # Display the output table
    st.table(output_table)

