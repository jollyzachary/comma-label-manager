# WELCOME TO COMMA

Created by Zachary Jolly

Version 1.1 - 2023

## About
Comma is a Python-based program designed to automate and streamline the process of generating labels for various items, with a specific configuration for labeling Sawzall blades. The program allows for the input of parameters regarding groups of samples, quantity in each group, and iterative procedures for each set of samples and groups. This automation greatly reduces man-hours spent on manual labeling and can be modified to suit a wide range of labeling needs.

## Setup and Usage

1. **Clone the repository** to your local machine. If you have git installed, you can do this by running the following command in your terminal or command prompt:
    ```
    git clone https://github.com/jollyzachary/comma-label-manager
    ```
    
2. **Navigate to the cloned repository**. You can do this from the terminal by running:
    ```
    cd comma-label-manager
    ```

3. **Ensure you have the required dependencies installed**. This program requires PyQt5 and openpyxl. You can install these with pip:
    ```
    pip install PyQt5 openpyxl
    ```

4. **Run the program** by executing the Comma1.1.py file. You can do this from the terminal by running:
    ```
    python Comma1.1.py
    ```

    If you have multiple versions of Python installed and python refers to Python 2.x on your machine, you might have to use `python3` instead.

## Using the Program

Once the program is running, you'll be presented with a user interface to generate labels:

- Select the blade length from the dropdown.
- Enter the number of samples.
- Enter the number of procedures.
- Click on "Run" to generate the labels.

## Updates in Version 1.1 

- Blade length selection for 6", 9", and 12" blades.
- Custom spacing option with up and down clickable arrows for fine-tuning.
- Improved combo box styling with grey selector color and white background.

## Feedback

If you encounter any issues or have any feedback, please contact [zach.j.jolly@gmail.com](mailto:zach.j.jolly@gmail.com)

Thank you for using Comma!

© 2023 Comma Inc.
