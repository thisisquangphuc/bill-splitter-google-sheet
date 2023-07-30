# Bill Splitter with Promotions and Shipping - Google Apps Script

# Overview

The Bill Splitter with Promotions and Shipping is a Google Apps Script application designed to simplify bill splitting among multiple members within a Google Sheet. The script provides functionalities to calculate each member's share of expenses, including promotions and shipping costs. It is a helpful tool for individuals, groups, or teams who frequently engage in joint expenses and need a convenient way to distribute the costs fairly.
Key Features

    - New Bill Initialization: This function sets up a new bill by clearing and initializing the relevant cells for item prices, promotions, and shipping.

    - Clear Table: The "clearTable" function allows users to clear the entire table, making it ready for a new bill or a fresh calculation.

    - Copy Range and Format with Insertion: This feature enables users to copy a range of data from the "Input" sheet and insert it into the "Report" sheet while maintaining the same formatting.

    - Bill Calculation: The core functionality of the script lies in the "calculateBill" function, which calculates the amount each member needs to pay, considering promotions and shipping costs. It distributes the expenses fairly among the members involved.

    - Automatic Cell Conversion: The "onEdit" function automatically converts certain cell values to multiples of 1000 for convenience when entering small decimal values.

# How to Use

    1. Setting up the Spreadsheet: Create a new Google Sheet and set up the necessary columns for item names, prices, quantities, promotions, and shipping details.

    2. Initializing a New Bill: Use the "newBill" function to reset the bill values, including promotions and shipping costs, to start a new expense calculation.

    3. Calculating the Bill: Once the item details and member information are entered, execute the "calculateBill" function to calculate each member's share, considering promotions and shipping.

    4. Clearing the Table: If needed, utilize the "clearTable" function to reset the table and prepare it for a new bill or fresh calculation.

# Contributions

Contributions to the Bill Splitter with Promotions and Shipping script are welcome! If you have any ideas for improvements, additional features, or bug fixes, feel free to fork this repository and create a pull request.
# License

This project is licensed under the MIT License.

# Disclaimer

The Bill Splitter with Promotions and Shipping script is provided as-is with no warranty or guarantee. Use it at your own risk.
