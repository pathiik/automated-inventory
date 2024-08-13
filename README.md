# Automated Inventory Addition Script
This script automates the process of logging into the Stockpile inventory management system, navigating to the item addition page, and adding multiple new items to the inventory based on data provided in an Excel file.

## Prerequisites
Before running this script, ensure that you have the following installed:
1. **Python 3.x**
2. **Google Chrome**
3. **Chromedriver**: Ensure that the `chromedriver` version is compatible with your installed Chrome version.
4. **pip**: Install the required Python libraries using `pip`.

## Required Python Libraries
Install the necessary libraries by running the following command:
```bash
pip install selenium 
pip install openpyxl 
pip install python-dotenv
```

## Usage
1. Clone this repository.
```bash
git clone https://github.com/pathiik/automated-inventory.git
```
2. Create a `.env` file in the same directory as the script and define the following environment variables:\
**(Replace the placeholders with your actual values)**
```bash
DRIVER_EXECUTABLE_PATH = <path_to_your_chromedriver>
WEB_LINK = <stockpile_web_link>
LOGIN_EMAIL = <your_stockpile_login_email>
LOGIN_PASSWORD = <your_stockpile_account_password>
FILE_PATH = <path_to_your_excel_file>
```

### Code Overview
**1. Initial Setup `initializing`:**
- Loads environment variables from the `.env` file.
- Initilizes the Chrome WebDriver.
- Navigates to the Stockpile login page.

**2. Login Function `login`:**
- Locates the email and password fields on the login page.
- Enters the provided login credentials and clicks the login button.

**3. Navigate to Item Addition Page `get_to_item_add_page`:**
- Clicks the *`add`* button to navigate to the item addition section.

**4. Navigate to "Add New" Page `get_to_add_new_page`:**
- Clicks the *`Add New`* link to open the form for adding a new item.

**5. Fetching Product Data `get_product_data`:**
- Opens the specified Excel file and retrives the product data.
- Retrives data from specific columns for a given number of rows and returns the data as a list of dictionaries.

**6. Adding New Items `adding_item`:**
- Iterates through each product's data and fills out the necessary fields on the form:
    - Adding the product name `new_item_name`.
        - Clicks on *`Name`* input field and enters the product name.
    - Adding the product SKU `new_item_sku`.
        - Clicks on *`SKU`* input field and enters the product SKU.
    - Selecting the product label `new_item_label`.
        - Clicks on the *`Labels`* dropdown and selects the product label.
    - Selecting the product manufacturer `new_item_manufacturer`.
        - Clicks on the *`Manufacturer`* dropdown and selects the product manufacturer.
    - Selecting the product transaction date `new_item_date`.
        - Clicks on the *`Transaction Date`* input field and selects the specific transaction date. (In this case, 28th day of 2 months ago).
    - Adding the product quantity `new_item_quantity`.
        - Clicks on the *`Quantity`* input field and enters the product quantity.
    - Adding the product cost `new_item_cost`.
        - Clicks on the *`Unit Cost ($)`* input field and enters the product cost.
- Clicks on the *`Save and Add Another`* button to save the new item and add another item `save_and_next`.

**7. Holding Page and Closing `hold_page`:**
- Holds the page for 20 seconds before quitting the WebDriver.

**8. Main Function `main`:**
- Calls the `initializing` function to set up the WebDriver.
- Calls the `login` function to log into the Stockpile account.
- Calls the `get_to_item_add_page` function to navigate to the item addition page.
- Calls the `get_to_add_new_page` function to navigate to the "Add New" page.
- Calls the `get_product_data` function to fetch the product data from the Excel file.
- Calls the `adding_item` function to add the new items to the inventory.
- Calls the `hold_page` function to hold the page for 20 seconds before quitting the WebDriver.


### Running the Script
1. Ensure that the `.env` file is correctly set up with the required environment variables.
2. Run the script using the following command:\
**(Replace `<script_name>` with the name of the script file.)**
```bash
python <script_name>.py
```

### Notes
1. Ensure the Excel file is properly formatted and contains the expected data in the correct columns.
2. If any issues occurs, the script will print an error message to the console, and the WebDriver will close automatically.

### Support
If you like this project, give it a ⭐ and share it with friends!

### Feedback
If you have any feedback or queries, please reach out to me at pathik.b45@gmail.com

###### Pathik Bhattarai