# Bond Valuation Tool

# 1. Overview
The Bond Valuation Tool is a Python application designed to perform comprehensive valuation of fixed-income securities (bonds). It calculates various key metrics for a bond, including its dirty price, clean price, accrued interest, and total market value. Additionally, it generates a detailed amortization schedule and outputs all this information into a well-formatted Excel report, complete with a visual chart of the bond's value progression over its coupon periods.
The tool is interactive, prompting the user for necessary bond parameters like face value, coupon rate, yield rate, maturity date, and settlement date.



# 2. Features
- **Comprehensive Bond Metrics**: Calculates dirty price, clean price, accrued interest, and overall bond value.
- **Bond Type Identification**: Classifies the bond as Premium, Discount, or Par.
- **Date Calculations**: Generates all coupon payment dates from settlement to maturity. Calculates the number of days between dates, days to next coupon (DSC), and days in the current coupon period (E).
- **Amortization Schedule**: Generates a detailed period-by-period amortization table showing:
                            - Beginning and ending dates for each period.
                            - Opening bond value.
                            - Interest payment.
                            - Coupon payment.
                            - Closing bond value.
- **Excel Reporting**:
Creates a professional Excel report (.xlsx) summarizing bond details and the amortization table.This includes a line chart visualizing the bond's value over time.

- **User-Friendly Interface**: Command-line prompts for easy input of bond parameters.
- **Custom Rounding**: Implements a specific rounding logic for numerical precision in financial calculations (note: see details in section 8.1).



# 3. Requirements
The script relies on the following Python libraries:
- 'datetime': For handling dates and times (standard library).
- 'calendar': For date-related calculations, specifically monthrange (standard library).
- 'math': For mathematical operations like ceil (standard library).
- 'xlsxwriter': For creating and writing to Excel files. This is an external library and needs to be installed.
- 'random': For generating a random number for the Excel file name (standard library).



# 4. Installation
Before running the script, you need to install the 'xlsxwriter' library if you haven't already. You can install it using pip:
'''
pip install xlsxwriter
'''

All other required libraries (datetime, calendar, math, random) are part of the standard Python library and do not require separate installation.



# 5. How to Use

# 5.1. Running the Script
Save the Python code as a .py file (e.g., Bond_Valuation_Tool.py). Open a terminal or command prompt. Navigate to the directory where you saved the file.
Run the script using the Python interpreter:
'''
python Bond_Valuation_Tool.py
'''

# 5.2. Input Parameters
The script will prompt you to enter the following details for the bond:
- **Face Value of the Bond**: The nominal or par value of the bond (e.g., 1000, 100000).
- **Annual Coupon Rate of the Bond**: The annual interest rate paid by the bond issuer, expressed as a decimal (e.g., 0.05 for 5%).
- **Annual Yield Rate of the Bond (YTM)**: The annual yield to maturity or market discount rate, expressed as a decimal (e.g., 0.06 for 6%).
- **Coupon Frequency**: How many times per year coupons are paid:
    - 12 for Monthly
    - 4 for Quarterly
    - 2 for Semi-annually
    - 1 for Annually
- **Maturity Date of the Bond**: The date when the bond expires and the face value is repaid. Format: YYYY-MM-DD (e.g., 2030-12-31).
- **Settlement/Valuation Date of the Bond**: The date on which the bond is being valued or traded. Format: YYYY-MM-DD (e.g., 2023-06-15).



# 6. Output

# 6.1. Console Output
After you provide the inputs, the script will print a summary of the calculated bond metrics to the console:
```
------------------------------------
Bond Valuation Report
------------------------------------

Face Value of the Bond : 100000
Annual Coupon Rate of the Bond (please enter in decimal form. Eg: 0.00) : 0.05
Annual Yield Rate of the Bond (please enter in decimal form. Eg: 0.00) : 0.06
Coupon Frequency (Monthly = 12, Quaterly = 4, Semi-annually=2, Annually=1) : 2
Maturity Date of the Bond (YYYY-MM-DD) : 2028-12-31
Settlement/Valuation Date of the Bond (YYYY-MM-DD) : 2023-07-15

------------------------------------
------------------------------------
Dirty Price:         95.9369
Accrued Interest:    0.1366
Clean Price:         95.7003
Bond Type:           Discounted Bond
------------------------------------

------------------------------------
Bond Value:          95,936.90
Settlement Date       2023-07-15
Last Coupon Date      2023-06-30
Next Coupon Date      2023-12-31
Maturity Date         2028-12-31
------------------------------------
Excel report generated and saved as 100000.0_20281231_xxxx.xlsx
Press Enter to close the window...

```

(Note: xxxx in the filename will be a random 4-digit number. Values are illustrative.)

# 6.2. Excel Report
An Excel file will be automatically generated and saved in the same directory where the script is run.
- **File Naming Convention**: 'BondFaceValue_MaturityDate_RandomNumber.xlsx' (e.g., 100000.0_20281231_1234.xlsx).
- **Content**: The Excel report ("Bond Report" sheet) includes:
    - Bond Summary Section:
        - Face Value (Original and Price per 100)
        - Bond Value (Market Value and Dirty Price per 100)
        - Clean Price (Value and Price per 100)
        - Accrued Interest (Value and Price per 100)
        - Key Dates: Settlement, Previous Coupon, Next Coupon, Maturity.
        - Other Details: Number of Coupons, Yield Rate, Coupon Rate, Coupon Frequency, Bond Type.
    - Bond Amortization Table:
        A detailed table with columns: No (Coupon Number), Beginning Date, Open Bond Value, Interest Payment, Coupon Payment, Closing Bond Value, End Date.The row corresponding to the settlement date is highlighted in red.
    - Bond Value Progression Chart:
        A line chart visualizing the "Open Bond Value" from the amortization table against "Beginning Date".




# 7. Code Deep Dive
The script is structured into a helper function, two main classes (Bond and ExcelReport), and a main execution block.

# 7.1. round_number(number, decimal_places) function
- **Purpose**: This function implements a custom rounding logic for floating-point numbers.
- **Parameters**:
    - number (float): The input number to be rounded.
    - decimal_places (int): The number of decimal places to round to.
- **Logic**:
    Converts the number to a string and splits it into integer and decimal parts.
    Iterates from the rightmost decimal digit, moving leftwards towards the desired decimal_places.
    In each step, it removes the last digit. If this digit is 5 or greater, it increments the new last digit.
    Reconstructs the number from the modified integer and decimal parts.
    This custom rounding attempts to "round half up" consistently from the right.

## 7.2. Bond Class
This class encapsulates all the properties and calculations related to a specific bond.

### 7.2.1. Initialization (__init__) and Core Attributes

- **Constructor Parameters**:
    - face_value (float): Nominal value of the bond.
    - coupon_rate (float): Annual coupon rate (decimal).
    - yield_rate (float): Annual yield to maturity (decimal).
    - coupon_frequency (int): Number of coupon payments per year.
    - maturity_date (datetime.date): Bond's maturity date.
    - settlement_date (datetime.date): Date of valuation.

- **Key Attributes Initialized**:
    'self.base_price' = 100: A standard reference price for quoting bond prices.
    
    'self.base_coupon_payments': Coupon payment amount per self.base_price (i.e., per 100 units of face value), calculated by __calculate_coupon_payments().
    
    self.coupon_dates: A dictionary generated by __generate_coupon_dates(). It contains:
    
    'settlement_date': The provided settlement date.
    
    'coupon_dates': A sorted list of all coupon payment dates from the coupon date immediately preceding or on the settlement date, up to and including the maturity date.
    
   'self.no_of_payment': Total number of coupon payments remaining from the first coupon date after or on the settlement date until maturity. Calculated by __calculate_no_compounding_periods().
    
    'self.DSC (Days to Settlement of Coupon)': Number of days from the settlement date to the next coupon payment date.
    
    'self.E (Days in Coupon Period)': Number of days in the coupon period in which the settlement date falls (i.e., days between the previous/current coupon date and the next coupon date).
    
    'self.dirty_price': The bond's price per 100 units of face value, including accrued interest. Calculated by __calculate_bond_price().
    
    'self.bond_value': The total market value of the bond (dirty_price * face_value / base_price). Calculated by __get_bond_value().
    
    'self.accrued_int': Accrued interest per 100 units of face value. Calculated by __accrued_int().
    
    'self.clean_price': The bond's price per 100 units of face value, excluding accrued interest (dirty_price - accrued_int).
    
    'self.bond_type': Categorizes the bond as 'Premium Bond' (coupon rate > yield rate), 'Discounted Bond' (coupon rate < yield 
    rate), or 'Par Bond' (coupon rate = yield rate).
    
    'self.compound_frequency': A string representation of the coupon frequency (e.g., "Semi-Annual"). Calculated by __compound_period().
    
    'self.modified_duration' = "": Initialized as an empty string. Note: This attribute is not currently calculated or used by the application.

### 7.2.2. Key Financial Calculations
__calculate_coupon_payments(self, face_value):
Calculates the periodic coupon payment amount: (face_value * coupon_rate) / coupon_frequency.
Used to calculate self.base_coupon_payments (with face_value=100) and actual coupon payments in the amortization table.

__calculate_bond_price(self, coupon_pmt, redemption, no_of_payment, DSC, E):
Calculates the bond's dirty price (present value of all future cash flows) per self.base_price (100).
    Formula:
        PV_Redemption = Redemption / ((1 + (Yield_Rate / Freq)) ^ (N - 1 + (DSC / E)))
        PV_Coupons = Sum [Coupon_Pmt / ((1 + (Yield_Rate / Freq)) ^ (k - 1 + (DSC / E)))] for k = 1 to N
        Dirty Price = PV_Redemption + PV_Coupons
    Where :
        coupon_pmt: Periodic coupon payment (for base price 100).
        redemption: Redemption value at maturity (typically self.base_price).
        no_of_payment (N): Number of remaining coupon payments.
        DSC: Days from settlement to the next coupon date.
        E: Days in the coupon period where settlement occurs.
        Yield_Rate: Periodic yield (self.yield_rate / self.coupon_frequency).
        The result is rounded using the custom round_number function.

__accrued_int(self, coupon_pmt, DSC, E):
Calculates the accrued interest per self.base_price (100).
    Formula: Coupon_Pmt * (E - DSC) / E
    This represents the portion of the next coupon payment that has "accrued" between the last coupon date and the settlement date.
    The result is rounded using the custom round_number function.

Clean Price: Calculated as self.dirty_price - self.accrued_int.

### 7.2.3. Date Handling
__add_months(self, sourcedate, months):
A utility function to add or subtract a specified number of months from a given datetime.date object.
Handles month-end conventions correctly (e.g., adding 1 month to Jan 31st results in Feb 28th/29th).

__generate_coupon_dates(self):
Generates all relevant coupon payment dates.
Starts from the maturity_date and works backward by months_interval = 12 / coupon_frequency.
Continues adding past coupon dates until a date earlier than or on the settlement_date is found.
The list of coupon dates is then sorted chronologically.
Returns a dictionary: {'settlement_date': self.settlement_date, 'coupon_dates': [list_of_dates]}. The first date in coupon_dates is the coupon date immediately preceding or on the settlement date.

__calculate_no_compounding_periods(self, start_date, end_date, compounding_frequency):
Calculates the total number of compounding (coupon) periods between a start_date and end_date.
It determines the total months and divides by the number of months per period (12 / compounding_frequency), rounding up (math.ceil) to ensure all partial periods are counted.

__no_of_days_between_dates(self, start_date, end_date):
A simple utility to calculate the absolute number of days between two datetime.date objects.

### 7.2.4. Amortization Table (get_bond_amortization_table)

- **Purpose**: Generates a period-by-period breakdown of the bond's value, interest, and coupon payments.
- **Logic**:
        - Retrieves all coupon dates and the settlement date.
        -Initializes an amortization_table list.
        - Backward Calculation for Opening Values: It iterates backward from the maturity date. For each coupon date, it        calculates the bond's dirty price as if that coupon date were the valuation date with DSC=1, E=1 (simplification for on-coupon-date valuation). This price is converted to the bond's opening value for that period.
        - The n-1 in self.__calculate_bond_price(base_coupon_payment, redemption, n-1, 1, 1) refers to the number of future payments from that point.
        - Forward Calculation for Payments: It then iterates through the table (excluding the last entry, which is maturity):
        - Interest Payment = Next Period's Open Bond Value - Current Period's Open Bond Value + Actual Coupon Payment. This is derived from the accounting identity: Opening Balance + Interest - Payment = Closing Balance.
        - Coupon Payment is the actual periodic coupon payment based on the bond's face_value.
        - Closing Bond Value is the Open Bond Value of the next period.
        - Settlement Date Adjustment: If the settlement_date is not a coupon payment date:
        - The first period in the table (from the previous coupon date to the next coupon date) is split into two:
        - One sub-period from the previous coupon date to the settlement_date.
        - Another sub-period from the settlement_date to the next coupon date.
        - Interest and coupon payments are adjusted accordingly for these partial periods. The first partial period up to settlement shows zero coupon payment (as it's not yet paid).
        - The bond value at the settlement_date is self.bond_value (the initially calculated market value).

### 7.2.5. Other Helper Methods
__compound_period(self):
Returns a string describing the compounding frequency (e.g., "Monthly", "Quarterly", "Semi-Annual", "Annual") based on self.coupon_frequency.

__get_bond_value(self, face_value, dirty_price):
Calculates the total market value of the bond: (dirty_price * face_value) / self.base_price.

## 7.3. ExcelReport Class
This class is responsible for generating the Excel output file.

### 7.3.1. generate_excel_report(...)
- **Parameters**: Takes numerous parameters, including calculated bond metrics (face_value, bond_value, dirty/clean prices,     accrued interest, rates, dates, etc.) and the amortization_table.

- **Functionality**:
    - File Creation: Creates a new Excel workbook and a worksheet named "Bond Report". The filename is  BondFaceValue_MaturityDate_RandomNumber.xlsx.
    - Formatting: Defines various cell formats (header, topics, currency, date, percentages, alignment, colors) for a professional look.
    - Hides Gridlines: Improves visual appeal.
    - Sets Column Widths: Adjusts column widths for better readability.
    - Writes Bond Summary: Populates cells with the bond's summary information, using the defined formats. This includes:
        - Face Value, Bond Value, Clean Price, Accrued Interest (both actual values and per 100 base price).
        - Key dates (Settlement, Previous/Next Coupon, Maturity).
        - Other details (No. of Coupons, Yield/Coupon Rates, Frequency, Bond Type).
    - Writes Amortization Table:
        - Writes headers for the amortization table.
        - Iterates through the amortization_table data, writing each row to the worksheet.
        - Applies currency and date formatting.
        - Highlights the row corresponding to the settlement date with a distinct format (red text, bold).
        - Adds a final row for the maturity date showing the final bond value (which should be the face value if held to maturity, or the redemption value used in calculations).
    - Creates Line Chart:
        - Adds a line chart to the worksheet.
        - Series Name: "Bond Value".
        - Categories (X-axis): "Beginning Date" column from the amortization table.
        - Values (Y-axis): "Open Bond Value" column from the amortization table.
        - Customizes line and marker appearance.
        - Sets chart title and axis labels.
        - Inserts the chart below the amortization table.
        - Hides Unused Columns/Rows: Cleans up the view by hiding columns beyond the report content.
        - Saves Workbook: Closes and saves the Excel file.
        - Prints a confirmation message to the console.

## 7.4. Main Execution Block (if __name__ == "__main__":)
This block runs when the script is executed directly.

- **User Input**:
    - Prints a header for the "Bond Valuation Report".
    - Prompts the user to enter all necessary bond parameters (face value, rates, frequency, dates).
    - Converts date strings to datetime.date objects.

- **Bond Object Creation**:
Creates an instance of the Bond class: my_Bond = Bond(...). This triggers all the internal calculations within the Bond object.

- **Data Retrieval**:
Extracts the calculated metrics (dirty price, bond value, accrued interest, clean price, dates, etc.) from the my_Bond object attributes.

- **Console Output**:
Prints a formatted summary of the key calculated bond metrics to the console.

- **Excel Report Generation**:
Creates an instance of the ExcelReport class: my_excel_report = ExcelReport().
Calls the generate_excel_report() method, passing all the necessary data to create the Excel file.

- **Pause**:
input("Press Enter to close the window...") keeps the console window open after execution until the user presses Enter, allowing them to review the output.



# 8. Underlying Financial Concepts (Briefly)

- **Dirty Price vs. Clean Price**:
    - Dirty Price (Full Price): The actual price paid for a bond. It includes the present value of all future cash flows, including any interest that has accrued since the last coupon payment.
    - Clean Price (Quoted Price): The price of a bond excluding accrued interest. Bond prices are typically quoted clean in the market.
    - Relationship: Dirty Price = Clean Price + Accrued Interest.
    - Accrued Interest: The interest earned on a bond since the last coupon payment date but not yet paid to the bondholder. If a bond is sold between coupon payment dates, the buyer usually compensates the seller for the accrued interest.
    - Bond Valuation Formula: The price of a bond is the sum of the present values of all its expected future cash flows (coupon payments and the final principal repayment). These cash flows are discounted using the yield to maturity (YTM) as the discount rate. The formula used in __calculate_bond_price accounts for settlement dates that fall between coupon payments using the DSC/E factor for fractional periods.
