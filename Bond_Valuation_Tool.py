import datetime
import calendar
import math
import xlsxwriter
import random


def round_number(number, decimal_places):
    """
    Rounds a floating-point number to a given number of decimal places
    using a custom rounding logic.

    Parameters:
    number (float): The input number to be rounded.
    decimal_places (int): The number of decimal places to round to.

    Returns:
    float: The rounded number.
    """

    # Convert the number to string and split into integer and decimal parts
    integer_part, decimal_part = str(number).split(".")

    # Convert the decimal part into a list of individual digits (as integers)
    decimal_digits = list(map(int, decimal_part))

    # Determine how many decimal places need to be processed
    rounding_iterations = len(decimal_digits) - decimal_places

    # Traverse from the rightmost digit to the required decimal place
    for _ in range(rounding_iterations):
        # Remove the last digit and check if it's 5 or greater
        if decimal_digits.pop(-1) >= 5:
            decimal_digits[-1] += 1  # Round up the next left digit

    # Convert the modified decimal part back to a string
    rounded_decimal_part = "".join(map(str, decimal_digits))

    # Combine the integer and decimal part to form the final number
    rounded_result = f"{integer_part}.{rounded_decimal_part}"

    return float(rounded_result)


class Bond:
  def __init__(self, face_value, coupon_rate, yield_rate, coupon_frequency, maturity_date, settlement_date):
    self.face_value = face_value
    self.coupon_rate = coupon_rate
    self.yield_rate = yield_rate
    self.coupon_frequency = coupon_frequency
    self.maturity_date = maturity_date
    self.settlement_date = settlement_date
    self.base_price = 100

    self.base_coupon_payments = self.__calculate_coupon_payments(self.base_price) #Compute the coupon payment base on the frequancy
    self.coupon_dates = self.__generate_coupon_dates() # This is dictionary which has settelemnt dates and all coupon dates from most recent coupon dates to maturity dates
    self.no_of_payment = self.__calculate_no_compounding_periods(self.coupon_dates['coupon_dates'][0], self.coupon_dates['coupon_dates'][-1], self.coupon_frequency) # No of coupon payments
    self.DSC = abs(self.__no_of_days_between_dates(settlement_date, self.coupon_dates['coupon_dates'][1]))       #No of days from settelement date to next coupon date
    self.E = abs(self.__no_of_days_between_dates(self.coupon_dates['coupon_dates'][0], self.coupon_dates['coupon_dates'][1]))  # number of days in coupon period in which the settlement date falls.
    self.dirty_price = self.__calculate_bond_price(self.base_coupon_payments,self.base_price,self.no_of_payment,self.DSC,self.E)
    self.bond_value = self.__get_bond_value(self.face_value, self.dirty_price) # Total value of the bond
    self.accrued_int= self.__accrued_int(self.base_coupon_payments,self.DSC,self.E) # Accrued interest base on 100
    self.clean_price = self.dirty_price - self.accrued_int    # Clean Price
    self.bond_type = 'Premium Bond' if self.coupon_rate > self.yield_rate else 'Discounted Bond' if self.coupon_rate < self.yield_rate else 'Par Bond'
    self.compound_frequency = self.__compound_period()  # Get compound frequency
    self.modified_duration = ""


  def __calculate_coupon_payments(self,face_value):
    """Calculate the coupon payment amount based on the face value, coupon rate, and coupon frequency.

    Args:
        face_value (float): The face value of the bond.

    Returns:
        float: The coupon payment amount.
    """
    coupon_pmt = face_value * self.coupon_rate / self.coupon_frequency
    return coupon_pmt

  def __compound_period(self):
    """Calculate the compound period based on the coupon frequency.

    Returns:
        str: The compound period (e.g., "Monthly", "Quarterly", "Semi-Annual", "Annual").
    """
    compound_period = "Monthly" if self.coupon_frequency == 12 else "Quarterly" if self.coupon_frequency == 4 else "Semi-Annual" if self.coupon_frequency == 2 else "Annual"
    return compound_period


  def __calculate_bond_price(self,coupon_pmt,redemption,no_of_payment,DSC,E):
    """Calculate the present value of the bond cash flows.

    Args:
        coupon_pmt (float): The coupon payment amount.

    Returns:
        dict: A dictionary containing the present value of the redemption and coupon payments.
    """

    pv_redemption = redemption / ((1 + (self.yield_rate/self.coupon_frequency)) ** (no_of_payment-1+(DSC/E)))

    pv_coupon = 0
    for n in range(1, no_of_payment+1):
      pv_coupon += coupon_pmt / ((1 + (self.yield_rate/self.coupon_frequency)) ** (n-1+(DSC/E)))

    result = round_number((pv_redemption + pv_coupon),4)
    return result


  def __accrued_int(self,coupon_pmt,DSC,E):
    """Calculate the accrued interest on the bond.

    Args:
        coupon_pmt (float): The coupon payment amount.
        DSC (int): The number of days since the last coupon payment.
        E (int): The number of days in the coupon period.

    Returns:
        float: The accrued interest amount.
    """
    return round_number((coupon_pmt * (E-DSC) / E),4)


  def __add_months(self,sourcedate, months):
    """
    Adds (or subtracts) months to a given date.
    Args:
        sourcedate (datetime.date): The source date.
        months (int): The number of months to add or subtract.

    Returns:
        datetime.date: The resulting date after adding or subtracting months.
    """
    month = sourcedate.month - 1 + months
    year = sourcedate.year + month // 12
    month = month % 12 + 1
    day = min(sourcedate.day, calendar.monthrange(year, month)[1])
    return datetime.date(year, month, day)


  def __generate_coupon_dates(self):
    """
    Generates a list of coupon payment dates (as datetime.date objects) from settlement until maturity.
    Assumes coupons are paid at regular intervals.

    Args:
        maturity_date (datetime.date): The maturity date of the bond.
        settlement_date (datetime.date): The settlement date of the bond.
        coupon_frequency (int): The number of coupon payments per year.

    Returns:
        dict: A dictionary containing the settlement date and a list of coupon payment dates.
    """
    result_coupon_dates = {'settlement_date': self.settlement_date,'coupon_dates':[self.maturity_date]}
    current_date = self.maturity_date
    #coupon_dates.append(current_date)

    months_interval = 12 // self.coupon_frequency
    # Generate coupon dates backwards from maturity until before the settlement date
    while True:
        current_date = self.__add_months(current_date,-months_interval)
        result_coupon_dates['coupon_dates'].append(current_date)
        if current_date <= self.settlement_date:
            break
    result_coupon_dates['coupon_dates'].sort()
    return result_coupon_dates


  def __calculate_no_compounding_periods(self, start_date, end_date, compounding_frequency):
    """Calculates the number of compounding periods between two dates.

    Args:
        start_date (datetime.date): The start date.
        end_date (datetime.date): The end date.
        compounding_frequency (int): The number of times compounding occurs per year.

    Returns:
        int: The number of compounding periods.
    """
    total_months = (end_date.year - start_date.year) * 12 + (end_date.month - start_date.month)
    compounding_periods = int(math.ceil(total_months / (12 / compounding_frequency))) # calculate the no periods between two dates and round up the result
    return compounding_periods


  def __no_of_days_between_dates(self,start_date, end_date):
    """Calculates the number of days between two dates.

    Args:
        start_date (datetime.date): The start date.
        end_date (datetime.date): The end date.

    Returns:
        int: The number of days between the two dates.
    """
    return (end_date - start_date).days


  def __get_bond_value(self,face_value, dirty_price):
    """
    Get the bond value based on the dirty price.

    Args:
        dirty_price (float): The dirty price of the bond.

    Returns:
        float: The bond value.
    """
    return dirty_price * face_value / self.base_price


  def get_bond_amortization_table(self):
    """
    Calculate the amortization table for the bond.

    Returns:
        list: A list of lists containing the amortization table data.
        Each inner list represents a row in the table and contains the following elements:
        - Coupon Number
        - Beginning Date
        - Open Bond Value
        - Interest Payment
        - Coupon Payment
        - End Bond Value
        - End Date
    """
    coupon_dates = self.coupon_dates['coupon_dates']                            # All coupon dates until maturiry
    settlement_date = self.coupon_dates['settlement_date']                      # Settlement date
    redemption = self.base_price                                                # Base redemption price which is 100
    base_coupon_payment = self.base_coupon_payments                             # Coupon payment base on face value 100
    face_value = self.face_value                                                # Actual Face value
    coupon_payment = self.__calculate_coupon_payments(face_value)               # Actual coupon payment base on actual face value

    # create initial amortization shedule with date and bond value
    amortization_table = []
    no_of_payment= len(coupon_dates)
    for n in range(no_of_payment, 0,-1):
      dirty_price = self.__calculate_bond_price(base_coupon_payment,redemption,n-1,1,1)
      bond_value = self.__get_bond_value(face_value, dirty_price)
      date=coupon_dates.pop(0)
      coupon_no = no_of_payment-n
      amortization_table.append([coupon_no, date, bond_value])

    # append each row of amortization shedule with interest payament, coupon payment, end bond value and end date
    no_of_periods = len(amortization_table)
    for index, row in enumerate(amortization_table[:-1]):
      next_bond_value = amortization_table[index+1][2]
      current_bond_value = row[2]
      interest = next_bond_value - current_bond_value + coupon_payment
      next_coupon_date = amortization_table[index+1][1]
      amortization_table[index].append(interest)
      amortization_table[index].append(coupon_payment)
      amortization_table[index].append(next_bond_value)
      amortization_table[index].append(next_coupon_date)

    # Delete the last item in the amortization shedule which is the face value. This is added to the end value of the previoust row
    del amortization_table[-1]

    if settlement_date != amortization_table[0][1]:
      coupon_no = round_number(self.DSC/self.E,2)

      # Change the interest, coupon,end value and end date of first coupon to adjest the settelement date
      intrest_to_settle_date = self.bond_value - amortization_table[0][2]
      amortization_table[0][3] = intrest_to_settle_date
      amortization_table[0][4] = 0
      amortization_table[0][5] = self.bond_value
      amortization_table[0][6] = settlement_date

      # Add new entry to settlement date values
      intrest_to_next_coupon_date = amortization_table[1][2] - self.bond_value  + coupon_payment
      end_bond_value = amortization_table[1][2]
      next_coupon_date = amortization_table[1][1]
      amortization_table.insert(1,[coupon_no,settlement_date,self.bond_value,intrest_to_next_coupon_date,coupon_payment,end_bond_value,next_coupon_date])

    return amortization_table


# ========================================================================================================================================================================================================================================================================
class ExcelReport:

    def generate_excel_report(
                                self,face_value, bond_value, dirty_price, clean_price,
                                accrued_int, yield_rate, coupon_rate, coupn_freq,
                                settlement_date, previous_coupon_date, next_coupon_date, maturity_date,
                                bond_type, no_of_coupon=0, modified_duration=0, amortization_table=None
                                ):
        """
        Creates an Excel report containing:
        - A bond summary (face value, dirty price, clean price, coupon interest, modified duration).
        - An amortization table listing coupon dates, bond values, and coupon interest.
        - A line chart of bond value progression over the coupon dates.

        The Excel file is saved automatically using the naming convention:
        BondFaceValue_MaturityDate_RandomNumber.xlsx
        Args:
            face_value (float): The face value of the bond.
            bond_value (float): The bond value.
            dirty_price (float): The dirty price of the bond.
            clean_price (float): The clean price of the bond.
            accrued_int (float): The accrued interest of the bond.
            yield_rate (float): The yield rate of the bond.
            coupon_rate (float): The coupon rate of the bond.
            coupn_freq (int): The coupon frequency of the bond.
            settlement_date (datetime.date): The settlement date of the bond.
            previous_coupon_date (datetime.date): The previous coupon date of the bond.
            next_coupon_date (datetime.date): The next coupon date of the bond.
            maturity_date (datetime.date): The maturity date of the bond.
            bond_type (str): The type of the bond.
            no_of_coupon (int, optional): The number of coupons of the bond. Defaults to 0.
            modified_duration (float, optional): The modified duration of the bond. Defaults to 0.
            amortization_table (list, optional): The amortization table of the bond.
        """
        file_name = f"{face_value}_{maturity_date.strftime('%Y%m%d')}_{random.randint(1000,9999)}.xlsx"
        workbook = xlsxwriter.Workbook(file_name)
        worksheet = workbook.add_worksheet("Bond Report")

        accrued_int_value = face_value/100*accrued_int
        clean_price_value = face_value/100*clean_price


        # Hide gridlines
        worksheet.hide_gridlines(2)

        # Define formatting for headers, dates, and numbers
        header_format = workbook.add_format({'bold': True,'font_size':20, 'font_color': 'white', 'bg_color': '#76933c','align': 'center'})
        report_topic_format = workbook.add_format({'bold': True,'font_size':11, 'font_color': 'white', 'bg_color': '#769042','align': 'center'})
        report_topic_format_2 = workbook.add_format({'bold': True,'font_size':11, 'font_color': 'white', 'bg_color': '#769042','align': 'left'})
        report_descriptions_format = workbook.add_format({'bold': True,'font_size':11, 'font_color': 'black', 'bg_color': '#d8e4bc'})
        amortization_table_topic_format = workbook.add_format({'bold': True,'font_size':11, 'font_color': 'white', 'bg_color': '#556B26','align': 'center'})
        date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
        currency_format = workbook.add_format({'num_format': '#,##0.00_);(#,##0.00)'})
        right_align = workbook.add_format({'align': 'right'})
        settlement_price_format = workbook.add_format({'bold': True,'font_color': 'red'})
        date_format_red = workbook.add_format({'bold': True,'num_format': 'yyyy-mm-dd','font_color': 'red'})
        currency_format_red = workbook.add_format({'bold': True,'num_format': '#,##0.00_);(#,##0.00)','font_color': 'red'})
        ratio_format = workbook.add_format({'num_format': '0.00%'})


        # Write bond summary details
        worksheet.merge_range('A1:G1', "Bond Valuation Report", header_format)
        worksheet.set_row(0, 25)  # Row indexing starts from 0. set o row height to 25

        # Set the width of column A to 24
        worksheet.set_column(0, 0, 24)

        # Set the width of column B-G to 18
        worksheet.set_column(1, 6, 18)


        worksheet.write("A3", "Descriptions", report_topic_format_2)
        worksheet.write("B3", "Original Value", report_topic_format)
        worksheet.write("C3", "Price (Base 100)", report_topic_format)

        worksheet.write("A4", "Face Value", report_descriptions_format)
        worksheet.write("B4", face_value, currency_format)
        worksheet.write("C4", 100, currency_format)

        worksheet.write("A5", "Bond Value", report_descriptions_format)
        worksheet.write("B5", bond_value, currency_format)
        worksheet.write("C5", dirty_price, currency_format)

        worksheet.write("A6", "Clean Price", report_descriptions_format)
        worksheet.write("B6", clean_price_value, currency_format)
        worksheet.write("C6", clean_price, currency_format)

        worksheet.write("A7", "Accrued Interest", report_descriptions_format)
        worksheet.write("B7", accrued_int_value, currency_format)
        worksheet.write("C7", accrued_int, currency_format)

        worksheet.write("A9", "Descriptions", report_topic_format_2)
        worksheet.write("B9", "Date", report_topic_format)

        worksheet.write("A10", "Settlement Date", report_descriptions_format)
        worksheet.write("B10", settlement_date, date_format)
        worksheet.write("A11", "Previouts Coupon Date", report_descriptions_format)
        worksheet.write("B11", previous_coupon_date, date_format)
        worksheet.write("A12", "Next Coupon Date", report_descriptions_format)
        worksheet.write("B12", next_coupon_date, date_format)
        worksheet.write("A13", "Maturity Date", report_descriptions_format)
        worksheet.write("B13", maturity_date, date_format)

        worksheet.write("A15", "Descriptions", report_topic_format_2)
        worksheet.write("B15", "Details", report_topic_format)

        worksheet.write("A16", "No of Coupon", report_descriptions_format)
        worksheet.write("B16", no_of_coupon, currency_format)
        worksheet.write("A17", "Yield Rate", report_descriptions_format)
        worksheet.write("B17", yield_rate, ratio_format)
        worksheet.write("A18", "Coupn Rate", report_descriptions_format)
        worksheet.write("B18", coupon_rate, ratio_format)
        worksheet.write("A19", "Coupon  Frequncy", report_descriptions_format)
        worksheet.write("B19", coupn_freq, right_align)
        worksheet.write("A20", "Bond Type", report_descriptions_format)
        worksheet.write("B20", bond_type,right_align)

        # Write Amortization Table headers
        worksheet.merge_range('A23:G23', "Bond Amortization Table", amortization_table_topic_format)
        worksheet.write("A24", "No", report_topic_format)
        worksheet.write("B24", "Beginning Date", report_topic_format)
        worksheet.write("C24", "Open Bond Value", report_topic_format)
        worksheet.write("D24", "Interest Payment", report_topic_format)
        worksheet.write("E24", "Coupon Payment", report_topic_format)
        worksheet.write("F24", "Closing Bond Value", report_topic_format)
        worksheet.write("G24", "End Date", report_topic_format)

        # Write table data
        start_row = 24
        end_row = 0
        for i, row in enumerate(amortization_table):
            # Convert date to datetime.datetime for Excel (if necessary)
            open_date = datetime.datetime.combine(row[1], datetime.time()) if isinstance(row[1], datetime.date) else row[1]
            end_date = datetime.datetime.combine(row[6], datetime.time()) if isinstance(row[6], datetime.date) else row[6]

            # Use different formating to settlement date values
            if row[1] == settlement_date:
                worksheet.write_number(start_row + i, 0, row[0],currency_format_red)
                worksheet.write_datetime(start_row + i, 1, open_date, date_format_red)
                worksheet.write_number(start_row + i, 2, row[2], currency_format_red)
                worksheet.write_number(start_row + i, 3, row[3], currency_format_red)
                worksheet.write_number(start_row + i, 4,-row[4], currency_format_red)
                worksheet.write_number(start_row + i, 5, row[5], currency_format_red)
                worksheet.write_datetime(start_row + i, 6, end_date, date_format_red)
            else:
                worksheet.write_number(start_row + i, 0, row[0])
                worksheet.write_datetime(start_row + i, 1, open_date, date_format)
                worksheet.write_number(start_row + i, 2, row[2], currency_format)
                worksheet.write_number(start_row + i, 3, row[3], currency_format)
                worksheet.write_number(start_row + i, 4,-row[4], currency_format)
                worksheet.write_number(start_row + i, 5, row[5], currency_format)
                worksheet.write_datetime(start_row + i, 6, end_date, date_format)

            end_row = start_row + i
        end_row +=1
        worksheet.write_number(end_row, 0, amortization_table[-1][0]+1)
        worksheet.write_datetime(end_row, 1, amortization_table[-1][-1],date_format)
        worksheet.write_number(end_row, 2, amortization_table[-1][-2],currency_format)


        # Create a line chart for Bond Value progression
        chart = workbook.add_chart({'type': 'line'})
        num_rows = len(amortization_table)
        chart.add_series({
            'name':       'Bond Value',
            'categories': [ "Bond Report", start_row, 1, start_row + num_rows, 1 ],
            'values':     [ "Bond Report", start_row, 2, start_row + num_rows, 2 ],
            'line':   {'color': '#769042'}, # setting the line colour
            'marker': {'type': 'square', 'size': 5,'border': {'color': 'black'}, 'fill':{'color': 'red'}}
        })
        chart.set_title({'name': 'Bond Value'})
        chart.set_x_axis({'name': 'Coupon Dates'})
        chart.set_y_axis({'name': 'Bond Value'})

        # Insert the chart into the worksheet
        chart_row = f"A{end_row + 6}"
        worksheet.insert_chart(chart_row, chart,{'x_scale': 1.5, 'y_scale': 1.5})


        # Hide columns from I to the end (104 is max column index in xlsxwriter)
        worksheet.set_column(8, 16383, None, None, {'hidden': True})
        # worksheet.set_row(chart_row+30, 1048576, None, None, {'hidden': True})

        workbook.close()
        print(f"Excel report generated and saved as {file_name}")


#===================================================================================================================
if __name__ == "__main__":

    print("------------------------------------")
    print("Bond Valuation Report")
    print("------------------------------------")
    print()

    face_value = float(input("Face Value of the Bond : "))
    coupon_rate = float(input("Annual Coupon Rate of the Bond (please enter in decimal form. Eg: 0.00) : "))
    yield_rate = float(input("Annual Yield Rate of the Bond (please enter in decimal form. Eg: 0.00) : "))
    coupon_frequency = int(input("Coupon Frequency (Monthly = 12, Quaterly = 4, Semi-annually=2, Annually=1) : "))
    maturity_date = input("Maturity Date of the Bond (YYYY-MM-DD) : ")
    settlement_date = input("Settlement/Valuation Date of the Bond (YYYY-MM-DD) : ")
    maturity_date = datetime.datetime.strptime(maturity_date, "%Y-%m-%d").date()
    settlement_date = datetime.datetime.strptime(settlement_date, "%Y-%m-%d").date()


    my_Bond=Bond(face_value,coupon_rate, yield_rate, coupon_frequency, maturity_date, settlement_date)

    dirty_price = my_Bond.dirty_price
    bond_value = my_Bond.bond_value
    accrued_int = my_Bond.accrued_int
    clean_price = my_Bond.clean_price
    settlement_date = my_Bond.settlement_date
    last_coupon_date = my_Bond.coupon_dates['coupon_dates'][0]
    next_coupon_date = my_Bond.coupon_dates['coupon_dates'][1]
    maturity_date = my_Bond.maturity_date
    amortization_table = my_Bond.get_bond_amortization_table()
    no_of_coupon = my_Bond.no_of_payment
    modified_duration = my_Bond.modified_duration
    bond_type = my_Bond.bond_type
    compound_frequency= my_Bond.compound_frequency

    print()
    print("------------------------------------")
    print("------------------------------------")
    print(f"Dirty Price:         {dirty_price:,.4f}")
    print(f"Accrued Interest:    {accrued_int:,.4f}")
    print(f"Clean Price:         {clean_price:,.4f}")
    print(f"Bond Type:           {my_Bond.bond_type}")
    print("------------------------------------")
    print()
    print("------------------------------------")
    print(f"Bond Value:        {bond_value:,.2f}")
    print(f"Settlement Date      {settlement_date}")
    print(f"Last Coupon Date     {last_coupon_date}")
    print(f"Next Coupon Date     {next_coupon_date}")
    print(f"Maturity Date        {maturity_date}")
    print("------------------------------------")


    my_excel_report = ExcelReport()
    my_excel_report.generate_excel_report(face_value, bond_value, dirty_price, clean_price, accrued_int, yield_rate,coupon_rate, compound_frequency, settlement_date, last_coupon_date, next_coupon_date, maturity_date,bond_type, no_of_coupon, modified_duration, amortization_table)

    input("Press Enter to close the window...")