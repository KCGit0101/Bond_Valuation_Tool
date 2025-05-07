# Bond Valuation Tool

## Overview

The **Bond Valuation Tool** is a Python-based application designed to calculate the present value of bonds, determine yield to maturity (YTM), and analyze various bond-related metrics. This tool is useful for financial analysts, investors, and students who want to understand the valuation of fixed-income securities.

---

## Features

1. **Present Value Calculation**: Computes the present value of a bond based on its coupon payments, face value, and discount rate.
2. **Yield to Maturity (YTM)**: Estimates the YTM, which is the internal rate of return (IRR) for the bond.
3. **Cash Flow Analysis**: Displays the bond's cash flow schedule.
4. **Customizable Inputs**: Allows users to input bond parameters such as coupon rate, face value, market price, and maturity period.
5. **Error Handling**: Ensures robust handling of invalid inputs and edge cases.

---

## Prerequisites

- Python 3.8 or higher
- Required libraries:
    - `numpy`
    - `pandas` (optional, for exporting data)
    - `matplotlib` (optional, for visualizations)

Install dependencies using:
```bash
pip install numpy pandas matplotlib
```

---

## Usage

### 1. Input Parameters
The tool requires the following inputs:
- **Face Value**: The bond's nominal value (e.g., $1,000).
- **Coupon Rate**: The annual interest rate paid by the bond (e.g., 5%).
- **Market Price**: The current price of the bond in the market.
- **Years to Maturity**: The remaining time until the bond matures.
- **Discount Rate**: The rate used to discount future cash flows (optional).

### 2. Running the Application
Run the script using:
```bash
python bond_valuation_tool.py
```

### 3. Output
The application provides:
- Present value of the bond.
- Yield to maturity (YTM).
- A detailed cash flow schedule.
- Optional: Graphical representation of cash flows.

---

## Code Explanation

### 1. `bond_valuation_tool.py`
This is the main script that ties all functionalities together.

#### Key Functions:
- **`calculate_present_value()`**:
    - Inputs: Face value, coupon rate, discount rate, years to maturity.
    - Outputs: Present value of the bond.
    - Formula: 
        \[
        PV = \sum \frac{C}{(1 + r)^t} + \frac{FV}{(1 + r)^T}
        \]
        Where:
        - \(C\): Coupon payment
        - \(r\): Discount rate
        - \(t\): Time period
        - \(FV\): Face value
        - \(T\): Total periods

- **`calculate_ytm()`**:
    - Inputs: Face value, coupon rate, market price, years to maturity.
    - Outputs: Approximate yield to maturity.
    - Uses numerical methods (e.g., Newton-Raphson) to solve for YTM.

- **`generate_cash_flow_schedule()`**:
    - Inputs: Face value, coupon rate, years to maturity.
    - Outputs: A list of cash flows for each period.

- **`plot_cash_flows()`**:
    - Inputs: Cash flow schedule.
    - Outputs: A bar chart visualizing cash flows over time.

---

## Example

### Input:
- Face Value: $1,000
- Coupon Rate: 5%
- Market Price: $950
- Years to Maturity: 10

### Output:
- Present Value: $950.23
- Yield to Maturity: 5.42%
- Cash Flow Schedule:
    | Year | Coupon Payment | Total Cash Flow |
    |------|----------------|-----------------|
    | 1    | $50            | $50             |
    | 2    | $50            | $50             |
    | ...  | ...            | ...             |
    | 10   | $50 + $1,000   | $1,050          |

---

## Future Enhancements

1. Add support for zero-coupon bonds.
2. Include inflation-adjusted bond valuation.
3. Export results to Excel or CSV.
4. Add a GUI for user-friendly interaction.

---

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.

---

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request with your changes.

---

## Contact

For questions or feedback, please contact:
- **Email**: your_email@example.com
- **GitHub**: [Your GitHub Profile](https://github.com/your-profile)
