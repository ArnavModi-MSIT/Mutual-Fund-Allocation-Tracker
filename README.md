# Portfolio Manager

A Python-based tool for managing and analyzing mutual fund portfolio data.

## Overview

Portfolio Manager is a command-line application that helps you track and analyze mutual fund investments. It processes Excel data files containing portfolio information and provides analysis of fund performance across different time periods.

## Features

- Import portfolio data from Excel files
- Track mutual fund investments across multiple months
- Calculate and analyze changes in:
  - Fund quantity
  - Market value
  - Percentage of NAV
- Identify new fund additions to the portfolio
- Persistent data storage using JSON

## Prerequisites

- Python 3.7 or higher
- Required Python packages:
  ```
  pandas
  openpyxl
  ```

## Installation

1. Clone the repository or download the source code
2. Install required packages:
   ```bash
   pip install pandas openpyxl
   ```

## Excel File Format Requirements

The Excel file should follow this structure:
- Data should start from row 7
- Required columns (in order from column C to H):
  1. Name (Fund Name)
  2. ISIN (Fund ISIN)
  3. Industry
  4. Quantity
  5. Market Value
  6. Percent NAV

## Usage

### Starting the Application

Run the main script:
```bash
python portfolio_manager.py
```

### Menu Options

1. **Import Excel Data**
   - Select option 1 from the main menu
   - Enter the path to your Excel file
   - Provide the month and year (e.g., "January 2024")

2. **Analyze Fund Performance**
   - Select option 2 from the main menu
   - Enter the fund name to search
   - Specify start and end months for comparison
   - View detailed analysis results

3. **Exit**
   - Select option 3 to exit the program

### Example Usage

```python
# Import data
Enter your choice (1-3): 1
Enter Excel file path: portfolio_jan_2024.xlsx
Enter month and year: January 2024

# Analyze performance
Enter your choice (1-3): 2
Enter fund name to search: HDFC
Enter start month: January 2024
Enter end month: February 2024
```

## Data Storage

- Portfolio data is stored in `portfolio_data.json`
- Data is organized by month-year
- Each fund entry contains:
  - Fund details (Name, ISIN, Industry)
  - Monthly data (Quantity, Market Value, NAV percentage)

## Error Handling

The application includes validation for:
- Month-year format
- Excel file structure
- Data types
- File existence
- Data integrity

## Output Format

Analysis results include:
- Fund identification details
- Status (Existing/New Addition)
- Changes in:
  - Quantity
  - Market value
  - Percentage to NAV
- Percentage changes for each metric

## Limitations

- Excel files must follow the specified format
- Dates must be in "Month YYYY" format
- All numerical data must be valid numbers

## Contributing

To contribute:
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## License

This project is open-source and available under the MIT License.
