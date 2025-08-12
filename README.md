# Automated Excel Reporting Tool

This project provides a solution to automate the process of generating reports from Excel data. It specifically targets the creation of pivot tables and bar charts, streamlining the workflow for analysis and presentation.

## Description

The `00_one_report.py` script is designed to take data from an Excel file, create a pivot table, and then generate a corresponding bar chart. This process, which typically requires manual effort in spreadsheet software, is automated to save time and reduce the potential for human error.

## Getting Started

### Dependencies

- Python 3.8 or higher
- Libraries: pandas, openpyxl
- Microsoft Windows 10 or higher (the project may work on other operating systems but has not been tested)

### Setup

Clone the repository to your local machine:

```bash
git clone https://github.com/yourusername/yourprojectname.git
cd yourprojectname
```

## Installation
Ensure you have Python installed and then set up a virtual environment:

```bash
python -m venv env
env\Scripts\activate  # On Windows
source env/bin/activate  # On Unix or MacOS
```

### Install the required packages:

```bash
pip install pandas openpyxl
```
### Executing the Program
Run the script with:

```bash
python 00_one_report.py
```

### Features
Automated pivot table creation from raw Excel data.
Bar chart generation based on the pivot table data.
Output saved as an enhanced Excel file with the pivot table and chart included.

### Help
If you encounter any issues, check the db directory to ensure your Excel data file is named correctly and placed in the correct path.

### Contributing
Contributions to this project are welcome. Please fork the repository and submit a pull request with your suggested changes.

### License
This project is licensed under the MIT License - see the LICENSE file for details.

### Acknowledgments
Thanks to Frank Andrade
Automate with Python

