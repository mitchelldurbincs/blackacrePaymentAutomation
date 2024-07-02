# Data Processing GUI

This project is a Python-based graphical user interface (GUI) application for processing and analyzing financial data from Stripe and other CSV files. It provides a user-friendly interface for selecting input files, specifying date ranges, and generating a final report in Excel format.

## Features

- File selection for Stripe CSV and other CSV files
- Date range selection for data filtering
- Data cleaning and processing
- Category and code mapping based on program names
- Progress bar for tracking processing status
- Dark-themed GUI for improved visibility
- Improved error handling and user feedback
- Configurable constants for easy customization

## Requirements

- Python 3.11 or higher
- pandas
- tkinter
- tkcalendar
- openpyxl
- xlrd

## Installation

1. Clone this repository or download the source code.
2. Install the required dependencies:

```
pip install -r requirements.txt
```

## Usage

1. Run the script:

```
python main.py
```

2. Use the "Browse" buttons to select the Stripe CSV and Other CSV files.
3. Set the start and end dates for the data range you want to process.
4. Click the "Process Data" button to start the data processing.
5. The application will display the progress and status of the operation.
6. Once complete, the final report will be saved as `final_report.xlsx` in the same directory.

## Data Processing Steps

1. Loads data from the selected CSV files and the `Codes.xlsx` file.
2. Cleans and processes the Stripe data, removing failed transactions and filtering by date range.
3. Maps program names to category codes and categories.
4. Processes each row of data, combining information from Stripe and the other CSV file.
5. Calculates amounts, fees, and other financial metrics.
6. Generates a final report with consolidated information.

## Customization

You can easily customize the application by modifying the constants at the top of the script:

- `TITLE`: The title of the application window
- `BACKGROUND`: The background color for the date selection widgets
- `FOREGROUND`: The text color for the date selection widgets
- `CODE_SHEET_NAME`: The name of the sheet in the Excel file containing category codes
- `FINAL_REPORT_NAME`: The name of the output Excel file

## Troubleshooting

- Ensure all required files are in the correct location.
- Check that the CSV files have the expected column names and data formats.
- If you encounter any errors, they will be displayed in a message box with details about the issue.
- Check the console output for any additional error information that may not be displayed in the GUI.

## Future Improvements

- Implement logging for better debugging and error tracking.
- Add a data preview feature to allow users to check input files before processing.
- Implement multithreading for improved performance with large datasets.
- Add more export options (e.g., CSV format) in addition to Excel.
- Implement unit tests to ensure reliability of data processing logic.

## Contributing

Contributions to improve the application are welcome. Please feel free to submit pull requests or open issues for any bugs or feature requests.

## License

This project is licensed under the MIT License.