# PDF to CSV Converter GUI Application

This Python script provides a graphical user interface (GUI) for converting PDF files to CSV format using the Aspose.PDF library. It allows users to select input and output files, toggle between light and dark modes, and provides feedback through logging.

## Dependencies
* Python 3.x
* tkinter
* aspose.pdf
* os

## Usage
* Run the script.
* The GUI application will open.
* Click the "Browse" button next to "Input PDF" to select a PDF file.
* Click the "Browse" button next to "Output CSV" to select the location to save the CSV file.
* Optionally, click "Store in Same Directory" to store the CSV file in the same directory as the input PDF.
* Click "Convert to CSV" to start the conversion process.
* The log will display the conversion status and details.

## Functions
* `toggle_mode()`: Toggles between light and dark modes for the GUI.
* `convert_to_csv()`: Converts the selected PDF file to CSV format.
* `store_csv_in_same_directory()`: Stores the CSV file in the same directory as the input PDF.
* `browse_input_pdf()`: Opens a file dialog to browse and select the input PDF file.
* `browse_output_csv()`: Opens a file dialog to select the location to save the output CSV file.

## GUI Elements
* Input PDF Entry: Field to display the path of the input PDF file.
* Output CSV Entry: Field to display the path of the output CSV file.
* Browse buttons: Allow users to select PDF and CSV files.
* Store in Same Directory button: Automatically sets the output CSV file path to the same directory as the input PDF.
* Convert to CSV button: Initiates the conversion process.
* Log Text Box: Displays conversion status and details.

## Troubleshooting
* Ensure Python 3.x is installed and accessible in your environment.
* Make sure tkinter and aspose.pdf libraries are installed.
* Verify that the selected input PDF file is accessible and in the correct format.
* Check file permissions and directory access if encountering issues with file handling.

## Future
* The following script was made as project done under my own leisure time. As a result, the GUI elements and all functionalities was made using the tkinter library and is made in a simple manner.
* The script will be further improved and implemented to allow converting into different formats with a more robust framework in the future.
