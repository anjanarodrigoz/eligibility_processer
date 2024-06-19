# Eligibility List Processor

Eligibility List Processor is a Python application that processes student marks and determines their eligibility based on specified weights and maximum marks for each subject. It also allows for the exclusion of certain students based on an exclusion list.

## Features

- **Upload Functionality**: Upload student marks and exclusion list from Excel files.
- **Weight and Max Marks Input**: Specify weights and maximum marks for each subject.
- **Eligibility Calculation**: Automatically calculates final marks and determines eligibility.
- **Save Results**: Save the processed eligibility list as an Excel file.
- **GUI Interface**: Easy-to-use graphical interface built with Tkinter.
- **Cross-Platform**: Runs on both Windows and macOS.


## File Formats
Student Marks Excel
The student marks file should be an Excel file with the following structure:

Index Number	Student Name	Subject1	Subject2	...
1	John Doe	85	90	...
2	Jane Smith	78	88	...

Exclusion List Excel
The exclusion list file should be an Excel file with the following structure:

Index Number
1
3
5


## Contributing
Contributions are welcome! Please open an issue or submit a pull request on GitHub.


### Prerequisites

- Python 3.x
- Tkinter (usually included with Python)
- Pandas
- PyExcel
- PyExcel-ODS3
- PyInstaller (for creating executables)

### Clone the Repository

```sh
git clone https://github.com/yourusername/eligibility-list-processor.git
cd eligibility-list-processor
