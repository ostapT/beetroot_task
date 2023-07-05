# PDF to Excel Converter
This project is a PDF to Excel converter. It allows you to extract text blocks from a PDF document, merge them, and save them into an Excel file.

## Installation
1. Clone the repository:
```shell
git clone https://github.com/ostapT/beetroot_task.git
```
2. Navigate to the project directory:
```shell
cd beetroot_task
```
3. Set virtual environment:
python3 -m venv venv
venv\Scripts\activate (Windows)
source venv/bin/activate (MacOS | Linux)
4. Install dependencies:
```shell
pip install requirements.txt
```
## Usage
Update the config.ini file with the file paths:

- Set the pdf_file key to the path of the PDF document.
- Set the excel_file key to the desired output path for the Excel file.

Run the main.py script to perform the PDF to Excel conversion:
```shell
python main.py
```
The script will extract the text blocks from the PDF, merge them, and save them into an Excel file.