from pathlib import Path
from datetime import datetime
import shutil
import tkinter as tk
from tkinter import filedialog
from llama_cloud_services import LlamaParse
from llama_index.core import SimpleDirectoryReader
import nest_asyncio
import requests
import io
import sys

PACKAGE_ROOT = Path(__file__).resolve().parent
TEMP_FOLDER = PACKAGE_ROOT.parent / "temp"
DATA_FOLDER = PACKAGE_ROOT.parent / "data"

API_KEY = "llx-..."

HEADER = {
    'accept': 'application/json',
    'Content-Type': 'multipart/form-data',
    'Authorization': 'Bearer ' + API_KEY,
}

nest_asyncio.apply()


def select_pdf_file():
    """Opens a file dialog to select a PDF file."""
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="Select PDF File", filetypes=[("PDF Files", "*.pdf")])
    
    return file_path

# set up parser
def set_parser():
    '''
    Setup parser with instruction to parse handwritten documents with tables

    Input: None
    Output: LlamaParse object
    '''

    parser = LlamaParse(
        api_key=API_KEY,
        result_type="markdown",
        premium_mode=True,  
        content_guideline_instruction=
        """
        These are handwritten files.
        These files contain tables.
        Tables are also handwritten. Some fully, some partially.
        Sometimes, the writting would be cursive so you must be extra carefull with those.
        Each file may contain a totally different handwriting so be careful about that too.
        These tables can have different structure from each other.
        Extract only the table from the files.
        If a column does not have a name, name it as a single space ' '. Do not remove the column.
        Do not mixup columns or make extra columns. Read writing and names properly
        Each page may have two table in exact same format. Extract them in a single sheet. Second table is just remaining rows of first table on the same page.
        Each page has a date and title on top of it, do not extract it as a separate table in a separate sheet. It should be in a single sheet.
        Extract each table into a seperate JSON looking format unless two tables are in exact same format (same column titles and handwriting).
        Maintain actual table format
        Final json must have all tables of PDF. 
        The final result should be a spreadsheet, each sheet corresponding to a single page
        """
    )

    return parser


#Parse files
def parse_files(parser: LlamaParse, file_path: str):
    '''
    Use the preset parser to read the given file 
    and extract table out of it in JSON
    format

    Input:
    parser: LlamaParse | LlamaParse object predefined with invoice specific values
    path: str          | Path to pdf files or folder

    '''
    time_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f'temp_{time_stamp}_document.pdf'
    shutil.copy(file_path, TEMP_FOLDER / file_name)

    file_extractor = {".pdf": parser}
    
    captured_output = io.StringIO()
    sys.stdout = captured_output  # Redirect stdout

    try:
        SimpleDirectoryReader(input_dir=TEMP_FOLDER, file_extractor=file_extractor).load_data()
         # Call test function
    finally:
        sys.stdout = sys.__stdout__  # Restore original stdout
        print("File parsed, downloading data...")  # Retrieve and process output
    
    output = captured_output.getvalue()
    job_id = output.split('\n')[0]
    job_id = job_id.split(' ')[-1].strip()

    url = f'https://api.cloud.llamaindex.ai/api/parsing/job/{job_id}/result/xlsx'

    try:
        response = requests.get(url=url, headers=HEADER)

    except:
        print("Could not download file, try again later")
        return
        
    return response

def remove_cache():
    try:
        for file in TEMP_FOLDER.iterdir():
            if file.is_file():
                file.unlink()

        print("All files deleted successfully.")
    except OSError:
        print("Error occurred while deleting files.")


def pdf2excel():
    '''
    Generate an excel file from tables stored in a PDF file

    Each page of PDF file is stored in a separate sheet in the excel file
    '''
    file_path = select_pdf_file()
    file_name = Path(file_path).name.split('.')[0]

    parser = set_parser()
    response = parse_files(parser=parser, file_path=file_path)

    if response:
        with open(DATA_FOLDER / f'{file_name}.xlsx', 'wb') as output:
            output.write(response.content)
            print(f"Saved file as {file_name}.xlsx in data folder")

    remove_cache()

if __name__ == '__main__':
    pdf2excel()
