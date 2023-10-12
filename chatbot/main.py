from io import BytesIO
from typing import Annotated

from fastapi import Body, FastAPI, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from openpyxl import load_workbook
from src.api_llm import conversation_with_powerbi
from src.handle_excel_file import get_first_row
from src.prompts import _prompt_sys_template, format_data
from src.utils import get_substring

app = FastAPI(
    title="API ChatBot",
    description="""**API ChatBot** est une IA qui permet de lire, traiter et modifier des fichiers Excel via l'injection de macros VBA générées par OpenAi.""",
    version="0.0.1",
    contact={
        "name": "YellowSys",
        "email": "mdougban@yellowsys.fr",
    },
)
origins = ["*"]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
app.mount("/data", StaticFiles(directory="data"), name="data")


@app.post("/handle_excel_file", tags=["QA"])
def handle_excel_file(
    username: Annotated[
        str,
        Body(
            title="A username",
            description="A valid username for the yellowsys llm api.",
        ),
    ],
    password: Annotated[
        str,
        Body(
            title="The user password",
            description="The user's password for using the yellowsys llm api.",
        ),
    ],
    query: Annotated[
        str,
        Body(
            title="The user query",
            description="The user query containing all the changes to be applied to the Excel file.",
        ),
    ],
    file: Annotated[
        bytes,
        File(
            title="The excel file to handle",
            description="The bytes object contains the Excel file you want to process.",
        ),
    ],
):
    """
    Generate VBA code using a yellowsys llm api and inject it into the excel file.

    Args:<br>
        username (str): A valid username for the yellowsys llm api.<br>
        password (str): The user's password for using the yellowsys llm api.<br>
        query (str): The user query containing all the changes to be applied to the Excel file.
        file (bytes): The bytes object contains the Excel file you want to process.

    Returns:<br>
        file_name (str): The path to the new Excel file to download.<br>
    """
    wb = load_workbook(filename=BytesIO(file))
    # We assume that the data are on the first sheet.
    # To view all existing sheets, use wb.sheetnames
    sheet_name = wb.sheetnames[0]
    sheet = wb[sheet_name]
    # Select the frist 5 rows
    first_rows = get_first_row(sheet)
    sys_role = _prompt_sys_template.format(
        sheet_name=sheet_name, first_rows=format_data(first_rows)
    )
    prompt = sys_role + """\n\n{history} \n\nHuman: {input}\n\nAssistant:"""
    response = {
        "response": "conversation_with_powerbi(prompt, query, username, password)"
    }
    vba_script = "get_substring(response['response'], start='Sub', end='End Sub')"
    with open("./data/history.csv", "a") as f:
        f.write(f"{query}[SEP]{vba_script}[SEP]{response['response']}[EOR]\n")

    file_name = "./data/output.xlsx"
    with open(file_name, "wb") as f:
        f.write(file)
    return {"url": file_name}
