import os
import re
import sys

import pythoncom
from win32com.client import Dispatch

sys.coinit_flags = 0  # comtypes.COINIT_MULTITHREADED appel CoInitialize.
XLS_SIGNATURE: bytes = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


class MacroExecutionError(Exception):
    pass


def get_first_rows(sheet) -> list[list]:
    """Extract the first five lines of the Excel sheet."""
    return [
        [str(sheet.cell(i, j).value) for j in range(1, sheet.max_column + 1)]
        for i in range(1, min(6, sheet.max_row) + 1)
    ]


def get_xls_first_rows(sheet) -> list[list]:
    return [
        [str(cell.value) for cell in sheet.row(i)] for i in range(min(6, sheet.nrows))
    ]


def extract_macro_name(macro: str):
    return (
        match[1]
        if (match := re.search(r"#[Ss][Uu][Bb]\s+(\w+)\s*\(", macro))
        else macro.split("\n")[0].split()[-1].split("(")[0].strip()
    )


def fix_macro_issue(macro: str):
    # TODO: try to fix the vba code using llm api
    macro = macro.replace("\\'", "'")
    return macro.replace('\\"', '"').strip()


def inject_macro(file: str, macro: str):
    """This function injects and executes a macro in a new xlsm file.

    Args:
        `file` (str): An input Excel file where you want to add and run a macro.
            It will be in one of these formats xls, xlsx, xlsm, xltx or xltm.
        `macro` (str): A vba script that you want to run as a macro.

    Return
        The path to the new Excel xlsm file.
    """
    macro_name = extract_macro_name(macro).strip()

    com_instance = Dispatch(
        "Excel.Application", pythoncom.CoInitialize()
    )  # USING WIN32COM
    com_instance.DisplayAlerts = False
    objworkbook = com_instance.Workbooks.Open(os.path.join(os.getcwd(), file))
    xlmodule = objworkbook.VBProject.VBComponents.Add(1)
    xlmodule.CodeModule.AddFromString(macro.strip())
    macro = fix_macro_issue(macro)
    try:
        objworkbook.Application.Run(macro_name)
        xlsm_file = re.sub("\.\w+", "_new.xlsm", file)  # .replace('\\', "/")
        objworkbook.SaveAs(os.path.join(os.getcwd(), xlsm_file))
    except Exception as e:
        # TODO: try again after fixing the vba issue using fix_macro_issue
        raise MacroExecutionError("Cannot run macro! Macro must contain errors.") from e
    finally:
        objworkbook.Close()
        com_instance.Quit()
    return xlsm_file
