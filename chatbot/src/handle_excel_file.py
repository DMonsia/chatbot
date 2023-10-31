import os
import re
import sys

import pythoncom
from win32com.client import Dispatch

sys.coinit_flags = 0  # comtypes.COINIT_MULTITHREADED appel CoInitialize.


class MacroExecutionError(Exception):
    pass


class ExcelFileProcessingError(Exception):
    pass


def get_first_rows_by_sheet(file: str) -> dict[str, list]:
    """Extract the first five lines of each sheet in the Excel file."""
    com_instance = Dispatch(
        "Excel.Application", pythoncom.CoInitialize()
    )  # USING WIN32COM
    com_instance.Visible = False
    com_instance.DisplayAlerts = False
    try:
        objworkbook = com_instance.Workbooks.Open(os.path.join(os.getcwd(), file))
        rows_by_sheet = {}
        for sheet in objworkbook.Sheets:
            temp = sheet.UsedRange()
            rows_by_sheet[sheet.Name] = [[str(val) for val in row] for row in temp[:5]]
    except Exception as e:
        print("\n", e, "\n")
        raise ExcelFileProcessingError(
            "An error occurred while processing the file."
        ) from e
    finally:
        com_instance.Quit()
    return rows_by_sheet


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
        `file` (str): An input Excel file path where you want to add and run a macro.
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
        xlsm_file = re.sub("\.\w+", ".xlsm", file)
        objworkbook.SaveAs(os.path.join(os.getcwd(), xlsm_file), FileFormat=52)
    except Exception as e:
        # TODO: try again after fixing the vba issue using fix_macro_issue
        print("\n")
        print(e)
        print("\n")
        raise MacroExecutionError("Cannot run macro! Macro must contain errors.") from e
    finally:
        objworkbook.Close()
        com_instance.Quit()
    return xlsm_file
