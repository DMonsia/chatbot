XLS_SIGNATURE: bytes = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


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
