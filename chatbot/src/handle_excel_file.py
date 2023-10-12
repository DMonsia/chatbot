def get_first_row(sheet) -> list[list]:
    """Extract the first five lines of the Excel sheet."""
    return [
        [str(sheet.cell(i, j).value) for j in range(1, sheet.max_column + 1)]
        for i in range(1, 7)
    ]
