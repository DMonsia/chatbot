class MacroNotFound(Exception):
    pass


def get_substring(text: str, start: str, end: str) -> str:
    """Extract a substring from a text based on the start and end token included."""
    try:
        idx_start = text.lower().index(start.lower())
        idx_end = text.lower().index(end.lower())
    except Exception as e:
        raise MacroNotFound(
            "Unable to generate vba code in response to your request. Please rephrase!"
        ) from e
    return text[idx_start : idx_end + len(end) + 1]
