def get_substring(text: str, start: str, end: str) -> str:
    """Extract a substring from a text based on the start and end token included."""
    idx_start = text.lower().index(start.lower())
    idx_end = text.lower().index(end.lower())
    return text[idx_start : idx_end + len(end) + 1]
