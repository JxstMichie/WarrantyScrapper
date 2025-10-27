from datetime import datetime
import re

def format_date(date_str: str) -> str:
    date_str = date_str.strip()
    date_str = re.sub(r'(?i)^ending on\s+', '', date_str).strip()

    for fmt in ("%B %d, %Y", "%Y-%m-%d", "%d %B %Y", "%b %d, %Y"):
        try:
            parsed_date = datetime.strptime(date_str, fmt)
            return parsed_date.strftime("%d.%m.%Y")
        except ValueError:
            continue

    return date_str