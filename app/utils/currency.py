"""Currency formatting utilities."""


def format_currency(amount):
    """Format a number as currency string."""
    if amount is None:
        return "$0.00"
    return f"${amount:,.2f}"


def parse_cost_string(cost_str):
    """
    Parse cost strings like "307" or "1671; 853" (multiple codes).

    Args:
        cost_str: Cost value (number or string)

    Returns:
        Float value of the cost
    """
    if cost_str is None:
        return 0.0

    if isinstance(cost_str, (int, float)):
        return float(cost_str)

    cost_str = str(cost_str).strip()

    if not cost_str:
        return 0.0

    # Handle semicolon-separated costs (sum them)
    if ';' in cost_str:
        parts = cost_str.split(';')
        total = 0.0
        for p in parts:
            cleaned = p.strip().replace(',', '').replace('$', '')
            if cleaned and cleaned.replace('.', '').replace('-', '').isdigit():
                total += float(cleaned)
        return total

    # Clean and parse single value
    cleaned = cost_str.replace(',', '').replace('$', '').strip()
    try:
        return float(cleaned)
    except ValueError:
        return 0.0
