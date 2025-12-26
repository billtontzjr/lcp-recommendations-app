"""Frequency string parsing utilities."""
import re

FREQUENCY_MULTIPLIERS = {
    "yearly": 1.0,
    "1x/year": 1.0,
    "1 time per year": 1.0,
    "2x/year": 2.0,
    "2 times per year": 2.0,
    "3 times per year": 3.0,
    "4 times per year": 4.0,
    "12 times a year": 12.0,
    "24 times a year": 24.0,
    "monthly": 12.0,
    "every 2 years": 0.5,
    "every 3 years": 0.333333,
    "every 5 years": 0.2,
    "every 8-10 years": 0.111111,
    "one time": 0.0,
    "one-time": 0.0,
}


def parse_frequency(freq_str):
    """
    Parse frequency string and return (is_annual, multiplier).

    Args:
        freq_str: Frequency string like "2x/year", "every 5 years", "one time"

    Returns:
        Tuple of (is_annual, multiplier):
        - is_annual: True if recurring, False if one-time
        - multiplier: Annual multiplier (e.g., 0.5 for every 2 years)
    """
    if not freq_str:
        return (False, 0)

    freq_lower = str(freq_str).lower().strip()

    # Check for one-time
    if "one time" in freq_lower or "one-time" in freq_lower:
        return (False, 1.0)

    # Check for visits pattern (e.g., "24 visits every 5 years")
    visits_match = re.search(r'(\d+)\s*visits?\s*(?:every\s*)?(\d+)?\s*years?', freq_lower)
    if visits_match:
        visits = int(visits_match.group(1))
        years = int(visits_match.group(2)) if visits_match.group(2) else 1
        return (True, visits / years)

    # Check for "X times per year" pattern
    times_match = re.search(r'(\d+)\s*(?:times?|x)\s*(?:per|a|/)\s*year', freq_lower)
    if times_match:
        return (True, float(times_match.group(1)))

    # Check for "every X years" pattern
    every_match = re.search(r'every\s*(\d+)(?:-(\d+))?\s*years?', freq_lower)
    if every_match:
        if every_match.group(2):
            # Range like "every 8-10 years" - use midpoint
            low = int(every_match.group(1))
            high = int(every_match.group(2))
            midpoint = (low + high) / 2
            return (True, 1.0 / midpoint)
        else:
            years = int(every_match.group(1))
            return (True, 1.0 / years)

    # Check standard patterns
    for pattern, mult in FREQUENCY_MULTIPLIERS.items():
        if pattern in freq_lower:
            if mult == 0.0:
                return (False, 1.0)
            return (True, mult)

    # Default to yearly if unrecognized
    return (True, 1.0)
