"""Cost calculation service."""
from app.utils.frequency_parser import parse_frequency
from app.utils.currency import parse_cost_string


def calculate_item_costs(item, patient_info, pfr_lookup, apc_lookup):
    """
    Calculate annual and one-time costs for an item.

    Args:
        item: Item dictionary from workbook parser
        patient_info: Patient info dictionary
        pfr_lookup: PFR pricing lookup dictionary
        apc_lookup: APC pricing lookup dictionary

    Returns:
        Dictionary with unit_cost, annual_cost, one_time_cost
    """
    geo_mult = float(patient_info.get('geographic_multiplier', 1.0) or 1.0)

    # Get base cost
    if item.get('cost'):
        base_cost = parse_cost_string(item['cost'])
    else:
        base_cost = lookup_cost(
            item.get('code', ''),
            item.get('code_type', ''),
            pfr_lookup,
            apc_lookup,
            geo_mult
        )

    # Parse frequency
    is_annual, multiplier = parse_frequency(item.get('frequency', ''))

    if is_annual:
        return {
            'unit_cost': base_cost,
            'annual_cost': round(base_cost * multiplier, 2),
            'one_time_cost': 0.0
        }
    else:
        return {
            'unit_cost': base_cost,
            'annual_cost': 0.0,
            'one_time_cost': round(base_cost * multiplier, 2) if multiplier else base_cost
        }


def lookup_cost(code, code_type, pfr_lookup, apc_lookup, geo_mult=1.0):
    """
    Look up cost from PFR/APC tables.

    Args:
        code: CPT code or codes (semicolon-separated)
        code_type: 'PFR', 'APC', or combination
        pfr_lookup: PFR pricing lookup dictionary
        apc_lookup: APC pricing lookup dictionary
        geo_mult: Geographic multiplier for facility fees

    Returns:
        Total cost as float
    """
    if not code:
        return 0.0

    code_str = str(code).strip()
    code_type_str = str(code_type).upper().strip() if code_type else ''

    total_cost = 0.0

    # Handle multiple codes (semicolon-separated)
    codes = [c.strip() for c in code_str.split(';') if c.strip()]

    for cpt in codes:
        # Try to look up in PFR
        pfr_price = pfr_lookup.get(cpt, 0.0)

        # Try to look up in APC
        apc_price = apc_lookup.get(cpt, 0.0)

        if 'APC' in code_type_str or 'FACILITY' in code_type_str:
            # Facility fee - apply geographic multiplier
            total_cost += pfr_price + (apc_price * geo_mult)
        else:
            # Professional fee only
            total_cost += pfr_price

    return total_cost


def calculate_all_costs(workbook_data):
    """
    Calculate costs for all items and generate summary.

    Args:
        workbook_data: Dictionary from workbook parser with patient_info, items, etc.

    Returns:
        Dictionary with items (with costs), category_totals, and grand_totals
    """
    patient_info = workbook_data['patient_info']
    items = workbook_data['items']
    pfr_lookup = workbook_data['pfr_lookup']
    apc_lookup = workbook_data['apc_lookup']

    life_expectancy = float(patient_info.get('life_expectancy', 0) or 0)

    # Calculate costs for each item
    calculated_items = []
    category_totals = {}

    for item in items:
        costs = calculate_item_costs(item, patient_info, pfr_lookup, apc_lookup)

        calculated_item = {
            **item,
            'unit_cost': costs['unit_cost'],
            'annual_cost': costs['annual_cost'],
            'one_time_cost': costs['one_time_cost'],
        }
        calculated_items.append(calculated_item)

        # Aggregate by category
        category = item.get('category', 'Uncategorized')
        if category not in category_totals:
            category_totals[category] = {
                'annual_cost': 0.0,
                'one_time_cost': 0.0,
                'items': []
            }

        category_totals[category]['annual_cost'] += costs['annual_cost']
        category_totals[category]['one_time_cost'] += costs['one_time_cost']
        category_totals[category]['items'].append(calculated_item)

    # Calculate grand totals
    total_annual = sum(ct['annual_cost'] for ct in category_totals.values())
    total_one_time = sum(ct['one_time_cost'] for ct in category_totals.values())
    lifetime_annual = total_annual * life_expectancy
    grand_total = lifetime_annual + total_one_time

    return {
        'items': calculated_items,
        'category_totals': category_totals,
        'totals': {
            'total_annual': round(total_annual, 2),
            'total_one_time': round(total_one_time, 2),
            'lifetime_annual': round(lifetime_annual, 2),
            'grand_total': round(grand_total, 2),
            'life_expectancy': life_expectancy,
        }
    }
