"""
shorten_ecoinvent_name.py

Shortens ecoinvent 3.x dataset names to a compact format:
    "<Product>, <M|T>, {<Geography>}"

Input format (standard ecoinvent 3.x naming):
    "<Product> {<GEO>}| <activity description> | <system model>, <type>"

Classification:
    'market for' in activity description -> M (Market)
    Anything else                        -> T (Transformation)

Examples:
    "Nylon 6-6 {RoW}| market for nylon 6-6 | Cut-off, U"  ->  "Nylon 6-6, M, {RoW}"
    "Brass {CH}| market for brass | Cut-off, U"            ->  "Brass, M, {CH}"
    "Nylon 6-6 {RoW}| nylon 6-6 production | Cut-off, U"  ->  "Nylon 6-6, T, {RoW}"

Usage:
    - Edit the list in the ENTER YOUR NAMES HERE section at the bottom.
    - Run: python shorten_ecoinvent_name.py
"""

import re


def shorten_ecoinvent_name(full_name):
    """
    Parse a single ecoinvent dataset name and return a compact label.
    Raises ValueError if the name cannot be parsed.
    """

    # Split on pipe character -- standard ecoinvent names have 3 pipe-separated segments
    segments = [seg.strip() for seg in full_name.split('|')]

    if len(segments) < 2:
        raise ValueError(
            f"Expected pipe-separated segments, got: '{full_name}'\n"
            "Check that the name follows standard ecoinvent format."
        )

    product_geo_segment = segments[0]   # e.g. "Nylon 6-6 {RoW}"
    activity_segment    = segments[1]   # e.g. "market for nylon 6-6"

    # Extract geographic tag: {RoW}, {CH}, {GLO}, etc.
    geo_match = re.search(r'\{[^}]+\}', product_geo_segment)

    if not geo_match:
        raise ValueError(
            f"No geographic tag found in: '{product_geo_segment}'\n"
            "Expected a tag like {RoW} or {CH}."
        )

    geo_tag      = geo_match.group()                                  # e.g. "{RoW}"
    product_name = product_geo_segment[: geo_match.start()].strip()  # e.g. "Nylon 6-6"

    if not product_name:
        raise ValueError(
            f"Could not extract a product name from: '{product_geo_segment}'"
        )

    # Classify as Market or Transformation
    if 'market for' in activity_segment.lower():
        process_type = 'M'
    else:
        process_type = 'T'

    return f'{product_name}, {process_type}, {geo_tag}'


def shorten_ecoinvent_names_batch(name_list):
    """
    Process a list of ecoinvent dataset names.
    Returns a list of dicts with keys: 'original', 'shortened', 'error'.
    """
    results = []
    for name in name_list:
        try:
            shortened = shorten_ecoinvent_name(name)
            results.append({'original': name, 'shortened': shortened, 'error': None})
        except ValueError as e:
            results.append({'original': name, 'shortened': None, 'error': str(e)})
    return results


# =============================================================================
# ENTER YOUR NAMES HERE -- add or remove lines as needed
# =============================================================================

my_names = [
    'Nylon 6-6 {RoW}| market for nylon 6-6 | Cut-off, U',
    'Brass {CH}| market for brass | Cut-off, U',
    # Add more names here...
]

# =============================================================================
# Run and display results -- no edits needed below this line
# =============================================================================

if __name__ == '__main__':
    results = shorten_ecoinvent_names_batch(my_names)

    print(f"{'INPUT':<70}  OUTPUT")
    print('-' * 100)
    for r in results:
        truncated = (r['original'][:65] + '...') if len(r['original']) > 65 else r['original']
        if r['error']:
            print(f"{truncated:<70}  ERROR: {r['error']}")
        else:
            print(f"{truncated:<70}  {r['shortened']}")
