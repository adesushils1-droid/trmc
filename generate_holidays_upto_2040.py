"""
Generate NSE India market holidays from 2026 to 2040.

Includes fixed-date holidays (Republic Day, Independence Day, Gandhi Jayanti,
Christmas) and computed holidays (Good Friday via Easter computus).
Lunar-based Hindu/Muslim holidays (Holi, Eid, Muharram, Diwali, etc.) are
marked as approximate since they depend on moon sighting and should be
verified against the official NSE holiday calendar each year.
"""

import json
from datetime import date, timedelta
from pathlib import Path

OUTPUT_FILE = Path(__file__).parent / "nse_holidays_upto_2040.json"
START_YEAR = 2026
END_YEAR = 2040


def easter_date(year: int) -> date:
    """Compute Easter Sunday using the Anonymous Gregorian computus algorithm."""
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return date(year, month, day)


def good_friday(year: int) -> date:
    """Good Friday is the Friday before Easter Sunday."""
    return easter_date(year) - timedelta(days=2)


def fixed_holidays(year: int) -> list[dict]:
    """Fixed-date NSE holidays that occur on the same calendar date every year."""
    return [
        {"date": f"{year}-01-26", "name": "Republic Day"},
        {"date": f"{year}-08-15", "name": "Independence Day"},
        {"date": f"{year}-10-02", "name": "Gandhi Jayanti"},
        {"date": f"{year}-12-25", "name": "Christmas"},
    ]


def generate_holidays() -> list[dict]:
    all_holidays: list[dict] = []
    for year in range(START_YEAR, END_YEAR + 1):
        # Fixed holidays (exact)
        for h in fixed_holidays(year):
            all_holidays.append(h)

        # Good Friday (computed, exact)
        gf = good_friday(year)
        all_holidays.append({
            "date": gf.isoformat(),
            "name": "Good Friday",
        })

        # Lunar-based holidays (approximate - moon sighting dependent)
        # These dates should be verified against the official NSE calendar.
        all_holidays.append({
            "date": f"{year}-01-01",
            "name": "New Year (Approximate)",
            "approximate": True,
        })
        all_holidays.append({
            "date": f"{year}-03-14",
            "name": "Holi (Approximate)",
            "approximate": True,
        })
        all_holidays.append({
            "date": f"{year}-03-30",
            "name": "Ram Navami (Approximate)",
            "approximate": True,
        })
        all_holidays.append({
            "date": f"{year}-04-10",
            "name": "Mahavir Jayanti (Approximate)",
            "approximate": True,
        })
        all_holidays.append({
            "date": f"{year}-04-14",
            "name": "Dr. Baba Saheb Ambedkar Jayanti",
        })
        all_holidays.append({
            "date": f"{year}-05-01",
            "name": "Maharashtra Day",
        })
        all_holidays.append({
            "date": f"{year}-05-25",
            "name": "Buddha Purnima (Approximate)",
            "approximate": True,
        })
        all_holidays.append({
            "date": f"{year}-06-15",
            "name": "Eid-ul-Adha / Bakri Id (Approximate)",
            "approximate": True,
        })
        all_holidays.append({
            "date": f"{year}-07-06",
            "name": "Muharram (Approximate)",
            "approximate": True,
        })
        all_holidays.append({
            "date": f"{year}-08-20",
            "name": "Raksha Bandhan (Approximate)",
            "approximate": True,
        })
        all_holidays.append({
            "date": f"{year}-09-05",
            "name": "Ganesh Chaturthi (Approximate)",
            "approximate": True,
        })
        all_holidays.append({
            "date": f"{year}-10-10",
            "name": "Dussehra / Vijayadashami (Approximate)",
            "approximate": True,
        })
        all_holidays.append({
            "date": f"{year}-10-20",
            "name": "Diwali (Approximate)",
            "approximate": True,
        })
        all_holidays.append({
            "date": f"{year}-11-15",
            "name": "Gurpurab / Guru Nanak Dev Jayanti (Approximate)",
            "approximate": True,
        })
        all_holidays.append({
            "date": f"{year}-12-26",
            "name": "Boxing Day (Approximate)",
            "approximate": True,
        })

    # Sort by date and remove exact duplicates (keep first occurrence)
    seen: set[str] = set()
    unique: list[dict] = []
    for h in sorted(all_holidays, key=lambda x: x["date"]):
        if h["date"] not in seen:
            seen.add(h["date"])
            unique.append(h)

    return unique


def main():
    holidays = generate_holidays()
    data = {
        "year_range": {"start": START_YEAR, "end": END_YEAR},
        "exchange": "NSE India",
        "total_holidays": len(holidays),
        "note": "Lunar-based holidays marked 'approximate: true' must be verified against official NSE calendar each year.",
        "holidays": holidays,
    }

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    print(f"Generated {len(holidays)} holidays from {START_YEAR} to {END_YEAR}")
    print(f"Saved to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
