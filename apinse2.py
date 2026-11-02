"""
NSE Stock Data Fetcher using NseIndiaApi
Handles NSE's session requirements automatically
"""

from datetime import datetime
import os
import sqlite3
from pathlib import Path
from typing import Any

from nse import NSE
from openpyxl import Workbook, load_workbook

stocks2 = [
    "HINDPETRO",
    "HINDUNILVR",
    "HINDZINC",
    "HINDWAREAP",
    "HIRECT",
    "HISARMETAL",
    "POWERINDIA",
    "HLEGLAS",
    "HLVLTD",
    "HMAAGRO",
    "HOMEFIRST",
    "HONASA",
    "HONDAPOWER",
    "HONAUT",
    "HORIZONIND",
    "HPAL",
    "HPL",
    "HSBCGOLD",
    "HUDCO",
    "HTMEDIA",
    "HUBTOWN",
    "HUHTAMAKI",
    "HTEL",
    "HYBRIDFIN",
    "HYUNDAI",
    "IGPL",
    "ISTLTD",
    "ICEMAKE",
    "ICICIBANK",
    "ICICILIFE",
    "ICICIGI",
    "LOWVOLIETF",
    "MIDSELIETF",
    "NIF100IETF",
    "NIFTYIETF",
    "SENSEXIETF",
    "ICICIAMC",
    "GOLDIETF",
    "NV20IETF",
    "ICICIB22",
    "CASHIETF",
    "EVIETF",
    "BSE500IETF",
    "GSEC5IETF",
    "ALPL30IETF",
    "AUTOIETF",
    "PVTBANIETF",
    "COMMOIETF",
    "CONSUMIETF",
    "FMCGIETF",
    "INFRAIETF",
    "LIQUIDIETF",
    "MIDCAPIETF",
    "MOM30IETF",
    "NEXT50IETF",
    "HEALTHIETF",
    "QUAL30IETF",
    "SILVERIETF",
    "ITIETF",
    "INSUREIETF",
    "BANKIETF",
    "METALIETF",
    "OILIETF",
    "PHARMAIETF",
    "PSUBNKIETF",
    "SMALLIETF",
    "TOP15IETF",
    "VAL30IETF",
    "FINIETF",
    "ICRA",
    "IDBI",
    "IDEAFORGE",
    "IDFCFIRSTB",
    "IFBAGRO",
    "IFBIND",
    "IFCI",
    "IFGLEXPOR",
    "IGARASHI",
    "IIFLCAPS",
    "IIFL",
    "IIRM",
    "IKIO",
    "IVC",
    "IMAGICAA",
    "INCREDIBLE",
    "IMPAL",
    "INDSWFTLAB",
    "INDTERRAIN",
    "INDBANK",
    "BAJAJINDEF",
    "INDGN",
    "INDIAGLYCO",
    "INDNIPPON",
    "IPL",
    "INDIASHLTR",
    "ITDC",
    "IBULLSLTD",
    "INDIAMART",
    "INDIANB",
    "INDIANCARD",
    "IEX",
    "INDIANHUME",
    "IMFA",
    "IOC",
    "IOB",
    "IRCTC",
    "IRFC",
    "IREDA",
    "IWP",
    "INDIGOPNTS",
    "INDIQUBE",
    "INDOAMIN",
    "INDOBORAX",
    "ICIL",
    "INDOFARM",
    "INDORAMA",
    "INDOTECH",
    "INDOTHAI",
    "INDOUS",
    "INDOMIM",
    "NIPPOBATRY",
    "INDOCO",
    "IGCL",
    "INDOKEM",
    "WAAREEINDO",
    "INDOSTAR",
    "INDOWIND",
    "IGL",
    "INDRAMEDCO",
    "INDUSTOWER",
    "INDUSINDBK",
    "INDPRUD",
    "IITL",
    "NAUKRI",
    "INFOBEAN",
    "INFOMEDIA",
    "INFY",
    "INGERRAND",
    "INNOVACAP",
    "INNOVANA",
    "INNOVISION",
    "INOXGREEN",
    "INOXINDIA",
    "INOXWIND",
    "INSECTICID",
    "INA",
    "INSPIRISYS",
    "ESSENTIA",
    "INTELLECT",
    "INTENTECH",
    "INTERARCH",
    "INDIGO",
    "INTLCONV",
    "IGIL",
    "INVENTURE",
    "IKS",
    "IVZINGOLD",
    "IVZINNIFTY",
    "INVPRECQ",
    "IOLCP",
    "IONEXCHANG",
    "IPRINGLTD",
    "IPCALAB",
    "IRB",
    "IRCON",
    "IRISDOREME",
    "IRIS",
    "IRMENERGY",
    "ISGEC",
    "ISHANCH",
    "ITCHOTELS",
    "ITC",
    "ITI",
    "ITL",
    "IVALUE",
    "IVP",
    "IZMO",
    "J&KBANK",
    "JGCHEM",
    "JAGRAN",
    "JAGSNPHARM",
    "JAIBALAJI",
    "JAICORPLTD",
    "JISLDVREQS",
    "JISLJALEQS",
    "JAINREC",
    "JPPOWER",
    "JAMNAAUTO",
    "JSFB",
    "JARO",
    "JASH",
    "JAYBARMARU",
    "JAYAGROGN",
    "JAYNECOIND",
    "JAYKAY",
    "JAYSREETEA",
    "JBMA",
    "JSLL",
    "JENBURPH",
    "JETFREIGHT",
    "JPOLYINVST",
    "JINDRILL",
    "JINDALPHOT",
    "JINDALPOLY",
    "JINDALSAW",
    "JSL",
    "JINDALSTEL",
    "JINDWORLD",
    "JKIPL",
    "JIOFIN",
    "JPL",
    "NIFTYJBLK",
    "JITFINFRA",
    "JKCEMENT",
    "JKLAKSHMI",
    "JKPAPER",
    "JKTYRE",
    "JKIL",
    "JMFINANCIL",
    "JNKINDIA",
    "JOCIL",
    "COCKERILL",
    "JSWCEMENT",
    "JSWDULUX",
    "JSWENERGY",
    "JSWHL",
    "JSWINFRA",
    "JSWSTEEL",
    "JTEKTINDIA",
    "JTLIND",
    "JUBLCPL",
    "JUBLFOOD",
    "JUBLINGREA",
    "JUBLPHARMA",
    "JMA",
    "JNPR",
    "JUNIPER",
    "JLHL",
    "JWL",
    "JUSTDIAL",
    "JYOTHYLAB",
    "JYOTICNC",
    "JYOTIRES",
    "JYOTISTRUC",
    "KJMCFIN",
    "KMCSHIL",
    "KMSUGAR",
    "KPEL",
    "KAJARIACER",
    "KAKATCEM",
    "KALPATARU",
    "KPIL",
    "KALYANKJIL",
    "KICL",
    "KSL",
    "KAMAHOLD",
    "KAMANWALA",
    "KAMATHOTEL",
    "KAMDHENU",
    "KAMOPAINTS",
    "KANANIIND",
    "KANCHI",
    "KANOHAR",
    "KANPRPLA",
    "KANSAINER",
    "KAPSTON",
    "KARAMTARA",
    "KARMAENG",
    "KTKBANK",
    "KARURVYSYA",
    "KAUSHALYA",
    "KSCL",
    "KAYNES",
    "KCP",
    "KDGREEN",
    "KDDL",
    "KEC",
    "KEI",
    "KELLTONTEC",
    "KENNAMET",
    "KERNEX",
    "KESORAMIND",
    "KKCL",
    "KEYFINSERV",
    "RUSTOMJEE",
    "KFINTECH",
    "KHADIM",
    "KHAITANLTD",
    "KHAICHEM",
    "KHANDSE",
    "KHAZANCHI",
    "KLBRENG-B",
    "KILITCH",
    "KINGFA",
    "KIOCL",
    "KIRANVYPAR",
    "KIRIINDUS",
    "KIRLOSBROS",
    "KECL",
    "KIRLFER",
    "KIRLOSIND",
    "KIRLOSENG",
    "KIRLPNU",
    "KITEX",
    "KNAGRI",
    "KNACK",
    "KMEW",
    "KNRCON",
    "KOHINOOR",
    "KOKUYOCMLN",
    "KOLTEPATIL",
    "KOPRAN",
    "GOLD1",
    "KOTAKBANK",
    "NIFTY1",
    "PSUBANK",
    "CHEMICAL",
    "ALPHA",
    "CONS",
    "IT",
    "LIQUID1",
    "LOWVOL1",
    "MIDCAP",
    "MNC",
    "SILVER1",
    "NV20",
    "MID150",
    "MOMENTUM30",
    "MSCIINDIA",
    "NEXT50ETF",
    "NIFTY100EW",
    "PVTBANK",
    "QUALITY30",
    "BANKNIFTY1",
    "KOTIC",
    "KOTHARIPET",
    "KOTHARIPRO",
    "KOTYARK",
    "KOVAI",
    "KPIGREEN",
    "KPITTECH",
    "KPRMILL",
    "KRBL",
    "KREBSBIO",
    "KRIDHANINF",
    "KRISHANA",
    "KRISHIVAL",
    "KRISHNADEF",
    "KIMS",
    "KRITI",
    "KRITINUT",
    "KRITIKA",
    "KRN",
    "KROSS",
    "KRSNAA",
    "KRYSTAL",
    "KSB",
    "KSE",
    "KSHINTL",
    "KSOLVES",
    "KUANTUM",
    "KUSUMGAR",
    "KPL",
    "KWIL",
    "LKPSEC",
    "LTF",
    "LTTS",
    "LAOPALA",
    "LAGNAM",
    "LAHOTIOV",
    "LFIC",
    "LAKSHMIMIL",
    "LALITHAA",
    "LAMBODHARA",
    "LANCER",
    "LANDMARK",
    "LPDC",
    "LT",
    "LASERPOWER",
    "LATENTVIEW",
    "LATTEYS",
    "LAURUSLABS",
    "LAXMICOT",
    "LAXMIDENTL",
    "LGHL",
    "LAXMIINDIA",
    "LXCHEM",
    "LCCPROJECT",
    "LEMERITE",
    "IXIGO",
    "LEAPIND",
    "THELEELA",
    "LEMONTREE",
    "LENSKART",
    "LEXUS",
    "LGBBROSLTD",
    "LGEINDIA",
    "LIBAS",
    "LIBERTSHOE",
    "LICHSGFIN",
    "LICMFGOLD",
    "LICNETFN50",
    "LICNETFSEN",
    "LICNFNHGP",
    "LICNETFGSC",
    "LICNMID100",
    "LICI",
    "LIKHITHA",
    "LINC",
    "LINCOLN",
    "LINDEINDIA",
    "LLOYDSENGG",
    "LLOYDSENT",
    "LLOYDSME",
    "LMW",
    "LODHA",
    "LCL",
    "LOKESHMACH",
    "LORDSCHLO",
    "LAL",
    "LOTUSCHO",
    "LOTUSEYE",
    "LOVABLE",
    "LOYALTEX",
    "LTFOODS",
    "LTM",
    "LUMAXTECH",
    "LUMAXIND",
    "LUMINO",
    "LUPIN",
    "LUXIND",
    "MBEL",
    "MKEXIM",
    "MKPL",
    "MCL",
    "M&MFIN",
    "MAANALU",
    "MCCHRLS-B",
    "MACPOWER",
    "MADHAVIPL",
    "MADHAV",
    "MADHUCON",
    "MBAPL",
    "MADRASFERT",
    "MAFATIND",
    "MAGADSUGAR",
    "MCLOUD",
    "MAGNUM",
    "MAHAPEXLTD",
    "MFML",
    "MGL",
    "MTNL",
    "MAHSCOOTER",
    "MAHSEAMLES",
    "MAHESHWARI",
    "M&M",
    "MAHEPC",
    "MHRIL",
    "MAHLIFE",
    "MAHLOG",
    "MAITHANALL",
    "MAKERSL",
    "MALLCOM",
    "MALUPAPER",
    "MAMATA",
    "MANAKCOAT",
    "MANINDS",
    "MANINFRA",
    "MANAKALUCO",
    "MANAKSIA",
    "MANALIPETC",
    "MANAPPURAM",
    "MANBA",
    "MANCREDIT",
    "MEIL",
    "MANGLMCEM",
    "MANGALAM",
    "MGEL",
    "MANORG",
    "MSL",
    "MWL",
    "MANIKA",
    "MANIPALHOS",
    "MPIMANIPAL",
    "MANKIND",
    "MVGJL",
    "MANOMAY",
    "MANORAMA",
    "MARALOVER",
    "MARATHON",
    "MARICO",
    "MARINE",
    "MARKOLINES",
    "MARKSANS",
    "MARSONS",
    "MARUTI",
    "MASFIN",
    "MASKINVEST",
    "MASTEK",
    "MASTERTR",
]


def get_value(data: dict[str, Any], path: tuple[str, ...], default: Any = None) -> Any:
    current = data
    for key in path:
        if not isinstance(current, dict):
            return default
        current = current.get(key)
        if current in (None, ""):
            return default
    return current


def first_value(*values: Any, default: Any = None) -> Any:
    for value in values:
        if value not in (None, ""):
            return value
    return default


def to_float(value: Any, default: float = 0.0) -> float:
    try:
        if value in (None, ""):
            return default
        return float(value)
    except (TypeError, ValueError):
        return default


def to_int(value: Any, default: int = 0) -> int:
    try:
        if value in (None, ""):
            return default
        return int(float(value))
    except (TypeError, ValueError):
        return default


def fetch_quote(nse: NSE, symbol: str) -> dict[str, Any]:
    if hasattr(nse, "quote"):
        return nse.quote(symbol)
    if hasattr(nse, "equityQuote"):
        return nse.equityQuote(symbol)
    if hasattr(nse, "equity_quote"):
        return nse.equity_quote(symbol)
    raise AttributeError("NSE object does not have quote(), equityQuote(), or equity_quote()")


def get_stock_data(nse: NSE, symbol: str) -> dict[str, Any] | None:
    try:
        print(f"Fetching {symbol} data from NSE...")
        quote = fetch_quote(nse, symbol)

        if not quote:
            print(f"NSE returned an empty response for {symbol}.")
            return None

        meta = quote.get("metaData", {})
        trade = quote.get("tradeInfo", {})
        sec = quote.get("secInfo", {})

        last_price = to_float(
            first_value(
                trade.get("lastPrice"),
                quote.get("lastPrice"),
                meta.get("closePrice"),
                meta.get("lastPrice"),
                default=0,
            )
        )

        change = to_float(meta.get("change", 0))
        percentage_change = to_float(meta.get("pChange", 0))
        market_cap = to_float(trade.get("totalMarketCap", 0))
        volume_traded = to_int(first_value(trade.get("totalTradedVolume"), trade.get("quantitytraded")))
        delivery_quantity = to_int(first_value(sec.get("deliveryQuantity"), trade.get("deliveryquantity")))
        delivery_percentage = to_float(
            first_value(
                trade.get("deliveryToTradedQuantity"),
                sec.get("deliveryTotradedQuantity"),
                default=0,
            )
        )

        return {
            "symbol": symbol,
            "company_name": meta.get("companyName", symbol),
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "last_price": last_price,
            "change": change,
            "percentage_change": percentage_change,
            "market_cap": market_cap,
            "market_cap_cr": round(market_cap / 10_000_000, 2) if market_cap else 0,
            "volume_traded": volume_traded,
            "delivery_quantity": delivery_quantity,
            "delivery_percentage": delivery_percentage,
            "open": to_float(meta.get("open", 0)),
            "high": to_float(meta.get("dayHigh", 0)),
            "low": to_float(meta.get("dayLow", 0)),
            "previous_close": to_float(meta.get("previousClose", 0)),
            "last_update_time": quote.get("lastUpdateTime"),
        }

    except Exception as e:
        print(f"Error fetching {symbol}: {e}")
        return None


def display_stock(data: dict[str, Any]):
    print("\n" + "=" * 70)
    print(f"{data['company_name']} ({data['symbol']})")
    print("=" * 70)
    print(f"As on: {data['timestamp']} IST")
    print("-" * 70)
    print(f"Last Price: Rs {data['last_price']:,.2f}")
    print(f"Change: {data['change']:+,.2f} ({data['percentage_change']:+.2f}%)")
    print(f"Open: Rs {data['open']:,.2f}")
    print(f"High: Rs {data['high']:,.2f}")
    print(f"Low: Rs {data['low']:,.2f}")
    print(f"Previous Close: Rs {data['previous_close']:,.2f}")
    print(f"Market Cap: Rs {data['market_cap_cr']:,.2f} Cr")
    print(f"Delivery %: {data['delivery_percentage']:.2f}%")
    print(f"Volume: {data['volume_traded']:,} shares")
    print(f"Last Update: {data.get('last_update_time') or data['timestamp']}")
    print("=" * 70)


def save_to_excel(sorted_results, date):
    stock_row = {}

    if os.path.exists('Stcapmark.xlsx'):
        wb = load_workbook('Stcapmark.xlsx')
        ws = wb.active
        for row in range(2, ws.max_row + 1):
            stock = ws.cell(row=row, column=1).value
            if stock:
                stock_row[stock] = row
        market_col = None
        col = 2
        while ws.cell(row=1, column=col).value is not None:
            if ws.cell(row=1, column=col).value == f'{date} Market Cap':
                market_col = col
                delivery_col = col + 1
                rank_col = col + 2
                break
            col += 3
        if market_col is None:
            col = 2
            while ws.cell(row=1, column=col).value is not None:
                col += 1
            market_col = col
            delivery_col = col + 1
            rank_col = col + 2
            ws.cell(row=1, column=market_col, value=f'{date} Market Cap')
            ws.cell(row=1, column=delivery_col, value=f'{date} Delivery Qty')
            ws.cell(row=1, column=rank_col, value=f'{date} Rank')
    else:
        wb = Workbook()
        ws = wb.active
        ws.cell(row=1, column=1, value='Stock Name')
        for row, stock in enumerate(stocks2, 2):
            ws.cell(row=row, column=1, value=stock)
            stock_row[stock] = row
        market_col = 2
        delivery_col = 3
        rank_col = 4
        ws.cell(row=1, column=market_col, value=f'{date} Market Cap')
        ws.cell(row=1, column=delivery_col, value=f'{date} Delivery Qty')
        ws.cell(row=1, column=rank_col, value=f'{date} Rank')

    for item in sorted_results:
        row = stock_row.get(item['stock'])
        if row:
            if 'element_int' in item:
                ws.cell(row=row, column=market_col, value=item['element_int'])
            if 'delivery' in item:
                ws.cell(row=row, column=delivery_col, value=item['delivery'])
            ws.cell(row=row, column=rank_col, value=item['rank'])

    wb.save('Stcapmark.xlsx')
    print(f"Data saved to Stcapmark.xlsx ({len(sorted_results)} stocks2)")


def save_to_database(sorted_results, date):
    db_path = 'stock_data.db'
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS stock_market_cap (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            stock TEXT NOT NULL,
            market_cap INTEGER,
            delivery_percentage REAL,
            rank INTEGER,
            date TEXT NOT NULL
        )
    ''')
    
    for item in sorted_results:
        cursor.execute('''
            INSERT OR REPLACE INTO stock_market_cap (stock, market_cap, delivery_percentage, rank, date)
            VALUES (?, ?, ?, ?, ?)
        ''', (item['stock'], item.get('element_int'), item.get('delivery'), item.get('rank'), date))
    
    conn.commit()
    conn.close()
    print(f"Data saved to {db_path}")


def main():
    print("\nNSE Stock Data Fetcher - Multiple Stocks")

    results = []
    with NSE(download_folder=Path.cwd(), server=False, timeout=20) as nse:
        for symbol in stocks2:
            data = get_stock_data(nse, symbol)
            if data:
                element_int = int(data['market_cap'] / 10_000_000) if data['market_cap'] else 0
                results.append({'stock': symbol, 'element_int': element_int, 'delivery': data['delivery_percentage']})
                print(f"{symbol}: Market Cap = Rs {data['market_cap_cr']:,.2f} Cr, Delivery % = {data['delivery_percentage']:.2f}%")

    # Sort and rank by market cap descending
    sorted_results = sorted(results, key=lambda x: x.get('element_int', 0) if x.get('element_int') else 0, reverse=True)
    for i, item in enumerate(sorted_results, int(os.environ.get("RANK_OFFSET", "1"))):
        item['rank'] = i

    date = datetime.now().date().isoformat()
    
    # Save to Excel
    save_to_excel(sorted_results, date)
    
    # Save to SQLite database
    save_to_database(sorted_results, date)


if __name__ == "__main__":
    main()