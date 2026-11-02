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

stocks1 = [
    "CAMS",
    "CONCORDBIO",
    "CEWATER",
    "CONFIPET",
    "CCCL",
    "CONCOR",
    "CONTROLPR",
    "CORALFINAC",
    "COROMANDEL",
    "CORONA",
    "COSMOFIRST",
    "CCHHL",
    "CPSEETF",
    "CRAFTSMAN",
    "CRAVATEX",
    "CREATIVEYE",
    "CNL",
    "CREDITACC",
    "MUFTI",
    "CREST",
    "CRISIL",
    "CRIZAC",
    "CROMPTON",
    "CROWN",
    "CSBBANK",
    "CSLFINANCE",
    "CSM",
    "CUMMINSIND",
    "CUPID",
    "CYBERTECH",
    "CYIENTDLM",
    "CYIENT",
    "DPWIRES",
    "DLINKINDIA",
    "DPABHUSHAN",
    "DBCORP",
    "DABUR",
    "DALBHARAT",
    "DALMIASUG",
    "DAMCAPITAL",
    "DAMODARIND",
    "DANGEE",
    "DATAPATTNS",
    "DATAMATICS",
    "DAVANGERE",
    "DBSTOCKBRO",
    "DCI",
    "DCBBANK",
    "DCMFINSERV",
    "DCM",
    "DCMSRIND",
    "DCMSIL",
    "DCMSHRIRAM",
    "DCW",
    "DCXINDIA",
    "DDEVPLSTIK",
    "DENORA",
    "DECCANCE",
    "DECNGOLD",
    "DEEDEV",
    "DEEPINDS",
    "DEEPA",
    "DBEIL",
    "DEEPAKFERT",
    "DEEPAKNTR",
    "DELHIVERY",
    "DELTACORP",
    "DELTAMAGNT",
    "DEN",
    "DENTA",
    "DEVX",
    "DEVIT",
    "DEVYANI",
    "DBOL",
    "DHAMPURSUG",
    "DHANBANK",
    "DHANUKA",
    "DHARMAJ",
    "DHATRE",
    "DHOOTTRANS",
    "DHRUV",
    "DHUNINV",
    "DVL",
    "DIAMINESQ",
    "DICIND",
    "DIFFNKG",
    "DGCONTENT",
    "DIGIDRIVE",
    "DIGISPICE",
    "DIGITIDE",
    "DNAMEDIA",
    "DBL",
    "DISAQ",
    "DISHTV",
    "DCAL",
    "DIVGIITTS",
    "DIVISLAB",
    "DIXON",
    "DJML",
    "DLF",
    "DMCC",
    "DODLA",
    "DOLATALGO",
    "DOLLAR",
    "DOLLEX",
    "DOLPHIN",
    "DOMS",
    "DONEAR",
    "AGARWALEYE",
    "DRAGARWQ",
    "LALPATHLAB",
    "DRREDDY",
    "DRCSYSTEMS",
    "DREAMFOLKS",
    "DREDGECORP",
    "KEEPLEARN",
    "LIQUIDETF",
    "BANK10ADD",
    "BANKADD",
    "GOLDADD",
    "ITADD",
    "EQUAL50ADD",
    "NIFTYADD",
    "PSUBANKADD",
    "PVTBANKADD",
    "MIDQ50ADD",
    "SENSEXADD",
    "SILVERADD",
    "FLEXIADD",
    "FMCGADD",
    "HEALTHADD",
    "INSURADD",
    "LIQUIDADD",
    "MIDBANKADD",
    "MIDCAPADD",
    "MSCIADD",
    "NEXT30ADD",
    "NEXT50ADD",
    "SMALLADD",
    "TOP10ADD",
    "DUCON",
    "DWARKESH",
    "DSSL",
    "DYNAMATECH",
    "DYCL",
    "DYNPRO",
    "EASTWEST",
    "EASEMYTRIP",
    "ECLERX",
    "ECORECO",
    "ECOSMOBLTY",
    "BBETF0432",
    "EBANK10",
    "EBANKNIFTY",
    "EBBETF0430",
    "EBBETF0431",
    "EBBETF0433",
    "ECAPINSURE",
    "EGOLD",
    "ELIQUID",
    "ELM250",
    "ELMDIV",
    "EMETAL",
    "EMULTIMQ",
    "ENEXT50",
    "ENIFTY",
    "ESENSEX",
    "ESILVER",
    "EDELWEISS",
    "EFCIL",
    "EICHERMOT",
    "EIDPARRY",
    "EIHAHOTELS",
    "EIHOTEL",
    "EIMCOELECO",
    "ELANTAS",
    "ELCIDIN",
    "ELDEHSG",
    "ELECON",
    "EMIL",
    "ELECTCAST",
    "ELECTHERM",
    "ELEVATE",
    "ELGIEQUIP",
    "ELGIRUBCO",
    "ELIN",
    "ELITECON",
    "ELLEN",
    "ELPROINTL",
    "EMAMILTD",
    "EMAMIPAP",
    "EMAMIREAL",
    "EMBDL",
    "EMCURE",
    "EMKAY",
    "EMMBI",
    "EMMVEE",
    "EMPIND",
    "EMPOWER",
    "EMSLIMITED",
    "EMUDHRA",
    "ENDURANCE",
    "ENERGYDEV",
    "ENGINERSIN",
    "ENKEIWHEL",
    "ENTERO",
    "ENIL",
    "EIEL",
    "EPACK",
    "EPACKPEB",
    "EPIGRAL",
    "EPL",
    "EQUITASBNK",
    "ERIS",
    "ESABINDIA",
    "ESCORTS",
    "ESDS",
    "ESSARSHPNG",
    "ESTER",
    "ETERNAL",
    "ETHOSLTD",
    "EUREKAFORB",
    "EIFFL",
    "EUROBOND",
    "EUROPRATIK",
    "EVEREADY",
    "EVERESTIND",
    "EKC",
    "EXCELINDUS",
    "EXCELSOFT",
    "EXICOM",
    "EXIDEIND",
    "EXPLEOSOL",
    "EXXARO",
    "FABTECH",
    "FACT",
    "FAIRCHEMOR",
    "FAZE3Q",
    "FCSSOFT",
    "FDC",
    "FEDFINA",
    "FEDDERSHOL",
    "FEDERALBNK",
    "FMGOETZE",
    "FERMENTA",
    "FIBERWEB",
    "FIEMIND",
    "FILATFASH",
    "FILATEX",
    "FINEORG",
    "FCL",
    "FINKURVE",
    "FINOPB",
    "FINCABLES",
    "FINPIPE",
    "FSL",
    "FISCHER",
    "FIVESTAR",
    "FLAIR",
    "FOCUS",
    "FOODSIN",
    "FORCEMOT",
    "FORTIS",
    "FOSECOIND",
    "FRACTAL",
    "FREDUN",
    "FRONTSP",
    "NYKAA",
    "UTLSOLAR",
    "FUSIONCX",
    "FUSION",
    "FMNL",
    "GMBREW",
    "GRINFRA",
    "GABRIEL",
    "GAIL",
    "GAJA",
    "GALAPREC",
    "GALAXYSURF",
    "GALLANTT",
    "GANDHAR",
    "GANDHITUBE",
    "GANESHCP",
    "GANESHHOU",
    "GANECOS",
    "GANGAFORGE",
    "GANGESSECU",
    "GRSE",
    "GARUDA",
    "GRWRHITECH",
    "GARFIBRES",
    "GATEWAY",
    "GAUDIUMIVF",
    "GAYAHWS",
    "GAYAPROJ",
    "GVPIL",
    "GVT&D",
    "GEECEE",
    "GEEKAYWIRE",
    "GEMAROMA",
    "GICRE",
    "GENCON",
    "GENESYS",
    "GENUSPAPER",
    "GENUSPOWER",
    "GEOJITFSL",
    "GERMAN",
    "GFLLIMITED",
    "GHCL",
    "GHCLTEXTIL",
    "GICHSGFIN",
    "GILLANDERS",
    "GILLETTE",
    "GINNIFILA",
    "GKENERGY",
    "GKB",
    "GKWLIMITED",
    "GLAND",
    "GLASSWALL",
    "GLAXO",
    "GLENMARK",
    "GLOBAL",
    "MEDANTA",
    "GLOBALVECT",
    "GLOBALE",
    "GLOBECIVIL",
    "GLOBE",
    "GLOBUSSPR",
    "GLOSTERLTD",
    "GLOTTIS",
    "GMMPFAUDLR",
    "GMRAIRPORT",
    "GMRP&UI",
    "GNA",
    "EBGNG",
    "GODIGIT",
    "GOCOLORS",
    "GOACARBON",
    "GOCLCORP",
    "GODAVARIB",
    "GPIL",
    "GODFRYPHLP",
    "GODREJAGRO",
    "GODREJCP",
    "GODREJIND",
    "GODREJPROP",
    "GOKEX",
    "GOKULAGRO",
    "GOKUL",
    "GOLDIAM",
    "GOODLUCK",
    "GOODYEAR",
    "GOPAL",
    "GOYALALUM",
    "GPTHEALTH",
    "GPTINFRA",
    "GRANDOAK",
    "GRANULES",
    "GRAPHITE",
    "GRASIM",
    "GRAUWEIL",
    "GRAVITA",
    "GREAVESCOT",
    "GREENLAM",
    "GREENPANEL",
    "GREENPLY",
    "GCSL",
    "GRINDWELL",
    "GRMOVER",
    "CEMNTGROWW",
    "GROWWCAPM",
    "GROWWCHEM",
    "GROWWDEFNC",
    "GROWWEV",
    "GROWWGOLD",
    "GROWWHOSPI",
    "GROWWLIQID",
    "GROWWLOVOL",
    "GROWWMC150",
    "GROWWMETAL",
    "GROWWMOM50",
    "GROWWN200",
    "GROWWNET",
    "GROWWNIFTY",
    "GROWWNXT50",
    "GROWWPOWER",
    "GROWWPSE",
    "GROWWPSUBK",
    "GROWWRAIL",
    "GROWWRLTY",
    "GROWWSC250",
    "GROWWSLVR",
    "PVTBKGROWW",
    "SMALLGROWW",
    "GRPLTD",
    "GSPCROP",
    "GSPLTRANS",
    "GSS",
    "GTLINFRA",
    "GTL",
    "GTPL",
    "GUFICBIO",
    "GIPCL",
    "GNFC",
    "GSFC",
    "GUJAPOLLO",
    "GUJALKALI",
    "GAEL",
    "GUJENERGY",
    "FLUOROCHEM",
    "GKSL",
    "GLFL",
    "GMDCLTD",
    "GNRL",
    "GPPL",
    "GUJTHEM",
    "GULFOILLUB",
    "GULPOLY",
    "GYFTR",
    "HBESD",
    "HGINFRA",
    "HALDER",
    "HALDYNGL",
    "HALEOSLABS",
    "HGM",
    "HAPPSTMNDS",
    "HAPPYFORGE",
    "HARDWYN",
    "HARIOMPIPE",
    "HRYNSHP",
    "HARRMALAYA",
    "HARSHA",
    "HATHWAY",
    "HATSUN",
    "HAVELLS",
    "HAWKINCOOK",
    "HAZOOR",
    "HBPOR",
    "HBSL",
    "HBLENGINE",
    "HCL-INSYS",
    "HCLTECH",
    "HDFIRE",
    "HDBFS",
    "HDFCAMC",
    "HDFCBANK",
    "HDFCGOLD",
    "HDFCLIFE",
    "HDFCBSE500",
    "HDFCGROWTH",
    "HDFCLIQUID",
    "HDFCLOWVOL",
    "HDFCMID150",
    "HDFCMOMENT",
    "HDFCNEXT50",
    "HDFCNIF100",
    "HDFCNIFBAN",
    "HDFCNIFIT",
    "HDFCNIFTY",
    "HDFCNIMEG",
    "HDFCPVTBAN",
    "HDFCQUAL",
    "HDFCSENSEX",
    "HDFCSILVER",
    "HDFCSML250",
    "HDFCVALUE",
    "HFTSEINDIA",
    "HDFCPSUBK",
    "HEADSUP",
    "HEALTHX",
    "HCG",
    "HECPROJECT",
    "HEGAM",
    "HEIDELBERG",
    "HEMIPROP",
    "HERANBA",
    "HERITGFOOD",
    "HEROMOTOCO",
    "HEROMOTORS",
    "HESTERBIO",
    "HEXATRADEX",
    "HEXAGON",
    "HEXT",
    "HITECH",
    "HILINFRA",
    "HIKAL",
    "HILTON",
    "HSCL",
    "HIMATSEIDE",
    "HINDALUMI",
    "HINDALCO",
    "HPIL",
    "HGS",
    "HINDADH",
    "HAL",
    "HINDCOMPOS",
    "HCC",
    "HINDCOPPER",
    "HNDFDS",
    "HINDOILEXP",
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
        for row, stock in enumerate(stocks1, 2):
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
    print(f"Data saved to Stcapmark.xlsx ({len(sorted_results)} stocks1)")


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
        for symbol in stocks1:
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