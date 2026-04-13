#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TWSE/TPEX Auction Data Collector
---------------------------------
Usage:
  python auction_fetcher.py           # Full historical download (2016~now)
  python auction_fetcher.py --update  # Incremental update
  python auction_fetcher.py --html    # Regenerate HTML viewer only
"""

import os
import sys
import time
import json
import logging
from datetime import datetime, date, timedelta

import requests
import pandas as pd
from bs4 import BeautifulSoup

# ── Settings ──────────────────────────────────────────────────────────────────
BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
STOCKS_JSON   = os.path.join(BASE_DIR, 'auction_stocks.json')
CBS_JSON      = os.path.join(BASE_DIR, 'auction_cbs.json')
HTML_FILE     = os.path.join(BASE_DIR, 'auction_viewer.html')
LEGACY_EXCEL  = os.path.join(BASE_DIR, 'auction_data.xlsx')
START_YEAR_AD = 2016

TWSE_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Referer': 'https://www.twse.com.tw/',
}
TPEX_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Referer': 'https://www.tpex.org.tw/',
}
THEFEW_HEADERS = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}

# ── Logging ───────────────────────────────────────────────────────────────────
log_file = os.path.join(BASE_DIR, 'auction_fetcher.log')
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s  %(levelname)-7s  %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)

# ── Date helpers ──────────────────────────────────────────────────────────────

def current_ad_year() -> int:
    return datetime.now().year

def ad_to_roc(year_ad: int) -> int:
    return year_ad - 1911

def _parse_date(date_str: str) -> datetime:
    parts = str(date_str).strip().split('/')
    return datetime(int(parts[0]), int(parts[1]), int(parts[2]))

def parse_tw_date(date_str: str):
    if not date_str or str(date_str).strip() in ('', 'nan', 'None'):
        return None
    try:
        return _parse_date(str(date_str))
    except Exception:
        return None

# ── Column normalization ───────────────────────────────────────────────────────

_COL_ALIASES: dict[str, str] = {
    '投標截止日':              '投標結束日',
    '競拍股數':                '競拍數量(張)',
    '最低每標單位(張)':         '最低每標單投標數量(張)',
    '最高投(得)標單位(張)':     '最高投(得)標數量(張)',
    '保證金成數':              '保證金成數(%)',
    '合格投標量(張)':           '合格投標數量(張)',
    '取消競價拍賣（流標或取消）':  '取消競價拍賣(流標或取消)',
    '承銷價格(元)':             '實際承銷價格(元)',   # 2016-2024 uses old name
}

def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.rename(columns=_COL_ALIASES)
    if '發行性質' not in df.columns:
        pos = list(df.columns).index('發行市場') + 1 if '發行市場' in df.columns else 4
        df.insert(pos, '發行性質', '')
    # If both old and new name survived a concat, merge them
    if '承銷價格(元)' in df.columns and '實際承銷價格(元)' in df.columns:
        df['實際承銷價格(元)'] = df['實際承銷價格(元)'].fillna(df['承銷價格(元)'])
        df.drop(columns=['承銷價格(元)'], inplace=True)
    return df

def _normalize_market(market: str, nature: str) -> str:
    m = str(market).strip()
    if '集中' in m or (m == '初上市'):
        return '集中交易市場'
    if '櫃' in m or (m == '初上櫃'):
        return '櫃檯買賣'
    return m

# ── Auction API ────────────────────────────────────────────────────────────────

def fetch_auction_year(year_ad: int) -> pd.DataFrame:
    url    = 'https://www.twse.com.tw/rwd/zh/announcement/auction'
    params = {'response': 'json', 'date': str(year_ad)}
    try:
        r = requests.get(url, params=params, headers=TWSE_HEADERS, timeout=30)
        r.raise_for_status()
        data = r.json()
        if data.get('stat') != 'OK' or not data.get('data'):
            log.warning(f'  {year_ad}: 無資料')
            return pd.DataFrame()
        df = pd.DataFrame(data['data'], columns=data['fields'])
        df = _normalize_columns(df)
        df = df[df['證券代號'].astype(str).str.strip() != '']
        df = df.dropna(subset=['投標結束日'])
        df['發行市場'] = df.apply(
            lambda r: _normalize_market(r['發行市場'], r.get('發行性質', '')), axis=1
        )
        log.info(f'  {year_ad}: {len(df)} 筆')
        return df
    except Exception as exc:
        log.error(f'  {year_ad} 下載失敗: {exc}')
        return pd.DataFrame()

# ── Result-backfill helpers ────────────────────────────────────────────────────

# These fields are 0 before the auction closes, non-zero after results are in
_RESULT_FIELDS = [
    '得標總金額(元)', '總合格件', '合格投標數量(張)',
    '最低得標價格(元)', '最高得標價格(元)', '得標加權平均價格(元)',
    '實際承銷價格(元)', '取消競價拍賣(流標或取消)',
]

def _is_empty(val) -> bool:
    """True if val is None, NaN, or a blank/nan/None string."""
    if val is None:
        return True
    try:
        import math
        if math.isnan(float(val)) if isinstance(val, float) else False:
            return True
    except Exception:
        pass
    return str(val).strip().lower() in ('', 'nan', 'none')

def _needs_result_update(row: pd.Series) -> bool:
    """True if auction has closed but we still have no result data."""
    try:
        bid_end = parse_tw_date(str(row.get('投標結束日', '')))
        if bid_end is None or bid_end.date() >= date.today():
            return False   # not closed yet
        # already has a cancel marker → no need to update
        cancel = row.get('取消競價拍賣(流標或取消)', '')
        if not _is_empty(cancel):
            return False
        qty = str(row.get('合格投標數量(張)', '0')).strip().replace(',', '')
        return qty in ('0', '', 'nan', 'None')
    except Exception:
        return False

def merge_results(exist_df: pd.DataFrame, fresh_df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """For records that are missing results, fill from freshly-fetched data."""
    if exist_df.empty or fresh_df.empty:
        return exist_df, 0
    fresh_index = {_record_key(r): r for _, r in fresh_df.iterrows()}
    count = 0
    for idx, row in exist_df.iterrows():
        if not _needs_result_update(row):
            continue
        key = _record_key(row)
        fr  = fresh_index.get(key)
        if fr is None:
            continue
        # Only update if the fresh data actually has results now
        fresh_qty    = str(fr.get('合格投標數量(張)', '0')).strip().replace(',', '')
        fresh_cancel = str(fr.get('取消競價拍賣(流標或取消)', '')).strip()
        if fresh_qty not in ('0', '', 'nan', 'None') or fresh_cancel:
            for field in _RESULT_FIELDS:
                if field in fresh_df.columns:
                    exist_df.at[idx, field] = fr.get(field, '')
            count += 1
    return exist_df, count

# ── Price cache (local JSON) ──────────────────────────────────────────────────
# Structure: { "YYYY-MM-DD": { "code": price_float } }
PRICE_CACHE_FILE = os.path.join(BASE_DIR, 'price_cache.json')
_price_cache: dict[str, dict[str, float]] = {}

def _load_price_cache() -> None:
    global _price_cache
    if os.path.exists(PRICE_CACHE_FILE):
        try:
            with open(PRICE_CACHE_FILE, encoding='utf-8') as f:
                _price_cache = json.load(f)
        except Exception:
            _price_cache = {}

def _save_price_cache() -> None:
    with open(PRICE_CACHE_FILE, 'w', encoding='utf-8') as f:
        json.dump(_price_cache, f, ensure_ascii=False)

def _cache_lookup(code: str, date_key: str):
    day = _price_cache.get(date_key, {})
    p   = day.get(code)
    return p  # float or None

def _cache_store(date_key: str, prices: dict[str, float]) -> None:
    if date_key not in _price_cache:
        _price_cache[date_key] = {}
    _price_cache[date_key].update(prices)

def _roc_to_ad_date(roc_str: str) -> str:
    """Convert '1150409' (ROC yyyymmdd) → 'YYYY-MM-DD'."""
    y = int(roc_str[:3]) + 1911
    m = roc_str[3:5]
    d = roc_str[5:7]
    return f'{y}-{m}-{d}'

# ── OpenAPI bulk download (today only) ───────────────────────────────────────

def download_today_prices() -> str:
    """Download all TWSE+TPEX closing prices for today via OpenAPI.
    Stores in local cache.  Returns date string 'YYYY-MM-DD' of the data.
    """
    date_key = None

    # TWSE
    try:
        r = requests.get('https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL',
                         headers={'User-Agent': 'Mozilla/5.0'}, timeout=30)
        r.raise_for_status()
        rows = r.json()
        if rows:
            date_key = _roc_to_ad_date(rows[0]['Date'])
            prices = {}
            for row in rows:
                c = row.get('Code', '').strip()
                p = row.get('ClosingPrice', '')
                if c and p and p not in ('', '--'):
                    try:
                        prices[c] = float(p.replace(',', ''))
                    except ValueError:
                        pass
            _cache_store(date_key, prices)
            log.info(f'  TWSE OpenAPI: {len(prices)} 筆收盤價 ({date_key})')
    except Exception as exc:
        log.warning(f'  TWSE OpenAPI 下載失敗: {exc}')

    time.sleep(0.5)

    # TPEX
    try:
        r = requests.get('https://www.tpex.org.tw/openapi/v1/tpex_mainboard_daily_close_quotes',
                         headers={'User-Agent': 'Mozilla/5.0'}, timeout=30)
        r.raise_for_status()
        rows = r.json()
        if rows:
            tpex_date_key = _roc_to_ad_date(rows[0]['Date'])
            if date_key is None:
                date_key = tpex_date_key
            prices = {}
            for row in rows:
                c = row.get('SecuritiesCompanyCode', '').strip()
                p = row.get('Close', '')
                if c and p and p not in ('', '--'):
                    try:
                        prices[c] = float(p.replace(',', ''))
                    except ValueError:
                        pass
            _cache_store(tpex_date_key, prices)
            log.info(f'  TPEX OpenAPI: {len(prices)} 筆收盤價 ({tpex_date_key})')
    except Exception as exc:
        log.warning(f'  TPEX OpenAPI 下載失敗: {exc}')

    if date_key:
        _save_price_cache()
    return date_key or ''

# ── Legacy per-stock fallback (for historical dates not in cache) ─────────────

def _get_twse_close_api(code: str, d: datetime):
    """One-stock TWSE STOCK_DAY API — single attempt, fast fail on 307."""
    url = 'https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY'
    target = f"{ad_to_roc(d.year)}/{d.month:02d}/{d.day:02d}"
    try:
        r = requests.get(url, params={'stockNo': code,
            'date': d.strftime('%Y%m%d'), 'response': 'json'},
            headers=TWSE_HEADERS, timeout=15)
        if r.status_code == 307:
            return None  # Rate limited — fall through to TPEX bulk
        data = r.json()
        if data.get('stat') != 'OK':
            return None
        for row in data.get('data', []):
            if row[0] == target:
                return float(row[6].replace(',', ''))
        return None
    except Exception as exc:
        log.warning(f'    TWSE {code} {d.date()}: {exc}')
        return None

def _get_tpex_close_api(code: str, d: datetime):
    """One-day TPEX bulk download for a specific date."""
    roc_date = f"{ad_to_roc(d.year)}/{d.month:02d}/{d.day:02d}"
    url = ('https://www.tpex.org.tw/web/stock/aftertrading/'
           'otc_quotes_no1430/stk_wn1430_result.php')
    params = {'l': 'zh-tw', 'o': 'json', 'd': roc_date, 'se': 'EW', 'st': 'stkno'}
    try:
        r = requests.get(url, params=params, headers=TPEX_HEADERS, timeout=30)
        r.raise_for_status()
        day_prices = {}
        tables = r.json().get('tables', [])
        if tables:
            for row in tables[0].get('data', []):
                c = str(row[0]).strip()
                p = str(row[2]).strip()
                if p and p not in ('', '--', 'N/A'):
                    try:
                        day_prices[c] = float(p.replace(',', ''))
                    except ValueError:
                        pass
        # Cache the whole day
        date_key = d.strftime('%Y-%m-%d')
        _cache_store(date_key, day_prices)
        return day_prices.get(code)
    except Exception as exc:
        log.warning(f'    TPEX {code} {d.date()}: {exc}')
        return None

# ── TPEX ESB (興櫃) current price ────────────────────────────────────────────

_ESB_CACHE: dict | None = None   # cached snapshot of tpex_esb_latest_statistics

def _get_esb_price(code: str) -> float | None:
    """Return today's average price from TPEX ESB (興櫃) for `code`.

    The ESB API has no historical endpoint, so this is only useful for
    stocks still on the emerging market whose auction date is today/recent.
    Uses `Average` (均價) as the proxy for closing price.
    """
    global _ESB_CACHE
    if _ESB_CACHE is None:
        try:
            r = requests.get(
                'https://www.tpex.org.tw/openapi/v1/tpex_esb_latest_statistics',
                headers=TPEX_HEADERS, timeout=20)
            r.raise_for_status()
            _ESB_CACHE = {d['SecuritiesCompanyCode']: d for d in r.json()}
            log.info(f'  ESB snapshot: {len(_ESB_CACHE)} 興櫃股票')
        except Exception as exc:
            log.warning(f'  ESB snapshot 下載失敗: {exc}')
            _ESB_CACHE = {}
    row = _ESB_CACHE.get(code)
    if not row:
        return None
    for field in ('Average', 'LatestPrice', 'PreviousAveragePrice'):
        v = str(row.get(field, '')).strip()
        if v and v not in ('', '--', '---', '0', '0.00'):
            try:
                return float(v)
            except ValueError:
                pass
    return None

# ── TWSE bulk historical download (one call = all ~1300 stocks for a date) ────

_TWSE_BULK_DATES: set[str] = set()  # dates where we've already done bulk download

def _get_twse_bulk_day(code: str, d: datetime):
    """Bulk-download ALL TWSE stocks for a historical date via STOCK_DAY_ALL.
    Caches the whole day (like TPEX bulk).  Returns price for `code` or None.
    """
    date_key = d.strftime('%Y-%m-%d')
    if date_key in _TWSE_BULK_DATES:
        return _cache_lookup(code, date_key)  # Already attempted; stock not on TWSE

    url = 'https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY_ALL'
    try:
        r = requests.get(url,
                         params={'date': d.strftime('%Y%m%d'), 'response': 'json'},
                         headers=TWSE_HEADERS, timeout=20)
        if r.status_code == 307:
            return None  # Rate limited — skip silently
        data = r.json()
        if data.get('stat') != 'OK':
            _TWSE_BULK_DATES.add(date_key)
            return None
        day_prices = {}
        for row in data.get('data', []):
            c = str(row[0]).strip()
            try:
                day_prices[c] = float(str(row[6]).replace(',', ''))
            except (ValueError, IndexError):
                pass
        _cache_store(date_key, day_prices)
        _TWSE_BULK_DATES.add(date_key)
        if day_prices:
            log.info(f'    TWSE bulk {date_key}: {len(day_prices)} 筆')
        return day_prices.get(code)
    except Exception as exc:
        log.warning(f'    TWSE bulk {d.date()}: {exc}')
        return None

# ── yfinance price lookup (reliable historical data, no rate limiting) ─────────

try:
    import yfinance as yf
    _YF_AVAILABLE = True
except ImportError:
    _YF_AVAILABLE = False

def _get_yfinance_price(code: str, d: datetime, market: str = '') -> float | None:
    """Fetch closing price via yfinance.  Uses raw (non-adjusted) close.

    If the exact date has no data (stock not yet listed — e.g. auction date is
    before the IPO listing date), extends the window up to 10 calendar days
    forward to find the first available post-listing close price.
    """
    if not _YF_AVAILABLE:
        return None
    import pandas as pd
    start_str = d.strftime('%Y-%m-%d')
    # Extend window: exact date + up to 20 calendar days to capture pre-listing auctions
    # (typical gap between 興櫃 auction end date and TWSE/TPEX listing date is 11~14 days)
    end_str   = (d + timedelta(days=21)).strftime('%Y-%m-%d')

    # Determine which exchange suffix to try first
    if '集中' in market:
        suffixes = ['.TW', '.TWO']
    elif '櫃' in market:
        suffixes = ['.TWO', '.TW']
    else:
        suffixes = ['.TW', '.TWO']  # CB underlying: try both

    for suffix in suffixes:
        ticker = f'{code}{suffix}'
        try:
            df = yf.download(ticker, start=start_str, end=end_str,
                             progress=False, auto_adjust=False)
            if df.empty:
                continue
            # Find price on target date; if not available, use the next trading day
            # (handles case where auction date is before IPO listing date)
            target = pd.Timestamp(d.date())
            if isinstance(df.columns, pd.MultiIndex):
                close_series = df[('Close', ticker)]
            else:
                close_series = df['Close']
            close_series = close_series.dropna()
            if close_series.empty:
                continue
            if target in close_series.index:
                val = float(close_series.loc[target])
            else:
                # Use first available date on or after target (within window)
                future = close_series[close_series.index >= target]
                if future.empty:
                    continue
                val = float(future.iloc[0])
                days_gap = (future.index[0] - target).days
                log.debug(f'    yfinance {ticker}: no data on {start_str}, '
                          f'using {future.index[0].date()} (+{days_gap}d)')
            if val and val > 0:
                return val
        except Exception as exc:
            log.debug(f'    yfinance {ticker} {start_str}: {exc}')
    return None

# ── Main price lookup ─────────────────────────────────────────────────────────

def get_closing_price(code: str, date_str: str, market: str = ''):
    """Get closing price for a stock on a given date.

    Strategy:
    1. Check local price cache (built from OpenAPI daily downloads)
    2. yfinance — reliable historical data for both TWSE and OTC
    3. TPEX bulk day download (fallback, caches whole day)
    """
    code = str(code).strip()
    if not code or str(date_str).strip() in ('', 'nan', 'None'):
        return None

    try:
        d        = _parse_date(date_str)
        date_key = d.strftime('%Y-%m-%d')
    except Exception:
        return None

    # 1. Cache hit — return immediately without any sleep
    cached = _cache_lookup(code, date_key)
    if cached is not None:
        return cached

    time.sleep(0.2)

    # 2. yfinance (handles both TWSE and OTC, historical dates, no rate limits)
    price = _get_yfinance_price(code, d, market)
    if price is not None:
        _cache_store(date_key, {code: price})
        _save_price_cache()
        return price

    # 3. TPEX bulk day download (caches ALL OTC stocks — useful when yfinance misses)
    price = _get_tpex_close_api(code, d)
    if price is not None:
        _save_price_cache()
        return price

    # 4. TPEX ESB (興櫃) — for stocks still on emerging market today.
    #    Only useful when auction date ≈ today; returns current average price.
    price = _get_esb_price(code)
    if price is not None:
        _cache_store(date_key, {code: price})
        _save_price_cache()
    return price

# ── CB helpers ────────────────────────────────────────────────────────────────

def is_cb(code: str) -> bool:
    return len(str(code).strip()) > 4

def cb_underlying_code(cb_code: str) -> str:
    return str(cb_code).strip()[:4]

def get_cb_conversion_price(cb_code: str):
    try:
        r    = requests.get(f'https://thefew.tw/quote/{cb_code}',
                            headers=THEFEW_HEADERS, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.content, 'lxml')
        tds  = soup.find_all('td')
        for i, td in enumerate(tds):
            if '發行時轉換價' in td.get_text() and i + 1 < len(tds):
                return float(tds[i + 1].get_text(strip=True).replace(',', ''))
        return None
    except Exception as exc:
        log.warning(f'    CB {cb_code} thefew.tw: {exc}')
        return None

# ── Enrich ────────────────────────────────────────────────────────────────────

def _clean_code(raw) -> str:
    s = str(raw).strip()
    if s.endswith('.0') and s[:-2].isdigit():
        return s[:-2]
    return s

def _p(val) -> str | None:
    """Convert a price float/None to str for DataFrame storage."""
    return str(val) if val is not None else None

def enrich_stocks(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df['投標結束日收盤價'] = pd.array([None] * len(df), dtype=object)
    total = len(df)
    for i, (idx, row) in enumerate(df.iterrows()):
        df.at[idx, '投標結束日收盤價'] = _p(get_closing_price(
            row['證券代號'], row['投標結束日'], row['發行市場']))
        if (i + 1) % 20 == 0 or (i + 1) == total:
            log.info(f'    股票收盤價: {i+1}/{total}')
    return df

def enrich_cbs(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df['投標結束日收盤價'] = pd.array([None] * len(df), dtype=object)
    df['發行時轉換價']    = pd.array([None] * len(df), dtype=object)
    total = len(df)
    for i, (idx, row) in enumerate(df.iterrows()):
        df.at[idx, '投標結束日收盤價'] = _p(get_closing_price(
            cb_underlying_code(row['證券代號']), row['投標結束日'], row['發行市場']))
        df.at[idx, '發行時轉換價'] = _p(get_cb_conversion_price(row['證券代號']))
        time.sleep(0.3)
        if (i + 1) % 10 == 0 or (i + 1) == total:
            log.info(f'    CB: {i+1}/{total}')
    return df

def split_and_enrich(combined: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    combined['證券代號'] = combined['證券代號'].apply(_clean_code)
    mask_cb = combined['證券代號'].apply(is_cb)
    stocks  = combined[~mask_cb].copy().reset_index(drop=True)
    cbs     = combined[mask_cb].copy().reset_index(drop=True)
    log.info(f'股票: {len(stocks)} 筆，CB: {len(cbs)} 筆')
    log.info('抓取股票收盤價…')
    stocks = enrich_stocks(stocks)
    log.info('抓取CB個股收盤價及發行時轉換價…')
    cbs = enrich_cbs(cbs)
    return stocks, cbs

# ── Storage ───────────────────────────────────────────────────────────────────

def save_data(stocks: pd.DataFrame, cbs: pd.DataFrame) -> None:
    stocks.to_json(STOCKS_JSON, orient='records', force_ascii=False, indent=2)
    cbs.to_json(CBS_JSON,    orient='records', force_ascii=False, indent=2)
    generate_html(stocks, cbs)
    log.info(f'已儲存 (股票:{len(stocks)}, CB:{len(cbs)})')

def load_data() -> tuple[pd.DataFrame, pd.DataFrame]:
    # Prefer JSON; fall back to legacy Excel for first-time migration
    if os.path.exists(STOCKS_JSON):
        stocks = pd.read_json(STOCKS_JSON, dtype=str)
        cbs    = pd.read_json(CBS_JSON, dtype=str) if os.path.exists(CBS_JSON) else pd.DataFrame()
        return stocks, cbs
    if os.path.exists(LEGACY_EXCEL):
        log.info('從舊版 Excel 遷移資料…')
        stocks = pd.read_excel(LEGACY_EXCEL, sheet_name='股票競拍', dtype=str)
        cbs    = pd.read_excel(LEGACY_EXCEL, sheet_name='CB競拍',   dtype=str)
        stocks = _normalize_columns(stocks)
        cbs    = _normalize_columns(cbs)
        # Drop any duplicate columns that survive normalization
        stocks = stocks.loc[:, ~stocks.columns.duplicated()]
        cbs    = cbs.loc[:, ~cbs.columns.duplicated()]
        return stocks, cbs
    return pd.DataFrame(), pd.DataFrame()

# ── Unique key ────────────────────────────────────────────────────────────────

def _record_key(row: pd.Series) -> tuple:
    return (str(row.get('投標開始日', '')).strip(),
            str(row.get('證券代號',   '')).strip())

# ── HTML generation ───────────────────────────────────────────────────────────

def generate_html(stocks: pd.DataFrame, cbs: pd.DataFrame) -> None:
    stocks_json = stocks.to_json(orient='records', force_ascii=False)
    cbs_json    = cbs.to_json(orient='records', force_ascii=False)
    now = datetime.now().strftime('%Y-%m-%d %H:%M')

    html = _HTML_TEMPLATE \
        .replace('__TIMESTAMP__', now) \
        .replace('__N_STOCKS__', str(len(stocks))) \
        .replace('__N_CBS__', str(len(cbs))) \
        .replace('__STOCKS_JSON__', stocks_json) \
        .replace('__CBS_JSON__', cbs_json)

    with open(HTML_FILE, 'w', encoding='utf-8') as f:
        f.write(html)
    log.info(f'HTML 產生 → {HTML_FILE}')

# ── Mode 1: Full download ─────────────────────────────────────────────────────

def download_history() -> None:
    log.info('=== 全量下載 (2016~現在) ===')
    _load_price_cache()
    log.info('下載今日全市場收盤價…')
    download_today_prices()
    all_dfs = []
    for year in range(START_YEAR_AD, current_ad_year() + 1):
        log.info(f'下載 {year} (民國{ad_to_roc(year)})…')
        df = fetch_auction_year(year)
        if not df.empty:
            all_dfs.append(df)
        time.sleep(1.2)
    if not all_dfs:
        log.error('無資料')
        return
    combined    = pd.concat(all_dfs, ignore_index=True)
    stocks, cbs = split_and_enrich(combined)
    save_data(stocks, cbs)
    log.info('=== 全量下載完成 ===')

# ── Mode 2: Incremental update ────────────────────────────────────────────────

def update_data() -> None:
    log.info('=== 增量更新 ===')
    _load_price_cache()
    log.info('下載今日全市場收盤價 (OpenAPI)…')
    today_date = download_today_prices()
    if today_date:
        log.info(f'  今日行情已存入快取 ({today_date})')
    stocks_exist, cbs_exist = load_data()

    current_year = current_ad_year()
    new_dfs = []
    for year in [current_year - 1, current_year]:
        log.info(f'下載 {year} (民國{ad_to_roc(year)})…')
        df = fetch_auction_year(year)
        if not df.empty:
            new_dfs.append(df)
        time.sleep(1.2)

    if not new_dfs:
        log.info('無法取得最新競拍清單（API 暫時無回應），僅補抓收盤價')
        fresh_stk = fresh_cb = pd.DataFrame()
        new_stocks = new_cbs = pd.DataFrame()
    else:
        fetched = pd.concat(new_dfs, ignore_index=True)
        fetched['證券代號'] = fetched['證券代號'].apply(_clean_code)

        mask_cb    = fetched['證券代號'].apply(is_cb)
        fresh_stk  = fetched[~mask_cb].copy()
        fresh_cb   = fetched[mask_cb].copy()

        # ── 1. Backfill auction results for existing records that now have data ──
        result_stk = result_cb = 0
        if not stocks_exist.empty:
            stocks_exist, result_stk = merge_results(stocks_exist, fresh_stk)
        if not cbs_exist.empty:
            cbs_exist, result_cb = merge_results(cbs_exist, fresh_cb)
        if result_stk or result_cb:
            log.info(f'  補填得標結果：股票 {result_stk} 筆，CB {result_cb} 筆')

        # ── 2. Find truly new records ──────────────────────────────────────────
        def find_new(new_df, exist_df):
            if exist_df.empty:
                return new_df
            exist_keys = {_record_key(r) for _, r in exist_df.iterrows()}
            return new_df[new_df.apply(lambda r: _record_key(r) not in exist_keys, axis=1)].copy()

        new_stocks = find_new(fresh_stk, stocks_exist)
        new_cbs    = find_new(fresh_cb,  cbs_exist)
        log.info(f'  新增：股票 {len(new_stocks)} 筆，CB {len(new_cbs)} 筆')

        if not new_stocks.empty:
            new_stocks = enrich_stocks(new_stocks.reset_index(drop=True))
        if not new_cbs.empty:
            new_cbs = enrich_cbs(new_cbs.reset_index(drop=True))

    # ── 3. Backfill missing closing prices ────────────────────────────────
    SAVE_INTERVAL = 30  # save progress every N items

    def backfill_close(df, is_cb_sheet, other_df, other_is_cb):
        """Backfill prices with incremental saves every SAVE_INTERVAL items."""
        if df.empty:
            return df
        if '投標結束日收盤價' not in df.columns:
            df['投標結束日收盤價'] = None
        mask = df['投標結束日收盤價'].isna() | df['投標結束日收盤價'].isin(['nan', 'None', ''])
        missing = df[mask]
        if missing.empty:
            return df
        total = len(missing)
        log.info(f'  補抓收盤價 {total} 筆…')
        for n, (idx, row) in enumerate(missing.iterrows(), 1):
            code  = cb_underlying_code(row['證券代號']) if is_cb_sheet else row['證券代號']
            price = get_closing_price(code, row['投標結束日'], row['發行市場'])
            df.at[idx, '投標結束日收盤價'] = str(price) if price is not None else None
            if is_cb_sheet and str(df.at[idx, '發行時轉換價']).strip() in ('', 'nan', 'None'):
                cp = get_cb_conversion_price(row['證券代號'])
                df.at[idx, '發行時轉換價'] = str(cp) if cp is not None else None
            if n % SAVE_INTERVAL == 0 or n == total:
                log.info(f'    進度 {n}/{total}，中途存檔…')
                if is_cb_sheet:
                    tmp_stocks = pd.concat([other_df, new_stocks], ignore_index=True)
                    tmp_cbs    = pd.concat([df, new_cbs],   ignore_index=True)
                else:
                    tmp_stocks = pd.concat([df, new_stocks], ignore_index=True)
                    tmp_cbs    = pd.concat([other_df, new_cbs], ignore_index=True)
                save_data(tmp_stocks, tmp_cbs)
        return df

    stocks_exist = backfill_close(stocks_exist, False, cbs_exist, True)
    cbs_exist    = backfill_close(cbs_exist,    True,  stocks_exist, False)

    stocks_all = pd.concat([stocks_exist, new_stocks], ignore_index=True)
    cbs_all    = pd.concat([cbs_exist,    new_cbs],    ignore_index=True)
    save_data(stocks_all, cbs_all)
    log.info('=== 增量更新完成 ===')

# ── HTML Template ─────────────────────────────────────────────────────────────

_HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>TWSE / TPEX 競拍資料庫</title>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Segoe UI', '微軟正黑體', Arial, sans-serif; font-size: 13px; background: #eef1f5; color: #222; }

/* ── Header ─────────────────── */
.hdr { background: #1a365d; color: #fff; padding: 14px 20px; display: flex; justify-content: space-between; align-items: center; }
.hdr h1 { font-size: 18px; font-weight: 600; letter-spacing: .3px; }
.hdr .meta { font-size: 11px; opacity: .75; }

/* ── Controls ───────────────── */
.ctrl { background: #fff; padding: 10px 20px; border-bottom: 1px solid #d0d5de;
        display: flex; gap: 12px; align-items: center; flex-wrap: wrap; position: sticky; top: 0; z-index: 100; }
.tabs { display: flex; gap: 4px; }
.tab  { padding: 6px 16px; border: 1px solid #bbb; border-radius: 4px; cursor: pointer;
        background: #fff; font: inherit; font-size: 13px; transition: .15s; }
.tab.active { background: #1a365d; color: #fff; border-color: #1a365d; }
.tab:hover:not(.active) { background: #edf0f7; }
#search { padding: 6px 11px; border: 1px solid #bbb; border-radius: 4px;
          font: inherit; font-size: 13px; width: 200px; outline: none; }
#search:focus { border-color: #1a365d; box-shadow: 0 0 0 2px #1a365d22; }
.stat-badge { margin-left: auto; background: #edf0f7; border-radius: 12px;
              padding: 3px 10px; font-size: 11px; color: #555; white-space: nowrap; }

/* ── Legend ─────────────────── */
.legend { display: flex; gap: 12px; align-items: center; font-size: 11px; color: #555; }
.dot { width: 10px; height: 10px; border-radius: 2px; display: inline-block; margin-right: 3px; }
.dot-done     { background: #d1fae5; border: 1px solid #6ee7b7; }
.dot-pending  { background: #fef9c3; border: 1px solid #fde047; }
.dot-noresult { background: #ffedd5; border: 1px solid #fdba74; }
.dot-cancel   { background: #f3f4f6; border: 1px solid #d1d5db; }

/* ── Table ──────────────────── */
.tbl-wrap { overflow: auto; height: calc(100vh - 95px); }
table { width: 100%; border-collapse: collapse; background: #fff; }
thead { position: sticky; top: 0; z-index: 10; }
thead tr:first-child { background: #243b5e; color: #fff; }
th { padding: 8px 10px; text-align: left; white-space: nowrap; cursor: pointer;
     user-select: none; font-weight: 500; font-size: 12px; border-right: 1px solid #2d4c72; }
th:hover { background: #2d5a8e; }
th:last-child { border-right: none; }
th .si { margin-left: 5px; opacity: .5; font-size: 10px; }
th.asc  .si::after { content: '▲'; opacity: 1; }
th.desc .si::after { content: '▼'; opacity: 1; }
th:not(.asc):not(.desc) .si::after { content: '⇅'; }
td { padding: 6px 10px; border-bottom: 1px solid #eaedf1; white-space: nowrap; vertical-align: middle; }
tr:hover td { background: #e8f0ff !important; }

/* row status */
tr.s-done     td { background: #f0fdf4; }
tr.s-pending  td { background: #fefce8; }
tr.s-noresult td { background: #fff7ed; }
tr.s-cancel   td { background: #f9fafb; color: #9ca3af; }

td.r  { text-align: right; font-variant-numeric: tabular-nums; }
td.mono { font-family: 'Consolas', monospace; font-size: 12px; }
td.bold { font-weight: 600; }

/* badges */
.badge { display: inline-block; padding: 1px 6px; border-radius: 3px; font-size: 11px; font-weight: 500; }
.b-tse   { background: #dbeafe; color: #1e40af; }
.b-otc   { background: #d1fae5; color: #065f46; }
.b-cncl  { background: #fee2e2; color: #991b1b; }
.b-pend  { background: #fef9c3; color: #92400e; }

/* highlight: 收盤 vs 得標均價 */
.hi-up  { color: #dc2626; font-weight: 600; }   /* 收盤 > 得標均價 (positive for bidders) */
.hi-dn  { color: #2563eb; font-weight: 600; }   /* 收盤 < 得標均價 */

.empty-msg { text-align: center; padding: 60px 20px; color: #9ca3af; font-size: 15px; }

/* ── Refresh button ──────────── */
#btn-refresh { padding: 5px 13px; border: 1px solid #bbb; border-radius: 4px; cursor: pointer;
               background: #fff; font: inherit; font-size: 13px; transition: .15s;
               display: inline-flex; align-items: center; gap: 5px; }
#btn-refresh:hover:not(:disabled) { background: #e0f0ff; border-color: #1a365d; color: #1a365d; }
#btn-refresh:disabled { opacity: .6; cursor: not-allowed; }
.spin { display: inline-block; }
#btn-refresh.spinning .spin { animation: spin .8s linear infinite; }
@keyframes spin { to { transform: rotate(360deg); } }
.toast { padding: 3px 11px; border-radius: 12px; font-size: 11px; font-weight: 500;
         white-space: nowrap; display: none; }
.toast.info    { display: inline-block; background: #dbeafe; color: #1e40af; }
.toast.success { display: inline-block; background: #d1fae5; color: #065f46; }
.toast.warn    { display: inline-block; background: #fef9c3; color: #92400e; }
.toast.error   { display: inline-block; background: #fee2e2; color: #991b1b; }

/* ── Column filter row ───────── */
.filter-row th { background: #1d3050; padding: 3px 4px;
                 border-right: 1px solid #2d4c72; border-bottom: 2px solid #0f1e38; }
.filter-row th:last-child { border-right: none; }
.fi { width: 100%; padding: 2px 5px; font-size: 11px; border: 1px solid #3a5070;
      background: #243b5e; color: #c8d8f0; border-radius: 3px; font-family: inherit;
      box-sizing: border-box; }
.fi::placeholder { color: #5a7898; font-size: 10px; }
.fi:focus { outline: none; border-color: #90b4e8; }
.fi.active { border-color: #fbbf24; background: #2e4a6e; color: #fde68a; }

/* ── Column panel ───────────── */
#btn-cols { padding: 5px 13px; border: 1px solid #bbb; border-radius: 4px; cursor: pointer;
            background: #fff; font: inherit; font-size: 13px; transition: .15s; }
#btn-cols:hover { background: #e0f0ff; border-color: #1a365d; color: #1a365d; }
#btn-cols.active { background: #1a365d; color: #fff; border-color: #1a365d; }

#col-panel { display: none; position: absolute; top: 100%; left: 0; right: 0;
             background: #fff; border-bottom: 2px solid #1a365d; padding: 10px 20px;
             z-index: 99; box-shadow: 0 4px 12px #0002; }
#col-panel.open { display: flex; gap: 8px; flex-wrap: wrap; align-items: flex-start; }
.col-item { display: flex; align-items: center; gap: 5px; padding: 4px 8px;
            border: 1px solid #dde; border-radius: 4px; background: #f7f8fa;
            cursor: grab; user-select: none; font-size: 12px; white-space: nowrap;
            transition: background .1s; }
.col-item:hover { background: #e8f0ff; border-color: #1a365d; }
.col-item.drag-over { border: 2px dashed #1a365d; background: #dbeafe; }
.col-item.dragging  { opacity: .4; }
.col-item .drag-h { color: #aaa; font-size: 13px; cursor: grab; }
.col-item input[type=checkbox] { cursor: pointer; width: 14px; height: 14px; }
.col-item label { cursor: pointer; }
.col-panel-hint { font-size: 11px; color: #888; align-self: center; margin-left: auto; }
</style>
</head>
<body>

<div class="hdr">
  <h1>TWSE / TPEX &nbsp;競價拍賣資料庫</h1>
  <div class="meta">資料來源：臺灣證券交易所 &nbsp;｜&nbsp; 最後更新：__TIMESTAMP__</div>
</div>

<div class="ctrl" style="position:relative;">
  <div class="tabs">
    <button class="tab active" onclick="switchTab('stocks')">股票競拍 (__N_STOCKS__)</button>
    <button class="tab"        onclick="switchTab('cbs')">CB競拍 (__N_CBS__)</button>
  </div>
  <input type="search" id="search" placeholder="搜尋代號、名稱…" oninput="render()">
  <button id="btn-refresh" onclick="doRefresh()" title="手動更新資料（需透過 auction_server.py 啟動）">
    <span class="spin">↻</span> 更新
  </button>
  <button id="btn-cols" onclick="toggleColPanel()" title="欄位顯示 / 排序">⚙ 欄位</button>
  <span id="toast" class="toast"></span>
  <div class="legend">
    <span><span class="dot dot-done"></span>已開標</span>
    <span><span class="dot dot-noresult"></span>待結果</span>
    <span><span class="dot dot-pending"></span>待投標</span>
    <span><span class="dot dot-cancel"></span>流標/取消</span>
  </div>
  <span class="stat-badge" id="stat"></span>
  <div id="col-panel"></div>
</div>

<div class="tbl-wrap">
  <table>
    <thead id="thead"></thead>
    <tbody id="tbody"></tbody>
  </table>
  <div class="empty-msg" id="empty" style="display:none">沒有符合的資料</div>
</div>

<script>
const RAW = {
  stocks: __STOCKS_JSON__,
  cbs:    __CBS_JSON__
};

const STOCK_COLS = [
  {k:'開標日期',                    lab:'開標日期',    t:'date'},
  {k:'證券代號',                    lab:'代號',        t:'str',  cls:'mono'},
  {k:'證券名稱',                    lab:'名稱',        t:'str',  cls:'bold'},
  {k:'發行市場',                    lab:'市場',        t:'str'},
  {k:'發行性質',                    lab:'性質',        t:'str'},
  {k:'投標結束日',                  lab:'投標結束',    t:'date'},
  {k:'競拍數量(張)',                 lab:'競拍量(張)',  t:'num'},
  {k:'最低投標價格(元)',             lab:'最低投標價',  t:'num'},
  {k:'合格投標數量(張)',             lab:'合格投標量',  t:'num'},
  {k:'最低得標價格(元)',             lab:'最低得標',    t:'num'},
  {k:'最高得標價格(元)',             lab:'最高得標',    t:'num'},
  {k:'得標加權平均價格(元)',         lab:'得標均價',    t:'num'},
  {k:'實際承銷價格(元)',             lab:'承銷價',      t:'num'},
  {k:'投標結束日收盤價',             lab:'結束日收盤',  t:'num',  cmp:true},
  {k:'得標折價率',                   lab:'得標折價率%', t:'num',  decimals:1,
   calc: r => { const l=nv(r['最低得標價格(元)']), c=nv(r['投標結束日收盤價']); return (c>0&&!isNaN(l))?l/c*100:NaN; }},
  {k:'取消競價拍賣(流標或取消)',      lab:'取消/流標',  t:'str'},
];

const CB_COLS = [
  {k:'開標日期',                    lab:'開標日期',    t:'date'},
  {k:'證券代號',                    lab:'代號',        t:'str',  cls:'mono'},
  {k:'證券名稱',                    lab:'名稱',        t:'str',  cls:'bold'},
  {k:'發行市場',                    lab:'市場',        t:'str'},
  {k:'發行性質',                    lab:'性質',        t:'str'},
  {k:'投標結束日',                  lab:'投標結束',    t:'date'},
  {k:'競拍數量(張)',                 lab:'競拍量(張)',  t:'num'},
  {k:'最低投標價格(元)',             lab:'最低投標價',  t:'num'},
  {k:'合格投標數量(張)',             lab:'合格投標量',  t:'num'},
  {k:'最低得標價格(元)',             lab:'最低得標',    t:'num'},
  {k:'最高得標價格(元)',             lab:'最高得標',    t:'num'},
  {k:'得標加權平均價格(元)',         lab:'得標均價',    t:'num'},
  {k:'實際承銷價格(元)',             lab:'承銷價',      t:'num'},
  {k:'發行時轉換價',                 lab:'發行轉換價',  t:'num'},
  {k:'投標結束日收盤價',             lab:'結束日收盤',  t:'num',  cmp:true},
  {k:'截拍日Parity',                lab:'截拍Parity%', t:'num',  decimals:1,
   calc: r => { const c=nv(r['投標結束日收盤價']), p=nv(r['發行時轉換價']); return (p>0&&!isNaN(c)&&!isNaN(p))?c/p*100:NaN; }},
  {k:'最低標溢價率',                 lab:'最低標溢價%', t:'num',  decimals:1,
   calc: r => { const l=nv(r['最低得標價格(元)']), c=nv(r['投標結束日收盤價']), p=nv(r['發行時轉換價']); const parity=(p>0&&!isNaN(c))?c/p*100:NaN; return (!isNaN(parity)&&parity>0&&!isNaN(l))?(l/parity-1)*100:NaN; }},
  {k:'取消競價拍賣(流標或取消)',      lab:'取消/流標',  t:'str'},
];

let tab     = 'stocks';
let sortKey = '開標日期';
let sortDir = -1;

// ── Column visibility & order ─────────────────────────────────────────────────
// Stored in localStorage as { order:[...keys], hidden:[...keys] }
const _colStateCache = {};

function _colStateKey(t) { return 'col_state_v1_' + t; }

function getColState(t) {
  if (!_colStateCache[t]) {
    try {
      const d = JSON.parse(localStorage.getItem(_colStateKey(t)) || 'null');
      _colStateCache[t] = { order: d?.order || null, hidden: new Set(d?.hidden || []) };
    } catch { _colStateCache[t] = { order: null, hidden: new Set() }; }
  }
  return _colStateCache[t];
}

function saveColState(t) {
  const s = _colStateCache[t];
  if (!s) return;
  try { localStorage.setItem(_colStateKey(t), JSON.stringify({ order: s.order, hidden: [...s.hidden] })); } catch {}
}

function getEffectiveCols() {
  const base = tab === 'stocks' ? STOCK_COLS : CB_COLS;
  const s = getColState(tab);
  let ordered = base;
  if (s.order) {
    const map = Object.fromEntries(base.map(c => [c.k, c]));
    const known = new Set(s.order);
    ordered = [
      ...s.order.map(k => map[k]).filter(Boolean),
      ...base.filter(c => !known.has(c.k)),   // newly added cols go to end
    ];
  }
  return ordered.filter(c => !s.hidden.has(c.k));
}

// Returns all cols (including hidden) in saved order — for the panel
function getAllColsOrdered() {
  const base = tab === 'stocks' ? STOCK_COLS : CB_COLS;
  const s = getColState(tab);
  if (!s.order) return base;
  const map = Object.fromEntries(base.map(c => [c.k, c]));
  const known = new Set(s.order);
  return [
    ...s.order.map(k => map[k]).filter(Boolean),
    ...base.filter(c => !known.has(c.k)),
  ];
}

// ── Column filters ────────────────────────────────────────────────────────────
const filters   = {};   // colKey → {op, v1, v2}
const filterRaw = {};   // colKey → raw string the user typed

function parseFilter(str) {
  str = str.trim();
  if (!str) return null;
  // range:  50~100  50～100  50-100
  const rng = str.match(/^(-?[\d.]+)\s*[~～\-]\s*(-?[\d.]+)$/);
  if (rng) return {op:'range', v1:+rng[1], v2:+rng[2]};
  // comparison:  >=50  <=50  >50  <50
  const cmp = str.match(/^(>=|<=|>|<)\s*(-?[\d.]+)$/);
  if (cmp) return {op:cmp[1], v1:+cmp[2]};
  // exact number
  if (/^-?[\d.]+$/.test(str)) return {op:'=', v1:+str};
  return null;
}

function applyFilter(f, n) {
  if (isNaN(n)) return false;
  if (f.op === 'range') return n >= f.v1 && n <= f.v2;
  if (f.op === '>')     return n >  f.v1;
  if (f.op === '<')     return n <  f.v1;
  if (f.op === '>=')    return n >= f.v1;
  if (f.op === '<=')    return n <= f.v1;
  if (f.op === '=')     return n === f.v1;
  return true;
}

function onFilterInput(inp) {
  const key = inp.dataset.key;
  filterRaw[key] = inp.value;          // keep raw text so rebuild can restore it
  const f = parseFilter(inp.value);
  if (f) filters[key] = f; else delete filters[key];
  inp.classList.toggle('active', !!f);
  render();
}

// ── Server / refresh ──────────────────────────────────────────────────────────
// Always try localhost:8787 — works from both http:// and file://
// (server sends Access-Control-Allow-Origin: * so CORS is not an issue)
const LOCAL    = 'http://localhost:8787';
const IS_LOCAL = (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1');
const IS_HTTP  = (window.location.protocol !== 'file:');
let   _srv       = IS_LOCAL ? '' : LOCAL;  // base URL; confirmed after ping
let   _srvReady  = IS_LOCAL;               // true once server confirmed reachable
let   _pollTimer = null;

function showToast(msg, cls, ms = 5000) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className   = 'toast ' + cls;
  if (ms > 0) setTimeout(() => { t.className = 'toast'; }, ms);
}

function _fetchT(url, ms = 3000) {
  const ctrl = new AbortController();
  const tid  = setTimeout(() => ctrl.abort(), ms);
  return fetch(url, {signal: ctrl.signal}).finally(() => clearTimeout(tid));
}

async function _pingServer() {
  try {
    const r = await _fetchT(LOCAL + '/api/refresh-status', 2500);
    _srvReady = r.ok;
  } catch { _srvReady = false; }
  return _srvReady;
}

async function refreshData() {
  const oldS = RAW.stocks.length, oldC = RAW.cbs.length;
  const oldPend = RAW.stocks.filter(r => rowStatus(r) === 'noresult').length
                + RAW.cbs.filter(r    => rowStatus(r) === 'noresult').length;
  try {
    const d = await _fetchT(_srv + '/api/data').then(r => r.json());
    RAW.stocks = d.stocks;
    RAW.cbs    = d.cbs;
    document.querySelectorAll('.tab')[0].textContent = `股票競拍 (${RAW.stocks.length})`;
    document.querySelectorAll('.tab')[1].textContent = `CB競拍 (${RAW.cbs.length})`;
    if (d.timestamp) {
      document.querySelector('.hdr .meta').innerHTML =
        `資料來源：臺灣證券交易所 &nbsp;｜&nbsp; 最後更新：${d.timestamp}`;
    }
    const newPend    = RAW.stocks.filter(r => rowStatus(r) === 'noresult').length
                     + RAW.cbs.filter(r    => rowStatus(r) === 'noresult').length;
    const addedTotal  = (RAW.stocks.length - oldS) + (RAW.cbs.length - oldC);
    const newlyOpened = Math.max(0, oldPend - newPend);
    let msg = '資料已是最新';
    if (addedTotal > 0)    msg = `新增 ${addedTotal} 筆競拍（股票+${RAW.stocks.length-oldS} CB+${RAW.cbs.length-oldC}）`;
    else if (newlyOpened > 0) msg = `${newlyOpened} 筆已補入開標結果`;
    showToast(msg, addedTotal + newlyOpened > 0 ? 'success' : 'info');
    render();
  } catch(e) {
    showToast('載入資料失敗', 'error');
  }
}

function _startPolling(btn) {
  if (_pollTimer) clearInterval(_pollTimer);
  _pollTimer = setInterval(async () => {
    const s = await _fetchT(_srv + '/api/refresh-status')
                    .then(r => r.json()).catch(() => ({running: false}));
    if (!s.running) {
      clearInterval(_pollTimer); _pollTimer = null;
      btn.disabled = false; btn.classList.remove('spinning');
      await refreshData();
    }
  }, 2000);
}

async function doRefresh() {
  const btn = document.getElementById('btn-refresh');

  // If opened as file://, probe the server first (quick 2.5s timeout)
  if (!IS_HTTP) {
    showToast('連線本機伺服器中…', 'info', 0);
    const ok = await _pingServer();
    if (!ok) {
      // 伺服器未啟動：顯示固定提示列，不做任何自動跳頁
      _showServerBanner();
      btn.disabled = false;
      return;
    }
    _srv = LOCAL;
  }

  btn.disabled = true;
  btn.classList.add('spinning');
  showToast('正在連線 TWSE 更新資料…', 'info', 0);
  try {
    await _fetchT(_srv + '/api/refresh');
    _startPolling(btn);
  } catch(e) {
    btn.disabled = false; btn.classList.remove('spinning');
    showToast('無法連線伺服器', 'error');
  }
}

function _showServerBanner() {
  if (document.getElementById('_srv-banner')) return;
  const d = document.createElement('div');
  d.id = '_srv-banner';
  d.style.cssText =
    'position:fixed;bottom:16px;left:50%;transform:translateX(-50%);' +
    'background:#1a365d;color:#fff;padding:12px 20px;border-radius:8px;' +
    'font-size:13px;box-shadow:0 4px 16px #0005;z-index:9999;' +
    'display:flex;align-items:center;gap:14px;max-width:480px;';
  d.innerHTML =
    '<span>伺服器未啟動。請雙擊 <b>競拍資料庫.bat</b>，之後按更新即可。' +
    '（已設定開機自動啟動，下次開機後無需此步驟）</span>' +
    '<button onclick="this.parentNode.remove()" ' +
    'style="background:transparent;border:1px solid #fff8;border-radius:4px;' +
    'color:#fff;cursor:pointer;padding:2px 8px;font:inherit;white-space:nowrap;">✕</button>';
  document.body.appendChild(d);
  // Auto-remove after 15s
  setTimeout(() => d.remove(), 15000);
}

// ── Column panel ──────────────────────────────────────────────────────────────
let _dragSrc = null;

function toggleColPanel() {
  const panel = document.getElementById('col-panel');
  const btn   = document.getElementById('btn-cols');
  if (panel.classList.contains('open')) {
    panel.classList.remove('open'); btn.classList.remove('active');
  } else {
    _buildColPanel(); panel.classList.add('open'); btn.classList.add('active');
  }
}

function _buildColPanel() {
  const panel = document.getElementById('col-panel');
  const s     = getColState(tab);
  const all   = getAllColsOrdered();

  panel.innerHTML = all.map(c => {
    const hidden  = s.hidden.has(c.k);
    const safeK   = c.k.replace(/"/g, '&quot;');
    const safeL   = c.lab.replace(/</g,'&lt;').replace(/>/g,'&gt;');
    return `<div class="col-item" draggable="true" data-k="${safeK}"
      ondragstart="_colDragStart(event)" ondragover="_colDragOver(event)"
      ondragleave="_colDragLeave(event)" ondrop="_colDrop(event)" ondragend="_colDragEnd(event)">
      <span class="drag-h">⠿</span>
      <input type="checkbox" id="chk_${safeK}" ${hidden?'':'checked'}
        onchange="_colToggle('${safeK}',this.checked)" onclick="event.stopPropagation()">
      <label for="chk_${safeK}" onclick="event.stopPropagation()">${safeL}</label>
    </div>`;
  }).join('') + '<span class="col-panel-hint">拖曳可調整順序</span>';
}

function _colToggle(k, visible) {
  const s = getColState(tab);
  if (visible) s.hidden.delete(k); else s.hidden.add(k);
  saveColState(tab);
  render();
}

function _colDragStart(e) {
  _dragSrc = e.currentTarget;
  _dragSrc.classList.add('dragging');
  e.dataTransfer.effectAllowed = 'move';
}
function _colDragOver(e) {
  e.preventDefault(); e.dataTransfer.dropEffect = 'move';
  const target = e.currentTarget;
  if (target !== _dragSrc) target.classList.add('drag-over');
}
function _colDragLeave(e) { e.currentTarget.classList.remove('drag-over'); }
function _colDragEnd(e)   { e.currentTarget.classList.remove('dragging'); }

function _colDrop(e) {
  e.preventDefault();
  const target = e.currentTarget;
  target.classList.remove('drag-over');
  if (!_dragSrc || _dragSrc === target) return;

  // Reorder DOM
  const panel = document.getElementById('col-panel');
  const items = [...panel.querySelectorAll('.col-item')];
  const fromIdx = items.indexOf(_dragSrc);
  const toIdx   = items.indexOf(target);
  if (fromIdx < 0 || toIdx < 0) return;
  if (fromIdx < toIdx) panel.insertBefore(_dragSrc, target.nextSibling);
  else                 panel.insertBefore(_dragSrc, target);

  // Save new order
  const newOrder = [...panel.querySelectorAll('.col-item')].map(el => el.dataset.k);
  getColState(tab).order = newOrder;
  saveColState(tab);
  render();
  _dragSrc = null;
}

// ── Core functions ────────────────────────────────────────────────────────────
function switchTab(t) {
  tab     = t;
  sortKey = '開標日期';
  sortDir = -1;
  Object.keys(filters).forEach(k => delete filters[k]);
  Object.keys(filterRaw).forEach(k => delete filterRaw[k]);
  document.getElementById('search').value = '';
  document.querySelectorAll('.tab').forEach((b,i) =>
    b.classList.toggle('active', (t==='stocks') === (i===0)));
  const panel = document.getElementById('col-panel');
  if (panel && panel.classList.contains('open')) _buildColPanel();
  render();
}

function nv(s) {
  if (s == null || s === '' || s === 'nan' || s === 'None') return NaN;
  return parseFloat(String(s).replace(/,/g,''));
}

function cmpV(a, b, t) {
  if (t === 'num') {
    const na = nv(a), nb = nv(b);
    if (isNaN(na) && isNaN(nb)) return 0;
    if (isNaN(na)) return 1;
    if (isNaN(nb)) return -1;
    return na - nb;
  }
  return String(a||'').localeCompare(String(b||''), 'zh-TW');
}

function rowStatus(row) {
  const cancel = String(row['取消競價拍賣(流標或取消)']||'').trim();
  if (cancel) return 'cancel';
  const bidEnd = String(row['投標結束日']||'').trim();
  if (!bidEnd) return 'done';
  const parts = bidEnd.split('/');
  const d = new Date(+parts[0], +parts[1]-1, +parts[2]);
  const today = new Date(); today.setHours(0,0,0,0);
  if (d >= today) return 'pending';
  const qty = nv(row['合格投標數量(張)']);
  if (isNaN(qty) || qty === 0) return 'noresult';
  return 'done';
}

function fmtN(val, decimals) {
  const n = nv(val);
  if (isNaN(n)) return '';
  return n.toLocaleString('zh-TW', {maximumFractionDigits: decimals ?? 2});
}

function colVal(col, row) {
  return col.calc ? col.calc(row) : row[col.k];
}

function cellHtml(col, row) {
  const raw = colVal(col, row);
  const str = (raw == null || raw === 'nan' || raw === 'None' || (typeof raw === 'number' && isNaN(raw))) ? '' : String(raw).trim();

  if (col.k === '發行市場') {
    if (!str) return '';
    const short = str.includes('集中') ? '上市' : str.includes('櫃') ? '上櫃' : str;
    const cls   = str.includes('集中') ? 'b-tse' : 'b-otc';
    return `<span class="badge ${cls}">${short}</span>`;
  }
  if (col.k === '取消競價拍賣(流標或取消)' && str) {
    return `<span class="badge b-cncl">${str}</span>`;
  }
  if (col.t === 'num') {
    const n = typeof raw === 'number' ? raw : nv(raw);
    if (isNaN(n)) return '';
    const decimals = col.decimals ?? 2;
    const fmt = v => (typeof v === 'number' ? v : nv(v)).toLocaleString('zh-TW', {maximumFractionDigits: decimals, minimumFractionDigits: decimals > 0 ? decimals : 0});
    if (col.cmp) {
      const avg = nv(row['得標加權平均價格(元)']);
      if (!isNaN(avg) && avg > 0) {
        if (n > avg) return `<span class="hi-up">${fmt(n)}</span>`;
        if (n < avg) return `<span class="hi-dn">${fmt(n)}</span>`;
      }
    }
    return fmt(n);
  }
  return str;
}

function render() {
  const cols    = getEffectiveCols();
  const allCols = tab === 'stocks' ? STOCK_COLS : CB_COLS;  // full list for filter/sort lookup
  const raw     = RAW[tab];
  const q       = document.getElementById('search').value.toLowerCase().trim();

  // ① text search
  let rows = q ? raw.filter(r =>
    (String(r['證券代號']||'') + String(r['證券名稱']||'') + String(r['開標日期']||'')).toLowerCase().includes(q)
  ) : [...raw];

  // ② column filters (numeric) — apply against ALL cols, even if currently hidden
  const activeFilters = Object.entries(filters);
  if (activeFilters.length) {
    rows = rows.filter(row =>
      activeFilters.every(([key, f]) => {
        const c = allCols.find(c => c.k === key);
        const v = c ? (c.calc ? c.calc(row) : nv(row[key])) : nv(row[key]);
        return applyFilter(f, typeof v === 'number' ? v : nv(v));
      })
    );
  }

  // ③ sort — look up col def from full list so computed cols sort correctly
  const colDef = allCols.find(c => c.k === sortKey) || allCols[0];
  rows.sort((a, b) => {
    const va = colDef.calc ? colDef.calc(a) : a[sortKey];
    const vb = colDef.calc ? colDef.calc(b) : b[sortKey];
    return sortDir * cmpV(va, vb, colDef.t);
  });

  // ── Header row ──
  const headerRow = '<tr>' + cols.map(c => {
    const active = c.k === sortKey;
    const cls    = active ? (sortDir === 1 ? 'asc' : 'desc') : '';
    return `<th class="${cls}" onclick="sortBy('${c.k}')">${c.lab}<span class="si"></span></th>`;
  }).join('') + '</tr>';

  // ── Filter row ──
  // Remember focused filter input so we can restore it after innerHTML rebuild
  const focusedEl  = document.activeElement;
  const focusedKey = focusedEl?.classList?.contains('fi') ? focusedEl.dataset.key : null;
  const focusedPos = focusedKey ? focusedEl.selectionStart : 0;

  const filterRow = '<tr class="filter-row">' + cols.map(c => {
    if (c.t !== 'num') return '<th></th>';
    const safeKey = c.k.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
    const curF    = filters[c.k];
    // Prefer raw text user typed (handles partial input like ">" or "90-")
    let curV = '';
    if (c.k in filterRaw) {
      curV = filterRaw[c.k];
    } else if (curF) {
      curV = (curF.op === 'range') ? `${curF.v1}~${curF.v2}` : `${curF.op}${curF.v1}`;
    }
    const safeV = curV.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
    return `<th><input class="fi${curF?' active':''}" data-key="${safeKey}" ` +
           `value="${safeV}" placeholder=">50 &lt;100 50~100" ` +
           `oninput="onFilterInput(this)" onclick="event.stopPropagation()" ` +
           `title="篩選：&gt;90  &lt;90  >=90  <=90  90~105  90-105"></th>`;
  }).join('') + '</tr>';

  document.getElementById('thead').innerHTML = headerRow + filterRow;

  // Restore focus & cursor to the filter input that was active before rebuild
  if (focusedKey) {
    const el = document.querySelector(`.fi[data-key="${focusedKey.replace(/\\/g,'\\\\').replace(/"/g,'\\"')}"]`);
    if (el) { el.focus(); el.setSelectionRange(focusedPos, focusedPos); }
  }

  // ── Body ──
  const empty = document.getElementById('empty');
  const tbody = document.getElementById('tbody');
  if (rows.length === 0) {
    tbody.innerHTML = '';
    empty.style.display = '';
  } else {
    empty.style.display = 'none';
    tbody.innerHTML = rows.map(row => {
      const s     = rowStatus(row);
      const cells = cols.map(c => {
        let cls = c.t === 'num' ? ' r' : '';
        if (c.cls) cls += ' ' + c.cls;
        return `<td class="${cls.trim()}">${cellHtml(c, row)}</td>`;
      }).join('');
      return `<tr class="s-${s}">${cells}</tr>`;
    }).join('');
  }
  const total = activeFilters.length ? `${rows.length} / ${raw.length}` : `${rows.length}`;
  document.getElementById('stat').textContent = `${total} 筆`;
}

function sortBy(key) {
  sortDir = (sortKey === key) ? -sortDir : -1;
  sortKey = key;
  render();
}

// ── Init ──────────────────────────────────────────────────────────────────────
// Close col panel when clicking outside
document.addEventListener('click', e => {
  const panel = document.getElementById('col-panel');
  const btn   = document.getElementById('btn-cols');
  if (panel && panel.classList.contains('open') &&
      !panel.contains(e.target) && e.target !== btn && !btn.contains(e.target)) {
    panel.classList.remove('open');
    btn.classList.remove('active');
  }
});

render();

// On page load: silently try the local server (works for both http:// and file://)
window.addEventListener('DOMContentLoaded', async () => {
  const ok = _srvReady || await _pingServer();
  if (!ok) return;   // server not running — use inline data as-is

  // Server is up: load latest data & check if auto-update is in progress
  await refreshData().catch(() => {});
  const s = await _fetchT(_srv + '/api/refresh-status').then(r => r.json()).catch(() => null);
  if (s && s.running) {
    const btn = document.getElementById('btn-refresh');
    btn.disabled = true; btn.classList.add('spinning');
    showToast('自動更新中…', 'info', 0);
    _startPolling(btn);
  }
});
</script>
</body>
</html>"""

# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == '__main__':
    if '--html' in sys.argv:
        _load_price_cache()
        stocks, cbs = load_data()
        if stocks.empty:
            log.error('無資料，請先執行全量下載')
        else:
            generate_html(stocks, cbs)
    elif '--update' in sys.argv:
        update_data()
    else:
        download_history()
