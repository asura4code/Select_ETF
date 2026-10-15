# -*- coding: utf-8 -*-
"""
Created on Sat Oct 10 19:09:04 2026

@author: Arsene
"""

# 程式目的；產出值得投資的台股ETF清單(全)
# 程式邏輯：
# 1. 考慮成立以來每日定額績效正報酬，且贏過大盤(0050)之ETF
# 2. K值 < 50
# 3. 紀錄近期首次入選日期

# 2025/6/18 因股利資訊有誤，改為只參考資本利得報酬(不含息)
# 2025/9/27 因原先抓取ETF名單網頁不見了，故改以excel維護名單，資料來源: https://tw.stock.yahoo.com/tw-etf/search
# 2025/10/13 調升K值門檻至50

"""
< 2026/10/3 >
ETF動態篩選與回測模型 (v7 整合版)
- 資料源：Yahoo Finance v8 Chart API (含配息解析，高效率下載)
- 定期定額：日買進 1 元，考慮除息當下的累積股數配息
- 回測機制：保留 v6 之動態滾動評分與重疊加碼攤平進出場
- 報表輸出：保留完整 Markdown 資訊提示與 Excel 動態清單更新

< 2026/10/10 > v8 雲端版
調整資料源：FinMind 抓日K + yfinance 抓配息
手動維護ETF分割資訊
"""
#%% 0. 匯入所需函式庫

import datetime as dt
import io
import os
import sys
import time
import warnings
from FinMind.data import DataLoader
from loguru import logger
import numpy as np
import pandas as pd
import requests
import talib
from tqdm import tqdm

warnings.filterwarnings("ignore")

# 壓制 FinMind 底層冗長的 debug/info log
logger.remove()
logger.add(sys.stderr, level="WARNING")

# 初始化 FinMind DataLoader
dl = DataLoader()
# 若有個人 Token，請在此帶入（未登入每小時限額約 300 次，帶 Token 為 600 次）：
dl.login_by_token(api_token = os.getenv('FM_DL_TOKEN'))

#%% 1. 自定義函數


# 歷史台股 ETF 分割紀錄表：{代號: [(生效日期, 分割倍數)]}
ETF_SPLITS_MAP = {
    # 1. 元大台灣50: 2025/06/18 起 1 拆 4
    "0050.TW": [
        ("2025-06-18", 4.0),
    ],
    # 2. 國泰臺灣加權正2: 2025/06/11 起 1 拆 7
    "00663L.TW": [
        ("2025-06-11", 7.0),
    ],
    # 3. 富邦科技 (小台積): 2025/11/26 起 1 拆 7
    "0052.TW": [
        ("2025-11-26", 7.0),
    ],
    # 4. 元大台灣50正2: 2023 年 1 拆 6，2026 年 1 拆 22
    "00631L.TW": [
        ("2026-03-31", 22.0),
    ],
    "00685L.TW": [
        ("2026-07-07", 24.0),
    ],
    # 5. 反向 ETF 反分割系列
    "00632R.TW": [
        ("2024-12-11", 1.0 / 7.0),  # 7 股併 1 股
    ],
    "00676R.TW": [
        ("2025-02-19", 1.0 / 6.0),  # 6 股併 1 股
    ],
    "00674R.TW": [
        ("2026-04-22", 0.2),  # 5 股併 1 股
    ],
}

def apply_split_adjustment(df_price, ticker):
    """依據分割對應表執行歷史行情向後還原 (Backward Adjustment)

    :param df_price: DataFrame，索引為 'YYYY-MM-DD' 格式之日期字串
    :param ticker: 包含後綴的股票代號 (如 '0050.TW')
    :return: 完成還原的 DataFrame
    """
    if ticker not in ETF_SPLITS_MAP or df_price is None or df_price.empty:
        return df_price

    price_cols = [
        col for col in ["Open", "High", "Low", "Close"] if col in df_price.columns
    ]

    # 依日期遞增排序處理，支援多階段分割
    events = sorted(ETF_SPLITS_MAP[ticker], key=lambda x: x[0])

    for split_date, ratio in events:
        if ratio <= 0 or ratio == 1.0:
            continue

        # 篩選分割生效日「之前」的所有交易日
        mask = df_price.index < split_date
        if not mask.any():
            continue

        # 1. 歷史價格等比縮放
        df_price.loc[mask, price_cols] = df_price.loc[mask, price_cols] / ratio

        # 2. 歷史成交量等比調整 (股數放大)
        if "Volume" in df_price.columns:
            df_price.loc[mask, "Volume"] = df_price.loc[mask, "Volume"] * ratio

        # 3. 歷史每股配息等比調整 (對齊分割後股數基準)
        if "Dividend" in df_price.columns:
            df_price.loc[mask, "Dividend"] = (
                df_price.loc[mask, "Dividend"] / ratio
            )

    return df_price

# 核心資料爬取函式 (FinMind 日K + yfinance 配息)

def fetch_yahoo_dividends(ticker, start_date="2009-01-01", end_date=None):
    """透過 Yahoo Finance v8 Chart API 抓取歷史配息資料"""
    if end_date is None:
        end_date = (dt.datetime.today() + dt.timedelta(days=1)).strftime("%Y-%m-%d")

    period1 = int(
        time.mktime(dt.datetime.strptime(start_date, "%Y-%m-%d").timetuple())
    )
    period2 = int(
        time.mktime(dt.datetime.strptime(end_date, "%Y-%m-%d").timetuple())
    )

    url = f"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}"
    params = {
        "period1": period1,
        "period2": period2,
        "interval": "1d",
        "events": "div,splits",
    }
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        ),
        "Accept": "application/json",
    }

    try:
        resp = requests.get(url, params=params, headers=headers, timeout=12)
        if resp.status_code != 200:
            return {}

        data = resp.json()
        chart = data.get("chart", {}).get("result")
        if not chart:
            return {}

        events = chart[0].get("events", {})
        div_events = events.get("dividends", {})
        div_dict = {}
        for _, v in div_events.items():
            d_str = dt.datetime.fromtimestamp(v.get("date", 0)).strftime("%Y-%m-%d")
            div_dict[d_str] = v.get("amount", 0.0)

        return div_dict
    except Exception:
        return {}

def fetch_etf_data(ticker, start_date="2009-01-01", end_date=None):
  if end_date is None:
    end_date = (dt.datetime.today() + dt.timedelta(days=1)).strftime("%Y-%m-%d")

  # 1. 取得純代號 (0050.TW -> 0050)
  stock_id = ticker.split(".")[0]

  try:
    # 2. FinMind 抓取原始日K
    df_price = dl.taiwan_stock_daily(
        stock_id=stock_id, start_date=start_date, end_date=end_date
    )
    if df_price is None or df_price.empty or len(df_price) < 20:
      return None

    df_price = df_price.rename(
        columns={
            "date": "Date",
            "open": "Open",
            "max": "High",
            "min": "Low",
            "close": "Close",
            "Trading_Volume": "Volume",
            "Trading_money": "Amount",
        }
    )
    df_price["Date"] = pd.to_datetime(df_price["Date"]).dt.strftime("%Y-%m-%d")
    df_price = df_price.set_index("Date").sort_index()

    # 確保數值型態正確
    numeric_cols = ["Open", "High", "Low", "Close", "Volume"]
    for col in numeric_cols:
      if col in df_price.columns:
        df_price[col] = pd.to_numeric(df_price[col], errors="coerce")
    df_price = df_price.dropna(subset=["Close"]).drop_duplicates()

    # 3. Yahoo Chart API 抓取配息資訊
    div_map = fetch_yahoo_dividends(
        ticker, start_date=start_date, end_date=end_date
    )
    df_price["Dividend"] = df_price.index.map(div_map).fillna(0.0)

    # 4. 執行分割與反分割還原 (確保 0050、00631L 等歷史走勢連續)
    df_price = apply_split_adjustment(df_price, ticker)

    # 5. 剔除特殊休市異常日
    if "2025-09-29" in df_price.index:
      df_price = df_price.drop("2025-09-29")

    return df_price

  except Exception:
    return None

# 定期定額計算函式 (模擬實際累積股數與配息收入)

def calc_dca_with_dividends(price_series, div_series):
    """
    計算每日定期定額買進 1 元，並在每次配息日依據當下累積股數計算股息收益。
    回傳：
        total_rt: 最終累積總報酬率
        ret_series: 每日滾動累積報酬率序列 (對齊日期)
    """
    df_calc = pd.DataFrame(
        {"close": price_series, "dividend": div_series}
    ).fillna(0.0)

    # 每日定額買進 1 元所換得的股數
    df_calc["volumes"] = 1.0 / df_calc["close"]
    # 累積持有股數
    df_calc["vol_ttl"] = df_calc["volumes"].cumsum()
    # 每次配息所產生的現金股利收入 (除息當天累積股數 * 當期每股配息)
    df_calc["val_dvd"] = df_calc["dividend"] * df_calc["vol_ttl"]
    df_calc["cum_val_dvd"] = df_calc["val_dvd"].cumsum()

    # 每日投入成本累計 = 1, 2, 3, ..., N
    cost_series = np.arange(1, len(df_calc) + 1)
    # 每日總資產市值 = 當日總持有股數 * 當日收盤價 + 累積現金股利
    df_calc["total_val"] = (
        df_calc["vol_ttl"] * df_calc["close"] + df_calc["cum_val_dvd"]
    )
    df_calc["return"] = (df_calc["total_val"] / cost_series) - 1.0

    total_rt = df_calc["return"].iloc[-1]
    return total_rt, df_calc["return"]

#%% 2. 抓取上市櫃ETF代號
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Referer": "https://www.tpex.org.tw/web/stock/aftertrading/daily_close_quotes/stk_quote.php?l=zh-tw",
}

# ---------------------------------------------------------
# 第1步：抓取上市 ETF (TWSE)
# ---------------------------------------------------------
print("正在取得上市 ETF 清單 (TWSE)...")
twse_direct_url = "https://www.twse.com.tw/exchangeReport/STOCK_DAY_ALL?response=open_data"

twse_list = []
try:
    resp = requests.get(twse_direct_url, headers=headers, timeout=20)
    resp.encoding = "utf-8"
    df_twse = pd.read_csv(io.StringIO(resp.text))
    
    code_col = [c for c in df_twse.columns if "代號" in c][0]
    name_col = [c for c in df_twse.columns if "名稱" in c][0]
    
    for _, row in df_twse.iterrows():
        code = str(row[code_col]).replace('"', '').strip()
        name = str(row[name_col]).replace('"', '').strip()
        # 只要以 00 開頭即為 ETF（涵蓋 0050, 0056 等 4 碼標的與 00878, 00632R 等）
        if code.startswith("00"):
            twse_list.append({
                "證券代號": f"{code}.TW",
                "證券名稱": name
            })
    print(f"成功取得上市 ETF: {len(twse_list)} 檔")
except Exception as e:
    print(f"上市 ETF 擷取異常: {e}")

# ---------------------------------------------------------
# 第2步：抓取上櫃 ETF (TPEx) - 改用穩定 JSON 介面
# ---------------------------------------------------------
print("正在取得上櫃 ETF 清單 (TPEx)...")
# 該端點專門回傳每日收盤行情 JSON，不需預先指定日期即可抓取最新一日
tpex_json_url = "https://www.tpex.org.tw/web/stock/aftertrading/daily_close_quotes/stk_quote_result.php?l=zh-tw&o=json"

tpex_list = []
try:
    tpex_resp = requests.get(tpex_json_url, headers=headers, timeout=20)
    tpex_data = tpex_resp.json()
    
    # 櫃買中心回傳結構中的資料列為 aaData 或 tables[0]['data']
    rows = tpex_data.get("aaData", [])
    if not rows and "tables" in tpex_data and len(tpex_data["tables"]) > 0:
        rows = tpex_data["tables"][0].get("data", [])

    for item in rows:
        # 每列陣列索引 0 為代號，索引 1 為名稱
        code = str(item[0]).strip()
        name = str(item[1]).strip()
        
        # 上櫃 ETF 特徵：代號以 00 開頭（如 00679B, 00720B 等多數為債券 ETF）
        if code.startswith("00"):
            tpex_list.append({
                "證券代號": f"{code}.TWO",
                "證券名稱": name
            })
    print(f"成功取得上櫃 ETF: {len(tpex_list)} 檔")
except Exception as e:
    print(f"上櫃 ETF 擷取異常: {e}")

# ---------------------------------------------------------
# 第3步：合併、排序
# ---------------------------------------------------------
all_list = twse_list + tpex_list


df_result = pd.DataFrame(all_list)
df_result = df_result.drop_duplicates(subset=["證券代號"]).sort_values(by="證券代號").reset_index(drop=True)

ticker_list = df_result.set_index('證券代號', inplace = False)

#%% 3. 主流程與資料抓取

start_time = time.time()
day_begin = "1980-01-01"
day_end = (dt.datetime.today() + dt.timedelta(days=1)).strftime("%Y-%m-%d")
k_th = 50

# 載入名單
# ticker_list = pd.read_excel("台股ETF名單.xlsx", sheet_name="ETF總表", index_col=0)
stock_list = ticker_list.index.tolist()
if "0050.TW" not in stock_list:
    stock_list.append("0050.TW")

print(f"開始下載 {len(stock_list)} 檔標的歷史行情 (FinMind) 與配息資訊 (Yahoo Chart API)...")
all_data = {}
for s in tqdm(stock_list, desc="下載行情"):
    # 傳入保留後綴的 ticker (如 '0050.TW', '00679B.TWO')
    df_s = fetch_etf_data(s, start_date=day_begin, end_date=day_end)
    if df_s is not None and len(df_s) >= 20:
        all_data[s] = df_s
    time.sleep(0.04)

print(
    f"下載完成，共成功取得 {len(all_data)} 檔資料，耗時：{(time.time() - start_time)/60:.1f} 分鐘"
)

# 整理收盤價、成交量與配息寬表格
prices_df = pd.DataFrame(
    {s: all_data[s]["Close"] for s in all_data}
).sort_index()
volume_df = pd.DataFrame(
    {s: all_data[s]["Volume"] for s in all_data}
).sort_index()
dividends_df = (
    pd.DataFrame({s: all_data[s]["Dividend"] for s in all_data})
    .sort_index()
    .fillna(0.0)
)

if "0050.TW" not in all_data:
    raise ValueError("未能順利取得 0050.TW 歷史資料，無法進行基準比較！")

#%% 4. 同期定期定額績效評估 (對決 0050)

t3 = time.time()
base_df = all_data["0050.TW"]

etf_list_records = []
score_dict = {}  # 存放各標的各交易日的動態評分

# candidates = [s for s in prices_df.columns if s != "0050.TW"]
candidates = [s for s in prices_df.columns]

for t in candidates:

    t_df = all_data[t][["Close", "Dividend"]].dropna(subset=["Close"])
    t_df = t_df[t_df['Close'] > 0]
    if len(t_df) < 20:
        continue

    start_date = t_df.index[0]
    # 取出 0050 同期資料
    b_df = base_df.loc[base_df.index >= start_date, ["Close", "Dividend"]]
    common_idx = t_df.index.intersection(b_df.index)
    if len(common_idx) < 20:
        continue

    t_prices = t_df.loc[common_idx, "Close"]
    t_divs = t_df.loc[common_idx, "Dividend"]
    b_prices = b_df.loc[common_idx, "Close"]
    b_divs = b_df.loc[common_idx, "Dividend"]

    # 計算標的與 0050 的定額累積績效
    rt, t_ret_series = calc_dca_with_dividends(t_prices, t_divs)
    rt_base, b_ret_series = calc_dca_with_dividends(b_prices, b_divs)

    # 計算歷史每日動態評分 (v6邏輯)
    eval_df = pd.DataFrame(
        {"return_x": t_ret_series, "return_y": b_ret_series}
    )
    score_col = np.where(eval_df["return_x"] >= 0, 1, 0)
    score_col = np.where(
        eval_df["return_x"] + 0.0001 >= eval_df["return_y"], 2, score_col
    )
    score_dict[t] = pd.Series(score_col, index=common_idx)

    # 彙整指標
    v_mean = volume_df[t].dropna().tail(20).mean()
    tot_dvd = t_divs.sum()

    diff = rt - rt_base
    final_score = 0
    if rt >= 0:
        final_score = 2 if diff >= -0.0001 else 1

    etf_list_records.append(
        {
            "ticker": t,
            "成立以來報酬": rt,
            "同期大盤報酬": rt_base,
            "報酬差": diff,
            "累計配息": tot_dvd,
            "平均成交金額": v_mean,  # 沿用 v6 欄位命名
            "評分": final_score,
        }
    )

etf_list = pd.DataFrame(etf_list_records).set_index("ticker")

# 篩選正報酬且評分=2 (勝過0050)
mvp_list = etf_list[
    (etf_list["成立以來報酬"] > 0) & (etf_list["評分"] == 2)
].copy()
print(f"比較大盤同期報酬耗時：{(time.time() - t3)/60:.2f} 分鐘")
print(f"滿足正報酬且打敗 0050 的 ETF 共：{len(mvp_list)} 檔")

#%% 5. 技術指標計算與回測 (保留 v6 多筆加碼攤平進出場機制)

def calc_taiwan_kd(high, low, close, n=9, k_period=3, d_period=3):
  """計算符合台灣市場慣例的 KD 指標 (初始值 K=50, D=50，遞迴平滑)

  :param high: 最高價序列 (numpy array 或 pandas Series)
  :param low: 最低價序列
  :param close: 收盤價序列
  :param n: RSV 週期 (預設 9)
  :param k_period: K 平滑週期 (預設 3)
  :param d_period: D 平滑週期 (預設 3)
  :return: (k_series, d_series) 均為 numpy array
  """
  length = len(close)
  k_arr = np.full(length, np.nan)
  d_arr = np.full(length, np.nan)

  if length < n:
    return k_arr, d_arr

  # 計算滾動最高價與最低價 (含當日的前 n 日)
  s_high = pd.Series(high)
  s_low = pd.Series(low)
  s_close = pd.Series(close)

  rolling_max = s_high.rolling(window=n).max().values
  rolling_min = s_low.rolling(window=n).min().values

  # 計算 RSV
  denominator = rolling_max - rolling_min
  # 避免最高價等於最低價除以 0 的情況
  rsv = np.where(denominator != 0, (s_close.values - rolling_min) / denominator * 100, 50.0)

  # 平滑係數： (3-1)/3 = 2/3 與 1/3
  alpha_k = 1.0 / k_period
  alpha_d = 1.0 / d_period

  # 第 n-1 筆（即第 n 個交易日，索引從 0 開始）為第一個有效 RSV
  # 其前一日（n-2）預設初值為 50
  prev_k = 50.0
  prev_d = 50.0

  for i in range(n - 1, length):
    curr_k = (1.0 - alpha_k) * prev_k + alpha_k * rsv[i]
    curr_d = (1.0 - alpha_d) * prev_d + alpha_d * curr_k
    k_arr[i] = curr_k
    d_arr[i] = curr_d
    prev_k = curr_k
    prev_d = curr_d

  return k_arr, d_arr

t4 = time.time()
tech = pd.DataFrame(
    columns=["price", "20dma", "Lower_band", "K值", "成交金額"]
)
back_test = pd.DataFrame(
    columns=[
        "勝率(%)",
        "總報酬率(%)",
        "進場次數",
        "總進場期間",
        "平均報酬率(%)",
        "最大報酬率(%)",
        "最小報酬率(%)",
    ]
)

for stock in tqdm(mvp_list.index, desc="計算技術指標與回測"):
    df_stock = all_data[stock].copy()
    close = df_stock["Close"].values.astype(float)
    high = df_stock["High"].values.astype(float)
    low = df_stock["Low"].values.astype(float)
    volume = df_stock["Volume"].values.astype(float)

    # 技術指標
    ma20 = talib.SMA(close, timeperiod=20)
    std20 = np.full(len(close), np.nan)
    for i in range(19, len(close)):
        std20[i] = np.std(close[i - 19 : i + 1])
    lower_band = ma20 - 2 * std20

    # slowk, slowd = talib.STOCH(
    #     high,
    #     low,
    #     close,
    #     fastk_period=9,
    #     slowk_period=3,
    #     slowk_matype=0,
    #     slowd_period=3,
    #     slowd_matype=0,
    # )
    slowk, slowd = calc_taiwan_kd(high, low, close, n=9, k_period=3, d_period=3)

    df_stock["20dma"] = ma20
    df_stock["20d_std"] = std20
    df_stock["Lower_band"] = lower_band
    df_stock["K值"] = slowk

    # 對齊動態評分
    if stock in score_dict:
        df_stock["評分"] = score_dict[stock]
    else:
        df_stock["評分"] = 0
    df_stock["評分"] = df_stock["評分"].fillna(0)

    # 進出場訊號 (v6 邏輯)
    df_stock["買進訊號"] = (
        (df_stock["評分"] > 0) & (df_stock["K值"] < 40)
    ).astype(int)
    df_stock["賣出訊號"] = (
        (df_stock["K值"] > 70)
        | (df_stock["Close"] > (df_stock["20dma"] + 2 * df_stock["20d_std"]))
    ).astype(int)

    # v6 重疊回測機制 (Nested Loop 逐筆買進獨立追蹤持有期間的加碼與配息)
    ret_records = []
    trade_sub_idx = 0
    price_indices = df_stock.index.tolist()

    for i in range(len(df_stock)):
        if df_stock["買進訊號"].iloc[i] == 1:
            buy_date = price_indices[i]
            # 尋找買進日之後的第一個賣出訊號日
            future_sells = df_stock[
                (df_stock.index > buy_date) & (df_stock["賣出訊號"] > 0)
            ]
            if len(future_sells) > 0:
                sell_date = future_sells.index[0]

                t_sub = df_stock.loc[
                    (df_stock.index >= buy_date) & (df_stock.index <= sell_date)
                ][["Close", "Dividend", "買進訊號"]].copy()

                # 當期買進所累積的股數 (若持有區間內再度觸發買進則加碼)
                t_sub["volumes"] = (1.0 / t_sub["Close"]) * t_sub["買進訊號"]
                t_sub["vol_ttl"] = t_sub["volumes"].cumsum()
                # 持有期間配息累加
                t_sub["val_dvd"] = t_sub["Dividend"] * t_sub["vol_ttl"]

                # 區間總結算報酬率
                buy_count = t_sub["買進訊號"].sum()
                if buy_count > 0:
                    rt_trade = (
                        t_sub["Close"].iloc[-1] * t_sub["vol_ttl"].iloc[-1]
                        + t_sub["val_dvd"].sum()
                    ) / buy_count - 1.0
                    dur = len(t_sub)
                    ret_records.append(
                        {
                            "買進日": buy_date,
                            "賣出日": sell_date,
                            "含息報酬": rt_trade,
                            "持有期間": dur,
                        }
                    )

    if len(ret_records) > 0:
        ret_df = pd.DataFrame(ret_records)
        win_rate = round(sum(ret_df["含息報酬"] > 0) / len(ret_df) * 100, 2)
        tot_ret = round(sum(ret_df["含息報酬"]) * 100, 2)
        avg_ret = round(sum(ret_df["含息報酬"]) * 100 / len(ret_df), 2)
        max_ret = round(ret_df["含息報酬"].max() * 100, 2)
        min_ret = round(ret_df["含息報酬"].min() * 100, 2)
        tot_dur = sum(ret_df["持有期間"])

        tem = pd.DataFrame(
            data={
                "勝率(%)": win_rate,
                "總報酬率(%)": tot_ret,
                "進場次數": len(ret_df),
                "總進場期間": tot_dur,
                "平均報酬率(%)": avg_ret,
                "最大報酬率(%)": max_ret,
                "最小報酬率(%)": min_ret,
            },
            index=[stock],
        )
        tem["報酬差(%)"] = tem["最大報酬率(%)"] + tem["最小報酬率(%)"]
        back_test = pd.concat([back_test, tem], axis=0)

    # 最新技術面行情摘要
    last_close = close[-1]
    last_vol = volume[-1]
    temp_tech = pd.DataFrame(
        data={
            "price": last_close,
            "20dma": ma20[-1],
            "Lower_band": lower_band[-1],
            "K值": slowk[-1],
            "成交金額": last_close * last_vol,
        },
        index=[stock],
    )
    tech = pd.concat([tech, temp_tech], axis=0)

# 年化報酬率與平均持有天數計算
back_test["年化報酬率_回測"] = (
    (back_test["總報酬率(%)"] / 100 + 1) ** (365 / back_test["總進場期間"]) - 1
) * 100 * (back_test["總報酬率(%)"] > 0)
back_test["平均持有期間"] = (
    back_test["總進場期間"] / back_test["進場次數"]
)

mvp_list = pd.concat([mvp_list, tech, back_test], axis=1)
print(f"\n計算技術指標與回測耗時：{(time.time() - t4)/60:.2f} 分鐘")

#%% 6. 產出報告與 Excel 動態清單更新 (完全保留 v6 輸出規格)

# 位階指標
mvp_list["p_20dma"] = mvp_list["price"] / mvp_list["20dma"]
mvp_list["p_Lower_band"] = mvp_list["price"] / mvp_list["Lower_band"]

result = mvp_list.sort_values(by=["平均報酬率(%)"], ascending=False)
result["ticker"] = ticker_list.loc[result.index, "證券名稱"]  # 對齊中文代號
result = result[
    (result["K值"] > 0)
    & (result["勝率(%)"] > 0)
    & (result["平均報酬率(%)"] > 0)
    # & ((result['報酬差'] >= 0.1) | (result['報酬差'] == 0))
]
result["年化報酬率"] = (
    (result["總報酬率(%)"] / 100 + 1) ** (365 / result["總進場期間"]) - 1
) * 100 * (result["總報酬率(%)"] > 0)

# 取出前次累計連續入選次數並記錄首次入選日期
p_file = "ETF動態清單.xlsx"
p_result = None
try:
    p_result = pd.read_excel(p_file, index_col=0).drop_duplicates()
    cml = result[result["K值"] < k_th].merge(
        p_result[p_result["K值"] < k_th],
        how="inner",
        left_index=True,
        right_index=True,
    )

    result["首次加入日期"] = np.nan
    result.loc[result["K值"] < k_th, "首次加入日期"] = (
        dt.datetime.today().strftime("%Y-%m-%d")
    )
    result.loc[cml.index, "首次加入日期"] = p_result.loc[
        cml.index, "首次加入日期"
    ]
except Exception:
    result["首次加入日期"] = np.nan
    result.loc[result["K值"] < k_th, "首次加入日期"] = (
        dt.datetime.today().strftime("%Y-%m-%d")
    )

result[["K值"]] = result[["K值"]].astype(float).round(2)
result[["年化報酬率"]] = (
    result[["年化報酬率"]].replace(0, np.nan).astype(float).round(2)
)

# 終端機格式化輸出
if p_result is not None:
    intersec = result[result["評分"] == 2][["ticker"]].merge(
        p_result[p_result["評分"] == 2][["ticker"]],
        how="inner",
        left_index=True,
        right_index=True,
    )
    insert = result[result["評分"] == 2][
        result[result["評分"] == 2][["ticker"]]
        .merge(intersec, how="left", left_index=True, right_index=True)
        .isna()
        .any(axis=1)
    ]
    remove = p_result[p_result["評分"] == 2][
        p_result[p_result["評分"] == 2][["ticker"]]
        .merge(intersec, how="left", left_index=True, right_index=True)
        .isna()
        .any(axis=1)
    ]

    print("[評分2]新增標的：")
    if len(insert) > 0:
        # print(
        #     etf_list.loc[insert.index][
        #         [
        #             "ticker",
        #             "成立以來報酬",
        #             "同期大盤報酬",
        #             "累計配息",
        #             "平均成交金額",
        #         ]
        #     ]
        #     .sort_values(by=["ticker"], ascending=True)
        #     .to_markdown()
        # )
        print(etf_list.loc[insert.index].sort_index().to_markdown())
    else:
        print("無新增標的")
    print("=" * 110)

    print("[評分2]移除標的：")
    if len(remove) > 0:
        # print(
        #     etf_list.loc[remove.index][
        #         [
        #             "ticker",
        #             "成立以來報酬",
        #             "同期大盤報酬",
        #             "累計配息",
        #             "平均成交金額",
        #         ]
        #     ]
        #     .sort_values(by=["ticker"], ascending=True)
        #     .to_markdown()
        # )
        print(etf_list.loc[remove.index].sort_index().to_markdown())
    else:
        print("無移除標的")
    print("=" * 110)

print("位階升溫標的：")
print(
    result[(result["K值"] >= 85)][
        ["ticker", "price", "K值", "勝率(%)", "平均報酬率(%)", "年化報酬率"]
    ]
    .sort_values(by=["K值"], ascending=False)
    .to_markdown()
)
print("=" * 110)

print("成立至今績效勝過同期大盤：")
print(
    result[(result["K值"] < k_th) & (result["評分"] == 2)][
        [
            "ticker",
            "price",
            "K值",
            "勝率(%)",
            "平均報酬率(%)",
            "年化報酬率",
            "首次加入日期",
        ]
    ]
    .sort_values(by=["年化報酬率"], ascending=True)
    .to_markdown()
)
print("=" * 110)

# 存檔輸出
col_list = [
    "ticker",
    "成立以來報酬",
    "同期大盤報酬",
    "price",
    "K值",
    "進場次數",
    "勝率(%)",
    "平均報酬率(%)",
    "平均持有期間",
    "平均成交金額",
    "首次加入日期",
    "評分",
    "最大報酬率(%)",
    "最小報酬率(%)",
    "報酬差(%)",
    "年化報酬率",
    "報酬差",
    "累計配息",
    "總報酬率(%)",
    "總進場期間",
    "p_20dma",
    "p_Lower_band",
    "20dma",
    "Lower_band",
    "成交金額",
]

result[col_list].to_excel(p_file)
end_time = time.time()
print(f"ETF模型全部運算完成，耗時：{(end_time - start_time)/60:.1f}分鐘")

#%% 7. 組裝推播內容並發送 Telegram 通知
import requests

def send_telegram_message(bot_token, chat_id, message_text, parse_mode='HTML'):
  """透過 Telegram Bot API 發送推播訊息

  :param bot_token: BotFather 給予的 HTTP API Token
  :param chat_id: 你的 Telegram 個人 Chat ID (純數字)
  :param message_text: 推播內文 (支援 HTML 粗體/斜體/換行)
  :param parse_mode: 格式化解析模式 ('HTML' 或 'Markdown')
  """
  url = f'https://api.telegram.org/bot{bot_token}/sendMessage'
  payload = {
      'chat_id': chat_id,
      'text': message_text,
      'parse_mode': parse_mode,
      'disable_web_page_preview': True,
  }

  try:
    resp = requests.post(url, json=payload, timeout=10)
    if resp.status_code == 200:
      print('✅ Telegram 訊息推播成功！')
      return True
    else:
      print(f'❌ Telegram 發送失敗，HTTP {resp.status_code}：{resp.text}')
      return False
  except Exception as e:
    print(f'❌ Telegram 連線異常: {e}')
    return False

today_str = dt.datetime.today().strftime('%Y-%m-%d')

# 1. 篩選勝過大盤且位階合理標的
win_base_df = result[(result['K值'] < k_th) & (result['評分'] == 2) & (result['報酬差'] >= 0.1)].sort_values(
    by=['年化報酬率'], ascending=False
)


# 2. 篩選位階升溫 (過熱) 標的
heated_df = result[(result['K值'] >= 85)].sort_values(
    by=['K值'], ascending=False
)

# 3. 建立 HTML 排版內容
lines = [
    f'📈 <b>【動態 ETF 清單】</b> <code>{today_str}</code>',
    '━━━━━━━━━━━━━━━━━━',
]

# 區塊一：勝過大盤標的
lines.append(f'🏆 <b>布局價值 共 {len(win_base_df)} 檔：</b>')
if len(win_base_df) > 0:
  win_base_df['平均成交金額_萬'] = (win_base_df['平均成交金額']/10000).astype(int)
  win_base_df['年化報酬率_int'] = win_base_df['年化報酬率'].astype(int)
  for sym, row in win_base_df.iterrows():
    # 呈現：名稱(代號) / 現價 / K值 / 勝率 / 年化
    lines.append(
        f'• <b>{row["ticker"]}</b> (<code>{sym}</code>)\n'
        f'  現價: <code>{row["price"]:.2f}</code> | K值: <code>{row["K值"]}</code>'
        f' | 平均成交金額: <code>{row["平均成交金額_萬"]}</code> 萬'
        f' | 勝率: <code>{row["勝率(%)"]}%</code>\n'
        f'  年化報酬率: <b>{row["年化報酬率_int"]}%</b>\n'
    )
else:
  lines.append('<i>（今日無符合條件標的）</i>')

lines.append('━━━━━━━━━━━━━━━━━━')

# 區塊二：位階過熱警戒
lines.append(f'🔥 <b>位階升溫 共 {len(heated_df)} 檔：</b>')
if len(heated_df) > 0:
  heated_df['年化報酬率_int'] = heated_df['年化報酬率'].astype(int)
  for sym, row in heated_df.iterrows():
    lines.append(
        f'• <b>{row["ticker"]}</b> (<code>{sym}</code>) | 現價:'
        f' <code>{row["price"]:.2f}</code> | K值: <code>{row["K值"]}</code>'
        f' | 年化報酬率: <code>{row["年化報酬率_int"]}%</code>\n'
    )
else:
  lines.append('<i>（今日無符合條件標的）</i>')

# 區塊三：新增與移除標的 (若有讀取到歷史檔案)
if p_result is not None:
  
  if len(insert) > 0:
    insert['insert'] = insert['ticker'] + '(' + insert.index + ')'
    new_tickers = ', '.join(insert['insert'].tolist())
    lines.append('━━━━━━━━━━━━━━━━━━')
    lines.append(f'✨ <b>今日新增</b>: \n{new_tickers}')
  if len(remove) > 0:
    remove['remove'] = remove['ticker'] + '(' + remove.index + ')'
    del_tickers = ', '.join(remove['remove'].tolist())
    lines.append('━━━━━━━━━━━━━━━━━━')
    lines.append(f'🔻 <b>今日移除</b>: \n{del_tickers}')

message_html = '\n'.join(lines)

# 填入由 @BotFather 與 @userinfobot 取得的資訊
TG_BOT_TOKEN = os.getenv('TG_BOT_TOKEN')
TG_CHAT_ID = os.getenv('TG_CHAT_ID')

send_telegram_message(TG_BOT_TOKEN, TG_CHAT_ID, message_html)