# =============================================================================
# tdcc_weekly.py
# TDCC 大戶持股比週報自動化腳本
#
# 這支程式就像一條自動化的生產線：
#   步驟1（下載）→ 取得原料（從 TDCC 網站下載本週股東持股調查報告）
#   步驟2（清洗）→ 清洗整理原料（去除無效資料、統一格式）
#   步驟3（計算）→ 加工製作（計算大戶/中戶/散戶的持股比例）
#   步驟4（寫入）→ 入庫存檔（將結果寫入歷史 Excel，保留近13週）
#   步驟5（篩選）→ 品質把關（找出連續三週大戶持股比上漲的股票）
#   步驟6（輸出）→ 包裝出貨（輸出可上傳 NotebookLM 的分析摘要）
# =============================================================================

import os
import sys
import time
import requests
import pandas as pd
import openpyxl
from pathlib import Path
from datetime import datetime, timedelta
from io import StringIO

# =============================================================================
# 路徑設定
# 說明：程式會自動判斷執行環境，選擇對應的儲存位置。
#       就像同一個員工在公司上班用公司電腦、在家用自己電腦，
#       工作內容一樣，但存檔位置不同。
#
#   GitHub Actions（雲端自動執行）→ 存放在 repo 的 output/ 資料夾
#   本機執行                      → 存放在家目錄的 TDCC_Analysis 資料夾
# =============================================================================
if os.environ.get("GITHUB_ACTIONS"):
    # 雲端執行：GITHUB_WORKSPACE 是 GitHub Actions 提供的 repo 根目錄路徑
    BASE_DIR = Path(os.environ.get("GITHUB_WORKSPACE", ".")) / "output"
else:
    # 本機執行：存放在使用者家目錄，方便在自己電腦上找到
    BASE_DIR = Path.home() / "Documents" / "TDCC_Analysis"

RAW_DIR     = BASE_DIR / "raw"
DATA_DIR    = BASE_DIR / "data"
SUMMARY_DIR = BASE_DIR / "summary"

# =============================================================================
# TDCC API 設定
# 說明：這是 TDCC（台灣集中保管結算所）的官方開放資料網址，
#       就像每週六都會更新的線上報告，免費公開供所有人下載。
#       TDCC_URL_TEMPLATE 支援帶入日期參數，用於下載歷史資料。
# =============================================================================
TDCC_URL          = "https://smart.tdcc.com.tw/opendata/getOD.ashx?id=1-5"
TDCC_URL_TEMPLATE = "https://smart.tdcc.com.tw/opendata/getOD.ashx?id=1-5&date={date}"

# =============================================================================
# 持股分級定義
# 說明：集保戶股權分散表就像一份「股票戶籍調查表」，
#       記錄了每一檔股票被哪些人持有、各自持有多少張。
#       TDCC 每週統計一次，把持有人按照持股張數分成15個等級，
#       就像把人依財富高低分成15個收入級距一樣。
#
#   級別1  = 1~999股（零股，不足1張）
#   級別2  = 1,000~5,000股（1~5張）
#   級別3  = 5,001~10,000股（5~10張）
#   ────────────────── 散戶與中戶分界線（10張）
#   級別4  = 10,001~15,000股（10~15張）
#   級別5  = 15,001~20,000股（15~20張）
#   級別6  = 20,001~30,000股（20~30張）
#   級別7  = 30,001~40,000股（30~40張）
#   級別8  = 40,001~50,000股（40~50張）
#   級別9  = 50,001~100,000股（50~100張）
#   級別10 = 100,001~200,000股（100~200張）
#   級別11 = 200,001~400,000股（200~400張）
#   級別12 = 400,001~600,000股（400~600張）
#   級別13 = 600,001~800,000股（600~800張）
#   級別14 = 800,001~1,000,000股（800~1,000張）
#   ────────────────── 中戶與大戶分界線（1,000張）
#   級別15 = 1,000,001股以上（1,000張以上）
# =============================================================================

# 大戶：持股1,000張以上（級別15）
# 就像百貨公司的頂級 VIP 會員，資金最雄厚，
# 他們的買賣動向往往決定整個市場的走勢方向
BIGSHOT_LEVELS = [15]

# 中戶：持股10張以上、未滿1,000張（級別4~14）
# 就像百貨公司的一般 VIP 會員，有一定資金規模，
# 對市場有影響力但不是主導者
MEDIUM_LEVELS = [4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]

# 散戶：持股未滿10張（級別1~3）
# 就像百貨公司的一般消費者，人數最多但個別影響力最小，
# 散戶大量追買往往是股價高點的警示訊號
RETAIL_LEVELS = [1, 2, 3]

# =============================================================================
# 其他設定常數
# =============================================================================
KEEP_WEEKS      = 13    # 歷史總表保留最近幾週（約三個月）
CONSEC_WEEKS    = 3     # 篩選連續幾週大戶持股比上漲
TIMEOUT_SECONDS = 30    # 網路請求逾時秒數

EXCEL_PATH = DATA_DIR / "TDCC_history.xlsx"

# HTTP 請求標頭，模擬一般瀏覽器行為，避免被伺服器拒絕
REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}


# =============================================================================
# 工具函式
# =============================================================================

def get_latest_saturday(reference_date: datetime = None) -> datetime:
    """
    計算指定日期（或今日）最近的週六日期。

    就像每週六出版的雜誌，不管你哪天去找，
    最新一期永遠是「上一個或當天的週六」出版的那期。

    Args:
        reference_date: 參考日期，預設為今日

    Returns:
        最近週六的 datetime 物件
    """
    if reference_date is None:
        reference_date = datetime.now()
    # 計算距離最近週六的天數（週六 = weekday 5）
    days_since_saturday = (reference_date.weekday() - 5) % 7
    return reference_date - timedelta(days=days_since_saturday)


def get_13_saturdays(reference_date: datetime = None) -> list:
    """
    從指定日期往回推算13個週六的日期清單（由舊到新）。

    就像翻閱一本13期的雜誌合訂本，
    從最舊一期排到最新一期。

    Args:
        reference_date: 參考日期，預設為今日

    Returns:
        list of datetime，共13個週六，由舊到新排列
    """
    latest = get_latest_saturday(reference_date)
    # 往前推12個週（共13個週六，包含本週）
    saturdays = [latest - timedelta(weeks=i) for i in range(KEEP_WEEKS - 1, -1, -1)]
    return saturdays


def ensure_dirs():
    """
    確保所有必要的資料夾都已存在，若不存在則自動建立。

    就像在開始整理資料之前，先確認書桌上的收納格都已準備好，
    缺少哪個格子就補上哪個，不需要手動一個個建立。
    """
    for d in [BASE_DIR, RAW_DIR, DATA_DIR, SUMMARY_DIR]:
        d.mkdir(parents=True, exist_ok=True)


def _to_roc_date(date: datetime) -> str:
    """
    將西曆日期轉換為民國年格式（TDCC API 日期參數所需格式）。

    TDCC open data API 的 date 參數採用民國年：
      西曆 2026-05-23 → 民國115年05月23日 → "1150523"
    """
    roc_year = date.year - 1911
    return f"{roc_year}{date.strftime('%m%d')}"


# =============================================================================
# 偵錯工具：TDCC API 歷史日期查詢能力探測
# =============================================================================

def probe_tdcc_historical_support() -> bool:
    """
    偵測 TDCC API 是否支援指定日期查詢歷史資料。

    就像在去報攤之前先打電話確認「你們有舊期雜誌嗎？」，
    確認好再出發，免得白跑一趟。

    這個函式會嘗試用一個「已知應該存在的歷史日期」
    對 TDCC API 發出請求，觀察回傳的資料是否與
    「不帶日期的最新資料」不同。
    若回傳資料明顯不同，代表 API 支援歷史查詢。
    若回傳資料完全相同，代表 API 忽略了日期參數，
    只會回傳最新資料。

    Args:
        無

    Returns:
        True  → API 支援歷史日期查詢
        False → API 不支援，每次只回傳最新資料
    """
    print("\n🔍 [偵錯] 探測 TDCC API 是否支援歷史日期查詢...")

    try:
        # 取得最新資料（不帶日期）
        resp_latest = requests.get(
            TDCC_URL,
            headers=REQUEST_HEADERS,
            timeout=TIMEOUT_SECONDS
        )
        resp_latest.raise_for_status()
        df_latest = _parse_tdcc_response(resp_latest.text)

        if df_latest is None or df_latest.empty:
            print("   ⚠️  無法取得最新資料，跳過 API 探測，預設為不支援歷史查詢")
            return False

        # 用「4週前的週六」測試歷史查詢（TDCC API 使用民國年格式）
        test_date = get_latest_saturday() - timedelta(weeks=4)
        test_date_roc  = _to_roc_date(test_date)
        test_date_greg = test_date.strftime("%Y%m%d")
        test_url = TDCC_URL_TEMPLATE.format(date=test_date_roc)
        print(f"   🔍 [偵錯] 以民國年格式測試歷史查詢：{test_date_roc}（西曆 {test_date_greg}）")
        print(f"      URL：{test_url}")

        resp_hist = requests.get(
            test_url,
            headers=REQUEST_HEADERS,
            timeout=TIMEOUT_SECONDS
        )
        resp_hist.raise_for_status()
        df_hist = _parse_tdcc_response(resp_hist.text)

        if df_hist is None or df_hist.empty:
            print("   ⚠️  歷史日期查詢回傳空資料，API 不支援歷史查詢")
            return False

        # 比較兩次回傳的「日期欄位」是否不同
        # TDCC 資料中通常有「資料日期」或「年度」欄位可供比較
        date_col = _detect_date_column(df_latest)
        if date_col:
            latest_dates = set(df_latest[date_col].astype(str).unique())
            hist_dates   = set(df_hist[date_col].astype(str).unique())

            if latest_dates != hist_dates:
                print(f"   ✅ [偵錯] TDCC API 支援歷史日期查詢")
                print(f"      最新資料日期：{latest_dates}")
                print(f"      歷史查詢日期：{hist_dates}")
                return True
            else:
                print(f"   ⚠️  [偵錯] API 回傳資料日期相同，可能不支援歷史查詢")
                print(f"      兩次查詢日期欄位值均為：{latest_dates}")
                return False
        else:
            # 無法從欄位判斷，改以資料筆數與內容雜湊比較
            hash_latest = pd.util.hash_pandas_object(df_latest).sum()
            hash_hist   = pd.util.hash_pandas_object(df_hist).sum()
            if hash_latest != hash_hist:
                print("   ✅ [偵錯] 兩次回傳資料內容不同，API 支援歷史查詢")
                return True
            else:
                print("   ⚠️  [偵錯] 兩次回傳資料完全相同，API 不支援歷史查詢")
                return False

    except requests.exceptions.Timeout:
        print("   ❌ [偵錯] API 探測逾時，預設為不支援歷史查詢")
        return False
    except Exception as e:
        print(f"   ❌ [偵錯] API 探測發生例外：{e}，預設為不支援歷史查詢")
        return False


def _detect_date_column(df: pd.DataFrame) -> str | None:
    """
    自動偵測 DataFrame 中代表「資料日期」的欄位名稱。

    就像在一堆標籤中找到寫有「日期」的那個，
    不同版本的 TDCC 資料可能用不同的欄位名稱。

    Args:
        df: 要偵測的 DataFrame

    Returns:
        日期欄位名稱，找不到則回傳 None
    """
    # 常見的日期欄位名稱關鍵字
    date_keywords = ["日期", "年度", "date", "week", "週"]
    for col in df.columns:
        for kw in date_keywords:
            if kw.lower() in col.lower():
                return col
    return None


def _parse_tdcc_response(text: str) -> pd.DataFrame | None:
    """
    解析 TDCC API 回傳的文字內容為 DataFrame。

    就像把剛收到的傳真紙上的表格，
    讀成電腦可以處理的表格格式。

    Args:
        text: API 回傳的原始文字（CSV 格式）

    Returns:
        解析後的 DataFrame，失敗則回傳 None
    """
    if not text or not text.strip():
        print("   🔍 [偵錯] API 回傳空字串")
        return None

    # 若回傳 HTML（錯誤頁或維護頁），直接判定失敗
    stripped = text.strip()
    if stripped.lower().startswith("<!") or stripped.lower().startswith("<html"):
        preview = stripped[:300].replace("\n", " ")
        print(f"   🔍 [偵錯] API 回傳 HTML（非 CSV）：{preview}")
        return None

    try:
        # TDCC 回傳的是 CSV 格式，直接用 pandas 讀取
        df = pd.read_csv(StringIO(text))
        if df.empty:
            print(f"   🔍 [偵錯] CSV 解析後為空，前300字元：{stripped[:300]}")
            return None
        return df
    except Exception as e:
        preview = stripped[:300].replace("\n", " ")
        print(f"   🔍 [偵錯] CSV 解析失敗（{e}），前300字元：{preview}")
        return None


# =============================================================================
# 功能1：下載本週（或指定週）TDCC CSV
# =============================================================================

def download_tdcc_csv(
    target_date: datetime = None,
    api_supports_history: bool = True
) -> tuple[pd.DataFrame, str]:
    """
    從 TDCC 官方網站下載指定週全市場股東持股調查報告。

    就像每週六去報攤買最新一期的財經雜誌，
    把本週全市場的股東持股調查報告下載回來存檔備用。
    如果這期雜誌已經買過了，就不需要再買一次。
    若需要補買舊期（歷史資料），
    會先確認報攤是否有舊期庫存（API 支援歷史查詢），
    沒有的話就略過並提醒使用者。

    Args:
        target_date: 目標週六日期，None 表示本週
        api_supports_history: API 是否支援歷史日期查詢

    Returns:
        tuple: (DataFrame 原始資料, 週六日期字串 YYYY-MM-DD)

    Raises:
        RuntimeError: 下載失敗或資料格式異常時拋出
    """
    # 計算目標週六日期
    if target_date is None:
        target_date = get_latest_saturday()

    week_date     = target_date.strftime("%Y-%m-%d")
    week_date_fmt = target_date.strftime("%Y%m%d")
    roc_date_str  = _to_roc_date(target_date)
    cache_path    = RAW_DIR / f"TDCC_{week_date_fmt}.csv"

    # 這個步驟就像去超市購物前先確認購物清單是否已經存在，
    # 若已經有了就不需要再重新整理一份，直接拿來用即可。
    if cache_path.exists():
        print(f"   📁 本週資料已存在，跳過下載：{cache_path.name}")
        try:
            df = pd.read_csv(cache_path, dtype=str)
            print(f"   ✅ 讀取快取完成：{cache_path.name}（共 {len(df):,} 筆資料）")
            return df, week_date
        except Exception as e:
            raise RuntimeError(f"讀取快取檔案失敗：{cache_path.name}\n原因：{e}")

    # 判斷是否為歷史日期查詢
    is_historical = (target_date.date() < get_latest_saturday().date())

    if is_historical and not api_supports_history:
        # 就像報攤告訴你「舊期雜誌已售完」，
        # 這週的資料無法取得，只能跳過
        print(f"   ⚠️  [偵錯] 跳過歷史日期 {week_date}：API 不支援歷史查詢")
        raise RuntimeError(
            f"API 不支援歷史查詢，無法下載 {week_date} 的資料。\n"
            "建議：每週執行一次，逐週累積資料。"
        )

    # 決定使用哪個 URL
    # 歷史資料使用民國年格式（TDCC API 正確的 date 參數格式）
    if is_historical:
        url = TDCC_URL_TEMPLATE.format(date=roc_date_str)
        print(f"   📡 下載歷史資料：{week_date}（民國年 {roc_date_str}）")
        print(f"      URL：{url}")
    else:
        url = TDCC_URL
        print(f"   📡 下載本週資料：{week_date}（{url}）")

    # 嘗試下載，最多重試 3 次（每次間隔 5 秒，避免限流）
    # 若本週無日期參數的 URL 失敗，也嘗試帶民國年格式的 URL
    urls_to_try = [url]
    if not is_historical:
        # 本週也試帶民國年格式的 URL，作為後備
        urls_to_try.append(TDCC_URL_TEMPLATE.format(date=roc_date_str))

    df = None
    last_error = None
    for attempt_url in urls_to_try:
        for retry in range(3):
            if retry > 0:
                wait = 5 * retry
                print(f"   ⏳ [重試 {retry}/2] 等待 {wait} 秒後重試...")
                time.sleep(wait)
            try:
                print(f"   📡 請求：{attempt_url}")
                resp = requests.get(
                    attempt_url, headers=REQUEST_HEADERS,
                    timeout=TIMEOUT_SECONDS
                )
                if resp.status_code != 200:
                    last_error = f"HTTP {resp.status_code}"
                    print(f"   ⚠️  HTTP {resp.status_code}，繼續重試")
                    continue
                df = _parse_tdcc_response(resp.text)
                if df is not None and not df.empty:
                    break  # 成功
                last_error = "資料為空"
            except requests.exceptions.Timeout:
                last_error = f"逾時（超過 {TIMEOUT_SECONDS} 秒）"
                print(f"   ⚠️  請求逾時，繼續重試")
            except requests.exceptions.ConnectionError as e:
                last_error = f"連線失敗：{e}"
                print(f"   ⚠️  連線失敗，繼續重試")
        if df is not None and not df.empty:
            break
        print(f"   ⚠️  URL {attempt_url} 所有重試均失敗，嘗試下一組 URL")
        time.sleep(3)

    if df is None or df.empty:
        raise RuntimeError(
            f"下載失敗（{last_error}），可能是該日期（{week_date}）無資料\n"
            "建議動作：確認該週六是否為 TDCC 公告日\n"
            "建議動作：請確認網路連線後重新執行"
        )

    # 偵錯：確認實際回傳的資料日期是否符合預期
    # 若民國年格式回傳錯誤日期，嘗試改用西曆格式作為後備
    date_col = _detect_date_column(df)
    if date_col and is_historical:
        actual_dates = df[date_col].astype(str).unique()
        print(f"   🔍 [偵錯] 回傳資料日期欄位值：{actual_dates[:3]}...")

        # 如果回傳的日期欄位包含民國年格式，確認是否為預期週別
        # 若回傳資料明顯與預期不符（日期相同或只有最新週），嘗試西曆格式
        actual_set = set(str(d).strip() for d in actual_dates)
        if roc_date_str not in actual_set and week_date_fmt not in actual_set:
            print(f"   ⚠️  [偵錯] 民國年格式未命中，嘗試西曆格式：{week_date_fmt}")
            fallback_url = TDCC_URL_TEMPLATE.format(date=week_date_fmt)
            try:
                fb_resp = requests.get(
                    fallback_url, headers=REQUEST_HEADERS,
                    timeout=TIMEOUT_SECONDS
                )
                if fb_resp.status_code == 200:
                    fb_df = _parse_tdcc_response(fb_resp.text)
                    if fb_df is not None and not fb_df.empty:
                        fb_dates = set(fb_df[date_col].astype(str).str.strip().unique())
                        if fb_dates != actual_set:
                            print(f"   ✅ [偵錯] 西曆格式回傳不同日期資料，改用此版本")
                            df = fb_df
            except Exception as fb_e:
                print(f"   ⚠️  [偵錯] 西曆格式後備查詢失敗：{fb_e}")

    # 儲存至快取
    try:
        df.to_csv(cache_path, index=False, encoding="utf-8-sig")
    except Exception as e:
        raise RuntimeError(f"儲存快取檔案失敗：{cache_path.name}\n原因：{e}")

    print(f"   ✅ 下載完成：{cache_path.name}（共 {len(df):,} 筆資料）")
    return df, week_date


# =============================================================================
# 功能2：清洗與整理全市場資料
# =============================================================================

def clean_market_data(df: pd.DataFrame, week_date: str) -> pd.DataFrame:
    """
    清除無效資料、統一欄位名稱、新增週別標籤。

    就像把剛買回來的食材清洗整理，
    去除不能吃的部分（無效資料），
    把食材分門別類放好（欄位標準化），
    貼上這週的日期標籤，
    讓後續的烹飪（計算）可以順利進行。

    Args:
        df:        功能1回傳的原始 DataFrame
        week_date: 週六日期字串（YYYY-MM-DD）

    Returns:
        清洗後的 DataFrame

    Raises:
        RuntimeError: 必要欄位不存在或資料清洗後為空時拋出
    """
    try:
        original_count = len(df)

        # 步驟一：欄位名稱標準化
        # 說明：TDCC 有時會用「佔」有時用「占」，
        #       就像同一個字有兩種寫法，需統一成同一個名稱
        df.columns = df.columns.str.strip()
        df = df.rename(columns={"佔集保庫存數比例%": "占集保庫存數比例%"})

        # 偵錯：印出實際欄位名稱，協助排查欄位對應問題
        print(f"   🔍 [偵錯] 資料欄位：{list(df.columns)}")

        # 步驟二：去除公債資料
        # 說明：代號開頭為 Y 的是公債不是股票，
        #       就像超市裡混入了非賣品，需要先挑出來
        stock_col = _find_column(df, ["證券代號", "stock_code", "代號"])
        if stock_col is None:
            raise RuntimeError(
                "找不到「證券代號」欄位，資料格式可能已變更\n"
                f"現有欄位：{list(df.columns)}"
            )
        df = df[~df[stock_col].astype(str).str.startswith("Y")].copy()

        # 步驟三：去除無效分級
        # 說明：持股分級=16是系統預設的空值標記，
        #       不代表任何真實的持股級距，需要去除
        level_col = _find_column(df, ["持股分級", "level", "分級"])
        if level_col is None:
            raise RuntimeError(
                "找不到「持股分級」欄位，資料格式可能已變更\n"
                f"現有欄位：{list(df.columns)}"
            )
        df[level_col] = pd.to_numeric(df[level_col], errors="coerce")
        df = df[df[level_col] != 16].copy()

        # 步驟四：數值型別轉換
        # 說明：確保人數、股數、比例欄位是數字格式，
        #       避免後續計算時出現「字串無法相加」的錯誤
        num_col_人數 = _find_column(df, ["人數"])
        num_col_股數 = _find_column(df, ["股數"])
        num_col_比例 = _find_column(df, ["占集保庫存數比例%", "佔集保庫存數比例%"])

        for col in [num_col_人數, num_col_股數, num_col_比例]:
            if col:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        # 步驟五：新增週別標籤
        # 說明：在每一筆資料加上這週的日期，
        #       就像在每張收據蓋上日期章，方便日後查閱
        df["週別"] = week_date

        # 標準化欄位名稱（供後續使用）
        rename_map = {}
        if stock_col != "證券代號":
            rename_map[stock_col] = "證券代號"
        name_col = _find_column(df, ["證券名稱", "stock_name", "名稱"])
        if name_col and name_col != "證券名稱":
            rename_map[name_col] = "證券名稱"
        if level_col != "持股分級":
            rename_map[level_col] = "持股分級"
        if num_col_人數 and num_col_人數 != "人數":
            rename_map[num_col_人數] = "人數"
        if num_col_股數 and num_col_股數 != "股數":
            rename_map[num_col_股數] = "股數"
        if num_col_比例 and num_col_比例 != "占集保庫存數比例%":
            rename_map[num_col_比例] = "占集保庫存數比例%"
        if rename_map:
            df = df.rename(columns=rename_map)

        if df.empty:
            raise RuntimeError("清洗後資料為空，請確認原始資料格式")

        stock_count = df["證券代號"].nunique()
        print(
            f"   ✅ 資料清洗完成：共 {stock_count:,} 檔股票，"
            f"{len(df):,} 筆分級資料"
            f"（原始 {original_count:,} 筆，去除 {original_count - len(df):,} 筆無效資料）"
        )
        return df

    except RuntimeError:
        raise
    except Exception as e:
        raise RuntimeError(f"資料清洗時發生未預期錯誤：{e}")


def _find_column(df: pd.DataFrame, candidates: list) -> str | None:
    """
    在 DataFrame 欄位中尋找候選名稱，回傳第一個找到的欄位名稱。

    就像找一個人，不管他叫「志明」還是「John」還是「阿明」，
    只要其中一個名字對上了就找到了。

    Args:
        df:         要搜尋的 DataFrame
        candidates: 候選欄位名稱清單

    Returns:
        找到的欄位名稱，或 None
    """
    for name in candidates:
        if name in df.columns:
            return name
        # 模糊比對：欄位名稱包含關鍵字
        for col in df.columns:
            if name in col:
                return col
    return None


# =============================================================================
# 功能3：計算全市場大戶/中戶/散戶持股比
# =============================================================================

def calculate_holder_ratio(df: pd.DataFrame) -> pd.DataFrame:
    """
    依照三種投資人定義，計算每支股票的持股比例彙總。

    就像幫每一家公司的股東做財富分級統計，
    計算出富豪群（大戶）、中產階級（中戶）、
    一般民眾（散戶）各自持有這家公司多少比例的股份。
    數字越高代表那個群體對這支股票的控制力越強。

    Args:
        df: 功能2回傳的清洗後 DataFrame

    Returns:
        每支股票的持股比彙總 DataFrame（每行為一支股票）

    Raises:
        RuntimeError: 必要欄位缺失或計算異常時拋出
    """
    try:
        results = []

        # 偵錯：確認計算前的欄位與資料筆數
        print(f"   🔍 [偵錯] 開始計算持股比，資料筆數：{len(df):,}")

        week_date = df["週別"].iloc[0]

        # 依股票代號分組後，對每支股票執行以下計算：
        # 1. 篩選出屬於大戶級別的資料列，加總其持股比例
        # 2. 篩選出屬於中戶級別的資料列，加總其持股比例
        # 3. 篩選出屬於散戶級別的資料列，加總其持股比例
        # 三者合計應接近100%
        # （差異來自零股待領回等特殊情況，屬正常現象）
        for stock_id, group in df.groupby("證券代號"):
            stock_name = group["證券名稱"].iloc[0] if "證券名稱" in group.columns else ""

            # 依持股分級篩選各群體資料
            big_data    = group[group["持股分級"].isin(BIGSHOT_LEVELS)]
            medium_data = group[group["持股分級"].isin(MEDIUM_LEVELS)]
            retail_data = group[group["持股分級"].isin(RETAIL_LEVELS)]
            all_data    = group[group["持股分級"].isin(BIGSHOT_LEVELS + MEDIUM_LEVELS + RETAIL_LEVELS)]

            results.append({
                "週別":        week_date,
                "股票代號":    stock_id,
                "股票名稱":    stock_name,
                "大戶持股比%": round(big_data["占集保庫存數比例%"].sum(), 4),
                "中戶持股比%": round(medium_data["占集保庫存數比例%"].sum(), 4),
                "散戶持股比%": round(retail_data["占集保庫存數比例%"].sum(), 4),
                "大戶人數":    int(big_data["人數"].sum()),
                "中戶人數":    int(medium_data["人數"].sum()),
                "散戶人數":    int(retail_data["人數"].sum()),
                "總股東人數":  int(all_data["人數"].sum()),
            })

        ratio_df = pd.DataFrame(results)

        # 偵錯：抽樣檢查計算結果
        if not ratio_df.empty:
            sample = ratio_df.iloc[0]
            total_pct = sample["大戶持股比%"] + sample["中戶持股比%"] + sample["散戶持股比%"]
            print(
                f"   🔍 [偵錯] 抽樣檢查（{sample['股票代號']} {sample['股票名稱']}）："
                f"大戶{sample['大戶持股比%']}% + 中戶{sample['中戶持股比%']}% + "
                f"散戶{sample['散戶持股比%']}% = {total_pct:.2f}%"
            )
            if total_pct < 90:
                print(
                    f"   ⚠️  [偵錯] 三類合計僅 {total_pct:.2f}%，"
                    "可能有未分類級別或資料異常，請確認"
                )

        print(f"   ✅ 持股比計算完成：共計算 {len(ratio_df):,} 檔股票")
        return ratio_df

    except Exception as e:
        raise RuntimeError(f"持股比計算時發生錯誤：{e}")


# =============================================================================
# 功能4：累積寫入歷史總表 Excel
# =============================================================================

def update_excel_history(ratio_df: pd.DataFrame, clean_df: pd.DataFrame):
    """
    將本週計算結果寫入歷史 Excel，並自動維護13週滾動視窗。

    就像維護一本帳簿，每週把新的一頁（本週資料）貼進去，
    同時把超過三個月的舊頁（13週前的資料）撕掉，
    讓帳簿永遠只保留最近13週的紀錄，不會越來越厚。

    Args:
        ratio_df: 功能3計算出的持股比彙總 DataFrame
        clean_df: 功能2清洗後的完整分級明細 DataFrame

    Returns:
        無（直接寫入 Excel 檔案）

    Raises:
        RuntimeError: Excel 讀寫失敗時拋出
    """
    try:
        current_week = ratio_df["週別"].iloc[0]

        # 步驟一：檢查是否為第一次執行
        # 說明：就像新買了一本帳簿，第一次使用時需要先畫好格線
        if not EXCEL_PATH.exists():
            print(f"   📖 首次執行，建立新的歷史總表：{EXCEL_PATH.name}")
            existing_ratio = pd.DataFrame()
            existing_clean = pd.DataFrame()
        else:
            # 步驟二：讀取現有資料
            # 說明：把帳簿現有的內容讀進來，準備加入新的一頁
            try:
                existing_ratio = pd.read_excel(
                    EXCEL_PATH, sheet_name="每週持股比", dtype={"股票代號": str}
                )
                existing_clean = pd.read_excel(
                    EXCEL_PATH, sheet_name="原始分級資料", dtype={"證券代號": str}
                )
                print(
                    f"   📖 讀取現有帳簿：{existing_ratio['週別'].nunique()} 週資料，"
                    f"{len(existing_ratio):,} 筆彙總，{len(existing_clean):,} 筆明細"
                )
            except Exception as e:
                print(f"   ⚠️  [偵錯] 讀取現有 Excel 發生問題：{e}")
                print("         將建立全新歷史總表")
                existing_ratio = pd.DataFrame()
                existing_clean = pd.DataFrame()

        # 步驟三：檢查本週是否已寫入
        # 說明：避免重複執行時把同一週的資料寫入兩次，
        #       就像避免同一筆帳重複記帳
        if not existing_ratio.empty and "週別" in existing_ratio.columns:
            if current_week in existing_ratio["週別"].values:
                print(f"   ⚠️  本週（{current_week}）資料已存在，跳過寫入")
                return

        # 步驟四：合併新舊資料
        # 說明：把本週新資料加到帳簿現有內容的後面
        combined_ratio = pd.concat([existing_ratio, ratio_df], ignore_index=True)
        combined_clean = pd.concat([existing_clean, clean_df], ignore_index=True)

        # 步驟五：保留最近13週，刪除舊資料
        # 說明：找出所有出現過的週別，只保留最新的13個，
        #       就像帳簿只保留最近13頁，舊的自動撕掉
        all_weeks = sorted(combined_ratio["週別"].unique())
        keep_weeks = all_weeks[-KEEP_WEEKS:]
        removed_weeks = [w for w in all_weeks if w not in keep_weeks]

        if removed_weeks:
            print(f"   🗑️  移除超過13週的舊資料：{removed_weeks}")

        combined_ratio = combined_ratio[combined_ratio["週別"].isin(keep_weeks)].copy()
        combined_clean = combined_clean[combined_clean["週別"].isin(keep_weeks)].copy()

        # 步驟六：寫回 Excel
        # 說明：把整理好的帳簿重新存檔
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
            combined_ratio.to_excel(writer, sheet_name="每週持股比", index=False)
            combined_clean.to_excel(writer, sheet_name="原始分級資料", index=False)

        print(
            f"   ✅ Excel 更新完成：共保留 {len(keep_weeks)} 週資料，"
            f"{len(combined_ratio):,} 筆彙總，{len(combined_clean):,} 筆分級明細"
        )

    except RuntimeError:
        raise
    except Exception as e:
        raise RuntimeError(f"寫入 Excel 時發生錯誤：{e}")


# =============================================================================
# 功能5：篩選連續三週大戶持股比上漲的股票
# =============================================================================

def filter_consecutive_bigshot_increase() -> pd.DataFrame:
    """
    從歷史總表中找出連續三週大戶持股比例遞增的股票。

    就像在全班同學中找出連續三週段考成績都進步的學生，
    不是只看最新一次的成績，而是要確認有持續進步的趨勢。
    這樣才能判斷大戶（主力）是真的在持續買進，
    而非偶然的短暫波動。

    Args:
        無（從 EXCEL_PATH 讀取）

    Returns:
        符合條件的股票 DataFrame，包含以下欄位：
        股票代號、股票名稱、
        第1週週別、第1週大戶持股比%、
        第2週週別、第2週大戶持股比%、
        第3週週別、第3週大戶持股比%、
        三週總增幅%（第三週 - 第一週，由大到小排序）

    Raises:
        RuntimeError: Excel 讀取失敗時拋出
    """
    try:
        # 步驟一：讀取歷史總表最近三週資料
        # 說明：只需要最近三週的資料就能判斷是否連續上漲
        if not EXCEL_PATH.exists():
            raise RuntimeError(f"歷史總表不存在：{EXCEL_PATH}")

        ratio_df = pd.read_excel(
            EXCEL_PATH, sheet_name="每週持股比", dtype={"股票代號": str}
        )

        all_weeks = sorted(ratio_df["週別"].unique())

        if len(all_weeks) < CONSEC_WEEKS:
            print(
                f"   ⚠️  歷史資料不足 {CONSEC_WEEKS} 週（目前僅 {len(all_weeks)} 週），"
                "無法篩選連續上漲條件"
            )
            return pd.DataFrame()

        # 取最新三週
        w1, w2, w3 = all_weeks[-3], all_weeks[-2], all_weeks[-1]
        print(f"   🔍 [偵錯] 篩選依據的三週：{w1} → {w2} → {w3}")

        df_w1 = ratio_df[ratio_df["週別"] == w1][["股票代號", "股票名稱", "大戶持股比%"]].copy()
        df_w2 = ratio_df[ratio_df["週別"] == w2][["股票代號", "大戶持股比%"]].copy()
        df_w3 = ratio_df[ratio_df["週別"] == w3][["股票代號", "大戶持股比%"]].copy()

        # 合併三週資料
        merged = df_w1.merge(df_w2, on="股票代號", suffixes=("_w1", "_w2"))
        merged = merged.merge(df_w3, on="股票代號")
        merged = merged.rename(columns={"大戶持股比%": "大戶持股比%_w3"})

        # 步驟二：對每支股票判斷連續上漲條件
        # 說明：逐一檢查每支股票在三個週別的大戶持股比，
        #       就像逐一翻開每位學生的三次成績單比較
        # 嚴格遞增條件：第1週 < 第2週 < 第3週
        cond = (
            (merged["大戶持股比%_w1"] < merged["大戶持股比%_w2"]) &
            (merged["大戶持股比%_w2"] < merged["大戶持股比%_w3"])
        )
        filtered = merged[cond].copy()

        # 計算三週總增幅
        filtered["三週總增幅%"] = (
            filtered["大戶持股比%_w3"] - filtered["大戶持股比%_w1"]
        ).round(4)

        # 步驟三：整理結果並排序
        # 說明：把符合條件的股票整理成清單，增幅最大的排在最前面
        result = filtered[[
            "股票代號", "股票名稱",
            "大戶持股比%_w1", "大戶持股比%_w2", "大戶持股比%_w3",
            "三週總增幅%"
        ]].copy()
        result.insert(2, "第1週週別", w1)
        result.insert(4, "第2週週別", w2)
        result.insert(6, "第3週週別", w3)
        result = result.rename(columns={
            "大戶持股比%_w1": "第1週大戶持股比%",
            "大戶持股比%_w2": "第2週大戶持股比%",
            "大戶持股比%_w3": "第3週大戶持股比%",
        })
        result = result.sort_values("三週總增幅%", ascending=False).reset_index(drop=True)

        total_stocks = ratio_df[ratio_df["週別"] == w3]["股票代號"].nunique()
        print(
            f"   ✅ 篩選完成：全市場 {total_stocks:,} 檔股票中，"
            f"共 {len(result):,} 檔符合連續三週大戶上漲條件"
        )
        return result

    except RuntimeError:
        raise
    except Exception as e:
        raise RuntimeError(f"篩選連續大戶上漲時發生錯誤：{e}")


# =============================================================================
# 功能6：輸出 NotebookLM 分析摘要文字檔
# =============================================================================

def export_notebooklm_summary(
    filtered_df: pd.DataFrame,
    week_date: str,
    total_stock_count: int
):
    """
    將篩選結果與歷史趨勢整理成 NotebookLM 可直接上傳的文字報告。

    就像幫老師整理出一份重點學習摘要，
    只把符合條件的學生（大戶持續增持的股票）的學習歷程，
    整理成一份清晰的報告。
    上傳到 NotebookLM 後就可以直接提問分析，
    不需要自己再整理資料。

    Args:
        filtered_df:       功能5篩選出的符合條件股票 DataFrame
        week_date:         本週週六日期字串（YYYY-MM-DD）
        total_stock_count: 全市場處理的股票總數

    Returns:
        無（直接寫出文字檔）

    Raises:
        RuntimeError: 讀取 Excel 或寫出文字檔失敗時拋出
    """
    try:
        week_date_fmt = week_date.replace("-", "")
        output_path   = SUMMARY_DIR / f"notebooklm_summary_{week_date_fmt}.txt"

        # 讀取13週詳細資料（供各股趨勢輸出使用）
        if EXCEL_PATH.exists():
            ratio_history = pd.read_excel(
                EXCEL_PATH, sheet_name="每週持股比", dtype={"股票代號": str}
            )
        else:
            ratio_history = pd.DataFrame()

        lines = []

        # ── 報告標頭 ──────────────────────────────────────────────────
        lines.append("=" * 40)
        lines.append(f"# 台股大戶持股比週報 {week_date}")
        lines.append("")
        lines.append("## 報告說明")
        lines.append("資料來源：臺灣集中保管結算所（TDCC）集保戶股權分散表")
        lines.append("更新頻率：每週日自動執行")
        lines.append("大戶定義：持股1,000張以上（第15級）")
        lines.append("中戶定義：持股10張以上、未滿1,000張（第4~14級）")
        lines.append("散戶定義：持股未滿10張（第1~3級）")
        lines.append("篩選條件：近三週大戶持股比連續上漲")
        lines.append(f"本週資料基準日：{week_date}（週六）")
        lines.append(f"全市場股票總數：{total_stock_count:,} 檔")
        lines.append(f"本週符合條件股票數：{len(filtered_df):,} 檔")
        lines.append("")

        # ── 符合條件股票總覽表 ────────────────────────────────────────
        lines.append("## 符合條件股票總覽")
        lines.append("（依三週總增幅由大到小排列）")
        lines.append("")

        if filtered_df.empty:
            lines.append("本週無符合條件之股票。")
        else:
            # 表格標題行
            lines.append(
                "| 排名 | 股票代號 | 股票名稱 | "
                "第1週大戶% | 第2週大戶% | 第3週大戶% | 三週增幅% |"
            )
            lines.append(
                "|------|---------|---------|"
                "-----------|-----------|-----------|---------|"
            )
            for i, row in filtered_df.iterrows():
                lines.append(
                    f"| {i+1:4d} | {row['股票代號']:8s} | {row['股票名稱']:8s} | "
                    f"{row['第1週大戶持股比%']:9.2f}% | "
                    f"{row['第2週大戶持股比%']:9.2f}% | "
                    f"{row['第3週大戶持股比%']:9.2f}% | "
                    f"{row['三週總增幅%']:7.2f}% |"
                )

        lines.append("")
        lines.append("=" * 40)

        # ── 各股13週詳細趨勢 ─────────────────────────────────────────
        for _, row in filtered_df.iterrows():
            stock_id   = row["股票代號"]
            stock_name = row["股票名稱"]

            lines.append("")
            lines.append(f"## [{stock_id} {stock_name}] 近13週詳細趨勢")
            lines.append("")
            lines.append(
                "| 週別       | 大戶持股比% | 大戶週變化% | "
                "中戶持股比% | 散戶持股比% | 大戶人數 | 總股東人數 |"
            )
            lines.append(
                "|-----------|-----------|-----------|"
                "-----------|-----------|---------|----------|"
            )

            if not ratio_history.empty:
                stock_hist = (
                    ratio_history[ratio_history["股票代號"] == stock_id]
                    .sort_values("週別")
                    .reset_index(drop=True)
                )

                prev_big = None
                max_big  = stock_hist["大戶持股比%"].max()
                min_big  = stock_hist["大戶持股比%"].min()
                max_week = stock_hist.loc[stock_hist["大戶持股比%"].idxmax(), "週別"]
                min_week = stock_hist.loc[stock_hist["大戶持股比%"].idxmin(), "週別"]

                for _, h in stock_hist.iterrows():
                    if prev_big is None:
                        change_str = "   -   "
                    else:
                        chg = h["大戶持股比%"] - prev_big
                        change_str = f"{chg:+.2f}%"
                    prev_big = h["大戶持股比%"]

                    lines.append(
                        f"| {h['週別']} | {h['大戶持股比%']:9.2f}% | "
                        f"{change_str:>9s} | "
                        f"{h['中戶持股比%']:9.2f}% | "
                        f"{h['散戶持股比%']:9.2f}% | "
                        f"{int(h['大戶人數']):7,d} | "
                        f"{int(h['總股東人數']):9,d} |"
                    )

                # 趨勢摘要
                latest = stock_hist.iloc[-1]
                prev   = stock_hist.iloc[-2] if len(stock_hist) >= 2 else latest
                latest_change = latest["大戶持股比%"] - prev["大戶持股比%"]

                lines.append("")
                lines.append("趨勢摘要：")
                lines.append(f"- 13週最高大戶持股比：{max_big:.2f}%（{max_week}）")
                lines.append(f"- 13週最低大戶持股比：{min_big:.2f}%（{min_week}）")
                lines.append("- 近三週大戶持股比方向：持續上升 ✅")
                lines.append(f"- 最新週大戶週變化：{latest_change:+.2f}%")
                lines.append(f"- 最新週大戶人數：{int(latest['大戶人數']):,} 人")
                lines.append(f"- 最新週總股東人數：{int(latest['總股東人數']):,} 人")
            else:
                lines.append("（歷史資料不足，無法顯示趨勢）")

            lines.append("")
            lines.append("=" * 40)

        # 寫出文字檔
        output_path.write_text("\n".join(lines), encoding="utf-8")
        print(f"   ✅ NotebookLM 摘要已輸出：{output_path.name}")

    except RuntimeError:
        raise
    except Exception as e:
        raise RuntimeError(f"輸出 NotebookLM 摘要時發生錯誤：{e}")


# =============================================================================
# 補足歷史資料（首次執行或資料不足時）
# =============================================================================

def backfill_history(api_supports_history: bool):
    """
    偵測歷史資料缺口，並自動補足至13週。

    就像新進一間公司接手帳本時，
    先清點帳本裡已有幾週的紀錄，
    再把缺少的那幾週一頁一頁補上，
    確保帳本從一開始就是完整的。

    若 API 不支援歷史查詢，
    會友善地告知使用者哪幾週需要手動補充，
    而不是讓程式崩潰。

    Args:
        api_supports_history: API 是否支援歷史日期查詢

    Returns:
        無
    """
    # 計算應有的13個週六清單
    target_saturdays = get_13_saturdays()
    target_dates     = [d.strftime("%Y-%m-%d") for d in target_saturdays]

    # 檢查 Excel 現有週數
    existing_weeks = []
    if EXCEL_PATH.exists():
        try:
            existing_df    = pd.read_excel(EXCEL_PATH, sheet_name="每週持股比")
            existing_weeks = list(existing_df["週別"].unique()) if "週別" in existing_df.columns else []
        except Exception:
            existing_weeks = []

    # 找出缺少的週別
    missing_dates = [d for d in target_dates if d not in existing_weeks]

    if not missing_dates:
        print("   ✅ 歷史資料已完整（13週），無需補足")
        return

    print(f"\n📋 歷史資料補足模式：需補足 {len(missing_dates)} 週")
    for d in missing_dates:
        print(f"   → {d}")

    if not api_supports_history:
        # probe 偵測為不支援，但仍嘗試下載（probe 可能因日期格式問題誤判）
        print(
            "\n   ⚠️  [偵錯] probe 顯示 API 可能不支援歷史查詢，"
            "但仍嘗試使用民國年格式補足歷史資料..."
        )

    # 依序補足缺少的週別（無論 probe 結果，皆嘗試；個別失敗時跳過）
    print(f"\n   開始逐週補足歷史資料...")
    succeeded = 0
    for i, date_str in enumerate(missing_dates, 1):
        # 跳過本週（由主流程處理）
        latest_sat = get_latest_saturday().strftime("%Y-%m-%d")
        if date_str == latest_sat:
            print(f"   [{i}/{len(missing_dates)}] {date_str} 為本週，由主流程處理")
            continue

        print(f"\n   [{i}/{len(missing_dates)}] 補足歷史資料：{date_str}")
        target_dt = datetime.strptime(date_str, "%Y-%m-%d")
        # 每次請求前等待，避免對 TDCC 伺服器造成過快的連續請求
        if i > 1:
            time.sleep(3)
        try:
            df, week_date = download_tdcc_csv(
                target_date=target_dt,
                api_supports_history=True   # 強制嘗試，不受 probe 結果影響
            )
            clean_df  = clean_market_data(df, week_date)
            ratio_df  = calculate_holder_ratio(clean_df)
            update_excel_history(ratio_df, clean_df)
            succeeded += 1
        except RuntimeError as e:
            print(f"   ⚠️  補足 {date_str} 失敗，跳過此週：{e}")
            continue

    print(f"\n   ✅ 歷史補足完成：成功補足 {succeeded}/{len(missing_dates)} 週")


# =============================================================================
# 主程式
# =============================================================================

def main():
    """
    主程式：依序執行六個功能，完成完整的 TDCC 週報流程。

    整個程式就像一條自動化的生產線：
      步驟1（下載）→ 取得原料
      步驟2（清洗）→ 清洗整理原料
      步驟3（計算）→ 加工製作
      步驟4（寫入）→ 入庫存檔
      步驟5（篩選）→ 品質把關
      步驟6（輸出）→ 包裝出貨
    """
    start_time = datetime.now()

    print("=" * 40)
    print("  TDCC 大戶持股比週報 開始執行")
    print(f"  執行時間：{start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 40)

    # 確保所有資料夾都已建立
    ensure_dirs()

    # ── 偵錯：探測 TDCC API 歷史查詢能力 ─────────────────────────
    api_supports_history = probe_tdcc_historical_support()

    # ── 補足歷史資料（首次執行或資料不足時）──────────────────────
    print("\n[補足檢查] 確認歷史資料完整性...")
    backfill_history(api_supports_history)

    # 補足結束後稍作等待，讓伺服器從連續請求中恢復
    print("\n   ⏳ 補足完成，等待 5 秒後繼續主流程...")
    time.sleep(5)

    # ── 主流程：本週資料 ──────────────────────────────────────────
    week_date         = None
    filtered_df       = pd.DataFrame()
    total_stock_count = 0
    current_weeks     = 0
    failed_step       = None
    error_message     = None

    try:
        # 步驟1：下載本週資料
        print("\n[步驟1] 下載本週 TDCC 資料...")
        df, week_date = download_tdcc_csv(api_supports_history=api_supports_history)

    except RuntimeError as e:
        failed_step   = "步驟1：下載 TDCC 資料"
        error_message = str(e)

    if not failed_step:
        try:
            # 步驟2：清洗資料
            print("\n[步驟2] 清洗整理全市場資料...")
            clean_df          = clean_market_data(df, week_date)
            total_stock_count = clean_df["證券代號"].nunique()

        except RuntimeError as e:
            failed_step   = "步驟2：清洗整理資料"
            error_message = str(e)

    if not failed_step:
        try:
            # 步驟3：計算持股比
            print("\n[步驟3] 計算全市場大戶/中戶/散戶持股比...")
            ratio_df = calculate_holder_ratio(clean_df)

        except RuntimeError as e:
            failed_step   = "步驟3：計算持股比"
            error_message = str(e)

    if not failed_step:
        try:
            # 步驟4：寫入歷史 Excel
            print("\n[步驟4] 更新歷史總表 Excel...")
            update_excel_history(ratio_df, clean_df)

        except RuntimeError as e:
            failed_step   = "步驟4：更新歷史 Excel"
            error_message = str(e)

    if not failed_step:
        try:
            # 步驟5：篩選連續三週大戶上漲
            print("\n[步驟5] 篩選連續三週大戶持股比上漲股票...")
            filtered_df = filter_consecutive_bigshot_increase()

        except RuntimeError as e:
            failed_step   = "步驟5：篩選連續大戶上漲"
            error_message = str(e)

    if not failed_step:
        try:
            # 步驟6：輸出 NotebookLM 摘要
            print("\n[步驟6] 輸出 NotebookLM 分析摘要...")
            export_notebooklm_summary(filtered_df, week_date, total_stock_count)

        except RuntimeError as e:
            failed_step   = "步驟6：輸出 NotebookLM 摘要"
            error_message = str(e)

    # 計算歷史累積週數
    if EXCEL_PATH.exists():
        try:
            hist = pd.read_excel(EXCEL_PATH, sheet_name="每週持股比")
            current_weeks = hist["週別"].nunique()
        except Exception:
            current_weeks = 0

    end_time     = datetime.now()
    week_date_fmt = week_date.replace("-", "") if week_date else "??????"

    # ── 輸出執行結果 ──────────────────────────────────────────────
    print("")
    if not failed_step:
        print("✅ TDCC 週報執行完成")
        print("─" * 24)
        print(f"執行時間：{end_time.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"本週資料基準日：{week_date}（週六）")
        print(f"全市場處理股票數：{total_stock_count:,} 檔")
        print(
            f"符合條件股票數：{len(filtered_df):,} 檔"
            f"（連續{CONSEC_WEEKS}週大戶持股比上漲）"
        )
        print(f"歷史累積週數：{current_weeks} 週（上限{KEEP_WEEKS}週）")
        print(f"輸出檔案：notebooklm_summary_{week_date_fmt}.txt")
        print(f"檔案位置：{SUMMARY_DIR}/")
        print("─" * 24)
    else:
        print("❌ 執行失敗")
        print("─" * 24)
        print(f"失敗步驟：{failed_step}")
        print(f"錯誤原因：{error_message}")
        print("建議動作：請確認網路連線後重新執行")
        print("─" * 24)
        sys.exit(1)


if __name__ == "__main__":
    main()
