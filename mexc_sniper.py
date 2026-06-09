#!/usr/bin/env python3
"""
MEXC 新規上場スナイパーボット
新しく上場した通貨ペアを検出して即座に買い注文を発注する
"""

import hashlib
import hmac
import json
import os
import time
import urllib.parse
import urllib.request
from datetime import datetime

# ─── 設定 ─────────────────────────────────────────────────────────────────────
API_KEY    = os.environ.get("MEXC_API_KEY", "")
API_SECRET = os.environ.get("MEXC_API_SECRET", "")

# 購入に使うクォート通貨 (USDT など)
QUOTE_CURRENCY = "USDT"

# 1回の購入金額 (USDT)
BUY_AMOUNT_USDT = 10.0

# 監視間隔 (秒)。短すぎるとレート制限に引っかかる
POLL_INTERVAL_SEC = 2

# 上場検出後に注文を出すまでの待機秒数 (0 = 即時)
ORDER_DELAY_SEC = 0

# ドライラン: True なら実際に注文は出さずログだけ出力
DRY_RUN = os.environ.get("DRY_RUN", "true").lower() != "false"

BASE_URL = "https://api.mexc.com"

# ─── HTTP ヘルパー ─────────────────────────────────────────────────────────────

def _request(method: str, path: str, params: dict = None, signed: bool = False) -> dict:
    params = params or {}
    if signed:
        params["timestamp"] = int(time.time() * 1000)
        query = urllib.parse.urlencode(params)
        sig = hmac.new(API_SECRET.encode(), query.encode(), hashlib.sha256).hexdigest()
        params["signature"] = sig

    query = urllib.parse.urlencode(params)
    url = f"{BASE_URL}{path}"
    if method == "GET" and query:
        url += "?" + query

    data = query.encode() if method == "POST" else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("Content-Type", "application/x-www-form-urlencoded")
    if API_KEY:
        req.add_header("X-MEXC-APIKEY", API_KEY)

    with urllib.request.urlopen(req, timeout=10) as resp:
        return json.loads(resp.read())


# ─── MEXC API ─────────────────────────────────────────────────────────────────

def get_all_symbols() -> set[str]:
    """取引所に現在存在する全シンボルを返す"""
    data = _request("GET", "/api/v3/exchangeInfo")
    return {s["symbol"] for s in data.get("symbols", [])}


def get_ticker_price(symbol: str) -> float | None:
    """シンボルの現在価格を返す。取得失敗時は None"""
    try:
        data = _request("GET", "/api/v3/ticker/price", {"symbol": symbol})
        return float(data["price"])
    except Exception:
        return None


def place_market_buy(symbol: str, quote_qty: float) -> dict:
    """成行買い注文を発注する (quoteOrderQty 指定)"""
    params = {
        "symbol":        symbol,
        "side":          "BUY",
        "type":          "MARKET",
        "quoteOrderQty": quote_qty,
    }
    return _request("POST", "/api/v3/order", params, signed=True)


# ─── スナイパーロジック ────────────────────────────────────────────────────────

def log(msg: str):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
    print(f"[{ts}] {msg}", flush=True)


def try_buy(symbol: str):
    if ORDER_DELAY_SEC > 0:
        log(f"  {ORDER_DELAY_SEC}秒待機中...")
        time.sleep(ORDER_DELAY_SEC)

    if DRY_RUN:
        price = get_ticker_price(symbol)
        log(f"  [DRY RUN] 買い注文スキップ: {symbol} 現在価格={price} "
            f"購入予定={BUY_AMOUNT_USDT} USDT")
        return

    if not API_KEY or not API_SECRET:
        log("  APIキー未設定のため注文をスキップします")
        return

    try:
        result = place_market_buy(symbol, BUY_AMOUNT_USDT)
        log(f"  注文完了: {json.dumps(result, ensure_ascii=False)}")
    except Exception as e:
        log(f"  注文エラー: {e}")


def run():
    log(f"MEXC スナイパー起動 | クォート={QUOTE_CURRENCY} "
        f"購入額={BUY_AMOUNT_USDT} USDT | DRY_RUN={DRY_RUN}")
    log("現在の全シンボルを取得中...")

    known = get_all_symbols()
    # 対象ペアだけフィルタ
    known = {s for s in known if s.endswith(QUOTE_CURRENCY)}
    log(f"監視開始: {len(known)} ペアを把握済み (間隔={POLL_INTERVAL_SEC}秒)")

    while True:
        time.sleep(POLL_INTERVAL_SEC)
        try:
            current = {s for s in get_all_symbols() if s.endswith(QUOTE_CURRENCY)}
        except Exception as e:
            log(f"シンボル取得エラー: {e}")
            continue

        new_symbols = current - known
        if new_symbols:
            for sym in sorted(new_symbols):
                log(f"★ 新規上場検出: {sym}")
                try_buy(sym)
            known = current
        else:
            removed = known - current
            if removed:
                log(f"上場廃止: {removed}")
                known = current


if __name__ == "__main__":
    run()
