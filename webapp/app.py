# -*- coding: utf-8 -*-
"""휴대폰 브라우저용 모바일 웹앱. stock_screener.py의 run_screen()을 그대로 재사용한다.

실행: python app.py  (기본 포트 5000)
같은 Wi-Fi에서 접속: http://<이 PC의 사설 IP>:5000
인터넷 어디서든 접속: cloudflared 터널을 함께 실행 (run_mobile.bat 참고)
"""

import math
import sys
from pathlib import Path
from types import SimpleNamespace

from flask import Flask, jsonify, render_template, request

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from stock_screener import COLOR_COLS, INT_COLS, PERCENT_COLS, run_screen, run_combined  # noqa: E402

app = Flask(__name__)

DEFAULT_ARGS = dict(
    period=20, market="ALL", market_cap_min=0, volume_min=0,
    cum_decline_min=None, up_days_min=None, consecutive_up_min=None,
    theme_keyword=None, min_data_days=None, code=None, name=None,
    search=None, watchlist=None, buy_price=None, quantity=None,
)

WATCHLIST_PATH = Path(__file__).resolve().parent.parent / "관심종목.txt"

PRESETS = {
    "watchlist": dict(watchlist=str(WATCHLIST_PATH)),
    "basic": dict(market_cap_min=1000, volume_min=500000),
    "decline": dict(cum_decline_min=20, period=20),
    "uptrend": dict(up_days_min=12, period=20),
}


def clean(v):
    """NaN/무한대를 JSON 직렬화 가능하도록 정리."""
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return None
    return v


def df_to_section(name, df):
    """DataFrame 하나를 프론트엔드가 그대로 그릴 수 있는 dict로 변환한다
    (표 컬럼/행 + 색상·퍼센트·정수 컬럼 힌트 + 매입가 모드 여부)."""
    columns = list(df.columns)
    rows = [[clean(v) for v in row] for row in df.itertuples(index=False, name=None)]
    return {
        "name": name,
        "columns": columns,
        "rows": rows,
        "color_cols": [c for c in columns if c in COLOR_COLS],
        "percent_cols": [c for c in columns if c in PERCENT_COLS],
        "int_cols": [c for c in columns if c in INT_COLS],
        "buy_price_mode": "평가손익" in columns,
        "count": len(rows),
    }


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/watchlist", methods=["GET"])
def get_watchlist():
    if not WATCHLIST_PATH.exists():
        return jsonify({"text": ""})
    return jsonify({"text": WATCHLIST_PATH.read_text(encoding="utf-8-sig")})


@app.route("/api/watchlist", methods=["POST"])
def save_watchlist():
    data = request.get_json(force=True)
    WATCHLIST_PATH.write_text(data.get("text", ""), encoding="utf-8")
    return jsonify({"ok": True})


@app.route("/api/run", methods=["POST"])
def api_run():
    data = request.get_json(force=True) or {}
    preset = data.get("preset")

    kwargs = dict(DEFAULT_ARGS)
    if preset and preset in PRESETS:
        kwargs.update(PRESETS[preset])

    for key in ("code", "name", "search", "theme_keyword"):
        if data.get(key):
            kwargs[key] = data[key]
    if data.get("period"):
        kwargs["period"] = int(data["period"])
    if data.get("custom_args"):
        kwargs.update(parse_custom_args(data["custom_args"]))
    if data.get("buy_price"):
        kwargs["buy_price"] = float(data["buy_price"])
    if data.get("quantity"):
        kwargs["quantity"] = int(data["quantity"])

    args = SimpleNamespace(**kwargs)

    try:
        result = run_screen(args)
    except (FileNotFoundError, ValueError) as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"오류가 발생했습니다: {e}"}), 500

    return jsonify(df_to_section(None, result))


@app.route("/api/run_combined", methods=["POST"])
def api_run_combined():
    """9) 통합 조회: 테마(콤마로 여러개 -> 테마별 섹션) + 관심종목/종목검색/종목조회(선택)를
    한 번에 실행해서 섹션 목록으로 반환한다. 모바일 웹앱에서는 엑셀 시트 대신
    화면에 섹션(카드)을 여러 개 순서대로 보여주는 방식으로 동일한 결과를 표현한다."""
    data = request.get_json(force=True) or {}

    try:
        sheets = run_combined(
            wl_input=data.get("wl_input") or None,
            watchlist_path=str(WATCHLIST_PATH),
            search_kw=data.get("search_kw") or None,
            lookup=data.get("lookup") or None,
            theme_kw=data.get("theme_keyword") or None,
            buy_price=float(data["buy_price"]) if data.get("buy_price") else None,
            quantity=int(data["quantity"]) if data.get("quantity") else None,
            period=int(data["period"]) if data.get("period") else 20,
        )
    except Exception as e:
        return jsonify({"error": f"오류가 발생했습니다: {e}"}), 500

    if not sheets:
        return jsonify({"error": "결과가 있는 항목이 하나도 없습니다."}), 400

    return jsonify({"sections": [df_to_section(name, df) for name, df in sheets]})


CUSTOM_ARG_MAP = {
    "--market-cap-min": ("market_cap_min", float),
    "--volume-min": ("volume_min", float),
    "--cum-decline-min": ("cum_decline_min", float),
    "--up-days-min": ("up_days_min", int),
    "--consecutive-up-min": ("consecutive_up_min", int),
    "--period": ("period", int),
    "--theme-keyword": ("theme_keyword", str),
    "--code": ("code", str),
    "--name": ("name", str),
    "--search": ("search", str),
    "--market": ("market", str),
}


def parse_custom_args(text):
    """'--market-cap-min 1000 --up-days-min 10' 같은 문자열을 kwargs로 변환."""
    tokens = text.split()
    out = {}
    i = 0
    while i < len(tokens):
        tok = tokens[i]
        if tok in CUSTOM_ARG_MAP and i + 1 < len(tokens):
            field, typ = CUSTOM_ARG_MAP[tok]
            try:
                out[field] = typ(tokens[i + 1])
            except ValueError:
                pass
            i += 2
        else:
            i += 1
    return out


if __name__ == "__main__":
    print("모바일 웹앱 서버 시작: http://0.0.0.0:5000")
    print("같은 Wi-Fi 폰에서: http://<이 PC의 사설 IP>:5000")
    app.run(host="0.0.0.0", port=5000, debug=False)
