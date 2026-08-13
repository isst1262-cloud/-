# -*- coding: utf-8 -*-
"""run.bat 에서 호출하는 대화형 메뉴. stock_screener.py를 옵션과 함께 실행한다."""

import subprocess
import sys
import traceback
import unicodedata

PY = sys.executable


BOM = "﻿"


def ask(prompt):
    return input(prompt).strip().lstrip(BOM)


def _width(s: str) -> int:
    """한글은 2칸, 영문/숫자는 1칸으로 계산한 터미널 표시 폭."""
    return sum(2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1 for ch in s)


def _menu_line(num: str, label: str, desc: str, label_width: int = 22) -> str:
    pad = " " * max(0, label_width - _width(label))
    return f"  {num}) {label}{pad}{desc}"


def ask_buy_price(args):
    """매입가(선택)를 물어보고 args에 --buy-price/--quantity를 추가한다."""
    buy_price = ask("매입가를 입력하세요 (손익표 계산, 안 하려면 그냥 Enter): ")
    if buy_price:
        args += ["--buy-price", buy_price]
        qty = ask("매입수량을 입력하세요 (총 손익 계산, 안 하려면 그냥 Enter): ")
        if qty:
            args += ["--quantity", qty]
    return args


def run(args, out_hint):
    cmd = [PY, "stock_screener.py"] + args
    print(f"\n> {' '.join(cmd)}\n")
    subprocess.run(cmd)
    print(f"\n결과 엑셀 파일이 이 폴더에 저장되었습니다: {out_hint} (자동으로 열립니다)")


def show_menu():
    print("=" * 50)
    print("  한국 주식 조건검색 스크리너")
    print("=" * 50)
    print(_menu_line("1", "관심종목 조회", "(직접 입력 또는 관심종목.txt)"))
    print(_menu_line("2", "기본 스크리닝", "(시총 1000억+ / 거래량 50만주+)"))
    print(_menu_line("3", "낙폭과대주", "(최근 20일 20% 이상 하락)"))
    print(_menu_line("4", "상승세 지속 종목", "(최근 20일 중 12일 이상 상승)"))
    print(_menu_line("5", "테마주 검색", "(키워드 직접 입력)"))
    print(_menu_line("6", "특정 종목 조회", "(정확한 종목코드/종목명 + 손익표)"))
    print(_menu_line("7", "직접 옵션 입력", "(고급 사용자용)"))
    print(_menu_line("8", "종목 검색", "(이름 일부만 입력 + 손익표)"))
    print(_menu_line("9", "통합 조회", "(테마 위주, 테마별 시트+관심/검색/조회)"))
    print("  0) 종료")
    print("=" * 50)


def main():
    while True:
        show_menu()
        sel = ask("원하는 번호를 입력하세요: ")

        if sel == "1":
            direct = ask("조회할 종목을 입력하세요 (콤마로 여러개, 예: 006400,삼성전자 / 비워두면 관심종목.txt 사용): ")
            if direct:
                args1 = ["--code", direct, "--name", direct, "--out", "result_관심종목.xlsx"]
            else:
                args1 = ["--watchlist", "관심종목.txt", "--out", "result_관심종목.xlsx"]
            args1 = ask_buy_price(args1)
            run(args1, "result_관심종목.xlsx")
        elif sel == "2":
            run(["--market-cap-min", "1000", "--volume-min", "500000", "--out", "result_기본스크리닝.xlsx"],
                "result_기본스크리닝.xlsx")
        elif sel == "3":
            run(["--cum-decline-min", "20", "--period", "20", "--out", "result_낙폭과대주.xlsx"],
                "result_낙폭과대주.xlsx")
        elif sel == "4":
            run(["--up-days-min", "12", "--period", "20", "--out", "result_상승세.xlsx"],
                "result_상승세.xlsx")
        elif sel == "5":
            kw = ask("테마 키워드를 입력하세요 (예: 2차전지,로봇): ")
            run(["--theme-keyword", kw, "--out", "result_테마주.xlsx"], "result_테마주.xlsx")
        elif sel == "6":
            code = ask("종목코드 또는 종목명을 입력하세요 (예: 006400 또는 삼성SDI): ")
            args6 = ["--name", code, "--code", code, "--out", "result_종목조회.xlsx"]
            args6 = ask_buy_price(args6)
            run(args6, "result_종목조회.xlsx")
        elif sel == "7":
            print("예시: --market-cap-min 1000 --volume-min 300000 --up-days-min 10")
            opt = ask("적용할 옵션을 입력하세요: ")
            run(opt.split(), "result_직접입력.xlsx")
        elif sel == "8":
            kw = ask("검색할 종목명(일부)을 입력하세요 (예: 삼성, SK, 전자): ")
            args8 = ["--search", kw, "--out", "result_종목검색.xlsx"]
            args8 = ask_buy_price(args8)
            run(args8, "result_종목검색.xlsx")
        elif sel == "9":
            print("테마주 위주 조회입니다. 각 항목은 비워두고 Enter만 누르면 건너뜁니다.")
            tkw = ask("[테마주] 테마 키워드 (콤마로 여러개, 예: 2차전지,로봇,반도체 -> 테마별로 시트가 따로 생성됨): ")
            wl = ask("[관심종목] 직접 입력 (선택, 콤마로 여러개): ")
            skw = ask("[종목검색] 검색 키워드 (선택): ")
            lk = ask("[종목조회] 종목코드/종목명 (선택): ")
            args9 = ["--combined", "--out", "result_통합조회.xlsx"]
            if tkw:
                args9 += ["--theme-keyword", tkw]
            if wl:
                args9 += ["--wl-input", wl]
            if skw:
                args9 += ["--search-kw", skw]
            if lk:
                args9 += ["--lookup", lk]
            args9 = ask_buy_price(args9)
            run(args9, "result_통합조회.xlsx (테마별 시트 + 관심종목/종목검색/종목조회)")
        elif sel == "0":
            break
        else:
            print("잘못된 입력입니다.")
            continue

        input("\n계속하려면 Enter를 누르세요...")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        print("\n[오류 발생]")
        traceback.print_exc()
        input("\nEnter를 누르면 창이 닫힙니다...")
