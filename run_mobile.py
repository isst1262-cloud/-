# -*- coding: utf-8 -*-
"""모바일 웹앱을 켜고, 접속 주소를 QR코드로 보여준다.
휴대폰으로 QR코드를 스캔하면 바로 접속되므로 주소를 직접 타이핑하지
않아도 되고, 오타나 예전에 죽은 주소를 잘못 쓰는 실수를 막을 수 있다.

run_mobile.bat 에서 이 스크립트를 실행한다.
"""

import os
import re
import shutil
import subprocess
import sys
import threading
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
WEBAPP_DIR = ROOT / "webapp"
QR_PATH = ROOT / "접속주소_QR코드.png"

sys.path.insert(0, str(WEBAPP_DIR))

# 콘솔이 아닌 곳(파일/파이프)으로 출력이 리다이렉트되는 경우에도 한 줄씩
# 바로바로 보이도록(파이썬 기본 블록 버퍼링 방지) 줄 단위 버퍼링으로 강제한다.
try:
    sys.stdout.reconfigure(line_buffering=True)
except Exception:
    pass


def start_flask():
    from app import app
    app.run(host="0.0.0.0", port=5000, debug=False, use_reloader=False)


def show_qr(url):
    try:
        import qrcode
    except ImportError:
        print("(qrcode 패키지가 없어 QR코드는 건너뜁니다. pip install qrcode[pil])", file=sys.stderr)
        return

    img = qrcode.make(url, box_size=10, border=2)
    img.save(QR_PATH)

    try:
        os.startfile(QR_PATH)
    except Exception:
        print(f"(QR코드 이미지 파일: {QR_PATH})")


URL_PATTERN = re.compile(r"https://(?!api\.)[a-zA-Z0-9-]+\.trycloudflare\.com")


def run_tunnel_once(cloudflared, protocol=None, timeout=25):
    """cloudflared를 한 번 실행해서 실제 접속 주소를 찾으면 (Popen, url)을,
    못 찾고 끝나면 (Popen, None)을 반환한다.
    'api.trycloudflare.com'은 클라우드플레어 내부 주소(오류 메시지 등에 등장)라
    URL_PATTERN에서 일부러 제외한다 - 실제 터널 주소는 절대 'api'가 아니다.

    protocol=None 이면 cloudflared 기본값(보통 QUIC/UDP)을 그대로 쓰고,
    "http2"를 주면 TCP 443 기반으로 강제한다. 터널 자체를 못 만드는 경우
    (드묾, 보통 api.trycloudflare.com DNS 조회 실패 같은 일시적 문제)에만
    이 함수를 다시 부른다 - 아래 wait_until_reachable의 "느리게 열림"과는
    다른 문제라 별도로 다룬다."""
    cmd = [cloudflared, "tunnel"]
    if protocol:
        cmd += ["--protocol", protocol]
    cmd += ["--url", "http://localhost:5000"]

    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        text=True, encoding="utf-8", errors="replace", bufsize=1,
    )

    found_url = None
    start = time.time()
    for line in proc.stdout:
        print(line, end="")
        if "failed to request quick Tunnel" in line:
            break
        m = URL_PATTERN.search(line)
        if m:
            found_url = m.group(0)
            break
        if time.time() - start > timeout:
            break

    return proc, found_url


def wait_until_reachable(url, tries=20, delay=2):
    """터널 주소는 만들어진 직후 전 세계 DNS에 퍼지는 데 보통 10~20초
    정도 걸린다(클라우드플레어도 "만들어지자마자는 안 열릴 수 있다"고
    안내함). 그래서 여기서는 새 터널을 다시 만드는 대신, 같은 주소로
    충분히(최대 tries*delay 초) 기다렸다가 응답이 오는지 확인만 한다."""
    import requests
    for _ in range(tries):
        try:
            r = requests.get(url, timeout=5)
            if r.status_code < 500:
                return True
        except requests.RequestException:
            pass
        time.sleep(delay)
    return False


def main():
    cloudflared = shutil.which("cloudflared")
    if not cloudflared:
        print("[오류] cloudflared를 찾을 수 없습니다.")
        print("설치 후 다시 실행해 주세요: winget install Cloudflare.cloudflared")
        print("(방금 설치했다면 이 창을 닫고 run_mobile.bat을 다시 실행해 보세요)")
        input("\nEnter를 누르면 창이 닫힙니다...")
        return

    print("로컬 서버를 시작하는 중...")
    threading.Thread(target=start_flask, daemon=True).start()
    time.sleep(2)

    # 터널 "생성" 자체가 실패하는 경우(드묾)에 대비한 재시도.
    # 기본 프로토콜로 먼저 시도하고, 계속 안 되면 http2로 바꿔본다.
    protocols = [None, None, "http2", "http2"]
    max_tries = len(protocols)
    proc = None
    url = None
    try:
        for attempt, protocol in enumerate(protocols, start=1):
            proto_label = protocol or "기본값"
            print(f"공개 접속 주소를 만드는 중입니다 (Cloudflare 터널, 시도 {attempt}/{max_tries}, 프로토콜: {proto_label})...\n")
            proc, url = run_tunnel_once(cloudflared, protocol=protocol)
            if url:
                break
            try:
                proc.terminate()
            except Exception:
                pass
            print("\n터널 생성에 실패했습니다. 몇 초 후 다시 시도합니다...\n", file=sys.stderr)
            time.sleep(3)

        if not url:
            print("\n[오류] 터널을 반복해서 만들지 못했습니다.")
            print("인터넷 연결 상태를 확인한 뒤 run_mobile.bat을 다시 실행해 주세요.")
            input("\nEnter를 누르면 창이 닫힙니다...")
            return

        print("\n" + "=" * 62)
        print("  주소가 만들어졌습니다. 아래 QR코드를 폰 카메라로 스캔하면")
        print("  됩니다. (지금 여는 이미지)")
        print(f"\n  주소를 직접 입력하려면: {url}")
        print("\n  * 이 주소는 전 세계에 퍼지는 데 10~20초 정도 걸릴 수")
        print("    있습니다. 방금 만들어졌다면 스캔 후 화면이 안 뜨더라도")
        print("    몇 초 기다렸다가 새로고침해 보세요.")
        print("=" * 62 + "\n")
        show_qr(url)

        print("실제로 열리는지 이 PC에서도 확인해 보는 중...", file=sys.stderr)
        if wait_until_reachable(url):
            print("확인 완료 - 정상적으로 접속됩니다.\n", file=sys.stderr)
        else:
            print("아직 응답이 없지만(전파 지연일 수 있음) 주소 자체는 살아",
                  "있으니 잠시 후 다시 스캔해 보세요.\n", file=sys.stderr)

        # 터널이 살아있는 동안 나머지 로그를 계속 화면에 보여준다.
        for line in proc.stdout:
            print(line, end="")
        proc.wait()
    except KeyboardInterrupt:
        if proc:
            proc.terminate()


if __name__ == "__main__":
    main()
