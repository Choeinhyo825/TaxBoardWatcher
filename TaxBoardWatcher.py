from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from PIL import Image, ImageDraw, ImageTk
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from win32com.client import Dispatch
from urllib.parse import urlencode
from selenium import webdriver
from bs4 import BeautifulSoup
from html import escape
from io import BytesIO
import tkinter as tk
import webbrowser
import threading
import datetime
import requests
import pystray
import ctypes
import random
import queue
import json
import time
import sys
import re
import os

try:
    from winotify import Notification as WinotifyNotification
except ImportError:
    WinotifyNotification = None

VERSION = "v.5.0.2"

# --- 자동 업데이트(GitHub Releases) ---
GITHUB_REPO = "Choeinhyo825/TaxBoardWatcher"
UPDATE_ASSET_NAME = "TaxBoardWatcher.exe"
UPDATE_CHECK_INTERVAL = 12 * 60 * 60  # 자동 업데이트 확인 최소 간격(12시간)

BOARD_DATA = "data/board_data.json"
LOG = "data/log.txt"
ICON = 'data/tax.png'
SEARCHING_ICON = 'data/searching.png'
CONFIG = "data/config.json"
SLEEP_TIME = 3*60*60 # 기본 3시간

# 게시판 개요 카드 헤더용 기관 로고(각 사이트 제공 이미지 URL)
BOARD_OVERVIEW_CARD_LOGOS = {
    "hometax": "https://hometax.go.kr/css/comm/bpr_portal_images/logo_hometax.svg",
    "moef": "https://mofe.go.kr/images/2026/logo.svg",
    "moleg": "https://www.moleg.go.kr/_kor/img/layout/logo.png",
    "gwanbo": "https://gwanbo.go.kr/image/common/logo.jpg",
    "mois": "https://www.mois.go.kr/frt2022/main/img/common/logo4.png",
}

# --- HTTP 공통 헤더/세션 ---
COMMON_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
    "Connection": "keep-alive",
}

def make_http_session():
    session = requests.Session()
    retry = Retry(
        total=3,
        connect=3,
        read=3,
        backoff_factor=1.5,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET", "POST"],
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    session.headers.update(COMMON_HEADERS)
    return session

tray_icon = None  # 전역 tray_icon 변수

_logo_pil_cache: dict = {}  # site_key -> PIL.Image | None
_interval_changed = threading.Event()  # 모니터링 주기 변경 시 sleep 중단용
_notify_queue = queue.Queue()  # 새 글 알림을 순차 표시하기 위한 큐

def enqueue_notification(*args, **kwargs):
    """새 글 알림을 큐에 넣는다. 모니터링 스레드가 알림창 대기(mainloop)로 멈추지 않도록 한다."""
    _notify_queue.put((args, kwargs))

def _notification_worker():
    """알림 큐를 소비해 한 번에 하나씩 알림창을 표시하는 전용 스레드."""
    while True:
        args, kwargs = _notify_queue.get()
        try:
            send_notification(*args, **kwargs)
        except Exception as e:
            tax_log("e", "", f"알림 표시 오류: {e}")
        finally:
            _notify_queue.task_done()

def get_base_dir() -> str:
    """실행 환경(스크립트/PyInstaller exe)에 관계없이 프로그램 루트 디렉토리를 반환."""
    if getattr(sys, "frozen", False):
        # PyInstaller --onefile: sys.executable = 실제 .exe 경로
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def _load_logo_photo(site_key: str, max_h: int = 38):
    if site_key not in _logo_pil_cache:
        base_dir = get_base_dir()
        logo_path_candidate = os.path.join(base_dir, "data", f"{site_key}_logo.png")
        logo_path = logo_path_candidate if os.path.exists(logo_path_candidate) else None
        if not logo_path:
            _logo_pil_cache[site_key] = None
        else:
            try:
                img = Image.open(logo_path).convert("RGBA")
                w, h = img.size
                if h > max_h:
                    img = img.resize((int(w * max_h / h), max_h), Image.LANCZOS)
                _logo_pil_cache[site_key] = img
            except Exception as e:
                _logo_pil_cache[site_key] = None
    pil_img = _logo_pil_cache.get(site_key)
    if pil_img is None:
        return None
    try:
        return ImageTk.PhotoImage(pil_img)
    except Exception:
        return None

# --- config.json 로드/저장 ---
def get_config():
    global SLEEP_TIME
    try:
        with open(CONFIG, "r", encoding="utf-8") as f:
            config = json.load(f)
            sleep_hour = config.get("sleep_hour")
            if isinstance(sleep_hour, int) and sleep_hour >= 0:
                SLEEP_TIME = sleep_hour * 60 * 60
            else:
                tax_log("w", "", f"설정된 sleep_hour 값이 올바르지 않습니다. 기본값(3)을 사용 합니다.")
    except Exception as e:
        tax_log("w","",f"설정 파일을 불러올 수 없습니다: {e}")

def save_config():
    try:
        with open(CONFIG, "w", encoding="utf-8") as f:
            json.dump({"sleep_hour": int(SLEEP_TIME / 3600)}, f, ensure_ascii=False, indent=2)
    except Exception as e:
        tax_log("w", "", f"설정 저장 실패: {e}")

def set_sleep_hour(hours):
    global SLEEP_TIME
    SLEEP_TIME = hours * 3600
    save_config()
    tax_log("i", "", f"모니터링 주기 변경: {hours}시간")
    _interval_changed.set()

# --- 공통 함수 ---
def tax_log(logType, category, message):
    if logType == "e":
        logType = "[ERROR]"
    elif logType == "w":
        logType = "[WARNING]"
    elif logType == "i":
        logType = "[INFO]"
    else:
        logType = ""
    
    if category == "hometax":
        category = "[ 홈택스 ]"
    elif category == "moef":
        category = "[ 기재부 ]"
    elif category == "moleg":
        category = "[ 법제처 ]"
    elif category == "gwanbo":
        category = "[ 관　보 ]"
    elif category == "mois":
        category = "[ 행안부 ]"
    else:
        category = "[ SYSTEM ]"

    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(f"[{now}] {logType} {category} {message}\n")

# 수동 모니터링 결과를 Windows 알림(토스트)으로 표시
def notify_manual_monitor_toast(title: str, msg: str, *, error: bool = False):
    if WinotifyNotification is None:
        tax_log("w", "", "Windows 알림(winotify)을 사용할 수 없습니다.")
        return
    max_len = 320
    if len(msg) > max_len:
        msg = msg[: max_len - 1] + "…"
    icon_path = ""
    abs_icon = os.path.abspath(ICON)
    if os.path.isfile(abs_icon):
        icon_path = abs_icon
    duration = "long" if error else "short"
    try:
        toast = WinotifyNotification(
            app_id="TaxBoardWatcher",
            title=title,
            msg=msg,
            icon=icon_path,
            duration=duration,
        )
        toast.show()
    except Exception as e:
        tax_log("e", "", f"Windows 토스트 표시 실패: {e}")


def send_notification(title, message, url=None, site_key=None, meta=None):
    BG     = "#ffffff"
    BG_MSG = "#f4f7fb"
    ACCENT = "#0078D7"
    FONT   = "Malgun Gothic"
    PRIMARY  = "#2563EB"
    TEXT_S   = "#333333"

    try:
        root = tk.Tk()
        try:
            dpi = ctypes.windll.user32.GetDpiForSystem()
            root.tk.call('tk', 'scaling', dpi / 72.0)
        except Exception:
            pass
        root.withdraw()
        root.overrideredirect(True)
        root.configure(bg="#c8cdd2")
        root.attributes("-topmost", True)
        root.attributes("-alpha", 0.0)

        def open_url():
            webbrowser.open(url)
            root.destroy()
        def on_close():
            root.destroy()

        # 1px 외곽 테두리
        outer = tk.Frame(root, bg="#c8cdd2", padx=1, pady=1)
        outer.pack(fill="both", expand=True)

        wrap = tk.Frame(outer, bg=BG)
        wrap.pack(fill="both", expand=True)

        # 드래그 가능한 상단 액센트 바
        drag_bar = tk.Frame(wrap, bg=ACCENT, height=36)
        drag_bar.pack(fill="x", side="top")
        drag_bar.pack_propagate(False)

        bar_lbl = tk.Label(drag_bar, text="TaxBoardWatcher - 새로운 게시물" if url else "TaxBoardWatcher", bg=ACCENT,
                           fg="white", font=(FONT, 9), padx=12)
        bar_lbl.pack(side="left", fill="y")

        def _drag_start(e):
            root._dx = e.x_root - root.winfo_x()
            root._dy = e.y_root - root.winfo_y()
        def _drag_move(e):
            root.geometry(f"+{e.x_root - root._dx}+{e.y_root - root._dy}")

        for w in (drag_bar, bar_lbl):
            w.bind("<ButtonPress-1>", _drag_start)
            w.bind("<B1-Motion>",     _drag_move)

        # 본문
        body = tk.Frame(wrap, bg=BG, padx=24, pady=20)
        body.pack(fill="both", expand=True)

        if url:
            # 로고 or 사이트명 텍스트
            logo_photo = _load_logo_photo(site_key, max_h=52) if site_key else None
            if logo_photo:
                logo_lbl = tk.Label(body, image=logo_photo, bg=BG)
                logo_lbl.image = logo_photo  # GC 방지
                logo_lbl.pack(anchor="w", pady=(0, 10))
            else:
                tk.Label(body, text=title, bg=BG,
                         font=(FONT, 13, "bold"), fg="#1a1a1a").pack(anchor="w", pady=(0, 10))

            # 구분선
            tk.Frame(body, bg="#e4e8ed", height=1).pack(fill="x", pady=(0, 12))

            # 메시지 박스
            msg_box = tk.Frame(body, bg=BG_MSG, padx=14, pady=12)
            msg_box.pack(fill="x", pady=(0, 20))
            tk.Label(msg_box, text=message, bg=BG_MSG,
                     font=(FONT, 10), fg=TEXT_S,
                     wraplength=550, justify="left").pack(anchor="w")

            # 날짜 등 부가 정보 (값이 있는 항목만 표시)
            meta_rows = [(lbl, val) for (lbl, val) in (meta or []) if val]
            if meta_rows:
                info = tk.Frame(msg_box, bg=BG_MSG)
                info.pack(anchor="w", fill="x", pady=(10, 0))
                for lbl, val in meta_rows:
                    row = tk.Frame(info, bg=BG_MSG)
                    row.pack(anchor="w", fill="x", pady=(2, 0))
                    tk.Label(row, text=lbl, bg=BG_MSG, font=(FONT, 9, "bold"),
                             fg="#5a6570", width=8, anchor="w").pack(side="left")
                    tk.Label(row, text=val, bg=BG_MSG, font=(FONT, 9),
                             fg=TEXT_S, wraplength=480, justify="left").pack(side="left")

            # 버튼 (중앙)
            btn_row = tk.Frame(body, bg=BG)
            btn_row.pack()
            tk.Button(btn_row, text="열기", command=open_url,
                      bg=ACCENT, fg="white", font=(FONT, 9),
                      relief="flat", padx=10, pady=1, cursor="hand2",
                      activebackground="#106EBE", activeforeground="white").pack(side="left", padx=4)
            tk.Button(btn_row, text="닫기", command=on_close,
                      bg="#e8e8e8", fg="#444444", font=(FONT, 9),
                      relief="flat", padx=10, pady=1, cursor="hand2",
                      activebackground="#d4d4d4").pack(side="left", padx=4)

        elif message == 'start':
            global SLEEP_TIME
            tk.Label(body, text="TaxBoardWatcher", bg=BG,
                     font=(FONT, 14, "bold"), fg="#1a1a1a").pack(anchor="w")
            tk.Label(body, text="모니터링을 시작합니다.", bg=BG,
                     font=(FONT, 9), fg="#999999").pack(anchor="w", pady=(3, 12))

            tk.Frame(body, bg="#e4e8ed", height=1).pack(fill="x", pady=(0, 12))

            info_box = tk.Frame(body, bg=BG_MSG, padx=14, pady=12)
            info_box.pack(fill="x", pady=(0, 20))
            tk.Label(info_box, text="🕓 모니터링 간격", bg=BG_MSG,
                     font=(FONT, 9), fg=TEXT_S).pack(side="left")
            tk.Label(info_box, text=f"{int(SLEEP_TIME / 3600)}시간", bg=BG_MSG,
                     font=(FONT, 9, "bold"), fg=PRIMARY).pack(side="right")

            btn_row = tk.Frame(body, bg=BG)
            btn_row.pack()
            tk.Button(btn_row, text="시작", command=on_close,
                      bg=ACCENT, fg="white", font=(FONT, 9),
                      relief="flat", padx=10, pady=1, cursor="hand2",
                      activebackground="#106EBE", activeforeground="white").pack()

        root.update_idletasks()
        min_w = 350 if message == "start" else 600
        w = max(root.winfo_reqwidth(), min_w)
        h = root.winfo_reqheight()
        x = (root.winfo_screenwidth() - w) // 2
        y = (root.winfo_screenheight() - h) // 2
        root.geometry(f"{w}x{h}+{x}+{y}")
        root.deiconify()

        for i in range(0, 11):
            root.attributes("-alpha", i / 10)
            root.update()
            time.sleep(0.03)

        root.mainloop()
    except Exception as e:
        tax_log("e", "", f" 알림 전송 오류: {e}")

# 모니터링 섹션 키 (board_data.json 값은 항상 dict[str, dict] 형태로 통일)
BOARD_DATA_SECTIONS = ("hometax", "moef", "moleg", "gwanbo", "mois")


def _normalize_board_data(data):
    """루트 dict의 모니터링 섹션을 문자열·불완전 dict 등 섞여 있어도 동일 스키마로 통일."""
    if not isinstance(data, dict):
        return {"hometax": {}, "moef": {}, "moleg": {}, "gwanbo": {}, "mois": {}}
    out = {}
    for k, v in data.items():
        if k not in BOARD_DATA_SECTIONS:
            out[k] = v
    for section in BOARD_DATA_SECTIONS:
        raw = data.get(section)
        if not isinstance(raw, dict):
            if raw is not None:
                tax_log(
                    "w",
                    "",
                    f"board_data.json의 {section}이(가) 객체가 아니어 빈 객체로 통일합니다.",
                )
            out[section] = {}
            continue
        if section == "hometax":
            out[section] = {k: _hometax_entry_from_value(k, v) for k, v in raw.items()}
        elif section == "moef":
            out[section] = {k: _moef_entry_from_value(v) for k, v in raw.items()}
        elif section == "moleg":
            out[section] = {k: _moleg_entry_from_value(v) for k, v in raw.items()}
        elif section == "gwanbo":
            out[section] = {k: _gwanbo_entry_from_value(v) for k, v in raw.items()}
        elif section == "mois":
            out[section] = {k: _mois_entry_from_value(v) for k, v in raw.items()}
    return out


def save_data(data):
    if not isinstance(data, dict):
        tax_log("e", "", "save_data: 루트가 dict가 아닙니다. 저장을 건너뜁니다.")
        return
    normalized = _normalize_board_data(data)
    for sec in BOARD_DATA_SECTIONS:
        blob = normalized.get(sec, {})
        cur = data.get(sec)
        if isinstance(cur, dict):
            cur.clear()
            cur.update(blob)
        else:
            data[sec] = blob
    with open(BOARD_DATA, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _moef_notice_period_strip(s: str) -> str:
    """MOEF 목록의 '예고기간 : …' 라벨을 제거하고 날짜 구간 문자열만 남긴다."""
    s = (s or "").strip()
    if not s:
        return ""
    s = re.sub(r"^\s*예고기간\s*[:：]\s*", "", s)
    return s.strip()


def _moef_entry_from_value(v):
    """moef 항목: dict(신규) 또는 str(구버전)을 {title, date, depart}로 통일."""
    if isinstance(v, dict):
        return {
            "title": (v.get("title") or "").strip(),
            "date": _moef_notice_period_strip(v.get("date") or ""),
            "depart": (v.get("depart") or "").strip(),
        }
    return {"title": str(v).strip(), "date": "", "depart": ""}


def _moleg_entry_from_value(v):
    """moleg 항목: dict(신규) 또는 str(구버전)을 표시용 필드로 통일."""
    if isinstance(v, dict):
        return {
            "law_type": (v.get("law_type") or "").strip(),
            "title": (v.get("title") or "").strip(),
            "ministry": (v.get("ministry") or "").strip(),
            "start_date": (v.get("start_date") or "").strip(),
            "end_date": (v.get("end_date") or "").strip(),
        }
    return {
        "law_type": "",
        "title": str(v).strip(),
        "ministry": "",
        "start_date": "",
        "end_date": "",
    }


def _hometax_entry_from_value(key, v):
    """hometax 항목: dict(신규) 또는 str(구버전)을 {number, title, changed_date}로 통일."""
    if isinstance(v, dict):
        return {
            "number": (v.get("number") or "").strip(),
            "title": (v.get("title") or "").strip(),
            "changed_date": (v.get("changed_date") or "").strip(),
        }
    title = str(v).strip()
    number = ""
    changed_date = ""
    m = re.match(r"^(\d{8})_(.+)$", str(key))
    if m:
        ymd = m.group(1)
        number = m.group(2)
        changed_date = f"{ymd[:4]}-{ymd[4:6]}-{ymd[6:8]}"
    return {"number": number, "title": title, "changed_date": changed_date}


def _gwanbo_entry_from_value(v):
    """gwanbo 항목: dict(신규) 또는 str(구버전)을 {title, published_date}로 통일."""
    if isinstance(v, dict):
        return {
            "title": (v.get("title") or "").strip(),
            "published_date": (v.get("published_date") or "").strip(),
        }
    return {"title": str(v).strip(), "published_date": ""}


def _mois_entry_from_value(v):
    """mois 항목: dict(신규) 또는 str(구버전)을 {title, date}로 통일."""
    if isinstance(v, dict):
        return {
            "title": (v.get("title") or "").strip(),
            "date": (v.get("date") or "").strip(),
        }
    return {"title": str(v).strip(), "date": ""}


def load_data():
    try:
        if os.path.exists(BOARD_DATA):
            with open(BOARD_DATA, 'r', encoding='utf-8') as f:
                data = json.load(f)
            moef = data.get("moef")
            if isinstance(moef, dict):
                data["moef"] = {k: _moef_entry_from_value(v) for k, v in moef.items()}
            moleg = data.get("moleg")
            if isinstance(moleg, dict):
                data["moleg"] = {k: _moleg_entry_from_value(v) for k, v in moleg.items()}
            hometax = data.get("hometax")
            if isinstance(hometax, dict):
                data["hometax"] = {
                    k: _hometax_entry_from_value(k, v) for k, v in hometax.items()
                }
            gwanbo = data.get("gwanbo")
            if isinstance(gwanbo, dict):
                data["gwanbo"] = {k: _gwanbo_entry_from_value(v) for k, v in gwanbo.items()}
            mois = data.get("mois")
            if isinstance(mois, dict):
                data["mois"] = {k: _mois_entry_from_value(v) for k, v in mois.items()}
            return data
        return {"hometax": {}, "moef": {}, "moleg": {}, "gwanbo": {}, "mois": {}}
    except Exception as e:
        tax_log("e", "", f"파일 로드 오류: {e}")
        return {"hometax": {}, "moef": {}, "moleg": {}, "gwanbo": {}, "mois": {}}

# --- 홈택스 스크래핑 ---
class HomeTaxScraper:
    # WebDriver 초기화
    def init_driver(self):
        # ChromeOptions 설정
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        # 기본 headless 창(800x600)에서는 그리드 우측 컬럼(변경일 등)이 overflow:hidden으로 잘려
        # Selenium .text가 빈 문자열을 반환하므로 창을 넓게 띄운다.
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        # ChromeDriverManager에서 설치된 드라이버 경로 가져오기
        service = Service(ChromeDriverManager().install())
        # WebDriver를 초기화할 때, service와 options를 명확히 전달
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        
    # __init__
    def __init__(self, shared_data):
        self.url = "https://hometax.go.kr/websquare/websquare.html?w2xPath=/ui/pp/index_pp.xml&tmIdx=16&tm2lIdx=1602000000&tm3lIdx="
        self.known_posts = shared_data.get("hometax", {})
        self.driver = None

    def fetch_latest_posts(self):
        self._fetch_ok = False
        try:
            self.init_driver()
            self.driver.get(self.url)

            # 홈택스는 WebSquare + 보안체크(serviceCheck/permission/token/wqAction) 단계를
            # 거쳐 그리드가 뜨기까지 ~10초가 걸린다. 고정 sleep 대신 요소가 채워질 때까지 대기.
            wait = WebDriverWait(self.driver, 30)
            tbody = wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, '#mf_txppWframe_grdList_body_tbody')
                )
            )
            # tbody가 생겨도 행 데이터가 아직 안 채워질 수 있으므로 최소 1행이 생길 때까지 대기
            wait.until(
                lambda d: len(tbody.find_elements(By.CSS_SELECTOR, 'tr')) > 0
            )

            # 모든 tr 요소 찾기
            rows = tbody.find_elements("css selector", 'tr')
            posts = []

            # .text는 화면에 보이는 텍스트만 반환하므로(잘린 셀은 ''), 표시 여부와 무관한 textContent를 읽는다.
            def cell_text(row, colindex, tag):
                td = row.find_element("css selector", f'td[data-colindex="{colindex}"]')
                return (td.find_element("css selector", tag).get_attribute("textContent") or "").strip()

            for row in rows:
                try:
                    code = cell_text(row, 0, 'nobr')        # 번호
                    title = cell_text(row, 2, 'a')          # 제목
                    date_raw = cell_text(row, 4, 'nobr')    # 변경일
                    date_digits = re.sub(r"\D", "", date_raw)
                    if len(date_digits) < 8:
                        tax_log("w", "hometax", f"변경일 파싱 생략(형식 불명): 번호={code!r} raw={date_raw!r}")
                        continue
                    date_compact = date_digits[:8]
                    changed_date = date_raw
                    posts.append(
                        {
                            "code": f"{date_compact}_{code}",
                            "number": code,
                            "title": title,
                            "changed_date": changed_date,
                        }
                    )
                except Exception as e:
                    tax_log("e", "hometax", f"행 파싱 오류: {e}")

            # 행은 있는데 하나도 파싱되지 않았다면 구조 변경 등으로 수집이 실패한 것이다.
            # 이를 성공으로 처리하면 "새로운 글이 없습니다"로 오인되어 감지가 멈춘 줄 모르게 된다.
            if rows and not posts:
                tax_log("w", "hometax", f"목록 {len(rows)}행 중 파싱된 글이 없습니다. 조회 실패로 처리합니다.")
                return []

            self._fetch_ok = True
            return posts
        except (NoSuchElementException, TimeoutException):
            tax_log("w", "hometax", f"필수 요소를 찾을 수 없습니다(로딩 지연/구조 변경). 30초 내 그리드가 나타나지 않았습니다.")
            return []
        except Exception as e:
            tax_log("e", "hometax", f"모니터링 오류: {e}")
            return []
        finally:
            if self.driver:
                self.driver.quit()
                self.driver = None

    def check_update(self):
        self.updated = False
        posts = self.fetch_latest_posts()
        
        latest_posts = {}
        for post in posts:
            post_code = post["code"]
            post_title = post["title"]
            latest_posts[post_code] = {
                "number": post.get("number", ""),
                "title": post_title,
                "changed_date": post.get("changed_date", ""),
            }

            if post_code not in self.known_posts:
                tax_log("i", "hometax", f"새 글 발견: {post_title} (코드: {post_code})")
                enqueue_notification(
                    "HomeTax", post_title, self.url, site_key="hometax",
                    meta=[("변경일", post.get("changed_date", ""))],
                )
                self.updated = True

        if latest_posts:
            self.known_posts = latest_posts
        if self._fetch_ok and not self.updated:
            tax_log("i", "hometax", "새로운 글이 없습니다.")
        return self.updated, self.known_posts

# --- 기획재정부 스크래핑 ---
class MoefScraper:
    def __init__(self, shared_data):
        self.url = "https://www.moef.go.kr/lw/lap/TbPrvntcList.do?bbsId=MOSFBBS_000000000055&menuNo=7050300&searchCondition3=1&searchKeyword3=%EC%86%8C%EB%93%9D%EC%84%B8"
        self.known_posts = shared_data.get("moef", {})
        self.session = make_http_session()

    def fetch_latest_posts(self):
        self._fetch_ok = False
        try:
            res = self.session.get(self.url, timeout=15)
            res.raise_for_status()
            soup = BeautifulSoup(res.text, "html.parser")

            posts = soup.select("ul.boardType3 li")
            new_posts = []
            
            for post in posts:
                link = post.find("a")
                if not link:
                    continue
                href = link.get("href") or ""
                mid = re.search(r'fn_egov_select\s*\(\s*["\']([^"\']+)["\']', href)
                if mid:
                    post_id = mid.group(1)
                else:
                    parts = href.split('"')
                    post_id = parts[1] if len(parts) > 1 else None
                if not post_id:
                    continue
                title = link.get_text(strip=True)
                date_el = post.select_one(".boardInfo .infoLeft span.date")
                date_str = ""
                if date_el:
                    date_str = re.sub(r"\s+", " ", date_el.get_text(separator=" ", strip=True))
                    date_str = _moef_notice_period_strip(date_str)
                depart_el = post.select_one(".boardInfo .infoLeft span.depart")
                depart_str = depart_el.get_text(strip=True) if depart_el else ""
                new_posts.append(
                    {"code": post_id, "title": title, "date": date_str, "depart": depart_str}
                )

            self._fetch_ok = True
            return new_posts
        except Exception as e:
            tax_log("e", "moef", f"모니터링 오류: {e}")
            return []

    def check_update(self):
        self.updated = False
        posts = self.fetch_latest_posts()

        latest_posts = {}
        for post in posts:
            post_code = post["code"]
            post_title = post["title"]
            latest_posts[post_code] = {
                "title": post_title,
                "date": post.get("date", ""),
                "depart": post.get("depart", ""),
            }
            if post_code not in self.known_posts:
                tax_log("i", "moef", f"새 글 발견: {post_title} (코드: {post_code})")
                enqueue_notification(
                    "기획재정부", post_title, self.url, site_key="moef",
                    meta=[("예고기간", post.get("date", "")), ("담당", post.get("depart", ""))],
                )
                self.updated = True

        if latest_posts:
            self.known_posts = latest_posts
        if self._fetch_ok and not self.updated:
            tax_log("i", "moef", "새로운 글이 없습니다.")
        return self.updated, self.known_posts

# --- 법제처 스크래핑 ---
class MolegScraper:
    def __init__(self, shared_data):
        self.url = "https://www.moleg.go.kr/lawinfo/makingList.mo?mid=a10104010000&pageCnt=10&lsClsCd=&cptOfiOrgCd=&keyField=lmNm&keyWord=%EC%86%8C%EB%93%9D%EC%84%B8%EB%B2%95&stYdFmt=&edYdFmt="
        self.known_posts = shared_data.get("moleg", {})
        self.session = make_http_session()

    def fetch_latest_posts(self):
        self._fetch_ok = False
        try:
            res = self.session.get(self.url, timeout=15)
            res.raise_for_status()
            soup = BeautifulSoup(res.text, "html.parser")

            rows = soup.select("#listForm table tbody tr")
            new_posts = []

            for tr in rows:
                tds = tr.find_all("td")
                if len(tds) < 5:
                    continue
                law_type = tds[0].get_text(strip=True)
                a = tds[1].find("a")
                if not a:
                    continue
                href = a.get("href", "")
                title = (a.get("title") or "").strip() or a.get_text(strip=True)
                if not href or not title:
                    continue
                law_seq = self._extract_law_seq(href)
                if not law_seq:
                    continue
                ministry = tds[2].get_text(strip=True)
                start_date = tds[3].get_text(strip=True)
                end_date = tds[4].get_text(strip=True)
                new_posts.append(
                    {
                        "code": law_seq,
                        "law_type": law_type,
                        "title": title,
                        "ministry": ministry,
                        "start_date": start_date,
                        "end_date": end_date,
                    }
                )

            self._fetch_ok = True
            return new_posts
        except Exception as e:
            tax_log("e", "moleg", f"모니터링 오류: {e}")
            return []

    def check_update(self):
        self.updated = False
        posts = self.fetch_latest_posts()
        
        latest_posts = {}
        for post in posts:
            post_code = post["code"]
            post_title = post["title"]
            latest_posts[post_code] = {
                "law_type": post.get("law_type", ""),
                "title": post_title,
                "ministry": post.get("ministry", ""),
                "start_date": post.get("start_date", ""),
                "end_date": post.get("end_date", ""),
            }
            if post_code not in self.known_posts:
                tax_log("i", "moleg", f"새 글 발견: {post_title} (코드: {post_code})")
                sd = (post.get("start_date") or "").strip()
                ed = (post.get("end_date") or "").strip()
                period = f"{sd}~{ed}" if sd and ed else (sd or ed)
                enqueue_notification(
                    "법제처", post_title, self.url, site_key="moleg",
                    meta=[("예고기간", period)],
                )
                self.updated = True

        if latest_posts:
            self.known_posts = latest_posts
        if self._fetch_ok and not self.updated:
            tax_log("i", "moleg", "새로운 글이 없습니다.")
        return self.updated, self.known_posts

    def _extract_law_seq(self, href: str) -> str:
        match = re.search(r"lawSeq=(\d+)", href)
        return match.group(1) if match else None

# --- 관보 스크래핑 ---
class GwanboScraper:
    def __init__(self, shared_data):
        self.url = "https://gwanbo.go.kr/SearchRestApi.jsp"
        # 웹 기본검색(searchKeyword.do)은 GET pKeyword 로 입력·자동검색 트리거됨 (hidden #pKeyword)
        self.search_keyword = "소득세법"
        self.headers = {
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        }
        self.known_posts = shared_data.get("gwanbo", {})
        self.updated = False
        self.session = make_http_session()

    def keyword_search_page_url(self):
        base = "https://gwanbo.go.kr/user/search/searchKeyword.do"
        return f"{base}?{urlencode({'pKeyword': self.search_keyword})}"

    def fetch_latest_posts(self):
        self._fetch_ok = False
        kw = self.search_keyword
        try:
            response = self.session.post(self.url, headers=self.headers, data={
                "mode": "keyword",
                "index": "gwanbo",
                "query": f"(unstored_field_subject:({kw})) AND keyword_category_order:(@@ORDER_NUM)",
                "pQuery_tmp": kw,
                "pageNo": "1",
                "listSize": "5",
                "sort": ""
            }, timeout=15)
            response.raise_for_status()
            json_data = response.json()
            new_posts = []
            for item in json_data["data"]:
                category_name = item.get("category_name", "")
                if not item.get("list"):
                    continue
                for entry in item["list"]:
                    raw_subject = entry.get("stored_field_subject", "")
                    subject = re.sub(r"<[^>]*>", "", raw_subject).strip()
                    key = entry.get("search_key", "")
                    reg = (entry.get("keyword_field_regdate") or "").strip()
                    if not reg:
                        y = (entry.get("stored_field_year") or "").strip()
                        m = (entry.get("stored_field_month") or "").strip()
                        d = (entry.get("stored_field_day") or "").strip()
                        if y and m and d:
                            reg = f"{y}{m.zfill(2)}{d.zfill(2)}"
                    new_posts.append(
                        {
                            "code": key,
                            "title": f"{category_name}-{subject}",
                            "published_date": reg,
                        }
                    )

            self._fetch_ok = True
            return new_posts

        except Exception as e:
            tax_log("e", "gwanbo", f"모니터링 오류: {e}")
            return []

    def check_update(self):
        self.updated = False
        posts = self.fetch_latest_posts()
        
        latest_posts = {}
        for post in posts:
            post_code = post["code"]
            post_title = post["title"]
            latest_posts[post_code] = {
                "title": post_title,
                "published_date": post.get("published_date", ""),
            }
            if post_code not in self.known_posts:
                tax_log("i", "gwanbo", f"새 글 발견: {post_title} (코드: {post_code})")
                enqueue_notification(
                    "대한민국 전자관보", post_title, self.keyword_search_page_url(), site_key="gwanbo",
                    meta=[("발행일", _gwanbo_published_display(post.get("published_date", "")))],
                )
                self.updated = True

        if latest_posts:
            self.known_posts = latest_posts
        if self._fetch_ok and not self.updated:
            tax_log("i", "gwanbo", "새로운 글이 없습니다.")
        return self.updated, self.known_posts

# --- 행정안전부 스크래핑 ---
class MoisScraper:
    def __init__(self, shared_data):
        self.url = "https://www.mois.go.kr/frt/bbs/type001/commonSelectBoardList.do?bbsId=BBSMSTR_000000000052"
        self.search_keyword = "행정기관(행정동) 및 관할구역(법정동)"
        self.known_posts = shared_data.get("mois", {})
        self.session = make_http_session()

    def keyword_search_page_url(self):
        return self.url

    def fetch_latest_posts(self):
        self._fetch_ok = False
        try:
            res = self.session.get(self.url, timeout=15)
            res.raise_for_status()
            soup = BeautifulSoup(res.text, "html.parser")

            rows = soup.select("table.table_style1.mobile tbody tr")
            new_posts = []

            for tr in rows:
                tds = tr.find_all("td")
                if len(tds) < 5:
                    continue
                a = tds[1].find("a")
                if not a:
                    continue
                href = a.get("href", "")
                title = a.get_text(strip=True)
                if not href or not title:
                    continue
                if self.search_keyword and self.search_keyword not in title:
                    continue
                ntt_id = self._extract_ntt_id(href)
                if not ntt_id:
                    continue
                date = tds[4].get_text(strip=True).rstrip(".") if len(tds) > 4 else ""
                new_posts.append({
                    "code": ntt_id,
                    "title": title,
                    "date": date,
                })

            self._fetch_ok = True
            return new_posts
        except Exception as e:
            tax_log("e", "mois", f"모니터링 오류: {e}")
            return []

    def check_update(self):
        self.updated = False
        posts = self.fetch_latest_posts()

        latest_posts = {}
        for post in posts:
            post_code = post["code"]
            post_title = post["title"]
            latest_posts[post_code] = {
                "title": post_title,
                "date": post.get("date", ""),
            }
            if post_code not in self.known_posts:
                tax_log("i", "mois", f"새 글 발견: {post_title} (코드: {post_code})")
                enqueue_notification(
                    "행정안전부", post_title, self.keyword_search_page_url(), site_key="mois",
                    meta=[("등록일", post.get("date", ""))],
                )
                self.updated = True

        if latest_posts:
            self.known_posts = latest_posts
        if self._fetch_ok and not self.updated:
            tax_log("i", "mois", "새로운 글이 없습니다.")
        return self.updated, self.known_posts

    def _extract_ntt_id(self, href: str) -> str:
        match = re.search(r"nttId=(\d+)", href)
        return match.group(1) if match else None


# --- 매니저 클래스 ---
class ScraperManager:
    def __init__(self):
        self.shared_data = load_data()
        self.hometax_scraper = HomeTaxScraper(self.shared_data)
        self.moef_scraper = MoefScraper(self.shared_data)
        self.moleg_scraper = MolegScraper(self.shared_data)
        self.gwanbo_scraper = GwanboScraper(self.shared_data)
        self.mois_scraper = MoisScraper(self.shared_data)

    def check_all(self):
        def _run(scraper, key):
            s_updated, s_data = scraper.check_update()
            if not getattr(scraper, "_fetch_ok", True):
                tax_log("i", key, "조회 실패, 3초 후 재시도...")
                time.sleep(3)
                s_updated, s_data = scraper.check_update()
            return s_updated, s_data

        ht_updated, ht_data = _run(self.hometax_scraper, "hometax")
        mf_updated, mf_data = _run(self.moef_scraper,    "moef")
        ml_updated, ml_data = _run(self.moleg_scraper,   "moleg")
        gb_updated, gb_data = _run(self.gwanbo_scraper,  "gwanbo")
        mo_updated, mo_data = _run(self.mois_scraper,    "mois")

        if ht_updated or mf_updated or ml_updated or gb_updated or mo_updated:
            self.shared_data["hometax"] = ht_data
            self.shared_data["moef"] = mf_data
            self.shared_data["moleg"] = ml_data
            self.shared_data["gwanbo"] = gb_data
            self.shared_data["mois"] = mo_data
            save_data(self.shared_data)
            tax_log("i", "", "통합 데이터 파일 업데이트 완료")

# --- 트레이 아이콘 세팅 ---
def create_image():
    image = Image.new('RGB', (64, 64), color=(0, 102, 204))
    d = ImageDraw.Draw(image)
    d.ellipse((8, 8, 56, 56), fill=(255, 255, 255))
    return image

# --- LOG 파일 미리보기 함수 ---
def open_log_file():
    try:
        os.startfile(f"{os.getcwd()}/{LOG}")
        tax_log("i", "System", f"로그 파일을 열었습니다.")
        
    except Exception as e:
        tax_log("e", "System", f"로그 파일 열기 실패: {e}")


def open_url_tray(url: str):
    """트레이 메뉴에서 모니터링 대상 페이지를 기본 브라우저로 연다."""
    def _open(icon, item):
        try:
            webbrowser.open(url)
            tax_log("i", "System", "모니터링 페이지를 열었습니다.")
        except Exception as e:
            tax_log("e", "System", f"브라우저에서 페이지 열기 실패: {e}")

    return _open


def _gwanbo_published_display(reg: str) -> str:
    """keyword_field_regdate(YYYYMMDD) 등을 화면용 날짜 문자열로."""
    reg = (reg or "").strip()
    if re.fullmatch(r"\d{8}", reg):
        return f"{reg[:4]}-{reg[4:6]}-{reg[6:8]}"
    return reg


def _format_gwanbo_overview_html(posts):
    """전자관보: 카테고리별로 목차|발행일 테이블."""
    if not posts:
        return (
            '<div class="table-wrap gwanbo-table-wrap">'
            '<table class="board-table gwanbo-table">'
            "<thead><tr><th class=\"th-gwanbo-cat\">목차</th><th>발행일</th></tr></thead>"
            "<tbody><tr><td colspan=\"2\" class=\"empty-cell\">저장된 목록이 없습니다.</td></tr></tbody>"
            "</table></div>"
        )

    groups = []
    cat_index = {}
    for _code, raw in posts.items():
        ent = _gwanbo_entry_from_value(raw)
        full_title = ent["title"]
        pub_raw = ent["published_date"]
        pub_disp = _gwanbo_published_display(pub_raw)
        pub_cell = escape(pub_disp) if pub_disp else "—"
        if "-" in full_title:
            cat, rest = full_title.split("-", 1)
            title_only = rest.strip() or full_title
        else:
            cat = "기타"
            title_only = full_title
        if cat not in cat_index:
            cat_index[cat] = len(groups)
            groups.append({"cat": cat, "items": []})
        groups[cat_index[cat]]["items"].append(
            {"title_only": title_only, "pub_cell": pub_cell}
        )

    parts = []
    for g in groups:
        rows = []
        for it in g["items"]:
            rows.append(
                "<tr>"
                f"<td class=\"col-gwanbo-toc\">{escape(it['title_only'])}</td>"
                f"<td class=\"col-gwanbo-pub\">{it['pub_cell']}</td>"
                "</tr>"
            )
        parts.append(
            f'<div class="gwanbo-group">'
            '<div class="table-wrap gwanbo-table-wrap">'
            '<table class="board-table gwanbo-table">'
            f"<thead><tr><th class=\"th-gwanbo-cat\">{escape(g['cat'])}</th><th>발행일</th></tr></thead>"
            f"<tbody>{''.join(rows)}</tbody>"
            "</table></div></div>"
        )
    return f'<div class="gwanbo-overview">{"".join(parts)}</div>'


def _format_moef_overview_html(posts):
    if not posts:
        return (
            '<div class="table-wrap moef-table-wrap">'
            '<table class="board-table moef-table">'
            "<thead><tr><th>예고명</th><th>예고기간</th><th>담당</th></tr></thead>"
            "<tbody><tr><td colspan=\"3\" class=\"empty-cell\">저장된 목록이 없습니다.</td></tr></tbody>"
            "</table></div>"
        )
    body_rows = []
    for code, raw in posts.items():
        ent = _moef_entry_from_value(raw)
        title = escape(ent["title"]) or "—"
        date_v = escape(ent["date"]) if ent["date"] else "—"
        depart_v = escape(ent["depart"]) if ent["depart"] else "—"
        body_rows.append(
            f'<tr data-code="{escape(code)}">'
            f'<td class="col-moef-title">{title}</td>'
            f'<td class="col-moef-period">{date_v}</td>'
            f'<td class="col-moef-depart">{depart_v}</td>'
            "</tr>"
        )
    return (
        '<div class="table-wrap moef-table-wrap">'
        '<table class="board-table moef-table">'
        "<thead><tr><th>예고명</th><th>기간</th><th>담당</th></tr></thead>"
        f"<tbody>{''.join(body_rows)}</tbody>"
        "</table></div>"
    )


def _format_hometax_overview_html(posts):
    if not posts:
        return (
            '<div class="table-wrap hometax-table-wrap">'
            '<table class="board-table hometax-table">'
            "<thead><tr><th>번호</th><th>제목</th><th>변경일</th></tr></thead>"
            "<tbody><tr><td colspan=\"3\" class=\"empty-cell\">저장된 목록이 없습니다.</td></tr></tbody>"
            "</table></div>"
        )
    body_rows = []
    for row_key, raw in posts.items():
        ent = _hometax_entry_from_value(row_key, raw)
        num = escape(ent["number"]) if ent["number"] else "—"
        title = escape(ent["title"]) if ent["title"] else "—"
        chg = escape(ent["changed_date"]) if ent["changed_date"] else "—"
        body_rows.append(
            f'<tr data-key="{escape(row_key)}">'
            f"<td class=\"col-num\">{num}</td>"
            f"<td class=\"col-title\">{title}</td>"
            f"<td class=\"col-date\">{chg}</td>"
            "</tr>"
        )
    return (
        '<div class="table-wrap hometax-table-wrap">'
        '<table class="board-table hometax-table">'
        "<thead><tr><th>번호</th><th>제목</th><th>변경일</th></tr></thead>"
        f"<tbody>{''.join(body_rows)}</tbody>"
        "</table></div>"
    )


def _format_moleg_overview_html(posts):
    if not posts:
        return (
            '<div class="table-wrap moleg-table-wrap">'
            '<table class="board-table moleg-table">'
            "<thead><tr><th>법령종류</th><th>입법예고명</th><th>소관부처</th><th>기간</th></tr></thead>"
            "<tbody><tr><td colspan=\"4\" class=\"empty-cell\">저장된 목록이 없습니다.</td></tr></tbody>"
            "</table></div>"
        )
    body_rows = []
    for code, raw in posts.items():
        ent = _moleg_entry_from_value(raw)
        lt = escape(ent["law_type"]) if ent["law_type"] else "—"
        title = escape(ent["title"]) if ent["title"] else "—"
        ministry = escape(ent["ministry"]) if ent["ministry"] else "—"
        sd_raw = (ent["start_date"] or "").strip()
        ed_raw = (ent["end_date"] or "").strip()
        if sd_raw and ed_raw:
            period = f"{escape(sd_raw)}~{escape(ed_raw)}"
        elif sd_raw or ed_raw:
            period = escape(sd_raw or ed_raw)
        else:
            period = "—"
        body_rows.append(
            f'<tr data-code="{escape(code)}">'
            f'<td class="col-law-type">{lt}</td>'
            f'<td class="col-law-title">{title}</td>'
            f'<td class="col-ministry">{ministry}</td>'
            f'<td class="col-period">{period}</td>'
            "</tr>"
        )
    return (
        '<div class="table-wrap moleg-table-wrap">'
        '<table class="board-table moleg-table">'
        "<thead><tr><th>법령종류</th><th>입법예고명</th><th>소관부처</th><th>기간</th></tr></thead>"
        f"<tbody>{''.join(body_rows)}</tbody>"
        "</table></div>"
    )


def _format_mois_overview_html(posts):
    if not posts:
        return (
            '<div class="table-wrap mois-table-wrap">'
            '<table class="board-table mois-table">'
            "<thead><tr><th>제목</th><th>등록일</th></tr></thead>"
            "<tbody><tr><td colspan=\"2\" class=\"empty-cell\">저장된 목록이 없습니다.</td></tr></tbody>"
            "</table></div>"
        )
    body_rows = []
    for code, raw in posts.items():
        ent = _mois_entry_from_value(raw)
        title = escape(ent["title"]) if ent["title"] else "—"
        date_v = escape(ent["date"]) if ent["date"] else "—"
        body_rows.append(
            f'<tr data-code="{escape(code)}">'
            f'<td class="col-title">{title}</td>'
            f'<td class="col-date">{date_v}</td>'
            "</tr>"
        )
    return (
        '<div class="table-wrap mois-table-wrap">'
        '<table class="board-table mois-table">'
        "<thead><tr><th>제목</th><th>등록일</th></tr></thead>"
        f"<tbody>{''.join(body_rows)}</tbody>"
        "</table></div>"
    )


def show_board_overview(icon, item):
    """HTML 파일을 생성해 4개 모니터링 목록을 브라우저로 연다."""
    try:
        base_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
        data = load_data()
        sections = [
            ("홈택스", "hometax", manager.hometax_scraper.url),
            ("기획재정부", "moef", manager.moef_scraper.url),
            ("법제처", "moleg", manager.moleg_scraper.url),
            ("전자관보", "gwanbo", manager.gwanbo_scraper.keyword_search_page_url()),
            ("행정안전부", "mois", manager.mois_scraper.keyword_search_page_url()),
        ]

        cards = []
        for title, key, url in sections:
            posts = data.get(key, {})
            if key == "gwanbo":
                list_block = _format_gwanbo_overview_html(posts)
            elif key == "hometax":
                list_block = _format_hometax_overview_html(posts)
            elif key == "moef":
                list_block = _format_moef_overview_html(posts)
            elif key == "moleg":
                list_block = _format_moleg_overview_html(posts)
            elif key == "mois":
                list_block = _format_mois_overview_html(posts)
            else:
                list_items = []
                if posts:
                    for _idx, (_, post_title) in enumerate(posts.items(), start=1):
                        list_items.append(f"<li>{escape(post_title)}</li>")
                else:
                    list_items.append("<li class='empty'>저장된 목록이 없습니다.</li>")
                list_block = f"<ul>{''.join(list_items)}</ul>"

            local_logo = os.path.join(base_dir, "data", f"{key}_logo.png")
            if os.path.exists(local_logo):
                logo_src = f"{key}_logo.png"
            else:
                logo_src = BOARD_OVERVIEW_CARD_LOGOS.get(key, "")
            if logo_src:
                heading_html = (
                    f'<h2 class="card-site-heading">'
                    f'<img class="card-site-logo" src="{escape(logo_src)}" alt="{escape(title)}" '
                    'loading="lazy" decoding="async" />'
                    f"</h2>"
                )
            else:
                heading_html = f"<h2>{escape(title)}</h2>"

            cards.append(
                f"""
                <section class="card">
                    <div class="head">
                        {heading_html}
                        <a href="{escape(url)}" target="_blank" rel="noopener noreferrer">페이지 열기</a>
                    </div>
                    {list_block}
                </section>
                """
            )

        rendered_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        html_doc = f"""<!doctype html>
<html lang="ko">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>TaxBoardWatcher</title>
    <style>
        body {{
            margin: 0;
            padding: clamp(12px, 3vw, 28px);
            background: #f4f6f8;
            color: #222;
            font-family: "Malgun Gothic", "Segoe UI", Arial, sans-serif;
        }}
        .container {{
            width: min(100%, 700px);
            margin: 0 auto;
        }}
        h1 {{
            margin: 0 0 8px 0;
            font-size: 26px;
        }}
        .sub {{
            margin-bottom: 16px;
            color: #666;
            font-size: 13px;
        }}
        .grid {{
            display: grid;
            grid-template-columns: 1fr;
            gap: 14px;
        }}
        .card {{
            background: #fff;
            border: 1px solid #d8dde3;
            border-radius: 10px;
            padding: 14px;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.03);
        }}
        .head {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 8px;
        }}
        .head h2 {{
            margin: 0;
            display: flex;
            align-items: center;
            min-height: 36px;
        }}
        .head h2.card-site-heading {{
            font-size: 0;
            line-height: 0;
        }}
        .head h2 .card-site-logo {{
            display: block;
            height: 34px;
            width: auto;
            max-width: min(260px, 55vw);
            object-fit: contain;
            object-position: left center;
        }}
        .head h2:not(.card-site-heading) {{
            font-size: 18px;
        }}
        .head a {{
            background: #0078d7;
            color: #fff;
            text-decoration: none;
            border-radius: 6px;
            padding: 5px 10px;
            font-size: 12px;
            white-space: nowrap;
        }}
        ul {{
            margin: 10px 0 0 0;
            padding-left: 18px;
        }}
        .gwanbo-overview {{
            margin-top: 10px;
        }}
        .gwanbo-group .gwanbo-table-wrap {{
            margin-top: 0;
        }}
        table.gwanbo-table thead th.th-gwanbo-cat {{
            font-weight: 700;
            min-width: 10em;
            word-break: keep-all;
        }}
        table.gwanbo-table .col-gwanbo-toc {{
            min-width: 12em;
        }}
        table.gwanbo-table .col-gwanbo-pub {{
            white-space: nowrap;
            width: 1%;
            font-variant-numeric: tabular-nums;
            color: #444;
        }}
        .gwanbo-overview .gwanbo-group + .gwanbo-group {{
            margin-top: 12px;
        }}
        li {{
            margin: 5px 0;
            line-height: 1.45;
            word-break: break-word;
            font-size: 12px;
        }}
        .empty {{
            color: #777;
            list-style: none;
            margin-left: -18px;
        }}
        .item-meta {{
            margin: 0;
            display: grid;
            grid-template-columns: 88px 1fr;
            gap: 4px 10px;
            font-size: 12px;
            align-items: start;
        }}
        .item-meta dd {{
            margin: 0;
            color: #222;
            line-height: 1.45;
            word-break: break-word;
        }}
        .item-meta dd.emph {{
            font-weight: 700;
        }}
        .table-wrap {{
            margin-top: 10px;
            overflow-x: auto;
            -webkit-overflow-scrolling: touch;
            border: 1px solid #e2e6eb;
            border-radius: 8px;
            background: #fafbfc;
        }}
        table.board-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 12px;
        }}
        table.board-table th,
        table.board-table td {{
            border-bottom: 1px solid #e8ecf0;
            padding: 8px 10px;
            text-align: left;
            vertical-align: middle;
            line-height: 1.45;
            word-break: break-word;
        }}
        table.board-table thead th {{
            background: #eef2f6;
            font-weight: 700;
            color: #333;
            white-space: nowrap;
            text-align: center;
        }}
        table.board-table tbody tr:last-child td {{
            border-bottom: none;
        }}
        table.board-table .col-num {{
            white-space: nowrap;
            width: 1%;
            color: #444;
        }}
        table.board-table .col-title {{
            min-width: 12em;
        }}
        table.board-table .col-date {{
            white-space: nowrap;
            width: 1%;
            color: #444;
        }}
        table.board-table .empty-cell {{
            text-align: center;
            color: #777;
            padding: 14px 10px;
        }}
        table.moef-table .col-moef-title {{
            min-width: 12em;
        }}
        table.moef-table .col-moef-period {{
            min-width: 10em;
            color: #333;
        }}
        table.moef-table .col-moef-depart {{
            white-space: nowrap;
            width: 1%;
            color: #444;
        }}
        table.board-table .col-law-type {{
            white-space: nowrap;
            width: 1%;
            color: #444;
        }}
        table.board-table .col-law-title {{
            min-width: 10em;
        }}
        table.board-table .col-ministry {{
            white-space: nowrap;
            width: 1%;
            color: #444;
        }}
        table.board-table .col-period {{
            white-space: nowrap;
            width: 1%;
            font-variant-numeric: tabular-nums;
            color: #444;
        }}
        .item-meta dt {{
            margin: 0;
            color: #5a6570;
            font-weight: 600;
        }}
        .meta-empty {{
            color: #999;
        }}
        @media (max-width: 760px) {{
            .container {{
                min-width: 0;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>TaxBoardWatcher - 게시물 목록</h1>
        <div class="sub">생성 시각: {escape(rendered_at)}</div>
        <div class="grid">
            {''.join(cards)}
        </div>
    </div>
</body>
</html>
"""

        output_dir = os.path.join(base_dir, "data")
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, "board_overview.html")
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html_doc)

        webbrowser.open(f"file:///{output_path.replace(os.sep, '/')}")
        tax_log("i", "System", f"게시물 목록 열기: {output_path}")
    except Exception as e:
        tax_log("e", "System", f"목록 통합 HTML 생성/열기 실패: {e}")


def on_exit(icon, item):
    tax_log("i", "", f"사용자 요청으로 종료")
    icon.stop()
    os._exit(0)

# --- 자동 업데이트 ---
_pending_update = None       # {"version", "url", "size"} — 설치 가능한 새 버전
_update_notified = None      # 토스트로 이미 안내한 버전(같은 버전 반복 안내 방지)
_last_update_check = 0.0
_update_lock = threading.Lock()


def _parse_version(s):
    """'v.5.0.2' / 'v5.0.2' 등을 (5, 0, 2)로 변환."""
    return tuple(int(n) for n in re.findall(r"\d+", s or ""))


def check_for_update(manual=False):
    """GitHub 최신 릴리스를 조회해 현재 버전보다 높으면 설치 대기 상태로 둔다."""
    global _pending_update, _update_notified, _last_update_check
    if not _update_lock.acquire(blocking=False):
        return
    try:
        _last_update_check = time.time()
        api = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        resp = requests.get(api, headers={"Accept": "application/vnd.github+json"}, timeout=15)
        if resp.status_code != 200:
            raise RuntimeError(f"릴리스 조회 실패(HTTP {resp.status_code})")
        release = resp.json()
        latest = release.get("tag_name", "")
        asset = next((a for a in release.get("assets", []) if a.get("name") == UPDATE_ASSET_NAME), None)

        if not asset or _parse_version(latest) <= _parse_version(VERSION):
            _pending_update = None
            if manual:
                notify_manual_monitor_toast("업데이트 확인", f"최신 버전을 사용 중입니다. ({VERSION})")
            return

        _pending_update = {
            "version": latest,
            "url": asset["browser_download_url"],
            "size": asset.get("size", 0),
        }
        tax_log("i", "", f"새 버전 발견: {latest} (현재 {VERSION})")
        if tray_icon is not None:
            tray_icon.update_menu()
        if manual or _update_notified != latest:
            _update_notified = latest
            notify_manual_monitor_toast(
                "새 버전이 있습니다",
                f"{latest} 업데이트가 있습니다. 트레이 메뉴의 '업데이트 설치'를 눌러 주세요.",
            )
    except Exception as e:
        tax_log("w", "", f"업데이트 확인 실패: {e}")
        if manual:
            notify_manual_monitor_toast("업데이트 확인 실패", str(e), error=True)
    finally:
        _update_lock.release()


def check_for_update_if_due():
    """모니터링 주기마다 호출. 마지막 확인 후 UPDATE_CHECK_INTERVAL이 지났을 때만 조회한다."""
    if time.time() - _last_update_check >= UPDATE_CHECK_INTERVAL:
        check_for_update()


def _ps_quote(s):
    return "'" + str(s).replace("'", "''") + "'"


def apply_update():
    """새 exe를 내려받아 두고, 프로그램 종료 후 교체·재시작하는 PowerShell을 띄운 뒤 종료한다.
    실행 중인 exe는 자기 자신을 덮어쓸 수 없으므로 교체는 외부 프로세스가 담당한다.
    data/ 폴더(게시글 데이터·설정·로그)는 건드리지 않는다."""
    upd = _pending_update
    if not upd:
        return
    if not getattr(sys, "frozen", False):
        notify_manual_monitor_toast("업데이트 불가", "exe로 실행 중일 때만 자동 업데이트할 수 있습니다.", error=True)
        return

    exe_path = sys.executable
    base_dir = os.path.dirname(exe_path)
    new_path = exe_path + ".new"
    try:
        tax_log("i", "", f"업데이트 다운로드 시작: {upd['version']}")
        with requests.get(upd["url"], stream=True, timeout=60) as r:
            r.raise_for_status()
            with open(new_path, "wb") as f:
                for chunk in r.iter_content(chunk_size=1024 * 256):
                    f.write(chunk)
        size = os.path.getsize(new_path)
        with open(new_path, "rb") as f:
            is_exe = f.read(2) == b"MZ"
        if not is_exe or (upd["size"] and size != upd["size"]):
            raise RuntimeError(f"다운로드 파일 검증 실패(size={size})")
    except Exception as e:
        tax_log("e", "", f"업데이트 다운로드 실패: {e}")
        notify_manual_monitor_toast("업데이트 실패", str(e), error=True)
        try:
            os.remove(new_path)
        except OSError:
            pass
        return

    # PyInstaller onefile은 부트로더(부모)와 파이썬(자식) 두 프로세스가 exe를 잡고 있으므로 둘 다 종료를 기다린다.
    pids = ",".join(str(p) for p in {os.getpid(), os.getppid()})
    script = f"""
$ErrorActionPreference = 'SilentlyContinue'
Wait-Process -Id {pids} -Timeout 60
$ok = $false
for ($i = 0; $i -lt 30; $i++) {{
    try {{ Move-Item -LiteralPath {_ps_quote(new_path)} -Destination {_ps_quote(exe_path)} -Force -ErrorAction Stop; $ok = $true; break }}
    catch {{ Start-Sleep -Seconds 1 }}
}}
Start-Process -FilePath {_ps_quote(exe_path)} -WorkingDirectory {_ps_quote(base_dir)}
"""
    import base64
    import subprocess
    encoded = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
    cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-EncodedCommand", encoded]
    # DETACHED_PROCESS로 띄우면 PowerShell이 콘솔 없이 시작돼 바로 종료되므로 CREATE_NO_WINDOW만 쓴다.
    flags = subprocess.CREATE_NO_WINDOW
    # 이 프로세스가 job object에 속해 있으면 종료 시 자식도 함께 정리되므로 job에서 분리해 띄운다.
    try:
        subprocess.Popen(cmd, creationflags=flags | subprocess.CREATE_BREAKAWAY_FROM_JOB, close_fds=True)
    except OSError:
        subprocess.Popen(cmd, creationflags=flags, close_fds=True)
    tax_log("i", "", f"업데이트 설치를 위해 종료합니다: {VERSION} -> {upd['version']}")
    if tray_icon is not None:
        tray_icon.stop()
    os._exit(0)


def on_update_menu(icon, item):
    """트레이 메뉴: 설치 대기 중인 버전이 있으면 설치, 없으면 수동 확인."""
    if _pending_update:
        threading.Thread(target=apply_update, daemon=True).start()
    else:
        threading.Thread(target=check_for_update, kwargs={"manual": True}, daemon=True).start()

# --- 모니터링 쓰레드 ---
def run_monitor():
    global tray_icon
    # 최초 실행시 즉시 한 번 체크
    tax_log("i", "", f"프로그램 시작, 최초 모니터링 실행")
    manager.check_all()
    check_for_update_if_due()
    while True:
        next_check_time = datetime.datetime.now() + datetime.timedelta(seconds=SLEEP_TIME)
        tax_log("i", "", f"다음 모니터링 예정 시간: {next_check_time.strftime('%Y-%m-%d %H:%M:%S')}")
        if tray_icon is not None:
            tray_icon.title = f"다음 모니터링 예정 시간: {next_check_time.strftime('%H시 %M분')}"
        interrupted = _interval_changed.wait(SLEEP_TIME)
        _interval_changed.clear()
        if interrupted:
            tax_log("i", "", f"모니터링 주기 변경, 타이머 재시작")
            continue
        tax_log("i", "", f"시간 경과, 모니터링 시작")
        manager.check_all()
        check_for_update_if_due()

# --- 시작프로그램'에 바로가기(.lnk) 파일 생성(최초 1회만 생성) --- 
def add_to_startup():
    startup_dir = os.path.join(os.getenv('APPDATA'), r"Microsoft\Windows\Start Menu\Programs\Startup")

    # Startup 폴더가 존재하는지 확인하고 없으면 생성
    if not os.path.exists(startup_dir):
        os.makedirs(startup_dir)

    shortcut_path = os.path.join(startup_dir, "TaxBoardWatcher.lnk")

    if not os.path.exists(shortcut_path):
        # 바로가기 생성
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.TargetPath = sys.executable  # 현재 실행 중인 exe 경로
        shortcut.WorkingDirectory = os.getcwd()
        shortcut.IconLocation = sys.executable
        shortcut.save()
        tax_log("i", "", f"시작프로그램 등록 완료: {shortcut_path}")
    # else:
        # tax_log("i", "", f"시작프로그램에 이미 등록되어 있습니다.")

def asciiart():
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(rf"=================================================================================================== {VERSION}""\n")
        f.write(r" _____                 ______                          _     _    _         _          _                  ""\n")
        f.write(r"|_   _|                | ___ \                        | |   | |  | |       | |        | |                 ""\n")
        f.write(r"  | |    __ _ __  __   | |_/ /  ___    __ _  _ __   __| |   | |  | |  __ _ | |_   ___ | |__    ___  _ __  ""\n")
        f.write(r"  | |   / _` |\ \/ /   | ___ \ / _ \  / _` || '__| / _` |   | |/\| | / _` || __| / __|| '_ \  / _ \| '__| ""\n")
        f.write(r"  | |  | (_| | >  <    | |_/ /| (_) || (_| || |   | (_| |   \  /\  /| (_| || |_ | (__ | | | ||  __/| |    ""\n")
        f.write(r"  \_/   \__,_|/_/\_\   \____/  \___/  \__,_||_|    \__,_|    \/  \/  \__,_| \__| \___||_| |_| \___||_|    ""\n")
        f.write(r"==========================================================================================================""\n")

# 모니터링 수동 실행
def manual_crawl(icon, item):
    tax_log("i", "", f"수동 모니터링 시작")
    set_tray_icon(SEARCHING_ICON) # 트레이 아이콘 "검색"
    err = None
    try:
        manager.check_all()
    except Exception as e:
        err = e
        tax_log("e", "", f"수동 모니터링 오류: {e}")
    finally:
        set_tray_icon(ICON) # 트레이 아이콘 기본값
        tax_log("i", "", f"수동 모니터링 완료")
    if err is None:
        notify_manual_monitor_toast("수동 모니터링", "모든 점검이 끝났습니다.", error=False)
    else:
        notify_manual_monitor_toast("수동 모니터링 오류", str(err), error=True)

# 트레이 아이콘 변경
def set_tray_icon(image_path):
    try:
        icon_image = Image.open(image_path)
        tray_icon.icon = icon_image
    except Exception as e:
        tax_log("e", "", f"트레이 아이콘 변경 오류: {e}")


# --- 메인 실행 ---
if __name__ == "__main__":
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # Per-Monitor DPI Aware
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()  # fallback (Vista+)
        except Exception:
            pass
    asciiart()
    add_to_startup()
    get_config()
    if not os.path.exists(LOG):
        open(LOG, "w", encoding="utf-8").close()
    if not os.path.exists(BOARD_DATA):
        open(BOARD_DATA, "w", encoding="utf-8").close()

    send_notification("TaxBoardWatcher","start") # 최초 실행 알림창
    manager = ScraperManager()

    notify_thread = threading.Thread(target=_notification_worker, daemon=True)
    notify_thread.start()

    monitor_thread = threading.Thread(target=run_monitor, daemon=True)
    monitor_thread.start()

    tray_icon = pystray.Icon("Web Monitor")
    tray_icon.icon = Image.open(ICON).resize((32, 32))
    tray_icon.title = "TaxBoardWatcher"

    def _make_interval_item(h):
        return pystray.MenuItem(
            f'{h}시간',
            lambda: set_sleep_hour(h),
            checked=lambda item: SLEEP_TIME == h * 3600,
            radio=True,
        )
    interval_submenu = pystray.Menu(*[_make_interval_item(h) for h in [1, 2, 3, 4, 6, 12, 24]])

    def _test_notify(kind):
        if kind == "start":
            threading.Thread(
                target=lambda: send_notification("", "start"),
                daemon=True,
            ).start()
        else:
            _TEST_SITES = [
                ("HomeTax",           "https://hometax.go.kr", "hometax"),
                ("기획재정부",        "https://www.moef.go.kr/lw/lap/TbPrvntcList.do?bbsId=MOSFBBS_000000000055&menuNo=7050300&searchCondition3=1&searchKeyword3=%EC%86%8C%EB%93%9D%EC%84%B8", "moef"),
                ("법제처",            "https://www.moleg.go.kr/lawinfo/makingList.mo?mid=a10104010000&pageCnt=10&lsClsCd=&cptOfiOrgCd=&keyField=lmNm&keyWord=%EC%86%8C%EB%93%9D%EC%84%B8%EB%B2%95&stYdFmt=&edYdFmt=", "moleg"),
                ("대한민국 전자관보", "https://gwanbo.go.kr/user/search/searchKeyword.do?pKeyword=%EC%86%8C%EB%93%9D%EC%84%B8%EB%B2%95", "gwanbo"),
                ("행정안전부",        "https://www.mois.go.kr/frt/bbs/type001/commonSelectBoardList.do?bbsId=BBSMSTR_000000000052", "mois"),
            ]
            _title, _url, _site_key = random.choice(_TEST_SITES)
            threading.Thread(
                target=lambda t=_title, u=_url, sk=_site_key: send_notification(
                    t,
                    "이 프로그램은 국세청, 기획재정부 등 정부의 주요 기관의 게시판을 정해진 주기마다 자동으로 확인하여 새 게시물이 등록되면 사용자에게 알림을 제공합니다.",
                    url=u,
                    site_key=sk,
                    meta=[("등록일", datetime.datetime.now().strftime("%Y-%m-%d"))],
                ),
                daemon=True,
            ).start()

    test_submenu = pystray.Menu(
        pystray.MenuItem('시작 알림', lambda: _test_notify("start")),
        pystray.MenuItem('새 게시물 알림', lambda: _test_notify("new")),
    )

    tray_icon.menu = pystray.Menu(
        pystray.MenuItem(f'TaxBoardWatcher  {VERSION}', None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem('게시물 목록', show_board_overview),
        pystray.MenuItem('모니터링 실행', manual_crawl),
        pystray.MenuItem('모니터링 주기', interval_submenu),
        pystray.MenuItem('로그 열기', open_log_file),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem('알림 테스트', test_submenu),
        pystray.MenuItem(
            lambda item: f"업데이트 설치 ({_pending_update['version']})" if _pending_update else '업데이트 확인',
            on_update_menu,
        ),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem('종료', on_exit),
    )
    tray_icon.run()