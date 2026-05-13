from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from win32com.client import Dispatch
from urllib.parse import urlencode
from PIL import Image, ImageDraw
from selenium import webdriver
from bs4 import BeautifulSoup
from html import escape
import tkinter as tk
import webbrowser
import threading
import datetime
import requests
import pystray
import json
import time
import sys
import re
import os

try:
    from winotify import Notification as WinotifyNotification
except ImportError:
    WinotifyNotification = None

VERSION = "v.4.0.0"

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

# --- config.json 로드 ---
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
        category = "[ 홈텍스 ]"
    elif category == "moef":
        category = "[ 기재부 ]"
    elif category == "moleg":
        category = "[ 법제처 ]"
    elif category == "gwanbo":
        category = "[ 관　보 ]"
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


def send_notification(title, message, url=None):
    try:
        # 알림창 생성
        root = tk.Tk()
        root.withdraw() # 창을 처음엔 숨김 상태로 시작
        root.title(title)
        root.resizable(False, False)
        root.configure(bg="#f0f4f7") # 배경색 설정
        root.attributes("-topmost", True) # 창을 항상 위로
        root.attributes("-alpha", 0.0)  # 처음엔 완전 투명

        # 알림창 아이콘
        icon_img = tk.PhotoImage(file=ICON)
        root.iconphoto(False, icon_img)

        def open_url():
            webbrowser.open(url)
            root.destroy()
        def on_close():
            root.destroy()

        # 새 글 발견시 url이 존재함.
        if url:
            # 느낌표 아이콘 
            canvas = tk.Canvas(root, width=50, height=50, bg="#f0f4f7", highlightthickness=0)
            canvas.create_oval(5, 5, 45, 45, fill="#0078D7", outline="#0078D7")
            canvas.create_text(25, 25, text="!", font=("Arial", 16, "bold"), fill="white")
            canvas.pack(pady=(15, 5))
            # 텍스트 라벨
            label1 = tk.Label(root, text="새로운 글이 등록되었습니다.", bg="#f0f4f7", font=("Arial", 12, "bold"))
            label1.pack()
            label2 = tk.Label(root, text=message, bg="#f0f4f7", font=("Arial", 10), wraplength=300)
            label2.pack(pady=(5, 15))
            # 버튼 영역
            button_frame = tk.Frame(root, bg="#f0f4f7")
            button_frame.pack(pady=(0, 10))
            open_button = tk.Button(button_frame, text="페이지 열기", command=open_url, width=12, bg="#0078D7", fg="white", font=("Arial", 10, "bold"))
            open_button.pack(side="left", padx=10)
            close_button = tk.Button(button_frame, text="닫기", command=on_close, width=12, bg="#cccccc", font=("Arial", 10, "bold"))
            close_button.pack(side="left", padx=10)
        elif message == 'start': # 최초 실행
            global SLEEP_TIME
            # 돋보기 아이콘
            canvas = tk.Canvas(root, width=50, height=50, bg="#f0f4f7", highlightthickness=0)
            canvas.create_oval(10, 10, 30, 30, outline="#0078D7", width=4)
            canvas.create_line(28, 28, 40, 40, fill="#0078D7", width=4)
            canvas.pack(pady=(15, 5))
            # 텍스트 라벨
            label1 = tk.Label(root, text="TaxBoardWatcher를 시작합니다.", bg="#f0f4f7", font=("Arial", 12, "bold"))
            label1.pack()
            label2 = tk.Label(root, text=f"현재 설정된 모니터링 간격은 {int(SLEEP_TIME/60/60)}시간입니다.", bg="#f0f4f7", font=("Arial", 10), wraplength=300)
            label2.pack(pady=(5, 15))
            # 버튼 영역
            button_frame = tk.Frame(root, bg="#f0f4f7")
            button_frame.pack()
            close_button = tk.Button(button_frame, text="시작", command=on_close, width=12, bg="#0078D7", fg="white", font=("Arial", 10, "bold"))
            close_button.pack(side="left", padx=10)

        root.update_idletasks()  # 내부 위젯 크기에 맞게 자동 조정
        root.minsize(350, 200)  # 최소 크기만 지정
        w = max(root.winfo_reqwidth(), 350)
        h = max(root.winfo_reqheight(), 200)
        x = (root.winfo_screenwidth() - w) // 2
        y = (root.winfo_screenheight() - h) // 2
        root.geometry(f"{w}x{h}+{x}+{y}")
        root.deiconify()  # withdraw() 이후 알림창 표시하기 위함

        # Fade-in 애니메이션
        for i in range(0, 11):  # 0.0 ~ 1.0
            root.attributes("-alpha", i / 10)
            root.update()
            time.sleep(0.03)  # 속도 조정 가능 (0.02~0.05 정도가 자연스러움)

        root.mainloop()
    except Exception as e:
        tax_log("e", "", f" 알림 전송 오류: {e}")

# 모니터링 섹션 키 (board_data.json 값은 항상 dict[str, dict] 형태로 통일)
BOARD_DATA_SECTIONS = ("hometax", "moef", "moleg", "gwanbo")


def _normalize_board_data(data):
    """루트 dict의 모니터링 섹션을 문자열·불완전 dict 등 섞여 있어도 동일 스키마로 통일."""
    if not isinstance(data, dict):
        return {"hometax": {}, "moef": {}, "moleg": {}, "gwanbo": {}}
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
            return data
        return {"hometax": {}, "moef": {}, "moleg": {}, "gwanbo": {}}
    except Exception as e:
        tax_log("e", "", f"파일 로드 오류: {e}")
        return {"hometax": {}, "moef": {}, "moleg": {}, "gwanbo": {}}

# --- 홈택스 스크래핑 ---
class HomeTaxScraper:
    # WebDriver 초기화
    def init_driver(self):
        # ChromeOptions 설정
        chrome_options = Options()
        chrome_options.add_argument("--headless")
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
        try:
            self.init_driver()
            tax_log("i", "hometax", f"새 글 찾는 중...")
            self.driver.get(self.url)
            time.sleep(3)  # 페이지 로딩 대기
            
            # tbody 요소 찾기
            tbody = self.driver.find_element("css selector", '#mf_txppWframe_grdList_body_tbody')
            
            # 모든 tr 요소 찾기
            rows = tbody.find_elements("css selector", 'tr')
            posts = []
            
            for row in rows:
                try:
                    # code값 (data-colindex="0"인 td의 nobr 텍스트)
                    code_td = row.find_element("css selector", 'td[data-colindex="0"]')
                    code = code_td.find_element("css selector", 'nobr').text.strip()
                    
                    # title값 (data-colindex="2"인 td의 a 텍스트)
                    title_td = row.find_element("css selector", 'td[data-colindex="2"]')
                    title = title_td.find_element("css selector", 'a').text.strip()
                    
                    # 변경일 (data-colindex="4"인 td의 nobr 텍스트)
                    date_td = row.find_element("css selector", 'td[data-colindex="4"]')
                    date_raw = date_td.find_element("css selector", 'nobr').text.strip()
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
            
            return posts
        except NoSuchElementException:
            tax_log("w", "hometax", f"필수 요소를 찾을 수 없습니다. 페이지 구조가 변경되었을 수 있습니다.")
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
                send_notification("HomeTax", post_title, self.url)
                self.updated = True

        if latest_posts:
            self.known_posts = latest_posts
        return self.updated, self.known_posts

# --- 기획재정부 스크래핑 ---
class MoefScraper:
    def __init__(self, shared_data):
        self.url = "https://www.moef.go.kr/lw/lap/TbPrvntcList.do?bbsId=MOSFBBS_000000000055&menuNo=7050300&searchCondition3=1&searchKeyword3=%EC%86%8C%EB%93%9D%EC%84%B8"
        self.known_posts = shared_data.get("moef", {})
        self.session = make_http_session()

    def fetch_latest_posts(self):
        try:
            tax_log("i", "moef", f"새 글 찾는 중...")
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
                send_notification("기획재정부", post_title, self.url)
                self.updated = True

        if latest_posts:
            self.known_posts = latest_posts
        return self.updated, self.known_posts

# --- 법제처 스크래핑 ---
class MolegScraper:
    def __init__(self, shared_data):
        self.url = "https://www.moleg.go.kr/lawinfo/makingList.mo?mid=a10104010000&pageCnt=10&lsClsCd=&cptOfiOrgCd=&keyField=lmNm&keyWord=%EC%86%8C%EB%93%9D%EC%84%B8%EB%B2%95&stYdFmt=&edYdFmt="
        self.known_posts = shared_data.get("moleg", {})
        self.session = make_http_session()

    def fetch_latest_posts(self):
        try:
            tax_log("i", "moleg", f"새 글 찾는 중...")
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
                send_notification("법제처", post_title, self.url)
                self.updated = True

        if latest_posts:
            self.known_posts = latest_posts
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
        tax_log("i", "gwanbo", f"새 글 찾는 중...")
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
                send_notification("대한민국 전자관보", post_title, self.keyword_search_page_url())
                self.updated = True

        if latest_posts:
            self.known_posts = latest_posts
        return self.updated, self.known_posts

# --- 매니저 클래스 ---
class ScraperManager:
    def __init__(self):
        self.shared_data = load_data()
        self.hometax_scraper = HomeTaxScraper(self.shared_data)
        self.moef_scraper = MoefScraper(self.shared_data)
        self.moleg_scraper = MolegScraper(self.shared_data)
        self.gwanbo_scraper = GwanboScraper(self.shared_data)

    def check_all(self):
        updated = False

        ht_updated, ht_data = self.hometax_scraper.check_update()
        mf_updated, mf_data = self.moef_scraper.check_update()
        ml_updated, ml_data = self.moleg_scraper.check_update()
        gb_updated, gb_data = self.gwanbo_scraper.check_update()

        if ht_updated or mf_updated or ml_updated or gb_updated:
            self.shared_data["hometax"] = ht_data
            self.shared_data["moef"] = mf_data
            self.shared_data["moleg"] = ml_data
            self.shared_data["gwanbo"] = gb_data
            save_data(self.shared_data)
            tax_log("i", "", f"통합 데이터 파일 업데이트 완료")

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


def show_board_overview(icon, item):
    """HTML 파일을 생성해 4개 모니터링 목록을 브라우저로 연다."""
    try:
        data = load_data()
        sections = [
            ("홈택스", "hometax", manager.hometax_scraper.url),
            ("기획재정부", "moef", manager.moef_scraper.url),
            ("법제처", "moleg", manager.moleg_scraper.url),
            ("전자관보", "gwanbo", manager.gwanbo_scraper.keyword_search_page_url()),
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
            else:
                list_items = []
                if posts:
                    for _idx, (_, post_title) in enumerate(posts.items(), start=1):
                        list_items.append(f"<li>{escape(post_title)}</li>")
                else:
                    list_items.append("<li class='empty'>저장된 목록이 없습니다.</li>")
                list_block = f"<ul>{''.join(list_items)}</ul>"

            logo_url = BOARD_OVERVIEW_CARD_LOGOS.get(key)
            if logo_url:
                heading_html = (
                    f'<h2 class="card-site-heading">'
                    f'<img class="card-site-logo" src="{escape(logo_url)}" alt="{escape(title)}" '
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

        base_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
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

# --- 모니터링 쓰레드 ---
def run_monitor():
    global tray_icon
    # 최초 실행시 즉시 한 번 체크
    tax_log("i", "", f"프로그램 시작, 최초 모니터링 실행")
    manager.check_all()
    while True:
        next_check_time = datetime.datetime.now() + datetime.timedelta(seconds=SLEEP_TIME)
        tax_log("i", "", f"다음 모니터링 예정 시간: {next_check_time.strftime("%Y-%m-%d %H:%M:%S")}")
        # tray_icon.title 업데이트
        if tray_icon is not None:
            tray_icon.title = f"다음 모니터링 예정 시간: {next_check_time.strftime('%H시 %M분')}"
        time.sleep(SLEEP_TIME)
        
        tax_log("i", "", f"시간 경과, 모니터링 시작")
        manager.check_all()

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
    asciiart()
    add_to_startup()
    get_config()
    if not os.path.exists(LOG):
        open(LOG, "w", encoding="utf-8").close()
    if not os.path.exists(BOARD_DATA):
        open(BOARD_DATA, "w", encoding="utf-8").close()

    send_notification("TaxBoardWatcher","start") # 최초 실행 알림창
    manager = ScraperManager()

    monitor_thread = threading.Thread(target=run_monitor, daemon=True)
    monitor_thread.start()

    tray_icon = pystray.Icon("Web Monitor")
    tray_icon.icon = Image.open(ICON).resize((32, 32))
    tray_icon.title = "TaxBoardWatcher"

    tray_icon.menu = pystray.Menu(
        pystray.MenuItem('✔ 게시물 목록', show_board_overview),
        pystray.MenuItem('▶ 모니터링 실행', manual_crawl),
        pystray.MenuItem('✱ 로그 열기', open_log_file),
        pystray.MenuItem('⏻ 종료', on_exit),
   
    )
    tray_icon.run()