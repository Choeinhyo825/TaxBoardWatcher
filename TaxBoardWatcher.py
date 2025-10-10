from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from win32com.client import Dispatch
from PIL import Image, ImageDraw
from selenium import webdriver
from bs4 import BeautifulSoup
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

VERSION = "v.3.0.0"

BOARD_DATA = "data/board_data.json"
LOG = "data/log.txt"
ICON = 'data/tax.png'
SEARCHING_ICON = 'data/searching.png'
CONFIG = "data/config.json"
SLEEP_TIME = 3*60*60 # 기본 3시간

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
            button_frame.pack()
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
            label2 = tk.Label(root, text=f"현재 설정된 크롤링 간격은 {int(SLEEP_TIME/60/60)}시간입니다.", bg="#f0f4f7", font=("Arial", 10), wraplength=300)
            label2.pack(pady=(5, 15))
            # 버튼 영역
            button_frame = tk.Frame(root, bg="#f0f4f7")
            button_frame.pack()
            close_button = tk.Button(button_frame, text="시작", command=on_close, width=12, bg="#0078D7", fg="white", font=("Arial", 10, "bold"))
            close_button.pack(side="left", padx=10)

        root.update_idletasks() # 내부 위젯 크기에 맞게 자동 조정
        root.minsize(350, 200)  # 최소 크기만 지정
        root.deiconify() # withdraw() 이후 알림창 표시하기 위함

        # Fade-in 애니메이션
        for i in range(0, 11):  # 0.0 ~ 1.0
            root.attributes("-alpha", i / 10)
            root.update()
            time.sleep(0.03)  # 속도 조정 가능 (0.02~0.05 정도가 자연스러움)

        root.mainloop()
    except Exception as e:
        tax_log("e", "", f" 알림 전송 오류: {e}")

# --- 저장 파일 관리 함수 ---
def save_data(data):
    with open(BOARD_DATA, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
def load_data():
    try:
        if os.path.exists(BOARD_DATA):
            with open(BOARD_DATA, 'r', encoding='utf-8') as f:
                return json.load(f)
        return {"hometax": {}, "moef": {}}
    except Exception as e:
        tax_log("e", "", f"파일 로드 오류: {e}")
        return {"hometax": {}, "moef": {}}

# --- 홈택스 크롤러 ---
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
                    
                    # update date값 (data-colindex="4"인 td의 nobr 텍스트)
                    date_td = row.find_element("css selector", 'td[data-colindex="4"]')
                    date = date_td.find_element("css selector", 'nobr').text.strip().replace("-", "")
                    
                    posts.append({"code": f"{date}_{code}", "title": title})
                except Exception as e:
                    tax_log("e", "hometax", f"행 파싱 오류: {e}")
            
            return posts
        except NoSuchElementException:
            tax_log("w", "hometax", f"필수 요소를 찾을 수 없습니다. 페이지 구조가 변경되었을 수 있습니다.")
            return []
        except Exception as e:
            tax_log("e", "hometax", f"크롤링 오류: {e}")
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
            latest_posts[post_code] = post_title
            
            if post_code not in self.known_posts:
                tax_log("i", "hometax", f"새 글 발견: {post_title} (코드: {post_code})")
                send_notification("HomeTax", post_title, self.url)
                self.updated = True

        if latest_posts:
            self.known_posts = latest_posts
        return self.updated, self.known_posts

# --- 기획재정부 크롤러 ---
class MoefScraper:
    def __init__(self, shared_data):
        self.url = "https://www.moef.go.kr/lw/lap/TbPrvntcList.do?bbsId=MOSFBBS_000000000055&menuNo=7050300&searchCondition3=1&searchKeyword3=%EC%86%8C%EB%93%9D%EC%84%B8"
        self.known_posts = shared_data.get("moef", {})

    def fetch_latest_posts(self):
        try:
            tax_log("i", "moef", f"새 글 찾는 중...")
            res = requests.get(self.url, timeout=10)
            res.raise_for_status()
            soup = BeautifulSoup(res.text, "html.parser")

            posts = soup.select("ul.boardType3 li")
            new_posts = []
            
            for post in posts:
                link = post.find("a")
                if link:
                    # fn_egov_select의 파라미터로 추출된 고유 ID
                    post_id = link['href'].split('"')[1]
                    title = link.get_text(strip=True)
                    new_posts.append({"code": post_id, "title": title})

            return new_posts
        except Exception as e:
            tax_log("e", "moef", f"크롤링 오류: {e}")
            return []

    def check_update(self):
        self.updated = False
        posts = self.fetch_latest_posts()

        latest_posts = {}
        for post in posts:
            post_code = post["code"]
            post_title = post["title"]
            latest_posts[post_code] = post_title
            if post_code not in self.known_posts:
                tax_log("i", "moef", f"새 글 발견: {post_title} (코드: {post_code})")
                send_notification("기획재정부", post_title, self.url)
                self.updated = True

        if latest_posts:
            self.known_posts = latest_posts
        return self.updated, self.known_posts

# --- 법제처 크롤러 ---
class MolegScraper:
    def __init__(self, shared_data):
        self.url = "https://www.moleg.go.kr/lawinfo/makingList.mo?mid=a10104010000&pageCnt=10&lsClsCd=&cptOfiOrgCd=&keyField=lmNm&keyWord=%EC%86%8C%EB%93%9D%EC%84%B8%EB%B2%95&stYdFmt=&edYdFmt="
        self.known_posts = shared_data.get("moleg", {})

    def fetch_latest_posts(self):
        try:
            tax_log("i", "moleg", f"새 글 찾는 중...")
            res = requests.get(self.url, timeout=10)
            res.raise_for_status()
            soup = BeautifulSoup(res.text, "html.parser")

            posts = soup.select("#listForm .wrap.title a")
            new_posts = []

            for post in posts:
                href = post.get("href", "")
                title = post.get("title", "").strip()
                if not href or not title:
                    continue
                law_seq = self._extract_law_seq(href)
                # if law_seq and law_seq not in self.known_posts:
                new_posts.append({"code": law_seq, "title": title})

            return new_posts
        except Exception as e:
            tax_log("e", "moleg", f"크롤링 오류: {e}")
            return []

    def check_update(self):
        self.updated = False
        posts = self.fetch_latest_posts()
        
        latest_posts = {}
        for post in posts:
            post_code = post["code"]
            post_title = post["title"]
            latest_posts[post_code] = post_title
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

# --- 관보 크롤러 ---
class GwanboScraper:
    def __init__(self, shared_data):
        self.url = "https://gwanbo.go.kr/SearchRestApi.jsp"
        self.headers = {
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "User-Agent": "Mozilla/5.0"
        }
        self.known_posts = shared_data.get("gwanbo", {})
        self.updated = False

    def fetch_latest_posts(self):
        tax_log("i", "gwanbo", f"새 글 찾는 중...")
        try:
            response = requests.post(self.url, headers=self.headers, data={
                "mode": "keyword",
                "index": "gwanbo",
                "query": "(unstored_field_subject:(소득세법)) AND keyword_category_order:(@@ORDER_NUM)",
                "pQuery_tmp": "소득세법",
                "pageNo": "1",
                "listSize": "5",
                "sort": ""
            })
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
                    # if key and key not in self.known_posts:
                    new_posts.append({"code": key, "title": f"{category_name}-{subject}(검색어:소득세법)"})

            return new_posts
        
        except Exception as e:
            tax_log("e", "gwanbo", f"크롤링 오류: {e}")
            return []

    def check_update(self):
        self.updated = False
        posts = self.fetch_latest_posts()
        
        latest_posts = {}
        for post in posts:
            post_code = post["code"]
            post_title = post["title"]
            latest_posts[post_code] = post_title
            if post_code not in self.known_posts:
                tax_log("i", "gwanbo", f"새 글 발견: {post_title} (코드: {post_code})")
                send_notification("대한민국 전자관보", post_title, "https://gwanbo.go.kr/user/search/searchKeyword.do")
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

def on_exit(icon, item):
    tax_log("i", "", f"사용자 요청으로 종료")
    icon.stop()
    os._exit(0)

# --- 모니터링 쓰레드 ---
def run_monitor():
    global tray_icon
    # 최초 실행시 즉시 한 번 체크
    tax_log("i", "", f"프로그램 시작, 최초 크롤링 실행")
    manager.check_all()
    while True:
        next_check_time = datetime.datetime.now() + datetime.timedelta(seconds=SLEEP_TIME)
        tax_log("i", "", f"다음 크롤링 예정 시간: {next_check_time.strftime("%Y-%m-%d %H:%M:%S")}")
        # tray_icon.title 업데이트
        if tray_icon is not None:
            tray_icon.title = f"다음 크롤링 예정 시간: {next_check_time.strftime('%H시 %M분')}"
        time.sleep(SLEEP_TIME)
        
        tax_log("i", "", f"시간 경과, 크롤링 시작")
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
        f.write(f"=================================================================================================== {VERSION}\n")
        f.write(f" _____                 ______                          _     _    _         _          _                  \n")
        f.write(f"|_   _|                | ___ \                        | |   | |  | |       | |        | |                 \n")
        f.write(f"  | |    __ _ __  __   | |_/ /  ___    __ _  _ __   __| |   | |  | |  __ _ | |_   ___ | |__    ___  _ __  \n")
        f.write(f"  | |   / _` |\ \/ /   | ___ \ / _ \  / _` || '__| / _` |   | |/\| | / _` || __| / __|| '_ \  / _ \| '__| \n")
        f.write(f"  | |  | (_| | >  <    | |_/ /| (_) || (_| || |   | (_| |   \  /\  /| (_| || |_ | (__ | | | ||  __/| |    \n")
        f.write(f"  \_/   \__,_|/_/\_\   \____/  \___/  \__,_||_|    \__,_|    \/  \/  \__,_| \__| \___||_| |_| \___||_|    \n")
        f.write(f"==========================================================================================================\n")

# 크롤링 수동 실행
def manual_crawl(icon, item):
    tax_log("i", "", f"수동 크롤링 시작")
    set_tray_icon(SEARCHING_ICON) # 트레이 아이콘 "검색"
    manager.check_all()
    set_tray_icon(ICON) # 트레이 아이콘 기본값
    tax_log("i", "", f"수동 크롤링 완료")

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
        pystray.MenuItem('크롤링 실행', manual_crawl),
        pystray.MenuItem('종료', on_exit)
    )
    tray_icon.run()