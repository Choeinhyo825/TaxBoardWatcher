# TaxBoardWatcher
## _홈텍스 등 게시판 모니터링_
### 대상
- [국세청 홈텍스 자료실][hometax]
- [기획재정부 입법 행정예고(검색어 "소득세")][moef]
- [법제처(검색어 "소득세법")][moleg]
- [전자관보(검색어 "소득세법")][gwanbo]

### 기본 명령어
```sh
# py 실행
python TaxBoardWatcher.py

# exe파일 생성 (winotify는 스크립트 상단 import로 PyInstaller에 포함됨. 토스트가 exe에서 안 뜨면 --hidden-import winotify 추가)
pyinstaller --noconsole --onefile TaxBoardWatcher.py

# 자동시작 생성 경로
C:\Users\inhyo\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup
```

### update 내역
dist/readme.txt 를 참고 하세요.

   [hometax]: <https://hometax.go.kr/websquare/websquare.html?w2xPath=/ui/pp/index_pp.xml&tmIdx=16&tm2lIdx=1602000000&tm3lIdx=>
   [moef]: <https://www.moef.go.kr/lw/lap/TbPrvntcList.do?bbsId=MOSFBBS_000000000055&menuNo=7050300&searchCondition3=1&searchKeyword3=%EC%86%8C%EB%93%9D%EC%84%B8>
   [moleg]: <https://www.moleg.go.kr/lawinfo/makingList.mo?mid=a10104010000&pageCnt=10&lsClsCd=&cptOfiOrgCd=&keyField=lmNm&keyWord=%EC%86%8C%EB%93%9D%EC%84%B8%EB%B2%95&stYdFmt=&edYdFmt=>
   [gwanbo]: <https://gwanbo.go.kr/user/search/searchKeyword.do?pKeyword=%EC%86%8C%EB%93%9D%EC%84%B8%EB%B2%95>
