# TaxBoardWatcher
## _홈텍스 등 게시판 모니터링_
### 대상
- [국세청 홈텍스 자료실][hometax]
- [기획재정부 입법 행정예고(검색어 "소득세")][moef]
- [법제처(검색어 "소득세법")][moleg]
- [전자관보(검색어 "소득세법")][gwanbo]
- [행정안전부][mois]

### 기본 명령어
```sh
# py 실행
python TaxBoardWatcher.py

# exe파일 생성 (winotify는 스크립트 상단 import로 PyInstaller에 포함됨. 토스트가 exe에서 안 뜨면 --hidden-import winotify 추가)
# dist/data의 아이콘·로고·기본 설정을 exe에 포함 (실행 시 data/에 없는 파일만 풀어 놓음)
pyinstaller --noconsole --onefile --name TaxBoardWatcher --add-data "dist/data/*.png;default_data" --add-data "dist/data/config.json;default_data" TaxBoardWatcher.py

# 자동시작 생성 경로
C:\Users\사용자명\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup
```

### 배포 (자동 업데이트)
`v*` 태그를 push하면 GitHub Actions([.github/workflows/release.yml](.github/workflows/release.yml))가 exe를 빌드해 Release에 올리고,
사용자 PC의 프로그램이 최신 Release를 확인해 트레이 메뉴에서 업데이트합니다.
```sh
# 1. TaxBoardWatcher.py 의 VERSION 과 dist/readme.txt 변경 이력 수정 후 커밋
# 2. 태그 push (태그와 VERSION 숫자가 다르면 빌드가 실패함)
git tag v5.0.2
git push origin master v5.0.2
```
- Release 자산 이름은 `TaxBoardWatcher.exe` 여야 합니다(앱이 이 이름으로 찾음).
- 저장소가 public 이어야 사용자 PC가 토큰 없이 Release 를 받을 수 있습니다.

### update 내역
dist/readme.txt 를 참고 하세요.

   [hometax]: <https://hometax.go.kr/websquare/websquare.html?w2xPath=/ui/pp/index_pp.xml&tmIdx=16&tm2lIdx=1602000000&tm3lIdx=>
   [moef]: <https://www.moef.go.kr/lw/lap/TbPrvntcList.do?bbsId=MOSFBBS_000000000055&menuNo=7050300&searchCondition3=1&searchKeyword3=%EC%86%8C%EB%93%9D%EC%84%B8>
   [moleg]: <https://www.moleg.go.kr/lawinfo/makingList.mo?mid=a10104010000&pageCnt=10&lsClsCd=&cptOfiOrgCd=&keyField=lmNm&keyWord=%EC%86%8C%EB%93%9D%EC%84%B8%EB%B2%95&stYdFmt=&edYdFmt=>
   [gwanbo]: <https://gwanbo.go.kr/user/search/searchKeyword.do?pKeyword=%EC%86%8C%EB%93%9D%EC%84%B8%EB%B2%95>
   [mois]: <https://www.mois.go.kr/frt/bbs/type001/commonSelectBoardList.do?bbsId=BBSMSTR_000000000052>
