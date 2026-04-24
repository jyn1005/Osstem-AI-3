# CLAUDE.md — Osstem AI 업무자동화 프로젝트

## 프로젝트 개요

오스템임플란트 회계팀 내부 업무 자동화 도구 모음.
SAP에서 추출한 Excel 데이터를 처리하고, 회계 결산 검증을 자동화하는 Python 스크립트 집합.

**주요 사용자**: 회계팀 재고매입결산 담당자 (비개발자)
**목표**: 반복적인 SAP 데이터 정제 및 GR/IR 매칭 검증 업무 자동화

---

## 기술 스택

- **Python** 3.10+
- **pandas** >= 2.0.0
- **openpyxl** >= 3.1.0
- 향후: `anthropic` SDK (Claude API 연동 예정, `.env.example` 참고)

---

## 프로젝트 구조

```
Osstem-AI-3/
├── src/
│   ├── sap_excel_cleaner.py          # SAP Excel 정제 (완성)
│   ├── utils/                         # 공용 유틸리티 (예정)
│   └── modules/
│       ├── analytics/
│       │   └── gr_ir_matcher.py       # GR/IR 매칭 엔진 (개발 중)
│       ├── chatbot/                   # 챗봇 모듈 (예정)
│       └── document/                  # 문서 처리 모듈 (예정)
├── tests/
├── docs/
│   └── 사용법.md
├── requirements.txt
├── .env.example
└── CLAUDE.md
```

---

## 핵심 함수 레퍼런스

### `src/sap_excel_cleaner.py`
| 함수 | 역할 |
|------|------|
| `clean_sap_excel(input_path, output_path)` | SAP Excel 정제 메인 함수. 신규 스크립트에서 재사용 가능 |
| `find_data_start(df_raw)` | SAP 메타 헤더를 건너뛰고 실제 데이터 시작 행 탐지 |
| `_apply_styles(ws)` | openpyxl 워크시트에 표준 서식 적용 (헤더 색상, 교차 행, 열 너비, 틀 고정) |

**서식 상수** (다른 모듈에서도 동일하게 사용):
```python
HEADER_COLOR  = "1F4E79"  # 헤더 배경 (진한 파랑)
HEADER_FONT   = "FFFFFF"  # 헤더 글자 (흰색)
ALT_ROW_COLOR = "D6E4F0"  # 짝수 행 배경 (연한 파랑)
```

---

## 개발 규칙

### 코드 스타일
- 함수·변수명: `snake_case` (Python 표준)
- 주석·출력 메시지: **한국어** (사용자가 비개발자이므로 진행 상황을 한국어로 출력)
- CLI 스크립트는 `if __name__ == "__main__":` 블록에 argparse 또는 sys.argv 처리 포함

### Excel 출력 원칙
- 원본 파일은 절대 덮어쓰지 않음 — 항상 새 파일로 저장
- 출력 파일명 패턴: `원본파일명_기능명_YYYYMMDD_HHMMSS.xlsx`
- 서식은 `_apply_styles()` 재사용, 시트가 여러 개일 경우 각 시트에 개별 적용

### 의존성
- `requirements.txt`에 버전 고정 필수 (`>=` 최소 버전 명시)
- 새 패키지 추가 시 `requirements.txt` 동시 업데이트

---

## 현재 개발 중인 기능

### GR/IR 매칭 엔진 (`src/modules/analytics/gr_ir_matcher.py`)

**목적**: SAP GR(입고) · IR(송장) Excel을 PO 번호 기준으로 3-way 매칭하여
잔액을 미착품/미확정채무/예외로 자동 분류하는 결산 검증 도구

**분류 로직**:
```
잔액 = GR금액합계 - IR금액합계  (PO번호별 groupby)

abs(잔액) <= tolerance(기본 1원)  → 완전매칭
GR합계 > 0, IR합계 == 0           → 미착품     (자산 계정)
IR합계 > 0, GR합계 == 0           → 미확정채무  (부채 계정)
그 외                              → 예외 (수동 검토 필요)
```

**출력 Excel 시트 구성**: 완전매칭 / 미착품 / 미확정채무 / 예외건 / 전체

**CLI**:
```bash
python src/modules/analytics/gr_ir_matcher.py \
  --gr <GR_Excel> --ir <IR_Excel> [--output <저장경로>] [--tolerance 1]
```

**미확정 사항** (SAP 실제 파일 수령 후 확정 필요):
- GR/IR Excel 실제 컬럼명
- PO 라인 단위 vs PO 번호 단위 집계 방식
- 외화 PO 환율 처리 기준 시점

---

## 개발 일지

### 2026-04-24

**완료 작업**:
- `mr11_gui.py` — GUI 전면 색상 재설계
  - 색상 팔레트 상수 도입: `BG`(연한 청회색) · `HEADER_BG`(딥 네이비) · `ACCENT`(블루) · `LOG_BG`(다크 네이비) 등
  - 헤더·조회조건·저장위치·로그 영역을 카드(흰색+테두리) 레이아웃으로 구성
  - 로그 텍스트 색상 흰색(`#FFFFFF`), 실행 중 버튼 `disabledforeground=white` 적용
  - 창 크기 조절 가능(`resizable(True, True)`, 최소 560×580)
- `mr11_gui.py` — 저장 위치 선택 기능 추가
  - `filedialog.askdirectory()` 기반 `폴더 선택` 버튼 추가
  - 기본값: exe 실행 폴더, 변경 시 선택한 경로로 저장
  - `run_download()` 함수 시그니처에 `save_dir` 인자 추가
- `mr11_gui.py` — 연말 전표 날짜 매칭 버그 수정
  - 전기일(col 17)·입력일(col 28) 중 회계연도 일치 날짜 우선 선택 → 1월 등 연말 전표 누락 방지

### 2026-04-22

**완료 작업**:
- `mr11_gui.py` — F4 matchcode 전표 조회 0개 버그 근본 수정
  - `read_page_direct()` row 2 gap 문제 수정: 예외 발생 시 `break` → `consecutive_fails >= 5` 방식으로 변경, 헤더(row 1)·row 2 누락 구간을 건너뛰고 row 3~부터 데이터 정상 읽기
  - `matchcode_popup_debug.txt` 분석으로 팝업 실제 컬럼 구조 확인 (col 1=전표번호, col 12=연도, col 17=전기일, col 28=입력일)
  - MAXRECORDS `""` → `"9999"` 로 변경, 날짜 범위 필터 제거 (오히려 0개 야기)
  - `usr.Children` 미사용 → SAP 스크롤 포커스 유지 → `sendVKey(82)` 페이지 넘김 정상 동작
- `mr11_gui.py` — `pythoncom.CoInitialize/CoUninitialize()` 추가: 연속 실행 시 `CoInitialize가 호출되지 않았습니다` 오류 수정
- `mr11_gui.py` — 엑셀 저장 경로 수정: `__file__`(임시폴더) → `sys.executable` 기준 exe 실행 위치로 저장
- `mr11_gui.py` — 저장 완료 시 알림 팝업 추가 (`messagebox.showinfo`)
- `mr11_gui.py` — 파일명 형식 변경: `MR11_{연도}_{월}월반제리스트_{타임스탬프}.xlsx` → `MR11반제리스트_{연도}_{월}월.xlsx`
- `mr11_gui.py` — 음수 표시 형식 변경: `(#,##0)` 괄호형 → `-#,##0` 마이너스 부호형
- `mr11_gui.py` — `parse_a_row()` I/J/K(Plnt/내역/OUn) 파싱 로직 전면 개선
  - Case A: `'1000 내역설명'` → 같은 part에 Plnt+내역
  - Case B: `'1000'` / `'내역설명'` → 별도 part로 분리
  - Case C: `'1000 내역설명 EA'` → Plnt+내역+OUn이 하나로 묶인 경우 (3SHAPE 등 일부 업체)

### 2026-04-21

**완료 작업**:
- `mr11_gui.py` — MR11 반제리스트 다운로더 Windows GUI 앱 신규 구현 (tkinter)
  - 회계연도·전기월 선택, SAP 로그인 안내 팝업, 실행 로그 실시간 출력
  - 오리 아이콘 적용 (duck.ico, 흰 배경 제거), 작업 표시줄 아이콘 적용
  - PyInstaller `--onefile --windowed` exe 빌드 (34MB)
- `download_march_all.py` — argparse로 `--month`/`--year` CLI 인자 추가, 월별 범용 실행 지원
- Excel 서식 개선: G·H열 회계 표시형식(`_(* #,##0_)`), 오른쪽 정렬, H1 `=SUM` 수식
- `MR11_2026_02월반제리스트` 59행 I/J/K 파싱 오류 수동 수정 (Plnt/내역/OUn)
- 연말 전표 필터링 버그 수정: 전기일(2025.12) vs 입력일(2026.01) 혼동 → 회계연도 일치 날짜 우선 선택
- F4 matchcode 팝업 33개 제한 원인 분석: `MAXRECORDS` 필드·연도 필터·`sendVKey(0)` 방식 적용
- `read_popup_labels()` → `read_page_direct()` 리팩터링: `usr.Children` 열람 제거, `findById` 직접 경로로 SAP 포커스 유지

**미확정 사항**:
- `read_page_direct()` 방식으로 80개 전표 전체 읽기 확인 필요 (2025.12 전표 포함 여부)

### 2026-04-17

**완료 작업**:
- `src/modules/analytics/mr11_processor.py` — SAP XLS(UTF-16 탭구분) 직접 읽기 지원, 컬럼명 수정(선택번호→전표번호, 항목→품목, 계정 키이름→계정키이름), 색상 회색 계열 변경
  - `_read_sap_xls()`: UTF-16 LE BOM + 탭구분 SAP 스프레드시트 직접 파싱
  - `_is_po_number()`: SAP 구매문서번호 식별 (8자리↑, 4로 시작)
- `src/modules/analytics/mr11show_sap_extractor.py` — SAP MR11SHOW 전체 자동화 v3
  - F4 matchcode → 전표목록 조회 → 전표별 스프레드시트 내보내기 → 반제리스트 Excel 누적 저장
  - `get_doc_list_via_matchcode()`: F4 팝업에서 전표번호·전기일자 자동 추출
  - `_sap_export_to_xls()`: SAP 메뉴 내보내기 → 파일 저장 경로 자동 감지
  - `extract_single_doc()`: 전표 1개 상세 데이터 추출 (내보내기 → `parse_rawdata()` 연동)
- `src/modules/analytics/find_alv_mr11show.py` — SAP GUI 컨트롤 트리 덤프 디버그 도구
- `download_march_all.py` — MR11SHOW 3월 전표 전체 일괄 다운로드 스크립트
  - SAP F4 matchcode로 전표 목록 읽기 → 3월 필터링 → 전표별 클립보드 내보내기 → 파싱 → 1개 Excel 저장
  - SAP 클립보드(CF_UNICODETEXT) 고정폭 텍스트 파싱: 2칸 이상 공백으로 필드 분리, `O.xxx.xxxx` 형식 재료코드 지원
  - 2026년 3월 전표 9개 / 99건 / 차이금액 합계 5,088,278 처리 완료

**미확정 사항**:
- 파싱 실패 엣지케이스: 내역이 매우 길거나 재료코드 사이 공백이 부족한 행 (86행 수동 수정)

---

### 2026-03-27

**완료 작업**:
- `src/modules/analytics/gr_ir_matcher.py` — GR/IR 매칭 엔진 초기 구현 완료
  - `load_gr_data()` / `load_ir_data()`: SAP Excel 로드 및 표준 컬럼 정규화
  - `match_gr_ir()`: PO 번호 기준 outer join 매칭 + 잔액 계산 + 5분류 레이블링
  - `add_aging()`: PO일자 기준 경과일수 및 Aging 구간(30일이하/31~60일/61~90일/90일초과) 산출
  - `export_report()`: 완전매칭·미착품·미확정채무·예외건·전체 5시트 Excel 리포트 생성
  - CLI: `--gr`, `--ir`, `--output`, `--tolerance` 인자 지원
- `tests/test_gr_ir_matcher.py` — pytest 단위 테스트 작성 (25개 케이스)
  - `TestMatchGrIr`: 완전매칭·미착품·미확정채무·예외 분류, 복수 PO, 합산 집계 등
  - `TestAddAging`: 4개 구간 + NaT 처리
  - `TestClassify`: `_classify()` 엣지케이스
- `src/modules/analytics/mr11_processor.py` — SAP MR11SHOW CSV 정제 및 반제리스트 Excel 누적 추가
- `.claude/skills/push/SKILL.md` — CLAUDE.md 정리 + 자동 커밋·푸시 슬래시 커맨드 생성
- `.claude/settings.json` — Claude Code 프로젝트 설정 파일 생성

**미확정 사항** (SAP 실제 파일 수령 후 확정 필요):
- GR/IR Excel 실제 컬럼명 (`GR_COL_MAP`, `IR_COL_MAP` 수정 필요)
- PO 라인 단위 vs PO 번호 단위 집계 방식 최종 결정
- 외화 PO 환율 처리 기준 시점

---

## 도메인 용어

| 용어 | 설명 |
|------|------|
| GR (Goods Receipt) | 입고처리. SAP 트랜잭션: MIGO |
| IR (Invoice Receipt) | 송장처리. SAP 트랜잭션: MIRO |
| PO (Purchase Order) | 구매발주서 |
| 미착품 | GR 완료, IR 미도착 → 자산 계정 |
| 미확정채무 | IR 완료, GR 미완 → 부채 계정 |
| 3-way Matching | PO · GR · IR 금액/수량 일치 검증 |
| GR/IR 계정 | 입고·송장 간 임시 중간 계정 (SAP: MB5S로 조회) |
| Aging | 미결 건 경과일수 (오늘 - PO 일자) |
