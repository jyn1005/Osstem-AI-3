# Osstem AI 업무자동화 프로젝트

오스템임플란트 회계팀 내부 업무 자동화 도구 모음입니다.

---

## 목차

- [프로젝트 개요](#프로젝트-개요)
- [기능 목록](#기능-목록)
- [설치 방법](#설치-방법)
- [사용법](#사용법)
- [프로젝트 구조](#프로젝트-구조)
- [기술 스택](#기술-스택)

---

## 프로젝트 개요

SAP에서 추출한 Excel 데이터를 처리하고, 회계 결산 검증을 자동화하는 Python 스크립트 집합입니다.

**주요 사용자**: 회계팀 재고매입결산 담당자  
**목표**: 반복적인 SAP 데이터 정제 및 GR/IR 매칭 검증 업무 자동화

---

## 기능 목록

### 1. MR11 반제리스트 다운로더 (`mr11_gui.py`)

SAP MR11SHOW 트랜잭션에서 월별 전표를 자동으로 다운로드하여 Excel로 저장하는 **Windows GUI 앱**입니다.

| 기능 | 설명 |
|------|------|
| GUI 인터페이스 | 회계연도·전기월 선택, 실행 버튼, 진행 상황 로그 |
| SAP 자동화 | MR11SHOW F4 matchcode → 전표 목록 조회 → 클립보드 내보내기 |
| 데이터 파싱 | SAP 고정폭 텍스트 파싱 (2칸 이상 공백 기준 필드 분리) |
| Excel 저장 | 회계 서식(#,##0), 교대 행 색상, SUM 수식 자동 적용 |
| exe 배포 | Python 설치 없이 실행 가능한 단일 exe 파일 |

**실행 방법**
```bash
# Python으로 실행
python mr11_gui.py

# 또는 dist 폴더의 exe 더블클릭 (Python 불필요)
```

---

### 2. MR11 반제리스트 CLI (`download_march_all.py`)

터미널에서 월·연도를 지정하여 반제리스트를 다운로드하는 CLI 스크립트입니다.

```bash
python download_march_all.py --month 4           # 2026년 4월
python download_march_all.py --month 12 --year 2025
python download_march_all.py -m 3 -y 2026
```

---

### 3. SAP Excel 자동 정리 (`src/sap_excel_cleaner.py`)

SAP에서 다운로드한 회계/비용 데이터 Excel을 자동으로 정리합니다.

```bash
python src/sap_excel_cleaner.py <SAP_Excel_파일경로>
```

---

### 4. GR/IR 매칭 엔진 (`src/modules/analytics/gr_ir_matcher.py`)

SAP GR(입고)·IR(송장) Excel을 PO 번호 기준으로 3-way 매칭하여 미착품/미확정채무/예외를 자동 분류합니다.

```bash
python src/modules/analytics/gr_ir_matcher.py --gr GR.xlsx --ir IR.xlsx
```

---

## 설치 방법

**Python 3.10 이상** 필요

```bash
pip install -r requirements.txt
```

**exe 빌드** (Python 설치 없이 배포할 경우)
```bash
pip install pyinstaller pillow
pyinstaller --onefile --windowed --icon duck.ico --add-data "duck.ico;." ^
  --hidden-import win32com --hidden-import win32com.client ^
  --hidden-import win32clipboard --hidden-import pywintypes ^
  --collect-submodules win32com --name MR11_반제리스트_다운로더 mr11_gui.py
```

---

## 프로젝트 구조

```
Osstem-AI-3/
├── mr11_gui.py                          # MR11 반제리스트 GUI 앱
├── download_march_all.py                # MR11 반제리스트 CLI
├── duck.ico / duck.png                  # 앱 아이콘
├── src/
│   ├── sap_excel_cleaner.py             # SAP Excel 자동 정리
│   └── modules/
│       └── analytics/
│           ├── gr_ir_matcher.py         # GR/IR 매칭 엔진
│           ├── mr11_processor.py        # MR11 XLS 파서
│           └── mr11show_sap_extractor.py # SAP MR11SHOW 자동화 v3
├── tests/
├── requirements.txt
└── README.md
```

---

## 기술 스택

- **Python** 3.10+
- **openpyxl** >= 3.1.0
- **pywin32** (win32com, win32clipboard) — SAP GUI 자동화
- **tkinter** — Windows GUI
- **PyInstaller** — exe 빌드

---

*본 프로젝트는 오스템임플란트 내부 AI 업무자동화 프로젝트입니다.*
