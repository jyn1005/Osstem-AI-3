# Osstem AI 업무자동화 프로젝트

오스템임플란트 내부 AI 기반 업무 자동화 프로젝트입니다.

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

AI 기술을 활용하여 오스템임플란트의 내부 업무 프로세스를 자동화하고, 업무 효율성을 극대화하는 것을 목표로 합니다.

---

## 기능 목록

### SAP Excel 자동 정리 (`src/sap_excel_cleaner.py`)

SAP에서 다운로드한 회계/비용 데이터 Excel을 자동으로 정리합니다.

| 기능 | 설명 |
|------|------|
| 헤더 자동 제거 | SAP 메타데이터 상단 헤더를 자동으로 탐지 후 제거 |
| 빈 행/열 삭제 | 불필요한 빈 행과 열 자동 정리 |
| 합계/소계 행 제거 | SAP 자동 생성 합계·소계 행 제거 |
| 시각적 서식 | 헤더 색상(진한 파랑), 교차 행 색상(연한 파랑) 적용 |
| 열 너비 자동 조정 | 내용에 맞게 열 너비 자동 설정 (최대 40자) |
| 틀 고정 | 헤더 행(1행) 틀 고정 |

---

## 설치 방법

**Python 3.10 이상** 필요

```bash
pip install -r requirements.txt
```

---

## 사용법

### SAP Excel 정리

```bash
# 기본 실행
python src/sap_excel_cleaner.py <SAP_Excel_파일경로>

# 저장 경로 직접 지정
python src/sap_excel_cleaner.py <SAP_Excel_파일경로> <저장경로>

# 예시
python src/sap_excel_cleaner.py C:/Downloads/SAP_report.xlsx
```

**실행 결과**
- 원본 파일과 같은 폴더에 `원본파일명_정리완료_날짜시간.xlsx` 파일이 생성됩니다.

---

## 프로젝트 구조

```
Osstem-AI-3/
├── src/
│   ├── sap_excel_cleaner.py   # SAP Excel 자동 정리 스크립트
│   ├── utils/                 # 공용 유틸리티 (예정)
│   └── modules/
│       ├── analytics/         # 분석 모듈 (예정)
│       ├── chatbot/           # 챗봇 모듈 (예정)
│       └── document/          # 문서 처리 모듈 (예정)
├── tests/                     # 테스트 코드
├── docs/
│   └── 사용법.md
├── requirements.txt
└── README.md
```

---

## 기술 스택

- **Python** 3.10+
- **pandas** >= 2.0.0
- **openpyxl** >= 3.1.0

---

*본 프로젝트는 오스템임플란트 내부 AI 업무자동화 프로젝트입니다.*
