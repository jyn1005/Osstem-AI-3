# Osstem AI 업무자동화 프로젝트

오스템임플란트의 AI 기반 업무 자동화 프로젝트입니다.

## 프로젝트 개요

AI 기술을 활용하여 오스템임플란트의 내부 업무 프로세스를 자동화하고, 업무 효율성을 극대화하는 것을 목표로 합니다.

## 기능

### SAP Excel 자동 정리 (`src/sap_excel_cleaner.py`)
SAP에서 다운로드한 회계/비용 데이터 Excel을 자동으로 정리합니다.

- SAP 메타데이터 헤더 자동 제거
- 빈 행/열 자동 삭제
- 헤더 색상, 교차 행 색상 등 시각적 서식 적용
- 열 너비 자동 조정 및 틀 고정

**사용법**
```bash
# 패키지 설치
pip install -r requirements.txt

# 실행
python src/sap_excel_cleaner.py <SAP_Excel_파일경로>

# 예시
python src/sap_excel_cleaner.py C:/Downloads/SAP_report.xlsx
```

**실행 결과**
- 원본 파일과 같은 폴더에 `_정리완료_날짜시간.xlsx` 파일이 생성됩니다.

## 기술 스택

- Python 3.10+
- pandas
- openpyxl

## 관련 링크

- [오스템임플란트 공식 홈페이지](https://www.osstem.com)

---

*본 프로젝트는 오스템임플란트 내부 AI 업무자동화 프로젝트입니다.*
