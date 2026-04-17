# -*- coding: utf-8 -*-
"""
MR11SHOW SAP GUI 자동화 스크립트 v3
SAP MR11SHOW (계정유지보수전표 조회/취소) 트랜잭션에서
전기일자 기준으로 전표를 자동 추출하여 반제리스트 Excel에 저장합니다.

동작 방식:
  Phase 1) MR11SHOW 초기화면에서 F8 → 전표 목록 리스트 읽기
  Phase 2) 목표 월(예: 3) 해당 전표번호만 필터링
  Phase 3) 각 전표번호로 MR11SHOW 재조회 → Classic ABAP List 읽기
  Phase 4) 반제리스트 Excel 누적 저장

사전 조건:
    1. SAP GUI가 열려 있고 로그인된 상태
    2. Alt+F12 → Scripting → Enable Scripting 체크
    3. pip install pywin32

사용법:
    python src/modules/analytics/mr11show_sap_extractor.py `
        --year 2026 --month 3 --master 반제리스트_3월.xlsx

    # 디버그 모드 (화면 구조 확인)
    python src/modules/analytics/mr11show_sap_extractor.py `
        --year 2026 --month 3 --master 반제리스트_3월.xlsx --debug
"""

import sys
import os
import re
import time
import argparse

SAP_TCODE = "MR11SHOW"
MAX_PAGES = 200

# ─────────────────────────────────────────────
# 상세 화면(Detail) GuiLabel 컬럼 매핑
# alv_tree_mr11show.txt 분석으로 확인된 실제값:
#   A행 col 1  = PO번호(4500XXXXXX, 10자리)
#   B행 col 1  = 선택번호(짧은 숫자)
# ─────────────────────────────────────────────
A_COL_MAP = {
    1:   "구매 문서",
    12:  "품목",
    18:  "PO 일자",
    29:  "이름 1",
    84:  "Plnt",
    89:  "내역",
    130: "OUn",
}

B_COL_MAP = {
    25: "계정키이름",
    46: "차이 수량",
    65: "차이 금액",
}

MASTER_COLS = [
    "전표번호", "구매 문서", "품목", "PO 일자",
    "계정키이름", "이름 1", "차이 수량", "차이 금액",
    "Plnt", "내역", "OUn",
]
NUMERIC_COLS = {"차이 수량", "차이 금액"}


# ════════════════════════════════════════════════════════
#  공통 헬퍼
# ════════════════════════════════════════════════════════

def connect_sap():
    """실행 중인 SAP GUI 세션에 연결합니다."""
    try:
        import win32com.client
    except ImportError:
        raise RuntimeError(
            "pywin32 패키지가 없습니다.\n설치: pip install pywin32"
        )
    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
        session = sap_gui.GetScriptingEngine.Children(0).Children(0)
        print(f"  [SAP] {session.Info.SystemName} / {session.Info.User}")
        return session
    except Exception as e:
        raise RuntimeError(f"SAP GUI 연결 실패: {e}")


def navigate_to_tcode(session, tcode: str):
    session.findById("wnd[0]/tbar[0]/okcd").text = f"/n{tcode}"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1.5)


def press_f8(session):
    """MR11SHOW 실행: VBScript 확인 기준 Enter(VKey 0) 우선, F8 폴백."""
    # VBScript 녹화: session.findById("wnd[0]").sendVKey 0 (Enter)
    # F8(VKey 8)은 SAP GUI에서 예외 없이 성공해도 화면 전환이 안 될 수 있으므로
    # Enter를 먼저 시도하고, 실패 시 F8 → tbar 버튼 순으로 시도
    for method in [
        lambda: session.findById("wnd[0]").sendVKey(0),    # Enter (VBScript 확인)
        lambda: session.findById("wnd[0]").sendVKey(8),    # F8
        lambda: session.findById("wnd[0]/tbar[1]/btn[8]").press(),
    ]:
        try:
            method()
            time.sleep(3.0)
            return
        except Exception:
            continue


def _read_page_labels(session) -> dict:
    """현재 화면의 모든 GuiLabel을 {(col, row): text} 로 반환합니다.
    SAP 화면 로딩 지연 시 최대 3회 재시도합니다."""
    for attempt in range(3):
        try:
            usr = session.findById("wnd[0]/usr")
            labels = {}
            for i in range(usr.Children.Count):
                try:
                    lbl = usr.Children.ElementAt(i)
                    lid = lbl.Id
                    bracket = lid[lid.rfind("[") + 1: lid.rfind("]")]
                    c, r = map(int, bracket.split(","))
                    labels[(c, r)] = lbl.Text
                except Exception:
                    pass
            return labels
        except Exception:
            if attempt < 2:
                time.sleep(1.5)
            continue
    return {}


def _get_val(labels: dict, col: int, row: int) -> str:
    val  = labels.get((col, row), "").strip()
    sign = labels.get((col - 1, row), "").strip()
    if sign and sign not in ("", " "):
        return sign + val
    return val


def _clean_number(s: str):
    """SAP 금액 문자열 → float ('280,000-' 형식 포함)."""
    s = str(s).strip().replace(",", "")
    if not s:
        return 0.0
    negative = s.endswith("-")
    s = s.rstrip("-").strip()
    try:
        val = float(s)
        return -val if negative else val
    except ValueError:
        return 0.0


def _debug_labels(labels: dict, title: str):
    print(f"\n{'─'*60}\n【DEBUG】{title}\n{'─'*60}")
    for (c, r), text in sorted(labels.items(), key=lambda x: (x[0][1], x[0][0])):
        if text.strip():
            print(f"  lbl[{c:>4},{r:>3}] = '{text}'")
    print(f"{'─'*60}")


# ════════════════════════════════════════════════════════
#  Phase 1 — 전표 목록 읽기 (F4 matchcode 방식)
# ════════════════════════════════════════════════════════

def _read_list_from_window(wnd_path: str, session, debug: bool = False,
                           page_label: str = "") -> list[dict]:
    """
    지정 윈도우(wnd[0] 또는 wnd[1])의 목록에서 전표번호·전기일자를 읽습니다.
    ALV Grid 또는 Classic List 모두 처리합니다.
    """
    docs: list[dict] = []

    # ── ALV Grid 시도 ─────────────────────────────────────
    grid_suffixes = [
        "/usr/cntlGRID1/shellcont/shell",
        "/usr/cntl00100/shellcont/shell",
        "/usr/cntlALV_GRID/shellcont/shell",
        "/usr/cntl*/shellcont/shell",
    ]
    for suffix in grid_suffixes:
        try:
            grid = session.findById(wnd_path + suffix)
            row_count = grid.RowCount
            for row in range(row_count):
                doc_no = ""
                post_date = ""
                for col_id in ["BELNR", "MBLNR", "DOCNO", "BELRNR", "MKPF_BELNR"]:
                    try:
                        v = grid.GetCellValue(row, col_id).strip()
                        if v and len(v.replace(" ", "")) >= 6:
                            doc_no = v.strip()
                            break
                    except Exception:
                        continue
                for col_id in ["BUDAT", "BLDAT", "CPUDT", "DATUM", "MKPF_BUDAT"]:
                    try:
                        v = grid.GetCellValue(row, col_id).strip()
                        if v:
                            post_date = v.strip()
                            break
                    except Exception:
                        continue
                if doc_no:
                    docs.append({"doc_no": doc_no, "posting_date": post_date})
            if docs:
                print(f"  [ALV Grid/{wnd_path}] {len(docs)}개 전표")
                return docs
        except Exception:
            continue

    # ── Classic List (GuiLabel) 읽기 ─────────────────────
    try:
        wnd = session.findById(wnd_path)
        wnd.sendVKey(82)   # Ctrl+Home
        time.sleep(0.5)
    except Exception:
        pass

    seen: set = set()
    prev_anchor = None
    page = 0

    while page < 50:
        # wnd[0]이면 _read_page_labels, wnd[1]이면 직접 읽기
        if wnd_path == "wnd[0]":
            labels = _read_page_labels(session)
        else:
            labels = _read_window_labels(session, wnd_path)

        if not labels:
            break

        if debug:
            _debug_labels(labels, f"{page_label} 페이지 {page + 1} ({wnd_path})")

        all_vals = sorted(
            (r, c, labels[(c, r)].strip())
            for (c, r) in labels if r >= 1 and labels[(c, r)].strip()
        )
        anchor = str(all_vals[0]) if all_vals else ""
        if not anchor or anchor == prev_anchor:
            break
        prev_anchor = anchor

        rows_data: dict = {}
        for (c, r), text in labels.items():
            if r < 1 or not text.strip():
                continue
            if r not in rows_data:
                rows_data[r] = {}
            rows_data[r][c] = text.strip()

        for row, cols in sorted(rows_data.items()):
            doc_no = ""
            post_date = ""
            for c, val in sorted(cols.items()):
                clean = val.replace(" ", "").replace(",", "")
                if 6 <= len(clean) <= 12 and clean.isdigit():
                    doc_no = clean
                    break
            for c, val in cols.items():
                m = re.search(r'\d{4}\.\d{2}\.\d{2}', val)
                if m:
                    post_date = m.group()
                    break
            if doc_no and doc_no not in seen:
                seen.add(doc_no)
                docs.append({"doc_no": doc_no, "posting_date": post_date})

        try:
            session.findById(wnd_path).sendVKey(77)  # Page Down
            time.sleep(0.5)
        except Exception:
            break
        page += 1

    return docs


def _read_window_labels(session, wnd_path: str) -> dict:
    """지정 윈도우의 /usr 컨테이너에서 GuiLabel 읽기."""
    try:
        usr = session.findById(wnd_path + "/usr")
    except Exception:
        return {}
    labels = {}
    for i in range(usr.Children.Count):
        lbl = usr.Children.ElementAt(i)
        try:
            lid = lbl.Id
            bracket = lid[lid.rfind("[") + 1: lid.rfind("]")]
            c, r = map(int, bracket.split(","))
            labels[(c, r)] = lbl.Text
        except Exception:
            pass
    return labels


def get_doc_list_via_matchcode(session, fiscal_year: str,
                                debug: bool = False) -> list[dict]:
    """
    계정유지보수전표 필드 옆 작은 네모(F4 matchcode) → Enter → 목록 읽기.

    흐름:
      1. 회계연도 입력
      2. 전표번호 필드에 포커스 → F4 키 (matchcode 버튼 역할)
      3. 팝업(wnd[1])에서 Enter → 목록 조회
      4. 팝업 내 목록에서 전표번호·전기일자 추출
      5. 팝업 닫기(ESC)
    """
    # 1) 회계연도 설정 (실제 필드 ID: txtKBKP-GJAHR)
    session.findById("wnd[0]/usr/txtKBKP-GJAHR").text = fiscal_year
    time.sleep(0.2)

    # 2) 전표번호 필드(ctxtKBKP-BELNR)에서 F4 실행
    print("  전표번호 필드 F4(matchcode) 실행 중...")
    opened = False
    try:
        field = session.findById("wnd[0]/usr/ctxtKBKP-BELNR")
        field.setFocus()
        field.text = ""
        session.findById("wnd[0]").sendVKey(4)   # F4 = matchcode
        time.sleep(2.0)
        opened = True
    except Exception as e:
        print(f"  [경고] F4 실행 오류: {e}")

    if not opened:
        print("  [경고] F4 실행 실패 — 필드 ID를 찾지 못했습니다.")
        return []

    # 3) wnd[1] 검색폼: 사용자 이름 필드(SPT151) 비우기 → 전체 목록 조회
    #    (사용자 이름이 있으면 해당 사용자 문서만 나옴)
    _USER_FIELD = (
        "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001"
        "/ssubSUBSCR_PRESEL:SAPLSDH4:0220"
        "/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[4,24]"
    )
    try:
        session.findById(_USER_FIELD).text = ""
        print("  사용자 이름 필드 초기화 완료")
        time.sleep(0.3)
    except Exception:
        pass

    # Enter 실행 → 결과 목록(wnd[2]) 표시
    print("  팝업에서 Enter 실행 중...")
    try:
        session.findById("wnd[1]").sendVKey(0)   # Enter
        time.sleep(2.5)
    except Exception:
        try:
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            time.sleep(2.5)
        except Exception:
            pass

    # debug: 검색 후 전체 창 트리 덤프
    if debug:
        tree_lines = []
        for wnd_try in ["wnd[2]", "wnd[1]", "wnd[0]"]:
            try:
                obj = session.findById(wnd_try)
                tree_lines.append(f"\n=== {wnd_try} 트리 ===")
                _dump_tree_lines(obj, tree_lines)
            except Exception:
                continue
        dump_path = "matchcode_popup_debug.txt"
        with open(dump_path, "w", encoding="utf-8") as f:
            f.write("\n".join(tree_lines))
        print(f"  [DEBUG] 검색 후 트리 저장: {dump_path}")
        # 핵심 부분만 출력 (wnd[2] 위주)
        in_wnd2 = False
        for line in tree_lines:
            if "=== wnd[2]" in line:
                in_wnd2 = True
            elif line.startswith("\n=== wnd["):
                in_wnd2 = False
            if in_wnd2:
                print(line)

    # 4) 결과 목록 읽기: wnd[2] 우선, 없으면 wnd[1]
    docs: list[dict] = []
    for wnd_path in ["wnd[2]", "wnd[1]"]:
        try:
            session.findById(wnd_path)
        except Exception:
            continue
        docs = _read_list_from_window(wnd_path, session,
                                       debug=debug, page_label="검색결과")
        if docs:
            print(f"  [{wnd_path}] {len(docs)}개 전표 발견")
            break

    # 5) 팝업 전체 닫기 (wnd[2] → wnd[1] 순서로 ESC)
    for wnd_path in ["wnd[2]", "wnd[1]"]:
        try:
            session.findById(wnd_path).sendVKey(12)
            time.sleep(0.5)
        except Exception:
            continue

    return docs


def _dump_tree_lines(obj, lines: list, depth: int = 0):
    """SAP GUI 컨트롤 트리를 재귀적으로 덤프합니다."""
    indent = "  " * depth
    try:
        obj_type = obj.Type
        obj_id   = obj.Id
        line     = f"{indent}[{obj_type}] {obj_id}"
        if obj_type in ("GuiShell", "GuiGridView"):
            try:
                line += f"  ← rows={obj.RowCount}, cols={obj.ColumnCount}, subtype={obj.SubType}"
            except Exception:
                pass
        elif obj_type in ("GuiTextField", "GuiCTextField", "GuiLabel"):
            try:
                line += f"  = '{obj.Text}'"
            except Exception:
                pass
        lines.append(line)
    except Exception as e:
        lines.append(f"{indent}[오류] {e}")
        return
    try:
        for i in range(obj.Children.Count):
            _dump_tree_lines(obj.Children.ElementAt(i), lines, depth + 1)
    except Exception:
        pass


# ════════════════════════════════════════════════════════
#  Phase 3 — 단일 전표 상세 데이터 추출 (SAP 스프레드시트 내보내기 방식)
# ════════════════════════════════════════════════════════
# ── 레이블 파싱 헬퍼 (내보내기 실패 시 폴백용으로 유지) ──
def _is_a_row(labels: dict, row: int) -> bool:
    """col 1에 PO번호(6자리 이상 숫자)가 있으면 A행."""
    val   = labels.get((1, row), "").strip()
    clean = val.replace(",", "").replace(".", "").replace(" ", "")
    return len(clean) >= 6 and clean.isdigit()


def _parse_item_pair(labels: dict, a_row: int) -> dict | None:
    """A행(a_row) + B행(a_row+1) 쌍에서 레코드 추출."""
    b_row  = a_row + 1
    record = {}
    for col, field in A_COL_MAP.items():
        record[field] = _get_val(labels, col, a_row)
    for col, field in B_COL_MAP.items():
        record[field] = _get_val(labels, col, b_row)

    # 구매 문서 없으면 유효하지 않은 행
    if not record.get("구매 문서", "").strip():
        return None

    for col in NUMERIC_COLS:
        record[col] = _clean_number(record.get(col, ""))

    # 전표번호는 run()에서 외부 주입 (doc_no)
    record["전표번호"] = ""
    return {k: v for k, v in record.items() if k in MASTER_COLS}


def _parse_b_only_row(labels: dict, b_row: int) -> dict | None:
    """A행 공란(관세반제 등): B행만으로 레코드 추출."""
    sel_no = _get_val(labels, 1, b_row)
    # 짧은 순번(선택번호)이 있어야 유효한 B행
    if not sel_no.strip():
        return None

    po = _get_val(labels, 8, b_row)
    if not po.strip():
        return None   # 구매문서 없으면 빈 행 — 스킵

    record = {
        "전표번호":   "",          # run()에서 외부 주입
        "구매 문서":  po,
        "품목":       _get_val(labels, 19, b_row),
        "PO 일자":    "",
        "계정키이름": _get_val(labels, 25, b_row),
        "이름 1":     "",
        "차이 수량":  _clean_number(_get_val(labels, 46, b_row)),
        "차이 금액":  _clean_number(_get_val(labels, 65, b_row)),
        "Plnt":       "",
        "내역":       "",
        "OUn":        "",
    }
    return {k: v for k, v in record.items() if k in MASTER_COLS}


def _parse_page_records(labels: dict) -> list:
    """현재 페이지에서 레코드 전체 추출 (A+B쌍 / B단독 모두 처리)."""
    col1_rows = sorted(
        r for (c, r) in labels
        if c == 1 and r >= 4 and labels[(c, r)].strip()
    )
    processed: set = set()
    records:   list = []

    for r in col1_rows:
        if r in processed:
            continue
        if _is_a_row(labels, r):
            rec = _parse_item_pair(labels, r)
            processed.add(r)
            processed.add(r + 1)
        else:
            rec = _parse_b_only_row(labels, r)
            processed.add(r)
        if rec:
            records.append(rec)

    return records


def _sap_export_to_xls(session, doc_no: str) -> str:
    """
    현재 MR11SHOW 상세 화면을 SAP 내장 스프레드시트 내보내기로 저장합니다.

    VBScript 녹화 기준 (저장 다이얼로그 없는 버전):
      메뉴[4/5/2/2] 선택 → 형식 다이얼로그(wnd[1]) radSPOPLI-SELFLAG[4,0] → OK
      → SAP가 기억한 경로에 자동 저장 (wnd[2] 없음)

    SAP는 마지막 저장 경로를 기억하여 다음번 저장 다이얼로그를 생략합니다.
    저장된 파일은 수정시각 변화로 감지합니다.

    Returns: 저장된 파일 경로 (실패 시 "")
    """
    import glob as _glob
    userprofile = os.environ.get("USERPROFILE", "")
    desktop     = os.path.join(userprofile, "Desktop")

    # SAP가 기억한 저장 경로 후보 (수정시각 변화로 감지)
    # 바탕화면 1.XLS, 기타 공통 경로 포함
    watch_candidates = []
    for d in [desktop,
              os.environ.get("TEMP", ""),
              os.path.join(userprofile, "Documents"),
              os.path.join(userprofile, "AppData", "Local", "Temp"),
              r"C:\Users\Public\Documents\ESTsoft\CreatorTemp",
              os.getcwd()]:
        if d and os.path.isdir(d):
            for pat in ("*.xls", "*.XLS", "*.xlsx"):
                watch_candidates.extend(_glob.glob(os.path.join(d, pat)))

    # 내보내기 전 수정시각 스냅샷
    before_mtime: dict = {}
    for f in watch_candidates:
        try:
            before_mtime[f] = os.path.getmtime(f)
        except Exception:
            pass
    t_before = time.time()

    try:
        # 0) 현재 화면 확인
        try:
            title = session.findById("wnd[0]").Text
            print(f"    [화면 확인] 현재 창 제목: '{title}'")
            info = session.Info
            print(f"    [화면 확인] TCode={info.Transaction}, Program={info.Program}")
        except Exception as ec:
            print(f"    [화면 확인 오류] {ec}")

        # 1) 메뉴: 목록 → 내보내기 → 스프레드시트
        print(f"    [내보내기 1] 메뉴 선택")
        session.findById("wnd[0]/mbar/menu[4]/menu[5]/menu[2]/menu[2]").select()
        time.sleep(1.0)

        # 2) 형식 선택 다이얼로그(wnd[1]): Spreadsheet [4,0] → OK
        print(f"    [내보내기 2] 형식 선택 다이얼로그")
        _RADIO = ("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150"
                  "/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]")
        session.findById(_RADIO).select()
        session.findById(_RADIO).setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        print(f"    [내보내기 2] OK 완료, SAP 저장 대기 중...")
        time.sleep(3.0)   # SAP 자동 저장 대기

        # 3) 추가 다이얼로그 처리 (저장 경로 설정 or 덮어쓰기 확인)
        #    - 첫 실행: wnd[2] = 저장 경로 입력 다이얼로그
        #    - 재실행: wnd[2] = "파일 덮어쓰기?" 확인 다이얼로그
        for wi in range(2, 4):
            try:
                w = session.findById(f"wnd[{wi}]")
                w_title = getattr(w, "Text", "")
                print(f"    [내보내기 3] wnd[{wi}] 발견: '{w_title}'")

                # 저장 경로 입력 필드가 있으면 경로 설정
                try:
                    session.findById(f"wnd[{wi}]/usr/ctxtDY_PATH").text     = desktop + "\\"
                    session.findById(f"wnd[{wi}]/usr/ctxtDY_FILENAME").text = f"mr11_{doc_no}.xls"
                    print(f"    [내보내기 3] 저장 경로 설정: {desktop}\\mr11_{doc_no}.xls")
                except Exception:
                    print(f"    [내보내기 3] 경로 필드 없음 (덮어쓰기 확인 다이얼로그로 처리)")

                # OK/예(Yes) 버튼 클릭
                session.findById(f"wnd[{wi}]/tbar[0]/btn[0]").press()
                print(f"    [내보내기 3] OK 클릭")
                time.sleep(2.0)
                break
            except Exception:
                pass

    except Exception as e:
        print(f"    [내보내기 오류] {e}")
        for w in ("wnd[2]", "wnd[1]"):
            try:
                session.findById(w).sendVKey(12)
                time.sleep(0.3)
            except Exception:
                pass
        return ""

    # 4) 수정된 파일 감지: 이전 스냅샷보다 mtime이 최신인 파일 선택
    modified: list = []
    for f in watch_candidates:
        try:
            mt = os.path.getmtime(f)
            if mt > t_before - 1:           # 내보내기 시작 시각 이후 저장된 파일
                modified.append((mt, f))
        except Exception:
            pass

    # 새로 생긴 파일(스냅샷에 없던 것)도 포함
    for d in [desktop,
              os.environ.get("TEMP", ""),
              os.path.join(userprofile, "Documents"),
              os.path.join(userprofile, "AppData", "Local", "Temp"),
              r"C:\Users\Public\Documents\ESTsoft\CreatorTemp",
              os.getcwd()]:
        if d and os.path.isdir(d):
            for pat in ("*.xls", "*.XLS", "*.xlsx"):
                for f in _glob.glob(os.path.join(d, pat)):
                    if f not in before_mtime:
                        try:
                            modified.append((os.path.getmtime(f), f))
                        except Exception:
                            pass

    if modified:
        modified.sort(reverse=True)
        found_path = modified[0][1]
        print(f"    SAP 저장 파일 감지: {found_path}")
        return found_path

    print(f"    [경고] 내보내기 파일 감지 실패")
    return ""


def extract_single_doc(session, doc_no: str, debug: bool = False) -> list:
    """
    MR11SHOW 상세 화면(단일 전표)에서 전체 레코드를 추출합니다.

    SAP 스프레드시트 내보내기(List→Export→Spreadsheet)로 임시 XLS를 만들고
    parse_rawdata()로 파싱합니다. 스크롤 문제가 없어 데이터가 완전히 추출됩니다.
    """
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from mr11_processor import parse_rawdata

    export_path = _sap_export_to_xls(session, doc_no)
    if not export_path:
        return []

    print(f"    내보내기 완료: {os.path.basename(export_path)}")

    try:
        records = parse_rawdata(export_path)
        return records
    except Exception as e:
        print(f"    [파싱 오류] {e}")
        return []


# ── 아래는 폴백용 레이블 파싱 함수 (현재 미사용) ─────────────
def _page_anchor(labels: dict) -> tuple:
    """현재 화면의 col 1 데이터 행 값 앞 5개 — 스크롤 여부 감지용."""
    rows = sorted(r for (c, r) in labels if c == 1 and r >= 4 and labels[(c, r)].strip())
    return tuple(labels.get((1, r), "") for r in rows[:5])


def _probe_scroll_method(session, wnd, page_size: int) -> str:
    """
    여러 스크롤 방법을 차례로 시도하여 실제로 화면이 바뀌는 방법을 탐지합니다.
    page_size: 탐지에 사용할 스크롤 단위 (= 페이지당 항목 수).
    탐지 후 Ctrl+Home으로 원위치 복귀.
    반환값: 'wnd_sb' | 'usr_sb' | 'none'
    """
    labels0 = _read_page_labels(session)
    anchor0 = _page_anchor(labels0)

    def _restore():
        try:
            wnd.sendVKey(82)
            time.sleep(0.8)
        except Exception:
            pass

    # 시도할 방법: wnd_sb → usr_sb (VKey는 모두 "not enabled" 확인됨)
    candidates = [
        ("wnd_sb", lambda p=page_size: setattr(
            session.findById("wnd[0]").VerticalScrollbar, "Position", p)),
        ("usr_sb", lambda p=page_size: setattr(
            session.findById("wnd[0]/usr").VerticalScrollbar, "Position", p)),
    ]

    for name, fn in candidates:
        try:
            fn()
            time.sleep(1.2)
        except Exception:
            _restore()
            continue

        labels1 = _read_page_labels(session)
        anchor1 = _page_anchor(labels1)

        if anchor1 != anchor0:
            print(f"    스크롤 방법 확인: {name} ✓")
            _restore()
            return name
        _restore()

    return "none"


def _do_scroll(session, method: str, pos: int, page_size: int, max_pos: int) -> int:
    """탐지된 방법으로 한 페이지 스크롤. 새 pos 반환, 실패 시 -1."""
    next_pos = min(pos + page_size, max_pos) if max_pos > 0 else pos + page_size
    try:
        if method == "wnd_sb":
            session.findById("wnd[0]").VerticalScrollbar.Position = next_pos
        elif method == "usr_sb":
            session.findById("wnd[0]/usr").VerticalScrollbar.Position = next_pos
        else:
            return -1
        time.sleep(1.2)
        return next_pos
    except Exception as e:
        print(f"    [스크롤 오류] {method}: {e}")
        return -1




def input_doc_number(session, doc_no: str, fiscal_year: str):
    """MR11SHOW 초기화면에 전표번호와 회계연도를 입력합니다."""
    session.findById("wnd[0]/usr/txtKBKP-GJAHR").text = fiscal_year
    session.findById("wnd[0]/usr/ctxtKBKP-BELNR").text = doc_no
    time.sleep(0.3)


# ════════════════════════════════════════════════════════
#  메인 실행
# ════════════════════════════════════════════════════════

def run(fiscal_year: str, target_month: str,
        master_path: str, debug: bool = False):
    """MR11SHOW 자동화 전체 파이프라인."""
    try:
        from mr11_processor import append_to_master
    except ImportError:
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from mr11_processor import append_to_master

    print(f"\n{'='*55}")
    print("  MR11SHOW SAP 자동화 v3")
    print(f"{'='*55}")
    print(f"  회계연도  : {fiscal_year}")
    print(f"  대상 월   : {target_month}월")
    print(f"  마스터파일: {master_path}")
    if debug:
        print("  [디버그 모드]")
    print(f"{'='*55}\n")

    print("[1/5] SAP GUI 연결 중...")
    session = connect_sap()

    # ── Phase 1: F4 matchcode → 전표 목록 조회 ──────────────
    print(f"\n[2/5] {SAP_TCODE} 화면으로 이동 중...")
    navigate_to_tcode(session, SAP_TCODE)

    print("[3/5] 전표번호 F4(작은 네모) → Enter → 목록 읽기...")
    all_docs = get_doc_list_via_matchcode(session, fiscal_year, debug=debug)

    if not all_docs:
        print("\n  전표 목록을 읽지 못했습니다.")
        print("  --debug 옵션으로 재실행하여 화면 구조를 확인하세요.")
        return

    # ── Phase 2: 목표 월 필터링 ─────────────────────────────
    month_str  = target_month.zfill(2)
    year_month = f"{fiscal_year}.{month_str}"

    # 전기일자가 없는 경우 전표번호 범위로 추론하지 않고 전부 포함
    target_docs = [
        d for d in all_docs
        if year_month in d.get("posting_date", "") or not d.get("posting_date", "")
    ]
    # 전기일자가 확인되는 문서와 안 되는 문서 분리
    dated_docs   = [d for d in all_docs if d.get("posting_date")]
    undated_docs = [d for d in all_docs if not d.get("posting_date")]

    march_docs = [
        d for d in dated_docs
        if year_month in d.get("posting_date", "")
    ]

    print(f"\n  전체 전표   : {len(all_docs)}개")
    print(f"  전기일 확인 : {len(dated_docs)}개")
    print(f"  {fiscal_year}년 {target_month}월 : {len(march_docs)}개")
    if undated_docs:
        print(f"  전기일 미확인: {len(undated_docs)}개 (→ 목록에 날짜 없음, 전부 처리)")
        march_docs.extend(undated_docs)

    if not march_docs:
        print("\n  해당 월 전표가 없습니다.")
        if all_docs:
            print("  발견된 전표 목록:")
            for d in all_docs[:20]:
                print(f"    {d['doc_no']}  {d.get('posting_date', '날짜미확인')}")
            if len(all_docs) > 20:
                print(f"    ... 외 {len(all_docs) - 20}개")
        return

    print(f"\n  처리 대상 전표번호: {[d['doc_no'] for d in march_docs]}")

    # ── Phase 3: 각 전표 상세 추출 ──────────────────────────
    print(f"\n[4/5] 전표별 상세 데이터 추출 중...")
    all_records: list = []

    for i, doc_info in enumerate(march_docs, 1):
        doc_no    = doc_info["doc_no"]
        post_date = doc_info.get("posting_date", "날짜미확인")
        print(f"\n  [{i}/{len(march_docs)}] 전표 {doc_no} ({post_date}) ...")

        navigate_to_tcode(session, SAP_TCODE)
        input_doc_number(session, doc_no, fiscal_year)
        press_f8(session)

        records = extract_single_doc(session, doc_no, debug=debug)

        # 전표번호를 실제 계정유지보수전표 번호로 설정
        for rec in records:
            rec["전표번호"] = doc_no

        # 유효 레코드 필터: 구매문서가 숫자이고 계정키이름이 있어야 함
        def _is_valid(r: dict) -> bool:
            po = r.get("구매 문서", "").strip().replace(" ", "")
            acct = r.get("계정키이름", "").strip()
            return bool(po) and po.replace("-", "").isdigit() and bool(acct)

        records = [r for r in records if _is_valid(r)]

        doc_sum = sum(r.get("차이 금액", 0) or 0 for r in records)
        print(f"    → {len(records)}건 추출  |  차이금액 합계: {doc_sum:,.0f}")

        if debug:
            for r in records:
                print(f"       PO {r.get('구매 문서','')} / 품목 {r.get('품목','')} "
                      f"/ {r.get('계정키이름','')} / 금액 {r.get('차이 금액',0):,.0f}")

        if not records and debug:
            labels = _read_page_labels(session)
            _debug_labels(labels, f"전표 {doc_no} — 상세 화면 구조")

        all_records.extend(records)

    if not all_records:
        print("\n  추출된 데이터가 없습니다.")
        print("  --debug 옵션으로 재실행하여 컬럼 위치를 확인하세요.")
        return

    print(f"\n  총 {len(all_records):,}건 추출 완료")

    # ── Phase 4: Excel 저장 ──────────────────────────────────
    print("[5/5] Excel 저장 중...")
    result = append_to_master(all_records, master_path)

    print(f"\n{'='*55}")
    print("  완료!")
    print(f"  추가 행수 : {result['added']:,}건")
    print(f"  중복 스킵 : {result['skipped']:,}건")
    print(f"  저장 위치 : {master_path}")
    print(f"{'='*55}\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="SAP MR11SHOW 자동 추출 → 반제리스트 Excel 누적 저장 v3"
    )
    parser.add_argument("--year",  "-y", default="2026",
                        help="회계연도 (기본값: 2026)")
    parser.add_argument("--month", "-m", default="",
                        help="대상 월 숫자 (예: 3)")
    parser.add_argument("--doc", "-d", default="",
                        help="단일 전표번호 검증용 (예: 5400004827). --master 없이도 동작")
    parser.add_argument("--master", "-M", default="",
                        help="반제리스트 마스터 Excel 파일 경로 (.xlsx)")
    parser.add_argument("--debug", action="store_true",
                        help="디버그: 화면 레이블 전체 출력")
    args = parser.parse_args()

    try:
        if args.doc:
            # 단일 전표 검증 모드
            session = connect_sap()
            navigate_to_tcode(session, SAP_TCODE)
            input_doc_number(session, args.doc, args.year)
            press_f8(session)
            records = extract_single_doc(session, args.doc, debug=True)
            for rec in records:
                rec["전표번호"] = args.doc
            records = [r for r in records
                       if r.get("구매 문서", "").strip().replace(" ","").replace("-","").isdigit()
                       and r.get("계정키이름", "").strip()]
            total = sum(r.get("차이 금액", 0) or 0 for r in records)
            print(f"\n전표 {args.doc} - 총 {len(records)}건, 차이금액 합계: {total:,.0f}")
            print(f"{'구매문서':>12}  {'품목':>6}  {'계정키이름':12}  {'차이금액':>12}")
            print("-" * 55)
            for r in records:
                print(f"  {r.get('구매 문서',''):>12}  {r.get('품목',''):>6}  "
                      f"{r.get('계정키이름',''):12}  {r.get('차이 금액',0):>12,.0f}")

            # --master 지정 시 Excel 파일로 저장
            if args.master:
                sys.path.insert(0, os.path.dirname(__file__))
                from mr11_processor import append_to_master
                result = append_to_master(records, args.master)
                added   = result.get("added",   0)
                skipped = result.get("skipped", 0)
                print(f"\n[저장 완료] {args.master}")
                print(f"  추가: {added}건 / 스킵(중복): {skipped}건")
            else:
                print(f"\n[안내] Excel로 저장하려면 --master <파일명.xlsx> 옵션을 추가하세요.")
                print(f"  예) --doc {args.doc} --master 반제리스트_{args.doc}.xlsx")
        else:
            if not args.month:
                print("[오류] --month 또는 --doc 중 하나를 지정하세요.")
                sys.exit(1)
            run(
                fiscal_year=args.year,
                target_month=args.month,
                master_path=args.master,
                debug=args.debug,
            )
    except RuntimeError as e:
        print(f"\n[오류] {e}")
        sys.exit(1)
