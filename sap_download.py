# -*- coding: utf-8 -*-
"""
SAP MR11SHOW 전표 5400004827 → XLS 다운로드
VBScript 녹화를 Python으로 1:1 변환
"""

import sys, os, time, glob
sys.stdout.reconfigure(encoding='utf-8')

import win32com.client

# ── SAP 연결 ──────────────────────────────────────────────
sap_gui    = win32com.client.GetObject("SAPGUI")
application = sap_gui.GetScriptingEngine
connection  = application.Children(0)
session     = connection.Children(0)
print(f"SAP 연결: {session.Info.SystemName} / {session.Info.User}")

# ── VBScript 원문 그대로 ────────────────────────────────
session.findById("wnd[0]").maximize()

session.findById("wnd[0]/usr/ctxtKBKP-BELNR").text = ""

# F4 — 매치코드 팝업 열기
session.findById("wnd[0]").sendVKey(4)
time.sleep(1.5)

# 매치코드 팝업 필터 초기화 후 검색
TAB = ("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001"
       "/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220")

fld5 = session.findById(f"{TAB}/txtG_SELFLD_TAB-LOW[5,24]")
fld5.text = ""
fld5.setFocus()
fld5.caretPosition = 0
session.findById("wnd[1]").sendVKey(0)
time.sleep(1.0)

fld4 = session.findById(f"{TAB}/txtG_SELFLD_TAB-LOW[4,24]")
fld4.text = ""
fld4.setFocus()
fld4.caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[0]").press()
time.sleep(2.0)

# 12번째 행 선택 (= 전표 5400004827)
lbl = session.findById("wnd[1]/usr/lbl[1,12]")
lbl.setFocus()
lbl.caretPosition = 8
session.findById("wnd[1]").sendVKey(2)   # 선택 전송
time.sleep(1.0)

# 메인 화면 실행 (Enter)
session.findById("wnd[0]").sendVKey(0)
time.sleep(3.0)

print("상세 화면 로딩 완료")

# ── 내보내기 ────────────────────────────────────────────
# 수정시각 스냅샷 (저장 위치 감지용)
desktop = os.path.join(os.environ.get("USERPROFILE", ""), "Desktop")
snap = {}
for d in [desktop,
          os.environ.get("TEMP",""),
          os.path.join(os.environ.get("USERPROFILE",""), "Documents")]:
    if d and os.path.isdir(d):
        for f in glob.glob(os.path.join(d, "*.xls")) + glob.glob(os.path.join(d, "*.XLS")):
            try:
                snap[f] = os.path.getmtime(f)
            except Exception:
                pass
t0 = time.time()

# 메뉴: 목록 → 내보내기 → 스프레드시트
session.findById("wnd[0]/mbar/menu[4]/menu[5]/menu[2]/menu[2]").select()
time.sleep(1.0)

# 형식 선택: Spreadsheet [4,0]
RADIO = ("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150"
         "/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]")
session.findById(RADIO).select()
session.findById(RADIO).setFocus()
session.findById("wnd[1]/tbar[0]/btn[0]").press()
time.sleep(2.0)

# 저장 다이얼로그 확인 (Enter로 승인)
shell = win32com.client.Dispatch("WScript.Shell")
shell.SendKeys("{ENTER}", 0)
print("저장 확인 Enter 전송")
time.sleep(3.0)

# ── 저장된 파일 찾기 ────────────────────────────────────
found = None

# 1) 수정된 파일 감지
for d in [desktop,
          os.environ.get("TEMP",""),
          os.path.join(os.environ.get("USERPROFILE",""), "Documents")]:
    if not d or not os.path.isdir(d):
        continue
    for pat in ("*.xls", "*.XLS", "*.xlsx"):
        for f in glob.glob(os.path.join(d, pat)):
            try:
                mt = os.path.getmtime(f)
                if mt >= t0 or (f in snap and mt > snap[f]):
                    found = f
                    break
            except Exception:
                pass
        if found:
            break
    if found:
        break

if not found:
    print("[경고] 파일 감지 실패. 바탕화면 확인 필요")
    sys.exit(1)

print(f"저장된 파일: {found}  ({os.path.getsize(found):,} bytes)")

# ── 프로젝트 폴더로 복사 ────────────────────────────────
import shutil
dest = os.path.join(os.path.dirname(__file__), "5400004827_raw.xls")
shutil.copy2(found, dest)
print(f"복사 완료: {dest}")
print("완료!")
