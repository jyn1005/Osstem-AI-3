# -*- coding: utf-8 -*-
"""
MR11SHOW ALV Grid / Classic List 경로 탐색 도구
SAP에서 MR11SHOW F8 결과 화면을 열어둔 상태에서 실행하세요.
결과를 alv_tree_mr11show.txt 파일로 저장합니다.

사용법:
    python src/modules/analytics/find_alv_mr11show.py
"""

import win32com.client

OUTPUT_FILE = "alv_tree_mr11show.txt"


def dump_tree(obj, lines: list, depth: int = 0):
    indent = "  " * depth
    try:
        obj_type = obj.Type
        obj_id   = obj.Id
        line     = f"{indent}[{obj_type}] {obj_id}"

        # GuiShell / GuiGridView 이면 행·열 정보 추가
        if obj_type in ("GuiShell", "GuiGridView"):
            try:
                line += f"  ← rows={obj.RowCount}, cols={obj.ColumnCount}, subtype={obj.SubType}"
            except Exception:
                try:
                    line += f"  ← subtype={obj.SubType}"
                except Exception:
                    pass

        lines.append(line)

    except Exception as e:
        lines.append(f"{indent}[읽기 오류] {e}")
        return

    # 자식 노드 재귀 탐색
    try:
        for i in range(obj.Children.Count):
            dump_tree(obj.Children.ElementAt(i), lines, depth + 1)
    except Exception:
        pass


def main():
    print("SAP GUI에 연결 중...")
    sap_gui = win32com.client.GetObject("SAPGUI")
    app     = sap_gui.GetScriptingEngine
    conn    = app.Children(0)
    session = conn.Children(0)
    print(f"연결 성공: {session.Info.SystemName} / {session.Info.User}")

    # ── 1. 컨트롤 트리 덤프 ──────────────────────────────────
    print("컨트롤 트리 덤프 중...")
    tree_lines = []
    dump_tree(session.findById("wnd[0]"), tree_lines)

    # ── 2. tbar[1] 버튼 툴팁 수집 ───────────────────────────
    tbar_lines = ["", "=" * 60, "【tbar[1] 버튼 툴팁】", "=" * 60]
    try:
        wnd  = session.findById("wnd[0]")
        tbar = wnd.findById("tbar[1]")
        for i in range(tbar.Children.Count):
            btn = tbar.Children.ElementAt(i)
            try:
                tbar_lines.append(
                    f"  btn[{btn.Id.split('/')[-1]}]  "
                    f"tooltip={btn.Tooltip}  text={btn.Text}"
                )
            except Exception:
                pass
    except Exception as e:
        tbar_lines.append(f"  툴팁 읽기 오류: {e}")

    # ── 3. GuiLabel 위치 + 값 전체 수집 ─────────────────────
    label_lines = ["", "=" * 60, "【GuiLabel 위치 및 값 전체】", "=" * 60]
    try:
        usr = session.findById("wnd[0]/usr")
        label_data = []
        for i in range(usr.Children.Count):
            lbl = usr.Children.ElementAt(i)
            try:
                lid    = lbl.Id
                bracket = lid[lid.rfind("[") + 1: lid.rfind("]")]
                c, r   = map(int, bracket.split(","))
                text   = lbl.Text
                label_data.append((r, c, text))
            except Exception:
                pass

        label_data.sort()  # (row, col, text) 순 정렬
        for r, c, text in label_data:
            label_lines.append(f"  lbl[{c:>4},{r:>3}] = '{text}'")

    except Exception as e:
        label_lines.append(f"  레이블 읽기 오류: {e}")

    # ── 파일 저장 ────────────────────────────────────────────
    all_lines = tree_lines + tbar_lines + label_lines
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write("\n".join(all_lines))

    print(f"\n저장 완료: {OUTPUT_FILE}")
    print(f"총 {len(tree_lines)}개 컨트롤 발견\n")

    # ── 콘솔 출력 ────────────────────────────────────────────
    print("=" * 60)
    print("【ALV Grid 후보 경로】")
    print("=" * 60)
    alv_found = [l for l in tree_lines if "GuiShell" in l or "GuiGridView" in l]
    if alv_found:
        print("\n".join(alv_found))
    else:
        print("없음 — Classic SAP ABAP List 화면 (GuiLabel 기반)")

    print()
    for line in tbar_lines:
        print(line)

    print()
    print("=" * 60)
    print("【GuiLabel 값 (비어있지 않은 것, row/col 순)】")
    print("=" * 60)
    for r, c, text in [(r, c, t) for r, c, t in label_data if t.strip()]:
        print(f"  lbl[{c:>4},{r:>3}] = '{text}'")


if __name__ == "__main__":
    main()
