# -*- coding: utf-8 -*-
r"""
accdb_state.py  --  .accdb の現在のVBA状態を読み取り専用で点検する

[v2 修正]
    - deadcode_delete_list.tsv を「accdbと同じフォルダ → カレント」の順で探索
      （従来はカレント固定。リポジトリ外から実行すると未検出→0件のまま
      「✓ すべて削除済みを確認」と“偽の緑”が出ていた）
    - リスト未検出/有効行なしの場合は【1】をスキップして ⚠ を明示
    - 見出しの「18件」ハードコードを実件数(len)に変更
    - 終了コード: 期待どおり 0 / NGあり・リスト未検証 1 / 実行エラー 2
      （バッチや pre-commit フックに組み込める）
    - --list で削除リストを明示指定可能
    - 出力を UTF-8 に統一（リダイレクト時の ✓/⚠ での UnicodeEncodeError を防止）

[前提] 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」ON。
       SaveAsText/LoadFromText/DeleteLines は一切呼ばない純粋な読取（無変更・バックアップ不要）。
[使用例]
    py accdb_state.py MAPG.accdb
    py accdb_state.py MAPG.accdb --list deadcode_delete_list.tsv
"""

import argparse, os, re, sys, unicodedata

try:
    import win32com.client as win32
    import pythoncom
except ImportError:
    sys.exit("pywin32 が必要です:  pip install pywin32")

norm = lambda s: unicodedata.normalize("NFKC", s)
HDR = re.compile(r"(?im)^\s*(?:Public|Private|Friend|Global)?\s*(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([^\s(]+)")

PRESENT_SAMPLE = [
    ("Form_試編生機出荷案内書", "落傷1_LostFocus"),
    ("Form_試編生機出荷案内書", "針折1_LostFocus"),
    ("Form_試編生機出荷案内書", "メーカ_AfterUpdate"),
    ("Form_設計Ｈメンテ",       "編機CD_AfterUpdate"),
    ("Form_設計Ｈメンテ",       "規格性量_仕上巾_LostFocus"),
    ("Form_設計Ｈメンテ",       "QR公称巾_AfterUpdate"),
    ("Form_設計Ｂメンテ",       "給糸NOU01_AfterUpdate"),
    ("Form_設計Ｂメンテ",       "原糸区分_Change"),
    ("Form_試編出荷案内書メール送信確認", "メールするCC牧野_Click"),
]


def _utf8_console():
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def find_list(accdb, explicit):
    """削除リストTSVの所在を解決。明示指定が無ければ accdbフォルダ→カレント の順で探索。"""
    if explicit:
        p = os.path.abspath(explicit)
        if not os.path.isfile(p):
            print("削除リストが見つかりません: " + p, file=sys.stderr)
            sys.exit(2)
        return p
    for base in (os.path.dirname(os.path.abspath(accdb)), os.getcwd()):
        p = os.path.join(base, "deadcode_delete_list.tsv")
        if os.path.isfile(p):
            return p
    return None


def load_absent(path):
    items = []
    for line in open(path, encoding="utf-8").read().splitlines()[1:]:
        c = line.split("\t")
        if len(c) >= 3:
            items.append((c[1], c[2]))
    return items


def parse_args():
    p = argparse.ArgumentParser(
        description=".accdb のVBA状態を読み取り専用で点検（終了コード 0=OK / 1=NG・未検証 / 2=エラー）")
    p.add_argument("accdb", nargs="?", default="MAPG.accdb", help="対象 .accdb（既定: MAPG.accdb）")
    p.add_argument("--list", help="削除リストTSVを明示指定（既定: accdbフォルダ→カレントの順に自動探索）")
    return p.parse_args()


def main():
    _utf8_console()
    args = parse_args()
    accdb = os.path.abspath(args.accdb)
    if not os.path.isfile(accdb):
        print("見つかりません: " + accdb, file=sys.stderr)
        return 2

    list_path = find_list(accdb, args.list)
    absent = load_absent(list_path) if list_path else []

    rc = 2
    pythoncom.CoInitialize()
    app = win32.DispatchEx("Access.Application")
    app.Visible = False
    try:
        app.OpenCurrentDatabase(accdb, False)
        try:
            proj = app.VBE.ActiveVBProject
            comps = list(proj.VBComponents)
        except Exception as e:
            print("VBAプロジェクトへアクセス不可。トラストセンターで信頼設定をONに。 " + str(e),
                  file=sys.stderr)
            return 2
        procs = {}
        for c in comps:
            cm = c.CodeModule
            n = cm.CountOfLines
            txt = cm.Lines(1, n) if n > 0 else ""
            procs[c.Name] = {norm(m.group(1)) for m in HDR.finditer(txt)}

        def has(mod, proc):
            for cn, ps in procs.items():
                if norm(cn) == norm(mod):
                    return norm(proc) in ps
            return None

        total = sum(len(v) for v in procs.values())
        print("=" * 60)
        print("対象: " + accdb)
        print("VBAコンポーネント数: %d / プロシージャ総数(概算): %d" % (len(procs), total))
        print("=" * 60)

        # 【1】削除済み検証（リストが正しく読めた時のみ判定に算入）
        still, list_checked = [], False
        if list_path is None:
            print("\n【1】削除済み検証: スキップ")
            print("  ⚠ deadcode_delete_list.tsv が見つかりません（accdbフォルダ→カレントを探索）")
        elif not absent:
            print("\n【1】削除済み検証: スキップ")
            print("  ⚠ 削除リストに有効な行がありません: " + list_path)
        else:
            list_checked = True
            print("\n【1】削除済みであるべき %d 件の検証  (%s)" % (len(absent), os.path.basename(list_path)))
            for mod, proc in absent:
                r = has(mod, proc)
                if r is True:
                    still.append((mod, proc)); print("  ⚠ まだ存在: [%s] %s" % (mod, proc))
                elif r is None:
                    print("  ?  モジュール未検出: [%s] %s" % (mod, proc))
            if not still:
                print("  ✓ %d件すべて削除済みを確認" % len(absent))

        print("\n【2】復元されているべき束縛ハンドラ（131件の代表サンプル）")
        missing = []
        for mod, proc in PRESENT_SAMPLE:
            r = has(mod, proc)
            if r is True:
                print("  ✓ 存在: [%s] %s" % (mod, proc))
            elif r is False:
                missing.append((mod, proc)); print("  ⚠ 欠落!: [%s] %s" % (mod, proc))
            else:
                print("  ?  モジュール未検出: [%s] %s" % (mod, proc))

        print("\n【3】コンポーネント別プロシージャ数（上位15）")
        for cn, ps in sorted(procs.items(), key=lambda kv: -len(kv[1]))[:15]:
            print("  %4d  %s" % (len(ps), cn))

        print("\n" + "=" * 60)
        ok = list_checked and (not still) and (not missing)
        if ok:
            print("総合: ✓ 期待どおり（%d件削除済み・サンプル復元確認）" % len(absent))
        else:
            if not list_checked:
                print("総合: ⚠ 削除リスト未検証（TSV未検出または空）")
            if still:
                print("総合: ⚠ 未削除が %d 件残存" % len(still))
            if missing:
                print("総合: ⚠ 復元されるべきハンドラが %d 件欠落" % len(missing))
        print("=" * 60)
        rc = 0 if ok else 1
        return rc
    finally:
        try:
            app.CloseCurrentDatabase()
        except Exception:
            pass
        app.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    sys.exit(main())
