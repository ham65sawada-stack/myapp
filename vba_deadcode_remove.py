# -*- coding: utf-8 -*-
r"""
vba_deadcode_remove.py  --  デッドコード(プロシージャ)を .accdb から外科的に削除

[方式] VBE オブジェクトモデルの CodeModule.DeleteLines を使い、対象プロシージャの
    行範囲だけを削除する。フォーム全体を LoadFromText し直さないため、埋め込み画像や
    レイアウト等のバイナリは一切失われない。行番号は保存値を使わず、削除直前に
    ProcStartLine/ProcCountLines で“現在値”を都度取得するためコード変動に強い。

[安全設計]
    1. --apply 時は .accdb をタイムスタンプ付きでバックアップ（1回）
    2. 既定はドライラン（計画表示のみ・無変更）。実削除は --apply が必須
    3. 1プロシージャずつ名前で現在位置を再取得 → DeleteLines（順序非依存・ドリフト耐性）
    4. --sync-src で src/ のテキストからも同プロシージャを除去（gitレビュー用ミラー）
    5. --git で 1プロシージャ=1コミット（--sync-src 前提）。.accdb の復元はバックアップで

[v2 修正]
    - .laccdb 検出時は中止（「Accessを閉じてから実行」の機械的担保）
    - 確認プロンプトをバックアップ作成より前に移動（中止時にバックアップのゴミを残さない）
    - ドライラン時も --sync-src の対象有無を表示（src 側の削除計画が事前に見える）
    - kind 列の表記ゆれを許容（"Property Get" → PropertyGet）。不明値は黙って
      Sub 扱いにせず、理由付きでスキップ表示
    - 集計の二重計上を解消: .accdb 側・src 側それぞれの“実削除”で計上し、
      src のみ削除された件数を別掲（従来は accdb スキップでも done に入っていた）
    - DeleteLines の例外を捕捉して失敗として集計。失敗ありなら終了コード 1
    - 出力を UTF-8 に統一（リダイレクト時の ✓/⚠ での UnicodeEncodeError を防止）

[前提]
    - Windows + Access、bit を揃えた Python、pywin32
    - 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を ON にすること
      Access オプション → トラスト センター → トラスト センターの設定
        → マクロの設定 → 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」
    - リポジトリ直下で実行。削除リストは deadcode_delete_list.tsv（編集して取捨選択可）

[使用例]
    py vba_deadcode_remove.py MAPG.accdb                       # ドライラン（既定）
    py vba_deadcode_remove.py MAPG.accdb --apply               # .accdb から削除
    py vba_deadcode_remove.py MAPG.accdb --apply --sync-src --git   # src同期＋1件1コミット
"""

import argparse, csv, datetime, os, re, shutil, subprocess, sys, unicodedata

try:
    import win32com.client as win32
    import pythoncom
except ImportError:
    sys.exit("pywin32 が必要です:  pip install pywin32")

norm = lambda s: unicodedata.normalize("NFKC", s)
PK = {"Sub": 0, "Function": 0, "PropertyGet": 3, "PropertyLet": 1, "PropertySet": 2}  # vbext_pk_*


def _utf8_console():
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def check_not_locked(accdb):
    lock = os.path.splitext(accdb)[0] + ".laccdb"
    if os.path.exists(lock):
        sys.exit(f".laccdb を検出: {lock}\n"
                 "Access がこのDBを開いています。Access を閉じてから再実行してください。\n"
                 "（Access 終了後も残る場合は前回異常終了の残骸なので、この .laccdb を削除して構いません）")


def pk_of(kind_field):
    """TSVのkind列 -> vbext_pk_*。空/None は 0 (Sub/Function共通)。不明値は None。"""
    k = (kind_field or "").strip().replace(" ", "")
    if k == "":
        return 0
    return PK.get(k)


def load_list(path):
    rows = []
    with open(path, encoding="utf-8") as f:
        for r in csv.DictReader(f, delimiter="\t"):
            rows.append(r)
    return rows


def get_project(vbe, accdb):
    target = os.path.normcase(os.path.abspath(accdb))
    for proj in vbe.VBProjects:
        try:
            if proj.FileName and os.path.normcase(os.path.abspath(proj.FileName)) == target:
                return proj
        except Exception:
            pass
    return vbe.VBProjects(1)


def find_component(proj, module_name):
    want = norm(module_name)
    for comp in proj.VBComponents:
        if norm(comp.Name) == want:
            return comp
    return None


def src_index(repo_dir):
    """NFKC正規化した basename -> src パス。フォーム/レポート/モジュールを横断。"""
    idx = {}
    for typ, pref in [("forms", "Form_"), ("reports", "Report_"), ("modules", "")]:
        d = os.path.join(repo_dir, "src", typ)
        if not os.path.isdir(d):
            continue
        for fn in os.listdir(d):
            if fn.lower().endswith(".txt"):
                base = fn[:-4]
                idx[(typ, norm(base))] = os.path.join(d, fn)
    return idx


def src_path_for(module, idx):
    if module.startswith("Form_"):
        return idx.get(("forms", norm(module[5:])))
    if module.startswith("Report_"):
        return idx.get(("reports", norm(module[7:])))
    return idx.get(("modules", norm(module)))


def _proc_hdr_re(proc):
    return re.compile(rf"^\s*(?:Public|Private|Friend)?\s*(?:Static\s+)?"
                      rf"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+{re.escape(proc)}\b", re.I)


def proc_in_text(path, proc):
    """src テキストに proc のヘッダ行が存在するか（無変更・ドライラン用）。"""
    hdr = _proc_hdr_re(proc)
    with open(path, encoding="utf-8") as f:
        for line in f:
            if hdr.match(line):
                return True
    return False


def remove_proc_from_text(path, proc):
    """src テキストから proc の Sub/Function/Property ブロックを除去。除去できたら True。"""
    with open(path, encoding="utf-8") as f:
        lines = f.read().split("\n")
    hdr = _proc_hdr_re(proc)
    endr = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\b", re.I)
    out, i, removed = [], 0, False
    while i < len(lines):
        if hdr.match(lines[i]):
            j = i
            while j < len(lines) and not endr.match(lines[j]):
                j += 1
            i = j + 1
            removed = True
            while i < len(lines) and lines[i].strip() == "":   # 後続の空行を1つ詰める
                i += 1
            continue
        out.append(lines[i]); i += 1
    if removed:
        with open(path, "w", encoding="utf-8", newline="") as f:
            f.write("\n".join(out))
    return removed


def git(args, repo_dir):
    return subprocess.run(["git"] + args, cwd=repo_dir,
                          capture_output=True, text=True, encoding="utf-8")


def main():
    _utf8_console()
    ap = argparse.ArgumentParser(
        description="デッドコードを .accdb から外科的に削除（終了コード 0=成功/1=失敗あり）")
    ap.add_argument("accdb")
    ap.add_argument("--list", default="deadcode_delete_list.tsv", help="削除リストTSV")
    ap.add_argument("--apply", action="store_true", help="実削除（既定はドライラン）")
    ap.add_argument("--sync-src", action="store_true", help="src/ テキストからも除去")
    ap.add_argument("--git", action="store_true", help="1プロシージャ=1コミット（--sync-src前提）")
    ap.add_argument("--yes", action="store_true", help="確認をスキップ")
    args = ap.parse_args()

    accdb = os.path.abspath(args.accdb)
    repo_dir = os.getcwd()
    if not os.path.isfile(accdb):
        sys.exit(f"見つかりません: {accdb}")
    if not os.path.isfile(args.list):
        sys.exit(f"削除リストが見つかりません: {args.list}")
    if args.git and not args.sync_src:
        sys.exit("--git は --sync-src と併用してください（コミット対象のテキスト差分が必要）。")
    check_not_locked(accdb)

    targets = load_list(args.list)
    print(f"削除リスト: {len(targets)} 件 ({args.list})")
    print(f"対象DB    : {accdb}")
    print(f"モード    : {'実削除(APPLY)' if args.apply else 'ドライラン(無変更)'}"
          f"{' +src同期' if args.sync_src else ''}{' +git' if args.git else ''}\n")

    srcidx = src_index(repo_dir) if args.sync_src else {}

    pythoncom.CoInitialize()
    app = win32.DispatchEx("Access.Application")
    app.Visible = False
    backup = None
    done, skipped, errors, src_only = [], [], [], 0
    try:
        app.OpenCurrentDatabase(accdb, False)
        try:
            vbe = app.VBE
            proj = get_project(vbe, accdb)
            _ = proj.VBComponents.Count            # ここでアクセス確認
        except Exception as e:
            sys.exit("VBProject にアクセスできません。Accessのトラストセンターで"
                     "「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」をONにしてください。\n"
                     f"  詳細: {str(e)[:100]}")

        # 確認 → バックアップ（実削除時のみ・1回）
        if args.apply:
            if not args.yes:
                ans = input(f"{len(targets)} 件のプロシージャを削除します。続行しますか? [y/N] ")
                if ans.strip().lower() not in ("y", "yes"):
                    sys.exit("中止しました。")
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            base, ext = os.path.splitext(accdb)
            backup = f"{base}_backup_{ts}{ext}"
            # OpenCurrentDatabase 中でもファイルコピーは可能（排他でなければ）。失敗時は閉じてコピー。
            try:
                shutil.copy2(accdb, backup)
            except Exception:
                app.CloseCurrentDatabase(); shutil.copy2(accdb, backup); app.OpenCurrentDatabase(accdb, False)
            print(f"バックアップ作成: {backup}\n")

        # モジュール毎にまとめ、同一モジュール内は開始行降順で処理（再取得もするので順序は安全側）
        by_mod = {}
        for t in targets:
            by_mod.setdefault(t["module"], []).append(t)

        for module, items in by_mod.items():
            comp = find_component(proj, module)
            cm = comp.CodeModule if comp else None
            spath = src_path_for(module, srcidx) if args.sync_src else None

            # 現在の開始行で降順
            def cur_start(it):
                try:
                    k = pk_of(it.get("kind"))
                    return cm.ProcStartLine(it["proc"], 0 if k is None else k) if cm else -1
                except Exception:
                    return -1

            for it in sorted(items, key=cur_start, reverse=True):
                proc = it["proc"]
                label = f"[{module}] {proc}"
                kind = pk_of(it.get("kind"))
                if kind is None:
                    skipped.append((label, f"不明なkind: {it.get('kind')!r}"))
                    print(f"  ？skip {label}  (kind={it.get('kind')!r})")
                    continue
                acc_ok, src_removed = False, False

                # .accdb 側
                if cm is None:
                    skipped.append((label, "モジュール未解決"))
                    print(f"  －skip {label}  (モジュール未解決)")
                else:
                    try:
                        start = cm.ProcStartLine(proc, kind)
                        count = cm.ProcCountLines(proc, kind)
                    except Exception:
                        start = count = None
                    if not start or not count:
                        skipped.append((label, ".accdbにプロシージャなし(削除済/改名?)"))
                        print(f"  －skip {label}  (.accdbに無し)")
                    else:
                        if args.apply:
                            try:
                                cm.DeleteLines(start, count)
                                acc_ok = True
                                print(f"  ✓削除 {label}  (行 {start}〜{start+count-1}, {count}行)")
                            except Exception as e:
                                errors.append((label, str(e)))
                                print(f"  ✗失敗 {label}  ({str(e)[:60]})")
                        else:
                            print(f"  ・対象 {label}  (行 {start}〜{start+count-1}, {count}行)")

                # src 側
                if args.sync_src:
                    if not spath:
                        print("        src: 対応ファイル未検出")
                    elif args.apply:
                        src_removed = remove_proc_from_text(spath, proc)
                        print(f"        src: {'除去' if src_removed else '該当なし'}")
                        if src_removed and args.git:
                            rel = os.path.relpath(spath, repo_dir)
                            git(["add", rel], repo_dir)
                            git(["commit", "-m", f"deadcode削除: {module}.{proc}"], repo_dir)
                    else:
                        print(f"        src: {'あり(除去予定)' if proc_in_text(spath, proc) else '該当なし'}")

                if args.apply:
                    if acc_ok or src_removed:
                        done.append(label)
                    if src_removed and not acc_ok:
                        src_only += 1

    finally:
        try:
            app.CloseCurrentDatabase()
        except Exception:
            pass
        app.Quit(); pythoncom.CoUninitialize()

    print(f"\n=== {'削除完了' if args.apply else 'ドライラン完了'} ===")
    if args.apply:
        note = f"（うちsrcのみ {src_only}）" if src_only else ""
        print(f"  対象 {len(targets)} / 削除 {len(done)}{note} / スキップ {len(skipped)} / 失敗 {len(errors)}")
    else:
        print(f"  対象 {len(targets)} / 削除予定 {len(targets)-len(skipped)} / スキップ {len(skipped)}")
    for lbl, why in skipped[:30]:
        print(f"  - skip: {lbl}  ({why})")
    for lbl, why in errors:
        print(f"  - 失敗: {lbl}  ({why[:80]})")
    if backup:
        print(f"\nバックアップ: {backup}")
    if args.apply:
        print("\nAccess の VBE で [デバッグ] → [コンパイル] を実行し、エラーが無いか確認してください。")
        if not args.sync_src:
            print("src/ を更新するには accdb_export.py を再実行してください。")
        print("問題があればバックアップから .accdb を復元できます。")
    else:
        print("\n実削除するには --apply を付けて再実行してください。")
    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()
