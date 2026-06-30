# -*- coding: utf-8 -*-
r"""
accdb_import.py  --  src/ テキスト -> Access .accdb 書き戻しツール（accdb_export.py の対）

[v2 修正] LoadFromText の文字コードを“環境のSaveAsTextに自動追従”させる方式に変更。
    日本語Accessでは SaveAsText/LoadFromText が CP932(BOMなし) を使うことが多く、
    UTF-16で書き込むと文字化けして対象オブジェクトが壊れる。これを防ぐため、
    各オブジェクトを一旦 SaveAsText して BOM を判定し、同じエンコーディングで
    書き戻す（CP932 / UTF-16LE / UTF-16BE / UTF-8 いずれの環境でも安全）。

[安全設計]
    1. 反映前に .accdb を必ずタイムスタンプ付きでバックアップ
    2. 既定は“指定したファイルだけ”反映（--all で全件、ただし要確認）
    3. オブジェクトごとに現行エンコーディングを検出し、それに合わせて書き戻し
    4. 実行前に対象一覧を表示し確認（--yes でスキップ）
    5. _manifest.tsv で安全名 -> 実オブジェクト名を正確に解決

[重要な注意]
    forms / reports は抽出時にバイナリ(埋め込み画像等)を間引いている。書き戻すと
    画像が失われ得る。コードだけ直す用途では modules / queries の往復が安全。
    tables は LoadFromText 非対応のため自動スキップ。

[前提] Windows + Access と bit を揃えた Python、pywin32。リポジトリ直下で実行。
[使用例]
    py accdb_import.py MAPG.accdb src/modules/共通モジュール.txt
    py accdb_import.py MAPG.accdb src/modules/フォーマット.txt src/queries/設計原糸一覧マスタ.txt
    py accdb_import.py MAPG.accdb --all
"""

import argparse, datetime, os, shutil, sys, tempfile

try:
    import win32com.client as win32
    import pythoncom
except ImportError:
    sys.exit("pywin32 が必要です:  pip install pywin32")

AC_FORM, AC_QUERY, AC_REPORT, AC_MACRO, AC_MODULE = 2, 1, 3, 4, 5
TYPE_AC = {"forms": AC_FORM, "reports": AC_REPORT, "modules": AC_MODULE,
           "macros": AC_MACRO, "queries": AC_QUERY}
LABEL = {"forms": "フォーム", "reports": "レポート", "modules": "モジュール",
         "macros": "マクロ", "queries": "クエリ", "tables": "テーブル"}
SAFE_TYPES = {"modules", "queries"}


def load_manifest(repo_dir):
    path = os.path.join(repo_dir, "src", "_manifest.tsv")
    m = {}
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            for line in f.read().splitlines()[1:]:
                parts = line.split("\t")
                if len(parts) == 3:
                    typ, name, rel = parts
                    m[rel.replace("\\", "/")] = (typ, name)
    return m


def resolve(path, repo_dir, manifest):
    ap = os.path.abspath(path)
    rel = os.path.relpath(ap, os.path.join(repo_dir, "src")).replace("\\", "/")
    if rel in manifest:
        return manifest[rel]
    return os.path.basename(os.path.dirname(ap)), os.path.splitext(os.path.basename(ap))[0]


def detect_encoding(app, actype, name, probe_path):
    """対象を一旦 SaveAsText して BOM からエンコーディングを判定。
    返り値: (python_encoding, bom_bytes)。失敗時は CP932(BOMなし) を仮定。"""
    try:
        if os.path.exists(probe_path):
            os.remove(probe_path)
        app.SaveAsText(actype, name, probe_path)
        with open(probe_path, "rb") as f:
            head = f.read(4)
        if head[:2] == b"\xff\xfe":   return "utf-16-le", b"\xff\xfe"
        if head[:2] == b"\xfe\xff":   return "utf-16-be", b"\xfe\xff"
        if head[:3] == b"\xef\xbb\xbf": return "utf-8", b"\xef\xbb\xbf"
        return "cp932", b""           # BOMなし＝日本語Accessの既定(ANSI/CP932)
    except Exception:
        return "cp932", b""           # 対象が未作成等。新規は既定で書く


def transcode(src_utf8, dst_path, enc, bom):
    """UTF-8(LF)のsrcを、検出エンコーディング + CRLF + 適切なBOM で書き出す。"""
    with open(src_utf8, encoding="utf-8") as f:
        text = f.read()
    text = text.replace("\r\n", "\n").replace("\n", "\r\n")
    try:
        data = text.encode(enc)
    except UnicodeEncodeError as e:
        # CP932で表現できない文字がある場合はUTF-16LEへフォールバック
        return ("retry-utf16", e)
    with open(dst_path, "wb") as f:
        f.write(bom)
        f.write(data)
    return ("ok", None)


def collect_targets(args, repo_dir, manifest):
    items = []
    if args.all:
        for rel, (typ, name) in sorted(manifest.items()):
            if typ != "tables":
                items.append((os.path.join(repo_dir, "src", rel), typ, name, typ in SAFE_TYPES))
    else:
        for p in args.files:
            if not os.path.isfile(p):
                print(f"  ! 見つかりません（スキップ）: {p}"); continue
            typ, name = resolve(p, repo_dir, manifest)
            items.append((os.path.abspath(p), typ, name, typ in SAFE_TYPES))
    return items


def main():
    args = parse_args()
    accdb = os.path.abspath(args.accdb)
    if not os.path.isfile(accdb):
        sys.exit(f"見つかりません: {accdb}")
    repo_dir = os.getcwd()

    manifest = load_manifest(repo_dir)
    if not manifest:
        print("  ! _manifest.tsv が見つかりません。フォルダ名/ファイル名から型を推定します。")

    targets = collect_targets(args, repo_dir, manifest)
    if not targets:
        sys.exit("反映対象がありません。ファイルを指定するか --all を付けてください。")

    skipped_tables = [t for t in targets if t[1] == "tables"]
    unknown        = [t for t in targets if t[1] not in TYPE_AC and t[1] != "tables"]
    targets        = [t for t in targets if t[1] in TYPE_AC]

    print(f"対象DB : {accdb}\n")
    print("=== 反映予定 ===")
    risky = []
    for path, typ, name, is_safe in targets:
        mark = "  " if is_safe else "⚠ "
        if not is_safe: risky.append(name)
        print(f"  {mark}[{LABEL.get(typ, typ)}] {name}")
    for path, typ, name in [(t[0], t[1], t[2]) for t in skipped_tables]:
        print(f"  --[テーブル] {name}（LoadFromText非対応のためスキップ）")
    for path, typ, name, _ in unknown:
        print(f"  --[不明:{typ}] {name}（型を解決できずスキップ）")
    if risky:
        print("\n⚠ フォーム/レポートを反映します。抽出時に間引いた埋め込み画像/プリンタ設定は"
              "既定値で再生成され、画像は失われる可能性があります。")

    if not args.yes:
        ans = input(f"\n{len(targets)} 件を {os.path.basename(accdb)} に反映します。続行しますか? [y/N] ")
        if ans.strip().lower() not in ("y", "yes"):
            sys.exit("中止しました。")

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    base, ext = os.path.splitext(accdb)
    backup = f"{base}_backup_{ts}{ext}"
    shutil.copy2(accdb, backup)
    print(f"\nバックアップ作成: {backup}")

    tmp_dir = tempfile.mkdtemp(prefix="accdb_imp_")
    tmp_in   = os.path.join(tmp_dir, "in.txt")
    tmp_prb  = os.path.join(tmp_dir, "probe.txt")
    ok, failures = 0, []
    force_enc = {"auto": None, "cp932": ("cp932", b""),
                 "utf-16": ("utf-16-le", b"\xff\xfe")}[args.encoding]

    print("Access を起動して反映中 …")
    pythoncom.CoInitialize()
    app = win32.DispatchEx("Access.Application")
    app.Visible = False
    try:
        app.OpenCurrentDatabase(accdb, False)
        for path, typ, name, _ in targets:
            try:
                actype = TYPE_AC[typ]
                enc, bom = force_enc if force_enc else detect_encoding(app, actype, name, tmp_prb)
                status, err = transcode(path, tmp_in, enc, bom)
                if status == "retry-utf16":
                    enc, bom = "utf-16-le", b"\xff\xfe"
                    status, err = transcode(path, tmp_in, enc, bom)
                app.LoadFromText(actype, name, tmp_in)
                ok += 1
                tag = {"cp932": "CP932", "utf-16-le": "UTF-16LE",
                       "utf-16-be": "UTF-16BE", "utf-8": "UTF-8"}.get(enc, enc)
                print(f"  ✓ [{LABEL.get(typ, typ)}] {name}  ({tag})")
            except Exception as e:
                failures.append((name, str(e)))
                print(f"  ✗ [{LABEL.get(typ, typ)}] {name}: {str(e)[:80]}")
    finally:
        try: app.CloseCurrentDatabase()
        except Exception: pass
        app.Quit()
        pythoncom.CoUninitialize()
        shutil.rmtree(tmp_dir, ignore_errors=True)

    print(f"\n=== 反映完了: 成功 {ok} / 失敗 {len(failures)} ===")
    if failures:
        for name, err in failures:
            print(f"  - {name}: {err[:80]}")
    print(f"バックアップ: {backup}")
    print("\n反映後は Access の VBE で [デバッグ] → [コンパイル] を実行し、")
    print("壊れがないか必ず確認してください（自動コンパイルは行いません）。")
    print("問題があれば、上記バックアップから復元するか git で元の版に戻せます。")


def parse_args():
    p = argparse.ArgumentParser(description="src/ のテキストを LoadFromText で .accdb に書き戻す")
    p.add_argument("accdb", help="反映先 .accdb のパス")
    p.add_argument("files", nargs="*", help="反映する src/ 配下のファイル（複数可）")
    p.add_argument("--all", action="store_true", help="manifest の全オブジェクトを反映（テーブル除く・要確認）")
    p.add_argument("--yes", action="store_true", help="確認プロンプトをスキップ")
    p.add_argument("--encoding", choices=["auto", "cp932", "utf-16"], default="auto",
                   help="書き込み文字コード。auto=現行SaveAsTextに自動追従(既定)")
    return p.parse_args()


if __name__ == "__main__":
    main()
