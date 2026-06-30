# -*- coding: utf-8 -*-
r"""
accdb_export.py  --  Access .accdb -> Git最適化テキスト抽出ツール

[目的]
    Access 365 (.accdb) の全コードオブジェクトを SaveAsText で抽出し、
    Git/GitHub での差分・レビュー・最適化に適した形に正規化して
    src/ ツリーへ出力する。VBAリファクタリングの基盤を作る。

[抽出対象]
    フォーム / レポート / 標準・クラスモジュール / マクロ / クエリ(SQL)
    （フォーム・レポート背後のVBAは各オブジェクトの出力内に含まれる）
    + テーブルスキーマ(DAO; --no-tables で抑止)

[Gitを汚す“ノイズ”を自動除去]
    常に除去:
        Checksum            … 保存の度に変化
        PrtDevMode/Names(W) … プリンタドライバ依存。マシンが違うだけで差分が出る
        NameMap / GUID      … バイナリGUIDブロック（再生成される）
    既定で除去（--keep-media で保持）:
        PictureData/OLEData/ImageData/dbLongBinary/dbBinary … 埋め込みバイナリ
    保持:
        PrtMip … 余白・段組などのページ幾何（レポートで意味を持つ・変化が少ない）

[重要な位置づけ]
    本ツールの出力は「完全な再構築用ソース」ではなく
    “コードレビュー/差分/最適化のためのミラー”。
    埋め込み画像やプリンタ設定などのバイナリ資産は .accdb 側が正本。
    個別オブジェクトの変更は LoadFromText で書き戻す運用（別スクリプト）。

[前提]
    Windows + Access と bit を揃えた Python（64bit Access -> 64bit Python）
        pip install pywin32

[使用例]
    python accdb_export.py "C:\work\myapp\M2発注一覧.accdb"
    python accdb_export.py ".\app.accdb" --out src --commit "VBA抽出: 初回"
    python accdb_export.py ".\app.accdb" --keep-media   # バイナリも残す
"""

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile

try:
    import win32com.client as win32
    import pythoncom
except ImportError:
    sys.exit("pywin32 が必要です:  pip install pywin32")


# --- Access acObjectType 定数 (SaveAsText 用) --------------------------------
AC_FORM, AC_QUERY, AC_REPORT, AC_MACRO, AC_MODULE = 2, 1, 3, 4, 5

SUBDIR = {
    AC_FORM:   "forms",
    AC_REPORT: "reports",
    AC_MODULE: "modules",
    AC_MACRO:  "macros",
    AC_QUERY:  "queries",
}
LABEL = {"forms": "フォーム", "reports": "レポート", "modules": "モジュール",
         "macros": "マクロ", "queries": "クエリ", "tables": "テーブル"}

# --- DAO データ型 -> 表示名 ---------------------------------------------------
DAO_TYPE = {
    1: "Boolean", 2: "Byte", 3: "Integer", 4: "Long", 5: "Currency",
    6: "Single", 7: "Double", 8: "Date", 9: "Binary", 10: "Text",
    11: "LongBinary(OLE)", 12: "Memo", 15: "GUID", 16: "BigInt",
    17: "VarBinary", 18: "Char", 19: "Numeric", 20: "Decimal",
    101: "Attachment", 109: "ComplexText",
}

# --- 正規化（ノイズ除去）-----------------------------------------------------
_DEVICE_KEYS = ("PrtDevMode", "PrtDevNames", "PrtDevModeW", "PrtDevNamesW",
                "NameMap", "GUID")
_MEDIA_KEYS = ("PictureData", "OLEData", "ImageData", "dbLongBinary", "dbBinary")
_CHECKSUM = re.compile(r'^\s*Checksum\s*=\s*-?\d+\s*$', re.I)
_END = re.compile(r'^\s*End\s*$', re.I)

# Windowsファイル名に使えない文字
_FORBIDDEN = re.compile(r'[\\/:*?"<>|]')


def make_block_matcher(strip_media: bool):
    """`Prop = Begin` 形式のバイナリブロック開始行を判定する関数を返す。"""
    keys = list(_DEVICE_KEYS) + (list(_MEDIA_KEYS) if strip_media else [])
    pat = re.compile(r'^\s*(' + '|'.join(keys) + r')\b', re.I)

    def is_block_start(line: str) -> bool:
        return bool(pat.match(line)) and line.rstrip().lower().endswith("begin")

    return is_block_start


def read_text_any(path: str) -> str:
    """SaveAsText の出力エンコーディング差（UTF-16 / CP932 / UTF-8）を吸収して読む。"""
    raw = open(path, "rb").read()
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return raw.decode("utf-16")
    for enc in ("cp932", "utf-8-sig", "utf-8"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("cp932", errors="replace")


def normalize_to_file(src_raw: str, dst: str, is_block_start) -> int:
    """ノイズ除去＋末尾空白除去のうえ UTF-8/LF で書き出す。書き込みバイト数を返す。"""
    text = read_text_any(src_raw)
    out, skip = [], False
    for ln in text.splitlines():
        if skip:                       # バイナリブロック内：End まで読み飛ばす
            if _END.match(ln):
                skip = False
            continue
        if is_block_start(ln):
            skip = True
            continue
        if _CHECKSUM.match(ln):
            continue
        out.append(ln.rstrip())        # 末尾空白を落として差分を安定化
    data = ("\n".join(out) + "\n").encode("utf-8")
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    with open(dst, "wb") as f:
        f.write(data)
    return len(data)


def safe_name(name: str) -> str:
    return _FORBIDDEN.sub("_", name)


# --- 抽出本体 ----------------------------------------------------------------
def export_objects(app, db, out_dir, is_block_start, manifest):
    """コードオブジェクトを SaveAsText -> 正規化 -> src/ へ。"""
    tmp = tempfile.mkdtemp(prefix="accdb_raw_")
    raw = os.path.join(tmp, "o.txt")
    counts = {v: 0 for v in SUBDIR.values()}
    failures = []
    proj = app.CurrentProject

    def do(obj_type, name):
        try:
            if os.path.exists(raw):
                os.remove(raw)
            app.SaveAsText(obj_type, name, raw)
            sub = SUBDIR[obj_type]
            dst = os.path.join(out_dir, sub, safe_name(name) + ".txt")
            normalize_to_file(raw, dst, is_block_start)
            manifest.append((sub, name, f"{sub}/{safe_name(name)}.txt"))
            counts[sub] += 1
        except Exception as e:           # 1件の失敗で全体を止めない
            failures.append((name, str(e)))

    for f in proj.AllForms:
        do(AC_FORM, f.Name)
    for r in proj.AllReports:
        do(AC_REPORT, r.Name)
    for m in proj.AllModules:            # 標準・クラスモジュール（背後コードは各オブジェクト内）
        do(AC_MODULE, m.Name)
    for mc in proj.AllMacros:
        do(AC_MACRO, mc.Name)
    for q in db.QueryDefs:               # システム/一時クエリは除外
        nm = q.Name
        if nm.startswith("~") or nm.startswith("MSys"):
            continue
        do(AC_QUERY, nm)

    shutil.rmtree(tmp, ignore_errors=True)
    return counts, failures


def export_tables(db, out_dir, manifest):
    """テーブルスキーマ（列・型・索引）をテキスト化。データは含めない。"""
    n = 0
    for td in db.TableDefs:
        name = td.Name
        if name.startswith("MSys") or name.startswith("~") or name.startswith("USys"):
            continue
        lines = [f"TABLE\t{name}"]
        try:
            if td.Connect:               # リンクテーブルは接続情報も記録
                lines.append(f"CONNECT\t{td.Connect}")
                lines.append(f"SOURCE\t{td.SourceTableName}")
            for fld in td.Fields:
                t = DAO_TYPE.get(fld.Type, f"Type{fld.Type}")
                lines.append(
                    f"FIELD\t{fld.Name}\t{t}\tSize={fld.Size}\t"
                    f"Req={fld.Required}\tZeroLen={fld.AllowZeroLength}")
            for idx in td.Indexes:
                cols = ",".join(f.Name for f in idx.Fields)
                lines.append(
                    f"INDEX\t{idx.Name}\t({cols})\tPK={idx.Primary}\tUnique={idx.Unique}")
        except Exception as e:
            lines.append(f"# 取得エラー: {e}")
        dst = os.path.join(out_dir, "tables", safe_name(name) + ".txt")
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        with open(dst, "wb") as fp:
            fp.write(("\n".join(lines) + "\n").encode("utf-8"))
        manifest.append(("tables", name, f"tables/{safe_name(name)}.txt"))
        n += 1
    return n


# --- 付帯ファイル ------------------------------------------------------------
def write_manifest(out_dir, manifest):
    """安全名 <-> 実オブジェクト名 の対応。LoadFromText で正確に書き戻すための索引。"""
    path = os.path.join(out_dir, "_manifest.tsv")
    with open(path, "wb") as f:
        f.write("type\toriginal_name\tfile\n".encode("utf-8"))
        for typ, name, rel in sorted(manifest):
            f.write(f"{typ}\t{name}\t{rel}\n".encode("utf-8"))


def ensure_gitattributes(repo_dir):
    """UTF-8/LF を Git に明示。既存があれば触らない。"""
    path = os.path.join(repo_dir, ".gitattributes")
    if os.path.exists(path):
        print(f"  .gitattributes は既存のため変更せず: {path}")
        return
    content = (
        "# Access SaveAsText 抽出物（accdb_export.py が UTF-8/LF 正規化）\n"
        "src/**/*.txt        text eol=lf\n"
        "src/_manifest.tsv   text eol=lf\n"
        "*.accdb   binary\n"
        "*.laccdb  binary\n"
    )
    with open(path, "wb") as f:
        f.write(content.encode("utf-8"))
    print(f"  .gitattributes を作成: {path}")


def git_commit(repo_dir, msg):
    rel = os.path.basename(repo_dir)  # unused; kept for clarity
    subprocess.run(["git", "-C", repo_dir, "add", "src", ".gitattributes"], check=False)
    subprocess.run(["git", "-C", repo_dir, "commit", "-m", msg], check=False)
    print("  git commit 実行。push は安全のため手動で:  git push")


# --- レポート ----------------------------------------------------------------
def report(counts, n_tbl, failures, out_dir):
    print("\n=== 抽出完了 ===")
    print(f"  出力先: {out_dir}")
    for sub in SUBDIR.values():
        print(f"  {LABEL[sub]:<6}: {counts[sub]:>4} 件")
    if n_tbl:
        print(f"  {LABEL['tables']:<6}: {n_tbl:>4} 件")
    total = sum(counts.values()) + n_tbl
    print(f"  合計  : {total:>4} オブジェクト")
    if failures:
        print(f"\n  ! 失敗 {len(failures)} 件:")
        for name, err in failures:
            print(f"    - {name}: {err[:80]}")


# --- エントリポイント --------------------------------------------------------
def parse_args():
    p = argparse.ArgumentParser(
        description="Access .accdb をGit最適化テキストに抽出する")
    p.add_argument("accdb", help="対象 .accdb のパス")
    p.add_argument("--out", default="src", help="出力先ディレクトリ（既定: src）")
    p.add_argument("--no-tables", action="store_true",
                   help="テーブルスキーマを抽出しない")
    p.add_argument("--keep-media", action="store_true",
                   help="埋め込み画像/OLE等のバイナリも残す（既定は除去）")
    p.add_argument("--commit", metavar="MSG",
                   help="抽出後に git add src && git commit -m MSG（push はしない）")
    return p.parse_args()


def main():
    args = parse_args()
    accdb = os.path.abspath(args.accdb)
    if not os.path.isfile(accdb):
        sys.exit(f"見つかりません: {accdb}")

    out_dir = os.path.abspath(args.out)
    repo_dir = os.path.dirname(out_dir)

    # 既存 src/ をクリア（accdb から削除済みのオブジェクトを残さない＝削除も差分に乗せる）
    for sub in list(SUBDIR.values()) + ["tables"]:
        d = os.path.join(out_dir, sub)
        if os.path.isdir(d):
            shutil.rmtree(d)

    print(f"対象 : {accdb}")
    print("Access を起動して抽出中 …（AutoExecマクロや起動フォームがある場合は注意）")

    pythoncom.CoInitialize()
    app = win32.DispatchEx("Access.Application")   # 既存セッションと干渉しない専用プロセス
    app.Visible = False
    manifest = []
    try:
        app.OpenCurrentDatabase(accdb, False)
        db = app.CurrentDb()
        is_block = make_block_matcher(strip_media=not args.keep_media)
        counts, failures = export_objects(app, db, out_dir, is_block, manifest)
        n_tbl = 0 if args.no_tables else export_tables(db, out_dir, manifest)
    finally:
        try:
            app.CloseCurrentDatabase()
        except Exception:
            pass
        app.Quit()
        pythoncom.CoUninitialize()

    write_manifest(out_dir, manifest)
    print("\n=== Git設定 ===")
    ensure_gitattributes(repo_dir)
    report(counts, n_tbl, failures, out_dir)

    if args.commit:
        print("\n=== コミット ===")
        git_commit(repo_dir, args.commit)

    print("\n次の手順:")
    print("  git status        # 差分を確認")
    print("  git add src .gitattributes")
    print('  git commit -m "VBA抽出"')
    print("  git push")


if __name__ == "__main__":
    main()
