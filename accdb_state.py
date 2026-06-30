# -*- coding: utf-8 -*-
import os, re, sys, unicodedata
import win32com.client as win32
import pythoncom

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

def load_absent():
    items = []
    p = "deadcode_delete_list.tsv"
    if os.path.exists(p):
        for line in open(p, encoding="utf-8").read().splitlines()[1:]:
            c = line.split("\t")
            if len(c) >= 3:
                items.append((c[1], c[2]))
    return items

def main():
    accdb = sys.argv[1] if len(sys.argv) > 1 else "MAPG.accdb"
    accdb = os.path.abspath(accdb)
    if not os.path.isfile(accdb):
        sys.exit("見つかりません: " + accdb)
    absent = load_absent()
    pythoncom.CoInitialize()
    app = win32.DispatchEx("Access.Application")
    app.Visible = False
    try:
        app.OpenCurrentDatabase(accdb, False)
        try:
            proj = app.VBE.ActiveVBProject
            comps = list(proj.VBComponents)
        except Exception as e:
            sys.exit("VBAプロジェクトへアクセス不可。トラストセンターで信頼設定をONに。 " + str(e))
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
        print("\n【1】削除済みであるべき18件の検証")
        still = []
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
        ok = (not still) and (not missing)
        if ok:
            print("総合: ✓ 期待どおり（18件削除済み・サンプル復元確認）")
        else:
            if still:   print("総合: ⚠ 未削除が %d 件残存" % len(still))
            if missing: print("総合: ⚠ 復元されるべきハンドラが %d 件欠落" % len(missing))
        print("=" * 60)
    finally:
        try: app.CloseCurrentDatabase()
        except Exception: pass
        app.Quit()
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()
