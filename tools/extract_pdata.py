"""把富邦試算 .xls 的 PDATA 工作表抽成 tables.json。

用法:
    python tools/extract_pdata.py <商品代號> <富邦.xls 路徑>

例如:
    python tools/extract_pdata.py PFA "PFA-美富紅運建議書 V2.6.xls"
    → 產生 products/PFA/tables.json

需要環境:
    - LibreOffice 已安裝 (用於把 .xls 轉成 .csv)
    - Python 3.x

抽出的表(命名範圍對應 PDATA 工作表):
    GP        每萬美元年繳費率
    DIE       身故保險金額(per 1萬)
    CV        現金價值/解約金(per 1萬)
    PA        減額繳清保額(per 1萬,前 N 年)
    PVFB      保單價值準備金(per 1萬)
    PVFB0     保單價值準備金(per 1萬,變體)
    _DIE2     身故保險金額第二版
    PV        現金價值變體
    BONU      年度保單紅利率(高/中/低三套)
    BONUDIE   終期紅利-身故率
    BONUCV    終期紅利-解約率
    EXP       (預留)

每張表的 key 格式:
    PFA + 年期(2碼) + 性別(2碼,1男/2女) + 年齡(2碼)
    例: PFA030201 = PFA 3年期 女 1歲
    BONU 系列再加 1 碼後綴: 1=低、2=中、3=高
"""
import sys
import os
import csv
import json
import re
import subprocess
import tempfile

csv.field_size_limit(10**7)


def excel_col_to_idx(letters):
    """A=0, B=1, ..., AA=26, AB=27 ..."""
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch.upper()) - ord('A') + 1)
    return n - 1


def to_num(s):
    if s == '' or s is None:
        return 0
    try:
        f = float(str(s).replace(',', ''))
        return int(f) if f.is_integer() else f
    except (ValueError, AttributeError):
        return 0


def extract_table(rows, c1_letter, r1, c2_letter, r2, key_prefix='PFA'):
    """在 CSV rows 範圍中,從 c1 開始取 key + (年度 1..N) 數值。"""
    c1 = excel_col_to_idx(c1_letter)
    c2 = excel_col_to_idx(c2_letter)
    out = {}
    for r in range(r1 - 1, r2):
        if r >= len(rows):
            break
        row = rows[r]
        if c1 >= len(row):
            continue
        key = row[c1]
        if not key or not key.startswith(key_prefix):
            continue
        vals = []
        for c in range(c1 + 1, c2 + 1):
            vals.append(to_num(row[c] if c < len(row) else ''))
        # 去掉尾端連續的 0 以節省空間
        while vals and vals[-1] == 0:
            vals.pop()
        out[key] = vals
    return out


def convert_xls_to_csv(xls_path, out_dir):
    """用 LibreOffice headless 把 .xls 的 PDATA 轉成 CSV。"""
    print(f"[1/3] 用 LibreOffice 轉 {xls_path} 為 CSV...")
    cmd = [
        'soffice', '--headless',
        '--convert-to',
        'csv:Text - txt - csv (StarCalc):44,34,76,1,,0,false,true,false,false,false,1',
        xls_path, '--outdir', out_dir,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    if result.returncode != 0:
        print('STDERR:', result.stderr)
        raise RuntimeError(f"LibreOffice 轉換失敗 (returncode={result.returncode})")
    # 找出 PDATA 那個 CSV
    base = os.path.splitext(os.path.basename(xls_path))[0]
    pdata_csv = os.path.join(out_dir, f"{base}-PDATA.csv")
    if not os.path.exists(pdata_csv):
        # 可能整本只有單一 sheet 名沒 -PDATA 後綴,fallback 找任何 .csv
        candidates = [f for f in os.listdir(out_dir) if f.endswith('.csv')]
        if not candidates:
            raise RuntimeError("找不到 PDATA CSV 輸出")
        pdata_csv = os.path.join(out_dir, candidates[0])
    return pdata_csv


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    code = sys.argv[1]
    xls_path = sys.argv[2]
    if not os.path.exists(xls_path):
        print(f"找不到檔案: {xls_path}")
        sys.exit(1)

    with tempfile.TemporaryDirectory() as tmp:
        csv_path = convert_xls_to_csv(xls_path, tmp)
        print(f"[2/3] 讀取 CSV: {csv_path}")
        with open(csv_path, 'r', encoding='utf-8') as f:
            rows = list(csv.reader(f))
        print(f"      共 {len(rows):,} 列, {len(rows[0]) if rows else 0} 欄")

        print("[3/3] 抽取命名範圍...")

        # PDATA 中各命名範圍位置(對應 PFA 系列;若新商品有不同位置,請另行調整)
        ranges = {
            'PA':       ('Q', 2,    'DX', 306),
            'DIE':      ('ED', 313, 'IK', 617),
            'CV':       ('Q', 626,  'DX', 930),
            'PVFB':     ('Q', 936,  'DY', 1239),
            'PVFB0':    ('ED', 936, 'IL', 1239),
            'SRV':      ('ED', 2,   'IK', 120),
            '_DIE2':    ('Q', 1244, 'DX', 1548),
            'PV':       ('ED', 1244,'IK', 1548),
            'BONU':     ('R', 1554, 'DY', 2445),
            'BONUDIE':  ('R', 2457, 'DY', 3348),
            'BONUCV':   ('EF', 2457,'IM', 3348),
            'EXP':      ('Q', 313,  'DX', 617),
        }

        tables = {}
        for name, (c1, r1, c2, r2) in ranges.items():
            t = extract_table(rows, c1, r1, c2, r2, key_prefix=code)
            tables[name] = t
            sample = next(iter(t.values()), [])
            print(f"  {name}: {len(t)} keys, 樣本長度 {len(sample)}")

        # GP (費率) 從前段直接讀
        # 列 1 開始: 商品代號(C1) + 商品代號性別年齡(C9) + 保費(C10)
        gp = {}
        for r in range(2, len(rows)):
            row = rows[r]
            if len(row) <= 10:
                continue
            prod = row[1]
            key = row[9]
            rate = to_num(row[10])
            if prod and prod.startswith(code) and key.startswith(code) and rate:
                gp[key] = rate
        tables['GP'] = gp
        print(f"  GP: {len(gp)} keys")

    out_dir = os.path.join('products', code)
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, 'tables.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(tables, f, ensure_ascii=False, separators=(',', ':'))
    print(f"\n✓ 已寫入 {out_path} ({os.path.getsize(out_path):,} bytes)")


if __name__ == '__main__':
    main()
