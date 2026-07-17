# -*- coding: utf-8 -*-
import pandas as pd, sys, json, datetime
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

TEST_DATE = datetime.date(2026, 3, 1)

def calc_age(birth_roc):
    try:
        s = str(int(birth_roc)).zfill(7)
        y = int(s[:3]) + 1911
        m = int(s[3:5])
        d = int(s[5:7])
        bd = datetime.date(y, m, d)
        age = (TEST_DATE - bd).days // 365
        return max(10, min(12, age))
    except:
        return None

def safe_float(v):
    if v is None: return None
    try:
        f = float(v)
        return None if f != f else f
    except:
        return None

def run_to_seconds(v):
    if v is None: return None
    try:
        f = float(v)
        if f != f: return None
        mins = int(f)
        secs = round((f - mins) * 100)
        return mins * 60 + secs
    except:
        return None

def safe_int_str(v):
    if v is None: return ''
    try:
        f = float(v)
        return '' if f != f else str(int(f))
    except:
        return str(v)

XLSM = r'G:\我的雲端硬碟\🎾體育組\114學年度體育組\114學年度體適能檢測\體適能檢測各班資料\114學年度體適能檢測資料各班上傳資料檔 (1).xlsm'

# Grade-to-default-age mapping (used when 出生年月日 not present)
GRADE_AGE = {'4': 10, '5': 11, '6': 12}

xf = pd.ExcelFile(XLSM, engine='openpyxl')
rows = []
for sheet in xf.sheet_names:
    if sheet[0] not in '456':
        continue
    df = pd.read_excel(XLSM, sheet_name=sheet, header=0, engine='openpyxl')
    cols = [str(c).strip() for c in df.columns.tolist()]
    has_birth = '出生年月日' in cols

    for _, row in df.iterrows():
        r = row.tolist()
        # Detect 班級 (first numeric column)
        cls_val = safe_int_str(r[0]) if r else ''
        if not cls_val or not cls_val.isdigit():
            continue

        # Use column-name-based lookup
        def col(name):
            try:
                idx = cols.index(name)
                return r[idx]
            except ValueError:
                return None

        birth = col('出生年月日') if has_birth else None
        age = calc_age(birth) if birth is not None else GRADE_AGE.get(sheet[0])
        gender = str(col('性別') or '').strip()

        rows.append({
            'cls': cls_val,
            'no': int(float(col('座號'))) if col('座號') is not None else 0,
            'sid': safe_int_str(col('學號')),
            'name': str(col('姓名') or ''),
            'gender': gender,
            'age': age,
            'bend':  safe_float(col('坐姿體前彎')),
            'jump':  safe_float(col('立定跳遠')),
            'situp': safe_float(col('仰臥起坐')),
            'run':   run_to_seconds(col('800公尺跑走')),
        })

with open(r'G:\我的雲端硬碟\pe-tools\fitness_data.json', 'w', encoding='utf-8') as f:
    json.dump(rows, f, ensure_ascii=False)
print(f'Exported {len(rows)} rows')
# Show sample of 601 class
sample = [r for r in rows if r['cls'] == '601'][:3]
for s in sample:
    print(s)
