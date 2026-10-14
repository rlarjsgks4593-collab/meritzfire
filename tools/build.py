#!/usr/bin/env python3
"""엑셀(data/latest.xlsx) + config.json → dist/index.html (설계사 코드별 암호화 데이터 포함)

사용: python3 build.py [엑셀경로]
엑셀 컬럼: 담당매니저명 / 지사명 / 설계사 코드 / 설계사명 / (…월 실적)* / (…시책금)*
 - '실적'으로 끝나는 컬럼은 모두 실적으로, '시책금'으로 끝나는 컬럼은 모두 시책 카드로 자동 인식
"""
import base64, json, os, re, sys, hashlib
import openpyxl
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

HERE = os.path.dirname(os.path.abspath(__file__))
xlsx = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, 'data', 'latest.xlsx')
cfg = json.load(open(os.path.join(HERE, 'config.json'), encoding='utf-8'))
ITER = int(cfg.get('pbkdf2_iterations', 150000))

ws = openpyxl.load_workbook(xlsx, data_only=True).active
rows = [r for r in ws.iter_rows(values_only=True)]
header = [str(h).strip() if h is not None else '' for h in rows[0]]
data = [r for r in rows[1:] if r and r[header.index('설계사 코드')] not in (None, '')]

def col(name):
    if name not in header:
        sys.exit(f'필수 컬럼이 없습니다: {name}  (현재 컬럼: {header})')
    return header.index(name)

c_mgr, c_br, c_code, c_name = col('담당매니저명'), col('지사명'), col('설계사 코드'), col('설계사명')
perf_idx = [i for i, h in enumerate(header) if h.endswith('실적')]
inc_idx = [i for i, h in enumerate(header) if h.endswith('시책금')]
if not perf_idx or not inc_idx:
    sys.exit(f'실적/시책금 컬럼을 찾지 못했습니다: {header}')

def num(v):
    if v in (None, ''): return 0
    return int(round(float(str(v).replace(',', ''))))

def norm_code(v):
    s = re.sub(r'\D', '', str(v))
    if isinstance(v, float): s = re.sub(r'\D', '', str(int(v)))
    return s

# 시책금이 특정 실적의 일정 배율이면 기준 문구 자동 생성 (예: 10월 실적 × 250%)
def basis_note(ii):
    best = None
    for pi in reversed(perf_idx):
        ratios = []
        for r in data:
            p, v = num(r[pi]), num(r[ii])
            if p > 0:
                ratios.append(v / p)
            elif v != 0:
                ratios = []; break
        if len(ratios) >= 10 and max(ratios) - min(ratios) < 1e-6:
            pct = ratios[0] * 100
            best = f'{header[pi]} × {pct:g}%'
            break
    return best

salt = os.urandom(16)
def derive(code):
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=48, salt=salt, iterations=ITER)
    return kdf.derive(code.encode())

rec, seen = {}, set()
for r in data:
    code = norm_code(r[c_code])
    if code in seen:
        print('중복 코드 건너뜀:', code); continue
    seen.add(code)
    d = derive(code)
    rid, key = d[:16].hex(), d[16:]
    payload = {
        'n': str(r[c_name]).strip(), 'b': str(r[c_br]).strip(), 'm': str(r[c_mgr]).strip(),
        'p': [num(r[i]) for i in perf_idx], 'i': [num(r[i]) for i in inc_idx],
    }
    iv = os.urandom(12)
    ct = AESGCM(key).encrypt(iv, json.dumps(payload, ensure_ascii=False).encode(), None)
    rec[rid] = base64.b64encode(iv + ct).decode()

perf_labels = [header[i] for i in perf_idx]
m = re.search(r'(\d+월)', perf_labels[-1])
meta = {
    'salt': base64.b64encode(salt).decode(), 'iter': ITER, 'rec': rec,
    'brand': cfg['brand'], 'page_title': cfg['page_title'],
    'title_login': '내 실적·시책 확인',
    'month': m.group(1) if m else '',
    'basis_date': cfg['basis_date'],
    'perf': perf_labels, 'inc': [header[i] for i in inc_idx],
    'note': [basis_note(i) for i in inc_idx],
    'footer_title': cfg['footer_title'], 'footer_lines': cfg['footer_lines'],
    'source_line': cfg['source_line'].replace('{basis_date}', cfg['basis_date']),
}
html = open(os.path.join(HERE, 'template.html'), encoding='utf-8').read()
blob = json.dumps(meta, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
html = html.replace('__DATA__', blob)
out = os.path.join(HERE, 'dist', 'index.html')
open(out, 'w', encoding='utf-8').write(html)
print(f'설계사 {len(rec)}명 · 실적 {perf_labels} · 시책 {meta["inc"]} · 기준문구 {meta["note"]}')
print(f'출력: {out} ({os.path.getsize(out)/1024:.0f} KB)')
