#!/usr/bin/env python3
"""월별 암호화 데이터(months/*.json) + template.html + config.json → dist/index.html

사용
  python3 build.py ingest 엑셀.xlsx [--basis "11월 3일"]   # 엑셀 1개를 해당 월 데이터로 저장(같은 달이면 덮어쓰기) 후 페이지 빌드
  python3 build.py                                        # 저장된 월별 데이터로 페이지만 다시 빌드 (문구 수정 등)

규칙
 - 엑셀 컬럼: 담당매니저명 / 지사명 / 설계사 코드 / 설계사명 / (…실적)* / (…시책금)*
 - 엑셀의 '맨 마지막 …월 실적' 컬럼이 그 파일의 월이 된다. 같은 월을 다시 올리면 그 월이 갱신되고, 이전 월은 그대로 고정된다.
 - months/*.json 에는 설계사 코드로 암호화된 값만 들어 있어 저장소에 올려도 된다. 엑셀 원본은 저장소에 올리지 않는다.
 - salt.txt 는 한 번 만들면 바꾸지 않는다 (바꾸면 저장된 월 데이터가 모두 열리지 않는다).
"""
import argparse, base64, json, os, re, sys
import openpyxl
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

HERE = os.path.dirname(os.path.abspath(__file__))
MONTHS = os.path.join(HERE, 'months')
OUT = os.environ.get('BUILD_OUT') or os.path.join(HERE, 'dist')
cfg = json.load(open(os.path.join(HERE, 'config.json'), encoding='utf-8'))
ITER = int(cfg.get('pbkdf2_iterations', 150000))


def get_salt():
    p = os.path.join(HERE, 'salt.txt')
    if not os.path.exists(p):
        open(p, 'w').write(base64.b64encode(os.urandom(16)).decode())
        print('salt.txt 새로 생성 (이후 바꾸지 마세요)')
    return base64.b64decode(open(p).read().strip())


def num(v):
    if v in (None, ''):
        return 0
    return int(round(float(str(v).replace(',', ''))))


def norm_code(v):
    if isinstance(v, float):
        v = int(v)
    return re.sub(r'\D', '', str(v))


def derive(code, salt):
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=48, salt=salt, iterations=ITER)
    return kdf.derive(code.encode())


def basis_from_name(path):
    m = re.search(r'(\d{2})(\d{2})(\d{2})', os.path.basename(path))
    return f'{int(m.group(2))}월 {int(m.group(3))}일' if m else None


def make_month(xlsx, basis, salt):
    ws = openpyxl.load_workbook(xlsx, data_only=True).active
    rows = [r for r in ws.iter_rows(values_only=True)]
    header = [str(h).strip() if h is not None else '' for h in rows[0]]

    def col(name):
        if name not in header:
            sys.exit(f'필수 컬럼이 없습니다: {name}  (현재 컬럼: {header})')
        return header.index(name)

    c_mgr, c_br, c_code, c_name = col('담당매니저명'), col('지사명'), col('설계사 코드'), col('설계사명')
    data = [r for r in rows[1:] if r and r[c_code] not in (None, '')]
    perf_idx = [i for i, h in enumerate(header) if h.endswith('실적')]
    fixed = {c_mgr, c_br, c_code, c_name}
    # 고정 컬럼과 '…실적' 컬럼을 뺀 나머지 이름 있는 컬럼은 모두 시책 카드로 본다 (…시책금, …합계 등)
    inc_idx = [i for i, h in enumerate(header) if h and i not in fixed and not h.endswith('실적')]
    if not perf_idx or not inc_idx:
        sys.exit(f'실적/시책 컬럼을 찾지 못했습니다: {header}')
    perf_labels = [header[i] for i in perf_idx]
    mm = re.search(r'(\d+월)', perf_labels[-1])
    if not mm:
        sys.exit(f'마지막 실적 컬럼에서 월을 찾지 못했습니다: {perf_labels[-1]}')
    month = mm.group(1)

    def basis_note(ii):
        for pi in reversed(perf_idx):
            ratios = []
            for r in data:
                p, v = num(r[pi]), num(r[ii])
                if p > 0:
                    ratios.append(v / p)
                elif v != 0:
                    ratios = []
                    break
            if len(ratios) >= 10 and max(ratios) - min(ratios) < 1e-6:
                return f'{header[pi]} × {ratios[0] * 100:g}%'
        return None

    rec, seen = {}, set()
    for r in data:
        code = norm_code(r[c_code])
        if code in seen:
            print('중복 코드 건너뜀:', code)
            continue
        seen.add(code)
        d = derive(code, salt)
        payload = {'n': str(r[c_name]).strip(), 'b': str(r[c_br]).strip(), 'm': str(r[c_mgr]).strip(),
                   'p': [num(r[i]) for i in perf_idx], 'i': [num(r[i]) for i in inc_idx]}
        iv = os.urandom(12)
        ct = AESGCM(d[16:]).encrypt(iv, json.dumps(payload, ensure_ascii=False).encode(), None)
        rec[d[:16].hex()] = base64.b64encode(iv + ct).decode()
    blob = {'month': month, 'basis_date': basis, 'perf': perf_labels,
            'inc': [header[i] for i in inc_idx], 'note': [basis_note(i) for i in inc_idx], 'rec': rec}
    print(f'[{month}] 설계사 {len(rec)}명 · 실적 {perf_labels} · 시책 {blob["inc"]} · 기준문구 {blob["note"]} · 기준일 {basis}')
    return blob


def sort_months(order):
    # 월 번호순 정렬. 연말~연초(예: 12월, 1월)가 같이 있으면 1~6월을 다음 해로 본다.
    nums = [int(re.match(r'(\d+)', m).group(1)) for m in order]
    wrap = any(n >= 10 for n in nums) and any(n <= 3 for n in nums)
    key = lambda m: (lambda n: n + 12 if wrap and n <= 6 else n)(int(re.match(r'(\d+)', m).group(1)))
    return sorted(order, key=key)


def order_path():
    return os.path.join(MONTHS, 'index.json')


def load_order():
    p = order_path()
    return json.load(open(p, encoding='utf-8'))['order'] if os.path.exists(p) else []


def ingest(xlsx, basis):
    os.makedirs(MONTHS, exist_ok=True)
    blob = make_month(xlsx, basis or basis_from_name(xlsx) or '', get_salt())
    if not blob['basis_date']:
        sys.exit('기준일을 알 수 없습니다. --basis "11월 3일" 처럼 지정하세요.')
    json.dump(blob, open(os.path.join(MONTHS, blob['month'] + '.json'), 'w', encoding='utf-8'),
              ensure_ascii=False, separators=(',', ':'))
    order = load_order()
    if blob['month'] not in order:
        order.append(blob['month'])
    order = sort_months(order)
    json.dump({'order': order}, open(order_path(), 'w', encoding='utf-8'), ensure_ascii=False)


def build_site():
    salt = get_salt()
    order = load_order()
    if not order:
        sys.exit('저장된 월 데이터가 없습니다. 먼저 ingest 로 엑셀을 넣어주세요.')
    months = []
    for i, label in enumerate(order):
        b = json.load(open(os.path.join(MONTHS, label + '.json'), encoding='utf-8'))
        months.append({'label': label, 'basis_date': b['basis_date'], 'final': i < len(order) - 1,
                       'perf': b['perf'], 'inc': [cfg.get('inc_labels', {}).get(x, x) for x in b['inc']],
                       'note': b['note'], 'rec': b['rec']})
    tabs = [{'l': m['label']} for m in months]
    for p in cfg.get('pending_months', []):
        if p not in order:
            tabs.append({'l': p, 'pending': True})
    meta = {
        'salt': base64.b64encode(salt).decode(), 'iter': ITER,
        'brand': cfg['brand'], 'page_title': cfg['page_title'], 'title_login': '내 실적·시책 확인',
        'footer_title': cfg['footer_title'], 'footer_lines': cfg['footer_lines'],
        'source_line': cfg['source_line'], 'tabs': tabs, 'months': months,
    }
    pr = cfg.get('promo')
    if pr:
        meta['promo'] = {'name': pr['name'], 'desc': pr.get('desc', ''), 'target': int(pr['target']), 'months': pr['months']}
    html = open(os.path.join(HERE, 'template.html'), encoding='utf-8').read()
    blob = json.dumps(meta, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
    html = html.replace('__DATA__', blob)
    os.makedirs(OUT, exist_ok=True)
    out = os.path.join(OUT, 'index.html')
    open(out, 'w', encoding='utf-8').write(html)
    print(f'출력: {out} ({os.path.getsize(out) / 1024:.0f} KB) · 월 {[t["l"] + ("(집계 전)" if t.get("pending") else "") for t in tabs]}')
    for folder in cfg.get('legacy_redirects', []):
        d = os.path.join(OUT, folder)
        os.makedirs(d, exist_ok=True)
        open(os.path.join(d, 'index.html'), 'w', encoding='utf-8').write(
            '<!doctype html><html lang="ko"><head><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width,initial-scale=1">'
            '<meta http-equiv="refresh" content="0; url=../"><title>메리츠화재 시책 현황</title></head>'
            '<body style="font-family:system-ui,sans-serif;padding:24px">'
            '<p>주소가 하나로 합쳐졌어요. <a href="../">시책 현황 열기</a></p>'
            '<script>location.replace("../")</script></body></html>')
        print(f'이전 주소 이동 페이지: {d}/index.html')


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest='cmd')
    a = sub.add_parser('ingest')
    a.add_argument('xlsx')
    a.add_argument('--basis')
    args = ap.parse_args()
    if args.cmd == 'ingest':
        ingest(args.xlsx, args.basis)
    build_site()
