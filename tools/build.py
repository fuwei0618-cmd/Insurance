#!/usr/bin/env python3
"""保險手冊 PWA 產生器

把三份 Claude 簡報（deck.json + slides/*.html + 圖片）轉成 PWA 用的資料檔。

用法：
  HANDBOOK_PASSWORD=xxxx python3 tools/build.py --src <簡報來源資料夾> --blobs <圖片資料夾> \
      --catalog <catalog.json> --selection "B94,CV4,..." --out <repo 根目錄>

<簡報來源資料夾> 內需有 ch1/project、ch2/project、ch3/project（各含 deck.json 與 slides/）。
講稿（<aside>）會用 HANDBOOK_PASSWORD 以 AES-GCM 加密後才寫入，原始碼裡看不到明文。
"""
import argparse, base64, hashlib, html, json, os, re, secrets, sys, time
from pathlib import Path

CHAPTERS = [
    (1, '第一章', '人生願景與風險觀念'),
    (2, '第二章', '保障與工具'),
    (3, '第三章', '保險規劃三步驟'),
]
CATALOG_ARTIFACT = 'https://claude.ai/artifact/WYAUYKNa4wA5EjBF9QNpYs'

# 第二章險種頁 → 商品庫分類（決定頁尾「國泰商品」顯示哪些已勾選商品）
FOOTER_RULES = {
    'ins-accident': lambda p: p['cat'] == '意外傷害',
    'ins-hosp': lambda p: p['cat'] == '住院手術' and '住院' in (p['items'] + p['name']),
    'ins-surgery': lambda p: p['cat'] == '住院手術' and '手術' in (p['items'] + p['name']),
    'ins-medexp': lambda p: p['cat'] == '實支實付',
    'ins-special': lambda p: p['cat'] == '住院手術' and ('特定處置' in p['name'] + p['desc'] + p['features']),
    'ins-cancer': lambda p: p['cat'] == '重大疾病/傷病' and ('癌' in p['name'] or '重大傷病' in p['name']),
    'ins-ci': lambda p: p['cat'] == '重大疾病/傷病' and not ('防癌' in p['name']),
    'ins-ltc': lambda p: p['cat'] == '長期照顧',
    'ins-life': lambda p: p['cat'] == '壽險',
}
CUT = ['住院', '重大', '防癌', '特定', '長期', '傷害', '定期', '終身', '利率', '醫療', '保險', '一年', '變額', '手術']


def short_name(name):
    n = re.sub(r'^國泰人壽', '', name)
    n = re.sub(r'[（(][^）)]*[）)]', '', n)
    idx = min([n.find(k) for k in CUT if n.find(k) > 1] or [len(n)])
    s = n[:idx] or n
    if '美元' in name: s += '（美元）'
    elif '澳幣' in name: s += '（澳幣）'
    return s


def footer_products(slide_id, catalog, selection):
    rule = FOOTER_RULES.get(slide_id)
    if not rule:
        return None
    names = []
    for p in catalog['products']:
        if p['code'] in selection and p.get('status') != '已下架' and rule(p):
            s = short_name(p['name'])
            if s not in names:
                names.append(s)
    return '・'.join(names) + '（依商品庫勾選）' if names else '尚未勾選商品，請到商品庫勾選'


def apply_footer(slide_id, s, catalog, selection):
    text = footer_products(slide_id, catalog, selection)
    if text is None:
        return s
    return re.sub(r'(國泰商品</p>\s*<p style="flex:1;[^"]*">)[^<]*(</p>)',
                  lambda m: m.group(1) + html.escape(text) + m.group(2), s, count=1)


def encrypt(obj, password):
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    salt, iv = secrets.token_bytes(16), secrets.token_bytes(12)
    key = hashlib.pbkdf2_hmac('sha256', password.encode(), salt, 310000, 32)
    ct = AESGCM(key).encrypt(iv, json.dumps(obj, ensure_ascii=False).encode(), None)
    b = lambda x: base64.b64encode(x).decode()
    return {'salt': b(salt), 'iv': b(iv), 'data': b(ct), 'iter': 310000}


def convert_images(blobs, out_img):
    from PIL import Image
    out_img.mkdir(parents=True, exist_ok=True)
    mapping = {}
    for f in sorted(Path(blobs).iterdir()):
        bid = f.stem
        im = Image.open(f)
        has_alpha = im.mode in ('RGBA', 'LA', 'P')
        im = im.convert('RGBA' if has_alpha else 'RGB')
        if max(im.size) > 1920:
            im.thumbnail((1920, 1920))
        dest = out_img / f'{bid}.webp'
        im.save(dest, 'WEBP', quality=84, method=6)
        mapping[bid] = f'img/{bid}.webp'
    return mapping


def slide_title(s):
    for tag in ('h1', 'h2', 'h3', 'p'):
        m = re.search(rf'<{tag}[^>]*>(.*?)</{tag}>', s, re.S)
        if m:
            t = re.sub(r'<br\s*/?>', ' ', m.group(1))
            t = html.unescape(re.sub(r'<[^>]+>', '', t)).strip()
            if t:
                return t[:40]
    return ''


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--src', required=True)
    ap.add_argument('--blobs', required=True)
    ap.add_argument('--catalog', required=True)
    ap.add_argument('--selection', required=True)
    ap.add_argument('--out', required=True)
    a = ap.parse_args()
    pw = os.environ.get('HANDBOOK_PASSWORD')
    if not pw:
        sys.exit('請設定 HANDBOOK_PASSWORD 環境變數')
    out = Path(a.out)
    catalog = json.load(open(a.catalog))
    selection = [c.strip() for c in a.selection.split(',') if c.strip()]
    img = convert_images(a.blobs, out / 'img')

    chapters, notes = [], {}
    for n, short, title in CHAPTERS:
        root = Path(a.src) / f'ch{n}' / 'project'
        deck = json.load(open(root / 'deck.json'))
        slides = []
        for sid in deck['order']:
            f = root / 'slides' / f'{sid}.html'
            if sid.startswith('tpl-') or not f.exists():
                continue
            s = f.read_text()
            if re.search(r'<section[^>]*\shidden[\s>=]', s):
                continue
            m = re.search(r'<aside>(.*?)</aside>', s, re.S)
            if m:
                notes[f'{n}/{sid}'] = html.unescape(m.group(1)).strip()
                s = s[:m.start()] + s[m.end():]
            s = re.sub(r'/_blob/([0-9a-f]{32})', lambda m: img.get(m.group(1), ''), s)
            if n == 2:
                s = apply_footer(sid, s, catalog, selection)
            s = s.replace(CATALOG_ARTIFACT + '#', '#catalog/')
            slides.append({'id': sid, 'title': slide_title(s), 'html': s.strip()})
        chapters.append({'n': n, 'short': short, 'title': title, 'slides': slides})

    version = time.strftime('%Y%m%d%H%M%S')
    data = {'version': version, 'chapters': chapters, 'notes': encrypt(notes, pw), 'selection': selection}
    (out / 'data').mkdir(exist_ok=True)
    (out / 'data' / 'handbook.js').write_text('window.HANDBOOK=' + json.dumps(data, ensure_ascii=False) + ';\n')
    (out / 'data' / 'catalog.json').write_text(json.dumps(catalog, ensure_ascii=False))
    sw = (out / 'sw.js').read_text()
    sw = re.sub(r"const VERSION = '[^']*'", f"const VERSION = '{version}'", sw)
    imgs = sorted(set(img.values()))
    sw = re.sub(r'const IMAGES = \[[^\]]*\]', 'const IMAGES = ' + json.dumps(imgs), sw)
    (out / 'sw.js').write_text(sw)
    print(f'ok {version}: ' + ', '.join(f"{c['short']} {len(c['slides'])} 頁" for c in chapters) + f'，講稿 {len(notes)} 則，圖片 {len(imgs)} 張')


if __name__ == '__main__':
    main()
