#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""PPT 渲染台 · 纯后端命令行（融合版）

一个程序，贯穿「素材 -> DeepSeek -> pptx」两步：

    scan    <文件夹>            扫描素材，出文件名清单（终端 + 剪贴板 + _文件名清单.txt）
    check   <spec.json> [素材]  静态校验 deck-spec，不产文件
    render  <spec.json> [素材]  渲染 pptx（deck-engine.js + 中文字体后处理）
    guide                       查看 deck-spec 格式速查

    直接拖文件夹或 .json 到本程序上 = 自动识别（scan / render）

融合来源
    1) Desktop\\ppt制作\\deepseek_python_20260917_830ee5.py   素材文件名清单
    2) D:\\工具\\ppt渲染台\\json2ppt.py + JSON转PPT.bat        JSON 转 PPT

依赖：仅 Python 标准库；render 需 node + D:\\工具\\ppt_work\\deck-engine.js
"""
import json
import os
import re
import shutil
import subprocess
import sys
import time
import unicodedata
import zipfile

# ──────────────────────────── 配置 ────────────────────────────
APP = 'PPT 渲染台'
VER = '1.0'

ROOT = os.path.dirname(os.path.abspath(__file__))
_LOCAL_ENGINE = os.path.join(ROOT, 'engine', 'deck-engine.js')
ENGINE = os.environ.get('PPT_ENGINE') or _LOCAL_ENGINE
SPECDIR = os.path.join(ROOT, '输出')
DESKTOP = os.path.join(os.path.expanduser('~'), 'Desktop')
FALLBACK_ASSETS = r'./demo-assets'
LIST_NAME = '_文件名清单.txt'

# 页面类型 -> 必填字段。运行时优先从 deck-engine.js 现读，读不到才用这张兜底表。
FALLBACK_TYPES = {
    'cover': ['title'], 'intro': ['title', 'rows'],
    'product': ['title', 'images', 'rows'], 'scene': ['image', 'chip'],
    'symbols': ['title', 'rows', 'images'], 'colors': ['title', 'swatches', 'rows'],
    'finale': ['title', 'thumbs', 'quote'],
    # 二期新增 18 种（引擎里现读的 TYPE_FIELDS 才是准的，这里只是读不到时的兜底）
    'toc': ['title', 'rows'], 'section': ['title'], 'bullets': ['title', 'rows'],
    'imageText': ['title', 'image', 'rows'], 'cards': ['title', 'cards'],
    'compare': ['title', 'columns'], 'timeline': ['title', 'nodes'], 'process': ['title', 'steps'],
    'matrix': ['title', 'quadrants'], 'pyramid': ['title', 'levels'], 'stats': ['title', 'stats'],
    'chart': ['title', 'chart'], 'table': ['title', 'head', 'rows'], 'quote': ['quote'],
    'team': ['title', 'images'], 'gallery': ['title', 'images'],
    'case': ['title', 'image', 'rows'], 'contact': ['title'],
}

# ──────────────────────────── 输出层 ────────────────────────────
_WIDTH_MIN, _WIDTH_MAX = 64, 108
_COLOR_ON = True
_CODES = {'red': 31, 'grn': 32, 'yel': 33, 'blu': 34,
          'mag': 35, 'cyn': 36, 'gry': 90, 'whi': 97}
_ANSI_RE = re.compile(r'\x1b\[[0-9;]*m')


def init_console():
    """Windows 控制台开 VT 色 + UTF-8；失败静默降级。"""
    global _COLOR_ON
    if os.name == 'nt':
        try:
            import ctypes
            k = ctypes.windll.kernel32
            h = k.GetStdHandle(-11)
            mode = ctypes.c_uint32()
            if k.GetConsoleMode(h, ctypes.byref(mode)):
                k.SetConsoleMode(h, mode.value | 0x0004)
        except Exception:
            _COLOR_ON = False
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding='utf-8', errors='replace')
        except (AttributeError, OSError):
            pass
    if os.environ.get('NO_COLOR') or not sys.stdout.isatty():
        _COLOR_ON = False


def cx(name, s):
    if not _COLOR_ON or name not in _CODES:
        return s
    return '\033[%dm%s\033[0m' % (_CODES[name], s)


def dw(s):
    """显示宽度：中日韩全角算 2 列。"""
    s = _ANSI_RE.sub('', s)
    n = 0
    for ch in s:
        if unicodedata.combining(ch):
            continue
        n += 2 if unicodedata.east_asian_width(ch) in ('W', 'F') else 1
    return n


def pad(s, n, align='<'):
    d = dw(s)
    if d >= n:
        return s
    gap = n - d
    if align == '>':
        return ' ' * gap + s
    if align == '^':
        l = gap // 2
        return ' ' * l + s + ' ' * (gap - l)
    return s + ' ' * gap


def trunc(s, n):
    if dw(s) <= n:
        return s
    out, w = '', 0
    for ch in (s if len(s) <= n else s[:n]):
        c = 2 if unicodedata.east_asian_width(ch) in ('W', 'F') else 1
        if w + c > n - 2:
            break
        out += ch
        w += c
    return out + '..'


def term_w():
    try:
        w = shutil.get_terminal_size().columns
    except Exception:
        w = 96
    return max(_WIDTH_MIN, min(w, _WIDTH_MAX))


def hr(label=''):
    """整宽分隔线；给 label 时居中嵌在线上。"""
    w = term_w()
    if not label:
        sys.stdout.write(cx('gry', '─' * w) + '\n')
        return
    t = ' ' + label + ' '
    n = max(w - dw(t), 0)
    l = n // 2
    sys.stdout.write(cx('gry', '─' * l) + cx('cyn', t) + cx('gry', '─' * (n - l)) + '\n')


def kv(k, v, kw=8, indent=2, vcolor=None):
    body = cx(vcolor, v) if vcolor else v
    sys.stdout.write(' ' * indent + cx('gry', pad(k, kw)) + '  ' + body + '\n')


def blank():
    sys.stdout.write('\n')


def step(idx, total, label, status, detail=''):
    """进度行：[1/6] 读取 spec ........ OK    4 页 · 2.0 KB —— OK 固定在同一列。"""
    w = term_w()
    col = max(w - 34, 42)            # OK 的起始列
    head = '  [%d/%d] %s ' % (idx, total, label)
    scolor = {'OK': 'grn', 'WARN': 'yel', 'FAIL': 'red', 'SKIP': 'gry'}.get(status, 'whi')
    stat = cx(scolor, pad(status, 4))
    tail = ('   ' + trunc(detail, max(w - col - 8, 12))) if detail else ''
    room = col - dw(head) - 1
    dots = cx('gry', '.' * room) if room >= 3 else ''
    sys.stdout.write(head + dots + ' ' + stat + tail + '\n')


def note(text, color='gry', indent=2):
    sys.stdout.write(' ' * indent + cx(color, text) + '\n')


def alert(title, lines, color='red'):
    """错误/提示框：左边界 + 底线，不闭合右边，避免宽度算错。"""
    w = term_w()
    top = '┌ ' + title + ' '
    sys.stdout.write('  ' + cx(color, top + '─' * max(w - 4 - dw(top), 0)) + '\n')
    for ln in lines:
        for sub in (ln.split('\n') if ln else ['']):
            sys.stdout.write('  ' + cx(color, '│') + ' ' + trunc(sub, w - 6) + '\n')
    sys.stdout.write('  ' + cx(color, '└' + '─' * max(w - 5, 10)) + '\n')


def die(msg, hint='', code=1):
    blank()
    alert('出错', [msg] + ([hint] if hint else []), 'red')
    blank()
    return code


# ──────────────────────────── 工具函数 ────────────────────────────
EXT_GROUPS = [
    ('图片',  'cyn', {'.png', '.jpg', '.jpeg', '.webp', '.gif', '.bmp',
                      '.tif', '.tiff', '.svg', '.heic', '.avif'}),
    ('PPT',   'mag', {'.pptx', '.ppt', '.potx', '.ppsx'}),
    ('文档',  'blu', {'.doc', '.docx', '.pdf', '.txt', '.md', '.rtf', '.wps'}),
    ('表格',  'blu', {'.xls', '.xlsx', '.csv', '.tsv'}),
    ('音视频', 'yel', {'.mp3', '.wav', '.m4a', '.flac', '.aac', '.mp4',
                       '.mov', '.mkv', '.avi', '.webm'}),
]
_TAG_MAP = {}


def _build_tag_map():
    for name, color, exts in EXT_GROUPS:
        for e in exts:
            _TAG_MAP[e] = (name, color)


_build_tag_map()


def tag_of(name):
    return _TAG_MAP.get(os.path.splitext(name)[1].lower(), ('其他', 'gry'))


def to_clipboard(text):
    try:
        p = subprocess.run('clip', input=text.encode('utf-16le'),
                           shell=True, capture_output=True)
        return p.returncode == 0
    except Exception:
        return False


def find_node():
    p = shutil.which('node')
    if p:
        return p
    for c in (r'C:\Program Files\nodejs\node.exe',
              r'C:\Program Files (x86)\nodejs\node.exe'):
        if os.path.isfile(c):
            return c
    return ''


def child_env():
    """引擎跑在 node 里，会 execFileSync('python', dims.py) —— 先把 python 塞进 PATH，
    并清掉代理（代理会拖慢/干扰本地子进程）。"""
    env = dict(os.environ)
    for k in ('HTTP_PROXY', 'HTTPS_PROXY', 'http_proxy', 'https_proxy'):
        env.pop(k, None)
    py_dir = os.path.dirname(sys.executable)
    if not shutil.which('python', path=env.get('PATH', '')):
        env['PATH'] = py_dir + os.pathsep + env.get('PATH', '')
    return env


# ──────────────────────────── spec 读取与校验 ────────────────────────────
_FENCE_HEAD = re.compile(r'^```[a-zA-Z]*\s*')
_FENCE_TAIL = re.compile(r'\s*```$')


def load_spec(path):
    """读 JSON：容忍 BOM / GBK / DeepSeek 的 ```json 围栏。"""
    raw = None
    with open(path, 'rb') as fp:
        data = fp.read()
    for enc in ('utf-8-sig', 'utf-8', 'gbk'):
        try:
            raw = data.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if raw is None:
        raw = data.decode('utf-8', errors='replace')
    text = raw.strip()
    if text.startswith('```'):
        text = _FENCE_TAIL.sub('', _FENCE_HEAD.sub('', text))
    try:
        spec = json.loads(text)
    except ValueError:
        # 网页版回复常带着寒暄和追问：把第一个 { 到最后一个 } 之间抠出来再试一次
        i, j = text.find('{'), text.rfind('}')
        if i < 0 or j <= i:
            raise
        spec = json.loads(text[i:j + 1])
    if not isinstance(spec, dict):
        raise ValueError('顶层必须是对象 { ... }')
    return spec


def engine_page_types():
    """从 deck-engine.js 现读 TYPE_FIELDS，保证校验口径和引擎一致。"""
    try:
        src = open(ENGINE, encoding='utf-8').read()
        m = re.search(r'TYPE_FIELDS\s*=\s*\{(.*?)\n\};', src, re.S)
        if m:
            out = {}
            for t, f in re.findall(r'(\w+)\s*:\s*\[([^\]]*)\]', m.group(1)):
                out[t] = re.findall(r'"([^"]+)"', f)
            if out:
                return out
    except Exception:
        pass
    return dict(FALLBACK_TYPES)


def page_images(pg):
    imgs = []
    if pg.get('image'):
        imgs += pg['image'] if isinstance(pg['image'], list) else [pg['image']]
    imgs += pg.get('thumbs') or []
    for x in (pg.get('images') or []):
        imgs.append(x if isinstance(x, str) else (x or {}).get('src'))
    ref = pg.get('ref')
    if isinstance(ref, dict):
        imgs.append(ref.get('src'))
    return [i for i in imgs if i]


def precheck(spec, assets):
    """返回 (errors, warns)；errors 里是 (页号或0, 类型, 描述)。"""
    types = engine_page_types()
    errors, warns = [], []
    pages = spec.get('pages')
    if not isinstance(pages, list) or not pages:
        errors.append((0, '-', 'pages 为空或缺失（至少要有一页）'))
        return errors, warns
    if not (spec.get('deck') or {}).get('title'):
        warns.append('deck.title 缺失（封面外的页面标题不受影响）')
    for i, pg in enumerate(pages, 1):
        if not isinstance(pg, dict):
            errors.append((i, '-', '这一页不是对象'))
            continue
        t = pg.get('type')
        if t not in types:
            errors.append((i, t or '(空)',
                           'type 非法，可选：' + ' / '.join(types.keys())))
            continue
        for f in types[t]:
            if pg.get(f) is None:
                errors.append((i, t, '缺字段 %s' % f))
        for r in (pg.get('rows') or []):
            if not (isinstance(r, list) and len(r) >= 2):
                errors.append((i, t, 'rows 每行必须是 [标签, 正文] 两元数组'))
                break
        if assets:
            for f in page_images(pg):
                if not os.path.isfile(os.path.join(assets, f)):
                    errors.append((i, t, '图片不存在: %s' % f))
    return errors, warns


def resolve_assets(spec, spec_path, cli_assets):
    """素材目录优先级：命令行 > spec.assets_dir > spec 同级目录 > 内置默认。"""
    cand = []
    if cli_assets:
        cand.append((cli_assets, '命令行'))
    if spec.get('assets_dir'):
        a = str(spec['assets_dir'])
        if not os.path.isabs(a):
            a = os.path.normpath(os.path.join(os.path.dirname(spec_path), a))
        cand.append((a, 'spec'))
    cand.append((os.path.dirname(spec_path), 'spec 同级'))
    cand.append((FALLBACK_ASSETS, '默认'))
    for path, src in cand:
        if os.path.isdir(path):
            return path, src
    return '', ''


# ──────────────────────────── 命令：scan ────────────────────────────
def walk_groups(folder):
    groups = {}
    for root, dirs, files in os.walk(folder):
        dirs[:] = sorted(d for d in dirs if not d.startswith('.'))
        rel = os.path.relpath(root, folder)
        if rel == '.':
            rel = '(根目录)'
        keep = sorted(f for f in files
                      if f != LIST_NAME and not f.startswith('~$'))
        if keep:
            groups[rel] = keep
    return [(k, groups[k]) for k in sorted(groups.keys())]


def cmd_scan(target, opt):
    folder = os.path.abspath(target)
    if not os.path.isdir(folder):
        return die('文件夹不存在：%s' % folder)

    groups = walk_groups(folder)
    total = sum(len(f) for _, f in groups)
    tally = {}
    img_total = 0
    for _, files in groups:
        for f in files:
            t, _c = tag_of(f)
            tally[t] = tally.get(t, 0) + 1
            if t == '图片':
                img_total += 1

    blank()
    hr('扫描素材')
    kv('目录', folder)
    kv('范围', '含全部子文件夹')
    hr()

    if not groups:
        note(cx('yel', '这个文件夹是空的（或只有隐藏文件）'))
    for rel, files in groups:
        kinds = {}
        for f in files:
            t, _c = tag_of(f)
            kinds[t] = kinds.get(t, 0) + 1
        stat = ' · '.join('%s %d' % (k, v) for k, v in
                          sorted(kinds.items(), key=lambda x: -x[1]))
        head = '  ' + cx('whi', rel)
        room = term_w() - dw(head) - dw(stat) - 3
        if room < 2:
            sys.stdout.write(head + '  ' + cx('gry', stat) + '\n')
        else:
            sys.stdout.write(head + ' ' + cx('gry', '.' * room) + ' ' + cx('gry', stat) + '\n')
        for f in files:
            t, c = tag_of(f)
            sys.stdout.write('    ' + cx(c, pad('[%s]' % t, 8)) + ' ' + f + '\n')
        blank()

    hr()
    kv('合计', '%d 个目录 · %d 个文件 · %s' % (
        len(groups), total, cx('cyn', '%d 张图片' % img_total)), vcolor='whi')
    order = [g[0] for g in EXT_GROUPS] + ['其他']
    kv('分类', ' · '.join('%s %d' % (k, tally.get(k, 0)) for k in order))

    # 清单文本（喂 DeepSeek 用，沿用旧格式，保持兼容）
    lines = ['素材根目录: %s' % folder, '',
             '共 %d 张图片，%d 个目录' % (img_total, len(groups)), '']
    for rel, files in groups:
        imgs = [f for f in files if tag_of(f)[0] == '图片']
        lines.append('== %s ==  （%d 个文件，其中 %d 张图片）' % (rel, len(files), len(imgs)))
        for f in imgs:
            lines.append('  [图片] %s' % f)
        for f in files:
            t = tag_of(f)[0]
            if t != '图片':
                lines.append('  [%s] %s' % (t, f))
        lines.append('')
    text = '\n'.join(lines)

    if opt['no_clip']:
        kv('剪贴板', '已跳过 (--no-clip)', vcolor='gry')
    elif to_clipboard(text):
        kv('剪贴板', cx('grn', '已复制') + '，可直接贴给 DeepSeek')
    else:
        kv('剪贴板', cx('yel', '复制失败') + '（清单仍已保存到文件）')

    if opt['no_save']:
        kv('清单', '已跳过 (--no-save)', vcolor='gry')
    else:
        out = os.path.join(folder, LIST_NAME)
        try:
            with open(out, 'w', encoding='utf-8') as fp:
                fp.write(text + '\n')
            kv('清单', out)
        except Exception as e:
            kv('清单', cx('red', '写入失败: %s' % e))
    blank()
    return 0


# ──────────────────────────── 命令：check ────────────────────────────
def cmd_check(target, opt):
    spec_path = os.path.abspath(target)
    if not os.path.isfile(spec_path):
        return die('spec 文件不存在：%s' % spec_path)
    try:
        spec = load_spec(spec_path)
    except Exception as e:
        return die('JSON 解析失败（原样发回 DeepSeek 修）', str(e))

    assets, src = resolve_assets(spec, spec_path, opt['assets'])
    pages = spec.get('pages') or []
    types = engine_page_types()

    blank()
    hr('校验 spec')
    kv('文件', spec_path)
    kv('规模', '%d 页 · %s' % (len(pages), human_size(os.path.getsize(spec_path))))
    if assets:
        kv('素材', assets + cx('gry', '  (%s)' % src))
    else:
        kv('素材', cx('yel', '未找到可用素材目录，已跳过图片检查'))
    hr()

    errors, warns = precheck(spec, assets if assets else '')
    bad = {}
    for idx, _t, msg in errors:
        bad.setdefault(idx, []).append(msg)

    for i, pg in enumerate(pages, 1):
        t = (pg or {}).get('type', '-')
        title = ''
        if isinstance(pg, dict):
            title = pg.get('title') or pg.get('chip') or ''
        ok = i not in bad
        mark = cx('grn', ' OK ') if ok else cx('red', 'FAIL')
        name = '第%-2d页  %s' % (i, pad(str(t), 10))
        sys.stdout.write('  [' + mark + '] ' + name)
        if ok:
            sys.stdout.write(cx('gry', trunc(title, max(term_w() - 24 - dw(name), 8))) + '\n')
        else:
            sys.stdout.write('\n')
            for m in bad[i]:
                sys.stdout.write('           ' + cx('red', '× ' + m) + '\n')

    for m in warns:
        sys.stdout.write('  [' + cx('yel', 'WARN') + '] ' + cx('yel', m) + '\n')

    blank()
    hr()
    if errors:
        kv('结果', cx('red', '%d 项未通过' % len(errors)), vcolor='whi')
        kv('下一步', '把上面的 × 行原样贴回 DeepSeek 让它改 JSON')
        blank()
        return 1
    kv('结果', cx('grn', '全部通过') + '（%d 页）' % len(pages), vcolor='whi')
    kv('下一步', 'deck.py render "%s"' % os.path.basename(spec_path))
    blank()
    return 0


# ──────────────────────────── 命令：render ────────────────────────────
def ea_patch(f):
    """中文字体后处理：给微软雅黑补 <a:ea> 东亚字体声明（原地，tmp 中转防锁）。"""
    latin = re.compile(r'(<a:latin typeface="微软雅黑"[^>]*/>)')
    zin = zipfile.ZipFile(f)
    tmp = f + '.tmp'
    zout = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename.endswith('.xml') and '微软雅黑'.encode() in data:
            data = latin.sub(r'\1<a:ea typeface="微软雅黑"/>',
                             data.decode('utf-8', errors='replace')).encode('utf-8')
        zout.writestr(item, data)
    zout.close()
    zin.close()
    try:
        os.replace(tmp, f)
    except PermissionError:
        time.sleep(1.5)
        try:
            os.replace(tmp, f)
        except PermissionError:
            with open(tmp, 'rb') as a, open(f, 'wb') as b:
                b.write(a.read())
            os.remove(tmp)


def cmd_render(target, opt):
    N = 6
    spec_path = os.path.abspath(target)
    blank()
    hr('渲染 PPT')
    blank()

    if not os.path.isfile(spec_path):
        step(1, N, '读取 spec', 'FAIL', '文件不存在')
        blank()
        return die('spec 文件不存在：%s' % spec_path)
    try:
        spec = load_spec(spec_path)
    except Exception as e:
        step(1, N, '读取 spec', 'FAIL', 'JSON 解析失败')
        blank()
        return die('JSON 解析失败（原样发回 DeepSeek 修）', str(e))
    pages = spec.get('pages') or []
    step(1, N, '读取 spec', 'OK',
         '%d 页 · %s' % (len(pages), human_size(os.path.getsize(spec_path))))

    assets, src = resolve_assets(spec, spec_path, opt['assets'])
    errors, warns = precheck(spec, assets if assets else '')
    if errors:
        step(2, N, '静态校验', 'FAIL', '%d 项未通过' % len(errors))
        blank()
        alert('校验未通过，退回修改（把这段原样贴给 DeepSeek）',
              ['第%d页 %s  %s' % (i, t, m) for i, t, m in errors])
        blank()
        return 1
    step(2, N, '静态校验', 'OK', '%d 页字段齐全' % len(pages))

    if not assets:
        step(3, N, '定位素材目录', 'FAIL', '没找到')
        blank()
        return die('找不到素材目录', '用 --assets "D:\\素材目录" 指定，或让 DeepSeek 在 JSON 里写好 assets_dir')
    n_img = len([f for f in os.listdir(assets)
                 if tag_of(f)[0] == '图片']) if os.path.isdir(assets) else 0
    step(3, N, '定位素材目录', 'OK',
         '%s [%s] %d 张图' % (os.path.basename(assets.rstrip('\\/')), src, n_img))

    if not os.path.isfile(ENGINE):
        blank()
        return die('渲染引擎不见了：%s' % ENGINE)
    node = find_node()
    if not node:
        blank()
        return die('找不到 node', '请安装 Node.js 或把它加入 PATH')

    stamp = time.strftime('%Y%m%d_%H%M%S')
    odir = os.path.abspath(opt['out'] or DESKTOP)
    if not os.path.isdir(odir):
        try:
            os.makedirs(odir, exist_ok=True)
        except Exception as e:
            blank()
            return die('输出目录不可用：%s' % odir, str(e))
    spec['assets_dir'] = assets
    spec['output'] = os.path.join(odir, '渲染_%s.pptx' % stamp)
    os.makedirs(SPECDIR, exist_ok=True)
    spec_file = os.path.join(SPECDIR, 'spec_%s.json' % stamp)
    with open(spec_file, 'w', encoding='utf-8') as fp:
        json.dump(spec, fp, ensure_ascii=False, indent=1)
    step(4, N, '写入渲染 spec', 'OK', os.path.basename(spec_file))

    t0 = time.time()
    try:
        r = subprocess.run([node, ENGINE, spec_file], capture_output=True,
                           encoding='utf-8', errors='replace',
                           timeout=180, env=child_env(),
                           cwd=os.path.dirname(ENGINE))
    except Exception as e:
        blank()
        return die('引擎调用失败：%s' % e)
    cost = time.time() - t0
    log = ((r.stdout or '') + (r.stderr or '')).strip()
    if r.returncode != 0 or not os.path.exists(spec['output']):
        step(5, N, '引擎渲染', 'FAIL', '%.1fs' % cost)
        blank()
        alert('引擎校验未通过（把这段原样贴给 DeepSeek）',
              (log or '(引擎没有输出)').split('\n'))
        blank()
        return 1
    step(5, N, '引擎渲染', 'OK', '%.1fs' % cost)

    try:
        ea_patch(spec['output'])
        step(6, N, '中文字体后处理', 'OK', 'a:ea 已注入')
    except Exception as e:
        step(6, N, '中文字体后处理', 'WARN', '跳过：%s' % e)

    blank()
    hr()
    out = spec['output']
    kv('成品', cx('grn', homeshort(out)), vcolor='whi')
    kv('大小', human_size(os.path.getsize(out)))
    kv('页数', str(len(pages)))
    kv('中间件', homeshort(spec_file))
    blank()
    return 0


# ──────────────────────────── 命令：guide ────────────────────────────
def cmd_guide():
    types = engine_page_types()
    optional = {
        'cover': 'kicker subtitle image strip tagline',
        'intro': 'badge pill image image_caption footer',
        'product': 'badge pill images[].label ref{src,caption} footer',
        'scene': '（整页大图，image + chip 必填）',
        'symbols': 'badge pill fig_caps[] footer',
        'colors': 'badge pill ratio[[名,色,占比]] ratio_note footer',
        'finale': 'badge pill strip footer',
        'toc': 'badge pill current footer（current=要高亮的章节名）',
        'section': 'badge subtitle image footer',
        'bullets': 'badge pill image image_caption footer',
        'imageText': 'badge pill image_side(left|right) image_caption footer',
        'cards': 'badge pill footer（cards[].icon 可用一个汉字当图标）',
        'compare': 'badge pill note footer（columns[].tone: up|down|neutral）',
        'timeline': 'badge pill footer（nodes 上下交替）',
        'process': 'badge pill footer',
        'matrix': 'badge pill x_label y_label note footer',
        'pyramid': 'badge pill base_note footer（levels 自下而上）',
        'stats': 'badge pill source footer',
        'chart': 'badge pill note footer（chart.type: bar|line|pie|doughnut|area）',
        'table': 'badge pill col_widths row_h note footer',
        'quote': 'image source footer',
        'team': 'badge pill footer（images[].label/.sub/.note）',
        'gallery': 'badge pill captions[] footer',
        'case': 'badge pill metrics[[值,标签]] image_caption footer',
        'contact': 'slogan lines[] image image_caption footer',
    }
    blank()
    hr('deck-spec 速查')
    blank()
    note(cx('whi', '顶层字段'), 'cyn')
    kv('assets_dir', '素材目录，图片按「文件名」引用', kw=12, indent=4)
    kv('output', '成品路径（本程序自动改成 桌面\\渲染_时间戳.pptx）', kw=12, indent=4)
    kv('deck', '{ title, subtitle, theme{ bg ink sub gold red hair muted font } }', kw=12, indent=4)
    kv('pages', '页面数组，见下', kw=12, indent=4)
    blank()
    note(cx('whi', '页面类型  %d 种' % len(types)), 'cyn')
    for t, req in types.items():
        sys.stdout.write('    ' + cx('mag', pad(t, 10)) +
                         cx('yel', '必填 ' + ','.join(req)) + '\n')
        sys.stdout.write('    ' + ' ' * 10 + cx('gry', '可选 ' +
                         optional.get(t, '-')) + '\n')
    blank()
    note(cx('whi', '常用片段'), 'cyn')
    sys.stdout.write(cx('gry', '''    rows   : [["文化来源","登高、佩茱萸…"], ["设计语言","…"]]
    strip  : [["菊金","#D9A441"], ["米白","#F3EDDF"]]
    images : [{"src":"x.png","label":"登高层"}, {"src":"y.png"}]
    theme  : {"bg":"22384E","ink":"F2EDE3","sub":"C9BFA9","gold":"C9A053",
              "red":"A63A2E","hair":"3A5566","muted":"9AA8B5","font":"微软雅黑"}
''') + '\n')
    hr()
    kv('预设说明', os.path.join(ROOT, '预设说明.md'))
    kv('引擎', ENGINE)
    kv('页面顺序', 'cover 起手 → toc → (section → 内容页若干)× → contact/finale 收尾')
    blank()
    return 0


# ──────────────────────────── 杂项 ────────────────────────────
def human_size(n):
    for u in ('B', 'KB', 'MB'):
        if n < 1024 or u == 'MB':
            return ('%d %s' % (n, u)) if u == 'B' else ('%.1f %s' % (n, u))
        n /= 1024.0
    return '?'


def homeshort(p):
    home = os.path.expanduser('~')
    if p.lower().startswith(home.lower()):
        return '~' + p[len(home):]
    return p


def print_help():
    w = term_w()
    blank()
    hr(APP + '  v' + VER)
    note('素材 -> DeepSeek -> pptx，一个程序跑完', 'gry')
    blank()
    hr('用法')
    items = [
        ('deck.py scan   <文件夹>', '扫描素材，出文件名清单（终端 + 剪贴板 + txt）'),
        ('deck.py check  <spec.json>', '静态校验 deck-spec，不产文件'),
        ('deck.py render <spec.json>', '渲染 pptx 到桌面'),
        ('deck.py guide', 'deck-spec 格式速查'),
        ('deck.py <文件夹|.json>', '拖拽用法：自动判断是 scan 还是 render'),
    ]
    for a, b in items:
        sys.stdout.write('  ' + cx('cyn', pad(a, 34)) + b + '\n')
    blank()
    hr('选项')
    opts = [
        ('-o, --out <目录>', '成品 pptx 输出目录（默认：桌面）'),
        ('    --assets <目录>', '强制指定素材目录（优先于 JSON 里的 assets_dir）'),
        ('    --no-clip', 'scan 时不写剪贴板'),
        ('    --no-save', 'scan 时不写 _文件名清单.txt'),
        ('    --no-color', '关闭彩色输出'),
        ('    --pause', '结束后暂停（拖拽运行时自动开启）'),
        ('-V, --version', '显示版本'),
        ('-h, --help', '显示本帮助'),
    ]
    for a, b in opts:
        sys.stdout.write('  ' + cx('gry', pad(a, 34)) + b + '\n')
    blank()
    hr('例子')
    note(r'deck.py scan "./demo-assets"', 'whi')
    note(r'deck.py render "spec.json" --assets "D:\素材"', 'whi')
    blank()
    note('把文件夹或 .json 直接拖到 PPT台.bat 图标上，也能跑。', 'gry')
    blank()


def parse_args(argv):
    opt = {'out': '', 'assets': '', 'no_clip': False, 'no_save': False,
           'no_color': False, 'pause': False}
    pos = []
    i = 0
    while i < len(argv):
        a = argv[i]
        if a in ('-h', '--help', 'help', '/?'):
            opt['help'] = True
        elif a in ('-V', '--version'):
            opt['version'] = True
        elif a in ('-o', '--out'):
            i += 1
            opt['out'] = argv[i] if i < len(argv) else ''
        elif a == '--assets':
            i += 1
            opt['assets'] = argv[i] if i < len(argv) else ''
        elif a == '--no-clip':
            opt['no_clip'] = True
        elif a == '--no-save':
            opt['no_save'] = True
        elif a == '--no-color':
            opt['no_color'] = True
        elif a == '--pause':
            opt['pause'] = True
        elif a.startswith('--') or (len(a) > 1 and a[0] == '-' and not os.path.exists(a)):
            sys.stderr.write('未知选项: %s\n' % a)
            opt['help'] = True
        else:
            pos.append(a)
        i += 1
    return opt, pos


def _run(opt, pos):
    global _COLOR_ON
    if opt.get('version'):
        sys.stdout.write('%s v%s\n' % (APP, VER))
        return 0
    if opt.get('help') or not pos:
        print_help()
        return 0 if opt.get('help') else 2

    # 容错：从帮助里复制命令时容易把程序名一起带上（deck.py scan xxx）
    while pos and os.path.basename(pos[0]).lower() in ('deck.py', 'deck', 'ppt台.bat'):
        pos = pos[1:]
    if not pos:
        print_help()
        return 0

    cmd = pos[0]
    rest = pos[1:]
    known = ('scan', 'check', 'render', 'guide')

    if cmd not in known:
        # 拖拽模式：第一个位置参数是路径 -> 按类型推断
        p = cmd
        if os.path.isdir(p):
            cmd, rest = 'scan', [p] + rest
        elif os.path.isfile(p) and p.lower().endswith('.json'):
            cmd, rest = 'render', [p] + rest
        elif os.path.isfile(p):
            return die('不认识的文件类型：%s' % os.path.basename(p),
                       '本程序只认 文件夹（scan）和 .json（render）')
        else:
            return die('找不到路径：%s' % p,
                       '用法见 deck.py -h')

    target = rest[0] if rest else ''

    if cmd == 'guide':
        code = cmd_guide()
    elif cmd == 'scan':
        if not target:
            if opt['pause']:
                code = die('没给文件夹', '把素材文件夹拖到本程序图标上，或写成 deck.py scan "D:\\素材"')
            else:
                target = os.getcwd()
                code = cmd_scan(target, opt)
        else:
            code = cmd_scan(target, opt)
    else:
        if not target:
            code = die('没给 spec 文件', '用法：deck.py %s "deck-spec.json"' % cmd)
        else:
            code = (cmd_check if cmd == 'check' else cmd_render)(target, opt)

    return code


def main(argv):
    """解析参数 -> 跑命令 -> 统一收尾（所有分支都会被 --pause 拦住，双击才不会闪退）。"""
    global _COLOR_ON
    opt, pos = parse_args(argv[1:])
    init_console()
    if opt.get('no_color'):
        _COLOR_ON = False

    try:
        code = _run(opt, pos)
    except KeyboardInterrupt:
        sys.stderr.write('\n已中断\n')
        code = 130

    if opt.get('pause') and not opt.get('version'):
        blank()
        try:
            input('按回车退出 ...')
        except (EOFError, KeyboardInterrupt):
            pass
    return code


if __name__ == '__main__':
    sys.exit(main(sys.argv))
