# -*- coding: utf-8 -*-
"""PPT 渲染台 · 专属图形界面

双击 = 开界面；把「素材文件夹」或「deck-spec .json」拖进窗口 / 拖到 exe 图标上 = 自动识别。
带 scan/check/render/guide 参数运行时仍可当命令行用，兼容原来的 PPT台.bat。

界面配色直接取 deck 主题本身：bg 22384E / ink F2EDE3 / gold C9A053 / red A63A2E。
"""
import os
import queue
import re
import shutil
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox

APP = 'PPT 渲染台'
VER = '1.0'

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

import deck  # noqa: E402


# ──────────────────────────── 打包/资源路径 ────────────────────────────
def res_dir():
    """只读资源（界面图标等）。打包后 = sys._MEIPASS。"""
    return getattr(sys, '_MEIPASS', None) or _HERE


def app_dir():
    """可写目录：exe 所在目录。绝不写 sys._MEIPASS（那是临时解压目录）。"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return _HERE


# 修掉「打包后 deck.ROOT 落在 _MEIPASS，中间件写进临时目录、退出即失」的老坑：
# 把中间件目录钉在 exe 旁边，单文件/文件夹两种形态都留得住。
deck.SPECDIR = os.path.join(app_dir(), '输出')
os.makedirs(deck.SPECDIR, exist_ok=True)


def resolve_engine():
    """按优先级找渲染引擎，找不到再用原程序的硬编码路径。"""
    cands = []
    if os.environ.get('PPT_ENGINE'):
        cands.append(os.environ['PPT_ENGINE'])
    cfg = os.path.join(app_dir(), '引擎路径.txt')
    if os.path.isfile(cfg):
        try:
            with open(cfg, encoding='utf-8-sig') as fp:
                p = fp.read().strip()
            if p:
                cands.append(p)
        except Exception:
            pass
    cands += [os.path.join(app_dir(), 'engine', 'deck-engine.js'),
              os.path.join(res_dir(), 'engine', 'deck-engine.js'),
              os.path.join(app_dir(), 'deck-engine.js'),
              deck.ENGINE]
    for c in cands:
        if c and os.path.isfile(c):
            return c
    return deck.ENGINE


# 模块级就定好：界面通道和命令行通道用的是同一个引擎
deck.ENGINE = resolve_engine()


# ──────────────────────────── 配色 / 字体 ────────────────────────────
C = {
    'bg': '#22384E', 'panel': '#1B2E42', 'card': '#26405A', 'log': '#16283A',
    'ink': '#F2EDE3', 'sub': '#C9BFA9', 'gold': '#C9A053', 'red': '#A63A2E',
    'hair': '#3A5566', 'muted': '#9AA8B5',
}
TAGS = {
    'def': C['ink'], 'whi': '#FFFFFF', 'gry': C['muted'], 'grn': '#8FD18F',
    'yel': '#E3C46B', 'cyn': '#7FD1D1', 'blu': '#7FB2E5', 'mag': '#D79BD7',
    'red': '#E06C5A',
}
UI = '微软雅黑'
MONO = 'Consolas'

_ANSI = re.compile(r'\x1b\[([0-9;]*)m')
_CODE2TAG = {'31': 'red', '32': 'grn', '33': 'yel', '34': 'blu',
             '35': 'mag', '36': 'cyn', '90': 'gry', '97': 'whi'}


class AnsiStream(object):
    """把 deck.py 的 ANSI 彩色输出转成 Tk 富文本片段，投进队列给主线程画。"""

    def __init__(self, q):
        self.q = q
        self._buf = ''
        self._raw = []

    def write(self, s):
        if not isinstance(s, str):
            s = str(s)
        self._buf += s
        m = re.search(r'\x1b(\[[0-9;]*)?$', self._buf)      # 转义序列可能被写一半
        hold = ''
        if m:
            hold = self._buf[m.start():]
            self._buf = self._buf[:m.start()]
        tag = 'def'
        for i, part in enumerate(_ANSI.split(self._buf)):
            if i % 2 == 0:
                if part:
                    self.q.put(('text', part, tag))
                    self._raw.append(part)
            else:
                code = part.split(';')[-1] if part else '0'
                if code in ('', '0'):
                    tag = 'def'
                elif code in _CODE2TAG:
                    tag = _CODE2TAG[code]
        self._buf = hold
        return len(s)

    def flush_hold(self):
        if self._buf:
            self.q.put(('text', self._buf, 'def'))
            self._raw.append(self._buf)
            self._buf = ''

    def text(self):
        return ''.join(self._raw)

    def flush(self):
        pass

    def isatty(self):
        return True

    def reconfigure(self, **kw):
        pass

    def fileno(self):
        raise OSError('AnsiStream has no fileno')


# ──────────────────────────── 界面 ────────────────────────────
class App(object):

    def __init__(self, preload=None):
        self.q = queue.Queue()
        self.busy = False
        self.last_pptx = ''
        self.last_list = ''
        self.cur = ('', '')

        deck.ENGINE = resolve_engine()

        self._dpi()
        self.root = self._make_root()          # 必须先有 root，再建 tk 变量
        self._load_icons()

        self.path_var = tk.StringVar()
        self.assets_var = tk.StringVar()
        self.out_var = tk.StringVar()
        self.mode_var = tk.StringVar(value='auto')

        self._build()
        self._probe()

        if preload:
            self.set_path(preload)
        self.root.after(60, self._drain)

    # ── 基础 ──
    @staticmethod
    def _dpi():
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

    def _make_root(self):
        self.dnd = False
        try:
            from tkinterdnd2 import TkinterDnD, DND_FILES
            self.dnd_files = DND_FILES
            root = TkinterDnD.Tk()
            self.dnd = True
        except Exception:
            root = tk.Tk()
        root.title('%s v%s' % (APP, VER))
        root.configure(bg=C['bg'])
        w, h = 920, 720
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        root.geometry('%dx%d+%d+%d' % (w, h, max((sw - w) // 2, 0), max((sh - h) // 3, 0)))
        root.minsize(840, 640)
        return root

    def _load_icons(self):
        self.icon32 = self.icon256 = None
        for name, attr in (('icon32.png', 'icon32'), ('icon256.png', 'icon256')):
            p = os.path.join(res_dir(), 'assets', name)
            if os.path.isfile(p):
                try:
                    img = tk.PhotoImage(file=p)
                    setattr(self, attr, img)
                except Exception:
                    pass
        if self.icon256 is not None:
            try:
                self.root.iconphoto(True, self.icon256)
            except Exception:
                pass

    def _btn(self, parent, text, cmd, width=None):
        b = tk.Button(parent, text=text, command=cmd, bg=C['panel'], fg=C['ink'],
                      activebackground=C['card'], activeforeground=C['gold'],
                      relief='flat', bd=0, font=(UI, 9), cursor='hand2',
                      padx=12, pady=4)
        if width:
            b.configure(width=width)
        return b

    def _entry(self, parent, var, width=None):
        e = tk.Entry(parent, textvariable=var, bg=C['panel'], fg=C['ink'],
                     insertbackground=C['ink'], relief='flat', font=(UI, 9))
        if width:
            e.configure(width=width)
        return e

    def _bind_drop(self, widget):
        if not self.dnd:
            return
        try:
            widget.drop_target_register(self.dnd_files)
            widget.dnd_bind('<<Drop>>', self._on_drop)
        except Exception:
            pass

    # ── 布局 ──
    def _build(self):
        main = tk.Frame(self.root, bg=C['bg'])
        main.pack(fill='both', expand=True, padx=14, pady=12)

        # 头部
        hd = tk.Frame(main, bg=C['panel'], highlightbackground=C['hair'], highlightthickness=1)
        hd.pack(fill='x')
        inner = tk.Frame(hd, bg=C['panel'])
        inner.pack(fill='x', padx=14, pady=10)
        if self.icon32 is not None:
            tk.Label(inner, image=self.icon32, bg=C['panel']).pack(side='left', padx=(0, 12))
        tbox = tk.Frame(inner, bg=C['panel'])
        tbox.pack(side='left')
        tk.Label(tbox, text=APP, bg=C['panel'], fg=C['gold'],
                 font=(UI, 17, 'bold')).pack(anchor='w')
        tk.Label(tbox, text='素材清单 · deck-spec 校验 · 一键出 pptx', bg=C['panel'],
                 fg=C['muted'], font=(UI, 9)).pack(anchor='w')
        chips = tk.Frame(inner, bg=C['panel'])
        chips.pack(side='right')
        self.chip = {}
        for key, label in (('dnd', '拖拽'), ('python', 'python'), ('node', 'node'), ('engine', '引擎')):
            lb = tk.Label(chips, text='\u25cf ' + label, bg=C['panel'], fg=C['muted'], font=(UI, 9))
            lb.pack(side='right', padx=5)
            self.chip[key] = lb

        # 拖拽区
        card = tk.Frame(main, bg=C['card'], highlightbackground=C['gold'], highlightthickness=1)
        card.pack(fill='x', pady=(12, 0))
        tk.Label(card, text='\u5c06「素材文件夹」或「deck-spec .json」拖到这里', bg=C['card'],
                 fg=C['ink'], font=(UI, 12, 'bold')).pack(pady=(14, 2))
        tk.Label(card, text='也可以点下面的按钮选；拖到 exe 图标上会带着路径直接开这个界面',
                 bg=C['card'], fg=C['muted'], font=(UI, 9)).pack()
        row = tk.Frame(card, bg=C['card'])
        row.pack(fill='x', padx=14, pady=(10, 14))
        self._btn(row, '选择文件夹…', self.pick_folder).pack(side='left')
        self._btn(row, '选择 JSON…', self.pick_json).pack(side='left', padx=(8, 12))
        self.path_entry = self._entry(row, self.path_var)
        self.path_entry.pack(side='left', fill='x', expand=True, ipady=4)
        self.path_entry.bind('<Return>', lambda e: self._start())

        # 选项
        opt = tk.Frame(main, bg=C['bg'])
        opt.pack(fill='x', pady=(10, 0))

        mrow = tk.Frame(opt, bg=C['bg'])
        mrow.pack(fill='x')
        tk.Label(mrow, text='动作', bg=C['bg'], fg=C['sub'], font=(UI, 9)).pack(side='left', padx=(0, 10))
        for val, text in (('auto', '自动识别'), ('scan', '素材清单'),
                          ('check', '校验 spec'), ('render', '渲染 pptx')):
            tk.Radiobutton(mrow, text=text, value=val, variable=self.mode_var,
                           bg=C['bg'], fg=C['ink'], selectcolor=C['panel'],
                           activebackground=C['bg'], activeforeground=C['gold'],
                           font=(UI, 9), bd=0, highlightthickness=0).pack(side='left', padx=(0, 12))

        for label, var, is_dir, hint in (
                ('素材目录', self.assets_var, True, '留空 = 用 spec 里写的'),
                ('输出目录', self.out_var, True, '留空 = 桌面')):
            r = tk.Frame(opt, bg=C['bg'])
            r.pack(fill='x', pady=(6, 0))
            tk.Label(r, text=label, bg=C['bg'], fg=C['sub'], font=(UI, 9),
                     width=8, anchor='w').pack(side='left')
            self._btn(r, '浏览…', lambda v=var, d=is_dir: self.pick_dir_into(v, d)).pack(side='left', padx=(0, 8))
            self._entry(r, var).pack(side='left', fill='x', expand=True, ipady=3)
            tk.Label(r, text=hint, bg=C['bg'], fg=C['muted'], font=(UI, 8)).pack(side='left', padx=(8, 0))

        # 开始
        self.run_btn = tk.Button(main, text='\u25b6   开始', command=self._start, bg=C['gold'],
                                 fg='#1B2E42', activebackground='#E0B96A', activeforeground='#1B2E42',
                                 font=(UI, 12, 'bold'), relief='flat', bd=0, cursor='hand2')
        self.run_btn.pack(fill='x', pady=(12, 8), ipady=8)

        # 底栏：先 pack 到最底部，否则会被 expand 的日志整块挤掉（连状态文字都看不见）
        ft = tk.Frame(main, bg=C['bg'])
        ft.pack(side='bottom', fill='x', pady=(8, 0))
        self._btn(ft, '打开输出', lambda: self.open_path(deck.SPECDIR)).pack(side='left')
        self._btn(ft, '打开成品', self.open_pptx).pack(side='left', padx=4)
        self._btn(ft, '打开清单', self.open_list).pack(side='left')
        self._btn(ft, '复制提示词', self.copy_prompt).pack(side='left', padx=4)
        self._btn(ft, '预设说明', self.open_presets).pack(side='left')
        self._btn(ft, '清空日志', self.clear_log).pack(side='left', padx=4)
        self.status = tk.Label(ft, text='就绪', bg=C['bg'], fg=C['muted'], font=(UI, 9))
        self.status.pack(side='right')

        # 日志
        lf = tk.Frame(main, bg=C['panel'], highlightbackground=C['hair'], highlightthickness=1)
        lf.pack(fill='both', expand=True)
        self.log = tk.Text(lf, bg=C['log'], fg=C['ink'], insertbackground=C['ink'],
                           font=(MONO, 9), relief='flat', bd=0, wrap='none', padx=10, pady=8,
                           height=9)
        ys = tk.Scrollbar(lf, command=self.log.yview, bd=0, bg=C['panel'], troughcolor=C['bg'])
        xs = tk.Scrollbar(lf, command=self.log.xview, orient='horizontal', bd=0,
                          bg=C['panel'], troughcolor=C['bg'])
        self.log.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        xs.pack(side='bottom', fill='x')
        ys.pack(side='right', fill='y')
        self.log.pack(side='left', fill='both', expand=True)
        for name, color in TAGS.items():
            self.log.tag_configure(name, foreground=color)
        self.log.configure(state='disabled')

        for w in (card, self.log, self.path_entry):
            self._bind_drop(w)


    # ── 环境探测 ──
    def _probe(self):
        node = deck.find_node()
        py = shutil.which('python') or ''
        eng_ok = os.path.isfile(deck.ENGINE)
        state = {'engine': eng_ok, 'node': bool(node), 'python': bool(py), 'dnd': self.dnd}
        for k, ok in state.items():
            self.chip[k].configure(fg='#8FD18F' if ok else '#E06C5A')

        self._append('\n', 'def')
        self._append('  环境检查\n', 'cyn')
        self._append('    引擎    %s\n' % deck.ENGINE, 'def' if eng_ok else 'red')
        self._append('    node    %s\n' % (node or '(未找到，装 Node.js 后才能渲染)'), 'def' if node else 'yel')
        self._append('    python  %s\n' % (py or '(未找到；图片会按方形兜底，仍能出稿)'), 'def' if py else 'yel')
        self._append('    拖拽    %s\n' % ('已启用（tkdnd）' if self.dnd else '不可用，只能用按钮选路径'),
                     'def' if self.dnd else 'yel')
        self._append('    中间件  %s\n' % deck.SPECDIR, 'def')
        self._append('\n  把素材文件夹或 deck-spec .json 拖进来，或点上面按钮选，然后按「开始」。\n', 'gry')

    # ── 路径 ──
    def set_path(self, p):
        p = (p or '').strip().strip('"')
        self.path_var.set(p)
        if os.path.isdir(p):
            self.mode_var.set('auto')
            self.status.configure(text='已选素材文件夹 · 将做「素材清单」')
        elif p.lower().endswith('.json'):
            self.mode_var.set('auto')
            self.status.configure(text='已选 deck-spec · 将做「渲染 pptx」')
        else:
            self.status.configure(text='这个路径既不是文件夹也不是 .json')

    def _on_drop(self, event):
        try:
            items = self.root.tk.splitlist(event.data)
        except Exception:
            items = [event.data]
        if items:
            self.set_path(items[0])
        return event.action if hasattr(event, 'action') else None

    def pick_folder(self):
        p = filedialog.askdirectory(title='选择素材文件夹', parent=self.root)
        if p:
            self.set_path(p)

    def pick_json(self):
        p = filedialog.askopenfilename(title='选择 deck-spec', parent=self.root,
                                       filetypes=[('deck-spec', '*.json'), ('全部文件', '*.*')])
        if p:
            self.set_path(p)

    def pick_dir_into(self, var, is_dir):
        p = filedialog.askdirectory(title='选择目录', parent=self.root, initialdir=var.get() or None)
        if p:
            var.set(os.path.normpath(p))

    # ── 干活 ──
    def _start(self):
        if self.busy:
            return
        p = self.path_var.get().strip().strip('"')
        if not p:
            messagebox.showwarning(APP, '先选一个素材文件夹或 deck-spec .json（也可以直接拖进来）',
                                   parent=self.root)
            return
        mode = self.mode_var.get()
        if mode == 'auto':
            if os.path.isdir(p):
                cmd = 'scan'
            elif os.path.isfile(p) and p.lower().endswith('.json'):
                cmd = 'render'
            else:
                messagebox.showwarning(APP, '认不出这个路径：只认 文件夹（素材）和 .json（deck-spec）',
                                       parent=self.root)
                return
        else:
            cmd = mode
        if cmd == 'scan' and not os.path.isdir(p):
            messagebox.showwarning(APP, '「素材清单」要给一个文件夹', parent=self.root)
            return
        if cmd in ('check', 'render') and not os.path.isfile(p):
            messagebox.showwarning(APP, '「%s」要给一个 .json 文件' % cmd, parent=self.root)
            return

        self.busy = True
        # tk 变量不是线程安全的：必须在主线程先读成纯字符串，再交给工作线程
        out = self.out_var.get().strip().strip('"')
        assets = self.assets_var.get().strip().strip('"')
        self.cur = (cmd, p, out, assets)
        self.last_pptx = ''
        self.last_list = ''
        self.run_btn.configure(state='disabled', text='\u25cf 正在跑…')
        self.status.configure(text='运行中…')
        self._append('\n' + '\u2500' * 92 + '\n', 'gry')
        threading.Thread(target=self._work, args=(cmd, p, out, assets), daemon=True).start()

    def _work(self, cmd, target, out, assets):
        stream = AnsiStream(self.q)
        old_out, old_err = sys.stdout, sys.stderr
        sys.stdout = sys.stderr = stream
        code = 1
        try:
            deck.init_console()
            opt = {'out': out, 'assets': assets,
                   'no_clip': False, 'no_save': False, 'no_color': False, 'pause': False,
                   'version': False}
            code = deck._run(opt, [cmd, target])
        except SystemExit as e:
            code = e.code if isinstance(e.code, int) else 0
        except Exception:
            stream.write('\n\x1b[31m内部错误：\x1b[0m\n' + traceback.format_exc())
            code = 1
        finally:
            sys.stdout, sys.stderr = old_out, old_err
            stream.flush_hold()
            self.q.put(('done', 0 if code is None else code, target))

    def _finished(self, code, target):
        self.busy = False
        self.run_btn.configure(state='normal', text='\u25b6   开始')
        cmd = self.cur[0]

        if cmd == 'render':
            outdir = self.cur[2] or deck.DESKTOP
            best, mt = '', -1
            try:
                for f in os.listdir(outdir):
                    if f.lower().endswith('.pptx'):
                        t = os.path.getmtime(os.path.join(outdir, f))
                        if t > mt:
                            best, mt = os.path.join(outdir, f), t
            except Exception:
                pass
            self.last_pptx = best
        elif cmd == 'scan':
            cand = os.path.join(target, deck.LIST_NAME)
            if os.path.isfile(cand):
                self.last_list = cand

        if code == 0:
            extra = ''
            if self.last_pptx:
                extra = '  ·  成品 %s' % os.path.basename(self.last_pptx)
            elif self.last_list:
                extra = '  ·  清单 %s' % os.path.basename(self.last_list)
            self.status.configure(text='完成' + extra)
        else:
            self.status.configure(text='没跑通（退出码 %s），把上面那段原样发给 DeepSeek 修' % code)

    # ── 日志 ──
    def _append(self, s, tag='def'):
        self.log.configure(state='normal')
        self.log.insert('end', s, tag)
        self.log.see('end')
        self.log.configure(state='disabled')

    def _drain(self):
        n = 0
        try:
            while n < 500:
                item = self.q.get_nowait()
                n += 1
                if item[0] == 'text':
                    self._append(item[1], item[2])
                elif item[0] == 'done':
                    self._append('\n', 'def')
                    self._finished(item[1], item[2])
        except queue.Empty:
            pass
        self.root.after(50, self._drain)

    def clear_log(self):
        self.log.configure(state='normal')
        self.log.delete('1.0', 'end')
        self.log.configure(state='disabled')

    # ── 打开 ──
    @staticmethod
    def open_path(p):
        try:
            if p and os.path.exists(p):
                os.startfile(p)          # noqa: S606  (Windows 专用)
                return True
        except Exception:
            pass
        return False

    def _folder_listing(self, folder):
        """按 scan 的口径列出文件夹里的图片文件名（给网页版 DeepSeek 看）。"""
        groups = deck.walk_groups(folder)
        out = []
        for rel, files in groups:
            imgs = [f for f in files if deck.tag_of(f)[0] == '图片']
            if imgs:
                out.append('== %s ==' % rel)
                for f in imgs:
                    out.append('  ' + f)
        return '\n'.join(out) if out else '(这个文件夹里没有图片)'

    def copy_prompt(self):
        """一键复制「给 DeepSeek 的提示词」＝ 输出格式要求 + 25 种版式规则 + 素材清单。"""
        doc = ''
        for cand in (os.path.join(app_dir(), '给DeepSeek的提示词.md'),
                     os.path.join(res_dir(), '给DeepSeek的提示词.md'),
                     os.path.join(app_dir(), '预设说明.md'),
                     os.path.join(res_dir(), '预设说明.md')):
            if os.path.isfile(cand):
                try:
                    with open(cand, encoding='utf-8') as fp:
                        doc = fp.read()
                except Exception:
                    pass
                break
        if not doc:
            messagebox.showinfo(APP, '没找到 给DeepSeek的提示词.md（应该和 exe 放在一起）', parent=self.root)
            return
        p = self.path_var.get().strip().strip('"')
        tail = []
        if p and os.path.isdir(p):
            try:
                tail = ['', '【素材目录】：', p.replace('\\', '/'), '', '【文件名】：', self._folder_listing(p)]
            except Exception as e:
                tail = ['', '（素材清单没生成出来：%s）' % e]
        text = doc + '\n'.join(tail)
        ok = deck.to_clipboard(text)
        self._append('\n', 'def')
        self._append('  提示词已复制（%d 字）%s\n'
                     % (len(text), '，粘到网页版 DeepSeek 就行' if ok else '，但写剪贴板失败'),
                     'grn' if ok else 'yel')

    def open_presets(self):
        for cand in (os.path.join(app_dir(), '预设说明.md'),
                     os.path.join(res_dir(), '预设说明.md')):
            if self.open_path(cand):
                return
        messagebox.showinfo(APP, '没找到 预设说明.md（应该和 exe 放在一起）', parent=self.root)

    def open_pptx(self):
        if not self.open_path(self.last_pptx):
            messagebox.showinfo(APP, '还没有成品。先渲染一次。', parent=self.root)

    def open_list(self):
        if not self.open_path(self.last_list):
            messagebox.showinfo(APP, '还没有清单。先做一次「素材清单」。', parent=self.root)

    def run(self):
        self.root.mainloop()


# ──────────────────────────── 自检 / 入口 ────────────────────────────
def selftest(out_path):
    """给打包后的 exe 用：把界面真正建起来 + 探一遍依赖，结果写成 JSON。"""
    import json
    rep = {'frozen': bool(getattr(sys, 'frozen', False)),
           'exe': sys.executable,
           'meipass': getattr(sys, '_MEIPASS', ''),
           'app_dir': app_dir(),
           'deck_module': getattr(deck, '__file__', ''),
           'deck_SPECDIR': deck.SPECDIR,
           'engine': deck.ENGINE,
           'engine_exists': os.path.isfile(deck.ENGINE),
           'node': deck.find_node(),
           'python': shutil.which('python') or '',
           'icon32': os.path.isfile(os.path.join(res_dir(), 'assets', 'icon32.png')),
           'presets_doc': os.path.isfile(os.path.join(app_dir(), '预设说明.md')),
           'type_count': 0, 'type_names': ''}
    root = None
    try:
        root = App()
        rep['tk'] = root.root.tk.call('info', 'patchlevel')
        rep['dnd'] = bool(root.dnd)
        try:
            pkgs = root.root.tk.call('package', 'names')
            rep['tkdnd_loaded'] = 'tkdnd' in str(pkgs)
        except Exception as e:
            rep['tkdnd_loaded'] = 'query failed: %r' % e
        rep['title'] = root.root.title()
        rep['geometry'] = root.root.winfo_geometry()
        rep['widgets'] = len(root.root.winfo_children())
        try:
            names = list(deck.engine_page_types().keys())
            rep['type_count'] = len(names)
            rep['type_names'] = ' '.join(names)
        except Exception as e:
            rep['type_count'] = 'ERR %r' % e
        root.root.update()
        rep['ok'] = True
    except Exception:
        rep['ok'] = False
        rep['error'] = traceback.format_exc()
    finally:
        try:
            if root is not None:
                root.root.destroy()
        except Exception:
            pass
    with open(out_path, 'w', encoding='utf-8') as fp:
        json.dump(rep, fp, ensure_ascii=False, indent=2)
    return 0 if rep.get('ok') else 1


def main():
    args = sys.argv[1:]
    if args and args[0] == '--selftest':
        return selftest(args[1] if len(args) > 1 else os.path.join(app_dir(), 'selftest.json'))
    if args and (args[0] in ('scan', 'check', 'render', 'guide', '-h', '--help', '-V', '--version')
                 or args[0] in ('--out', '-o', '--assets')):
        for name in ('stdout', 'stderr'):
            if getattr(sys, name, None) is None:        # 窗口版 exe 在无控制台时可能是 None
                try:
                    setattr(sys, name, open(os.devnull, 'w', encoding='utf-8'))
                except Exception:
                    pass
        return deck.main([sys.argv[0]] + args)
    App(args[0] if args else None).run()
    return 0


if __name__ == '__main__':
    sys.exit(main() or 0)
