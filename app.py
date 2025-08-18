import os, re, json, time, threading, queue
from urllib.parse import urlparse, urljoin
import tkinter as tk
import ttkbootstrap as tb
from ttkbootstrap import ttk
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledText
from tkinter import filedialog, messagebox


import requests
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# NEW: 用于 4:3 转换
from PIL import Image
# ---- DPI awareness & scaling helpers (Windows + 通用) ----
# --- 放在 imports 后面 ---
import sys, os, tarfile, pathlib, platform

APP_NAME = "WebImageSaver"

def _user_data_dir():
    home = pathlib.Path.home()
    if platform.system() == "Windows":
        base = pathlib.Path(os.environ.get("LOCALAPPDATA", home / "AppData/Local"))
        return base / APP_NAME
    elif platform.system() == "Darwin":
        return home / "Library/Application Support" / APP_NAME
    else:
        return home / ".local/share" / APP_NAME

APP_ROOT   = pathlib.Path(getattr(sys, "_MEIPASS", os.path.dirname(sys.argv[0])))
RUNTIME_DIR = _user_data_dir()
MS_DIR      = RUNTIME_DIR / "ms-playwright"
MS_TGZ_APP  = APP_ROOT / "ms-playwright.tgz"   # 跟 exe 同目录打包进去
MS_TGZ_USER = RUNTIME_DIR / "ms-playwright.tgz"

# 让 Playwright 永远使用“用户目录”的浏览器（可写）
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(MS_DIR)

def ensure_local_browsers():
    RUNTIME_DIR.mkdir(parents=True, exist_ok=True)
    if MS_DIR.exists() and any(MS_DIR.iterdir()):
        return

    if MS_TGZ_APP.exists():
        try:
            with tarfile.open(MS_TGZ_APP, "r:gz") as tf:
                tf.extractall(RUNTIME_DIR)
            # 可选：留一份副本
            try:
                if not MS_TGZ_USER.exists():
                    MS_TGZ_USER.write_bytes(MS_TGZ_APP.read_bytes())
            except Exception:
                pass
        except Exception as e:
            raise SystemExit(f"[ERROR] 解压浏览器失败：{e}")

        if MS_DIR.exists() and any(MS_DIR.iterdir()):
            return
        raise SystemExit("[ERROR] 浏览器解压后内容缺失。")

    raise SystemExit("[ERROR] 缺少浏览器内核（ms-playwright.tgz）。")

ensure_local_browsers()



def _apply_win_dpi_awareness():
    if platform.system() == "Windows":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # system DPI aware
        except Exception:
            pass

def set_ui_scaling(root, factor: float):
    """Tk 的缩放：1.0=100%，1.25=125%，1.5=150%"""
    try:
        root.call('tk', 'scaling', factor)
    except Exception:
        pass

def set_scaling_from_system(root, user_factor=1.0):
    # Windows: 用系统 DPI 自动换算 tk scaling；其它平台用默认
    try:
        import ctypes
        dpi = ctypes.windll.user32.GetDpiForSystem()  # 96=100%
        tk_scale = (dpi / 72.0) * user_factor         # tk 以 72dpi 为 1.0 基准
        root.call('tk', 'scaling', tk_scale)
    except Exception:
        # 失败就退回到手动倍率
        set_ui_scaling(root, 1.6 * user_factor)


def convert_to_4_3(input_path, output_path, background_color=(0, 0, 0, 0)):
    """
    将图片转为 4:3 的 PNG（居中填充，不裁切），默认透明背景。
    """
    try:
        with Image.open(input_path) as original_img:
            # 统一到 RGBA，保证可以有透明背景
            if original_img.mode != "RGBA":
                original_img = original_img.convert("RGBA")

            ow, oh = original_img.size
            target_ratio = 4 / 3
            original_ratio = ow / oh

            if abs(original_ratio - target_ratio) < 1e-3:
                # 已经接近 4:3，直接输出 PNG
                original_img.save(output_path, "PNG")
                return

            if original_ratio > target_ratio:
                # 图片更“宽”，补高度
                new_h = int(round(ow / target_ratio))
                new_w = ow
            else:
                # 图片更“高”，补宽度
                new_w = int(round(oh * target_ratio))
                new_h = oh

            new_img = Image.new("RGBA", (new_w, new_h), background_color)
            paste_x = (new_w - ow) // 2
            paste_y = (new_h - oh) // 2
            new_img.paste(original_img, (paste_x, paste_y))

            new_img.save(output_path, "PNG")
    except Exception as e:
        # 不阻塞主流程，只记录
        print(f"[WARN] 4:3 转换失败: {input_path} -> {e}")


ILLEGAL = r'[\\/:*?"<>|]'
CLICK_MORE_TEXTS = [
    "加载更多","更多","下一页","更多内容","查看更多","展开",
    "Load more","More","Next","Show more","See more","View more","Continue"
]

def sanitize(text: str, max_len=30):
    if not text: return ""
    text = re.sub(r'\s+', ' ', text).strip()
    text = re.sub(ILLEGAL, ' ', text)
    return text[:max_len]

def choose_name(heading, caption, alt, fallback):
    for s in (heading, caption, alt, fallback):
        s = sanitize(s)
        if s: return s
    return "image"

def ext_from_url(u: str):
    path = urlparse(u).path.lower()
    for ext in (".png",".jpg",".jpeg",".webp",".gif",".bmp",".svg"):
        if path.endswith(ext): return ".png" if ext==".svg" else ext
    return ".jpg"

def fetch_bytes(url, headers=None):
    r = requests.get(url, headers=headers or {}, timeout=20)
    r.raise_for_status()
    return r.content

JS_COLLECT = r"""
() => {
  function textOf(el){ if(!el) return ""; const t = el.innerText||el.textContent||""; return t.trim().replace(/\s+/g,' '); }
  function nearestHeading(el){
    let n = el;
    while(n){
      let cur = n.previousElementSibling;
      while(cur){
        if(['H2','H3','H1','H4'].includes(cur.tagName) && textOf(cur)) return textOf(cur);
        cur = cur.previousElementSibling;
      }
      n = n.parentElement;
    }
    return "";
  }
  function captionAround(el){
    let p = el.closest('figure');
    if(p){ const cap = p.querySelector('figcaption'); if(cap) return textOf(cap); }
    const aria = el.getAttribute('aria-label')||''; if(aria.trim()) return aria.trim();
    const descId = el.getAttribute('aria-describedby');
    if(descId){ const d = document.getElementById(descId); if(d) return textOf(d); }
    const near = el.closest('[class], [role], section, article, div');
    if(near){ const cand = textOf(near); if(cand && cand.length<=80) return cand; }
    return "";
  }
  function absUrl(u){ try { return new URL(u, location.href).href; } catch(e){ return ""; } }
  function cssPath(el){
    if (!(el instanceof Element)) return "";
    const parts = [];
    while (el && el.nodeType === Node.ELEMENT_NODE && parts.length < 8) {
      let sel = el.nodeName.toLowerCase();
      if (el.id) { sel += "#" + el.id; parts.unshift(sel); break; }
      else {
        let i = 1, sib = el;
        while ((sib = sib.previousElementSibling) != null) if (sib.nodeName === el.nodeName) i++;
        sel += `:nth-of-type(${i})`;
      }
      parts.unshift(sel);
      el = el.parentElement;
    }
    return parts.join(" > ");
  }

  const out = [];

  // <img>
  document.querySelectorAll('img').forEach(img=>{
    const src = img.currentSrc || img.src || "";
    if(!src) return;
    out.push({
      kind:"img",
      url:absUrl(src),
      alt:img.alt||"",
      caption:captionAround(img),
      nearestHeading:nearestHeading(img),
      css: cssPath(img)
    });
  });

  // CSS 背景图
  document.querySelectorAll('*').forEach(el=>{
    const bg = getComputedStyle(el).backgroundImage;
    if(bg && bg.includes('url(')){
      const m = bg.match(/url\((['"]?)(.*?)\1\)/);
      if(m && m[2]){
        out.push({
          kind:"bg",
          url:absUrl(m[2]),
          alt:"",
          caption:captionAround(el),
          nearestHeading:nearestHeading(el),
          css: cssPath(el)
        });
      }
    }
  });

  // 按 url 去重
  const seen = new Set(); const dedup=[];
  for(const it of out){
    if(!it.url) continue;
    const key = it.url.split('#')[0];
    if(seen.has(key)) continue;
    seen.add(key); dedup.push(it);
  }
  return dedup;
}
"""

class App:
    def __init__(self, root):
        self.root = root
        root.title("Web Image Saver (GUI)")
        root.geometry("800x560")
        self.q = queue.Queue()
        self.worker = None
        self.stop_flag = threading.Event()

        # ===== 主题与基础样式 =====
        style = ttk.Style(self.root)
        style.configure("TFrame", padding=12)
        style.configure("TLabelframe", padding=12, labeloutside=False)
        style.configure("TButton", padding=(14, 8))
        style.configure("TLabel", padding=(0, 2))

        # ===== 顶部表单（两列自适应）=====
        frm = ttk.Frame(root)
        frm.pack(fill="x")
        ttk.Label(frm, text="目标网址：").grid(row=0, column=0, sticky=E)
        self.url_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.url_var).grid(row=0, column=1, sticky=EW, padx=8)

        ttk.Label(frm, text="输出目录：").grid(row=1, column=0, sticky=E, pady=(6, 0))
        self.dir_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.dir_var).grid(row=1, column=1, sticky=EW, padx=8, pady=(6, 0))
        ttk.Button(frm, text="选择…", bootstyle="secondary", command=self.pick_dir) \
            .grid(row=1, column=2, sticky=W, padx=6, pady=(6, 0))
        frm.columnconfigure(1, weight=1)

        # ===== 选项分组 =====
        opt = ttk.Labelframe(root, text="选项")
        opt.pack(fill="x", pady=(10, 6))

        self.headless = tk.BooleanVar(value=False)
        self.try_more = tk.BooleanVar(value=True)
        self.mode = tk.StringVar(value="download")  # download / screenshot / both
        self.max_scrolls = tk.IntVar(value=30)
        self.pad_43 = tk.BooleanVar(value=False)

        ttk.Checkbutton(opt, text="可视浏览器（推荐）",
                        variable=self.headless, onvalue=False, offvalue=True,
                        bootstyle="round-toggle") \
            .grid(row=0, column=0, sticky=W)
        ttk.Checkbutton(opt, text="尝试点击“加载更多/下一页”",
                        variable=self.try_more, bootstyle="round-toggle") \
            .grid(row=0, column=1, sticky=W, padx=18)

        ttk.Label(opt, text="最大滚动次数：").grid(row=0, column=2, sticky=E)
        ttk.Spinbox(opt, from_=5, to=200, textvariable=self.max_scrolls, width=6) \
            .grid(row=0, column=3, sticky=W, padx=(6, 0))

        ttk.Label(opt, text="保存方式：").grid(row=1, column=0, sticky=E, pady=(8, 0))
        ttk.Radiobutton(opt, text="下载原图", variable=self.mode, value="download", bootstyle="toolbutton") \
            .grid(row=1, column=1, sticky=W, pady=(8, 0))
        ttk.Radiobutton(opt, text="截图元素", variable=self.mode, value="screenshot", bootstyle="toolbutton") \
            .grid(row=1, column=2, sticky=W, pady=(8, 0))
        ttk.Radiobutton(opt, text="同时保存（二者都要）", variable=self.mode, value="both", bootstyle="toolbutton") \
            .grid(row=1, column=3, sticky=W, pady=(8, 0))

        ttk.Checkbutton(opt, text="将全部图片转为 4:3 的 PNG（补边居中）",
                        variable=self.pad_43, bootstyle="square-toggle") \
            .grid(row=2, column=0, columnspan=4, sticky=W, pady=(10, 0))

        for c in range(4):
            opt.columnconfigure(c, weight=1)

        # ===== 操作区（右侧对齐）=====
        bar = ttk.Frame(root)
        bar.pack(fill="x", pady=(6, 4))
        ttk.Separator(root, orient=HORIZONTAL).pack(fill="x")

        right = ttk.Frame(bar)
        right.pack(side="right")
        ttk.Button(right, text="开始", bootstyle="primary", command=self.start).pack(side="left")
        ttk.Button(right, text="停止", bootstyle="secondary-outline", command=self.stop) \
            .pack(side="left", padx=(10, 0))

        # ===== 进度条 + 日志 =====
        self.pbar = ttk.Progressbar(root, mode="indeterminate", bootstyle="striped")
        self.pbar.pack(fill="x", pady=(8, 6))
        self.pbar.stop()  # 任务开始时可以 self.pbar.start(10)

        log_frame = ttk.Labelframe(root, text="任务日志")
        log_frame.pack(fill="both", expand=True, pady=(0, 10))
        self.log = ScrolledText(log_frame, height=16, autohide=True)
        self.log.pack(fill="both", expand=True)
        self.tick()

    def pick_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.dir_var.set(d)

    def log_put(self, msg):
        self.q.put(msg)

    def tick(self):
        try:
            while True:
                msg = self.q.get_nowait()
                self.log.insert("end", msg + "\n")
                self.log.see("end")
        except queue.Empty:
            pass
        self.root.after(100, self.tick)

    def start(self):
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("提示", "任务正在进行中…")
            return
        url = self.url_var.get().strip()
        out_dir = self.dir_var.get().strip()
        if not url or not out_dir:
            messagebox.showwarning("提示", "请填写网址并选择输出目录")
            return
        os.makedirs(out_dir, exist_ok=True)
        self.stop_flag.clear()
        self.worker = threading.Thread(target=self.run_job, args=(url, out_dir), daemon=True)
        self.worker.start()
        self.pbar.start(12)

    def stop(self):
        self.stop_flag.set()
        self.log_put("[INFO] 已请求停止，当前步骤完成后结束。")

    def try_click_more(self, page):
        if not self.try_more.get(): return False
        clicked = False
        buttons = page.locator("button, a, div[role=button]")
        count = min(buttons.count(), 200)
        for i in range(count):
            el = buttons.nth(i)
            try:
                txt = el.inner_text(timeout=400).strip().lower()
            except Exception:
                continue
            if any(k.lower() in txt for k in CLICK_MORE_TEXTS):
                try:
                    el.scroll_into_view_if_needed(timeout=1000)
                    el.click(timeout=1500)
                    clicked = True
                except Exception:
                    pass
        return clicked

    def auto_scroll_and_collect(self, page, max_scrolls):
        seen_count, no_new = 0, 0
        all_items = []
        for r in range(max_scrolls):
            if self.stop_flag.is_set(): break
            if self.try_click_more(page):
                time.sleep(0.8)
            page.evaluate("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(0.8)
            try:
                items = page.evaluate(JS_COLLECT)
            except Exception:
                items = []
            if len(items) > seen_count:
                seen_count, no_new = len(items), 0
            else:
                no_new += 1
            by_url = { it["url"]: it for it in all_items }
            for it in items:
                by_url[it["url"]] = it
            all_items = list(by_url.values())
            self.log_put(f"[SCROLL] 第{r+1}次，累计图片{len(all_items)}")
            if no_new >= 2: break
        return all_items

    def run_job(self, url, out_dir):
        mode = self.mode.get()
        max_scrolls = self.max_scrolls.get()
        self.log_put("[INFO] 启动浏览器…")
        used = set()

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=self.headless.get())
            ctx = browser.new_context()
            page = ctx.new_page()
            try:
                page.goto(url, timeout=45000, wait_until="domcontentloaded")
                page.wait_for_load_state("networkidle", timeout=15000)
                self.log_put("[INFO] 页面已加载，开始自动滚动/加载更多…")
            except PWTimeout:
                self.log_put("[WARN] 页面加载超时，继续尝试采集…")

            items = self.auto_scroll_and_collect(page, max_scrolls)

            # 保存清单
            manifest = os.path.join(out_dir, "images_manifest.json")
            with open(manifest, "w", encoding="utf-8") as f:
                json.dump(items, f, ensure_ascii=False, indent=2)
            self.log_put(f"[INFO] 采集到 {len(items)} 张图片，已写入清单：{manifest}")

            self.pbar.config(maximum=max(1, len(items)))
            idx = 0
            for i, it in enumerate(items, start=1):
                if self.stop_flag.is_set(): break
                heading = it.get("nearestHeading") or ""
                caption = it.get("caption") or ""
                alt = it.get("alt") or ""
                fallback = os.path.splitext(os.path.basename(urlparse(it["url"]).path))[0]
                base = f"{i:03d}-" + choose_name(heading, caption, alt, fallback)
                name = base
                k = 2
                while name.lower() in used:
                    name = f"{base}-{k}"; k += 1
                used.add(name.lower())

                # ------- 从这里开始替换：三种模式（含 both） -------
                def do_download(name_base):
                    ext = ext_from_url(it["url"])
                    raw_path = os.path.join(out_dir, name_base + ("-orig" + ext if mode == "both" else ext))
                    try:
                        content = fetch_bytes(it["url"])
                        with open(raw_path, "wb") as fp:
                            fp.write(content)
                        self.log_put(f"[SAVE] {raw_path}")
                        # 4:3：下载原图时，转为 PNG（输出 .png）
                        if self.pad_43.get():
                            png_out = raw_path.rsplit(".", 1)[0] + ".png"
                            convert_to_4_3(raw_path, png_out, background_color=(0, 0, 0, 0))
                            self.log_put(f"[4:3] {png_out}")
                        return True
                    except Exception as e:
                        self.log_put(f"[ERR ] 下载失败：{it['url']}  {e}")
                        return False

                def do_capture(name_base):
                    cap_path = os.path.join(out_dir, name_base + ("-cap.png" if mode == "both" else ".png"))
                    try:
                        # 有 css 选择器就做元素级截图；没有就退化到视窗截图
                        if it.get("css"):
                            page.locator(it["css"]).first.scroll_into_view_if_needed(timeout=2000)
                            page.locator(it["css"]).first.screenshot(path=cap_path)
                        else:
                            page.screenshot(path=cap_path, full_page=False)
                        self.log_put(f"[CAP ] {cap_path}")
                        if self.pad_43.get():
                            convert_to_4_3(cap_path, cap_path, background_color=(0, 0, 0, 0))
                            self.log_put(f"[4:3] {cap_path}")
                        return True
                    except Exception as e:
                        self.log_put(f"[ERR ] 截图失败：{e}")
                        return False

                mode = self.mode.get()
                if mode == "download":
                    do_download(name)
                elif mode == "screenshot":
                    do_capture(name)
                else:  # both
                    do_download(name)
                    do_capture(name)
                # ------- 替换到这里结束 -------
                idx += 1
                self.pbar["value"] = idx
                self.root.update_idletasks()

            browser.close()
            self.log_put("[DONE] 任务完成。")
            self.pbar.stop()
            self.pbar["value"] = 0

def _log_crash():
    try:
        (RUNTIME_DIR).mkdir(parents=True, exist_ok=True)
        (RUNTIME_DIR / "last_crash.txt").write_text(traceback.format_exc(), encoding="utf-8")
    except Exception:
        pass

if __name__ == "__main__":
    try:
        _apply_win_dpi_awareness()

        import ttkbootstrap as tb
        # 主题可换：flatly / cosmo / lumen（亮）或 darkly / superhero（暗）
        root = tb.Window(themename="flatly")

        # 首次运行的全局缩放（125% 更舒适；想更大改 1.35/1.5）
        set_scaling_from_system(root, user_factor=1.0)  # 想再大一点改 1.1/1.2
        # 根据屏幕尺寸设置更大的窗口，并居中
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        w = int(sw * 0.65)  # 宽度占屏幕 65%
        h = int(sh * 0.75)  # 高度占屏幕 75%
        root.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

        root.minsize(940, 620)         # 防止被压得太小

        # 可选：Ctrl+/Ctrl- 缩放；Ctrl+0 复位
        def _zoom(delta):
            cur = float(root.call('tk', 'scaling'))
            set_ui_scaling(root, max(0.8, min(2.0, cur + delta)))
        root.bind("<Control-plus>",   lambda e: (_zoom(0.1), None))
        root.bind("<Control-KP_Add>", lambda e: (_zoom(0.1), None))
        root.bind("<Control-minus>",  lambda e: (_zoom(-0.1), None))
        root.bind("<Control-KP_Subtract>", lambda e: (_zoom(-0.1), None))
        root.bind("<Control-0>",      lambda e: (set_ui_scaling(root, 1.25), None))

        App(root)
        root.mainloop()
    except Exception:
        _log_crash()
        raise

