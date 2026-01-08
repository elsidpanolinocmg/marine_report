"""
Newspaper OCR Annotator - Multi-article per page, zoom/pan, annotation -> OCR -> article mapping
+ Auto-detect defaults (Title, Body, Author, Date, Issue)

Dependencies:
- pytesseract
- pdf2image
- Pillow
- python-docx

System deps:
- Tesseract OCR (install separately)
- Poppler (for pdf2image)
"""

import os
import re
import json
import uuid
import threading
import pytesseract
from tkinter import (
    Tk, Frame, Canvas, Button, Label, Entry, Listbox, Scrollbar,
    END, StringVar, filedialog, messagebox
)
from pdf2image import convert_from_path
from PIL import Image, ImageTk
from docx import Document

# --------- CONFIG ----------
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
POPPLER_PATH = None  # set if needed

OUTPUT_DIR = "Processed_Articles"
ANNOT_DIR = "Annotations"
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(ANNOT_DIR, exist_ok=True)

FIELD_COLORS = {
    "Article Title": "blue",
    "Author": "orange",
    "Date Published": "purple",
    "Issue No": "gray",
    "Article Body": "green",
}
FIELD_NAMES = list(FIELD_COLORS.keys())

# ---------- UTIL ----------
def safe_filename(s: str) -> str:
    if not s:
        return "untitled"
    s2 = re.sub(r'[\n\r\t]+', ' ', s)
    s2 = re.sub(r'[\\/*?:"<>|]', '', s2).strip()
    if not s2:
        return "untitled"
    return s2[:120]

def unique_title_for_page(page_articles: dict, title: str) -> str:
    t = title or "untitled"
    base = t
    i = 1
    while t in page_articles:
        t = f"{base} #{i}"
        i += 1
    return t

def clean_ocr_text(text: str) -> str:
    lines = text.splitlines()
    cleaned, buffer = [], ""
    for line in lines:
        line = line.strip()
        if not line:
            if buffer:
                cleaned.append(buffer.strip())
                buffer = ""
            cleaned.append("")
            continue
        if buffer and not re.search(r'[.!?:"”]$', buffer):
            buffer += " " + line
        else:
            if buffer:
                cleaned.append(buffer.strip())
            buffer = line
    if buffer:
        cleaned.append(buffer.strip())
    return "\n".join(cleaned)

# ---------- APP ----------
class NewspaperOCRApp:
    def __init__(self, root):
        self.root = root
        root.title("Newspaper OCR Annotator (multi-article per page + auto-detect)")
        root.geometry("1250x820")

        # State
        self.pages = []
        self.current_page_index = -1
        self.base_scale = 1.0
        self.zoom = 1.0
        self.current_pil_image = None
        self.current_image_tk = None
        self.annotations = []
        self.page_articles = {}

        # UI Layout
        main = Frame(root); main.pack(fill="both", expand=True)

        # Left side: canvas
        left = Frame(main); left.pack(side="left", fill="both", expand=True)
        self.canvas = Canvas(left, bg="#111", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<ButtonPress-1>", self.canvas_mouse_down)
        self.canvas.bind("<B1-Motion>", self.canvas_mouse_move)
        self.canvas.bind("<ButtonRelease-1>", self.canvas_mouse_up)
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)
        self.canvas.bind("<ButtonPress-2>", self.start_pan)
        self.canvas.bind("<B2-Motion>", self.do_pan)
        self.canvas.bind("<ButtonPress-3>", self.start_pan)
        self.canvas.bind("<B3-Motion>", self.do_pan)

        # Toolbar
        tool_frame = Frame(left); tool_frame.pack(fill="x")
        Button(tool_frame, text="Load PDF", command=self.load_pdf).pack(side="left")
        Button(tool_frame, text="Prev Page", command=self.show_prev_page).pack(side="left")
        Button(tool_frame, text="Next Page", command=self.show_next_page).pack(side="left")
        for name in FIELD_NAMES:
            Button(tool_frame, text=name, bg=FIELD_COLORS[name],
                   command=lambda n=name: self.set_selected_field(n)).pack(side="left")
        Button(tool_frame, text="Zoom In", command=lambda: self.zoom_by(1.25)).pack(side="left")
        Button(tool_frame, text="Zoom Out", command=lambda: self.zoom_by(0.8)).pack(side="left")
        self.selected_field = "Article Body"
        self.set_status_var = StringVar(value="Ready")
        Label(tool_frame, textvariable=self.set_status_var).pack(side="right")

        # Right: editor
        right = Frame(main, width=420); right.pack(side="right", fill="y")
        Label(right, text="Articles on this page").pack(anchor="w")
        self.article_listbox = Listbox(right, height=8); self.article_listbox.pack(fill="x")
        self.article_listbox.bind("<<ListboxSelect>>", self.on_article_select)
        Button(right, text="Create Article", command=self.create_blank_article).pack()
        Button(right, text="Delete Article", command=self.delete_selected_article).pack()

        Label(right, text="Article Title:").pack(anchor="w")
        self.title_entry = Entry(right); self.title_entry.pack(fill="x")
        Label(right, text="Author:").pack(anchor="w")
        self.author_entry = Entry(right); self.author_entry.pack(fill="x")
        Label(right, text="Date Published:").pack(anchor="w")
        self.date_entry = Entry(right); self.date_entry.pack(fill="x")
        Label(right, text="Issue No:").pack(anchor="w")
        self.issue_entry = Entry(right); self.issue_entry.pack(fill="x")

        Label(right, text="Article Body Parts:").pack(anchor="w")
        self.body_listbox = Listbox(right, height=10); self.body_listbox.pack(fill="both")

        Button(right, text="Remove Selected Body", command=self.remove_selected_body).pack()
        Button(right, text="Save Article (Word)", command=self.save_selected_article_word).pack()
        Button(right, text="Save Annotations (JSON)", command=self.save_annotations_json).pack()

        self.draw_start = None
        self.temp_rect_id = None

    # ---------- STATUS ----------
    def set_status(self, s: str):
        self.set_status_var.set(s); self.root.update_idletasks()

    # ---------- FIELD ----------
    def set_selected_field(self, field_name):
        if field_name in FIELD_NAMES:
            self.selected_field = field_name
            self.set_status(f"Selected {field_name}")

    # ---------- LOAD PDF ----------
    def load_pdf(self):
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if not pdf_path: return
        self.pages.clear(); self.page_articles.clear(); self.annotations.clear()
        self.current_page_index = -1
        name = os.path.basename(pdf_path)
        self.set_status("Loading PDF...")

        def worker():
            pages = convert_from_path(pdf_path, dpi=200, poppler_path=POPPLER_PATH)
            for i, page in enumerate(pages):
                self.pages.append({"image": page.convert("RGB"), "pdf": name, "page_no": i+1})
            self.root.after(0, lambda: self.display_page(0))
        threading.Thread(target=worker, daemon=True).start()

    # ---------- DISPLAY PAGE ----------
    def display_page(self, idx):
        if idx < 0 or idx >= len(self.pages): return
        self.current_page_index = idx
        page = self.pages[idx]
        pil = page["image"]
        cw, ch = 600, 800
        sw, sh = cw/pil.width, ch/pil.height
        self.base_scale = min(sw, sh, 1.0); self.zoom = 1.0
        self._render_image()

        # auto-detect defaults
        detected = self.auto_detect(pil)
        if detected:
            aid = unique_title_for_page(self.page_articles.setdefault(idx, {}), detected["title"])
            self.page_articles[idx][aid] = detected
            self.refresh_article_listbox()
            self.select_article_by_id(aid)

    def _render_image(self):
        pil = self.pages[self.current_page_index]["image"]
        scale = self.base_scale * self.zoom
        resized = pil.resize((int(pil.width*scale), int(pil.height*scale)), Image.LANCZOS)
        self.current_image_tk = ImageTk.PhotoImage(resized)
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor="nw", image=self.current_image_tk)

    def show_prev_page(self):
        if self.current_page_index > 0: self.display_page(self.current_page_index-1)
    def show_next_page(self):
        if self.current_page_index < len(self.pages)-1: self.display_page(self.current_page_index+1)

    # ---------- AUTO-DETECT ----------
    def auto_detect(self, pil_img):
        data = pytesseract.image_to_data(pil_img, output_type=pytesseract.Output.DICT)
        blocks = {"title": [], "bodies": [], "author": "", "date": "", "issue": ""}
        buffer, last_line = [], -1
        for i, word in enumerate(data["text"]):
            if not word.strip(): continue
            h, line = data["height"][i], data["line_num"][i]
            if h > 40: blocks["title"].append(word)
            if re.search(r"Vol\s*\d+\s*No\.\s*\d+", word, re.I): blocks["issue"] = word
            if re.search(r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)", word, re.I): blocks["date"] = word
            if word.lower().startswith("by "): blocks["author"] = word
            if line != last_line and buffer:
                blocks["bodies"].append(" ".join(buffer)); buffer = []
            buffer.append(word); last_line = line
        if buffer: blocks["bodies"].append(" ".join(buffer))
        return {
            "title": " ".join(blocks["title"]) or "untitled",
            "author": blocks["author"], "date": blocks["date"],
            "issue": blocks["issue"], "bodies": blocks["bodies"], "annotations": []
        }

    # ---------- ARTICLES ----------
    def refresh_article_listbox(self):
        self.article_listbox.delete(0, END)
        for aid in self.page_articles.get(self.current_page_index, {}):
            self.article_listbox.insert(END, aid)

    def get_current_selected_article_id(self):
        sel = self.article_listbox.curselection()
        return self.article_listbox.get(sel[0]) if sel else None

    def select_article_by_id(self, aid):
        L = self.article_listbox.get(0, END)
        for i, v in enumerate(L):
            if v == aid:
                self.article_listbox.selection_clear(0, END)
                self.article_listbox.selection_set(i)
                self.show_article_fields(aid)
                return

    def on_article_select(self, e=None):
        aid = self.get_current_selected_article_id()
        if aid: self.show_article_fields(aid)

    def show_article_fields(self, aid):
        art = self.page_articles.get(self.current_page_index, {}).get(aid)
        if not art: return
        self.title_entry.delete(0, END); self.title_entry.insert(0, art["title"])
        self.author_entry.delete(0, END); self.author_entry.insert(0, art["author"])
        self.date_entry.delete(0, END); self.date_entry.insert(0, art["date"])
        self.issue_entry.delete(0, END); self.issue_entry.insert(0, art["issue"])
        self.body_listbox.delete(0, END)
        for b in art["bodies"]: self.body_listbox.insert(END, b)

    def create_blank_article(self):
        page_articles = self.page_articles.setdefault(self.current_page_index, {})
        aid = unique_title_for_page(page_articles, "untitled")
        page_articles[aid] = {"title": aid, "author": "", "date": "", "issue": "", "bodies": [], "annotations": []}
        self.refresh_article_listbox(); self.select_article_by_id(aid)

    def delete_selected_article(self):
        aid = self.get_current_selected_article_id()
        if not aid: return
        self.page_articles[self.current_page_index].pop(aid, None)
        self.refresh_article_listbox()

    def remove_selected_body(self):
        sel = self.body_listbox.curselection()
        if not sel: return
        idx = sel[0]
        aid = self.get_current_selected_article_id()
        art = self.page_articles.get(self.current_page_index, {}).get(aid)
        if art and idx < len(art["bodies"]):
            art["bodies"].pop(idx); self.body_listbox.delete(idx)

    # ---------- SAVE ----------
    def save_annotations_json(self):
        page = self.pages[self.current_page_index]
        arts = self.page_articles.get(self.current_page_index, {})
        out = {"pdf": page["pdf"], "page_no": page["page_no"], "articles": arts}
        fname = f"{page['pdf']}_p{page['page_no']}.json"
        path = os.path.join(ANNOT_DIR, fname)
        with open(path, "w", encoding="utf-8") as f: json.dump(out, f, indent=2, ensure_ascii=False)
        messagebox.showinfo("Saved", f"Annotations saved to {path}")

    def save_selected_article_word(self):
        aid = self.get_current_selected_article_id()
        if not aid: return
        art = self.page_articles.get(self.current_page_index, {}).get(aid)
        if not art: return
        doc = Document()
        table = doc.add_table(rows=5, cols=2)
        labels = ["Article Title:", "Author:", "Date Published:", "Issue No:", "Article Body:"]
        values = [art["title"], art["author"], art["date"], art["issue"], "\n\n".join(art["bodies"])]
        for i, lab in enumerate(labels): table.cell(i,0).text, table.cell(i,1).text = lab, values[i]
        outpath = os.path.join(OUTPUT_DIR, f"{safe_filename(art['title'])}.docx")
        doc.save(outpath)
        messagebox.showinfo("Saved", f"Article exported to {outpath}")

    # ---------- CANVAS EVENTS ----------
    def canvas_mouse_down(self, e): self.draw_start = (self.canvas.canvasx(e.x), self.canvas.canvasy(e.y))
    def canvas_mouse_move(self, e):
        if self.draw_start and not self.temp_rect_id:
            x0,y0=self.draw_start;x1,y1=self.canvas.canvasx(e.x),self.canvas.canvasy(e.y)
            self.temp_rect_id=self.canvas.create_rectangle(x0,y0,x1,y1,outline="yellow")
    def canvas_mouse_up(self, e): self.draw_start=None; self.temp_rect_id=None
    def zoom_by(self, f): self.zoom=max(0.2,min(5.0,self.zoom*f)); self._render_image()
    def on_mousewheel(self, e): self.zoom_by(1.1 if e.delta>0 else 0.9)
    def start_pan(self, e): self.canvas.scan_mark(e.x,e.y)
    def do_pan(self, e): self.canvas.scan_dragto(e.x,e.y, gain=1)

# ---------- RUN ----------
if __name__=="__main__":
    root=Tk(); app=NewspaperOCRApp(root); root.mainloop()
