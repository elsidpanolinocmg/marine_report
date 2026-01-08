"""
Newspaper OCR Annotator - Multi-article per page, zoom/pan, annotation -> OCR -> article mapping

Dependencies:
- pytesseract
- pdf2image
- Pillow
- python-docx

System deps:
- Tesseract OCR (install separately)
- Poppler (for pdf2image) - install and add bin to PATH or set POPPLER_PATH below
"""

import os
import re
import json
import uuid
import threading
import pytesseract
from tkinter import (
    Tk, Frame, Canvas, Button, Label, Entry, Text, Listbox, Scrollbar,
    END, StringVar, filedialog, messagebox
)
from pdf2image import convert_from_path
from PIL import Image, ImageTk
from docx import Document

# --------- CONFIG ----------
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"  # adjust if needed
POPPLER_PATH = None  # e.g. r"C:\poppler-xx\bin" or None if poppler in PATH

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
    s2 = re.sub(r'[\\/*?:"<>|]', '', s2)
    s2 = s2.strip()
    if not s2:
        return "untitled"
    return s2[:120]

def unique_title_for_page(page_articles: dict, title: str) -> str:
    # ensure title is unique as key; if duplicate, append suffix
    t = title or "untitled"
    base = t
    i = 1
    while t in page_articles:
        t = f"{base} #{i}"
        i += 1
    return t

def clean_ocr_text(text: str) -> str:
    import re
    lines = text.splitlines()
    cleaned = []
    buffer = ""

    for line in lines:
        line = line.strip()
        if not line:
            # Empty line → paragraph break
            if buffer:
                cleaned.append(buffer.strip())
                buffer = ""
            cleaned.append("")  # keep blank line
            continue

        if buffer:
            # If previous line doesn’t end with sentence punctuation, join
            if not re.search(r'[.!?:"”]$', buffer):
                buffer += " " + line
            else:
                cleaned.append(buffer.strip())
                buffer = line
        else:
            buffer = line

    if buffer:
        cleaned.append(buffer.strip())

    return "\n".join(cleaned)

# ---------- APP ----------
class NewspaperOCRApp:
    def __init__(self, root):
        self.root = root
        root.title("Newspaper OCR Annotator (multi-article per page)")
        root.geometry("1250x820")

        # State
        self.pages = []  # list of dicts: {"image": PIL.Image, "pdf": filename, "page_no": int}
        self.current_page_index = -1

        self.base_scale = 1.0  # scale to fit
        self.zoom = 1.0        # dynamic zoom multiplier
        self.current_pil_image = None
        self.current_image_tk = None

        # annotations: list of dicts: {id, page_index, bbox_orig (x1,y1,x2,y2), label, ocr_text, article_id (optional)}
        self.annotations = []

        # page_articles: dict page_index -> dict article_id -> article data
        # article data: {"title","author","date","issue","bodies": [str], "annotations": [ann_id]}
        self.page_articles = {}

        # UI - layout (left: image canvas, right: article editor/list)
        main = Frame(root)
        main.pack(fill="both", expand=True)

        # Left: canvas area
        left = Frame(main)
        left.pack(side="left", fill="both", expand=True)

        canvas_frame = Frame(left)
        canvas_frame.pack(fill="both", expand=True)
        self.canvas = Canvas(canvas_frame, bg="#111", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        vsb = Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(left, orient="horizontal", command=self.canvas.xview)
        hsb.pack(side="bottom", fill="x")
        self.canvas.config(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Bind canvas events: drawing, pan/zoom
        self.canvas.bind("<ButtonPress-1>", self.canvas_mouse_down)
        self.canvas.bind("<B1-Motion>", self.canvas_mouse_move)
        self.canvas.bind("<ButtonRelease-1>", self.canvas_mouse_up)

        # pan with middle or right mouse
        self.canvas.bind("<ButtonPress-2>", self.start_pan)
        self.canvas.bind("<B2-Motion>", self.do_pan)
        self.canvas.bind("<ButtonPress-3>", self.start_pan)
        self.canvas.bind("<B3-Motion>", self.do_pan)

        # mouse wheel for zoom
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)

        # Tools above canvas
        tool_frame = Frame(left)
        tool_frame.pack(fill="x")
        Button(tool_frame, text="Load PDF", command=self.load_pdf).pack(side="left", padx=4, pady=4)
        Button(tool_frame, text="Prev Page", command=self.show_prev_page).pack(side="left", padx=4)
        Button(tool_frame, text="Next Page", command=self.show_next_page).pack(side="left", padx=4)
        Label(tool_frame, text="   Select field:").pack(side="left", padx=(8,4))
        for name in FIELD_NAMES:
            Button(tool_frame, text=name, bg=FIELD_COLORS[name],
                   command=lambda n=name: self.set_selected_field(n)).pack(side="left", padx=2)
        Button(tool_frame, text="Zoom In", command=lambda: self.zoom_by(1.25)).pack(side="left", padx=6)
        Button(tool_frame, text="Zoom Out", command=lambda: self.zoom_by(1/1.25)).pack(side="left", padx=2)

        self.selected_field = "Article Body"
        self.set_status_var = StringVar(value="Ready")
        Label(tool_frame, textvariable=self.set_status_var).pack(side="right", padx=6)

        # Right: editor & article list
        right = Frame(main, width=420)
        right.pack(side="right", fill="y")

        Label(right, text="Articles on this page").pack(anchor="w", padx=8, pady=(8,0))
        self.article_listbox = Listbox(right, height=8)
        self.article_listbox.pack(fill="x", padx=8)
        self.article_listbox.bind("<<ListboxSelect>>", self.on_article_select)

        art_btns = Frame(right)
        art_btns.pack(fill="x", padx=8, pady=(4,6))
        Button(art_btns, text="Create Article", command=self.create_blank_article).pack(side="left", padx=4)
        Button(art_btns, text="Delete Article", command=self.delete_selected_article).pack(side="left", padx=4)
        Button(art_btns, text="Prev Article", command=self.prev_article).pack(side="left", padx=4)
        Button(art_btns, text="Next Article", command=self.next_article).pack(side="left", padx=4)

        Label(right, text="Article Title:").pack(anchor="w", padx=8, pady=(8,0))
        self.title_entry = Entry(right)
        self.title_entry.pack(fill="x", padx=8)

        Label(right, text="Author:").pack(anchor="w", padx=8, pady=(8,0))
        self.author_entry = Entry(right)
        self.author_entry.pack(fill="x", padx=8)

        Label(right, text="Date Published:").pack(anchor="w", padx=8, pady=(8,0))
        self.date_entry = Entry(right)
        self.date_entry.pack(fill="x", padx=8)

        Label(right, text="Issue No:").pack(anchor="w", padx=8, pady=(8,0))
        self.issue_entry = Entry(right)
        self.issue_entry.pack(fill="x", padx=8)

        Label(right, text="Article Body Parts (multiple):").pack(anchor="w", padx=8, pady=(8,0))
        self.body_listbox = Listbox(right, height=10)
        self.body_listbox.pack(fill="both", padx=8, pady=(0,8), expand=False)

        body_btn_frame = Frame(right)
        body_btn_frame.pack(fill="x", padx=8, pady=(0,8))
        Button(body_btn_frame, text="Remove Selected Body", command=self.remove_selected_body).pack(side="left", padx=4)
        Button(body_btn_frame, text="Save Article (Word)", command=self.save_selected_article_word).pack(side="right", padx=4)
        Button(body_btn_frame, text="Save Annotations (JSON)", command=self.save_annotations_json).pack(side="right", padx=6)

        # drawing state
        self.draw_start = None
        self.temp_rect_id = None

        # initial status
        self.set_status("Ready")

    # ------- status -------
    def set_status(self, s: str):
        self.set_status_var.set(s)
        self.root.update_idletasks()

    # ------- field selection -------
    def set_selected_field(self, field_name):
        if field_name in FIELD_NAMES:
            self.selected_field = field_name
            self.set_status(f"Selected field: {field_name}")

    # ------- load PDFs (background) -------
    def load_pdf(self):
        pdf_path = filedialog.askopenfilename(
            title="Select a PDF file",
            filetypes=[("PDF files", "*.pdf")]
        )
        if not pdf_path:
            return

        # Clear previous
        self.pages.clear()
        self.annotations.clear()
        self.page_articles.clear()
        self.current_page_index = -1
        self.canvas.delete("all")
        self.body_listbox.delete(0, END)
        self.title_entry.delete(0, END)
        self.author_entry.delete(0, END)
        self.date_entry.delete(0, END)
        self.issue_entry.delete(0, END)

        name = os.path.basename(pdf_path)

        self.set_status("Loading PDF pages...")
        thread = threading.Thread(target=self._load_pdf_thread, args=(pdf_path, name), daemon=True)
        thread.start()

    def _load_pdf_thread(self, pdf_path, name):
        try:
            pages = convert_from_path(pdf_path, dpi=200, poppler_path=POPPLER_PATH) \
                if POPPLER_PATH else convert_from_path(pdf_path, dpi=200)
        except Exception as e:
            print("convert_from_path error:", e)
            return

        for i, page in enumerate(pages):
            self.pages.append({"image": page.convert("RGB"), "pdf": name, "page_no": i+1})
            idx = len(self.pages)-1
            self.page_articles[idx] = {
                # "title": "", "author": "", "date": "",
                # "issue": name.replace(".pdf",""), "bodies": []
            }
            if self.current_page_index == -1:
                self.root.after(0, lambda: self.display_page(0))
            self.root.after(0, lambda msg=f"Loaded {name} page {i+1}": self.set_status(msg))

        self.root.after(0, lambda: self.set_status("PDF loaded successfully."))

    # ------- display & navigation -------
    def display_page(self, page_index: int):
        if page_index < 0 or page_index >= len(self.pages):
            return
        self.current_page_index = page_index
        page = self.pages[page_index]
        pil = page["image"]

        # compute base scale (fit)
        canvas_w = max(600, int(self.root.winfo_width() * 0.55))
        canvas_h = max(400, int(self.root.winfo_height() * 0.75))
        orig_w, orig_h = pil.size
        scale_w = canvas_w / orig_w
        scale_h = canvas_h / orig_h
        self.base_scale = min(scale_w, scale_h, 1.0)
        self.zoom = 1.0

        self._render_image()  # draws image and annotations

        # Load article list for page
        self.refresh_article_listbox()
        # Load first article's fields if exists
        page_articles = self.page_articles.get(page_index, {})
        if page_articles:
            # pick first article
            first_id = next(iter(page_articles))
            self.select_article_by_id(first_id)
        else:
            # clear right panel
            self.title_entry.delete(0, END)
            self.author_entry.delete(0, END)
            self.date_entry.delete(0, END)
            self.issue_entry.delete(0, END)
            self.body_listbox.delete(0, END)

        self.set_status(f"Showing page {page['pdf']} (page {page['page_no']})")

    def _render_image(self):
        page = self.pages[self.current_page_index]
        pil = page["image"]
        scale = self.base_scale * self.zoom
        disp_w = int(pil.width * scale)
        disp_h = int(pil.height * scale)
        resized = pil.resize((disp_w, disp_h), Image.LANCZOS)
        self.current_pil_image = pil  # keep original for cropping
        self.current_image_tk = ImageTk.PhotoImage(resized)

        self.canvas.delete("all")
        self.canvas.config(scrollregion=(0, 0, disp_w, disp_h))
        self.canvas.create_image(0, 0, anchor="nw", image=self.current_image_tk, tags=("page_image",))
        # draw annotations for this page (scaled)
        self.redraw_annotations_for_page(self.current_page_index)

    def show_prev_page(self):
        if self.current_page_index > 0:
            self.display_page(self.current_page_index - 1)

    def show_next_page(self):
        if self.current_page_index < len(self.pages) - 1:
            self.display_page(self.current_page_index + 1)

    # zoom helpers
    def zoom_by(self, factor: float):
        self.zoom *= factor
        # clamp zoom
        self.zoom = max(0.2, min(self.zoom, 5.0))
        self._render_image()

    def on_mousewheel(self, event):
        # zoom centered on viewport (simple)
        if event.delta > 0:
            self.zoom_by(1.1)
        else:
            self.zoom_by(1/1.1)

    # pan helpers
    def start_pan(self, event):
        self.canvas.scan_mark(event.x, event.y)

    def do_pan(self, event):
        self.canvas.scan_dragto(event.x, event.y, gain=1)

    # ------- annotation drawing (display coords -> original coords) -------
    def canvas_mouse_down(self, event):
        if self.current_page_index == -1:
            return
        self.canvas.focus_set()
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        self.draw_start = (x, y)
        self.temp_rect_id = self.canvas.create_rectangle(x, y, x, y, outline="yellow", width=2)

    def canvas_mouse_move(self, event):
        if not getattr(self, "draw_start", None) or not self.temp_rect_id:
            return
        x0, y0 = self.draw_start
        x1 = self.canvas.canvasx(event.x)
        y1 = self.canvas.canvasy(event.y)
        self.canvas.coords(self.temp_rect_id, x0, y0, x1, y1)

    def canvas_mouse_up(self, event):
        if not getattr(self, "draw_start", None) or not self.temp_rect_id:
            return
        x0, y0 = self.draw_start
        x1 = self.canvas.canvasx(event.x)
        y1 = self.canvas.canvasy(event.y)
        # clean up temporary rect
        self.canvas.delete(self.temp_rect_id)
        self.temp_rect_id = None
        self.draw_start = None

        # normalize
        x1f, x2f = sorted((x0, x1))
        y1f, y2f = sorted((y0, y1))
        if abs(x2f - x1f) < 8 or abs(y2f - y1f) < 6:
            self.set_status("Selection too small, ignored.")
            return

        # convert display coords -> original image coords (orig = disp / (base_scale*zoom))
        scale = self.base_scale * self.zoom
        orig_x1 = int(x1f / scale)
        orig_y1 = int(y1f / scale)
        orig_x2 = int(x2f / scale)
        orig_y2 = int(y2f / scale)
        pil = self.pages[self.current_page_index]["image"]
        orig_x1 = max(0, min(orig_x1, pil.width - 1))
        orig_x2 = max(1, min(orig_x2, pil.width))
        orig_y1 = max(0, min(orig_y1, pil.height - 1))
        orig_y2 = max(1, min(orig_y2, pil.height))

        label = self.selected_field
        ann_id = str(uuid.uuid4())
        # draw permanent rectangle in display coords (so user sees it immediately)
        color = FIELD_COLORS.get(label, "yellow")
        disp_coords = (x1f, y1f, x2f, y2f)
        rect_id = self.canvas.create_rectangle(*disp_coords, outline=color, width=2, tags=("annotation", ann_id))
        self.canvas.create_text(x1f + 6, y1f + 6, text=label[0], anchor="nw", fill=color,
                                font=("Arial", 10, "bold"), tags=("annotation", ann_id))

        ann = {
            "id": ann_id,
            "page_index": self.current_page_index,
            "bbox_orig": (orig_x1, orig_y1, orig_x2, orig_y2),
            "label": label,
            "ocr_text": "[processing...]",
            "article_id": None,
            "canvas_ids": [rect_id],
        }
        self.annotations.append(ann)

        # OCR in background and assign to article if possible
        self.set_status(f"OCR {label} on page {self.current_page_index+1} ...")
        threading.Thread(target=self._ocr_crop_and_assign, args=(ann,), daemon=True).start()

    def _ocr_crop_and_assign(self, ann):
        try:
            pil = self.pages[ann["page_index"]]["image"]
            x1, y1, x2, y2 = ann["bbox_orig"]
            crop = pil.crop((x1, y1, x2, y2))

            raw_text = pytesseract.image_to_string(crop, config="--psm 6")
            cleaned_text = clean_ocr_text(raw_text)
            text = cleaned_text.strip() or "[no text]"
            
        except Exception as e:
            text = f"[ocr error: {e}]"

        def finish():
            ann["ocr_text"] = text
            # If this is a Title annotation -> create (or update) an article keyed by that title
            page_idx = ann["page_index"]
            page_articles = self.page_articles.setdefault(page_idx, {})

            if ann["label"] == "Article Title":
                # use OCR'd title (sanitized). Ensure uniqueness.
                candidate = text.strip() or "untitled"
                article_id = unique_title_for_page(page_articles, candidate)
                # create article if absent
                art = page_articles.setdefault(article_id, {"title": article_id, "author": "", "date": self.pages[page_idx]["pdf"].replace(".pdf", ""), "issue": "", "bodies": [], "annotations": []})
                # link annotation to article
                ann["article_id"] = article_id
                art["annotations"].append(ann["id"])
                # set title field UI if current page/article
                if self.current_page_index == page_idx:
                    self.refresh_article_listbox()
                    self.select_article_by_id(article_id)
                self.set_status(f"Created/updated article: {article_id}")
            else:
                # Non-title: try attach to currently selected article on page, else last created article
                current_article_id = self.get_current_selected_article_id()
                if not current_article_id:
                    # pick last article on this page if any
                    if page_articles:
                        current_article_id = next(reversed(page_articles))
                if current_article_id:
                    ann["article_id"] = current_article_id
                    art = page_articles[current_article_id]
                    art["annotations"].append(ann["id"])
                    # map text to correct field
                    if ann["label"] == "Author":
                        art["author"] = text
                    elif ann["label"] == "Date Published":
                        art["date"] = text
                    elif ann["label"] == "Issue No":
                        art["issue"] = text
                    elif ann["label"] == "Article Body":
                        art["bodies"].append(text)
                    # if current article is selected, refresh UI
                    if self.current_page_index == page_idx and self.get_current_selected_article_id() == current_article_id:
                        self.show_article_fields(current_article_id)
                    self.set_status(f"Assigned annotation to article: {current_article_id}")
                else:
                    # no article to assign; keep unassigned (user can assign later)
                    self.set_status("OCR done: annotation unassigned. Create/select article to assign.")
            # refresh annotation overlays text label etc
            self.redraw_annotations_for_page(self.current_page_index)

        self.root.after(0, finish)

    # redraw annotations (converted from orig coords -> display coords)
    def redraw_annotations_for_page(self, page_index):
        # remove annotation-tagged items
        for item in self.canvas.find_withtag("annotation"):
            try:
                self.canvas.delete(item)
            except Exception:
                pass
        # draw annotations belonging to this page
        scale = self.base_scale * self.zoom
        for ann in self.annotations:
            if ann["page_index"] != page_index:
                continue
            x1, y1, x2, y2 = ann["bbox_orig"]
            dx1 = x1 * scale; dy1 = y1 * scale; dx2 = x2 * scale; dy2 = y2 * scale
            color = FIELD_COLORS.get(ann["label"], "yellow")
            cid = self.canvas.create_rectangle(dx1, dy1, dx2, dy2, outline=color, width=2, tags=("annotation",))
            self.canvas.create_text(dx1 + 6, dy1 + 6, text=ann["label"][0], anchor="nw", fill=color, font=("Arial", 10, "bold"), tags=("annotation",))
            ann["canvas_ids"] = [cid]

    # ------- article management UI -------
    def refresh_article_listbox(self):
        self.article_listbox.delete(0, END)
        page_idx = self.current_page_index
        if page_idx == -1:
            return
        page_articles = self.page_articles.get(page_idx, {})
        for aid in page_articles:
            self.article_listbox.insert(END, aid)

    def get_current_selected_article_id(self):
        sel = self.article_listbox.curselection()
        if not sel:
            return None
        idx = sel[0]
        return self.article_listbox.get(idx)

    def on_article_select(self, event=None):
        aid = self.get_current_selected_article_id()
        if aid:
            self.show_article_fields(aid)

    def select_article_by_id(self, article_id: str):
        # select in listbox if present
        L = self.article_listbox.get(0, END)
        for i, v in enumerate(L):
            if v == article_id:
                self.article_listbox.selection_clear(0, END)
                self.article_listbox.selection_set(i)
                self.article_listbox.see(i)
                self.show_article_fields(article_id)
                return

    def show_article_fields(self, article_id: str):
        page_idx = self.current_page_index
        page_articles = self.page_articles.get(page_idx, {})
        art = page_articles.get(article_id)
        if not art:
            return
        # populate UI fields
        self.title_entry.delete(0, END); self.title_entry.insert(0, art.get("title", ""))
        self.author_entry.delete(0, END); self.author_entry.insert(0, art.get("author", ""))
        self.date_entry.delete(0, END); self.date_entry.insert(0, art.get("date", ""))
        self.issue_entry.delete(0, END); self.issue_entry.insert(0, art.get("issue", ""))
        # bodies
        self.body_listbox.delete(0, END)
        for b in art.get("bodies", []):
            self.body_listbox.insert(END, b)
        self.set_status(f"Editing article: {article_id}")

    def create_blank_article(self):
        if self.current_page_index == -1:
            messagebox.showinfo("No page", "Load a page first.")
            return
        page_idx = self.current_page_index
        page_articles = self.page_articles.setdefault(page_idx, {})
        new_title = unique_title_for_page(page_articles, "untitled")
        page_articles[new_title] = {"title": new_title, "author": "", "date": "", "issue": self.pages[page_idx]["pdf"].replace(".pdf", ""), "bodies": [], "annotations": []}
        self.refresh_article_listbox()
        self.select_article_by_id(new_title)
        self.set_status(f"Created blank article: {new_title}")

    def delete_selected_article(self):
        aid = self.get_current_selected_article_id()
        if not aid:
            messagebox.showinfo("Select article", "Select an article to delete.")
            return
        page_idx = self.current_page_index
        page_articles = self.page_articles.get(page_idx, {})
        art = page_articles.pop(aid, None)
        # detach annotations referencing this article
        if art:
            for ann_id in art.get("annotations", []):
                for ann in self.annotations:
                    if ann["id"] == ann_id:
                        ann["article_id"] = None
        self.refresh_article_listbox()
        # clear right panel
        self.title_entry.delete(0, END); self.author_entry.delete(0, END); self.date_entry.delete(0, END); self.issue_entry.delete(0, END)
        self.body_listbox.delete(0, END)
        self.set_status(f"Deleted article: {aid}")

    def prev_article(self):
        size = self.article_listbox.size()
        if size == 0:
            return
        sel = self.article_listbox.curselection()
        idx = sel[0] if sel else 0
        idx = max(0, idx - 1)
        self.article_listbox.selection_clear(0, END); self.article_listbox.selection_set(idx); self.article_listbox.see(idx)
        self.on_article_select()

    def next_article(self):
        size = self.article_listbox.size()
        if size == 0:
            return
        sel = self.article_listbox.curselection()
        idx = sel[0] if sel else -1
        idx = min(size - 1, idx + 1)
        self.article_listbox.selection_clear(0, END); self.article_listbox.selection_set(idx); self.article_listbox.see(idx)
        self.on_article_select()

    # when user edits title/fields and wants to persist them, provide a simple "save fields" on focus-out or when saving article
    def persist_current_article_fields(self):
        aid = self.get_current_selected_article_id()
        if not aid:
            return
        page_idx = self.current_page_index
        page_articles = self.page_articles.get(page_idx, {})
        art = page_articles.get(aid)
        if not art:
            return
        # read current entries
        new_title = self.title_entry.get().strip() or art.get("title", "")
        if new_title != aid:
            # need to change the article id (title used as id)
            new_id = unique_title_for_page(page_articles, new_title)
            # re-key
            page_articles[new_id] = art
            del page_articles[aid]
            art["title"] = new_id
            # update annotations linking
            for ann_id in art.get("annotations", []):
                for ann in self.annotations:
                    if ann["id"] == ann_id:
                        ann["article_id"] = new_id
            self.refresh_article_listbox()
            self.select_article_by_id(new_id)
            aid = new_id
        # update other fields
        art["author"] = self.author_entry.get().strip()
        art["date"] = self.date_entry.get().strip()
        art["issue"] = self.issue_entry.get().strip()
        # bodies are maintained via body_listbox and remove operation

    def remove_selected_body(self):
        sel = self.body_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        aid = self.get_current_selected_article_id()
        if not aid:
            return
        page_idx = self.current_page_index
        art = self.page_articles.get(page_idx, {}).get(aid)
        if not art:
            return
        if idx < len(art.get("bodies", [])):
            art["bodies"].pop(idx)
        self.body_listbox.delete(idx)
        self.set_status("Removed body part.")

    # ------- saving: annotations json & article docx -------
    def save_annotations_json(self):
        if self.current_page_index == -1:
            messagebox.showinfo("No page", "Nothing to save.")
            return
        page = self.pages[self.current_page_index]
        page_idx = self.current_page_index
        anns = [a for a in self.annotations if a["page_index"] == page_idx]
        articles = self.page_articles.get(page_idx, {})
        out = {
            "pdf": page["pdf"],
            "page_no": page["page_no"],
            "annotations": [{"id": a["id"], "label": a["label"], "bbox_orig": a["bbox_orig"], "ocr_text": a.get("ocr_text",""), "article_id": a.get("article_id")} for a in anns],
            "articles": {aid: {"title":ad.get("title"), "author":ad.get("author"), "date":ad.get("date"), "issue":ad.get("issue"), "bodies": ad.get("bodies",[])} for aid,ad in articles.items()}
        }
        fname = f"{page['pdf']}_p{page['page_no']}.json"
        path = os.path.join(ANNOT_DIR, fname)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2, ensure_ascii=False)
        self.set_status(f"Annotations saved: {path}")
        messagebox.showinfo("Saved", f"Annotations saved to\n{path}")

    def save_selected_article_word(self):
        if self.current_page_index == -1:
            messagebox.showinfo("No page", "No page selected.")
            return
        aid = self.get_current_selected_article_id()
        if not aid:
            messagebox.showinfo("No article", "Select or create an article first.")
            return
        # persist any edits first
        self.persist_current_article_fields()

        page_idx = self.current_page_index
        art = self.page_articles.get(page_idx, {}).get(aid)
        if not art:
            messagebox.showinfo("No article", "Article not found.")
            return

        # build docx
        doc = Document()
        table = doc.add_table(rows=5, cols=2)
        labels = ["Article Title:", "Author:", "Date Published:", "Issue No:", "Article Body:"]
        body_text = "\n\n".join(art.get("bodies", [])) or "[no body captured]"
        values = [art.get("title",""), art.get("author",""), art.get("date",""), art.get("issue",""), body_text]
        for i, lab in enumerate(labels):
            table.cell(i,0).text = lab
            table.cell(i,1).text = values[i]

        fname = safe_filename(art.get("title") or f"page{page_idx+1}_{uuid.uuid4().hex[:6]}")
        outpath = os.path.join(OUTPUT_DIR, f"{fname}.docx")
        try:
            doc.save(outpath)
        except Exception as e:
            messagebox.showerror("Save error", f"Error saving file:\n{e}")
            return
        self.set_status(f"Saved Word: {outpath}")
        messagebox.showinfo("Saved", f"Article exported to\n{outpath}")

# ---------- run ----------
def main():
    root = Tk()
    app = NewspaperOCRApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
