"""
Newspaper OCR Annotator - Multi-article per page, zoom/pan, annotation -> article mapping.

Dependencies:
 - pytesseract
 - pdf2image
 - Pillow
 - python-docx
 - tkinter (part of stdlib)

System deps:
 - Tesseract OCR (install separately)
 - Poppler (for pdf2image) - set POPPLER_PATH if not in PATH
"""

import os
import re
import json
import threading
import uuid
import pytesseract
from tkinter import (
    Tk, Frame, Canvas, Button, Label, Entry, Text, Listbox, Scrollbar,
    END, StringVar, filedialog, messagebox, LEFT, RIGHT, X, Y, BOTH
)
from tkinter import ttk
from pdf2image import convert_from_path
from PIL import Image, ImageTk
from docx import Document

# ----------------- CONFIG -----------------
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

# ----------------- UTIL -----------------
def safe_filename(s: str) -> str:
    if not s:
        return "untitled"
    cleaned = re.sub(r'[\n\r\t]+', ' ', s)
    cleaned = re.sub(r'[\\/*?:"<>|]', "", cleaned)
    cleaned = cleaned.strip()
    return cleaned[:150] or "untitled"

# ----------------- APP -----------------
class NewspaperOCRApp:
    def __init__(self, root):
        self.root = root
        root.title("Newspaper OCR — Multi-article")
        root.geometry("1300x820")

        # Data structures
        self.pages = []  # list of dict {"image": PIL.Image, "pdf": name, "page_no": n}
        self.current_page_index = -1

        # annotations = list of dicts:
        # {id, page_index, bbox_orig:(x1,y1,x2,y2), label, ocr_text, article_title_or_None}
        self.annotations = []

        # page_articles: { page_index: { title: {title, author, date, issue, bodies:[...], annotations:[ann_id,...]} } }
        self.page_articles = {}

        # display scaling / zoom
        self.base_scale = 1.0
        self.zoom = 1.0
        self.current_pil_image = None
        self.current_image_tk = None

        # drawing state
        self.draw_start = None
        self.temp_rect_id = None

        # selected label
        self.selected_field = "Article Body"

        # selected article title currently shown on right panel
        self.selected_article_title = None

        # ---------------- UI Layout ----------------
        main = Frame(root)
        main.pack(fill=BOTH, expand=True)

        # LEFT: canvas + tools
        left_col = Frame(main)
        left_col.pack(side=LEFT, fill=BOTH, expand=True)

        # Canvas frame (with scrollbars)
        canvas_frame = Frame(left_col)
        canvas_frame.pack(fill=BOTH, expand=True)
        self.canvas = Canvas(canvas_frame, bg="#222")
        self.canvas.pack(side=LEFT, fill=BOTH, expand=True)
        vsb = Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        vsb.pack(side=RIGHT, fill=Y)
        hsb = Scrollbar(left_col, orient="horizontal", command=self.canvas.xview)
        hsb.pack(side="bottom", fill=X)
        self.canvas.config(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Bind canvas events
        self.canvas.bind("<ButtonPress-1>", self.canvas_mouse_down)
        self.canvas.bind("<B1-Motion>", self.canvas_mouse_move)
        self.canvas.bind("<ButtonRelease-1>", self.canvas_mouse_up)
        # pan with middle or right mouse
        self.canvas.bind("<ButtonPress-2>", self.start_pan)
        self.canvas.bind("<B2-Motion>", self.do_pan)
        self.canvas.bind("<ButtonPress-3>", self.start_pan)
        self.canvas.bind("<B3-Motion>", self.do_pan)
        # zoom with mouse wheel (Windows)
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)
        # mac/linux might need Button-4/5 binds but many environments handle MouseWheel

        # Tools above canvas
        tool_frame = Frame(left_col)
        tool_frame.pack(fill=X)
        Button(tool_frame, text="Load Folder", command=self.load_folder).pack(side=LEFT, padx=4, pady=4)
        Button(tool_frame, text="Prev Page", command=self.prev_page).pack(side=LEFT, padx=4)
        Button(tool_frame, text="Next Page", command=self.next_page).pack(side=LEFT, padx=4)
        Label(tool_frame, text=" | Select label: ").pack(side=LEFT)
        for fname in FIELD_NAMES:
            btn = ttk.Radiobutton(tool_frame, text=fname, value=fname, variable=StringVar(value=self.selected_field),
                                  command=lambda f=fname: self.select_field(f))
            # ttk Radiobutton with StringVar per-call is clumsy — use simple Button for reliability
        # We'll create label buttons instead (keeps things simple)
        for fname in FIELD_NAMES:
            Button(tool_frame, text=fname, bg=FIELD_COLORS[fname], command=lambda f=fname: self.select_field(f)).pack(side=LEFT, padx=2)

        # Status label
        self.status_var = StringVar(value="Ready")
        Label(tool_frame, textvariable=self.status_var).pack(side="right", padx=6)

        # RIGHT: article editor & article list
        right_col = Frame(main, width=420)
        right_col.pack(side=RIGHT, fill="y")

        Label(right_col, text="Articles on page", font=("Arial", 11, "bold")).pack(anchor="w", padx=8, pady=(8,0))
        self.article_listbox = Listbox(right_col, height=8)
        self.article_listbox.pack(fill=X, padx=8)
        self.article_listbox.bind("<<ListboxSelect>>", self.on_article_select)

        art_btns = Frame(right_col)
        art_btns.pack(fill=X, padx=8, pady=(4,8))
        Button(art_btns, text="Create Blank Article", command=self.create_blank_article).pack(side=LEFT, padx=2)
        Button(art_btns, text="Delete Article", command=self.delete_selected_article).pack(side=LEFT, padx=2)
        Button(art_btns, text="Prev Article", command=self.prev_article).pack(side=LEFT, padx=2)
        Button(art_btns, text="Next Article", command=self.next_article).pack(side=LEFT, padx=2)

        # Article fields
        Label(right_col, text="Article Title:").pack(anchor="w", padx=8, pady=(8,0))
        self.title_entry = Entry(right_col)
        self.title_entry.pack(fill=X, padx=8)
        Label(right_col, text="Author:").pack(anchor="w", padx=8, pady=(8,0))
        self.author_entry = Entry(right_col)
        self.author_entry.pack(fill=X, padx=8)
        Label(right_col, text="Date Published:").pack(anchor="w", padx=8, pady=(8,0))
        self.date_entry = Entry(right_col)
        self.date_entry.pack(fill=X, padx=8)
        Label(right_col, text="Issue No:").pack(anchor="w", padx=8, pady=(8,0))
        self.issue_entry = Entry(right_col)
        self.issue_entry.pack(fill=X, padx=8)

        Label(right_col, text="Article Body Parts:").pack(anchor="w", padx=8, pady=(8,0))
        self.body_listbox = Listbox(right_col, height=8)
        self.body_listbox.pack(fill=BOTH, padx=8, pady=(0,6), expand=False)

        body_btns = Frame(right_col)
        body_btns.pack(fill=X, padx=8, pady=(0,8))
        Button(body_btns, text="Append Selected Annotation", command=self.append_selected_annotation_to_article).pack(side=LEFT, padx=2)
        Button(body_btns, text="Remove Body Part", command=self.remove_body_part).pack(side=LEFT, padx=2)

        save_frame = Frame(right_col)
        save_frame.pack(fill=X, padx=8, pady=8)
        Button(save_frame, text="Save Article to Word", command=self.save_selected_article_to_word).pack(side=LEFT, padx=4)
        Button(save_frame, text="Save Page Annotations (JSON)", command=self.save_page_annotations_json).pack(side=LEFT, padx=4)

        # Annotation list
        Label(right_col, text="Annotations (page)", font=("Arial", 10, "bold")).pack(anchor="w", padx=8)
        self.ann_listbox = Listbox(right_col, height=6)
        self.ann_listbox.pack(fill=X, padx=8, pady=(0,6))
        ann_btns = Frame(right_col)
        ann_btns.pack(fill=X, padx=8, pady=(0,8))
        Button(ann_btns, text="Delete Annotation", command=self.delete_selected_annotation).pack(side=LEFT, padx=2)
        Button(ann_btns, text="Assign to Article", command=self.assign_annotation_to_selected_article).pack(side=LEFT, padx=2)

        # keep UI responsive
        root.update_idletasks()

    # ----------------- Helpers -----------------
    def set_status(self, msg: str):
        self.status_var.set(msg)
        self.root.update_idletasks()

    def select_field(self, field_name):
        self.selected_field = field_name
        self.set_status(f"Selected label: {field_name}")

    # ----------------- Load PDFs -----------------
    def load_folder(self):
        folder = filedialog.askdirectory(title="Select folder with PDF files")
        if not folder:
            return
        pdf_files = [os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith(".pdf")]
        if not pdf_files:
            messagebox.showinfo("No PDFs", "No PDF files found in folder.")
            return

        # reset
        self.pages.clear()
        self.annotations.clear()
        self.page_articles.clear()
        self.current_page_index = -1
        self.canvas.delete("all")
        self.article_listbox.delete(0, END)
        self.ann_listbox.delete(0, END)
        self.title_entry.delete(0, END)
        self.author_entry.delete(0, END)
        self.date_entry.delete(0, END)
        self.issue_entry.delete(0, END)
        self.body_listbox.delete(0, END)

        self.set_status("Loading PDFs...")
        threading.Thread(target=self._load_pages_thread, args=(pdf_files,), daemon=True).start()

    def _load_pages_thread(self, pdf_files):
        for pdf in pdf_files:
            try:
                pages = convert_from_path(pdf, dpi=200, poppler_path=POPPLER_PATH) if POPPLER_PATH else convert_from_path(pdf, dpi=200)
            except Exception as e:
                print("pdf convert error:", e)
                continue
            for i, pilpage in enumerate(pages):
                idx = len(self.pages)
                self.pages.append({"image": pilpage.convert("RGB"), "pdf": os.path.basename(pdf), "page_no": i+1})
                # initialize page_articles mapping
                self.page_articles[idx] = {}
                # show first page immediately
                if self.current_page_index == -1:
                    self.root.after(0, lambda i=0: self.display_page(i))
                self.root.after(0, lambda msg=f"Loaded {os.path.basename(pdf)} page {i+1}": self.set_status(msg))
        self.root.after(0, lambda: self.set_status("All PDFs loaded."))

    # ----------------- Page display, zoom, pan -----------------
    def display_page(self, page_index: int):
        if page_index < 0 or page_index >= len(self.pages):
            return
        self.current_page_index = page_index
        page = self.pages[page_index]
        pil = page["image"]
        # compute base_scale to fit canvas area
        canvas_w = max(600, int(self.root.winfo_width() * 0.6))
        canvas_h = max(400, int(self.root.winfo_height() * 0.8))
        orig_w, orig_h = pil.size
        scale_w = canvas_w / orig_w
        scale_h = canvas_h / orig_h
        self.base_scale = min(scale_w, scale_h, 1.0)
        self.zoom = 1.0
        self._render_image()
        # refresh article list for this page
        self.refresh_article_listbox()
        self.refresh_annotation_listbox()
        self.load_article_fields_for_current_selection()
        self.set_status(f"Showing {page['pdf']} page {page['page_no']}")

    def _render_image(self):
        page = self.pages[self.current_page_index]
        pil = page["image"]
        scale = self.base_scale * self.zoom
        disp_w = int(pil.width * scale)
        disp_h = int(pil.height * scale)
        resized = pil.resize((disp_w, disp_h), Image.LANCZOS)
        self.current_pil_image = pil
        self.current_image_tk = ImageTk.PhotoImage(resized)
        self.canvas.delete("all")
        self.canvas.config(scrollregion=(0, 0, disp_w, disp_h))
        self.canvas.create_image(0, 0, anchor="nw", image=self.current_image_tk, tags=("page_image",))
        # redraw annotations (scaled)
        self.redraw_annotations_for_page(self.current_page_index)

    def zoom_in(self):
        self.zoom *= 1.25
        self._render_image()

    def zoom_out(self):
        self.zoom /= 1.25
        self._render_image()

    def on_mousewheel(self, event):
        # Zoom while ctrl key not required for simplicity
        if event.delta > 0:
            self.zoom_in()
        else:
            self.zoom_out()

    def start_pan(self, event):
        self.canvas.scan_mark(event.x, event.y)

    def do_pan(self, event):
        self.canvas.scan_dragto(event.x, event.y, gain=1)

    def prev_page(self):
        if self.current_page_index > 0:
            self.display_page(self.current_page_index - 1)

    def next_page(self):
        if self.current_page_index < len(self.pages) - 1:
            self.display_page(self.current_page_index + 1)

    # ----------------- Annotation drawing -----------------
    def canvas_mouse_down(self, event):
        if self.current_page_index == -1:
            return
        self.canvas.focus_set()
        x = self.canvas.canvasx(event.x); y = self.canvas.canvasy(event.y)
        self.draw_start = (x, y)
        self.temp_rect_id = self.canvas.create_rectangle(x, y, x, y, outline="yellow", width=2)

    def canvas_mouse_move(self, event):
        if not self.draw_start or not self.temp_rect_id:
            return
        x0, y0 = self.draw_start
        x1 = self.canvas.canvasx(event.x); y1 = self.canvas.canvasy(event.y)
        self.canvas.coords(self.temp_rect_id, x0, y0, x1, y1)

    def canvas_mouse_up(self, event):
        if not self.draw_start or not self.temp_rect_id:
            return
        x0, y0 = self.draw_start
        x1 = self.canvas.canvasx(event.x); y1 = self.canvas.canvasy(event.y)
        # remove temp rect
        self.canvas.delete(self.temp_rect_id)
        self.temp_rect_id = None
        self.draw_start = None

        # normalize coords & check minimum size
        x1f, x2f = sorted((x0, x1)); y1f, y2f = sorted((y0, y1))
        if abs(x2f - x1f) < 8 or abs(y2f - y1f) < 6:
            self.set_status("Selection too small, ignored.")
            return

        # map display coords back to original image coords
        scale = self.base_scale * self.zoom
        orig_x1 = int(x1f / scale); orig_y1 = int(y1f / scale)
        orig_x2 = int(x2f / scale); orig_y2 = int(y2f / scale)
        # clamp
        pil = self.pages[self.current_page_index]["image"]
        orig_x1 = max(0, min(orig_x1, pil.width - 1))
        orig_x2 = max(1, min(orig_x2, pil.width))
        orig_y1 = max(0, min(orig_y1, pil.height - 1))
        orig_y2 = max(1, min(orig_y2, pil.height))

        label = self.selected_field
        ann_id = str(uuid.uuid4())
        # draw permanent rectangle in display coords
        color = FIELD_COLORS.get(label, "yellow")
        disp_coords = (x1f, y1f, x2f, y2f)
        rect_id = self.canvas.create_rectangle(*disp_coords, outline=color, width=2, tags=("annotation", ann_id))
        text_id = self.canvas.create_text(x1f + 6, y1f + 6, text=label[0], anchor="nw", fill=color, font=("Arial", 10, "bold"), tags=("annotation", ann_id))

        ann = {
            "id": ann_id,
            "page_index": self.current_page_index,
            "bbox_orig": (orig_x1, orig_y1, orig_x2, orig_y2),
            "label": label,
            "ocr_text": "[processing...]",
            "article_title": None,  # will be assigned
            "canvas_items": [rect_id, text_id],
        }
        self.annotations.append(ann)
        self.ann_listbox.insert(END, f"{label} p{ann['page_index']+1}: (processing...)")

        # OCR in background
        self.set_status(f"OCR {label} on page {self.current_page_index+1} ...")
        threading.Thread(target=self._ocr_and_assign, args=(ann,), daemon=True).start()

    def _ocr_and_assign(self, ann):
        try:
            pil = self.pages[ann["page_index"]]["image"]
            x1, y1, x2, y2 = ann["bbox_orig"]
            crop = pil.crop((x1, y1, x2, y2))
            text = pytesseract.image_to_string(crop, config="--psm 6").strip()
            text = text or "[no text]"
        except Exception as e:
            text = f"[ocr error: {e}]"

        # run UI updates in main thread
        def finish():
            ann["ocr_text"] = text
            # update ann_listbox text (find index for this ann on current page)
            self.refresh_annotation_listbox()
            # if Title -> create article and assign automatically
            if ann["label"] == "Article Title":
                title = text.strip()
                if not title:
                    title = f"untitled_{uuid.uuid4().hex[:6]}"
                # create article for page if not exists
                page_map = self.page_articles.setdefault(ann["page_index"], {})
                if title not in page_map:
                    page_map[title] = {"title": title, "author": "", "date": "", "issue": self.pages[ann["page_index"]]["pdf"].replace(".pdf",""), "bodies": [], "annotations": []}
                    # add to listbox if current page
                    if ann["page_index"] == self.current_page_index:
                        self.refresh_article_listbox()
                # assign annotation to that article
                page_map[title]["annotations"].append(ann["id"])
                ann["article_title"] = title
                # select the newly created article in UI
                if ann["page_index"] == self.current_page_index:
                    titles = list(self.page_articles[self.current_page_index].keys())
                    try:
                        idx = titles.index(title)
                        self.article_listbox.selection_clear(0, END)
                        self.article_listbox.selection_set(idx)
                        self.article_listbox.see(idx)
                        self.on_article_select(None)
                    except Exception:
                        pass
                self.set_status(f"Created article '{title}' from Title annotation.")
            else:
                # If an article is selected, auto-assign to that article.
                sel = self.article_listbox.curselection()
                if sel:
                    idx = sel[0]
                    titles = list(self.page_articles.get(self.current_page_index, {}).keys())
                    if titles and idx < len(titles):
                        title = titles[idx]
                        # attach
                        page_map = self.page_articles.setdefault(ann["page_index"], {})
                        art = page_map.setdefault(title, {"title": title, "author": "", "date": "", "issue": self.pages[ann["page_index"]]["pdf"].replace(".pdf",""), "bodies": [], "annotations": []})
                        art["annotations"].append(ann["id"])
                        ann["article_title"] = title
                        # map text to field
                        if ann["label"] == "Author":
                            art["author"] = text
                            if ann["page_index"] == self.current_page_index:
                                self.author_entry.delete(0, END); self.author_entry.insert(0, text)
                        elif ann["label"] == "Date Published":
                            art["date"] = text
                            if ann["page_index"] == self.current_page_index:
                                self.date_entry.delete(0, END); self.date_entry.insert(0, text)
                        elif ann["label"] == "Issue No":
                            art["issue"] = text
                            if ann["page_index"] == self.current_page_index:
                                self.issue_entry.delete(0, END); self.issue_entry.insert(0, text)
                        elif ann["label"] == "Article Body":
                            art["bodies"].append(text)
                            if ann["page_index"] == self.current_page_index:
                                self.body_listbox.insert(END, text)
                        self.set_status(f"Assigned annotation to article '{title}'.")
                        return
                # if no article selected and not title, leave unassigned (user can assign later)
                self.set_status("Annotation OCR done — unassigned. Select article and 'Assign to Article' to attach.")
        self.root.after(0, finish)

    # ----------------- Annotations & Article list UI -----------------
    def refresh_annotation_listbox(self):
        self.ann_listbox.delete(0, END)
        for ann in self.annotations:
            if ann["page_index"] != self.current_page_index:
                continue
            short = ann.get("ocr_text", "").replace("\n", " ")[:80]
            assigned = ann.get("article_title") or ""
            self.ann_listbox.insert(END, f"{ann['label']} p{ann['page_index']+1} {assigned}: {short}")

    def refresh_article_listbox(self):
        self.article_listbox.delete(0, END)
        page_map = self.page_articles.get(self.current_page_index, {})
        for title in page_map.keys():
            self.article_listbox.insert(END, title)

    def on_article_select(self, event):
        sel = self.article_listbox.curselection()
        if not sel:
            self.selected_article_title = None
            self.clear_article_fields()
            return
        idx = sel[0]
        titles = list(self.page_articles.get(self.current_page_index, {}).keys())
        if idx >= len(titles):
            return
        title = titles[idx]
        self.selected_article_title = title
        # fill right side fields from model
        art = self.page_articles[self.current_page_index][title]
        self.title_entry.delete(0, END); self.title_entry.insert(0, art.get("title",""))
        self.author_entry.delete(0, END); self.author_entry.insert(0, art.get("author",""))
        self.date_entry.delete(0, END); self.date_entry.insert(0, art.get("date",""))
        self.issue_entry.delete(0, END); self.issue_entry.insert(0, art.get("issue",""))
        self.body_listbox.delete(0, END)
        for b in art.get("bodies", []):
            self.body_listbox.insert(END, b)

    def clear_article_fields(self):
        self.title_entry.delete(0, END); self.author_entry.delete(0, END)
        self.date_entry.delete(0, END); self.issue_entry.delete(0, END)
        self.body_listbox.delete(0, END)

    def create_blank_article(self):
        if self.current_page_index == -1:
            messagebox.showinfo("No page", "Load a page first.")
            return
        title = f"untitled_{uuid.uuid4().hex[:6]}"
        page_map = self.page_articles.setdefault(self.current_page_index, {})
        page_map[title] = {"title": title, "author": "", "date": "", "issue": self.pages[self.current_page_index]["pdf"].replace(".pdf",""), "bodies": [], "annotations": []}
        self.refresh_article_listbox()
        # select it
        titles = list(page_map.keys())
        idx = titles.index(title)
        self.article_listbox.selection_clear(0, END)
        self.article_listbox.selection_set(idx)
        self.on_article_select(None)
        self.set_status(f"Created blank article '{title}'")

    def delete_selected_article(self):
        sel = self.article_listbox.curselection()
        if not sel:
            messagebox.showinfo("Select article", "Select an article to delete.")
            return
        idx = sel[0]
        titles = list(self.page_articles.get(self.current_page_index, {}).keys())
        if idx >= len(titles):
            return
        title = titles[idx]
        # remove article (annotations remain but are detached)
        art = self.page_articles[self.current_page_index].pop(title, None)
        # detach annotations
        if art:
            for aid in art.get("annotations", []):
                for ann in self.annotations:
                    if ann["id"] == aid:
                        ann["article_title"] = None
        self.refresh_article_listbox()
        self.refresh_annotation_listbox()
        self.clear_article_fields()
        self.set_status(f"Deleted article '{title}'")

    def append_selected_annotation_to_article(self):
        sel_ann = self.ann_listbox.curselection()
        sel_art = self.article_listbox.curselection()
        if not sel_ann or not sel_art:
            messagebox.showinfo("Select both", "Select an annotation and an article to append.")
            return
        ann_idx = sel_ann[0]
        # find nth annotation on page
        # build list of anns for current page
        page_anns = [a for a in self.annotations if a["page_index"] == self.current_page_index]
        if ann_idx >= len(page_anns):
            return
        ann = page_anns[ann_idx]
        art_titles = list(self.page_articles.get(self.current_page_index, {}).keys())
        art = self.page_articles[self.current_page_index][art_titles[sel_art[0]]]
        # attach
        ann["article_title"] = art["title"]
        art.setdefault("annotations", []).append(ann["id"])
        # if body, append text
        if ann["label"] == "Article Body":
            art.setdefault("bodies", []).append(ann.get("ocr_text",""))
            if self.selected_article_title == art["title"]:
                self.body_listbox.insert(END, ann.get("ocr_text",""))
        elif ann["label"] == "Author":
            art["author"] = ann.get("ocr_text","")
            if self.selected_article_title == art["title"]:
                self.author_entry.delete(0, END); self.author_entry.insert(0, art["author"])
        elif ann["label"] == "Date Published":
            art["date"] = ann.get("ocr_text","")
            if self.selected_article_title == art["title"]:
                self.date_entry.delete(0, END); self.date_entry.insert(0, art["date"])
        self.refresh_annotation_listbox()
        self.set_status("Annotation assigned to article.")

    def assign_annotation_to_selected_article(self):
        # same as append_selected_annotation_to_article (kept for button)
        self.append_selected_annotation_to_article()

    def delete_selected_annotation(self):
        sel = self.ann_listbox.curselection()
        if not sel:
            messagebox.showinfo("Select annotation", "Select an annotation to delete.")
            return
        idx = sel[0]
        page_anns = [a for a in self.annotations if a["page_index"] == self.current_page_index]
        if idx >= len(page_anns):
            return
        ann = page_anns[idx]
        # remove canvas items if visible
        for item in ann.get("canvas_items", []):
            try:
                self.canvas.delete(item)
            except Exception:
                pass
        # remove from annotations list
        self.annotations = [a for a in self.annotations if a["id"] != ann["id"]]
        # remove from any article references
        for page_map in self.page_articles.values():
            for art in page_map.values():
                if "annotations" in art and ann["id"] in art["annotations"]:
                    art["annotations"].remove(ann["id"])
                if ann.get("ocr_text") and ann.get("ocr_text") in art.get("bodies", []):
                    art["bodies"] = [b for b in art.get("bodies", []) if b != ann.get("ocr_text")]
        self.refresh_annotation_listbox()
        self.refresh_article_listbox()
        self.set_status("Annotation deleted.")

    def remove_body_part(self):
        sel = self.body_listbox.curselection()
        if not sel or self.selected_article_title is None:
            return
        idx = sel[0]
        art = self.page_articles[self.current_page_index].get(self.selected_article_title)
        if not art:
            return
        if idx < len(art.get("bodies", [])):
            art["bodies"].pop(idx)
        self.body_listbox.delete(idx)
        self.set_status("Removed body part.")

    # ----------------- Article navigation -----------------
    def prev_article(self):
        titles = list(self.page_articles.get(self.current_page_index, {}).keys())
        if not titles:
            return
        sel = self.article_listbox.curselection()
        idx = sel[0] if sel else 0
        new = max(0, idx - 1)
        self.article_listbox.selection_clear(0, END)
        self.article_listbox.selection_set(new)
        self.article_listbox.see(new)
        self.on_article_select(None)

    def next_article(self):
        titles = list(self.page_articles.get(self.current_page_index, {}).keys())
        if not titles:
            return
        sel = self.article_listbox.curselection()
        idx = sel[0] if sel else -1
        new = min(len(titles) - 1, idx + 1)
        self.article_listbox.selection_clear(0, END)
        self.article_listbox.selection_set(new)
        self.article_listbox.see(new)
        self.on_article_select(None)

    # ----------------- Redraw annotations -----------------
    def redraw_annotations_for_page(self, page_index: int):
        # remove existing annotation items then draw annotations for this page scaled to display
        # delete items with tag "annotation"
        for item in self.canvas.find_withtag("annotation"):
            try: self.canvas.delete(item)
            except: pass
        scale = self.base_scale * self.zoom
        for ann in self.annotations:
            if ann["page_index"] != page_index:
                continue
            x1, y1, x2, y2 = ann["bbox_orig"]
            dx1 = x1 * scale; dy1 = y1 * scale; dx2 = x2 * scale; dy2 = y2 * scale
            color = FIELD_COLORS.get(ann["label"], "yellow")
            rect = self.canvas.create_rectangle(dx1, dy1, dx2, dy2, outline=color, width=2, tags=("annotation",))
            txt = self.canvas.create_text(dx1 + 6, dy1 + 6, text=ann["label"][0], anchor="nw", fill=color, font=("Arial", 10, "bold"), tags=("annotation",))
            ann["canvas_items"] = [rect, txt]

    # ----------------- Save / Export -----------------
    def save_page_annotations_json(self):
        if self.current_page_index == -1:
            messagebox.showinfo("No page", "Nothing to save.")
            return
        page = self.pages[self.current_page_index]
        anns = [a for a in self.annotations if a["page_index"] == self.current_page_index]
        out = {
            "pdf": page["pdf"],
            "page_no": page["page_no"],
            "annotations": [{"id": a["id"], "label": a["label"], "bbox_orig": a["bbox_orig"], "ocr_text": a.get("ocr_text",""), "article_title": a.get("article_title")} for a in anns],
            "articles": self.page_articles.get(self.current_page_index, {})
        }
        fname = f"{page['pdf']}_p{page['page_no']}.json"
        path = os.path.join(ANNOT_DIR, fname)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2, ensure_ascii=False)
        self.set_status(f"Annotations & articles saved: {path}")
        messagebox.showinfo("Saved", f"Saved annotations to\n{path}")

    def save_selected_article_to_word(self):
        if self.current_page_index == -1:
            messagebox.showinfo("No page", "No page selected.")
            return
        if self.selected_article_title is None:
            messagebox.showinfo("No article", "Select an article from the list first.")
            return
        art = self.page_articles[self.current_page_index].get(self.selected_article_title)
        if not art:
            messagebox.showinfo("No article", "Article not found.")
            return
        # allow edits from UI to override model
        art["title"] = self.title_entry.get().strip() or art.get("title","")
        art["author"] = self.author_entry.get().strip() or art.get("author","")
        art["date"] = self.date_entry.get().strip() or art.get("date","")
        art["issue"] = self.issue_entry.get().strip() or art.get("issue","")
        art["bodies"] = [self.body_listbox.get(i) for i in range(self.body_listbox.size())] or art.get("bodies",[])

        doc = Document()
        table = doc.add_table(rows=5, cols=2)
        labels = ["Article Title:", "Author:", "Date Published:", "Issue No:", "Article Body:"]
        values = [
            art.get("title",""),
            art.get("author",""),
            art.get("date",""),
            art.get("issue",""),
            "\n\n".join(art.get("bodies",[])) or "[no body]"
        ]
        for i, lab in enumerate(labels):
            table.cell(i,0).text = lab
            table.cell(i,1).text = values[i]

        fname = safe_filename(art.get("title") or f"page{self.current_page_index+1}_{uuid.uuid4().hex[:6]}")
        out = os.path.join(OUTPUT_DIR, f"{fname}.docx")
        doc.save(out)
        self.set_status(f"Saved Word: {out}")
        messagebox.showinfo("Saved", f"Article exported to\n{out}")

# ----------------- Run -----------------
def main():
    root = Tk()
    app = NewspaperOCRApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
