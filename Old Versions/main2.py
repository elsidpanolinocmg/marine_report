"""
Simplified Newspaper OCR Annotator

- Left: image canvas (draw colored rectangles -> OCR crops)
- Right: simple fields (Title, Author, Date, Issue, multiple Body parts)
- Save Selected Page Article -> Word file (2-column table)
- Annotations saved to JSON per page
- Zoom/pan + proper scaling for annotations
- OCR runs in background threads

Dependencies: pytesseract, pdf2image, Pillow, python-docx
System: Tesseract OCR program, Poppler for pdf2image
"""

import os
import re
import json
import threading
import uuid
import pytesseract
from tkinter import (
    Tk,
    Frame,
    Canvas,
    Button,
    Label,
    Entry,
    Text,
    Listbox,
    Scrollbar,
    END,
    StringVar,
    filedialog,
    messagebox,
)
from pdf2image import convert_from_path
from PIL import Image, ImageTk
from docx import Document

# ---------- CONFIG ----------
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"  # adjust if needed
POPPLER_PATH = None  # set to poppler bin path or leave None if poppler in PATH

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

# ---------- UTILS ----------
def safe_filename(s):
    return re.sub(r'[\\/*?:"<>|]', "", s).strip()[:150] or "untitled"

# ---------- APP ----------
class NewspaperOCRApp:
    def __init__(self, root):
        self.root = root
        root.title("Newspaper OCR — Simplified")
        root.geometry("1200x800")

        # Data
        self.pages = []  # list of dicts: {"image": PIL.Image, "pdf": name, "page_no": int}
        self.current_page_index = -1
        self.current_pil_image = None
        self.current_image_tk = None
        self.base_scale = 1.0   # fit scale
        self.zoom = 1.0         # zoom multiplier
        # annotations: list of {id, bbox_orig:(x1,y1,x2,y2), label, ocr_text, page_index, canvas_ids}
        self.annotations = []
        # per-page simple article model: dict page_index -> article dict
        # article dict: {title, author, date, issue, bodies: [str], pdf_name}
        self.page_articles = {}

        # Active field selection (radio-like)
        self.selected_field = "Article Body"

        # --- UI Layout ---
        main = Frame(root)
        main.pack(fill="both", expand=True)

        # left: canvas
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

        # Bind mouse events for drawing and pan/zoom
        self.canvas.bind("<ButtonPress-1>", self.canvas_mouse_down)
        self.canvas.bind("<B1-Motion>", self.canvas_mouse_move)
        self.canvas.bind("<ButtonRelease-1>", self.canvas_mouse_up)
        # pan via middle button or right button
        self.canvas.bind("<ButtonPress-2>", self.start_pan)
        self.canvas.bind("<B2-Motion>", self.do_pan)
        self.canvas.bind("<ButtonPress-3>", self.start_pan)
        self.canvas.bind("<B3-Motion>", self.do_pan)
        # zoom with ctrl+mousewheel
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)

        # Tools frame above canvas
        tool = Frame(left)
        tool.pack(fill="x")
        Button(tool, text="Load Folder", command=self.load_folder).pack(side="left", padx=4, pady=4)
        Button(tool, text="Prev Page", command=self.prev_page).pack(side="left", padx=4)
        Button(tool, text="Next Page", command=self.next_page).pack(side="left", padx=4)
        Label(tool, text=" | Select field: ").pack(side="left")
        for fname in FIELD_NAMES:
            b = Button(tool, text=fname, bg=FIELD_COLORS[fname],
                       command=lambda f=fname: self.select_field(f))
            b.pack(side="left", padx=2)

        # status bar
        self.status_var = StringVar(value="Ready")
        Label(tool, textvariable=self.status_var).pack(side="right", padx=6)

        # right: editor
        right = Frame(main, width=380)
        right.pack(side="right", fill="y")

        Label(right, text="Article Title:").pack(anchor="w", padx=6, pady=(8,0))
        self.title_entry = Entry(right)
        self.title_entry.pack(fill="x", padx=6)

        Label(right, text="Author:").pack(anchor="w", padx=6, pady=(8,0))
        self.author_entry = Entry(right)
        self.author_entry.pack(fill="x", padx=6)

        Label(right, text="Date Published:").pack(anchor="w", padx=6, pady=(8,0))
        self.date_entry = Entry(right)
        self.date_entry.pack(fill="x", padx=6)

        Label(right, text="Issue No:").pack(anchor="w", padx=6, pady=(8,0))
        self.issue_entry = Entry(right)
        self.issue_entry.pack(fill="x", padx=6)

        Label(right, text="Article Body Parts (multiple):").pack(anchor="w", padx=6, pady=(8,0))
        self.body_list = Listbox(right, height=10)
        self.body_list.pack(fill="both", padx=6, pady=(0,6), expand=False)

        btn_frame = Frame(right)
        btn_frame.pack(fill="x", padx=6, pady=(0,8))
        Button(btn_frame, text="Remove Selected Body", command=self.remove_body_part).pack(side="left", padx=4)
        Button(btn_frame, text="Save Article (Word)", command=self.save_article_word).pack(side="right", padx=4)
        Button(btn_frame, text="Save Annotations (JSON)", command=self.save_annotations_json).pack(side="right", padx=4)

        # make sure the canvas has a minimum size
        root.update_idletasks()

        # drawing state
        self.draw_start = None
        self.temp_rect = None

    # --------------- UI helpers ---------------
    def set_status(self, msg):
        self.status_var.set(msg)
        self.root.update_idletasks()

    def select_field(self, field_name):
        self.selected_field = field_name
        self.set_status(f"Selected field: {field_name}")

    # --------------- Loading ---------------
    def load_folder(self):
        folder = filedialog.askdirectory(title="Select folder with PDF files")
        if not folder:
            return
        pdfs = [os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith(".pdf")]
        if not pdfs:
            messagebox.showinfo("No PDFs", "No PDF files found in folder.")
            return

        # Clear previous
        self.pages.clear()
        self.annotations.clear()
        self.page_articles.clear()
        self.current_page_index = -1
        self.canvas.delete("all")
        self.body_list.delete(0, END)
        self.title_entry.delete(0, END)
        self.author_entry.delete(0, END)
        self.date_entry.delete(0, END)
        self.issue_entry.delete(0, END)

        # Load first PDF pages progressively in background to avoid freezing
        self.set_status("Loading pages...")
        thread = threading.Thread(target=self._load_pages_thread, args=(pdfs,), daemon=True)
        thread.start()

    def _load_pages_thread(self, pdf_paths):
        for pdf_path in pdf_paths:
            name = os.path.basename(pdf_path)
            try:
                pages = convert_from_path(pdf_path, dpi=200, poppler_path=POPPLER_PATH) if POPPLER_PATH else convert_from_path(pdf_path, dpi=200)
            except Exception as e:
                print("convert_from_path error:", e)
                continue
            for i, page in enumerate(pages):
                self.pages.append({"image": page.convert("RGB"), "pdf": name, "page_no": i+1})
                # set up empty article model for that page
                idx = len(self.pages)-1
                self.page_articles[idx] = {"title": "", "author": "", "date": "", "issue": name.replace(".pdf",""), "bodies": []}
                # show first page asap
                if self.current_page_index == -1:
                    self.root.after(0, lambda: self.display_page(0))
                self.root.after(0, lambda msg=f"Loaded {name} page {i+1}": self.set_status(msg))
        self.root.after(0, lambda: self.set_status("All pages loaded."))

    # --------------- Page display, zoom, pan ---------------
    def display_page(self, page_index):
        if page_index < 0 or page_index >= len(self.pages):
            return
        self.current_page_index = page_index
        page = self.pages[page_index]
        pil = page["image"]
        # compute base scale so the image fits in the canvas area
        canvas_w = max(600, int(self.root.winfo_width() * 0.55))
        canvas_h = max(400, int(self.root.winfo_height() * 0.75))
        orig_w, orig_h = pil.size
        scale_w = canvas_w / orig_w
        scale_h = canvas_h / orig_h
        self.base_scale = min(scale_w, scale_h, 1.0)
        self.zoom = 1.0
        self._render_image()
        # load article fields for this page
        article = self.page_articles.get(page_index, {"title":"", "author":"", "date":"", "issue":"", "bodies":[]})
        self.title_entry.delete(0, END); self.title_entry.insert(0, article.get("title",""))
        self.author_entry.delete(0, END); self.author_entry.insert(0, article.get("author",""))
        self.date_entry.delete(0, END); self.date_entry.insert(0, article.get("date",""))
        self.issue_entry.delete(0, END); self.issue_entry.insert(0, article.get("issue",""))
        self.body_list.delete(0, END)
        for b in article.get("bodies",[]):
            self.body_list.insert(END, b)
        self.set_status(f"Showing page {page['pdf']} (page {page['page_no']})")

    def _render_image(self):
        page = self.pages[self.current_page_index]
        pil = page["image"]
        scale = self.base_scale * self.zoom
        disp_w = int(pil.width * scale)
        disp_h = int(pil.height * scale)
        resized = pil.resize((disp_w, disp_h), Image.LANCZOS)
        self.current_pil_image = pil  # original
        self.current_image_tk = ImageTk.PhotoImage(resized)
        self.canvas.delete("all")
        self.canvas.config(scrollregion=(0,0,disp_w,disp_h))
        self.canvas.create_image(0,0,anchor="nw",image=self.current_image_tk, tags=("page_image",))
        # redraw annotations for this page
        self.redraw_annotations_for_page(self.current_page_index)

    def zoom_in(self):
        self.zoom *= 1.25
        self._render_image()

    def zoom_out(self):
        self.zoom /= 1.25
        self._render_image()

    def on_mousewheel(self, event):
        # ctrl+wheel is conventional; here we always zoom on wheel for simplicity
        if event.delta > 0:
            self.zoom_in()
        else:
            self.zoom_out()

    # pan support
    def start_pan(self, event):
        self.canvas.scan_mark(event.x, event.y)

    def do_pan(self, event):
        self.canvas.scan_dragto(event.x, event.y, gain=1)

    # next / prev
    def next_page(self):
        if self.current_page_index < len(self.pages)-1:
            self.display_page(self.current_page_index+1)

    def prev_page(self):
        if self.current_page_index > 0:
            self.display_page(self.current_page_index-1)

    # --------------- Annotation drawing and OCR ---------------
    def canvas_mouse_down(self, event):
        if self.current_page_index == -1:
            return
        self.canvas.focus_set()
        x = self.canvas.canvasx(event.x); y = self.canvas.canvasy(event.y)
        self.draw_start = (x,y)
        self.temp_rect = self.canvas.create_rectangle(x,y,x,y, outline="yellow", width=2)

    def canvas_mouse_move(self, event):
        if not self.draw_start or not self.temp_rect:
            return
        x0,y0 = self.draw_start
        x1 = self.canvas.canvasx(event.x); y1 = self.canvas.canvasy(event.y)
        self.canvas.coords(self.temp_rect, x0,y0,x1,y1)

    def canvas_mouse_up(self, event):
        if not self.draw_start or not self.temp_rect:
            return
        x0,y0 = self.draw_start
        x1 = self.canvas.canvasx(event.x); y1 = self.canvas.canvasy(event.y)
        self.canvas.delete(self.temp_rect)
        self.temp_rect = None
        self.draw_start = None

        # normalize and small-check
        x1f, x2f = sorted((x0,x1))
        y1f, y2f = sorted((y0,y1))
        if abs(x2f-x1f) < 8 or abs(y2f-y1f) < 6:
            self.set_status("Selection too small, ignored.")
            return

        # convert displayed coords -> original image coords
        # displayed coordinate = orig_coord * base_scale * zoom
        scale = self.base_scale * self.zoom
        orig_x1 = int(x1f / scale)
        orig_y1 = int(y1f / scale)
        orig_x2 = int(x2f / scale)
        orig_y2 = int(y2f / scale)
        # clamp to image
        pil = self.pages[self.current_page_index]["image"]
        orig_x1 = max(0, min(orig_x1, pil.width-1))
        orig_x2 = max(1, min(orig_x2, pil.width))
        orig_y1 = max(0, min(orig_y1, pil.height-1))
        orig_y2 = max(1, min(orig_y2, pil.height))

        label = self.selected_field
        ann_id = str(uuid.uuid4())
        # draw a permanent rectangle in display coords
        color = FIELD_COLORS.get(label, "yellow")
        disp_coords = (x1f, y1f, x2f, y2f)
        canvas_rect = self.canvas.create_rectangle(*disp_coords, outline=color, width=2, tags=("annotation", ann_id))
        # small label letter
        self.canvas.create_text(x1f+6, y1f+6, text=label[0], anchor="nw", fill=color, font=("Arial",10,"bold"), tags=("annotation", ann_id))

        ann = {
            "id": ann_id,
            "bbox_orig": (orig_x1, orig_y1, orig_x2, orig_y2),
            "label": label,
            "ocr_text": "[processing...]",
            "page_index": self.current_page_index,
            "canvas_ids": [canvas_rect],
        }
        self.annotations.append(ann)

        # show placeholder in UI and run OCR in background
        self.set_status(f"OCR {label} on page {self.current_page_index+1} ...")
        threading.Thread(target=self._ocr_and_assign, args=(ann,), daemon=True).start()

    def _ocr_and_assign(self, ann):
        try:
            pil = self.pages[ann["page_index"]]["image"]
            x1,y1,x2,y2 = ann["bbox_orig"]
            crop = pil.crop((x1,y1,x2,y2))
            text = pytesseract.image_to_string(crop, config="--psm 6")
            text = text.strip() or "[no text]"
        except Exception as e:
            text = f"[ocr error: {e}]"

        # update on main thread
        def finish():
            ann["ocr_text"] = text
            # assign directly to page_article fields (simplified)
            art = self.page_articles.setdefault(ann["page_index"], {"title":"", "author":"", "date":"", "issue": self.pages[ann["page_index"]]["pdf"].replace(".pdf",""), "bodies":[]})
            label = ann["label"]
            if label == "Article Title":
                art["title"] = text
                self.title_entry.delete(0, END); self.title_entry.insert(0, text)
            elif label == "Author":
                art["author"] = text
                self.author_entry.delete(0, END); self.author_entry.insert(0, text)
            elif label == "Date Published":
                art["date"] = text
                self.date_entry.delete(0, END); self.date_entry.insert(0, text)
            elif label == "Issue No":
                art["issue"] = text
                self.issue_entry.delete(0, END); self.issue_entry.insert(0, text)
            elif label == "Article Body":
                # append body part
                art["bodies"].append(text)
                self.body_list.insert(END, text)
            self.set_status(f"OCR done for {label}")
        self.root.after(0, finish)

    # redraw annotations for current page (used after zoom / display)
    def redraw_annotations_for_page(self, page_index):
        # remove existing annotation items, then draw those for this page
        # note: we won't try to reuse old canvas_ids because sizes change; we'll just redraw
        self.canvas.delete("annotation")
        scale = self.base_scale * self.zoom
        for ann in self.annotations:
            if ann["page_index"] != page_index:
                continue
            x1,y1,x2,y2 = ann["bbox_orig"]
            # convert to display coords
            dx1 = x1 * scale; dy1 = y1 * scale; dx2 = x2 * scale; dy2 = y2 * scale
            color = FIELD_COLORS.get(ann["label"], "yellow")
            cid = self.canvas.create_rectangle(dx1,dy1,dx2,dy2, outline=color, width=2, tags=("annotation",))
            self.canvas.create_text(dx1+6, dy1+6, text=ann["label"][0], anchor="nw", fill=color, font=("Arial",10,"bold"), tags=("annotation",))
            # store last canvas id (for possible future deletion)
            ann["canvas_ids"] = [cid]

    # --------------- Save / Export ---------------
    def save_annotations_json(self):
        if self.current_page_index == -1:
            messagebox.showinfo("No page", "Nothing to save.")
            return
        page = self.pages[self.current_page_index]
        page_idx = self.current_page_index
        anns = [a for a in self.annotations if a["page_index"] == page_idx]
        out = {
            "pdf": page["pdf"],
            "page_no": page["page_no"],
            "annotations": [{"id": a["id"], "label": a["label"], "bbox_orig": a["bbox_orig"], "ocr_text": a.get("ocr_text","")} for a in anns]
        }
        fname = f"{page['pdf']}_p{page['page_no']}.json"
        path = os.path.join(ANNOT_DIR, fname)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2, ensure_ascii=False)
        self.set_status(f"Annotations saved: {path}")
        messagebox.showinfo("Saved", f"Annotations saved to\n{path}")

    def save_article_word(self):
        if self.current_page_index == -1:
            messagebox.showinfo("No page", "No page selected.")
            return

        # take fields from UI (user might have edited)
        art = self.page_articles.setdefault(
            self.current_page_index,
            {"title": "", "author": "", "date": "",
            "issue": self.pages[self.current_page_index]["pdf"].replace(".pdf", ""),
            "bodies": []}
        )
        art["title"] = self.title_entry.get().strip() or art.get("title", "")
        art["author"] = self.author_entry.get().strip() or art.get("author", "")
        art["date"] = self.date_entry.get().strip() or art.get("date", "")
        art["issue"] = self.issue_entry.get().strip() or art.get("issue", "")
        # bodies come from listbox
        art["bodies"] = [self.body_list.get(i) for i in range(self.body_list.size())] or art.get("bodies", [])

        # build docx
        doc = Document()
        table = doc.add_table(rows=5, cols=2)
        labels = ["Article Title:", "Author:", "Date Published:", "Issue No:", "Article Body:"]
        values = [
            art.get("title", ""),
            art.get("author", ""),
            art.get("date", ""),
            art.get("issue", ""),
            "\n\n".join(art.get("bodies", [])) or "[no body]"
        ]
        for i, lab in enumerate(labels):
            table.cell(i, 0).text = lab
            table.cell(i, 1).text = values[i]

        # use safe_filename
        fname = safe_filename(art.get("title") or f"page{self.current_page_index+1}_{uuid.uuid4().hex[:6]}")
        out = os.path.join(OUTPUT_DIR, f"{fname}.docx")

        doc.save(out)
        self.set_status(f"Saved Word: {out}")
        messagebox.showinfo("Saved", f"Article exported to\n{out}")


    def remove_body_part(self):
        sel = self.body_list.curselection()
        if not sel:
            return
        idx = sel[0]
        self.body_list.delete(idx)
        # update page_articles structure
        art = self.page_articles.get(self.current_page_index)
        if art:
            if idx < len(art["bodies"]):
                art["bodies"].pop(idx)
        self.set_status("Removed body part.")

# ---------- run ----------
def main():
    root = Tk()
    app = NewspaperOCRApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
