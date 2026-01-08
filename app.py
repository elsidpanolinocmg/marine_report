# app.py
import os
import uuid
import threading
from tkinter import *
from tkinter import filedialog, messagebox
from pdf2image import convert_from_path
from PIL import Image, ImageTk
import pytesseract

import ocr_utils
import storage

# tweak these to your environment
pytesseract_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"  # change if needed
# If you need to set this for pytesseract inside ocr_utils, uncomment:
# import pytesseract
pytesseract.pytesseract.tesseract_cmd = pytesseract_path

OUTPUT_DIR = "Processed_Articles"
SESSION_DIR = "Sessions"
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(SESSION_DIR, exist_ok=True)

class NewspaperOCRApp:
    def __init__(self, root):
        self.root = root
        root.title("Newspaper OCR Annotator (modular)")
        root.geometry("1300x760")

        # state
        self.pdf_path = None
        self.pages = []  # list of PIL images
        self.current_page_index = 0
        self.scale = 1.0

        # annotations: list of dicts: {id, article_id, page_index, bbox_orig(tuple), type, text}
        self.annotations = []

        # articles: dict id -> article data
        # article: {id, title, author, date, issue, bodies(list), annotations(list), manual_modified(bool)}
        self.articles = {}

        self.current_article_id = None

        # UI layout
        left = Frame(root)
        left.pack(side="left", fill="both", expand=True)

        self.canvas = Canvas(left, bg="#333")
        self.canvas.pack(fill="both", expand=True)

        # navigation controls
        nav = Frame(left)
        nav.pack(fill="x")
        Button(nav, text="Open PDF", command=self.open_pdf).pack(side="left", padx=4)
        Button(nav, text="Open Session (JSON)", command=self.open_session).pack(side="left", padx=4)
        Button(nav, text="Prev Page", command=self.prev_page).pack(side="left", padx=4)
        Button(nav, text="Next Page", command=self.next_page).pack(side="left", padx=4)
        self.page_label = Label(nav, text="Page 0/0")
        self.page_label.pack(side="left", padx=8)
        Button(nav, text="Save Session (JSON)", command=self.save_session).pack(side="right", padx=4)
        Button(nav, text="Auto Save Now", command=self.auto_save).pack(side="right", padx=4)

        # canvas events (zoom/pan)
        self.canvas.bind("<MouseWheel>", self.on_mouse_wheel)
        self.canvas.bind("<ButtonPress-1>", self.pan_start)
        self.canvas.bind("<B1-Motion>", self.pan_move)

        # draw/annotate with right mouse button (works with canvas coords)
        self.start_x = self.start_y = None
        self.temp_rect_id = None
        self.canvas.bind("<ButtonPress-3>", self.rect_start)
        self.canvas.bind("<B3-Motion>", self.rect_draw)
        self.canvas.bind("<ButtonRelease-3>", self.rect_release)

        # Right side UI
        right = Frame(root, width=420)
        right.pack(side="right", fill="y")

        Label(right, text="Articles").pack(anchor="w", padx=6, pady=(6,0))
        self.article_listbox = Listbox(right)
        self.article_listbox.pack(fill="x", padx=6)
        self.article_listbox.bind("<<ListboxSelect>>", self.on_article_select)

        btn_frame = Frame(right)
        btn_frame.pack(fill="x", padx=6, pady=4)
        Button(btn_frame, text="New Article", command=self.new_article).pack(side="left")
        Button(btn_frame, text="Delete Article", command=self.delete_article).pack(side="left", padx=6)
        Button(btn_frame, text="Prev", command=self.prev_article).pack(side="left", padx=6)
        Button(btn_frame, text="Next", command=self.next_article).pack(side="left", padx=6)

        Label(right, text="Article Editor").pack(anchor="w", padx=6, pady=(8,0))
        self.editor = Text(right, height=20, wrap="word")
        self.editor.pack(fill="both", expand=True, padx=6, pady=(0,6))

        edit_btns = Frame(right)
        edit_btns.pack(fill="x", padx=6)
        Button(edit_btns, text="Apply Editor -> Article", command=self.apply_editor).pack(side="left")
        Button(edit_btns, text="Save All Word", command=self.save_all_docx).pack(side="right")

        # keep track of autosave path for convenience
        self.autosave_path = os.path.join(SESSION_DIR, "autosave_session.json")

    # ---------------- PDF / session load/save ----------------
    def open_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if not path:
            return
        self.pdf_path = path
        self._load_pdf_pages(path)
        self._reset_state_for_new_pdf()
        self.show_page()

    def _load_pdf_pages(self, pdf_path):
        # convert pages in background thread (UI remains responsive)
        try:
            pages = convert_from_path(pdf_path, dpi=150)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open PDF: {e}")
            return
        self.pages = pages
        self.current_page_index = 0

    def _reset_state_for_new_pdf(self):
        # keep articles/annotations empty for new PDF
        self.annotations.clear()
        self.articles.clear()
        self.current_article_id = None
        self.article_listbox.delete(0, END)
        self.editor.delete("1.0", END)

    def save_session(self):
        if not self.pdf_path:
            messagebox.showinfo("No PDF", "Open a PDF before saving a session.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")], initialdir=SESSION_DIR)
        if not path:
            return
        session = self._compose_session_dict()
        storage.save_session_json(path, session)
        messagebox.showinfo("Saved", f"Session saved to {path}")

    def open_session(self):
        path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")], initialdir=SESSION_DIR)
        if not path:
            return
        session = storage.load_session_json(path)
        # session must include pdf_path - prompt user if missing or not found
        pdf_path = session.get("pdf_path")
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showinfo("PDF not found", "Original PDF path not found. Please select the PDF file to load pages.")
            pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
            if not pdf_path:
                return
        self.pdf_path = pdf_path
        self._load_pdf_pages(pdf_path)
        # restore articles and annotations (keep ids)
        self.articles.clear()
        for aid, art in session.get("articles", {}).items():
            # ensure structure
            art.setdefault("annotations", [])
            # mark manual_modified if present
            art["manual_modified"] = art.get("manual_modified", False)
            self.articles[aid] = art
        self.annotations = session.get("annotations", [])
        # map annotations into article lists (ensure pointers)
        for art in self.articles.values():
            art["annotations"] = [a for a in self.annotations if a.get("article_id") == art.get("id")]
        # populate UI
        self.refresh_article_listbox()
        self.current_page_index = 0
        self.show_page()
        messagebox.showinfo("Loaded", "Session loaded. Check article list and pages.")

    def auto_save(self):
        if not self.pdf_path:
            return
        session = self._compose_session_dict()
        storage.save_session_json(self.autosave_path, session)
        # silent autosave – optional popups can be enabled if desired

    def _compose_session_dict(self):
        # build a JSON-serializable session dict
        return {
            "pdf_path": self.pdf_path,
            "articles": self.articles,
            "annotations": self.annotations
        }

    # ---------------- show page / drawing ----------------
    def show_page(self):
        if not self.pages:
            self.canvas.delete("all")
            self.page_label.config(text="Page 0/0")
            return

        pil = self.pages[self.current_page_index]
        w = max(1, int(pil.width * self.scale))
        h = max(1, int(pil.height * self.scale))
        resized = pil.resize((w, h))
        self.tkpage = ImageTk.PhotoImage(resized)

        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor="nw", image=self.tkpage)

        # draw annotations for this page
        for ann in self.annotations:
            if ann.get("page_index") == self.current_page_index:
                x1, y1, x2, y2 = [c * self.scale for c in ann["bbox_orig"]]
                color = {"Title": "blue", "Author": "orange", "Date": "purple", "Issue": "brown"}.get(ann.get("type"), "green")
                self.canvas.create_rectangle(x1, y1, x2, y2, outline=color, width=2)

        self.page_label.config(text=f"Page {self.current_page_index+1}/{len(self.pages)}")

        # ensure article editor shows current selected article content (if any)
        if self.current_article_id:
            self.load_article_to_editor(self.current_article_id)

    def prev_page(self):
        if self.current_page_index > 0:
            self.current_page_index -= 1
            self.show_page()

    def next_page(self):
        if self.current_page_index < len(self.pages) - 1:
            self.current_page_index += 1
            self.show_page()

    # ---------------- canvas helpers (pan / zoom) ----------------
    def pan_start(self, event):
        self.canvas.scan_mark(event.x, event.y)

    def pan_move(self, event):
        self.canvas.scan_dragto(event.x, event.y, gain=1)

    def on_mouse_wheel(self, event):
        # zoom in/out around center - keep simple
        if event.delta > 0:
            self.scale *= 1.1
        else:
            self.scale /= 1.1
        # clamp scale
        self.scale = max(0.2, min(self.scale, 5.0))
        self.show_page()

    # ---------------- rectangle drawing (right mouse) ----------------
    def rect_start(self, event):
        # use canvas coordinates so the displayed rectangle matches cursor even when scrolled
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        self.temp_rect_id = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline="yellow", width=2)

    def rect_draw(self, event):
        if not self.temp_rect_id:
            return
        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        self.canvas.coords(self.temp_rect_id, self.start_x, self.start_y, cur_x, cur_y)

    def rect_release(self, event):
        if not self.temp_rect_id:
            return
        x2 = self.canvas.canvasx(event.x)
        y2 = self.canvas.canvasy(event.y)
        x1, y1 = self.start_x, self.start_y
        # remove temp rect (we'll redraw when needed)
        try:
            self.canvas.delete(self.temp_rect_id)
        except Exception:
            pass
        self.temp_rect_id = None

        # small selection -> ignore
        if abs(x2 - x1) < 8 or abs(y2 - y1) < 6:
            return

        if not self.current_article_id:
            messagebox.showerror("No article", "Create or select an article first before annotating.")
            return

        # normalize and convert to original image coordinates
        nx1, nx2 = sorted((x1, x2))
        ny1, ny2 = sorted((y1, y2))
        bbox_orig = (nx1 / self.scale, ny1 / self.scale, nx2 / self.scale, ny2 / self.scale)

        # OCR in background to avoid UI blocking
        threading.Thread(target=self._ocr_and_attach, args=(bbox_orig,), daemon=True).start()

    def _ocr_and_attach(self, bbox_orig):
        # get OCR text
        pil = self.pages[self.current_page_index]
        try:
            text = ocr_utils.ocr_image_crop(pil, bbox_orig)
        except Exception as e:
            text = f"[ocr error: {e}]"

        # ask user for annotation type on main thread (must not call dialogs from background thread)
        def ask_type_and_finish():
            top = Toplevel(self.root)
            top.title("Annotation Type")
            Label(top, text="Choose type for this annotation:").pack(padx=8, pady=6)
            var = StringVar(value="Body")
            for t in ["Title", "Author", "Date", "Issue", "Body"]:
                Radiobutton(top, text=t, variable=var, value=t).pack(anchor="w", padx=10)
            def on_ok():
                top.destroy()
            Button(top, text="OK", command=on_ok).pack(pady=6)
            top.grab_set()
            self.root.wait_window(top)
            atype = var.get()

            ann = {
                "id": str(uuid.uuid4()),
                "article_id": self.current_article_id,
                "page_index": self.current_page_index,
                "bbox_orig": bbox_orig,
                "type": atype,
                "text": text
            }
            self.annotations.append(ann)
            # attach annotation to article object (by id)
            art = self.articles.get(self.current_article_id)
            if art is None:
                # shouldn't happen, but safeguard
                return
            art.setdefault("annotations", []).append(ann)

            # Now integrate OCR text into article fields — but **do not overwrite manual edits**.
            # If user manually modified the article (manual_modified True), do not auto-set title/author/date/issue
            manual = art.get("manual_modified", False)
            if atype == "Title":
                if not manual or not art.get("title"):
                    art["title"] = text or art.get("title", "")
            elif atype == "Author":
                if not manual or not art.get("author"):
                    art["author"] = text or art.get("author", "")
            elif atype == "Date":
                if not manual or not art.get("date"):
                    art["date"] = text or art.get("date", "")
            elif atype == "Issue":
                # some newspapers write "Vol 46 No. 1" etc - keep OCR as issue if not manually modified
                if not manual or not art.get("issue"):
                    art["issue"] = text or art.get("issue", "")
            elif atype == "Body":
                # append body segment (always append)
                art.setdefault("bodies", []).append(text)

            # refresh UI
            self.refresh_article_listbox()
            # if current article selected in UI, reload editor
            if self.current_article_id == art.get("id"):
                self.load_article_to_editor(self.current_article_id)
            # autosave
            self.auto_save()
            # redraw page to show created rectangle
            self.show_page()

        # call on main/UI thread
        self.root.after(0, ask_type_and_finish)

    # ---------------- articles management ----------------
    def new_article(self):
        aid = str(uuid.uuid4())
        art = {
            "id": aid,
            "title": f"Article {len(self.articles)+1}",
            "author": "",
            "date": "",
            "issue": "",
            "bodies": [],
            "annotations": [],
            "manual_modified": False
        }
        self.articles[aid] = art
        self.refresh_article_listbox()
        # select it
        index = list(self.articles.keys()).index(aid)
        self.article_listbox.selection_clear(0, END)
        self.article_listbox.selection_set(index)
        self.current_article_id = aid
        self.load_article_to_editor(aid)

    def delete_article(self):
        sel = self.article_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        aid = list(self.articles.keys())[idx]
        # remove article; also remove annotations referencing it
        self.annotations = [a for a in self.annotations if a.get("article_id") != aid]
        del self.articles[aid]
        self.refresh_article_listbox()
        self.editor.delete("1.0", END)
        self.current_article_id = None
        self.show_page()
        self.auto_save()

    def on_article_select(self, event=None):
        sel = self.article_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        aid = list(self.articles.keys())[idx]
        self.current_article_id = aid
        self.load_article_to_editor(aid)

    def prev_article(self):
        sel = self.article_listbox.curselection()
        if not sel:
            # nothing selected -> select first maybe
            if self.article_listbox.size() > 0:
                self.article_listbox.selection_set(0)
                self.on_article_select()
            return
        idx = sel[0]
        if idx > 0:
            self.article_listbox.selection_clear(0, END)
            self.article_listbox.selection_set(idx-1)
            self.on_article_select()

    def next_article(self):
        sel = self.article_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx < self.article_listbox.size() - 1:
            self.article_listbox.selection_clear(0, END)
            self.article_listbox.selection_set(idx+1)
            self.on_article_select()

    def load_article_to_editor(self, aid):
        art = self.articles.get(aid)
        if not art:
            return
        self.editor.delete("1.0", END)
        header = f"Title: {art.get('title','')}\nAuthor: {art.get('author','')}\nDate: {art.get('date','')}\nIssue: {art.get('issue','')}\n\n"
        body = "\n\n".join(art.get("bodies", []))
        self.editor.insert("1.0", header + body)

    def apply_editor(self):
        """
        When user edits the editor text and clicks Apply:
        - parse Title/Author/Date/Issue from top lines if present
        - set manual_modified = True so later OCR won't clobber fields
        """
        if not self.current_article_id:
            return
        content = self.editor.get("1.0", END).rstrip()
        lines = content.splitlines()
        art = self.articles[self.current_article_id]
        # parse header if present
        if lines and lines[0].startswith("Title:"):
            art["title"] = lines[0][6:].strip()
        if len(lines) > 1 and lines[1].startswith("Author:"):
            art["author"] = lines[1][7:].strip()
        if len(lines) > 2 and lines[2].startswith("Date:"):
            art["date"] = lines[2][5:].strip()
        if len(lines) > 3 and lines[3].startswith("Issue:"):
            art["issue"] = lines[3][6:].strip()
        # body is the rest after a blank line (detect first blank after header)
        try:
            blank_idx = lines.index("", 4)
            body_lines = lines[blank_idx+1:]
        except ValueError:
            # no blank - assume body from line 5 onward
            body_lines = lines[4:]

        paragraphs = []
        current_para = []
        for line in body_lines:
            if line.strip() == "":
                if current_para:
                    paragraphs.append(" ".join(current_para).strip())
                    current_para = []
            else:
                current_para.append(line.strip())
        if current_para:
            paragraphs.append(" ".join(current_para).strip())


        art["bodies"] = ocr_utils.clean_paragraphs(body_lines)
        # user explicitly applied edits => mark manual_modified
        art["manual_modified"] = True
        # persist back to list UI
        self.refresh_article_listbox()
        self.auto_save()

    def refresh_article_listbox(self):
        self.article_listbox.delete(0, END)
        for aid, art in self.articles.items():
            title = art.get("title") or f"Article {list(self.articles.keys()).index(aid)+1}"
            self.article_listbox.insert(END, title)

    # ---------------- saving exports ----------------
    def save_all_docx(self):
        if not self.articles:
            messagebox.showinfo("No articles", "No articles to save.")
            return
        for aid, art in self.articles.items():
            safe = storage.safe_filename(art.get("title") or f"article_{aid[:6]}")
            out = os.path.join(OUTPUT_DIR, f"{safe}.docx")
            storage.save_article_docx(out, art)
        messagebox.showinfo("Saved", f"All articles exported to {OUTPUT_DIR}")

    # ---------------- misc ----------------
    def auto_save(self):
        # silent autosave session (overwrites autosave file)
        if not self.pdf_path:
            return
        session = self._compose_session_dict()
        storage.save_session_json(self.autosave_path, session)

def main():
    root = Tk()
    app = NewspaperOCRApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
