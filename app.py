# app.py
import os
import uuid
import threading
from pathlib import Path
from tkinter import *
from tkinter import filedialog, messagebox
from pdf2image import convert_from_path
from PIL import Image, ImageTk
import pytesseract
import re

import ocr_utils
import storage

# ML model
from ultralytics import YOLO

# tweak these to your environment
pytesseract_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
pytesseract.pytesseract.tesseract_cmd = pytesseract_path

OUTPUT_DIR = "Processed_Articles"
SESSION_DIR = "Sessions"
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(SESSION_DIR, exist_ok=True)


class NewspaperOCRApp:
    def __init__(self, root):
        self.root = root
        root.title("Newspaper OCR Annotator (ML Integrated)")
        root.geometry("1400x760")

        # state
        self.pdf_path = None
        self.pages = []
        self.current_page_index = 0
        self.scale = 1.0

        self.annotations = []
        self.articles = {}
        self.current_article_id = None

        # ML model lazy load
        self.model_path = "models/best.pt"  # replace with your YOLOv8 trained model
        self.yolo_model = None

        # ------------------- UI -------------------
        left = Frame(root)
        left.pack(side="left", fill="both", expand=True)

        self.canvas = Canvas(left, bg="#333")
        self.canvas.pack(fill="both", expand=True)

        nav = Frame(left)
        nav.pack(fill="x")
        Button(nav, text="Open PDF", command=self.open_pdf).pack(side="left", padx=4)
        Button(nav, text="Open Session (JSON)", command=self.open_session).pack(
            side="left", padx=4
        )
        Button(nav, text="Prev Page", command=self.prev_page).pack(side="left", padx=4)
        Button(nav, text="Next Page", command=self.next_page).pack(side="left", padx=4)
        self.page_label = Label(nav, text="Page 0/0")
        self.page_label.pack(side="left", padx=8)
        Button(nav, text="Save Session (JSON)", command=self.save_session).pack(
            side="right", padx=4
        )
        Button(nav, text="Auto Save Now", command=self.auto_save).pack(
            side="right", padx=4
        )
        Button(
            nav,
            text="Show Selected Article Only",
            command=self.show_selected_boxes_only,
        ).pack(side="right", padx=4)
        Button(
            nav, text="Auto Detect All Pages", command=self.auto_detect_all_pages
        ).pack(side="left", padx=4)
        Button(
            nav, text="Auto Detect Page", command=self.auto_detect_current_page
        ).pack(side="left", padx=4)

        # Canvas events
        self.canvas.bind("<MouseWheel>", self.on_mouse_wheel)
        self.canvas.bind("<ButtonPress-1>", self.pan_start)
        self.canvas.bind("<B1-Motion>", self.pan_move)
        self.start_x = self.start_y = None
        self.temp_rect_id = None
        self.canvas.bind("<ButtonPress-3>", self.rect_start)
        self.canvas.bind("<B3-Motion>", self.rect_draw)
        self.canvas.bind("<ButtonRelease-3>", self.rect_release)

        # Right side
        right = Frame(root, width=420)
        right.pack(side="right", fill="y")
        Label(right, text="Articles").pack(anchor="w", padx=6, pady=(6, 0))
        self.article_listbox = Listbox(right)
        self.article_listbox.pack(fill="x", padx=6)
        self.article_listbox.bind("<<ListboxSelect>>", self.on_article_select)

        btn_frame = Frame(right)
        btn_frame.pack(fill="x", padx=6, pady=4)
        Button(btn_frame, text="New Article", command=self.new_article).pack(
            side="left"
        )
        Button(btn_frame, text="Delete Article", command=self.delete_article).pack(
            side="left", padx=6
        )
        Button(btn_frame, text="Prev", command=self.prev_article).pack(
            side="left", padx=6
        )
        Button(btn_frame, text="Next", command=self.next_article).pack(
            side="left", padx=6
        )

        Label(right, text="Article Editor").pack(anchor="w", padx=6, pady=(8, 0))
        self.editor = Text(right, height=20, wrap="word")
        self.editor.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        edit_btns = Frame(right)
        edit_btns.pack(fill="x", padx=6)
        Button(
            edit_btns, text="Apply Editor -> Article", command=self.apply_editor
        ).pack(side="left")
        Button(edit_btns, text="Save All Word/JSON", command=self.save_all_docx).pack(
            side="right"
        )

        self.autosave_path = os.path.join(SESSION_DIR, "autosave_session.json")

    # ---------------- PDF / session ----------------
    def open_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if not path:
            return
        self.pdf_path = path
        self._load_pdf_pages(path)
        self._reset_state_for_new_pdf()
        self.show_page()

    def _load_pdf_pages(self, pdf_path):
        try:
            self.pages = convert_from_path(pdf_path, dpi=150)
            self.current_page_index = 0
        except Exception as e:
            messagebox.showerror("Error", f"Could not open PDF: {e}")

    def _reset_state_for_new_pdf(self):
        self.annotations.clear()
        self.articles.clear()
        self.current_article_id = None
        self.article_listbox.delete(0, END)
        self.editor.delete("1.0", END)

    def save_session(self):
        if not self.pdf_path:
            messagebox.showinfo("No PDF", "Open a PDF first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            initialdir=SESSION_DIR,
        )
        if not path:
            return
        storage.save_session_json(path, self._compose_session_dict())
        messagebox.showinfo("Saved", f"Session saved to {path}")

    def open_session(self):
        path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json")], initialdir=SESSION_DIR
        )
        if not path:
            return
        session = storage.load_session_json(path)
        pdf_path = session.get("pdf_path")
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showinfo("PDF not found", "Please select original PDF.")
            pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
            if not pdf_path:
                return
        self.pdf_path = pdf_path
        self._load_pdf_pages(pdf_path)
        self.articles.clear()
        for aid, art in session.get("articles", {}).items():
            art.setdefault("annotations", [])
            art["manual_modified"] = art.get("manual_modified", False)
            self.articles[aid] = art
        self.annotations = session.get("annotations", [])
        for art in self.articles.values():
            art["annotations"] = [
                a for a in self.annotations if a.get("article_id") == art.get("id")
            ]
        self.refresh_article_listbox()
        self.current_page_index = 0
        self.show_page()
        messagebox.showinfo("Loaded", "Session loaded.")

    def auto_save(self):
        if not self.pdf_path:
            return
        storage.save_session_json(self.autosave_path, self._compose_session_dict())

    def _compose_session_dict(self):
        return {
            "pdf_path": self.pdf_path,
            "articles": self.articles,
            "annotations": self.annotations,
        }

    # ---------------- show page ----------------
    def show_page(self, only_selected_article=False):
        if not self.pages:
            self.canvas.delete("all")
            self.page_label.config(text="Page 0/0")
            return
        pil = self.pages[self.current_page_index]
        w, h = int(pil.width * self.scale), int(pil.height * self.scale)
        resized = pil.resize((w, h))
        self.tkpage = ImageTk.PhotoImage(resized)
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor="nw", image=self.tkpage)

        for ann in self.annotations:
            if ann["page_index"] != self.current_page_index:
                continue
            if (
                only_selected_article
                and ann.get("article_id") != self.current_article_id
            ):
                continue
            x1, y1, x2, y2 = [c * self.scale for c in ann["bbox_orig"]]
            color = {
                "Title": "blue",
                "Author": "orange",
                "Date": "purple",
                "Issue": "brown",
                "Body": "green",
                "Image": "red",
            }.get(ann.get("type"), "yellow")
            self.canvas.create_rectangle(x1, y1, x2, y2, outline=color, width=2)

        self.page_label.config(
            text=f"Page {self.current_page_index+1}/{len(self.pages)}"
        )
        if self.current_article_id:
            self.load_article_to_editor(self.current_article_id)

    def show_selected_boxes_only(self):
        self.show_page(only_selected_article=True)

    def prev_page(self):
        if self.current_page_index > 0:
            self.current_page_index -= 1
            self.show_page()

    def next_page(self):
        if self.current_page_index < len(self.pages) - 1:
            self.current_page_index += 1
            self.show_page()

    # ---------------- canvas helpers ----------------
    def pan_start(self, event):
        self.canvas.scan_mark(event.x, event.y)

    def pan_move(self, event):
        self.canvas.scan_dragto(event.x, event.y, gain=1)

    def on_mouse_wheel(self, event):
        self.scale = max(
            0.2, min(self.scale * (1.1 if event.delta > 0 else 1 / 1.1), 5.0)
        )
        self.show_page()

    # ---------------- rectangle drawing ----------------
    def rect_start(self, event):
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        self.temp_rect_id = self.canvas.create_rectangle(
            self.start_x,
            self.start_y,
            self.start_x,
            self.start_y,
            outline="yellow",
            width=2,
        )

    def rect_draw(self, event):
        if not self.temp_rect_id:
            return
        cur_x, cur_y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        self.canvas.coords(self.temp_rect_id, self.start_x, self.start_y, cur_x, cur_y)

    def rect_release(self, event):
        if not self.temp_rect_id:
            return
        x2, y2 = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        x1, y1 = self.start_x, self.start_y
        try:
            self.canvas.delete(self.temp_rect_id)
        except:
            pass
        self.temp_rect_id = None
        if abs(x2 - x1) < 8 or abs(y2 - y1) < 6:
            return
        if not self.current_article_id:
            messagebox.showerror("No article", "Create/select an article first.")
            return
        bbox_orig = (
            min(x1, x2) / self.scale,
            min(y1, y2) / self.scale,
            max(x1, x2) / self.scale,
            max(y1, y2) / self.scale,
        )
        threading.Thread(
            target=self._ocr_and_attach, args=(bbox_orig,), daemon=True
        ).start()

    def _ocr_and_attach(self, bbox_orig):
        pil = self.pages[self.current_page_index]
        try:
            text = ocr_utils.ocr_image_crop(pil, bbox_orig)
        except Exception as e:
            text = f"[ocr error: {e}]"

        def ask_type_and_finish():
            top = Toplevel(self.root)
            top.title("Annotation Type")
            Label(top, text="Choose type:").pack(padx=8, pady=6)
            var = StringVar(value="Body")
            for t in ["Title", "Author", "Date", "Issue", "Body", "Image"]:
                Radiobutton(top, text=t, variable=var, value=t).pack(
                    anchor="w", padx=10
                )
            Button(top, text="OK", command=top.destroy).pack(pady=6)
            top.grab_set()
            self.root.wait_window(top)
            atype = var.get()
            ann = {
                "id": str(uuid.uuid4()),
                "article_id": self.current_article_id,
                "page_index": self.current_page_index,
                "bbox_orig": bbox_orig,
                "type": atype,
                "text": text,
            }
            art = self.articles.get(self.current_article_id)
            if not art:
                return
            if atype == "Image":
                img_dir = Path(OUTPUT_DIR) / "images"
                img_dir.mkdir(parents=True, exist_ok=True)
                crop = pil.crop(bbox_orig)
                img_path = img_dir / f"{art['id']}_{uuid.uuid4().hex}.png"
                crop.save(img_path)
                ann["path"] = str(img_path)
                art.setdefault("images", []).append(
                    {
                        "path": str(img_path),
                        "bbox": bbox_orig,
                        "page_index": self.current_page_index,
                    }
                )
            else:
                art.setdefault("bodies", []).append(text)
                if atype == "Title":
                    art["title"] = text
                elif atype == "Author" and not art.get("author"):
                    art["author"] = text
                elif atype == "Date" and not art.get("date"):
                    art["date"] = text
                elif atype == "Issue" and not art.get("issue"):
                    art["issue"] = text
                elif atype == "Body":
                    art.setdefault("bodies", []).append(text)
            self.annotations.append(ann)
            art.setdefault("annotations", []).append(ann)
            self.refresh_article_listbox()
            self.load_article_to_editor(self.current_article_id)
            self.auto_save()
            self.show_page()

        self.root.after(0, ask_type_and_finish)

    # ---------------- Articles ----------------
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
            "manual_modified": False,
        }
        self.articles[aid] = art
        self.refresh_article_listbox()
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
        if not sel and self.article_listbox.size() > 0:
            self.article_listbox.selection_set(0)
            self.on_article_select()
            return
        idx = sel[0]
        if idx > 0:
            self.article_listbox.selection_clear(0, END)
            self.article_listbox.selection_set(idx - 1)
            self.on_article_select()

    def next_article(self):
        sel = self.article_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx < self.article_listbox.size() - 1:
            self.article_listbox.selection_clear(0, END)
            self.article_listbox.selection_set(idx + 1)
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
        if not self.current_article_id:
            return
        content = self.editor.get("1.0", END).rstrip()
        lines = content.splitlines()
        art = self.articles[self.current_article_id]
        if lines and lines[0].startswith("Title:"):
            art["title"] = lines[0][6:].strip()
        if len(lines) > 1 and lines[1].startswith("Author:"):
            art["author"] = lines[1][7:].strip()
        if len(lines) > 2 and lines[2].startswith("Date:"):
            art["date"] = lines[2][5:].strip()
        if len(lines) > 3 and lines[3].startswith("Issue:"):
            art["issue"] = lines[3][6:].strip()
        try:
            blank_idx = lines.index("", 4)
            body_lines = lines[blank_idx + 1 :]
        except ValueError:
            body_lines = lines[4:]
        art["bodies"] = ocr_utils.clean_paragraphs(body_lines)
        art["manual_modified"] = True
        self.refresh_article_listbox()
        self.auto_save()

    def refresh_article_listbox(self):
        self.article_listbox.delete(0, END)
        for aid, art in self.articles.items():
            title = (
                art.get("title") or f"Article {list(self.articles.keys()).index(aid)+1}"
            )
            self.article_listbox.insert(END, title)

    def save_all_docx(self):
        if not self.articles:
            messagebox.showinfo("No articles", "Nothing to save.")
            return
        for aid, art in self.articles.items():
            safe = storage.safe_filename(art.get("title", "untitled"))
            out = os.path.join(OUTPUT_DIR, f"{safe}.docx")
            storage.save_article_docx(out, art)
        messagebox.showinfo("Saved", f"All articles saved to {OUTPUT_DIR}")

    # ---------------- ML detection ----------------
    def auto_detect_all_pages(self):
        if not self.pdf_path:
            return
        if self.yolo_model is None:
            try:
                self.yolo_model = YOLO(self.model_path)
            except Exception as e:
                messagebox.showerror("YOLO Error", f"Could not load YOLO model: {e}")
                return
        threading.Thread(target=self._detect_all_pages_thread, daemon=True).start()

    def _detect_all_pages_thread(self):
        for page_idx, pil in enumerate(self.pages):
            temp_path = os.path.join(SESSION_DIR, f"page_{page_idx}.png")
            pil.save(temp_path)

            results = self.yolo_model(temp_path, conf=0.05, iou=0.5)

            for box, cls_id in zip(results[0].boxes.xyxy, results[0].boxes.cls):
                x1, y1, x2, y2 = map(float, box)

                class_map = {0: "Title", 1: "Body", 2: "Image"}
                atype = class_map.get(int(cls_id), "Body")
                bbox_orig = (x1, y1, x2, y2)

                # 🔹 Create a NEW article whenever a Title is detected
                if atype == "Title":
                    self.new_article()

                if not self.current_article_id:
                    self.new_article()

                art = self.articles[self.current_article_id]

                ann = {
                    "id": str(uuid.uuid4()),
                    "article_id": self.current_article_id,
                    "page_index": page_idx,
                    "bbox_orig": bbox_orig,
                    "type": atype,
                    "text": "",
                }

                # 🖼 IMAGE HANDLING
                if atype == "Image":
                    img_dir = Path(OUTPUT_DIR) / "images"
                    img_dir.mkdir(parents=True, exist_ok=True)

                    crop = pil.crop(bbox_orig)
                    img_filename = safe_filename(f"{art['title']+"-"+ art['id']}_{uuid.uuid4().hex}")+".png"
                    img_path = (
                        img_dir
                        / img_filename
                    )
                    crop.save(img_path)

                    ann["path"] = str(img_path)
                    art.setdefault("images", []).append(
                        {
                            "path": str(img_path),
                            "bbox": bbox_orig,
                            "page_index": page_idx,
                        }
                    )

                # 🧠 TEXT / OCR HANDLING
                else:
                    try:
                        text = ocr_utils.ocr_image_crop(pil, bbox_orig)
                    except Exception as e:
                        text = f"[ocr error: {e}]"

                    ann["text"] = text

                    # ✅ ALWAYS override title if detected
                    if atype == "Title":
                        art["title"] = text.strip()

                    elif atype == "Author" and not art.get("author"):
                        art["author"] = text

                    elif atype == "Date" and not art.get("date"):
                        art["date"] = text

                    elif atype == "Issue" and not art.get("issue"):
                        art["issue"] = text

                    elif atype == "Body":
                        art.setdefault("bodies", []).append(text)

                self.annotations.append(ann)
                art.setdefault("annotations", []).append(ann)

            self.current_page_index = page_idx
            self.root.after(0, self.show_page)

        messagebox.showinfo("ML Detection", "Auto-detection completed for all pages.")

    def auto_detect_current_page(self):
        if not self.pdf_path:
            return
        if self.yolo_model is None:
            try:
                self.yolo_model = YOLO(self.model_path)
            except Exception as e:
                messagebox.showerror("YOLO Error", f"Could not load YOLO model: {e}")
                return
        threading.Thread(
            target=self._detect_page_thread,
            args=(self.current_page_index,),
            daemon=True,
        ).start()

    def _detect_page_thread(self, page_idx):
        pil = self.pages[page_idx]
        temp_path = os.path.join(SESSION_DIR, f"page_{page_idx}.png")
        pil.save(temp_path)
        results = self.yolo_model(temp_path, conf=0.05, iou=0.5)

        # sort boxes by top-left Y coordinate so Titles appear first top-to-bottom
        boxes = [
            (box, int(cls_id))
            for box, cls_id in zip(results[0].boxes.xyxy, results[0].boxes.cls)
        ]
        boxes.sort(key=lambda b: b[0][1])  # sort by y1

        current_article_id = None

        for box, cls_id in boxes:
            x1, y1, x2, y2 = map(float, box)
            class_map = {0: "Title", 1: "Body", 2: "Image"}
            atype = class_map.get(int(cls_id), "Body")
            bbox_orig = (x1, y1, x2, y2)

            # Title starts a new article
            if atype == "Title":
                self.new_article()
                current_article_id = self.current_article_id

            # If first box is not Title, create a temp article
            if current_article_id is None:
                self.new_article()
                current_article_id = self.current_article_id

            art = self.articles[current_article_id]

            ann = {
                "id": str(uuid.uuid4()),
                "article_id": current_article_id,
                "page_index": page_idx,
                "bbox_orig": bbox_orig,
                "type": atype,
                "text": "",
            }

            if atype == "Image":
                img_dir = Path(OUTPUT_DIR) / "images"
                img_dir.mkdir(parents=True, exist_ok=True)
                crop = pil.crop(bbox_orig)
                img_filename = safe_filename(f"{art['title']+"-"+ art['id']}_{uuid.uuid4().hex}")+".png"
                img_path = img_dir / img_filename
                crop.save(img_path)
                ann["path"] = str(img_path)
                art.setdefault("images", []).append(
                    {"path": str(img_path), "bbox": bbox_orig, "page_index": page_idx}
                )
            else:
                try:
                    text = ocr_utils.ocr_image_crop(pil, bbox_orig)
                except Exception as e:
                    text = f"[ocr error: {e}]"
                ann["text"] = text
                if atype == "Title":
                    art["title"] = text
                elif atype == "Body":
                    art.setdefault("bodies", []).append(text)

            self.annotations.append(ann)
            art.setdefault("annotations", []).append(ann)

        self.current_page_index = page_idx
        self.root.after(0, lambda: self.show_page())
        messagebox.showinfo(
            "ML Detection", f"Auto-detection completed for page {page_idx+1}."
        )


def safe_filename(text, max_len=60):
    if not text:
        return "untitled"
    text = text.strip()
    text = re.sub(r"[^\w\s-]", "", text)  # remove special chars
    text = re.sub(r"\s+", "_", text)  # spaces → _
    return text[:max_len]


def main():
    root = Tk()
    app = NewspaperOCRApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
