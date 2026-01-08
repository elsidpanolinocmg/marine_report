import os
import re
import pytesseract
from tkinter import filedialog, Tk, Button, Text, Label, Scrollbar, Frame, Canvas, BOTH, RIGHT, Y, LEFT, END
from pdf2image import convert_from_path
from PIL import Image, ImageTk, ImageEnhance

# OCR setup (make sure Tesseract is installed)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


class NewspaperOCRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Newspaper OCR App v2")
        self.root.geometry("1200x800")

        # === Panels ===
        self.left_frame = Frame(self.root, width=600, height=800)
        self.left_frame.pack(side=LEFT, fill=BOTH, expand=True)

        self.right_frame = Frame(self.root, width=600, height=800)
        self.right_frame.pack(side=RIGHT, fill=BOTH, expand=True)

        # === Canvas for image ===
        self.canvas = Canvas(self.left_frame, bg="black")
        self.canvas.pack(fill=BOTH, expand=True)

        # === Text area ===
        self.text_area = Text(self.right_frame, wrap="word", font=("Arial", 12))
        self.text_area.pack(expand=True, fill=BOTH)

        scroll = Scrollbar(self.text_area)
        scroll.pack(side=RIGHT, fill=Y)
        scroll.config(command=self.text_area.yview)
        self.text_area.config(yscrollcommand=scroll.set)

        # === Status bar ===
        self.status = Label(self.root, text="Ready", anchor="w")
        self.status.pack(side="bottom", fill="x")

        # === Buttons ===
        frame = Frame(self.root)
        frame.pack(side="bottom", fill="x")

        Button(frame, text="Load Folder", command=self.load_folder).pack(side="left", padx=5, pady=5)
        Button(frame, text="Prev Page", command=self.prev_page).pack(side="left", padx=5)
        Button(frame, text="Next Page", command=self.next_page).pack(side="left", padx=5)
        Button(frame, text="Zoom In", command=lambda: self.zoom(1.2)).pack(side="left", padx=5)
        Button(frame, text="Zoom Out", command=lambda: self.zoom(0.8)).pack(side="left", padx=5)

        # Highlight buttons
        Button(frame, text="Title", command=lambda: self.set_highlight_type("title")).pack(side="right", padx=5)
        Button(frame, text="Author", command=lambda: self.set_highlight_type("author")).pack(side="right", padx=5)
        Button(frame, text="Date", command=lambda: self.set_highlight_type("date")).pack(side="right", padx=5)
        Button(frame, text="Issue", command=lambda: self.set_highlight_type("issue")).pack(side="right", padx=5)
        Button(frame, text="Body", command=lambda: self.set_highlight_type("body")).pack(side="right", padx=5)

        # === Variables ===
        self.pdf_files = []
        self.pages = []
        self.current_page_index = 0
        self.zoom_level = 1.0
        self.current_image = None
        self.tk_image = None

        # Highlights storage (dict: page -> list of annotations)
        self.highlights = {}

        # Highlight colors
        self.highlight_colors = {
            "title": "blue",
            "author": "red",
            "date": "purple",
            "issue": "orange",
            "body": "green"
        }
        self.current_highlight_type = "body"

        # Mouse interaction
        self.start_x = None
        self.start_y = None
        self.canvas.bind("<Button-1>", self.on_click)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.temp_rect = None

    # === File loading ===
    def load_folder(self):
        folder = filedialog.askdirectory()
        if not folder:
            return
        self.pdf_files = [os.path.join(folder, f) for f in os.listdir(folder) if f.endswith(".pdf")]
        if not self.pdf_files:
            self.status.config(text="No PDF found in folder")
            return

        self.process_pdf(self.pdf_files[0])

    def process_pdf(self, filepath):
        self.status.config(text=f"Loading {os.path.basename(filepath)} …")
        self.root.update_idletasks()

        self.pages = convert_from_path(filepath, dpi=150)
        self.current_page_index = 0
        self.show_page()

    # === Page display ===
    def show_page(self):
        if not self.pages:
            return

        page = self.pages[self.current_page_index]

        # Apply zoom
        w, h = page.size
        zoomed = page.resize((int(w * self.zoom_level), int(h * self.zoom_level)))
        self.current_image = zoomed
        self.tk_image = ImageTk.PhotoImage(zoomed)

        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor="nw", image=self.tk_image)

        # Run OCR
        text = pytesseract.image_to_string(page, config="--psm 4")
        self.text_area.delete("1.0", END)
        self.text_area.insert("1.0", text)

        # Status
        self.status.config(text=f"Page {self.current_page_index+1}/{len(self.pages)}")

        # Draw saved highlights
        self.draw_highlights()

    def next_page(self):
        if self.current_page_index < len(self.pages) - 1:
            self.current_page_index += 1
            self.show_page()

    def prev_page(self):
        if self.current_page_index > 0:
            self.current_page_index -= 1
            self.show_page()

    def zoom(self, factor):
        self.zoom_level *= factor
        self.show_page()

    # === Highlighting ===
    def set_highlight_type(self, htype):
        self.current_highlight_type = htype
        self.status.config(text=f"Highlight type: {htype}")

    def on_click(self, event):
        self.start_x, self.start_y = event.x, event.y

    def on_drag(self, event):
        if self.temp_rect:
            self.canvas.delete(self.temp_rect)
        self.temp_rect = self.canvas.create_rectangle(
            self.start_x, self.start_y, event.x, event.y,
            outline=self.highlight_colors[self.current_highlight_type],
            width=2
        )

    def on_release(self, event):
        if not self.pages:
            return
        x1, y1, x2, y2 = self.start_x, self.start_y, event.x, event.y
        rect = (x1, y1, x2, y2)

        # Save highlight in original image coordinates
        page = self.pages[self.current_page_index]
        orig_w, orig_h = page.size
        disp_w, disp_h = self.current_image.size

        scale_x = orig_w / disp_w
        scale_y = orig_h / disp_h

        rect_orig = (int(x1 * scale_x), int(y1 * scale_y), int(x2 * scale_x), int(y2 * scale_y))

        self.highlights.setdefault(self.current_page_index, []).append(
            {"type": self.current_highlight_type, "rect": rect_orig}
        )

        self.draw_highlights()

    def draw_highlights(self):
        self.canvas.delete("highlight")
        if self.current_page_index not in self.highlights:
            return

        page = self.pages[self.current_page_index]
        orig_w, orig_h = page.size
        disp_w, disp_h = self.current_image.size
        scale_x = disp_w / orig_w
        scale_y = disp_h / orig_h

        for hl in self.highlights[self.current_page_index]:
            x1, y1, x2, y2 = hl["rect"]
            self.canvas.create_rectangle(
                x1 * scale_x, y1 * scale_y, x2 * scale_x, y2 * scale_y,
                outline=self.highlight_colors[hl["type"]],
                width=2,
                tags="highlight"
            )


if __name__ == "__main__":
    root = Tk()
    app = NewspaperOCRApp(root)
    root.mainloop()
