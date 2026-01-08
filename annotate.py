import os
import json
import tkinter as tk
from tkinter import ttk, filedialog
from PIL import Image, ImageTk
from pdf2image import convert_from_path

# ==== SETTINGS ====
PDF_PATH = r"PDFS/29th July 1977.pdf"
OUTPUT_DIR = "annotations_dataset"
LABELS = ["title", "body", "image"]
POPPLER_PATH = r"C:\poppler-25.07.0\Library\bin"

os.makedirs(OUTPUT_DIR, exist_ok=True)

# --- Convert PDF to images ---
print("Converting PDF to images...")
pages = convert_from_path(PDF_PATH, dpi=200, poppler_path=POPPLER_PATH)
image_paths = []
for i, page in enumerate(pages):
    img_path = os.path.join(OUTPUT_DIR, f"page_{i+1:03d}.jpg")
    page.save(img_path, "JPEG")
    image_paths.append(img_path)
print(f"Converted {len(image_paths)} pages.")


class Annotator(tk.Tk):
    def __init__(self, image_paths):
        super().__init__()
        self.title("📰 Newspaper Annotator")
        self.geometry("1200x900")

        self.image_paths = image_paths
        self.page_index = 0
        self.annotations = []
        self.current_label = tk.StringVar(value=LABELS[0])
        self.start_x = None
        self.start_y = None
        self.rect_id = None
        self.scale = 1.0
        self.pan_start = None

        # UI layout
        self.canvas = tk.Canvas(self, bg="gray")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.bind("<ButtonPress-1>", self.on_click)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.canvas.bind("<MouseWheel>", self.on_zoom)
        self.canvas.bind("<ButtonPress-3>", self.pan_start_event)
        self.canvas.bind("<B3-Motion>", self.pan_move_event)

        toolbar = ttk.Frame(self)
        toolbar.pack(fill=tk.X, pady=5)

        ttk.Label(toolbar, text="Label:").pack(side=tk.LEFT, padx=4)
        label_box = ttk.Combobox(toolbar, textvariable=self.current_label, values=LABELS, state="readonly")
        label_box.pack(side=tk.LEFT, padx=4)

        ttk.Button(toolbar, text="Save", command=self.save_annotation).pack(side=tk.LEFT, padx=4)
        ttk.Button(toolbar, text="Next ▶", command=self.next_page).pack(side=tk.LEFT, padx=4)
        ttk.Button(toolbar, text="⟲ Reset", command=self.reset_boxes).pack(side=tk.LEFT, padx=4)

        self.zoom_label = ttk.Label(toolbar, text="Zoom: 100%")
        self.zoom_label.pack(side=tk.RIGHT, padx=4)

        self.load_page()

    # --- Load and display image ---
    def load_page(self):
        path = self.image_paths[self.page_index]
        self.annotations = []
        self.img = Image.open(path)
        self.tk_img = ImageTk.PhotoImage(self.img)
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, image=self.tk_img, anchor="nw", tags="IMG")
        self.canvas.config(scrollregion=self.canvas.bbox("all"))
        self.title(f"Annotating: {os.path.basename(path)}")

    # --- Mouse events for drawing boxes ---
    def on_click(self, event):
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        self.rect_id = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y,
                                                    outline="red", width=2, tags="BOX")

    def on_drag(self, event):
        if self.rect_id:
            curX, curY = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
            self.canvas.coords(self.rect_id, self.start_x, self.start_y, curX, curY)

    def on_release(self, event):
        if self.rect_id:
            x1, y1, x2, y2 = self.canvas.coords(self.rect_id)
            self.annotations.append({
                "x1": int(x1 / self.scale),
                "y1": int(y1 / self.scale),
                "x2": int(x2 / self.scale),
                "y2": int(y2 / self.scale),
                "label": self.current_label.get()
            })
            self.rect_id = None

    # --- Zoom and pan ---
    def on_zoom(self, event):
        factor = 1.1 if event.delta > 0 else 0.9
        self.scale *= factor
        self.zoom_label.config(text=f"Zoom: {int(self.scale*100)}%")
        self.redraw_image()

    def pan_start_event(self, event):
        self.pan_start = (event.x, event.y)

    def pan_move_event(self, event):
        if self.pan_start:
            dx = event.x - self.pan_start[0]
            dy = event.y - self.pan_start[1]
            self.canvas.move("all", dx, dy)
            self.pan_start = (event.x, event.y)

    def redraw_image(self):
        resized = self.img.resize((int(self.img.width * self.scale), int(self.img.height * self.scale)))
        self.tk_img = ImageTk.PhotoImage(resized)
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, image=self.tk_img, anchor="nw", tags="IMG")
        for ann in self.annotations:
            x1, y1, x2, y2 = [v * self.scale for v in (ann["x1"], ann["y1"], ann["x2"], ann["y2"])]
            color = "green" if ann["label"] == "title" else "blue" if ann["label"] == "body" else "red"
            self.canvas.create_rectangle(x1, y1, x2, y2, outline=color, width=2, tags="BOX")

    # --- Buttons ---
    def reset_boxes(self):
        self.annotations.clear()
        self.redraw_image()

    def save_annotation(self):
        img_path = self.image_paths[self.page_index]
        json_path = img_path.replace(".jpg", ".json")
        with open(json_path, "w") as f:
            json.dump(self.annotations, f, indent=2)
        print(f"✅ Saved annotations → {json_path}")

    def next_page(self):
        self.save_annotation()
        if self.page_index + 1 < len(self.image_paths):
            self.page_index += 1
            self.load_page()
        else:
            print("✅ All pages annotated!")
            self.destroy()


# --- Run the app ---
if __name__ == "__main__":
    app = Annotator(image_paths)
    app.mainloop()
