

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import pymupdf

# --- Helper function for PyInstaller ---
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- Dimensions in points (1 inch = 72 points) ---
PAGE_SIZES = {
    "Original Size (No Scaling)": None,
    "ANSI A (8.5 x 11 in)": (612, 792),
    "ANSI B (11 x 17 in)": (792, 1224),
    "ANSI C (17 x 22 in)": (1224, 1584),
    "ISO A4 (210 x 297 mm)": (595, 842),
    "ISO A3 (297 x 420 mm)": (842, 1191),
    "ISO A2 (420 x 594 mm)": (1191, 1684),
}


class RotationEditor(tk.Toplevel):
    """A Toplevel window for editing the rotation of a single PDF's pages."""
    def __init__(self, parent_app, file_object):
        super().__init__(parent_app.root)
        self.parent_app = parent_app
        self.file_object = file_object
        
        self.transient(parent_app.root)
        self.title("Rotate Pages")
        self.geometry("600x700")

        self.doc = pymupdf.open(self.file_object["path"])
        self.current_page = 0
        self.preview_image = None
        self.temp_rotations = self.file_object["rotations"].copy()
        
        self.create_widgets()
        self.after(100, self.render_page)

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(main_frame, bg="lightgrey")
        self.canvas.pack(fill=tk.BOTH, expand=True, pady=5)
        self.canvas.bind("<Configure>", self.render_page)
        
        nav_frame = ttk.Frame(main_frame)
        nav_frame.pack(fill=tk.X, pady=5)
        self.prev_button = ttk.Button(nav_frame, text="<< Prev Page", command=self.prev_page)
        self.prev_button.pack(side=tk.LEFT, padx=5)
        self.page_label = ttk.Label(nav_frame, text="Page 0/0", anchor="center")
        self.page_label.pack(side=tk.LEFT, expand=True, fill=tk.X)
        self.next_button = ttk.Button(nav_frame, text="Next Page >>", command=self.next_page)
        self.next_button.pack(side=tk.RIGHT, padx=5)

        control_frame = ttk.LabelFrame(main_frame, text="Rotate Current Page", padding="10")
        control_frame.pack(fill=tk.X, pady=5)
        ttk.Button(control_frame, text="Rotate Left (-90°)", command=lambda: self.rotate_page(-90)).pack(side=tk.LEFT, expand=True, padx=5)
        ttk.Button(control_frame, text="Rotate Right (+90°)", command=lambda: self.rotate_page(90)).pack(side=tk.RIGHT, expand=True, padx=5)

        all_pages_frame = ttk.LabelFrame(main_frame, text="Rotate All Pages", padding="10")
        all_pages_frame.pack(fill=tk.X, pady=5)
        ttk.Button(all_pages_frame, text="Rotate All +90°", command=self.rotate_all).pack(fill=tk.X, expand=True)
        
        ttk.Button(main_frame, text="Apply & Close", command=self.apply_and_close).pack(pady=10)

    def render_page(self, event=None):
        self.canvas.delete("all")
        if not self.doc: return
        
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        if canvas_width <= 1 or canvas_height <= 1: return

        page = self.doc.load_page(self.current_page)
        
        current_rotation = self.temp_rotations.get(self.current_page, 0)
        page.set_rotation(current_rotation)
        
        page_rect = page.rect
        zoom_x = canvas_width / page_rect.width
        zoom_y = canvas_height / page_rect.height
        zoom = min(zoom_x, zoom_y) * 0.98
        
        matrix = pymupdf.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        img_data = pix.tobytes("ppm")
        self.preview_image = tk.PhotoImage(data=img_data)
        
        self.canvas.create_image(canvas_width/2, canvas_height/2, image=self.preview_image)
        self.update_nav()

    def update_nav(self):
        total_pages = len(self.doc)
        current_rotation = self.temp_rotations.get(self.current_page, 0)
        self.page_label.config(text=f"Page {self.current_page + 1} of {total_pages} (Rotation: {current_rotation}°)")
        self.prev_button.config(state=tk.NORMAL if self.current_page > 0 else tk.DISABLED)
        self.next_button.config(state=tk.NORMAL if self.current_page < total_pages - 1 else tk.DISABLED)

    def rotate_page(self, angle):
        current_rotation = self.temp_rotations.get(self.current_page, 0)
        new_rotation = (current_rotation + angle) % 360
        self.temp_rotations[self.current_page] = new_rotation
        self.render_page()

    def rotate_all(self):
        for i in range(len(self.doc)):
            current_rotation = self.temp_rotations.get(i, 0)
            new_rotation = (current_rotation + 90) % 360
            self.temp_rotations[i] = new_rotation
        self.render_page()

    def prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.render_page()

    def next_page(self):
        if self.current_page < len(self.doc) - 1:
            self.current_page += 1
            self.render_page()

    def apply_and_close(self):
        self.file_object["rotations"] = self.temp_rotations
        self.doc.close()
        self.parent_app.update_output_preview()
        self.destroy()


class PDFMergerApp:
    def __init__(self, root_window):
        self.root = root_window
        self.root.title("SwettyPDF 📄")
        self.root.geometry("1100x700")
        self.root.minsize(900, 600)

        try:
            self.root.iconbitmap(resource_path("icon.ico"))
        except tk.TclError:
            print("Icon file 'icon.ico' not found.")
        
        self.style = ttk.Style(self.root)
        if 'xpnative' in self.style.theme_names(): self.style.theme_use('xpnative')
        elif 'aqua' in self.style.theme_names(): self.style.theme_use('aqua')
        else: self.style.theme_use('clam')

        self.file_objects = []
        self.selected_page_size = tk.StringVar()
        self.orientation_var = tk.StringVar(value="Portrait") # **NEW** state for orientation
        
        self.output_preview_doc = None
        self.current_page = 0
        self.preview_image = None

        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        left_column = ttk.Frame(main_frame, width=400)
        left_column.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        left_column.pack_propagate(False)
        right_column = ttk.Frame(main_frame)
        right_column.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        step1_frame = ttk.LabelFrame(left_column, text="Step 1: Add Files to Merge")
        step1_frame.pack(fill='x', pady=(5, 10))
        ttk.Button(step1_frame, text="Browse for PDF Files...", command=self.select_files).pack(fill='x', expand=True, padx=10, pady=10)
        reorder_frame = ttk.LabelFrame(left_column, text="Step 2: Review and Reorder Files")
        reorder_frame.pack(fill='both', expand=True, pady=5)
        self.listbox = tk.Listbox(reorder_frame, height=10)
        self.listbox.pack(side=tk.LEFT, fill='both', expand=True, padx=(10, 5), pady=10)
        reorder_btn_frame = ttk.Frame(reorder_frame)
        reorder_btn_frame.pack(side=tk.RIGHT, fill='y', padx=(5, 10), pady=10)
        ttk.Button(reorder_btn_frame, text="↑ Move Up", command=self.move_up).pack(pady=2, fill='x')
        ttk.Button(reorder_btn_frame, text="↓ Move Down", command=self.move_down).pack(pady=2, fill='x')
        ttk.Button(reorder_btn_frame, text="Rotate File...", command=self.open_rotation_editor).pack(pady=10, fill='x')
        ttk.Separator(reorder_btn_frame, orient='horizontal').pack(pady=5, fill='x')
        ttk.Button(reorder_btn_frame, text="Remove", command=self.remove_file).pack(pady=2, fill='x')
        step3_frame = ttk.LabelFrame(left_column, text="Step 3: Choose Output Page Size (Optional)")
        step3_frame.pack(fill='x', pady=10)

        # **NEW** Frame to hold both dropdown and radio buttons
        size_control_frame = ttk.Frame(step3_frame)
        size_control_frame.pack(fill='x', expand=True, padx=10, pady=10)

        self.selected_page_size.set("Original Size (No Scaling)")
        self.selected_page_size.trace_add("write", self._on_page_size_change)
        self.orientation_var.trace_add("write", self.update_output_preview)

        ttk.Combobox(size_control_frame, textvariable=self.selected_page_size, values=list(PAGE_SIZES.keys()), state="readonly").pack(side=tk.LEFT, fill='x', expand=True)

        # **NEW** Frame for orientation radio buttons
        orientation_frame = ttk.Frame(size_control_frame)
        orientation_frame.pack(side=tk.LEFT, padx=(10, 0))
        self.portrait_radio = ttk.Radiobutton(orientation_frame, text="Portrait", variable=self.orientation_var, value="Portrait")
        self.portrait_radio.pack(anchor="w") # anchor="w" aligns the text to the left
        self.landscape_radio = ttk.Radiobutton(orientation_frame, text="Landscape", variable=self.orientation_var, value="Landscape")
        self.landscape_radio.pack(anchor="w")

        step4_frame = ttk.LabelFrame(left_column, text="Step 4: Create Final PDF")
        step4_frame.pack(fill='x', pady=(5, 10))
        ttk.Button(step4_frame, text="Merge & Save PDF", command=self.merge_pdfs).pack(fill='x', expand=True, padx=10, pady=10)
        preview_frame = ttk.LabelFrame(right_column, text="Live Output Preview")
        preview_frame.pack(fill=tk.BOTH, expand=True)
        self.preview_canvas = tk.Canvas(preview_frame, bg="lightgrey")
        self.preview_canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.preview_canvas.bind("<Configure>", self.render_page)
        nav_frame = ttk.Frame(preview_frame)
        nav_frame.pack(fill=tk.X, pady=5)
        self.prev_button = ttk.Button(nav_frame, text="<< Previous", command=self.prev_page)
        self.prev_button.pack(side=tk.LEFT, padx=10)
        self.page_label = ttk.Label(nav_frame, text="Page 0 of 0")
        self.page_label.pack(side=tk.LEFT, expand=True)
        self.next_button = ttk.Button(nav_frame, text="Next >>", command=self.next_page)
        self.next_button.pack(side=tk.RIGHT, padx=10)

        # Initialize the state of the radio buttons
        self._on_page_size_change()
        
    def _on_page_size_change(self, *args):
        """**NEW** Handles changes to the page size dropdown."""
        is_original = self.selected_page_size.get() == "Original Size (No Scaling)"
        new_state = tk.DISABLED if is_original else tk.NORMAL
        
        self.portrait_radio.config(state=new_state)
        self.landscape_radio.config(state=new_state)
        
        self.update_output_preview()

    def open_rotation_editor(self):
        selected_indices = self.listbox.curselection()
        if not selected_indices:
            messagebox.showinfo("Information", "Please select a file from the list to rotate.")
            return
        selected_index = selected_indices[0]
        file_object = self.file_objects[selected_index]
        RotationEditor(self, file_object)

    def update_output_preview(self, *args):
        """Creates an in-memory merged PDF, preserving all vector and annotation data."""
        if not self.file_objects:
            self.output_preview_doc = None
            self.current_page = 0
            self.render_page()
            return
            
        try:
            self.output_preview_doc = pymupdf.open() 
            size_key = self.selected_page_size.get()
            target_size = PAGE_SIZES.get(size_key)

            # **MODIFIED** - Apply orientation setting
            if target_size:
                orientation = self.orientation_var.get()
                w, h = target_size
                # Ensure we have canonical portrait dimensions first (width < height)
                p_w, p_h = (w, h) if w < h else (h, w)
                
                if orientation == "Landscape":
                    target_size = (p_h, p_w) # Swap for landscape
                else:
                    target_size = (p_w, p_h) # Use portrait

            for file_obj in self.file_objects:
                with pymupdf.open(file_obj["path"]) as source_doc:
                    rotations = file_obj["rotations"]
                    
                    if target_size is None:
                        # If not resizing, insert the pristine doc...
                        start_page = self.output_preview_doc.page_count
                        self.output_preview_doc.insert_pdf(source_doc)
                        # ...then apply rotations to the newly added pages in the output doc.
                        for page_num, angle in rotations.items():
                            if angle != 0:
                                self.output_preview_doc[start_page + page_num].set_rotation(angle)
                    else:
                        # --- VECTOR-PRESERVING RESIZING LOGIC ---
                        target_width, target_height = target_size
                        for source_page in source_doc:
                            angle = rotations.get(source_page.number, 0)
                            
                            # Determine the source page's dimensions AFTER rotation
                            r_width, r_height = source_page.rect.width, source_page.rect.height
                            if angle in [90, 270]:
                                r_width, r_height = r_height, r_width # Swap dimensions
                            
                            # Calculate the rectangle to draw into, preserving aspect ratio
                            zoom_x = target_width / r_width
                            zoom_y = target_height / r_height
                            zoom = min(zoom_x, zoom_y)

                            final_width = r_width * zoom
                            final_height = r_height * zoom
                            x0 = (target_width - final_width) / 2
                            y0 = (target_height - final_height) / 2
                            final_rect = pymupdf.Rect(x0, y0, x0 + final_width, y0 + final_height)
                            
                            # Create new page and draw the source page content into the calculated rect
                            new_page = self.output_preview_doc.new_page(width=target_width, height=target_height)
                            # Pass the rotation directly to the drawing function
                            new_page.show_pdf_page(final_rect, source_doc, source_page.number, rotate=-angle)
            
            self.current_page = 0
            self.render_page()

        except Exception as e:
            self.output_preview_doc = None
            messagebox.showerror("Preview Error", f"Could not generate preview:\n{e}")
            self.render_page()

    def render_page(self, event=None):
        self.preview_canvas.delete("all")
        if not self.output_preview_doc or len(self.output_preview_doc) == 0:
            self.page_label.config(text="Page 0 of 0")
            self.prev_button.config(state=tk.DISABLED)
            self.next_button.config(state=tk.DISABLED)
            return
        canvas_width = self.preview_canvas.winfo_width()
        canvas_height = self.preview_canvas.winfo_height()
        if canvas_width <= 1 or canvas_height <= 1: return
        try:
            page = self.output_preview_doc.load_page(self.current_page)
            page_rect = page.rect
            zoom_x = canvas_width / page_rect.width
            zoom_y = canvas_height / page_rect.height
            zoom = min(zoom_x, zoom_y) * 0.98
            matrix = pymupdf.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            img_data = pix.tobytes("ppm")
            self.preview_image = tk.PhotoImage(data=img_data)
            self.preview_canvas.create_image(canvas_width/2, canvas_height/2, image=self.preview_image)
            total_pages = len(self.output_preview_doc)
            self.page_label.config(text=f"Page {self.current_page + 1} of {total_pages}")
            self.prev_button.config(state=tk.NORMAL if self.current_page > 0 else tk.DISABLED)
            self.next_button.config(state=tk.NORMAL if self.current_page < total_pages - 1 else tk.DISABLED)
        except Exception as e:
            messagebox.showerror("Render Error", f"Could not display page:\n{e}")

    def prev_page(self):
        if self.output_preview_doc and self.current_page > 0:
            self.current_page -= 1
            self.render_page()

    def next_page(self):
        if self.output_preview_doc and self.current_page < len(self.output_preview_doc) - 1:
            self.current_page += 1
            self.render_page()

    def select_files(self):
        filepaths = filedialog.askopenfilenames(title="Select PDF files", filetypes=(("PDF files", "*.pdf"),))
        if filepaths:
            for fp in filepaths:
                self.file_objects.append({"path": fp, "rotations": {}})
            self.update_listbox()

    def update_listbox(self):
        self.listbox.delete(0, tk.END)
        for f_obj in self.file_objects:
            self.listbox.insert(tk.END, os.path.basename(f_obj["path"]))
        self.update_output_preview()

    def move_up(self):
        selected_indices = self.listbox.curselection()
        if not selected_indices: return
        idx = selected_indices[0]
        if idx > 0:
            self.file_objects.insert(idx - 1, self.file_objects.pop(idx))
            self.update_listbox()
            self.listbox.selection_set(idx - 1)
            self.listbox.activate(idx - 1)

    def move_down(self):
        selected_indices = self.listbox.curselection()
        if not selected_indices: return
        idx = selected_indices[0]
        if idx < len(self.file_objects) - 1:
            self.file_objects.insert(idx + 1, self.file_objects.pop(idx))
            self.update_listbox()
            self.listbox.selection_set(idx + 1)
            self.listbox.activate(idx + 1)

    def remove_file(self):
        selected_indices = self.listbox.curselection()
        if not selected_indices: return
        self.file_objects.pop(selected_indices[0])
        self.update_listbox()

    def merge_pdfs(self):
        if not self.file_objects:
            messagebox.showerror("Error", "No PDF files selected!")
            return
        output_filename = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], title="Save merged PDF as...")
        if not output_filename: return
        if self.output_preview_doc and len(self.output_preview_doc) > 0:
            try:
                self.output_preview_doc.save(output_filename, garbage=4, deflate=True, clean=True)
                messagebox.showinfo("Success", f"Successfully merged PDFs into:\n{output_filename}")
                self.file_objects.clear()
                self.update_listbox()
            except Exception as e:
                messagebox.showerror("Save Error", f"Could not save the PDF:\n{e}")
        else:
            messagebox.showerror("Error", "No documents selected or preview generated to save.")

def main() -> None:
    try:
        from ctypes import windll
        myappid = 'SWETT.SWETT-PDF-TOOL.1.0' 
        windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except (ImportError, AttributeError):
        pass
    root = tk.Tk()
    PDFMergerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()