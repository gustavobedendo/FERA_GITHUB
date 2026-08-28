import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
import sqlite3
import io
import time
from functools import partial
#codereview
def fetch_images_from_db(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    cursor.execute("SELECT path, img_blob FROM thumbnails LIMIT 50")
    rows = cursor.fetchall()
    
    images_info = []
    for row in rows:
        path = row[0]
        img_blob = row[1]
        images_info.append((path, img_blob))
    
    conn.close()
    return images_info

class ThumbnailGallery(tk.Tk):
    def __init__(self, images_info):
        super().__init__()
        self.after_loadthumbnail = None
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.idframe = None
        self.previous_width = None
        self.previous_height = None
        self.previous_configure = time.time()
        self.title("Thumbnail Gallery")
        self.zoom_factors = [1, 2, 3]
        self.current_zoom_index = 1  # Start with zoom factor 2
        self.thumbnail_sizes = [48, 64, 96]

        self.images_info = images_info
        self.thumbnail_size = self.thumbnail_sizes[self.current_zoom_index]
        self.thumbnail_cache = {}  # Cache to store thumbnails
        self.update_idletasks()
        self.max_thumbnails_per_row = self.calculate_max_thumbnails_per_row()
        self.current_page = 0
        self.thumbnails_per_page = 150  # Adjust based on how many thumbnails fit in your window
        self.resizing = False  # Flag to track resizing state
        self.setup_ui()
        self.bind('<Configure>', self.on_window_resize)
        #self.canvas.bind('<Configure>', self.on_canvas_configure)
        self.load_thumbnails()

    def calculate_max_thumbnails_per_row(self):
        # Calculate the number of columns that fit in the initial window size
        available_width = self.winfo_width()
        horizontalpadding = 24
        return available_width // (self.thumbnail_size + horizontalpadding)  # Include padding

    def setup_ui(self):
        # Create canvas
        self.canvas = tk.Canvas(self, bg="white")
        self.canvas.grid(row=0, column=0, sticky='nsew', pady=20, padx=20)

        # Add scrollbar
        self.scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scrollbar.grid(row=0, column=1, sticky='ns')
        
        self.hscrollbar = ttk.Scrollbar(self, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.hscrollbar.grid(row=1, column=0, sticky='ew')
        #self.canvas.configure(yscrollincrement='1')

        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.configure(xscrollcommand=self.hscrollbar.set)

        # Zoom and pagination controls
        control_frame = ttk.Frame(self)
        control_frame.rowconfigure(0, weight=1)
        control_frame.columnconfigure((0,1,2,3), weight=1)
        control_frame.grid(row=2, column=0, sticky='nsew')

        zoom_in_button = ttk.Button(control_frame, text="Zoom In", command=self.zoom_in)
        zoom_in_button.grid(row=0, column=0, sticky='nsew')

        zoom_out_button = ttk.Button(control_frame, text="Zoom Out", command=self.zoom_out)
        zoom_out_button.grid(row=0, column=1, sticky='nsew')

        prev_page_button = ttk.Button(control_frame, text="Previous Page", command=self.prev_page)
        prev_page_button.grid(row=0, column=2, sticky='nsew')

        # Entry for current page
        self.current_page_entry = ttk.Entry(control_frame, width=20)
        self.current_page_entry.grid(row=0, column=4, sticky='ns')
        self.current_page_entry.insert(0, str(self.current_page + 1))  # Initialize with current page number
        self.current_page_entry.bind("<Return>", self.change_page)

        next_page_button = ttk.Button(control_frame, text="Next Page", command=self.next_page)
        next_page_button.grid(row=0, column=3, sticky='nsew')

        # Total pages label
        total_pages_label = ttk.Label(control_frame, text=f"of {len(self.images_info) // self.thumbnails_per_page + 1}")
        total_pages_label.grid(row=0, column=5, sticky='nsew')


    def change_page(self, event=None):
        try:
            new_page = int(self.current_page_entry.get()) - 1  # Convert to 0-based index
            if new_page >= 0 and new_page * self.thumbnails_per_page < len(self.images_info):
                self.current_page = new_page
                self.load_thumbnails()
            else:
                # Invalid page number, reset entry to current page
                self.current_page_entry.delete(0, tk.END)
                self.current_page_entry.insert(0, str(self.current_page + 1))
        except ValueError:
            # If entry contains non-integer value, reset to current page
            self.current_page_entry.delete(0, tk.END)
            self.current_page_entry.insert(0, str(self.current_page + 1))


    def load_thumbnails(self, pidx=None, pcurrent_page=None):
        print("loading")
        if(self.after_loadthumbnail!=None):
            self.after_cancel(self.after_loadthumbnail)
        # Clear current thumbnails
        self.canvas.delete("all")

        # Ensure there are images to display
        if not self.images_info:
            return

        # Update idle tasks to ensure the window layout is up-to-date
        #self.canvas.update_idletasks()
        #self.update_idletasks()
        # Determine the number of columns that fit in the current window size
        available_width = self.winfo_width()
        available_height = self.winfo_height()
        if available_width == 1 or available_height == 1:  # Prevent zero division error during initial load
            available_width = self.winfo_width()
            available_height = self.winfo_height()

        # Calculate the number of columns based on thumbnail size and available width
        verticalpadding = 48
        horizontalpadding = 24
        #max_thumbnails_per_row = available_width // (self.thumbnail_size + horizontalpadding)  # Include padding
        print(self.max_thumbnails_per_row)
        #thumbnails_per_row = min(max_thumbnails_per_row, len(self.images_info))

        start_idx = self.current_page * self.thumbnails_per_page
        current_page = self.current_page
        if(pcurrent_page!=None and pidx!=None):
            start_idx = pidx
        end_idx = start_idx + self.thumbnails_per_page
        
        current_images_info = self.images_info[start_idx:end_idx]
        init = time.time()
        thumbnail_frame = ttk.Frame(width=self.thumbnail_size, height=self.thumbnail_size + verticalpadding)
        thumbnail_frame.rowconfigure(0, weight=1)
        thumbnail_frame.columnconfigure(0, weight=1)
        thumbnail_frame.grid(row=0, column=0, sticky='nsew')
        # Load and display thumbnails
        row=1
        col=1
        for idx, (path, img_blob) in enumerate(current_images_info):
            row = idx // self.max_thumbnails_per_row
            col = idx % self.max_thumbnails_per_row
            #thumbnail_frame = ttk.Frame(self.canvas, width=self.thumbnail_size, height=self.thumbnail_size + verticalpadding)
            self.load_thumbnail(path, img_blob, thumbnail_frame, row, col)
        
        # Update canvas configuration after adding all thumbnails
        self._frame_id = self.canvas.create_window(0, 0, window=thumbnail_frame, anchor=tk.NW)
        self.update_idletasks()
        print(self.canvas.bbox("all"))
        self.canvas.config(scrollregion=thumbnail_frame.bbox())
        #self.canvas.config(scrollregion=(0, 0, 
        ##                                 (self.max_thumbnails_per_row+1)*(self.thumbnail_size+8),
        #                                 (col+1)*(self.thumbnail_size+8)))
        #print(0, 0, (self.max_thumbnails_per_row+1)*(self.thumbnail_size+8),
        #                                 (row+1)*(self.thumbnail_size+8))



            

    def load_thumbnail(self, path, img_blob, thumbnail_frame, row, col):
        if (path, self.thumbnail_size) in self.thumbnail_cache:
            img_tk = self.thumbnail_cache[(path, self.thumbnail_size)]
        else:
            img = Image.open(io.BytesIO(img_blob))
            img.thumbnail((self.thumbnail_size, self.thumbnail_size), Image.LANCZOS)
            img_tk = ImageTk.PhotoImage(img)
            self.thumbnail_cache[(path, self.thumbnail_size)] = img_tk

        img_label = ttk.Label(thumbnail_frame, image=img_tk)
        img_label.image = img_tk  # Keep a reference to avoid garbage collection
        img_label.grid(row=row, column=col, sticky='nsew', padx=4, pady=4)
        #nome = path.split('/')[-1]
        #limit = self.thumbnail_size // 4
        #if(len(nome)> limit):
        #    nome = nome[:limit]+"..."
        #name_label = ttk.Label(thumbnail_frame, text=nome, wraplength=self.thumbnail_size)
        #name_label.pack(side=tk.BOTTOM)

    #def on_canvas_configure(self, event):
    #    self.canvas.config(scrollregion=self.canvas.bbox("all"))

    def on_window_resize(self, event=None):
        current_time = time.time()
        if current_time - self.previous_configure < 0.3:
            return
        if self.previous_width == self.winfo_width() and self.previous_height == self.winfo_height():
            return
        self.previous_width = self.winfo_width()
        self.previous_height = self.winfo_height()
        self.previous_configure = current_time
        #self.canvas.config(scrollregion=self.canvas.bbox("all"))
        #self.update_idletasks()
        #self.canvas.config(xscrollcommand=self.hscrollbar.set, yscrollcommand=self.scrollbar.set)
        print('1')
        self.load_thumbnails()

    def zoom_in(self):
        if self.current_zoom_index < len(self.zoom_factors) - 1:
            self.current_zoom_index += 1
            self.thumbnail_size = self.thumbnail_sizes[self.current_zoom_index]
            self.load_thumbnails()

    def zoom_out(self):
        if self.current_zoom_index > 0:
            self.current_zoom_index -= 1
            self.thumbnail_size = self.thumbnail_sizes[self.current_zoom_index]
            self.load_thumbnails()

    def prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.load_thumbnails()

    def next_page(self):
        if (self.current_page + 1) * self.thumbnails_per_page < len(self.images_info):
            self.current_page += 1
            self.load_thumbnails()

if __name__ == "__main__":
    # Example usage with SQLite database path
    db_path = r"D:\28875-23-Anexo\Anexo\Eq01\sources\ferathumbs.db"
    images_info = fetch_images_from_db(db_path)

    app = ThumbnailGallery(images_info)
    app.mainloop()
