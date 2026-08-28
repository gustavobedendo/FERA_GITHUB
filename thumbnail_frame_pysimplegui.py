import sqlite3
import math
import base64
import io
from PIL import Image
import PySimpleGUI as sg

def fetch_images_from_db(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute("SELECT path, img_blob FROM thumbnails")
    rows = cursor.fetchall()

    images_info = []
    for row in rows:
        path = row[0]
        img_blob = row[1]
        images_info.append((path, img_blob))

    conn.close()
    return images_info

class PySimpleGUIWidget:
    def __init__(self, images_info):
        self.images_info = images_info
        self.thumbnail_cache = {}
        self.current_page = 0
        self.thumbnails_per_page = 600
        
        self.thumbnail_sizes = [64, 96, 128, 160, 196]
        self.thumbnail_sizes_index = 1
        self.thumbnail_size = self.thumbnail_sizes[self.thumbnail_sizes_index]

        self.layout = [
            [sg.Button("Zoom In", key="-ZOOM_IN-"), sg.Button("Zoom Out", key="-ZOOM_OUT-"), 
             sg.Button("Previous Page", key="-PREV_PAGE-"), sg.InputText("1", size=(5, 1), key="-PAGE_INPUT-"), 
             sg.Button("Next Page", key="-NEXT_PAGE-"), sg.Text("", key="-PAGE_LABEL-", size=(20, 1))],
            [sg.Column([[]], key="-THUMBNAIL_AREA-", scrollable=True, size=(800, 600), vertical_scroll_only=False)]
        ]

        self.window = sg.Window("Image Viewer", self.layout, resizable=True, finalize=True)
        self.load_thumbnails()

    def load_thumbnails(self):
        self.window['-THUMBNAIL_AREA-'].Widget.pack_forget()
        self.window['-THUMBNAIL_AREA-'].Widget.pack(fill='both', expand=True)

        thumbnails_per_row = self.window['-THUMBNAIL_AREA-'].get_size()[0] // self.thumbnail_size
        self.window['-THUMBNAIL_AREA-'].Widget.grid_columnconfigure(0, weight=1)
        self.window['-THUMBNAIL_AREA-'].Widget.grid_rowconfigure(0, weight=1)
        start_idx = self.current_page * self.thumbnails_per_page
        end_idx = start_idx + self.thumbnails_per_page
        current_images_info = self.images_info[start_idx:end_idx]

        for i in range(0, len(current_images_info), thumbnails_per_row):
            row_layout = []
            for path, img_blob in current_images_info[i:i + thumbnails_per_row]:
                thumbnail_image = self.load_thumbnail(path, img_blob)
                if thumbnail_image:
                    row_layout.append(sg.Image(data=thumbnail_image))
            self.window.extend_layout(self.window['-THUMBNAIL_AREA-'], [row_layout])

        max_page = math.ceil(len(self.images_info) / self.thumbnails_per_page)
        self.window['-PAGE_LABEL-'].update(f"Page {self.current_page + 1}/{max_page}")
        self.window['-PAGE_INPUT-'].update(str(self.current_page + 1))

    def load_thumbnail(self, path, img_blob):
        try:
            if (path, self.thumbnail_size) in self.thumbnail_cache:
                thumbnail = self.thumbnail_cache[(path, self.thumbnail_size)]
            else:
                img = Image.open(io.BytesIO(img_blob))
                img.thumbnail((self.thumbnail_size, self.thumbnail_size))
                bio = io.BytesIO()
                img.save(bio, format="PNG")
                thumbnail = base64.b64encode(bio.getvalue()).decode('utf-8')
                self.thumbnail_cache[(path, self.thumbnail_size)] = thumbnail
            return thumbnail
        except Exception as e:
            print(f"Error loading thumbnail for {path}: {e}")
            return None

    def run(self):
        while True:
            event, values = self.window.read()
            if event == sg.WIN_CLOSED:
                break
            elif event == "-ZOOM_IN-":
                if self.thumbnail_sizes_index < len(self.thumbnail_sizes) - 1:
                    self.thumbnail_sizes_index += 1
                    self.thumbnail_size = self.thumbnail_sizes[self.thumbnail_sizes_index]
                    self.load_thumbnails()
            elif event == "-ZOOM_OUT-":
                if self.thumbnail_sizes_index > 0:
                    self.thumbnail_sizes_index -= 1
                    self.thumbnail_size = self.thumbnail_sizes[self.thumbnail_sizes_index]
                    self.load_thumbnails()
            elif event == "-PREV_PAGE-":
                if self.current_page > 0:
                    self.current_page -= 1
                    self.load_thumbnails()
            elif event == "-NEXT_PAGE-":
                max_page = math.ceil(len(self.images_info) / self.thumbnails_per_page)
                if (self.current_page + 1) < max_page:
                    self.current_page += 1
                    self.load_thumbnails()
            elif event == "-PAGE_INPUT-":
                try:
                    new_page = int(values["-PAGE_INPUT-"]) - 1
                    max_page = math.ceil(len(self.images_info) / self.thumbnails_per_page)
                    if 0 <= new_page < max_page:
                        self.current_page = new_page
                        self.load_thumbnails()
                    else:
                        self.window['-PAGE_INPUT-'].update(str(self.current_page + 1))
                except ValueError:
                    self.window['-PAGE_INPUT-'].update(str(self.current_page + 1))

        self.window.close()

def main():
    db_path = r"D:\28875-23-Anexo\Anexo\Eq01\sources\ferathumbs.db"
    images_info = fetch_images_from_db(db_path)
    viewer = PySimpleGUIWidget(images_info=images_info)
    viewer.run()

if __name__ == "__main__":
    main()
