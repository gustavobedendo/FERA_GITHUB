import sys
from PySide2.QtWidgets import QSizePolicy, QApplication, QScrollArea, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QLineEdit, QFrame
from PySide2.QtCore import Qt, QSize, QObject, QTimer, QEvent
from PySide2.QtGui import QPixmap, QImage, QIntValidator
import sqlite3, math

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

class PySide2Widget(QWidget):
    def __init__(self, parent=None, images_info=None):
        super().__init__(parent)
        self.images_info = images_info
        self.thumbnail_cache = {}
        self.current_page = 0
        self.thumbnails_per_page = 600
        
        self.thumbnail_sizes = [64,96,128,160,196]
        self.thumbnail_sizes_index = 1
        self.thumbnail_size = self.thumbnail_sizes[self.thumbnail_sizes_index]
        self.previous_horizontal_limit = None
        self.setup_ui()
        self.load_thumbnails()
        

    def setup_ui(self):
        layout = QVBoxLayout(self)

        # Scroll Area Widget Contents
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_content = QWidget(self.scroll_area)
        self.scroll_layout = QVBoxLayout(self.scroll_content)
        self.scroll_area.setWidget(self.scroll_content)

        # Scroll Bars
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)  # Add horizontal scrollbar

        # Restrict horizontal expansion
        self.scroll_content.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Preferred)

        # Zoom and pagination controls
        self.control_frame = QFrame()
        control_layout = QHBoxLayout(self.control_frame)
        control_layout.setContentsMargins(0, 0, 0, 0)

        zoom_in_button = QPushButton("Zoom In")
        zoom_in_button.clicked.connect(self.zoom_in)
        control_layout.addWidget(zoom_in_button)

        zoom_out_button = QPushButton("Zoom Out")
        zoom_out_button.clicked.connect(self.zoom_out)
        control_layout.addWidget(zoom_out_button)

        prev_page_button = QPushButton("Previous Page")
        prev_page_button.clicked.connect(self.prev_page)
        control_layout.addWidget(prev_page_button)

        self.current_page_entry = QLineEdit()
        self.current_page_entry.setMaximumWidth(50)  # Set maximum width
        self.current_page_entry.setValidator(QIntValidator())  # Only accept integers
        self.current_page_entry.returnPressed.connect(self.change_page)
        control_layout.addWidget(self.current_page_entry)

        next_page_button = QPushButton("Next Page")
        next_page_button.clicked.connect(self.next_page)
        control_layout.addWidget(next_page_button)

        self.page_label = QLabel()
        page_layout = QHBoxLayout()
        page_layout.addStretch()  # Add stretch to push the label to the right
        page_layout.addWidget(self.page_label)
        control_layout.addLayout(page_layout)

        layout.addWidget(self.control_frame)
        layout.addWidget(self.scroll_area)

        # Connect resize event of scroll area to update thumbnails layout
        self.scroll_area.installEventFilter(self)



    def change_page(self):
        print("Teste")
        try:
            new_page = int(self.current_page_entry.text()) - 1
            max_page = math.ceil(len(self.images_info) / self.thumbnails_per_page)
            
            if 0 <= new_page < max_page:
                self.current_page = new_page
                self.load_thumbnails(True)
            else:
                self.current_page_entry.setText(str(self.current_page + 1))
        except ValueError:
            self.current_page_entry.setText(str(self.current_page + 1))

    def load_thumbnails(self, reload=False):
        self.clear_thumbnails()
        start_idx = self.current_page * self.thumbnails_per_page
        end_idx = start_idx + self.thumbnails_per_page
        current_images_info = self.images_info[start_idx:end_idx]

        # Get the width of the parent component
        parent_width = self.control_frame.width()-32 if self.control_frame else 64

        # Set the width of the scroll content to the width of the parent component
        self.scroll_content.setFixedWidth(parent_width)

        max_thumbnails_per_row = math.ceil(parent_width / self.thumbnail_size)
        if(reload or self.previous_horizontal_limit != max_thumbnails_per_row):
            self.previous_horizontal_limit = max_thumbnails_per_row
        else:
            return
        row_layout = QHBoxLayout()

        for path, img_blob in current_images_info:
            pixmap = self.load_thumbnail(path, img_blob)
            thumbnail_label = QLabel()
            thumbnail_label.setPixmap(pixmap)
            thumbnail_label.setScaledContents(True)
            thumbnail_label.setFixedSize(self.thumbnail_size, self.thumbnail_size)  # Set fixed size
            name_label = QLabel(path.split('/')[-1][:self.thumbnail_size // 4] + "...")

            thumbnail_frame = QFrame()
            thumbnail_layout = QVBoxLayout(thumbnail_frame)
            thumbnail_layout.addWidget(thumbnail_label)
            thumbnail_layout.addWidget(name_label)

            row_layout.addWidget(thumbnail_frame)

            # Add row to scroll layout when it's filled
            if row_layout.count() >= max_thumbnails_per_row:
                self.scroll_layout.addLayout(row_layout)
                row_layout = QHBoxLayout()

        # Add the last row if it's not fully filled
        if row_layout.count() > 0:
            self.scroll_layout.addLayout(row_layout)

        # Update page label
        max_page = math.ceil(len(self.images_info) / self.thumbnails_per_page)
        self.page_label.setText(f"Page {self.current_page + 1}/{max_page}")

    def load_thumbnail(self, path, img_blob):
        try:
            if (path, self.thumbnail_size) in self.thumbnail_cache:
                pixmap = self.thumbnail_cache[(path, self.thumbnail_size)]
            else:
                img = QImage()
                assert img.loadFromData(img_blob)
                pixmap = QPixmap.fromImage(img).scaled(self.thumbnail_size, self.thumbnail_size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                self.thumbnail_cache[(path, self.thumbnail_size)] = pixmap

            return pixmap
        except Exception as e:
            print(f"Error loading thumbnail for {path}: {e}")
            return None

    def clear_thumbnails(self):
        for i in reversed(range(self.scroll_layout.count())):
            item = self.scroll_layout.itemAt(i)
            if isinstance(item, QHBoxLayout):
                for j in reversed(range(item.count())):
                    widget = item.itemAt(j).widget()
                    if widget:
                        widget.deleteLater()
                self.scroll_layout.removeItem(item)
            else:
                widget = item.widget()
                if widget:
                    widget.deleteLater()
                self.scroll_layout.removeItem(item)

    def zoom_in(self):
        if self.thumbnail_sizes_index < len(self.thumbnail_sizes) - 1:
            self.thumbnail_sizes_index += 1
            self.thumbnail_size = self.thumbnail_sizes[self.thumbnail_sizes_index]
            self.load_thumbnails()

    def zoom_out(self):
        if self.thumbnail_sizes_index > 0:
            self.thumbnail_sizes_index -= 1
            self.thumbnail_size = self.thumbnail_sizes[self.thumbnail_sizes_index]
            self.load_thumbnails()

    def prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.load_thumbnails(True)

    def next_page(self):
        max_page = math.ceil(len(self.images_info) / self.thumbnails_per_page)
        if (self.current_page + 1) < max_page:
            self.current_page += 1
            self.load_thumbnails(True)

    # Event filter for handling resize event of the scroll area
    def eventFilter(self, obj, event):
        if event.type() == QEvent.Resize and obj is self.scroll_area:
            QTimer.singleShot(0, self.update_thumbnail_layout)  # Trigger thumbnail layout update after resize event
        return super().eventFilter(obj, event)

    def update_thumbnail_layout(self):
        self.load_thumbnails()  # Reload thumbnails layout after resize event

def main():
    db_path = r"D:\28875-23-Anexo\Anexo\Eq01\sources\ferathumbs.db"
    images_info = fetch_images_from_db(db_path)
    #print(images_info)
    app = QApplication([])
    pyside2_widget = PySide2Widget(parent=None, images_info=images_info)
    #pyside2_widget.images_info = images_info  # Assign images_info to the widget
    pyside2_widget.show()  # Ensure the widget is visible
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()

