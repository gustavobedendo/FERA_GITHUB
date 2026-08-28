import tkinter as tk
from tkinter import ttk

class CustomPanedWindow(ttk.PanedWindow):
    def __init__(self, parent, columns, *args, **kwargs):
        super().__init__(parent, orient=tk.VERTICAL, *args, **kwargs)

        # Top Frame: Treeview
        self.top_frame = ttk.Frame(self)
        self.add(self.top_frame, weight=1)

        self.treeview = self._create_treeview(self.top_frame, columns)
        self.treeview.pack(fill=tk.BOTH, expand=True)

        # Bottom Frame: Notebook
        self.bottom_frame = ttk.Frame(self)
        self.add(self.bottom_frame, weight=1)

        self.notebook = self._create_notebook(self.bottom_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

    def _create_treeview(self, parent, columns):
        # Create the Treeview with Scrollbars
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.BOTH, expand=True)

        y_scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        x_scrollbar = ttk.Scrollbar(frame, orient=tk.HORIZONTAL)
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        treeview = ttk.Treeview(
            frame,
            columns=[col[0] for col in columns],
            show="headings",
            yscrollcommand=y_scrollbar.set,
            xscrollcommand=x_scrollbar.set
        )

        y_scrollbar.config(command=treeview.yview)
        x_scrollbar.config(command=treeview.xview)

        # Configure Columns and Headings
        for col_name, col_type in columns:
            treeview.heading(col_name, text=col_name, anchor=tk.W, command=lambda c=col_name: self._sort_treeview(treeview, c, col_type))
            treeview.column(col_name, anchor=tk.W, width=100)

        # Customize Heading Style
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Helvetica", 12, "bold"))

        treeview.pack(fill=tk.BOTH, expand=True)
        return treeview

    def _create_notebook(self, parent):
        notebook = ttk.Notebook(parent)

        # Tabs: Preview, Hexadecimal, Metadados
        preview_frame = ttk.Frame(notebook)
        hexadecimal_frame = ttk.Frame(notebook)
        metadados_frame = ttk.Frame(notebook)
        hits_frame = ttk.Frame(notebook)

        notebook.add(preview_frame, text="Preview")
        notebook.add(hexadecimal_frame, text="Hexadecimal")
        notebook.add(metadados_frame, text="Metadados")
        notebook.add(hits_frame, text="Hits Relatórios")

        # Placeholder for when the frames are implemented
        ttk.Label(preview_frame, text="Preview Content Goes Here").pack(pady=10)
        ttk.Label(hexadecimal_frame, text="Hexadecimal Content Goes Here").pack(pady=10)
        ttk.Label(metadados_frame, text="Metadados Content Goes Here").pack(pady=10)
        ttk.Label(hits_frame, text="Metadados Content Goes Here").pack(pady=10)

        return notebook

    def _sort_treeview(self, treeview, column, col_type):
        """Sort the Treeview column when heading is clicked."""
        data = [(treeview.set(child, column), child) for child in treeview.get_children()]

        if col_type in ("int", "float"):
            data.sort(key=lambda t: float(t[0]) if t[0].replace('.', '', 1).isdigit() else float('-inf'))
        elif col_type == "date":
            data.sort(key=lambda t: t[0])  # Add date parsing if needed
        else:
            data.sort(key=lambda t: t[0])

        for index, (val, child) in enumerate(data):
            treeview.move(child, "", index)

        # Reverse sorting toggle (not implemented but can be added)

