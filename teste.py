import tkinter as tk
from tkinter import ttk

class Tooltip:
    """A simple tooltip implementation."""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        self.widget.bind("<Motion>",self.show_tooltip)
        self.last_id = None
        #self.widget.bind("<Enter>", self.show_tooltip)
        #self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event=None):
        item_id = self.widget.identify_row(event.y)
        
            
        #print(item_id)
        if item_id:
            values = self.widget.item(item_id, "values")
            if values[1] != "Não":
                self.hide_tooltip()
                return
        else:
            self.hide_tooltip()
            return
        if(item_id == self.last_id):
            return 
        self.hide_tooltip()
        x = self.widget.winfo_pointerx() + 20
        y = self.widget.winfo_pointery() + 10

        self.tooltip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")

        label = tk.Label(
            tw, text=self.text, justify="left",
            background="#ffffe0", relief="solid", borderwidth=1,
            font=("tahoma", "10", "normal")
        )
        label.pack(ipadx=1)

    def hide_tooltip(self, event=None):
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None

def create_popup(data):
    def add_tooltip(widget, text):
        Tooltip(widget, text)

    popup = tk.Tk()
    popup.title("FERA - Informações")
    popup.geometry("700x500")
    popup.configure(bg="#f0f0f5")

    # Title Label
    title_label = tk.Label(popup, text="FERA - Informações da Ferramenta", font=("Arial", 16, "bold"), bg="#4a90e2", fg="white")
    title_label.pack(fill="x")

    # List Section
    list_frame = tk.Frame(popup, bg="#ffffff", relief="groove", borderwidth=2)
    list_frame.pack(fill="both", expand=True, padx=10, pady=10)

    columns = ("Equipamento", "Conteúdo Indexado")
    tree = ttk.Treeview(list_frame, columns=columns, show="headings")
    tree.heading("Equipamento", text="Equipamento")
    tree.heading("Conteúdo Indexado", text="Conteúdo Indexado")
    # Scrollbar for Treeview
    tree_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=tree_scrollbar.set)
    tree_scrollbar.pack(side="right", fill="y")
    tree.pack(fill="both", expand=True, padx=10, pady=10)
    
    # Example data
    """ data = [
        ("Equipamento 1", "Sim"),
        ("Equipamento 2", "Não"),
        ("Equipamento 3", "Sim"),
        ("Equipamento 4", "Não"),
        ("Equipamento 1", "Sim"),
        ("Equipamento 2", "Não"),
        ("Equipamento 3", "Sim"),
        ("Equipamento 4", "Não"),
        ("Equipamento 1", "Sim"),
        ("Equipamento 2", "Não"),
        ("Equipamento 3", "Sim"),
        ("Equipamento 4", "Não"),
    ]
 """
    for parent_alias in data:
        tree.insert("", "end", values=(data, data[parent_alias]))
    add_tooltip(tree, "Detalhes: Não conteúdo para indexar (diretório 'files' associado ao equipamento inexistente).")

            

    

    # Textbox Section
    text_frame = tk.Frame(popup, bg="#f9f9f9", relief="groove", borderwidth=2)
    text_frame.pack(fill="both", expand=True, padx=10, pady=10)

    text_label = tk.Label(text_frame, text="O que há de novo?", font=("Arial", 14, "bold"), bg="#f9f9f9")
    text_label.pack(anchor="w", padx=10, pady=5)

    text_box = tk.Text(text_frame, wrap="word", font=("Arial", 10), bg="#ffffff", relief="sunken", borderwidth=2)
    text_scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_box.yview)
    text_box.configure(yscrollcommand=text_scrollbar.set)
    text_scrollbar.pack(side="right", fill="y")
    text_box.pack(fill="both", expand=True, padx=10, pady=10)

    # Prepopulate the textbox with an example
    text_box.insert("1.0", "Bem-vindo à nova atualização!\n\n- Correção de bugs\n- Novas funcionalidades adicionadas")

    popup.mainloop()

#create_popup()