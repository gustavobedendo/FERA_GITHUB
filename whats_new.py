import tkinter as tk
from tkinter import ttk
import global_settings
import os, traceback, sys
import fera_config
texto = " FERA - Forensics Evidence Report Analyzer \n"+\
                    "* License: GNU Affero General Public License v3.0\n\n"+\
                    "STATE DEPARTMENT OF PUBLIC SECURITY -- SCIENTIFIC POLICE OF PARANÁ\n\n"+\
                    "  CODED BY by:\nGustavo Borelli Bedendo <gustavo.bedendo@gmail.com>\n\n"+\
                    "  SUPPORTERS :\nAlexandre Vrubel\nRoger Roberto Rocha Duarte\nWellerson Jeremias Colombari\n\n\n\n"+\
                    "  MAIN TESTERS AND USAGE IDEAS:\nConrado Pinto Rebessi\nJacson Gluzezak\nJoel Koster\nLaercio Silva de Campos Junior\nMarcus Fabio Fontenelle do Carmo\nRaphael Zago\n"+\
                    "\n\nJanuary 2024\n\n"+\
                    "It is a work in progress, the code, \ndespite the ugliness and some bugs, is available on:\n"+\
                    "https://github.com/gustavobedendo/FERA"

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

def on_quit(popup):
    popup.destroy()
    global_settings.popup_whatsnew = None
    
def get_application_path():
    application_path = None
    if getattr(sys, 'frozen', False):
        application_path = sys._MEIPASS
    elif __file__:
        application_path = os.path.dirname(os.path.abspath(__file__))
    return application_path

def create_popup():
    
    def add_tooltip(widget, text):
        Tooltip(widget, text)

    popup = tk.Toplevel(global_settings.root)
    notebook = ttk.Notebook(popup)
    var = tk.IntVar(value=fera_config.get_dontshow())
    cb = ttk.Checkbutton(
        popup,
        text="Não mostrar novamente",
        variable=var,
        command=lambda var=var: fera_config.on_toggle(var)
    )
    notebook.pack(fill="both", expand=True)
    cb.pack(padx=20, pady=20)
    # Tab 1: Fera Index
    tab1 = ttk.Frame(notebook)
    notebook.add(tab1, text="Novidades")
    popup.title("FERA - What's new?")
    popup.geometry("700x700")
    popup.configure(bg="#f0f0f5")

    # Title Label
    #title_label = tk.Label(tab1, text="FERA - Informações da Ferramenta", font=("Arial", 16, "bold"), bg="#4a90e2", fg="white")
    #title_label.pack(fill="x")

               

    

    # Textbox Section
    text_frame = tk.Frame(tab1, bg="#f9f9f9", relief="groove", borderwidth=2)
    text_frame.pack(fill="both", expand=True, padx=10, pady=10)

    text_label = tk.Label(text_frame, text="O que há de novo?", font=("Arial", 14, "bold"), bg="#f9f9f9")
    text_label.pack(anchor="w", padx=10, pady=5)

    text_box = tk.Text(text_frame, wrap="word", font=("Arial", 10), bg="#ffffff", relief="sunken", borderwidth=2)
    text_scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_box.yview)
    text_box.configure(yscrollcommand=text_scrollbar.set)
    text_scrollbar.pack(side="right", fill="y")
    text_box.pack(fill="both", expand=True, padx=10, pady=10)

    # Prepopulate the textbox with an example
    texto = "Bem-vindo à nova atualização!\n\n- Correção de bugs\n- Novas funcionalidades adicionadas"
    whatsnewtxt = os.path.join(get_application_path(), "whats_new.txt")
    try:
        if(os.path.exists(whatsnewtxt)):
            with open(whatsnewtxt, encoding='utf8') as whatsnew:
                texto = whatsnew.read()
    except:
        traceback.print_exc()
    text_box.insert("1.0", texto)
    
    tab2 = ttk.Frame(notebook)
    notebook.add(tab2, text="Equipamentos")
    # List Section
    list_frame = tk.Frame(tab2, bg="#ffffff", relief="groove", borderwidth=2)
    list_frame.pack(fill="both", expand=True, padx=10, pady=10)

    columns = ("Equipamento", "Relatórios Indexados", "Arquivos Indexados")
    tree = ttk.Treeview(list_frame, columns=columns, show="headings")
    tree.heading("Equipamento", text="Equipamento")
    tree.heading("Relatórios Indexados", text="Relatórios Indexados")
    tree.heading("Arquivos Indexados", text="Arquivos Indexados")
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
    for parent_alias in global_settings.info_index:
        tree.insert("", "end", values=(parent_alias, "Sim", global_settings.info_index[parent_alias][0]))
    add_tooltip(tree, "Detalhes: Não conteúdo para indexar (diretório 'files' associado ao equipamento inexistente).")
    
    tab3 = ttk.Frame(notebook)
    notebook.add(tab3, text="Sobre o FERA")
    label = tk.Label(tab3, font=global_settings.Font_tuple_Arial_10, text=texto, image=global_settings.tkphotologo, compound='top')
    label.pack(fill='x', padx=5, pady=5)
    popup.protocol("WM_DELETE_WINDOW", lambda : on_quit(popup))
    return popup
    #popup.mainloop()

#create_popup()