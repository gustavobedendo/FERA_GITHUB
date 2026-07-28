# -*- coding: utf-8 -*-
"""
Portable IPED launcher.

When a bundle contains one IPED per equipment, this launcher lets the user
choose which case to open. Legacy bundles with a single IPED still open
directly.
"""

import os
import platform
import re
import subprocess
import traceback
import tkinter
from tkinter import messagebox, ttk


TARGET_NAME = "IPED-SearchApp.exe"


def find_iped_exe(root_dir, max_depth=3):
    root_dir = os.path.abspath(root_dir)
    root_depth = root_dir.rstrip(os.sep).count(os.sep)

    for current_dir, dirnames, filenames in os.walk(root_dir):
        current_depth = current_dir.rstrip(os.sep).count(os.sep) - root_depth
        if current_depth >= max_depth:
            dirnames[:] = []

        if TARGET_NAME in filenames:
            return os.path.join(current_dir, TARGET_NAME)

    return None


def equipment_sort_key(case):
    match = re.search(r"(?i)eq\s*([0-9]{1,3})", case["label"])
    if match:
        return int(match.group(1))
    return 9999


def discover_equipment_cases(root_dir):
    iped_root = os.path.join(root_dir, "IPED")
    cases = []
    if not os.path.isdir(iped_root):
        return cases

    for name in os.listdir(iped_root):
        case_root = os.path.join(iped_root, name)
        if not os.path.isdir(case_root):
            continue
        if re.match(r"(?i)^Eq[0-9]{1,3}$", name) is None:
            continue
        exe_path = find_iped_exe(case_root, max_depth=2)
        if exe_path is None:
            continue
        eq_number = re.sub(r"\D", "", name).zfill(2)
        cases.append({
            "label": f"Eq{eq_number}",
            "title": f"Equipamento {eq_number}",
            "path": case_root,
            "exe": exe_path,
        })

    return sorted(cases, key=equipment_sort_key)


def short_path(path, root_dir):
    try:
        return os.path.relpath(path, root_dir)
    except Exception:
        return path


def open_case(case):
    print(case["exe"])
    subprocess.Popen([case["exe"]], cwd=case["path"])


def configure_style(root):
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("Root.TFrame", background="#f4f7fb")
    style.configure("Header.TFrame", background="#172033")
    style.configure("Card.TFrame", background="#ffffff", relief="solid", borderwidth=1)
    style.configure("Title.TLabel", background="#172033", foreground="#ffffff", font=("Segoe UI", 18, "bold"))
    style.configure("Subtitle.TLabel", background="#172033", foreground="#cbd5e1", font=("Segoe UI", 10))
    style.configure("CardTitle.TLabel", background="#ffffff", foreground="#172033", font=("Segoe UI", 13, "bold"))
    style.configure("CardPath.TLabel", background="#ffffff", foreground="#667085", font=("Segoe UI", 9))
    style.configure("Open.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 8))
    style.map("Open.TButton", background=[("active", "#2563eb")], foreground=[("active", "#ffffff")])
    return style


def create_icon(parent, text, background, foreground):
    canvas = tkinter.Canvas(parent, width=44, height=44, highlightthickness=0, background="#ffffff")
    canvas.create_oval(4, 4, 40, 40, fill=background, outline="")
    canvas.create_text(22, 22, text=text, fill=foreground, font=("Segoe UI", 16, "bold"))
    return canvas


def show_selector(cases, root_dir):
    window = tkinter.Tk()
    window.title("Abrir IPED")
    window.geometry("560x%s" % min(720, max(360, 210 + len(cases) * 92)))
    window.minsize(500, 320)
    window.configure(background="#f4f7fb")
    configure_style(window)

    root = ttk.Frame(window, style="Root.TFrame")
    root.pack(fill="both", expand=True)

    header = ttk.Frame(root, style="Header.TFrame", padding=(24, 22, 24, 20))
    header.pack(fill="x")
    ttk.Label(header, text="Escolha o IPED", style="Title.TLabel").pack(anchor="w")
    ttk.Label(
        header,
        text="Este anexo possui um caso IPED por equipamento.",
        style="Subtitle.TLabel",
    ).pack(anchor="w", pady=(6, 0))

    body = ttk.Frame(root, style="Root.TFrame", padding=(20, 18, 20, 20))
    body.pack(fill="both", expand=True)

    def launch_and_close(case):
        try:
            open_case(case)
            window.destroy()
        except Exception as ex:
            traceback.print_exc()
            messagebox.showerror("Erro ao abrir IPED", str(ex), parent=window)

    for index, case in enumerate(cases):
        card = ttk.Frame(body, style="Card.TFrame", padding=(16, 14, 16, 14))
        card.pack(fill="x", pady=(0, 12))
        card.columnconfigure(1, weight=1)

        icon_text = re.sub(r"\D", "", case["label"]) or str(index + 1)
        icon = create_icon(card, icon_text[-2:], "#dbeafe", "#1d4ed8")
        icon.grid(row=0, column=0, rowspan=2, sticky="n", padx=(0, 14))

        ttk.Label(card, text=case["title"], style="CardTitle.TLabel").grid(row=0, column=1, sticky="w")
        ttk.Label(card, text=short_path(case["path"], root_dir), style="CardPath.TLabel").grid(row=1, column=1, sticky="w", pady=(4, 0))
        ttk.Button(card, text="Abrir", style="Open.TButton", command=lambda c=case: launch_and_close(c)).grid(row=0, column=2, rowspan=2, padx=(16, 0))

    footer = ttk.Frame(root, style="Root.TFrame", padding=(20, 0, 20, 18))
    footer.pack(fill="x")
    ttk.Button(footer, text="Cancelar", command=window.destroy).pack(side="right")

    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = int((window.winfo_screenwidth() - width) / 2)
    y = int((window.winfo_screenheight() - height) / 2)
    window.geometry(f"{width}x{height}+{x}+{y}")
    window.mainloop()


def main():
    root_dir = os.getcwd()
    cases = discover_equipment_cases(root_dir)
    if len(cases) == 1:
        open_case(cases[0])
        return
    if len(cases) > 1:
        show_selector(cases, root_dir)
        return

    exe_path = find_iped_exe(root_dir)
    if exe_path is None:
        raise FileNotFoundError(
            "IPED-SearchApp.exe nao encontrado em ate 3 niveis a partir do CWD"
        )

    print(exe_path)
    subprocess.Popen([exe_path], cwd=root_dir)


try:
    main()
except Exception:
    traceback.print_exc()
    if "Windows" in platform.system():
        try:
            messagebox.showerror("Erro ao abrir IPED", traceback.format_exc())
        except Exception:
            pass
