# -*- coding: utf-8 -*-
"""
Created on Fri Oct 30 10:41:59 2020

@author: gustavo.bedendo
"""
from tkinter.filedialog import askopenfilename, asksaveasfilename, askopenfilenames
import tkinter 
from tkinter import ttk
from pathlib import Path
import os
import fitz
import math
import re
from functools import partial
import platform 
import global_settings, utilities_general, classes_general
from process_functions import insertThread
import multiprocessing as mp
from threading import Thread
import time, sys
import traceback


plt = platform.system()

class Processar():
    def __init__(self, idpdf, rel_path_pdf, pdf, init, fim, mt, mb, me, md):
        self.rel_path_pdf = rel_path_pdf
        self.pdf = pdf
        self.idpdf = idpdf
        self.paginit = init
        self.pagfim = fim
        self.me = me
        self.mt = mt
        self.mb = mb
        self.md = md



def sairpdfdif(varok, valor, window):
        varok.set(valor)
        window.destroy()
        
def popup(window, varok, texto = 'Os arquivos não possuem HASH compatível.\n\nDeseja prosseguir?'):
        window.rowconfigure((0,1), weight=1)
        window.columnconfigure((0,1), weight=1)
        w = 400 # width for the Tk root
        h = 200 # height for the Tk root    
        label = tkinter.Label(window, text=texto, image=global_settings.warningimage, compound='top')    
        label.grid(row=0, column=0, sticky='ew', pady=20, columnspan=2)  
        # get screen width and height
        ws = global_settings.root.winfo_screenwidth() # width of the screen
        hs = global_settings.root.winfo_screenheight() # height of the screen        
        # calculate x and y coordinates for the Tk root window
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)
        window.geometry('%dx%d+%d+%d' % (w, h, x, y))        
        button_ok = tkinter.Button(window, font=global_settings.Font_tuple_Arial_8, text="Prosseguir", command= lambda : sairpdfdif(varok, True, window))
        button_ok.grid(row=1, column=1, pady=20) 
        button_cancel = tkinter.Button(window, font=global_settings.Font_tuple_Arial_8, text="Cancelar", command= lambda : sairpdfdif(varok, False, window))
        button_cancel.grid(row=1, column=0, pady=20)         

class App():
    def __init__(self, version, gotoviewer=False):
        self.root = tkinter.Toplevel()
        self.root.bind('<Control-Shift-L>', lambda e: global_settings.log_window.deiconify())
        self.root.protocol("WM_DELETE_WINDOW", lambda : global_settings.on_quit())
        
        self.version = global_settings.version
        self.root.tk.call('wm', 'iconphoto', self.root._w, tkinter.PhotoImage(data=global_settings.icon))
        self.root.geometry("1200x500")
        utilities_general.center(self.root)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.root.title("FERA "+ self.version+" - Forensics Evidence Report Analyzer -- Polícia Científica do Paraná") 
        try:
            self.finalizados = 0
            self.paginasindexadas = 0
            self.globalFrame = tkinter.Frame(self.root)
            self.globalFrame.grid(row=0, column=0, sticky="nsew")
            self.globalFrame.rowconfigure(0, weight=1)
            self.globalFrame.columnconfigure(0, weight=1)
            self.directoriesFrame = tkinter.Frame(self.globalFrame, borderwidth=1, relief='groove')
            self.directoriesFrame.grid(row=0, column=0, sticky="nsew")
            self.directoriesFrame.rowconfigure(0, weight=1)
            self.directoriesFrame.columnconfigure(0, weight=1)            
            self.dirs = ttk.Treeview(self.directoriesFrame, selectmode='extended', columns=('ID', '% Indexado', 'Status', 'Tipo'))
            self.dirs['column']=("zero","one", "two", "three")
            self.dirs.heading("#0", text="Relatorio", anchor="n")
            self.dirs.heading("zero", text="ID", anchor="n")
            self.dirs.heading("one", text="% Indexado", anchor="n")
            self.dirs.heading("two", text="Status", anchor="n")
            self.dirs.heading("three", text="Tipo", anchor="n") 
            self.hscroll = tkinter.Scrollbar(self.directoriesFrame, orient="horizontal")
            self.hscroll.config( command = self.dirs.xview )
            self.dirs.configure(xscrollcommand=self.hscroll.set)
            self.hscroll.grid(row=1, column=0, sticky='ew')            
            self.vscroll = tkinter.Scrollbar(self.directoriesFrame, orient="vertical")
            self.vscroll.config( command = self.dirs.yview )
            self.dirs.configure(yscrollcommand=self.vscroll.set)
            self.vscroll.grid(row=0, column=1, sticky='ns')
            self.dirs.grid(row=0, column=0, sticky="nsew")
            self.dirs.bind("<<TreeviewSelect>>", lambda e: self.treeview_selection(e))
            self.root.update()
            imfecharb = b'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAhElEQVRYhe2WsQ3AIAwEbwr2H4eKJkXGSRojpUiigIEv4pco/fcSYBtCodCzMlCA5PBI5pF7igtwAFtniGS1h3m5DFpDeGrdRsPgPYbD4S3G0+BfANPhb6Bl8LsQu51l8GuICq5BlsHlAaRXIH2E0m8obUTSViwdRvJxLF9I5CtZKPQPnYIqZ80MhoLJAAAAAElFTkSuQmCC'
            imfecharb = b'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'
            self.imfechar = tkinter.PhotoImage(data=imfecharb)
            imopenb = b'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAABkElEQVRYhe3VwUtUURTH8Q9ECPoPtPIfEKKVf0C0FcyiRYxoqdRiaqFSq6BFqxZDLUSDVvoHSAutVbgQg5RBhIJW7VsZWhElr8W90uPpDL4382ZczA8O53Deu7wvv3vuffR0jjWGN7iU6r3Gbs64hoEiAB+Q4F6qtxF7eWIyrssNUcE8rrQBICkCsYkfTjqwhWqDWD4FYCX1fCQPwCimMJQByOtAOt7lAdiJi6oZgE3cLxAfzwowiyV8awBQugNLMTdyoNkMNIvtdgH0HCjdgYVuA/S3AHDYDoBjnRXgFx5jAnMxP8HvTgD8wU18yvR3cQtHZQPUhFOxj2nh6p7BAd5jsWyAiZgf4Wusv+BprCfLBqic8qEj4e5PcLtsgCn8xJowiHU8FOz/LvzKCwNcxR0MNgGoCxOfYA+v/B/IeXxuBeCGMFz9qV4WIME67gp7/zcCjAsutHQMLwjDldZb4bLJxiie4wFe4HqD91bzAOTRtJPD2VEdA9SEbegKwMtYP+sGwDD6Yn0RlzsN0FNh/QMhUGALjgWQ1AAAAABJRU5ErkJggg=='
            imopenb = b'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'
            self.imopen = tkinter.PhotoImage(data=imopenb)
            imupdateb = b'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAACFklEQVRYhdXXTYjNURjH8c/cmWksjJDXkCy8ZCg0KZESGzZsiMJKxE5KUZItNt5JysJC0WQKE7FAKI1SxCgiUfJSxshLGtfiHLlN9/+f+597b5dfnc25v+d5vp1zzznPn39cTWiuVfEGfEA3ZhfM5zAO87AKy9GK4ZUGWIR8HI8xF3vxumC+cPzCLWzFkEoAHEwp9ADtOIpTuNwH7B22oH6gxevwqkjx0xifEjMrgv+I/g4MHQhAa5HieXwU9r8/TcT1GNOFMVkBziUA5HGjxByNOB5j7ggnqmTtSwF4miFPDm0x7nAWAJiGhViCZViBlRibMc9gvEQvWrJCVErrhVW4UCuAerzBTwxLM+7AVWHvKq0TwiqsSzM9i6aK3GR9tDrmPpJkyOE7eoRLpdJaEAHakgyjouFJFYoTTlUet5MMzdHwvEoA82P+9jRTN76qzhasiQCH0kyd0TSligCb00y7o2lbFQBywh+xMc00MwI8UsZbXq4uRoi1tQKYITwcbzEhQ1yDcMZf4CHu467QPV2Lv5esncIqdCr9VmyR/ITnhZe0ZNXhjL8N6eQSYy4lFO81gNasEcdigi/Yo//VGIn3RQA6shYv1CZ8iol6cB4bsBjTMUf4NtiOe0WK57GxHAAYgQP4nFCgcHQJfWPh3OhyAf6oCUuxH2dxE1dwEruEbroOg4RH7ZtwKmqiScKqTa0VwP+h3y5rsNtpU0mxAAAAAElFTkSuQmCC'
            self.imupdate = tkinter.PhotoImage(data=imupdateb)
            imaddb = b'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAS0lEQVRYhWNgoAzMhOIBA2egeNQBow4YdcCoA0YdMOqAoe2AmUgGkYq/QjG5+mcOCgdQAoZHGhh1wKgDRh0w6oBRB4w6gBJAcfccAKlCYd1gjUu5AAAAAElFTkSuQmCC'
            imremoveb = b'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAA4klEQVRYhe2XIQ7CQBBFXwKCpAZRUcNVCIpwjapeoqKiB0DiOEAFpqpcAIHgDpBUcwIQHZJm2TZtGdy85Jvm78xLtmbBzwwIO7IHXh059pybd+zyUvQsmZrzGIENkCtnN0ZgCaTAQSmpzBzMP66gGCNwByqaH0cjlcwczAMonW8LT++zwMXtljJzskACPIHA6Z0kbQLpJpoCOc09Rk7vKmkTSTc3ARMwARMwARMwARMwAU2BTIaGTu8iaRNKN9MUWAGxp7eWuMRyRk3gV0YL3IAavadZLTMHsxVjrWfZQ2Z+8QZsqyF6+vcXyAAAAABJRU5ErkJggg=='
            imbrowseb = b'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAiElEQVRYhe3SOwqAMBCE4R/sPZkHsbDyGLYWXiRH8gJqqRCbBIQYH2CyhTswkG4+woIGOmAF7EU3oE8FuBs/IpLkybhvIQ1oUiDeAL5ocE+5AcE9SQBsDFAhkFOVNMCQ9vt9TQyQ9RYUoAAFKEABClCAOGBxj8mNlxnHZ4AWGIHajQ8Zx1t+nx2V3cDT6bQk5wAAAABJRU5ErkJggg=='
            self.imadd = tkinter.PhotoImage(data=imaddb)
            self.imremove = tkinter.PhotoImage(data=imremoveb)
            self.imbrowse = tkinter.PhotoImage(data=imbrowseb)
            self.continueFrame = tkinter.Frame(self.globalFrame, borderwidth=4, relief='flat', pady=5)
            self.continueFrame.grid(row=2, column=0, sticky="nsew")

            self.continueFrame.rowconfigure((0,1,2), weight=1)
            self.continueFrame.columnconfigure((0,1,2), weight=1)          
            self.blocate = tkinter.Button(self.continueFrame, text="Localizar relatorio(s)", font=global_settings.Font_tuple_Arial_8, command=self.locarel, state='disabled', image=self.imbrowse, compound="right")
            self.blocate.grid(row=1, column=0, sticky="n")           
            self.badd = tkinter.Button(self.continueFrame, text="Adicionar relatorio(s)", font=global_settings.Font_tuple_Arial_8, command=self.addrelsPopup, image=self.imadd, compound="right")
            self.badd.grid(row=1, column=1, sticky="n")          
            self.brem = tkinter.Button(self.continueFrame, text="Remover relatorio", font=global_settings.Font_tuple_Arial_8, command=self.remrels, image=self.imremove, compound="right", state="disabled")
            self.brem.grid(row=1, column=2, sticky="n")          
            
            self.bopenviewer = tkinter.Button(self.continueFrame, text="Abrir", font=global_settings.Font_tuple_ArialBold_10, command=self.abrir, image=self.imopen, compound="top")
            self.bopenviewer.grid(row=2, column=1, sticky="s", pady=(15, 0))           
            
            self.largura = self.directoriesFrame.winfo_width()
            self.dirs.column("#0", width=math.floor(10*(self.largura/20)))
            self.dirs.column("zero", width=math.floor(2*(self.largura/20)))
            self.dirs.column("one", width=math.floor(2*(self.largura/20)))
            self.dirs.column("two", width=math.floor(2*(self.largura/20)))
            self.dirs.column("three", width=math.floor(2*(self.largura/20)))
            
            self.bfechar = tkinter.Button(self.continueFrame, text="Fechar", font=global_settings.Font_tuple_Arial_8, command= lambda : on_quit(self.root), image=self.imfechar, compound="right")
            self.bfechar.grid(row=2, column=3, sticky="se", pady=(15, 0))  
            self.populateEqs()
            
            self.root.after(1000, self.update_indexing_status)
            if(not gotoviewer):
                self.root.wait_window()
            else:
                self.abrir()
            
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        finally:
            None
            
    def abrir(self):
        if(len(self.dirs.get_children(''))==0):
            return
        for relatorio in global_settings.infoLaudo:
            if(global_settings.infoLaudo[relatorio].status=='erro' or global_settings.infoLaudo[relatorio].status=='incompativel'):
                return
        on_quit(self.root, sair = False)
        
       
    def sairpdfdif(self, varok, valor, window):
        varok.set(valor)
        window.destroy()
    
    
    def popup_pdfdif(self, window, varok, texto = 'Os arquivos não possuem HASH compatível.\n\nDeseja prosseguir?'):
        warningb = b'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAChklEQVRYR8XXt44UQRDG8d/hibEiIsC7iICHgAjvrQAhEjwhBAhPBhLeexIkXoMA762EgATxABygknpOc6PdnZm9W5hope2p+tf3VVdPd/nPT1cf8w9K7/9qN067AANxGqvRjVPYgT91QdoFOIkt+J2SBtB+7PsXADPxKCWegx94goCYiM91INpR4B7m4w4Wp2QnsC3ZsrmTAPnqZ+FZSjYGHxBNOaGOCnUVuIsFuI0lhUqPYzvOYFNVFeoAzMDj5H0o8byQJFR4j8GpFz5VgagDEJ4vbFJ9litT4Sw29ifA9NTpsc9DiRdNgo9OvRAqTMLHMoiqCoTni3ALS3NBs8GTj3MsDaXz2NAfAFF9eB9PsfpGAJkKQ5IKsTuaPlUUyKq/iWWFSI0AYslR7MQFrO8LwLTkfcQIJV5WBBiV/A8VJqfd0ZCjTIHwPKbdDSxvEKGZArH0CHbhItY1U6EVQL76+P2qJkCoEP4PxRS8awTRCiA8j2nXrPqI10qBvAqXsLYOwFQ8TQnC+0bVVwEoVaGZAln117GiRReXKRCvHsZuXMaaYqxGAFWrL5sx2f8j044Ylnrhbf7FRgDheUy7a1hZNUvJukPYgyvpM65neREgujXO+JA2Ov91i8BzcS79HyP3QYu1oULsiOEIhd9ka4sA4XlMu6tYVVLVN8QRHM93jC1ZfxB7i7HzADGx4oyP6ntRNglcF2BE6oVQoUfdPEDM7dirVaoPpjoWZDVkKvT0Qh7gC8alE6zHo35qwixMqBDK/UTMCHmAr8nH2XjYz4mzcPNwPwfQnQfIDo8O5e4VNi42W4sKDMCBNP/Hd4giLi1xr4jJGFe6XhZ0KGfrsGXfAx2H+gthzokhxc9aDgAAAABJRU5ErkJggg=='
        warningimage = tkinter.PhotoImage(data=warningb)
        window.rowconfigure((0,1), weight=1)
        window.columnconfigure((0,1), weight=1)
        w = 300 # width for the Tk root
        h = 200 # height for the Tk root    
        label = tkinter.Label(window, text=texto, image=warningimage, compound='top')    
        label.grid(row=0, column=0, sticky='ew', pady=20, columnspan=2)
        # get screen width and height
        ws = self.root.winfo_screenwidth() # width of the screen
        hs = self.root.winfo_screenheight() # height of the screen        
        # calculate x and y coordinates for the Tk root window
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)
        window.geometry('%dx%d+%d+%d' % (w, h, x, y))        
        button_ok = tkinter.Button(window, text="Prosseguir", font=global_settings.Font_tuple_Arial_8, command= lambda : self.sairpdfdif(varok, True, window))
        button_ok.grid(row=1, column=1, pady=20) 
        button_cancel = tkinter.Button(window, text="Cancelar", font=global_settings.Font_tuple_Arial_8, command= lambda : self.sairpdfdif(varok, False, window))
        button_cancel.grid(row=1, column=0, pady=20) 
    
    
    def locarel(self):
        try:
            tipos = [('Arquivo PDF', '*.pdf')]
            path = (askopenfilename(filetypes=tipos, defaultextension=tipos))
            if(path!=None and path!=''):
                
                
                drivepdf = os.path.splitdrive(path)[0]
                drivebd = os.path.splitdrive(global_settings.pathdb)[0]
                if(drivepdf!=drivebd):
                    window = tkinter.Toplevel(global_settings.root)
                    window.protocol("WM_DELETE_WINDOW", lambda : None)
                    window.rowconfigure((0,1), weight=1)
                    window.columnconfigure((0,1,2), weight=1)
                    labelask = tkinter.Label(window, image=global_settings.difdrive, compound="top", text="Não foi possível adicionar o relatório:"+
                                             "\n\n<{}>\n ao banco de dados:\n <{}>\n\n pois ambos se encontram em partições diferentes! ({} e {})"\
                                                 .format(path, global_settings.pathdb, drivepdf, drivebd))
                    labelask.grid(row=0, column=0, columnspan=3, sticky='nse', pady=5)
                    answer0 = tkinter.Button(window, font=global_settings.Font_tuple_Arial_8, text="OK", command=  lambda : window.destroy())
                    answer0.grid(row=1, column=1, sticky='s', pady=5, padx=10)
                    global_settings.root.wait_window(window)
                    raise Exception()
                selecao = self.dirs.focus()
                values = self.dirs.item(selecao, 'values')
                hashpdf = str(utilities_general.md5(path))
                sqliteconn = utilities_general.connectDB(str(global_settings.pathdb))
                cursor = sqliteconn.cursor()
                try:
                    cursor.custom_execute("PRAGMA foreign_keys = ON")
                    cursor.custom_execute('''SELECT P.rel_path_pdf, P.indexado, P.id_pdf, P.hash FROM Anexo_Eletronico_Pdfs P WHERE P.id_pdf = ? ''', (values[0],))
                    record = cursor.fetchone()
                    relpathpdf_old = record[0]
                    pathpdf_old = utilities_general.get_normalized_path(os.path.join(global_settings.pathdb.parent, relpathpdf_old))                    
                    relpathpdf = os.path.relpath(path, global_settings.pathdb.parent)
                    pathpdf_new = utilities_general.get_normalized_path(os.path.join(global_settings.pathdb.parent, relpathpdf))
                    name = os.path.basename(relpathpdf)
                    if(record[3]==hashpdf):
                        name = os.path.basename(relpathpdf)
                        cursor.custom_execute("UPDATE Anexo_Eletronico_Pdfs set rel_path_pdf = ? WHERE id_pdf = ?", (relpathpdf, record[2],))
                        parentpath = Path(path).parent
                        
                        try:
                            if(not self.dirs.exists(str(parentpath))):
                                self.dirs.insert('', index='end', iid=str(parentpath), text=str(parentpath), values=("","","",""))
                        except Exception as ex:
                            utilities_general.printlogexception(ex=ex)
                        try:
                            relative_path = os.path.relpath(path, parentpath)
                            self.dirs.item('idpdf'+str(values[0]), text=str(relative_path))
                            self.dirs.item('idpdf'+str(values[0]), values=(values[0],"100%","OK", values[3], path))
                            self.dirs.see('idpdf'+str(values[0]))
                            #self.dirs.delete(selecao)
                        except Exception as ex:
                            utilities_general.printlogexception(ex=ex)
                        
                        pathpdf = utilities_general.get_normalized_path(pathpdf_new)
                        #class RelatorioSuccint:
                        #    def __init__(self, idpdf, toc, lenpdf, pixorgw, pixorgh, mt, mb, me, md, paginasindexadas, rel_path_pdf, abs_path_pdf, tipo):
                        relatorio_proxy = classes_general.RelatorioSuccint(values[0], global_settings.infoLaudo[pathpdf_old].toc,\
                                                                           global_settings.infoLaudo[pathpdf_old].len,\
                                                                           global_settings.infoLaudo[pathpdf_old].pixorgw,\
                                                                           global_settings.infoLaudo[pathpdf_old].pixorgh,\
                                                                           global_settings.infoLaudo[pathpdf_old].mt,\
                                                                           global_settings.infoLaudo[pathpdf_old].mb,\
                                                                           global_settings.infoLaudo[pathpdf_old].me,\
                                                                           global_settings.infoLaudo[pathpdf_old].md,\
                                                                           global_settings.infoLaudo[pathpdf_old].len,\
                                                                           relpathpdf, pathpdf, global_settings.infoLaudo[pathpdf_old].tipo)
                                                                                                             
                        global_settings.infoLaudo[pathpdf]=global_settings.infoLaudo[pathpdf_old]
                        global_settings.listaRELS[pathpdf]=relatorio_proxy
                        
                        global_settings.listaRELS[pathpdf].status="indexado"
                        global_settings.infoLaudo[pathpdf].status="indexado"
                        global_settings.infoLaudo[pathpdf].paginasindexadas=global_settings.infoLaudo[pathpdf_old].len
                        
                        global_settings.infoLaudo[pathpdf].rel_path_pdf = relpathpdf
                        
                        del global_settings.infoLaudo[pathpdf_old]
                        del global_settings.listaRELS[pathpdf_old]
                        sqliteconn.commit()
                    else:
                        varok = tkinter.BooleanVar()
                        varok.set(False)
                        window = tkinter.Toplevel()
                        window.protocol("WM_DELETE_WINDOW", disable_event)
                        self.popup_pdfdif(window, varok)
                        self.root.wait_window(window)
                        
                        if(varok.get()):
                            global_settings.send_vacuum = True
                            doc = fitz.open(pathpdf_new)
                            try:
                                document_margin = classes_general.Document_Margin(doc)     
                                if(not document_margin.marginsok):
                                    return
                                cursor.custom_execute("UPDATE Anexo_Eletronico_Pdfs set rel_path_pdf = ?, indexado = ?, hash = ?, doclen = ?  WHERE id_pdf = ?", \
                                                      (relpathpdf, 0,'', len(doc), record[2],))
                                cursor.custom_execute("DELETE FROM Anexo_Eletronico_Tocs WHERE id_pdf = ?", (record[2],))
                                cursor.custom_execute("UPDATE Anexo_Eletronico_Obsitens set status = 'alterado'  WHERE id_pdf = ?", (record[2],))
                                parentpath = Path(pathpdf_new).parent
                                idpdf = record[2]
                                relp = Path(utilities_general.get_normalized_path(os.path.join(global_settings.pathdb.parent, str(relpathpdf))))
                                relpdir = relp.parent
                                pixorg = doc[0].get_pixmap()
                                mmtopxtop = math.floor(document_margin.mt/25.4*72)
                                mmtopxbottom = math.ceil(pixorg.height-(document_margin.mb/25.4*72))
                                mmtopxleft = math.floor(document_margin.me/25.4*72)
                                mmtopxright = math.ceil(pixorg.width-(document_margin.md/25.4*72))
                                
                                global_settings.infoLaudo[pathpdf_new] = classes_general.Relatorio() 
                                global_settings.infoLaudo[pathpdf_new].mt = document_margin.mt
                                global_settings.infoLaudo[pathpdf_new].mb = document_margin.mb
                                global_settings.infoLaudo[pathpdf_new].me = document_margin.me
                                global_settings.infoLaudo[pathpdf_new].md = document_margin.md
                                global_settings.infoLaudo[pathpdf_new].rel_path_pdf = relpathpdf
                                global_settings.infoLaudo[pathpdf_new].hash = ''
                                global_settings.infoLaudo[pathpdf_new].status = 'naoindexado'
                                global_settings.documents_to_index.append(pathpdf_new)
                                global_settings.infoLaudo[pathpdf_new].id = idpdf
                                global_settings.infoLaudo[pathpdf_new].len = len(doc)
                                global_settings.infoLaudo[pathpdf_new].paginasindexadas = 0
                                global_settings.infoLaudo[pathpdf_new].pixorgw = pixorg.width
                                global_settings.infoLaudo[pathpdf_new].pixorgh = pixorg.height
                                #T.toc_unit, T.pagina, T.deslocy, T.init
                                nameddests = {}
                                try:
                                    nameddests = grabNamedDestinations(doc)
                                except Exception as ex:
                                    utilities_general.printlogexception(ex=ex)
                                    nameddests = {}
                                toc = doc.get_toc(simple=False)
                                listatocs = []
                                for entrada in toc:
                                    novotexto = ""
                                    init = 0
                                    tocunit = entrada[1]
                                    pagina = None
                                    deslocy = None
                                    init = 0
                                    if('page' in entrada[3]):
                                        pagina = entrada[3]['page'] 
                                        deslocy = entrada[3]['to'].y
                                    elif('file' in entrada[3]):
                                        arquivocomdest = entrada[3]['file'].split("#")
                                        arquivo = arquivocomdest[0]
                                        dest = arquivocomdest[1]
                                        if(os.path.basename(pathpdf_new)==arquivo):
                                            if(dest in nameddests):
                                            #for sec in nameddests:
                                                #if(dest==sec[0]):
                                                pagina = int(nameddests[dest][0])
                                                deslocy = pixorg.height-round(float(nameddests[dest][2]))
                                                #break
                                    if(pagina==None):
                                        None
                                        continue
                                    extrair_texto = utilities_general.extract_text_from_page(doc, pagina, deslocy, mmtopxtop, mmtopxbottom, mmtopxleft, mmtopxright)
                                    init = extrair_texto[0]
                                    novotexto = extrair_texto[1]
                                    listatocs.append((entrada[1], idpdf, pagina, deslocy, init,))
                                for toc in listatocs:
                                    global_settings.infoLaudo[pathpdf_new].toc.append((toc[0], int(toc[2]), int(toc[3]), int(toc[4])))
                                global_settings.infoLaudo[pathpdf_new].ultimaPosicao=0.0
                                global_settings.infoLaudo[pathpdf_new].tipo = global_settings.infoLaudo[pathpdf_old].tipo
                                
                                global_settings.infoLaudo[pathpdf_new].parent_alias = global_settings.infoLaudo[pathpdf_old].parent_alias
                                global_settings.infoLaudo[pathpdf_new].zoom_pos = global_settings.infoLaudo[pathpdf_old].zoom_pos
                                #class RelatorioSuccint:
                                #    def __init__(self, idpdf, toc, lenpdf, pixorgw, pixorgh, mt, mb, me, md, paginasindexadas, rel_path_pdf, abs_path_pdf, tipo):
                                relatorio_proxy = classes_general.RelatorioSuccint(idpdf, global_settings.infoLaudo[pathpdf_new].toc, global_settings.infoLaudo[pathpdf_new].len, \
                                                                                   pixorg.width, pixorg.height, document_margin.mt, document_margin.mb, document_margin.me, \
                                                                                       document_margin.md, 0, \
                                                                                       relpathpdf, pathpdf_new, global_settings.infoLaudo[pathpdf_old].tipo)   
                                global_settings.listaRELS[pathpdf_new] = relatorio_proxy
                                if(pathpdf_old!=pathpdf_new):
                                    del global_settings.infoLaudo[pathpdf_old]
                                    del global_settings.listaRELS[pathpdf_old]
                                if pathpdf_old in global_settings.documents_to_index:
                                    global_settings.documents_to_index.remove(pathpdf_old)
                                insert_query_toc = """INSERT INTO Anexo_Eletronico_Tocs
                                        (toc_unit, id_pdf , pagina, deslocy, init) VALUES
                                        (?,?,?,?,?)
                                """
                                cursor.custom_executemany(insert_query_toc, listatocs)
                            except Exception as ex:
                                utilities_general.printlogexception(ex=ex)
                            finally:
                                doc.close()
                            
                            try:
                                if(not self.dirs.exists(str(parentpath))):
                                    self.dirs.insert('', index='end', iid=str(parentpath), text=str(parentpath), values=("","","",""))
                            except Exception as ex:
                                utilities_general.printlogexception(ex=ex)
                            try:
                                idpdfitem = 'idpdf'+str(record[2])
                                values = self.dirs.item(idpdfitem, 'values')
                                self.dirs.item(idpdfitem, values=(values[0], '0%', 'Indexando', values[3], values[4]))
                                #self.dirs.insert(str(parentpath), index='end', iid='idpdf'+str(values[0]), text=str(relative_path), values=(values[0],"0%","Indexando", values[3], path))
                                #self.dirs.delete(selecao)
                                self.dirs.see(idpdfitem)
                            except Exception as ex:
                                utilities_general.printlogexception(ex=ex)
                            cursor.custom_execute("DELETE FROM Anexo_Eletronico_SearchResults where id_pdf = ?", (record[2],))                
                            check_previous_search =  "SELECT DISTINCT C.termo, C.tipobusca, C.id_termo, C.fixo, C.pesquisado  FROM Anexo_Eletronico_SearchTerms C ORDER by 3"
                            cursor.custom_execute(check_previous_search)
                            termosbuscados = cursor.fetchall()
                            for termo in termosbuscados:
                                id_termo = termo[2]
                                pesquisados = termo[4].replace("({})".format(record[2]), "")
                                updateinto2 = "UPDATE Anexo_Eletronico_SearchTerms set pesquisado = ? WHERE id_termo = ?"
                                cursor.custom_execute(updateinto2, (pesquisados, id_termo))
                            sqliteconn.commit()
                            global_settings.documents_to_index.append(pathpdf_new)
                            print(global_settings.infoLaudo)
                            utilities_general.initiate_indexing_thread()

                            
                except Exception as ex:
                    None
                    utilities_general.printlogexception(ex=ex)
                finally:
                    cursor.close()
                    
                    if(sqliteconn):
                        sqliteconn.close()
        except Exception as ex:
            None
            utilities_general.printlogexception(ex=ex)
    
    def treeview_selection(self, event=None):
        try:
            selecao = self.dirs.focus()
            #
            selecaoiid = self.dirs.selection()[0]
            #
            #
            if(len(self.dirs.get_children(selecao))==0 and len(self.dirs.selection())==1):
                values = self.dirs.item(selecao, 'values')
                #
                if(values[2]=='PDF não encontrado' or values[2]=='Hash incompatível'):
                    self.blocate.config(state='normal')
                else:
                    self.blocate.config(state='disabled')
            else:
                self.blocate.config(state='disabled')
            mesmopai = self.dirs.parent(selecaoiid)
            podehabilitar=True
            for selec in self.dirs.selection():
                if(self.dirs.parent(selec)!=mesmopai):
                    podehabilitar = False
                    break
            if(podehabilitar):
                self.brem.config(state='normal')
            else:
                self.brem.config(state='disabled')
        except Exception as ex:
            self.blocate.config(state='disabled')
            
    def update_indexing_status(self):
        try:
            tempo = 1000
            if(not global_settings.processados.empty()):
                tempo= 10
                proc = global_settings.processados.get(0)
                #print(proc[0])
                if(proc[0]=='update'):
                    idpdf = global_settings.infoLaudo[proc[1]].id
                    global_settings.infoLaudo[proc[1]].paginasindexadas += proc[2]
                    global_settings.infoLaudo[proc[1]].paginasindexadas = min(global_settings.infoLaudo[proc[1]].paginasindexadas, global_settings.infoLaudo[proc[1]].len)
                    #None
                    porcent = round(global_settings.infoLaudo[proc[1]].paginasindexadas / global_settings.infoLaudo[proc[1]].len * 100)
                    idpdfitem = 'idpdf'+str(idpdf)
                    values = self.dirs.item(idpdfitem, 'values')
                    self.dirs.item(idpdfitem, values=(values[0], str(porcent)+'%', values[2], values[3], values[4]))
                if(proc[0]=='saving'):
                    idpdf = proc[2]
                    idpdfitem = 'idpdf'+str(idpdf)
                    values = self.dirs.item(idpdfitem, 'values')
                    self.dirs.item(idpdfitem, values=(values[0], "Finalizando...", values[2], values[3], values[4]))
                elif(proc[0]=='ok'):
                    idpdf = global_settings.infoLaudo[proc[2]].id
                    idpdfitem = 'idpdf'+str(idpdf)
                    values = self.dirs.item(idpdfitem, 'values')
                    #idpdf,"100%",'OK', tipo, relatorio
                    self.dirs.item(idpdfitem, values=(idpdf,"100%","OK",values[3],values[4]))
                elif(proc[0]=='running_im'):
                    idpdf = global_settings.infoLaudo[proc[1]].id
                    percentage = proc[3]
                    idpdfitem = 'idpdf'+str(idpdf)
                    values = self.dirs.item(idpdfitem, 'values')
                    #idpdf,"100%",'OK', tipo, relatorio
                    self.dirs.item(idpdfitem, values=(idpdf,"100%","Image_match {}".format(percentage),values[3],values[4]))

        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
            None
        finally:
            self.root.after(tempo, self.update_indexing_status)
    
    def populateEqs(self):
        self.progressinit = ttk.Progressbar(self.continueFrame, mode='indeterminate')
        self.bopenviewer.config(state='disabled')
        self.progressinit.grid(row=0, column=1, sticky='nsew', pady=10)
        relpathsdir = set()
        qtos = 0
        self.progressinit['maximum'] = len(global_settings.infoLaudo)
        self.progressinit['mode'] = 'determinate'
        self.progressinit.start()
        for relatorio in global_settings.infoLaudo:
            self.progressinit['value'] = qtos
            qtos += 1
            self.progressinit.update()
            idpdf = global_settings.infoLaudo[relatorio].id
            relp = Path(relatorio)
            relpdir = relp.parent
            tipo = global_settings.infoLaudo[relatorio].tipo
            if(not self.dirs.exists(relpdir)):
                self.dirs.insert('', index='end', iid=relpdir, text=str(relpdir), values=("","","","","","","",""))
            relative_path = os.path.relpath(relatorio, relpdir)
            if(global_settings.infoLaudo[relatorio].status=='erro'):
                self.dirs.insert(relpdir, index='end', iid='idpdf'+str(idpdf), text=str(relative_path), values=(idpdf, "-","PDF não encontrado", tipo, relatorio))
                self.bopenviewer.config(state='disabled')                        
            else:
                if(global_settings.infoLaudo[relatorio].status=='naoindexado'):  
                    self.dirs.insert(relpdir, index='end', iid='idpdf'+str(idpdf), text=str(relative_path), values=(idpdf,"0%","Indexando",tipo, relatorio))                             
                else:
                    if(global_settings.infoLaudo[relatorio].status=='incompativel'):
                        self.dirs.insert(relpdir, index='end', iid='idpdf'+str(idpdf), text=str(relative_path), values=(idpdf, '-',"Hash incompatível",tipo, relatorio))
                        self.bopenviewer.config(state='disabled')
                    else:
                        self.dirs.insert(relpdir, index='end', iid='idpdf'+str(idpdf), text=str(relative_path), values=(idpdf,"100%",'OK', tipo, relatorio))
            
            self.dirs.see('idpdf'+str(idpdf))
            
        self.bopenviewer.config(state='normal')
        self.progressinit.grid_forget()

    def addrelsPopup(self):   
        try:
            self.opcao = tkinter.Menu(self.root, tearoff=0)
            self.opcao.add_command(label='Laudo', command=partial(addrels,'laudo', self))
            self.opcao.add_command(label='Relatorio', command=partial(addrels,'relatorio',  self))
            self.opcao.add_command(label='Outros', command=partial(addrels,'outros', self))
            self.opcao.tk_popup(self.badd.winfo_rootx(),self.badd.winfo_rooty())         
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
            None
        finally:
            self.opcao.grab_release() 
        
    def remrels(self):
        sqliteconn = utilities_general.connectDB(str(global_settings.pathdb))
        cursor = sqliteconn.cursor()
        try:
            for selection in self.dirs.selection():
                texto = self.dirs.item(selection, 'text')
                value = self.dirs.item(selection, 'values')[0]
                cursor.custom_execute("PRAGMA foreign_keys = ON")
                cursor.custom_execute("SELECT P.id_pdf, P.rel_path_pdf, P.indexado FROM Anexo_Eletronico_Pdfs P where :pdf = P.id_pdf ",{'pdf': value})
                record = cursor.fetchone()
                relpathpdf = record[1]
                relp = Path(utilities_general.get_normalized_path(os.path.join(global_settings.pathdb.parent, str(relpathpdf))))
                abs_path_pdf = utilities_general.get_normalized_path(relp)
                cursor.custom_execute("DROP TABLE IF EXISTS Anexo_Eletronico_Conteudo_id_pdf_"+str(record[0]))
                #cursor.custom_execute("DROP TABLE IF EXISTS Anexo_Eletronico_Conteudo_id_pdf_"+str(record[0]))
                cursor.custom_execute("DROP TABLE IF EXISTS Anexo_Eletronico_Conteudo_id_pdf_"+str(record[0])+"_config")
                cursor.custom_execute("DROP TABLE IF EXISTS Anexo_Eletronico_Conteudo_id_pdf_"+str(record[0])+"_content")
                cursor.custom_execute("DROP TABLE IF EXISTS Anexo_Eletronico_Conteudo_id_pdf_"+str(record[0])+"_data")
                cursor.custom_execute("DROP TABLE IF EXISTS Anexo_Eletronico_Conteudo_id_pdf_"+str(record[0])+"_docsize")
                cursor.custom_execute("DROP TABLE IF EXISTS Anexo_Eletronico_Conteudo_id_pdf_"+str(record[0])+"_idx")
                cursor.custom_execute("DELETE FROM Anexo_Eletronico_Pdfs where id_pdf = ?", (record[0],))
                cursor.custom_execute("DELETE FROM Anexo_Eletronico_SearchResults where id_pdf = ?", (record[0],))
                
                check_previous_search =  "SELECT DISTINCT C.termo, C.tipobusca, C.id_termo, C.fixo, C.pesquisado  FROM Anexo_Eletronico_SearchTerms C ORDER by 3"
                cursor.custom_execute(check_previous_search)
                termosbuscados = cursor.fetchall()
                for termo in termosbuscados:
                    id_termo = termo[2]
                    pesquisados = termo[4].replace("({})".format(record[0]),'')
                    updateinto2 = "UPDATE Anexo_Eletronico_SearchTerms set pesquisado = ? WHERE id_termo = ?"
                    cursor.custom_execute(updateinto2, (pesquisados, id_termo))
                del global_settings.infoLaudo[abs_path_pdf]
                del global_settings.listaRELS[abs_path_pdf]
                if abs_path_pdf in global_settings.documents_to_index:
                    global_settings.documents_to_index.remove(abs_path_pdf)
            for selection in self.dirs.selection():
                self.dirs.delete(selection)
            global_settings.send_vacuum = True
            sqliteconn.commit()

        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        finally:
            cursor.close()
           # None
            if(sqliteconn):
                sqliteconn.close()
          



def on_quit(root, sair=True):
    if(not sair):
        root.destroy()
    else:        
        global_settings.on_quit()

        
def popup_window(toplevel, texto):
    def popupcomandook1(window):
        window.destroy()
    window = tkinter.Toplevel()
    window.protocol("WM_DELETE_WINDOW", disable_event)
    label = tkinter.Label(window, text=texto, image=global_settings.warningimage, compound='top')
    label.pack(fill='x', padx=50, pady=20)
    button_close = tkinter.Button(window, font=global_settings.Font_tuple_Arial_8, text="OK", command= lambda : popupcomandook1(window))
    button_close.pack(fill='y', pady=20)   
    return window
 


def iterateXREF_NamedDests(doc, xref, pismm, p3, p4, pnotmm, xreftopage, listax):
    chaves = doc.xref_get_keys(xref)
    if("Names" in chaves):
        names = doc.xref_get_key(xref, "Names")[1]
        
        namesdest = p3.findall(names)
        #if("sec" in names):
        #print(names)
        for namedd, xr, num, letra  in namesdest:
            
            #keys = doc.xref_get_keys(int(xr))
            folhas = p4.findall(doc.xref_object(int(xr)))
            
            for pageref, x, y in folhas:
                #print(folhas, namedd, xreftopage[int(pageref)])
                listax[namedd] = (xreftopage[int(pageref)], x, y)

    elif("Kids" in chaves):
        destinations_kids = doc.xref_get_key(xref, "Kids")
        destinations_limits = doc.xref_get_key(xref, "Limits")
        quaislimites = pismm.findall(destinations_limits[1])
        splitted = destinations_kids[1].split(" ")
        grauavore = int(len(splitted)/3)
        for i in range(grauavore):
            if('null'==destinations_limits[0]):
                indice = i * 3
                novoxref = int(splitted[indice].replace("[", "").replace("]", ""))
                iterateXREF_NamedDests(doc, novoxref, pismm, p3, p4, pnotmm, xreftopage, listax)
            elif(len(quaislimites)>1):
                if(('sec' >= quaislimites[0][:3].lower() and 'sec' <= quaislimites[1][:3].lower()) or \
                   ('sub' >= quaislimites[0][:3].lower() and 'sub' <= quaislimites[1][:3].lower())):
                    #print(quaislimites)
                    indice = i * 3
                    novoxref = int(splitted[indice].replace("[", "").replace("]", ""))
                    iterateXREF_NamedDests(doc, novoxref, pismm, p3, p4, pnotmm, xreftopage, listax)
                else:
                    None
                    #print(quaislimites)
            elif(len(quaislimites)>0):
                if('sec' >= quaislimites[0][:3].lower() or 'sub' >= quaislimites[0][:3].lower()):
                    #print(quaislimites)
                    indice = i * 3
                    novoxref = int(splitted[indice].replace("[", "").replace("]", ""))
                    iterateXREF_NamedDests(doc, novoxref, pismm, p3, p4, pnotmm, xreftopage, listax)
                else:
                    None
                    #print(quaislimites)
            else:
                None
                #print(quaislimites)
    else:
        None
        #print(chaves)
        '''
        if(len(destinations_limits)>1):
            quaislimites = pnotmm.findall(destinations_limits[1])
            if(len(quaislimites)>0 or 'null'==destinations_limits[0]):
                splitted = destinations_kids[1].split(" ")
                grauavore = int(len(splitted)/3)
                for i in range(grauavore):
                    indice = i * 3
                    novoxref = int(splitted[indice].replace("[", "").replace("]", ""))
                    iterateXREF_NamedDests(doc, novoxref, p1, p2, p3, p4, pnotmm, xreftopage, listax)
        '''
    #return listax

def iteratetreepages( xreftopage, doc, numberregex, xref, count):
    objrootpages = doc.xref_get_key(int(xref), "Type")[1]
    if(objrootpages=="/Pages"):
        objrootkids = doc.xref_get_key(int(xref), "Kids")[1]
        for indobj, gen in numberregex.findall(objrootkids):
           count = iteratetreepages(xreftopage, doc, numberregex, indobj, count)  
        #return count
    elif(objrootpages=="/Page"):
        xreftopage[int(xref)] = count
        count += 1
    
    return count

def loadPages(xreftopage, doc, numberregex):
    rootpdf  = doc.pdf_catalog()
    objpagesr = numberregex.findall(doc.xref_get_key(rootpdf, "Pages")[1])[0][0]
    objrootpages = doc.xref_get_key(int(objpagesr), "Type")[1]
    if(objrootpages=="/Pages"):
       objrootkids = doc.xref_get_key(int(objpagesr), "Kids")[1] 
       count = 0
       for indobj, gen in numberregex.findall(objrootkids):
           count = iteratetreepages(xreftopage, doc, numberregex, indobj, count)

def grabNamedDestinations(doc):
    xreftopage = {}
    numbercompile = re.compile(r"([0-9]+)\s([0-9]+)")
    loadPages(xreftopage, doc, numbercompile)
    #
    listnameddestinations = []
    #for p in range(len(doc)):
    #    xref = doc.page_xref(p)
    #    xreftopage[xref] = p
    regexismm = r"\(([a-zA-Z0-9_\.\-]+)\)"
    pismm = re.compile(regexismm)
    regex2= r"\/F\([-_\.a-zA-Z0-9]+\)"
    regex1= r"([0-9]+)(\s[0-9]+\s)([A-Z])"
    regex3= r"\(([subsection|subsubsection|section]+\*?\.[a-zA-Z0-9_\.\-]+)\)([0-9]+)(\s[0-9]+\s)([A-Z])"
    regex3 = r"\(([^\(\)]+)\)([0-9]+)(\s[0-9]+\s)([A-Z])"
    regexnotmm = r"\(([subsection|subsubsection|section]+\*?\.[a-zA-Z0-9_\.\-]+)\)"
    regexlimits = r"\[\(([^\)]*)\)\(([^\)]*)\)\]"
    regex4 = r"\[\s+([0-9]+)\s+0\s+R\s+/XYZ\s+([0-9\.]+)\s+([0-9\.]+)"
    p1 = None
    p2 = None
    p3 = re.compile(regex3)
    p4 = re.compile(regex4)
    pnotmm = re.compile(regexlimits)
    pagexref = doc.page_xref(1)
    
    
    key = doc.xref_get_key(doc.pdf_catalog(), "Names")
    lista = {}
    try:
        dests = doc.xref_get_key(int(key[1].split(" ")[0]), "Dests")
        #print(dests)
        if("xref" in key[0]):
            iterateXREF_NamedDests(doc, int(dests[1].split(" ")[0]),pismm, p3, p4, pnotmm, xreftopage, lista)
        else:    
            #print(dests)
            None
    except Exception as ex:
        #utilities_general.printlogexception(ex=ex)
        traceback.print_exc()
        
    
        #sys.exit(0)
    #print(lista)
    return lista
        
def popupcomandook(window, datatop, databottom, dataleft, dataright):
        marginsok = True
        mt = datatop.get()
        mb = databottom.get()
        me = dataleft.get()
        md = dataright.get()
        window.destroy()

       
def popupcomandocancel(window):
    marginsok = False
    window.destroy()


def re_fn(expr, item):
    reg = re.compile(expr, re.I)
    return reg.search(item) is not None    

        
def disable_event():
    pass


def addrel_commandLine(patpdf):
    goon = True
    def update_indexing_status():
        print()
        try:
            while goon:
                if(not global_settings.processados.empty()):
                    tempo= 10
                    proc = global_settings.processados.get(0)
                    
                    if(proc[0].strip()=='update'):
                        #print(proc[0])
                        idpdf = global_settings.infoLaudo[proc[1]].id
                        global_settings.infoLaudo[proc[1]].paginasindexadas += proc[2]
                        global_settings.infoLaudo[proc[1]].paginasindexadas = min(global_settings.infoLaudo[proc[1]].paginasindexadas, global_settings.infoLaudo[proc[1]].len)
                        #None
                        porcent = round(global_settings.infoLaudo[proc[1]].paginasindexadas / global_settings.infoLaudo[proc[1]].len * 100)
                        idpdfitem = 'idpdf'+str(idpdf)
                        #values = self.dirs.item(idpdfitem, 'values')
                        print(f"{patpdf} Indexado: {porcent}%")
                        sys.stdout.flush()
                        #self.dirs.item(idpdfitem, values=(values[0], str(porcent)+'%', values[2], values[3], values[4]))
                    elif(proc[0].strip()=='saving'):
                        print()
                        idpdf = proc[2]
                        idpdfitem = 'idpdf'+str(idpdf)
                        #values = self.dirs.item(idpdfitem, 'values')
                        #self.dirs.item(idpdfitem, values=(values[0], "Finalizando...", values[2], values[3], values[4]))
                        print(f"{patpdf} Indexado: Finalizando")
                    elif(proc[0].strip()=='ok'):
                        print()
                        idpdf = global_settings.infoLaudo[proc[2]].id
                        idpdfitem = 'idpdf'+str(idpdf)
                        #values = self.dirs.item(idpdfitem, 'values')
                        #idpdf,"100%",'OK', tipo, relatorio
                        #self.dirs.item(idpdfitem, values=(idpdf,"100%","OK",values[3],values[4]))
                        print(f"{patpdf} Indexado: OK")
                        return
                    else:
                        print(proc)
                #time.sleep(5)
        except:
            traceback.print_exc()
    try:
        insert_query_pdf = """INSERT INTO Anexo_Eletronico_Pdfs
                    (rel_path_pdf , indexado, tipo, lastpos, margemsup, margeminf, margemesq, margemdir,
                     pixorgw, pixorgh, doclen, parent_alias, zoom_pos) VALUES
                    (?,?,?, '0.0', ?,?,?,?,?,?,?,?,?)
        """
        pdf = None

        doc = fitz.open(patpdf)
        pixorg = doc[0].get_pixmap()
        doclen = len(doc)
        #doc.close()
        pdf = (patpdf, global_settings.default_margin_top, global_settings.default_margin_bottom, \
               global_settings.default_margin_left, global_settings.default_margin_right, pixorg, doclen, '', 0)
            
            
        mt = pdf[1]
        mb = pdf[2]
        me = pdf[3]
        md = pdf[4]
        pixorg = pdf[5]
        doclen = pdf[6]
        #return
        sqliteconn = None
        try:
            sqliteconn = utilities_general.connectDB(str(global_settings.pathdb))
            relpathpdf = os.path.relpath(patpdf, global_settings.pathdb.parent)
            tipo = 'relatorio'
            cursor = sqliteconn.cursor()
            print(f"Inserting {patpdf} to DB -->")
            cursor.custom_execute(insert_query_pdf, (relpathpdf, 0,tipo, mt, mb, me, md, int(pixorg.width), int(pixorg.height), doclen, '', 0))
            print(f"Inserting {patpdf} to DB --> OK")
            mmtopxtop = math.floor(mt/25.4*72)
            mmtopxbottom = math.ceil(pixorg.height-(mb/25.4*72))
            mmtopxleft = math.floor(me/25.4*72)
            mmtopxright = math.ceil(pixorg.width-(md/25.4*72))
            idpdf = cursor.lastrowid
            relp = Path(utilities_general.get_normalized_path(os.path.join(global_settings.pathdb.parent, str(relpathpdf))))
            relpdir = relp.parent
            pathpdf2 = str(patpdf)
            pathpdf2 = utilities_general.get_normalized_path(pathpdf2)
            abs_path_pdf = pathpdf2
            try:
                nameddests = grabNamedDestinations(doc)
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)
                nameddests = {}
            
            toc = doc.get_toc(simple=False)
            listatocs = []
            for entrada in toc:
                novotexto = ""
                init = 0
                tocunit = entrada[1]
                pagina = None
                deslocy = None
                init = 0
                if('page' in entrada[3]):
                    pagina = entrada[3]['page'] 
                    deslocy = entrada[3]['to'].y
                elif('file' in entrada[3]):
                    arquivocomdest = entrada[3]['file'].split("#")
                    arquivo = arquivocomdest[0]
                    dest = arquivocomdest[1]
                    if(os.path.basename(pathpdf2)==arquivo):
                        if(dest in nameddests):
                        #for sec in nameddests:
                            #if(dest==sec[0]):
                            pagina = int(nameddests[dest][0])
                            deslocy = pixorg.height-round(float(nameddests[dest][2]))
                            #break
                if(pagina==None):
                    None
                    continue
                extrair_texto = utilities_general.extract_text_from_page(doc, pagina, deslocy, mmtopxtop, mmtopxbottom, mmtopxleft, mmtopxright)
                init = extrair_texto[0]
                novotexto = extrair_texto[1]
                
                listatocs.append((entrada[1], idpdf, pagina, deslocy, init,))
            print(f"Inserting TOC {patpdf} to DB -->")
            insert_query_toc = """INSERT INTO Anexo_Eletronico_Tocs
                    (toc_unit, id_pdf , pagina, deslocy, init) VALUES
                    (?,?,?,?,?)
            """
            cursor.custom_executemany(insert_query_toc, listatocs)
            print(f"Inserting TOC {patpdf} to DB --> OK")
            #sqliteconn.commit()
            
            cursor.custom_execute("DROP TABLE IF EXISTS Anexo_Eletronico_Conteudo_id_pdf_"+str(idpdf), None, False, False)
            cursor.custom_execute("DROP TABLE IF EXISTS Anexo_Eletronico_Conteudo_id_pdf_"+str(idpdf)+"_config", None, False, False)
            cursor.custom_execute("DROP TABLE IF EXISTS Anexo_Eletronico_Conteudo_id_pdf_"+str(idpdf)+"_content", None, False, False)
            cursor.custom_execute("DROP TABLE IF EXISTS Anexo_Eletronico_Conteudo_id_pdf_"+str(idpdf)+"_data", None, False, False)
            cursor.custom_execute("DROP TABLE IF EXISTS Anexo_Eletronico_Conteudo_id_pdf_"+str(idpdf)+"_docsize", None, False, False)
            cursor.custom_execute("DROP TABLE IF EXISTS Anexo_Eletronico_Conteudo_id_pdf_"+str(idpdf)+"_idx", None, False, False)
            cursor.custom_execute("DROP TABLE IF EXISTS Anexo_Eletronico_Images_id_pdf_"+str(idpdf), None, False, False)
            create_table_content = 'CREATE VIRTUAL TABLE Anexo_Eletronico_Conteudo_id_pdf_'+str(idpdf)+' USING fts4(texto, pdf_id UNINDEXED, pagina' ')'
            cursor.custom_execute(create_table_content)
            
            create_table_images = '''CREATE TABLE Anexo_Eletronico_Images_id_pdf_{} (
            id_image INTEGER PRIMARY KEY AUTOINCREMENT,
            hash_image TEXT NOT NULL,
            bbox_x0 INTEGER NOT NULL,
            bbox_y0 INTEGER NOT NULL,
            bbox_x1 INTEGER NOT NULL,
            bbox_y1 INTEGER NOT NULL,   
            pagina INTEGER NOT NULL
            )
            '''
            cursor.custom_execute(create_table_images.format(str(idpdf)))
            sqliteconn.commit()
            fim = 0  
            
            for i in range(global_settings.nthreads):
                init = fim
                fim = math.ceil((i+1) * (len(doc)/global_settings.nthreads))
                #__init__(self, idpdf, rel_path_pdf, pdf, init, fim, mt, mb, me, md)
                proc = Processar(idpdf, abs_path_pdf, os.path.basename(abs_path_pdf), init, min(fim, len(doc)), mt, mb, me, md)
                #print()
                global_settings.processar.put(proc)
            #return 
            #print(global_settings.processar)
            insert_content= global_settings.manager.list()
            insert_content_images= global_settings.manager.list()   
            global_settings.infoLaudo[pathpdf2] = classes_general.Relatorio() 
            global_settings.infoLaudo[pathpdf2].mt = mt
            global_settings.infoLaudo[pathpdf2].mb = mb
            global_settings.infoLaudo[pathpdf2].me = me
            global_settings.infoLaudo[pathpdf2].md = md
            global_settings.infoLaudo[pathpdf2].rel_path_pdf = relpathpdf
            global_settings.infoLaudo[pathpdf2].hash = ''
            global_settings.infoLaudo[pathpdf2].status = 'naoindexado'
            #global_settings.documents_to_index.append(abs_path_pdf)
            global_settings.infoLaudo[pathpdf2].id = idpdf
            global_settings.infoLaudo[pathpdf2].len = len(doc)
            global_settings.infoLaudo[pathpdf2].parent_alias = ''
           
            global_settings.infoLaudo[pathpdf2].zoom_pos = 0
            global_settings.infoLaudo[pathpdf2].pixorgw = pixorg.width
            global_settings.infoLaudo[pathpdf2].pixorgh = pixorg.height
            #global_settings.infoLaudo[pathpdf2].paginasindexadas = 0
            #T.toc_unit, T.pagina, T.deslocy, T.init
            for toc in listatocs:
                global_settings.infoLaudo[pathpdf2].toc.append((toc[0], int(toc[2]), int(toc[3]), int(toc[4])))
            global_settings.infoLaudo[pathpdf2].ultimaPosicao=0.0
            global_settings.infoLaudo[pathpdf2].tipo = tipo
            #class RelatorioSuccint:
            #    def __init__(self, idpdf, toc, lenpdf, pixorgw, pixorgh, mt, mb, me, md, paginasindexadas, rel_path_pdf, abs_path_pdf, tipo):
            relatorio_proxy = classes_general.RelatorioSuccint(idpdf, global_settings.infoLaudo[pathpdf2].toc, global_settings.infoLaudo[pathpdf2].len, \
                                                               pixorg.width, pixorg.height, mt, mb, me, md, 0, \
                                                                   relpathpdf, pathpdf2, tipo)   
            global_settings.listaRELS[pathpdf2] = relatorio_proxy
            for i in range(global_settings.nthreads):
                
                #insertThread(processar, processados, listaRELS, pathdb, inserts, status)
                global_settings.processing_threads[i] = mp.Process(target=insertThread, name = "FERA INDEXING PROCESS {}".format(i+1),
                                                                   args=(global_settings.processar, global_settings.processados, \
                                                                                              global_settings.listaRELS, \
                                                                                              global_settings.pathdb,insert_content, \
                                                                                                  insert_content_images,None,), daemon=True)
                global_settings.processing_threads[i].start()
            cancommit = True
            new_thread = Thread(target=update_indexing_status)
            new_thread.start()
            
            

            for i in range(global_settings.nthreads):
                None
                global_settings.processing_threads[i].join()
            print()
            if(len(insert_content)==doclen):
                goon = False
                new_thread.join()
                sqliteconn = utilities_general.connectDB(str(global_settings.pathdb))
                cursor = sqliteconn.cursor()
                hashpdf = str(utilities_general.md5(abs_path_pdf))
                cursor.custom_execute("UPDATE Anexo_Eletronico_Pdfs set indexado = 1, hash = ? WHERE id_pdf = ?", (hashpdf, idpdf,))
                print(f"Indexing processes finalized {abs_path_pdf} - saving to database -->")
                sql_insert_content = "INSERT INTO Anexo_Eletronico_Conteudo_id_pdf_" + str(idpdf) +\
                                    " (texto, pagina) VALUES (?,?)"
                pdfsql = 'Anexo_Eletronico_Conteudo_id_pdf_'+str(idpdf)
                #sql_insert_content2 = "INSERT INTO {}({}) VALUES (?)".format(pdfsql,pdfsql)
                #insert_content.put('STOP')
                #insert_contents=dump_queue(insert_content)
                
                sql_insert_content_images = "INSERT INTO Anexo_Eletronico_Images_id_pdf_" + str(idpdf) +\
                                    " (hash_image, bbox_x0, bbox_y0, bbox_x1, bbox_y1, pagina) VALUES (?,?,?,?,?,?)"
                pdfsql_images = 'Anexo_Eletronico_Images_id_pdf_'+str(idpdf)
                cursor.custom_executemany(sql_insert_content, insert_content)
                sqliteconn.commit()
                print(f"Indexing processes finalized {abs_path_pdf} - saving to database --> OK DONE")
                return True
            return False
            
            
        except:
            traceback.print_exc()
            return False
        finally:
            goon = False
            if(sqliteconn):
                sqliteconn.close()
        
        
    except:
        return False


def addrels(tipo, view=None, pathpdfinput = None, pathdbext=None, rootx=None, sqliteconnx=None):
    
    idpdfs = []

    if(pathpdfinput==None):
        pathpdfs = askopenfilenames(filetypes=(("PDF", "*.pdf"), ("Todos os arquivos", "*")))
    else:
        pathpdfs = pathpdfinput
        #global_settings.pathdb = Path(pathdbext) 
    if(pathpdfs!=None and pathpdfs!=''):
        global_settings.send_vacuum = True
        sqliteconn = None
        
        if(sqliteconnx==None):
            sqliteconn = utilities_general.connectDB(str(global_settings.pathdb))
        else:
            sqliteconn = sqliteconnx
        cursor = sqliteconn.cursor()
        try:
            global_settings.askformargins = True
            for patpdf in pathpdfs:
                pathpdf = Path(patpdf)
                
                drivepdf = os.path.splitdrive(pathpdf)[0]
                drivebd = os.path.splitdrive(global_settings.pathdb)[0]
                if(drivepdf!=drivebd):
                    window = tkinter.Toplevel(global_settings.root)
                    window.protocol("WM_DELETE_WINDOW", lambda : None)
                    window.rowconfigure((0,1), weight=1)
                    window.columnconfigure((0,1,2), weight=1)
                    labelask = tkinter.Label(window, image=global_settings.difdrive, compound="top", text="Não foi possível adicionar o relatório:"+
                                             "\n\n<{}>\n ao banco de dados:\n <{}>\n\n pois ambos se encontram em partições diferentes! ({} e {})"\
                                                 .format(pathpdf, global_settings.pathdb, drivepdf, drivebd))
                    labelask.grid(row=0, column=0, columnspan=3, sticky='nse', pady=5)
                    answer0 = tkinter.Button(window, font=global_settings.Font_tuple_Arial_8, text="OK", command=  lambda : window.destroy())
                    answer0.grid(row=1, column=1, sticky='s', pady=5, padx=10)
                    global_settings.root.wait_window(window)
                    raise Exception()
                
                relpathpdf = os.path.relpath(pathpdf, global_settings.pathdb.parent)
                filename, file_extension = os.path.splitext(patpdf)
                abs_path_pdf = utilities_general.get_normalized_path(patpdf)
                if(abs_path_pdf in global_settings.infoLaudo):
                    continue
                global_settings.infoLaudo[abs_path_pdf] = classes_general.Relatorio()
                if(file_extension.lower()==".pdf"):
                    cursor.custom_execute("PRAGMA foreign_keys = ON")
                    cursor.custom_execute('''SELECT P.id_pdf, P.indexado, P.tipo, P.margemsup, P.margeminf, P.margemesq, P.margemdir, 
                                   P.lastpos, P.hash, P.pixorgw, P.pixorgh, P.doclen, P.parent_alias, P.zoom_pos FROM
                                   Anexo_Eletronico_Pdfs P where :pdf = P.rel_path_pdf ''',{'pdf': str(relpathpdf)})
                    r = cursor.fetchone()
                    idpdf = None
                    if(r!=None):
                        if(not os.path.exists(abs_path_pdf)):
                            global_settings.infoLaudo[abs_path_pdf].status = 'erro'
                        else:
                            if(r[8]=='' or r[8]==None):
                                global_settings.infoLaudo[abs_path_pdf].status = 'naoindexado'
                                global_settings.documents_to_index.append(abs_path_pdf)
                            else:
                                hashpdf = str(utilities_general.md5(abs_path_pdf))
                                if(hashpdf.lower()!=r[8].lower()):
                                    print(abs_path_pdf)
                                    print(hashpdf.lower())
                                    global_settings.infoLaudo[abs_path_pdf].status = 'incompativel'
                                else:
                                    global_settings.infoLaudo[abs_path_pdf].status = 'indexado'
                        doclen = None
                        if(r[9]!=None):
                            global_settings.infoLaudo[abs_path_pdf].pixorgw = r[9]
                            global_settings.infoLaudo[abs_path_pdf].pixorgh = r[10]
                            doclen = r[11]
                            global_settings.infoLaudo[abs_path_pdf].len = doclen
                        else:
                            doc = fitz.open(patpdf)
                            pixorg = doc[0].get_pixmap()
                            global_settings.infoLaudo[abs_path_pdf].pixorgw = pixorg.width
                            global_settings.infoLaudo[abs_path_pdf].pixorgh = pixorg.height
                            doclen = len(doc)
                            global_settings.infoLaudo[abs_path_pdf].len = doclen
                            doc.close()
                            updateinto2 = "UPDATE Anexo_Eletronico_Pdfs set pixorgw = ?, pixorgh= ?, doclen = ? WHERE id_pdf = ?"
                            cursor.custom_execute(updateinto2, (int(pixorg.width), int(pixorg.height), doclen, r[0],))
                            sqliteconn.commit()
                        
                        global_settings.infoLaudo[abs_path_pdf].parent_alias = r[12]
                        #print('v', global_settings.infoLaudo[abs_path_pdf].parent_alias)
                        global_settings.infoLaudo[abs_path_pdf].zoom_pos = r[13]
                        global_settings.infoLaudo[abs_path_pdf].mt = r[3]
                        global_settings.infoLaudo[abs_path_pdf].mb = r[4]
                        global_settings.infoLaudo[abs_path_pdf].me = r[5]
                        global_settings.infoLaudo[abs_path_pdf].md = r[6]
                        global_settings.infoLaudo[abs_path_pdf].tipo = r[2]
                        global_settings.infoLaudo[abs_path_pdf].rel_path_pdf = relpathpdf
                        global_settings.infoLaudo[abs_path_pdf].hash = r[8]
                        global_settings.infoLaudo[abs_path_pdf].id = r[0]
                        idpdf = r[0]
                        
                        select_tocs = '''SELECT  T.toc_unit, T.pagina, T.deslocy, T.init FROM 
                        Anexo_Eletronico_Tocs T WHERE T.id_pdf = ? ORDER BY 2,3
                        '''              
                        cursor.custom_execute(select_tocs, (r[0],))
                        tocs = cursor.fetchall()
                        for toc in tocs:
                            global_settings.infoLaudo[abs_path_pdf].toc.append((toc[0], int(toc[1]), int(toc[2]), int(toc[3])))
                        global_settings.infoLaudo[abs_path_pdf].ultimaPosicao=float(r[7])
                        paginasindexadas = 0
                        if(global_settings.infoLaudo[abs_path_pdf].status == 'indexado'):
                            paginasindexadas = doclen
                        global_settings.infoLaudo[abs_path_pdf].paginasindexadas = paginasindexadas
                          
                        #class RelatorioSuccint:
                        #    def __init__(self, idpdf, toc, lenpdf, pixorgw, pixorgh, mt, mb, me, md, paginasindexadas, rel_path_pdf, abs_path_pdf, tipo):
                        relatorio_proxy = classes_general.RelatorioSuccint(r[0], global_settings.infoLaudo[abs_path_pdf].toc, \
                                                                           global_settings.infoLaudo[abs_path_pdf].len, \
                                                                           global_settings.infoLaudo[abs_path_pdf].pixorgw, \
                                                                               global_settings.infoLaudo[abs_path_pdf].pixorgh, \
                                                                                   r[3], r[4], r[5], r[6], paginasindexadas, \
                                                                               relpathpdf, abs_path_pdf, r[2])   
                        global_settings.listaRELS[abs_path_pdf] = relatorio_proxy

                    else:
                        insert_query_pdf = """INSERT INTO Anexo_Eletronico_Pdfs
                                    (rel_path_pdf , indexado, tipo, lastpos, margemsup, margeminf, margemesq, margemdir,
                                     pixorgw, pixorgh, doclen, parent_alias, zoom_pos) VALUES
                                    (?,?,?, '0.0', ?,?,?,?,?,?,?,?,?)
                        """
                        pdf = None
                        if(global_settings.askformargins):
                            margin_pdf_window_class = classes_general.Ask_for_margins(patpdf)
                            global_settings.root.wait_window(margin_pdf_window_class.window)
                            pdf = margin_pdf_window_class.info_pdf_margin
                        else:
                            doc = fitz.open(patpdf)
                            pixorg = doc[0].get_pixmap()
                            doclen = len(doc)
                            doc.close()
                            pdf = (patpdf, global_settings.default_margin_top, global_settings.default_margin_bottom, \
                                   global_settings.default_margin_left, global_settings.default_margin_right, pixorg, doclen, '', 0)
                        if(pdf==None):
                            print("ERRO")
                            continue
                        mt = pdf[1]
                        mb = pdf[2]
                        me = pdf[3]
                        md = pdf[4]
                        pixorg = pdf[5]
                        doclen = pdf[6]
                        cursor.custom_execute(insert_query_pdf, (relpathpdf, 0,tipo, mt, mb, me, md, int(pixorg.width), int(pixorg.height), doclen, '', 0))
                        mmtopxtop = math.floor(mt/25.4*72)
                        mmtopxbottom = math.ceil(pixorg.height-(mb/25.4*72))
                        mmtopxleft = math.floor(me/25.4*72)
                        mmtopxright = math.ceil(pixorg.width-(md/25.4*72))
                        idpdf = cursor.lastrowid
                        relp = Path(utilities_general.get_normalized_path(os.path.join(global_settings.pathdb.parent, str(relpathpdf))))
                        relpdir = relp.parent
                        pathpdf2 = str(pathpdf)
                        pathpdf2 = utilities_general.get_normalized_path(pathpdf2)
                        doc = fitz.open(pathpdf2)
                        idpdfs.append((idpdf, len(doc)))
                        try:
                            nameddests = {}
                            try:
                                nameddests = grabNamedDestinations(doc)
                            except Exception as ex:
                                traceback.print_exc()
                                utilities_general.printlogexception(ex=ex)
                                
                            toc = doc.get_toc(simple=False)
                            listatocs = []
                            for entrada in toc:
                                novotexto = ""
                                init = 0
                                tocunit = entrada[1]
                                pagina = None
                                deslocy = None
                                init = 0
                                if('page' in entrada[3]):
                                    pagina = entrada[3]['page'] 
                                    deslocy = entrada[3]['to'].y
                                elif('file' in entrada[3]):
                                    arquivocomdest = entrada[3]['file'].split("#")
                                    arquivo = arquivocomdest[0]
                                    dest = arquivocomdest[1]
                                    if(os.path.basename(pathpdf2)==arquivo):
                                        if(dest in nameddests):
                                        #for sec in nameddests:
                                            #if(dest==sec[0]):
                                            pagina = int(nameddests[dest][0])
                                            deslocy = pixorg.height-round(float(nameddests[dest][2]))
                                            #break
                                if(pagina==None):
                                    None
                                    continue
                                extrair_texto = utilities_general.extract_text_from_page(doc, pagina, deslocy, mmtopxtop, mmtopxbottom, mmtopxleft, mmtopxright)
                                init = extrair_texto[0]
                                novotexto = extrair_texto[1]
                                
                                listatocs.append((entrada[1], idpdf, pagina, deslocy, init,))
                            insert_query_toc = """INSERT INTO Anexo_Eletronico_Tocs
                                    (toc_unit, id_pdf , pagina, deslocy, init) VALUES
                                    (?,?,?,?,?)
                            """
                            cursor.custom_executemany(insert_query_toc, listatocs)
                            global_settings.infoLaudo[pathpdf2] = classes_general.Relatorio() 
                            global_settings.infoLaudo[pathpdf2].mt = mt
                            global_settings.infoLaudo[pathpdf2].mb = mb
                            global_settings.infoLaudo[pathpdf2].me = me
                            global_settings.infoLaudo[pathpdf2].md = md
                            global_settings.infoLaudo[pathpdf2].rel_path_pdf = relpathpdf
                            global_settings.infoLaudo[pathpdf2].hash = ''
                            global_settings.infoLaudo[pathpdf2].status = 'naoindexado'
                            global_settings.documents_to_index.append(abs_path_pdf)
                            global_settings.infoLaudo[pathpdf2].id = idpdf
                            global_settings.infoLaudo[pathpdf2].len = len(doc)
                            global_settings.infoLaudo[pathpdf2].parent_alias = ''
                           
                            global_settings.infoLaudo[pathpdf2].zoom_pos = 0
                            global_settings.infoLaudo[pathpdf2].pixorgw = pixorg.width
                            global_settings.infoLaudo[pathpdf2].pixorgh = pixorg.height
                            #global_settings.infoLaudo[pathpdf2].paginasindexadas = 0
                            #T.toc_unit, T.pagina, T.deslocy, T.init
                            for toc in listatocs:
                                global_settings.infoLaudo[pathpdf2].toc.append((toc[0], int(toc[2]), int(toc[3]), int(toc[4])))
                            global_settings.infoLaudo[pathpdf2].ultimaPosicao=0.0
                            global_settings.infoLaudo[pathpdf2].tipo = tipo
                            #class RelatorioSuccint:
                            #    def __init__(self, idpdf, toc, lenpdf, pixorgw, pixorgh, mt, mb, me, md, paginasindexadas, rel_path_pdf, abs_path_pdf, tipo):
                            relatorio_proxy = classes_general.RelatorioSuccint(idpdf, global_settings.infoLaudo[pathpdf2].toc, global_settings.infoLaudo[pathpdf2].len, \
                                                                               pixorg.width, pixorg.height, mt, mb, me, md, 0, \
                                                                                   relpathpdf, pathpdf2, tipo)   
                            global_settings.listaRELS[pathpdf2] = relatorio_proxy
                            
                                                       
                            
                        except Exception as ex:
                            utilities_general.printlogexception(ex=ex)
                            None
                            global_settings.on_quit()
                        finally:
                            None
                    if(view!=None): 
                        try:
                            if(not view.dirs.exists(str(relpdir))):
                                view.dirs.insert('', index='end', iid=str(relpdir), text=str(relpdir), values=("","","","","","","",""))
                        except Exception as ex:
                            None
                            utilities_general.printlogexception(ex=ex)
                            None
                        try:
                            relative_path = os.path.relpath(pathpdf2, relpdir)
                            idpdfitem =  'idpdf'+str(idpdf)
                            view.dirs.insert(relpdir, index='end', iid=idpdfitem, text=str(relp), values=(idpdf,"0%","-","-","-", "-", tipo, pathpdf2))
                            view.dirs.see(idpdfitem)
                        except Exception as ex:
                            None
                            utilities_general.printlogexception(ex=ex)  
                            None
            sqliteconn.commit()
            utilities_general.initiate_indexing_thread()

        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        finally:
            
            if(sqliteconnx==None):                
                if(sqliteconn):
                    sqliteconn.close()
    
    return idpdfs




def validate_annotation(sqliteconn, cursor):
    resulttable = '''SELECT name FROM sqlite_master WHERE type="table" AND name="Anexo_Eletronico_Annotations"'''
    cursor.custom_execute(resulttable)
    tableannotacount = cursor.fetchone()
    if(tableannotacount==None):
        create_table_annotations = '''CREATE TABLE Anexo_Eletronico_Annotations (
        id_annot INTEGER PRIMARY KEY AUTOINCREMENT,
        id_obs INTEGER NOT NULL,
        id_pdf INTEGER NOT NULL,
        paginainit INTEGER NOT NULL,
        p0x INTEGER NOT NULL,
        p0y INTEGER NOT NULL,
        paginafim INTEGER NOT NULL,
        p1x INTEGER NOT NULL,
        p1y INTEGER NOT NULL,
        link TEXT DEFAULT '',
        conteudo TEXT DEFAULT '',
        CONSTRAINT fk_obs
            FOREIGN KEY (id_obs)
                REFERENCES Anexo_Eletronico_Obsitens (id_obs)
                ON DELETE CASCADE,
        CONSTRAINT fk_pdf                    
            FOREIGN KEY (id_pdf)
                REFERENCES Anexo_Eletronico_Pdfs (id_pdf)
                ON DELETE CASCADE
        )
        '''
        doc = None
        cursor.custom_execute(create_table_annotations)  
        try:
            updateannots = '''SELECT P.rel_path_pdf, O.paginainit, O.p0x, O.p0y, O.paginafim, O.p1x, O.p1y, O.tipo, O.id_obs, O.fixo, O.status, 
            O.conteudo, O.arquivo, P.id_pdf, O.id_obs FROM Anexo_Eletronico_Obsitens O, 
            Anexo_Eletronico_Pdfs P  WHERE
                O.id_pdf  = P.id_pdf ORDER by 1'''
            cursor.custom_execute(updateannots)
            obsitens = cursor.fetchall()
            
            pathpdfatual_local = None
            for obsitem in obsitens:
                paginainit = obsitem[1]
                p0x = obsitem[2]
                p0y = obsitem[3]
                paginafim = obsitem[4]
                p1x = obsitem[5]
                p1y = obsitem[6]
                tipo = obsitem[7]
                relpath = obsitem[0]
                status = obsitem[10]
                conteudo = obsitem[11]
                idpdf = obsitem[13]
                arquivo = obsitem[12]
                idobs = obsitem[14]
                ident = ' '
                pathpdf = utilities_general.get_normalized_path(os.path.join(global_settings.pathdb.parent, relpath))

                if(pathpdf!=pathpdfatual_local):
                    pathpdfatual_local = pathpdf
                    doc = fitz.open(pathpdfatual_local)
                
                
                insert_annot = '''INSERT INTO Anexo_Eletronico_Annotations
                                        (id_pdf, id_obs, paginainit, p0x, p0y, paginafim, p1x, p1y, link, conteudo) VALUES
                                        (?,?,?,?,?,?,?,?,?,?)'''
                p0x = obsitem[2]
                p0y = obsitem[3]
                p1x = obsitem[5]
                p1y = obsitem[6]
                #extract_links_from_page(doc, idpdf, idobs, pathpdf, paginainit, paginafim, p0x, p0y, p1x, p1y)
                links_tratados = utilities_general.extract_links_from_page(doc, idpdf, idobs, pathpdf, paginainit, paginafim, p0x, p0y, p1x, p1y)
                cursor.custom_execute("PRAGMA journal_mode=WAL")
                cursor.custom_executemany(insert_annot, links_tratados)
  
                                
            sqliteconn.commit()
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        finally:
            try:
                doc.close()
            except:
                None
def update_db_version(sqliteconn, cursor):
    updateinto2 = "UPDATE FERA_CONFIG set param = ? WHERE config = ?"
    cursor.custom_execute(updateinto2, (global_settings.dbversion,'dbversion',))
    sqliteconn.commit()

def necessity_to_validate(cursor):
    try:
        select_query = """SELECT config, param FROM FERA_CONFIG """
        cursor.custom_execute(select_query)    
        records = cursor.fetchall()
        nodbversion = True
        actualdbversion = None
        for conf in records:
            if(conf[0]=='dbversion'):
                nodbversion = False
                actualdbversion = conf[1]
                try:
                    actualdbversion = float(actualdbversion)
                    if(actualdbversion < float(global_settings.dbversion)):
                        return True
                    else:
                        return False
                except:
                    actualdbversion = "1.0"
                    updateinto2 = "UPDATE FERA_CONFIG set param = ? WHERE config = ?"
                    cursor.custom_execute(updateinto2, (actualdbversion,'dbversion',))
                    return True
        return True
    except Exception as ex:
        utilities_general.printlogexception(ex=ex)

def validate_new_db_columns(cursor, must_commit=False):
       
    commit = must_commit  
    try:
        cursor.custom_execute("ALTER TABLE Anexo_Eletronico_Obsitens ADD COLUMN withalt INTEGER DEFAULT 0", None, False, False)
        commit = True
    except Exception as ex:
        None
    try:
        cursor.custom_execute("ALTER TABLE Anexo_Eletronico_Obsitens ADD COLUMN conteudo TEXT DEFAULT ''", None, False, False)
        commit = True
    except Exception as ex:
        None
    try:
        cursor.custom_execute("ALTER TABLE Anexo_Eletronico_Obsitens ADD COLUMN arquivo TEXT DEFAULT ''", None, False, False)
        commit = True
    except Exception as ex:
        None
    try:
       cursor.custom_execute('ALTER TABLE Anexo_Eletronico_SearchTerms ADD COLUMN pesquisado', None, False, False)
       commit = True
    except Exception as ex:
        None
    resulttable = '''SELECT name FROM sqlite_master WHERE type="table" AND name="Anexo_Eletronico_SearchResults"'''
    cursor.custom_execute(resulttable)
    tableresultcount = cursor.fetchone()
    if(tableresultcount==None):
        create_table_searchesresults = '''CREATE TABLE Anexo_Eletronico_SearchResults (
        id_termo INTEGER NOT NULL,
        id_pdf INTEGER NOT NULL,
        pagina INTEGER NOT NULL,
        init INTEGER NOT NULL,
        fim INTEGER NOT NULL,
        toc TEXT,
        snippetantes TEXT,
        snippetdepois TEXT,
        termo TEXT,
        CONSTRAINT fk_termo
            FOREIGN KEY (id_termo)
                REFERENCES Anexo_Eletronico_SearchTerms (id_termo)
                ON DELETE CASCADE,
        CONSTRAINT fk_pdf
        FOREIGN KEY (id_pdf)
            REFERENCES Anexo_Eletronico_Pdfs (id_pdf)
            ON DELETE CASCADE
        )
        '''
        cursor.custom_execute(create_table_searchesresults)  
        commit = True  

        #commit = True 
    try:
        addcolumn = "ALTER TABLE Anexo_Eletronico_Obscat ADD COLUMN ordem INTEGER NOT NULL DEFAULT 0"
        cursor.custom_execute(addcolumn, None, False, False)
        commit = True  
        obscats = "SELECT id_obscat FROM Anexo_Eletronico_Obscat"        
        cursor.custom_execute(obscats)
        obscats = cursor.fetchall()
        ordem = 0
        for obscat in obscats:
            updateinto2 = "UPDATE Anexo_Eletronico_Obscat set ordem = ? WHERE id_obscat = ?"
            cursor.custom_execute(updateinto2, (ordem, obscat[0],))
            ordem += 1
    except:
        None
    try:
        addcolumn2 = "ALTER TABLE Anexo_Eletronico_Obsitens ADD COLUMN conteudo TEXT"
        cursor.custom_execute(addcolumn2, None, False, False)
        commit = True    
    except Exception as ex:
        None
    
    try:
        addcolumn2 = "ALTER TABLE Anexo_Eletronico_Pdfs ADD COLUMN pixorgw INTEGER"
        cursor.custom_execute(addcolumn2, None, False, False)
        commit = True    
    except Exception as ex:
        None
    try:
        addcolumn2 = "ALTER TABLE Anexo_Eletronico_Pdfs ADD COLUMN pixorgh INTEGER"
        cursor.custom_execute(addcolumn2, None, False, False)
        commit = True    
    except Exception as ex:
        None
    try:
        addcolumn2 = "ALTER TABLE Anexo_Eletronico_Pdfs ADD COLUMN doclen INTEGER"
        cursor.custom_execute(addcolumn2, None, False, False)
        commit = True    
    except Exception as ex:
        None
    try:
        addcolumn2 = "ALTER TABLE Anexo_Eletronico_Pdfs ADD COLUMN parent_alias TEXT DEFAULT ''"
        cursor.custom_execute(addcolumn2, None, False, False)
        commit = True    
    except Exception as ex:
        None
    try:
        addcolumn2 = "ALTER TABLE Anexo_Eletronico_Pdfs ADD COLUMN zoom_pos INTEGER DEFAULT 0"
        cursor.custom_execute(addcolumn2, None, False, False)
        commit = True    
    except Exception as ex:
        None    
    
    
        
    return commit
 
def create_edit_recent_cases():
    try:
        count = 0
        add = []
        add.append(str(global_settings.pathdb)+"\n")
        try:
            with open(global_settings.recenttxt, 'r', encoding="utf-8") as f:                   
               for line in f:              
                   if(os.path.isfile(line.strip()) and  line not in add):
                       add.append(line)
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        with open(global_settings.recenttxt, 'w', encoding="utf-8") as f:    
            for linha in add:                    
                f.write(linha)
    except IOError as ex:
        utilities_general.printlogexception(ex=ex)
    except Exception as ex:
        utilities_general.printlogexception(ex=ex)
    
def loaddb(toplevel, event=None, lb1=None, addrel = None):
    try:
        if(event!=None):
            try:
                widget = event.widget
                selection=widget.curselection()
                value = widget.get(selection[0]).strip()
                pp =value                
                global_settings.pathdb = Path(pp)
            except Exception as ex:
                None
        elif(lb1!=None and lb1.curselection()!=()):
            try:
                selection=lb1.curselection()
                value = lb1.get(selection[0]).strip()
                pp =value
                global_settings.pathdb = Path(pp)
            except Exception as ex:
                None            
        elif(global_settings.pathdb==None):
            global_settings.pathdb = Path(askopenfilename(filetypes=(("SQLite", "*.db"), ("Todos os arquivos", "*"))))
            if(os.path.isfile(global_settings.pathdb)):               
                None
            else:                
                global_settings.pathdb = None
        if(os.path.isfile(global_settings.pathdb)):
            if(os.path.exists(str(global_settings.pathdb)+'.lock')):
                window = popup_window(toplevel, "O banco de dados aparentemente está aberto em outra execução!\n\n"+\
                             "Favor fechar execuções utilizando o mesmo banco de dados para evitar inconsistências!")
                global_settings.root.wait_window(window)
                try:
                    os.remove(str(global_settings.pathdb)+'.lock')
                except:
                    None
            else:
                with open(str(global_settings.pathdb)+'.lock', 'w') as fp:
                    fp.write(str(os.getpid()))
                    
                    #pass
            create_edit_recent_cases()
            tocommit = False
            sqliteconn = utilities_general.connectDB(str(global_settings.pathdb))
            try:
                cursor = sqliteconn.cursor()
                must_validate = necessity_to_validate(cursor)
                validate_annotation(sqliteconn, cursor)
                if(must_validate):
                    tocommit = validate_new_db_columns(cursor, must_validate)
                    update_db_version(sqliteconn, cursor)
                if(isinstance(toplevel, tkinter.Toplevel)):                   
                   toplevel.destroy()                
            except Exception as ex:
                popup_window(toplevel, "Erro na abertura do Banco de Dados")
                utilities_general.printlogexception(ex=ex)
                global_settings.pathdb=None
            finally:
                if(tocommit):
                    sqliteconn.commit()
                if(sqliteconn):
                    sqliteconn.close()
        else:
            global_settings.pathdb = None
    except Exception as ex:
        utilities_general.printlogexception(ex=ex)
        global_settings.pathdb = None
    finally:
        None
        
def check_clicklistbox(event=None, lb1=None):
    if (event.y < lb1.bbox(0)[1]) or (event.y > lb1.bbox("end")[1]+lb1.bbox("end")[3]): # if not between it        
        lb1.select_clear(0, 'end')
        
        
def gather_information_fromdb(sqliteconn):    
    #doc = None  
    #sqliteconn = utilities_general.connectDB(str(global_settings.pathdb))
    cursor = sqliteconn.cursor()
    must_validate = necessity_to_validate(cursor)
    tocommit = False
    if(must_validate):
        tocommit = validate_new_db_columns(cursor, must_validate)
        update_db_version(sqliteconn, cursor)
    if(tocommit):
       sqliteconn.commit() 
    totalpaginas = 0
    global_settings.splash_window.window.attributes("-alpha", 255)
    try:
        None
        global_settings.splash_window.window.wm_attributes("-alpha", 255)
    except:
        None   
    select_all_pdfs = '''SELECT  P.id_pdf, P.rel_path_pdf, P.lastpos, P.tipo, P.margemsup, P.margeminf,
    P.margemesq, P.margemdir, P.hash, P.indexado, P.pixorgw, P.pixorgh, P.doclen, P.parent_alias, P.zoom_pos FROM 
    Anexo_Eletronico_Pdfs P ORDER BY 4,2
    '''
    porcento = 0
    global_settings.splash_window.label['text'] = f"Reunindo informações ({porcento}%)"
    cursor.custom_execute(select_all_pdfs)
    relats = cursor.fetchall()
    qtos = 0
    verificados = {}            
    cont = 0
    abs_path_pdf = None
    for r in relats: 
        abs_path_pdf = utilities_general.get_normalized_path(os.path.join(global_settings.pathdb.parent, str(r[1])))
        qtos+=1
        porcento = round(qtos/len(relats)*100, 0)
        global_settings.splash_window.label['text'] = f"Reunindo informações ({porcento}%)"
        global_settings.splash_window.label.update()        
        global_settings.infoLaudo[abs_path_pdf] = classes_general.Relatorio()
        filename, file_extension = os.path.splitext(abs_path_pdf)
        if(file_extension.lower()==".pdf"):  
            idpdf= r[0]
            doclen = r[12]
            pixmapw = r[10]
            pixmaph = r[11]
            parent_alias = r[13]
            if(r[12]==None):
                
                doc = fitz.open(abs_path_pdf)
                try:
                    doclen = len(doc)
                    pixorg = doc[0].get_pixmap()
                    pixmapw = int(pixorg.width)
                    pixmaph = int(pixorg.height)
                    updateinto2 = "UPDATE Anexo_Eletronico_Pdfs set pixorgw = ?, pixorgh= ?, doclen = ? WHERE id_pdf = ?"
                    cursor.custom_execute(updateinto2, (int(pixorg.width), int(pixorg.height), doclen, r[0],))
                    sqliteconn.commit()
                except Exception as ex:
                    utilities_general.printlogexception(ex=ex)
                finally:
                    doc.close()
            global_settings.infoLaudo[abs_path_pdf].zoom_pos = r[14]
            global_settings.infoLaudo[abs_path_pdf].mt = r[4]
            global_settings.infoLaudo[abs_path_pdf].mb = r[5]
            global_settings.infoLaudo[abs_path_pdf].me = r[6]
            global_settings.infoLaudo[abs_path_pdf].md = r[7]
            global_settings.infoLaudo[abs_path_pdf].rel_path_pdf = r[1]
            global_settings.infoLaudo[abs_path_pdf].hash = r[8]
            global_settings.infoLaudo[abs_path_pdf].id = idpdf
            global_settings.infoLaudo[abs_path_pdf].len = doclen
           
            global_settings.infoLaudo[abs_path_pdf].parent_alias = parent_alias
            if(r[8]==1):
                global_settings.infoLaudo[abs_path_pdf].pagiansprocessadas = global_settings.infoLaudo[abs_path_pdf].len
            else:
                global_settings.infoLaudo[abs_path_pdf].paginasprocessadas = 0
            totalpaginas += global_settings.infoLaudo[abs_path_pdf].len
            global_settings.infoLaudo[abs_path_pdf].tipo = r[3]
            global_settings.infoLaudo[abs_path_pdf].pixorgw = pixmapw
            global_settings.infoLaudo[abs_path_pdf].pixorgh = pixmaph
            select_tocs = '''SELECT  T.toc_unit, T.pagina, T.deslocy, T.init FROM 
            Anexo_Eletronico_Tocs T WHERE T.id_pdf = ? ORDER BY 2,3
            '''              
            cursor.custom_execute(select_tocs, (r[0],))
            tocs = cursor.fetchall()
            for toc in tocs:
                global_settings.infoLaudo[abs_path_pdf].toc.append((toc[0], int(toc[1]), int(toc[2]), int(toc[3])))
            
            #    global_settings.listaRELS[abs_path_pdf] = (r[0], r[1], abs_path_pdf, (toc[0], int(toc[1]), int(toc[2]), int(toc[3])), 0) 
            global_settings.infoLaudo[abs_path_pdf].ultimaPosicao=float(r[2])
            global_settings.infoLaudo[abs_path_pdf].tipo = r[3]
            global_settings.infoLaudo[abs_path_pdf].id = r[0] 
            paginasindexadas = 0
            if(not os.path.exists(abs_path_pdf)):
                global_settings.infoLaudo[abs_path_pdf].status = 'erro'
            else:
                if(r[8]=='' or r[8]==None):
                    global_settings.infoLaudo[abs_path_pdf].status = 'naoindexado'
                    global_settings.documents_to_index.append(abs_path_pdf)
                else:
                    
                    hashpdf = str(utilities_general.md5(abs_path_pdf))
                    if(hashpdf.lower()!=r[8].lower()):
                        print(abs_path_pdf)
                        print(hashpdf.lower())
                        global_settings.infoLaudo[abs_path_pdf].status = 'incompativel'
                    else:
                        global_settings.infoLaudo[abs_path_pdf].status = 'indexado'
                        paginasindexadas = r[12]
           
            relatorio_proxy = classes_general.RelatorioSuccint(r[0], global_settings.infoLaudo[abs_path_pdf].toc, global_settings.infoLaudo[abs_path_pdf].len, \
                                                               pixmapw, pixmaph, r[3], r[4], r[5], r[6], paginasindexadas, \
                                                                   r[1], abs_path_pdf, r[2])   
            global_settings.listaRELS[abs_path_pdf] = relatorio_proxy
            verificados[str(idpdf)] = "OK"              
            cont+=1  
    validate_annotation(sqliteconn, cursor)
    
def createNewDbFile(toplevel=None, sqliteconnx=None):
    if(global_settings.pathdb !=None):
        
        sqliteconn = None
        if(sqliteconnx==None):
            if(os.path.exists(str(global_settings.pathdb))):
                os.remove(str(global_settings.pathdb))
            sqliteconn = utilities_general.connectDB(str(global_settings.pathdb))
        else:
            sqliteconn = sqliteconnx
        cursor = sqliteconn.cursor()
        try:
            cursor.custom_execute("PRAGMA foreign_keys = ON")
            cursor.custom_execute(''' DROP TABLE IF EXISTS Anexo_Eletronico_Equips''')
            cursor.custom_execute(''' DROP TABLE IF EXISTS Anexo_Eletronico_Pdfs''')
            cursor.custom_execute(''' DROP TABLE IF EXISTS Anexo_Eletronico_Tocs''')
            cursor.custom_execute(''' DROP TABLE IF EXISTS Anexo_Eletronico_Conteudo''')
    
            create_table_pdfs = '''CREATE TABLE Anexo_Eletronico_Pdfs (
            id_pdf INTEGER PRIMARY KEY AUTOINCREMENT,
            rel_path_pdf TEXT NOT NULL,
            indexado INTEGER NOT NULL,
            hash TEXT,
            tipo TEXT NOT NULL,
            lastpos TEXT NOT NULL,
            margemsup INTEGER NOT NULL,
            margeminf INTEGER NOT NULL,
            margemesq INTEGER NOT NULL,
            margemdir INTEGER NOT NULL,
            pixorgw INTEGER,
            pixorgh INTEGER,
            doclen INTEGER,
            zoom_pos INTEGER DEFAULT 0,
            parent_alias TEXT DEFAULT '')
            '''
    
            create_table_tocs = '''CREATE TABLE Anexo_Eletronico_Tocs (
            id_toc INTEGER PRIMARY KEY AUTOINCREMENT,
            toc_unit TEXT NOT NULL,
            id_pdf TEXT NOT NULL,
            pagina INTEGER NOT NULL,
            deslocy INTEGER NOT NULL,
            init INTEGER NOT NULL,
            external_pdf TEXT,
            CONSTRAINT fk_pdf
                FOREIGN KEY (id_pdf)
                    REFERENCES Anexo_Eletronico_Pdfs (id_pdf)
                    ON DELETE CASCADE
            )'''
    
            create_table_pdflinks = '''CREATE TABLE Anexo_Eletronico_Links(
            id_link INTEGER PRIMARY KEY AUTOINCREMENT,
            paginainit INTEGER NOT NULL,
            p0x INTEGER NOT NULL,
            p0y INTEGER NOT NULL,
            paginafim INTEGER NOT NULL,
            p1x INTEGER NOT NULL,
            p1y INTEGER NOT NULL,
            tipo TEXT NOT NULL,
            id_obs INTEGER NOT NULL,
            id_pdf INTEGER NOT NULL,
            fixo INTEGER NOT NULL,
            CONSTRAINT fk_obs
                FOREIGN KEY (id_obs)
                    REFERENCES Anexo_Eletronico_Obsitens (id_obs)
                    ON DELETE CASCADE
            CONSTRAINT fk_pdf
                FOREIGN KEY (id_pdf)
                    REFERENCES Anexo_Eletronico_Pdfs (id_pdf)
                    ON DELETE CASCADE
            )
            '''
    
            create_table_searchterms = '''CREATE TABLE Anexo_Eletronico_SearchTerms (
            id_termo INTEGER PRIMARY KEY AUTOINCREMENT,
            termo TEXT NOT NULL,
            tipobusca TEXT NOT NULL, 
            fixo INTEGER NOT NULL,
            pesquisado TEXT)  
            '''
    
            create_table_searchpdfs = '''CREATE TABLE Anexo_Eletronico_SearchPdfs (
            id_termo_pdf INTEGER PRIMARY KEY AUTOINCREMENT,
            id_pdf INTEGER NOT NULL,
            id_termo INTEGER NOT NULL,
            CONSTRAINT fk_pdf
                FOREIGN KEY (id_pdf)
                    REFERENCES Anexo_Eletronico_Pdfs (id_pdf)
                    ON DELETE CASCADE,
            CONSTRAINT fk_id_termo
                FOREIGN KEY (id_termo)
                    REFERENCES Anexo_Eletronico_SearchTerms (id_termo)
                    ON DELETE CASCADE
            )  
            '''
            
            create_table_searchesresults = '''CREATE TABLE Anexo_Eletronico_SearchResults (
            id_termo INTEGER NOT NULL,
            id_pdf INTEGER NOT NULL,
            pagina INTEGER NOT NULL,
            init INTEGER NOT NULL,
            fim INTEGER NOT NULL,
            toc TEXT,
            snippetantes TEXT,
            snippetdepois TEXT,
            termo TEXT,
            CONSTRAINT fk_termo
                FOREIGN KEY (id_termo)
                    REFERENCES Anexo_Eletronico_SearchTerms (id_termo)
                    ON DELETE CASCADE,
            CONSTRAINT fk_pdf
            FOREIGN KEY (id_pdf)
                REFERENCES Anexo_Eletronico_Pdfs (id_pdf)
                ON DELETE CASCADE
            )
            '''
            
            create_table_obscat = '''CREATE TABLE Anexo_Eletronico_Obscat (
            id_obscat INTEGER PRIMARY KEY AUTOINCREMENT,
            obscat TEXT NOT NULL,
            fixo INTEGER NOT NULL,
            ordem INTEGER NOT NULL
            )
            '''
            
            create_table_obsitens = '''CREATE TABLE Anexo_Eletronico_Obsitens (
            id_obs INTEGER PRIMARY KEY AUTOINCREMENT,
            id_obscat INTEGER NOT NULL,
            id_pdf INTEGER NOT NULL,
            paginainit INTEGER NOT NULL,
            p0x INTEGER NOT NULL,
            p0y INTEGER NOT NULL,
            paginafim INTEGER NOT NULL,
            p1x INTEGER NOT NULL,
            p1y INTEGER NOT NULL,
            tipo TEXT NOT NULL,
            fixo INTEGER NOT NULL,
            status TEXT NOT NULl DEFAULT 'ok',
            conteudo TEXT DEFAULT '',
            arquivo TEXT DEFAULT '',
            withalt INTEGER DEFAULT 0,
            CONSTRAINT fk_obs
                FOREIGN KEY (id_obscat)
                    REFERENCES Anexo_Eletronico_Obscat (id_obscat)
                    ON DELETE CASCADE,
            CONSTRAINT fk_pdf                    
                FOREIGN KEY (id_pdf)
                    REFERENCES Anexo_Eletronico_Pdfs (id_pdf)
                    ON DELETE CASCADE
            )
            '''
            
            create_table_annotations = '''CREATE TABLE Anexo_Eletronico_Annotations (
            id_annot INTEGER PRIMARY KEY AUTOINCREMENT,
            id_obs INTEGER NOT NULL,
            id_pdf INTEGER NOT NULL,
            rel_path_pdf NOT NULL,
            paginainit INTEGER NOT NULL,
            p0x INTEGER NOT NULL,
            p0y INTEGER NOT NULL,
            paginafim INTEGER NOT NULL,
            p1x INTEGER NOT NULL,
            p1y INTEGER NOT NULL,
            link TEXT DEFAULT '',
            conteudo TEXT DEFAULT '',
            CONSTRAINT fk_obs
                FOREIGN KEY (id_obs)
                    REFERENCES Anexo_Eletronico_Obsitens (id_obs)
                    ON DELETE CASCADE,
            CONSTRAINT fk_pdf                    
                FOREIGN KEY (id_pdf)
                    REFERENCES Anexo_Eletronico_Pdfs (id_pdf)
                    ON DELETE CASCADE
            )
            '''
            
            
            
            
            create_table_config = '''CREATE TABLE FERA_CONFIG (
            id_conf INTEGER PRIMARY KEY AUTOINCREMENT,
            config TEXT NOT NULL,
            param TEXT NOT NULL
            )'''
            cursor.custom_execute(create_table_pdfs)
            cursor.custom_execute(create_table_tocs)
            cursor.custom_execute(create_table_searchterms)
            cursor.custom_execute(create_table_searchesresults)
            cursor.custom_execute(create_table_obscat)
            cursor.custom_execute(create_table_obsitens)
            cursor.custom_execute(create_table_pdflinks)
            cursor.custom_execute(create_table_config)
            #cursor.custom_execute(create_table_images)
            sqliteconn.commit()
            insert_query_pdf = """INSERT INTO FERA_CONFIG
            (config , param) VALUES (?,?)   """
            cursor.custom_execute(insert_query_pdf, ('dbversion', str(global_settings.dbversion),))
            sqliteconn.commit()
            if(toplevel!=None):
                loaddb(toplevel)
            else:
                return True
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
            global_settings.pathdb = None
            if(toplevel==None):
                traceback.print_exc()
                return False
        finally:
            if(sqliteconnx==None):            
                sqliteconn.close()

        
class import_create_toplevel():
    def __init__(self):
        def solicitarDiretorio(toplevel):
            tipos = [('SQLite DB', '*.db')]
            path = (asksaveasfilename(filetypes=tipos, defaultextension=tipos))
            try:
                with open(path, 'w'):
                    pass
                global_settings.pathdb = Path(path) 
                createNewDbFile(toplevel)                                  
            except:
                global_settings.pathdb = None
            
        toplevel = tkinter.Toplevel(global_settings.root)  
        toplevel.bind('<Control-Shift-L>', lambda e: global_settings.log_window.deiconify())
        toplevel.protocol("WM_DELETE_WINDOW", lambda : on_quit(toplevel))
        toplevel.geometry("600x600")
        toplevel.title("FERA "+global_settings.version+" - Forensics Evidence Report Analyzer -- Polícia Científica do Paraná")
        toplevel.columnconfigure(0, weight=1)
        toplevel.rowconfigure(0, weight=1)
        
        frameinitial = tkinter.Frame(toplevel, highlightbackground="green", highlightcolor="green", highlightthickness=1, bd=0)
        frameinitial.grid(row=0, column=0, sticky='nsew')
        frameinitial.columnconfigure((0,1), weight=1)
        frameinitial.rowconfigure((0,1), weight=1)
        frameinitial.rowconfigure(2, weight=10)
        labelopcao = tkinter.Label(frameinitial, text="Clique na opção desejada")
        labelopcao.grid(row=0, column=0, columnspan=2, sticky='nsew')
        
        newdb = tkinter.Button(frameinitial, font=global_settings.Font_tuple_Arial_8, text="Criar caso", command= lambda : solicitarDiretorio(toplevel))
        newdb.grid(row=1, column=0, sticky='n')
        recenttxt = os.path.join(os.getcwd(),"recents.txt")
        Lb1 = tkinter.Listbox(frameinitial)
        olddb = tkinter.Button(frameinitial, font=global_settings.Font_tuple_Arial_8, text="Importar caso", command= lambda lb=Lb1: loaddb(toplevel, None, lb))
        olddb.grid(row=1, column=1, sticky='n')
        
        Lb1.bindtags(('.Lb1', 'Listbox', 'post-listbox-bind-search','.','all'))
        Lb1.grid(row=2, column=0, columnspan=2, sticky='nsew')
        Lbhs = tkinter.Scrollbar(frameinitial, orient="horizontal")
        Lbhs.config( command = Lb1.xview )
        Lbhs.grid(row=3, column=0, columnspan=2, sticky='ew')
        Lb1.configure(xscrollcommand=Lbhs.set)    
        
        Lbvs = tkinter.Scrollbar(frameinitial, orient="vertical")
        Lbvs.config( command = Lb1.yview )
        Lbvs.grid(row=2, column=2, rowspan=2, sticky='ns')
        Lb1.configure(yscrollcommand=Lbvs.set)
        try:        
            with open(recenttxt, 'r', encoding="utf-8") as f:
                count = 1
                for line in f:
                    if(os.path.isfile(line.strip())):
                        Lb1.insert(count, line)
                        count += 1
        except Exception as ex:
            None
            #utilities_general.utilities_general.printlogexception(ex=ex)
        Lb1.bind_class('post-listbox-bind-search', '<Double-1>', lambda e, lb=Lb1: loaddb(toplevel, e, lb))
        Lb1.bind_class('post-listbox-bind-search','<1>',  lambda e,  lb=Lb1: check_clicklistbox(e, lb))
        global_settings.root.wait_window(toplevel)
    