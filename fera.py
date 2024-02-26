# -*- coding: utf-8 -*-
"""
Created on Sun Sep 27 14:39:11 2020

@author: gustavo.bedendo
"""

# -*- coding: utf-8 -*-
"""
Created on Wed Sep 16 13:45:26 2020

@author: gustavo.bedendo
"""

from threading import Thread
import tkinter
import global_settings
import time, re
from tkinter import ttk
import fitz
import math
import multiprocessing as mp
from PIL import Image, ImageTk, ImageFile
ImageFile.LOAD_TRUNCATED_IMAGES = True
try:
    import keyboard
except:
    pass
from tkinter.filedialog import askopenfilename, asksaveasfilename, askdirectory
import subprocess, os, platform
from pathlib import Path
import webbrowser
import clipboard
from io import BytesIO
import traceback
import sqlite3
import sys
#import indexador_fera
from functools import partial
import shutil
#import getopt

import xlsxwriter
import binascii, io
import utilities_general, utilities_export
import classes_general, process_functions
import rtfunicode
plt = platform.system()
try:
    import pygetwindow as gw
except:
    None
""" try:
    import gi
except:
    None
try:
    from gi.repository import Gtk
except:
    None
try:
    import pywinctl as gw
except:
    None """
try:
    import win32clipboard
except:
    None


try:
    from pywinauto import Application
except:
    None

class MainWindow():    
    def fixed_map(self, option):
        return [elm for elm in self.style.map('Treeview', query_opt=option) if
          elm[:2] != ('!disabled', '!selected')]
    
    def __init__(self):
        def _executable_exists(name):
            return subprocess.call(["which", name],
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE) == 0
        self.locations_toplevels = {}
        self.previousSearchesWindow = None
        self.previous_simple_searches = []
        self.lista_executaveis = ['Escutar e transcrever']        
        self.paginaSearchSimple = -1
        self.termossimplespesquisados = {}
        self.primeiroresetbuscar = True
        self.somasnippet = 0
        self.alreadyenhanced = set()
        self.othertags = []
        self.allimages = {}
        self.totalMov = 16
        self.globalFrame = ttk.PanedWindow(orient=tkinter.HORIZONTAL)
        self.positions = [None] * 10
        self.indiceposition = 0
        self.globalFrame.grid(row=0, column=0, sticky="nsew")
        self.globalFrame.rowconfigure(1, weight=1)
        self.globalFrame.columnconfigure((0, 1, 2, 3, 4, 5), weight=1)
        self.ininCanvasesid = [None] * global_settings.minMaxLabels * global_settings.divididoEm
        self.tkimgs = [None] * global_settings.minMaxLabels * global_settings.divididoEm
        self.fakeImage = None
        self.fakePages = [None] * global_settings.minMaxLabels
        self.fakeLines = [None] * global_settings.minMaxLabels
        self.initialPos = None
        self.bg = "#%02x%02x%02x" % (145, 145, 145)
        self.maiorresult = 0       
        self._jobscrollpagebymouse = None
        self.style = ttk.Style()
        self.style.configure("Treeview.Heading", font=global_settings.Font_tuple_Arial_10, foreground='#525252')
        self.style.configure("Treeview", font=global_settings.Font_tuple_Arial_10, rowheight=22, indent=10)
        self.style.configure("boldify-results", font=global_settings.Font_tuple_Arial_10)
        self.style.configure("unboldify-results", font=global_settings.Font_tuple_Arial_10)
        self.style.configure("TNotebook.Tab", font=global_settings.Font_tuple_ArialBold_12)
        self.style.configure("TNotebook.Tab", borderwidth=1)
        self.style.configure("TNotebook", tabposition='n')
        self.style.map('Treeview', foreground=self.fixed_map('foreground'), background=self.fixed_map('background'))
        self.simplesearching = False       
        self.afterpaint = None
        self.afterquads = None
        global_settings.root.attributes("-alpha", 0)
        try:
            global_settings.root.wm_attributes("-alpha", 0)
        except:
            None 
        global_settings.root.deiconify()
        w, h = global_settings.root.winfo_screenwidth(),global_settings. root.winfo_screenheight()
        if plt == "Linux":
            try: 
                global_settings.root.attributes('-zoomed', True)
                
            except Exception as ex:
                global_settings.root.geometry("%dx%d+0+0" % (w, h-20))
                utilities_general.printlogexception(ex=ex)                
        elif plt=="Windows":
            try:
                global_settings.root.state("zoomed")
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)
                global_settings.root.geometry("%dx%d+0+0" % (w, h-20))
        self.color = (103, 245, 134, 75)
        self.colorenhancebookmark = (103, 245, 134, 35)
        self.colorenhanceannotation = (144, 59, 255, 25)
        self.colorquad = (21, 71, 150, 85)
        self.colorlink = (175, 200, 240, 95)
        self.initUI()
        
            
        #global_settings.root.update_idletasks() 
        self.window_search_info = classes_general.Search_Info_Window(global_settings.root, self.docInnerCanvas.winfo_rooty())
        global_settings.splash_window.window.destroy()
        global_settings.root.attributes("-alpha", 1.0)
        try:
            None
            global_settings.root.wm_attributes("-alpha", 1.0)
        except:
            None 
        if(global_settings.indexing_thread!=None and global_settings.indexing_thread.is_alive()):
            x = global_settings.root.winfo_x()
            y = global_settings.root.winfo_y()
            global_settings.indexingwindow.window.deiconify()
            global_settings.indexingwindow.window.geometry("+%d+%d" % (x + 5, y + 5))
            global_settings.indexingwindow.window.overrideredirect(True)
            global_settings.indexingwindow.window.lift()
            
        self.exporting = False
        self.exportinterval = utilities_export.ExportInterval(global_settings.root)
        self.exportinterval.window.withdraw()
        
        self.winfox = self.docInnerCanvas.winfo_x()
        self.winfoy= self.docInnerCanvas.winfo_y()
        self.docInnerCanvas.bind("<Configure>", self.configureWindow)
        self.labeldocname.config(font=global_settings.Font_tuple_Arial_10, text=os.path.basename(global_settings.pathpdfatual))
        global_settings.root.focus_set()
        if(plt == 'Linux' and not _executable_exists("xclip")):
            utilities_general.popup_window("A biblioteca XCLIP parece não estar instalada!\nAlgumas funcionalidades podem apresentar problemas.\nFavor instalar o pacote XCLIP (sudo apt install xclip)", False)
        #self.globalFrame.sash_place(0, 450,self.winfoy)

    
        #partial   

    def initUI(self):        
        start_time = time.time()
        try:
            
            self.searchedTerms = []
            self.leftPanel()
            self.createTopBar()
            self.drawCanvas()
            global_settings.root.update() 
            global_settings.root.resizable(True, True)
            self.selectReport(self.primeiro)
            self.checkUpdates()
            self.treeSeachAfter()
            self.checkPages()
            self.populationSearches = None
            self.getlastPos()
            #self.showAllBookmarks()  
            self.cleanup()
            self.free_searches()
            self.check_errors()
            global_settings.label_warning_error.bind('<Button-1>',  self.open_log_window)
            
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
 
    def open_log_window(self, event=None):
        global_settings.log_window.deiconify()
        global_settings.root.after(10, lambda: global_settings.log_window.lift())
        #global_settings.log_window.lift()

    def loadDocOnCanvas(self):
        try:
            global_settings.processed_pages[0] = 0
            global_settings.pathpdfatual2 = global_settings.pathpdfatual
            pedido1 = classes_general.classes_general.PedidoDePagina(qualLabel = 0, qualPdf = global_settings.pathpdfatual2, qualPagina = 0, matriz = self.mat, \
                  pixheight = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh, pixwidth = \
                      global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw, zoom = self.zoom_x*global_settings.zoom, \
                      scrollvalue = self.vscrollbar.get()[0] ,\
                      scrolltotal = self.scrolly, canvash = self.canvash, mt = global_settings.infoLaudo[global_settings.pathpdfatual].mt, \
                          mb = global_settings.infoLaudo[global_settings.pathpdfatual].mb, me = global_settings.infoLaudo[global_settings.pathpdfatual].me,\
                              md = global_settings.infoLaudo[global_settings.pathpdfatual].md)
            global_settings.request_queuexml.put(pedido1)
            global_settings.request_queue.put(pedido1)
            global_settings.processed_requests[0] = pedido1
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
            
    def checkLink(self, event):        
        try:
            try:
                if self.tw:
                    self.tw.destroy()
            except Exception as ex:
                None
            self.tw = tkinter.Toplevel(self.docInnerCanvas, background='#ededd3')  
            self.tw.rowconfigure(2, weight=1)
            self.tw.columnconfigure(0, weight=1)
            # Leaves only the label and removes the app window
            self.tw.wm_overrideredirect(True)
            x = event.x_root + 15
            y = event.y_root + 10
            self.tw.wm_geometry("+%d+%d" % (x, y))
            self.tw.withdraw()   
            labels = []
            posicaoRealY0Canvas = self.vscrollbar.get()[0] * self.scrolly + event.y
            posicaoRealX0Canvas = self.hscrollbar.get()[0] * (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw * self.zoom_x * global_settings.zoom) + event.x
            posicaoRealY0 = (posicaoRealY0Canvas % (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom)) / (self.zoom_x*global_settings.zoom)
            posicaoRealX0 = posicaoRealX0Canvas / (self.zoom_x*global_settings.zoom)
            pagina = math.floor(posicaoRealY0Canvas / (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh *self.zoom_x*global_settings.zoom))
            self.labelmousepos.config(font=global_settings.Font_tuple_Arial_10, text="(Pagina:{},X:{},Y:{})".format(pagina+1, round(posicaoRealX0), round(posicaoRealY0)))
            cont_row = 0
            if(not self.selectionActive and not self.areaselectionActive):
                if(pagina in global_settings.infoLaudo[global_settings.pathpdfatual].links):
                    for link in global_settings.infoLaudo[global_settings.pathpdfatual].links[pagina]:
                        r = link['from']
                        
                        if(posicaoRealY0 >= r.y0 and posicaoRealY0 <= r.y1 and posicaoRealX0 >= r.x0 and posicaoRealX0 <= r.x1):
                            #print(link)
                            self.docInnerCanvas.config(cursor='hand2')
                            ehLink=True
                            try:
                                text_link = ""
                                if('file' in link):
                                    text_link = link['file']
                                elif('to' in link):
                                    text_link = link['to']
                                elif('uri' in link):
                                    text_link = link['uri']

                                if(text_link==""):
                                    xref = link['xref']
                                    info = global_settings.docatual.xref_get_key(xref, 'A')
                                    grupos_search = global_settings.regex_actions_compiled.search(info[1])
                                    #print(info, grupos_search)
                                    if(grupos_search==None):
                                        continue
                                    grupos = grupos_search.groups()
                                    if(grupos[-1]=="Launch"):
                                        text_link += grupos[2]
                                    elif(grupos[-1]=="GoToR"):
                                        text_link += f"{grupos[2]}#{grupos[1]}"
                                texto_file = "Link: " + text_link
                                label1 = tkinter.Label(self.tw, font=global_settings.Font_tuple_Arial_10, text=texto_file, justify='left',
                                               background='#ededd3', relief='solid', borderwidth=0)
                                label1.grid(row=0, column=0, sticky='w', pady=0)
                                cont_row += 1
                            except Exception as ex:
                                utilities_general.printlogexception(ex=ex)
                            break

                if(cont_row==0):
                        self.docInnerCanvas.config(cursor='')
            if(self.showbookmarsboolean):
                listaquads = self.docInnerCanvas.find_withtag('enhanceobs'+global_settings.pathpdfatual+str(pagina))
                listaquads_annot = self.docInnerCanvas.find_withtag('enhanceannot'+global_settings.pathpdfatual+str(pagina))
                for quadelement in listaquads:  
                    tags = self.docInnerCanvas.gettags(quadelement)
                    idobsitem = "obsitem"+(tags[1].replace("enhanceobs",""))
                    observation = self.allobsbyitem[idobsitem]
                    bbox = self.docInnerCanvas.bbox(quadelement)
                    if(self.docInnerCanvas.canvasx(event.x) >= bbox[0] and self.docInnerCanvas.canvasy(event.y) >= bbox[1] \
                        and self.docInnerCanvas.canvasx(event.x) <= bbox[2] and self.docInnerCanvas.canvasy(event.y) <= bbox[3]):
                        
                        textoobscat = (self.treeviewObs.item(observation.idobscat, 'values')[2] + "\n").strip()
                        
                        label = tkinter.Label(self.tw, font=global_settings.Font_tuple_ArialBold_10, text="Categoria:" + textoobscat, justify='left',
                                       background='#ededd3', relief='solid', borderwidth=0)
                        label.grid(row=cont_row, column=0, sticky='w', pady=0) 
                        cont_row += 1
                        conteudo = observation.conteudo
                        colocarbreak = False
                        cont = 0
                        new_input  = ''
                        for i, letter in enumerate(conteudo):
                            if i % 50 == 0 and i> 0:
                                colocarbreak = True
                            if(colocarbreak):
                                new_input += letter
                                if(letter== ' '):                                    
                                    new_input += '\n'
                                    colocarbreak = False
                                else:
                                    cont += 1
                                    if(cont>10):
                                       new_input += '\n'
                                       cont = 0
                                       colocarbreak = False
                            else:
                                new_input += letter
                        if(new_input!=''):
                            label = tkinter.Label(self.tw, font=global_settings.Font_tuple_ArialBold_10, text="Anotação Principal:", justify='left',
                                           background='#ededd3', relief='solid', borderwidth=0)
                            label.grid(row=cont_row, column=0, sticky='w', pady=0) 
                            cont_row += 1
                            label = tkinter.Label(self.tw, font=global_settings.Font_tuple_Arial_10_italic, text=new_input, justify='left',
                                           background='#ededd3', relief='solid', borderwidth=0)
                            label.grid(row=cont_row, column=0, sticky='w', pady=0) 
                            cont_row += 1
                for quadelement in listaquads_annot:  
                    tags = self.docInnerCanvas.gettags(quadelement)
                    tags_split = tags[1].split("enhanceannot")
                    idobsitem = "obsitem"+tags_split[0]
                    idannot = int(tags_split[1])
                    annotation = self.allobsbyitem[idobsitem].annotations[idannot]
                    conteudo = annotation.conteudo
                    
                    if(conteudo!=''):
                        link = "Link:{}".format(annotation.link)
                        bbox = self.docInnerCanvas.bbox(quadelement)
                        if(self.docInnerCanvas.canvasx(event.x) >= bbox[0] and self.docInnerCanvas.canvasy(event.y) >= bbox[1] \
                            and self.docInnerCanvas.canvasx(event.x) <= bbox[2] and self.docInnerCanvas.canvasy(event.y) <= bbox[3]):
                            colocarbreak = False
                            cont = 0
                            new_input  = ''
                            for i, letter in enumerate(conteudo):
                                if i % 50 == 0 and i> 0:
                                    colocarbreak = True
                                if(colocarbreak):
                                    new_input += letter
                                    if(letter== ' '):                                    
                                        new_input += '\n'
                                        colocarbreak = False
                                    else:
                                        cont += 1
                                        if(cont>10):
                                           new_input += '\n'
                                           cont = 0
                                           colocarbreak = False
                                else:
                                    new_input += letter
                            if(new_input!=''):
                                label = tkinter.Label(self.tw, font=global_settings.Font_tuple_ArialBold_10, text="Anotação do link:", justify='left',
                                               background='#ededd3', relief='solid', borderwidth=0)
                                label.grid(row=cont_row, column=0, sticky='w', pady=0) 
                                cont_row += 1
                                label = tkinter.Label(self.tw, font=global_settings.Font_tuple_Arial_10_italic, text=new_input, justify='left',
                                               background='#ededd3', relief='solid', borderwidth=0)
                                label.grid(row=cont_row, column=0, sticky='w', pady=0) 
                                cont_row += 1
                        
            if(pagina in global_settings.infoLaudo[global_settings.pathpdfatual].widgets):
               for wid in global_settings.infoLaudo[global_settings.pathpdfatual].widgets[pagina]:
                   if(posicaoRealX0 >= wid[1][0] and posicaoRealX0 <= wid[1][2] and posicaoRealY0 >= wid[1][1] and posicaoRealY0 <= wid[1][3]):                       
                       
                       label = tkinter.Label(self.tw, font=global_settings.Font_tuple_Arial_10_italic, text="\n"+wid[0], justify='left',
                                      background='#ededd3', relief='solid', borderwidth=0)
                       label.grid(row=cont_row, column=0, sticky='w', pady=0) 
                       cont_row += 1
           
            if(cont_row>=1):
                self.tw.deiconify()
        except Exception as ex:
            None
            utilities_general.printlogexception(ex=ex)
            
    def configureWindow(self, event=None):        
        sobraEspaco = 0
        if(self.docFrame.winfo_width() > global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw * self.zoom_x * global_settings.zoom):
            sobraEspaco = self.docInnerCanvas.winfo_x()
        self.maiorw = self.docFrame.winfo_width()
        if(global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw*self.zoom_x*global_settings.zoom>self.maiorw):
            self.maiorw = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw*self.zoom_x *global_settings.zoom           
        self.scrolly = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh*self.zoom_x*global_settings.zoom*global_settings.infoLaudo[global_settings.pathpdfatual].len  - 35
        self.docInnerCanvas.config(scrollregion=(sobraEspaco, 0, sobraEspaco+global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw * self.zoom_x * global_settings.zoom, self.scrolly))
        dx = 0
        #print(self.docFrame.winfo_width())
        if(self.docFrame.winfo_width()<1070):
            #self.toolbar = tkinter.Frame(self.docOuterFrame, borderwidth=4, bg=self.bg, relief='groove')     
            self.toolbar.columnconfigure((0, 1, 2, 3), weight=1)
            self.toolbar.rowconfigure((0,1), weight=1)
            self.toolbar.grid(column=0, row=0, sticky='ew')   
            
            #self.manipulationTool = tkinter.Frame(self.toolbar, bg=self.bg, borderwidth=4, relief='groove')
            self.manipulationTool.grid(row=1, column=0, columnspan=2, sticky='w')
            #self.manipulationTool.columnconfigure((0,1,2,3,4,5,6), weight=1)
            
            #self.basicTool = tkinter.Frame(self.toolbar, bg=self.bg)  
            self.basicTool.grid(row=0, column=0, columnspan=4, sticky='n')
            #self.basicTool.columnconfigure((0,1,2,3,4,5,6,7), weight=1)
            
            #self.rightFrame  = tkinter.Frame(self.toolbar, bg=self.bg)
            self.rightFrame.grid(row=1, column=2, columnspan=2, sticky='e')
            #self.rightFrame.columnconfigure((0,1,2,3), weight=1)  
        else:
            #self.toolbar.columnconfigure((0, 1, 2, 3), weight=1)
            self.toolbar.rowconfigure(0, weight=1)
            self.toolbar.columnconfigure((0,1,2), weight=1)
            self.toolbar.grid(column=0, row=0, sticky='ew')  
            #self.manipulationTool = tkinter.Frame(self.toolbar, bg=self.bg, borderwidth=4, relief='groove')
            
            #global_settings.root.update_idle
            
            self.manipulationTool.grid(row=0, column=0, columnspan=1, sticky='w')
            #self.manipulationTool.columnconfigure((0,1,2,3,4,5,6,7), weight=1)
            
            #self.basicTool = tkinter.Frame(self.toolbar, bg=self.bg)  
            #self.basicTool.grid(row=0, column=1, sticky='w')
            self.basicTool.grid(row=0, column=1, columnspan=1, sticky='n')
            #self.basicTool.columnconfigure((0,1,2,3,4,5,6,7), weight=1)
            
            #self.rightFrame  = tkinter.Frame(self.toolbar, bg=self.bg)
            self.rightFrame.grid(row=0, column=3, columnspan=1,sticky='e')
            #self.rightFrame.columnconfigure((0,1,2,3), weight=1)
        try:
            for indice in range(global_settings.minMaxLabels):
                coords = self.docInnerCanvas.coords(self.ininCanvasesid[indice])
                pos_h = self.docInnerCanvas.winfo_x()
                dx = pos_h - coords[0] 
                self.docInnerCanvas.move(self.ininCanvasesid[indice], dx, 0)        
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        
        for quad in self.docInnerCanvas.find_withtag('quad'):
            coords = self.docInnerCanvas.coords(self.ininCanvasesid[indice])
            self.docInnerCanvas.move(quad, dx, 0)        
        for link in self.docInnerCanvas.find_withtag('link'):            
            self.docInnerCanvas.move(link, dx, 0)
        for simplesearch in self.docInnerCanvas.find_withtag('simplesearch'):            
            self.docInnerCanvas.move(simplesearch, dx, 0)
        for search in self.docInnerCanvas.find_withtag('obsitem'):            
            self.docInnerCanvas.move(search, dx, 0)
        listaobj = self.docInnerCanvas.find_all()
        for tag in self.allimages:
            if 'enhanceobs' in tag:
                for search in self.docInnerCanvas.find_withtag(tag):            
                    self.docInnerCanvas.move(search, dx, 0)
            elif 'enhanceannot' in tag:
                for search in self.docInnerCanvas.find_withtag(tag):            
                    self.docInnerCanvas.move(search, dx, 0)
        atual = round((self.vscrollbar.get()[0]*global_settings.infoLaudo[global_settings.pathpdfatual].len))
        
    def getlastPos(self):
        global_settings.infoLaudo[global_settings.pathpdfatual].ultimaPosicao=(self.vscrollbar.get()[0])
        global_settings.root.after(4000, self.getlastPos)
        

    def process_links(self, links, pagina, pdfatualnorm):
        sobraEspaco = 0
        if(self.docFrame.winfo_width() > global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw * self.zoom_x * global_settings.zoom):
            sobraEspaco = self.docInnerCanvas.winfo_x()
        deslocy = (math.floor(pagina) * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom)
        for link in links:
            try:
                r = link['from'] 
                arquivo = ''
                if('file' in link):
                    file = link['file']     
                    if(file==""):
                        xref = link['xref']
                        info = global_settings.docatual.xref_get_key(xref, 'A')
                        grupos_search = global_settings.regex_actions_compiled.search(info[1])
                        if(grupos_search==None):
                            continue
                        grupos = grupos_search.groups()
                        if(grupos[-1]=="Launch"):
                            file = grupos[2]         
                    arquivo = str(Path(utilities_general.get_normalized_path(os.path.join(Path(utilities_general.get_normalized_path(pdfatualnorm)).parent,file))))  
                rect = classes_general.Rect()
                x0k = math.floor(r.x0*self.zoom_x*global_settings.zoom +sobraEspaco)
                x1k = math.ceil(r.x1*self.zoom_x*global_settings.zoom +sobraEspaco)
                y0k = math.ceil(((r.y0*self.zoom_x*global_settings.zoom)  +deslocy))
                y1k = math.ceil(((r.y1*self.zoom_x*global_settings.zoom)  +deslocy))
                rect.image = utilities_general.create_rectanglex(min(x0k, x1k), min(y0k, y1k), max(x0k, x1k), max(y0k,y1k), (100,100,100,0), link=False, withborder=False, transparent=False)  
                tag = ('extlink', 'extlink'+pdfatualnorm,'extlink'+pdfatualnorm+str(pagina))
                rect.idrect = self.docInnerCanvas.create_image(min(x0k, x1k), min(y0k, y1k), image=rect.image, anchor='nw', tags=tag)
                self.addImagetoList(tag[0], rect) 
                classes_general.FileTooltip(self.docInnerCanvas, rect.idrect, arquivo, global_settings.pathdb)
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)
    
    def check_errors(self):
        try:
            erro = None
            
            if(not global_settings.erros_queue.empty()):   
                 
                 erro = global_settings.erros_queue.get(0)
                 #print(erro)
                 utilities_general.printlogexception(ex=erro[1])
                 if(erro[0]=='2'):
                     print('xyz')
                     global_settings.log_window_text.insert('end',erro[1])
                     global_settings.log_window_text.insert('end',"\n")
                     
                 elif(erro[0]=='3'): 
                     #print('abc')
                     global_settings.log_window_text.insert('end',erro[1] + " " + erro[2] + "\n")
                     global_settings.log_window_text.insert('end',erro[4] + "\n")
                     global_settings.log_window_text.insert('end',"\n")
                     utilities_general.popup_window("{} - {} \n {}".format(erro[1], erro[2], erro[4]), False)
                     idtermo = str(erro[3])
                     idx = 't'+idtermo
                     #print(idx, (self.treeviewSearches.exists(idx)))
                     if(self.treeviewSearches.exists(idx)):
                         self.treeviewSearches.item(idx, tags=("termowitherror",))
                     
            else:
                time.sleep(0.1)    
            #if(erro[0]=='erro'):
            #    utilities_general.printlogexception(ex=erro[1])
            #else:
            #    None
        except:
            None
        finally:
            global_settings.root.after(3000, self.check_errors)
            
    def cleanup(self):
        try:
            if(len(global_settings.infoLaudo[global_settings.pathpdfatual].mapeamento)>= 100):
                remove = [previous for previous in global_settings.infoLaudo[global_settings.pathpdfatual].mapeamento if not previous in global_settings.processed_pages]
                for previous in remove: 
                    del global_settings.infoLaudo[global_settings.pathpdfatual].links[previous]
                    del global_settings.infoLaudo[global_settings.pathpdfatual].widgets[previous]
                    del global_settings.infoLaudo[global_settings.pathpdfatual].mapeamento[previous]
                    del global_settings.infoLaudo[global_settings.pathpdfatual].quadspagina[previous]
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        finally:
            global_settings.root.after(60000, self.cleanup)
    
    def previousSearchSelected(self, event=None):
        self.previousSearchesWindow.withdraw()
        selection = event.widget.curselection()
        if selection:
            index = selection[0]
            data = event.widget.get(index)
            self.simplesearch.delete("0", "end")
            self.simplesearch.insert("0", data)
            #label.configure(text=data)
        else:
            None
            #label.configure(text="")
        
    
    def showPreviousSearches(self, event=None):
        if(len(self.previous_simple_searches) > 0 and self.previousSearchesWindow==None):
            #x = y = 0
            x = self.simplesearch.winfo_rootx() + 0
            y = self.simplesearch.winfo_rooty() + self.simplesearch.winfo_height()
            # creates a toplevel window
            self.previousSearchesWindow = tkinter.Toplevel(global_settings.root, background='white')
            #self.previousSearchesWindow.lift()

            self.previousSearchesWindow.rowconfigure(0, weight=1)
            self.previousSearchesWindow.columnconfigure(0, weight=1)
            # Leaves only the label and removes the app window
            self.previousSearchesWindow.overrideredirect(True)
            self.previousSearchesWindow.wm_geometry("+%d+%d" % (x+5, y))
            #label1 = tkinter.Label(self.previousSearchesWindow, font=global_settings.Font_tuple_Arial_10, text='TESTE', justify='left',
            #               background='#ededd3', relief='solid', borderwidth=0)
            #label1.grid(row=0, column=0, sticky='w', pady=0)
            self.previoussearcheslistbox=None
        if(len(self.previous_simple_searches) > 0 and self.previoussearcheslistbox==None):            
            self.previoussearcheslistbox = tkinter.Listbox(self.previousSearchesWindow, listvariable = None, selectmode= 'simple', exportselection=False)
            self.previoussearcheslistbox.bind('<<ListboxSelect>>', self.previousSearchSelected)
            
            self.previoussearcheslistbox.grid(row=0, column=0, sticky='nsew', pady=0, padx=0)
            self.scrollprevioussearchesv = ttk.Scrollbar(self.previousSearchesWindow, orient="vertical")
            self.scrollprevioussearchesv.grid(row=0, column=1, sticky='ns')
            self.scrollprevioussearchesv.config( command = self.previoussearcheslistbox.yview )
            self.previoussearcheslistbox.configure(yscrollcommand=self.scrollprevioussearchesv.set)
            #--
            self.scrollprevioussearchesh = ttk.Scrollbar(self.previousSearchesWindow, orient="horizontal")
            self.scrollprevioussearchesh.grid(row=1, column=0, sticky='ew')
            self.scrollprevioussearchesh.config( command = self.previoussearcheslistbox.xview )
            self.previoussearcheslistbox.configure(xscrollcommand=self.scrollprevioussearchesh.set)
        if(len(self.previous_simple_searches) > 0):
            self.previoussearcheslistbox.delete(0,tkinter.END) 
            for k in range(len(self.previous_simple_searches)-1,-1, -1):
                termoprevious = self.previous_simple_searches[k]
                self.previoussearcheslistbox.insert(tkinter.END, termoprevious)
            self.previousSearchesWindow.lift()
            self.previousSearchesWindow.deiconify()
            
                
    def free_searches(self):
        try:  
            if(global_settings.indexing_thread!=None and global_settings.indexing_thread.is_alive()):
                proc = None
                self.notebook.tab(1, state="disabled")
                self.simplesearch.config(state='disabled')
                self.searchbutton.config(state='disabled')
                self.nhp.config(state='disabled')
                self.php.config(state='disabled')
                try:
                    if(not global_settings.processados.empty()):
                        proc = global_settings.processados.get(0)
                        if(proc[0]=='clear_searches' and proc[1]):
                            self.notebook.tab(1, state="normal")
                            self.simplesearch.config(state='normal')
                            self.searchbutton.config(state='normal')
                            self.nhp.config(state='normal')
                            self.php.config(state='normal')
                            global_settings.indexingwindow.window.withdraw()
                            global_settings.initiate_search_process()
                        else:
                            if(proc[0]=='clear_searches' and not proc[1]):
                                global_settings.indexingwindow.window.lift()
                                self.notebook.tab(1, state="disabled")
                                self.simplesearch.config(state='disabled')
                                self.searchbutton.config(state='disabled')
                                self.nhp.config(state='disabled')
                                self.php.config(state='disabled')
                                global_settings.indexingwindow.window.lift()
                            elif(proc[0]=='update'): 
                                global_settings.indexingwindow.window.lift()
                                global_settings.indexingwindow.progressbar['value'] += proc[2]
                            elif(proc[0]=='erro'):
                                utilities_general.printlogexception(ex=proc[1])  
                                
                except:
                    None
                finally:
                    global_settings.root.after(300, self.free_searches)
            
            else:
                naoindexados = []
                for abs_path_pdf in global_settings.infoLaudo:
                    if(global_settings.infoLaudo[abs_path_pdf].status!='indexado'):
                        naoindexados.append(abs_path_pdf)
                if(len(naoindexados)> 0):
                    relsni = ""
                    for ni in naoindexados:
                        relsni += ni + "\n"
                    texto = "As buscas estão desativadas, pois ocorreu um erro de indexação nos relatorios:\n\n{}\nFavor remover e adicionar novamente os mesmos.".format(relsni)
                    utilities_general.below_right(utilities_general.popup_window(texto, False),self.docInnerCanvas.winfo_rooty())
                else:
                    self.notebook.tab(1, state="normal")
                    self.simplesearch.config(state='normal')
                    self.searchbutton.config(state='normal')
                    self.nhp.config(state='normal')
                    self.php.config(state='normal')
                    global_settings.indexingwindow.window.withdraw()
                    global_settings.initiate_search_process()
                
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
            global_settings.root.after(300, self.free_searches)
        finally:
            None
            
    def enhance_observations(self, respostaPagina):
        if(global_settings.pathpdfatual in self.allobs):
            for observation in self.allobs[global_settings.pathpdfatual]:
                if(respostaPagina.qualPagina < observation.paginainit or respostaPagina.qualPagina > observation.paginafim):     
                    continue
                enhancearea = False
                enhancetext = False
                iiditem = observation.idobs
                if(observation.tipo=='area'):
                    enhancearea = True
                elif(observation.tipo=='texto'):
                    enhancetext = True
                for p in range(observation.paginainit, observation.paginafim+1):  
                    if(p != respostaPagina.qualPagina or p not in global_settings.processed_pages):     
                        continue
                    posicoes_normalizadas = self.normalize_positions(p, observation.paginainit, observation.paginafim, \
                                             observation.p0x, observation.p0y, observation.p1x, observation.p1y)
                    posicaoRealX0 = posicoes_normalizadas[0]
                    posicaoRealY0 = posicoes_normalizadas[1]
                    posicaoRealX1 = posicoes_normalizadas[2]
                    posicaoRealY1 = posicoes_normalizadas[3]
                   
                    self.prepararParaQuads(p, posicaoRealX0, posicaoRealY0, posicaoRealX1, posicaoRealY1, color=self.colorenhancebookmark, \
                                           tag=['enhanceobs'+global_settings.pathpdfatual+str(p),'enhanceobs'+str(iiditem), 'enhanceobs'+str(iiditem)+str(p), "enhanceobs"], \
                                               apagar=False,  enhancetext=enhancetext, enhancearea=enhancearea, withborder=False, alt=observation.withalt)
                for annot in observation.annotations:
                    annotation = observation.annotations[annot]
                    if(annotation.conteudo!=''):
                        idannot = annotation.idannot
                        
                        for p in range(annotation.paginainit, annotation.paginafim+1): 
                            if(p != respostaPagina.qualPagina or p not in global_settings.processed_pages):    
                                continue                            
                            posicoes_normalizadas = self.normalize_positions(p, annotation.paginainit, annotation.paginafim, \
                                                     annotation.p0x, annotation.p0y, annotation.p1x, annotation.p1y)
                            posicaoRealX0 = posicoes_normalizadas[0]
                            posicaoRealY0 = posicoes_normalizadas[1]
                            posicaoRealX1 = posicoes_normalizadas[2]
                            posicaoRealY1 = posicoes_normalizadas[3]
                            self.prepararParaQuads(p, posicaoRealX0, posicaoRealY0, posicaoRealX1, posicaoRealY1, self.colorenhanceannotation, \
                                                   tag=['enhanceannot'+global_settings.pathpdfatual+str(p),\
                                                        str(iiditem)+'enhanceannot'+str(idannot),'enhanceannot'], apagar=False,  \
                                                       enhancetext=enhancetext, enhancearea=enhancearea, withborder=True, alt=False)
  

    def checkUpdates(self, event=None):       
        atual = round((self.vscrollbar.get()[0]*global_settings.infoLaudo[global_settings.pathpdfatual].len))  
        if(True):
            try:        
                atual = (self.vscrollbar.get()[0]*global_settings.infoLaudo[global_settings.pathpdfatual].len)    
                                           
                cl = math.ceil(atual)     
                fl = math.floor(atual)  
                atual = round(atual)
                coords = self.docInnerCanvas.coords(self.fakeLines[0])
                coords1 = self.docInnerCanvas.coords(self.fakeLines[1])
                desloc_em_paginas = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x * global_settings.zoom
                dy = ((atual) * desloc_em_paginas) - coords[1]                
                dy1 = ((atual+1) * desloc_em_paginas) - coords1[1]                
                if(atual >= global_settings.infoLaudo[global_settings.pathpdfatual].len-1):
                    dy1 = (self.scrolly-(self.hscrollbar.winfo_height()/2)  - coords1[1])
                    self.docInnerCanvas.move(self.fakeLines[0], 0, dy)
                else:
                    self.docInnerCanvas.move(self.fakeLines[0], 0, dy)
                    self.docInnerCanvas.move(self.fakeLines[1], 0, dy1)            
                    
                if(not global_settings.response_queuexml.empty()):                    
                    pdfatualnorm = utilities_general.get_normalized_path(global_settings.pathpdfatual)
                    respostaPagina = global_settings.response_queuexml.get()
                    if((respostaPagina.qualPagina >= (atual-math.floor(global_settings.minMaxLabels/2)) and \
                            respostaPagina.qualPagina <= (atual+math.ceil(global_settings.minMaxLabels/2))) and \
                            respostaPagina.qualPdf == global_settings.pathpdfatual):  
                        global_settings.infoLaudo[global_settings.pathpdfatual].images[respostaPagina.qualPagina] = respostaPagina.images
                        global_settings.infoLaudo[global_settings.pathpdfatual].links[respostaPagina.qualPagina] = respostaPagina.links
                        global_settings.infoLaudo[global_settings.pathpdfatual].widgets[respostaPagina.qualPagina] = respostaPagina.widgets
                        global_settings.infoLaudo[global_settings.pathpdfatual].mapeamento[respostaPagina.qualPagina] = respostaPagina.mapeamento
                        global_settings.infoLaudo[global_settings.pathpdfatual].quadspagina[respostaPagina.qualPagina] = respostaPagina.quadspagina
                        if(self.showbookmarsboolean):
                            #None
                            self.enhance_observations(respostaPagina)                         
                if(not global_settings.response_queue.empty()):
                    
                    respostaPagina = global_settings.response_queue.get()
                    if((respostaPagina.qualPagina >= (atual-math.floor(global_settings.minMaxLabels/2)) and \
                            respostaPagina.qualPagina <= (atual+math.ceil(global_settings.minMaxLabels/2))) and \
                            respostaPagina.qualPdf == global_settings.pathpdfatual and respostaPagina.zoom == self.zoom_x*global_settings.zoom):   
                        indice = (respostaPagina.qualLabel * global_settings.divididoEm) + respostaPagina.qualGrid
                        zoom = global_settings.zoom
                        altura = math.floor((respostaPagina.qualPagina*global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh*self.zoom_x*global_settings.zoom) + \
                                            ((respostaPagina.qualGrid/global_settings.divididoEm)*global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh*self.zoom_x*global_settings.zoom))
                        coords = self.docInnerCanvas.coords(self.ininCanvasesid[indice])
                        pos_h = self.docInnerCanvas.winfo_x()
                        self.docInnerCanvas.coords(self.ininCanvasesid[indice], pos_h, altura)                        
                        self.tkimgs[indice] = tkinter.PhotoImage(data = respostaPagina.imgdata)
                        self.docInnerCanvas.itemconfig(self.ininCanvasesid[indice], image = self.tkimgs[indice])
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)
            finally:                
                global_settings.root.after(10, self.checkUpdates)
                
    def treeSeachAfter(self):
        resultsearch = None
        intervalo = 300
        try:
            contagem = 0
            starttime = time.process_time_ns()
            endtime = time.process_time_ns()
            while(((endtime-starttime)/1000000) < 200): 
                endtime = time.process_time_ns()
                if(not global_settings.result_queue.empty()):
                    try:
                        contagem +=1
                        res = global_settings.result_queue.get()
                        #adicionado a llista de buscas
                        intervalo = 20
                        if(res[0]==0):
                            resultsearch = res[1]
                            idtermo = resultsearch.idtermo
                            termo = resultsearch.termo.strip()
                            iidx = 't'+idtermo
                            if(not iidx in global_settings.idtermopdfs):
                                global_settings.idtermopdfs[iidx] = {}
                            tipobusca = resultsearch.tipobusca
                            if(not self.treeviewSearches.exists(iidx)):
                                self.treeviewSearches.insert(parent='', index=-1*int(idtermo), iid=iidx, text=termo.upper(), tag=('termosearching',), image=global_settings.lupa, \
                                                             values=(termo, tipobusca, 0))
                                global_settings.searchResultsDict['t'+idtermo] = resultsearch
                        elif(res[0]==0.5):
                            idtermo = str(res[1])
                            idx = 't'+idtermo
                            if(not idx in global_settings.idtermopdfs):
                                global_settings.idtermopdfs[idx] = {}
                            if(self.treeviewSearches.exists(idx)):
                                self.treeviewSearches.item(idx, tags=("termoinprocess",))
                        #terminou de receber resultados de buscas para o termo
                        elif(res[0]==2):
                            resultsearch = res[1]
                            idtermo = resultsearch.idtermo
                            resultsearch = res[1]
                            idtermo = resultsearch.idtermo
                            idx = 't'+resultsearch.idtermo
                            termo = resultsearch.termo
                            tipobusca = resultsearch.tipobusca
                            self.treeviewSearches.item(idx, tags=("termosearch",))
                            termo  = resultsearch.termo
                            th = utilities_general.countChildren(self.treeviewSearches, idx)   
                            self.treeviewSearches.item(idx, text=termo.upper() + ' (' + str(th) + ')'  + " - "+tipobusca)  
                            self.treeviewSearches.item(idx, tags=("termosearch",))
                            valores = self.treeviewSearches.item(idx, 'values')
                            self.treeviewSearches.item(idx, values=(valores[0], valores[1], th, termo.upper(), tipobusca,))
                            #for child in self.treeviewSearches.get_children(idx):
                                
                            #    self.treeviewSearches.item(idx, values=(valores[0], valores[1], th, termo.upper(),))
                            
                        elif(res[0]==1):
                            try:
                                for resultsearch in res[1]:
                                #resultsearch = res[1]
                                    idtermo = resultsearch.idtermo
                                    if(not self.treeviewSearches.exists('t'+idtermo)):
                                        continue
                                    else:
                                        termo = resultsearch.termo
                                        idtermopdf = resultsearch.idtermopdf
                                        pathpdfbase = os.path.basename(resultsearch.pathpdf)
                                        
                                        snippet = resultsearch.snippet[0] + resultsearch.snippet[1] + resultsearch.snippet[2]                                    
                                        pagina = resultsearch.pagina  
                                        t = 't'+resultsearch.idtermo
                                        tp = 'tp'+resultsearch.idtermopdf
                                        if(not t in global_settings.idtermopdfs):
                                            global_settings.idtermopdfs[t] = {}
                                        
                                        global_settings.idtermopdfs[t][(tp, pathpdfbase)] = []
                                        tptoc = resultsearch.tptoc
                                        if(not self.treeviewSearches.exists(tp)):  
                                            index_insert = self.insertIndex(self.treeviewSearches, t, pathpdfbase, 0)
                                            self.treeviewSearches.insert(parent=t, index=index_insert, iid=tp, \
                                                                         text=pathpdfbase, tag=('relsearch'), image=global_settings.imagereportb, \
                                                                             values=(pathpdfbase, resultsearch.idpdf,))                                     
                                        
                                        if(tptoc!=None):
                                            
                                            desloc = resultsearch.pagina * global_settings.infoLaudo[resultsearch.pathpdf].pixorgh
                                            #tocname = str(utilities_general.locateToc(resultsearch.pagina, resultsearch.pathpdf, init=resultsearch.init))
                                            tocname = resultsearch.toc
                                            snippettotal = str(pagina+1)+' - '+snippet
                                            if(not self.treeviewSearches.exists(tptoc)):
                                                 self.treeviewSearches.insert(parent=tp, iid=tptoc, text=tocname, index='end', tag=('relsearchtoc'), values=(tocname,))    
                                            
                                            idx = self.treeviewSearches.insert(parent=tptoc, index='end', text=snippettotal, tag='resultsearch', \
                                                                                image=global_settings.snippet,values=(resultsearch.snippet[0], resultsearch.snippet[1], \
                                                                                                            resultsearch.snippet[2],str(pagina+1),))
                                            tamanho = len(snippettotal)*4+150
                                            
                                            if(tamanho>self.maiorresult):
                                                self.maiorresult = tamanho                                                
                                                self.treeviewSearches.column("#0", width=self.maiorresult, stretch=True, minwidth=self.maiorresult, anchor="w")
                                            resultsearch.snippet = ""
                                            global_settings.searchResultsDict[idx] = resultsearch
                                        else:
                                            idx = self.treeviewSearches.insert(parent=tp, index='end', text=' '+str(pagina+1)+' - '+snippet, tag='resultsearch', \
                                                                                image=global_settings.snippet, values=(resultsearch.snippet[0], resultsearch.snippet[1], \
                                                                                                            resultsearch.snippet[2],str(pagina+1),))
                                            snippettotal = str(pagina+1)+' - '+snippet                                        
                                            tamanho = len(snippettotal)*4+150                                        
                                            if(tamanho>self.maiorresult):
                                                self.maiorresult = tamanho                                                
                                                self.treeviewSearches.column("#0", width=self.maiorresult, stretch=True, minwidth=self.maiorresult, anchor="w")
                                            resultsearch.snippet = ""
                                            global_settings.searchResultsDict[idx] = resultsearch
                            except Exception as ex:
                                utilities_general.printlogexception(ex=ex)
                    except Exception as ex:
                        utilities_general.printlogexception(ex=ex)
                else: 
                    break            
            if(contagem>0):
                endtime = time.process_time_ns()
        except:
            None
        finally:    
            #if(resultsearch!=None):        
            global_settings.root.after(intervalo, self.treeSeachAfter)   
            #else:
            #    global_settings.root.after(300, self.treeSeachAfter)
          
    def checkPages(self):
        try:
            atual = math.ceil((self.vscrollbar.get()[0]*global_settings.infoLaudo[global_settings.pathpdfatual].len))
            if(atual not in global_settings.processed_pages): 
                self.alreadyenhanced = set()
                for i in range(global_settings.minMaxLabels):
                   global_settings.processed_pages[i] = -1
                while not global_settings.request_queuexml.empty():
                    try:
                        global_settings.request_queuexml.get(0)
                    except Exception as ex:
                        None
                
                while not global_settings.request_queue.empty():
                    try:
                        global_settings.request_queue.get(0) 
                    except Exception as ex:
                        None
                while not global_settings.response_queue.empty():
                    try:
                        global_settings.response_queue.get(0) 
                    except Exception as ex:
                        None
                self.docInnerCanvas.delete("enhanceobs"+global_settings.pathpdfatual+str(global_settings.processed_pages[atual % global_settings.minMaxLabels]))
                self.clearSomeImages(["enhanceobs"+global_settings.pathpdfatual+str(global_settings.processed_pages[atual % global_settings.minMaxLabels])])                
                self.docInnerCanvas.delete('enhanceannot'+global_settings.pathpdfatual+str(global_settings.processed_pages[atual % global_settings.minMaxLabels]))
                self.clearSomeImages(["enhanceannot"+global_settings.pathpdfatual+str(global_settings.processed_pages[atual % global_settings.minMaxLabels])])
                global_settings.processed_pages[atual % global_settings.minMaxLabels] = atual
                pedido2 = classes_general.PedidoDePagina(qualLabel = atual % global_settings.minMaxLabels, qualPdf = global_settings.pathpdfatual, qualPagina = atual, matriz = self.mat, \
                  pixheight = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh, pixwidth = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw, zoom = self.zoom_x*global_settings.zoom, \
                      scrollvalue = self.vscrollbar.get()[0] ,\
                      scrolltotal = self.scrolly, canvash = self.canvash, mt = global_settings.infoLaudo[global_settings.pathpdfatual].mt, \
                          mb = global_settings.infoLaudo[global_settings.pathpdfatual].mb, me = global_settings.infoLaudo[global_settings.pathpdfatual].me, md = global_settings.infoLaudo[global_settings.pathpdfatual].md)
                global_settings.request_queuexml.put(pedido2)
                global_settings.request_queue.put(pedido2) 
                global_settings.processed_requests[pedido2.qualLabel] = pedido2
            for i in range(1, math.ceil(global_settings.minMaxLabels/2)):                  
                if(atual + i < global_settings.infoLaudo[global_settings.pathpdfatual].len):
                    if((atual + i) not in global_settings.processed_pages):                      
                       
                        global_settings.pathpdfatual2 = global_settings.pathpdfatual
                        self.docInnerCanvas.delete("enhanceobs"+global_settings.pathpdfatual+str(global_settings.processed_pages[(atual + i) % global_settings.minMaxLabels]))
                        self.clearSomeImages(["enhanceobs"+global_settings.pathpdfatual+str(global_settings.processed_pages[(atual + i) % global_settings.minMaxLabels])])
                        self.docInnerCanvas.delete("enhanceannot"+global_settings.pathpdfatual+str(global_settings.processed_pages[(atual + i) % global_settings.minMaxLabels]))
                        self.clearSomeImages(["enhanceannot"+global_settings.pathpdfatual+str(global_settings.processed_pages[(atual + i) % global_settings.minMaxLabels])])
                        if(global_settings.processed_pages[(atual + i) % global_settings.minMaxLabels] in self.alreadyenhanced):
                            self.alreadyenhanced.remove(global_settings.processed_pages[(atual + i) % global_settings.minMaxLabels])
                        global_settings.processed_pages[(atual + i) % global_settings.minMaxLabels] = (atual+i)
                        pedido3 = classes_general.PedidoDePagina(qualLabel = (atual + i) % global_settings.minMaxLabels, qualPdf = global_settings.pathpdfatual2, qualPagina = atual + i, matriz = self.mat, \
                                                 pixheight = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh, pixwidth = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw, zoom = self.zoom_x*global_settings.zoom, \
                                                     scrollvalue = self.vscrollbar.get()[0] ,\
                                                         scrolltotal = self.scrolly, canvash = self.canvash, mt = global_settings.infoLaudo[global_settings.pathpdfatual].mt, \
                                                             mb = global_settings.infoLaudo[global_settings.pathpdfatual].mb, me = global_settings.infoLaudo[global_settings.pathpdfatual].me, md = global_settings.infoLaudo[global_settings.pathpdfatual].md)
                        global_settings.request_queuexml.put(pedido3)
                        global_settings.request_queue.put(pedido3) 
                        global_settings.processed_requests[pedido3.qualLabel] = pedido3
                if(atual-i >= 0):
                    if((atual - i) not in global_settings.processed_pages):                        
                        global_settings.pathpdfatual2 = global_settings.pathpdfatual
                        self.docInnerCanvas.delete("enhanceobs"+global_settings.pathpdfatual+str(global_settings.processed_pages[(atual - i) % global_settings.minMaxLabels]))
                        self.clearSomeImages(["enhanceobs"+global_settings.pathpdfatual+str(global_settings.processed_pages[(atual - i) % global_settings.minMaxLabels])])
                        self.docInnerCanvas.delete("enhanceannot"+global_settings.pathpdfatual+str(global_settings.processed_pages[(atual + i) % global_settings.minMaxLabels]))
                        self.clearSomeImages(["enhanceannot"+global_settings.pathpdfatual+str(global_settings.processed_pages[(atual + i) % global_settings.minMaxLabels])])
                        if(global_settings.processed_pages[(atual - i) % global_settings.minMaxLabels] in self.alreadyenhanced):
                            self.alreadyenhanced.remove(global_settings.processed_pages[(atual - i) % global_settings.minMaxLabels])
                            
                        pedido4 = classes_general.PedidoDePagina(qualLabel = (atual - i) % global_settings.minMaxLabels, qualPdf = global_settings.pathpdfatual2, qualPagina = atual -i, matriz = self.mat, \
                                                 pixheight = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh, pixwidth = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw, zoom = self.zoom_x*global_settings.zoom, \
                                                     scrollvalue = self.vscrollbar.get()[0] ,\
                                                         scrolltotal = self.scrolly, canvash = self.canvash, mt = global_settings.infoLaudo[global_settings.pathpdfatual].mt, \
                                                             mb = global_settings.infoLaudo[global_settings.pathpdfatual].mb, me = global_settings.infoLaudo[global_settings.pathpdfatual].me, md = global_settings.infoLaudo[global_settings.pathpdfatual].md)
                        global_settings.processed_pages[(atual - i) % global_settings.minMaxLabels] = (atual-i)
                        global_settings.request_queuexml.put(pedido4)
                        global_settings.request_queue.put(pedido4)
                        global_settings.processed_requests[(atual - i) % global_settings.minMaxLabels] = pedido4
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        finally:
            global_settings.root.after(10, self.checkPages)
    
    def treeview_open(self, event):
        iid = self.treeviewEqs.focus()
        valores = (self.treeviewEqs.item(iid, 'values'))
        opcao = valores[0] 
        if(opcao=="eq"):
            children = self.treeviewEqs.get_children(iid)
            for child in children:
                self.treeviewEqs.item(child, open=False)
            self.treeviewEqs.item(iid, open=True)
            
    def treeview_selection_loc_right(self, event=None, item=None):
        iid = self.treeviewLocs.identify_row(event.y)  
        self.treeviewLocs.selection_set(iid)
        if(self.treeviewLocs.tag_has('locationlp', iid)):
            return
        valores = (self.treeviewLocs.item(iid, 'values'))[1:]
        if("FERA" in self.treeviewLocs.item(iid)['text']):
            self.treeview_selection_loc(valores)
        else:
            webbrowser.open(valores[0])
            
    def treeview_selection_loc(self, valores):
        
        try: 
            
            pid = os.getpid()
            if(valores not in self.locations_toplevels or not self.locations_toplevels[valores].is_alive()):
                
                processo = mp.Process(target=utilities_general.show_locations, args=(valores[0], valores[1], pid,), daemon=True)
                processo.start()
                self.locations_toplevels[valores] = processo
            else:
                try:
                    
                        win = gw.getWindowsWithTitle(f"P{pid} - {valores[1]}")[0]
                        win.activate()
                except:
                    None
            
                
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
            
    def treeview_selection(self, event=None, item=None):
        if(event!=None):
            iid = self.treeviewEqs.identify_row(event.y)
            if(iid==None or iid==''):
                return
        try: 
            for pdf in global_settings.infoLaudo:
                global_settings.infoLaudo[pdf].retangulosDesenhados = {}
            selecao = None
            if(item==None):
                selecao = self.treeviewEqs.focus()                
            else:
                selecao = item
            self.treeviewEqs.selection_set(selecao)
            self.treeviewEqs.focus(selecao)
            valores = (self.treeviewEqs.item(selecao, 'values'))
            opcao = None       
            global_settings
            if(valores!=None and valores!='' and len(valores)>0):
                opcao = valores[0] 
            if(opcao=='pdf'):
                newpath = utilities_general.get_normalized_path(valores[1])
                try:
                    self.positions[self.indiceposition] = (global_settings.pathpdfatual, self.vscrollbar.get()[0])
                    self.indiceposition += 1
                    if(self.indiceposition>=10):
                        self.indiceposition = 0
                except Exception as ex:
                    None
                if(global_settings.pathpdfatual!=newpath or item!=None):
                    pdfantigo = global_settings.pathpdfatual
                    for i in range(global_settings.minMaxLabels):
                        global_settings.processed_pages[i] = -1
                    sobraEspaco = 0
                    if(self.docFrame.winfo_width() > global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw * self.zoom_x * global_settings.zoom):
                        sobraEspaco = self.docInnerCanvas.winfo_x()
                    self.maiorw = self.docFrame.winfo_width()
                    if(global_settings.infoLaudo[newpath].pixorgw*self.zoom_x*global_settings.zoom>self.maiorw):
                        self.maiorw = global_settings.infoLaudo[newpath].pixorgw*self.zoom_x *global_settings.zoom           
                    self.scrolly = global_settings.infoLaudo[newpath].pixorgh*self.zoom_x*global_settings.zoom*global_settings.infoLaudo[newpath].len  - 35
                    self.docInnerCanvas.config(scrollregion=(sobraEspaco, 0, sobraEspaco+ (global_settings.infoLaudo[newpath].pixorgw*global_settings.zoom*self.zoom_x), self.scrolly))
                    pagina = round(global_settings.infoLaudo[newpath].ultimaPosicao*global_settings.infoLaudo[newpath].len)   
                    global_settings.pathpdfatual = utilities_general.get_normalized_path(newpath)
                    if(global_settings.infoLaudo[pdfantigo].zoom_pos!=global_settings.infoLaudo[global_settings.pathpdfatual].zoom_pos):
                        self.zoomx(None, None, pdfantigo, global_settings.infoLaudo[newpath].ultimaPosicao)
                    else:
                        self.docInnerCanvas.yview_moveto(global_settings.infoLaudo[newpath].ultimaPosicao)
                    if(str(pagina+1)!=self.pagVar.get()):
                        self.pagVar.set(str(pagina+1))
                    
                    try:
                        global_settings.docatual.close()
                    except Exception as ex:
                        None
                    global_settings.docatual = fitz.open(global_settings.pathpdfatual)
                    self.labeldocname.config(font=global_settings.Font_tuple_Arial_10, text=os.path.basename(global_settings.pathpdfatual))
                    self.totalPgg.config(font=global_settings.Font_tuple_Arial_10, text="/ "+str(global_settings.infoLaudo[global_settings.pathpdfatual].len))
                    self.clearAllImages()
                    for pdf in global_settings.infoLaudo:
                        global_settings.infoLaudo[pdf].retangulosDesenhados = {}
                    
                if(event!=None):
                    self.treeviewEqs.selection_set(iid)
                else:
                    self.treeviewEqs.selection_set(item)                                    
            elif(opcao=="toc"): 
                pdfantigo = global_settings.pathpdfatual
                newpath = valores[1]
                eq = selecao[3]
                toc = selecao[2]
                self.positions[self.indiceposition] = (global_settings.pathpdfatual, self.vscrollbar.get()[0])
                self.indiceposition += 1
                if(self.indiceposition>=10):
                    self.indiceposition = 0
                if(global_settings.pathpdfatual!=newpath):
                    global_settings.pathpdfatual =utilities_general.get_normalized_path(newpath)
                    self.totalPgg.config(font=global_settings.Font_tuple_Arial_10, text="/ "+str(global_settings.infoLaudo[global_settings.pathpdfatual].len))
                    try:
                        global_settings.docatual.close()
                    except Exception as ex:
                        None
                    global_settings.docatual = fitz.open(global_settings.pathpdfatual)
                    self.clearAllImages()
                    self.docwidth = self.docOuterFrame.winfo_width()
                    for i in range(global_settings.minMaxLabels):
                        global_settings.processed_pages[i] = -1
                    sobraEspaco = 0
                    if(self.docFrame.winfo_width() > global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw * self.zoom_x * global_settings.zoom):
                        sobraEspaco = self.docInnerCanvas.winfo_x() 
                    self.maiorw = self.docFrame.winfo_width()
                    if(global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw*self.zoom_x*global_settings.zoom>self.maiorw):
                        self.maiorw = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw*self.zoom_x *global_settings.zoom           
                    self.scrolly = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh*self.zoom_x*global_settings.zoom*global_settings.infoLaudo[global_settings.pathpdfatual].len  - 35
                    self.docInnerCanvas.config(scrollregion=(sobraEspaco, 0, sobraEspaco+ (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw*global_settings.zoom*self.zoom_x), self.scrolly))
                    self.labeldocname.config(font=global_settings.Font_tuple_Arial_10, text=os.path.basename(global_settings.pathpdfatual))
                    self.totalPgg.config(font=global_settings.Font_tuple_Arial_10, text="/ "+str(global_settings.infoLaudo[global_settings.pathpdfatual].len))
                pagina = int(valores[3])
                deslocy = float(valores[4])
                ondeir = (float(pagina) / (global_settings.infoLaudo[global_settings.pathpdfatual].len)+(deslocy*self.zoom_x*global_settings.zoom)/self.scrolly)
                if(global_settings.infoLaudo[pdfantigo].zoom_pos!=global_settings.infoLaudo[global_settings.pathpdfatual].zoom_pos):
                    self.zoomx(None, None, pdfantigo)
                #global_settings.root.after(1, lambda: self.docInnerCanvas.yview_moveto(ondeir))
                self.docInnerCanvas.yview_moveto(ondeir)
                if(str(pagina+1)!=self.pagVar.get()):
                    self.pagVar.set(str(pagina+1))
                if(event!=None):
                    self.treeviewEqs.selection_set(iid)
                else:
                    self.treeviewEqs.selection_set(item)            
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
    
    def copyName(self, iid):
        iid = os.path.basename(utilities_general.get_normalized_path(iid))
        clipboard.copy(iid.strip())
    
    def copyTOC(self, iid):
        memoriainterna = False
        string = ""
        tocs = []
        iid = utilities_general.get_normalized_path(iid)
        for tindex in range(len(global_settings.infoLaudo[iid].toc)):
            toc = global_settings.infoLaudo[iid].toc[tindex]
            if("1 Aparelho" in toc[0]):
                memoriainterna = True
            elif("2 "==toc[0][0:2]):
                tocs.append(toc)
                break
            elif(memoriainterna):
                tocunit = toc[0].split(" ")[0]
                tocunitsplit = tocunit.split(".")
                if(len(tocunitsplit)!=2):
                    continue
                else:
                    tocs.append(toc)
        memoriainterna = False
        for tindex in range(len(tocs)-1):
            toc = tocs[tindex]
            tocunit = toc[0].split(" ")[0]
            tocunitsplit = tocunit.split(".")
            if(len(tocunitsplit)!=2):
                continue
            tocunit = toc[0].split(" ")[0]
            toctext = toc[0][len(tocunit+" "):]
            if(tindex==len(tocs)-2):
                tocnext = tocs[tindex+1]
                tocnextpage = tocnext[1]
                string += "O(s) registro(s) de {} encontra(m)-se na subseção {}, Fls. {} a {}.\n\r".format(toctext, tocunit, toc[1]+1, tocnextpage)
            else:
                tocnext = tocs[tindex+1]
                if(tocnext[2]<135):  
                    if(toc[1] == tocnext[1]-1):
                        string += "O(s) registro(s) de {} encontra(m)-se na subseção {}, Fls. {}.\r".format(toctext, tocunit, toc[1]+1)                        
                    else:
                        string += "O(s) registro(s) de {} encontra(m)-se na subseção {}, Fls. {} a {}.\r".format(toctext, tocunit, toc[1]+1, tocnext[1])
                else:
                    if(toc[1]+1 == tocnext[1]+1):
                        string += "O(s) registro(s) de {} encontra(m)-se na subseção {}, Fls. {}.\r".format(toctext, tocunit, toc[1]+1)
                        
                    else:
                       string += "O(s) registro(s) de {} encontra(m)-se na subseção {}, Fls. {} a {}.\r".format(toctext, tocunit, toc[1]+1, tocnext[1]+1) 
        clipboard.copy(string.strip())
    def treeview_eqs_right(self, event=None):
        try:
            iid = self.treeviewEqs.identify_row(event.y)  
            if(self.treeviewEqs.tag_has('reportlp', iid)):
                basepdf = self.treeviewEqs.item(iid, 'values')[1]
                self.eqmenu = tkinter.Menu(global_settings.root, tearoff=0)
                
               
                self.eqmenu.add_command(label='Copiar Índice', image=global_settings.copycat, compound='left',  command=lambda iid=iid: self.copyTOC(iid)) 
                self.eqmenu.add_command(label='Copiar Nome', image=global_settings.copycat, compound='left',  command=lambda iid=iid: self.copyName(iid))
                self.eqmenu.add_command(label='Mover para', image=global_settings.copycat, compound='left',  command=lambda iid=iid: self.change_eq_name(iid))
                try:
                    if(isinstance(event.widget, ttk.Treeview)):
                        self.eqmenu.tk_popup(event.x_root, event.y_root) 
                except Exception as ex:
                    utilities_general.printlogexception(ex=ex) 
                finally:
                    self.eqmenu.grab_release()  
            if(self.treeviewEqs.tag_has('equipmentlp', iid)):
                self.eqmenu = tkinter.Menu(global_settings.root, tearoff=0)
                self.eqmenu.add_command(label='Renomear', image=global_settings.copycat, compound='left',  command=lambda iid=iid: self.change_eq_name(iid))
                try:
                    if(isinstance(event.widget, ttk.Treeview)):
                        self.eqmenu.tk_popup(event.x_root, event.y_root) 
                except Exception as ex:
                    utilities_general.printlogexception(ex=ex) 
                finally:
                    self.eqmenu.grab_release() 
        except Exception as ex: 
            utilities_general.printlogexception(ex=ex)
            
    def invert_first_second_layer_tree():
        None
    
    def change_eq_name(self, iid):
        texto_candidato = ""
        idpdfs = []
        if(self.treeviewEqs.tag_has('equipmentlp', iid)):
            texto_candidato = self.treeviewEqs.item(iid)['text']
            w=classes_general.popupWindow(global_settings.root,texto_candidato)      
            global_settings.root.wait_window(w.top)
            if(w.value!=None and w.value.strip()!=''):
                children = None
                children = self.treeviewEqs.get_children(iid)
                
                new_alias = w.value.strip().upper()
                exists = self.treeviewEqs.exists(new_alias)
                for child in children:
                    p = self.treeviewEqs.item(child, 'values')[1]
                    global_settings.infoLaudo[p].alias_parent = w.value.upper()
                    idpdfs.append(global_settings.infoLaudo[p].id)                
                    if(exists):
                        index_insert = self.insertIndex(self.treeviewEqs, new_alias, os.path.basename(p), 0)
                        self.treeviewEqs.move(child, new_alias, index_insert)
                if(not exists):
                    self.treeviewEqs.item(iid, text = w.value.upper())
                    self.treeviewEqs.item(iid, values = ('eq', w.value.upper()))                   
                else:
                    self.treeviewEqs.delete(iid)
        elif(self.treeviewEqs.tag_has('reportlp', iid)):
            texto_candidato = self.treeviewEqs.item(self.treeviewEqs.parent(iid))['text']
            w=classes_general.popupWindow(global_settings.root,texto_candidato)      
            global_settings.root.wait_window(w.top)
            if(w.value!=None and w.value.strip()!=''):
                new_alias = w.value.strip().upper()
                exists = self.treeviewEqs.exists(new_alias)
                old_parent = self.treeviewEqs.parent(iid)
                p = self.treeviewEqs.item(iid, 'values')[1]
                global_settings.infoLaudo[p].alias_parent = new_alias
                idpdfs.append(global_settings.infoLaudo[p].id) 
                if(not exists):
                    index_insert = self.insertIndex(self.treeviewEqs, '', new_alias, 0)
                    self.treeviewEqs.insert(parent='', index=index_insert, iid=new_alias, text=new_alias, \
                                            image=global_settings.imageequip, tag='equipmentlp', values=('eq', str(w.value.upper()), global_settings.infoLaudo[p].id,))
                index_insert = self.insertIndex(self.treeviewEqs, new_alias, os.path.basename(p), 0)
                self.treeviewEqs.move(iid, new_alias, index_insert)
                if(len(self.treeviewEqs.get_children(old_parent))==0):
                    self.treeviewEqs.delete(old_parent)
                    
            
        updateinto2 = "UPDATE Anexo_Eletronico_Pdfs set parent_alias = '{new_alias}' WHERE id_pdf IN ({seq})".\
            format(new_alias=new_alias, seq=','.join(['?']*len(idpdfs)))      
                
        sqliteconn = utilities_general.connectDB(str(global_settings.pathdb))
        try:
            cursor = sqliteconn.cursor()
            cursor.custom_execute(updateinto2, idpdfs)
            sqliteconn.commit()
        except Exception as ex:
            utilities_general.printlogexception(ex=ex) 
        finally:
            if(sqliteconn):
                sqliteconn.close()
    
    def tabOpened(self, event=None):
        texto = (self.notebook.tab(self.notebook.select(), 'text'))        
        if(texto=='Buscas (*)'):
            self.notebook.tab(self.notebook.select(), font=global_settings.Font_tuple_Arial_10, text="Buscas")
            self.eqmenu = tkinter.Menu(global_settings.root, tearoff=0)

    
    def exportMidiasFromObs(self):  
        
        iids = self.treeviewObs.selection()
        lista = self.IterateChildObs(iids, [])
        path = (askdirectory(initialdir=global_settings.pathdb.parent))
        
        window = tkinter.Toplevel()
        try:
            label = tkinter.Label(window, font=global_settings.Font_tuple_Arial_10, text="Exportando arquivos de observações!")
            label.pack(fill='x', padx=50, pady=20)
            progresssearch = ttk.Progressbar(window, mode='indeterminate')
            progresssearch.pack(fill='x', padx=50, pady=20)
            #progresssearch['maximum'] = pageend.get() - pageinit.get() +1
            window.lift()
            window.update_idletasks()
            pdfs = {}
            for l in lista:
                if(l[1] not in pdfs):
                    pdfs[l[1]] = {}
                    
            for l in lista:
                pi = int(l[2])
                pf = int(l[5])
                pdf = utilities_general.get_normalized_path(l[1])
                margemsup = (global_settings.infoLaudo[pdf].mt/25.4)*72
                margeminf = global_settings.infoLaudo[pdf].pixorgh-((global_settings.infoLaudo[pdf].mb/25.4)*72)
                for p in range(pi, pf+1):
                    if(p not in pdfs[pdf]):
                        pdfs[pdf][p] = []    
                    if (p> pi and p <pf):                    
                        pdfs[pdf][p].append((margemsup, margeminf))
                    elif (p==pi):
                        yinit = round(float(l[4]), 0)
                        yfim = round(float(l[7]), 0) if p==pf else margeminf
                        pdfs[pdf][p].append((yinit, yfim))    
                    elif (p==pf):
                        yinit = round(float(l[4]), 0) if p==pi else margemsup
                        yfim = round(float(l[7]), 0)
                        pdfs[pdf][p].append((margemsup, margeminf))
            
            for temppdf in pdfs:
                temp = fitz.open(temppdf)
                try:
                    for page in pdfs[temppdf]:
                        loadedPage = temp[page]
                        links = loadedPage.get_links()
                        for link in links:
                            r = link['from']
                            if('file' not in link):
                                continue
                            try:
                                arquivo  = link['file']
                                if("#" in arquivo):
                                    continue
                                if(arquivo==""):
                                    xref = link['xref']
                                    info = global_settings.docatual.xref_get_key(xref, 'A')
                                    grupos_search = global_settings.regex_actions_compiled.search(info[1])
                                    if(grupos_search==None):
                                        continue
                                    grupos = grupos_search.groups()
                                    if(grupos[-1]=="GoToR"):
                                        continue
                                    elif(grupos[-1]=="Launch"):
                                        arquivo = grupos[2]
                                pdfatualnorm = utilities_general.get_normalized_path(temppdf)
                                filepath = str(Path(utilities_general.get_normalized_path(os.path.join(Path(utilities_general.get_normalized_path(pdfatualnorm)).parent,arquivo))))
                                for obs in pdfs[temppdf][page]:
                                    pm = (r.y0 + r.y1) / 2
                                    if((r.y0 > obs[0] and r.y1 > obs[1] and r.y0 <= obs[1]) or\
                                       (r.y0 < obs[0] and r.y1 < obs[1] and r.y1 <= obs[1]) \
                                       or (pm >= obs[0] and pm <= obs[1])):
                                        shutil.copy2(filepath, os.path.join(path, os.path.basename(arquivo)))
                                        break
                            except Exception as ex:
                                utilities_general.printlogexception(ex=ex)
                            
                except Exception as ex:
                    utilities_general.printlogexception(ex=ex)
                finally:
                    temp.close()
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        finally:
            window.destroy()
        
    def IterateChildObs(self, iids, lista):
        for iid in iids:
            if(self.treeviewObs.tag_has('obsitem',iid)):
                values = self.treeviewObs.item(iid,'values')
                lista.append(values)
            else:
                children = self.treeviewObs.get_children(iid)
                self.IterateChildObs(children, lista)
        return lista
    
    def treeview_obs_right(self, event=None):        
        iid = self.treeviewObs.identify_row(event.y)  
        if(self.treeviewObs.tag_has('obscat',iid)):
            self.editdelcat = tkinter.Menu(global_settings.root, tearoff=0)            
            self.editdelcat.add_command(label="Editar Categoria", image=global_settings.editcat, compound='left', \
                                        command=lambda: self.add_edit_category('edit'))
            self.editdelcatmove = tkinter.Menu(self.editdelcat, tearoff=0)
            self.editdelcatmove.add_command(label="Mover para o topo", image=global_settings.movecattop, \
                                            compound='left', command=lambda: self.moveCategory('top', self.treeviewObs.selection()[0]))
            self.editdelcatmove.add_command(label="Mover para cima", image=global_settings.movecatup, \
                                            compound='left', command=lambda: self.moveCategory('up', self.treeviewObs.selection()[0]))
            self.editdelcatmove.add_command(label="Mover para baixo", image=global_settings.movecatdown, \
                                            compound='left', command=lambda: self.moveCategory('down', self.treeviewObs.selection()[0]))
            self.editdelcatmove.add_command(label="Mover para o fundo", image=global_settings.movecatbottom, \
                                            compound='left', command=lambda: self.moveCategory('bottom', self.treeviewObs.selection()[0]))
            self.editdelcat.add_cascade(label='Mover Categoria', menu=self.editdelcatmove, image=global_settings.movecat, compound='left')  
            self.copyobsto = tkinter.Menu(self.editdelcat, tearoff=0)
            self.copyobsto.add_command(label="Para o clipboard", image=global_settings.copycat, \
                                       compound='left', command=lambda: self.addcatpopup(None, 'copiarclip'))       
            self.copyobsto.add_command(label="Em formato CSV", image=global_settings.copycat, \
                                       compound='left', command=lambda: self.addcatpopup(None, 'copiarcsv'))       
            self.copyobsto.add_command(label="Clipboard (RTF)", image=global_settings.copycat, \
                                       compound='left', command=lambda: self.copy_obs_rtf(False))
            self.copyobsto.add_command(label="Clipboard (RTF) com anotações", image=global_settings.copycat, \
                                       compound='left', command=lambda: self.copy_obs_rtf(True))
            self.copyobsto.add_command(label="Clipboard (Imagens em Tabela - RTF)", image=global_settings.copycat, \
                                       compound='left', command=lambda: self.export_images_table_rtf(False))
            self.copyobsto.add_command(label="Clipboard (Imagens em Tabela - RTF) com anotações", image=global_settings.copycat, \
                                       compound='left', command=lambda: self.export_images_table_rtf(True))
            self.editdelcat.add_cascade(label='Copiar Páginas', menu=self.copyobsto, image=global_settings.copycat, compound='left') 
            self.editdelcat.add_command(label="Exportar mídias", image=global_settings.imtoFile, \
                                        compound='left', command= self.exportMidiasFromObs)
            if(self.treeviewObs.item(iid,'values')[0]=='0' or True):                
                self.editdelcat.add_separator()
                self.editdelcat.add_command(label="Excluir Categoria", image=global_settings.delcat, \
                                            compound='left', command=lambda: self.addcatpopup(None, 'exclude'))    
            self.treeviewObs.selection_set(iid)
            self.treeviewObs.focus(iid)
            try:
                if(isinstance(event.widget, ttk.Treeview)):
                    self.editdelcat.tk_popup(event.x_root, event.y_root) 
            except Exception as ex:
                utilities_general.printlogexception(ex=ex) 
            finally:
                self.editdelcat.grab_release()
        elif(self.treeviewObs.tag_has('tocobs',iid) or self.treeviewObs.tag_has('relobs',iid)):
            basepdf = None
            
            self.editdelcat = tkinter.Menu(global_settings.root, tearoff=0)              
            self.copyobsto = tkinter.Menu(self.editdelcat, tearoff=0)
            self.copyobsto.add_command(label="Para o clipboard", image=global_settings.copycat, compound='left', command=lambda: self.addcatpopup(None, 'copiarclip'))      
            self.copyobsto.add_command(label="Em formato CSV", image=global_settings.copycat, compound='left', command=lambda: self.addcatpopup(None, 'copiarcsv'))       
            self.copyobsto.add_command(label="Clipboard (RTF)", image=global_settings.copycat, \
                                       compound='left', command=lambda: self.copy_obs_rtf(False))
            self.copyobsto.add_command(label="Clipboard (RTF) com anotações", image=global_settings.copycat, \
                                       compound='left', command=lambda: self.copy_obs_rtf(True))
            self.copyobsto.add_command(label="Clipboard (Imagens em Tabela - RTF)", image=global_settings.copycat, \
                                       compound='left', command=lambda: self.export_images_table_rtf(False))
            self.copyobsto.add_command(label="Clipboard (Imagens em Tabela - RTF) com anotações", image=global_settings.copycat, \
                                       compound='left', command=lambda: self.export_images_table_rtf(True))
            if (self.treeviewObs.tag_has('relobs',iid)):
                basepdf = self.treeviewObs.item(iid,'values')[0]
                self.editdelcat.add_command(label='Copiar Nome do Relatório', image=global_settings.copycat, compound='left',  command=lambda iid=iid: self.copyName(basepdf))
            self.editdelcat.add_cascade(label='Copiar Páginas', menu=self.copyobsto, image=global_settings.copycat, compound='left')  
            self.editdelcat.add_command(label="Exportar mídias", image=global_settings.imtoFile, compound='left', command= self.exportMidiasFromObs)
            self.treeviewObs.selection_set(iid)
            self.treeviewObs.focus(iid)
            try:
                if(isinstance(event.widget, ttk.Treeview)):
                    self.editdelcat.tk_popup(event.x_root, event.y_root)  
            except Exception as ex:
                utilities_general.printlogexception(ex=ex) 
            finally:
                self.editdelcat.grab_release()
        elif((self.treeviewObs.tag_has('obsitem', iid) and len(self.treeviewObs.selection())==1)):
            self.treeviewObs.selection_set(iid)
            self.treeviewObs.focus(iid)
            #self.treeview_selection_obs(item=iid)
            try:
                if(isinstance(event.widget, ttk.Treeview)):
                    tagsss = self.treeviewObs.item(iid, 'tags')
                    alterado = False
                    self.delitemcat = tkinter.Menu(global_settings.root, tearoff=0)
                    if(self.treeviewObs.item(iid,'values')[0]=='0' or True):  
                        self.delitemcat.add_command(label="Excluir Marcação", image=global_settings.delcat, compound='left', command=lambda: self.addcatpopup(None, 'excludeitem'))
                    self.delitemcat.add_command(label="Alterar Categoria", image=global_settings.editcat, compound='left', command=lambda: self.addcatpopup(None, 'changecat'))
                    self.copyobsto = tkinter.Menu(self.delitemcat, tearoff=0)
                    self.copyobsto.add_command(label="Clipboard (RTF)", image=global_settings.copycat, \
                                               compound='left', command=lambda: self.copy_obs_rtf(False))
                    self.copyobsto.add_command(label="Clipboard (RTF) com anotações", image=global_settings.copycat, \
                                               compound='left', command=lambda: self.copy_obs_rtf(True))
                    self.copyobsto.add_command(label="Clipboard (Imagens em Tabela - RTF)", image=global_settings.copycat, \
                                               compound='left', command=lambda: self.export_images_table_rtf(False))
                    self.copyobsto.add_command(label="Clipboard (Imagens em Tabela - RTF) com anotações", image=global_settings.copycat, \
                                               compound='left', command=lambda: self.export_images_table_rtf(True))
                    
                    self.delitemcat.add_cascade(label='Copiar Páginas', menu=self.copyobsto, image=global_settings.copycat, compound='left')  
                    self.delitemcat.add_command(label="Exportar mídias", image=global_settings.imtoFile, compound='left', command= self.exportMidiasFromObs)
                    self.delitemcat.add_command(label="Adicionar/Editar Anotação", image=global_settings.withnote, compound='left', command= lambda : self.editadd_anottation(iid))
                    for tg in tagsss:
                        if('alterado' in tg):
                            alterado = True
                    if(self.treeviewObs.item(self.treeviewObs.selection()[0],'values')[9])=='0':
                        if(not alterado):
                            None
                        else:
                            self.delitemcat.add_command(label="Validar Observação", \
                                                        image=global_settings.checki, compound='left', command=lambda: self.addcatpopup(None, 'validarobs'))
                        self.delitemcat.tk_popup(event.x_root, event.y_root) 
                    elif(True):
                        if(not alterado):
                            None
                        else:
                            self.delitemcat.add_command(label="Validar Observação", image=global_settings.checki, compound='left', command=lambda: self.addcatpopup(None, 'validarobs'))
                        self.delitemcat.tk_popup(event.x_root, event.y_root) 
                
            except Exception as ex:
               utilities_general.printlogexception(ex=ex) 
            finally:
                self.delitemcat.grab_release()            
        elif(isinstance(event.widget, ttk.Treeview) and (self.treeviewObs.tag_has('obsitem', iid) and len(self.treeviewObs.selection())>1)): 
            id0 = self.treeviewObs.selection()[0]
            pai1 = self.treeviewObs.parent(iid)            
            for k in self.treeviewObs.selection():
                pai2 = self.treeviewObs.parent(k)
                if(pai2!=pai1):
                    return
            tagsss = self.treeviewObs.item(iid, 'tags')
            alterado = False
            self.delitemcat = tkinter.Menu(global_settings.root, tearoff=0)
            if(self.treeviewObs.item(iid,'values')[0]=='0' or True):  
                self.delitemcat.add_command(label="Excluir Marcações", image=global_settings.delcat, compound='left', \
                                            command=lambda: self.addcatpopup(None, 'excludeitems', self.treeviewObs.item(self.treeviewObs.selection()[0], 'text')))
            self.delitemcat.add_command(label="Alterar Categorias", image=global_settings.editcat, compound='left', command=lambda: self.addcatpopup(None, 'changecats'))
            self.copyobsto = tkinter.Menu(self.delitemcat, tearoff=0)
            self.copyobsto.add_command(label="Clipboard (RTF)", image=global_settings.copycat, \
                                       compound='left', command=lambda: self.copy_obs_rtf(False))
            self.copyobsto.add_command(label="Clipboard (RTF) com anotações", image=global_settings.copycat, \
                                       compound='left', command=lambda: self.copy_obs_rtf(True))
            self.copyobsto.add_command(label="Clipboard (Imagens em Tabela - RTF)", image=global_settings.copycat, \
                                       compound='left', command=lambda: self.export_images_table_rtf(False))
            self.copyobsto.add_command(label="Clipboard (Imagens em Tabela - RTF) com anotações", image=global_settings.copycat, \
                                       compound='left', command=lambda: self.export_images_table_rtf(True))
            self.delitemcat.add_cascade(label='Copiar Páginas', menu=self.copyobsto, image=global_settings.copycat, compound='left') 
            self.delitemcat.add_command(label="Exportar mídias", image=global_settings.imtoFile, compound='left', command= self.exportMidiasFromObs)
            for tg in tagsss:
                if('alterado' in tg):
                    alterado = True
            if(self.treeviewObs.item(self.treeviewObs.selection()[0],'values')[9])=='0':
                if(not alterado):
                    None
                else:
                    self.delitemcat.add_command(label="Validar Observação", image=global_settings.checki, compound='left', command=lambda: self.addcatpopup(None, 'validarobs'))
    
                self.delitemcat.tk_popup(event.x_root, event.y_root) 
            elif(True):
                if(not alterado):
                    None
                else:
                    self.delitemcat.add_command(label="Validar Observação", image=global_settings.checki, compound='left', command=lambda: self.addcatpopup(None, 'validarobs'))
    
                self.delitemcat.tk_popup(event.x_root, event.y_root) 
    
             
    
    def treeview_search_right(self, event=None):
        iid = self.treeviewSearches.identify_row(event.y)  
        if(self.treeviewSearches.parent(iid)=='' and self.treeviewSearches.item(iid, 'text') != '' and False):
            self.treeviewSearches.selection_set(iid)
            try:
                if(isinstance(event.widget, ttk.Treeview)):
                    resultsearch = global_settings.searchResultsDict[iid]
                    if(resultsearch.fixo):
                        self.menuexcludesearch.tk_popup(event.x_root, event.y_root) 
                    elif(True):
                        self.menuexcludesearch.tk_popup(event.x_root, event.y_root) 
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)
            finally:
                self.menuexcludesearch.grab_release()
        elif(iid != ''):
            #self.treeviewSearches.selection_set(iid)
            try:
                if(isinstance(event.widget, ttk.Treeview)):
                    
                    self.menuexportsearchtobs = tkinter.Menu(global_settings.root, tearoff=0)
                    if(self.treeviewSearches.tag_has('relsearch', iid)):
                        basepdf = self.treeviewSearches.item(iid,'values')[0]
                        self.menuexportsearchtobs.add_command(label='Copiar Nome do Relatório', image=global_settings.copycat, compound='left',  command=lambda iid=iid: self.copyName(basepdf))
                    getobscatas =  self.treeviewObs.get_children('')
                    menucats = tkinter.Menu(self.menuexportsearchtobs, tearoff=0)
                    for obscat in getobscatas:
                        #texto = self.treeviewObs.item(obscat, 'text')
                        cat_text = self.treeviewObs.item(obscat, 'text')
                        idobscat = self.treeviewObs.item(obscat, 'values')[1]
                        menucats.add_command(label=cat_text, image=global_settings.itemimage, compound='left', command=partial(self.addmarkerFromSearch,idobscat,event, False))
                    menucats.add_separator()
                    menucats.add_command(label="Nova categoria", image=global_settings.addcat, compound='left', command=partial(self.addmarkerFromSearch,None, event, False))
                    menucats.add_separator()
                    menucats.add_command(label="Nova categoria (Utilizar Termo de Busca)", image=global_settings.addcat, compound='left', command=partial(self.addmarkersFromSearch,event))
                    self.menuexportsearchtobs.add_cascade(label='Enviar para:', menu=menucats, image=global_settings.catimage, compound='left')
                    if(self.treeviewSearches.tag_has('termosearch', iid)):
                        #self.menuexportsearchtobs.add_separator()
                        self.menuexportsearchtobs.add_command(label="Excluir Busca", image=global_settings.delcat, compound='right', command=self.exclude_search)   
                    self.menuexportsearchtobs.tk_popup(event.x_root, event.y_root)
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)
            finally:
                self.menuexportsearchtobs.grab_release()

        
    def exclude_search(self, event=None, lista=None):
        try:
            qtos = 0
            listadel1 = set()
            listadel2 = []
            
            for selecao in self.treeviewSearches.selection():
                if(self.treeviewSearches.parent(selecao)=='' and self.treeviewSearches.item(selecao, 'text') != ''):
                    try:
                        del global_settings.idtermopdfs[selecao]
                    except:
                        None
                    qtos += 1
                    if(selecao in global_settings.searchResultsDict):
                        resultsearch = global_settings.searchResultsDict[selecao]
                        if(resultsearch.idtermo in listadel1):
                            continue
                        listadel1.add(resultsearch.idtermo)
                        tipobusca = resultsearch.tipobusca 
                        listadel2.append(resultsearch.idtermo)
                        oldqueue = []
                        try:
                            self.termosearchVar.set("")                              
                        except Exception as ex:
                            utilities_general.printlogexception(ex=ex)
                    try:
                        del global_settings.listaTERMOS[(resultsearch.termo.upper(), tipobusca)]
                    except:
                        None
            self.treeviewSearches.delete(*self.treeviewSearches.selection())
            self.deleteSearchProcess(global_settings.result_queue, global_settings.pathdb, listadel2, global_settings.erros_queue, global_settings.queuesair)
            
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        
    def deleteSearchProcess(self, result_queue, pathdb, idtermos, erros_queue, queuesair):
        sqliteconn = None
        cursor = None
        notok = True
        window, progressbar = self.windSearchResults("Excluindo buscas!")
        progressbar['maximum'] = len(self.treeviewSearches.selection())
        progressbar['mode'] = 'indeterminate' 
        progressbar.update_idletasks()
        while(notok):
            sqliteconn = None        
            try:
                sqliteconn =  utilities_general.connectDB(str(pathdb), 5, maxrepeat=-1)
                if(sqliteconn==None):
                    utilities_general.popup_window("O banco de dados está ocupado.\n A operação não foi concluída, tente novamente em alguns segundos.", False)
                    return
                cursor = sqliteconn.cursor()
                cursor.custom_execute("PRAGMA journal_mode=WAL")
                cursor.custom_execute("PRAGMA foreign_keys = ON")
                termostr = ','.join(('?') for t in idtermos)
                cursor.custom_execute("DELETE FROM Anexo_Eletronico_SearchTerms WHERE id_termo IN ({})".format(termostr), idtermos)
                sqliteconn.commit()
                cursor.close()
                notok = False
                #return None
            except sqlite3.OperationalError as ex:
                utilities_general.printlogexception(ex=ex)
                time.sleep(2)
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)       
            finally:
                try:
                    window.destroy()
                except:
                    None
                try:
                    sqliteconn.close() 
                except:
                    None
            
    def querySql(self):
        self.w=classes_general.querySqlWindow(global_settings.root,'')
        self.querysql["state"] = "disabled" 
        global_settings.root.wait_window(self.w.top)
        self.querysql["state"] = "normal"
    
    def changecatpopupresult(self, event=None, operacao=None, window=None, valornovo=None, itens=None, novacatset=None):
        if(operacao=='ok'):
            sqliteconn =  utilities_general.connectDB(str(global_settings.pathdb), 5)
            try:
                if(sqliteconn==None):
                    utilities_general.popup_window("O banco de dados está ocupado.\n A operação não foi concluída, tente novamente em alguns segundos.", False)
                else:
                    for item in itens: 
                        iiditem = self.treeviewObs.item(item, 'values')[8]
                        iidantigo = self.treeviewObs.item(item, 'values')[10]
                        iidnovo = self.treeviewObs.item(novacatset[valornovo], 'values')[1]
                        if(str(iidantigo!=str(iidnovo))):
                            paginainit = self.treeviewObs.item(item, 'values')[2]
                            relpath = self.treeviewObs.item(item, 'values')[1]
                            p0y = float(self.treeviewObs.item(item, 'values')[4])
                            basepdf = utilities_general.get_normalized_path(os.path.join(global_settings.pathdb.parent, relpath))
                            basepdf = utilities_general.get_normalized_path(basepdf)
                            cursor = sqliteconn.cursor()
                            cursor.custom_execute("PRAGMA journal_mode=WAL")
                            updateinto2 = "UPDATE Anexo_Eletronico_Obsitens set id_obscat = ? WHERE id_obs = ?"
                            cursor.custom_execute(updateinto2, (iidnovo,iiditem,))                            
                            cursor.close()
                            tocpdf = global_settings.infoLaudo[basepdf].toc
                            try:
                                #tocnameprev = 
                                tocx = utilities_general.locateToc(int(paginainit), basepdf, p0y=p0y, tocpdf=tocpdf)
                                tocname = tocx[0]
                                
                                novoiidtoc = str(iidnovo)+basepdf+tocname
                                ident= ' '
                                if(not self.treeviewObs.exists(str(iidnovo)+basepdf)):
                                    indexrelobs = self.qualIndexTreeObs( None, None,str(iidnovo), ident+os.path.basename(basepdf))
                                    self.treeviewObs.insert(parent=str(iidnovo), iid=(str(iidnovo)+basepdf), text=ident+os.path.basename(basepdf), image=global_settings.imagereportb, index=indexrelobs, tag=('relobs'), values=(basepdf,))
                                    self.treeviewObs.tag_configure('relobs', background='#e3e1e1')
                                if(not self.treeviewObs.exists(str(iidnovo)+basepdf+tocname)):
                                    indextocobs = self.qualIndexTreeObs( None, None,str(iidnovo)+basepdf, ident+ident+tocname)
                                    novoiidtoc = self.treeviewObs.insert(parent=str(iidnovo)+basepdf, iid=(str(iidnovo)+basepdf+tocname), text=ident+ident+tocname, index=indextocobs,\
                                                                         tag=('tocobs'), values=(0,0,tocx[1],0, tocx[2], basepdf))                                
                                novoiidtocindex = self.qualIndexTreeObs( paginainit, p0y, (str(iidnovo)+basepdf+tocname))
                                parenteantigo = self.treeviewObs.parent(item)
                                self.treeviewObs.move(item, novoiidtoc, novoiidtocindex)
                                children = self.treeviewObs.get_children(parenteantigo)
                                if(len(children)==0 and self.treeviewObs.parent(parenteantigo)!=''):
                                    self.treeviewObs.delete(parenteantigo)
                            except Exception as ex:
                                utilities_general.printlogexception(ex=ex)
                                if(not self.treeviewObs.exists(str(iidnovo)+basepdf)):
                                    indexrelobs = self.qualIndexTreeObs( None, None,str(iidnovo), ident+os.path.basename(basepdf))
                                    self.treeviewObs.insert(parent=str(iidnovo), iid=(str(iidnovo)+basepdf), image=global_settings.imagereportb, text=ident+os.path.basename(basepdf), index=indexrelobs, tag=('relobs'), values=(basepdf,))
                                    self.treeviewObs.tag_configure('relobs', background='#e3e1e1')
                                novoiidindex = self.qualIndexTreeObs( paginainit, p0y, (str(iidnovo)+basepdf))
                                self.treeviewObs.move(item, (str(iidnovo)+basepdf), novoiidindex)
                                parenteantigo = self.treeviewObs.parent(item)
                                #self.treeviewObs.move(item, novoiidtoc, novoiidtocindex)
                                children = self.treeviewObs.get_children(parenteantigo)
                                if(len(children)==0 and self.treeviewObs.parent(parenteantigo)!=''):
                                    self.treeviewObs.delete(parenteantigo)
                    sqliteconn.commit()
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)
            finally:
                if(sqliteconn):
                    sqliteconn.close()                
            window.destroy()
        elif(operacao=='cancel'):
            window.destroy()    
            
    def orderObsBy(self, mode="histdec"):   
        if(mode=="histdec"):
            l = [(k, k) for k in self.treeviewObs.get_children('')]
            l.sort(key=lambda tid: int(tid[1]))
        elif(mode=="histcres"):
            l = [(k, k) for k in self.treeviewObs.get_children('')]
            l.sort(reverse=True, key=lambda tid: int(tid[1]))
        elif(mode=="azcres"):
            l = [(k, self.treeviewObs.item(k, 'text')) for k in self.treeviewObs.get_children('')]
            l.sort(key=lambda az: az[1])            
        elif(mode=="azdec"):
            l = [(k, self.treeviewObs.item(k, 'text')) for k in self.treeviewObs.get_children('')]
            l.sort(reverse=True, key=lambda az: az[1])
        for ln, _ in l:
            self.treeviewObs.move(ln, '', 'end')
    def orderpopupObs(self, event): 
        try:
            if(self.menuorderbyobs != None):
                None
        except:
            self.menuorderbyobs = tkinter.Menu(global_settings.root, tearoff=0)
            self.menuorderbyobs.add_command(label="Histórica Crescente", compound='left', image=global_settings.ordergenericup, command= lambda: self.orderObsBy('histdec'))
            self.menuorderbyobs.add_command(label="Histórica Decrescente", compound='left', image=global_settings.ordergenericdown, command= lambda: self.orderObsBy('histcres'))
            self.menuorderbyobs.add_command(label="Alfabética Crescente", compound='left', image=global_settings.orderazup, command= lambda: self.orderObsBy('azcres'))
            self.menuorderbyobs.add_command(label="Alfabética Decrescente", compound='left', image=global_settings.orderazdown, command= lambda: self.orderObsBy('azdec'))
        try:
            self.menuorderbyobs.tk_popup(event.x_root, event.y_root)         
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        finally:
            self.menuorderbyobs.grab_release()
    
    
    def orderSeachesBy(self, mode="histdec"):  
        if(mode=="histdec"):
            l = [(k, k) for k in self.treeviewSearches.get_children('')]
            l.sort(key=lambda tid: int(tid[1].replace("t", "")))
        elif(mode=="histcres"):
            l = [(k, k) for k in self.treeviewSearches.get_children('')]
            l.sort(reverse=True, key=lambda tid: int(tid[1].replace("t", "")))
        elif(mode=="azcres"):
            l = [(k, self.treeviewSearches.item(k, 'values')[0]) for k in self.treeviewSearches.get_children('')]
            l.sort(key=lambda az: az[1])            
        elif(mode=="azdec"):
            l = [(k, self.treeviewSearches.item(k, 'values')[0]) for k in self.treeviewSearches.get_children('')]
            l.sort(reverse=True, key=lambda az: az[1])
        elif(mode=="hitcres"):
            l = [(k, self.treeviewSearches.item(k, 'values')[2]) for k in self.treeviewSearches.get_children('')]
            l.sort(key=lambda hits: int(hits[1]))
        elif(mode=="hitdec"):
            l = [(k, self.treeviewSearches.item(k, 'values')[2]) for k in self.treeviewSearches.get_children('')]
            l.sort(reverse=True, key=lambda hits: int(hits[1]))
        for ln, _ in l:
            self.treeviewSearches.move(ln, '', 'end')
            
    def orderpopup(self, event): 
        try:
            if(self.menuorderby != None):
                None
        except:
            self.menuorderby = tkinter.Menu(global_settings.root, tearoff=0)
            self.orderazdownb = b'iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAABPUlEQVRIieWUP0oDQRjF18LCRi8QGN7vGxZs0uwRbALamBvYhBxATMDGMlews7QRvIOlq2ITb2CRRrDyD6w2uzAsmzAhY+WDV8zy8X7fPJbJskBmNgLK7K8UAwDKLicDmNkoNHAHvCUDhAIOgE/gKDlAkgMWks5jF4oG9Hq9HaCUdJtl2VZygKQr4MV7vxsdHguQdAq8O+f21wqPBQDfwJOky9DJAO3g5ICN9H8AeZ7TvEF5npMcAEzNbG5mc2CSHCDp0czOgMlalcYA6noq77055wRUZuaTAYApcB+cy5U1AbNlAEl9MztuzT8APy0vXwr4Ai7aAEl9YAGcNLPee6srGZhZUXsAVEv/JmDYQBpAEH5TFMV2MDvu2rauabzqFg3kGXjtCg9m92K+dUEOgY+6087wjQUMJV2nDP8FDNeWH0s/W/oAAAAASUVORK5CYII='
            self.orderazupb = b'iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAABUUlEQVRIieXUsUoDQRDG8UshgiCkSCxS7fxn2HRCvIdIrW9gkzyAZRCu01fwAWwEsbELVrYBtbASWwOCSZ8I2mThOJLLxWznwMDdcfv9uOF2kyRXwAlwZ2a7SexahM+AH+A+KpILfwE+gE/gNk3TneK7InK1rKuEZ6raA0YicrgKKQYDT8C8DJgBWZIkSQAWQQE5LVnrgamInJUBWbjOAwFR1eNl69rt9j7wClyvDC9WESipmqreAM+tVmsvOqCq58CX957K4VUB51wXmDnnuhuFVwWAKfCoqr18xwRGyzoasFX9H8DMmuGo8N43ogOq2gfGwLjyH7QJADwAF8AlMIwKeO8bwNzMjlQ1Bb5V9SAaoKp9EXkP98Bb7I02BCa5TTapPKZ1gJk1gbmIDMIRISKDxciaWwO58dQKX1VtTOsA51zdzDrF52bWcc7Vtwb+Ur//XZa7TzC8LAAAAABJRU5ErkJggg=='
            self.ordernumberdownb = b'iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAvUlEQVRIie3UIQ7CQBSE4RoOMv/4mj0GiuoaTM+BqOAK3ADTm9RwCQQnIJhi2mTTtAnQrSDwkpE73+6Kl2XR2K6ANltrXgGAdiopgc52bbvqUwNdaiBEZ0JyQFJpO9gOksrkwFSSARNn0n7R6oCkUxygWRUYkgxYNL8B/HfRFwLAcQ6QlNveLQUewGEMSMqBG7BfChQDMgBReRNC2IwBoHlrF0XIBbjOlfcv+2wXAVvgPtxwqnzxAIWkc8ryJ/IKpNezywMHAAAAAElFTkSuQmCC'
            self.ordernumberupb = b'iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAx0lEQVRIie3UsQ3CMBCFYTeIOe6/1DSeBGoamsyRIkswRETDGGlYgoIR0oQmRpblACGHhBROeoUl6z6fLJ1zUQE74FQUxdpZ19C8A3rgbIpEzS/AFbgBjfd+ld4VkWMu7zSvVLUEWhHZjCHDhE1oDDRA/wzogMo55wIwvDQghxRQVR/OqupfAVV0+QEERFW3s4C4UmDkQX0uZoCq+jgisjcFcqD1BGWS2voP2lzMgFm1DGDyLpoKTN5FnwBfXxV/4CeAejm76A67BqV0goeW0gAAAABJRU5ErkJggg=='
            self.ordergenericdownb = b'iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAmUlEQVRIie3TsQ3CMBCFYTfMYd1/FU0aj5Eq2YAmc1BmFBo2ScMSFMkCiAYaR7KQsJDOpImfdO377opzLomqDsDk/pUKVGAbAJhyo6qdCVDVITfe+6MJMKUCFdgrAIzfABFpsq//I/AEzp+AiDTADJysQAs8gHEFYvkCXEMIBxMQkT5ecgPucfMy5QmyXvIqXp4gvYhcSpa/AUdjWDfP8ZtHAAAAAElFTkSuQmCC'
            self.ordergenericupb = b'iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAsUlEQVRIie3VPQoCMRCG4WnEY4S8k2Ybm5zEvYHNXsO9i7XYeIxtvIE2CvY2NtokMFgIkgj+7AeBMMU8mRSJiAnQApsQwlRqJzW/AjdgWxUxzXfAETgD6xjjpGbzpap2wOCca4BTlUlS815EJAMiIt77WZpkUQr0eW+BjKjqvAiweQSqZwRGoPxA/wGoapcXsAIOtuaca4oAYDBrD1xs7enT8hFX9FtA+pjatwEv5yuBO7yaZfPtKzdJAAAAAElFTkSuQmCC'
            self.orderazdown = tkinter.PhotoImage(data=self.orderazdownb)
            self.orderazup = tkinter.PhotoImage(data=self.orderazupb)
            self.ordernumberdown = tkinter.PhotoImage(data=self.ordernumberdownb)
            self.ordernumberup = tkinter.PhotoImage(data=self.ordernumberupb)
            self.ordergenericup = tkinter.PhotoImage(data=self.ordergenericupb)
            self.ordergenericdown = tkinter.PhotoImage(data=self.ordergenericdownb)
            self.menuorderby.add_command(label="Histórica Crescente", compound='left', image=global_settings.ordergenericup, command= lambda: self.orderSeachesBy('histdec'))
            self.menuorderby.add_command(label="Histórica Decrescente", compound='left', image=global_settings.ordergenericdown, command= lambda: self.orderSeachesBy('histcres'))
            self.menuorderby.add_command(label="Alfabética Crescente", compound='left', image=global_settings.orderazup, command= lambda: self.orderSeachesBy('azcres'))
            self.menuorderby.add_command(label="Alfabética Decrescente", compound='left', image=global_settings.orderazdown, command= lambda: self.orderSeachesBy('azdec'))
            self.menuorderby.add_command(label="Qtde. Hits. Crescente", compound='left', image=global_settings.ordernumberup, command= lambda: self.orderSeachesBy('hitcres'))
            self.menuorderby.add_command(label="Qtde. Hits. Decrescente", compound='left', image=global_settings.ordernumberdown, command= lambda: self.orderSeachesBy('hitdec'))
        
        try:
            self.menuorderby.tk_popup(event.x_root, event.y_root)         
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        finally:
            self.menuorderby.grab_release()
     
    def add_edit_category(self, operacao, valor=''):
        w = None
        sqliteconn = None
        try:
            if(operacao=='edit'):
                item = self.treeviewObs.selection()[0]
                valor = self.treeviewObs.item(item, 'text')
                idobscat = self.treeviewObs.item(item, 'values')[1]
                w=classes_general.popupWindow(global_settings.root,valor)            
                self.menuaddcat["state"] = "disabled" 
                global_settings.root.wait_window(w.top)
                if(w.value!=None and w.value.strip()!=''):
                    newcat = (w.value.upper())
                    sqliteconn = utilities_general.connectDB(str(global_settings.pathdb), 5)
                    cursor = sqliteconn.cursor()
                    self.menuaddcat["state"] = "normal"
                    updateinto2 = "UPDATE Anexo_Eletronico_Obscat set obscat = ? WHERE id_obscat = ?"
                    cursor.custom_execute(updateinto2, (newcat,int(idobscat),))
                    self.treeviewObs.item(self.treeviewObs.selection()[0], text=newcat)
                    self.treeviewObs.item(item)['values'] = (1, idobscat, newcat)
                    sqliteconn.commit() 
            elif(operacao=='add'):
                w=classes_general.popupWindow(global_settings.root,valor)            
                self.menuaddcat["state"] = "disabled" 
                global_settings.root.wait_window(w.top)
                if(w.value!=None and w.value.strip()!=''):
                    sqliteconn = utilities_general.connectDB(str(global_settings.pathdb), 5)
                    cursor = sqliteconn.cursor()
                    newcat = (w.value.upper())
                    insertinto =  "INSERT INTO Anexo_Eletronico_Obscat (obscat, fixo, ordem) values (?,?,0)"
                    fixo = 0
                    if(True):
                        fixo = 1
                    cursor.custom_execute(insertinto, (newcat.upper(), fixo,))
                    idnovo = cursor.lastrowid
                    self.treeviewObs.insert(parent='', index='end', iid=idnovo, text=newcat.upper(), values=(str(fixo), idnovo, newcat,), \
                                            image=global_settings.catimage, tag='obscat')
                    self.treeviewObs.tag_configure('obscat', background='#a1a1a1', font=global_settings.Font_tuple_ArialBoldUnderline_12)
                    sqliteconn.commit()
                    return (idnovo, newcat)
        except Exception as ex:
            utilities_general.popup_window("Erro na inserção/edição de categoria.", False)
            utilities_general.printlogexception(ex=ex)
        finally:
            self.menuaddcat["state"] = "normal" 
            if(sqliteconn):
                sqliteconn.close()
    
    def exclude_itens(self, items, database=True):
        sqliteconn = None
        if(database):
            sqliteconn =  utilities_general.connectDB(str(global_settings.pathdb), 5)
        try:
            if(sqliteconn==None and database):
                utilities_general.popup_window("O banco de dados está ocupado.\n A operação não foi concluída, tente novamente em alguns segundos.", False)
            else:
                for item in items:
                    valores = self.treeviewObs.item(item, 'values')
                    if(database):
                        cursor = sqliteconn.cursor()
                        cursor.custom_execute("PRAGMA foreign_keys = ON")
                        deletefrom = "DELETE FROM Anexo_Eletronico_Obsitens WHERE id_obs = ?"
                        
                        cursor.custom_execute("PRAGMA journal_mode=WAL")
                        cursor.custom_execute(deletefrom, (valores[8],))
                        cursor.close()
                    parenteantigo = self.treeviewObs.parent(item)    
                    observation = self.allobsbyitem['obsitem'+str(valores[8])]
                    pathpdf = observation.pathpdf
                    del self.allobsbyitem['obsitem'+str(valores[8])]
                    self.allobs[pathpdf].remove(observation)
                    self.docInnerCanvas.delete('obsitem'+str(valores[8]))
                    self.docInnerCanvas.delete('enhanceobs'+str(valores[8]))
                    self.docInnerCanvas.delete('note'+str(valores[8]))
                    for annot in observation.annotations:
                        annotation = observation.annotations[annot]
                        self.docInnerCanvas.delete(str(valores[8])+'enhanceannot'+str(annotation.idannot))
                    self.treeviewObs.delete(item)
                    children = self.treeviewObs.get_children(parenteantigo)        
                    while(parenteantigo!=''):
                        children = self.treeviewObs.get_children(parenteantigo)
                        temp = self.treeviewObs.parent(parenteantigo)
                        if(len(children)>0 or temp==''):
                            break
                        self.treeviewObs.delete(parenteantigo)
                        parenteantigo = temp
                if(database):
                    sqliteconn.commit()
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        finally:
            if(sqliteconn and database):
                sqliteconn.close()
                
    def copy_obs_rtf_multiple_tables(self, iid, withnotes=False):
        docespecial = None

        try:              
            pathdocespecial = None                
            pagatual = None
            textonatabela = ""
            
            cont = 1  
            lista = []
            self.descendants_obsitem(iid, lista)
            for item, pdf in lista:
                pinit = None
                pfim = None
                pinit2 = math.inf
                pfim2 = -1
                observation = self.allobsbyitem[item]
                estaselecao = ""
                if(observation.conteudo!=''):
                    observacao_p1 = "Observação: "
                    estaselecao += '{\\rtlch\\ltrch\\loch\\fs20\\li72\\f2\\b\\ul{'+observacao_p1.encode('rtfunicode').decode('ascii') + '}}'
                    estaselecao += '{\\rtlch\\ltrch\\loch\\fs20\\li72\\f2\\i{'+observation.conteudo.encode('rtfunicode').decode('ascii') + '}\\line}'
                valoresPecial = self.treeviewObs.item(item, 'values')
                pagi = int(valoresPecial[2].strip())+1
                estaselecao+=f"\\f2(...)\\line"
                
                
                if(pathdocespecial!=observation.pathpdf):
                    pathdocespecial1 = utilities_general.get_normalized_path(observation.pathpdf)
                    pagatual = None
                    if(docespecial!=None):
                        docespecial.close()
                    docespecial = fitz.open(pathdocespecial1)
                tiposelecao = observation.tipo          
                pinit = observation.paginainit
                pfim = observation.paginafim
                pinit2 = min(pinit, pinit2)
                pfim2 = max(pfim, pfim2)
                p0xinit = observation.p0x
                p0yinit = observation.p0y+2
                p1xinit = observation.p1x
                p1yinit = observation.p1y-2 
                if(tiposelecao=='area'):
                    textonatabela += self.ObstoRTf(item, docespecial, pathdocespecial1, tiposelecao, pinit, \
                                                   pfim, p0xinit, p0yinit, p1xinit, p1yinit, pmin=pinit2, pmax = pfim2, withnote=withnotes)[0]
                elif(tiposelecao=='texto'):
                    textonatabela += self.ObstoRTf(item, docespecial, pathdocespecial1, tiposelecao, pinit, pfim, p0xinit,\
                                                   p0yinit, p1xinit, p1yinit, estaselecao=estaselecao, pmin=pinit2, pmax = pfim2, withnote=withnotes)[0]  
            textofinal = ("{{\\rtf1\\ansi\\deff0{{\\fonttbl{{\\f0\\froman\\fprq2\\fcharset0 Times New Roman;}}{{\\f1\\froman\\fprq2\\fcharset2 Symbol;}}"+
"{{\\f2\\fswiss\\fprq2\\fcharset0 Arial;}}}}{{\\colortbl;\\red240\\green240\\blue240;\\red221\\green221\\blue221;\\red255\\green255\\blue255;\\red190\\green0\\blue0;}} {}}}")\
            .format(textonatabela)
            conteudo = bytearray(textofinal, 'utf8')
            utilities_general.copy_to_clipboard("rtf", conteudo)  
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        finally:
            try:
                docespecial.close()
            except Exception as ex:
                None
    def copy_obs_rtf_single_table(self, withnotes=False):  
        pathdocespecial = None
        docespecial = None
        pagatual = None
        textonatabela = ""
        pinit = None
        pfim = None
        pinit2 = math.inf
        pfim2 = -1
        textoselecao = ""
        for iid in self.treeviewObs.selection():
            observation = self.allobsbyitem[iid]
            if(observation.conteudo!=''):
                observacao_p1 = "Observação: "
                textoselecao += '{\\rtlch\\ltrch\\loch\\fs20\\li72\\f2\\b\\ul{'+observacao_p1.encode('rtfunicode').decode('ascii') + '}}'
                textoselecao += '{\\rtlch\\ltrch\\loch\\fs20\\li72\\f2\\i{'+observation.conteudo.encode('rtfunicode').decode('ascii') + '}}'
            
            pdf = self.treeviewObs.parent(iid)
            child2 = iid
            while(self.treeviewObs.parent(pdf)!=''):
                pdf = self.treeviewObs.parent(pdf)
            #valoresPecial = self.treeviewObs.item(child2, 'values')
            #pagi = int(valoresPecial[2].strip())+1
            if(pathdocespecial!=observation.pathpdf):
                pathdocespecial1 = utilities_general.get_normalized_path(observation.pathpdf)
                pagatual = None
                if(docespecial!=None):
                    docespecial.close()
                docespecial = fitz.open(pathdocespecial1)
            tiposelecao = observation.tipo          
            pinit = observation.paginainit
            pfim = observation.paginafim
            pinit2 = min(pinit, pinit2)
            pfim2 = max(pfim, pfim2)
            p0xinit = observation.p0x
            p0yinit = observation.p0y+2
            p1xinit = observation.p1x
            p1yinit = observation.p1y-2   
            if(tiposelecao=='texto'):
                if(textoselecao.endswith("(...)\\line")):
                    textoselecao += f"(Fls. {observation.paginainit+1})\\line"
                else:
                    textoselecao += f"\\line\\f2(...) (Fls. {observation.paginainit+1})\\line"
            else:
                if(textoselecao.endswith("(...)\\line")):
                    textoselecao += f"\\f2(Fls. {observation.paginainit+1})\\line"
                else:
                    textoselecao += f"\\line\\f2(Fls. {observation.paginainit+1})\\line"
            
            textonatabela, textoselecao = self.ObstoRTf(iid, docespecial, pathdocespecial1, \
                                                        tiposelecao, pinit, pfim, p0xinit, p0yinit, \
                                                            p1xinit, p1yinit, estaselecao=textoselecao, pmin=pinit2, pmax = pfim2, withnote=withnotes)
        textofinal = ("{{\\rtf1\\ansi\\deff0{{\\fonttbl{{\\f0\\froman\\fprq2\\fcharset0 Times New Roman;}}{{\\f1\\froman\\fprq2\\fcharset2 Symbol;}}"+
           "{{\\f2\\fswiss\\fprq2\\fcharset0 Arial;}}}}{{\\colortbl;\\red240\\green240\\blue240;\\red221\\green221\\blue221;\\red255\\green255\\blue255;\\red190\\green0\\blue0;}} {}}}").format(textonatabela)
        conteudo = bytearray(textofinal, 'utf8')
        utilities_general.copy_to_clipboard("rtf", conteudo) 
     
    def export_images_table_rtf(self, withnotes=False):    
        utilities_export.Export_Images_To_Table(self, withnotes)
        
    def copy_obs_rtf(self, withnotes=False):  
        selected_obs = self.treeviewObs.selection()
        if(len(selected_obs)>1):
            self.copy_obs_rtf_single_table(withnotes)
        elif(len(selected_obs)==1):
            self.copy_obs_rtf_multiple_tables(selected_obs[0], withnotes)
    
    def addcatpopup(self, event=None, operacao=None, valor='') :
        try:
            iid = self.treeviewObs.selection()[0]
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
            return
        if(operacao=='excludeitems'):
            self.exclude_itens(self.treeviewObs.selection())
        elif(operacao=='excludeitem'):
            self.exclude_itens([self.treeviewObs.selection()[0]])
        elif(operacao=='exclude'):
            sqliteconn =  utilities_general.connectDB(str(global_settings.pathdb), 5)
            try:
                if(sqliteconn==None):
                    utilities_general.popup_window("O banco de dados está ocupado.\n A operação não foi concluída, tente novamente em alguns segundos.", False)
                    return                
                else:
                    item = self.treeviewObs.selection()[0]
                    idobs = self.treeviewObs.item(item, 'values')[1]
                    cursor = sqliteconn.cursor()
                    cursor.custom_execute("PRAGMA journal_mode=WAL")
                    cursor.custom_execute("PRAGMA foreign_keys = ON")
                    deletefrom2 = "DELETE FROM Anexo_Eletronico_Obscat WHERE id_obscat = ?"
                    cursor.custom_execute(deletefrom2, (idobs,))
                    leaves = []
                    self.colectLeaves(self.treeviewObs, item, leaves)
                    self.exclude_itens(leaves, database=False)
                    self.treeviewObs.delete(self.treeviewObs.selection()[0])
                    sqliteconn.commit()
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)
            finally:
                if(sqliteconn):
                    sqliteconn.close()
        elif(operacao=='changecats'):
            window = tkinter.Toplevel()
            window.rowconfigure((0,1), weight=1)
            window.columnconfigure((0,1), weight=1)
            label = tkinter.Label(window, font=global_settings.Font_tuple_Arial_10, text='Nova Categoria')
            label.grid(row=0, column=0, padx=50, pady=20, sticky='ns')            
            n = tkinter.StringVar() 
            novacat = ttk.Combobox(window, width = 27,  
                            textvariable = n, exportselection=0, state="readonly")
            values = []
            filhos = self.treeviewObs.get_children('')
            novacatset = {}
            for filho in filhos:
                texto = self.treeviewObs.item(filho, 'text')  
                novacatset[texto] = filho
                values.append(texto)
            novacat['values'] = (values)
            novacat.current(0)
            novacat.grid(row=0, column=1, padx=50, pady=20, sticky='ns')
            item = self.treeviewObs.focus()
            button_ok = tkinter.Button(window, font=global_settings.Font_tuple_Arial_10, text="OK", command= lambda : self.changecatpopupresult(event=None, operacao='ok', \
                                                                          window=window, valornovo=n.get(), itens=self.treeviewObs.selection(), novacatset=novacatset))
            button_ok.grid(row=1, column=0, padx=50, pady=20, sticky='ns')
            button_close = tkinter.Button(window, font=global_settings.Font_tuple_Arial_10, text="Cancelar", command= lambda : self.changecatpopupresult(event=None, operacao='cancel', \
                                                                                                           window=window, valornovo=None, itens=None, novacatset=None))
            button_close.grid(row=1, column=1, padx=50, pady=20, sticky='ns')
        elif(operacao=='changecat'):
            window = tkinter.Toplevel()
            window.rowconfigure((0,1), weight=1)
            window.columnconfigure((0,1), weight=1)
            label = tkinter.Label(window, font=global_settings.Font_tuple_Arial_10, text='Nova Categoria')
            label.grid(row=0, column=0, padx=50, pady=20, sticky='ns')            
            n = tkinter.StringVar() 
            novacat = ttk.Combobox(window, width = 27,  
                            textvariable = n, exportselection=0, state="readonly")
            values = []
            filhos = self.treeviewObs.get_children('')
            novacatset = {}
            for filho in filhos:
                texto = self.treeviewObs.item(filho, 'text')  
                novacatset[texto] = filho
                values.append(texto)
            novacat['values'] = (values)
            novacat.current(0)
            novacat.grid(row=0, column=1, padx=50, pady=20, sticky='ns')
            item = self.treeviewObs.focus()
            button_ok = tkinter.Button(window, font=global_settings.Font_tuple_Arial_10, text="OK", command= lambda : self.changecatpopupresult(event=None, operacao='ok', \
                                                                                              window=window, valornovo=n.get(), itens=[item], novacatset=novacatset))
            button_ok.grid(row=1, column=0, padx=50, pady=20, sticky='ns')
            button_close = tkinter.Button(window, font=global_settings.Font_tuple_Arial_10, text="Cancelar", command= lambda : self.changecatpopupresult(event=None, operacao='cancel', \
                                                                                                               window=window, valornovo=None, itens=None, novacatset=None))
            button_close.grid(row=1, column=1, padx=50, pady=20, sticky='ns')
        elif(operacao=='copiarclip'):
            utilities_export.Export_to_Clipboard(self, iid)       
        elif(operacao=='copiarcsv'):
            utilities_export.Export_to_CSV(self, iid)                        
        elif(operacao=="validarobs"):
            item = self.treeviewObs.selection()[0]
            valores = self.treeviewObs.item(item, 'values')
            iiditem = valores[8]
            sqliteconn =  utilities_general.connectDB(str(global_settings.pathdb), 5)
            try:
                if(sqliteconn==None):
                    utilities_general.popup_window("O banco de dados está ocupado.\n A operação não foi concluída, tente novamente em alguns segundos.", False)
                    return                
                else:
                    cursor = sqliteconn.cursor()
                    cursor.custom_execute("PRAGMA journal_mode=WAL")
                    cursor.custom_execute("UPDATE Anexo_Eletronico_Obsitens SET status = 'ok' WHERE id_obs = ?", (iiditem,))
                    self.treeviewObs.tag_configure('alterado'+str(iiditem), background='#ffffff')
                    self.treeviewObs.item(item, tags=('obsitem', 'ok'+str(iiditem),))                    
                    sqliteconn.commit()                
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)
            finally:
                if(sqliteconn):
                    sqliteconn.close()
        

            
        
            
    
    def add_obscat(self, catname):
        newcat = catname.upper()                
        check_previous_search =  "SELECT C.id_obscat, C.obscat FROM Anexo_Eletronico_Obscat C WHERE C.obscat = ?"
        sqliteconn = utilities_general.connectDB(str(global_settings.pathdb), 5)
        try:
            if(sqliteconn==None):
                utilities_general.popup_window("O banco de dados está ocupado.\n A operação não foi concluída, tente novamente em alguns segundos.", False)
                return (None, None)            
            else:
                cursor = sqliteconn.cursor()
                cursor.custom_execute("PRAGMA journal_mode=WAL")
                cursor.custom_execute("PRAGMA foreign_keys = ON")
                cursor.custom_execute(check_previous_search, (newcat.upper(),))
                termos = cursor.fetchone()
                if(termos==None):
                    insertinto =  "INSERT INTO Anexo_Eletronico_Obscat (obscat, fixo, ordem) values (?,?,0)"
                    fixo = 0
                    if(True):
                        fixo = 1
                    cursor.custom_execute(insertinto, (newcat.upper(), fixo,))
                    idnovo = cursor.lastrowid
                    self.treeviewObs.insert(parent='', index='end', iid=idnovo, text=newcat.upper(), values=(str(fixo), idnovo, newcat.upper(),), image=global_settings.catimage, tag='obscat')
                    self.treeviewObs.tag_configure('obscat', background='#a1a1a1', font=global_settings.Font_tuple_ArialBoldUnderline_12)
                    sqliteconn.commit()
                    return (idnovo, newcat.upper())
                else:                    
                    return (termos[0], newcat.upper())                
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
            return (None, None)
        finally:
            if(sqliteconn):
                sqliteconn.close()                    
    
    def descendants_obsitem(self, item, lista):       
        if(self.treeviewObs.tag_has('obsitem', item)):
            if(item not in lista):
                pdf = os.path.basename((self.treeviewObs.item(item, 'values'))[1])
                lista.append((item, pdf))            
        else:
            children = self.treeviewObs.get_children(item)
            for child in children:
                self.descendants_obsitem(child, lista)        

    def get_thumbnail(self, filepath):
        size = 340, 340
        try:
            filename, extension = os.path.splitext(filepath)
            if(extension.lower() in global_settings.listavidformats):
                #ffmpeg -ss 01:23:45 -i input -frames:v 1 -q:v 2 output.jpg
                executavel = os.path.join(utilities_general.get_application_path(), "ffmpeg")
                comando = f"\"{executavel}\" -y -ss 1 -i \"{filepath}\" -frames:v 1 -q:v 2 teste.png"  
                #popen = subprocess.Popen(comando, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                popen = subprocess.run(comando,universal_newlines=True, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True)
                #return_code = popen.wait()
                
                if(not os.path.exists("teste.png")):
                    comando = f"\"{executavel}\" -y -ss 0 -i \"{filepath}\" -frames:v 1 -q:v 2 teste.png"                                              
                    #popen = subprocess.Popen(comando, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    popen = subprocess.run(comando,universal_newlines=True, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True)
                    
                    lines = popen.stderr.split("\n")
                    if(len(lines)>0):
                        raise classes_general.FFMPEGException(lines)
                    
                    #return_code = popen.wait()
                if(os.path.exists("teste.png")): 
                    Image.LOAD_TRUNCATED_IMAGES = True
                    with Image.open("teste.png") as imx:
                        imx.thumbnail(size, Image.ANTIALIAS)
                        
                        imx.save('teste.png','PNG')
                        width, height = imx.size
                    with open('teste.png', 'rb') as f:
                        content = f.read()
                        pngtohex = binascii.hexlify(content).decode('utf8')
                        #os.remove('teste.png')
                        
                        return (340, 340, pngtohex)
                    
            else:
                Image.LOAD_TRUNCATED_IMAGES = True
                with Image.open(filepath) as imx:
                    
                    imx.thumbnail(size, Image.ANTIALIAS)
                    imx.save('teste.png','PNG')
                    width, height = imx.size
                with open('teste.png', 'rb') as f:
                    content = f.read()
                    pngtohex = binascii.hexlify(content).decode('utf8')
                    #os.remove('teste.png')
                    
                    return (340, 340, pngtohex)
        except classes_general.FFMPEGException as ex:
            utilities_general.printlogexception(ex=ex)
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        finally:
            try:
                os.remove('teste.png')
            except:
                None 

    def rtf_encode_char(unichar):
        code = ord(unichar)        
        return '\\u' + str(code if code <= 32767 else code-65536) + '?'
    
    def ObstoRTf(self, idobsitem, docespecial, pathdocespecial, tiposelecao, pinit, pfim, p0xinit, p0yinit, p1xinit, p1yinit, estaselecao="", pmin=None, pmax=None, withnote=False):
         textonatabela = ""
         pathdocespecial = utilities_general.get_normalized_path(pathdocespecial)
         try:
             observation = self.allobsbyitem[idobsitem]
             annotations = observation.annotations
             #print(annotations)
             margemsup = (global_settings.infoLaudo[pathdocespecial].mt/25.4)*72
             margeminf = global_settings.infoLaudo[pathdocespecial].pixorgh-((global_settings.infoLaudo[pathdocespecial].mb/25.4)*72)
             margemesq = (global_settings.infoLaudo[pathdocespecial].me/25.4)*72
             margemdir = global_settings.infoLaudo[pathdocespecial].pixorgw-((global_settings.infoLaudo[pathdocespecial].md/25.4)*72)
             if(tiposelecao=='area'):                 
                                                   
                 images = []
                 for pagina in range(pinit, pfim+1):
                     p0x = max(p0xinit, margemesq)
                     p1x = min(p1xinit, margemdir)
                     p0y = None
                     p1y = None
                     if(pinit==pfim):
                         p0y = p0yinit
                         p1y = p1yinit
                     elif(pagina==pinit):
                         p0y = p0yinit
                         p1y = margeminf
                     elif(pagina==pfim):
                         p0y = margemsup
                         p1y = p1yinit
                     else:
                         p0y = margemsup
                         p1y = margeminf
                     loadedPage = docespecial[pagina]
                     box = fitz.Rect(p0x, p0y, p1x, p1y)
                     matriz = fitz.Matrix(2.5, 2.5)
                     pix = loadedPage.get_pixmap(alpha=False, matrix=matriz, clip=box, dpi=1200)
                     mode = "RGBA" if pix.alpha else "RGB"
                     imgdata = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
                     pix = None
                     images.append(imgdata)
                 if(len(images) > 0):
                     imagem = utilities_general.concatVertical(images)
                    
                     if platform.system() == 'Windows' or  platform.system() == 'Linux':    # Windows
                        output = io.BytesIO()
                        imagem.save('teste.png','PNG')
                        with open('teste.png', 'rb') as f:
                            content = f.read()
                        os.remove('teste.png')
                        pngtohex = binascii.hexlify(content).decode('utf8')                         
                        width, height = imagem.size
                        ratio = width / height
                        #preencher = "\\picwgoal9070"
                        preencher = ""
                        if(height>width):
                            preencher = f"\\picwgoal{int(4070/ratio)}\\pichgoal4070"
                        else:
                            preencher = f"\\picwgoal9070\\pichgoal{int(9070/ratio)}"
                        estaselecao  += "{{\\rtlch \\ltrch\\loch\\qc\\pict\\pngblip\\picscalex100\\picscaley100\\piccropl0\\piccropr0\\piccropt0\\piccropb0{}\\pngblip\n{}}}\\line"\
                            .format(preencher, pngtohex)
                        docname = os.path.basename(pathdocespecial)
                        docpagina = ""
                        if(pmin!=None and pmax!=None and pmin!=pmax):
                            docpagina = "{{\\fs22\\f2{{ Relat}}\\\'F3rio \\\'22{}\\\'22 -- Fls. {} a {}}}".format(docname, pmin+1, pmax+1)
                        else:
                            docpagina = "{{\\fs22\\f2{{ Relat}}\\\'F3rio \\\'22{}\\\'22 -- Fls. {}}}".format(docname, pagina+1)   
                        
                        #docpagina = "{{\\fs22\\f2{{ Relat}}\\\'F3rio \\\'22{}\\\'22 -- Fls. {}}}".format(docname, pagi)
                        textonatabela += ("\\par\\trowd\\clbrdrb\\brdrs\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrr\\brdrs\\trautofit1\\intbl\\clftsWidth3\\clwWidth9070\\cellx9070{{\\cbpat2\\qc\\loch\\i\\b\\fs22\\f2 TABELA }}{{\\hich\\af2\\af2\\loch\\f2\\f2\\loch\\field{{\\*\\fldinst{{SEQ Tabela\\\\* ARABIC }}}}{{\\fldrslt{{ }}}}}}"+\
                        "{{\\f2\\qc:\\i\\f2{}}}\\f2\\cell\\row"+\
                            # "\\trowd\\clftsWidth1\\clbrdrb\\brdrs\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrr\\brdrs\\cellx1\\intbl{{\\cbpat2\\qc\\loch\\b{{TABELA}}}}\\cell\\row"+\
                            "\\par\\trowd\\clbrdrb\\brdrs\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrr\\brdrs\\trautofit1\\intbl\\clftsWidth3\\clwWidth9070\\cellx9070 {}\\cell\\row\\pard\\line").format(docpagina, estaselecao) + "\n"
                        return (textonatabela, estaselecao)
             elif(tiposelecao=='texto'):
                 addreticencias = False
                 if(estaselecao==""):
                     addreticencias = True
                 reg12 = "\[[0-9]{2}\/[0-9]{2}\/[0-9]{4}\s[0-9]{2}:[0-9]{2}:[0-9\?]{2}\]\s<.*>:\s"
                 reg2 = "\[[0-9]{2}\/[0-9]{2}\/[0-9]{4}\s[0-9]{2}:[0-9]{2}:[0-9\?]{2}\]\s<este\saparelho>:\s"
                 reganexo = r"Anexo: .+"
                 pagatual = None
                 links_img = []
                 links = []
                 dictx = None
                 for pagina in range(pinit, pfim+1):
                     imagens = set()
                     width, height = 340, 340
                     size = 340, 340
                     if(pagina!=pagatual):
                         pagatual = docespecial[pagina]
                         dictx = docespecial[pagina].get_text("rawdict", flags=2+4+64)
                         links = docespecial[pagina].get_links()
                     realce = False                     
                     p0x = max(p0xinit, margemesq)
                     p1x = min(p1xinit, margemdir)
                     p0y = None
                     p1y = None
                     if(pinit==pfim):
                         p0y = p0yinit
                         p1y = p1yinit
                     elif(pagina==pinit):
                         p0y = p0yinit
                         p1y = margeminf
                     elif(pagina==pfim):
                         p0y = margemsup
                         p1y = p1yinit
                     else:
                         p0y = margemsup
                         p1y = margeminf
                     
                     for block in dictx['blocks']: 
                        bboxb = block['bbox']
                        x0b = bboxb[0] 
                        y0b = float(bboxb[1])
                        x1b = bboxb[2]
                        y1b = float(bboxb[3])
                        if(y1b < p0y-1):
                            continue
                        elif(y0b > p1y+1):
                            break
                        ymedio = (y0b+y1b)/2.0
                        if(block['type']==0):
                            for line in block['lines']:
                                linha = ""
                                linhaorg = ""
                                bboxl = line['bbox']
                                x0l = bboxl[0] 
                                y0l = bboxl[1]
                                x1l = bboxl[2]
                                y1l = bboxl[3]
                                
                                for span in line['spans']:
                                   if(p0y > y0l and p1y < y1l):
                                      primeiroemoji = True
                                      x0 = min(p0x, p1x)
                                      x1 = max(p0x, p1x) 
                                      for char in span['chars']:
                                          quad = char['bbox']
                                          c = char['c']
                                          if(ord(c)>=55296 and primeiroemoji):
                                              primeiroemoji = False
                                              c = ' ' + c
                                          if(float(quad[2]) <= x1 and (float(quad[0])+float(quad[2]))/2 >= x0):       
                                              linhaorg += c
                                   elif(p0y <= y0l and p1y >= y1l):
                                       primeiroemoji = True
                                       for char in span['chars']:                                             
                                           quad = char['bbox']
                                           c = char['c']                                                 
                                           if(ord(c)>=55296 and primeiroemoji):
                                               primeiroemoji = False
                                               c = ' ' + c
                                           #linha += c1
                                           linhaorg += c
                                   elif(p0y <= y1l and p1y > y1l):  
                                       primeiroemoji = True
                                       for char in span['chars']:
                                           quad = char['bbox']
                                           if((float(quad[0])+float(quad[2]))/2 >= p0x):  
                                               c = char['c'] 
                                               if(ord(c)>=55296 and primeiroemoji):
                                                   primeiroemoji = False
                                                   c = ' ' + c
                                               linhaorg += c
                                   elif(p1y >= y0l and p0y < y0l):
                                       primeiroemoji = True
                                       for char in span['chars']:                                                 
                                           quad = char['bbox']  
                                           if((float(quad[0])+float(quad[2]))/2 <= p1x):  
                                               c = char['c'] 
                                               
                                               if(ord(c)>=55296 and primeiroemoji):
                                                   primeiroemoji = False
                                                   c = ' ' + c
                                               linhaorg += c
                                   
                                matchorigem = re.search(reg2, linhaorg)
                                matchdestino2= re.search(reg12, linhaorg)
                                if(linhaorg.strip()!=''):
                                    if(matchorigem!=None):
                                        realce = True
                                        linha1 = '{\\rtlch \\ltrch\\loch\\fs20\\li72\\f2\\highlight1{'+ linhaorg[:matchorigem.start()].encode('rtfunicode').decode('ascii')+ '}}'
                                        linha2 = '{\\rtlch \\ltrch\\loch\\fs20\\li72\\f2\\highlight1{'+linhaorg[matchorigem.start():matchorigem.end()].encode('rtfunicode').decode('ascii')+ '}}'
                                        linha3 = '{\\rtlch \\ltrch\\loch\\fs20\\li72\\f2\\highlight1{'+linhaorg[matchorigem.end():].encode('rtfunicode').decode('ascii') + '}\\line}'
                                        estaselecao += linha1+linha2+linha3 
                                    else:
                                        if(matchdestino2!=None):
                                            realce = False
                                            linha1 = '{\\rtlch \\ltrch\\loch\\fs20\\li72\\f2{'+ linhaorg[:matchdestino2.start()].encode('rtfunicode').decode('ascii')+ '}}'
                                            linha2 = '{\\rtlch \\ltrch\\loch\\fs20\\li72\\f2{'+linhaorg[matchdestino2.start():matchdestino2.end()].encode('rtfunicode').decode('ascii')+ '}}'
                                            linha3 = '{\\rtlch \\ltrch\\loch\\fs20\\li72\\f2{'+linhaorg[matchdestino2.end():].encode('rtfunicode').decode('ascii') + '}\\line}'
                                            estaselecao += linha1+linha2+linha3
                                        elif(realce):
                                            linha3 = '{\\rtlch \\ltrch\\loch\\fs20\\li72\\f2\\highlight1{'+linhaorg.encode('rtfunicode').decode('ascii') + '}\\line}'
                                            estaselecao += linha3 
                                        else:
                                            realce = False 
                                            linha3 = '{\\rtlch \\ltrch\\loch\\fs20\\li72\\f2{'+linhaorg.encode('rtfunicode').decode('ascii') + '}\\line}'
                                            estaselecao += linha3 
                                
                                if(withnote):
                                    for annot_id in annotations:
                                        annot_obj = annotations[annot_id]
                                        pymedio = (annot_obj.p1y + annot_obj.p0y) / 2.0
                                        if(pymedio >= y0l and pymedio <= y1l):  
                                            if(annot_obj.conteudo!=''):
                                                observacao_p0 = "          "
                                                observacao_p1 = "↳ Observação: "
                                                estaselecao +=\
                                                    '{\\rtlch\\ltrch\\loch\\fs20\\li72\\f2{'+observacao_p0.encode('rtfunicode').decode('ascii') + '}}'
                                                estaselecao +=\
                                                    '{\\rtlch\\ltrch\\loch\\fs20\\li72\\f2\\b\\ul{'+observacao_p1.encode('rtfunicode').decode('ascii') + '}}'
                                                estaselecao += \
                                                    '{\\rtlch\\ltrch\\loch\\fs20\\li72\\f2\\cf4\\i{'+annot_obj.conteudo.encode('rtfunicode').decode('ascii') + '}\\line}'
                                            break
                        elif(block['type']==1):
                            
                            bboxb = block['bbox']
                            x0b = bboxb[0] 
                            y0b = float(bboxb[1])
                            x1b = bboxb[2]
                            y1b = float(bboxb[3])
                            mediay = (p0y + p1y) / 2.0 
                            for annot_id in annotations:
                                annot_obj = annotations[annot_id]
                                pymedio = (annot_obj.p1y + annot_obj.p0y) / 2.0
                                if(pymedio >= bboxb[1] and pymedio <= bboxb[3]):
                                    if("#" in annot_obj.link):
                                        None
                                    else:
                                        pngtohex = ""
                                        content = ""
                                        width, height = None, None
                                        filepath = str(Path(os.path.join(Path(pathdocespecial).parent,annot_obj.link)))        
                                        lines = ''
                                        try:
                                            filename, extension = os.path.splitext(filepath)
                                            if(extension.lower() in global_settings.listavidformats):
                                                #ffmpeg -ss 01:23:45 -i input -frames:v 1 -q:v 2 output.jpg
                                                executavel = os.path.join(utilities_general.get_application_path(), "ffmpeg")
                                                comando = f"\"{executavel}\" -y -ss 1 -i \"{filepath}\" -frames:v 1 -q:v 2 teste.png"  
                                                #popen = subprocess.Popen(comando, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                                                popen = subprocess.run(comando,universal_newlines=True, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True)
                                                #return_code = popen.wait()
                                                
                                                if(not os.path.exists("teste.png")):
                                                    comando = f"\"{executavel}\" -y -ss 0 -i \"{filepath}\" -frames:v 1 -q:v 2 teste.png"                                              
                                                    #popen = subprocess.Popen(comando, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                                                    popen = subprocess.run(comando,universal_newlines=True, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True)
                                                    
                                                    lines = popen.stderr.split("\n")
                                                    if(len(lines)>0):
                                                        raise classes_general.FFMPEGException(lines)
                                                    
                                                    #return_code = popen.wait()
                                                if(os.path.exists("teste.png")): 
                                                    Image.LOAD_TRUNCATED_IMAGES = True
                                                    with Image.open("teste.png") as imx:
                                                        imx.thumbnail(size, Image.ANTIALIAS)
                                                        
                                                        imx.save('teste.png','PNG')
                                                        width, height = imx.size
                                                    with open('teste.png', 'rb') as f:
                                                        content = f.read()
                                                        pngtohex = binascii.hexlify(content).decode('utf8')
                                                        estaselecao += \
                                    "{{\\rtlch \\ltrch\\loch\\qc\\pict\\pngblip\\picscalex100\\picscaley100\\piccropl0\\piccropr0\\piccropt0\\piccropb0\\picw{}\\pich{}\n{}}}\\line"\
                                        .format(width, height, pngtohex)
                                                        #os.remove('teste.png')
                                                       
                                                        imagens.add((annot_obj.link, pathdocespecial))
                                                        links_img.append((annot_obj.p0y, annot_obj.p1y, pngtohex))
                                                	
                                            else:
                                                Image.LOAD_TRUNCATED_IMAGES = True
                                                with Image.open(filepath) as imx:
                                                   
                                                    imx.thumbnail(size, Image.ANTIALIAS)
                                                    imx.save('teste.png','PNG')
                                                    width, height = imx.size
                                                with open('teste.png', 'rb') as f:
                                                    content = f.read()
                                                    pngtohex = binascii.hexlify(content).decode('utf8')
                                                    estaselecao += \
                                "{{\\rtlch \\ltrch\\loch\\qc\\pict\\pngblip\\picscalex100\\picscaley100\\piccropl0\\piccropr0\\piccropt0\\piccropb0\\picw{}\\pich{}\n{}}}\\line"\
                                    .format(width, height, pngtohex)
                                                    #os.remove('teste.png')
                                                   
                                                    imagens.add((annot_obj.link, pathdocespecial))
                                                    links_img.append((annot_obj.p0y, annot_obj.p1y, pngtohex))
                                        except classes_general.FFMPEGException as ex:
                                            utilities_general.printlogexception(ex=ex)
                                        except Exception as ex:
                                            utilities_general.printlogexception(ex=ex)
                                        finally:
                                            try:
                                                os.remove('teste.png')
                                            except:
                                                None                             
                                        if(withnote and annot_obj.conteudo!=''):
                                            observacao_p0 = "          "
                                            observacao_p1 = "↳ Observação: "
                                            estaselecao += '{\\rtlch\\ltrch\\loch\\fs20\\li72\\f2{'+observacao_p0.encode('rtfunicode').decode('ascii') + '}}'
                                            #observacao_p1 = "Observação do Perito: "
                                            estaselecao += '{\\rtlch\\ltrch\\loch\\fs20\\li72\\f2\\b\\ul{'+observacao_p1.encode('rtfunicode').decode('ascii') + '}}'
                                            estaselecao += '{\\rtlch\\ltrch\\loch\\fs20\\li72\\f2\\i{'+annot_obj.conteudo.encode('rtfunicode').decode('ascii') + '}\\line}'
                                        #break

                 pagi = pagina+1
                 if(estaselecao!=""):
                    estaselecao = estaselecao+f"\\f2(...)\\line"           
                 docname = os.path.basename(pathdocespecial)
                 if(pmin!=None and pmax!=None and pmin!=pmax):
                     docpagina = "{{\\fs22\\f2{{ Relat}}\\\'F3rio \\\'22{}\\\'22 -- Fls. {} a {}}}".format(docname, pmin+1, pmax+1)
                 else:
                     docpagina = "{{\\fs22\\f2{{ Relat}}\\\'F3rio \\\'22{}\\\'22 -- Fls. {}}}".format(docname, pagina+1)   
                 
                 textonatabela += ("\\par\\trowd\\clbrdrb\\brdrs\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrr\\brdrs\\trautofit1\\intbl\\clftsWidth3\\clwWidth9070\\cellx9070{{\\cbpat2\\qc\\loch\\i\\b\\fs22\\f2 TABELA }}{{\\hich\\af2\\af2\\loch\\f2\\f2\\loch\\field\\flddirty{{\\*\\fldinst{{SEQ Tabela\\\\* ARABIC }}}}{{\\fldrslt{{ }}}}}}"+\
                 "{{\\f2\\qc:\\i\\f2{}}}\\f2\\cell\\row"+\
                    # "\\trowd\\clftsWidth1\\clbrdrb\\brdrs\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrr\\brdrs\\cellx1\\intbl{{\\cbpat2\\qc\\loch\\b{{TABELA}}}}\\cell\\row"+\
                     "\\par\\trowd\\clbrdrb\\brdrs\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrr\\brdrs\\trautofit1\\intbl\\clftsWidth3\\clwWidth9070\\cellx9070 {}\\cell\\row\\pard\\line").format(docpagina, estaselecao) + "\n"
                 return (textonatabela, estaselecao)
         except Exception as ex:
             utilities_general.printlogexception(ex=ex)
             return None
                 
            
    def treeview_selection_obs(self, event=None, item=None):        
        try:           
            valores = None
            if(item==None):
                selecao = self.treeviewObs.focus()
                item = selecao
                valores = (self.treeviewObs.item(selecao, 'values'))
            else:
                valores = (self.treeviewObs.item(item, 'values')) 
            tags = self.treeviewObs.item(item, 'tag')
            region = ""
            if(event!=None):
                region = self.treeviewObs.identify("region", event.x, event.y)
            if region == "heading":
                self.orderpopupObs(event)
            elif("obsitem" in tags): 
                for pdf in global_settings.infoLaudo:
                    global_settings.infoLaudo[pdf].retangulosDesenhados = {}                
                self.docInnerCanvas.delete("quad")
                self.docInnerCanvas.delete("simplesearch")
                self.docInnerCanvas.delete("obsitem")
                self.docInnerCanvas.delete("link")
                self.clearSomeImages(["quad", "simplesearch", "obsitem", "link"])
                sobraEspaco = 0
                if(self.docFrame.winfo_width() > global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw * self.zoom_x * global_settings.zoom):
                    sobraEspaco = self.docInnerCanvas.winfo_x()
                posicaoRealY1 = round(float(valores[7]))
                posicaoRealX1 = round(float(valores[6]))
                posicaoRealY0 = round(float(valores[4]))
                posicaoRealX0 = round(float(valores[3]))                
                pp = round(float(valores[2]))
                up = round(float(valores[5]))
                pathpdf = utilities_general.get_normalized_path(valores[1])
                try:
                    self.positions[self.indiceposition] = (global_settings.pathpdfatual, self.vscrollbar.get()[0])
                    self.indiceposition += 1
                    if(self.indiceposition>=10):
                        self.indiceposition = 0
                except Exception as ex:
                    None                
                if(pathpdf!=global_settings.pathpdfatual):
                    pdfantigo = global_settings.pathpdfatual
                    self.docwidth = self.docOuterFrame.winfo_width()                    
                    self.clearAllImages()
                    for i in range(global_settings.minMaxLabels):
                        global_settings.processed_pages[i] = -1
                    pathpdf = utilities_general.get_normalized_path(pathpdf)
                    global_settings.pathpdfatual = pathpdf
                    try:
                        global_settings.docatual.close()
                    except Exception as ex:
                        None
                    global_settings.docatual = fitz.open(global_settings.pathpdfatual)
                    self.labeldocname.config(font=global_settings.Font_tuple_Arial_10, text=os.path.basename(global_settings.pathpdfatual))
                    self.totalPgg.config(font=global_settings.Font_tuple_Arial_10, text="/ "+str(global_settings.infoLaudo[global_settings.pathpdfatual].len))                    
                    if(global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw*self.zoom_x*global_settings.zoom>self.maiorw):
                        self.maiorw = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw*self.zoom_x *global_settings.zoom           
                    self.scrolly = round((global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh*self.zoom_x*global_settings.zoom), 16)*global_settings.infoLaudo[global_settings.pathpdfatual].len - 35
                    self.docInnerCanvas.config(scrollregion=(sobraEspaco, 0, sobraEspaco+ (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw*global_settings.zoom*self.zoom_x), self.scrolly))                
                    if(global_settings.infoLaudo[pdfantigo].zoom_pos!=global_settings.infoLaudo[global_settings.pathpdfatual].zoom_pos):
                        #global_settings.root.after(1, lambda: self.zoomx(None, None, pdfantigo))
                        self.zoomx(None, None, pdfantigo)
                atual = ((self.vscrollbar.get()[0]*global_settings.infoLaudo[global_settings.pathpdfatual].len))
                ondeir = (pp / (global_settings.infoLaudo[global_settings.pathpdfatual].len)+(posicaoRealY0*self.zoom_x*global_settings.zoom- global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw/2*self.zoom_x*global_settings.zoom)/self.scrolly)
                deslocy = (math.floor(pp) * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom) + (posicaoRealY0 *  self.zoom_x * global_settings.zoom)  
                deslocy1 = (math.floor(up) * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom) + (posicaoRealY1 *  self.zoom_x * global_settings.zoom)                   
                desloctotalmenor =  (atual * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom) 
                desloctotalmaior =   desloctotalmenor + self.docFrame.winfo_height() - 20
                if(deslocy < desloctotalmenor or deslocy > desloctotalmaior or deslocy1 > desloctotalmaior):
                    ondeir = ((pp) / (global_settings.infoLaudo[global_settings.pathpdfatual].len)) + (posicaoRealY0*self.zoom_x*global_settings.zoom-self.docFrame.winfo_height()/2)/self.scrolly
                    self.docInnerCanvas.yview_moveto(ondeir)
                    if(str(pp+1)!=self.pagVar.get()):
                        self.pagVar.set(str(pp+1))             
                enhancearea = False
                enhancetext = False
                if(valores[0]=='area'):
                    enhancearea = True
                elif(valores[0]=='texto'):
                    enhancetext = True
                obsobject = self.allobsbyitem[item]
                for p in range(pp, up+1): 
                    posicoes_normalizadas = self.normalize_positions(p, pp, up, \
                                             posicaoRealX0, posicaoRealY0, posicaoRealX1,posicaoRealY1)
                    posicaoRealX0_temp = posicoes_normalizadas[0]
                    posicaoRealY0_temp = posicoes_normalizadas[1]
                    posicaoRealX1_temp = posicoes_normalizadas[2]
                    posicaoRealY1_temp = posicoes_normalizadas[3]
                    self.prepararParaQuads(p, posicaoRealX0_temp, posicaoRealY0_temp, posicaoRealX1_temp, posicaoRealY1_temp, self.color, tag=['obsitem'], apagar=True, \
                                           enhancetext=enhancetext, enhancearea=enhancearea, alt=obsobject.withalt)
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)

    def moveCategory(self, operacao, item):
        if(operacao=='top'):
            self.treeviewObs.move(item, '', 0)
        elif(operacao=='bottom'):
            self.treeviewObs.move(item, '', 'end')
        elif(operacao=='up'):
            self.treeviewObs.move(item, '', self.treeviewObs.index(item)-1)
        elif(operacao=='down'):
            self.treeviewObs.move(item, '', self.treeviewObs.index(item)+1)

    def showhideresults(self):
        try:
            if(self.hideresultsvar.get()):
                termos = self.treeviewSearches.get_children('')
                for termo in termos:
                    results = self.treeviewSearches.get_children(termo)
                    if(len(results)==0):
                        indice = self.treeviewSearches.index(termo)
                        self.detachedSearchResults.append((termo, indice))
                        self.treeviewSearches.detach(termo)
            else:
                for tupla in self.detachedSearchResults:
                    self.treeviewSearches.move(tupla[0], '', tupla[1])
                self.detachedSearchResults = []
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
            
    def importNativeSearchList(self, categoria):        
        application_path = utilities_general.get_application_path()        
        searchlist = os.path.join(application_path, "ListasDeBusca", categoria)
        searchlist = utilities_general.get_normalized_path(searchlist)
        with open(searchlist, "r", encoding='utf-8') as a_file:            
            try:
                for line in a_file:    
                    stripped_line = line.strip()
                    if(stripped_line[0]=="#"):
                        continue
                    tipo = stripped_line.split(" ")[0]
                    termo = stripped_line[len(tipo):len(stripped_line)].strip()
                    if((termo.upper(), tipo.upper()) in global_settings.listaTERMOS):
                        continue  
                    else:
                        global_settings.listaTERMOS[(termo.upper(),tipo.upper())] = []                            
                        if("MATCH" in tipo.upper()):
                            pedidosearch = classes_general.PedidoSearch(global_settings.contador_buscas_incr*-1, "MATCH", termo)
                            #global_settings.searchqueue.append((termo, "MATCH", None))
                            global_settings.searchqueue.put(pedidosearch)
                            global_settings.contador_buscas_incr += 1
                        elif("REGEX" in tipo.upper()):
                            pedidosearch = classes_general.PedidoSearch(global_settings.contador_buscas_incr*-1, "REGEX", termo)
                            #global_settings.searchqueue.append((termo,"REGEX", None))
                            global_settings.searchqueue.put(pedidosearch)
                            global_settings.contador_buscas_incr += 1
                        else:
                            if(len(termo)>=3):
                                pedidosearch = classes_general.PedidoSearch(global_settings.contador_buscas_incr*-1, "LIKE", termo)
                                #global_settings.searchqueue.append((termo, "LIKE", None))
                                global_settings.searchqueue.put(pedidosearch)
                                global_settings.contador_buscas_incr += 1
                self.primeiroresetbuscar = True
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)  
            
    def importListPopUp(self):   
        try:
            self.opcaoimportlist = tkinter.Menu(global_settings.root, tearoff=0)
            self.opcaoimportlist.add_command(label='Pornografia Infantil', image=global_settings.childpornb, compound='left', command=partial(self.importNativeSearchList, "lista_vulneraveis.txt"))
            self.opcaoimportlist.add_command(label='Armas/Drogas', image=global_settings.gunb, compound='left', command=partial(self.importNativeSearchList, "lista_reupreso.txt"))
            self.opcaoimportlist.add_command(label='Violência', image=global_settings.violenceb, compound='left', command=partial(self.importNativeSearchList, "lista_homicidios.txt"))
            self.opcaoimportlist.add_separator()
            self.opcaoimportlist.add_command(label='Arquivo de texto', image=global_settings.textb, compound='left', command=self.openSearchlist)
            self.opcaoimportlist.tk_popup(self.bfromFile.winfo_rootx()+50,self.bfromFile.winfo_rooty())         
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        finally:
            self.opcaoimportlist.grab_release()
            
    def exportInterval(self, event=None):
        doctoexport = tkinter.Label(self.exportinterval.window, font=global_settings.Font_tuple_Arial_10, text=global_settings.pathpdfatual)
        doctoexport.grid(row=0, column=0, columnspan=2, sticky='nsew', pady=5, padx=5)
        
        self.exportinterval.show()
            
    def filterdocWindow(self, event=None):        
        if(self.windowfilter==None):
            docs = []
            tupla = []
            linha = 0
            for eq in self.treeviewEqs.get_children(""):
                for doc in self.treeviewEqs.get_children(eq):
                    nomedoc = self.treeviewEqs.item(doc, "text")
                    docs.append(nomedoc)
                    tupla.append(linha)
                    linha += 1
            tupla.append(linha+1)
            tupla.append(linha+2)
            self.windowfilter = tkinter.Toplevel()  
            self.windowfilter.rowconfigure(0, weight=1)
            self.windowfilter.columnconfigure((0,1), weight=1)
            self.filterframedcanvas = tkinter.Canvas(self.windowfilter)
            self.filterframedcanvas.grid(row=0, column=0, columnspan=2, sticky='nsew', pady=(0,10))
            self.filterframedoc = tkinter.Frame(self.windowfilter)
            self.filterframedoc.grid(row=0, column=0, sticky='nsew')
            self.filterframedoc.rowconfigure(tuple(tupla), weight=1)
            self.filterframedoc.columnconfigure(0, weight=1)            
            self.filterframedcanvasreturn = self.filterframedcanvas.create_window((0,0), window=self.filterframedoc, anchor = "nw")            
            self.vsbfilter = tkinter.Scrollbar(self.windowfilter, orient="vertical")            
            self.vsbfilter.config(command=self.filterframedcanvas.yview)
            self.vsbfilter.grid(row=0, column=2, sticky='ns')
            self.filterframedcanvas.config(yscrollcommand = self.vsbfilter.set)
            linha = 0
            alldocs = tkinter.Checkbutton(self.filterframedoc, font=global_settings.Font_tuple_Arial_10, text="SELECIONAR TODOS")
            nonedocs = tkinter.Checkbutton(self.filterframedoc, font=global_settings.Font_tuple_Arial_10, text="DESMARCAR TODOS")
            self.cbsfilters = []
            self.cbsfilters.append(alldocs, None)
            self.cbsfilters.append(nonedocs, None)
            for nomedoc in docs:
                cb = tkinter.Checkbutton(self.filterframedoc, font=global_settings.Font_tuple_Arial_10, text="{}".format(nomedoc.strip()))
                self.cbsfilters.append(cb, nomedoc)
                cb.grid(row=linha+2, column=0, sticky='w', pady=(0,10))
                linha += 1          
            self.windowfilter.title("Filtrar observações por documento")
            botaoaplicar = tkinter.button(self.windowfilter, font=global_settings.Font_tuple_Arial_10, text='Aplicar', image=global_settings.checki, compound='right', command= self.applyFilterDoc)
            botaocancelar = tkinter.button(self.windowfilter, font=global_settings.Font_tuple_Arial_10, text='Cancelar', image=global_settings.checki, compound='right', command= lambda: self.windowfilter.withdraw())          
        self.windowfilter.deiconify()  
   
    def saveSearchResults(self):        
        window = tkinter.Toplevel()
        try:           
            window.geometry("800x600")
            var = tkinter.BooleanVar()
            window.protocol("WM_DELETE_WINDOW", lambda: var.set(False))
        
            window.title("Exportar resultados (CSV)")
            window.rowconfigure(1, weight=1)
            window.columnconfigure((0,1), weight=1)
            
            labeltermos = tkinter.Label(window, font=global_settings.Font_tuple_Arial_10, text="Termos a serem exportados:")
            labeltermos.grid(row=0, column=0, sticky='ns', pady=5)
            labeldocs = tkinter.Label(window, font=global_settings.Font_tuple_Arial_10, text="Nos documentos:")
            labeldocs.grid(row=0, column=1, sticky='ns', pady=5)
            
            frametermos = tkinter.Frame(window)
            frametermos.rowconfigure(0, weight=1)
            frametermos.columnconfigure(0, weight=1)
            frametermos.grid(row=1, column=0, sticky='nsew', pady=5)
            termosvar = tkinter.StringVar()
            lbtermos = tkinter.Listbox(frametermos, listvariable = termosvar, selectmode=tkinter.EXTENDED, exportselection=False)
            lbtermos.grid(row=0, column=0, sticky='nsew', pady=2)
                        
            framedocs = tkinter.Frame(window)
            framedocs.rowconfigure(0, weight=1)
            framedocs.columnconfigure(0, weight=1)
            framedocs.grid(row=1, column=1, sticky='nsew', pady=5, padx=10)
            docsvar = tkinter.StringVar()
            lbdocs = tkinter.Listbox(framedocs, listvariable = docsvar, selectmode=tkinter.EXTENDED, exportselection=False)
            lbdocs.grid(row=0, column=0, sticky='nsew', pady=5)
            insertdocs = []
            
            scroltermos = ttk.Scrollbar(frametermos, orient="vertical")
            scroltermos.grid(row=0, column=1, sticky='ns')
            scroltermos.config( command = lbtermos.yview )
            lbtermos.configure(yscrollcommand=scroltermos.set)
            #--
            scroltermos2 = ttk.Scrollbar(frametermos, orient="horizontal")
            scroltermos2.grid(row=1, column=0, sticky='ew')
            scroltermos2.config( command = lbtermos.xview )
            lbtermos.configure(xscrollcommand=scroltermos2.set)
                        
            scroldocs = ttk.Scrollbar(framedocs, orient="vertical")
            scroldocs.grid(row=0, column=1, sticky='ns')
            scroldocs.config( command = lbdocs.yview )
            lbdocs.configure(yscrollcommand=scroldocs.set)
            #--
            scroldocs2 = ttk.Scrollbar(framedocs, orient="horizontal")
            scroldocs2.grid(row=1, column=0, sticky='ew')
            scroldocs2.config( command = lbdocs.xview )
            lbdocs.configure(xscrollcommand=scroldocs2.set)
            termosdict = []
            docsdict = []
            index = 0
            for child in self.treeviewEqs.get_children(''):
                for child2 in self.treeviewEqs.get_children(child):
                    pdf = os.path.basename(self.treeviewEqs.item(child2, 'values')[1])
                    lbdocs.insert(tkinter.END, pdf)
                    docsdict.append(pdf)
                
            inserttermos = []
            for child in self.treeviewSearches.get_children(''):
                texto = self.treeviewSearches.item(child, 'text')
                lbtermos.insert(tkinter.END, texto)
                termosdict.append(child)
            answeryes = tkinter.Button(window, font=global_settings.Font_tuple_Arial_10, text="Exportar", command = lambda: var.set(True))
            answeryes.grid(row=2, column=0, columnspan=4, sticky='ns', pady=15)
            global_settings.root.wait_variable(var)
            if(var.get()):
                tipos = [('XLSX', '*.xls')]
                path = (asksaveasfilename(filetypes=tipos, defaultextension=tipos))
                if(path!=None and path!=''):
                    workbook = xlsxwriter.Workbook(path)
                    try:
                    #with open(path, mode='w', newline='', encoding='utf-8') as csv_file:
                        worksheet = workbook.add_worksheet()
                        bold = workbook.add_format({'bold': True})
                        cell_formatdarkgray = workbook.add_format()                       
                        cell_formatdarkgray.set_bg_color("#bfbfbf")
                        cell_formatdarkgray.set_bold()
                        cell_formatlightgray = workbook.add_format()                       
                        cell_formatlightgray.set_bg_color("#e6e6e6")
                        worksheet.write(0, 0, "Termo", cell_formatdarkgray )
                        worksheet.write(0, 1, "Tipo de busca", cell_formatdarkgray )
                        worksheet.write(0, 2, "Documento", cell_formatdarkgray )
                        worksheet.write(0, 3, "Seção", cell_formatdarkgray )
                        worksheet.write(0, 4, "Página", cell_formatdarkgray )
                        worksheet.write(0, 5, "Trecho", cell_formatdarkgray )
                        termosexport = lbtermos.curselection()
                        docsexport = lbdocs.curselection()
                        linha= 1
                        gray = False
                        maiortermo = ""
                        maiortipo = "Tipo de busca"
                        maiornomepdf = ""
                        maiortoctext = ""
                        maiorpagina = ""
                        maiortrecho = ""
                        if(len(termosexport)>0 and len(docsexport) > 0):
                            docsselecionados = []
                            for docindex in docsexport:
                                docsselecionados.append(docsdict[docindex])
                            for term in termosexport:
                                childsearched = termosdict[term]
                                valores = self.treeviewSearches.item(childsearched, 'values')
                                termo = valores[0]
                                tipo= valores[1]
                                filhospdf = self.treeviewSearches.get_children(childsearched)
                                for pdfsearched in filhospdf:
                                    nomepdf = os.path.basename(self.treeviewSearches.item(pdfsearched, 'values')[0])
                                    if(nomepdf in docsselecionados):
                                        childofpdf = self.treeviewSearches.get_children(pdfsearched)
                                        for child in childofpdf:
                                            if(self.treeviewSearches.tag_has('relsearchtoc', child)):
                                                toc = self.treeviewSearches.get_children(child)
                                                toctext = self.treeviewSearches.item(child, 'text')
                                                for res in toc:
                                                    valoresres = self.treeviewSearches.item(res, 'values')
                                                    textores =  self.treeviewSearches.item(res, 'text')
                                                    pagina = textores.split(" - ")[0]
                                                    snippet = valoresres[0] + " <b>" + valoresres[1] + "<\\b> " + valoresres[2]
                                                    if(gray):
                                                        worksheet.write(linha, 0, termo, cell_formatlightgray )
                                                        worksheet.write(linha, 1, tipo, cell_formatlightgray )
                                                        worksheet.write(linha, 2, nomepdf, cell_formatlightgray )
                                                        worksheet.write(linha, 3, toctext, cell_formatlightgray )
                                                        worksheet.write(linha, 4, pagina, cell_formatlightgray )
                                                        worksheet.write_rich_string(linha, 5, valoresres[0]+" ", bold, valoresres[1]+" ", valoresres[2]+" ", cell_formatlightgray )
                                                        if(len(termo) > len(maiortermo)):
                                                            maiortermo = termo
                                                        if(len(tipo) > len(maiortipo)):
                                                            maiortipo = tipo
                                                        if(len(nomepdf) > len(maiornomepdf)):
                                                            maiornomepdf = nomepdf
                                                        if(len(toctext) > len(maiortoctext)):
                                                            maiortoctext = toctext
                                                        if(len(pagina) > len(maiorpagina)):
                                                            maiorpagina = pagina
                                                        if(len(snippet) > len(maiortrecho)):
                                                            maiortrecho = snippet
                                                        linha += 1
                                                        gray = not gray
                                                    else:
                                                        worksheet.write(linha, 0, termo )
                                                        worksheet.write(linha, 1, tipo)
                                                        worksheet.write(linha, 2, nomepdf )
                                                        worksheet.write(linha, 3, toctext )
                                                        worksheet.write(linha, 4, pagina )
                                                        worksheet.write_rich_string(linha, 5, valoresres[0]+" ", bold, valoresres[1]+" ", valoresres[2]+" " )
                                                        linha += 1
                                                        if(len(termo) > len(maiortermo)):
                                                            maiortermo = termo
                                                        if(len(tipo) > len(maiortipo)):
                                                            maiortipo = tipo
                                                        if(len(nomepdf) > len(maiornomepdf)):
                                                            maiornomepdf = nomepdf
                                                        if(len(toctext) > len(maiortoctext)):
                                                            maiortoctext = toctext
                                                        if(len(pagina) > len(maiorpagina)):
                                                            maiorpagina = pagina
                                                        if(len(snippet) > len(maiortrecho)):
                                                            maiortrecho = snippet
                                                        gray = not gray
                                                    #writer.writerow([termo+tipo, nomepdf, toctext, pagina, snippet])
                                            elif(self.treeviewSearches.tag_has('resultsearch', child)):
                                                res = child
                                                valoresres = self.treeviewSearches.item(res, 'values')
                                                textores =  self.treeviewSearches.item(res, 'text')
                                                pagina = textores.split(" - ")[0]
                                                snippet = valoresres[0] + " <b>" + valoresres[1] + "<\\b> " + valoresres[2]
                                                toctext = ""
                                                # writer.writerow([termo+tipo, nomepdf, "-", pagina, snippet])
                                                if(gray):
                                                    worksheet.write(linha, 0, termo, cell_formatlightgray )
                                                    worksheet.write(linha, 1, tipo, cell_formatlightgray )
                                                    worksheet.write(linha, 2, nomepdf, cell_formatlightgray )
                                                    worksheet.write(linha, 3, toctext, cell_formatlightgray )
                                                    worksheet.write(linha, 4, pagina, cell_formatlightgray )
                                                    worksheet.write_rich_string(linha, 5, valoresres[0]+" ", bold, valoresres[1]+" ", valoresres[2]+" ", cell_formatlightgray )
                                                    linha += 1
                                                    gray = not gray
                                                    if(len(termo) > len(maiortermo)):
                                                        maiortermo = termo
                                                    if(len(tipo) > len(maiortipo)):
                                                        maiortipo = tipo
                                                    if(len(nomepdf) > len(maiornomepdf)):
                                                        maiornomepdf = nomepdf
                                                    if(len(toctext) > len(maiortoctext)):
                                                        maiortoctext = toctext
                                                    if(len(pagina) > len(maiorpagina)):
                                                        maiorpagina = pagina
                                                    if(len(snippet) > len(maiortrecho)):
                                                        maiortrecho = snippet
                                                else:
                                                    worksheet.write(linha, 0, termo )
                                                    worksheet.write(linha, 1, tipo)
                                                    worksheet.write(linha, 2, nomepdf )
                                                    worksheet.write(linha, 3, toctext )
                                                    worksheet.write(linha, 4, pagina )
                                                    worksheet.write_rich_string(linha, 5, valoresres[0]+" ", bold, valoresres[1]+" ", valoresres[2]+" ")
                                                    linha += 1
                                                    gray = not gray
                                                    if(len(termo) > len(maiortermo)):
                                                        maiortermo = termo
                                                    if(len(tipo) > len(maiortipo)):
                                                        maiortipo = tipo
                                                    if(len(nomepdf) > len(maiornomepdf)):
                                                        maiornomepdf = nomepdf
                                                    if(len(toctext) > len(maiortoctext)):
                                                        maiortoctext = toctext
                                                    if(len(pagina) > len(maiorpagina)):
                                                        maiorpagina = pagina
                                                    if(len(snippet) > len(maiortrecho)):
                                                        maiortrecho = snippet
                            worksheet.set_column(0, 0, len(maiortermo)+5)
                            worksheet.set_column(1, 1, len(maiortipo)+5)
                            worksheet.set_column(2, 2, len(maiornomepdf)+5)
                            worksheet.set_column(3, 3, len(maiortoctext)+5)
                            worksheet.set_column(4, 4, len(maiorpagina)+5)
                            worksheet.set_column(5, 5, len(maiortrecho)+5)
                    except Exception as ex:
                        utilities_general.printlogexception(ex=ex)
                    finally:
                        workbook.close()
        except Exception as ex:
            utilities_general.printlogexception(ex=ex) 
        finally:
            window.destroy()
     
            
    def windSearchResults(self, texto=""):
        window = tkinter.Toplevel()
        window.overrideredirect(True)
        window.columnconfigure(0, weight=1)
        window.rowconfigure((0,1), weight=1)
        label = tkinter.Label(window, font=global_settings.Font_tuple_Arial_10, text=texto, image=global_settings.sync, compound='top')
        label.grid(row=0, column=0, sticky='ew', padx=50, pady=10)
        progresssearch = ttk.Progressbar(window, mode='indeterminate')
        progresssearch.grid(row=1, column=0, sticky='ew', padx=50, pady=10)
        return  (window, progresssearch)
    def saveSearchlist(self):    
        tipos = [('Texto', '*.txt')]
        path = (asksaveasfilename(filetypes=tipos, defaultextension=tipos))
        try:
            if(path!=None and path!=''):
                with open(path, "w", encoding='utf8') as a_file:
                    for (termo, tipobusca) in global_settings.listaTERMOS:
                        a_file.write('{} {}\n'.format(tipobusca, termo))
        except Exception as ex:
            utilities_general.printlogexception(ex=ex) 
    
           
    
    def insertIndex(self, tree, parent, texto_candidato, index=0):
        children = tree.get_children(parent)
        
        for child in children:
            texto = tree.item(child, 'text')
            if(len(texto_candidato) < len(texto)):
                break
            elif(len(texto_candidato) >= len(texto) and texto_candidato < texto):
                break
            index += 1
        return index
    
    def sanitize_search(self):
        
        
        if("\n" in self.searchVar.get()):
            #print(self.searchVar.get(), self.searchtypeVar.get())
            old = self.searchVar.get().replace("\n",' ').replace("\r",' ')
            self.entrysearch.delete("0", "end")
            self.entrysearch.insert("0", old)
            #self.searchVar.set(self.searchVar.get().replace("\n",'').replace("\r",''))
            
    def leftPanel(self):
          
        global_settings.pathpdfatual = None
        try:
            self.windowfilter = None            
            self.detachedSearchResults = []
            self.infoFrame = tkinter.Frame(borderwidth=2, bg='white', relief='ridge')   
            self.menuexcludesearch = tkinter.Menu(global_settings.root, tearoff=0)
            self.menuexcludesearch.add_command(label="Excluir Busca", image=global_settings.delcat, compound='right', command=self.exclude_search)            
            self.infoFrame.rowconfigure(1, weight=1)
            self.infoFrame.columnconfigure(0, weight=1)            
            self.logoframe = tkinter.Frame(self.infoFrame, highlightthickness=0)
            self.logoframe.rowconfigure(0, weight=1)
            self.logoframe.columnconfigure(0, weight=1)
            self.logoframe.grid(row=0, column=0, sticky='nswe')
            self.labelpcp = tkinter.Label(self.logoframe, image=global_settings.tkphotologo2)
            self.labelpcp.grid(row=0, column=0, sticky='n')            
            self.notebook = ttk.Notebook(self.infoFrame, padding=8)
            self.notebook.bind("<ButtonRelease-1>", self.tabOpened)
            self.notebook.grid(row=1, column=0, sticky='nsew')
            self.tocOuterFrame = tkinter.Frame()
            self.tocOuterFrame.rowconfigure(0, weight=1)
            self.tocOuterFrame.columnconfigure(0, weight=1)
            self.canvastoc = tkinter.Canvas(self.tocOuterFrame)
            self.canvastoc.grid(row=0, column=0, sticky="nsew")
            self.tocFrame = tkinter.Frame(borderwidth=2, relief='ridge')
            self.tocFrame.rowconfigure(1, weight=1)
            self.tocFrame.columnconfigure((0, 1), weight=1)
            self.treeviewEqs = ttk.Treeview(self.tocFrame, selectmode='browse')
            self.collapseeqs = tkinter.Button(self.tocFrame, font=global_settings.Font_tuple_Arial_10, text='Colapsar todos', \
                                              image=global_settings.collapseimg, compound="right", command=self.collapsealleqs)
            self.collapseeqs.grid(row=0, column=0,sticky='n', padx=10, pady=5)   
            
            self.expandeqs = tkinter.Button(self.tocFrame, font=global_settings.Font_tuple_Arial_10, text='Expandir todos', \
                                              image=global_settings.expandimg, compound="right", command=self.expandalleqs)
            self.expandeqs.grid(row=0, column=1,sticky='n', padx=10, pady=5)   
            
            self.treeviewEqs.grid(row=1, column=0, sticky='nsew', columnspan=2)
            self.treeviewEqs.heading("#0", text="Equipamentos / Relatorios", anchor="w")
            self.scrolltoc = ttk.Scrollbar(self.tocFrame, orient="vertical")
            self.scrolltoc.config( command = self.treeviewEqs.yview )
            self.treeviewEqs.configure(yscrollcommand=self.scrolltoc.set)
            self.scrolltoch = ttk.Scrollbar(self.tocFrame, orient="horizontal")
            self.scrolltoch.config( command = self.treeviewEqs.xview )
            
            self.locFrame = tkinter.Frame(borderwidth=2, relief='ridge')
            self.locFrame.rowconfigure(1, weight=1)
            self.locFrame.columnconfigure((0, 1), weight=1)
            self.treeviewLocs = ttk.Treeview(self.locFrame, selectmode='browse')
            #self.collapseeqs = tkinter.Button(self.tocFrame, font=global_settings.Font_tuple_Arial_10, text='Colapsar todos', \
            #                                  image=global_settings.collapseimg, compound="right", command=self.collapsealleqs)
            #self.collapseeqs.grid(row=0, column=0,sticky='n', padx=10, pady=5)   
            
            #self.expandeqs = tkinter.Button(self.tocFrame, font=global_settings.Font_tuple_Arial_10, text='Expandir todos', \
            #                                  image=global_settings.expandimg, compound="right", command=self.expandalleqs)
            #self.expandeqs.grid(row=0, column=1,sticky='n', padx=10, pady=5)   
            
            self.treeviewLocs.grid(row=1, column=0, sticky='nsew', columnspan=2)
            self.treeviewLocs.heading("#0", text="Localizações", anchor="w")
            self.scrollloc = ttk.Scrollbar(self.locFrame, orient="vertical")
            self.scrollloc.config( command = self.treeviewLocs.yview )
            self.treeviewLocs.configure(yscrollcommand=self.scrollloc.set)
            self.scrollloch = ttk.Scrollbar(self.tocFrame, orient="horizontal")
            self.scrollloch.config( command = self.treeviewLocs.xview )
            #self.treeviewLocs.bindtags(('.self.treeviewLocs', 'Treeview', 'post-tree-bind','.','all'))
            self.treeviewLocs.bind("<1>", lambda e: self.treeview_selection_loc_right(e))
            
            self.treeviewEqs.configure(xscrollcommand=self.scrolltoch.set)
            self.scrolltoch.grid(row=2, column=0, sticky='ew', columnspan=2)
            self.treeviewEqs.bindtags(('.self.treeviewEqs', 'Treeview', 'post-tree-bind', 'post-tree-teste','.','all'))
            self.treeviewEqs.bind_class('post-tree-bind', "<1>", lambda e: self.treeview_selection(e))
            self.treeviewEqs.bind_class('post-tree-bind','<Right>',lambda e: self.treeview_selection(e))
            self.treeviewEqs.bind_class('post-tree-bind','<Left>',lambda e: self.treeview_selection(e))
            self.treeviewEqs.bind_class('post-tree-bind','<Up>', lambda e: self.treeview_selection(e))
            self.treeviewEqs.bind_class('post-tree-bind','<Down>', lambda e: self.treeview_selection(e))
            self.treeviewEqs.bind_class('post-tree-bind', "<3>", self.treeview_eqs_right)
            treevieweqtt = classes_general.CreateToolTip(self.treeviewEqs, "Equipamentos / Relatórios", istreeview=True, classe='post-tree-bind')
            self.scrolltoc.grid(row=1, column=2, sticky='ns')
            maiorresult = 0
            self.primeiro = None
            #rels = global_settings.infoLaudo.keys()
            index = 0
            locations = set()
            for relatorio in sorted(global_settings.infoLaudo):
                parente = os.path.dirname(relatorio)
                if(os.path.exists(os.path.join(parente, 'locations'))):
                    nomeeq = os.path.basename(parente)
                    if(nomeeq not in locations): 
                        locations.add(nomeeq)
                        location = os.path.join(parente, 'locations', 'index.html')
                        
                        id = self.treeviewLocs.insert(parent='', index='end', iid=nomeeq, text=f"Localizações {nomeeq} ",\
                                                image=global_settings.locationsmallicon, tag='locationlp', values=('location', location, f"Localizações {nomeeq}",))
                        if plt=="Windows":
                            self.treeviewLocs.insert(parent=id, index='end', text=f"Abrir no FERA",\
                                                    image=global_settings.feraiconbi, tag='locationlpchild', values=('location', location, f"Localizações {nomeeq}",))
                        self.treeviewLocs.insert(parent=id, index='end', text=f"Abrir no navegador padrão",\
                                                image=global_settings.browsericonbi, tag='locationlpchild', values=('location', location, f"Localizações {nomeeq}",))
                        self.treeviewLocs.item(id, open=True)
                p = Path(relatorio)
                pai = p.parent
                paibase = os.path.basename(pai)
                ok = False
                if(global_settings.infoLaudo[relatorio].parent_alias != None and \
                   global_settings.infoLaudo[relatorio].parent_alias!=''):
                    paibase = global_settings.infoLaudo[relatorio].parent_alias
                    ok = True
                else:
                    for k in range(3):
                        if("EQ" in paibase.upper()):
                            ok = True
                            break
                        else:
                            pai = pai.parent
                            paibase = os.path.basename(pai)
                if(not ok):
                    paibase = "Documentos"
                    pai = p.parent
                pdfbase = os.path.basename(p)
                tipo = "pdf"
                paibase = paibase.strip()
                if(not paibase.upper() in global_settings.equipments):
                    global_settings.equipments[paibase.upper()] = []
                global_settings.equipments[paibase.upper()].append(pdfbase)
                try:   
                    if(not self.treeviewEqs.exists(paibase.upper())):
                        if(global_settings.infoLaudo[relatorio].tipo=='laudo'):
                            index += 1
                            self.treeviewEqs.insert(parent='', index='0', iid=paibase.upper(), text='LAUDO',\
                                                    image=global_settings.imageequip, tag='equipmentlp', values=('eq', str(paibase),))                        
                        else:
                            index_insert = self.insertIndex(self.treeviewEqs, '', paibase.upper(), index)
                            self.treeviewEqs.insert(parent='', index=index_insert, iid=paibase.upper(), text=paibase.upper(), \
                                                    image=global_settings.imageequip, tag='equipmentlp', values=('eq', str(paibase),))
                except Exception as ex:
                    #None
                    utilities_general.printlogexception(ex=ex)
                index_insert = self.insertIndex(self.treeviewEqs, paibase.upper(), pdfbase, 0)
                
                self.treeviewEqs.insert(parent=paibase.upper(), index=index_insert, iid=str(p), text=pdfbase, tag='reportlp',\
                                        image=global_settings.imagereportb, values=(tipo, str(p), global_settings.infoLaudo[relatorio].id,))                
                self.treeviewEqs.see(str(p))
                for t in global_settings.infoLaudo[relatorio].toc:
                    #try:
                        nivel = t[0].split(' ')[0].split('.')
                        ident = ''
                        for k in range(len(nivel)):
                            ident += '     '
                        self.treeviewEqs.insert(parent=str(p), index='end', iid=str(p)+t[0]+str(t[1])+str(t[2]),\
                                                text=ident+t[0], tag="tocpdf", values=('toc', str(p), t[0], t[1], t[2],))
                        somatexto = paibase.upper()+pdfbase+t[0]
                        tamanho = global_settings.Font_tuple_ArialBold_10.measure(pdfbase)+150
                        if(tamanho>maiorresult):
                            maiorresult = tamanho
                            self.treeviewEqs.column("#0", width=maiorresult, stretch=True, minwidth=maiorresult, anchor="w")
                    #except:
                    #    None
                #self.treeviewEqs.bind_class('post-tree-teste','<<TreeviewClose>>', lambda e: self.treeview_open(e))
            firstreport = self.treeviewEqs.get_children(self.treeviewEqs.get_children('')[0])[0]
            firstreportpdf = self.treeviewEqs.item(firstreport, 'values')[1]
            #if(global_settings.pathpdfatual==None):
            global_settings.pathpdfatual = utilities_general.get_normalized_path(firstreportpdf)              
            self.primeiro = firstreportpdf
            try:
                global_settings.docatual.close()
            except Exception as ex:
                None
            global_settings.docatual = fitz.open(global_settings.pathpdfatual)
            global_settings.zoom = global_settings.listaZooms[global_settings.infoLaudo[global_settings.pathpdfatual].zoom_pos]
            self.treeviewLocs.tag_configure('locationlp', background='#a1a1a1', font=global_settings.Font_tuple_ArialBoldUnderline_16)
            self.treeviewLocs.tag_configure('locationlpchild', background='#a1a1a1', font=global_settings.Font_tuple_ArialBold_14)
            self.treeviewEqs.tag_configure('equipmentlp', background='#a1a1a1', font=global_settings.Font_tuple_ArialBoldUnderline_12)
            self.treeviewEqs.tag_configure('reportlp', background='#ebebeb', font=global_settings.Font_tuple_ArialBold_10)     
            self.treeviewEqs.tag_configure('tocpdf', font=global_settings.Font_tuple_Arial_8)   
            self.canvastoc.create_window((0,0), window=self.tocFrame, anchor="nw")            
            self.searchFrame = tkinter.Frame(borderwidth=2, bg='white')
            self.searchFrame.rowconfigure(1, weight=1)
            self.searchFrame.columnconfigure(0, weight=1)
            self.searchBox = tkinter.Frame(self.searchFrame, borderwidth=2, relief='ridge')
            self.searchBox.grid(row=0, column=0, sticky='new', pady=(0, 5))
            self.searchBox.rowconfigure((0,1,2), weight=1)
            self.searchBox.columnconfigure(0, weight=2)
            self.searchBox.columnconfigure(1, weight=1)
            self.searchVar = tkinter.StringVar()
            self.searchVar.set("")
            self.searchtypeVar = tkinter.StringVar()
            self.searchtypeVar.set("LIKE")
            self.searchVar.trace_add("write", lambda name, index, mode: self.sanitize_search())
            
            self.entrysearch = classes_general.PlaceholderEntry(self.searchBox, placeholder='Buscar...', justify='center', textvariable=self.searchVar, state='normal',\
                                                                exportselection=False)
            self.entrysearch.bind('<Return>',  lambda e: self.searchTerm(event=e))
            self.entrysearch.grid(row=1, column=0, sticky='nsew', padx=2, pady=5)            
            self.limitSearchFrame = tkinter.Frame(self.searchBox)
            self.limitSearchFrame.grid(row=0, column=0, sticky='w', pady=2)
            self.limitSearchFrame.rowconfigure(0, weight=1)
            self.limitSearchFrame.columnconfigure(0, weight=1)
            self.limitsearchVar = tkinter.IntVar()
            self.limitsearchVar.set(3000)
            self.limitsearchlabel = tkinter.Label(self.limitSearchFrame, font=global_settings.Font_tuple_Arial_10, text='*Max. Resultados por seção: 3000')
            self.limitsearchlabel.grid(row=0, column=0, sticky='n', pady=2)
            self.querysql = tkinter.Button(self.searchBox, font=global_settings.Font_tuple_Arial_10, text='Ajuda', \
                                           image=global_settings.querysqlim, compound="right", command=self.querySql)
            self.querysql.grid(row=0, column=1, sticky='e', pady=5)
            self.searchbutton = tkinter.Button(self.searchBox, font=global_settings.Font_tuple_Arial_10, text='Pesquisar', \
                                               image=global_settings.lupa, compound="right", state='normal', command= lambda: self.searchTerm())
            self.searchbutton.grid(row=1, column=1, sticky='e', pady=2)
            self.searchtypeFrame = tkinter.Frame(self.searchBox)
            self.searchtypeFrame.grid(row=2, column=0, columnspan=2, pady=0, padx=0, sticky='nsew')
            self.searchtypeFrame.rowconfigure(0, weight=1)
            self.searchtypeFrame.columnconfigure((0,1,2), weight=1)
            
            self.searchtypeLike = tkinter.Radiobutton(self.searchtypeFrame,variable=self.searchtypeVar, font=global_settings.Font_tuple_Arial_10, text="LIKE", value="LIKE")
            self.searchtypeMatch = tkinter.Radiobutton(self.searchtypeFrame,variable=self.searchtypeVar, font=global_settings.Font_tuple_Arial_10, text="MATCH", value="MATCH")
            self.searchtypeRegex = tkinter.Radiobutton(self.searchtypeFrame,variable=self.searchtypeVar, font=global_settings.Font_tuple_Arial_10, text="REGEX", value="REGEX")
            self.searchtypeLike.grid(row=0, column=0, sticky='ns')
            self.searchtypeMatch.grid(row=0, column=1, sticky='ns')
            self.searchtypeRegex.grid(row=0, column=2,  sticky='ns')            
            sep = ttk.Separator(self.searchBox)
            sep.grid(row=3, column=0, columnspan=2, sticky='nsew', pady=5)            
            self.searchlistframe = tkinter.Frame(self.searchBox)
            self.searchlistframe.grid(row=4, column=0, columnspan=2, sticky='nsew', pady=(0, 5))
            self.searchlistframe.rowconfigure(0, weight=1)
            self.searchlistframe.columnconfigure((0,1,2), weight=1)
            self.bfromFile = tkinter.Button(self.searchlistframe, font=global_settings.Font_tuple_Arial_10, text="Importar", \
                                            image=global_settings.imfromFile, compound="right", state='normal', command=self.importListPopUp)
            self.bfromFile.image = global_settings.imfromFile
            self.bfromFile.grid(row=0, column=0, pady=2, sticky='n')  
            self.btoFile = tkinter.Button(self.searchlistframe, font=global_settings.Font_tuple_Arial_10, text="Exportar", \
                                          image=global_settings.imtoFile, compound="right", state='normal', command=self.saveSearchlist)
            self.btoFile.image = global_settings.imtoFile
            self.btoFile.grid(row=0, column=1, sticky='n', pady=2) 
            self.saveresulttocsv = tkinter.Button(self.searchlistframe, font=global_settings.Font_tuple_Arial_10, text="Salvar (CSV)", \
                                                  image=global_settings.copyrestoblip, compound="right", state='normal', command=self.saveSearchResults)
            self.saveresulttocsv.image = global_settings.copyrestoblip
            self.saveresulttocsv.grid(row=0, column=2, sticky='n', pady=2) 
            self.searchEnv = tkinter.Frame(self.searchFrame, borderwidth=2, relief='ridge')
            self.searchEnv.grid(row=1, column=0, sticky='nsew')
            self.searchEnv.rowconfigure(1, weight=1)
            self.searchEnv.columnconfigure(0, weight=1)
            self.searchbuttonsframe = tkinter.Frame(self.searchEnv, borderwidth=2, relief='ridge')
            self.searchbuttonsframe.grid(row=0, column=0, sticky='nsew')
            self.searchbuttonsframe.rowconfigure((0,1,2,3), weight=1)
            self.searchbuttonsframe.columnconfigure((0,1,2), weight=1)
            bnextfind = tkinter.Button(self.searchbuttonsframe, image=global_settings.imnextfind, font=global_settings.Font_tuple_Arial_10, text="", \
                                       compound='left', command=lambda: self.iterateSearchList(None, 'proximo'))
            bnextfind.image = global_settings.imnextfind
            bnextfind.grid(column=2, row=0, sticky='ns', padx=10, pady=5, rowspan=2) 
            
            self.ocorrenciasLabel = tkinter.Label(self.searchbuttonsframe, font=global_settings.Font_tuple_Arial_10, text="-- de -----")
            self.ocorrenciasLabel.grid(row=1, column=1, sticky='nsew', pady=2)            
            self.termosearchVar = tkinter.StringVar(self.searchbuttonsframe)
            self.termosearchVar.set("")
            self.termosearched = tkinter.Entry(self.searchbuttonsframe, justify='center', textvariable=self.termosearchVar, state='disabled', exportselection=False)
            self.termosearched.grid(row=0, column=1, sticky='nsew', pady=5)            
            self.hideresultsvar = tkinter.BooleanVar()
            self.hideresultsvar.set(0)
            self.hideshow = ttk.Checkbutton(self.searchbuttonsframe,text='Esconder termos sem resultados', command=lambda:self.showhideresults(), \
                                            image=global_settings.showhide, compound='right', variable=self.hideresultsvar)
            self.hideshow.grid(column=0, row=2, sticky='ns', padx=10, pady=2, columnspan=3)            
            self.collapse = tkinter.Button(self.searchbuttonsframe, font=global_settings.Font_tuple_Arial_10, text='Colapsar todos', \
                                           image=global_settings.collapseimg, compound="right", command=self.collapseall)
            self.collapse.grid(row=3, column=1,sticky='ns', padx=10, pady=2)
            
            self.filter_docs_search = tkinter.Button(self.searchbuttonsframe, font=global_settings.Font_tuple_Arial_10, text='', \
                                           image=global_settings.filterimage, compound="right", command=self.filter_eqs)
            self.filter_docs_search.grid(row=3, column=2,sticky='e', padx=10, pady=2)
            self.filter_docs_search_text = tkinter.Label(self.searchbuttonsframe, font=global_settings.Font_tuple_Arial_10, text="Filtro Aplicado", fg='red')
            self.filter_docs_search_text.grid(row=4, column=2,sticky='e')
            self.filter_docs_search_text.grid_forget() 
            
            bprevfind = tkinter.Button(self.searchbuttonsframe, image=global_settings.imprevfind, font=global_settings.Font_tuple_Arial_10, text="", compound='right', command=lambda: self.iterateSearchList(None, 'anterior'))
            bprevfind.image = global_settings.imprevfind
            bprevfind.grid(column=0, row=0, sticky='ns', padx=10, pady=5, rowspan=2)  
            self.searchtreeframe = tkinter.Frame(self.searchEnv, borderwidth=2, relief='ridge')
            self.searchtreeframe.grid(row=1, column=0, sticky='nsew')
            self.searchtreeframe.rowconfigure(0, weight=1)
            self.searchtreeframe.columnconfigure(0, weight=1)
            self.treeviewSearches = ttk.Treeview(self.searchtreeframe, selectmode='extended')
            self.treeviewSearches.bind_class('post-tree-bind-search', '<F2>', lambda e: self.iterateSearchList(None, 'anterior'))
            self.treeviewSearches.bind_class('post-tree-bind-search', '<F3>', lambda e: self.iterateSearchList(None, 'proximo'))
            self.treeviewSearches.bindtags(('.self.treeviewEqs', 'Treeview', 'post-tree-bind-search','.','all'))
            self.treeviewSearches.bind_class('post-tree-bind-search', "<1>", lambda e: self.treeview_selection_search(e))
            self.treeviewSearches.bind_class('post-tree-bind-search', "<3>", self.treeview_search_right)
            self.treeviewSearches.bind_class('post-tree-bind-search','<Up>', lambda e: self.treeview_selection_search())
            self.treeviewSearches.bind_class('post-tree-bind-search','<Down>', lambda e: self.treeview_selection_search())
            treeviewsearchtt = classes_general.CreateToolTip(self.treeviewSearches, "Buscas", istreeview=True, classe='post-tree-bind-search')
            self.treeviewSearches.heading("#0", text="Resultados", anchor="w")
            self.treeviewSearches.column("#0", width=200, stretch=True, minwidth=200, anchor="w")
            self.treeviewSearches.grid(row=0, column=0, sticky='nsew')
            self.scrolltreeviewSearches = ttk.Scrollbar(self.searchtreeframe, orient="vertical")
            self.scrolltreeviewSearches.grid(row=0, column=1, sticky='ns')
            self.scrolltreeviewSearches.config( command = self.treeviewSearches.yview )
            self.treeviewSearches.tag_configure('termosearch', foreground="#000000", background='#d0d0d0', font=global_settings.Font_tuple_ArialBoldUnderline_10)
            self.treeviewSearches.tag_configure('termosearching', foreground='#d0d0d0', font=global_settings.Font_tuple_ArialBold_10)
            self.treeviewSearches.tag_configure('termoinprocess', background='#7dffd8', font=global_settings.Font_tuple_ArialBold_10)
            self.treeviewSearches.tag_configure('termowitherror', background='#db041a', font=global_settings.Font_tuple_ArialBold_10)
            self.treeviewSearches.tag_configure('resultsearch',font=global_settings.Font_tuple_Arial_8)
            self.treeviewSearches.tag_configure('relsearch', background='#f0f0f0',font=global_settings.Font_tuple_Arial_8)
            self.treeviewSearches.tag_configure('relsearchtoc',font=global_settings.Font_tuple_Arial_8)
            self.treeviewSearches.configure(yscrollcommand=self.scrolltreeviewSearches.set)
            self.scrolltreeviewSearchesH = ttk.Scrollbar(self.searchtreeframe, orient="horizontal")
            self.scrolltreeviewSearchesH.config(command = self.treeviewSearches.xview )
            self.treeviewSearches.configure(xscrollcommand=self.scrolltreeviewSearchesH.set)
            self.scrolltreeviewSearchesH.grid(row=1, column=0, sticky='ew')
            self.obsFrame = tkinter.Frame(borderwidth=2, bg='white')
            self.obsFrame.rowconfigure(1, weight=1)
            self.obsFrame.columnconfigure(0, weight=1)
            self.treeviewSearches.bind_class('post-tree-bind-search', '<Delete>', lambda e: self.deleteSearchDel(e))            
            self.obsButtonFrame = tkinter.Frame(self.obsFrame, borderwidth=2, relief='ridge', pady=5)
            self.obsButtonFrame.rowconfigure((0,1,2), weight=1)
            self.obsButtonFrame.columnconfigure((0,1), weight=1)
            self.obsButtonFrame.grid(row=0, column=0, sticky='nsew')
            self.menuaddcat = tkinter.Button(self.obsButtonFrame, font=global_settings.Font_tuple_Arial_10, \
                                             text="Adicionar Categoria  ", image=global_settings.addcat, compound='right', command=lambda: self.add_edit_category('add',''))
            self.menuaddcat.grid(row=0, column=0, columnspan=2, sticky='ns')    
            self.sep1 = ttk.Separator(self.obsButtonFrame,orient='horizontal')
            self.sep1.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(5,0)) 
            self.collapseobs = tkinter.Button(self.obsButtonFrame, font=global_settings.Font_tuple_Arial_10, text='Colapsar todos', image=global_settings.collapseimg, compound="right", command=self.collapseallobs)
            self.collapseobs.grid(row=2, column=0, columnspan=2, sticky='ns', pady=(5,5))   
            self.TreeviewobsFrame = tkinter.Frame(self.obsFrame, borderwidth=2, bg='white', relief='ridge')
            self.TreeviewobsFrame.rowconfigure(0, weight=1)
            self.TreeviewobsFrame.columnconfigure(0, weight=1)
            self.TreeviewobsFrame.grid(row=1, column=0, sticky='nsew')
            self.treeviewObs = ttk.Treeview(self.TreeviewobsFrame, selectmode='extended')
            self.treeviewObs.heading("#0", text="Categorias", anchor="w")
            self.treeviewObs.bindtags(('.self.treeviewObs', 'Treeview', 'post-tree-bind-obs','.','all'))
            self.treeviewObs.bind_class('post-tree-bind-obs', "<3>", self.treeview_obs_right)
            self.treeviewObs.bind_class('post-tree-bind-obs', "<1>", lambda e: self.treeview_selection_obs(e))
            self.treeviewObs.bind_class('post-tree-bind-obs','<Up>', lambda e: self.treeview_selection_obs(e))
            self.treeviewObs.bind_class('post-tree-bind-obs','<Down>', lambda e: self.treeview_selection_obs(e))
            treeviewobstt = classes_general.CreateToolTip(self.treeviewObs, "Observações", istreeview=True, classe='post-tree-bind-obs')
            self.treeviewObs.grid(row=0, column=0, sticky='nsew')
            self.scrolltreeviewObs = ttk.Scrollbar(self.TreeviewobsFrame, orient="vertical")
            self.scrolltreeviewObs.grid(row=0, column=1, sticky='ns')
            self.scrolltreeviewObs.config( command = self.treeviewObs.yview )
            self.treeviewObs.configure(yscrollcommand=self.scrolltreeviewObs.set)
            self.scrolltreeviewObsH = ttk.Scrollbar(self.TreeviewobsFrame, orient="horizontal")
            self.scrolltreeviewObsH.config( command = self.treeviewObs.xview )
            self.treeviewObs.configure(xscrollcommand=self.scrolltreeviewObsH.set)
            self.scrolltreeviewObsH.grid(row=1, column=0, sticky='ew')
            #self.commenticonb = b'iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAdElEQVRIie3UsQnAMAxE0T9jspA3yhYZKimSyiCMg4h91/lAjQs/JIQATuAR1kET5ee1uoAiKWAZVwQs4+oBiixAB4zu/+rgfweft2QW6J0HKZA+OoCRuoAtA2av6R0Q5TQAKAHZHUCLWICI2IAWsaUi1pQXDcnofAiAy1cAAAAASUVORK5CYII='
            self.notebook.add(self.tocFrame, text="Relatorios", sticky='nsew', image=global_settings.repicon, compound='top')
            self.notebook.add(self.searchFrame, text="Buscas", sticky='nsew', image=global_settings.searchicon, compound='top')
            self.notebook.add(self.obsFrame, text="Marcadores", sticky='nsew', image=global_settings.commenticon, compound='top')
            self.notebook.add(self.locFrame, text="Localizações", sticky='nsew', image=global_settings.locationicon, compound='top')
            self.globalFrame.add(self.infoFrame)
            self.treeviewObs.tag_configure('relobs', background='#e3e1e1')
            self.allobs = {}
            self.allobsbyitem = {}
            sqliteconn =  utilities_general.connectDB(str(global_settings.pathdb), 5)
            try:
                cursor = sqliteconn.cursor()   
                check_previous_search =  "SELECT DISTINCT C.termo, C.tipobusca, C.id_termo, C.fixo, C.pesquisado  FROM Anexo_Eletronico_SearchTerms C ORDER by 3"
                cursor.custom_execute(check_previous_search)
                termos = cursor.fetchall()
                cursor.close()
                for termox in termos:
                    tipobusca = termox[1]                    
                    termo = termox[0].strip()
                    idtermo = termox[2]
                    pesquisados = termox[4]
                    global_settings.listaTERMOS[(termo.upper(),tipobusca)] = [termo,tipobusca, idtermo, pesquisados]
                    pedidosearch = classes_general.PedidoSearch(int(idtermo)*-1, tipobusca, termo)
                    global_settings.searchqueue.put(pedidosearch)
                    #global_settings.searchqueue.append((termo, tipobusca, None))
                    global_settings.contador_buscas_incr += 1
                check_previous_obscat =  "SELECT C.obscat, C.id_obscat, C.fixo, C.ordem FROM Anexo_Eletronico_Obscat C ORDER BY 4"
                cursor = sqliteconn.cursor()
                cursor.custom_execute("PRAGMA foreign_keys = ON")
                cursor.custom_execute(check_previous_obscat)
                obscats = cursor.fetchall()
                cursor.close()
                for obscat in obscats:
                    self.treeviewObs.insert(parent='', index='end', iid=str(obscat[1]), text=obscat[0].upper(), values=(str(obscat[2]), obscat[1], obscat[0],), \
                                            image=global_settings.catimage, tag='obscat')
                    self.treeviewObs.tag_configure('obscat', background='#a1a1a1', font=global_settings.Font_tuple_ArialBoldUnderline_12)
                    check_previous_obsitens =  '''SELECT P.rel_path_pdf, O.paginainit, O.p0x, O.p0y, O.paginafim, O.p1x, O.p1y, O.tipo, O.id_obs, O.fixo, O.status, 
                    O.conteudo, O.arquivo, O.withalt FROM Anexo_Eletronico_Obsitens O, 
                    Anexo_Eletronico_Pdfs P  WHERE
                        O.id_pdf  = P.id_pdf AND
                        O.id_obscat = ? ORDER BY 2,9'''
                    cursor = sqliteconn.cursor() 
                    cursor.custom_execute(check_previous_obsitens, (obscat[1],))
                    obsitens = cursor.fetchall()
                    for obsitem in obsitens:
                        check_previous_annots =  '''SELECT A.id_annot, A.paginainit, A.p0x, A.p0y, A.paginafim, A.p1x, A.p1y, A.link, A.conteudo
                        FROM Anexo_Eletronico_Annotations A WHERE
                            A.id_obs  = ? ORDER BY 1'''
                        cursor.custom_execute(check_previous_annots, (obsitem[8],))
                        annots = cursor.fetchall()
                        annotation_list = {}
                        relpath = obsitem[0]
                        basepdf = utilities_general.get_normalized_path(os.path.join(global_settings.pathdb.parent, relpath))
                        pathpdf = utilities_general.get_normalized_path(basepdf)
                        for annot in annots:
                            annots_object = classes_general.Annotation(pathpdf, annot[0], annot[1], annot[4], annot[2], annot[3],\
                                                                       annot[5], annot[6], obsitem[8], annot[8], annot[7])
                            annotation_list[annot[0]] = (annots_object)
                        paginainit = obsitem[1]
                        p0x = obsitem[2]
                        p0y = obsitem[3]
                        paginafim = obsitem[4]
                        p1x = obsitem[5]
                        p1y = obsitem[6]
                        tipo = obsitem[7]                       
                        status = obsitem[10]
                        conteudo = obsitem[11]
                        arquivo = obsitem[12]
                        withalt = True if obsitem[13] == 1 else False
                        ident = ' '
                        
                        if(pathpdf in global_settings.infoLaudo and pathpdf not in self.allobs):
                            self.allobs[pathpdf] = []
                        
                        obsobject = classes_general.Observation(paginainit, paginafim, p0x, p0y, p1x, p1y, tipo, pathpdf, obsitem[8], \
                                                                obscat[1], conteudo, annotation_list, withalt)
                        self.allobs[pathpdf].append(obsobject)
                        self.allobsbyitem['obsitem'+str(obsitem[8])] = obsobject
                        tocpdf = global_settings.infoLaudo[pathpdf].toc
                        #treeobs_insert(self, tipo, fixo, paginainit, paginafim, basepdf, p0x, p0y, p1x, p1y, tocpdf, idobscat, idobsitem, ident)
                        self.treeobs_insert(tipo, 1, paginainit, paginafim, pathpdf, p0x, p0y, p1x, p1y, tocpdf, obscat[1], obsitem[8],ident = ' ', tag_alt=status+str(obsitem[8]))
                        if(status=='alterado'):
                            self.treeviewObs.tag_configure(status+str(obsitem[8]), background='#ff4747')
                            
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)
            finally:
                try:
                    cursor.close() 
                except Exception as ex:
                    None
                try:
                    sqliteconn.close()
                except Exception as ex:
                    None
                self.treeviewObs.tag_configure('relobs', background='#e3e1e1')
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
            
    def filter_eqs(self, event=None):
        filtro_window = classes_general.Filter_Window(self.treeviewEqs, self.treeviewSearches, self)
    
    def collapseallobs(self, event=None):
        for child in self.treeviewObs.get_children(''): 
            self.treeviewObs.item(child, open=False)

    def collapsealleqs(self, event=None):
        for child in self.treeviewEqs.get_children(''):
            for child2 in self.treeviewEqs.get_children(child): 
                self.treeviewEqs.item(child2, open=False)
                
    def expandalleqs(self, event=None):
        for child in self.treeviewEqs.get_children(''):
            for child2 in self.treeviewEqs.get_children(child): 
                self.treeviewEqs.item(child2, open=True)
    
    def collapseall(self, event=None):
        for child in self.treeviewSearches.get_children(''):
            self.treeviewSearches.item(child, open=False)   

    def deleteSearchDel(self, event=None):
        iids = self.treeviewSearches.selection()
        if(len(iids)==1):
            
            self.treeviewSearches.selection_set(iids[0])
            if(self.treeviewSearches.parent(iids)=='' and self.treeviewSearches.item(iids[0], 'text') != ''):
                nxt = self.treeviewSearches.next(iids[0])
                prev = self.treeviewSearches.prev(iids[0])
                try:
                    if(isinstance(event.widget, ttk.Treeview)):
                        self.exclude_search(event)                        
                        if(nxt!=''):
                            self.treeviewSearches.selection_set(nxt)
                            self.treeviewSearches.focus(nxt)
                        elif(prev!=''):
                            self.treeviewSearches.selection_set(prev)
                            self.treeviewSearches.focus(prev)
                except Exception as ex:
                    utilities_general.printlogexception(ex=ex) 
        elif(len(iids)>1):           
            #lista = []
            #for item in iids:
            #    if(self.treeviewSearches.parent(item)=='' and self.treeviewSearches.item(item, 'text') != ''):
            #        lista.append(item)
            self.exclude_search(event)
            nxt = self.treeviewSearches.next(iids[0])
            prev = self.treeviewSearches.prev(iids[0])
            try:
                if(isinstance(event.widget, ttk.Treeview)):
                    if(nxt!=''):
                        self.treeviewSearches.selection_set(nxt)
                        self.treeviewSearches.focus(nxt)
                    elif(prev!=''):
                        self.treeviewSearches.selection_set(prev)
                        self.treeviewSearches.focus(prev)
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)
                
    def openSearchlist(self):        
        searchlist = None
        searchlist = Path(askopenfilename(filetypes=(("Texto", "*.txt"), ("Todos os arquivos", "*"))))
        if(searchlist!=None and searchlist!=''):
            with open(searchlist, "r", encoding='utf-8') as a_file:
                try:
                    for line in a_file: 
                        #print(line)
                        stripped_line = line.strip()
                        if(stripped_line[0]=="#"):
                            continue
                        tipo = stripped_line.split(" ")[0].strip()
                        termo = stripped_line[len(tipo):len(stripped_line)].strip()
                        
                         
                        if(len(stripped_line)==1):
                            termo = stripped_line
                            tipo = "LIKE"
                            if((termo.upper(), tipo.upper()) in global_settings.listaTERMOS):
                                continue 
                            if(len(termo)<3):
                                utilities_general.popup_window(f'BUSCA - LIKE : O tamanho da palavra <{termo}> deve possuir ao menos 3 caractéres!', False)
                                continue
                            global_settings.listaTERMOS[(termo.upper(),tipo.upper())] = [] 
                            pedidosearch = classes_general.PedidoSearch(global_settings.contador_buscas_incr*-1, "LIKE", termo)
                            global_settings.searchqueue.put(pedidosearch)
                            global_settings.contador_buscas_incr += 1
                        else:
                            
                                                                                  
                            if("MATCH" == tipo.upper()):
                                if((termo.upper(), tipo.upper()) in global_settings.listaTERMOS):
                                    continue 
                                global_settings.listaTERMOS[(termo.upper(),tipo.upper())] = []  
                                pedidosearch = classes_general.PedidoSearch(global_settings.contador_buscas_incr*-1, "MATCH", termo)
                                global_settings.searchqueue.put(pedidosearch)
                                global_settings.contador_buscas_incr += 1
                            elif("REGEX" == tipo.upper()):
                                if((termo.upper(), tipo.upper()) in global_settings.listaTERMOS):
                                    continue 
                                global_settings.listaTERMOS[(termo.upper(),tipo.upper())] = []  
                                pedidosearch = classes_general.PedidoSearch(global_settings.contador_buscas_incr*-1, "REGEX", termo)
                                global_settings.searchqueue.put(pedidosearch)
                                global_settings.contador_buscas_incr += 1
                            elif("LIKE" == tipo.upper()):
                                
                                if((termo.upper(), tipo.upper()) in global_settings.listaTERMOS):
                                    continue 
                                if(len(termo)<3):
                                    utilities_general.popup_window(f'BUSCA - LIKE : O tamanho da palavra <{termo}> deve possuir ao menos 3 caractéres!', False)
                                    continue
                                global_settings.listaTERMOS[(termo.upper(),tipo.upper())] = []  
                                pedidosearch = classes_general.PedidoSearch(global_settings.contador_buscas_incr*-1, "LIKE", termo)
                                global_settings.searchqueue.put(pedidosearch)
                                global_settings.contador_buscas_incr += 1
                            else:
                                termo = stripped_line
                                tipo = "LIKE"
                                if((termo.upper(), tipo.upper()) in global_settings.listaTERMOS):
                                    continue 
                                if(len(termo)<3):
                                    utilities_general.popup_window(f'BUSCA - LIKE : O tamanho da palavra <{termo}> deve possuir ao menos 3 caractéres!', False)
                                    continue
                                global_settings.listaTERMOS[(termo.upper(),tipo.upper())] = []  
                                pedidosearch = classes_general.PedidoSearch(global_settings.contador_buscas_incr*-1, "LIKE", termo)
                                global_settings.searchqueue.put(pedidosearch)
                                global_settings.contador_buscas_incr += 1
                    #self.uniquesearchprocess2.start() 
                    #self.primeiroresetbuscar = True
                except Exception as ex:
                    #print(ex)
                    utilities_general.printlogexception(ex=ex)  
            
    def iterateSearchList(self, event=None, tipo=None):
        try:
            for pdf in global_settings.infoLaudo:
                global_settings.infoLaudo[pdf].retangulosDesenhados = {}            
            mudar=''
            if(tipo=='proximo'):
                mudar = self.treeviewSearches.next(self.treeviewSearches.selection()[0])
            elif(tipo=='anterior'):
                mudar = self.treeviewSearches.prev(self.treeviewSearches.selection()[0])
            if(mudar==''):
                paiultimo = self.treeviewSearches.parent(self.treeviewSearches.selection()[0])
                if(tipo=='proximo'):
                    proximopai =  self.treeviewSearches.next(paiultimo)
                elif(tipo=='anterior'):
                    proximopai =  self.treeviewSearches.prev(paiultimo)                
                if(proximopai==''):
                    return                
                if(tipo=='proximo'):
                    mudar =  self.treeviewSearches.get_children(proximopai)[0]
                elif(tipo=='anterior'):
                    mudar =  self.treeviewSearches.get_children(proximopai)[-1]
            self.treeviewSearches.see(mudar)
            self.treeviewSearches.selection_set(mudar)
            self.treeview_selection_search()            
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)            
            
    def treeview_selection_search(self, event=None):        
        try:
            region = ""
            if(event!=None):
                region = self.treeviewSearches.identify("region", event.x, event.y)
            if region == "heading":
                self.orderpopup(event)
            else:
                for pdf in global_settings.infoLaudo:
                    global_settings.infoLaudo[pdf].retangulosDesenhados = {}
                if(event!=None):
                    searchresultiid = self.treeviewSearches.identify_row(event.y)
                else:
                    searchresultiid = self.treeviewSearches.selection()[0]
                try:
                    resultsearch = global_settings.searchResultsDict[searchresultiid]
                except Exception as ex:
                    return                
                parent = self.treeviewSearches.parent(searchresultiid)
                if(len(self.treeviewSearches.get_children(searchresultiid))==0 and parent != ''):                      
                    raiz = self.treeviewSearches.parent(searchresultiid)                    
                    while(raiz!=''):
                        if(self.treeviewSearches.parent(raiz)==''):
                            break
                        raiz = self.treeviewSearches.parent(raiz)
                    newpath = utilities_general.get_normalized_path(resultsearch.pathpdf)
                    sobraEspaco = 0
                    if(self.docFrame.winfo_width() > global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw * self.zoom_x * global_settings.zoom):
                        sobraEspaco = self.docInnerCanvas.winfo_x()  
                    self.maiorw = self.docFrame.winfo_width()
                    self.positions[self.indiceposition] = (global_settings.pathpdfatual, self.vscrollbar.get()[0])
                    self.indiceposition += 1
                    if(self.indiceposition>=10):
                        self.indiceposition = 0
                    if(global_settings.pathpdfatual!=newpath):
                        pdfantigo = global_settings.pathpdfatual
                        self.docInnerCanvas.delete("quad")
                        self.docInnerCanvas.delete("simplesearch")
                        self.docInnerCanvas.delete("obsitem")
                        self.docInnerCanvas.delete("link")
                        self.clearSomeImages(["quad", "simplesearch", "obsitem", "link"])
                        self.docwidth = self.docOuterFrame.winfo_width()
                        global_settings.pathpdfatual = utilities_general.get_normalized_path(newpath)
                        try:
                            global_settings.docatual.close()
                        except Exception as ex:
                            None
                        global_settings.docatual = fitz.open(global_settings.pathpdfatual)
                        self.labeldocname.config(font=global_settings.Font_tuple_Arial_10, text=os.path.basename(global_settings.pathpdfatual))
                        for i in range(global_settings.minMaxLabels):
                            global_settings.processed_pages[i] = -1
                        self.totalPgg.config(font=global_settings.Font_tuple_Arial_10, text="/ "+str(global_settings.infoLaudo[global_settings.pathpdfatual].len))
                        if(global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw*self.zoom_x*global_settings.zoom>self.maiorw):
                            self.maiorw = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw*self.zoom_x *global_settings.zoom           
                        self.scrolly = round((global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh*self.zoom_x*global_settings.zoom), 16)*global_settings.infoLaudo[global_settings.pathpdfatual].len  - 35
                        self.docInnerCanvas.config(scrollregion=(sobraEspaco, 0, sobraEspaco + (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw*global_settings.zoom*self.zoom_x), self.scrolly))
                        if(global_settings.infoLaudo[pdfantigo].zoom_pos!=global_settings.infoLaudo[global_settings.pathpdfatual].zoom_pos):
                            #global_settings.root.after(1, lambda: self.zoomx(None, None, pdfantigo))
                            self.zoomx(None, None, pdfantigo)
                        #self.docInnerCanvas.update_idletasks()                    
                    totalhits = self.treeviewSearches.item(parent, 'text').split(' ')
                    self.ocorrenciasLabel.config(font=global_settings.Font_tuple_Arial_10, text=str(self.treeviewSearches.index(searchresultiid)+1) + ' de ' + totalhits[len(totalhits)-1])
                    self.termosearchVar.set(self.treeviewSearches.item(raiz, 'text'))
                    self.docInnerCanvas.delete("simplesearch")
                    self.clearSomeImages(["simplesearch"])
                    pagina = int(resultsearch.pagina)-1
                    if(self.afterpaint!=None):
                        global_settings.root.after_cancel(self.afterpaint)
                    self.window_search_info.window.deiconify()
                    if(pagina in global_settings.processed_pages):
                        listaresultados = [resultsearch]                        
                        self.paintsearchresult(listaresultados)
                    else:
                        ondeir = ((pagina) / (global_settings.infoLaudo[global_settings.pathpdfatual].len))
                        self.docInnerCanvas.yview_moveto(ondeir)
                        if(str(pagina+1)!=self.pagVar.get()):
                            self.pagVar.set(str(pagina+1))
                        listaresultados = [resultsearch]                        
                        self.paintsearchresult(listaresultados)              
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
            
    def paintsearchresult(self, listaresultados, simplesearch=False, first=True):
        if(len(listaresultados)>0):            
            pagina = int(listaresultados[0].pagina)
            if(pagina not in global_settings.processed_pages):
                 ondeir = ((pagina) / (global_settings.infoLaudo[global_settings.pathpdfatual].len))
                 self.docInnerCanvas.yview_moveto(ondeir)
                 if(str(pagina+1)!=self.pagVar.get()):
                     self.pagVar.set(str(pagina+1))
                 self.afterpaint = global_settings.root.after(100, lambda: self.paintsearchresult(listaresultados, simplesearch, first=False))
            
            elif(pagina not in global_settings.infoLaudo[global_settings.pathpdfatual].quadspagina):
                if(first or pagina in global_settings.processed_pages):
                    self.afterpaint = global_settings.root.after(100, lambda: self.paintsearchresult(listaresultados, simplesearch, first=False))
                    
            else:
                self.window_search_info.window.withdraw()
                self.docInnerCanvas.delete("simplesearch")
                self.clearSomeImages(["simplesearch"])
                for resultsearch in reversed(listaresultados):
                    sobraEspaco = 0
                    if(self.docFrame.winfo_width() > global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw * self.zoom_x * global_settings.zoom):
                        sobraEspaco = self.docInnerCanvas.winfo_x() 
                    posicoes = global_settings.infoLaudo[global_settings.pathpdfatual].quadspagina[pagina]
                    init = posicoes[resultsearch.init]
                    fim = posicoes[resultsearch.fim-1]
                    p0x = init[0]
                    p0y = (init[1])+5
                    
                    p1x = fim[2]
                    p1y = (fim[3])
                    if(p0y>=p1y-1):
                        p0y = p1y-4
                    #print(p0x, p0y, p1x, p1y)
                    self.prepararParaQuads(pagina, int(p0x), int(p0y), math.ceil((p1x)), int(p1y), color=self.color, tag=["simplesearch"], apagar=False, enhancetext=True, enhancearea=False, alt=False)
                    atual = ((self.vscrollbar.get()[0]*global_settings.infoLaudo[global_settings.pathpdfatual].len))
                    deslocy = (math.floor(pagina) * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom) + (p0y *  self.zoom_x * global_settings.zoom)                    
                    desloctotalmenor =  (atual * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom) 
                    desloctotalmaior =   desloctotalmenor + self.docFrame.winfo_height() - self.hscrollbar.winfo_height() -  self.labeldocname.winfo_height()
                    if(deslocy < desloctotalmenor or deslocy > desloctotalmaior):
                        ondeir = ((pagina) / (global_settings.infoLaudo[global_settings.pathpdfatual].len)) + (p0y*self.zoom_x*global_settings.zoom-self.docFrame.winfo_height()/2)/self.scrolly
                        self.docInnerCanvas.yview_moveto(ondeir)
                        if(str(pagina+1)!=self.pagVar.get()):
                            self.pagVar.set(str(pagina+1))
                    if(simplesearch):
                        self.simplesearching = False
                        self.nhp.config(relief='raised', state='normal')
                        self.php.config(relief='raised', state='normal')
                        
    def _on_mousewheel(self, event):
        self.docInnerCanvas.yview_scroll(-1*int((event.delta/120)), "units")
        try:
            if (event.num==4):
                 self.docInnerCanvas.yview_scroll(-1, "units")
                 
            elif(event.num==5):
                 self.docInnerCanvas.yview_scroll(1, "units")
        except Exception as ex:
            None
        finally:
            try:
                at = round(self.vscrollbar.get()[0]*global_settings.infoLaudo[global_settings.pathpdfatual].len)
                posicaoRealY0Canvas = self.vscrollbar.get()[0] * self.scrolly + event.y
                posicaoRealX0Canvas = self.hscrollbar.get()[0] * (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw * self.zoom_x * global_settings.zoom) + event.x
                posicaoRealY0 = round((posicaoRealY0Canvas % (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom)) / (self.zoom_x*global_settings.zoom), 0)
                posicaoRealX0 = round(posicaoRealX0Canvas / (self.zoom_x*global_settings.zoom), 0)
                pagina = math.floor(posicaoRealY0Canvas / (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh *self.zoom_x*global_settings.zoom)) + 1
                self.labelmousepos.config(font=global_settings.Font_tuple_Arial_10, text="(Pagina:{},X:{},Y:{})".format(pagina, round(posicaoRealX0), round(posicaoRealY0)))
                if(self.initialPos!=None):
                    global_settings.root.after(10, lambda e=event:self._selectingText(e))
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)
         
    def addImagetoList(self, tag, image, obsitem=None):        
        if(not tag in self.allimages):
            self.allimages[tag] = []
        if(obsitem!=None):
            self.allimages[tag].append((image, obsitem))
        else:            
            self.allimages[tag].append(image)        
        
    def clearEnhanceObs(self):
        apagar = []
        self.docInnerCanvas.delete('note')
        for tag in self.allimages:            
            if "enhanceobs" in tag:                
                self.docInnerCanvas.delete(tag)
                apagar.append(tag)
            elif "enhanceannot" in tag:
                self.docInnerCanvas.delete(tag)
                apagar.append(tag)
        for tag in apagar:
            if('note' in tag):
                del self.allimages[tag]
                
    def clearAllImages(self):        
        for tag in self.allimages:
            self.docInnerCanvas.delete(tag)
        self.allimages = {}
    def clearSomeImages(self, listatags):
        for tag in listatags:
            if tag in self.allimages:
                del self.allimages[tag] 
            #self.docInnerCanvas.delete(tag)
    
    def hideprevioussearcheswindow(self):
        try:
            self.previousSearchesWindow.withdraw()
        except:
            None
            
    def clearSelectedTextByCLick(self, tipo, event):       
       try:
           
           if(event.widget!=None):
               try:
                   event.widget.focus_set()  
                   #if(not isinstance(event.widget, tkinter.Button)):
                   #    print(event.widget)
                       #self.hideprevioussearcheswindow()
                   #else:
                   #    N
               except:
                   None
           
           if(isinstance(event.widget, classes_general.CustomCanvas) or isinstance(event.widget, classes_general.CustomFrame)):
               self.hideprevioussearcheswindow()
               self.docInnerCanvas.focus_set()
               posicaoRealY0Canvas = self.vscrollbar.get()[0] * (self.scrolly) + event.y
               posicaoRealX0Canvas = self.hscrollbar.get()[0] * (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw * self.zoom_x * global_settings.zoom) + event.x
               posicaoRealY0 = (posicaoRealY0Canvas % (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom)) / (self.zoom_x*global_settings.zoom)
               posicaoRealX0 = posicaoRealX0Canvas / (self.zoom_x*global_settings.zoom)    
               if(self.selectionActive):
                   if(tipo=="press"): 
                       
                       for pdf in global_settings.infoLaudo:
                           global_settings.infoLaudo[pdf].retangulosDesenhados = {}                           
                       self.docInnerCanvas.delete("simplesearch")
                       self.docInnerCanvas.delete("quad")
                       self.docInnerCanvas.delete("obsitem")
                       self.clearSomeImages(["simplesearch", "quad", "obsitem"])                       
                       posicaoRealY0Canvas = self.docInnerCanvas.canvasy(event.y)
                       posicaoRealX0Canvas = self.docInnerCanvas.canvasx(event.x)                       
                       pagina = math.floor(posicaoRealY0Canvas / (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh *self.zoom_x*global_settings.zoom))
                       self.initialPos = (posicaoRealX0Canvas, posicaoRealY0Canvas, posicaoRealX0, posicaoRealY0, pagina)
                   elif(tipo=="release"):
                       self.paginaSearchSimple = -1
                       self.initialPos = None
               elif(self.areaselectionActive):
                   if(tipo=="press"):
                       for pdf in global_settings.infoLaudo:
                           global_settings.infoLaudo[pdf].retangulosDesenhados = {}                           
                       self.docInnerCanvas.delete("simplesearch")
                       self.docInnerCanvas.delete("quad")
                       self.docInnerCanvas.delete("obsitem")
                       self.clearSomeImages(["simplesearch", "quad", "obsitem"])
                       pagina = math.floor(posicaoRealY0Canvas / (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh *self.zoom_x*global_settings.zoom))
                       self.initialPos = (posicaoRealX0Canvas, posicaoRealY0Canvas, posicaoRealX0, posicaoRealY0, pagina)
                   elif(tipo=="release"):
                       self.paginaSearchSimple = -1
                       self.initialPos = None
               else:
                   if(tipo=="press"):
                       self.docInnerCanvas.scan_mark(event.x, event.y)
 
                       pagina = math.floor(posicaoRealY0Canvas / (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh *self.zoom_x*global_settings.zoom))
                       self.initialPos = (posicaoRealX0Canvas, posicaoRealY0Canvas, posicaoRealX0, posicaoRealY0, pagina)
                   elif(tipo=="release"):
                       self.paginaSearchSimple = -1
                       if(posicaoRealX0Canvas == self.initialPos[0] and posicaoRealY0Canvas == self.initialPos[1]): 
                            linkcustom = False
                            self.initialPos = None
                            
                            if(not linkcustom):
                                pagina = math.floor(posicaoRealY0Canvas / (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh *self.zoom_x*global_settings.zoom))
                                for link in global_settings.infoLaudo[global_settings.pathpdfatual].links[pagina]:
                                     r = link['from']                                     
                                     if(posicaoRealX0 >= r.x0 and posicaoRealX0 <= r.x1 and posicaoRealY0 >= r.y0 and posicaoRealY0 <= r.y1):
                                         if('page' in link):
                                             pageint = int(link['page'])
                                             to = link['to']                                             
                                             if(pageint > 0 and pageint<=global_settings.infoLaudo[global_settings.pathpdfatual].len):
                                                 ondeir = (pageint) / global_settings.infoLaudo[global_settings.pathpdfatual].len + (to.y / (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh*global_settings.infoLaudo[global_settings.pathpdfatual].len))
                                                 self.positions[self.indiceposition] = (global_settings.pathpdfatual, self.vscrollbar.get()[0])
                                                 self.indiceposition += 1
                                                 if(self.indiceposition>=10):
                                                     self.indiceposition = 0
                                                 self.docInnerCanvas.yview_moveto(ondeir)
                                                 if(str(pageint+1)!=self.pagVar.get()):
                                                     self.pagVar.set(str(pageint+1))                                                 
                                             else:
                                                 atual = round((self.vscrollbar.get()[0]*global_settings.infoLaudo[global_settings.pathpdfatual].len))                                                 
                                         elif('file' in link):
                                             arquivo = link['file']
                                             if(arquivo==""):
                                                xref = link['xref']
                                                info = global_settings.docatual.xref_get_key(xref, 'A')
                                                grupos_search = global_settings.regex_actions_compiled.search(info[1])
                                                if(grupos_search==None):
                                                    continue
                                                grupos = grupos_search.groups()
                                                arquivo = grupos[2]
                                                if(grupos[-1]=="Launch"):
                                                    arquivo = grupos[2]
                                                elif(grupos[-1]=="GoToR"):
                                                    arquivo = f"{grupos[2]}#{grupos[1]}"
                                             arquivosplit = arquivo.split("#")
                                             arquivo = utilities_general.get_normalized_path(arquivosplit[0])
                                             if(len(arquivosplit)>1):
                                                 arquivo = utilities_general.get_normalized_path(os.path.join(Path(global_settings.pathpdfatual).parent, arquivosplit[0]))
                                                 aprocurar = arquivosplit[1]
                                                 if(arquivo in global_settings.infoLaudo):
                                                     texto = ""
                                                     if("mm.chat" in aprocurar):
                                                         recttext = fitz.Rect(80, r.y0-10, 148, r.y1+20)
                                                         texto = global_settings.docatual[pagina].get_textbox(recttext)
                                                         texto = texto.replace("\n", " ")                                                         
                                                     if(arquivo!=global_settings.pathpdfatual):
                                                         try:
                                                             global_settings.docatual.close()
                                                         except Exception as ex:
                                                             None
                                                         global_settings.docatual=fitz.open(arquivo)
                                                     retorno = utilities_general.processDocXREF(arquivo, global_settings.docatual, aprocurar)    
                                                     if(retorno!=None):                                                 
                                                         to = retorno[3]
                                                         page_dest = int(retorno[1])
                                                         pdfantigo = global_settings.pathpdfatual
                                                         if(arquivo!=global_settings.pathpdfatual):
                                                             
                                                             sobraEspaco = 0
                                                             if(self.docFrame.winfo_width() > global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw * self.zoom_x * global_settings.zoom):
                                                                 sobraEspaco = self.docInnerCanvas.winfo_x()  
                                                             self.docwidth = self.docOuterFrame.winfo_width()
                                                             
                                                             self.clearAllImages()
                                                             for i in range(global_settings.minMaxLabels):
                                                                 global_settings.processed_pages[i] = -1
                                                             self.positions[self.indiceposition] = (global_settings.pathpdfatual, self.vscrollbar.get()[0])
                                                             global_settings.pathpdfatual = utilities_general.get_normalized_path(arquivo)
                                                             self.labeldocname.config(font=global_settings.Font_tuple_Arial_10, text=os.path.basename(global_settings.pathpdfatual))
                                                             self.totalPgg.config(font=global_settings.Font_tuple_Arial_10, text="/ "+str(global_settings.infoLaudo[global_settings.pathpdfatual].len))                    
                                                             if(global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw*self.zoom_x*global_settings.zoom>self.maiorw):
                                                                 self.maiorw = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw*self.zoom_x *global_settings.zoom           
                                                             self.scrolly = round((global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh*self.zoom_x*global_settings.zoom), 16)*global_settings.infoLaudo[global_settings.pathpdfatual].len  - 35
                                                             self.docInnerCanvas.config(scrollregion=(sobraEspaco, 0, sobraEspaco+ (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw*global_settings.zoom*self.zoom_x), self.scrolly))                
                                                             self.treeviewEqs.selection_set(global_settings.pathpdfatual)
                                                         else:
                                                             self.positions[self.indiceposition] = (global_settings.pathpdfatual, self.vscrollbar.get()[0])
                                                         ondeir = (page_dest) / global_settings.infoLaudo[global_settings.pathpdfatual].len + (to / (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh*global_settings.infoLaudo[global_settings.pathpdfatual].len))
                                                         #if()
                                                         #self.positions[self.indiceposition] = (global_settings.pathpdfatual, self.vscrollbar.get()[0])
                                                         self.indiceposition += 1
                                                         if(self.indiceposition>=10):
                                                             self.indiceposition = 0
                                                         if(global_settings.infoLaudo[pdfantigo].zoom_pos!=global_settings.infoLaudo[global_settings.pathpdfatual].zoom_pos):
                                                             #global_settings.root.after(1, lambda: self.zoomx(None, None, pdfantigo))
                                                             self.zoomx(None, None, pdfantigo)
                                                         self.docInnerCanvas.yview_moveto(ondeir)
                                                         #global_settings.root.after(1, lambda: self.docInnerCanvas.yview_moveto(ondeir))
                                                         
                                                         if(str(page_dest+1)!=self.pagVar.get()):
                                                             self.pagVar.set(str(page_dest+1))
                                                         if("mm.chat" in aprocurar and not texto==""):
                                                             texto = texto.strip()
                                                             regexdatahora = "([0-9]{2}\/[0-9]{2}\/[0-9]{4})\s+([0-9]{2}\:[0-9]{2}\:[0-9]{2})"
                                                             datahora = re.findall(regexdatahora, texto)
                                                             texto = "["+datahora[0][0]+" "+datahora[0][1]+"]"
                                                             global_settings.root.after(100, lambda: self.dosearchsimple('next', termo=texto))
                                                         
                                                             
                                                 else:
                                                     if plt == "Linux":
                                                         arquivo = str(arquivo).replace("\\","/")
                                                         pdfatualnorm = str(global_settings.pathpdfatual).replace("\\","/")
                                                     elif plt=="Windows":
                                                         arquivo = str(arquivo).replace("/","\\")
                                                         pdfatualnorm = str(global_settings.pathpdfatual).replace("/","\\")
                                                     
                                                     filepath = str(Path(utilities_general.get_normalized_path(os.path.join(Path(utilities_general.get_normalized_path(pdfatualnorm)).parent,arquivo))))
                                                     try:
                                                         if(not os.path.exists(filepath)):
                                                             utilities_general.popup_window(f'O arquivo não selecionado não existe, favor verifique: \n<{filepath}>', False)
                                                         
                                                         elif platform.system() == 'Darwin':       # macOS
                                                             subprocess.call(('open', filepath), shell=True)
                                                         elif platform.system() == 'Windows':    # Windows
                                                             os.startfile(filepath)
                                                         else:    
                                                             openfile = ['xdg-open', filepath]
                                                             try:                                                                 
                                                                 myenv = dict(os.environ)  # make a copy of the environment
                                                                 HOME = os.path.expanduser("~")    
                                                                 # Single directory where user-specific data files should be written
                                                                 XDG_DATA_HOME = os.environ.get("XDG_DATA_HOME", os.path.join(HOME, ".local", "share"))                                                                
                                                                 # Single directory where user-specific configuration files should be written
                                                                 XDG_CONFIG_HOME = os.environ.get("XDG_CONFIG_HOME", os.path.join(HOME, ".config"))                                                                
                                                                 # List of directories where data files should be searched.
                                                                 XDG_DATA_DIRS_LIST = [XDG_DATA_HOME] + "/usr/local/share:/usr/share".split(":")
                                                                 XDG_DATA_DIRS = ':'.join((t) for t in XDG_DATA_DIRS_LIST)
                                                                 # List of directories where configuration files should be searched.
                                                                 XDG_CONFIG_DIRS_LIST = [XDG_CONFIG_HOME] + "/etc/xdg".split(":")
                                                                 XDG_CONFIG_DIRS = ':'.join((t) for t in XDG_CONFIG_DIRS_LIST)
                                                                 #lp_key = 'LD_LIBRARY_PATH'  # for GNU/Linux and *BSD.
                                                                 myenv['XDG_DATA_HOME'] = XDG_DATA_HOME
                                                                 myenv['XDG_CONFIG_HOME'] = XDG_CONFIG_HOME
                                                                 myenv['XDG_DATA_DIRS'] = XDG_DATA_DIRS
                                                                 myenv['XDG_CONFIG_DIRS'] = XDG_CONFIG_DIRS                                                                 
                                                                 capture_output = subprocess.run(openfile, check=True, env=myenv, capture_output=True)
                                                                 #print(capture_output)
                                                                 utilities_general.printlogexception(ex=capture_output)
                                                             except Exception as ex:
                                                                 None
                                                                 utilities_general.printlogexception(ex=ex)
                                                                 try:
                                                                     webbrowser.open_new_tab(filepath)
                                                                 except:
                                                                     utilities_general.popup_window('O arquivo não possui um \nprograma associado para abertura!', False)
                                                     except Exception as ex:
                                                         utilities_general.printlogexception(ex=ex)
                                                         utilities_general.popup_window('O arquivo não possui um \nprograma associado para abertura!', False)                                                              
                                             else:
                                                 if plt == "Linux":
                                                     arquivo = str(arquivo).replace("\\","/")
                                                     pdfatualnorm = str(global_settings.pathpdfatual).replace("\\","/")
                                                 elif plt=="Windows":
                                                     arquivo = str(arquivo).replace("/","\\")
                                                     pdfatualnorm = str(global_settings.pathpdfatual).replace("/","\\")                                                 
                                                 filepath = str(Path(utilities_general.get_normalized_path(os.path.join(Path(utilities_general.get_normalized_path(pdfatualnorm)).parent,arquivo))))
                                                 try:
                                                     if(not os.path.exists(filepath)):
                                                         utilities_general.popup_window(f'O arquivo não selecionado não existe, favor verifique: \n<{filepath}>', False)
                                                     elif platform.system() == 'Darwin':       # macOS
                                                         subprocess.call(('open', filepath), shell=True)
                                                     elif platform.system() == 'Windows':    # Windows
                                                         os.startfile(filepath)
                                                     else:  
                                                         openfile = ['xdg-open', filepath]
                                                         try:                                                             
                                                             myenv = dict(os.environ)  # make a copy of the environment
                                                             HOME = os.path.expanduser("~")
                                                             # Single directory where user-specific data files should be written
                                                             XDG_DATA_HOME = os.environ.get("XDG_DATA_HOME", os.path.join(HOME, ".local", "share"))                                                            
                                                             # Single directory where user-specific configuration files should be written
                                                             XDG_CONFIG_HOME = os.environ.get("XDG_CONFIG_HOME", os.path.join(HOME, ".config"))                                                            
                                                             # List of directories where data files should be searched.
                                                             XDG_DATA_DIRS_LIST = [XDG_DATA_HOME] + "/usr/local/share:/usr/share".split(":")
                                                             XDG_DATA_DIRS = ':'.join((t) for t in XDG_DATA_DIRS_LIST)
                                                             # List of directories where configuration files should be searched.
                                                             XDG_CONFIG_DIRS_LIST = [XDG_CONFIG_HOME] + "/etc/xdg".split(":")
                                                             XDG_CONFIG_DIRS = ':'.join((t) for t in XDG_CONFIG_DIRS_LIST)
                                                             #lp_key = 'LD_LIBRARY_PATH'  # for GNU/Linux and *BSD.
                                                             myenv['XDG_DATA_HOME'] = XDG_DATA_HOME
                                                             myenv['XDG_CONFIG_HOME'] = XDG_CONFIG_HOME
                                                             myenv['XDG_DATA_DIRS'] = XDG_DATA_DIRS
                                                             myenv['XDG_CONFIG_DIRS'] = XDG_CONFIG_DIRS                                                             
                                                             proc = subprocess.check_output(openfile)
                                                             #time.sleep(0.4)
                                                             #poll = proc.poll()
                                                             #print(proc.pid)

                                                             #if poll is not None:
                                                             #    raise Exception("Poll is None")
                                                         except Exception as ex:
                                                             utilities_general.printlogexception(ex=ex)
                                                             if(global_settings.open_with_window == None):
                                                                 global_settings.open_with_window = classes_general.Linux_Open_With()
                                                             global_settings.open_with_window.open_with_window_visible(filepath)
                                                             #
                                                             try:
                                                                 None
                                                                 #mimelist_path = os.path.join(HOME, ".local", "share", "applications", "mimeapps.list")
                                                                 #mimelist_path = os.path.join(HOME, ".config", "mimeapps.list")
                                                                 #self.open_mimelist(mimelist_path)
                                                                 #webbrowser.open_new_tab(filepath)
                                                             except:      
                                                                 None
                                                                 #utilities_general.popup_window('O arquivo não possui um \nprograma associado para abertura!', False)
                                                 except Exception as ex:
                                                     utilities_general.printlogexception(ex=ex)
                                                     utilities_general.popup_window('O arquivo não possui um \nprograma associado para abertura!', False)
                                         elif('uri' in link):
                                             webbrowser.open(link['uri'])
                       self.initialPos = None      
       except TypeError as ex:
           None
       except Exception as ex:
           utilities_general.printlogexception(ex=ex)
           
    def open_mimelist(self, path):
        with open(path, 'r') as mime:
            lines = mime.readlines()
            #print(lines)
          
    def find_similar(self, bytes_from_image, arquivo):
        
        name_pdf_output = global_settings.pathpdfatual+"-temp_out"
        #diretorio_temp_output =global_settings.pathpdfatua l
        similar_each_pdf = process_functions.find_similar(bytes_from_image)
        #print(similar_each_pdf)
        idobscat = None
        relatorioprev = None
        doc = None
        try:
            for relatorio in similar_each_pdf:
                id_pdf = global_settings.infoLaudo[relatorio].id
                for match in similar_each_pdf[relatorio]:
                    hash_image = match[2]
                    if(idobscat==None):
                       idobscat, obscat  = self.add_edit_category('add',arquivo)
                    #SELECT id_image, hash_image, pagina, bbox_x0, bbox_y0, bbox_x1, bbox_y1"
                    pagina = match[2]
                    p0x = match[3]
                    p0y = match[4]
                    p1x = match[5]
                    p1y = match[6]
                    if(relatorio!=relatorioprev):
                        try:
                            doc.close()
                        except:
                            None
                        doc = fitz.open(relatorio)
                        relatorioprev = relatorio
                    self.addmarker2(relatorio, doc, idobscat, id_pdf, pagina, p0x, p0y, p1x, p1y)
            
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        finally:
            try:
                doc.close()
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)
                
    def show_open_with_window(self, filepath):
        if(global_settings.open_with_window == None):
            #if global_settings.plt == "Linux":
            global_settings.open_with_window = classes_general.Linux_Open_With()
        global_settings.open_with_window.open_with_window_visible(filepath)

    def rightClick_link_or_obs(self, event=None):       
        posicaoRealY0Canvas = self.vscrollbar.get()[0] * (self.scrolly) + event.y
        posicaoRealX0Canvas = self.hscrollbar.get()[0] * (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw * self.zoom_x * global_settings.zoom) + event.x
        posicaoRealY0 = (posicaoRealY0Canvas % (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom)) / (self.zoom_x*global_settings.zoom)
        posicaoRealX0 = posicaoRealX0Canvas / (self.zoom_x*global_settings.zoom)   
        pagina = math.floor(posicaoRealY0Canvas / (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh *self.zoom_x*global_settings.zoom))
        menusaveas = None
        for link in global_settings.infoLaudo[global_settings.pathpdfatual].links[pagina]:
            r = link['from']
            xref = link['xref']
            if(posicaoRealX0 >= r.x0 and posicaoRealX0 <= r.x1 and posicaoRealY0 >= r.y0 and posicaoRealY0 <= r.y1):
                if('file' in link and os.path.basename(global_settings.pathpdfatual) not in link['file']):
                    arquivo = link['file']
                    if(arquivo==""):
                        xref = link['xref']
                        info = global_settings.docatual.xref_get_key(xref, 'A')
                        grupos_search = global_settings.regex_actions_compiled.search(info[1])
                        if(grupos_search==None):
                            continue
                        grupos = grupos_search.groups()
                        arquivo = grupos[2]
                        if(grupos[-1]=="Launch"):
                            arquivo = grupos[2]
                    filepath = Path(utilities_general.get_normalized_path(os.path.join(Path(global_settings.pathpdfatual).parent,arquivo)))    
                    menusaveas = tkinter.Menu(global_settings.root, tearoff=0)
                    menusaveas.add_command(label="Salvar como", command= lambda : self.saveas(os.path.basename(filepath), filepath))
                    if global_settings.plt == "Linux":
                        #print("Teste")
                        
                        menusaveas.add_command(label="Abrir com:", command= partial(self.show_open_with_window, filepath))
                    elif global_settings.plt == "Windows" and False:
                        menusaveas.add_command(label="Abrir com:", command= partial(self.show_open_with_window, filepath))
                     
                    try:
                        img = global_settings.docatual.extract_image(xref)
                        #print(img)
                    except:
                        None
                    
                    for link2 in global_settings.infoLaudo[global_settings.pathpdfatual].images[pagina]:
                        #r = link['from']
                        xref2 = link2[0]
                        #print(xref2, xref)
                        img = global_settings.docatual.extract_image(xref2)
                        #print(img)
                        img_bbox = global_settings.docatual[pagina].get_image_bbox(link2)
                        #print(img_bbox)
                        if(posicaoRealX0 >= img_bbox.x0 and posicaoRealX0 <= img_bbox.x1 and posicaoRealY0 >= img_bbox.y0 and posicaoRealY0 <= img_bbox.y1):
                            #print("Arquivo:", arquivo)
                            #menusaveas.add_command(label="Procurar semelhantes", command= lambda : self.find_similar(img['image'], arquivo))
                            #print(img)
                            try:
                                None
                                #img = global_settings.docatual.extract_image(xref)
                                
                            except:
                                None
                            break
                    break
         
        
        deslocy = pagina * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom
        
        sobraEspaco = self.docInnerCanvas.winfo_x()   
        x0 = sobraEspaco + event.x -1
        y0 = math.floor(posicaoRealY0Canvas -1)
        x1 = sobraEspaco + event.x +1
        y1 = math.ceil(posicaoRealY0Canvas +1)
        itens_enclosed = self.docInnerCanvas.find_overlapping(x0, y0, x1, y1)
        for item in itens_enclosed:
            tags = self.docInnerCanvas.gettags(item)
            if('obsitem' in tags):
                if(menusaveas==None):
                    menusaveas = tkinter.Menu(global_settings.root, tearoff=0)
                idobsitem = tags[1]
                menusaveas.add_command(label="Criar/Editar Anotações", image=global_settings.withnote, compound='left', command=partial(self.editadd_anottation,idobsitem))
                break
            elif('enhanceobs' in tags):
                if(menusaveas==None):
                    menusaveas = tkinter.Menu(global_settings.root, tearoff=0)
                idobsitem = tags[1].replace("enhanceobs","obsitem")
                menusaveas.add_command(label="Criar/Editar Anotações", image=global_settings.withnote, compound='left', command=partial(self.editadd_anottation,idobsitem))
                break
        if(menusaveas!=None):
            try:
                menusaveas.tk_popup(event.x_root, event.y_root) 
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)
            finally:
                menusaveas.grab_release()
    
    def already_there(self, x1,y1,x2,y2, tag):
        return False
        encloseds1 = self.docInnerCanvas.find_overlapping(x1-10, y1, x2-10, y2)
        encloseds2 = self.docInnerCanvas.find_enclosed(x1-10, y1, x2-10, y2)
        #print(encloseds1, encloseds2, tag)
        for enclosed in encloseds1:
            #print(self.docInnerCanvas.gettags(enclosed), tag in self.docInnerCanvas.gettags(enclosed))
            if(tag in self.docInnerCanvas.gettags(enclosed)):
                return True
        for enclosed in encloseds2:
            #print(self.docInnerCanvas.gettags(enclosed), tag in self.docInnerCanvas.gettags(enclosed))
            if(tag in self.docInnerCanvas.gettags(enclosed)):
                return True
        return False
                 
    def pintarQuads(self, pagina, p0x, p0y, p1x, p1y, sobraEspaco, enhancetext=False, enhancearea=False, color=None, tag=["quad"], \
                    apagar=True, custom=False, altpressed=False, withborder=True, alt=True):

        if(p1y-p0y<=5):
            p0y -=2
            p1y +=2 
        
        #print(tag)
        quads_on_canvas = []
        try:
            global_settings.zoom = global_settings.listaZooms[global_settings.infoLaudo[global_settings.pathpdfatual].zoom_pos]
            global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[pagina] = {}            
            if(enhancetext):                
                global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[pagina]['text']=[]
                for block in global_settings.infoLaudo[global_settings.pathpdfatual].mapeamento[pagina]:                    
                    x0b = block[0]
                    y0b = block[1]
                    x1b = block[2]
                    y1b = block[3]
                    if(x0b > max(p0x, p1x) or x1b < min(p0x, p1x)):
                        if not (y1b < p1y and y0b > p0y):
                            continue
                    if keyboard.is_pressed('alt'):
                        
                        for line in global_settings.infoLaudo[global_settings.pathpdfatual].mapeamento[pagina][block]:
                            
                            x0l = line[0] 
                            y0l = line[1]
                            x1l = line[2] 
                            y1l = line[3] 
                            if(y1l < p0y or y0l > p1y):
                                continue
                            x0 = min(p0x, p1x)
                            x1 = max(p0x, p1x)
                            rects = []
                            for quad in global_settings.infoLaudo[global_settings.pathpdfatual].mapeamento[pagina][block][line]:
                                qtosrects = len(rects)
                                if( (quad[0]+quad[2])/2 <= x1 and (quad[0]+quad[2])/2 >= x0 and quad[3] >= p0y and quad[1] <= p1y):                                    
                                    if(qtosrects==0):
                                        rect = classes_general.Rect()
                                        rect.x0 = quad[0]
                                        rect.y0 = quad[1]
                                        rect.x1 = quad[2]
                                        rect.y1 = quad[3]
                                        rect.char.append(quad[4])
                                        rects.append(rect)
                                    else:
                                        ultimorect = rects[qtosrects-1]
                                        if(ultimorect.x1+100 >= quad[0]):
                                            ultimorect.char.append(quad[4])
                                            ultimorect.x1 = quad[2]
                                        else:
                                            rect = classes_general.Rect()
                                            rect.x0 = quad[0]
                                            rect.y0 = quad[1]
                                            rect.x1 = quad[2]
                                            rect.y1 = quad[3]
                                            rect.char.append(quad[4])
                                            rects.append(rect)
                            for r in rects:
                                deslocy = pagina * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom
                                x0k = math.floor(r.x0*self.zoom_x*global_settings.zoom +sobraEspaco)
                                x1k = math.ceil(r.x1*self.zoom_x*global_settings.zoom +sobraEspaco)
                                y0k = math.ceil(((r.y0*self.zoom_x*global_settings.zoom)  +deslocy))
                                y1k = math.ceil(((r.y1*self.zoom_x*global_settings.zoom)  +deslocy))
                                r.image = utilities_general.create_rectanglex(min(x0k, x1k), min(y0k, y1k), max(x0k, x1k), max(y0k,y1k), color, withborder=withborder)  
                                if(self.already_there(min(x0k, x1k), min(y0k, y1k), max(x0k, x1k), max(y0k,y1k), tag[0])):
                                    continue
                                r.idrect = self.docInnerCanvas.create_image(min(x0k, x1k), min(y0k, y1k), image=r.image, anchor='nw', tags=(tag))
                                self.addImagetoList(tag[0], r)  
                                #quads_on_canvas.append(r)
                                global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[pagina]['text'].append((line, r))
                    
                        
                    elif(y1b < p1y and y0b > p0y):
                        
                        for line in global_settings.infoLaudo[global_settings.pathpdfatual].mapeamento[pagina][block]:
                            
                            rects = []
                            for quad in global_settings.infoLaudo[global_settings.pathpdfatual].mapeamento[pagina][block][line]:
                                qtosrects = len(rects) 
                                if(qtosrects==0):
                                    rect = classes_general.Rect()
                                    rect.x0 = quad[0]
                                    rect.y0 = quad[1]
                                    rect.x1 = quad[2]
                                    rect.y1 = quad[3]
                                    rect.char.append(quad[4])
                                    rects.append(rect)
                                else:
                                    ultimorect = rects[qtosrects-1]
                                    if(ultimorect.x1+100 >= quad[0]):
                                        ultimorect.char.append(quad[4])
                                        ultimorect.x1 = quad[2]
                                    else:
                                        rect = classes_general.Rect()
                                        rect.x0 = quad[0]
                                        rect.y0 = quad[1]
                                        rect.x1 = quad[2]
                                        rect.y1 = quad[3]
                                        rect.char.append(quad[4])
                                        rects.append(rect)
                            for r in rects:
                                deslocy = pagina * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom
                                deslocx =  self.hscrollbar.get()[0] * self.canvasw        
                                x0k = math.floor(r.x0*self.zoom_x*global_settings.zoom +sobraEspaco)
                                x1k = math.ceil(r.x1*self.zoom_x*global_settings.zoom+sobraEspaco)
                                y0k = math.ceil(((r.y0*self.zoom_x*global_settings.zoom)  +deslocy))
                                y1k = math.ceil(((r.y1*self.zoom_x*global_settings.zoom)  +deslocy))
                                if(self.already_there(min(x0k, x1k), min(y0k, y1k), max(x0k, x1k), max(y0k,y1k), tag[0])):
                                    continue
                                r.image = utilities_general.create_rectanglex(min(x0k, x1k), min(y0k, y1k), max(x0k, x1k), max(y0k,y1k), color, withborder=withborder)
                                r.idrect = self.docInnerCanvas.create_image(min(x0k, x1k), min(y0k, y1k), image=r.image, anchor='nw', tags=(tag))
                                quads_on_canvas.append(r)
                                self.addImagetoList(tag[0], r)
                                global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[pagina]['text'].append((line, r))
                    else:
                        
                        
                        linetop = True
                        linebottom = True
                        not_painted_solo_line = True
                        not_painted_end_block = True
                        not_painted_start_block = True
                        whole_line = False
                        for line in global_settings.infoLaudo[global_settings.pathpdfatual].mapeamento[pagina][block]:
                            x0l = line[0] 
                            y0l = line[1]
                            x1l = line[2] 
                            y1l = line[3]
                            #if(x0l > max(p0x, p1x) or x1l < min(p0x, p1x)):
                            #   continue
                            if(y1l < p0y):                                
                                continue
                            if(y0l > p1y):
                                continue                            
                            if(p0y >= y0l and p1y <= y1l and not_painted_solo_line):    
                                x0 = min(p0x, p1x)
                                x1 = max(p0x, p1x)                                
                                rects = []
                                for quad in global_settings.infoLaudo[global_settings.pathpdfatual].mapeamento[pagina][block][line]:
                                    qtosrects = len(rects)
                                    if( quad[2] <= x1 and (quad[0]+quad[2])/2 >= x0):                                        
                                        if(qtosrects==0):
                                            rect = classes_general.Rect()
                                            rect.x0 = quad[0]
                                            rect.y0 = quad[1]
                                            rect.x1 = quad[2]
                                            rect.y1 = quad[3]
                                            rect.char.append(quad[4])
                                            rects.append(rect)
                                        else:
                                            ultimorect = rects[qtosrects-1]
                                            if(ultimorect.x1+100 >= quad[0]):
                                                ultimorect.char.append(quad[4])
                                                ultimorect.x1 = quad[2]
                                            else:
                                                rect = classes_general.Rect()
                                                rect.x0 = quad[0]
                                                rect.y0 = quad[1]
                                                rect.x1 = quad[2]
                                                rect.y1 = quad[3]
                                                rect.char.append(quad[4])
                                                rects.append(rect)
                                for r in rects:
                                    not_painted_solo_line=False
                                    deslocy = pagina * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom
                                    deslocx =  self.hscrollbar.get()[0] * self.canvasw
        
                                    x0k = math.floor(r.x0*self.zoom_x*global_settings.zoom +sobraEspaco)
                                    x1k = math.ceil(r.x1*self.zoom_x*global_settings.zoom +sobraEspaco)
                                    y0k = math.ceil(((r.y0*self.zoom_x*global_settings.zoom)  +deslocy))
                                    y1k = math.ceil(((r.y1*self.zoom_x*global_settings.zoom)  +deslocy))
                                    if(self.already_there(min(x0k, x1k), min(y0k, y1k), max(x0k, x1k), max(y0k,y1k), tag[0])):
                                        continue
                                    r.image = utilities_general.create_rectanglex(min(x0k, x1k), min(y0k, y1k), max(x0k, x1k), max(y0k,y1k), color, withborder=withborder)
                                    r.idrect = self.docInnerCanvas.create_image(min(x0k, x1k), min(y0k, y1k), image=r.image, anchor='nw', tags=(tag))
                                    self.addImagetoList(tag[0], r)
                                    quads_on_canvas.append(r)
                                    global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[pagina]['text'].append((line, r))
                            elif(p0y < y0l and p1y > y1l):  
                                rects = []
                                for quad in global_settings.infoLaudo[global_settings.pathpdfatual].mapeamento[pagina][block][line]:
                                    qtosrects = len(rects) 
                                    if(qtosrects==0):
                                        rect = classes_general.Rect()
                                        rect.x0 = quad[0]
                                        rect.y0 = quad[1]
                                        rect.x1 = quad[2]
                                        rect.y1 = quad[3]
                                        rect.char.append(quad[4])
                                        rects.append(rect)
                                    else:
                                        ultimorect = rects[qtosrects-1]
                                        if(ultimorect.x1+100 >= quad[0]):
                                            ultimorect.char.append(quad[4])
                                            ultimorect.x1 = quad[2]
                                        else:
                                            rect = classes_general.Rect()
                                            rect.x0 = quad[0]
                                            rect.y0 = quad[1]
                                            rect.x1 = quad[2]
                                            rect.y1 = quad[3]
                                            rect.char.append(quad[4])
                                            rects.append(rect)
                                for r in rects:
                                    deslocy = pagina * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom
                                    deslocx =  self.hscrollbar.get()[0] * self.canvasw        
                                    x0k = math.floor(r.x0*self.zoom_x*global_settings.zoom +sobraEspaco)
                                    x1k = math.ceil(r.x1*self.zoom_x*global_settings.zoom+sobraEspaco)
                                    y0k = math.ceil(((r.y0*self.zoom_x*global_settings.zoom)  +deslocy))
                                    y1k = math.ceil(((r.y1*self.zoom_x*global_settings.zoom)  +deslocy))
                                    if(self.already_there(min(x0k, x1k), min(y0k, y1k), max(x0k, x1k), max(y0k,y1k), tag[0])):
                                        continue
                                    r.image = utilities_general.create_rectanglex(min(x0k, x1k), min(y0k, y1k), max(x0k, x1k), max(y0k,y1k), color, withborder=withborder)
                                    r.idrect = self.docInnerCanvas.create_image(min(x0k, x1k), min(y0k, y1k), image=r.image, anchor='nw', tags=(tag))
                                    self.addImagetoList(tag[0], r)
                                    quads_on_canvas.append(r)
                                    global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[pagina]['text'].append((line, r))
                            elif(p0y < y1l and p1y > y1l and not_painted_start_block and not_painted_solo_line):                                  
                                linetop = False
                                rects = []
                                for quad in global_settings.infoLaudo[global_settings.pathpdfatual].mapeamento[pagina][block][line]:
                                    qtosrects = len(rects)
                                    if((quad[0]+quad[2])/2 >= p0x):                                        
                                        if(qtosrects==0):
                                            rect = classes_general.Rect()
                                            rect.x0 = quad[0]
                                            rect.y0 = quad[1]
                                            rect.x1 = quad[2]
                                            rect.y1 = quad[3]
                                            rect.char.append(quad[4])
                                            rects.append(rect)
                                        else:
                                            ultimorect = rects[qtosrects-1]
                                            if(ultimorect.x1+100 >= quad[0]):
                                                ultimorect.char.append(quad[4])
                                                ultimorect.x1 = quad[2]
                                            else:
                                                rect = classes_general.Rect()
                                                rect.x0 = quad[0]
                                                rect.y0 = quad[1]
                                                rect.x1 = quad[2]
                                                rect.y1 = quad[3]
                                                rect.char.append(quad[4])
                                                rects.append(rect)
                                for r in rects:
                                    not_painted_start_block = False
                                    deslocy = pagina * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom                                    
                                    x0k = math.floor(r.x0*self.zoom_x*global_settings.zoom +sobraEspaco)
                                    x1k = math.ceil(r.x1*self.zoom_x*global_settings.zoom+sobraEspaco )
                                    y0k = math.ceil(((r.y0*self.zoom_x*global_settings.zoom)  +deslocy))
                                    y1k = math.ceil(((r.y1*self.zoom_x*global_settings.zoom)  +deslocy))
                                    if(self.already_there(min(x0k, x1k), min(y0k, y1k), max(x0k, x1k), max(y0k,y1k), tag[0])):
                                        continue
                                    r.image = utilities_general.create_rectanglex(min(x0k, x1k), min(y0k, y1k), max(x0k, x1k), max(y0k,y1k), color, withborder=withborder)
                                    r.idrect = self.docInnerCanvas.create_image(min(x0k, x1k), min(y0k, y1k), image=r.image, anchor='nw', tags=(tag))
                                    self.addImagetoList(tag[0], r)
                                    quads_on_canvas.append(r)
                                    global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[pagina]['text'].append((line, r))   
                            elif(p1y > y0l and p0y < y0l and  not_painted_end_block and not_painted_solo_line):
                                                                
                                linebottom = False
                                rects = []
                                for quad in global_settings.infoLaudo[global_settings.pathpdfatual].mapeamento[pagina][block][line]:
                                    qtosrects = len(rects)
                                    if((quad[0]+quad[2])/2 <= p1x):                                        
                                        if(qtosrects==0):
                                            rect = classes_general.Rect()
                                            rect.x0 = quad[0]
                                            rect.y0 = quad[1]
                                            rect.x1 = quad[2]
                                            rect.y1 = quad[3]
                                            rect.char.append(quad[4])
                                            rects.append(rect)
                                        else:
                                            ultimorect = rects[qtosrects-1]
                                            if(ultimorect.x1+100 >= quad[0]):
                                                ultimorect.char.append(quad[4])
                                                ultimorect.x1 = quad[2]
                                            else:
                                                rect = classes_general.Rect()
                                                rect.x0 = quad[0]
                                                rect.y0 = quad[1]
                                                rect.x1 = quad[2]
                                                rect.y1 = quad[3]
                                                rect.char.append(quad[4])
                                                rects.append(rect)
                                for r in rects:
                                    not_painted_end_block = False
                                    deslocy = pagina * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom                                    
                                    x0k = math.floor(r.x0*self.zoom_x*global_settings.zoom +sobraEspaco)
                                    x1k = math.ceil(r.x1*self.zoom_x*global_settings.zoom+sobraEspaco )
                                    y0k = math.ceil(((r.y0*self.zoom_x*global_settings.zoom)  +deslocy))
                                    y1k = math.ceil(((r.y1*self.zoom_x*global_settings.zoom)  +deslocy))
                                    if(self.already_there(min(x0k, x1k), min(y0k, y1k), max(x0k, x1k), max(y0k,y1k), tag[0])):
                                        continue
                                    r.image = utilities_general.create_rectanglex(min(x0k, x1k), min(y0k, y1k), max(x0k, x1k), max(y0k,y1k), color, withborder=withborder)
                                    r.idrect = self.docInnerCanvas.create_image(min(x0k, x1k), min(y0k, y1k), image=r.image, anchor='nw', tags=(tag))
                                    self.addImagetoList(tag[0], r)
                                    quads_on_canvas.append(r)
                                    global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[pagina]['text'].append((line, r))
                            elif(p0y >= y0l and p1y <= y1l):
                                for line in global_settings.infoLaudo[global_settings.pathpdfatual].mapeamento[pagina][block]:
                                    
                                    x0l = line[0] 
                                    y0l = line[1]
                                    x1l = line[2] 
                                    y1l = line[3] 
                                    if(y1l < p0y or y0l > p1y):
                                        continue
                                    x0 = min(p0x, p1x)
                                    x1 = max(p0x, p1x)
                                    rects = []
                                    for quad in global_settings.infoLaudo[global_settings.pathpdfatual].mapeamento[pagina][block][line]:
                                        qtosrects = len(rects)
                                        if( (quad[0]+quad[2])/2 <= x1 and (quad[0]+quad[2])/2 >= x0 and quad[3] >= p0y and quad[1] <= p1y):                                    
                                            if(qtosrects==0):
                                                rect = classes_general.Rect()
                                                rect.x0 = quad[0]
                                                rect.y0 = quad[1]
                                                rect.x1 = quad[2]
                                                rect.y1 = quad[3]
                                                rect.char.append(quad[4])
                                                rects.append(rect)
                                            else:
                                                ultimorect = rects[qtosrects-1]
                                                if(ultimorect.x1+100 >= quad[0]):
                                                    ultimorect.char.append(quad[4])
                                                    ultimorect.x1 = quad[2]
                                                else:
                                                    rect = classes_general.Rect()
                                                    rect.x0 = quad[0]
                                                    rect.y0 = quad[1]
                                                    rect.x1 = quad[2]
                                                    rect.y1 = quad[3]
                                                    rect.char.append(quad[4])
                                                    rects.append(rect)
                                    for r in rects:
                                        deslocy = pagina * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom
                                        x0k = math.floor(r.x0*self.zoom_x*global_settings.zoom +sobraEspaco)
                                        x1k = math.ceil(r.x1*self.zoom_x*global_settings.zoom +sobraEspaco)
                                        y0k = math.ceil(((r.y0*self.zoom_x*global_settings.zoom)  +deslocy))
                                        y1k = math.ceil(((r.y1*self.zoom_x*global_settings.zoom)  +deslocy))
                                        if(self.already_there(min(x0k, x1k), min(y0k, y1k), max(x0k, x1k), max(y0k,y1k), tag[0])):
                                            continue
                                        r.image = utilities_general.create_rectanglex(min(x0k, x1k), min(y0k, y1k), max(x0k, x1k), max(y0k,y1k), color, withborder=withborder)    
                                        
                                        r.idrect = self.docInnerCanvas.create_image(min(x0k, x1k), min(y0k, y1k), image=r.image, anchor='nw', tags=(tag))
                                        self.addImagetoList(tag[0], r)  
                                        #quads_on_canvas.append(r)
                                        global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[pagina]['text'].append((line, r))
                            else:
                                #print(p0x, p0y, p1x, p1y)
                                #print("Here")
                                #continue
                                #p0y += 1
                                for line in global_settings.infoLaudo[global_settings.pathpdfatual].mapeamento[pagina][block]:
                                    
                                    x0l = line[0] 
                                    y0l = line[1]
                                    x1l = line[2] 
                                    y1l = line[3] 
                                    if(y1l < p0y or y0l > p1y):
                                        continue
                                    x0 = min(p0x, p1x)
                                    x1 = max(p0x, p1x)
                                    rects = []
                                    for quad in global_settings.infoLaudo[global_settings.pathpdfatual].mapeamento[pagina][block][line]:
                                        qtosrects = len(rects)
                                        if( (quad[0]+quad[2])/2 <= x1 and (quad[0]+quad[2])/2 >= x0 and quad[3] >= p0y and quad[1] <= p1y):                                    
                                            if(qtosrects==0):
                                                rect = classes_general.Rect()
                                                rect.x0 = quad[0]
                                                rect.y0 = quad[1]
                                                rect.x1 = quad[2]
                                                rect.y1 = quad[3]
                                                rect.char.append(quad[4])
                                                rects.append(rect)
                                            else:
                                                ultimorect = rects[qtosrects-1]
                                                if(ultimorect.x1+100 >= quad[0]):
                                                    ultimorect.char.append(quad[4])
                                                    ultimorect.x1 = quad[2]
                                                else:
                                                    rect = classes_general.Rect()
                                                    rect.x0 = quad[0]
                                                    rect.y0 = quad[1]
                                                    rect.x1 = quad[2]
                                                    rect.y1 = quad[3]
                                                    rect.char.append(quad[4])
                                                    rects.append(rect)
                                    for r in rects:
                                        deslocy = pagina * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom
                                        x0k = math.floor(r.x0*self.zoom_x*global_settings.zoom +sobraEspaco)
                                        x1k = math.ceil(r.x1*self.zoom_x*global_settings.zoom +sobraEspaco)
                                        y0k = math.ceil(((r.y0*self.zoom_x*global_settings.zoom)  +deslocy))
                                        y1k = math.ceil(((r.y1*self.zoom_x*global_settings.zoom)  +deslocy))
                                        if(self.already_there(min(x0k, x1k), min(y0k, y1k), max(x0k, x1k), max(y0k,y1k), tag[0])):
                                            continue
                                        r.image = utilities_general.create_rectanglex(min(x0k, x1k), min(y0k, y1k), max(x0k, x1k), max(y0k,y1k), color, withborder=withborder)                                
                                        r.idrect = self.docInnerCanvas.create_image(min(x0k, x1k), min(y0k, y1k), image=r.image, anchor='nw', tags=(tag))
                                        self.addImagetoList(tag[0], r)  
                                        #quads_on_canvas.append(r)
                                        global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[pagina]['text'].append((line, r))
            elif(enhancearea):
                global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[pagina]['areaSelection'] = []
                deslocy = pagina * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom
                rect = classes_general.Rect()
                rect.x0 = p0x
                rect.y0 = p0y
                rect.x1 = p1x
                rect.y1 = p1y    
                x0k = math.floor(rect.x0*self.zoom_x*global_settings.zoom+ sobraEspaco)
                x1k = math.ceil(rect.x1*self.zoom_x*global_settings.zoom+ sobraEspaco)
                y0k = math.ceil(((rect.y0*self.zoom_x*global_settings.zoom)  +deslocy))
                y1k = math.ceil(((rect.y1*self.zoom_x*global_settings.zoom)  +deslocy))
                rect.image = utilities_general.create_rectanglex(min(x0k, x1k), min(y0k, y1k), max(x0k, x1k), max(y0k,y1k), color, withborder=withborder)
                rect.idrect = self.docInnerCanvas.create_image(min(x0k, x1k), min(y0k, y1k), image=rect.image, anchor='nw', tags=(tag))
                self.addImagetoList(tag[0], rect)
                quads_on_canvas.append(rect)
                global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[pagina]['areaSelection'].append((None, rect))
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        return quads_on_canvas
        

    def prepararParaQuads(self, pagina, posicaoRealX0, posicaoRealY0, posicaoRealX1, posicaoRealY1, \
                          color=(21, 71, 150, 85),tag=["quad"], apagar=True, enhancetext=False, enhancearea=False, withborder=True, alt=False, first=True):        
        margemsup = (global_settings.infoLaudo[global_settings.pathpdfatual].mt/25.4)*72
        margeminf = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh-((global_settings.infoLaudo[global_settings.pathpdfatual].mb/25.4)*72)
        margemesq = (global_settings.infoLaudo[global_settings.pathpdfatual].me/25.4)*72
        margemdir = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw-((global_settings.infoLaudo[global_settings.pathpdfatual].md/25.4)*72)
        if(pagina not in global_settings.infoLaudo[global_settings.pathpdfatual].mapeamento):
            if(first or pagina in global_settings.processed_pages or pagina in global_settings.infoLaudo[global_settings.pathpdfatual].mapeamento):
                self.afterquads = global_settings.root.after(500, lambda: self.prepararParaQuads( pagina, posicaoRealX0, posicaoRealY0, posicaoRealX1, \
                                                                         posicaoRealY1, color=color,tag=tag, apagar=apagar, enhancetext=enhancetext, \
                                                                             enhancearea=enhancearea, withborder=withborder, alt=alt, first=False))
        else:  
            if("enhanceobs" in tag[0]):
                listaquads = self.docInnerCanvas.find_withtag(tag[2])
                #alt = Tru
                if(len(listaquads)>0):
                    return []
            elif("enhanceannot" in tag[0]):
                enhancetext = False
                enhancearea = True
                listaquads = self.docInnerCanvas.find_withtag(tag[1])
                if(len(listaquads)>0):
                    return []
                
            if(posicaoRealX0 <= posicaoRealX1 and posicaoRealY0 <= posicaoRealY1):
                p0x = posicaoRealX0
                p0y = posicaoRealY0
                p1x = posicaoRealX1
                p1y = posicaoRealY1
            elif(posicaoRealX0 > posicaoRealX1 and posicaoRealY0 <= posicaoRealY1):                
                p0x = posicaoRealX0
                p0y = posicaoRealY0
                p1x = posicaoRealX1
                p1y = posicaoRealY1
            elif(posicaoRealX0 <= posicaoRealX1 and posicaoRealY0 > posicaoRealY1):
                p0x = posicaoRealX1
                p0y = posicaoRealY1
                p1x = posicaoRealX0
                p1y = posicaoRealY0
            elif (posicaoRealX0 > posicaoRealX1 and posicaoRealY0 > posicaoRealY1):
                p0x = posicaoRealX1
                p0y = posicaoRealY1
                p1x = posicaoRealX0
                p1y = posicaoRealY0
                
            
            sobraEspaco = self.docInnerCanvas.winfo_x()   
            
            
            quads_on_canvas = self.pintarQuads(pagina=pagina, p0x=p0x, p0y=p0y, p1x=p1x, p1y=p1y, sobraEspaco=sobraEspaco, color=color, apagar=apagar, \
                             custom=False, tag=tag, altpressed=self.altpressed and alt, enhancetext=enhancetext, enhancearea=enhancearea, withborder=withborder, alt=alt)
            return quads_on_canvas
        
    def scrollByMouseOutCanvas(self, dif):
        if(self.initialPos != None):
            if(global_settings.root.winfo_pointery()-dif > self.docInnerCanvas.winfo_height()):
                self.docInnerCanvas.yview_scroll(1, "units")
            elif(global_settings.root.winfo_pointery()-dif < 0):     
                self.docInnerCanvas.yview_scroll(-1, "units") 
            self._jobscrollpagebymouse = global_settings.root.after(50, lambda d=dif: self.scrollByMouseOutCanvas(d))
        
    def _selectingText(self, evento):        
        margemsup = (global_settings.infoLaudo[global_settings.pathpdfatual].mt/25.4)*72
        margeminf = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh-((global_settings.infoLaudo[global_settings.pathpdfatual].mb/25.4)*72)
        margemesq = (global_settings.infoLaudo[global_settings.pathpdfatual].me/25.4)*72
        margemdir = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw-((global_settings.infoLaudo[global_settings.pathpdfatual].md/25.4)*72)
        try:
            if(self.selectionActive or self.areaselectionActive):            
                if(isinstance(evento.widget, tkinter.Canvas) and self.initialPos==None):
                    None
                        
                if(self.initialPos!=None and isinstance(evento.widget, tkinter.Canvas)):
                    dif = global_settings.root.winfo_pointery() - evento.y
                    try:
                        global_settings.root.after_cancel(self._jobscrollpagebymouse)
                    except Exception as ex:
                        None
                    self._jobscrollpagebymouse = global_settings.root.after(50, lambda d=dif: self.scrollByMouseOutCanvas(d))
                    self.docInnerCanvas.delete("quad")
                    self.clearSomeImages(["quad"])
                    posicaoRealX0=self.initialPos[2]
                    posicaoRealY0=self.initialPos[3]
                    posicaoRealX0Canvas=self.initialPos[0]
                    posicaoRealY0Canvas=self.initialPos[1]                    
                    posicaoRealY1Canvas = self.vscrollbar.get()[0] * (self.scrolly) + evento.y 
                    posicaoRealX1Canvas = self.hscrollbar.get()[0] * (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw * self.zoom_x * global_settings.zoom) + evento.x
                    posicaoRealY1 = (posicaoRealY1Canvas % (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom)) / (self.zoom_x*global_settings.zoom)
                    posicaoRealX1 = posicaoRealX1Canvas / (self.zoom_x*global_settings.zoom)                    
                    if(posicaoRealX0Canvas <= posicaoRealX1Canvas and posicaoRealY0Canvas <= posicaoRealY1Canvas):
                        p0xc = posicaoRealX0Canvas
                        p0yc = posicaoRealY0Canvas
                        p1xc = posicaoRealX1Canvas
                        p1yc = posicaoRealY1Canvas
                        p0x = posicaoRealX0
                        p0y = posicaoRealY0
                        p1x = posicaoRealX1
                        p1y = posicaoRealY1
                    elif(posicaoRealX0Canvas > posicaoRealX1Canvas and posicaoRealY0Canvas <= posicaoRealY1Canvas):                
                        p0xc = posicaoRealX0Canvas
                        p0yc = posicaoRealY0Canvas
                        p1xc = posicaoRealX1Canvas
                        p1yc = posicaoRealY1Canvas
                        p0x = posicaoRealX0
                        p0y = posicaoRealY0
                        p1x = posicaoRealX1
                        p1y = posicaoRealY1
                    elif(posicaoRealX0Canvas <= posicaoRealX1Canvas and posicaoRealY0Canvas > posicaoRealY1Canvas):
                        p0xc = posicaoRealX1Canvas
                        p0yc = posicaoRealY1Canvas
                        p1xc = posicaoRealX0Canvas
                        p1yc = posicaoRealY0Canvas
                        p0x = posicaoRealX1
                        p0y = posicaoRealY1
                        p1x = posicaoRealX0
                        p1y = posicaoRealY0
                    elif (posicaoRealX0Canvas > posicaoRealX1Canvas and posicaoRealY0Canvas > posicaoRealY1Canvas):
                        p0xc = posicaoRealX1Canvas 
                        p0yc = posicaoRealY1Canvas 
                        p1xc = posicaoRealX0Canvas 
                        p1yc = posicaoRealY0Canvas 
                        p0x = posicaoRealX1
                        p0y = posicaoRealY1
                        p1x = posicaoRealX0
                        p1y = posicaoRealY0
                    pp = math.floor(p0yc / (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh *self.zoom_x*global_settings.zoom))
                    up = math.floor(p1yc / (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh *self.zoom_x*global_settings.zoom))
                    del_pages = []
                    for pagina in global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados:
                        if(pagina<pp or pagina>up):
                            del_pages.append(pagina)
                    for pagina in del_pages:
                        del global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[pagina]
                    self.exportinterval.initpageVar.set(pp+1)
                    self.exportinterval.inityVar.set(round(p0y,0))
                    self.exportinterval.endpageVar.set(up+1)
                    self.exportinterval.endyVar.set(round(p1y, 0))
                    origemx = 0
                    origemx1 = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw
                    if(self.altpressed or self.areaselectionActive):
                        origemx=self.initialPos[2]
                        origemx1=posicaoRealX1
                    
                    tags = ["quad"]
                    if(self.altpressed):
                        tags.append("withalt")
                    for p in range(pp, up+1):
                        posicoes_normalizadas = self.normalize_positions(p, pp, up, \
                                                 p0x, p0y, p1x, p1y)
                        posicaoRealX0 = posicoes_normalizadas[0]
                        posicaoRealY0 = posicoes_normalizadas[1]
                        posicaoRealX1 = posicoes_normalizadas[2]
                        posicaoRealY1 = posicoes_normalizadas[3]
                        
                        self.prepararParaQuads(p, posicaoRealX0=posicaoRealX0, posicaoRealY0=posicaoRealY0, posicaoRealX1=posicaoRealX1, posicaoRealY1=posicaoRealY1, \
                                               color=(21, 71, 150, 85),tag=tags, apagar=True, enhancetext=self.selectionActive, enhancearea=self.areaselectionActive)
                else:
                    None
            else:
                self.docInnerCanvas.scan_dragto(evento.x, evento.y, gain=1)
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        
    def doubleClickSelection(self, evento):         
         global_settings.zoom = global_settings.listaZooms[global_settings.infoLaudo[global_settings.pathpdfatual].zoom_pos]
         if(isinstance(evento.widget, tkinter.Canvas)):
                self.docInnerCanvas.delete("simplesearch")
                self.docInnerCanvas.delete("quad")
                self.docInnerCanvas.delete("obsitem")
                self.clearSomeImages(["simplesearch", "quad", "obsitem"])
                posicaoRealY0Canvas = self.vscrollbar.get()[0] * (self.scrolly) + evento.y
                posicaoRealX0Canvas = self.hscrollbar.get()[0] * (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw * self.zoom_x * global_settings.zoom) + evento.x
                posicaoRealY0 = (posicaoRealY0Canvas % (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom)) / (self.zoom_x*global_settings.zoom)
                posicaoRealX0 = posicaoRealX0Canvas / (self.zoom_x*global_settings.zoom)
                pagina = math.floor(posicaoRealY0Canvas / (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh *self.zoom_x*global_settings.zoom))
                p0x = posicaoRealX0
                p0y = posicaoRealY0
                p1x = posicaoRealX0
                p1y = posicaoRealY0
                sobraEspaco = 0
                if(self.docFrame.winfo_width() > global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw * self.zoom_x * global_settings.zoom):
                    sobraEspaco = self.docInnerCanvas.winfo_x()
                if(self.selectionActive):
                    rect = None 
                    for block in global_settings.infoLaudo[global_settings.pathpdfatual].mapeamento[pagina]:
                        x0b = block[0]
                        y0b = block[1]
                        x1b = block[2]
                        y1b = block[3]
                        if(y0b <= p0y and y1b >= p1y):
                            for line in global_settings.infoLaudo[global_settings.pathpdfatual].mapeamento[pagina][block]:
                                x0l = line[0]
                                y0l = line[1]
                                x1l = line[2]
                                y1l = line[3]                                
                                if(y0l <= p0y and y1l >= p1y and x0l <= p0x and x1l >=p1x):
                                    for quad in global_settings.infoLaudo[global_settings.pathpdfatual].mapeamento[pagina][block][line]:
                                        if(quad[4] == " "):   
                                            if(quad[0] <= p0x):
                                                rect = None
                                            else:
                                                break
                                        else:
                                            if(rect==None):
                                                rect = classes_general.Rect()
                                                rect.x0 = quad[0]
                                                rect.y0 = quad[1]
                                                rect.y1 = quad[3]
                                                rect.x1 = quad[2]
                                            rect.char.append(quad[4])
                                            rect.x0 = min(rect.x0, quad[0])
                                            rect.y0 = min(rect.y0, quad[1])
                                            rect.x1 = max(rect.x1, quad[2])
                                            rect.y1 = max(rect.y1, quad[3])
                    if(rect!=None):
                        deslocy = pagina * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom
                        x0k = math.floor(rect.x0*self.zoom_x*global_settings.zoom + sobraEspaco)
                        x1k = math.ceil(rect.x1*self.zoom_x*global_settings.zoom + sobraEspaco)
                        y0k = math.ceil(((rect.y0*self.zoom_x*global_settings.zoom)  +deslocy))
                        y1k = math.ceil(((rect.y1*self.zoom_x*global_settings.zoom)  +deslocy))
                        rect.image = utilities_general.create_rectanglex(min(x0k, x1k), min(y0k, y1k), max(x0k, x1k), max(y0k,y1k), (21, 71, 150, 85), withborder=False)
                        rect.idrect = self.docInnerCanvas.create_image(min(x0k, x1k), min(y0k, y1k), image=rect.image, anchor='nw', tags=("quad"))
                        self.addImagetoList("quad", rect.image)
                        global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[pagina] = {}
                        global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[pagina]['text'] = []
                        global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[pagina]['text'].append((line, rect))
            
    def _clearClick(self, event):
        self.initialPos = None
    
    def zoomx(self, event=None, tipozoom='', pathpdfantigo = None, moveto = None):
        
        if((tipozoom=='plus' and global_settings.infoLaudo[global_settings.pathpdfatual].zoom_pos < len(global_settings.listaZooms)-1) \
           or (tipozoom=='minus' and global_settings.infoLaudo[global_settings.pathpdfatual].zoom_pos > 0) or
           pathpdfantigo != None):
            
            self.winfox = self.docInnerCanvas.winfo_x()
            valor = moveto
            if(valor == None):
                valor = self.vscrollbar.get()[0]
            for k in range(global_settings.minMaxLabels):
                self.docInnerCanvas.itemconfig(self.ininCanvasesid[k], image = None)
                self.tkimgs[k] = None
            oldzoom = global_settings.listaZooms[global_settings.infoLaudo[global_settings.pathpdfatual].zoom_pos]
            if(pathpdfantigo!=None):
                oldzoom = global_settings.listaZooms[global_settings.infoLaudo[pathpdfantigo].zoom_pos]
            else:
                global_settings.infoLaudo[global_settings.pathpdfatual].something_changed = True
            if(tipozoom=='plus'):
                global_settings.infoLaudo[global_settings.pathpdfatual].zoom_pos += 1
            elif(tipozoom=='minus'):
                global_settings.infoLaudo[global_settings.pathpdfatual].zoom_pos -= 1
            global_settings.zoom = global_settings.listaZooms[global_settings.infoLaudo[global_settings.pathpdfatual].zoom_pos]

            self.mat = fitz.Matrix(self.zoom_x*global_settings.zoom, self.zoom_x*global_settings.zoom)
            
            sobraEspacoold = self.docInnerCanvas.winfo_x()
            self.docInnerCanvas.config(width= (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw * self.zoom_x * global_settings.zoom))
            for i in range(global_settings.minMaxLabels):
                global_settings.processed_pages[i] = -1   
            self.docInnerCanvas.delete(self.fakeLines[0])
            self.docInnerCanvas.delete(self.fakeLines[1])
            self.fakeLines[0] = self.docInnerCanvas.create_line(0,0, max(self.docFrame.winfo_width(), global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw*\
                                                                         self.zoom_x*global_settings.zoom),\
                                                                0, width=5, fill=self.bg)
            self.fakeLines[1] = self.docInnerCanvas.create_line(0,global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x * global_settings.zoom, \
                                                                max(self.docFrame.winfo_width(), \
                                                            global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh*self.zoom_x*global_settings.zoom), \
                                                            global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x * global_settings.zoom, width=5, fill=self.bg) 
            sobraEspaco = self.docInnerCanvas.winfo_x()
            self.docInnerCanvas.configure(scrollregion = (sobraEspaco,0,sobraEspaco+global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw * self.zoom_x*global_settings.zoom, \
                                                          global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh*self.zoom_x*global_settings.zoom*\
                                                              global_settings.infoLaudo[global_settings.pathpdfatual].len))
            self.docInnerCanvas.yview_moveto(valor)
            
            self.scrolly = round((global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh*self.zoom_x*global_settings.zoom), 16)*\
                global_settings.infoLaudo[global_settings.pathpdfatual].len  - 35
            try:  
                listasimplesearch = self.docInnerCanvas.find_withtag("simplesearch")
                listaquads = self.docInnerCanvas.find_withtag("quad")
                listalinks = self.docInnerCanvas.find_withtag("link")
                listaobs = self.docInnerCanvas.find_withtag("obsitem")
                #listaenhanceobs = self.docInnerCanvas.find_withtag("enhanceobs")
                self.clearEnhanceObs() 
                #listaenhanceannots = self.docInnerCanvas.find_withtag("enhanceannots")
                self.clearSomeImages(["simplesearch", "quad", "link", "obsitem", "enhanceobs", "enhanceannot"])
                for quadelement in listasimplesearch:
                    box = (self.docInnerCanvas.bbox(quadelement))
                    pagina = math.floor(box[1]/(global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*oldzoom))
                    deslocy = pagina * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom
                    deslocyold = pagina * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*oldzoom
                    x0novo = round(((box[0]-sobraEspacoold)/(self.zoom_x*oldzoom))*self.zoom_x*global_settings.zoom + sobraEspaco)
                    x1novo = round(((box[2]-sobraEspacoold)/(self.zoom_x*oldzoom))*self.zoom_x*global_settings.zoom + sobraEspaco)
                    y0novo =round(((box[1]-deslocyold)/(self.zoom_x*oldzoom))*self.zoom_x*global_settings.zoom  +deslocy)
                    y1novo = round(((box[3]-deslocyold)/(self.zoom_x*oldzoom))*self.zoom_x*global_settings.zoom  +deslocy)    
                    
                    tempimagem = (utilities_general.create_rectanglex(x0novo, y0novo, x1novo, y1novo, self.color, link=False))                    
                    self.docInnerCanvas.itemconfig(quadelement, image=tempimagem)#
                    self.addImagetoList("simplesearch", tempimagem)
                    coords = self.docInnerCanvas.coords(quadelement)
                    dx = x0novo -coords[0]
                    dy = y0novo -coords[1]
                    self.docInnerCanvas.move(quadelement, dx, dy)                    
                for quadelement in listaquads:                    
                    box = (self.docInnerCanvas.bbox(quadelement)) 
                    pagina = math.floor(box[1]/((global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*oldzoom)))
                    deslocy = pagina * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom
                    deslocyold = pagina * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*oldzoom
                    x0novo = round(((box[0]-sobraEspacoold)/(self.zoom_x*oldzoom))*self.zoom_x*global_settings.zoom + sobraEspaco)
                    x1novo = round(((box[2]-sobraEspacoold)/(self.zoom_x*oldzoom))*self.zoom_x*global_settings.zoom + sobraEspaco)
                    y0novo = round(((box[1]-deslocyold)/(self.zoom_x*oldzoom))*self.zoom_x*global_settings.zoom  +deslocy)
                    y1novo = round(((box[3]-deslocyold)/(self.zoom_x*oldzoom))*self.zoom_x*global_settings.zoom  +deslocy)                       
                    tempimagem = (utilities_general.create_rectanglex(x0novo, y0novo, x1novo, y1novo, (21, 71, 150, 85), link=False))
                    self.docInnerCanvas.itemconfig(quadelement, image=tempimagem)
                    self.addImagetoList("quad", tempimagem)
                    coords = self.docInnerCanvas.coords(quadelement)
                    dx = x0novo -coords[0]
                    dy = y0novo -coords[1]
                    self.docInnerCanvas.move(quadelement, dx, dy) 
                for quadelement in listalinks:
                    box = (self.docInnerCanvas.bbox(quadelement))
                    pagina = math.floor(box[1]/((global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*oldzoom)))
                    deslocy = pagina * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom
                    deslocyold = pagina * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*oldzoom
                    x0novo = round(((box[0]-sobraEspacoold)/(self.zoom_x*oldzoom))*self.zoom_x*global_settings.zoom + sobraEspaco)
                    x1novo = round(((box[2]-sobraEspacoold)/(self.zoom_x*oldzoom))*self.zoom_x*global_settings.zoom + sobraEspaco)
                    y0novo = round(((box[1]-deslocyold)/(self.zoom_x*oldzoom))*self.zoom_x*global_settings.zoom  +deslocy)
                    y1novo = round(((box[3]-deslocyold)/(self.zoom_x*oldzoom))*self.zoom_x*global_settings.zoom  +deslocy)                       
                    tempimagem = (utilities_general.create_rectanglex(x0novo, y0novo, x1novo, y1novo, (175, 200, 240, 95), link=True))
                    self.addImagetoList("link", tempimagem)
                    self.docInnerCanvas.itemconfig(quadelement, image=tempimagem)
                    coords = self.docInnerCanvas.coords(quadelement)
                    dx = x0novo -coords[0]
                    dy = y0novo -coords[1]
                    self.docInnerCanvas.move(quadelement, dx, dy)               
                for quadelementx in listaobs:                    
                    box = (self.docInnerCanvas.bbox(quadelementx))                    
                    pagina = math.floor(box[1]/((global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*oldzoom)))
                    deslocy = pagina * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom
                    deslocyold = pagina * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*oldzoom
                    x0novo = round(((box[0]-sobraEspacoold)/(self.zoom_x*oldzoom))*self.zoom_x*global_settings.zoom + sobraEspaco)
                    x1novo = round(((box[2]-sobraEspacoold)/(self.zoom_x*oldzoom))*self.zoom_x*global_settings.zoom + sobraEspaco)
                    y0novo = round(((box[1]-deslocyold)/(self.zoom_x*oldzoom))*self.zoom_x*global_settings.zoom  +deslocy)
                    y1novo = round(((box[3]-deslocyold)/(self.zoom_x*oldzoom))*self.zoom_x*global_settings.zoom  +deslocy)                       
                    tempimagem = (utilities_general.create_rectanglex(x0novo, y0novo, x1novo, y1novo, self.color, link=False))
                    self.addImagetoList("obsitem", tempimagem)
                    self.docInnerCanvas.itemconfig(quadelementx, image=tempimagem)
                    coords = self.docInnerCanvas.coords(quadelementx)
                    dx = x0novo -coords[0]
                    dy = y0novo -coords[1]
                    self.docInnerCanvas.move(quadelementx, dx, dy)                
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)     
    
    def freemanipulation(self):
        self.lockmanipulation = False     
    
    def manipulatePagesByClick(self, tipo, event=None):        
        if(not self.lockmanipulation):
            self.lockmanipulation = True
            try:                
                at = round(self.vscrollbar.get()[0]*global_settings.infoLaudo[global_settings.pathpdfatual].len)
                atfloor = math.floor((self.vscrollbar.get()[0]*global_settings.infoLaudo[global_settings.pathpdfatual].len))
                if(tipo=="next"): 
                    if(at+1 > global_settings.infoLaudo[global_settings.pathpdfatual].len):
                       self.docInnerCanvas.yview_moveto(1.0)
                       self.pagVar.set(str(global_settings.infoLaudo[global_settings.pathpdfatual].len))                       
                    else:
                        if(global_settings.zoom>=1.3 and atfloor <= at):
                            self.docInnerCanvas.yview_scroll(10, "units")
                        else: 
                            self.docInnerCanvas.yview_scroll(16, "units")
                            self.pagVar.set(str(at+2))
                elif(tipo=="prev"):
                    if(at-1 <= 0):
                        self.docInnerCanvas.yview_moveto(0)                       
                        self.pagVar.set(str(1))                       
                    else:
                        if(global_settings.zoom>=1.3 and atfloor <= at):
                            self.docInnerCanvas.yview_scroll(-10, "units")
                        else:
                            self.docInnerCanvas.yview_scroll(-16, "units")
                            self.pagVar.set(str(at))
                elif(tipo=="next10"):
                    if(at+10 > global_settings.infoLaudo[global_settings.pathpdfatual].len):
                        self.docInnerCanvas.yview_moveto(1.0)
                        self.pagVar.set(str(global_settings.infoLaudo[global_settings.pathpdfatual].len))                       
                    else:
                        self.docInnerCanvas.yview_scroll(160, "units")
                        self.pagVar.set(str(at+11))                       
                elif(tipo=="prev10"):
                    if(at-10 <= 0):
                        self.docInnerCanvas.yview_moveto(0)                        
                        self.pagVar.set(str(1))                        
                    else:
                        self.docInnerCanvas.yview_scroll(-160, "units")
                        self.pagVar.set(str(at-9))                        
                elif(tipo=="first"):
                    self.docInnerCanvas.yview_moveto(0)
                    self.pagVar.set(str(1))                    
                elif(tipo=="last"):
                    self.docInnerCanvas.yview_moveto(1.0)
                    self.pagVar.set(str(global_settings.infoLaudo[global_settings.pathpdfatual].len))
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)
                atual = round((self.vscrollbar.get()[0]*global_settings.infoLaudo[global_settings.pathpdfatual].len))                
            finally:
                global_settings.root.after(50, self.freemanipulation)
    
    def ssv(self, name=None, index=None, mode=None, sv=None):
        self.docInnerCanvas.delete("simplesearch")
        self.clearSomeImages(["simplesearch"])
        
    def dosearchsimple(self, tipo, termo=""):  
        #if(self.window_search_info==None):
        #    self.window_search_info = classes_general.Search_Info_Window(global_settings.root)
        #else:
        self.window_search_info.window.deiconify() 
        self.nhp.config(relief='sunken', state='disabled')
        self.php.config(relief='sunken', state='disabled')
        
        global_settings.root.after(50, lambda: self.dosearchsimpleend(self.window_search_info, tipo, termo))
        
    def dosearchsimpleend(self, window_search_info, tipo, termo=""):
        if(not self.simplesearching):           
            try:
                sqliteconn= None
                if(termo==""):
                    termo= self.simplesearchvar.get()
                if(len(termo)<3):
                    self.window_search_info.window.withdraw()
                    utilities_general.below_right(utilities_general.popup_window('<{}> - O termo pesquisado deve possuir 3 caracteres ou mais!'.format(termo), False),\
                                                      self.docInnerCanvas.winfo_rooty())
                    return
                self.simplesearching = True
                
                novotermo = ""
                for char in termo:
                    codePoint = ord(char)
                    if(codePoint<256):
                        codePoint += global_settings.lowerCodeNoDiff[codePoint]
                    novotermo += chr(codePoint) 
                termo = novotermo
                if(termo not in self.previous_simple_searches):
                    self.previous_simple_searches.append(termo)                    
                else:
                    self.previous_simple_searches.remove(termo) 
                    self.previous_simple_searches.append(termo) 
                  
                self.showPreviousSearches()
                idpdf = global_settings.infoLaudo[global_settings.pathpdfatual].id
                rects = []
                listapintados = self.docInnerCanvas.find_withtag("simplesearch") 
                recordsx = []
                if(tipo=='prev'):
                    atual = math.floor((self.vscrollbar.get()[0]*global_settings.infoLaudo[global_settings.pathpdfatual].len))
                    if(termo in self.termossimplespesquisados and idpdf in self.termossimplespesquisados[termo]):                        
                        listapaginas = self.termossimplespesquisados[termo][idpdf]
                        pagref = None
                        for i in range(len(listapaginas)-1, -1, -1):
                            pagnow = listapaginas[i]
                            if(self.paginaSearchSimple!= int(self.pagVar.get())-1):
                                if(pagnow[0]<=int(self.pagVar.get())-1):
                                    recordsx.append(pagnow)
                                    self.paginaSearchSimple = pagnow[0]
                                    break
                            else:
                                if(pagnow[0]<int(self.pagVar.get())-1):
                                    recordsx.append(pagnow)
                                    self.paginaSearchSimple = pagnow[0]
                                    break
                    else:
                        sqliteconn =  utilities_general.connectDB(str(global_settings.pathdb), 5)
                        try:
                            if(sqliteconn==None):
                                utilities_general.popup_window("O banco de dados está ocupado.\n A operação não foi concluída, tente novamente em alguns segundos.", False)
                                return
                            cursor = sqliteconn.cursor()
                            comando = 'SELECT C.pagina, C.texto  FROM Anexo_Eletronico_Conteudo_id_pdf_'+str(idpdf)+' C where C.texto like :termo ESCAPE :escape ORDER BY 1'
                            cursor.custom_execute(comando, {'termo':'%'+termo+'%', 'escape': '\\'})
                            records2 = cursor.fetchall()
                            if(termo not in self.termossimplespesquisados):
                                self.termossimplespesquisados[termo] ={}
                            self.termossimplespesquisados[termo][idpdf] = records2
                            cursor.close()
                            listapaginas = self.termossimplespesquisados[termo][idpdf]
                            pagref = None
                            for i in range(len(listapaginas)-1, -1, -1):
                                pagnow = listapaginas[i]
                                if(self.paginaSearchSimple!= int(self.pagVar.get())-1):
                                    if(pagnow[0]<=int(self.pagVar.get())-1):
                                        recordsx.append(pagnow)
                                        self.paginaSearchSimple = pagnow[0]
                                        break
                                else:
                                    if(pagnow[0]<int(self.pagVar.get())-1):
                                        recordsx.append(pagnow)
                                        self.paginaSearchSimple = pagnow[0]
                                        break
                        except Exception as ex:
                            utilities_general.printlogexception(ex=ex)
                        finally:
                            if(sqliteconn):
                                sqliteconn.close()
                elif(tipo=='next'):
                    atual = math.ceil((self.vscrollbar.get()[0]*global_settings.infoLaudo[global_settings.pathpdfatual].len))
                    if(termo in self.termossimplespesquisados and idpdf in self.termossimplespesquisados[termo]):
                        listapaginas = self.termossimplespesquisados[termo][idpdf]
                        pagref = None
                        for i in range(len(listapaginas)):
                            pagnow = listapaginas[i]
                            if(self.paginaSearchSimple!= int(self.pagVar.get())-1):
                                if(pagnow[0]>=int(self.pagVar.get())-1):
                                    recordsx.append(pagnow)
                                    self.paginaSearchSimple = pagnow[0]
                                    break
                            else:
                                if(pagnow[0]>int(self.pagVar.get())-1):
                                    recordsx.append(pagnow)
                                    self.paginaSearchSimple = pagnow[0]
                                    break
                    else:
                        sqliteconn =  utilities_general.connectDB(str(global_settings.pathdb), 5)
                        try:
                            if(sqliteconn==None):
                                utilities_general.popup_window("O banco de dados está ocupado.\n A operação não foi concluída, tente novamente em alguns segundos.", False)
                                return
                            cursor = sqliteconn.cursor()
                            comando = 'SELECT C.pagina, C.texto FROM Anexo_Eletronico_Conteudo_id_pdf_'+str(idpdf)+' C where C.texto like :termo ESCAPE :escape ORDER BY 1 '
                            cursor.custom_execute(comando, {'termo':'%'+termo+'%', 'escape': '\\'})
                            records2 = cursor.fetchall()
                            if(termo not in self.termossimplespesquisados):
                                self.termossimplespesquisados[termo] ={}
                            self.termossimplespesquisados[termo][idpdf] = records2
                            cursor.close()
                            listapaginas = self.termossimplespesquisados[termo][idpdf]
                            pagref = None
                            for i in range(len(listapaginas)):
                                pagnow = listapaginas[i]
                                if(self.paginaSearchSimple!= int(self.pagVar.get())-1):
                                    if(pagnow[0]>=int(self.pagVar.get())-1):
                                        recordsx.append(pagnow)
                                        self.paginaSearchSimple = pagnow[0]
                                        break
                                else:
                                    if(pagnow[0]>int(self.pagVar.get())-1):
                                        recordsx.append(pagnow)
                                        self.paginaSearchSimple = pagnow[0]
                                        break
                        except Exception as ex:
                            utilities_general.printlogexception(ex=ex)
                        finally:
                            if(sqliteconn):
                                sqliteconn.close()
                try:
                    if(len(recordsx)>0):
                        results = utilities_general.searchsqlite("LIKE", termo, global_settings.pathpdfatual, global_settings.pathdb, idpdf, simplesearch=True, \
                                                                 erros_queue = global_settings.erros_queue, jarecords=recordsx)
                        if(results != None and len(results) >0):  
                            pagina = results[0].pagina
                            ondeir = ((pagina) / (global_settings.infoLaudo[global_settings.pathpdfatual].len))
                            self.docInnerCanvas.yview_moveto(ondeir)
                            if(str(pagina+1)!=self.pagVar.get()):
                               self.pagVar.set(str(pagina+1))
                            self.positions[self.indiceposition] = (global_settings.pathpdfatual, self.vscrollbar.get()[0])
                            self.indiceposition += 1
                            if(self.indiceposition>=10):
                                self.indiceposition = 0
                            self.paintsearchresult(results, True)
                        else:
                            self.simplesearching = False
                            self.nhp.config(relief='raised', state='normal')
                            self.php.config(relief='raised', state='normal')                        
                            utilities_general.below_right_edge(utilities_general.popup_window('<{}> - Nenhuma ocorrência encontrada!'.format(termo), False),\
                                                              self.docInnerCanvas.winfo_rooty())
                                                              
                    else:
                        self.window_search_info.window.withdraw() 
                        self.simplesearching = False                                                
                        utilities_general.below_right_edge(utilities_general.popup_window('<{}> - Nenhuma ocorrência encontrada!'.format(termo), False),\
                                                          self.docInnerCanvas.winfo_rooty())
                except Exception as ex:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    utilities_general.popup_window(traceback.format_exception(exc_type, exc_value, exc_tb), False)
                                  
            except Exception as ex:
                exc_type, exc_value, exc_tb = sys.exc_info()
                utilities_general.popup_window(traceback.format_exception(exc_type, exc_value, exc_tb), False)
            finally:
                self.nhp.config(relief='raised', state='normal')
                self.php.config(relief='raised', state='normal')
                self.simplesearching = False
        
    def searchTerm(self, termo=None, event=None):
        try:
            if(termo==None):
                termo = self.searchVar.get().strip()
            novotermo = ""
            for char in termo:
                codePoint = ord(char)
                if(codePoint<256):
                    codePoint += global_settings.lowerCodeNoDiff[codePoint]
                novotermo += chr(codePoint) 
            novotermo = novotermo.strip()
            tipobusca = self.searchtypeVar.get()
            if tipobusca=="REGEX":
                termo = termo
            elif tipobusca=="MATCH":
                termo = termo
            else:
                termo = novotermo
            if(not (termo.upper(), tipobusca) in global_settings.listaTERMOS):
                global_settings.listaTERMOS[(termo.upper(),tipobusca)] = []
                pedidosearch = classes_general.PedidoSearch(-math.inf, tipobusca, termo)
                global_settings.searchqueue.put(pedidosearch)
                #global_settings.searchqueue.insert(0, (termo, tipobusca, None, '0')) 
                global_settings.contador_buscas_incr += 1                
            elif(len(global_settings.listaTERMOS[termo.upper(), tipobusca])>0):
                idtermo = str(global_settings.listaTERMOS[(termo.upper(),tipobusca)][2])
                self.treeviewSearches.selection_set('t'+idtermo)
                self.treeviewSearches.move('t'+idtermo, '', 0)
                self.treeviewSearches.focus('t'+idtermo)                    
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
            
    def printer(self):
        _filename = global_settings.pathpdfatual
        widthdoc = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw
        heightdoc = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh
        atual = math.floor((self.vscrollbar.get()[0]*global_settings.infoLaudo[global_settings.pathpdfatual].len))+1
        application_path = utilities_general.get_application_path()
        try:
            if plt == "Linux":
                subprocess.Popen([os.path.join(application_path, 'printer_interface_linux', 'printer_interface'), str(_filename), str(widthdoc), str(heightdoc), str(atual)])
            elif plt=="Windows":
                subprocess.Popen([os.path.join(application_path, 'printer_interface_windows', 'printer_interface.exe'), str(_filename), str(widthdoc), str(heightdoc), str(atual)])
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        
    def createTopBar(self):      
        try:
            self.clipboardgtk = None
            self.swapframes = tkinter.Frame(bg=self.bg, highlightthickness=0)
            self.swapframes.rowconfigure(0, weight=1)
            self.swapframes.columnconfigure(0, weight=1)
            self.globalFrame.add(self.swapframes)            
            self.docOuterFrame = tkinter.Frame(self.swapframes, bg=self.bg, highlightthickness=0)
            self.docOuterFrame.grid(row=0, column=0, sticky='nsew', padx=0, pady=0)
            self.docOuterFrame.rowconfigure(1, weight=1)
            self.docOuterFrame.columnconfigure(0, weight=1)
            #self.globalFrame.paneconfig(self.swapframes, minsize=global_settings.root.winfo_screenwidth()/2)
            
            self.toolbar = tkinter.Frame(self.docOuterFrame,bg=self.bg, borderwidth=2, relief='groove')     
            
            #self.toolbar.columnconfigure((0, 1, 2, 3), weight=1)
            self.toolbar.rowconfigure(0, weight=1)
            self.toolbar.columnconfigure((0,1,2), weight=1)
            self.toolbar.grid(column=0, row=0, sticky='ew')  
            
            self.manipulationTool = tkinter.Frame(self.toolbar, bg=self.bg, borderwidth=2, relief='groove')
            #self.manipulationTool.grid(row=0, column=0, sticky='w')
            self.manipulationTool.columnconfigure((0,1,2,3,4,5,6), weight=1)
            
            self.basicTool = tkinter.Frame(self.toolbar, bg=self.bg, borderwidth=2, relief='groove')
            #self.basicTool.grid(row=0, column=1, sticky='n')
            self.basicTool.columnconfigure((0,1,2,3,4,5,6,7), weight=1)
            
            self.rightFrame  = tkinter.Frame(self.toolbar,  bg=self.bg, borderwidth=2, relief='groove')
            #self.rightFrame.grid(row=0, column=3, sticky='ew')
            self.rightFrame.columnconfigure((0,1,2,3), weight=1)    
            
            #global_settings.root.update_idletasks()
            
            #self.manipulationTool = tkinter.Frame(self.toolbar, bg=self.bg, borderwidth=4, relief='groove')
            self.manipulationTool.grid(row=0, column=0, sticky='w')
            #self.manipulationTool.columnconfigure((0,1,2,3,4,5,6,7), weight=1)
            
            #self.basicTool = tkinter.Frame(self.toolbar, bg=self.bg)  
            #self.basicTool.grid(row=0, column=1, sticky='w')
            self.basicTool.grid(row=0, column=1, sticky='n')
            #self.basicTool.columnconfigure((0,1,2,3,4,5,6,7), weight=1)
            
            #self.rightFrame  = tkinter.Frame(self.toolbar, bg=self.bg)
            self.rightFrame.grid(row=0, column=2, sticky='e')
            
            #self.manipulationTool = tkinter.Frame(self.toolbar, bg=self.bg, borderwidth=4, relief='groove')
            
            #self.manipulationTool.columnconfigure((0,1,2,3,4,5,6,7), weight=1)

            
            
            self.reportbut = tkinter.Menubutton(self.rightFrame, image=global_settings.reportsicon, relief='raised')
            self.menuReports()
            self.reportbut.grid(row=0, column=3, sticky='e')
            
            self.nhp = tkinter.Button(self.rightFrame, image=global_settings.downhiti)
            self.nhp.image = global_settings.downhiti
            self.nhp.grid(column=2, row=0, sticky='ns', padx=5) 
            self.nhp.config(command=lambda: self.dosearchsimple('next'))
            nhp = classes_general.CreateToolTip(self.nhp, "Próxima página com ocorrências")
            
            self.php = tkinter.Button(self.rightFrame, image=global_settings.uphiti)
            self.php.image = global_settings.uphiti
            self.php.grid(column=0, row=0, sticky='ns', padx=5) 
            self.php.config(command=lambda: self.dosearchsimple('prev'))
            php = classes_general.CreateToolTip(self.php, "Página anterior com ocorrências")
            
            self.simplesearchvar = tkinter.StringVar()
            self.simplesearch = tkinter.Entry(self.rightFrame, justify='center', textvariable=self.simplesearchvar, exportselection=False)
            self.simplesearchvar.set("")
            self.simplesearchvar.trace_add("write", self.ssv)
            self.simplesearch.grid(row=0, column=1, sticky='ns', padx=5)
            self.simplesearch.bind('<Return>',  lambda e: self.dosearchsimple('next'))
            self.simplesearch.bind('<Button-1>', lambda e: self.showPreviousSearches(e))
            
            zoomp = tkinter.Button(self.manipulationTool, image=global_settings.im, command= lambda: self.zoomx(tipozoom='plus'))
            zoomp.image = global_settings.im
            zoomp.grid(column=4, row=0, sticky='w', padx=(1,1))  
            zoomp_ttp = classes_general.CreateToolTip(zoomp, "Aumentar Zoom")            
            zoomm = tkinter.Button(self.manipulationTool, image=global_settings.im2, command= lambda: self.zoomx(tipozoom='minus'))
            zoomm.image = global_settings.im2
            zoomm.grid(column=5, row=0, sticky='w', padx=(1,1))  
            zoom_ttp = classes_general.CreateToolTip(zoomm, "Dominuir Zoom")
            if(plt=="Linux"):
                 printb = tkinter.Button(self.manipulationTool, image=global_settings.printeri, command= lambda: self.printer(), state='normal')
            elif(plt=='Windows'):
                 printb = tkinter.Button(self.manipulationTool, image=global_settings.printeri, command= lambda: self.printer(), state='normal')
            printb.image = global_settings.printeri
            printb.grid(column=6, row=0, sticky='w', padx=(1,1))  
            printb_ttp = classes_general.CreateToolTip(printb, "Imprimir")  

            self.lockmanipulation = False
            self.fp = tkinter.Button(self.basicTool, image=global_settings.imFP)
            self.fp.image = global_settings.imFP
            self.fp.grid(column=0, row=0, sticky='n')  
            self.fp.config(command=lambda: self.manipulatePagesByClick('first'))
            fp_ttp = classes_general.CreateToolTip(self.fp, "Ir para primeira página")
            
            self.pp10 = tkinter.Button(self.basicTool, image=global_settings.imPP10)
            self.pp10.image = global_settings.imPP10
            self.pp10.grid(column=1, row=0, sticky='n') 
            self.pp10.config(command=lambda: self.manipulatePagesByClick('prev10'))
            pp10_ttp = classes_general.CreateToolTip(self.pp10, "Voltar DEZ páginas")
            
            self.pp = tkinter.Button(self.basicTool, image=global_settings.imPP)
            self.pp.image = global_settings.imPP
            self.pp.grid(column=2, row=0, sticky='n') 
            self.pp.config(command=lambda: self.manipulatePagesByClick('prev')) 
            pp_ttp = classes_general.CreateToolTip(self.pp, "Pagina anterior")
            
            self.pagVar = tkinter.StringVar()
            self.pag = tkinter.Entry(self.basicTool, justify='right', textvariable=self.pagVar, exportselection=False)
            self.pagVar.set("1")
            self.pag.bind('<Return>', self.gotoPage)
            self.pag.grid(row=0, column=3, sticky='ns')
            
            self.totalPgg = tkinter.Label(self.basicTool, font=global_settings.Font_tuple_Arial_10, text="/ "+str(global_settings.infoLaudo[global_settings.pathpdfatual].len))
            self.totalPgg.grid(row=0, column=4, sticky='ns')
                
            self.np = tkinter.Button(self.basicTool, image=global_settings.imNP)
            self.np.image = global_settings.imNP
            self.np.grid(column=5, row=0, sticky='n') 
            self.np.config(command=lambda: self.manipulatePagesByClick('next'))
            np_ttp = classes_general.CreateToolTip(self.np, "Pagina seguinte")
            
            self.np10 = tkinter.Button(self.basicTool, image=global_settings.imNP10)
            self.np10.image = global_settings.imNP10
            self.np10.grid(column=6, row=0, sticky='n') 
            self.np10.config(command=lambda: self.manipulatePagesByClick('next10'))
            np10_ttp = classes_general.CreateToolTip(self.np10, "Avançar DEZ páginas")
            
            self.lp = tkinter.Button(self.basicTool, image=global_settings.imLP)
            self.lp.image = global_settings.imLP
            self.lp.grid(column=7, row=0, sticky='n') 
            self.lp.config(command=lambda: self.manipulatePagesByClick('last'))
            lp_ttp = classes_general.CreateToolTip(self.lp, "Ir para última página")
                  
            self.bdrag = tkinter.Button(self.manipulationTool, image=global_settings.imdrag, state='disabled', padx=20, command=self.activateDrag)
            self.bdrag.image = global_settings.imdrag
            self.bdrag.grid(column=0, row=0, sticky='w', padx=(1,1))
            bdrag_ttp = classes_general.CreateToolTip(self.bdrag, "Modo de navegação")
            self.bselect = tkinter.Button(self.manipulationTool, image=global_settings.imselect, command=self.activateSelection)
            self.bselect.image = global_settings.imselect
            self.bselect.grid(column=1, row=0, sticky='w', padx=(1,1)) 
            bselect_ttp = classes_general.CreateToolTip(self.bselect, "Modo de seleção de texto")
            self.selectionActive = False
            self.areaselectionActive = False
            self.bdrag.config(relief='sunken')
            self.baselect = tkinter.Button(self.manipulationTool, image=global_settings.imareaselect, command=self.activateareaSelection)
            self.baselect.image = global_settings.imareaselect
            baselect_ttp = classes_general.CreateToolTip(self.baselect, "Modo de seleção de área")
            self.baselect.grid(column=2, row=0, sticky='w', padx=(1,1))             
            self.showbookmarks = tkinter.Button(self.manipulationTool, image=global_settings.showbookmarksi, command=self.showAllBookmarks)
            self.showbookmarks.image = global_settings.showbookmarksi
            showbookmarks_ttp = classes_general.CreateToolTip(self.showbookmarks, "Realçar Marcadores")
            self.showbookmarks.grid(column=3, row=0, sticky='w', padx=(1,1)) 
            self.docwidth = self.docOuterFrame.winfo_width()
            self.showbookmarsboolean = True            
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)    
    
    def gotoPage(self,event):        
        page = self.pagVar.get()
        try:
            pageint = int(page)
            if(pageint > 0 and pageint<=global_settings.infoLaudo[global_settings.pathpdfatual].len):
                ondeir = (pageint-1) / global_settings.infoLaudo[global_settings.pathpdfatual].len
                self.docInnerCanvas.yview_moveto(ondeir)
                if(str(pageint)!=self.pagVar.get()):
                    self.pagVar.set(str(pageint) )
            else:
                atual = round((self.vscrollbar.get()[0]*global_settings.infoLaudo[global_settings.pathpdfatual].len))
        except Exception as ex:
            atual = round((self.vscrollbar.get()[0]*global_settings.infoLaudo[global_settings.pathpdfatual].len))
            
    def showAllBookmarks(self):
        if(self.showbookmarsboolean):
            self.showbookmarsboolean = False
            self.showbookmarks.config(relief='raised', state='normal')
            self.clearEnhanceObs()            
        else:
            self.showbookmarsboolean = True
            self.showbookmarks.config(relief='sunken', state='normal')
            self.clearEnhanceObs()
            if(global_settings.pathpdfatual in self.allobs):
                for observation in self.allobs[global_settings.pathpdfatual]:
                    None
                    if(observation.paginainit in global_settings.processed_pages and observation.paginafim in global_settings.processed_pages):
                        enhancearea = False
                        enhancetext = False
                        if(observation.tipo=='area'):
                            enhancearea = True
                        elif(observation.tipo=='texto'):
                            enhancetext = True
                        for p in range(observation.paginainit, observation.paginafim+1): 
                            if(p not in global_settings.processed_pages):
                                continue
                            posicoes_normalizadas = self.normalize_positions(p, observation.paginainit, observation.paginafim, \
                                                     observation.p0x, observation.p0y, observation.p1x, observation.p1y)
                            posicaoRealX0 = posicoes_normalizadas[0]
                            posicaoRealY0 = posicoes_normalizadas[1]
                            posicaoRealX1 = posicoes_normalizadas[2]
                            posicaoRealY1 = posicoes_normalizadas[3]
                            iiditem = observation.idobs 
                            self.prepararParaQuads(p, posicaoRealX0, posicaoRealY0, posicaoRealX1, posicaoRealY1, self.colorenhancebookmark, \
                                                   tag=['enhanceobs'+global_settings.pathpdfatual+str(p),'enhanceobs'+str(iiditem), 'enhanceobs'+str(iiditem)+str(p),'enhanceobs'], apagar=False,  \
                                                       enhancetext=enhancetext, enhancearea=enhancearea, withborder=False, alt=observation.withalt)

                    for annot in observation.annotations:
                        annotation = observation.annotations[annot]
                        if(annotation.conteudo!=''):
                            idannot = annotation.idannot
                            for p in range(annotation.paginainit, annotation.paginafim+1): 
                                if(p not in global_settings.processed_pages):
                                    continue
                                posicoes_normalizadas = self.normalize_positions(p, annotation.paginainit, annotation.paginafim, \
                                                         annotation.p0x, annotation.p0y, annotation.p1x, annotation.p1y)
                                posicaoRealX0 = posicoes_normalizadas[0]-2
                                posicaoRealY0 = posicoes_normalizadas[1]-2
                                posicaoRealX1 = posicoes_normalizadas[2]+2
                                posicaoRealY1 = posicoes_normalizadas[3]+2
                                self.prepararParaQuads(p, posicaoRealX0, posicaoRealY0, posicaoRealX1, posicaoRealY1, self.colorenhanceannotation, \
                                                       tag=['enhanceannot'+global_settings.pathpdfatual+str(p),str(iiditem)+'enhanceannot'+str(idannot)\
                                                            , 'enhanceannot'], apagar=False,  \
                                                           enhancetext=False, enhancearea=True, withborder=True, alt=False)
                                
    
    def normalize_positions(self, p, pinit, pfim, posx0, posy0, posx1, posy1):
        margemsup = (global_settings.infoLaudo[global_settings.pathpdfatual].mt/25.4)*72
        margeminf = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh-((global_settings.infoLaudo[global_settings.pathpdfatual].mb/25.4)*72)
        margemesq = (global_settings.infoLaudo[global_settings.pathpdfatual].me/25.4)*72
        margemdir = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw-((global_settings.infoLaudo[global_settings.pathpdfatual].md/25.4)*72)
        if(pinit==pfim):
            return (posx0, posy0, posx1, posy1) 
        elif(p==pinit):
            return (posx0, posy0, posx1, margeminf)
        elif(p==pfim):
            return (posx0, margemsup, posx1, posy1)
        else:
            return (posx0, margemsup, posx1, margeminf)
    
    def activateareaSelection(self, event=None):
        self.bdrag.config(relief='raised', state='normal')
        self.bselect.config(relief='raised', state='normal')
        self.baselect.config(relief='sunken', state='disabled')
        self.docFrame.config(cursor="")
        self.docInnerCanvas.config(cursor="crosshair")
        self.selectionActive = False
        self.areaselectionActive = True  
        
    
    def activateSelection(self, event=None):
        self.bselect.config(relief='sunken', state='disabled')
        self.baselect.config(relief='raised', state='normal')
        self.bdrag.config(relief='raised', state='normal')
        self.docFrame.config(cursor="")
        self.docInnerCanvas.config(cursor="xterm")
        self.areaselectionActive = False
        self.selectionActive = True
        
    def activateDrag(self, event=None):
        self.bdrag.config(relief='sunken', state='disabled')
        self.bselect.config(relief='raised', state='normal')
        self.baselect.config(relief='raised', state='normal')
        self.docFrame.config(cursor="")
        self.docInnerCanvas.config(cursor="")
        self.selectionActive = False
        self.areaselectionActive = False     
   
        
    def copiar(self, event=None):       
        doc = None
        try:
            pinit = min(global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados)
            pfim = max(global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados)
            if(self.selectionActive):
                tudo = []                
                for p in global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados:
                    if(p in global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados):
                        ultimatupla = None
                        for tupla in global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[p]['text']:
                            linha = tupla[0]
                            rect = tupla[1]
                            if(ultimatupla!=None):
                                if(ultimatupla[1].y0+2 >= rect.y0 and ultimatupla[1].y1-2 <= rect.y1):
                                    tudo.append(" ")
                                else:
                                    tudo.append("\n")
                            else:
                                tudo.append("\n")
                            for char in rect.char:
                                tudo.append(char)                            
                            ultimatupla = tupla    
                string = ''.join(tudo)
                clipboard.copy(string.strip())
            if(self.areaselectionActive):
                images = []
                for p in range(pinit, pfim+1):
                    if(p in global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados):
                        for tupla in global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[p]['areaSelection']:
                            linha = tupla[0]
                            rect = tupla[1]
                            pathpdf2 = global_settings.pathpdfatual
                            doc = fitz.open(pathpdf2)
                            loadedPage = doc[p]
                            box = fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y1)
                            zoom = 2.5
                            mat = fitz.Matrix(zoom, zoom)
                            pix = loadedPage.get_pixmap(alpha=False, matrix=mat, clip=box, dpi=1200) 
                            mode = "RGBA" if pix.alpha else "RGB"
                            
                            imgdata = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
                            pix = None
                            images.append(imgdata)
                if(len(images) > 0):
                    imagem = utilities_general.concatVertical(images)
                    if platform.system() == 'Darwin':       # macOS
                        None
                    elif platform.system() == 'Windows':    # Windows
                        output = BytesIO()
                        imagem.convert("RGB").save(output, "BMP")
                        data = output.getvalue()[14:]
                        output.close()
                        win32clipboard.OpenClipboard()
                        win32clipboard.EmptyClipboard()
                        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
                        win32clipboard.CloseClipboard()
                    elif plt == "Linux": 
                        output = BytesIO()
                        imagem.save(output, format="png")
                        clip = subprocess.Popen(("xclip", "-selection", "clipboard", "-t", "image/png", "-i"), 
                          stdin=subprocess.PIPE)
                        clip.stdin.write(output.getvalue())
                        clip.stdin.close()
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        finally:
            if(doc!=None):
                doc.close()
                
    
     
    def colectLeaves(self, treeview, item, leaves):
        children = treeview.get_children(item)
        if(len(children)==0):
            leaves.append(item)
        for child in children:
            self.colectLeaves(treeview, child, leaves)          
     
    def colectLeavesFromIteminit(self, treeview, item, leaves):
        children = treeview.get_children(item)
        if(len(children)>0):
            for child in children:
                self.colectLeavesFromIteminit(treeview, child, leaves)
        elif(item in global_settings.searchResultsDict):
            leaves.append(item)
        
    def addSeveralMarkers(self, idobscat, items):
        listadeitenscompleto = global_settings.manager.list()
        added = set()
        allitens = []
        for iteminit in items:
            leaves = []
            self.colectLeavesFromIteminit(self.treeviewSearches, iteminit, leaves)
            
            for leaf in leaves:
                if(leaf in added):
                    continue
                added.add(leaf)
                #if(global_settings.searchResultsDict[leaf].idtermopdf+str(global_settings.searchResultsDict[leaf].tptoc) in added):
                #    continue
                #added.add(global_settings.searchResultsDict[leaf].idtermopdf)
                pdf = utilities_general.get_normalized_path(global_settings.searchResultsDict[leaf].pathpdf)
                allitens.append((global_settings.searchResultsDict[leaf], global_settings.infoLaudo[pdf].mt, \
                                 global_settings.infoLaudo[pdf].mb, global_settings.infoLaudo[pdf].me, global_settings.infoLaudo[pdf].md,\
                                     global_settings.infoLaudo[pdf].pixorgw, global_settings.infoLaudo[pdf].pixorgh, idobscat))
        addserveralobs = mp.Process(target=process_functions.processBatchInsertObs, args=(listadeitenscompleto, allitens,), daemon=True)
        addserveralobs.start()  
        addingmarkers = tkinter.Toplevel()
        x = global_settings.root.winfo_x()
        y = global_settings.root.winfo_y()
        addingmarkers.geometry("+%d+%d" % (x + 10, y + 10))
        addingmarkers.rowconfigure((0,1), weight=1)
        addingmarkers.columnconfigure(0, weight=1)
        label = tkinter.Label(addingmarkers, font=global_settings.Font_tuple_Arial_10, text="Adicionando marcadores!", image=global_settings.processing, compound='left')
        label.image = global_settings.processing
        label.grid(row=0, column=0, sticky='ew', pady=5, padx=5)
        progressindex = ttk.Progressbar(addingmarkers, mode='indeterminate')
        progressindex.grid(row=1, column=0, sticky='ew', pady=5)
        progressindex.start()
        self.checkWhenAddSeveralIsDone(addserveralobs, listadeitenscompleto, addingmarkers, None)
        
  
        
    
        
    def qualIndexTreeObs(self, paginaAinserir, p0yarg, imediateParent, name=None):
   
        children = self.treeviewObs.get_children(imediateParent)
        index = 0
            
        if(len(children)>0):
            for child in children:
                valores = self.treeviewObs.item(child, 'values')
                texto = self.treeviewObs.item(child, 'text')
                
                if(name!=None):
                    if(texto > name):
                        return index
                    else:
                        index += 1
                else:
                    pagina = int(valores[2])
                    p0y = float(valores[4])
                    if(pagina == int(paginaAinserir) and p0y > p0yarg):
                        return index
                    elif(pagina > int(paginaAinserir)):
                        return index
                    else:
                        index += 1
            return index
                
        else:
            return index
        
        
    def checkWhenAddSeveralIsDone(self, processo, lista, addingmarkers, doctemppdf):
        if(processo.is_alive()):
            global_settings.root.after(1000, lambda: self.checkWhenAddSeveralIsDone(processo, lista, addingmarkers, doctemppdf))
        else:
            try:
                addingmarkers.destroy()
            except:
                None
            try:
                doctemp = None
                sqliteconn =  utilities_general.connectDB(str(global_settings.pathdb), 5)
                cursor = None
                try:
                    if(sqliteconn==None):
                        utilities_general.popup_window("O banco de dados está ocupado.\n A operação não foi concluída, tente novamente em alguns segundos.", False)
                        return
                    processo.join()  
                    
                    cursor = sqliteconn.cursor() 
                    #listaobsitens_add = []
                    for itensproc in lista:
                        idobscat = itensproc[6]
                        textoobscat = self.treeviewObs.item(idobscat, 'values')[2]  
                        p0x = itensproc[0]
                        p0y = itensproc[1]
                        p1x = itensproc[2]
                        p1y = itensproc[3]
                        tipo = 'texto'
                        paginainit = itensproc[4]
                        paginafim  = itensproc[4]
                        
                        cursor.custom_execute("PRAGMA foreign_keys = ON")
                        iid = idobscat
                        insert_query_pdf = """INSERT INTO Anexo_Eletronico_Obsitens
                                                (id_obscat, id_pdf, paginainit, p0x, p0y, paginafim, p1x, p1y, tipo, fixo, conteudo) VALUES
                                                (?,?,?,?,?,?,?,?,?,?,?)
                        """
                        fixo = 0
                        if(True):
                            fixo = 1
                        id_pdf = global_settings.infoLaudo[utilities_general.get_normalized_path(itensproc[5])].id
                         
                        cursor.custom_execute("PRAGMA journal_mode=WAL")
                        cursor.custom_execute(insert_query_pdf, (iid, id_pdf, paginainit, p0x, p0y, paginafim, p1x, p1y, tipo, fixo, '',))
                        iiditem = str(cursor.lastrowid)
                        #cursor.close()
                        pathpdf = utilities_general.get_normalized_path(itensproc[5])
                        if(pathpdf in global_settings.infoLaudo and pathpdf not in self.allobs):
                            self.allobs[pathpdf] = []
                            
                        #treeobs_insert(self, tipo, fixo, paginainit, paginafim, basepdf, p0x, p0y, p1x, p1y, tocpdf, idobscat, idobsitem, ident)
                        tocpdf = global_settings.infoLaudo[pathpdf].toc
                        self.treeobs_insert(tipo, 1, paginainit, paginafim, pathpdf, p0x, p0y, p1x, p1y, tocpdf, iid, iiditem,ident = ' ')
                        
                        if(pathpdf!=doctemppdf):
                            if(doctemp!=None):
                                doctemp.close()
                                
                            doctemp = fitz.open(pathpdf)
                            doctemppdf = pathpdf
                        links_tratados = utilities_general.extract_links_from_page(doctemp, id_pdf, iiditem, global_settings.pathpdfatual,\
                                                                                   paginainit, paginafim, p0x, p0y, p1x, p1y)
                        insert_annot = '''INSERT INTO Anexo_Eletronico_Annotations
                                                 (id_pdf, id_obs, paginainit, p0x, p0y, paginafim, p1x, p1y, link, conteudo) VALUES
                                                 (?,?,?,?,?,?,?,?,?,?)'''
                        #cursor.custom_execute("PRAGMA journal_mode=WAL")
                        annotation_list = {}
                        for link_t in links_tratados:
                            cursor.custom_execute(insert_annot, link_t)
                            idannot = cursor.lastrowid                            
                            paginainit = link_t[2]
                            paginafim = link_t[5]
                            p0xa = link_t[3]
                            p0ya = link_t[4]
                            p1xa = link_t[6]
                            p1ya = link_t[7]
                            iiditem = iiditem
                            conteudo = ''
                            link = link_t[8]
                            annots_object = classes_general.Annotation(pathpdf, idannot, paginainit, paginafim, p0xa, p0ya, p1xa, p1ya, iiditem, conteudo, link)
                            annotation_list[idannot] = annots_object
                        textoobscat = self.treeviewObs.item(idobscat, 'values')[2]
                        # def __init__(self, paginainit, paginafim, p0x, p0y, p1x, p1y, tipo, pathpdf, idobs, idobscat, obscat, conteudo, anottations):
                        obsobject = classes_general.Observation(paginainit, paginafim, p0x, p0y, p1x, p1y, tipo,\
                                                                pathpdf, iiditem, iid, '', annotation_list)
                        self.allobs[pathpdf].append(obsobject)
                        self.allobsbyitem['obsitem'+str(iiditem)] = obsobject
                        
                    #cursor.close()    
                    sqliteconn.commit()
                except Exception as ex:
                    utilities_general.printlogexception(ex=ex)
                finally:
                    #cursor.close()
                    if(doctemp!=None):
                        doctemp.close()
                    if(sqliteconn):
                        sqliteconn.close()
            except sqlite3.OperationalError:
                global_settings.root.after(1000, lambda: self.checkWhenAddSeveralIsDone(processo, lista, doctemppdf))                                           
    
    def treeobs_insert(self, tipo, fixo, paginainit, paginafim, basepdf, p0x, p0y, p1x, p1y, tocpdf, idobscat, idobsitem, ident, tag_alt=''):
       
        idobscat = str(idobscat)
        tags_to_item =None
        if(tag_alt!=''):
            tags_to_item = ('obsitem', 'obsitem'+str(idobsitem), tag_alt,)
        else:
            tags_to_item = ('obsitem', 'obsitem'+str(idobsitem),)
        #idobsitem = str(idobsitem)
        try:
            tocx = utilities_general.locateToc(paginainit, basepdf, p0y=p0y, tocpdf=tocpdf)
            tocname = tocx[0]
            if(not self.treeviewObs.exists(idobscat+basepdf)):
                indexrelobs = self.qualIndexTreeObs( None, None,idobscat, ident+os.path.basename(basepdf))
                self.treeviewObs.insert(parent=idobscat, iid=idobscat+basepdf, text=ident+os.path.basename(basepdf), index=indexrelobs, image=global_settings.imagereportb, tag=('relobs'), values=(basepdf,))
                self.treeviewObs.tag_configure('relobs', background='#e3e1e1')
            if(not self.treeviewObs.exists(idobscat+basepdf+tocname)):
                indextocobs = self.qualIndexTreeObs(tocx[1], tocx[2],idobscat+basepdf)
                self.treeviewObs.insert(parent=idobscat+basepdf, iid=idobscat+basepdf+tocname, text=ident+ident+tocname,\
                                        index=indextocobs, tag=('tocobs'), values=(0,0,tocx[1],0, tocx[2], basepdf))
            indexinserir = self.qualIndexTreeObs(paginainit, p0y, (idobscat+basepdf+tocname))
            
            if(paginainit==paginafim):
                labeltext = ident+ident+ident+'Pg.'+str(paginainit+1)+' - '+ os.path.basename(basepdf)
                self.treeviewObs.insert(parent=(idobscat+basepdf+tocname), index=indexinserir, iid='obsitem'+str(idobsitem), text=labeltext,\
                                image=global_settings.itemimage, \
                                    values=(tipo, basepdf,str(paginainit), str(p0x), str(p0y), str(paginafim), str(p1x), str(p1y), idobsitem, str(fixo), idobscat, '',),\
                                    tag=tags_to_item)
            else:
                self.treeviewObs.insert(parent=(idobscat+basepdf+tocname), index=indexinserir, iid='obsitem'+str(idobsitem), \
                                        text=ident+ident+ident+'Pg.'+str(paginainit+1)+' - '+'Pg.'+str(paginafim+1)+' - '+os.path.basename(basepdf), \
                                image=global_settings.itemimage, \
                                    values=(tipo, basepdf,str(paginainit), str(p0x), str(p0y), str(paginafim), str(p1x), str(p1y), idobsitem, str(fixo), idobscat, '',), \
                                            tag=tags_to_item)
        except Exception as ex:
            #utilities_general.printlogexception(ex=ex)
            if(not self.treeviewObs.exists(idobscat+basepdf)):
                indexrelobs = self.qualIndexTreeObs( None, None,idobscat, os.path.basename(basepdf))
                self.treeviewObs.insert(parent=idobscat, iid=(idobscat+basepdf), text=os.path.basename(basepdf), index=indexrelobs, image=global_settings.imagereportb, tag=('relobs'), values=(basepdf,))
                self.treeviewObs.tag_configure('relobs', background='#e3e1e1')
            indexinserir = self.qualIndexTreeObs( paginainit, p0y, (idobscat+basepdf))
            
            if(paginainit==paginafim):
                self.treeviewObs.insert(parent=(idobscat+basepdf), index=indexinserir, iid='obsitem'+str(idobsitem), text=ident+ident+'Pg.'+str(paginainit+1)+' - '+\
                                        os.path.basename(basepdf), \
                                image=global_settings.itemimage, \
                                    values=(tipo, basepdf,str(paginainit), str(p0x), str(p0y), str(paginafim), str(p1x), str(p1y), idobsitem, str(fixo), idobscat, '',),\
                                        tag=tags_to_item)
            else:
                self.treeviewObs.insert(parent=(idobscat+basepdf), index=indexinserir, iid='obsitem'+str(idobsitem), text=ident+ident+'Pg.'+str(paginainit+1)\
                                        +' - '+'Pg.'+str(paginafim+1)+' - '+os.path.basename(basepdf), \
                                image=global_settings.itemimage,\
                                    values=(tipo, basepdf,str(paginainit), str(p0x), str(p0y), str(paginafim), str(p1x), str(p1y), idobsitem, str(fixo), idobscat, '',),\
                                       tag=tags_to_item )
    def get_result_searches(self, item, leaves):
        if(self.treeviewSearches.tag_has('resultsearch', item)):
            if(item not in leaves):
                leaves[item] = []
        else:
            children = self.treeviewSearches.get_children(item)
            for child in children:
                self.get_result_searches(child, leaves)
                
    def addSeveralMarkersVersion2(self, idobscat_leaves):
        listadeitenscompleto = global_settings.manager.list()
        allitens = []
        for idobscat in idobscat_leaves:
            for leaf in idobscat_leaves[idobscat]:
                pdf = utilities_general.get_normalized_path(global_settings.searchResultsDict[leaf].pathpdf)
                allitens.append((global_settings.searchResultsDict[leaf], global_settings.infoLaudo[pdf].mt, \
                                 global_settings.infoLaudo[pdf].mb, global_settings.infoLaudo[pdf].me, global_settings.infoLaudo[pdf].md,\
                                     global_settings.infoLaudo[pdf].pixorgw, global_settings.infoLaudo[pdf].pixorgh, idobscat))
        addserveralobs = mp.Process(target=process_functions.processBatchInsertObs, args=(listadeitenscompleto, allitens,), daemon=True)
        addserveralobs.start()  
        addingmarkers = tkinter.Toplevel()
        x = global_settings.root.winfo_x()
        y = global_settings.root.winfo_y()
        addingmarkers.geometry("+%d+%d" % (x + 10, y + 10))
        addingmarkers.rowconfigure((0,1), weight=1)
        addingmarkers.columnconfigure(0, weight=1)
        label = tkinter.Label(addingmarkers, font=global_settings.Font_tuple_Arial_10, text="Adicionando marcadores!", image=global_settings.processing, compound='left')
        label.image = global_settings.processing
        label.grid(row=0, column=0, sticky='ew', pady=5, padx=5)
        progressindex = ttk.Progressbar(addingmarkers, mode='indeterminate')
        progressindex.grid(row=1, column=0, sticky='ew', pady=5)
        progressindex.start()
        self.checkWhenAddSeveralIsDone(addserveralobs, listadeitenscompleto, addingmarkers, None)
                
    def addmarkersFromSearch(self, event):
        sqliteconn = None
        try:
            iids = self.treeviewSearches.selection()
            novas_obs = {}
            novas_obs_v2 = {}
            for iid in iids:
                if(self.treeviewSearches.tag_has('resultsearch', iid)):
                    resultsearch = global_settings.searchResultsDict[iid]
                    if(resultsearch.termo not in novas_obs):
                        novas_obs[resultsearch.termo] = []
                    if(iid not in novas_obs[resultsearch.termo]):
                        novas_obs[resultsearch.termo].append(iid)
                else:
                    leafs = {}
                    self.get_result_searches(iid, leafs)
                    for leaf in leafs:
                        if(self.treeviewSearches.tag_has('resultsearch', leaf)):
                            resultsearch = global_settings.searchResultsDict[leaf]
                            if(resultsearch.termo not in novas_obs):
                                novas_obs[resultsearch.termo] = []
                            if(leaf not in novas_obs[resultsearch.termo]):
                                novas_obs[resultsearch.termo].append(leaf)
            sqliteconn = utilities_general.connectDB(str(global_settings.pathdb), 5)
            if(sqliteconn==None):
                utilities_general.popup_window("O banco de dados está ocupado.\n A operação não foi concluída, tente novamente em alguns segundos.", False)
                return
            cursor = sqliteconn.cursor()
            for key in novas_obs:
                
                newcat = key
                insertinto =  "INSERT INTO Anexo_Eletronico_Obscat (obscat, fixo, ordem) values (?,?,0)"
                fixo = 0
                if(True):
                    fixo = 1
                cursor.custom_execute(insertinto, (newcat.upper(), fixo,))
                idnovo = cursor.lastrowid
                sqliteconn.commit()
                novas_obs_v2[idnovo] = novas_obs[key]
                self.treeviewObs.insert(parent='', index='end', iid=idnovo, text=newcat.upper(), values=(str(fixo), idnovo, newcat,), \
                                    image=global_settings.catimage, tag='obscat')
                self.treeviewObs.tag_configure('obscat', background='#a1a1a1', font=global_settings.Font_tuple_ArialBoldUnderline_12)
                
            self.addSeveralMarkersVersion2(novas_obs_v2)
            
            
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
    
    def addmarkerFromSearch(self, idobscat, event, first=True, use_name = False):
        if(idobscat==None):
           idobscat, obscat  = self.add_edit_category('add','')
        if(idobscat==None):
           return
        items = self.treeviewSearches.selection()
        item = self.treeviewSearches.identify_row(event.y)
        children = utilities_general.countChildren(self.treeviewSearches, item, putcount=False)  
        #print(children)
        if(len(items)>1):
            #for item in items:
            self.addSeveralMarkers(idobscat, items)
        elif(children>0):# and first):
            self.addSeveralMarkers(idobscat, [item])
        else:
            resultsearch = global_settings.searchResultsDict[self.treeviewSearches.identify_row(event.y)]
            pagina = int(resultsearch.pagina)            
            if(pagina not in global_settings.infoLaudo[global_settings.pathpdfatual].quadspagina):
                if(first or pagina in global_settings.processed_pages):
                    global_settings.root.after(100, lambda: self.addmarkerFromSearch(idobscat, event, first=False, use_name=use_name))
            else:
                try:
                    sqliteconn =  utilities_general.connectDB(str(global_settings.pathdb), 5)
                    cursor = sqliteconn.cursor() 
                    try:
                        if(sqliteconn==None):
                            utilities_general.popup_window("O banco de dados está ocupado.\n A operação não foi concluída, tente novamente em alguns segundos.", False)
                            return
                        sobraEspaco = 0
                        if(self.docFrame.winfo_width() > global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw * self.zoom_x * global_settings.zoom):
                            sobraEspaco = self.docInnerCanvas.winfo_x() 
                        posicoes = global_settings.infoLaudo[global_settings.pathpdfatual].quadspagina[pagina]
                        init = posicoes[resultsearch.init]
                        fim = posicoes[resultsearch.fim-1]
                        p0x = round(init[0])
                        p0y = round((init[1]+init[3])/2)
                        p1x = round(fim[2])
                        p1y = round((fim[1]+fim[3])/2)
                        tipo = 'texto'
                        paginainit = pagina
                        paginafim  = pagina    
                        
                        cursor.custom_execute("PRAGMA foreign_keys = ON")
                        insert_query_pdf = """INSERT INTO Anexo_Eletronico_Obsitens
                                                (id_obscat, id_pdf, paginainit, p0x, p0y, paginafim, p1x, p1y, tipo, fixo, conteudo) VALUES
                                                (?,?,?,?,?,?,?,?,?,?,?)
                        """
                        fixo = 0
                        if(True):
                            fixo = 1
                        id_pdf = global_settings.infoLaudo[utilities_general.get_normalized_path(resultsearch.pathpdf)].id
                        
                        cursor.custom_execute("PRAGMA journal_mode=WAL")
                        cursor.custom_execute(insert_query_pdf, (idobscat, id_pdf, paginainit, p0x, p0y, paginafim, p1x, p1y, tipo, fixo, '',))
                        iiditem = str(cursor.lastrowid)
                        #cursor.close()
                        basepdf = utilities_general.get_normalized_path(resultsearch.pathpdf)
                        ident = ' '
                        pathpdf = utilities_general.get_normalized_path(basepdf)
                        if(pathpdf in global_settings.infoLaudo and pathpdf not in self.allobs):
                            self.allobs[pathpdf] = []
                        links_tratados = utilities_general.extract_links_from_page(global_settings.docatual, id_pdf, iiditem, global_settings.pathpdfatual,\
                                                                                   paginainit, paginafim, p0x, p0y, p1x, p1y)
                        insert_annot = '''INSERT INTO Anexo_Eletronico_Annotations
                                                 (id_pdf, id_obs, paginainit, p0x, p0y, paginafim, p1x, p1y, link, conteudo) VALUES
                                                 (?,?,?,?,?,?,?,?,?,?)'''
                        #cursor.custom_execute("PRAGMA journal_mode=WAL")
                        annotation_list = {}
                        for link_t in links_tratados:
                            cursor.custom_execute(insert_annot, link_t)
                            idannot = cursor.lastrowid                            
                            paginainit = link_t[2]
                            paginafim = link_t[5]
                            p0xa = link_t[3]
                            p0ya = link_t[4]
                            p1xa = link_t[6]
                            p1ya = link_t[7]
                            iiditem = iiditem
                            conteudo = ''
                            link = link_t[8]
                            annots_object = classes_general.Annotation(pathpdf, idannot, paginainit, paginafim, p0xa, p0ya, p1xa, p1ya, iiditem, conteudo, link)
                            annotation_list[idannot] = annots_object
                        textoobscat = self.treeviewObs.item(idobscat, 'values')[2]
                   # def __init__(self, paginainit, paginafim, p0x, p0y, p1x, p1y, tipo, pathpdf, idobs, idobscat, obscat, conteudo, anottations):
                        obsobject = classes_general.Observation(paginainit, paginafim, p0x, p0y, p1x, p1y, tipo,\
                                                                pathpdf, iiditem, idobscat, '', annotation_list)
                        self.allobs[pathpdf].append(obsobject)
                        self.allobsbyitem['obsitem'+str(iiditem)] = obsobject
                        tocpdf = global_settings.infoLaudo[pathpdf].toc
                        #treeobs_insert(self, tipo, fixo, paginainit, paginafim, basepdf, p0x, p0y, p1x, p1y, tocpdf, idobscat, idobsitem, ident)
                        self.treeobs_insert(tipo, 1, paginainit, paginafim, pathpdf, p0x, p0y, p1x, p1y, tocpdf, int(idobscat), int(iiditem),ident = ' ')
 
                        sqliteconn.commit()
                    except Exception as ex:
                        utilities_general.printlogexception(ex=ex)                            
                    finally:
                        if(sqliteconn):
                            sqliteconn.close()
                except sqlite3.OperationalError:
                    global_settings.root.after(1000, lambda: self.addmarkerFromSearch(idobscat, event, first=False, use_name=use_name))    

    def editmarkercontent(self, newcontent, idobs, arquivo):
        updateinto2 = "UPDATE Anexo_Eletronico_Obsitens set conteudo = ? WHERE id_obs = ?"
        sqliteconn =  utilities_general.connectDB(str(global_settings.pathdb), 5)
        try:
            if(sqliteconn==None):
                utilities_general.popup_window("O banco de dados está ocupado.\n A operação não foi concluída, tente novamente em alguns segundos.", False)
                return
            valores = self.treeviewObs.item(idobs, 'values')
            cursor = sqliteconn.cursor() 
            cursor.custom_execute(updateinto2, (newcontent, valores[8],))
            sqliteconn.commit()            
            tupla = (valores[0],valores[1],valores[2],valores[3],valores[4],valores[5],valores[6],valores[7],valores[8],valores[9],valores[10],newcontent)
            self.treeviewObs.item(idobs, values=tupla)
            self.allobsbyitem[idobs].conteudo = newcontent
            pagina = self.allobsbyitem[idobs].paginainit
            sobraEspaco = 0
            if(self.docFrame.winfo_width() > global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw * self.zoom_x * global_settings.zoom):
                sobraEspaco = self.docInnerCanvas.winfo_x()
            deslocy = (math.floor(pagina) * global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom)# + (posicaoRealY0 *  self.zoom_x * global_settings.zoom) 

                #classes_general.ObsTooltip(self.docInnerCanvas, 'enhanceobs'+str(idobs), texto, newcontent, self.allobsbyitem['obsitem'+str(idobs)])
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)                
        finally:    
            try:
                cursor.close() 
            except Exception as ex:
                None
            try:
                sqliteconn.close()
            except Exception as ex:
                None
                
    def addmarker2(self, relatorio, doc, idobscat, id_pdf, pagina, p0x, p0y, p1x, p1y):
        sqliteconn =  utilities_general.connectDB(str(global_settings.pathdb), 5)
        try:
            conteudo='' 
            withalt = True
            tipo = 'area' 
            if(sqliteconn==None):
                utilities_general.popup_window("O banco de dados está ocupado.\n A operação não foi concluída, tente novamente em alguns segundos.", False)
                return
            cursor = sqliteconn.cursor()   
            cursor.custom_execute("PRAGMA journal_mode=WAL")
            insert_query_pdf = """INSERT INTO Anexo_Eletronico_Obsitens
                                    (id_obscat, id_pdf, paginainit, p0x, p0y, paginafim, p1x, p1y, tipo, fixo, conteudo, withalt) VALUES
                                    (?,?,?,?,?,?,?,?,?,?,?,?)
            """
            
            relpath = os.path.relpath(relatorio, global_settings.pathdb.parent)
            #id_pdf = global_settings.infoLaudo[global_settings.pathpdfatual].id
            
            cursor.custom_execute(insert_query_pdf, (idobscat, id_pdf, pagina, p0x, p0y, pagina, p1x, p1y, tipo, 1, conteudo, 1 if withalt else False,))
            iiditem = str(cursor.lastrowid)
            links_tratados = utilities_general.extract_links_from_page(doc, id_pdf, iiditem, relatorio,\
                                                                       pagina, pagina, p0x, p0y, p1x, p1y)
            
            insert_annot = '''INSERT INTO Anexo_Eletronico_Annotations
                                     (id_pdf, id_obs, paginainit, p0x, p0y, paginafim, p1x, p1y, link, conteudo) VALUES
                                     (?,?,?,?,?,?,?,?,?,?)'''
            cursor.custom_execute("PRAGMA journal_mode=WAL")
            annotation_list = {}
            paginainit = pagina
            paginafim = pagina
            for link_t in links_tratados:
                cursor.custom_execute(insert_annot, link_t)
                idannot = cursor.lastrowid                            
                paginainit = link_t[2]
                paginafim = link_t[5]
                p0xa = link_t[3]
                p0ya = link_t[4]
                p1xa = link_t[6]
                p1ya = link_t[7]
                iiditem = iiditem
                conteudo = ''
                link = link_t[8]
                annots_object = classes_general.Annotation(relatorio, idannot, paginainit, paginafim, p0xa, p0ya, p1xa, p1ya, iiditem, conteudo, link)
                annotation_list[idannot] = annots_object
            textoobscat = self.treeviewObs.item(idobscat, 'values')[2]
       # def __init__(self, paginainit, paginafim, p0x, p0y, p1x, p1y, tipo, pathpdf, idobs, idobscat, obscat, conteudo, anottations):
            obsobject = classes_general.Observation(paginainit, paginafim, p0x, p0y, p1x, p1y, tipo,\
                                                    relatorio, iiditem, idobscat, conteudo, annotation_list, withalt)
            cursor.close() 
            sqliteconn.commit()
            try:
                sqliteconn.close()
            except Exception as ex:
                None
            enhancearea = False
            enhancetext = False
            if(tipo=='area'):
                enhancearea = True
            elif(tipo=='texto'):
                enhancetext = True
            for p in range(paginainit, paginafim+1): 
                posicoes_normalizadas = self.normalize_positions(p, paginainit, paginafim, \
                                         p0x, p0y, p1x, p1y)
                posicaoRealX0 = posicoes_normalizadas[0]
                posicaoRealY0 = posicoes_normalizadas[1]
                posicaoRealX1 = posicoes_normalizadas[2]
                posicaoRealY1 = posicoes_normalizadas[3]
                tags = ['enhanceobs'+relatorio+str(p),'enhanceobs'+str(iiditem), 'enhanceobs'+str(iiditem)+str(p), 'enhanceobs']
                self.prepararParaQuads(p, posicaoRealX0, posicaoRealY0, posicaoRealX1, posicaoRealY1, color=self.colorenhancebookmark, \
                                       tag=tags, apagar=False, \
                                           enhancetext=enhancetext, enhancearea=enhancearea, withborder=False, alt=withalt)
            basepdf = utilities_general.get_normalized_path(relatorio)
            pathpdf = utilities_general.get_normalized_path(basepdf)
            if(pathpdf in global_settings.infoLaudo and pathpdf not in self.allobs):
                self.allobs[pathpdf] = []  
            self.allobs[relatorio].append(obsobject)
            self.allobsbyitem['obsitem'+str(iiditem)] = obsobject
            tocpdf = global_settings.infoLaudo[pathpdf].toc
            #treeobs_insert(self, tipo, fixo, paginainit, paginafim, basepdf, p0x, p0y, p1x, p1y, tocpdf, idobscat, idobsitem, ident)
            self.treeobs_insert(tipo, 1, paginainit, paginafim, pathpdf, p0x, p0y, p1x, p1y, tocpdf, int(idobscat), int(iiditem),ident = ' ')
            return str(iiditem)
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)  
            return None
        finally:    
            try:
                cursor.close() 
            except Exception as ex:
                None
            try:
                sqliteconn.close()
            except Exception as ex:
                None
        
    def addmarker(self, idobscat):
        
        conteudo=''    
        if(idobscat==None):
           idobscat, obscat  = self.add_edit_category('add','')
        if(idobscat==None):
           return
        tipo = None
        p0x = None
        p0y = None
        p1x = None
        p1y = None
        paginainit = min(global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados)
        paginafim = max(global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados)
        #print(paginainit, paginafim)
        withalt = False
        if(len(self.docInnerCanvas.find_withtag('withalt'))):
            withalt = True
        if(self.selectionActive):
            tipo = 'texto'
            pagina =paginainit
            p0x = 1000000000000
            p0y = 1000000000000
            p1x = -100000000000
            p1y = -1000000000000
            for tupla in global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[pagina]['text']:
                if(tupla[1].y0 <= p0y):
                    p0x = min(p0x, tupla[1].x0)
                    p0y = math.floor(tupla[1].y0)
            pagina2 = paginafim
            for tupla in global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[pagina2]['text']:
                if(tupla[1].y1>= p1y):
                    p1x = tupla[1].x1
                    p1y = math.ceil(tupla[1].y1)
            
        elif(self.areaselectionActive):
            p0x = global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[paginainit]['areaSelection'][0][1].x0
            p0y = global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[paginainit]['areaSelection'][0][1].y0         
            p1x = global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[paginafim]['areaSelection'][0][1].x1
            p1y = global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[paginafim]['areaSelection'][0][1].y1 
            tipo = 'area' 

        try:
            sqliteconn =  utilities_general.connectDB(str(global_settings.pathdb), 5)
            try:
                if(sqliteconn==None):
                    utilities_general.popup_window("O banco de dados está ocupado.\n A operação não foi concluída, tente novamente em alguns segundos.", False)
                    return
                cursor = sqliteconn.cursor()   
                cursor.custom_execute("PRAGMA journal_mode=WAL")
                insert_query_pdf = """INSERT INTO Anexo_Eletronico_Obsitens
                                        (id_obscat, id_pdf, paginainit, p0x, p0y, paginafim, p1x, p1y, tipo, fixo, conteudo, withalt) VALUES
                                        (?,?,?,?,?,?,?,?,?,?,?,?)
                """
                p0y = math.floor(p0y)
                p1y = math.ceil(p1y)
                relpath = os.path.relpath(global_settings.pathpdfatual, global_settings.pathdb.parent)
                id_pdf = global_settings.infoLaudo[global_settings.pathpdfatual].id
                #print(idobscat, id_pdf, paginainit, p0x, p0y, paginafim, p1x, p1y, tipo, 1)
                cursor.custom_execute(insert_query_pdf, (idobscat, id_pdf, paginainit, p0x, p0y, paginafim, p1x, p1y, tipo, 1, conteudo, 1 if withalt else False,))
                iiditem = str(cursor.lastrowid)
                links_tratados = utilities_general.extract_links_from_page(global_settings.docatual, id_pdf, iiditem, global_settings.pathpdfatual,\
                                                                           paginainit, paginafim, p0x, p0y, p1x, p1y)
                
                insert_annot = '''INSERT INTO Anexo_Eletronico_Annotations
                                         (id_pdf, id_obs, paginainit, p0x, p0y, paginafim, p1x, p1y, link, conteudo) VALUES
                                         (?,?,?,?,?,?,?,?,?,?)'''
                cursor.custom_execute("PRAGMA journal_mode=WAL")
                annotation_list = {}
                for link_t in links_tratados:
                    #print(link_t)
                    cursor.custom_execute(insert_annot, link_t)
                    idannot = cursor.lastrowid                            
                    paginainit = link_t[2]
                    paginafim = link_t[5]
                    p0xa = link_t[3]
                    p0ya = link_t[4]
                    p1xa = link_t[6]
                    p1ya = link_t[7]
                    iiditem = iiditem
                    conteudo = ''
                    link = link_t[8]
                    annots_object = classes_general.Annotation(global_settings.pathpdfatual, idannot, paginainit, paginafim, p0xa, p0ya, p1xa, p1ya, iiditem, conteudo, link)
                    annotation_list[idannot] = annots_object
                textoobscat = self.treeviewObs.item(idobscat, 'values')[2]
           # def __init__(self, paginainit, paginafim, p0x, p0y, p1x, p1y, tipo, pathpdf, idobs, idobscat, obscat, conteudo, anottations):
                obsobject = classes_general.Observation(paginainit, paginafim, p0x, p0y, p1x, p1y, tipo,\
                                                        global_settings.pathpdfatual, iiditem, idobscat, conteudo, annotation_list, withalt)
                cursor.close() 
                sqliteconn.commit()
                try:
                    sqliteconn.close()
                except Exception as ex:
                    None
                enhancearea = False
                enhancetext = False
                if(tipo=='area'):
                    enhancearea = True
                elif(tipo=='texto'):
                    enhancetext = True
                for p in range(paginainit, paginafim+1): 
                    posicoes_normalizadas = self.normalize_positions(p, paginainit, paginafim, \
                                             p0x, p0y, p1x, p1y)
                    posicaoRealX0 = posicoes_normalizadas[0]
                    posicaoRealY0 = posicoes_normalizadas[1]
                    posicaoRealX1 = posicoes_normalizadas[2]
                    posicaoRealY1 = posicoes_normalizadas[3]
                    tags = ['enhanceobs'+global_settings.pathpdfatual+str(p),'enhanceobs'+str(iiditem), 'enhanceobs'+str(iiditem)+str(p), 'enhanceobs']
                    self.prepararParaQuads(p, posicaoRealX0, posicaoRealY0, posicaoRealX1, posicaoRealY1, color=self.colorenhancebookmark, \
                                           tag=tags, apagar=False, \
                                               enhancetext=enhancetext, enhancearea=enhancearea, withborder=False, alt=withalt)
                basepdf = utilities_general.get_normalized_path(global_settings.pathpdfatual)
                pathpdf = utilities_general.get_normalized_path(basepdf)
                if(pathpdf in global_settings.infoLaudo and pathpdf not in self.allobs):
                    self.allobs[pathpdf] = []  
                self.allobs[global_settings.pathpdfatual].append(obsobject)
                self.allobsbyitem['obsitem'+str(iiditem)] = obsobject
                tocpdf = global_settings.infoLaudo[pathpdf].toc
                #treeobs_insert(self, tipo, fixo, paginainit, paginafim, basepdf, p0x, p0y, p1x, p1y, tocpdf, idobscat, idobsitem, ident)
                self.treeobs_insert(tipo, 1, paginainit, paginafim, pathpdf, p0x, p0y, p1x, p1y, tocpdf, int(idobscat), int(iiditem),ident = ' ')
                return str(iiditem)
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)  
                return None
            finally:    
                try:
                    cursor.close() 
                except Exception as ex:
                    None
                try:
                    sqliteconn.close()
                except Exception as ex:
                    None
        except sqlite3.OperationalError:
            global_settings.root.after(1000, lambda: self.addmarker(idobscat=idobscat, p0x = p0x, p0y = p0y, p1x = p1x, p1y = p1y, paginainit = paginainit, paginafim = paginafim, tipo = tipo))    
                
    def selectReport(self, item):
        self.treeviewEqs.selection_set(item)
        self.treeview_selection(item=item)

    def openHelp(self):
        arquivo = "FERA.pdf"
        application_path = utilities_general.get_application_path()     
        filepath = os.path.join(application_path,arquivo)
        try:            
            if platform.system() == 'Darwin':       # macOS
                subprocess.call(('open', filepath), shell=True)
            elif platform.system() == 'Windows':    # Windows
                os.startfile(filepath)
            else:           
                openfile = ['xdg-open', filepath]
                subprocess.run(openfile, check=True)
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
            utilities_general.popup_window('O arquivo não possui um \nprograma associado para abertura!', False)

    def menuReports(self, event=None):
        try:
            self.reportbut.menu = tkinter.Menu(self.reportbut, tearoff=0)
            self.reportbut["menu"]= self.reportbut.menu  
            self.menureportsbyeq = {}
            geteqs =  self.treeviewEqs.get_children('')
            for eq in geteqs:
                patheq = self.treeviewEqs.item(eq, 'text')
                self.menureportsbyeq[patheq] = tkinter.Menu(global_settings.root, tearoff=0)
                primeiracamada =  self.treeviewEqs.get_children(eq)
                for reports in primeiracamada:
                    self.menureportsbyeq[patheq].add_command(label=reports, image=global_settings.catimage, compound='left', command=partial(self.selectReport,reports))
                self.reportbut.menu.add_cascade(label=patheq, menu=self.menureportsbyeq[patheq], image=global_settings.imageequip, compound='left')
            self.reportbut.menu.add_separator()
            texto = " FERA - Forensics Evidence Report Analyzer \n"+\
                    "* License: GNU Affero General Public License v3.0\n\n"+\
                    "STATE DEPARTMENT OF PUBLIC SECURITY -- SCIENTIFIC POLICE OF PARANÁ\n\n"+\
                    "  CODED BY by:\nGustavo Borelli Bedendo <gustavo.bedendo@gmail.com>\n\n"+\
                    "  SUPPORTERS :\nAlexandre Vrubel\nRoger Roberto Rocha Duarte\nWellerson Jeremias Colombari\n\n\n\n"+\
                    "  MAIN TESTERS AND USAGE IDEAS:\nConrado Pinto Rebessi\nJacson Gluzezak\nJoel Koster\nLaercio Silva de Campos Junior\nMarcus Fabio Fontenelle do Carmo\nRaphael Zago\n"+\
                    "\n\nJanuary 2024\n\n"+\
                    "It is a work in progress, the code, \ndespite the ugliness and some bugs, is available on:\n"+\
                    "https://github.com/gustavobedendo/FERA"
            
            self.reportbut.menu.add_command(label='Logs', image=global_settings.logimage , compound='left', command= lambda : global_settings.log_window.deiconify())
            self.reportbut.menu.add_command(label='Ajuda', image=global_settings.helpimage , compound='left', command= self.openHelp)  
            

            self.reportbut.menu.add_separator()
            self.reportbut.menu.add_command(label='Sobre', image=global_settings.aboutimage , compound='left', command= lambda: utilities_general.popup_window(texto, False, imagepcp=global_settings.tkphotologo))
            self.reportbut.pack()    
        except Exception as ex:
            utilities_general.printlogexception(ex=ex) 
        finally:
            None 
            

    
    def saveas(self, initialf, asbpathfile):
        path = (asksaveasfilename(initialfile=initialf))
        if(path!=None and path!=''):
            shutil.copy2(asbpathfile, path)
            
    def menuExportInterval(self, event=None):
        self.menuExport = tkinter.Menu(global_settings.root, tearoff=0)
        try:
            self.menuExport.add_command(label="Export arquivos em intervalo", command= lambda : self.exportInterval())
            self.menuExport.tk_popup(event.x_root, event.y_root) 
        except Exception as ex:
            None
        finally:
            self.menuExport.grab_release()
            
    
    def editadd_anottation(self, iiditem):
        observation = self.allobsbyitem[iiditem]
        classes_general.Annotation_Window(observation, self.allobsbyitem[iiditem].annotations, self)

    
    def specialCopy(self):
        pinit = min(global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados)
        linha1 = global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[pinit]['text'][0][1]
        linha2 = global_settings.infoLaudo[global_settings.pathpdfatual].retangulosDesenhados[pinit]['text'][-1][1]
        #print(linha1.x0, linha1.y0, linha2.x1, linha2.y1)
        pagina = global_settings.docatual[pinit]
        tabelas = pagina.find_tables(fitz.Rect(0,0,999,999))
        links = pagina.get_links()
                        
        for tabela in tabelas.tables:
            for linha in tabela.rows:
                if(linha1.y0 > linha.bbox[1] and linha2.y1 < linha.bbox[3]):
                    celulas = filter(lambda x: x is not None, linha.cells)
                    textos = []
                    for cell in celulas:
                        textos.append((pagina.get_textbox(cell).replace("\n", " "), cell))
                    for link in links:
                        r = link['from']
                        if('file' not in link):
                            continue
                        try:
                            arquivo  = link['file']
                            if("#" in arquivo):
                                continue
                            pdfatualnorm = utilities_general.get_normalized_path(global_settings.pathpdfatual)
                            filepath = str(Path(utilities_general.get_normalized_path(os.path.join(Path(utilities_general.get_normalized_path(pdfatualnorm)).parent,arquivo))))
                            #print(self.get_thumbnail(filepath))

                                
                        except:
                            None
                    #print(textos[0], textos[1], textos[2])
                        #print()
                        #print(cell) 

    def menuPopup(self, event):
        if(self.areaselectionActive or self.selectionActive):
            self.menu = tkinter.Menu(global_settings.root, tearoff=0)
            self.menu.add_command(label="Copiar", command=self.copiar)
            self.menucats_copy = tkinter.Menu(self.menu, tearoff=0)
            self.menucats_copy.add_command(label="Linha do Tempo (mobileMerger)", command= lambda : self.specialCopy())
            self.menu.add_cascade(label="Copiar Especial (RTF)", menu=self.menucats_copy)
            self.menu.add_command(label="Export arquivos em intervalo", command= lambda : self.exportInterval())
            self.menu.add_separator()           
            getobscatas =  self.treeviewObs.get_children('')
            self.menucats = tkinter.Menu(self.menu, tearoff=0)
            for obscat in getobscatas:
                cat_text = self.treeviewObs.item(obscat, 'text')
                idobscat = self.treeviewObs.item(obscat, 'values')[1]
                self.menucats.add_command(label=cat_text, image=global_settings.catimage, compound='left', command=partial(self.addmarker,idobscat))
            self.menucats.add_separator()
            self.menucats.add_command(label="Nova categoria", image=global_settings.addcat, compound='left', command=partial(self.addmarker,None))
            self.menu.add_cascade(label='Adicionar marcador', menu=self.menucats, image=global_settings.itemimage, compound='left')
            try:
                if(len(self.docInnerCanvas.find_withtag('quad')) == 0):
                    self.menu.entryconfig(0, state='disabled')
                    #if(not ehLink):
                    self.menu.entryconfig(0, state='disabled')
                else:
                    self.menu.entryconfig(0, state='normal')
                self.menu.tk_popup(event.x_root, event.y_root)         
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)
            finally:
                self.menu.grab_release()
        else:            
            self.rightClick_link_or_obs(event)
                
    def scrollzoom(self, event=None):
        try:
            if (event.delta>0):
                 self.zoomx(tipozoom='plus')
            else:
                 self.zoomx(tipozoom='minus')
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
            
    def focusSimpleSearch(self, event):
        self.simplesearch.focus()
        self.simplesearch.selection_range(0, 'end')
    
    def drawCanvas(self):        
        try:
            self.f12 = False
            self.docFrame = classes_general.CustomFrame(self.docOuterFrame, bg=self.bg, highlightthickness=0)
            self.docFrame.grid(column=0, row=1, sticky='nsew',padx=0, pady=0,)  
            self.docFrame.rowconfigure(0, weight=1)
            self.docFrame.columnconfigure(0, weight=1)            
            self.docInnerCanvas = classes_general.CustomCanvas(self.docFrame, bg='white', highlightthickness=0, relief="raised")    
            self.docInnerCanvas.grid(row=0, column=0, padx=0, pady=0, sticky='ns')
            self.docInnerCanvas.rowconfigure(0, weight=1)
            self.docInnerCanvas.columnconfigure(0, weight=1)    
            self.vscrollbar = tkinter.Scrollbar(self.docFrame, orient='vertical', cursor="left_ptr")
            self.vscrollbar.grid(row=0, column=1, sticky='nse')
            self.vscrollbar.config( command = self.docInnerCanvas.yview )
            self.labeldocframe = tkinter.Frame(self.docFrame)
            self.labeldocframe.grid(row=2, column=0, sticky='ew')
            self.labeldocframe.rowconfigure(0, weight=1)
            self.labeldocframe.columnconfigure((0,1), weight=1)
            global_settings.label_warning_error = tkinter.Label(self.labeldocframe, bg="#%02x%02x%02x" % (145, 145, 145), font=global_settings.Font_tuple_Arial_10, text=" ")
            global_settings.label_warning_error.grid(row=0, column=0, sticky='w')
            self.labeldocname = tkinter.Label(self.labeldocframe, font=global_settings.Font_tuple_Arial_10, text="")
            self.labeldocname.grid(row=0, column=0, sticky='ns')
            self.labelmousepos = tkinter.Label(self.labeldocframe, font=global_settings.Font_tuple_Arial_10, text="")
            self.labelmousepos.grid(row=0, column=1, sticky='e')
            self.docInnerCanvas.bind("<MouseWheel>", self._on_mousewheel)
            self.docInnerCanvas.bind("<Button-4>", self._on_mousewheel)
            self.docInnerCanvas.bind("<Button-5>", self._on_mousewheel)
            
            self.docInnerCanvas.bind("1", lambda e : self.activateDrag(e))
            self.docInnerCanvas.bind("2", lambda e : self.activateSelection(e))
            self.docInnerCanvas.bind("3", lambda e :  self.activateareaSelection(e))
            self.docFrame.bind("<MouseWheel>", self._on_mousewheel)
            self.docFrame.bind("<Button-4>", self._on_mousewheel)
            self.docFrame.bind("<Button-5>", self._on_mousewheel)
            self.docFrame.bind_all("<1>", lambda event: self.clearSelectedTextByCLick("press", event))
            self.docInnerCanvas.bind("<B1-Motion>", self._selectingText)
            self.docFrame.bind_all('<Double-Button-1>', self.doubleClickSelection)            
            global_settings.root.bind('<Control-c>', self.copiar)    
            global_settings.root.bind('<Control-C>', self.copiar) 
            self.docFrame.bind_all("<ButtonRelease-1>", lambda event: self.clearSelectedTextByCLick("release", event))
            self.docInnerCanvas.bind("<Button-3>", self.menuPopup)
            self.hscrollbar = tkinter.Scrollbar(self.docFrame, orient='horizontal', cursor="left_ptr")
            self.hscrollbar.grid(row=1, column=0, sticky='ew')
            self.hscrollbar.config( command = self.docInnerCanvas.xview )            
            self.docInnerCanvas.bind('<Right>', lambda event: self.docInnerCanvas.xview_scroll(1, "units"))
            self.docInnerCanvas.bind('<Left>', lambda event: self.docInnerCanvas.xview_scroll(-1, "units"))
            self.docInnerCanvas.bind('<Up>', lambda event: self.docInnerCanvas.yview_scroll(-1, "units"))
            self.docInnerCanvas.bind('<Down>', lambda event: self.docInnerCanvas.yview_scroll(1, "units"))
            self.docInnerCanvas.bind('<Prior>', lambda event: self.manipulatePagesByClick('prev', event))
            self.docInnerCanvas.bind('<Next>', lambda event: self.manipulatePagesByClick('next', event))
            self.docInnerCanvas.bind('<Home>', lambda event: self.manipulatePagesByClick('first', event))
            self.docInnerCanvas.bind('<End>', lambda event: self.manipulatePagesByClick('last', event))
            #self.docInnerCanvas.bind('<KeyPress-Alt_L>', self.altPressed)
            #self.docInnerCanvas.bind('<KeyRelease-Alt_L>', self.altRelease)
            global_settings.root.bind('<Tab>', self.altRelease2)
            self.docInnerCanvas.bind("<Motion>", self.checkLink)
            self.docInnerCanvas.bind("<Control-MouseWheel>", self.scrollzoom)
            self.docInnerCanvas.bind("<Control-4>", lambda event: self.zoomx(event, tipozoom='plus'))
            self.docInnerCanvas.bind("<Control-5>", lambda event: self.zoomx(event, tipozoom='minus'))
            self.docFrame.bind("<Control-MouseWheel>", self.scrollzoom)
            self.docFrame.bind("<Control-4>", lambda event: self.zoomx(event, tipozoom='plus'))
            self.docFrame.bind("<Control-5>", lambda event: self.zoomx(event, tipozoom='minus'))
            
            self.docFrame.bind_all("<Control-f>", self.focusSimpleSearch)
            self.docFrame.bind_all("<Control-F>", self.focusSimpleSearch)
            self.docFrame.bind_all("<Control-Down>", lambda event: self.dosearchsimple('next'))
            self.docFrame.bind_all("<Control-Up>", lambda event: self.dosearchsimple('prev'))
            global_settings.root.update_idletasks()            
            self.zoom_x =round((self.vscrollbar.winfo_height()-2)/global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh, 16)#  zoom
            self.canvash = self.docInnerCanvas.winfo_height()-self.hscrollbar.winfo_height()-self.labeldocname.winfo_height()
            self.canvasw = self.docFrame.winfo_width()
            margemesq = 0
            margemdir = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw
            self.maiorw = self.canvasw
            if(global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw*self.zoom_x*global_settings.zoom>self.maiorw):
                self.maiorw = global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw*self.zoom_x *global_settings.zoom
    
            self.docInnerCanvas.config(width=global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw*self.zoom_x*global_settings.zoom)    
            sobraEspaco = self.docInnerCanvas.winfo_x()    
            self.mat = fitz.Matrix(self.zoom_x*global_settings.zoom, self.zoom_x*global_settings.zoom)
            self.scrolly = round((global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh*self.zoom_x*global_settings.zoom), 16)*\
                global_settings.infoLaudo[global_settings.pathpdfatual].len  - 35
            self.docInnerCanvas.config(scrollregion=(sobraEspaco, 0, sobraEspaco+ (global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw*\
                                                                                   global_settings.zoom*self.zoom_x), self.scrolly))
            self.docInnerCanvas.configure(xscrollcommand=self.hscrollbar.set)
            self.docInnerCanvas.configure(yscrollcommand=self.vscrollbar.set)
            self.docInnerCanvas.configure(yscrollincrement=str(round((global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh*self.zoom_x)/self.totalMov, 8)))
            self.altpressed=False          
            anc_h = 'nw'
            pos_h =  (self.docFrame.winfo_width() - global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw*self.zoom_x*global_settings.zoom) / 2            
            self.create_fakeimage()
            init = self.docInnerCanvas.winfo_width()/2
            for k in range(global_settings.minMaxLabels):
                for d in range(global_settings.divididoEm):
                    indice = (k*global_settings.divididoEm) + d
                    altura = (k*global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh*self.zoom_x*global_settings.zoom) + \
                        ((d/global_settings.divididoEm)*global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh*self.zoom_x*global_settings.zoom)
                    self.ininCanvasesid[indice] = self.docInnerCanvas.create_image((pos_h,altura), \
                                                                                   anchor=anc_h, tag="canvas")
                self.fakePages[k] = self.docInnerCanvas.create_image((pos_h,(k*global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh*\
                                                                             self.zoom_x*global_settings.zoom)), anchor=anc_h, image=self.fakeImage)
                self.docInnerCanvas.tag_lower(self.fakePages[k])
            self.docInnerCanvas.program = self  
            
            self.topLine = self.docInnerCanvas.create_line(0,0, global_settings.root.winfo_width(), 0, width=10, fill=self.bg)            
            self.fakeLines[0] = self.docInnerCanvas.create_line(0,0, global_settings.root.winfo_width(), 0, width=5, fill=self.bg)
            self.fakeLines[1] = self.docInnerCanvas.create_line(0,global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x * \
                                                                global_settings.zoom, global_settings.root.winfo_width(), \
                                                            global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x * \
                                                                global_settings.zoom, width=5, fill=self.bg)
            self.docInnerCanvas.tag_raise(self.topLine)
            global_settings.root.bind('<Alt-Left>', self.altleft)
            global_settings.root.bind('<Alt-Right>', self.altright)   
            global_settings.root.bind('<Control-g>', lambda e: self.pag.focus_set())
            global_settings.root.bind('<Control-G>', lambda e: self.pag.focus_set())            
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
     
    def altright(self, event):        
        try:
            temp = self.indiceposition+1
            if(self.indiceposition>=9):
                temp = 0
            if(self.positions[temp]!=None):
                newpath = utilities_general.get_normalized_path(self.positions[temp][0])
                novoscroll = self.positions[temp][1]
                self.positions[self.indiceposition] = (global_settings.pathpdfatual, self.vscrollbar.get()[0])
                if(global_settings.pathpdfatual!=newpath):
                    pdfantigo = global_settings.pathpdfatual
                    for i in range(global_settings.minMaxLabels):
                        global_settings.processed_pages[i] = -1
                    sobraEspaco = 0
                    if(self.docFrame.winfo_width() > global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw * self.zoom_x * global_settings.zoom):
                        sobraEspaco = self.docInnerCanvas.winfo_x()
                    self.maiorw = self.docFrame.winfo_width()
                    if(global_settings.infoLaudo[newpath].pixorgw*self.zoom_x*global_settings.zoom>self.maiorw):
                        self.maiorw = global_settings.infoLaudo[newpath].pixorgw*self.zoom_x *global_settings.zoom           
                    self.scrolly = global_settings.infoLaudo[newpath].pixorgh*self.zoom_x*global_settings.zoom*global_settings.infoLaudo[newpath].len  - 35
                    self.docInnerCanvas.config(scrollregion=(sobraEspaco, 0, sobraEspaco+ (global_settings.infoLaudo[newpath].pixorgw*global_settings.zoom*self.zoom_x), self.scrolly))
                    pagina = round(global_settings.infoLaudo[newpath].ultimaPosicao*global_settings.infoLaudo[newpath].len)   
                    self.docInnerCanvas.yview_moveto(global_settings.infoLaudo[newpath].ultimaPosicao)
                    if(str(pagina+1)!=self.pagVar.get()):
                        self.pagVar.set(str(pagina+1))
                    global_settings.pathpdfatual = utilities_general.get_normalized_path(newpath)
                    try:
                        global_settings.docatual.close()
                    except Exception as ex:
                        None
                    global_settings.docatual = fitz.open(global_settings.pathpdfatual)
                    self.labeldocname.config(font=global_settings.Font_tuple_Arial_10, text=os.path.basename(global_settings.pathpdfatual)) 
                    self.clearAllImages()
                    self.totalPgg.config(font=global_settings.Font_tuple_Arial_10, text="/ "+str(global_settings.infoLaudo[global_settings.pathpdfatual].len))                    
                    for pdf in global_settings.infoLaudo:
                        global_settings.infoLaudo[pdf].retangulosDesenhados = {}
                    if(global_settings.infoLaudo[pdfantigo].zoom_pos!=global_settings.infoLaudo[global_settings.pathpdfatual].zoom_pos):
                        #global_settings.root.after(1, lambda: self.zoomx(None, None, pdfantigo))
                        self.zoomx(None, None, pdfantigo)
                #novoscroll = self.positions[temp][1]
                #global_settings.root.after(1, self.docInnerCanvas.yview_moveto(novoscroll))
                self.docInnerCanvas.yview_moveto(novoscroll)
                
                pagina = round(novoscroll*global_settings.infoLaudo[newpath].len)
                if(str(pagina+1)!=self.pagVar.get()):
                    self.pagVar.set(str(pagina+1))
                self.totalPgg.config(font=global_settings.Font_tuple_Arial_10, text="/ "+str(global_settings.infoLaudo[global_settings.pathpdfatual].len))
                self.indiceposition += 1
                if(self.indiceposition>9):
                    self.indiceposition = 0
                global_settings.root.update_idletasks()
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        
    def altleft(self, event):        
        try:
            temp = self.indiceposition-1
            if(self.indiceposition<=0):
                temp = 9
            if(self.positions[temp]!=None):
                newpath = utilities_general.get_normalized_path(self.positions[temp][0])
                novoscroll = self.positions[temp][1]
                self.positions[self.indiceposition] = (global_settings.pathpdfatual, self.vscrollbar.get()[0])
                #print(global_settings.pathpdfatual, self.positions[temp])
                if(global_settings.pathpdfatual!=newpath):
                    pdfantigo = global_settings.pathpdfatual
                    for i in range(global_settings.minMaxLabels):
                        global_settings.processed_pages[i] = -1
                    sobraEspaco = 0
                    if(self.docFrame.winfo_width() > global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw * self.zoom_x * global_settings.zoom):
                        sobraEspaco = self.docInnerCanvas.winfo_x()
                    self.maiorw = self.docFrame.winfo_width()
                    if(global_settings.infoLaudo[newpath].pixorgw*self.zoom_x*global_settings.zoom>self.maiorw):
                        self.maiorw = global_settings.infoLaudo[newpath].pixorgw*self.zoom_x *global_settings.zoom           
                    self.scrolly = global_settings.infoLaudo[newpath].pixorgh*self.zoom_x*global_settings.zoom*global_settings.infoLaudo[newpath].len  - 35
                    self.docInnerCanvas.config(scrollregion=(sobraEspaco, 0, sobraEspaco+ (global_settings.infoLaudo[newpath].pixorgw*global_settings.zoom*self.zoom_x), self.scrolly))
                    pagina = round(global_settings.infoLaudo[newpath].ultimaPosicao*global_settings.infoLaudo[newpath].len)   
                    global_settings.pathpdfatual = utilities_general.get_normalized_path(newpath)
                    try:
                        global_settings.docatual.close()
                    except Exception as ex:
                        None
                    self.docInnerCanvas.yview_moveto(global_settings.infoLaudo[global_settings.pathpdfatual].ultimaPosicao)
                    if(str(pagina+1)!=self.pagVar.get()):
                        self.pagVar.set(str(pagina+1))
                    
                    global_settings.docatual = fitz.open(global_settings.pathpdfatual)
                   
                    self.labeldocname.config(font=global_settings.Font_tuple_Arial_10, text=os.path.basename(global_settings.pathpdfatual))
                    self.clearAllImages()
                    self.totalPgg.config(font=global_settings.Font_tuple_Arial_10, text="/ "+str(global_settings.infoLaudo[global_settings.pathpdfatual].len))                    
                    for pdf in global_settings.infoLaudo:
                        global_settings.infoLaudo[pdf].retangulosDesenhados = {} 
                    if(global_settings.infoLaudo[pdfantigo].zoom_pos!=global_settings.infoLaudo[global_settings.pathpdfatual].zoom_pos):
                        #global_settings.root.after(1, lambda: self.zoomx(None, None, pdfantigo))
                        novoscroll = self.positions[temp][1]
                        self.zoomx(None, None, pdfantigo)
                #
                pagina = round(novoscroll*global_settings.infoLaudo[global_settings.pathpdfatual].len)
                
                global_settings.root.after(10, lambda: self.docInnerCanvas.yview_moveto(novoscroll))
                self.docInnerCanvas.yview_moveto(novoscroll)
                if(str(pagina+1)!=self.pagVar.get()):
                    self.pagVar.set(str(pagina+1))
                    
                self.totalPgg.config(font=global_settings.Font_tuple_Arial_10, text="/ "+str(global_settings.infoLaudo[global_settings.pathpdfatual].len))
                self.indiceposition -= 1 
                if(self.indiceposition<0):
                    self.indiceposition = 9
                global_settings.root.update_idletasks()
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
    
    def altRelease2(self, event):
        #print("RELEASE")
        if(self.altpressed):
            self.altpressed=False
    
    def altRelease(self, event):
        #print("RELEASE")
        self.altpressed=False
    
    def altPressed(self, event):
        self.altpressed=True
                   
    def create_fakeimage(self):
        altura = math.ceil(global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh * self.zoom_x*global_settings.zoom)
        largura = math.ceil(global_settings.infoLaudo[global_settings.pathpdfatual].pixorgw * self.zoom_x*global_settings.zoom)
        image = Image.new('RGBA', (largura, altura), (255, 255, 255, 255))       
        self.fakeImage = ImageTk.PhotoImage(image)

 




def start_fera_app():
    start_time = time.time()
    try:
        
        global_settings.splash_window.window.deiconify()
        global_settings.splash_window.label['text'] = "Carregando ferramenta..."
        global_settings.splash_window.label.update()
        try:
            sqliteconn = utilities_general.connectDB(str(global_settings.pathdb), 5)
            pos = 0
            try:
                cursor = sqliteconn.cursor()
                teste = 'SELECT id_conf, config, param FROM FERA_CONFIG'
                cursor.custom_execute(teste, None, False, False)
                configs = cursor.fetchall()
                for config in configs:
                    if(config[1]=='zoom'):
                        pos = int(config[2])
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)                   
            finally:
                if(sqliteconn):
                    sqliteconn.close()            
            #global_settings.posicaoZoom = pos
            #global_settings.zoom = global_settings.listaZooms[global_settings.posicaoZoom]
            global_settings.indexingwindow = classes_general.Indexing_window(global_settings.root)
            if(global_settings.indexing_thread!=None and global_settings.indexing_thread.is_alive()):
                totlendoc = 0
                already_indexed = 0
                for relatorio in global_settings.infoLaudo:
                    if(global_settings.infoLaudo[relatorio].status=='naoindexado'):
                        totlendoc += global_settings.infoLaudo[relatorio].len
                        already_indexed += global_settings.infoLaudo[relatorio].paginasindexadas
                global_settings.indexingwindow.window.rowconfigure((0,1), weight=1)
                global_settings.indexingwindow.window.columnconfigure(0, weight=1)
                label = tkinter.Label(global_settings.indexingwindow.window, font=global_settings.Font_tuple_Arial_10, text="Indexando documento(s)", image=global_settings.processing, compound='left')
                label.image = global_settings.processing
                label.grid(row=0, column=0, sticky='ew', pady=5, padx=5)
                global_settings.indexingwindow.progressbar = ttk.Progressbar(global_settings.indexingwindow.window, mode='determinate')
                global_settings.indexingwindow.progressbar.grid(row=1, column=0, sticky='ew', pady=5)
                global_settings.indexingwindow.progressbar['value'] = already_indexed
                global_settings.indexingwindow.progressbar['maximum'] = totlendoc
                global_settings.indexingwindow.window.protocol("WM_DELETE_WINDOW", lambda: None)
                global_settings.indexingwindow.window.resizable(False, False)                
                
            else:
                None
                #global_settings.indexingwindow.window.destroy()
            global_settings.splash_window.label['text'] = "Carregando interface..."
            global_settings.splash_window.label.update()
            global_settings.fera_main_window = MainWindow() 
            global_settings.root.mainloop()
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        try:
            global_settings.splash_window.window.destroy()
        except:
            None
        
            
    except Exception as ex:
        utilities_general.printlogexception(ex=ex)
        global_settings.on_quit()
    finally:
        try:
            global_settings.splash_window.window.destroy()
        except:
            None
            

    
        



