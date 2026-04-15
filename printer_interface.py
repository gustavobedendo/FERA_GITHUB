# -*- coding: utf-8 -*-
"""
Created on Wed Mar 31 10:17:25 2021

@author: gustavo.bedendo
"""

import traceback
import subprocess
import PIL.ImageWin
import PIL.Image
import io
from tkinter import *
import tkinter
import time
import sys
from tkinter import Frame
from tkinter import font # * doesn't import font or messagebox
from tkinter import messagebox
from tkinter import ttk
from tkinter import PhotoImage
import multiprocessing as mp
import os
import fitz
import platform
plt = platform.system()
if plt=="Windows":
    import win32print
    import win32ui
    import win32gui


#
# Constants for GetDeviceCaps
#
#
# HORZRES / VERTRES = printable area
#
HORZRES = 8
VERTRES = 10
#
# LOGPIXELS = dots per inch
#
LOGPIXELSX = 88
LOGPIXELSY = 90
#
# PHYSICALWIDTH/HEIGHT = total area
#
PHYSICALWIDTH = 110
PHYSICALHEIGHT = 111
#
# PHYSICALOFFSETX/Y = left / top margin
#
PHYSICALOFFSETX = 112
PHYSICALOFFSETY = 113




class PlaceholderEntry(ttk.Entry):
    def __init__(self, container, placeholder, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        self.placeholder = placeholder
        self.insert("0", self.placeholder)
        self.bind("<FocusIn>", self._clear_placeholder)
        self.bind("<FocusOut>", self._add_placeholder)

    def _clear_placeholder(self, e):
        if(super().get()=='Buscar...'):
            self.delete("0", "end")

    def _add_placeholder(self, e):
        if not self.get():
            self.insert("0", self.placeholder)
            
def font_size(fs):
    return font.Font(family='Helvetica', size=fs, weight='bold')

class Printer():
    
    def __init__(self):
        None
        self.printer_process = None

    def validatePages(self, name=None, index=None, mode=None, sv=None):
        #print('teste')
        try:
            regex = "([0-9]+)((-[0-9]+)|,[0-9]+)*\Z"
            matched = re.match(regex, self.pagvar.get())
        
            is_match = bool(matched)
            #print(is_match)
            if(is_match):
                self.p_button['relief'] = 'raised'
                self.p_button['state'] = 'normal'
                pageswithcomma = self.pagvar.get().split(',')
                self.listadepaginas = []
                for pages in pageswithcomma:
                    if('-' in pages):
                        pagessplit = pages.split('-')
                        init = min(int(pagessplit[0]),int(pagessplit[1]))
                        end = max(int(pagessplit[0]),int(pagessplit[1]))
                        for k in range(init, end+1):
                            if(k > len(self.docprint) or k <= 0):
                                self.avisolabel['text'] = 'Intervalo de páginas não está no documento - Mínimo: {} -- Máximo: {} '.format(1, len(self.docprint))
                                self.p_button['relief'] = 'sunken'
                                self.p_button['state'] = 'disabled'
                                return
                            self.listadepaginas.append(k)
                    else:
                        if(int(pages) > len(self.docprint) or int(pages) <= 0):
                            self.avisolabel['text'] = 'Intervalo de páginas não está no documento - Mínimo: {} -- Máximo: {} '.format(1, len(self.docprint))
                            self.p_button['relief'] = 'sunken'
                            self.p_button['state'] = 'disabled'
                            return
                        self.listadepaginas.append(int(pages))
                if(self.prop==None and self.canvaspageframe.winfo_height() != 1):
                    self.prop =  self.canvaspageframe.winfo_height() / self.heightdoc
                    #print(self.canvaspageframe.winfo_height(), self.heightdoc, self.prop)
                    self.matprop = fitz.Matrix(self.prop*0.9, self.prop*0.9)
                elif(self.prop!=None):
                    None
                else:
                    return
                self.avisolabel['text'] = ""
                self.indice = 0
                paginadoc = self.listadepaginas[self.indice]
                pix = self.docprint[paginadoc-1].getPixmap(alpha=False, matrix=self.matprop)  
                #print(primeirapagina, self.matprop, self.prop)
                imgdata = pix.getImageData("ppm")
                self.imagemavista = PhotoImage(data=imgdata)
                desloc = self.canvaspageframe.winfo_width() / 2
                self.paginaavista =  self.canvasfirstpage.create_image((desloc,0), image=self.imagemavista, anchor='n')
                self.varpageview.set("{} / {} - Pagina: {}".format(str(1), len(self.listadepaginas), paginadoc))
            else:
                self.avisolabel['text'] = "Páginas inválidas\n Separar páginas por ';' - Exemplo (Páginas 1 a 10 e 12): '1-10;12' "
                self.p_button['relief'] = 'sunken'
                self.p_button['state'] = 'disabled'
        except:
            traceback.print_exc()
    def manipulatePagesByClick(self, evento=None):
        try:
            if(evento=='next'):
                self.indice += 1
            elif(evento=='prev'):
                self.indice -= 1
            elif(evento=='first'):
                self.indice = 0
            elif(evento=='last'):
                self.indice = len(self.listadepaginas)-1
            self.indice = max(self.indice, 0)
            self.indice = min(self.indice, len(self.listadepaginas)-1)
            paginadoc = self.listadepaginas[self.indice]
            pix = self.docprint[paginadoc-1].getPixmap(alpha=False, matrix=self.matprop)  
            #print(primeirapagina, self.matprop, self.prop)
            imgdata = pix.getImageData("ppm")
            self.imagemavista = PhotoImage(data=imgdata)
            desloc = self.canvaspageframe.winfo_width() / 2
            self.paginaavista =  self.canvasfirstpage.create_image((desloc,0), image=self.imagemavista, anchor='n')
            self.varpageview.set("{} / {} - Pagina: {}".format(str(1), len(self.listadepaginas), paginadoc))
        except:
            None
    def sel_printer(self, name=None, index=None, mode=None, sv=None):
        if(self.popupMenu2!=None):
            
            if plt == "Linux":
                print(name)
                comando = "echo kali | lpoptions -d "+self._printer.get() +" -l"
                #comando = ["lpoptions", "-d", name, "-l"]
                process = subprocess.Popen(comando, shell=True, stdout=subprocess.PIPE, universal_newlines=True)
                #outputOptions = process.stdout
                outputOptions = process.communicate()[0].strip()
                choicesinteiras = outputOptions.splitlines()
                temcor = False
                choices = []
                for choice in choicesinteiras:
                    if('COLOR' in choice.upper()):
                        temcor = True
                        colors = choice.split(": ")[1]
                        choices=colors.split(" ")
                        break
                if(len(choices)>0):
                    self._color.set(choices[0])
                estado = 'normal'
                if not temcor:
                    estado = 'disabled'    
                #self.popupMenu2['choices']=choices
                
                self.popupMenu2.configure(state=estado)    
                
                self.popupMenu2['menu'].delete(0, 'end')
                for choicex in choices:
                    self.popupMenu2['menu'].add_command(label=choicex, command= lambda x=choicex: self._color.set(x))
            
        '''
        printer = self._printer.get()
        PRINTER_DEFAULTS = {"DesiredAccess":win32print.PRINTER_ALL_ACCESS}
        handle = win32print.OpenPrinter(printer, PRINTER_DEFAULTS)
        try:
            level = 2
            attributes = win32print.GetPrinter(handle, level)
            #devmode=pDevModeObj  
            print(dir(attributes['pDevMode']))  
            print(attributes['pDevMode'].PrintQuality)
            print(attributes['pDevMode'].PaperSize)
            print(attributes['pDevMode'].PaperWidth)
            print(attributes['pDevMode'].Scale)
            print(attributes['pDevMode'].Size)
            print(attributes['pDevMode'].Orientation)
            print(attributes['pDevMode'].PaperLength)
            
                
        except:
            None
        finally:
            win32print.ClosePrinter(handle)
        '''
    
    def printFunction(self, _filename, widthdoc, heightdoc, atual, root):
        self.atual = atual
        self.window = root
        self.window.title("Python Printer")
        self.window.geometry("480x720")
        self.window.resizable(False, False)
        #root.tk.call('encoding', 'system', 'utf-8')
        self.window.protocol("WM_DELETE_WINDOW", self.on_quit)
        self.root = root
        self.heightdoc = heightdoc
        self.widthdoc = heightdoc
        # Add a grid
        self.window.columnconfigure(0, weight = 1)
        self.window.rowconfigure(0, weight = 1)
        mainframe = Frame(self.window)
        self.popupMenu2 = None
        #mainframe.grid(column=0,row=0, sticky=(N,W,E,S) )
        
        mainframe.grid(column=0,row=0, sticky='nsew')
        mainframe.columnconfigure(0, weight = 1)
        mainframe.rowconfigure(1, weight = 1)
        
        configframe = Frame(mainframe)
        configframe.grid(row=0, column=0, sticky='nsew', padx = 10, pady=10)
        configframe.columnconfigure(0, weight = 1)
        configframe.rowconfigure((0,1,2,3,4,5,6,7,8), weight = 1)
        
        # Create a _printer variable
        self._printer = StringVar(self.window)
        self._printer.trace('w', self.sel_printer)
        # Create a _color variable
        self._color = StringVar(self.window)
        self._color.trace('w', self.sel_color)
        self.filename = _filename
        self.docprint = fitz.open(_filename)
        self.paginaavista = None
        self.prop = None
        try:    
            if plt == "Linux":
                comando = "echo kali | lpstat -a"
                process = subprocess.Popen(comando, shell=True, stdout=subprocess.PIPE, universal_newlines=True)
                outputPrinters = process.communicate()[0].strip()
                print("Impressoras = ", outputPrinters)
                choicesinteiras = outputPrinters.splitlines()
                choices = []
                for linhainteira in choicesinteiras:
                    choices.append(linhainteira.split(" ")[0])
                if(len(choices) > 0):
                    self._printer.set(choices[0])
            elif plt=="Windows":
                choices = [printer[2] for printer in win32print.EnumPrinters(2)]
                self._printer.set(win32print.GetDefaultPrinter()) # set the default option
            
            Label(configframe, text="DOCUMENTO:", font=font_size(12)).grid(row = 0, column = 0, sticky='n', pady=5)
            docvar = StringVar(configframe)
            docvar.set(os.path.basename(_filename))
            #print(_filename)
            docentry = Entry(configframe, justify='center', text=os.path.basename(_filename), state='disabled',  textvariable=docvar, exportselection=False)
            docentry.grid(row = 1, column =0, pady=(0, 10,), sticky='ew', padx=30)
            
            Label(configframe, text="SELECIONE A IMPRESSORA", font=font_size(10)).grid(row = 2, column = 0, sticky='n', pady=5)
            self.popupMenu = OptionMenu(configframe, self._printer, *choices)
            self.popupMenu['font'] = font_size(12)
            self.popupMenu.grid(row =3, column =0, sticky='n', pady=5)
            
            # Dictionary with options
            choices = []
            temcor = False
            if plt == "Linux":
                comando = "echo kali | lpoptions -d "+self._printer.get()+ " -l"
                process = subprocess.Popen(comando, shell=True, stdout=subprocess.PIPE, universal_newlines=True)
                outputOptions = process.communicate()[0].strip()
                choicesinteiras = outputOptions.splitlines()
                
                for choice in choicesinteiras:
                    if('COLOR' in choice.upper()):
                        temcor = True
                        colors = choice.split(": ")[1]
                        choices=colors.split(" ")
                        break
                if(len(choices)>0):
                    self._color.set(choices[0])
            elif plt=="Windows":
                temcor = True
                choices = ["COLORIDO", "PRETO/BRANCO"]
                self._color.set("PRETO/BRANCO") # set the default option
            
            estado = 'normal'
            if not temcor:
                estado = 'disabled'
            self.popupMenu2 = OptionMenu(configframe, self._color, *choices)
            self.popupMenu2.configure(state=estado)
            self.popupMenu2['font'] = font_size(12)
            Label(configframe, text="COR", font=font_size(10)).grid(row = 4, column = 0)
            self.popupMenu2.grid(row = 5, column =0)
            
            paglabel = tkinter.Label(configframe, text="Páginas de impressão - Total {}):".format(len(self.docprint)))
            paglabel['font'] = font_size(10)
            paglabel.grid(row = 6, column = 0, pady=(10, 5))
           # 
           
            self.avisolabel = tkinter.Label(configframe, text="", fg='red')
            self.avisolabel.grid(row = 8, column =0, pady=(0, 10,), sticky='ew')
            self.pagvar = tkinter.StringVar(configframe)
            
            
            
            
            self.pagentry = PlaceholderEntry(configframe, justify='center', placeholder="",  textvariable=self.pagvar,\
                                        exportselection=False)
            self.pagentry.grid(row = 7, column =0, pady=(0, 10,), sticky='ew', padx=30)
            
            
            #------------------------------------------------------------
            
            self.pageframe = tkinter.Frame(mainframe)
            self.pageframe.grid(row=1, column=0, sticky='nsew', padx = 10, pady=10)
            self.pageframe.columnconfigure(0, weight = 1)
            self.pageframe.rowconfigure(2, weight = 1)
            
            self.progressprint = ttk.Progressbar(self.pageframe, mode='indeterminate')                
            self.progressprint.grid(row=0, column=0, sticky='nsew', pady=5)
            self.progressprint.grid_remove()
            
            controlpageframe = tkinter.Frame(self.pageframe)
            controlpageframe.grid(row=1, column=0, sticky='nsew')
            controlpageframe.columnconfigure((0,1,3,4), weight = 1)
            controlpageframe.columnconfigure(2, weight = 2)
            controlpageframe.rowconfigure(1, weight = 1)
            
            self.canvaspageframe = tkinter.Frame(self.pageframe)
            self.canvaspageframe.grid(row=2, column=0, sticky='nsew')
            self.canvaspageframe.columnconfigure(0, weight = 1)
            self.canvaspageframe.rowconfigure(0, weight = 1)
            
            firstPage = b'iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAABgklEQVRIS92VS07DMBCGZ5J0X8dOJbgAqni3Yk0OwpZz8LgLG64AbZHYwwGgFNi1aeME8VRjD4rUIoTSxBHtpt56PJ//8fxjhAUvXHB+WGKAK8ShjdgJguA+r4xCiIYCqMvh8CwrbmaJXM4JiI7CMDydBeCc7xHABWndlVI2ywMQT8Lh8DjroBCiqYkuAVGSUr6U8nlugKrn7VpatwDgRSu1H0XR0yyV+SXKUFD1vJ1J8tdJ8se8NyoFYIxtoWW1AeBdOY4f9/sPRT4yBjDGNtC2O5rok2zbjweDblHydN8IIIRY1VrfaqKk4jj7Ra37G2wKWBkrdYMAuuI4fhAEdya3N1aQBnLO60R0pRHHFctKIbkGnF7ASME02HXddUTsKIAv03coBUhBjLFNtO02EH2YdFJpQAqp1mrbllKp0d5UkvhxHPfmZrRpoj9u9qMoyjTcv4ZdOkk1UQsBeqPRqFFqFjHOD0ip67w581OuJFkLw/C8FMC0z4vilvjLLJJuuv8NJqHsGRXGYe0AAAAASUVORK5CYII='
            previousPage = b'iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAA30lEQVRIS+XUvQrCMBAH8Dv6BOb6fuLQQUdt8RUEvyYHN/EpxHcRnEsS4l6MtGihoCYXmkWzBv4/7oNDiPwwcj78N4BENKmS5HQry8unVoe2CIUQe0DMEGColDr2CbThYO1Ka118WxRuBazwGuYA7HAOEBTuCwSHewFCiCUg5j4DfTds5wyIaGEB5tGAehGIaGcBxmDtWmudc+6Xs4JnWDDiCzTzaitB3GgpZz6VcIAOckfcGimnLoQLsJEQoIMgwEgpdejz2L2ycJCmma2qszHmGgNwtb/5D22RV/hvAA91028ZZhfsxAAAAABJRU5ErkJggg=='   
            nextPage = b'iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAA3ElEQVRIS+3UUQqCQBAG4H/GA7QqnaOiO1RE0GN0x6jnlOpCzp5A2TAQIqidXfOpfPf/5p9dJQz80MD5+AEgHY/X1DS1tbaIWad3RSbPDwxs4dzeWnsIRbxAmqYjMJ8JmMcgXqCduA+iAvogaiAWCQJikGDgFSFgJSKXd7crGnDMBQMzJlpWVXX7GtDeqC6cgJ2InD59G0ENHuHOlZwkU014C6sBY4wBUISEq4E2nJlLEE20k3dr8zboE65qkOX5EcAmdHJ1gyzLFkRUi8g19E+qahAT+vyO9wz+wB0mD2MZRUQZ9AAAAABJRU5ErkJggg=='
            lastPage = b'iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAABbUlEQVRIS93VzU7CQBAH8JlZkPiRlC5LOZgQlahXzh5MQB/Bg08jvIlXH6O8gOInvoGxbFvudLtmjTWElCaky4Uem2Z+zX92ZhE2/OCG68MWA64Qt5poMguC56IYHc/rsCS5jKLoPu+7lRFxzseA2CbEaynl4yqEcz4AxEEUhrQW4LpuG4lGAOCkRFez6fQprwAXYghaGyD3ZwubXK/Xj4gxgxz8IeNlpBRgijmt1glLEh8A9nSa9uM4fllESgO/iOd1UKkRIda0UgZ5zRArgCnWbDZP50niM6Kq1roXRdG7eW8NMMWEEOdzpXyGSETUlVJ+bQQwWK1a7QZB8G0NEEKcJUr5oDUxxnphGE6sRZTlT4gVADD5f1hrcnaCGMDOYnOtAP8zgLi7fDxLA47jHLNKxQzYft6AlQZ4o2Em9jBlrF+0UTnnd4A4XHvZcc5vtNafcRy/Fa1rsxSBsYtYyoe1tqmtq3SLr0xbEf0AkgncGV5n0L4AAAAASUVORK5CYII='
            imFP = PhotoImage(data=firstPage)
            imPP = PhotoImage(data=previousPage)
            imNP = PhotoImage(data=nextPage)
            imLP = PhotoImage(data=lastPage)
            
            self.fp = tkinter.Button(controlpageframe, image=imFP)
            self.fp.image = imFP
            self.fp.grid(column=0, row=0, sticky='e', padx = 5)  
            self.fp.config(command=lambda: self.manipulatePagesByClick('first'))
            self.pp = Button(controlpageframe, image=imPP)
            self.pp.image = imPP
            self.pp.grid(column=1, row=0, sticky='w', padx = 5) 
            self.pp.config(command=lambda: self.manipulatePagesByClick('prev')) 
            
            self.varpageview = StringVar()
            self.varpageview.set("0 / 0")
            self.pageviewentry = Entry(controlpageframe, justify='center', text="0 / 0", state='disabled',  textvariable=self.varpageview, exportselection=False)
            self.pageviewentry.grid(column=2, row=0, sticky='nsew', padx = 5) 
            
            self.np = Button(controlpageframe, image=imNP)
            self.np.image = imNP
            self.np.grid(column=3, row=0, sticky='e', padx = 5) 
            self.np.config(command=lambda: self.manipulatePagesByClick('next'))
            self.lp = Button(controlpageframe, image=imLP)
            self.lp.image = imLP
            self.lp.grid(column=4, row=0, sticky='w', padx = 5) 
            self.lp.config(command=lambda: self.manipulatePagesByClick('last'))
            
            self.canvasfirstpage = Canvas(self.canvaspageframe, background='lightgrey')
            self.canvasfirstpage.grid(row=0, column=0, sticky='nsew', padx = 10, pady=10)
            
            
            
            #------------------------------------------------------------
            
            #lastpageframe = Frame(mainframe)
            #lastpageframe.grid(row=1, column=1, s\ticky='nsew', padx = 10, pady=10)
            #lastpageframe.columnconfigure(0, weight = 1)
            #lastpageframe.rowconfigure(0, weight = 1)
            #canvassecondpage = Canvas(lastpageframe, background='yellow')
            #canvassecondpage.grid(row=0, column=0, sticky='nsew', padx = 10, pady=10)
            
            #------------------------------------------------------------
            
            #Label(mainframe).grid(row = 2, column = 0, columnspan=2)
            #p_button = Button(mainframe, text=u'\uD83D\uDDB6' + " IMPRIMIR", command=PrintAction, fg="dark green", relief='sunken', bg = "white")
            
            printerimageb = b'iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAABMUlEQVRIicXVvUrDUBiH8R9d3Do5FS+joKvg4OSgbl1qEdShziJegBfQwUHB3Y/JwcHNScQPsDdQoQUnHUWFOuQUQtu0aUzxDw/hkPO+T8jhTRjMMrop2BxSmypV/KCcwEYQfGM9q6A54v4CPrGDL6xNSwDbWSSTCGArSFbzEsyLzuA4Ru/gF/MQzOAQR328o56HICnNUDuQS9ziJvCEdmydlnao7a3v8ED07s6xlzNXobeuaIDyyD5Ow7U8DUEdjXCdiiCe/xMUMJeRQhrBrnSf62GcJAkaoiGp4gLPmJ2QR5zF+jR6ghe0RJPYxBvuw1OUYgVJlMLea7zG+rRC74EchM2wEitIYiUmOBjWcJRgkmQSVPAxhspfBEUsjaGYRdCR/NNPopNWUJN9Dmr9zX4BoCKvbw1oarEAAAAASUVORK5CYII='
            printerimage = PhotoImage(data=printerimageb)
            self.p_button = Button(mainframe, text="IMPRIMIR", image=printerimage, compound="left", command= lambda : self.PrintAction('print'), fg="dark green", state='disabled', relief='sunken', bg = "white")
            self.p_button.image = printerimage
            #self.p_button = Button(mainframe, text="IMPRIMIR", command= lambda : self.PrintAction('print'), fg="dark green", state='disabled', relief='sunken', bg = "white")

            self.p_button['font'] = font_size(18)
            self.p_button.grid(row = 3, column =0, pady=10)
            self.pagvar.trace_add('write', self.validatePages)
            #self.root.update_idletasks()
            #self.pagvar.set(str(self.atual))
            
            #
            self.root.after(100, lambda : self.setpagvar(str(self.atual)))
           
            
            #pagvar.trace_add("write", lambda pagvar, p_button: )
        except:
            traceback.print_exc()
            
    def setpagvar(self, pagina):
        try:
            self.pagvar.set(pagina)
        except:
            traceback.print_exc()
    # on change dropdown value
   
    # link function to change dropdown
    
    
    def sel_color(self, *args):
       
        None
        
    
    
    
    def getPrinterOutput(self):
        error_state = self.error_state
        impressos = self.paginasImpressas
        try:
            self.progressprint['value'] = impressos
            #print("{} / {}".format(impressos, self.progressprint['maximum']))
            if(self.progressprint['value']==self.progressprint['maximum'] and not error_state):
                self.on_quit()
            elif(not error_state):
                self.root.after(100, self.getPrinterOutput)
            else:            
                self.progressprint.grid_remove()
                self.progressprint = Label(self.pageframe, text='ERRO DE IMPRESSÃO', fg='red')                
                self.progressprint.grid(row=0, column=0, sticky='nsew', pady=5)
                self.p_button['text']='FECHAR'
        except:
            traceback.print_exc()    
    
    def PrintAction(self, action=None):   
        if(action=='print'):
            if not self.filename:
                messagebox.showerror("Error", "No File Selected")
                return
            elif not self._printer.get():
                messagebox.showerror("Error", "No Printer Selected")
                return
            else:
                self.popupMenu['state'] = 'disabled' 
                self.popupMenu2['state'] = 'disabled' 
                self.pagentry['state'] = 'disabled' 
                impressora = self._printer.get()
                cor = str(self._color.get())
                filename = self.filename   
                self.p_button['fg']="dark red"
                self.p_button['text']='CANCELAR'
                self.p_button['command']=lambda : self.PrintAction('cancelar')
                self.progressprint.grid()
                #time.sleep(3)
                #self.on_quit()
                #return
                self.progressprint['maximum'] = len(self.listadepaginas)
                self.progressprint['mode'] = 'determinate'
                #self.retorno = mp.Queue()
                #self.input = mp.Queue()
                self.printerProcess(impressora, cor, filename, self.listadepaginas)    
                #self.printer_process.start()
                #printerProcess(impressora, cor, filename, self.listadepaginas, self.retorno, self.input)
                self.root.after(100, self.getPrinterOutput)
                
        elif(action=='cancelar'):
            #self.input.put(('cancelar'))
            self.on_quit()
    def on_quit(self):
        global root
        try:
            self.docprint.close()
        except:
            None
        if plt=="Windows":
            try:
                print('Encerrando doc')
                self.hDC.EndDoc()
                
            except:
                traceback.print_exc()
            try:
                print('Encerrando hdc1')
                #self.hDC1.DeleteDC()
                #self.hDC1.DeleteDC()
            except:
                traceback.print_exc()
            try:
                print('Deletando hdc')
                self.hDC.DeleteDC()
                #self.hDC.DeleteDC()
            except:
                traceback.print_exc()
            try:
                None
                win32print.ClosePrinter(self.hPrinter)
                print('Fechando impressora')
            except:
                traceback.print_exc()
        try:
            None
            root.destroy()
        except:
            traceback.print_exc()
    
    def imprimirPaginas(self, indice):
        page = self.listadepaginas[indice]
        pagina = page-1
        printable_area = self.hDC.GetDeviceCaps (HORZRES), self.hDC.GetDeviceCaps (VERTRES)
        printer_size = self.hDC.GetDeviceCaps (PHYSICALWIDTH), self.hDC.GetDeviceCaps (PHYSICALHEIGHT)
        printer_margins = self.hDC.GetDeviceCaps (PHYSICALOFFSETX), self.hDC.GetDeviceCaps (PHYSICALOFFSETY)
        try:
            #if(not inputcommand.empty()):
            #    break
            
            
            pix = self.docprint[pagina].getPixmap(alpha=False, matrix=fitz.Matrix(3,3))  
            pix.setResolution(600, 600)                    
            imgdata = pix.pillowData(format="PNG", optimize=True)
            image = PIL.Image.open(io.BytesIO(imgdata))
            if image.size[0] > image.size[1]:
                image = image.rotate (90)     
            ratios = [1.0 * printable_area[0] / image.size[0], 1.0 * printable_area[1] / image.size[1]]
            scale = min (ratios)    
            dib = PIL.ImageWin.Dib(image)
            scaled_width, scaled_height = [int (scale * i) for i in image.size]
            x1 = int ((printer_size[0] - scaled_width) / 2)
            y1 = int ((printer_size[1] - scaled_height) / 2)
            x2 = x1 + scaled_width
            y2 = y1 + scaled_height
            tentativas = 0
            max_tentativas = 3  
            while(tentativas <= max_tentativas):
                try:
                    self.hDC.StartPage()                        
                    dib.draw(self.hDC.GetHandleOutput(), (x1, y1, x2, y2))
                    self.paginasImpressas += 1
                    print("Pagina: {}".format(pagina))
                    break
                except:
                    tentativas += 1
                    time.sleep(3)
                    traceback.print_exc()
                    if(tentativas == max_tentativas):
                        self.error_state = True
                        return
                    #messagebox.showerror("Error", "Erro de impressão -1")
                finally:
                    self.hDC.EndPage()
                
            if(indice==len(self.listadepaginas)-1):
                
                None
            else:
                #print('proxima pagina')
                root.after(100, lambda :self.imprimirPaginas(indice+1))
        except:
            
            traceback.print_exc()
            #messagebox.showerror("Error", "Erro de impressão")
            return
        finally:
            None
            #retorno.put((error_state, paginasImpressas))
            #root.update_idletasks()
            #time.sleep(.200)
            #hDC.EndPage()
            #win32print.EndDocPrinter (hPrinter)
        
                
    def printerProcess(self, impressora, cor, filename, listadepaginas): 
        self.error_state = False
        self.paginasImpressas = 0
        if plt=="Windows":
            PRINTER_DEFAULTS = {"DesiredAccess":win32print.PRINTER_ALL_ACCESS}
            self.hPrinter  = win32print.OpenPrinter(impressora, PRINTER_DEFAULTS)
            devmode = win32print.GetPrinter(self.hPrinter, 2)
            devmode["pDevMode"].Color = 1 if cor == "PRETO/BRANCO" else 2
            devmode["pDevMode"].Copies = 1
            devmode["pDevMode"].PrintQuality = 600
            try:                    
                win32print.SetPrinter(self.hPrinter, 2, devmode, 0)
            except:
                print("win32print.SetPrinter: settings could not be changed")
            
            hDC1 = win32gui.CreateDC("WINSPOOL", impressora, devmode["pDevMode"])
            self.hDC = win32ui.CreateDCFromHandle(hDC1)
            #self.docprintx = fitz.open(filename)
            #messagebox.showerror("Info", "printer_margins")
            #root = Tk()
            try:
                self.hDC.CreatePrinterDC(impressora)
                printable_area = self.hDC.GetDeviceCaps (HORZRES), self.hDC.GetDeviceCaps (VERTRES)
                printer_size = self.hDC.GetDeviceCaps (PHYSICALWIDTH), self.hDC.GetDeviceCaps (PHYSICALHEIGHT)
                printer_margins = self.hDC.GetDeviceCaps (PHYSICALOFFSETX), self.hDC.GetDeviceCaps (PHYSICALOFFSETY)
                #
                try:
                    self.hDC.StartDoc (filename)
                    #messagebox.showerror("Error", "start doc")
                    self.error_state = False
                    self.paginasImpressas = 0
                    self.imprimirPaginas(0)
                except:
                    try:
                        self.hDC.EndDoc()
                    except:
                        None
                    traceback.print_exc()
                    #messagebox.showerror("Error", "Erro de impressão 2")
                finally:
                    None
            except:
                traceback.print_exc()
            finally:
                
                

                None
        if plt=="Linux":
            try:
                vet = ["lp", "-d", impressora, "-o", "fit-to-page", "-o", "media=A4", filename,  "-o", "page-ranges="+self.pagvar.get()]
                print(vet)
                subprocess.Popen(["lp", "-d", impressora, "-o", "fit-to-page", "-o", "media=A4", filename,  "-o", "page-ranges="+self.pagvar.get()])
                self.on_quit()
            except:
                traceback.print_exc()
            None
    
            
def go():
    global root
    try:
        global root
        root = tkinter.Tk()
        printerapp = Printer()
        _filename = sys.argv[1]
        widthdoc = int(sys.argv[2])
        heightdoc = int(sys.argv[3])
        atual = int(sys.argv[4])
        #root = None
        printerapp.printFunction(_filename, widthdoc, heightdoc, atual, root)
        root.protocol("WM_DELETE_WINDOW", printerapp.on_quit)
        root.mainloop()
    except:
        traceback.print_exc()
        
if __name__ == '__main__':
    mp.freeze_support()
    go()
    

