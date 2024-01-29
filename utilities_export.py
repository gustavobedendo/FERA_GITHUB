# -*- coding: utf-8 -*-
"""
Created on Tue Feb  1 13:49:31 2022

@author: labinfo
"""
import tkinter, subprocess, PIL
import global_settings
from tkinter import ttk
from functools import partial
from tkinter.filedialog import askdirectory
import fitz, os
import utilities_general, classes_general
from pathlib import Path
import shutil
import binascii, platform
import clipboard
import csv
from PIL import Image, ImageFile
ImageFile.LOAD_TRUNCATED_IMAGES = True
from tkinter.filedialog import asksaveasfilename
plt = platform.system()

class ExportInterval():
    def show(self):
        self.pathdoc = global_settings.pathpdfatual
        self.doctoexport['text'] = global_settings.pathpdfatual
        self.window.deiconify()
    def __init__(self, root):
        self.root = root
        #self.pathdoc = pathpdfatual
        self.window = tkinter.Toplevel()  
        self.window.rowconfigure(1, weight=1)
        self.window.columnconfigure((0,1), weight=1)
        self.window.protocol("WM_DELETE_WINDOW", self.exportCancel)
        #print(pathpdfatual)
        self.doctoexport = tkinter.Label(self.window, font=global_settings.Font_tuple_Arial_10)
        self.doctoexport.grid(row=0, column=0, columnspan=2, sticky='nsew', pady=5, padx=5)
        
        self.exportframeinit = tkinter.Frame(self.window)
        self.exportframeinit.grid(row=1, column=0, sticky='nsew', pady=5, padx=5)
        self.exportframeinit.rowconfigure((0,1), weight=1)
        self.exportframeinit.columnconfigure((0,1), weight=1)
        
        self.exportframeend = tkinter.Frame(self.window)
        self.exportframeend.grid(row=1, column=1, sticky='nsew', pady=5, padx=5)
        self.exportframeend.rowconfigure((0,1), weight=1)
        self.exportframeend.columnconfigure((0,1), weight=1)
        
        self.initpageVar = tkinter.IntVar()
        self.inityVar = tkinter.IntVar()
        self.endpageVar = tkinter.IntVar()
        self.endyVar = tkinter.IntVar()
        self.initpagelabel = tkinter.Label(self.exportframeinit, font=global_settings.Font_tuple_Arial_10, text="Pagina inicial:")
        self.initpagelabel.grid(row=0, column=0, sticky='e', pady=5, padx=5)
        self.initpage = tkinter.Entry(self.exportframeinit, justify='center', textvariable=self.initpageVar, exportselection=False)
        self.initpage.grid(row=0, column=1, sticky='w', pady=5, padx=5)
        self.initylabel = tkinter.Label(self.exportframeinit, font=global_settings.Font_tuple_Arial_10, text="Pos. Y inicial:")
        self.initylabel.grid(row=1, column=0, sticky='e', pady=5, padx=5)
        self.inity = tkinter.Entry(self.exportframeinit, justify='center', textvariable=self.inityVar, exportselection=False)
        self.inity.grid(row=1, column=1, sticky='w', pady=5, padx=5)
        
        self.endpagelabel = tkinter.Label(self.exportframeend, font=global_settings.Font_tuple_Arial_10, text="Pagina final:")
        self.endpagelabel.grid(row=0, column=0, sticky='e', pady=5, padx=5)
        self.endpage = tkinter.Entry(self.exportframeend, justify='center', textvariable=self.endpageVar, exportselection=False)
        self.endpage.grid(row=0, column=1, sticky='w', pady=5, padx=5)
        self.endylabel = tkinter.Label(self.exportframeend, font=global_settings.Font_tuple_Arial_10, text="Pos. Y final:")
        self.endylabel.grid(row=1, column=0, sticky='e', pady=5, padx=5)
        self.endy = tkinter.Entry(self.exportframeend, justify='center', textvariable=self.endyVar, exportselection=False)
        self.endy.grid(row=1, column=1, sticky='w', pady=5, padx=5)
        #botaoaplicar = None
        self.botaoaplicar = tkinter.Button(self.window, font=global_settings.Font_tuple_Arial_10, text='Exportar')
        self.botaoaplicar['command'] =  partial(self.exportOk, self.initpageVar, self.inityVar, self.endpageVar, self.endyVar, self.botaoaplicar)
        self.botaocancelar = tkinter.Button(self.window, font=global_settings.Font_tuple_Arial_10, text='Cancelar', command= self.exportCancel)
        self.botaoaplicar.grid(row=3, column=0, sticky='nsew', pady=5, padx=5)
        self.botaocancelar.grid(row=3, column=1, sticky='nsew', pady=5, padx=5)
        self.progresssearch = ttk.Progressbar(self.window, mode='determinate', maximum = 1)
        self.progresssearch.grid(row=2, column=0, columnspan=2, sticky='nsew', pady=5, padx=5)
        self.count = 0
        self.progresssearch['value'] = self.count
        self.docatual = None
        
        
    def exportOk(self, pageinit, yinit, pageend, yend, botaoaplicar):        
        self.pathtosave = (askdirectory(initialdir=global_settings.pathdb.parent))
        
        if(self.docatual==None):
            self.docatual = fitz.open(self.pathdoc)
        self.progresssearch['maximum'] = pageend.get() - pageinit.get() +1
        self.count = 0
        self.window.lift()
        self.exporting = True
        pagenow = self.initpageVar.get()
        self.botaoaplicar.config(state='disabled')
        self.root.after(1, lambda h=pagenow : self.exportRootAfter(h))
       
    def exportCancel(self, event=None):
        self.window.withdraw()
        self.botaoaplicar.config(state='normal')
        self.exporting = False
        self.count = 0
        self.progresssearch['value'] = self.count
        try:
            self.docatual.close()
        except:
            None
        self.docatual = None
        #self.window = None
        
    def exportRootAfter(self, pagenow):
        self.window.lift()
        if(pagenow > self.endpageVar.get()):
            self.window.withdraw()
            self.count = 0
            self.progresssearch['value'] = self.count
            self.botaoaplicar.config(state='normal')
            self.exporting = False
            try:
                self.docatual.close()
            except:
                None
            self.docatual = None
            return
            #self.window = None
        try:
            loadedPage = self.docatual[pagenow-1]
            
            links = loadedPage.get_links()
            for link in links:
                r = link['from']
                if('file' not in link):
                    continue
                try:
                    arquivo  = link['file']
                    pdfatualnorm = self.pathdoc
                    filepath = str(Path(os.path.join(Path(pdfatualnorm).parent,arquivo)))
                    if(self.initpageVar.get()==self.endpageVar.get()):
                        if(r.y0 >= self.inityVar.get() and r.y0 <= self.endyVar.get()):
                            shutil.copy2(filepath, os.path.join(self.pathtosave, os.path.basename(filepath)))
                    else:
                       
                        if(pagenow==self.initpageVar.get()):
                            if(r.y0 >= self.inityVar.get()):
                                shutil.copy2(filepath, os.path.join(self.pathtosave, os.path.basename(filepath)))
                        elif(pagenow==self.endpageVar.get()):
                            if(r.y0 <= self.endyVar.get()):
                                shutil.copy2(filepath, os.path.join(self.pathtosave, os.path.basename(filepath)))
                        else:
                            shutil.copy2(filepath, os.path.join(self.pathtosave, os.path.basename(filepath)))
                except Exception as ex:
                    utilities_general.printlogexception(ex=ex)  
            self.count += 1
            self.progresssearch['value'] = self.count
            pagenow = pagenow+1
            if(self.exporting):
                self.root.after(1, lambda h=pagenow : self.exportRootAfter(h))
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
            
class Export_Images_To_Table():
    def __init__(self, fera, withnnotes):
        try:
            window = tkinter.Toplevel()
            window.columnconfigure(0, weight=1)
            window.rowconfigure((0,1,2), weight=1)
            label = tkinter.Label(window, font=global_settings.Font_tuple_Arial_10, text="Numero de colunas")
            label.grid(row=0, column=0, sticky='ew', padx=50, pady=10)
            current_value = tkinter.IntVar(value=2)
            spin = tkinter.Scale(window, from_=1, to=4, variable=current_value, orient=tkinter.HORIZONTAL)
            spin.grid(row=1, column=0, sticky='ew', padx=50, pady=10)
            def exit_btn():
               window.destroy()
            button_close = tkinter.Button(window, font=global_settings.Font_tuple_Arial_10, text="OK", command= exit_btn)
            button_close.grid(row=2, column=0, sticky='ew', padx=50, pady=10)                
            global_settings.root.wait_window(window)
            try:
                window.destroy()
            except:
                None
            qtascols = int(current_value.get())
            intervalo = int(9070 / float(current_value.get()))
            largura = int(9070 / float(current_value.get()))
            rtftotal = bytearray("", 'utf8')                    
            pathdocespecial = None
            docespecial = None
            pagatual = None
            xmlatual = None
            rootxml = None
            textonatabela = ""
            pinit = None
            pfim = None
            textoselecao = "(...)\\line"
            listalinks = {}
            listaaprocessar = []                
            lista = []
            imagens= []
            images = []                
            anexosset = []
            anexos = ""
            for iid in fera.treeviewObs.selection():                    
                fera.descendants_obsitem(iid, lista)
            sorted_by_second = sorted(lista, key=lambda tup: tup[1])
            for iidx in sorted_by_second:
                pdf = iidx[1]
                child2 = iidx[0]
                valoresPecial = fera.treeviewObs.item(child2, 'values')
                pagi = int(valoresPecial[2].strip())+1
                if(pathdocespecial!=valoresPecial[1]):
                    pathdocespecial = utilities_general.get_normalized_path(valoresPecial[1])
                    pagatual = None
                    if(docespecial!=None):
                        docespecial.close()
                    docespecial = fitz.open(pathdocespecial)
                tiposelecao = valoresPecial[0]                    
                pinit = int(valoresPecial[2])      
                pfim = int(valoresPecial[5])
                pinit2 = min(pinit, pfim)
                pfim2 = max(pfim, pinit)
                p0xinit = (int(float(valoresPecial[3])))
                p0yinit = (int(float(valoresPecial[4])))-2
                p1xinit = (int(float(valoresPecial[6])))
                p1yinit = (int(float(valoresPecial[7])))+2
                listaaprocessar.append((pathdocespecial, ))                    
                margemsup = (global_settings.infoLaudo[pathdocespecial].mt/25.4)*72
                margeminf = global_settings.infoLaudo[pathdocespecial].pixorgh-((global_settings.infoLaudo[pathdocespecial].mb/25.4)*72)
                margemesq = (global_settings.infoLaudo[pathdocespecial].me/25.4)*72
                margemdir = global_settings.infoLaudo[pathdocespecial].pixorgw-((global_settings.infoLaudo[pathdocespecial].md/25.4)*72)
                size = 0.0666666667*intervalo-10, 450    
                print(size)
                japegos = set()   
                japegos2 = set()
                for pagina in range(pinit2, pfim2+1):
                    p0x = max(p0xinit, margemesq)
                    if(pagina>pinit2):  
                        p0y = int(float(margemsup))
                    else:
                        p0y = max(p0yinit, margemsup)                    
                    p1x = min(p1xinit, margemdir)
                    if(pagina < pfim2):
                        p1y = int(float(margeminf))
                    else:
                        p1y = min(p1yinit, margeminf)
                    links = docespecial[pagina].get_links()
                    for link in links:
                        r = link['from']
                        if('file' in link):
                            file = link['file']
                            rect = fitz.Rect(r.x0-1, r.y0-5, r.x1, r.y1+5)
                            file_on_pdf = docespecial[pagina].get_textbox(rect).strip()
                            
                            #print(file, 'x', file_on_pdf, 'y', rect)
                            mediay = (r.y1 + r.y0) / 2.0    
                            if(p1y >= mediay and p0y <= mediay and (file, pathdocespecial) not in imagens):
                                if("#" in file):
                                    continue
                                
                                file = utilities_general.get_normalized_path(file)
                                filepath = str(Path(utilities_general.get_normalized_path\
                                                    (os.path.join(Path(utilities_general.get_normalized_path(pathdocespecial)).parent,file))))  
                                if(filepath in japegos):
                                    if(not (filepath, '') in japegos2):
                                        continue
                                japegos2.add((filepath, file_on_pdf))
                                japegos.add(filepath)
                                pngtohex = ""
                                content = ""
                                width, height = None, None
                                try:
                                    filename, extension = os.path.splitext(filepath)
                                    if(extension in global_settings.listavidformats):
                                        executavel = os.path.join(utilities_general.get_application_path(), "ffmpeg")
                                        comando = f"\"{executavel}\" -y -ss 1 -i \"{filepath}\" -frames:v 1 -q:v 2 teste.png"  
                                        #popen = subprocess.Popen(comando, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                                        popen = subprocess.run(comando,universal_newlines=True, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True)
                                        #return_code = popen.wait()
                                        if(not os.path.exists("teste.png")):
                                            comando = f"\"{executavel}\" -y -ss 0 -i \"{filepath}\" -frames:v 1 -q:v 2 teste.png"                                              
                                            #popen = subprocess.Popen(comando, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                                            popen = subprocess.run(comando,universal_newlines=True, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True)
                                            #return_code = popen.wait()
                                            lines = popen.stderr.split("\n")
                                            if(len(lines)>0):
                                                raise classes_general.FFMPEGException(lines)
                                        if(os.path.exists("teste.png")):                                               
                                            with Image.open("teste.png") as imx:
                                                imx.thumbnail(size, Image.ANTIALIAS)
                                                imx.save('teste.png','PNG')
                                                width, height = imx.size
                                            with open('teste.png', 'rb') as f:
                                                content = f.read()
                                                pngtohex = binascii.hexlify(content).decode('utf8')
                                            if(os.path.basename(pathdocespecial) not in anexosset):
                                                anexosset.append(os.path.basename(pathdocespecial))
                                    else:
                                        with Image.open(filepath) as imx:
                                            #imx.thumbnail(size,Image.Resampling.LANCZOS)
                                            
                                            imx.save('teste.png','PNG')
                                            width, height = imx.size
                                        with open('teste.png', 'rb') as f:
                                            content = f.read()
                                            pngtohex = binascii.hexlify(content).decode('utf8')
                                       
                                        
                                        if(os.path.basename(pathdocespecial) not in anexosset):
                                            anexosset.append(os.path.basename(pathdocespecial))
                                    imagens.append((filepath, pathdocespecial, pagina, pngtohex, width, height))
                                except PIL.UnidentifiedImageError:
                                    None
                                except Exception as ex:
                                    None
                                    utilities_general.printlogexception(ex=ex)
                                finally:
                                    try:
                                        os.remove('teste.png')
                                    except:
                                        None
                                #utilities_general.printlogexception(ex=ex)

            anexosmap = {}
            superscript = ""
            for i in range(0, len(anexosset)-1):
                anexosmap[anexosset[i]] = i+1
                if(len(anexosset)>1):
                    superscript = "{{\\hich\\af2\\loch\\super\\b\\fs32\\f2\\loch ({})}}".format(i+1)
                anexos += superscript + "\\\'22" + anexosset[i] + "\\\'22" + "{{\\fs22\\f2, }}{{ }}"                
            if(len(anexosset)>1):
                superscript = "{{\\hich\\af2\\loch\\super\\b\\fs32\\f2\\loch ({})}}{{\\fs22\\f2}}".format(len(anexosset))
            anexos += superscript + "\\\'22" + anexosset[len(anexosset)-1] + "\\\'22"
            anexosmap[anexosset[len(anexosset)-1]] = len(anexosset)
            docpagina = "{{\\fs22\\f2IMAGENS DO ANEXO {}}}".format(anexos)  
            if(len(anexosset)>1):
                docpagina = "{{\\fs22\\f2IMAGENS DOS ANEXOS:\\line {}}}".format(anexos) 
            if plt == 'Windows' or  plt == 'Linux':
                textonatabela = "\\pard\\fs22\\f2\\par\\pard\\fs22\\f2\\tcelld\\fs22\\f2\\clbrdrb\\brdrs\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrr\\brdrs\\trautofit1\\intbl\\clftsWidth3\\clwWidth9070\\cellx9070{{\\cbpat2\\qc\\loch\\b\\fs22\\f2 TABELA }}{{\\qc\\f2\\b\\fs22\\field{{\\*\\fldinst  SEQ Tabela \\\\* ARABIC }}{{\\b\\fldrslt}}}}"+\
                "{{\\b\\qc:}}{{ {}}}\\cell\\row\\pard".format(docpagina)
                cont = 1
                picts = ""
                poscols = "\\tcelld\\clbrdrb\\brdrs\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrr\\brdrs\\cellx9070"
                poscols
                if qtascols == 2:
                    poscols = "\\tcelld\\clbrdrb\\brdrs\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrr\\brdrs\\cellx4535\\tcelld\\clbrdrb\\brdrs\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrr\\brdrs\\cellx9070"
                elif qtascols == 3:
                    poscols = "\\tcelld\\clbrdrb\\brdrs\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrr\\brdrs\\cellx3023\\tcelld\\clbrdrb\\brdrs\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrr\\brdrs\\cellx6047\\tcelld\\clbrdrb\\brdrs\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrr\\brdrs\\cellx9070"
                elif qtascols == 4:
                    poscols = "\\tcelld\\clbrdrb\\brdrs\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrr\\brdrs\\cellx2267\\tcelld\\clbrdrb\\brdrs\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrr\\brdrs\\cellx4535\\tcelld\\clbrdrb\\brdrs\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrr\\brdrs\\cellx6802\\tcelld\\clbrdrb\\brdrs\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrr\\brdrs\\cellx9070"
                textonatabela += poscols
                japroc = set()
                for im in imagens:                        
                    if((im[0], im[1]) in japroc):
                        continue                        
                    japroc.add((im[0], im[1]))
                    try:                            
                        if(cont>qtascols):
                            textonatabela += "\\row"
                            textonatabela += "\\trautofit1\\intbl\\clftsWidth3\\clwWidth9070"
                            textonatabela += poscols
                            cont = 1
                        #twips = cont * largura
                        img = im[0]
                        #pdforg = im[1]
                        width, height = im[4], im[5]
                        prop = height / width 
                        twipsw = round(9070 / qtascols) -(56 * 6)
                        twipsh = round(prop * twipsw)
                        if(twipsh / 11200 > 1):
                            twipsh = 11200
                            twipsw /= 11200
                        #print(twipsw, twipsh)
                        pngtohex = im[3]
                        superscript = ""
                        if(len(anexosset)>1):
                            superscript = "{{\\hich\\af2\\loch\\super\\b\\fs28\\f2\\loch ({})\\fs22\\f2}}".format(anexosmap[os.path.basename(im[1])])
                        flslegenda = "{{ }}{{\\fs18\\f2(Fls. {})}}".format(str(im[2]+1))+superscript
                        legenda = r'\qc\sa120{{\hich\af2\loch\fs18\b\f2\loch Imagem }}{{{{\fs18\f2\field{{\*\fldinst  SEQ Figura \\* ARABIC }}{{\b\fldrslt{{ }}}}}}{{\fs18\f2\b\qc:{{ }}}}{{\fs18\f2\i"{}"}}}}'.format(os.path.basename(img))
                        
                        picts = "\\clbrdrb\\brdrs\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrr\\brdrs{{\\fs18\\f2\\pict\\picscalex100\\picscaley100\\piccropl0\\piccropr0\\piccropt0\\piccropb0\\picwgoal{}\\pichgoal{}\\pngblip\n{}}}\\line{}\\cell".format(twipsw, twipsh, pngtohex, legenda + flslegenda)                           
                        textonatabela += picts
                        cont += 1
                    except Exception as ex:
                        utilities_general.printlogexception(ex=ex)
                textonatabela += "\\row\\line"
                textofinal = ("{{\\rtf1\\ansi\\deff0{{\\fonttbl{{\\f0\\froman\\fprq2\\fcharset0 Times New Roman;}}{{\\f1\\froman\\fprq2\\fcharset2 Symbol;}}"+
               "{{\\f2\\fswiss\\fprq2\\fcharset0 Arial;}}}}{{\\colortbl;\\red240\\green240\\blue240;\\red221\\green221\\blue221;\\red255\\green255\\blue255;}} {}}}").format(textonatabela)
                conteudo = bytearray(textofinal, 'utf8')
                utilities_general.copy_to_clipboard("rtf", conteudo)                  
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        finally:
            try:
                docespecial.close()
            except Exception as ex:
                None

class Export_to_Clipboard():
    def __init__(self, mw, iid):
        texto = ""
        if(mw.treeviewObs.tag_has('obscat', iid)):
            children = mw.treeviewObs.get_children(iid)
            
            for pdf in children:
                paginas = set()
                texto += mw.treeviewObs.item(pdf, 'text').strip() + "\n"
                children2 = mw.treeviewObs.get_children(pdf)
                if(len(children2)>0):
                    for toc in children2: 
                        texto += mw.treeviewObs.item(toc, 'text').strip() + "\n"
                        children3 = mw.treeviewObs.get_children(toc)
                        primeiro = True
                        for child2 in children3:
                            pagi = int(mw.treeviewObs.item(child2, 'values')[2].strip())+1                            
                            if(not pagi in paginas):
                                if(primeiro):                                    
                                    texto += str(int(mw.treeviewObs.item(child2, 'values')[2].strip())+1)
                                    primeiro = False
                                else:
                                    texto += ", "+  str(int(mw.treeviewObs.item(child2, 'values')[2].strip())+1)
                            paginas.add(pagi)
                        texto += "\n"
                else:
                    primeiro = True
                    for child2 in children2:
                        pagi = int(mw.treeviewObs.item(child2, 'values')[2].strip())+1                        
                        if(not pagi in paginas):
                            if(primeiro):                                
                                texto += str(int(mw.treeviewObs.item(child2, 'values')[2].strip())+1)
                                primeiro = False
                            else:
                                texto += ", "+ str(int(mw.treeviewObs.item(child2, 'values')[2].strip())+1)
                        paginas.add(pagi)
                        texto += "\n"
                texto += "\n"
        elif(mw.treeviewObs.tag_has('relobs', iid)):
            pdf = iid
            children2 = mw.treeviewObs.get_children(pdf)
            if(len(children2)>0):
                paginas = set()
                for toc in children2: 
                    texto += mw.treeviewObs.item(toc, 'text').strip() + "\n"
                    children3 = mw.treeviewObs.get_children(toc)
                    primeiro = True
                    for child2 in children3:
                        pagi = int(mw.treeviewObs.item(child2, 'values')[2].strip())+1                            
                        if(not pagi in paginas):
                            if(primeiro):                                    
                                texto += str(int(mw.treeviewObs.item(child2, 'values')[2].strip())+1)
                                primeiro = False
                            else:
                                texto += ", "+  str(int(mw.treeviewObs.item(child2, 'values')[2].strip())+1)
                        paginas.add(pagi)
                    texto += "\n"
            else:
                paginas = set()
                primeiro = True
                for child2 in children2:
                    pagi = int(mw.treeviewObs.item(child2, 'values')[2].strip())+1                        
                    if(not pagi in paginas):
                        if(primeiro):                                
                            texto += str(int(mw.treeviewObs.item(child2, 'values')[2].strip())+1)
                            primeiro = False
                        else:
                            texto += ", "+ str(int(mw.treeviewObs.item(child2, 'values')[2].strip())+1)
                    paginas.add(pagi)
                    texto += "\n"
            texto += "\n"
        elif(mw.treeviewObs.tag_has('tocobs', iid)):
            paginas = set()
            primeiro = True
            pdf = mw.treeviewObs.parent(iid)
            toc = iid
            children3 = mw.treeviewObs.get_children(toc)
            for child2 in children3:
                pagi = int(mw.treeviewObs.item(child2, 'values')[2].strip())+1                            
                if(not pagi in paginas):
                    if(primeiro):                                    
                        texto += str(int(mw.treeviewObs.item(child2, 'values')[2].strip())+1)
                        primeiro = False
                    else:
                        texto += ", "+  str(int(mw.treeviewObs.item(child2, 'values')[2].strip())+1)
                paginas.add(pagi)
            texto += "\n"
        clipboard.copy(texto.strip()) 

class Export_to_CSV():
    def __init__(self, mw, iid):
        tipos = [('CSV', '*.csv')]
        path = (asksaveasfilename(filetypes=tipos, defaultextension=tipos))
        if(path!=None and path!=''):
            with open(path, mode='w', newline='', encoding='utf-8') as csv_file:
                writer = csv.writer(csv_file, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                writer.writerow(['Documento', 'Seção', 'Página'])
                children = mw.treeviewObs.get_children(iid)
                if(mw.treeviewObs.tag_has('obscat', iid)):
                    for pdf in children:
                        paginas = set()
                        children2 = mw.treeviewObs.get_children(pdf)
                        if(len(children2)>0):
                            for toc in children2:
                                children3 = mw.treeviewObs.get_children(toc)
                                for child2 in children3:
                                    pagi = int(mw.treeviewObs.item(child2, 'values')[2].strip())+1                            
                                    if(not pagi in paginas):
                                        writer.writerow([mw.treeviewObs.item(pdf, 'text').strip(), mw.treeviewObs.item(toc, 'text').strip(),\
                                                         int(mw.treeviewObs.item(child2, 'values')[2].strip())+1])
                                    paginas.add(pagi)
                        else:
                            for child2 in children2:
                                pagi = int(mw.treeviewObs.item(child2, 'values')[2].strip())+1                        
                                if(not pagi in paginas):
                                    writer.writerow([mw.treeviewObs.item(pdf, 'text').strip(), '-', int(mw.treeviewObs.item(child2, 'values')[2].strip())+1])
                                paginas.add(pagi)
                elif(mw.treeviewObs.tag_has('relobs', iid)):
                    pdf = iid
                    children2 = mw.treeviewObs.get_children(pdf)
                    if(len(children2)>0):
                        paginas = set()
                        for toc in children2:
                            children3 = mw.treeviewObs.get_children(toc)
                            for child2 in children3:
                                pagi = int(mw.treeviewObs.item(child2, 'values')[2].strip())+1                            
                                if(not pagi in paginas):
                                    writer.writerow([mw.treeviewObs.item(pdf, 'text').strip(), mw.treeviewObs.item(toc, 'text').strip(),\
                                                     int(mw.treeviewObs.item(child2, 'values')[2].strip())+1])
                                paginas.add(pagi)
                    else:
                        paginas = set()
                        for child2 in children2:
                            pagi = int(mw.treeviewObs.item(child2, 'values')[2].strip())+1                        
                            if(not pagi in paginas):
                                writer.writerow([mw.treeviewObs.item(pdf, 'text').strip(), '-',
                                                 int(mw.treeviewObs.item(child2, 'values')[2].strip())+1])
                            paginas.add(pagi)
                elif(mw.treeviewObs.tag_has('tocobs', iid)):
                    pdf = mw.treeviewObs.parent(iid)
                    toc = iid
                    children3 = mw.treeviewObs.get_children(toc)
                    paginas = set()
                    for child2 in children3:
                        pagi = int(mw.treeviewObs.item(child2, 'values')[2].strip())+1                            
                        if(not pagi in paginas):
                            writer.writerow([mw.treeviewObs.item(pdf, 'text').strip(), mw.treeviewObs.item(toc, 'text').strip(),
                                             int(mw.treeviewObs.item(child2, 'values')[2].strip())+1])
                        paginas.add(pagi)
 
        