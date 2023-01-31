# -*- coding: utf-8 -*-
"""
Created on Tue Feb  1 13:52:40 2022

@author: labinfo
"""
import global_settings, utilities_general, classes_general, process_functions
import sys, os
import traceback
from pathlib import Path
import threading as thr
import sqlite3
import tkinter, math, re, fitz, time
from PIL import Image, ImageTk, ImageFile
ImageFile.LOAD_TRUNCATED_IMAGES = True
import hashlib, subprocess, platform
try:
    import win32clipboard
except:
    None
plt = platform.system()

def initiate_indexing_thread():
    if(global_settings.indexing_thread==None or not global_settings.indexing_thread.is_alive()):
        global_settings.indexing_thread = thr.Thread(target=process_functions.indexing_thread_func, daemon=True)
        global_settings.indexing_thread.start()
        print('Starting: Indexing Thread')
        #break

def popup_window(texto, sair):
    
    window = tkinter.Toplevel()
    label = tkinter.Label(window, font=global_settings.Font_tuple_Arial_10, text=texto, image=global_settings.warningimage, compound='top')
    label.pack(fill='x', padx=50, pady=20)
    button_close = tkinter.Button(window, font=global_settings.Font_tuple_Arial_10, text="OK", command= lambda : popupcomandook(sair, window))
    button_close.pack(fill='y', pady=20) 
    return window


#def log_window(texto):
#    
#    window = tkinter.Toplevel()
#    label = tkinter.Label(window, font=global_settings.Font_tuple_Arial_10, text=texto, image=global_settings.warningimage, compound='top')
##    label.pack(fill='x', padx=50, pady=20)
#    button_close = tkinter.Button(window, font=global_settings.Font_tuple_Arial_10, text="OK", command= lambda : popupcomandook(sair, window))
#    button_close.pack(fill='y', pady=20) 
#    return window

def popupcomandook(sair, window):
    
    if(sair):
        try:
            window.destroy()
            global_settings.on_quit()
        except:
            None
        finally:
            global_settings.root.destroy()
            os._exit(1)
    else:
        window.destroy()
        


def extract_links_from_page(doc, idpdf, idobs, pathpdf, paginainit, paginafim, p0x, p0y, p1x, p1y):
    p0x = round(p0x)
    p1x = round(p1x)
    p0y = round(p0y)
    p1y = round(p1y)
    links_tratados = []
    
    for p in range(paginainit, paginafim+1):
        '''
        if(paginainit!=paginafim and False):
            if(p==paginainit):
                p1y = global_settings.infoLaudo[pathpdf].mb
            elif(p==paginafim):
                p0y = global_settings.infoLaudo[pathpdf].mt
            else:
                p0y = global_settings.infoLaudo[pathpdf].mt
                p1y = global_settings.infoLaudo[pathpdf].mb
            if(p0y==p1y):
                p0y -= 1
                p1y += 1
        '''
        loadedpage_links = doc[p].get_links()
        for link in loadedpage_links:
            #print(link)
            r = link['from']
            rymedio = math.ceil((math.ceil(r.y1) + math.ceil(r.y0))/2.0)
            link_tratado = None
            if((paginainit==paginafim and p0y <= rymedio and p1y >= rymedio) or
               (p==paginainit and paginainit!=paginafim  and p0y <= rymedio) or \
               (p==paginafim and paginainit!=paginafim and p1y >= rymedio)  or \
               (p!=paginainit and paginainit!=paginafim and p!=paginafim)):
                
                if('file' in link):
                    link_tratado = link['file']
                    #if(link['file'] not in dict_of_anottations):
                    #    dict_of_anottations[link_tratado] = []
                    #dict_of_anottations[link_tratado].append((r.x0, r.y0, r.x1, r.y1, ''))
                elif('to' in link):
                    link_tratado =link['to']
                    if("Point" in str(link_tratado)):
                        link_tratado = None
                    #if(link['to'] not in dict_of_anottations):
                    #    dict_of_anottations[link['to']] = []
                    #dict_of_anottations[link['to']].append((r.x0, r.y0, r.x1, r.y1, ''))
                elif('uri' in link):
                    link_tratado =link['uri']
                    #if(link['uri'] not in dict_of_anottations):
                    #    dict_of_anottations[link['uri']] = []
                    #dict_of_anottations[link['uri']].append((r.x0, r.y0, r.x1, r.y1, ''))
                
            if(link_tratado!=None):        
                
                links_tratados.append((idpdf, idobs, paginainit, math.ceil(r.x0), \
                                              math.ceil(r.y0), paginafim, math.floor(r.x1), math.floor(r.y1), link_tratado, '',))
    return    links_tratados       
                    
        
def below_right(win, dist):
    """
    centers a tkinter window
    :param win: the main window or Toplevel window to center
    """
    win.update_idletasks()
    width = win.winfo_width()
    frm_width = win.winfo_rootx() - win.winfo_x()
    win_width = width + 2 * frm_width
    height = win.winfo_height()
    titlebar_height = win.winfo_rooty() - win.winfo_y()
    win_height = height + titlebar_height + frm_width
    x = win.winfo_screenwidth() // 2 - win_width // 2
    #y = who.winfo_rooty()
    win.geometry('{}x{}-{}+{}'.format(width, height, 0, dist))
    win.deiconify()  

def below_right_edge(win, dist):
    """
    centers a tkinter window
    :param win: the main window or Toplevel window to center
    """
    win.update_idletasks()
    width = win.winfo_width()
    #frm_width = win.winfo_rootx() - win.winfo_x()
    #win_width = width + 2 * frm_width
    height = win.winfo_height()
    #titlebar_height = win.winfo_rooty() - win.winfo_y()
    #win_height = height + titlebar_height + frm_width
    #x = win.winfo_screenwidth() // 2 - win_width // 2
    #y = who.winfo_rooty()
    win.geometry('{}x{}-{}+{}'.format(width, height, width+10, dist))
    win.deiconify() 

def center(win):
    """
    centers a tkinter window
    :param win: the main window or Toplevel window to center
    """
    win.update_idletasks()
    width = win.winfo_width()
    frm_width = win.winfo_rootx() - win.winfo_x()
    win_width = width + 2 * frm_width
    height = win.winfo_height()
    titlebar_height = win.winfo_rooty() - win.winfo_y()
    win_height = height + titlebar_height + frm_width
    x = win.winfo_screenwidth() // 2 - win_width // 2
    y = win.winfo_screenheight() // 2 - win_height // 2
    win.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    win.deiconify()   
        
def popup_window(texto, sair, imagepcp=None):
    global warningimage, windowpopup
    try:
        windowpopup.destroy()
        windowpopup = None
    except Exception as ex:
        None
    windowpopup = tkinter.Toplevel()
    windowpopup.focus_set()
    #w = 300 # width for the Tk root
    #h = 200 # height for the Tk root
    if(imagepcp!=None):
        label = tkinter.Label(windowpopup, font=global_settings.Font_tuple_Arial_10, text=texto, image=imagepcp, compound='top')
    else:
        label = tkinter.Label(windowpopup, font=global_settings.Font_tuple_Arial_10, text=texto, image=global_settings.warningimage, compound='top')
    label.pack(fill='x', padx=5, pady=5)
    # get screen width and height
    #ws = global_settings.root.winfo_screenwidth() # width of the screen
    #hs = global_settings.root.winfo_screenheight() # height of the screen
    
    # calculate x and y coordinates for the Tk root window
    #x = (ws/2) - (w/2)
    #y = (hs/2) - (h/2)
    #window.geometry('%dx%d+%d+%d' % (w, h, x, y))

    button_close = tkinter.Button(windowpopup, font=global_settings.Font_tuple_Arial_10, text="OK", command= lambda : popupcomandook(sair, windowpopup))
    button_close.pack(fill='y', pady=20) 
    windowpopup.bind('<Return>',  lambda e: popupcomandook(sair, windowpopup))
    windowpopup.bind('<Escape>',  lambda e: popupcomandook(sair, windowpopup))
    return windowpopup

def printlogexception(printorlog='print', ex=None):
    
    #if(global_settings.log_window==None):
    try:
        global_settings.log_window_text.insert('end', traceback.format_exc())
        global_settings.log_window_text.insert('end',"\n")
        global_settings.label_warning_error.config(bg='red')
    except:
        None
    
    if(printorlog=='log'):
        global_settings.logging.exception('!')
    elif(printorlog=='print'):
        exc_type, exc_value, exc_tb = sys.exc_info()
        traceback.print_exception(exc_type, exc_value, exc_tb)
    
    #else:
    #    None   
def get_application_path():
    application_path = None
    if getattr(sys, 'frozen', False):
        application_path = sys._MEIPASS
    elif __file__:
        application_path = os.path.dirname(os.path.abspath(__file__))
    return application_path

def get_normalized_path(path):
    pdfrep = str(path)
    if global_settings.plt == "Linux":
        pdfrep = pdfrep.replace("\\","/")
    elif global_settings.plt=="Windows":
        pdfrep = pdfrep.replace("/","\\")
    return os.path.normpath(pdfrep)

def concatVertical(images):
    if(len(images) > 0):
        altura = 0
        for im in images:
            altura += im.height
        dst = Image.new('RGB', (images[0].width, altura))
        posicao = 0
        imagem = 0
        while(imagem < len(images)):
            dst.paste(images[imagem], (0, posicao))
            posicao += images[imagem].height
            imagem += 1                
        return dst
    else:
        return None
    

    
def create_rectanglex(x1, y1, x2, y2, color, link=False, withborder=True, transparent=False, **kwargs):
    try:
        if(transparent):
            dst = Image.new('RGBA', (x2-x1, y2-y1))            
            image = Image.new('RGBA', (x2-x1, y2-y1), color)
            dst.paste(image, (0, 0))         
            return ImageTk.PhotoImage(dst)
        elif(link):
            dst = Image.new('RGBA', (x2-x1, y2-y1))
            border1 = Image.new('RGBA', (x2-x1, 1), (35, 129, 166,255)) 
            image = Image.new('RGBA', (x2-x1, y2-y1), color) 
            dst.paste(image, (0, 0))
            dst.paste(border1, (0,  y2-y1-1))
            return ImageTk.PhotoImage(dst)
        elif(not withborder):
            dst = Image.new('RGBA', (x2-x1, y2-y1))            
            image = Image.new('RGBA', (x2-x1, y2-y1), color)
            dst.paste(image, (0, 0))         
            return ImageTk.PhotoImage(dst)
        elif(withborder):
            dst = Image.new('RGBA', (x2-x1, y2-y1))            
            bordertopbottom = Image.new('RGBA', (x2-x1, 1), (0, 0, 0,255)) 
            bordersides = Image.new('RGBA', (1, (y2-y1)), (0, 0, 0,255)) 
            image = Image.new('RGBA', (x2-x1, y2-y1), color)
            dst.paste(image, (0, 0))
            dst.paste(bordertopbottom, (0,  y2-y1-1))
            dst.paste(bordertopbottom, (0,  0))
            dst.paste(bordersides, (x2-x1-1,  0))
            dst.paste(bordersides, (0,  0))            
            return ImageTk.PhotoImage(dst)
    except Exception as ex:
        None

def insertIndex(tree, parent, texto_candidato, index=0):
    children = tree.get_children(parent)
    
    for child in children:
        texto = tree.item(child, 'text')
        if(texto_candidato < texto):
            break
        index += 1
    return index 

def countChildren(treeview, treenode, putcount=True):    
    th = 0           
    if(treeview.tag_has("resultsearch",treenode)):
        th = 1
    else:
        
        for termonode in treeview.get_children(treenode): 
            th += countChildren(treeview, termonode, putcount=putcount) 
        if(treeview.tag_has("relsearchtoc",treenode)):
            textotoc = treeview.item(treenode, 'text')
            if(putcount):
                if(th>=3000):
                    treeview.item(treenode, text=textotoc + ' (' + str(th) + ')*') 
                else:
                    treeview.item(treenode, text=textotoc + ' (' + str(th) + ')') 
        else:
            textoother = treeview.item(treenode, 'text')
            if(putcount):
                treeview.item(treenode, text=textoother + ' (' + str(th) + ')')  
            if(treeview.tag_has("relsearch",treenode)):
                valores = treeview.item(treenode, 'values')
                treeview.item(treenode, values=(valores[0], valores[1], th, textoother,))
    return th

def locateToc(pagina, pdf, p0y=None, init=None, tocpdf=None):
        pdf = utilities_general.get_normalized_path(pdf)
        pdfx = (str(Path(pdf)))
        pdfx = get_normalized_path(pdfx)
        t = 0
        napagina = False
        naoachou = True
        if(init!=None):
            for t in range(len(tocpdf)-1):
                if(pagina >= tocpdf[t][1] and pagina < tocpdf[t+1][1]):
                    naoachou = False
                    break   
                elif(pagina >= tocpdf[t][1] and pagina <= tocpdf[t+1][1]):
                    napagina = True
                    
                if(napagina and tocpdf[t+1][3] > init):  
                    naoachou = False
                    break
            
            if(naoachou):
                if(pagina==0):
                    t=0
                else:
                    t=len(tocpdf)-1
                    
        elif(p0y!=None):
             for t in range(len(tocpdf)-1):
                if(pagina >= tocpdf[t][1] and pagina < tocpdf[t+1][1]):
                    naoachou = False
                    break   
                elif(pagina >= tocpdf[t][1] and pagina <= tocpdf[t+1][1]):
                    napagina = True
                    
                if(napagina and tocpdf[t+1][2] > p0y):  
                    naoachou = False
                    break
            
             if(naoachou):
                if(pagina==0):
                    t=0
                else:
                    t=len(tocpdf)-1
        
        t = min(t, len(tocpdf)-1)
        t = max(0, t)
        tocc=[pdf,'','']
        if(len(tocpdf) > 0 and len(tocpdf[t])>0):
            tocc = tocpdf[t]
        return tocc
  
def iterateXREF_Names(doc, xref, abs_path_pdf, pismm, aprocurar, rereference, rename_dest, regex):
    chaves = doc.xref_get_keys(xref)
    abs_path_pdf = get_normalized_path(abs_path_pdf)
    #regex = "\([A-Za-z0-9\.]+\)[0-9]+\s[0-9]\sR"
    if("Names" in chaves):
        named_kids = doc.xref_get_key(xref, "Names")[1]
        found = regex.findall(named_kids)
        ##print(named_kids)
        for f in found:
            name_dest, reference = f
            ##print(name_dest, aprocurar)
            if(name_dest==aprocurar):
                destination_final = doc.xref_object(int(reference)).split(" ")
                abs_path_pdf = get_normalized_path(abs_path_pdf)
                dest_page = global_settings.infoLaudo[abs_path_pdf].ref_to_page[int(destination_final[1])]
                cropbox = doc.page_cropbox(dest_page)
                return (name_dest, dest_page, math.floor(float(destination_final[5])), math.floor((cropbox.y1-float(destination_final[6]))))
        return None

    elif("Kids" in chaves):
        destinations_kids = doc.xref_get_key(xref, "Kids")
        destinations_limits = doc.xref_get_key(xref, "Limits")
        retorno = None
        
        if(len(destinations_limits)>1):
                
            quaislimites = pismm.findall(destinations_limits[1]) 
            ##print(destinations_limits, quaislimites, aprocurar)
            if('null'==destinations_limits[0]):
                splitted = destinations_kids[1].split(" ")
                grauavore = int(len(splitted)/3)
                for i in range(grauavore):
                    indice = i * 3
                    novoxref = int(splitted[indice].replace("[", "").replace("]", ""))
                    retorno = iterateXREF_Names(doc, novoxref, abs_path_pdf, pismm, aprocurar, rereference, rename_dest, regex)
                    if(retorno != None):
                        break
            elif(len(quaislimites)>1):
                if(aprocurar >= quaislimites[0] and aprocurar <= quaislimites[1]):
                    splitted = destinations_kids[1].split(" ")
                    grauavore = int(len(splitted)/3)
                    for i in range(grauavore):
                        indice = i * 3
                        novoxref = int(splitted[indice].replace("[", "").replace("]", ""))
                        retorno = iterateXREF_Names(doc, novoxref, abs_path_pdf, pismm, aprocurar, rereference, rename_dest, regex)
                        if(retorno != None):
                            break
            elif(len(quaislimites)>0):
                if(aprocurar >= quaislimites[0]):
                    splitted = destinations_kids[1].split(" ")
                    grauavore = int(len(splitted)/3)
                    for i in range(grauavore):
                        indice = i * 3
                        novoxref = int(splitted[indice].replace("[", "").replace("]", ""))
                        retorno = iterateXREF_Names(doc, novoxref, abs_path_pdf, pismm, aprocurar, rereference, rename_dest, regex)
                        if(retorno != None):
                            break
        #elif(len(destinations_limits)==1):
        return retorno  
  
def iteratetreepages(abs_path_doc, doc, numberregex, xref, count):
    objrootpages = doc.xref_get_key(int(xref), "Type")[1]
    abs_path_doc = get_normalized_path(abs_path_doc)
    if(objrootpages=="/Pages"):
        objrootkids = doc.xref_get_key(int(xref), "Kids")[1]
        for indobj, gen in numberregex.findall(objrootkids):
           count = iteratetreepages(abs_path_doc, doc, numberregex, indobj, count)  
        #return count
    elif(objrootpages=="/Page"):
        global_settings.infoLaudo[abs_path_doc].ref_to_page[int(xref)] = count
        count += 1
    return count
        
    
def loadPages(abs_path_pdf, doc, numberregex):
    rootpdf  = doc.pdf_catalog()
    objpagesr = numberregex.findall(doc.xref_get_key(rootpdf, "Pages")[1])[0][0]
    objrootpages = doc.xref_get_key(int(objpagesr), "Type")[1]
    if(objrootpages=="/Pages"):
       objrootkids = doc.xref_get_key(int(objpagesr), "Kids")[1]
       count = 0
       for indobj, gen in numberregex.findall(objrootkids):
           count = iteratetreepages(abs_path_pdf, doc, numberregex, indobj, count)
    else:
        None    
  
def processDocXREF(abs_path_pdf, doc, aprocurar):
    regex = "\(([A-Za-z0-9\.]+)\)([0-9]+)"
    abs_path_pdf = get_normalized_path(abs_path_pdf)
    if(len(global_settings.infoLaudo[abs_path_pdf].ref_to_page)==0):
        numbercompile = re.compile(r"([0-9]+)\s([0-9]+)")
        loadPages(abs_path_pdf, doc, numbercompile)
    rootpdf  = doc.pdf_catalog()
    tupla_names1 = doc.xref_get_key(rootpdf, "Names")
    
    regexismm = r"\(([a-zA-Z0-9_\.\-]+)\)"
    pismm = re.compile(regexismm)
    tupla_dests = doc.xref_get_key(int(tupla_names1[1].split(" ")[0]), "Dests")
    destinations = doc.xref_get_keys(int(tupla_dests[1].split(" ")[0]))
    if("Kids" in destinations):
        rereference = re.compile("[0-9]+\s")
        rename_dest = re.compile("\([A-Za-z0-9\.]+\)")
        regex = re.compile("\(([A-Za-z0-9\.]+)\)([0-9]+)")
        retorno = iterateXREF_Names(doc, int(tupla_dests[1].split(" ")[0]), abs_path_pdf, pismm, aprocurar, rereference, rename_dest, regex)
        
        return retorno
    else:
        regex = re.compile("\(([A-Za-z0-9\.]+)\)([0-9]+)")
        named_kids = doc.xref_get_key(int(tupla_dests[1].split(" ")[0]), "Names")[1]
        found = regex.findall(named_kids)
        
        for f in found:
            name_dest, reference = f
            
            if(name_dest==aprocurar):
                
                destination_final = doc.xref_object(int(reference)).split(" ")

                dest_page = global_settings.infoLaudo[abs_path_pdf].ref_to_page[int(destination_final[1])]
                cropbox = doc.page_cropbox(dest_page)
                return (name_dest, dest_page, math.floor(float(destination_final[5])), math.floor((cropbox.y1-float(destination_final[6]))))
    return None 

def extract_text_from_page(doc, pagina, deslocy, topmargin, bottommargin, leftmargin, rightmargin, \
                           flags=2+64, replace_accent = True, extract_image=False, diretorio_temp_input=None):
    #print("Extracing image:", extract_image)
    if(extract_image):
        flags = 2+4+64
    lowerCodeNoDiff = [ 
      #00-0F #0
       0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,\
       #00-0F #16
       0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,\
       #10-1F #32
       0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,\
       #20-2F #48
       0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,\
       #30-3F #64
       0,  32,  32,  32,  32,  32,  32,  32,  32,  32,  32,  32,  32,  32,  32,  32,\
       #40-4F #80
      32,  32,  32,  32,  32,  32,  32,  32,  32,  32,  32,   0,   0,   0,   0,   0,\
      #50-5F #96
       0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,\
       #60-6F #112
       0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,\
       #70-7F #128
       0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,\
       #80-8F #144
       0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,\
       #90-9F #160
       0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,\
       #A0-AF #176
       0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,\
       #B0-BF #192
       -95, -96, -97, -98, -99,-100,  32,-100, -99,-100,-101,-102, -99,-100,-101,-102,\
     #C0-CF #208
      32, -99, -99,-100,-101,-102,-103,   0,   0,-100,-101,-102,-103,-100,  32,   0,\
      #D0-DF #224
      -127,-128,-129,-130,-131,-132,   0,-132,-131,-132,-133,-134,-131,-132,-133,-134,\
    #E0-EF #240
       0,-131,-131,-132,-133,-134,-135,   0,   0,-132,-133,-134,-135,-132,   0,-134 \
       #F0-FF #256
       ]  
    quadspagina = []
    mapeamento = {}
    dictx = doc[pagina].get_text("rawdict", flags=flags)  
    images_extracted = []
    novotexto = ''
    init = 0
    x0 = -1
    y0 = -1
    x1 = -1
    y1 = -1
    for block in dictx['blocks']:
        if(block['type']==0):
            pontosBlock = block['bbox']
            
            bloco = (math.floor(float(pontosBlock[0])), math.floor(float(pontosBlock[1])), \
                     math.ceil(float(pontosBlock[2])), math.floor(float(pontosBlock[3])), 'text')
            if(pontosBlock[1]>deslocy and pontosBlock[3] > 0):
                break 
            mapeamento[bloco] = {}
            for line in block['lines']:
                pontosLine = line['bbox']
                linha = (math.floor(float(pontosLine[0])), math.ceil(float(pontosLine[1])+1), \
                     math.ceil(float(pontosLine[2])), math.floor(float(pontosLine[3])-1))
                mapeamento[bloco][linha] = []
                for span in line['spans']:
                    a = span["ascender"]
                    d = span["descender"]
                    #r = fitz.Rect(span["bbox"])
                    #r.y1 = r.y1 
                    #r.y0 = r.y0 
                    r = fitz.Rect(span["bbox"])
                    o = fitz.Point(span["origin"])  # its y-value is the baseline
                    r.y1 = o.y - span["size"] * d / (a - d)
                    r.y0 = r.y1 - span["size"]
                    x0 = y0 = x1 = y1 = None
                    for char in span['chars']:
                        bboxchar = char['bbox']
                        bboxxmedio = (bboxchar[0]+bboxchar[2])/2
                        bboxymedio = (bboxchar[1]+bboxchar[3])/2
                        if(bboxxmedio < leftmargin or bboxxmedio > rightmargin or bboxymedio < topmargin or bboxymedio > bottommargin):
                            continue
                        x0 = math.floor(float(bboxchar[0]))
                        #y0 = math.floor(r.y0)
                        y0 = r.y0 -1
                        x1 = math.ceil(float(bboxchar[-2]))
                        #y1 = math.floor(r.y1)
                        y1 = r.y1 -1
                        c = char['c']
                        if(replace_accent):
                            codePoint = ord(c)
                            if(codePoint<256):
                                codePoint += lowerCodeNoDiff[codePoint]
                            c = chr(codePoint)
                        mapeamento[bloco][linha].append((x0, y0, x1, y1, c))
                        quadspagina.append((x0, y0, x1, y1, c))
                        novotexto += c
                        init += 1
                if(len(novotexto) > 0 and novotexto[-1]!=' '):
                    novotexto += ' '
                    quadspagina.append((x0, y0, x1, y1, ' '))
                    init += 1
            if(len(novotexto) > 0 and novotexto[-1]!=' '):
                novotexto += ' '
                quadspagina.append((x0, y0, x1, y1, ' '))
                init += 1
        elif(block['type']==1):
            pontosBlock = block['bbox']
            bloco = (math.floor(float(pontosBlock[0])), math.floor(float(pontosBlock[1])), \
                     math.ceil(float(pontosBlock[2])), math.floor(float(pontosBlock[3])), 'image')
            mapeamento[bloco] = {}
            mapeamento[bloco][bloco] = []
            mapeamento[bloco][bloco].append(bloco)
            #print(extract_image)
            if(extract_image):
                hash_object = hashlib.md5(block['image'])
                image_hash = hash_object.hexdigest()
                bbox_x0= math.floor(int(pontosBlock[0]))
                bbox_y0= math.floor(int(pontosBlock[1]))
                bbox_x1= math.floor(int(pontosBlock[2]))
                bbox_y1= math.floor(int(pontosBlock[3]))
                
                try:
                    None
                    f = open(os.path.join(diretorio_temp_input, "{}.png".format(image_hash)), 'wb')
                    f.write(block['image'])
                    f.close()
                    images_extracted.append((image_hash, bbox_x0, bbox_y0, bbox_x1, bbox_y1, pagina))
                except:
                    traceback.print_exc()
    return (init, novotexto, quadspagina, mapeamento, images_extracted)


def copy_to_clipboard(tipo, conteudo):
    if(tipo=="rtf"):
        if plt == 'Windows':                        
            CF_RTF = win32clipboard.RegisterClipboardFormat("Rich Text Format")
            win32clipboard.OpenClipboard(0)
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(CF_RTF, conteudo)
            win32clipboard.CloseClipboard()
        elif plt == 'Linux':
            subprocess.Popen(['xclip', '-selection', 'clipboard', '-t', 'text/rtf'], stdin=subprocess.PIPE).communicate(conteudo)

def connectDB(dbpath='', timeout=10, maxrepeat=-1,  check_same_thread_arg=True):
    #hasconn = False
    repeat = 0
    #print("Connecting: {}".format(dbpath))
    while(repeat < maxrepeat or maxrepeat==-1):
        try:
            sqliteconn = sqlite3.connect(str(dbpath), timeout=timeout, factory=classes_general.Custom_Database, check_same_thread=check_same_thread_arg)
            return sqliteconn
        except Exception as ex:
            printlogexception(ex=ex)
            repeat += 1
            None
    return None    

def md5(path_pdf):
    hash_md5 = hashlib.md5()
    with open(path_pdf, "rb") as f:
        cont = 0
        if(os.path.getsize(path_pdf)>4096 * 1024 * 8 + 1):
            f.seek(- 4096 * 1024 * 8, 2)
        for chunk in iter(lambda: f.read(4096), b""):
            hash_md5.update(chunk)
    digest = hash_md5.hexdigest()
    print(path_pdf, digest)
    return digest
                
def searchsqlite(tipobusca, termo, pathpdf, pathdb, idpdf, simplesearch = False, queuesair = None, \
                 idtermo = None, idtermopdf = None, erros_queue = None, fixo = None, result_queue = None,\
                     jarecords=None, sqliteconnx=None, tocs_pdf=None, listaTERMOS=None):
    def re_fn(expr, item):
        reg = re.compile(expr, re.I)
        return reg.search(item) is not None
    pathdocespecial1 = pathpdf
    doc = fitz.open(pathdocespecial1)
    #destepdf = 0
    resultados_para_banco = []
    resultadosx = []
    try:       
        records2 = []
        if(tipobusca=="MATCH"):            
            notok = True
            
            while(notok):
                #sqliteconn = None
                #cursor = None
                try:
                    if(sqliteconnx==None):
                        sqliteconn = connectDB(str(pathdb))
                    else:
                        sqliteconn=sqliteconnx
                    cursor = sqliteconn.cursor()
                    pdfsql = 'Anexo_Eletronico_Conteudo_id_pdf_'+str(idpdf)
                    novabusca =  "SELECT  C.pagina, C.texto, offsets({}) FROM {} C where texto MATCH ? ORDER BY 1".format(pdfsql,pdfsql)
                    cursor.execute(novabusca, (termo.upper(),))                
                    records2 = cursor.fetchall()
                    notok = False
                except sqlite3.OperationalError as ex:
                    utilities_general.printlogexception(ex=ex)
                    time.sleep(2)
                except Exception as ex:
                    utilities_general.printlogexception(ex=ex)
                finally:
                    try:
                        sqliteconn.close()
                    except Exception as ex:
                        None
            #rectspagina = {}
            results = []
            countpagina = 0
            counter = 0
            parar = False
            
            for pages in records2:
                if(listaTERMOS != None and not (termo.upper(),tipobusca) in listaTERMOS):
                    break
                #resultporsecao = 0
                if(parar):
                    inserts = []
                    break                
                offsets = str(pages[2]).split(' ')
                qualcharinit = None
                qualcharfim = None
                contchar = 0
                contagem = 0
                textoembytes = pages[1].encode('utf-8')
                for offset in range(0, len(offsets),4):
                    
                    
                    init = int(offsets[offset+2])
                    fim = int(init+int(offsets[offset+3]))                    
                    slicebytesinit = textoembytes[:init]
                    slicebytesdif =  textoembytes[init:fim]
                    devoltainit = slicebytesinit.decode('utf-8')
                    devoltadif = slicebytesdif.decode('utf-8')
                    ##busca simples
                    toc = None
                    if(tocs_pdf==None):
                        toc = None
                    else:
                        toc = locateToc(pages[0], pathpdf, None, len(devoltainit), tocs_pdf)[0]
                    counter += 1
                    resultsearch = classes_general.ResultSearch()
                    resultsearch.toc = toc
                    resultsearch.idtermopdf = str(idtermopdf)
                    resultsearch.init = len(devoltainit)
                    resultsearch.fim = resultsearch.init + len(devoltadif)
                    resultsearch.pagina = pages[0]
                    resultsearch.pathpdf = pathpdf
                    resultsearch.idpdf = str(idpdf)
                    resultsearch.termo = termo
                    resultsearch.tipobusca = tipobusca
                    resultsearch.idtermo = str(idtermo)
                    resultsearch.prior=int(resultsearch.idtermo)*-1
                    resultsearch.tptoc = 'tp'+str(idtermopdf)+resultsearch.toc
                    snippet = ''.join(char if len(char.encode('utf-8')) <= 3 else '�' for char in pages[1])
                    snippetantes = ""
                    snippetdepois = ""
                    espacos = 0
                    for k in range(len(devoltainit)-1, -1, -1):
                        if(snippet[k]== ' '):
                            espacos+=1
                        snippetantes = snippet[k] + snippetantes
                        if(espacos>=7):
                            break
                    espacos = 0
                    for k in range(len(devoltainit)+len(devoltadif)+1, len(snippet)):
                        if(snippet[k]==' '):
                            espacos+=1
                        snippetdepois += snippet[k] 
                        if(espacos>=7):
                            break    
                    resultsearch.snippet =  (snippetantes, snippet[len(devoltainit):len(devoltainit)+len(devoltadif)], snippetdepois)                    
                    resultsearch.fixo = fixo
                    resultsearch.counter = counter
                    resultados_para_banco.append((resultsearch.idtermo, resultsearch.idpdf, \
                                                 resultsearch.pagina, resultsearch.init, resultsearch.fim, resultsearch.toc, snippetantes, snippetdepois, termo))
                    if(queuesair != None and not queuesair.empty()):
                        x = queuesair.get()    
                        if(x[0]=='pararbusca' and str(x[1])==str(idtermo)):                             
                            parar = True
                            resultadosx = []
                        elif(x[0]=='sairtudo'):
                            if(cursor):
                                cursor.close()              
                            if(sqliteconn):
                                sqliteconn.close()
                            parar = True
                            queuesair.put(x)
                            erros_queue.put(('2', "Parar busca"))
                            return 
                        else:
                            queuesair.put(x)
                    if(not simplesearch):
                        if(parar):
                            resultadosx = []
                            break
                        resultadosx.append(resultsearch)
                        
                    else:
                        results.append((resultsearch))
                    contchar += 1
                countpagina += 1
            
            if(simplesearch):
                return results
            else:
                #result_queue.put((1, resultadosx))
                return [resultados_para_banco, resultadosx] 
        elif(tipobusca=="LIKE"):  
            notok = True
            records2 = []
            while(notok):
                sqliteconn = None
                cursor = None
                try:
                    if(not simplesearch):
                        sqliteconn = connectDB(str(pathdb))
                        cursor = sqliteconn.cursor()
                        cursor.custom_execute("PRAGMA journal_mode=WAL")
                        novabusca =  'SELECT  C.pagina, C.texto FROM Anexo_Eletronico_Conteudo_id_pdf_'+str(idpdf)+''' C where texto like :termo ESCAPE :escape ORDER BY 1'''
                        
                        cursor.custom_execute(novabusca, {'termo':'%'+termo+'%', 'escape': '\\'})
                        records2 = cursor.fetchall()
                    else:
                        records2 = jarecords
                    
                    notok = False
                except sqlite3.OperationalError as ex:
                    utilities_general.printlogexception(ex=ex)
                    time.sleep(2)
                except Exception as ex:
                    utilities_general.printlogexception(ex=ex)
                finally:
                    
                    try:
                        sqliteconn.close()
                    except Exception as ex:
                        None
                    
            results = []
            countpagina = 0
            counter = 0
            inserts = []
            parar = False  
            resultporsecao = {}
            for pagina in records2:
                if(listaTERMOS != None and not (termo.upper(),tipobusca) in listaTERMOS):
                    break
                if(parar):
                    break
                jaachados = set()
                qualcharinit = None
                qualcharfim = None
                init = 0
                resultfind = pagina[1].find(termo, init, len(pagina[1]))
                
                while resultfind!=-1:
                    qualcharinit = resultfind
                    qualcharfim = qualcharinit + len(termo)
                    
                    ##busca simples
                    toc = None
                    if(tocs_pdf==None):
                        toc = None
                    else:
                        toc = locateToc(pagina[0], pathpdf, None, resultfind, tocs_pdf)[0]
                    if(toc not in resultporsecao):
                        resultporsecao[toc]=0
                    if(resultporsecao[toc]>=3000):
                        break  
                    resultporsecao[toc]+=1
                    resultsearch = classes_general.ResultSearch()
                    if(str(qualcharinit)+'-'+str(qualcharfim) in jaachados):
                        init = resultfind+len(termo)
                        resultfind = pagina[1].find(termo, init, len(pagina[1]))
                    else:
                        jaachados.add(str(qualcharinit)+'-'+str(qualcharfim))
                        counter += 1
                                                
                        resultsearch.init = qualcharinit
                        resultsearch.fim = qualcharfim
                        resultsearch.pagina = pagina[0]

                        pathpdf = get_normalized_path(pathpdf)
                        resultsearch.pathpdf = pathpdf
                        resultsearch.idpdf = str(idpdf)
                        resultsearch.termo = termo
                        resultsearch.tipobusca = tipobusca
                        
                        
                        if(not simplesearch):
                            snippetantes = ""
                            snippetdepois = ""
                            espacos = 0
                            for k in range(resultfind-1, -1, -1):
                                char = pagina[1][k]
                                if(char== ' '):
                                    espacos+=1                            
                                if(len(char.encode('utf-8')) <= 3):
                                    snippetantes = char + snippetantes
                                else:
                                    snippetantes = '�' + snippetantes
                                    #snippetantes = char + snippetantes
                                if(espacos>=4):
                                    break
                            espacos = 0
                            for k in range(resultfind+(len(termo)), len(pagina[1])):
                                char = pagina[1][k]
                                if(char== ' '):
                                    espacos+=1 
                                if(len(char.encode('utf-8')) <= 3):
                                    snippetdepois += char 
                                else:
                                    snippetdepois += '�'
                                    #snippetdepois += char
                                #snippetdepois += snippet[k] 
                                if(espacos>=4):
                                    break    
                            #snippetantes = ''.join(char if len(char.encode('utf-8')) < 3 else '�' for char in snippetantes)
                            termo2 = ''.join(char if len(char.encode('utf-8')) <= 3 else '�' for char in termo)
                            resultsearch.idtermopdf = idtermopdf
                            resultsearch.idtermo = idtermo
                            resultsearch.prior=int(resultsearch.idtermo)*-1
                            resultsearch.fixo = fixo
                            resultsearch.counter = counter
                            resultsearch.toc = toc
                            resultsearch.tptoc = 'tp'+str(idtermopdf)+resultsearch.toc
                            resultsearch.snippet =  (snippetantes, termo2, snippetdepois)
                        
                            resultados_para_banco.append((resultsearch.idtermo, resultsearch.idpdf, \
                                                     resultsearch.pagina, resultsearch.init, resultsearch.fim, resultsearch.toc, snippetantes, snippetdepois, termo))
                        else:
                            resultsearch.idtermo = -math.inf
                            resultsearch.idtermopdf = -math.inf
                            resultsearch.prior=-math.inf
                            
                        init = resultfind+len(termo)-1
                        resultfind = pagina[1].find(termo, init)
                        if(queuesair != None and not queuesair.empty()):
                            x = queuesair.get()    
                            if(x[0]=='pararbusca' and str(x[1])==str(idtermo)): 
                                parar = True
                                resultadosx = []
                            elif(x[0]=='sairtudo'):                            
                                parar = True
                                queuesair.put(x)
                                return False
                            else:
                                queuesair.put(x)
                        if(not simplesearch):
                            if(parar):
                                resultadosx = []
                                break
                            resultadosx.append(resultsearch)
                        else:
                            results.append((resultsearch))
                countpagina += 1
            #for resu in resultadosx:
                  
            if(simplesearch):
                return results  
            else:
                #result_queue.put((1, resultadosx))
                return [resultados_para_banco, resultadosx] 
        elif(tipobusca=="REGEX"): 
            
            notok = True
            while(notok):
                sqliteconn = None
                cursor = None
                try:
                    if(not simplesearch):
                        sqliteconn = connectDB(str(pathdb), check_same_thread_arg=False)
                        sqliteconn.create_function("REGEXP", 2, re_fn)
                        cursor = sqliteconn.cursor()
                        cursor.custom_execute("PRAGMA journal_mode=WAL")
                        novabusca =  'SELECT  C.pagina, C.texto FROM Anexo_Eletronico_Conteudo_id_pdf_'+str(idpdf)+''' C where C.texto REGEXP :regex ORDER BY 1'''
                        cursor.custom_execute(novabusca, {'regex':termo}, timeout=30)
                        records2 = cursor.fetchall()
                    
                    notok = False
                except classes_general.TimeLimitExecuteException as ex:
                    sqliteconn.interrupt()
                    #except:
                    #    None
                    raise# classes_general.TimeLimitExecuteException()
                except sqlite3.OperationalError as ex:
                    erros_queue.put(('2', traceback.format_exc()))
                    #utilities_general.printlogexception(ex=ex)
                    time.sleep(2)
                except Exception as ex:
                    erros_queue.put(('2', traceback.format_exc()))
                #    utilities_general.printlogexception(ex=ex)
                finally:                    
                    try:                     
                        sqliteconn.close()
                    except Exception as ex:
                        None
            results = []
            countpagina = 0
            counter = 0
            inserts = []
            parar = False  
            resultporsecao = {}
            for pagina in records2:
                if(listaTERMOS != None and not (termo.upper(),tipobusca) in listaTERMOS):
                    break
                matches = [x.group() for x in re.finditer(termo, pagina[1])]
                jamatched = set()
                for match in matches:
                    if(listaTERMOS != None and not (termo.upper(),tipobusca) in listaTERMOS):
                        break
                    if(match in jamatched):
                        continue
                    jamatched.add(match)
                    qualcharinit = None
                    qualcharfim = None
                    jaachados = set()
                    init = 0
                    resultfind = pagina[1].find(match, init, len(pagina[1]))
                    #
                    while resultfind!=-1:
                       
                        ##busca simples
                        qualcharinit = resultfind
                        qualcharfim = qualcharinit + len(match)
                        
                        if(tocs_pdf==None):
                            toc = None
                        else:
                            toc = locateToc(pagina[0], pathpdf, None, resultfind, tocs_pdf)[0]
                        if(toc not in resultporsecao):
                            resultporsecao[toc]=0
                        if(resultporsecao[toc]>=3000):
                            break  
                        resultporsecao[toc]+=1
                        if(str(qualcharinit)+'-'+str(qualcharfim) in jaachados):
                            init = resultfind+len(match)
                            resultfind = pagina[1].find(match, init, len(pagina[1]))
                        else:                            
                            jaachados.add(str(qualcharinit)+'-'+str(qualcharfim))
                            counter += 1
                            #qualcharinit = resultfind
                            #qualcharfim = qualcharinit + len(match)
                            resultsearch = classes_general.ResultSearch()
                                                    
                            resultsearch.init = qualcharinit
                            resultsearch.fim = qualcharinit + len(match)
                            #erros_queue.put(('2', termo, match, pagina[0], qualcharinit, qualcharfim))
                            resultsearch.pagina = pagina[0]
    
                            pathpdf = get_normalized_path(pathpdf)
                            resultsearch.pathpdf = pathpdf
                            resultsearch.idpdf = str(idpdf)
                            resultsearch.termo = termo
                            resultsearch.tipobusca = tipobusca
                            
                            
                            #snippetantes = ''.join(char if len(char.encode('utf-8')) < 3 else '�' for char in snippetantes[1])
                            if(not simplesearch):
                                snippetantes = ""
                                snippetdepois = ""
                                espacos = 0
                                for k in range(resultfind-1, -1, -1):
                                    char = pagina[1][k]
                                    if(char== ' '):
                                        espacos+=1                            
                                    if(len(char.encode('utf-8')) <= 3):
                                        snippetantes = char + snippetantes
                                    else:
                                        snippetantes = '�' + snippetantes
                                    if(espacos>=4):
                                        break
                                espacos = 0
                                for k in range(resultfind+(len(match)), len(pagina[1])):
                                    char = pagina[1][k]
                                    if(char== ' '):
                                        espacos+=1 
                                    if(len(char.encode('utf-8')) <= 3):
                                        snippetdepois += char 
                                    else:
                                        snippetdepois += '�'
                                    #snippetdepois += snippet[k] 
                                    if(espacos>=4):
                                        break    
                                #snippetantes = ''.join(char if len(char.encode('utf-8')) < 3 else '�' for char in snippetantes)
                                match2 = ''.join(char if len(char.encode('utf-8')) <= 3 else '�' for char in match)
                                resultsearch.idtermopdf = idtermopdf
                                resultsearch.idtermo = idtermo
                                resultsearch.prior=int(resultsearch.idtermo)*-1
                                resultsearch.fixo = fixo
                                resultsearch.counter = counter
                                resultsearch.toc = toc
                                resultsearch.tptoc = 'tp'+str(idtermopdf)+resultsearch.toc
                                resultsearch.snippet =  (snippetantes, match2, snippetdepois)
                                #resultsearch.snippet =  ("", match, "")
                            
                                resultados_para_banco.append((resultsearch.idtermo, resultsearch.idpdf, \
                                                         resultsearch.pagina, resultsearch.init, resultsearch.fim, resultsearch.toc, snippetantes, snippetdepois, match))
                            else:
                                resultsearch.idtermo = -math.inf
                                resultsearch.idtermopdf = -math.inf
                                resultsearch.prior=-math.inf
                                
                            init = resultfind+len(match)-1
                            resultfind = pagina[1].find(match, init, len(pagina[1]))
                            if(not simplesearch):
                                if(parar):
                                    resultadosx = []
                                    break
                                resultadosx.append(resultsearch)
                            else:
                                results.append((resultsearch))
                        if(queuesair != None and not queuesair.empty()):
                            x = queuesair.get()    
                            if(x[0]=='pararbusca' and str(x[1])==str(idtermo)): 
                                parar = True
                                resultadosx = []
                            elif(x[0]=='sairtudo'):                            
                                parar = True
                                queuesair.put(x)
                                return False
                            else:
                                queuesair.put(x)
                        
                countpagina += 1
            #for resu in resultadosx:
            #    result_queue.put(resu)  
            #result_queue.put((1, resultadosx)) 
            return [resultados_para_banco, resultadosx]  
    except classes_general.TimeLimitExecuteException as ex:
        raise
            
    except sqlite3.Error as ex:
        raise
        #erros_queue.put(('3', traceback.format_exc()))
        #utilities_general.printlogexception(ex=ex)
                    
    except Exception as ex:
        raise
        #erros_queue.put(('3', traceback.format_exc()))
        #utilities_general.printlogexception(ex=ex)
    
    finally: 
        try:
            doc.close()
        except:
            None
