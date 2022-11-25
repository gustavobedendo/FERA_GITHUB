# -*- coding: utf-8 -*-
"""
Created on Tue Feb  1 13:47:47 2022

@author: labinfo
"""

import time, math, sqlite3, sys, multiprocessing
import utilities_general, classes_general, global_settings, os, re
import fitz, traceback#, setproctitle
from queue import PriorityQueue
from pathlib import Path
import subprocess, shutil

import multiprocessing as mp

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
        
        
def logging_proces2s(erros_queue):   
    while True:
        try:
            erro = None
            if(not erros_queue.empty()):             
                 erro = erros_queue.get(0)
            else:
                time.sleep(0.1)    
            if(erro[0]=='erro'):
                #print(erro)
                #sys.stdout.flush()
                utilities_general.printlogexception(ex=erro[1])
            else:
                None
        except:
            None
        finally:
            None


def indexing_thread_func():
    def dump_queue(queue):
        """
        Empties all pending items in a queue and returns them in a list.
        """
        result = []
    
        for i in iter(queue.get, 'STOP'):
            #print(i[0])
            result.append(i)
        return result
    global_settings.processados.put(('clear_searches', False))
    while(len(global_settings.documents_to_index)>0):
        abs_path_pdf = global_settings.documents_to_index.pop(0)
        if(global_settings.infoLaudo[abs_path_pdf].status=='naoindexado'):
            try:
                mt = global_settings.infoLaudo[abs_path_pdf].mt
                mb = global_settings.infoLaudo[abs_path_pdf].mb
                me = global_settings.infoLaudo[abs_path_pdf].me
                md = global_settings.infoLaudo[abs_path_pdf].md
                idpdf = global_settings.infoLaudo[abs_path_pdf].id
                rel_path_pdf = global_settings.infoLaudo[abs_path_pdf].id
                totalPaginas = 0
                cont = 0
                doc = fitz.open(abs_path_pdf)
                pdf = os.path.basename(abs_path_pdf)
                fim = 0                    
                totalPaginas += len(doc)
                sqliteconn = utilities_general.connectDB(str(global_settings.pathdb))
                print("Indexing thread got connection")
                cursor = sqliteconn.cursor()
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
                sqliteconn.close()
                cont+=1
                for i in range(global_settings.nthreads):
                    init = fim
                    fim = math.ceil((i+1) * (len(doc)/global_settings.nthreads))
                    #__init__(self, idpdf, rel_path_pdf, pdf, init, fim, mt, mb, me, md)
                    proc = Processar(idpdf, abs_path_pdf, pdf, init, min(fim, len(doc)), mt, mb, me, md) 
                    global_settings.processar.put(proc)
                doc.close()
                #insert_content = [global_settings.manager.list()*global_settings.nthreads]
                insert_content= global_settings.manager.list()
                insert_content_images= global_settings.manager.list()
                name_pdf_input = os.path.basename(abs_path_pdf)+"-temp"
                name_pdf_output = os.path.basename(abs_path_pdf)+"-temp_out"
                diretorio_temp_input =os.path.join(Path(abs_path_pdf).parent, name_pdf_input)
                diretorio_temp_output =os.path.join(Path(abs_path_pdf).parent, name_pdf_output)
                try:
                    None
                    #os.makedirs(diretorio_temp_input)
                except:
                    None
                for i in range(global_settings.nthreads):
                    
                    #insertThread(processar, processados, listaRELS, pathdb, inserts, status)
                    global_settings.processing_threads[i] = mp.Process(target=insertThread, name = "FERA INDEXING PROCESS {}".format(i+1),
                                                                       args=(global_settings.processar, global_settings.processados, \
                                                                                                  global_settings.listaRELS, \
                                                                                                  global_settings.pathdb,insert_content, \
                                                                                                      insert_content_images,diretorio_temp_input,), daemon=True)
                    global_settings.processing_threads[i].start()
                cancommit = True
                for i in range(global_settings.nthreads):
                    global_settings.processing_threads[i].join()
                if(global_settings.listaRELS[abs_path_pdf].continuar_a_indexar and len(insert_content)==global_settings.infoLaudo[abs_path_pdf].len):
                    sqliteconn = utilities_general.connectDB(str(global_settings.pathdb))
                    cursor = sqliteconn.cursor()
                    hashpdf = str(utilities_general.md5(abs_path_pdf))
                    cursor.custom_execute("UPDATE Anexo_Eletronico_Pdfs set indexado = 1, hash = ? WHERE id_pdf = ?", (hashpdf, idpdf,))
                    print("Indexing processes finalized - saving to database")
                    sql_insert_content = "INSERT INTO Anexo_Eletronico_Conteudo_id_pdf_" + str(idpdf) +\
                                        " (texto, pagina) VALUES (?,?)"
                    pdfsql = 'Anexo_Eletronico_Conteudo_id_pdf_'+str(idpdf)
                    #sql_insert_content2 = "INSERT INTO {}({}) VALUES (?)".format(pdfsql,pdfsql)
                    #insert_content.put('STOP')
                    #insert_contents=dump_queue(insert_content)
                    
                    sql_insert_content_images = "INSERT INTO Anexo_Eletronico_Images_id_pdf_" + str(idpdf) +\
                                        " (hash_image, bbox_x0, bbox_y0, bbox_x1, bbox_y1, pagina) VALUES (?,?,?,?,?,?)"
                    pdfsql_images = 'Anexo_Eletronico_Images_id_pdf_'+str(idpdf)
                    #sql_insert_content2_images = "INSERT INTO {}({}) VALUES (?)".format(pdfsql_images,pdfsql_images)
                    
                    #return_code = run_image_image_match(diretorio_temp_input, diretorio_temp_output, abs_path_pdf, idpdf)
                    
                    global_settings.processados.put(('saving', abs_path_pdf, idpdf))
                    #cursor.custom_executemany(sql_insert_content_images, insert_content_images)
                    #for i in range(global_settings.nthreads):
                    cursor.custom_executemany(sql_insert_content, insert_content)
                    global_settings.infoLaudo[abs_path_pdf].status = 'indexado'
                    sqliteconn.commit()
                    global_settings.processados.put(('ok', idpdf, abs_path_pdf))
                    print("Indexing processes finalized - saved")
                else:
                    print("Indexing processes finalized - an error has occured")
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)
            finally:
                try:
                    doc.close()
                except:
                    None
                try:
                    sqliteconn.close()
                except:
                    None
    print("Indexing thread going down!")
    global_settings.processados.put(('clear_searches', True))


def run_image_image_match(diretorio_temp_input, diretorio_temp_output, abs_path_pdf, idpdf):
    image_match_dir = os.path.join(utilities_general.get_application_path(), 'image_match')
    image_match_exe_abs = os.path.join(image_match_dir, 'image_match.exe')
   # cmd = r"{} calc_descriptor33 {} {}".format(image_match_exe_abs, diretorio_temp_input, diretorio_temp_output)
    cmd = "\"{}\" calc_descriptor33 \"{}\" \"{}\"".format(image_match_exe_abs, diretorio_temp_input, diretorio_temp_output)
    #print(cmd)
    popen = subprocess.Popen(cmd, universal_newlines=True, shell=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
    #print(cmd)
    regex_percentage = r"File\s[0-9]+\s(\(.*\%\))"
    regex_percentage_compiled = re.compile(regex_percentage)
    while True:
        try:
            line = popen.stdout.readline()
            #print("Line: ", line)
            if not line:
                #None
                break
       
            try:     
                percentages = regex_percentage_compiled.findall(line)
                for percent  in percentages:
                    global_settings.processados.put(('running_im', abs_path_pdf, idpdf, str(percent)))
                #print(line)
            except:
                traceback.print_exc()
        except UnicodeDecodeError:
            None
    return_code = popen.wait()
    #print(return_code)
    try:
        shutil.rmtree(diretorio_temp_input)
    except:
        traceback.print_exc()
    return return_code

def insertThread(processar, processados, listaRELS, pathdb, inserts, inserts_images, diretorio_temp_input):
        
        sqliteconn = None
        #None
        while not processar.empty():
            proc = processar.get()
            ppaginas = 0
            mt = proc.mt
            mb = proc.mb
            me = proc.me
            md = proc.md
            
            abs_path_pdf = proc.rel_path_pdf
            #abs_path_pdf = utilities_general.get_normalized_path(os.path.join(pathdb.parent, proc.rel_path_pdf))
            abs_path_pdf = utilities_general.get_normalized_path(abs_path_pdf)
            
            doc2 = fitz.open(abs_path_pdf)
            pixorg = doc2[0].get_pixmap()
            mmtopxtop = math.floor(mt/25.4*72)
            mmtopxbottom = math.ceil(pixorg.height-(mb/25.4*72))
            mmtopxleft = math.floor(me/25.4*72)
            mmtopxright = math.ceil(pixorg.width-(md/25.4*72))
            try:
                for p in range(proc.paginit, proc.pagfim):
                    if(abs_path_pdf not in listaRELS):
                        raise Exception()
                    if(p==len(doc2)):
                        continue
                    extrair_texto = utilities_general.extract_text_from_page(doc2, p, sys.maxsize, mmtopxtop, mmtopxbottom, mmtopxleft, \
                                                                             mmtopxright, extract_image=False, diretorio_temp_input=diretorio_temp_input)
                    init = extrair_texto[0]
                    novotexto = extrair_texto[1]
                    tupla = (novotexto, p)
                    inserts.append(tupla)
                    for image in extrair_texto[4]:
                        inserts_images.append(image)
                    ppaginas += 1
                    if(ppaginas%100==0):
                        processados.put(('update', abs_path_pdf, 100))
                if(ppaginas%100!=0):
                    processados.put(('update', abs_path_pdf, ppaginas%100))
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)
                processados.put(('erro', ex))
                #listaRELS[abs_path_pdf].continuar_a_indexar = False
                raise Exception()
            finally:
                doc2.close()

def backgroundRendererImage(processed_pages, request_queue, response_queue, queuesair, listaRELS, erros_queue):   
    #setproctitle.setproctitle(multiprocessing.current_process().name)

    docs = {}
    doc = None
    pathatual = None
    lastpos = 0
    qtdeCache = 1
    docs = [None]*qtdeCache    
    while True:
        try:
            pedidoPagina = None

            if(not request_queue.empty()):             
                 pedidoPagina = request_queue.get(0)
            else:
                time.sleep(0.1)
      
            if(pedidoPagina!=None): 
                if(pedidoPagina.qualPdf!=pathatual):
                    pathatual = pedidoPagina.qualPdf
                    pathatual = utilities_general.get_normalized_path(pathatual)
                    doc = None
                    for aberto in docs:
                        if(aberto==None):
                            continue
                        if(aberto[0]==pathatual):
                            doc = aberto[1]
                            break
                    if(doc==None):
                        
                        doc = fitz.open(pathatual)
                        try:
                            docs[lastpos%qtdeCache][1].close()
                        except Exception as ex:
                            None
                        docs[lastpos%qtdeCache] = (pathatual, doc)
                        
                        lastpos+=1
                        if(lastpos==qtdeCache):
                            lastpos=0
                if(not pedidoPagina.qualPagina in processed_pages):
                    continue
                if pedidoPagina.qualPagina >= len(doc):
                    continue
                loadedPage = doc[pedidoPagina.qualPagina]
                if(not pedidoPagina.qualPagina in processed_pages):
                    continue
                pix = loadedPage.get_pixmap(alpha=False, matrix=pedidoPagina.matriz)
                if(pix.width > pix.height):                    
                    pix = loadedPage.get_pixmap(alpha=False, matrix=pedidoPagina.matriz.prescale(pix.height/pix.width, pix.height/pix.width))
                imgdata = pix.tobytes("ppm")
                respostaPagina = classes_general.RespostaDePagina()
                respostaPagina.links = loadedPage.get_links()
                wids = loadedPage.widgets()
                respostaPagina.widgets = []
                if(not pedidoPagina.qualPagina in processed_pages):
                    continue
                for wid in wids:
                    tup = (wid.field_label, wid.rect)
                    respostaPagina.widgets.append(tup)
                respostaPagina.mapeamento = {}
                i = 0
                respostaPagina.qualPagina = pedidoPagina.qualPagina
                respostaPagina.qualGrid = i     
                respostaPagina.imgdata = imgdata
                respostaPagina.qualLabel = pedidoPagina.qualLabel
                respostaPagina.qualPdf = pedidoPagina.qualPdf
                respostaPagina.zoom = pedidoPagina.zoom
                respostaPagina.height = pix.height
                respostaPagina.width = pix.width
                pix = None
                response_queue.put(respostaPagina)               
            else:
                time.sleep(0)
                
        except Exception as ex:
            erros_queue.put(('2', traceback.format_exc()))
        finally:
            for abs_path_pdf in listaRELS.keys():
                try:
                    docs[abs_path_pdf].close()
                except Exception as ex:
                    None
                    
def backgroundRendererXML(request_queuexml, response_queuexml, queuesair, listaRELS, erros_queue, listadeobs): 
    #setproctitle.setproctitle(multiprocessing.current_process().name)
    docs = {}
    doc = None
    pathatual = None
    lastpos = 0
    qtdeCache = 1
    docs = [None]*qtdeCache    
    while True:
        try:
            pedidoPagina = None
            if(not request_queuexml.empty()):             
                 pedidoPagina = request_queuexml.get(0)
            else:
                time.sleep(0.1)
            if(pedidoPagina!=None): 
                if(pedidoPagina.qualPdf!=pathatual):
                    pathatual = pedidoPagina.qualPdf
                    pathatual = utilities_general.get_normalized_path(pathatual)
                    doc = None
                    for aberto in docs:
                        if(aberto==None):
                            continue
                        if(aberto[0]==pathatual):
                            doc = aberto[1]
                            break
                    if(doc==None):
                        doc = fitz.open(pathatual)
                        try:
                            docs[lastpos%qtdeCache][1].close()
                        except Exception as ex:
                            None
                        docs[lastpos%qtdeCache] = (pathatual, doc)
                        
                        lastpos+=1
                        if(lastpos==qtdeCache):
                            lastpos=0
                mt = pedidoPagina.mt
                mb = pedidoPagina.mb
                me = pedidoPagina.me
                md = pedidoPagina.md
                if pedidoPagina.qualPagina >= len(doc):
                    continue
                loadedPage = doc[pedidoPagina.qualPagina]                
                mmtopxtop = math.floor(mt/25.4*72)
                mmtopxbottom = math.ceil(pedidoPagina.pixheight-(mb/25.4*72))
                mmtopxleft = math.floor(me/25.4*72)
                mmtopxright = math.ceil(pedidoPagina.pixwidth-(md/25.4*72))                
                respostaPagina = classes_general.RespostaDePaginaXML()
                respostaPagina.qualPdf = pedidoPagina.qualPdf
                respostaPagina.links = loadedPage.get_links()
                respostaPagina.images = loadedPage.get_images(full=True)
                respostaPagina.qualPagina = pedidoPagina.qualPagina
                wids = loadedPage.widgets()
                respostaPagina.widgets = []
                for wid in wids:
                    tup = (wid.field_label, wid.rect)
                    respostaPagina.widgets.append(tup)
                
                #def extract_text_from_page(doc, pagina, deslocy, topmargin, bottommargin, leftmargin, rightmargin, flags=2+64):
                #return (init, novotexto, quadspagina, mapeamento)
                text_extracted = utilities_general.extract_text_from_page(doc, pedidoPagina.qualPagina, sys.maxsize, mmtopxtop, mmtopxbottom,\
                                                                          mmtopxleft, mmtopxright, flags=2+4+64, replace_accent = False)
                respostaPagina.mapeamento = text_extracted[3]
                respostaPagina.quadspagina = text_extracted[2]

                response_queuexml.put(respostaPagina) 
        except Exception as ex:
            None
            #erros_queue.put(('info', traceback.format_exc()))
        finally:
            for abs_path_pdf in listaRELS.keys():
                try:
                    docs[abs_path_pdf].close()
                except Exception as ex:
                    None    
                    

def find_similar(bytes_from_image):
    image_match_dir = os.path.join(utilities_general.get_application_path(), 'image_match')
    image_match_exe_abs = os.path.join(image_match_dir, 'image_match.exe')
    tempfile = os.path.join(utilities_general.get_application_path(), "temp_im.png")
    #print(tempfile)
    f = open(tempfile, 'wb')
    f.write(bytes_from_image)
    f.close()
    regex = r"Match:\s([^\.]+)"
    match_img = re.compile(regex)
    similar_full = {}
    for relatorio in global_settings.infoLaudo:
        similar_full[relatorio] = []
        similar = []
        relatorio_full = global_settings.infoLaudo[relatorio]
        diretorio_temp_output = relatorio+"-temp_out"
        idpdf = relatorio_full.id
        diretorio_temp_output_exec = diretorio_temp_output+"-exec"
        try:
            None
            #os.makedirs(diretorio_temp_output_exec)
        except:
            None
        try:
            cmd = "\"{}\" brute_search33 \"{}\" \"{}\" 7 {}".format(image_match_exe_abs, diretorio_temp_output, tempfile, diretorio_temp_output_exec)
            #print(cmd)
            popen = subprocess.Popen(cmd, universal_newlines=True, shell=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
            #print(cmd)
            while True:
                try:
                    line = popen.stdout.readline()
                    #print("Line: ", line)
                    if not line:
                        #None
                        break
               
                    try:                
                        None
                        #print(line)
                        matches = match_img.findall(line)
                        for match  in matches:
                            similar.append(match)
                            #global_settings.processados.put(('running_im', abs_path_pdf, idpdf, str(percent)))
                    except:
                        traceback.print_exc()
                except UnicodeDecodeError:
                    None
            return_code = popen.wait()   
        except:
            None 
        #placeholder= '?'
        placeholders= ', '.join(f'\"{unused}\"' for unused in similar)
        get_search_results =  "SELECT id_image, hash_image, pagina, bbox_x0, bbox_y0, bbox_x1, bbox_y1"+\
            " FROM Anexo_Eletronico_Images_id_pdf_{}  where hash_image IN ({}) ORDER by 3".format(idpdf, placeholders)
        #print(get_search_results)
        sqliteconn = utilities_general.connectDB(str(global_settings.pathdb))
        try:
            cursor = sqliteconn.cursor()
            cursor.custom_execute(get_search_results)
            search_results = cursor.fetchall()
            for result in search_results:
                similar_full[relatorio].append(result)
        except:
            None
        finally:
            sqliteconn.close()
            try:
                shutil.rmtree(diretorio_temp_output_exec)
            except:
                traceback.print_exc()
    try:
        None
        #os.remove(tempfile)
    except:
        traceback.print_exc()
    return similar_full
                         
 
def checkSearchQueue(searchqueue, listaTERMOS, result_queue, pedidos, pedidos_to_db, pathdb, erros_queue):
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
    try:
        sqliteconn = None
        commit = False
        cursor = None
        send_back = []
        try:
            while not searchqueue.empty():
                #if(len(searchqueue) ==0):
                #    return
                pedidosearch = searchqueue.get()  
                
                termo = pedidosearch.termo
                termoorg = termo  
                novotermo = ""
                novotermo2 = ""
                tipobusca = pedidosearch.tipo
                for char in termo:
                    codePoint = ord(char)
                    if tipobusca=="MATCH" and (codePoint<256) and (codePoint>=192):
                        codePoint += lowerCodeNoDiff[codePoint]
                    elif(codePoint<256):
                        codePoint += lowerCodeNoDiff[codePoint]
                    novotermo += chr(codePoint) 
                    novotermo2 += chr(codePoint) 
                novotermo = novotermo.strip()
                novotermo2 = novotermo2.strip()
                
                #termo = termo if tipobusca=="REGEX" else novotermo
                if tipobusca=="REGEX":
                    termo = termo
                elif tipobusca=="MATCH":
                    termo = novotermo2
                else:
                    termo = novotermo
                #advancedsearchbool = int(advancedsearch)==1
                
                pesquisados = ""
                if((termo.upper(),tipobusca) in listaTERMOS and len(listaTERMOS[(termo.upper(), tipobusca)])>0):                            
                    idtermo = listaTERMOS[(termo.upper(),tipobusca)][2]
                    pesquisados = listaTERMOS[(termo.upper(),tipobusca)][3]
                else:
                    if(sqliteconn == None):
                        sqliteconn = utilities_general.connectDB(str(pathdb))
                    if(cursor==None):
                        cursor = sqliteconn.cursor()
                        cursor.custom_execute("PRAGMA journal_mode=WAL")
                    try:                        
                        sql_insert_searchterm = "INSERT INTO Anexo_Eletronico_SearchTerms (termo, tipobusca, fixo, pesquisado) VALUES (?,?,?,?)" 
                        pesquisados = ""
                        cursor.custom_execute(sql_insert_searchterm, (termo,tipobusca, 1,"",))
                        idtermo = cursor.lastrowid
                        listaTERMOS[(termo.upper(),tipobusca)] = [termo,tipobusca, idtermo, ""]
                        commit = True
                                    
                        #cursor.close() 
                    except Exception as ex:
                        None
                    
                if((termo.upper(),tipobusca) in listaTERMOS):
                        
                    resultsearch = classes_general.ResultSearch()
                    resultsearch.termo = termo
                    resultsearch.tipobusca = tipobusca
                    resultsearch.idtermo = str(idtermo)
                    resultsearch.fixo = 1
                    resultsearch.prior=int(resultsearch.idtermo)*-1
                    pedido_in_queue = classes_general.PedidoInQueue(idtermo, tipobusca, termo, pesquisados)
                    pedidos.put(pedido_in_queue)    
                    #send_back.append((0,  resultsearch))
                    result_queue.put((0,  resultsearch))  
        except Exception as ex:
            None
        if(commit):
            sqliteconn.commit()
        try:
            sqliteconn.close()
        except Exception as ex:
            None
        #for res in send_back:
        #    result_queue.put(res)  
    except Exception as ex:
        None
        #exc_type, exc_value, exc_tb = sys.exc_info()
        erros_queue.put(('2', traceback.format_exc()))
        #erros_queue.put(('2', ex))  
        
def searchProcess(result_queue, pathdb, erros_queue, queuesair, searchqueue, update_queue, listaRELS, listaTERMOS, estavel=False):  
    #setproctitle.setproctitle(multiprocessing.current_process().name)
    historicoDeParsing = {}
    pedidos = PriorityQueue()
    while(True): 
        if(not searchqueue.empty()):           
            notok = True
            while(notok):
                sqliteconn = None
                cursor = None
                try:                    
                    notok = False
                    
                    pedidos_to_db = []
                    adv = []
                    notadv = []
                    #while len(searchqueue)>0:
                    checkSearchQueue(searchqueue, listaTERMOS, result_queue, pedidos, pedidos_to_db, pathdb, erros_queue)
                    while not pedidos.empty():
                        #while  len(searchqueue)>0:
                        checkSearchQueue(searchqueue, listaTERMOS, result_queue, pedidos, pedidos_to_db, pathdb, erros_queue)                            
                        pedidosearch_off_queue = pedidos.get()                   
                        termo =  pedidosearch_off_queue.termo
                        tipobusca = pedidosearch_off_queue.tipobusca
                        idtermo = pedidosearch_off_queue.idtermo
                        
                        #search_results = []
                        pesquisadoadd = pedidosearch_off_queue.pesquisados
                        if(pesquisadoadd==None):
                            pesquisadoadd=""
                        try:
                            sqliteconn.close()
                        except Exception as ex:
                            None
                        sqliteconn = utilities_general.connectDB(str(pathdb), 5, maxrepeat=4)
                        try:
                            #if(sqliteconn
                            needcommit = False
                            counter = 0
                            tocs_pdf = {}
                            for key_in_listaRELS in listaRELS.keys():
                                tocs_pdf[key_in_listaRELS] = []
                                for toc_unit in listaRELS[key_in_listaRELS].tocpdf:
                                    tocs_pdf[key_in_listaRELS].append(toc_unit)
                                    
                            #erros_queue.put(('2', tocs_pdf[key_in_listaRELS]))
                            result_queue.put((0.5, idtermo))
                            for key_in_listaRELS in listaRELS.keys():
                                if((termo.upper(),tipobusca) not in listaTERMOS):
                                    needcommit = False
                                    break
                                abs_path_pdf = listaRELS[key_in_listaRELS].abs_path_pdf
                                idpdf = listaRELS[key_in_listaRELS].idpdf
                              
                                idtermopdf = str(idpdf)+'-'+str(idtermo)
                                cursor = sqliteconn.cursor()
                                if("({})".format(idpdf) in pesquisadoadd):
                                    get_search_results =  "SELECT id_termo, id_pdf, pagina, init, fim, toc, snippetantes, snippetdepois, termo "+\
                                        "FROM Anexo_Eletronico_SearchResults  where id_termo = ? AND id_pdf = ? ORDER by 1,2,3,4"
                                    cursor.custom_execute(get_search_results, (idtermo, idpdf,))
                                    search_results = cursor.fetchall()
                                    resultadosx = []
                                    for result_res in search_results:
                                        counter += 1
                                        resultsearch = classes_general.ResultSearch()
                                        resultsearch.toc = result_res[5]
                                        resultsearch.idtermopdf = str(idtermopdf)
                                        resultsearch.init = result_res[3]
                                        resultsearch.fim = result_res[4]
                                        resultsearch.pagina = result_res[2]
                                        resultsearch.pathpdf = abs_path_pdf
                                        resultsearch.idpdf = str(idpdf)
                                        resultsearch.termo = termo
                                        resultsearch.tipobusca = tipobusca
                                        resultsearch.idtermo = str(idtermo)
                                        resultsearch.prior=int(resultsearch.idtermo)*-1
                                        snippetantes = result_res[6]
                                        snippetdepois = result_res[7]
                                        resultsearch.snippet =  (snippetantes, result_res[8], snippetdepois)                    
                                        resultsearch.fixo = 1
                                        resultsearch.counter = counter
                                        #result_queue.put((0,  resultsearch)) 
                                        try:                                        
                                            resultsearch.tptoc = 'tp'+str(idtermopdf)+result_res[5]
                                        except:
                                            resultsearch.tptoc = None
                                            #resultsearch.tptoc = 'tp'+str(idtermopdf)+result_res[5]
                                        resultadosx.append(resultsearch)
                                    if((termo.upper(),tipobusca) in listaTERMOS):
                                        result_queue.put((1, resultadosx))
                                    else:
                                        break
                                    
                                else:
                                    retorno_sqlite_search = utilities_general.searchsqlite(tipobusca, termo, abs_path_pdf, pathdb, idpdf, queuesair=queuesair, idtermo=str(idtermo), \
                                                                                    idtermopdf=str(idtermopdf), \
                                                                 erros_queue = erros_queue, fixo = 1, result_queue = result_queue, \
                                                                     sqliteconnx=None, tocs_pdf=tocs_pdf[abs_path_pdf], listaTERMOS=listaTERMOS)
                                    search_results = retorno_sqlite_search[0]
                                    resultadosx = retorno_sqlite_search[1]
                                    
                                    if(len(search_results)>0):
                                        
                                        
                                        sql_insert_searchresukt =\
                        "INSERT INTO Anexo_Eletronico_SearchResults (id_termo, id_pdf, pagina, init, fim, toc, snippetantes, snippetdepois, termo) VALUES (?,?,?,?,?,?,?,?,?)"
                                        cursor.custom_executemany(sql_insert_searchresukt, (search_results))
                                    pesquisadoadd += "({})".format(idpdf)
                                    updateinto2 = "UPDATE Anexo_Eletronico_SearchTerms set pesquisado = ? WHERE id_termo = ?"  
                                    cursor.custom_execute(updateinto2, (pesquisadoadd,idtermo,))
                                    needcommit = True
                                    if((termo.upper(),tipobusca) in listaTERMOS):
                                        result_queue.put((1, resultadosx))
                                    else:
                                        break
                                    #result_queue.put((1, search_results))
                            if((termo.upper(),tipobusca) in listaTERMOS):
                                resultsearch = classes_general.ResultSearch()
                                resultsearch.termo = termo
                                resultsearch.tipobusca = tipobusca
                                resultsearch.idtermo = str(idtermo)
                                resultsearch.prior=int(resultsearch.idtermo)*-1
                                resultsearch.end=True
                                resultsearch.idpdf = str(math.inf)
                                resultsearch.counter = counter + 1
                                try:                                        
                                    resultsearch.tptoc = 'tp'+str(idtermopdf)+str(result_res[5])
                                except:
                                    resultsearch.tptoc = None
                                result_queue.put((2,  resultsearch)) 
                                if(needcommit):
                                    sqliteconn.commit()
                                listaTERMOS[(termo.upper(),tipobusca)] = [termo,tipobusca, idtermo, pesquisadoadd]
                                pesquisadoadd = ""
                        except classes_general.TimeLimitExecuteException as ex:
                            erros_queue.put(('3', termo, tipobusca, idtermo, "Ocorreu um erro inesperado e a operação não foi concluída.\n{}".format("Tempo limite de execução excedido.")) ) 
                            
                        except Exception as ex:
                            erros_queue.put(('3', termo, tipobusca, idtermo, traceback.format_exc()))
                            #erros_queue.put(('2', ex))
                        finally:
                            if(sqliteconn):
                                sqliteconn.close()
                
                except sqlite3.Error as ex:
                    None
                    erros_queue.put(('2', traceback.format_exc()))
                    #exc_type, exc_value, exc_tb = sys.exc_info()
                    #erros_queue.put(('2', str(exc_type) + "\n" + str(exc_value)+ "\n" + str(exc_tb)))
                except sqlite3.OperationalError as ex:
                    None
                    erros_queue.put(('2', traceback.format_exc()))
                    #exc_type, exc_value, exc_tb = sys.exc_info()
                    #erros_queue.put(('2', str(exc_type) + "\n" + str(exc_value)+ "\n" + str(exc_tb)))
                    time.sleep(1)
                except Exception as ex:
                    erros_queue.put(('2', traceback.format_exc()))
                    #utilities_general.printlogexception(ex=ex) 
                finally:
                    try:
                        cursor.close() 
                    except Exception as ex:
                        None
                    try:
                        sqliteconn.close()
                    except Exception as ex:
                        None
        else:
            if(not estavel):
                break
            else:
                time.sleep(1)          

def processBatchInsertObs(listadeitenscompleto, allitens):
        doc = None
        global_settings.pathpdfatual = None
        paginaatual = None
        quadspagina = []
        try:
            for item in allitens:
                resultsearch = item[0]
                mt = item[1]
                mb = item[2]
                me = item[3]
                md =  item[4]
                idobscat = item[7]
                mmtopxtop = math.floor(mt/25.4*72)
                mmtopxbottom = math.ceil(item[6]-(mb/25.4*72))
                mmtopxleft = math.floor(me/25.4*72)
                mmtopxright = math.ceil(item[5]-(md/25.4*72))
                pathpdf = resultsearch.pathpdf
                if(global_settings.pathpdfatual!=pathpdf):
                    if(doc!=None):
                        doc.close()
                    global_settings.pathpdfatual=pathpdf
                    doc = fitz.open(global_settings.pathpdfatual)
                pagina = int(resultsearch.pagina)
                if(pagina !=paginaatual):
                    quadspagina = []
                    paginaatual = pagina
                    loadedPage = doc[paginaatual]
                    
                    #def extract_text_from_page(doc, pagina, deslocy, topmargin, bottommargin, leftmargin, rightmargin, flags=2+64):
                    #return (init, novotexto, quadspagina, mapeamento)
                    text_extracted = utilities_general.extract_text_from_page(doc, paginaatual, sys.maxsize, mmtopxtop, mmtopxbottom, mmtopxleft, mmtopxright, flags=2+64)
                    quadspagina = text_extracted[2]

                posicoes = quadspagina
                init = posicoes[resultsearch.init]
                fim = posicoes[resultsearch.fim-1]
                p0x = round(init[0])
                p0y = round((init[1]+init[3])/2)
                p1x = round(fim[2])
                p1y = round((fim[1]+fim[3])/2)
                listadeitenscompleto.append((p0x, p0y, p1x, p1y, pagina, global_settings.pathpdfatual, idobscat))
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        finally:
            if(doc!=None):
                doc.close()  