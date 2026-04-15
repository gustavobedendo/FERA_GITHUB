# -*- coding: utf-8 -*-
"""
Created on Tue Feb  1 13:52:40 2022

@author: labinfo
"""
#codereview
import traceback
try:
    import webview
except:
    traceback.print_exc()
import global_settings, utilities_general, classes_general, process_functions
import sys, os

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
        
def get_text_from_obs(observations):
    pathdocespecial = None
    docespecial = None
    fitzpage = None
    regex = r"Hash\s+MD5:\s+([0-9a-zA-Z]{32})"
    hashes = []
    for observation in observations: 
        if(pathdocespecial!=observation.pathpdf):
            pathdocespecial = utilities_general.get_normalized_path(observation.pathpdf)
            pagatual = None
            if(docespecial!=None):
                docespecial.close()
            docespecial = fitz.open(pathdocespecial)
        margemsup = (global_settings.infoLaudo[pathdocespecial].mt/25.4)*72
        margeminf = global_settings.infoLaudo[pathdocespecial].pixorgh-((global_settings.infoLaudo[pathdocespecial].mb/25.4)*72)
        margemesq = (global_settings.infoLaudo[pathdocespecial].me/25.4)*72
        margemdir = global_settings.infoLaudo[pathdocespecial].pixorgw-((global_settings.infoLaudo[pathdocespecial].md/25.4)*72)
        for pagina in range(observation.paginainit, observation.paginafim+1):
            x0 = None
            x1 = margemdir
            y0 = None
            y1 = None
            if(pagina>observation.paginainit and pagina < observation.paginainit):
                x0 = margemesq
                #x1 = margemdir
                y0 = margemsup
                y1 = margeminf
            else:
                x0 = observation.p0x
                y0 = observation.p0y+2
                #x1 = observation.p1x
                y1 = observation.p1y-2 
            rect = fitz.Rect(x0, y0, x1, y1)
            if(pagatual!=pagina):
                pagatual = pagina
                fitzpage = docespecial[pagatual]
            texto = fitzpage.get_textbox(rect)
            results = re.findall(regex, texto)
            for result in results:
                if(result not in hashes):
                    hashes.append(result)
    return hashes
            

def popup_window(texto, sair):
    
    window = tkinter.Toplevel()
    label = tkinter.Label(window, font=global_settings.Font_tuple_Arial_10, text=texto, image=global_settings.warningimage, compound='top')
    label.pack(fill='x', padx=50, pady=20)
    button_close = tkinter.Button(window, font=global_settings.Font_tuple_Arial_10, text="OK", command= lambda : popupcomandook(sair, window))
    button_close.pack(fill='y', pady=20) 
    return window

def show_locations(url, titulo, pid):
    #local_root = tkinter.Tk()
    try:
        webview.create_window(f"P{pid} - {titulo}", url, width=1024, height=768, text_select=True)
        webview.start()
        print("webview started")
    except Exception as ex:
        utilities_general.printlogexception(ex=ex)
    #local_root.mainloop()

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
                    actualdbversion = global_settings.dbversion
                    updateinto2 = "UPDATE FERA_CONFIG set param = ? WHERE config = ?"
                    cursor.custom_execute(updateinto2, (actualdbversion,'dbversion',))
                    return True
        return True
    except Exception as ex:
        global_settings.allok = 774
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
       cursor.custom_execute('ALTER TABLE Anexo_Eletronico_SearchTerms ADD COLUMN pesquisado TEXT', None, False, False)
       commit = True
    except Exception as ex:
        utilities_general.printlogexception(ex=ex)
    try:
       cursor.custom_execute("ALTER TABLE Anexo_Eletronico_SearchTerms ADD COLUMN tipo TEXT NOT NULL default 'relatorio'", None, False, False)
       commit = True
    except Exception as ex:
        utilities_general.printlogexception(ex=ex)
    resulttable = '''SELECT name FROM sqlite_master WHERE type="table" AND name="Anexo_Eletronico_SearchResults"'''
    cursor.custom_execute(resulttable)
    tableresultcount = cursor.fetchone()
    if(tableresultcount==None):
        create_table_searchesresults = '''CREATE TABLE Anexo_Eletronico_SearchResults (
        id_termo INTEGER NOT NULL,
        id_pdf INTEGER NOT NULL,
        pagina INTEGER NOT NULL,
        init INTEGER,
        fim INTEGER,
        x0 INTEGER,
        y0 INTEGER,
        x1 INTEGER,
        y1 INTEGER,
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
    """ try:
        addcolumn = "ALTER TABLE Anexo_Eletronico_SearchResults ADD COLUMN type TEXT NOT NULL default 'relatorio'"
        cursor.custom_execute(addcolumn)
        commit = True  
    except:
        None """
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

def get_eq_base(arquivo):
    arquivo_pasta = os.path.dirname(arquivo)
    for k in range(3):
        if("EQ" in os.path.basename(arquivo_pasta).upper()):
            return os.path.basename(arquivo_pasta)
        else:
            arquivo_pasta = os.path.dirname(arquivo_pasta)
    return "Documentos"

def update_db_version(sqliteconn, cursor):
    updateinto2 = "UPDATE FERA_CONFIG set param = ? WHERE config = ?"
    cursor.custom_execute(updateinto2, (global_settings.dbversion,'dbversion',))
    sqliteconn.commit()
    

    
def ensure_table_exists(sqliteconn, abs_path_pdf, idpdf):
    table_schema = f"""
        CREATE TABLE Anexo_Eletronico_Pdf_Hashes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            id_pdf INTEGER,
            hash TEXT NOT NULL,
            page INTEGER NOT NULL,
            x0 INTEGER NOT NULL,
            y0 INTEGER NOT NULL,
            x1 INTEGER NOT NULL,
            y1 INTEGER NOT NULL,
            original_path TEXT NOT NULL,
            saved_as TEXT NOT NULL,
            parent_alias TEXT NOT NULL, 
            CONSTRAINT fk_idpdf
                FOREIGN KEY (id_pdf)
                    REFERENCES Anexo_Eletronico_Pdfs (id_pdf)
                    ON DELETE CASCADE
        );
        """
    """
    Ensures the specified table exists. If it does not, creates it.

    :param db_name: The SQLite database file name.
    :param table_name: The name of the table to check/create.
    :param table_schema: The SQL schema to create the table if it doesn't exist.
    """
    connection = sqliteconn
    cursor = connection.cursor()
    
    # Check if the table exists
    cursor.execute(f"""
        SELECT name FROM sqlite_master WHERE type='table' AND name='Anexo_Eletronico_Pdf_Hashes';
    """)
    table_exists = cursor.fetchone()
    
    if not table_exists:
        # Create the table if it does not exist
        cursor.execute(table_schema)
        print(f"Table 'Anexo_Eletronico_Pdf_Hashes' created.")
        filelist = os.path.join(os.path.dirname(abs_path_pdf), 'sources', 'fileList.txt')
        global_settings.texto_splash = f"Processando Filelist{os.path.basename(os.path.dirname(abs_path_pdf))}"
        if(filelist not in global_settings.processed_filelist):
            global_settings.processed_filelist[filelist] = global_settings.manager.list([None]*2)
            global_settings.processed_filelist[filelist] = extract_hash_and_saved_as(filelist,  
                                                                                     f"Processando Filelist{os.path.basename(os.path.dirname(abs_path_pdf))}",
                                                                                     idpdf)
        #connection.commit()
        return False
    else:
        print(f"Table 'Anexo_Eletronico_Pdf_Hashes' already exists.")
        return True

    
    
def extract_hash_and_saved_as(file_path, texto, idpdf):
    """
    Extracts the hash and saved_as information from a file formatted as the example provided.

    :param file_path: Path to the file to be processed.
    :return: List of dictionaries with 'hash' and 'saved_as' keys.
    """
    results = {}
    results2 = {}
    pattern = re.compile(r"^([\da-f]{32})\s+(.+?)\[saved as (.+?)\]$")
    inittime = time.time()
    inittime2 = time.time()
    cont = 0
    with open(file_path, "r", encoding="utf-8") as file:
        linhas = file.readlines()
        for line in linhas:
            cont += 1
            if(time.time()-inittime>5):
                global_settings.processados.put(('indexando links - hashes', idpdf, f"FileList {cont}/{len(linhas)}"))
                global_settings.texto_splash = f"{texto} {cont}/{len(linhas)}"
                inittime = time.time()
            if(time.time()-inittime2>20):
                print(f"{texto} {cont}/{len(linhas)}")
                inittime2 = time.time()
            match = pattern.match(line.strip())
            if match:
                hash_value, original_path, saved_as = match.groups()
                results[saved_as] = (hash_value.upper(), original_path)
                results2[hash_value] = (saved_as, original_path)

    return results, results2
    
def extract_links_from_pdf(pdf_path, idpdf, filelist):
    """
    Extract links from all pages of a PDF and return their 'from' and 'file' keys if they exist.

    :param pdf_path: Path to the PDF file.
    :return: List of dictionaries containing extracted link details.
    """
    links2 = []
    inittime = time.time()
    inittime2 = time.time()
    with fitz.open(pdf_path) as doc:
        for page_num, page in enumerate(doc):
            if(time.time()-inittime>2):
                inittime = time.time()
                global_settings.processados.put(('indexando links - hashes', idpdf, f"{os.path.basename(pdf_path)} {page_num}/{len(doc)}"))
                global_settings.texto_splash = f"Processando Links {os.path.basename(pdf_path)} {page_num}/{len(doc)}"
            if(time.time()-inittime2>15):
                inittime2 = time.time()
                print(f"Processando Links {os.path.basename(pdf_path)} {page_num}/{len(doc)}")
            # Extract links from the current page
            for link in page.get_links():
                if('file' not in link): 
                    continue
                file = link['file']
                if(file not in global_settings.processed_filelist[filelist][0]):
                    continue
                if(file[0:5]!= "files"):
                    continue
                hash = global_settings.processed_filelist[filelist][0][file][0]
                original_path = global_settings.processed_filelist[filelist][0][file][1]
                saved_as = global_settings.processed_filelist[filelist][1][hash][0]
                x0, y0, x1, y1 = link.get("from", (0, 0, 0, 0))
                """ if(hash not in global_settings.hashes_to_position):
                    global_settings.hashes_to_position[hash] = {}
                if(idpdf not in global_settings.hashes_to_position[hash]):
                    global_settings.hashes_to_position[hash][idpdf] = []
                global_settings.hashes_to_position[hash][idpdf].append((idpdf, hash, int(page_num),int(x0),int(y0),int(x1),int(y1))) """
                links2.append((idpdf, hash, int(page_num),int(x0),int(y0),int(x1),int(y1), original_path, saved_as))
    return links2


def add_links_from_pdf(sqliteconn, cursor, abs_path_pdf, parent_alias, idpdf):
    hashes_to_insert = {}
    index_fts = os.path.join(os.path.dirname(abs_path_pdf), 'sources', 'index_fts.db')
    files_folder = os.path.join(os.path.dirname(abs_path_pdf), 'files')
    global_settings.info_index_boolean = False
    if(os.path.exists(files_folder)):
        if(os.path.exists(index_fts)):
            global_settings.info_index_boolean = True
            global_settings.info_index[parent_alias] = ("Sim", index_fts)
            jaexiste = ensure_table_exists(sqliteconn, abs_path_pdf, idpdf)
            if(not jaexiste):
                filelist = os.path.join(os.path.dirname(abs_path_pdf), 'sources', 'fileList.txt')
                hashes_to_insert = extract_links_from_pdf(abs_path_pdf, idpdf, filelist)
                insert_hash = f"INSERT INTO Anexo_Eletronico_Pdf_Hashes (id_pdf, hash, page, x0, y0, x1, y1, original_path, saved_as, parent_alias) VALUES (?,?,?,?,?,?,?,?,?,?)"
                cont = 0
                inittime = time.time()
                if(len(hashes_to_insert)>0):
                        #cont += 1
                        #if(time.time()-inittime>2):
                        #    inittime = time.time()
                        #    global_settings.texto_splash = f"Inserindo hashes {cont}{len(hashes_to_insert)}"
                    cursor.custom_executemany(insert_hash, hashes_to_insert)
                sqliteconn.commit()
            else:
                None#get_hashes_from_db()
        else:
            global_settings.info_index[parent_alias] = ("Não","")
    else:
        global_settings.info_index[parent_alias] = ("-","")
    
    
def get_hashes_from_db():
    check_previous_search =  "SELECT id_pdf, hash, page, x0, y0, x1, y1, original_path, saved_as, parent_alias FROM Anexo_Eletronico_Pdf_Hashes"
    sqliteconn = None
    global_settings.splash_window.label['text'] = "Carregando relação Links->Hashes..."
    global_settings.splash_window.label.update_idletasks()
    try:
        sqliteconn = utilities_general.connectDB(str(global_settings.pathdb))
        cursor = sqliteconn.cursor()
        cursor.custom_execute(check_previous_search)    
        records = cursor.fetchall()
        
        for idpdf, hash, pagenum, x0, y0, x1, y1 in records:
            if(hash not in global_settings.hashes_to_position):
                global_settings.hashes_to_position[hash] = {}
            if(idpdf not in global_settings.hashes_to_position[hash]):
                global_settings.hashes_to_position[hash][idpdf] = []
            global_settings.hashes_to_position[hash][idpdf].append((idpdf, hash, pagenum, x0, y0, x1, y1))
    except Exception as ex:
        utilities_general.printlogexception(ex=ex)
    finally:
        if(sqliteconn):
            sqliteconn.close()
    

def gather_information_fromdb(sqliteconn=None):    
    #doc = None  
    if(sqliteconn==None):
        sqliteconn = utilities_general.connectDB(str(global_settings.pathdb))
    #sqliteconn = utilities_general.connectDB(str(global_settings.pathdb))
    cursor = sqliteconn.cursor()
    must_validate = necessity_to_validate(cursor)
    print("Will update db schema", must_validate)
    tocommit = False
    if(must_validate):
        #global_settings.allok = 775
        tocommit = validate_new_db_columns(cursor, must_validate)
        update_db_version(sqliteconn, cursor)
    if(tocommit):
       sqliteconn.commit() 
    totalpaginas = 0
      
    select_all_pdfs = '''SELECT  P.id_pdf, P.rel_path_pdf, P.lastpos, P.tipo, P.margemsup, P.margeminf,
    P.margemesq, P.margemdir, P.hash, P.indexado, P.pixorgw, P.pixorgh, P.doclen, P.parent_alias, P.zoom_pos FROM 
    Anexo_Eletronico_Pdfs P ORDER BY 4,2
    '''
    porcento = 0
    global_settings.texto_splash = f"Reunindo informações ({porcento}%)"
    cursor.custom_execute(select_all_pdfs)
    relats = cursor.fetchall()
    qtos = 0
    verificados = {}            
    cont = 0
    abs_path_pdf = None
    index_fts_set = set()
    
    #no_files_set = set()
    for r in relats: 
        idpdf= r[0]
        
        abs_path_pdf = utilities_general.get_normalized_path(os.path.join(global_settings.pathdb.parent, str(r[1])))
        global_settings.idpdf_to_pathpdf[idpdf] = abs_path_pdf
        global_settings.texto_splash = f"Processando {os.path.basename(abs_path_pdf)}"
        parent_alias = r[13]
        if(parent_alias==None or parent_alias==''):
            parent_alias = utilities_general.get_eq_base(abs_path_pdf)
        
        add_links_from_pdf(sqliteconn, cursor, abs_path_pdf, parent_alias, idpdf)
            
        porcento = round(qtos/len(relats)*100, 0)
        global_settings.texto_splash =  f"Reunindo informações ({porcento}%)"       
        global_settings.infoLaudo[abs_path_pdf] = classes_general.Relatorio()
        filename, file_extension = os.path.splitext(abs_path_pdf)
        if(file_extension.lower()==".pdf"):  
            
            doclen = r[12]
            pixmapw = r[10]
            pixmaph = r[11]
            
            if(r[12]==None):
                
                doc = fitz.open(abs_path_pdf)
                try:
                    doclen = len(doc)
                    pixorg = doc[0].get_pixmap()
                    pixmapw = int(pixorg.width)
                    pixmaph = int(pixorg.height)
                    updateinto2 = "UPDATE Anexo_Eletronico_Pdfs set pixorgw = ?, pixorgh= ?, doclen = ? WHERE id_pdf = ?"
                    cursor.custom_execute(updateinto2, (int(pixorg.width), int(pixorg.height), doclen, r[0],))
                    #sqliteconn.commit()
                except Exception as ex:
                    global_settings.allok = 776
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
                global_settings.allok = 777
            else:
                if(r[8]=='' or r[8]==None):
                    global_settings.infoLaudo[abs_path_pdf].status = 'naoindexado'
                    global_settings.documents_to_index.append(abs_path_pdf)
                    global_settings.allok = 778
                else:
                    
                    hashpdf = str(utilities_general.md5(abs_path_pdf))
                    if(hashpdf.lower()!=r[8].lower()):
                        print(abs_path_pdf)
                        print(hashpdf.lower())
                        global_settings.infoLaudo[abs_path_pdf].status = 'incompativel'
                        global_settings.allok = 779
                    else:
                        global_settings.infoLaudo[abs_path_pdf].status = 'indexado'
                        paginasindexadas = r[12]
            if(parent_alias==None or parent_alias==''):
                parent_alias = utilities_general.get_eq_base(abs_path_pdf)
            relatorio_proxy = classes_general.RelatorioSuccint(r[0], global_settings.infoLaudo[abs_path_pdf].toc, global_settings.infoLaudo[abs_path_pdf].len, \
                                                               pixmapw, pixmaph, r[3], r[4], r[5], r[6], paginasindexadas, \
                                                                   r[1], abs_path_pdf, r[2], parent_alias)   
            global_settings.listaRELS[abs_path_pdf] = relatorio_proxy
            verificados[str(idpdf)] = "OK"              
            cont+=1  
        
    validate_annotation(sqliteconn, cursor)
    global_settings.finished_gathering_info = True
    
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
            global_settings.allok = 780
            utilities_general.printlogexception(ex=ex)
        finally:
            try:
                doc.close()
            except:
                None

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
                    if(link_tratado==""):
                        xref = link['xref']
                        info = global_settings.docatual.xref_get_key(xref, 'A')
                        grupos_search = global_settings.regex_actions_compiled.search(info[1])
                        if(grupos_search==None):
                            continue
                        grupos = grupos_search.groups()
                        link_tratado = grupos[2]
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
    #print(ex)
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
        images = [im for im in images if im.width > 0 and im.height > 0]
        if not images:
            return None

        # Compute a canvas that can hold the widest image
        dst_width  = max(im.width for im in images)
        dst_height = sum(im.height for im in images)

        dst = Image.new('RGB', (dst_width, dst_height))
        y_offset = 0
        for im in images:
            # Optionally, you could center narrower images:
            # x_offset = (dst_width - im.width) // 2
            dst.paste(im, (0, y_offset))
            y_offset += im.height

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
                if(global_settings.limit_search > 0 and th>=global_settings.limit_search):
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
            elif(treeview.tag_has("arqsearch",treenode)):
                valores = treeview.item(treenode, 'values')
                treeview.item(treenode, values=(valores[0], valores[1], th, textoother,))
    return th

def locateToc(pagina, pdf, p0y=None, init=None, tocpdf=None):
    pagina = int(pagina)
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
                        y0 = r.y0 -2
                        x1 = math.ceil(float(bboxchar[-2]))
                        #y1 = math.floor(r.y1)
                        y1 = r.y1 -2
                        c = char['c']
                        if(replace_accent):
                            codePoint = ord(c)
                            if(codePoint<256):
                                codePoint += global_settings.lowerCodeNoDiff[codePoint]
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
    novotexto = novotexto.encode('utf-8', 'surrogatepass').decode('utf-8', 'ignore')
    return (init, novotexto, quadspagina, mapeamento, images_extracted)


def copy_to_clipboard(tipo, conteudo):
    def _executable_exists(name):
        return subprocess.call(["which", name],
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE) == 0
    if(tipo=="rtf"):
        if plt == 'Windows':                        
            CF_RTF = win32clipboard.RegisterClipboardFormat("Rich Text Format")
            win32clipboard.OpenClipboard(0)
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(CF_RTF, conteudo)
            win32clipboard.CloseClipboard()
        elif plt == 'Linux':
            if(_executable_exists("xclip")):
                subprocess.Popen(['xclip', '-selection', 'clipboard', '-t', 'text/rtf'], stdin=subprocess.PIPE).communicate(conteudo)
            else:
                popup_window("Não foi identificada biblioteca compatível para CLIPBOARD - Favor instalar o pacote XCLIP", False)

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
    #print(path_pdf, digest)
    return digest
                
def searchsqlite(tipobusca, termo, pathpdf, pathdb, idpdf, simplesearch = False, queuesair = None, \
                 idtermo = None, idtermopdf = None, erros_queue = None, fixo = None, result_queue = None,\
                     jarecords=None, sqliteconnx=None, tocs_pdf=None, listaTERMOS=None, parent_alias=None, info_index=None, fontebusca='relatorio'):
    def re_fn(expr, item):
        reg = re.compile(expr, re.I)
        return reg.search(item) is not None
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
                    resultsearch = classes_general.ResultSearch(fontebusca)
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
                    if(parent_alias!=None):
                        resultsearch.parent_alias = parent_alias
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
                    if(global_settings.limit_search > 0 and resultporsecao[toc]>=global_settings.limit_search):
                        break  
                    resultporsecao[toc]+=1
                    resultsearch = classes_general.ResultSearch(fontebusca)
                    if(parent_alias!=None):
                        resultsearch.parent_alias = parent_alias
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
                        if(global_settings.limit_search > 0 and resultporsecao[toc]>=global_settings.limit_search):
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
                            resultsearch = classes_general.ResultSearch(fontebusca)
                            if(parent_alias!=None):
                                resultsearch.parent_alias = parent_alias
                                                    
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
        None
        
def searchsqlite_Files(idpdf_to_pathpdf, hashes_to_position, listaTERMOS, tocs_pdf, 
                       tipobusca, pathdb, pathdb_files, termo, idtermo, parent_alias, queuesair, erros_queue):
    def re_fn(expr, item):
        reg = re.compile(expr, re.I)
        return reg.search(item) is not None
    #destepdf = 0
    print(pathdb_files)
    resultados_para_banco = []
    resultadosx = []
    sqliteconn = connectDB(str(pathdb_files), check_same_thread_arg=False)
    try:       
        cursor = sqliteconn.cursor()
        cursor.execute(f"attach '{pathdb_files}' as db1")
        cursor.execute(f"attach '{pathdb}' as db2")
        records2 = []
        novabusca = ""
        if(tipobusca=="MATCH"):  
            novabusca =  f"SELECT  H.hash, trim(replace(snippet(documents, -1, '', '', '...', 200), char(13), ' ')), H.id_pdf, H.page, H.x0, H.y0, H.x1, H.y1 \
                FROM db1.documents C INNER JOIN \
                db2.Anexo_Eletronico_Pdf_Hashes H on (C.hash = H.hash) where C.content MATCH ? ORDER BY 1"          
        if(tipobusca=="LIKE"):  
            novabusca =  f"SELECT  H.hash, trim(replace(snippet(documents, -1, '', '', '...', 200), char(13), ' ')), H.id_pdf, H.page, H.x0, H.y0, H.x1, H.y1 FROM db1.documents C \
                INNER JOIN db2.Anexo_Eletronico_Pdf_Hashes H \
                on (C.hash = H.hash) where C.content like :termo ESCAPE :escape ORDER BY 1"        
        if(tipobusca=="REGEX"):  
            sqliteconn.create_function("REGEXP", 2, re_fn)
            novabusca =  f"SELECT H.hash, trim(replace(snippet(documents, -1, '', '', '...', 200), char(13), ' ')), H.id_pdf, H.page, H.x0, H.y0, H.x1, H.y1 \
                FROM db2.documents INNER JOIN db2.Anexo_Eletronico_Pdf_Hashes H \
                on (C.hash = H.hash) where C.content REGEXP :regex ORDER BY 1"
        notok = True    
        while(notok):
            try:
                
                #cursor = sqliteconn.cursor()
                if(tipobusca=="MATCH"):
                    
                    cursor.custom_execute("PRAGMA journal_mode=WAL")
                    cursor.execute(novabusca, (termo.upper(),))                
                    records2 = cursor.fetchall()
                    notok = False
                elif(tipobusca=="LIKE"):
                    cursor.custom_execute("PRAGMA journal_mode=WAL")
                    cursor.custom_execute(novabusca, {'termo':'%'+termo+'%', 'escape': '\\'})            
                    records2 = cursor.fetchall()
                    notok = False
                elif(tipobusca=="REGEX"):
                    cursor.custom_execute("PRAGMA journal_mode=WAL")
                    cursor.custom_execute(novabusca, {'regex':termo}, timeout=30)
                    records2 = cursor.fetchall()
                    notok = False
            except sqlite3.OperationalError as ex:
                utilities_general.printlogexception(ex=ex)
                time.sleep(2)
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)
                time.sleep(2)
                
            #rectspagina = {}
        results = []
        countpagina = 0
        counter = 0
        parar = False
        
        for hash_hit, snippet, idpdf, pagenum, x0, y0, x1, y1 in records2:
            #if(hash_hit not in hashes_to_position):
            #    continue
            #for idpdf in hashes_to_position[hash_hit]:
            #    for idpdf, hash, pagenum, x0, y0, x1, y1 in hashes_to_position[hash_hit][idpdf]:
            if(listaTERMOS != None and not (termo.upper(),tipobusca) in listaTERMOS):
                break
            #resultporsecao = 0
            if(parar):
                inserts = []
                break                
            idtermopdf = str(idpdf)+'-'+str(idtermo)
            toc = None
            if(tocs_pdf==None):
                toc = None
            else:
                toc = locateToc(int(pagenum), idpdf_to_pathpdf[idpdf], float(y0), None, tocs_pdf[idpdf_to_pathpdf[idpdf]])[0]
            counter += 1
            resultsearch = classes_general.ResultSearch('arquivo')
            resultsearch.toc = toc
            resultsearch.idtermopdf = str(idtermopdf)
            resultsearch.link_position = (x0, y0, x1, y1)
            resultsearch.pagina = int(pagenum)
            resultsearch.pathpdf = idpdf_to_pathpdf[idpdf]
            resultsearch.idpdf = str(idpdf)
            resultsearch.termo = termo
            resultsearch.tipobusca = tipobusca
            resultsearch.idtermo = str(idtermo)
            resultsearch.prior=int(resultsearch.idtermo)*-1
            resultsearch.tptoc = 'tp'+str(idtermopdf)+resultsearch.toc
            resultsearch.parent_alias = parent_alias
            snippet = ''.join(char if len(char.encode('utf-8')) <= 3 else '�' for char in snippet)
            snippetantes = ""
            snippetdepois = ""
            espacos = 0   
            resultsearch.snippet =  ("",snippet ,"")                    
            resultsearch.fixo = 1
            resultsearch.counter = counter
            resultados_para_banco.append((resultsearch.idtermo, resultsearch.idpdf, \
                                            resultsearch.pagina, x0, y0, x1, y1, resultsearch.toc, snippetantes, snippetdepois, termo))
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
            if(parar):
                resultadosx = []
                break
            resultadosx.append(resultsearch)
                    
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
        
        if(sqliteconn):
            sqliteconn.close()
