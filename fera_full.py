# -*- coding: utf-8 -*-
"""
Created on Fri Jun 23 17:07:07 2023

@author: labinfo
"""
import multiprocessing as mp
import sys, getopt, global_settings, indexador_fera, setproctitle, fera, utilities_general, global_settings, classes_general
import os,sqlite3, time, shutil
from pathlib import Path
import traceback


def start_up_app():
    sqliteconn = None
    try:
        if(len(sys.argv) == 1): 
            indexador_fera.import_create_toplevel()            
            if(global_settings.pathdb == None):
                return  
            else:
                sys.argv.append(str(global_settings.pathdb))
        global_settings.splash_window = classes_general.Splash_window(global_settings.root)
        if(len(sys.argv) >= 2): 
            filename, extension = os.path.splitext(sys.argv[1])
            if(".pdf" == extension.lower()):
                pathp = sys.argv[1]
                global_settings.pathdb = Path(os.path.abspath(sys.argv[1])+".db")  
                sqliteconn = None
                tocommit = False
                if(not os.path.exists(global_settings.pathdb)):
                    
                    #print(1)
                    notindexed = []
                    notindexed.append(pathp)
                    global_settings.splash_window.window.deiconify()
                    global_settings.splash_window.label['text'] = "Criando banco de dados..."
                    global_settings.splash_window.label.update_idletasks()
                    indexador_fera.createNewDbFile(global_settings.root)   
                    sqliteconn = utilities_general.connectDB(str(global_settings.pathdb))
                    global_settings.splash_window.label['text'] = "Aguardando definição de margens..."
                    global_settings.splash_window.label.update_idletasks()
                    tupleinfo = indexador_fera.addrels('relatorio', view=None, pathpdfinput = notindexed, \
                                                       pathdbext=global_settings.pathdb, rootx=global_settings.root, sqliteconnx=sqliteconn)
                    global_settings.splash_window.label['text'] = "Carregando ferramenta..."
                    global_settings.splash_window.label.update_idletasks()
                else:
                    sqliteconn = utilities_general.connectDB(str(global_settings.pathdb))
                    #print(2)
                    cursor = sqliteconn.cursor()
                    if(os.path.exists(str(global_settings.pathdb)+'.lock')):
                        window = utilities_general.popup_window(sair=True, texto = \
                                    "O banco de dados aparentemente está aberto em outra execução!\nO programa irá encerrar para evitar inconsistências.\n"+\
                                     "Para corrigir esse problema:\nVerifique outras execuções utilizando o mesmo banco de dados\n ou \nApague o arquivo <{}>".\
                                         format(str(global_settings.pathdb)+'.lock'))
                        global_settings.root.wait_window(window)                                            
                    indexador_fera.validate_annotation(sqliteconn, cursor)
                    must_validate = indexador_fera.necessity_to_validate(cursor)
                    
                    if(must_validate):
                        tocommit = indexador_fera.validate_new_db_columns(cursor, must_validate)
                        indexador_fera.update_db_version(sqliteconn, cursor)
                    notindexed = []
                    select_all_pdfs = '''SELECT  P.id_pdf, P.indexado, P.rel_path_pdf FROM 
                    Anexo_Eletronico_Pdfs P 
                    '''
                    try:
                        cursor.custom_execute(select_all_pdfs, None, True, False)
                        relats = cursor.fetchall()
                        notindexed.append(pathp)
                        for rel in relats:                                   
                            notindexed.append(os.path.join(global_settings.pathdb.parent, rel[2]))
                        if(len(notindexed)>0): 
                            tupleinfo = indexador_fera.addrels('relatorio', pathpdfinput = notindexed, pathdbext=sys.argv[1], \
                                                                   rootx=global_settings.root, sqliteconnx=sqliteconn)
                    except sqlite3.OperationalError as ex:  
                        notindexed.append(pathp)
                        try:
                            sqliteconn.close()
                        except:
                            None
                        indexador_fera.createNewDbFile(global_settings.root) 
                        sqliteconn = utilities_general.connectDB(str(global_settings.pathdb))                           
                        tupleinfo = indexador_fera.addrels('relatorio', pathpdfinput = notindexed, pathdbext=sys.argv[1], rootx=global_settings.root, sqliteconnx=sqliteconn)                            
                        indexing = True
                #indexador_fera.gather_information_fromdb()
                if(tocommit):
                    sqliteconn.commit()
                utilities_general.initiate_indexing_thread()
                global_settings.initiate_processes()
                start_time = time.time()
                fera.start_fera_app() 
            elif(".db" == extension.lower()):
                global_settings.pathdb = Path(sys.argv[1])    
                gotoviewer = False
                if(len(sys.argv) >= 3 and sys.argv[2]=='1'):
                    gotoviewer = True  
                sqliteconn = utilities_general.connectDB(str(global_settings.pathdb))
                
                utilities_general.gather_information_fromdb(sqliteconn)
                sqliteconn.close()
                utilities_general.initiate_indexing_thread()
                if(not gotoviewer):
                    global_settings.splash_window.window.withdraw()
                    indexador_fera.App(global_settings.version, gotoviewer)
                if(global_settings.pathdb==None):
                    return
                global_settings.initiate_processes()
                fera.start_fera_app()  
        else:
            print(1)
    except Exception as ex:
        utilities_general.printlogexception(ex=ex)
    finally:
        try:
            sqliteconn.close()
        except:
            None

def find_reports_shortcut(path):
    pdfs = []
    for filename in os.listdir(path):
        if(filename.endswith(".pdf")):
            pdfs.append(os.path.join(path, filename))
    return pdfs

if __name__ == '__main__':  
    status = 1
    try:        
        mp.freeze_support()    
        commandline = False
        long_options = ["commandline", "relatorio=", "pathdb=", "shortcut"]
        argumentList = sys.argv[1:]
        arguments, values = getopt.getopt(argumentList, [], long_options)
        pathdb = None
        reports = []
        status = 1
        shortcut = False
        pathdbparent = ""
        print(sys.argv)
        for currentArgument, currentValue in arguments:
            if currentArgument in ("--commandline"):
                commandline = True
            if currentArgument in ("--shortcut"):
                shortcut = True
                commandline = True
                pathdbparent = os.path.join(sys.argv[2], "FERA")
                nomedb = f"fera.db"
                pathdb = os.path.join(pathdbparent, nomedb)
                reports.extend(find_reports_shortcut(sys.argv[2]))
                print(reports)
            if currentArgument in ("--pathdb"):
                pathdb = str(currentValue)
            if currentArgument in ("--relatorio"):
                reports.append(currentValue)
        if(commandline):
            if(reports==None or pathdb==None):
                print(f"Erro pathdb: {pathdb} - reports: {reports}")
                sys.exit(1)
            global_settings.initiate_variables(commandline)
            if(shortcut):
                #rep = input("REP:")
                #eq = int(input("Eq:"))
                
                fera_windows_src = os.path.join(utilities_general.get_application_path(), "FERA_WINDOWS")
                fera_windows_dst = os.path.join(pathdbparent, "FERA_WINDOWS")
                shutil.copytree(fera_windows_src, fera_windows_dst)
                fera_linux_src = os.path.join(utilities_general.get_application_path(), "FERA_LINUX")
                fera_linux_dst = os.path.join(pathdbparent, "FERA_LINUX")
                shutil.copytree(fera_linux_src, fera_linux_dst)
                shutil.copy(os.path.join(utilities_general.get_application_path(), "FERA_CALLERS", "FERA-Windows.exe"), os.path.dirname(sys.argv[2]))
                shutil.copy(os.path.join(utilities_general.get_application_path(), "FERA_CALLERS", "FERA-Linux.sh"), os.path.dirname(sys.argv[2]))
                shutil.copy(os.path.join(utilities_general.get_application_path(), "FERA_CALLERS", "FERA.pdf"), os.path.dirname(sys.argv[2]))
            status = indexador_fera.build_db_with_reports_commandline(pathdb, reports)
            
            
        else:
            global_settings.initiate_variables()
            setproctitle.setproctitle("FERA "+global_settings.version+" - Forensics Evidence Report Analyzer -- Polícia Científica do Paraná")
            start_up_app()
    except Exception as ex:
        if(commandline):
            traceback.print_exc()
        utilities_general.printlogexception(ex=ex)
    finally:
        sys.exit(status)
        
