# -*- coding: utf-8 -*-
"""
Created on Fri Jun 23 17:07:07 2023

@author: labinfo
"""
import multiprocessing as mp
import sys, getopt, global_settings, fera, utilities_general, global_settings, classes_general
import os,sqlite3, time
from pathlib import Path
import threading



def start_up_app(testar1=False, testar2=False):
    sqliteconn = None
    try:
        global_settings.splash_window = classes_general.Splash_window(global_settings.root)
        global_settings.pathdb = Path(sys.argv[1])    
        global_settings.splash_window.window.attributes("-alpha", 255)
        try:
            None
            global_settings.splash_window.window.wm_attributes("-alpha", 255)
        except:
            None 
        #tilities_general.gather_information_fromdb()
        thread = threading.Thread(target=utilities_general.gather_information_fromdb)
        thread.start()
        global_settings.root.after(5, lambda: wait_to_open(testar1, testar2))
        global_settings.root.mainloop()
        
            
                  
    except Exception as ex:
        utilities_general.printlogexception(ex=ex)
        global_settings.allok = 71
    finally:
        try:
            if(sqliteconn):
                sqliteconn.close()
        except:
            None

def wait_to_open(testar1=False, testar2=False):
    if(global_settings.finished_gathering_info):
        if(global_settings.pathdb==None):
            return
        global_settings.initiate_processes()
        fera.start_fera_app(testar1, testar2)
    else:
        #print('waiting')
        global_settings.splash_window.label['text'] = global_settings.texto_splash
        global_settings.root.update_idletasks()
        global_settings.root.after(5, lambda: wait_to_open(testar1, testar2))

if __name__ == '__main__':  
    status = 1
    try:        
        mp.freeze_support()    
        commandline = False
        #sys.argv.append(r"D:\TESTEVALIDATION\95962-24-Anexo\FERA\fera-95962-24.db")
        #sys.argv.append(r"--test2")
        long_options = ["commandline", "relatorio=", "pathdb=", "test1", "test2"]
        argumentList = sys.argv[1:]
        arguments, values = getopt.getopt(argumentList, [], long_options)
        pathdb = None
        reports = []
        status = 0
        testar1 = False
        testar2 = False
        if("--test1" in argumentList):
            testar1 = True
        if("--test2" in argumentList):
            testar2 = True
        #print('test', testar, argumentList)
        global_settings.initiate_variables()
        start_up_app(testar1, testar2)
    except Exception as ex:
        global_settings.allok = 72
        utilities_general.printlogexception(ex=ex)
    finally:
        if(global_settings.allok>0):
            status = global_settings.allok
        print(status)
        sys.exit(status)
        
