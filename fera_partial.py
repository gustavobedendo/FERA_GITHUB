# -*- coding: utf-8 -*-
"""
Created on Fri Jun 23 17:07:07 2023

@author: labinfo
"""
import multiprocessing as mp
import sys, getopt, global_settings, setproctitle, fera, utilities_general, global_settings, classes_general
import os,sqlite3, time
from pathlib import Path


def start_up_app():
    sqliteconn = None
    try:
        global_settings.splash_window = classes_general.Splash_window(global_settings.root)
        global_settings.pathdb = Path(sys.argv[1])    
        sqliteconn = utilities_general.connectDB(str(global_settings.pathdb))
        utilities_general.gather_information_fromdb(sqliteconn)
        if(global_settings.pathdb==None):
            return
        global_settings.initiate_processes()
        fera.start_fera_app()
            
                  
    except Exception as ex:
        utilities_general.printlogexception(ex=ex)
    finally:
        try:
            sqliteconn.close()
        except:
            None

if __name__ == '__main__':  
    status = 1
    try:        
        mp.freeze_support()    
        commandline = False
        long_options = ["commandline", "relatorio=", "pathdb="]
        argumentList = sys.argv[1:]
        arguments, values = getopt.getopt(argumentList, [], long_options)
        pathdb = None
        reports = []
        status = 1

        global_settings.initiate_variables()
        setproctitle.setproctitle("FERA "+global_settings.version+" - Forensics Evidence Report Analyzer -- Polícia Científica do Paraná")
        start_up_app()
    except Exception as ex:
        utilities_general.printlogexception(ex=ex)
    finally:
        sys.exit(status)
        
