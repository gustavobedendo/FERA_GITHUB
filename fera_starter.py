
import os
import subprocess, platform, traceback
try:
    WINDOWS = "Windows" in platform.system()
    LINUX = "Linux" in platform.system()
    #MAC = (platform.system() == "Darwin")
    if(WINDOWS):
        exe_path = os.path.join(os.getcwd(), 'FERA', 'FERA_WINDOWS', 'fera.exe')
    elif(LINUX):
        exe_path = os.path.join(os.getcwd(), 'FERA', 'FERA_LINUX', 'fera')
    
    pathdb = None
    for filename in os.listdir(os.path.join(os.getcwd(), 'FERA')):
        if(os.path.isdir(filename)):
            continue
        filenamenoext, extension = os.path.splitext(filename)
        if ("fera" in filenamenoext.lower() and ".db" == extension.lower()):
            pathdb = os.path.join(os.getcwd(), 'FERA', filename)
            break
    
    if(pathdb!=None):
        print(exe_path, pathdb)
        subprocess.Popen([exe_path, pathdb, '1'])
except:
    traceback.print_exc()
