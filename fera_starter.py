
import os
import subprocess, platform, traceback, sys
#codereview
ok = 1
try:
    WINDOWS = "Windows" in platform.system()
    LINUX = "Linux" in platform.system()
    #MAC = (platform.system() == "Darwin")
    if(WINDOWS):
        exe_path = os.path.join(os.getcwd(), 'FERA', 'FERA_WINDOWS', 'fera.exe')
    elif(LINUX):
        exe_path = os.path.join(os.getcwd(), 'FERA', 'FERA_LINUX', 'fera')
    
    pathdb = None
    if(os.path.exists(os.path.join(os.getcwd(), 'FERA'))):
        for filename in os.listdir(os.path.join(os.getcwd(), 'FERA')):
            if(os.path.isdir(filename)):
                continue
            filenamenoext, extension = os.path.splitext(filename)
            if ("fera" in filenamenoext.lower() and ".db" == extension.lower()):
                pathdb = os.path.join(os.getcwd(), 'FERA', filename)
                break
    #exe_path = r"D:\TESTEVALIDATION\95962-24-Anexo\FERA\FERA_WINDOWS\fera.exe"
    #pathdb = r"D:\TESTEVALIDATION\95962-24-Anexo\FERA\fera-95962-24.db"
    #sys.argv.append("--test")
    if(pathdb!=None):
        print(exe_path, pathdb, 'with timeout 1200s')
        try:
            if("--test" in sys.argv[1:]):
                print("Testing 1", [exe_path, pathdb, '1', '--test1'])
                popen = subprocess.Popen([exe_path, pathdb, '1', '--test1'])
                ok = popen.wait(timeout=1200)
                if(ok!=0):
                    raise Exception(f"FERA Fail on Test 1: {ok}")
                print("Testing 2", [exe_path, pathdb, '1', '--test2'])
                popen = subprocess.Popen([exe_path, pathdb, '1', '--test2'])
                ok = popen.wait(timeout=1200)
                if(ok!=0):
                    raise Exception(f"FERA Fail on Test 2: {ok}")
            else:
                subprocess.Popen([exe_path, pathdb, '1'])
                ok = 0
        except subprocess.TimeoutExpired:
            print("Process timed out. Killing it...")
            popen.kill()
            raise
except:
    traceback.print_exc()
    ok = 61
finally:
    sys.exit(ok)
