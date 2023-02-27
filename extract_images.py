# -*- coding: utf-8 -*-
"""
Created on Thu Jun  2 08:57:06 2022

@author: labinfo
"""

import fitz, traceback, os, sys, subprocess, hashlib, shutil
from pathlib import Path

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

def get_application_path():
    application_path = None
    if getattr(sys, 'frozen', False):
        application_path = sys._MEIPASS
    elif __file__:
        application_path = os.path.dirname(os.path.abspath(__file__))
    return application_path

def extract_images(hash_doc):
    pdf = r"D:\Laudos\Designados\2919-22\REP2919-2022-Anexo\Equipamentos\EQ01\Anexo_2919-2022_EQ01_6.pdf"
    #pdf = r"B:\VISUALIZADOR\FERA\FERA.pdf"
    diretorio = Path(pdf).parent
    diretorio_temp_input = os.path.join(diretorio, "image_match-input-{}".format(hash_doc))
    diretorio_temp_output = os.path.join(diretorio, "image_match-output-{}".format(hash_doc))
    image_match_dir = os.path.join(get_application_path(), 'image_match')
    image_match_exe_abs = os.path.join(image_match_dir, 'image_match.exe')
    try:
        os.makedirs(diretorio_temp_input)
    except:
        None
    hashes = {}
    doc = fitz.open(pdf)
    try:
        count = 0
        
        for p in range(len(doc)):
            page = doc[p]
            d = page.get_text("dict")
            blocks = d["blocks"]  # the list of block dictionaries
            for b in blocks:
                if b["type"] == 1:
                    #d = doc.extract_image(1373)
                    hash_object = hashlib.md5(b['image'])
                    image_hash = hash_object.hexdigest()
                    try:
                        f = open(os.path.join(diretorio_temp_input, "{}.png".format(image_hash)), 'wb')
                        f.write(b['image'])
                        f.close()
                        #pix.save(os.path.join(diretorio_temp_input, "{}.png".format(image_hash)))
                    except:
                        None
            print("{} / {}".format(p+1, len(doc)))
        cmd = r"\"{}\" calc_descriptor33 \"{}\" \"{}\"".format(image_match_exe_abs, diretorio_temp_input, diretorio_temp_output)
        #pritn(cmd)
        popen = subprocess.Popen(cmd, universal_newlines=True, shell=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
        print(cmd)
        while True:
            try:
                line = popen.stdout.readline()
                print("Line: ", line)
                if not line:
                    #None
                    break
           
                try:                
                    None
                    print(line)
                except:
                    traceback.print_exc()
            except UnicodeDecodeError:
                None
        return_code = popen.wait()
        print(return_code)
        try:
            shutil.rmtree(diretorio_temp_input)
        except:
            traceback.print_exc()
        
    except:
        traceback.print_exc()
    finally:
        doc.close()
        
extract_images("fffffff_hash")