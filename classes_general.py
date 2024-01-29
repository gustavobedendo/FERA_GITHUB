# -*- coding: utf-8 -*-
"""
Created on Tue Feb  1 13:59:33 2022

@author: labinfo
"""
import utilities_general, classes_general
import global_settings, sqlite3, subprocess
import tkinter, platform, webbrowser
from functools import total_ordering
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename
import shutil, os, math, fitz
from pathlib import Path
from PIL import Image, ImageTk, ImageFile
ImageFile.LOAD_TRUNCATED_IMAGES = True
#import speech_recognition as sr
#from vosk import Model, KaldiRecognizer
import wave
import json
from tkinter import messagebox
#import vosk
from tkinter import ttk
from functools import partial
import threading
plt = platform.system()



class TimeLimitExecuteException(Exception):
    pass

class FFMPEGException(Exception):
    pass

class Custom_Cursor(sqlite3.Cursor):
    def execute_with_timeout(self, command, arguments):
        try:
            if(arguments==None):
                self.execute(command)
            else:
                self.execute(command, arguments)
        except sqlite3.OperationalError as ex:
            raise TimeLimitExecuteException()
    def custom_execute(self, command, arguments=None, raise_ex=True, show_popup=True, timeout=0):
        
        try:
            try:
                quais = ['delete']
                primeiro = command.split(" ")[0].lower()
                if(primeiro in quais):
                    global_settings.is_dirty = True
            except Exception as ex:
                utilities_general.utilities_general.printlogexception(ex=ex)

            #print(command)
            if(timeout!=0):
                execute_thread = threading.Thread(target=self.execute_with_timeout, args=(command, arguments,), daemon=True)
                execute_thread.start()
                execute_thread.join(timeout)
                if(execute_thread.isAlive()):
                    #erros_queue.put(('2', 'traceback.format_exc'))
                    raise TimeLimitExecuteException()
                
                    
            elif(arguments==None):
                self.execute(command)
            else:
                self.execute(command, arguments)
        except TimeLimitExecuteException as ex:
            raise
        except sqlite3.OperationalError as ex:
            if(show_popup):
                utilities_general.popup_window("Ocorreu um erro inesperado e a operação não foi concluída. \nTente novamente em alguns segundos.\n{}".format(ex), False)
            if(raise_ex):
                raise ex
        except sqlite3.DatabaseError as ex:
            if(show_popup):
                utilities_general.popup_window("Ocorreu um erro inesperado e a operação não foi concluída. \nTente novamente em alguns segundos.\n{}".format(ex), False)
            if(raise_ex):
                raise ex
    def custom_executemany(self, command, arguments=None, raise_ex=True, show_popup=True, timeout=0):
        try:
            if(timeout!=0):
                execute_thread = threading.Thread(target=self.execute_with_timeout, args=(command, arguments,), daemon=True)
                execute_thread.start()
                execute_thread.join(timeout)
                if(execute_thread.isAlive()):
                    raise TimeLimitExecuteException()
            elif(arguments==None):
                self.executemany(command)
            else:
                self.executemany(command, arguments)
        except sqlite3.OperationalError as ex:
            if(show_popup):
                utilities_general.popup_window("Ocorreu um erro inesperado e a operação não foi concluída. \nTente novamente em alguns segundos.\n{}".format(ex), False)
            if(raise_ex):
                raise ex
        except sqlite3.DatabaseError as ex:
            if(show_popup):
                utilities_general.popup_window("Ocorreu um erro inesperado e a operação não foi concluída. \nTente novamente em alguns segundos.\n{}".format(ex), False)
            if(raise_ex):
                raise ex
                
        
class Custom_Database(sqlite3.Connection):
    def cursor(self):
        return super(Custom_Database, self).cursor(Custom_Cursor)




class Splash_window():
    def __init__(self, root):
        self.window = tkinter.Toplevel(root)
        self.window.overrideredirect(True)
        self.window.withdraw()
        self.label = None
        root.eval(f'tk::PlaceWindow {str(self.window)} center')
        self.window.rowconfigure((0,1), weight=1)
        self.window.columnconfigure(0, weight=1)
        self.label = tkinter.Label(self.window, bg="#464646", fg="#bc9e20", \
                                                            font=global_settings.Font_tuple_ArialBold_10, image=global_settings.loading, text="Carregando", compound="top", borderwidth=0)
        self.label.image = global_settings.loading
        self.label.grid(row=0, column=0, sticky='nsew')
        self.window.protocol("WM_DELETE_WINDOW", lambda: None)
        self.window.resizable(False, False)
        utilities_general.center(self.window)
        
class Indexing_window():
    def __init__(self, root):
        self.window = tkinter.Toplevel(root)
        self.window.overrideredirect(True)
        self.window.withdraw()
        self.progressbar = None
        

class PixorgFake():
    def __init__(self):
        self.width = 596
        self.height = 842

class Linux_Open_With():
    def __init__(self):
        try:
            self.window = tkinter.Toplevel(global_settings.root)
            self.window.protocol("WM_DELETE_WINDOW", lambda: None)
            self.window.rowconfigure(1, weight=1)
            self.window.geometry('600x800')
            self.window.columnconfigure(1, weight=1)
            self.window.title("Abrir com:")
            self.labelask = tkinter.Label(self.window, text="Lista de Programas")
            self.labelask.grid(row=0, column=0, columnspan=3, sticky='nsew', pady=5)
            self.file_path = None
            self.programslistbox = tkinter.Listbox(self.window, listvariable = None, selectmode= 'simple', exportselection=False)
            self.filled = False
            self.allprogramas = set()
            self.programslistbox.grid(row=1, column=0, columnspan=2, sticky='nsew', pady=0, padx=0)
            self.programslistboxscroll = ttk.Scrollbar(self.window, orient="vertical")
            self.programslistboxscroll.grid(row=1, column=2, sticky='nsew')
            self.programslistboxscroll.config( command = self.programslistbox.yview )
            self.programslistbox.configure(yscrollcommand=self.programslistboxscroll.set)
            self.frame = tkinter.Frame(self.window)
            self.frame.grid(row=2, column=0, columnspan=2, sticky='nsew', pady=0, padx=0)
            self.frame.rowconfigure(0, weight=1)
            self.frame.columnconfigure((0,1), weight=1)
            self.answer0 = tkinter.Button(self.frame, font=global_settings.Font_tuple_Arial_8,\
                                          image = global_settings.openbi, compound="right", text="Abrir", command=  partial(self.open_file))
            self.answer1 = tkinter.Button(self.frame, font=global_settings.Font_tuple_Arial_8,\
                                          image = global_settings.closebi, compound="right", text="Cancelar", command=  partial(self.cancel))
                
                
            self.select_exec = tkinter.Button(self.frame, font=global_settings.Font_tuple_Arial_8,\
                                          image = global_settings.openfilebi, compound="right", text="Selecionar executável", command=  partial(self.select_exec))
            self.select_exec.grid(row=2, column=0, columnspan=3, sticky='ns', pady=5, padx=10)
            self.answer0.grid(row=3, column=0, sticky='ns', pady=5, padx=10)
            self.answer1.grid(row=3, column=2, sticky='ns', pady=5, padx=10)
            self.executable = None
            
            
            self.fill_list()
        except Exception as ex:
            utilities_general.utilities_general.printlogexception(ex=ex)
            
    def select_exec(self):
        executavel = askopenfilename(filetypes=[("Todos os arquivos", "*")])
        print("executavel", executavel)
        if(executavel!=None and executavel!='' and os.path.exists(executavel)):
            if global_settings.plt == "Windows":
                aplicativo_selecionado = executavel.replace("/","\\")
            try:
                aplicativo_selecionado = executavel
                result = subprocess.Popen([aplicativo_selecionado, self.file_path])
                if(aplicativo_selecionado not in self.allprogramas):
                    res=messagebox.askquestion('FERA', 'Deseja adicionar o programa selecionado aos programas recentes?')
                    if res.lower() == 'sim' or res.lower() == 'yes':
                        self.programslistbox.insert(tkinter.END, aplicativo_selecionado)
                        if global_settings.plt == "Windows":
                            programas = os.path.join(utilities_general.get_application_path(), "programas.txt")
                            if(os.path.exists(programas)):
                                with open(programas, 'a') as programs:     
                                    programs.write(f"\n{aplicativo_selecionado};\n")
            except Exception as ex:
                utilities_general.utilities_general.printlogexception(ex=ex)
            finally:
                self.window.withdraw()
                self.file_path = None
                self.executable = None
            
    def fill_list(self):
        if(not self.filled):
            if global_settings.plt == "Linux":
                command = "for app in /usr/share/applications/*.desktop ~/.local/share/applications/*.desktop; do app=\"${app##/*/}\"; echo \"${app::-8}\"; done"
                result = subprocess.run([command], shell=True, executable="/bin/bash", stdout=subprocess.PIPE)
                output = result.stdout.decode('utf-8')
                for linha in output.split("\n"):
                    if(linha.strip()==""):
                        continue
                    if(linha=="vi" or linha=="vim"):
                        continue
                    self.programslistbox.insert(tkinter.END, linha)
                    self.allprogramas.add(linha)
            elif global_settings.plt == "Windows":
                programas = os.path.join(utilities_general.get_application_path(), "programas.txt")
                if(os.path.exists(programas)):
                    with open(programas, 'r') as programs:
                        linhas = programs.readlines()
                        for linha in linhas:
                            if(linha.strip()==""):
                                continue
                            programapath = linha.split(";")[0]
                            if(os.path.exists(programapath)):
                                self.programslistbox.insert(tkinter.END, programapath.replace("/","\\"))
                                self.allprogramas.add(programapath)
            self.filled = True
            
    def open_file(self):
        try:
            aplicativo_selecionado_index = self.programslistbox.curselection()[0]
            aplicativo_selecionado = self.programslistbox.get(aplicativo_selecionado_index)
            #print([aplicativo_selecionado, self.file_path])
            result = subprocess.Popen([aplicativo_selecionado, self.file_path])
        except Exception as ex:
            utilities_general.utilities_general.printlogexception(ex=ex)
        finally:
            self.window.withdraw()
            self.file_path = None
            self.executable = None
        
    def cancel(self):
        self.window.withdraw()
        self.executable = None
        
    def open_with_window_visible(self, file_path):
        self.executable = None
        basename = os.path.basename(file_path)
        self.file_path = file_path
        
        self.window.title(f"{basename} - Abrir com:")
        self.window.lift()
        self.window.deiconify()
        

class Ask_for_margins():
    
    def __init__(self, patpdf):
        
        try:
            self.pathpdf = patpdf
            self.info_pdf_margin = None
            nomarginb = b'iVBORw0KGgoAAAANSUhEUgAAAGYAAACACAYAAAD5whz9AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsQAAA7EAZUrDhsAAAUhSURBVHhe7d3JL3NtGMfxy9A8NcRKYtfEypa/QCLUHGqeVggRRNroX1EhImZ25pnEEBZCscPWSkIsLAwJSsXQ59wn95v39ajT8rR1efv7JIfet0Sq3/RMPSdCXAqHw0FWq5UWFhbI6XTS/9Xz87P69728vFBYWBgVFhbS+Pi4/CkzIozJZHKJh8G4GI1G8RKwQycnJ26fcDAtZrNZvhx8kN1ud/tkg21paWlxvb6+ypfl+4UqT0Z5XtDZ2UlKHLFqlzPfK1R+B0VXVxc1NzeziBOytbXlSk5OlsO3LBYL6XQ6Ofq5NjY26ODgQI48a2hooO7ubgoJCZEz30CEUb65XZTdaHV999M1Nja6/fu0ltraWpeyWy1/Q+BhVfaBoaEhqquro+/aBgd1mJiYGPnIveHh4W+LE9RhxLbVZrPJkXsiTlVVlXq2IJCCflUmTkW1tbXJkXvitI2II07pBAq2MYrW1lZqb2+XI/cmJiYCGgdhJHFo0NHRobmLPDk5SZWVlQGJgzD/YTabqaenRzPO1NQUVVRU+D0OwvxBHFz29fVpxpmenqby8nJ6enqSM76HMG7U19dTf38/hYZ+/PLMzMyo7xx/xUGYD4jjF/HO8RSnrKzML3EQRoOIMzAwoBlnbm6OTCYTPT4+yhnfQBgPamtraXBwUDPO8vIyFRQU+DQOwnihpqZGPXemFWdlZUV95/jqmgmE8VJ1dTWNjIyoF3F8ZHV11WdxEOYTxF6YiBMeHi5n3ltbW/NJHIT5JHH84k2cvLw8enh4kDOfhzBfIHaRR0dHNeOsr69Tfn7+l+MgzBeVlpbS2NiYxziZmZl0d3cnZ7yHMH+hpKRE/UhA67qIra0tys7O/nQchPlLxcXFHuNsb29TVlbWp+IgjA8UFRWpZwB+/folZ96z2+3qau329lbOaEMYH8nNzaXZ2VnNODs7O17HQRgfysnJ8fjO2d3d9SoOwviY2NDPz8+TXq+XM++JOBkZGXRzcyNn3kMYPxAbek9x9vb2KCUlha6vr+XMWwjjJ2J1JW4E04qzv79PaWlpdHV1JWf+hTB+JFZXi4uLFBERIWfe+ygOwvhZenq6xzjigvfU1FS6vLyUMwgTEEajkZaWljTjHB4eUkJCAp2enqpjzdswxE2zkZGRcvRzNTU1qbdV/Ck2NpYSExPlyP+Ojo7o7OxMjtyLi4uj8/Nz3IbBcRFNgmJV5umqfm6Oj4+DYxuTlJQkH/0MBoMhOMKIY4r4+Hg54k98xhMUYaKjo9VPHMWG9acIir2yf1xcXFBvb6963PCVTxV9SdwItbm5KUdviQ/XgioMJ/f39xQVFSVHb4kwOMBkCmGYQhimEIYphGEKYZhCGKYQhimEYQphmEIYphCGKYRhCmGYQhimEIYphGEKYZhCGKYQhimEYQphmEIYphCGKYRhCmGYQhimEIYphGEKYZhCGKYQhimEYQphmEIYphCGKYRhCmGYQhimEIYphGEKYZhCGKYQhimEYQphmEIYphCGKYRhCmGYQhimEIYphGEKYZhCGKYQhimEYQphmEIYphCGKYRhCmGYQhimEIYphGEKYZhCGKYQhimEYQphmEIYphCGKYRhCmGYQhimEIYpzX+CbbFYSKfTyRH40tPTE7W3t8vRW+KfYIsvLuUxFkaLaIJVGVOhCvkQuBBNQg0GgxwCF2oTl8JkMrld12EJ/CJaCCHii8PhIKvVSgsLC+R0OpWfQ6Dp9XrKz88nm81GUVFR9BuTeESk5mol/gAAAABJRU5ErkJggg=='
            mmmarginb = b'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'
            cmarginb = b'iVBORw0KGgoAAAANSUhEUgAAAGYAAACACAIAAAB2oIuqAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAT+SURBVHhe7dlJS+tQGMbxWinqQuGiIAU3fgTrrk4Il+KAYMVxrRSKiNjNxa1bRbsTxWHnwq1aJ1xIwS/RVdcOoK0DTr0n9hWbNql5Yltz0vcHVxtpkpM/OWmT60ilUslkMhgMut3uPzKora11uVxOp1P8HB8fF+MvMSWZ3+93SMvn86WPpGQc8Xicdi6tubk5OpqScESjUdqzzGZnZ9/f3+mYisxxfn5Ou5XczMxMaarZJ5kwPT1dgmrayUKh0D/r8Xg8ND594qO/2NW0k93f39MrKxEnEb3Ka2pq6u3tjY6vCGyYTJicnCxeNbmT1dXV0ascxasmd7KBgYHFxUVayDExMfH6+koHWjjSJxM/l5aW0ou5xB3Vy8sLHWuB2CGZsLy8nH6Ra2xsrLDVbJJMWFlZqaiooAW10dHRAlazTzJhdXVVr9rIyEihqtkqmbC2tqZXbXh4+Pn5mY77B+yWTFhfX3c6nbSgVpBqNkwm5Kk2NDT0w2r2TCZsbGzoVevv7396eqIAONsmEzY3N/Wq9fX1ma5m52TC1taWXrXe3t7Hx0fKgLB5MmFnZ6eyspIW1Hp6ekxUkylZMZioVu7JBJ/P9/DwQD0M4GQKqBonI11dXYlEgqrkxcm+dHZ2GqlWyGS0yaIx94mpyeVy0Su1jo6Ob6uVabK9vb2qqipaUGtvb7+7u6NdainTZML+/r5etba2tjzVyjeZcHBwYKJaWScTIpFIdXU1Lah5vd7b21vad4ZyTyYcHh7qVWttbb25uaHdf+JkiqOjozzVrq+vaQQfOBk5Pj6uqamhBTWPx5NZjZN9OTk50avW0tJydXWVHgYnUzk9PdWrVl9fH4/HxTDkTtbQ0PDXGFrhE/1VS1NTE70pR2NjoxiG3MlKT8klUbL5+Xna0+/Z3t6WKdnu7i7t6fecnZ3JlCyRSDQ3N9POfolkE1O4uLgQ12Da32+QL5lweXm5sLAwODhIn3DG0BA/0V/z6u7upndnkDKZOTREhGYETpYPJ4NxMhgng3EyGCeDcTJY6ZPF6Lc10BARnAzGyWClT2YtNEQEJ4NJkSwWi4QDAa/XSxv8IBYDgXDkh/OcNoawdjIllaqTJm/AfDjaBMLCySIGan3xhk1lo7UR1k0WC2cV+5iJaeop+ikQoVUBtCpCgmSiVCSmcQrFck9D/FSjFRHWTSYmprhKfXd9F29SQc80Wg1h4WQGZUUDm9FaCPmTZV30wLlJayFskCyrGXaa0UoIOyRTz03sNKOVEJwMZr9kPDENUF3L+PJvwE+u/uWZTP29zJZfZaF58x31GVb2N0x5fTxCU+UyE8z2ySLqUyoLOiXTaGWERMmybiYziLt3s9OdtoCQOZnyAC2s+VzIONoUQtJkZi5bmmh7CE4G42QwiZJlfgfjZIZwMhgng8W+LmbesKnvrRpogwiJkhUFDRHByWCcDCZRsljGowuvuZvwXLQ9hDzJsm4xC/SZSVtDSJMs62Gi2Yc92WhjCE4GkyYZT0wT+PJvDTREBCeDcTIYJ4NxMhgng3EyGCeDcTIYJ4NxMhgng3EyWCmS2Qwng3EyGCeDcTIYJ4NxMhgng3EyGCeDcTIYJ4NxMhgng3EyGJYsFAr9K3siAuXIoJuM6eFkMCVXNBqlJWaAkisej9MSM0DJlUql/H4//YHlJUIp//Mi/iWTyWAw6Ha7/zAdIo5IJEKlUqn/JXQXaRZ8z7QAAAAASUVORK5CYII='
            self.nomargin = tkinter.PhotoImage(data=nomarginb)
            self.mmmargin = tkinter.PhotoImage(data=mmmarginb)
            self.cmargin = tkinter.PhotoImage(data=cmarginb)
            
            self.window = tkinter.Toplevel(global_settings.root)
            self.window.protocol("WM_DELETE_WINDOW", lambda: None)
            self.window.rowconfigure((0,1,2), weight=1)
            self.window.columnconfigure((0,1,2), weight=1)
            self.labelask = tkinter.Label(self.window, text="Deseja definir a(s) margens do(s) documento(s)?")
            self.labelask.grid(row=0, column=0, columnspan=3, sticky='ns', pady=5)
            self.definirmargens = [True]
            self.answer0 = tkinter.Button(self.window, font=global_settings.Font_tuple_Arial_8,\
                                          image = self.mmmargin, compound="top", text="Margens mobileMerger", command=  partial(self.defineMargins, 0))
            self.answer1 = tkinter.Button(self.window, font=global_settings.Font_tuple_Arial_8,\
                                          image = self.nomargin, compound="top", text="Sem Margens", command=  partial(self.defineMargins, 1))
            self.answer2 = tkinter.Button(self.window, font=global_settings.Font_tuple_Arial_8,\
                                          image = self.cmargin, compound="top", text="Margens Personalizadas", command=  partial(self.defineMargins, 2))
            
            self.answer0.grid(row=1, column=0, sticky='ns', pady=5, padx=10)
            self.answer1.grid(row=1, column=1, sticky='ns', pady=5, padx=10)
            self.answer2.grid(row=1, column=2, sticky='ns', pady=5, padx=10)
            self.askformargins = tkinter.BooleanVar()
            self.askformargins.set(0)
            self.dontask = ttk.Checkbutton(self.window,text='Não perguntar novamente', variable=self.askformargins)
            self.dontask.grid(row=2, column=0, columnspan=3, sticky='ns', pady=5)
            self.info_pdf_margin = None
        except Exception as ex:
            utilities_general.utilities_general.printlogexception(ex=ex)
            
        
    def defineMargins(self, tipomargem):
        pathpdf = Path(self.pathpdf)
        
        doc = None
        try:
            topmargemmm = math.floor(115/72 * 25.4)     
            bottommargemmm = max(0, math.floor((842-813)/72 * 25.4))
            leftmargemmm = 30
            rightmarmmm = 10
            doc = fitz.open(self.pathpdf)
            
            if(tipomargem==2):
                
                pixorg= doc[0].get_pixmap()
                document_margin = classes_general.Document_Margin(doc)
                if(not document_margin.marginsok):
                    raise Exception()
                else:
                    topmargemmm = document_margin.mt     
                    bottommargemmm = document_margin.mb
                    leftmargemmm = document_margin.me
                    rightmarmmm = document_margin.md
                    self.info_pdf_margin= (self.pathpdf, document_margin.mt, document_margin.mb, document_margin.me, document_margin.md, document_margin.pixorg, len(doc))
    
            elif(tipomargem==1):                  
                pixorg= doc[0].get_pixmap()
                topmargemmm = 0     
                bottommargemmm = 0
                leftmargemmm = 0
                rightmarmmm = 0
                self.info_pdf_margin= (self.pathpdf, 0, 0, 0, 0, pixorg, len(doc))
                
            elif(tipomargem==0):
                
                fakepixorg = PixorgFake()                
                pixorg= doc[0].get_pixmap()
                topmargemmm = math.floor(115/72 * 25.4)     
                bottommargemmm = max(0, math.floor((842-813)/72 * 25.4))
                leftmargemmm = 30
                rightmarmmm = 10
                self.info_pdf_margin = (self.pathpdf, topmargemmm, bottommargemmm, leftmargemmm, rightmarmmm, fakepixorg, len(doc))
            else:
                None
            global_settings.askformargins = not self.askformargins.get()
            if(self.askformargins):
                global_settings.default_margin_top = topmargemmm
                global_settings.default_margin_bottom = bottommargemmm
                global_settings.default_margin_left = leftmargemmm
                global_settings.default_margin_right = rightmarmmm
            self.window.destroy()
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        finally:
            try:
                doc.close()
            except:
                None

class Filter_Window():
    def close(self):
        self.window.withdraw()
        
    def select(self, todos=False):
        for key in self.intvars.keys():
            if(todos):
                self.intvars[key][0].set(1)
            else:
                self.intvars[key][0].set(0)
    def apply(self):
       
        pdfs_to_filter_out = []
        for key in self.intvars.keys():
            intvar = self.intvars[key][0]
            if(intvar.get()==0):
               idpdf1 = self.intvars[key][1] 
               pdfs_to_filter_out.append(idpdf1)
        """
        #termos = self.treeview_to_filter.get_children('')
        for termo in global_settings.listaTERMOS.keys():
            soma = 0
            idtermo = global_settings.listaTERMOS[termo][2]
            idtermo_item = "t"+str(idtermo)
            pdfs = sorted(global_settings.idtermopdfs[idtermo_item].keys(), key=lambda x: x[1])
            index_insert = 0
            for pdf in pdfs:
                idpdf2 = self.treeview_to_filter.item(pdf[0], 'values')[1]
                texto = self.treeview_to_filter.item(pdf[0], 'text')
                if(idpdf2 in pdfs_to_filter_out):
                    self.treeview_to_filter.detach(pdf[0])
                else:      
                    #print(pdf, index_insert)                               
                    self.treeview_to_filter.move(pdf[0], idtermo_item, index_insert)
                    index_insert += 1 
                    soma += int(self.treeview_to_filter.item(pdf[0], 'values')[2])
            valores = self.treeview_to_filter.item(idtermo_item, 'values')
            self.treeview_to_filter.item(idtermo_item, text = valores[3] + ' (' + str(soma) + ')'  + " - "+valores[4])
            index_insert = utilities_general.insertIndex(self.treeview_to_filter, '', valores[3], index=0)
            self.treeview_to_filter.move(idtermo_item, '', index_insert)
            if(soma==0 and self.ferawindow.hideresultsvar.get()):
                self.treeview_to_filter.detach(idtermo_item)
            """    
        try:
            inserts_termo = {}
            inserts_termo_doc = {}
            inserts_termo_doc_toc = {}
            change = False
            for item in global_settings.searchResultsDict:
                if(item[0]=='t'):
                    continue
                
                resultsearch = global_settings.searchResultsDict[item]
                t = 't'+resultsearch.idtermo
                tp = 'tp'+resultsearch.idtermopdf
                tptoc = resultsearch.tptoc
                attrs = vars(resultsearch)
                if(t not in inserts_termo):
                    inserts_termo[t] = 0
                if(tp not in inserts_termo_doc):
                    inserts_termo_doc[tp] = 0
                if(tptoc not in inserts_termo_doc_toc):
                    inserts_termo_doc_toc[tptoc] = 0

                #print(', '.join("%s: %s" % item for item in attrs.items()))
                #print(resultsearch)
                if(resultsearch.idpdf in pdfs_to_filter_out):
                    change = True
                    #print(change)
                    self.treeview_to_filter.detach(item)
                else:
                    docu = resultsearch.idpdf
                    if(resultsearch.pagina+1 >= self.intvars[docu][2].get() and resultsearch.pagina+1 <= self.intvars[docu][3].get()):
                        
                        #print(tptoc, resultsearch.idpdf, resultsearch.idtermopdf)
                        
                        self.treeview_to_filter.move(item, tptoc, inserts_termo_doc_toc[tptoc])
                        inserts_termo[t] += 1
                        inserts_termo_doc[tp] += 1
                        inserts_termo_doc_toc[tptoc] += 1
                    else:
                        #print(resultsearch.pagina, self.intvars[docu][2].get(),  self.intvars[docu][3].get())
                        change = True
                        print(change)
                        self.treeview_to_filter.detach(item)
            for item in inserts_termo:
                valores = self.treeview_to_filter.item(item, 'values')
                termo = valores[3]
                tipobusca = valores[4]
                th = inserts_termo[item]
                self.treeview_to_filter.item(item, values=(valores[0], valores[1], th, termo.upper(), tipobusca,))
                self.treeview_to_filter.item(item, text=termo.upper() + ' (' + str(th) + ')'  + " - "+tipobusca)
            for item in inserts_termo_doc:
                valores = self.treeview_to_filter.item(item, 'values')
                pathpdfbase = valores[0]
                th = inserts_termo_doc[item]
                self.treeview_to_filter.item(item, text=pathpdfbase + ' (' + str(th) + ')') 
            for item in inserts_termo_doc_toc:
                valores = self.treeview_to_filter.item(item, 'values')
                tocname = valores[0]
                th = inserts_termo_doc_toc[item]
                self.treeview_to_filter.item(item, text=tocname + ' (' + str(th) + ')') 
            print(change)
            if(change):
                self.ferawindow.filter_docs_search_text.grid(row=4, column=2,sticky='e')
            else:
                self.ferawindow.filter_docs_search_text.grid_forget()
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)        
        self.window.withdraw()
        
    def __init__(self, treeview, treeview_to_filter, ferawindow):
        def enable_disable(docu):
            try:
                if(self.intvars[docu][0].get()==0):
                    self.intvars[docu][4].config(state='disabled')
                    self.intvars[docu][5].config(state='disabled')
                else:
                    self.intvars[docu][4].config(state='normal')
                    self.intvars[docu][5].config(state='normal')
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)
        def max_page_set(docu):
            try:
                self.intvars[docu][3].set(self.intvars[docu][6])
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)
                
            
        try:
            if(global_settings.filter_window_searches!=None):
                self.window = global_settings.filter_window_searches 
            if(global_settings.filter_window_searches==None):
                self.treeview = treeview
                self.treeview_to_filter = treeview_to_filter
                self.ferawindow = ferawindow
                self.window = tkinter.Toplevel(global_settings.root)
                global_settings.filter_window_searches  = self.window 
                self.window.protocol("WM_DELETE_WINDOW", lambda: self.close())
                self.window.rowconfigure(1, weight=1)
                self.window.columnconfigure((0,1), weight=1)
                
                self.botaoselectall = tkinter.Button(self.window, image = global_settings.selectall, compound='left', text="Marcar Todos", \
                                                  command= lambda: self.select(True))
                self.botaoselectall.grid(row=0, column=0, sticky='ns', pady=5, padx=5)
                
                self.botaounselectall = tkinter.Button(self.window, image = global_settings.unselectall, compound='left', text="Desmarcar Todos", \
                                                  command=lambda: self.select(False))
                self.botaounselectall.grid(row=0, column=1, sticky='ns', pady=5)
                
                
                self.check_docs_frame = tkinter.Frame(self.window, borderwidth=2)
                self.check_docs_frame.grid(row=1, column=0, columnspan=2, sticky='nsew', pady=15)
                row = 0
                children1 = treeview.get_children('')
                self.intvars = {}
                for eq in children1:
                    eqtext = treeview.item(eq, 'text')
                    eqlabel = tkinter.Label(self.check_docs_frame, font=global_settings.Font_tuple_ArialBold_12, text=eqtext)
                    eqlabel.grid(row=row, column=0, sticky='w', pady=5)
                    children2 = treeview.get_children(eq)
                    row += 1
                    for child in children2:
                        docu = treeview.item(child, 'text')
                        idpdf = treeview.item(child, 'values')[2]
                        path_pdf = treeview.item(child, 'values')[1]
                        tamanho_doc = global_settings.infoLaudo[path_pdf].len
                        svar_init = tkinter.IntVar()
                        svar_init.set(1)
                        svar_end = tkinter.IntVar()
                        svar_end.set(tamanho_doc)
                        entry1 = tkinter.Entry(self.check_docs_frame, justify='center', textvariable=svar_init, exportselection=False)
                        entry2 = tkinter.Entry(self.check_docs_frame, justify='center', textvariable=svar_end, exportselection=False)
                        self.intvars[idpdf] = [tkinter.IntVar(), idpdf,svar_init,svar_end,entry1, entry2, tamanho_doc]
                        self.intvars[idpdf][0].set(1)
                        chk = tkinter.Checkbutton(self.check_docs_frame, text=docu, variable=self.intvars[idpdf][0], offvalue=0, onvalue=1, command=lambda:enable_disable(idpdf))
                        chk.grid(row=row, column=0, sticky='w', pady=1)
                        label = tkinter.Label(self.check_docs_frame, font=global_settings.Font_tuple_Arial_10, text="Fls. ")
                        label.grid(row=row, column=1, sticky='w', padx=5)
                        entry1.grid(row=row, column=2, sticky='w', pady=1, padx = 5)
                        label2 = tkinter.Label(self.check_docs_frame, font=global_settings.Font_tuple_Arial_10, text=" a ")
                        label2.grid(row=row, column=3, sticky='w', padx=5)
                        
                        entry2.grid(row=row, column=4, sticky='w', pady=1, padx = 5)
                        max_page = tkinter.Button(self.check_docs_frame, text="Max.", command= lambda: max_page_set(idpdf))
                        max_page.grid(column=5, row=row, sticky='w', padx = 5) 
                        row += 1
                        
                self.botaoapply = tkinter.Button(self.window, image = global_settings.imageapply, compound='left', text="Aplicar", \
                                                  command=self.apply)
                self.botaoapply.grid(row=2, column=0, sticky='ns', pady=5)
                
                self.botaoclose = tkinter.Button(self.window, image = global_settings.imagedontapply, compound='left', text="Fechar", \
                                                  command=self.close)
                self.botaoclose.grid(row=2, column=1, sticky='ns', pady=5)
            else:
                #print(teste)
                self.window.deiconify()
                global_settings.root.after(1, lambda:self.window.lift())
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        
        
                
class Annotation_Window():
    def __init__(self, observation, annotations, feraapp):
        self.audio_list_formats = [".3gp", ".aac", ".aiff",".amr", ".flac",".m4a", ".mp3",".ogg", ".opus",".ra", ".wav",".wave", ".wma", ".mp4"]
        self.global_annot_window = tkinter.Toplevel(global_settings.root)
        self.global_annot_window.geometry('800x500')
        self.global_annot_window.rowconfigure(0, weight=1)
        self.global_annot_window.columnconfigure(0, weight=1)
        
        self.globalFrame = tkinter.PanedWindow(self.global_annot_window, sashwidth=8)
        self.globalFrame.grid(row=0, column=0, sticky="nsew")
        self.globalFrame.rowconfigure(0, weight=1)
        self.globalFrame.columnconfigure(0, weight=1)
        self.globalFrame.columnconfigure(1, weight=4)
        
        
        self.annot_tree_frame = tkinter.Frame(self.global_annot_window, borderwidth=2)
        self.annot_tree_frame.grid(row=0, column=0, sticky='nsew')
        self.annot_tree_frame.rowconfigure(0, weight=1)
        self.annot_tree_frame.columnconfigure(0, weight=1)
        self.treeview_link = ttk.Treeview(self.annot_tree_frame, selectmode='browse')
        self.treeview_link.heading("#0", text="Anotações", anchor="w")
        self.treeview_link.grid(row=0, column=0, sticky='nsew')
        self.treeview_link.bindtags(('.self.treeview_link', 'Treeview', 'post-tree-bind-link','.','all'))
        self.treeview_link.bind_class('post-tree-bind-link', "<1>", lambda e: self.treeview_selection(e, observation, annotations))
        
        
        self.annot_content_frame = tkinter.Frame(self.global_annot_window, borderwidth=2, relief='ridge')
        self.annot_content_frame.grid(row=0, column=1, sticky='nsew')
        self.annot_content_frame.rowconfigure(1, weight=1)
        self.annot_content_frame.columnconfigure(0, weight=1)
        
        self.annot_content_frame_buttons= tkinter.Frame(self.annot_content_frame, borderwidth=2, relief='ridge')
        self.annot_content_frame_buttons.grid(row=0, column=0, sticky='nsew')
        self.annot_content_frame_buttons.rowconfigure((0,1), weight=1)
        self.annot_content_frame_buttons.columnconfigure((0,1,2), weight=1)
        #self.annot_content_frame_buttons.columnconfigure((0,1,2), weight=1)
        
        self.text_box = tkinter.Text(self.annot_content_frame)
        self.content = "" 
        self.text_box.insert('end', self.content)
        
        self.text_box.grid(row=1, column=0, sticky='nsew', pady=5)
        self.text_box.bind('<Key>',self.Text_edited)
        self.processing = {}
        
        playb = b'iVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAAAAXNSR0IArs4c6QAAAyZJREFUaEPtmctrE1EUxr8zTTW2BetCoWijFnShuNOFf4gIgqBCrZoYFEQFQd24EBFbEyNYhSJIXfjY+NjoRhEEURKoj1bb2GBta1qImVc6mTtHJnXcKCamScYLvcsZhvl+95x77v3OJUg+SHL9WATwO4IVR2B5YnyFcGiLcJQxM7rmi9/Cvf9XBNByObOPCFcABMEsSKHrgaI4mTu6Puc3SFmAYGK8KyDoQ4AQ2NaxhN7OFPm7xQTGFIHCWrTznp8QZQFa4+O7wXRz58YW7NnUCtVi9A9peJIpgAG3Ctx3gIhxOPTVD5CyAG2xzCEGruzb3IYdG5b90jg0W0RfUsUXVbjP8gCf1mdDMZwlp5EgVQO4IucE485HA7dHDLYdEAMvFMXp1sLr3jcKYkEAnsh03kbvGxUjORtg2ABd1FE4g+iGuXqD1ATAFSkc4EHawMA7nQvCjQYPMTd1m9E1L+sJUTMAT+SUIRBLanjzzQKYmUjpDxrBYzMnVqr1AKk5gCfy+cQc4kmV80UmAiYFUdiMdN6vNUTdAFyhbskdeKfj0WdzXrdDD1jhnlqW3LoCeLP9atpCPKnhm1kqud+Z+LgRDvWD3MK1sNEQAFdiQTAGPxilsju/UfAzRQT2q0dWDy8EoWEAnsjRnI3epIpPOdvdxQsAn9eWZM+hZ2uxGpCGA3gl9+6ogVvvdbackidJsYJuIxx69a8QvgB4Iid1gb6khlTWclPKBimJFlp6KhtepVUK4itAaSUAeJw2ceOtzobN7qoeJeIePbL2aSUQvgN4ImcLDq6mVLyYLEXDYcJBI7L2WjmI/wbAE/psYg4XXudhO7Ag7JB+pGv6bxD/HYArNp7S8DBtuvm1S4+GBhcB/mRoyuVmte9/phDbAkU4EqXQb4uYlQNGtLO/3ET4vgakLqPSbmSue5P2KCHtYU7q47S0hkZqSymtqZe2rSJ1Y0va1qLUzV1p2uvSX3AEY5/XB6AMNysIbO9YSsmsJdcVk3seb41l9oKRAEl4yecZivZL6XbR1LTFbqYx82DnRDmj0aj3ZQ1No4RU+59FgGpnrlbfSR+BH0y+wk9jwOEuAAAAAElFTkSuQmCC'
        #playb = b'iVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAAAAXNSR0IArs4c6QAAA0dJREFUaEPtWUtrFEEQrppREDx5F7Z7Fm9REjVREEUF9eRRPBhQED0LOSUoRpTkInoO8eYDPHuLmBzzUHwEvIXpXvAnKAi6XVIyEzrD7E5PZzY7ITvH6arq76uvuqe6B2GXP7jL8cOAQL8V3DsKSCmPEdFpADjULeuI2DbGfGi1Wl9su0ajcTwIgotEFJZRrVO8NEahAlLKs0T0HBFPuE5MRH8AYExr/ZV9GDwArAZBsM81hm2XjWePdSUghLhFRPM+EyPiRBzHz3gyKeUEADz1Ab+ZacSpOI5nszE6EuDMG2MWfcBzxhBxVCn1LVFgBBFXEXG/LwlEfBTH8bQzASHEp5yyWQCAda7LTkCI6CcAvEvBp3ZCiGFEvIqIB7v4XgGA4bzxUgR4wQLA/+xZEo7HcfzaN4NFfkKI60T0qpPipQgIIe4i4pw16YJSirPTk6cIPE9aioCU8j4APE7REtGs1nqqF+jzwBtj/iLiZ0QcsyrAfQ1EUTRNRA+LnLdLSEp5zRjzJlM2bUS8CQBHXDDk7kI7QaAbeF5rrhj6QqAIPCtbWwIu4GtLwBV8LQmUAV87Ao1GYwQA1uzdJtkqx7XWb/N2s1qtgWwzVwQ+UWCSiGasrdy9mXNl7/od4D6IFeBmzgU8x2XV0gYw2xxuaXG2I58rgRRQGIbnjTFL6TmhyJ+JB0FwgYgWs83hpjI7RaAIrO94Xz5kvmDz/GpNgI+iXHbtdnspe8aufQnZ5+jSZ+KqdyGfkomiqD7bqCcBp5a+tmvAtQoGBHzKw8VnoIDLedQlk742e1uBnbxW6aSQEGIGESet8QdKqSdZ+9xdKHuxRUTvtdaXfcvBx08IsYCIlyzfO0qpF04Eoig6SkTrW/puxJ5eLdpzRVF0g68Z7XdhGA5tbGx8dyKQ9O8fgyA4aTuwEoi4QkS/fLJa5MMXv/wTJZN5IKI1rfUp526UDZvN5hnuArdzJV4E2GU8aeTOaa1XShFgYyEEX/HN94tEcpS8rZR62Yls4S+mRAn+xTTqkrGqbLhsiOheq9Va7hazkEDq3Gw2h4wxXJ+HAeBAVUAzcX4T0Y8wDJfzFmzpEuoRyErDOitQ6awVBhsQqDCZXqF2vQL/AJflpE8JZ3oTAAAAAElFTkSuQmCC'
        saveb = b'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'
        googleb = b'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'
        self.play = tkinter.PhotoImage(data=playb)
        self.save = tkinter.PhotoImage(data=saveb)
        self.google = tkinter.PhotoImage(data=googleb)
        
        self.botaosalvar = tkinter.Button(self.annot_content_frame_buttons, image = self.save, compound='left', text="", \
                                          command=partial(self.save_transcription, annotations, observation, feraapp))
        self.botaosalvar.grid(row=0, column=0, sticky='ns', pady=5)
        self.botaosalvar.image = self.save
        self.botaosalvar.config(relief='sunken', state='disabled')
        
        self.botaoplay = tkinter.Button(self.annot_content_frame_buttons, image = self.play, compound='left', text="", command=partial(self.play_file))
        self.botaoplay.grid(row=0, column=1, sticky='ns', pady=5)
        self.botaoplay.image = self.play
        
        """ self.botaogoogle = tkinter.Button(self.annot_content_frame_buttons, image = self.google, compound='left', text="", command=partial(self.execute_dir, True))
        self.botaogoogle.grid(row=0, column=2, sticky='ns', pady=5)
        self.botaogoogle.image = self.google """
        
        #self.botaovosk = tkinter.Button(self.annot_content_frame_buttons, text="VOSK", command=partial(self.execute_dir, False))
        #self.botaovosk.grid(row=0, column=3, sticky='ns', pady=5)
        #self.botaovosk.image = self.google
        
        self.progressindex = ttk.Progressbar(self.annot_content_frame_buttons, mode='indeterminate')
        self.progressindex.grid(row=1, column=0, columnspan=3, sticky='nsew', pady=5)
        self.progressindex.grid_forget() 
        
        self.treeview_link.insert(parent='', index='0', iid='main'+str(observation.idobs), text='Observação Principal', image=global_settings.itemimage)
        self.obs_info = {}
        self.obs_info['main'+str(observation.idobs)] = [observation.conteudo, 0]
        for annot in annotations:
            #coord = "{}-{}-{}-{}".format(coords[0],coords[1],coords[2],coords[3])
            idannot = str(annotations[annot].idannot)
            texto = os.path.basename(annotations[annot].link).strip()
            if(texto==''):
                texto = annotations[annot].link
            try:
                self.obs_info[idannot] = [annotations[annot].conteudo, 0]
                self.treeview_link.insert(parent='main'+str(observation.idobs), index='end', iid=idannot, text=texto, image=global_settings.itemimage)
                self.treeview_link.see(idannot)
            except:
                None
        self.annotations = annotations
        self.globalFrame.add(self.annot_tree_frame, minsize=100)
        self.globalFrame.add(self.annot_content_frame, minsize=100)
        self.treeview_link.selection_set('main'+str(observation.idobs))
        self.text_box.insert('end', observation.conteudo)
        #self.botaogoogle.config(relief='sunken', state='disabled')
        #self.botaovosk.config(relief='sunken', state='disabled')
    
    def execute_dir(self, google):
        self.progressindex.grid(row=1, column=0, columnspan=3, sticky='nsew', pady=5)
        self.progressindex.start()
        annotation = self.annotations[int(self.treeview_link.selection()[0])]
        #self.botaogoogle.config(state='disabled')
        #self.botaovosk.config(state='disabled')
        
        try:
            self.processing[annotation.link] = ""
            filepath = str(Path(utilities_general.get_normalized_path(os.path.join(Path(utilities_general.get_normalized_path(annotation.pathpdf)).parent,annotation.link))))
            self.thread1 = threading.Thread(target=self.executecommands, args=(annotation, filepath, google,))
            self.thread1.start()
            self.global_annot_window.after(1, self.execute_after)
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        finally:
            None
    
            
    def executecommands(self, annotation, filepath, google):        
        
        try:
            temp_path = None
            filename, file_extension = os.path.splitext(filepath)            
            output_data = ""
            print(file_extension)
            if file_extension.lower() not in self.audio_list_formats:
                output_data = f"Arquivo {filepath} desconsiderado"
            else:
                temp_path = os.path.join(utilities_general.get_application_path(), os.path.basename(filepath) + ".wav")
                
                if os.path.exists(output_data):
                    os.remove(output_data)
                #google = False
                executavel = os.path.join(utilities_general.get_application_path(), "ffmpeg")
                if(google):
                    cmd = f"\"{executavel}\" -nostdin -y -i \"{filepath}\" -ac 2 -f wav \"{temp_path}\""
                else:
                    sample_rate=16000                    
                    cmd = f"\"{executavel}\" -nostdin -y -i \"{filepath}\" -ar {sample_rate} -ac 1 -f wav \"{temp_path}\""
                print(cmd)
                    
                if subprocess.run(cmd, capture_output=True).returncode != 0:                
                    output_data = f"Falha na conversão do arquivo {filepath} pelo ffmpeg. Arquivo desconsiderado"
                
                if os.path.exists(temp_path):
                    # use the audio file as the audio source
                    
                    
                    google = True
                    if(google):
                        r = sr.Recognizer()
                        with sr.AudioFile(temp_path) as source:
                            audio = r.record(source)  # read the entire audio file
                
                        try:
                            # for testing purposes, we're just using the default API key
                            # to use another API key, use `r.recognize_google(audio, key="GOOGLE_SPEECH_RECOGNITION_API_KEY")`
                            # instead of `r.recognize_google(audio)`
                            output_data = "Transcrição: \n"+r.recognize_google(audio, language='pt-BR')
                            
                            #print(f"{filepath}: \n\t" + tmp)
                        except sr.UnknownValueError:
                            output_data = "Google Speech Recognition could not understand audio"
                        except sr.RequestError as e:
                            output_data = "Could not request results from Google Speech Recognition service; {0}".format(e)
                    '''
                    else:
                        wf = wave.open(temp_path, "rb")
                        rec = KaldiRecognizer(global_settings.model, wf.getframerate())
                        
                        text = ""
                        while True:
                            data = wf.readframes(1000)
                            if len(data) == 0:
                                break
                            if rec.AcceptWaveform(data):
                                jres = json.loads(rec.Result())
                                text = text + " " + jres['text']
                        output_data = "Transcrição: \n"+text
                    '''
            self.text_box.insert('end', output_data)
            
            self.obs_info[str(annotation.idannot)][0] += output_data
            self.obs_info[str(annotation.idannot)][1] = 1
            self.botaosalvar.config(relief='raised', state='normal')
            del self.processing[annotation.link]
            try:
                os.remove(temp_path)
            except:
                None
        except Exception as ex:
            print(temp_path)
            utilities_general.printlogexception(ex=ex)
        finally:
            #self.botaogoogle.config(state='normal')
            #self.botaovosk.config(state='normal')
            
            self.botaosalvar.config(relief='raised', state='normal')
            
    def execute_after(self):
        if(self.thread1.is_alive()):
            self.global_annot_window.after(10, self.execute_after)
        else:
            self.progressindex.grid_forget()        
            
            #tkinter.Text_box.update()    
    
    def Text_edited(self, e):
        #botaogoogle.config(state='normal')
        content =self.text_box.get("1.0",tkinter.END).strip()        
        idannot = str((self.treeview_link.selection()[0]))
        self.obs_info[idannot][0] = content
        self.obs_info[idannot][1] = 1
        self.botaosalvar.config(relief='raised', state='normal')
    def treeview_selection(self, e, observation, annotations):
        idannot = str(self.treeview_link.selection()[0])
        self.text_box.delete("1.0","end")
        self.text_box.insert('end', self.obs_info[idannot][0])
        self.text_box.update_idletasks()
        if(self.obs_info[idannot][1]==1):
            self.botaosalvar.config(relief='raised', state='normal')
        else:
            self.botaosalvar.config(relief='sunken', state='disabled')
        if('main' in idannot):
            #self.botaogoogle.config(relief='sunken', state='disabled')
            #self.botaovosk.config(relief='sunken', state='disabled')
            self.botaoplay.config(relief='sunken', state='disabled')
        else:
            self.botaoplay.config(relief='raised', state='normal')
            try:
                filename, file_extension = os.path.splitext(annotations[int(idannot)].link)
                if(file_extension in self.audio_list_formats):
                    if(len(self.processing)>0):
                        None#self.botaogoogle.config(relief='sunken', state='disabled')
                        #self.botaovosk.config(relief='sunken', state='disabled')
                    else:
                        None#self.botaogoogle.config(relief='raised', state='normal')
                        #self.botaovosk.config(relief='raised', state='normal')
                    #self.botaogoogle.grid(row=0, column=2, sticky='ns', pady=5)
                else:
                    None#self.botaogoogle.config(relief='sunken', state='disabled')
                    #self.botaovosk.config(relief='sunken', state='disabled')
                    #self.botaogoogle.grid_forget()
            except:
                None#self.botaogoogle.config(relief='sunken', state='disabled')
                #self.botaovosk.config(relief='sunken', state='disabled')
                #self.botaogoogle.grid_forget()
                
        
        
    def save_transcription(self, annotations, observation, feraapp):
        #print('teste')
        idannot = str((self.treeview_link.selection()[0]))
        content =self.text_box.get("1.0",tkinter.END).strip()
        if('main' in idannot):            
            sqliteconn =  utilities_general.connectDB(str(global_settings.pathdb), 5)
            try:
                if(sqliteconn==None):
                    utilities_general.popup_window("O banco de dados está ocupado.\n A operação não foi concluída, tente novamente em alguns segundos.", False)
                    return
                updateinto2 = "UPDATE Anexo_Eletronico_Obsitens set conteudo = ? WHERE id_obs = ?"
                cursor = sqliteconn.cursor() 
                cursor.custom_execute(updateinto2, (content, observation.idobs,))
                idannot = idannot.replace("main","")
                cursor.custom_execute(updateinto2, (content, int(idannot),))
                self.botaosalvar.config(relief='sunken', state='disabled')
                observation.conteudo = content
                
                sqliteconn.commit()
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)
            finally:
                if(sqliteconn):
                    sqliteconn.close()
        else:
            sqliteconn =  utilities_general.connectDB(str(global_settings.pathdb), 5)
            try:
                if(sqliteconn==None):
                    utilities_general.popup_window("O banco de dados está ocupado.\n A operação não foi concluída, tente novamente em alguns segundos.", False)
                    return
                updateinto2 = "UPDATE Anexo_Eletronico_Annotations set conteudo = ? WHERE id_annot = ?"
                cursor = sqliteconn.cursor() 
                cursor.custom_execute(updateinto2, (content, int(idannot),))
                ##cursor.custom_execute(updateinto2, (content, int(idannot),))
                annotations[int(idannot)].conteudo = content
                self.botaosalvar.config(relief='sunken', state='disabled') 
                if(observation.pathpdf==global_settings.pathpdfatual):
                    for p in range(annotations[int(idannot)].paginainit, annotations[int(idannot)].paginafim+1): 
                        if(p not in global_settings.processed_pages):
                            continue
                        posicoes_normalizadas = feraapp.normalize_positions(p, annotations[int(idannot)].paginainit, annotations[int(idannot)].paginafim, \
                                                 annotations[int(idannot)].p0x, annotations[int(idannot)].p0y, annotations[int(idannot)].p1x, annotations[int(idannot)].p1y)
                        posicaoRealX0 = posicoes_normalizadas[0]
                        posicaoRealY0 = posicoes_normalizadas[1]
                        posicaoRealX1 = posicoes_normalizadas[2]
                        posicaoRealY1 = posicoes_normalizadas[3]
                        feraapp.prepararParaQuads(p, posicaoRealX0, posicaoRealY0, posicaoRealX1, posicaoRealY1, feraapp.colorenhanceannotation, \
                                               tag=['enhanceannot'+global_settings.pathpdfatual+str(p),str(observation.idobs)+'enhanceannot'+str(idannot)], apagar=False,  \
                                                   enhancetext=False, enhancearea=True, withborder=True, alt=False)
                sqliteconn.commit()
                if(content==''):
                    feraapp.docInnerCanvas.delete(str(annotations[int(idannot)].idobs)+'enhanceannot'+str(idannot))
            except Exception as ex:
                utilities_general.printlogexception(ex=ex)  
            finally:
                if(sqliteconn):
                    sqliteconn.close()
    def play_file(self):
        annotation = self.annotations[int(self.treeview_link.selection()[0])]
        filepath = str(Path(utilities_general.get_normalized_path(os.path.join(Path(utilities_general.get_normalized_path(annotation.pathpdf)).parent,annotation.link))))
        if(not os.path.exists(filepath)):
            utilities_general.popup_window(f'O arquivo não selecionado não existe, favor verifique: \n<{filepath}>', False)
        
        elif plt == 'Darwin':       # macOS
            subprocess.call(('open', filepath), shell=True)
        elif plt == 'Windows':    # Windows
            os.startfile(filepath)    
        else:           
            openfile = ['xdg-open', filepath]
            try:                
                myenv = dict(os.environ)  # make a copy of the environment
                HOME = os.path.expanduser("~")
    
                # Single directory where user-specific data files should be written
                XDG_DATA_HOME = os.environ.get("XDG_DATA_HOME", os.path.join(HOME, ".local", "share"))
               
                # Single directory where user-specific configuration files should be written
                XDG_CONFIG_HOME = os.environ.get("XDG_CONFIG_HOME", os.path.join(HOME, ".config"))
               
                # List of directories where data files should be searched.
                XDG_DATA_DIRS_LIST = [XDG_DATA_HOME] + "/usr/local/share:/usr/share".split(":")
                XDG_DATA_DIRS = ':'.join((t) for t in XDG_DATA_DIRS_LIST)
                # List of directories where configuration files should be searched.
                XDG_CONFIG_DIRS_LIST = [XDG_CONFIG_HOME] + "/etc/xdg".split(":")
                XDG_CONFIG_DIRS = ':'.join((t) for t in XDG_CONFIG_DIRS_LIST)
                #lp_key = 'LD_LIBRARY_PATH'  # for GNU/Linux and *BSD.
                myenv['XDG_DATA_HOME'] = XDG_DATA_HOME
                myenv['XDG_CONFIG_HOME'] = XDG_CONFIG_HOME
                myenv['XDG_DATA_DIRS'] = XDG_DATA_DIRS
                myenv['XDG_CONFIG_DIRS'] = XDG_CONFIG_DIRS
                
                subprocess.run(openfile, check=True, env=myenv)
                #outs, errs = proc.communicate()
            except Exception as ex:
                webbrowser.open_new_tab(filepath)
                utilities_general.printlogexception(ex=ex)
                utilities_general.popup_window('O arquivo não possui um \nprograma associado para abertura!', False) 

class Document_Margin():
    
    def popupcomandook(self):
        self.marginsok = True
        self.mt = self.datatop.get()
        self.mb = self.databottom.get()
        self.me = self.dataleft.get()
        self.md = self.dataright.get()
        if(self.save_default_margins_var.get()):
            global_settings.default_margin_top = self.datatop.get()
            global_settings.default_margin_bottom = self.databottom.get()
            global_settings.default_margin_left = self.dataleft.get()
            global_settings.default_margin_right = self.dataright.get()
        self.window.destroy()
    
    def popupcomandocancel(self):
        self.marginsok = False
        self.window.destroy()
    
    def create_rectanglex(self, x1, y1, x2, y2, color, **kwargs):
            image = Image.new('RGBA', (x2-x1, y2-y1), color)   
            return ImageTk.PhotoImage(image)
    
    def checkmargin(self):
        try:
            self.mmtopxtop = math.floor(self.datatop.get()/25.4*72)
            self.mmtopxbottom = math.ceil(self.pixorg.height-(self.databottom.get()/25.4*72))
            self.mmtopxleft = math.floor(self.dataleft.get()/25.4*72)
            self.mmtopxright = math.ceil(self.pixorg.width-(self.dataright.get()/25.4*72))
            if(self.mmtopxbottom>=0 and self.mmtopxbottom> self.mmtopxtop and
                self.mmtopxright>=0 and self.mmtopxright > self.mmtopxleft):
                self.fotocanvas.delete("margem")
                self.x0k = math.floor(self.mmtopxleft/self.prop)
                self.y0k = math.floor(self.mmtopxtop/self.prop)
                self.x1k = math.ceil(self.mmtopxright/self.prop)
                self.y1k =  math.ceil(self.mmtopxbottom/self.prop)
                self.margemimg[0] = self.create_rectanglex(self.x0k, self.y0k, self.x1k, self.y1k, (21, 71, 150, 85))
                self.margempreimage[0] = self.fotocanvas.create_image(self.x0k, self.y0k, image=self.margemimg[0], anchor='nw', tags=("margem"))
                self.button_ok.config(relief='raised', state='active')
            else:
                #None
                self.button_ok.config(relief='sunken', state='disabled')
        except Exception as ex:
            self.button_ok.config(relief='sunken', state='disabled')
    
    def __init__(self, pdf):
        self.marginsok = True
        self.window = tkinter.Toplevel(global_settings.root)
        self.window.protocol("WM_DELETE_WINDOW", lambda : None)
        self.window.geometry("800x640")
        self.window.rowconfigure(1, weight=1)
        self.window.columnconfigure(0, weight=1)
        self.prop = 1.5
        self.loadedpage = pdf[math.floor(len(pdf)/self.prop)]
        self.pixorg = self.loadedpage.get_pixmap()    
        #72pixels = 1 inch = 25,4mm
        self.inch = 2.54
        
        self.ajustar = tkinter.Frame(self.window, borderwidth=2, bg='white', relief='ridge')
        self.ajustar.grid(row=0, column=0, sticky='nsew')
        self.ajustar.rowconfigure(0, weight=1)
        self.ajustar.columnconfigure((0,1,2,3,4,5,6,7), weight=1)
        self.datatop = tkinter.IntVar()
        self.databottom = tkinter.IntVar()
        self.dataleft = tkinter.IntVar()
        self.dataright = tkinter.IntVar()
        
        self.datatop.set(global_settings.default_margin_top)     
        self.databottom.set(global_settings.default_margin_bottom)
        self.dataleft.set(global_settings.default_margin_left)       
        self.dataright.set(global_settings.default_margin_right)
        
        self.top = tkinter.Label(self.ajustar, text='Superior (mm): ')
        self.top.grid(row=0, column=0, sticky='e')
        self.entrytop = tkinter.Entry(self.ajustar, textvariable=self.datatop)
        self.entrytop.grid(row=0, column=1, sticky='nsw')
        self.bottom = tkinter.Label(self.ajustar, text='Inferior (mm): ')
        self.bottom.grid(row=0, column=2, sticky='e')
        self.entrybottom = tkinter.Entry(self.ajustar, textvariable=self.databottom)
        self.entrybottom.grid(row=0, column=3, sticky='nsw')
        self.left = tkinter.Label(self.ajustar, text='Esquerda (mm): ')
        self.left.grid(row=0, column=4, sticky='e')
        self.entryleft = tkinter.Entry(self.ajustar, textvariable=self.dataleft)
        self.entryleft.grid(row=0, column=5, sticky='nsw')
        self.right = tkinter.Label(self.ajustar, text='Direita (mm): ')
        self.right.grid(row=0, column=6, sticky='e')
        self.entryright = tkinter.Entry(self.ajustar, textvariable=self.dataright)
        self.entryright.grid(row=0, column=7, sticky='nsw')
        
        self.fotodoc = tkinter.Frame(self.window, borderwidth=2, relief='ridge')
        self.fotodoc.grid(row=1, column=0, sticky='nsew')
        self.fotodoc.rowconfigure(0, weight=1)
        self.fotodoc.columnconfigure(0, weight=1)
        self.fotocanvas = tkinter.Canvas(self.fotodoc, bg='gray', highlightthickness=0, relief="raised")
        self.fotocanvas.grid(row=0, column=0, sticky='nsew')
        self.mat = fitz.Matrix(1/self.prop, 1/self.prop)
        self.pix = self.loadedpage.get_pixmap(alpha=False, matrix=self.mat) 
        self.imgdata = self.pix.tobytes("ppm")
        self.docimg = tkinter.PhotoImage(data = self.imgdata)    
        self.docvanvas = self.fotocanvas.create_image((0,0), anchor='nw', image=self.docimg)
        
        self.bottomframe = tkinter.Frame(self.window, borderwidth=2, relief='ridge')
        self.bottomframe.grid(row=2, column=0, sticky='nsew')
        self.bottomframe.rowconfigure((0,1), weight=1)
        self.bottomframe.columnconfigure((0,1), weight=1)
        self.save_default_margins_var = tkinter.BooleanVar()
        self.save_default_margins_var.set(0)
        self.save_default_margins_check = ttk.Checkbutton(self.bottomframe,text='Salvar configurações de margem', variable=self.save_default_margins_var)
        self.save_default_margins_check.grid(row=0, column=0, sticky='w', padx=5)
        self.button_ok = tkinter.Button(self.bottomframe, text="OK", font=global_settings.Font_tuple_Arial_8, command= \
                                        lambda : self.popupcomandook())
        self.button_ok.grid(row=1, column=0, sticky='ns', pady=5)
        self.button_cancel = tkinter.Button(self.bottomframe, text="Cancelar", font=global_settings.Font_tuple_Arial_8, command= lambda : self.popupcomandocancel())
        self.button_cancel.grid(row=1, column=1, sticky='ns', pady=5)
        
        self.mmtopxtop = math.floor(self.datatop.get()/25.4*72)
        self.mmtopxbottom = math.ceil(max(0, self.pixorg.height-(self.databottom.get()/25.4*72)))
        self.mmtopxleft = math.floor(self.dataleft.get()/25.4*72)
        self.mmtopxright = math.ceil(self.pixorg.width-(self.dataright.get()/25.4*72))
        
        self.x0k = math.floor(self.mmtopxleft/self.prop)
        self.y0k = math.floor(self.mmtopxtop/self.prop)
        self.x1k = math.ceil(self.mmtopxright/self.prop)
        self.y1k =  math.ceil(self.mmtopxbottom/self.prop)
        
        self.margemimg = [self.create_rectanglex(self.x0k, self.y0k, self.x1k, self.y1k, (21, 71, 150, 85))]
        self.margempreimage = [self.fotocanvas.create_image(self.x0k, self.y0k, image=self.margemimg[0], anchor='nw', tags=("margem"))]
        self.dataleft.trace_add("write", lambda *args: self.checkmargin())
        self.databottom.trace_add("write", lambda *args : self.checkmargin())
        self.datatop.trace_add("write", lambda *args : self.checkmargin())
        self.dataright.trace_add("write", lambda *args : self.checkmargin())
        global_settings.root.wait_window(self.window)
        

@total_ordering
class PedidoSearch():
    def __init__(self, prior, tipo, termo):
        self.prior = prior
        self.tipo = tipo
        self.termo = termo
    def __eq__(self, other):
       return isinstance(other, PedidoSearch) and self.prior == other.prior
    def __lt__(self, other):       
       return isinstance(other, PedidoSearch) and self.prior < other.prior
   
@total_ordering
class PedidoInQueue():
    def __init__(self, idtermo, tipobusca, termo, pesquisados):
        self.idtermo = idtermo
        self.tipobusca = tipobusca
        self.termo = termo
        self.pesquisados = pesquisados
    def __eq__(self, other):
       return isinstance(other, PedidoInQueue) and self.idtermo*-1 == other.idtermo*-1
    def __lt__(self, other):       
       return isinstance(other, PedidoInQueue) and self.idtermo*-1 < other.idtermo*-1

@total_ordering
class ResultSearch():
    def __init__(self):
        self.idtermopdf = None
        self.idtermo = None
        self.idpdf = None
        self.snippet = None
        self.init = None
        self.fim = None
        self.pathpdf =None
        self.pagina = None
        self.termo = None
        self.tipobusca = None
        self.counter = None
        self.fixo = None
        self.lenresults = None
        self.toc = None
        self.t = None
        self.tp = None
        self.tptoc = None
        self.prior= None
        self.end= False
    def __eq__(self, other):
       if(self.prior == other.prior):
           if(self.idpdf == other.idpdf):
               return isinstance(other, ResultSearch) and self.counter == other.counter
           else:
               return isinstance(other, ResultSearch) and self.idpdf == other.idpdf
           
       else:
           return isinstance(other, ResultSearch) and self.prior == other.prior

    def __lt__(self, other):
       if(self.prior == other.prior):
           if(self.idpdf == other.idpdf):
               return isinstance(other, ResultSearch) and self.counter < other.counter
           else:
               return isinstance(other, ResultSearch) and self.idpdf < other.idpdf
           
       else:
           return isinstance(other, ResultSearch) and self.prior < other.prior

class Rect():
    def __init__(self):
        self.image = None
        self.photoimage = None
        self.idrect = None
        self.x0 = None
        self.x1 = None
        self.y0 = None
        self.y1 = None
        self.quads = []
        self.quadsCanvas = []
        self.pagina = None
        self.offset = None
        self.char = []        

class RelatorioSuccint:
    def __init__(self, idpdf, toc, lenpdf, pixorgw, pixorgh, mt, mb, me, md, paginasindexadas, rel_path_pdf, abs_path_pdf, tipo):
        self.idpdf = idpdf
        self.tocpdf = global_settings.manager.list(toc)
        self.lenpdf = lenpdf
        self.pixorgw = pixorgw
        self.pixorgh = pixorgh
        self.mt = mt
        self.mb = mb
        self.me = me
        self.md = md
        self.paginasindexadas = paginasindexadas
        self.rel_path_pdf = rel_path_pdf
        self.abs_path_pdf = abs_path_pdf 
        self.tipo = tipo
        self.continuar_a_indexar = True

class Relatorio():
    def __init__(self):
        self.id = None
        self.toc = []
        self.len = -1
        self.pixorgw = None
        self.pixorgh = None 
        self.mt = None
        self.mb = None
        self.me = None
        self.md = None
        self.mapeamento = {}
        self.quadspagina = {}
        self.links = {}
        self.images = {}
        self.linkscustom = {}
        self.linksporpagina = {}
        self.retangulosDesenhados = {}
        self.widgets = {}
        self.ultimaPosicao = {}
        self.ref_to_page = {}
        self.name_to_dest = {}
        self.tipo = None
        self.rel_path_pdf = ''
        self.paginasindexadas = 0
        self.hash = ''
        self.status = ''
        self.parent_alias = ""
        self.zoom_pos = 0
        self.something_changed = False
        
class RespostaDePaginaXML():
    def __init__(self):
        self.qualPagina = None
        self.mapeamento = None
        self.quadspagina = None
        self.links = None
        self.widgets = None
        self.qualPdf = None        

class RespostaDePagina():
    def __init__(self):
        self.pix = None
        self.imgdata = None
        self.qualPagina = None
        self.qualLabel = None
        self.qualGrid = None
        self.qualPdf = None
        self.zoom = None
        self.height = None
        self.width = None
        
class PedidoDePagina():
     def __init__(self, qualLabel = None, qualPdf = None, qualPagina = None, matriz = None, \
                  pixheight = None, pixwidth = None, zoom = None, scrollvalue = None ,\
                      scrolltotal = None, canvash = None, mt = None, mb = None, me = None, md = None):
        self.qualLabel = qualLabel
        self.qualPdf = qualPdf
        self.qualPagina = qualPagina
        self.matriz = matriz
        self.pixheight = pixheight
        self.pixwidth = pixwidth
        self.zoom = zoom
        self.scrollvalue = scrollvalue
        self.scrolltotal = scrolltotal
        self.canvash = canvash
        self.mt = mt
        self.mb = mb
        self.me = me
        self.md = md
        

class ExternalLink():
    def __init__(self, pagina, p0x, p0y, p1x, p1y, arquivo, imagem=None):
        self.pagina = pagina
        self.p0x = p0x
        self.p0y = p0y
        self.p1x = p1x
        self.p1y = p1y
        self.arquivo = arquivo
        self.imagem = imagem

class Observation():
    def __init__(self, paginainit, paginafim, p0x, p0y, p1x, p1y, tipo, pathpdf, idobs, idobscat, conteudo, annotations, withalt=False):
        self.paginainit = paginainit
        self.paginafim = paginafim
        self.p0x = p0x
        self.p0y = p0y
        self.p1x = p1x
        self.p1y = p1y
        self.tipo = tipo
        self.pathpdf = pathpdf
        self.idobs = idobs
        self.conteudo = conteudo
        self.idobscat = str(idobscat)
        #self.obscat = str(obscat)
        self.withalt = withalt
        self.annotations = annotations
        
class Annotation():
    def __init__(self, pathpdf, idannot, paginainit, paginafim, p0x, p0y, p1x, p1y, idobs, conteudo, link):
        self.idannot = idannot
        self.paginainit = paginainit
        self.paginafim = paginafim
        self.p0x = p0x
        self.p0y = p0y
        self.p1x = p1x
        self.p1y = p1y
        self.idobs = idobs
        self.pathpdf = pathpdf
        self.conteudo = conteudo
        self.link = link

class NoteTooltip(object):
    '''
    create a tooltip for a given widget
    '''
    def __init__(self, canv, widget, conteudo, text=''):
        try:
            self.widget = widget
            self.text = text
            self.conteudo = conteudo
            self.tw = None
            self.canv = canv
            self.canv.tag_bind(widget, "<Enter>", self.enter)
            self.canv.tag_bind(widget, "<Leave>", self.close)
            self.canv.tag_bind(widget,"<Motion>",self.enter)
                
        except Exception as ex:
            utilities_general.utilities_general.printlogexception(ex=ex)
            
    
            
    def enter(self, event=None):
        try:
            if self.tw:
                self.tw.destroy()
            x = y = 0
            x = event.x_root + 10
            y = event.y_root + 10
            # creates a toplevel window
            self.tw = tkinter.Toplevel(self.canv, background='#ededd3')
            self.tw.rowconfigure(1, weight=1)
            self.tw.columnconfigure(0, weight=1)
            # Leaves only the label and removes the app window
            self.tw.wm_overrideredirect(True)
            self.tw.wm_geometry("+%d+%d" % (x, y))
            
            
            conteudo = self.conteudo
            colocarbreak = False
            cont = 0
            new_input = ""
            for i, letter in enumerate(conteudo):
                if i % 100 == 0 and i> 0:
                    colocarbreak = True
                if(colocarbreak):
                    new_input += letter
                    if(letter== ' '):                                    
                        new_input += '\n'
                        colocarbreak = False
                    else:
                        cont += 1
                        if(cont>10):
                           new_input += '\n'
                           cont = 0
                           colocarbreak = False
                else:
                    new_input += letter
            
            #label1 = tkinter.Label(self.tw, font=Font_tuple_Arial_10, text='', justify='left',
            #               background='#ededd3', relief='solid', borderwidth=0)
            #label1.grid(row=0, column=0, sticky='w', pady=0)
            label2 = tkinter.Label(self.tw, font=global_settings.Font_tuple_ArialBold_10, text="Nota:", justify='left',
                           background='#ededd3', relief='solid', borderwidth=0)
            label2.grid(row=0, column=0, sticky='w', pady=0)
            label3 = tkinter.Label(self.tw, font=global_settings.Font_tuple_Arial_10_italic, text=new_input, justify='left',
                           background='#ededd3', relief='solid', borderwidth=0)
            label3.grid(row=1, column=0, sticky='w', pady=0)
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
            
    def close(self, event=None):
        if self.tw:
            self.tw.destroy()
            
class FileTooltip(object):
    '''
    create a tooltip for a given widget
    '''
    def __init__(self, canv, widget, arquivo, pathdb, imagem=None):
        try:
            self.widget = widget
            self.tw = None
            self.arquivo = arquivo
            self.canv = canv
           
                
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
    
    def saveas(self, initialf, asbpathfile):
        path = (asksaveasfilename(initialfile=initialf))
        if(path!=None and path!=''):
            shutil.copy2(asbpathfile, path)        
    
    def menuAnnot(self, initialf, asbpathfile, event, pagina, x0, y0, x1, y1):
        self.menusaveas = tkinter.Menu(global_settings.root, tearoff=0)
        try:
            texto = 'Criar Anotação'
            if(self.conteudo!=''):
                texto = 'Editar Anotação'
            self.menusaveas.add_command(label="Salvar como", command= lambda : self.saveas(initialf, asbpathfile))
            #if(len(self.lista_executaveis)>0):
            #    self.listaexecs = tkinter.Menu(self.menusaveas, tearoff=0)
                #for executavel in self.lista_executaveis:   
                    
                #    self.listaexecs.add_command(label=executavel, command=partial(self.process_with,executavel, asbpathfile, pagina, x0, y0, x1, y1, event))
            #self.menusaveas.add_command(label=texto, font=Font_tuple_Arial_8, command=partial(self.process_with,'Escutar e Transcrever', asbpathfile, pagina, x0, y0, x1, y1, event))  
            #self.menusaveas.tk_popup(event.x_root, event.y_root) 
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        finally:
            self.menusaveas.grab_release()
            
    def enter(self, event=None):
        try:
            #print('1')
            if self.tw:
                self.tw.destroy()
            self.canv.config(cursor='hand2')
            x = y = 0
            x = event.x_root + 10
            y = event.y_root + 10
            # creates a toplevel window
            self.tw = tkinter.Toplevel(self.canv, background='#ededd3')
            self.tw.rowconfigure(2, weight=1)
            self.tw.columnconfigure(0, weight=1)
            # Leaves only the label and removes the app window
            self.tw.wm_overrideredirect(True)
            self.tw.wm_geometry("+%d+%d" % (x, y))
            label1 = tkinter.Label(self.tw, font=global_settings.Font_tuple_Arial_10, text=self.arquivo, justify='left',
                           background='#ededd3', relief='solid', borderwidth=0)
            label1.grid(row=0, column=0, sticky='w', pady=0)

        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
            
    def close(self, event=None):
        self.canv.config(cursor="fleur")
        if self.tw:
            self.tw.destroy()
            

    
            
class ObsTooltip(object):
    '''
    create a tooltip for a given widget
    '''
    def __init__(self, canv, widget, text='', conteudo='', obsobject=None):
        try:
            self.widget = widget
            self.text = text
            self.conteudo = conteudo
            self.tw = None
            self.canv = canv
            self.canv.tag_bind(widget, "<Enter>", self.enter)
            self.canv.tag_bind(widget, "<Leave>", self.close)
            self.canv.tag_bind(widget, "<Motion>",self.enter)
            self.canv.tag_bind(widget, "<3>",self.menuAnnot)
            self.obsobject = obsobject
                
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
            
    def menuAnnot(self):
        self.menusaveas = tkinter.Menu(global_settings.root, tearoff=0)
        try:
            texto = 'Criar Anotação'
            if(self.obsobject.conteudo!=''):
                texto = 'Editar Anotação'
            x0 = self.obsobject.p0x
            y0 = self.obsobject.p0y
            x1 = self.obsobject.p1x
            y1 = self.obsobject.p1x
            pagina = self.obsobject.paginainit
            abspathfile = self.obsobject.arquivo
            #self.menusaveas.add_command(label="Salvar como", command= lambda : self.saveas(initialf, asbpathfile))
            #if(len(self.lista_executaveis)>0):
            #    self.listaexecs = tkinter.Menu(self.menusaveas, tearoff=0)
                #for executavel in self.lista_executaveis:   
                    
                #    self.listaexecs.add_command(label=executavel, command=partial(self.process_with,executavel, asbpathfile, pagina, x0, y0, x1, y1, event))
            #self.menusaveas.add_command(label=texto, font=global_settings.Font_tuple_Arial_8, command=partial(self.process_with,'Escutar e Transcrever', asbpathfile, pagina, x0, y0, x1, y1, event))  
            #self.menusaveas.tk_popup(event.x_root, event.y_root) 
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        finally:
            self.menusaveas.grab_release()
            
    def enter(self, event=None):
        try:
            if self.tw:
                self.tw.destroy()
            x = y = 0
            x = event.x_root + 10
            y = event.y_root + 10
            # creates a toplevel window
            self.tw = tkinter.Toplevel(self.canv, background='#ededd3')
            self.tw.rowconfigure(2, weight=1)
            self.tw.columnconfigure(0, weight=1)
            # Leaves only the label and removes the app window
            self.tw.wm_overrideredirect(True)
            self.tw.wm_geometry("+%d+%d" % (x, y))
            
            
            
            label1 = tkinter.Label(self.tw, font=global_settings.Font_tuple_Arial_10, text=self.text, justify='left',
                           background='#ededd3', relief='solid', borderwidth=0)
            label1.grid(row=0, column=0, sticky='w', pady=0)
            if(self.conteudo!=''):
                conteudo = self.conteudo
                colocarbreak = False
                cont = 0
                new_input = ""
                for i, letter in enumerate(conteudo):
                    if i % 100 == 0 and i> 0:
                        colocarbreak = True
                    if(colocarbreak):
                        new_input += letter
                        if(letter== ' '):                                    
                            new_input += '\n'
                            colocarbreak = False
                        else:
                            cont += 1
                            if(cont>10):
                               new_input += '\n'
                               cont = 0
                               colocarbreak = False
                    else:
                        new_input += letter
                label2 = tkinter.Label(self.tw, font=global_settings.Font_tuple_ArialBold_10, text="Nota:", justify='left',
                               background='#ededd3', relief='solid', borderwidth=0)
                label2.grid(row=1, column=0, sticky='w', pady=0)
                label3 = tkinter.Label(self.tw, font=global_settings.Font_tuple_Arial_10_italic, text=new_input, justify='left',
                               background='#ededd3', relief='solid', borderwidth=0)
                label3.grid(row=2, column=0, sticky='w', pady=0)
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
            
    def close(self, event=None):
        if self.tw:
            self.tw.destroy()
    
class CreateToolTip(object):
    '''
    create a tooltip for a given widget
    '''
    def __init__(self, widget, text='widget info', istreeview=False, classe=''):
        try:
            self.istreeview = istreeview
            self.widget = widget
            self.text = text
            self.tw = None
            if(istreeview):
                self.widget.bind_class(classe,"<Motion>",self.enter)
                self.widget.bind_class(classe,"<Enter>",self.enter)
                self.widget.bind_class(classe,"<Leave>",self.close)
            else:
                self.widget.bind("<Enter>", self.enter)
                self.widget.bind("<Leave>", self.close)
                
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)
        

    def enter(self, event=None):
        try:
            if self.tw:
                self.tw.destroy()
            x = y = 0
            x = event.x
            y = event.y
            x += self.widget.winfo_rootx() + 25
            y += self.widget.winfo_rooty() + 20
            
            if(self.istreeview):
                iid = self.widget.identify_row(event.y)
                if(self.widget.tag_has('resultsearch', iid)):
                    texto = self.widget.item(iid, 'values')
                    if(len(texto)<=1):
                       return
                    self.tw = tkinter.Toplevel(self.widget)
                    self.tw.rowconfigure(0, weight=1)
                    self.tw.columnconfigure((0, 2), weight=1)
                    # Leaves only the label and removes the app window
                    self.tw.wm_overrideredirect(True)
                    self.tw.wm_geometry("+%d+%d" % (x, y))
                    label1 = tkinter.Label(self.tw, text=texto[0], justify='left', padx = 0,
                                   background='#ededd3', relief='solid', borderwidth=0,
                                   font=global_settings.Font_tuple_Arial_10)
                    label2 = tkinter.Label(self.tw, text=texto[1], justify='left',
                                   background='#ededd3', relief='solid', borderwidth=0,padx = 0,
                                   font=global_settings.Font_tuple_ArialBold_10)
                    label3 = tkinter.Label(self.tw, text=texto[2], justify='left',
                                   background='#ededd3', relief='solid', borderwidth=0,padx = 0,
                                   font=global_settings.Font_tuple_Arial_10)
                    label1.grid(row=0, column=0, sticky='ew', padx=0)
                    label2.grid(row=0, column=1, sticky='ew', padx=0)
                    label3.grid(row=0, column=2, sticky='ew', padx=0)
                    return
                else:
                    if(len(self.widget.item(iid, 'text'))<=1):
                       return
                    self.text = self.widget.item(iid, 'text')    
            self.tw = tkinter.Toplevel(self.widget)
            # Leaves only the label and removes the app window
            self.tw.wm_overrideredirect(True)
            self.tw.wm_geometry("+%d+%d" % (x, y))
            label = tkinter.Label(self.tw, font=global_settings.Font_tuple_Arial_10, text=self.text, justify='left',
                           background='#ededd3', relief='solid', borderwidth=1)
            label.pack(ipadx=1)
        except Exception as ex:
            utilities_general.printlogexception(ex=ex)

    def close(self, event=None):
        if self.tw:
            self.tw.destroy()

class PlaceholderEntry(ttk.Entry):
    def __init__(self, container, placeholder, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        self.placeholder = placeholder
        self.insert("0", self.placeholder)
        self.bind("<FocusIn>", self._clear_placeholder)
        self.bind("<FocusOut>", self._add_placeholder)

    def _clear_placeholder(self, e):
        if(super().get()=='Buscar...' or super().get()=="Aguarde, pesquisando..."):
            self.delete("0", "end")

    def _add_placeholder(self, e):
        if not self.get():
            self.insert("0", self.placeholder)
        
 
class querySqlWindow():
    def __init__(self,master, valor):
        self.value=None
        top=self.top=tkinter.Toplevel(master)
        self.top.rowconfigure((0,1,2,3,4), weight=1)
        self.top.columnconfigure((0,1), weight=1)
        #self.l=tkinter.Label(top,text="SELECT (........) FROM (......) WHERE (.....) MATCH <CONTINUAR ABAIXO>")
        #self.l.grid(row=0, column=0, columnspan=2, sticky='ns', pady=5)
        self.cattextvariable = tkinter.StringVar()
        self.cattextvariable.set(valor)
        #self.e=tkinter.Entry(top, width=100, textvariable=self.cattextvariable, justify='center')
        #self.e.focus_set()
        #self.e.grid(row=1, column=0, columnspan=2, sticky='nsew', pady=5)
        application_path = utilities_general.get_application_path()
        try:
            fts5tut = os.path.join(application_path,"fts4tutorial.png")
            self.imgtutorial = ImageTk.PhotoImage(file=fts5tut)
            self.tutorial = tkinter.Label(top, image=self.imgtutorial)
            self.tutorial.grid(row=2, column=0, sticky='nsew', columnspan=2, pady=5)
        except Exception as ex:
            fts5tut = os.path.join(os.getcwd(),"fts4tutorial.png")
            self.imgtutorial = ImageTk.PhotoImage(file=fts5tut)
            self.tutorial = tkinter.Label(top, image=self.imgtutorial)
            self.tutorial.grid(row=2, column=0, sticky='nsew', columnspan=2, pady=5)
       #self.aviso = tkinter.Label(top, font=Font_tuple_Arial_10, text='As aspas simples EXTERNAS não são necessárias, pois são adicionadas automaticamente!')
       # self.aviso.grid(row=3, column=0, sticky='ns', columnspan=2, pady=5)
        self.bok=tkinter.Button(top,text='Ok',command=self.cancel, font=global_settings.Font_tuple_Arial_10)
        self.bok.grid(row=4, column=0, columnspan=2, sticky='ns', pady=5)
        #self.bcancel=tkinter.Button(top,text='Cancelar',command=self.cancel)
        #self.bcancel.grid(row=4, column=1, sticky='ns', pady=5)
        
    def ok(self):
        self.value=self.cattextvariable.get()
        self.top.destroy()
    def cancel(self):
        self.value=None
        self.top.destroy()
class Search_Info_Window():
    def __init__(self, root, dist):
        self.window = tkinter.Toplevel(root)    
        self.window.rowconfigure((0,1), weight=1)
        self.window.columnconfigure(0, weight=1)
        self.label = tkinter.Label(self.window, font=global_settings.Font_tuple_Arial_10, text='Processando...')            
        self.label.grid(row=0, column=0, sticky= 'ns')
        self.progresssearch = ttk.Progressbar(self.window, mode='indeterminate', maximum = 1)
        self.progresssearch.grid(row=1, column=0, sticky='nsew', pady=5, padx=5)
        self.progresssearch.start()
        utilities_general.below_right(self.window, dist)
        self.window.overrideredirect(True)
        self.window.withdraw()
        #print("GO!")

class popupWindow(object):
    def __init__(self,master, valor):
        self.value=None
        top=self.top=tkinter.Toplevel(master)
        self.top.rowconfigure((0,1,2), weight=1)
        self.top.columnconfigure((0,1), weight=1)
        self.l=tkinter.Label(top,text="Nome da categoria:", font=global_settings.Font_tuple_Arial_10)
        self.l.grid(row=0, column=0, columnspan=2, sticky='ns', pady=20)
        self.cattextvariable = tkinter.StringVar()
        self.cattextvariable.set(valor)
        self.e=tkinter.Entry(top, width=100, textvariable=self.cattextvariable, justify='center')
        self.e.focus_set()
        self.e.grid(row=1, column=0, columnspan=2, sticky='nsew')
        self.bok=tkinter.Button(top,text='Ok',command=self.ok, font=global_settings.Font_tuple_Arial_10)
        self.bok.grid(row=2, column=0, sticky='ns', pady=20)
        self.bcancel=tkinter.Button(top,text='Cancelar',command=self.cancel, font=global_settings.Font_tuple_Arial_10)
        self.bcancel.grid(row=2, column=1, sticky='ns', pady=20)
        self.e.bind('<Return>',  lambda e: self.ok())
    def ok(self):
        self.value=self.cattextvariable.get()
        self.top.destroy()
    def cancel(self):
        self.value=None
        self.top.destroy()

class CustomFrame(tkinter.Frame):
    def __init__(self, master=None, cnf={}, **kw):
        super(CustomFrame, self).__init__(master=master, cnf={}, **kw)

class CustomCanvas(tkinter.Canvas):
    def __init__(self, master=None, scroll=None, **kw):
        super(CustomCanvas, self).__init__(master=master, **kw)        
    def yview(self, *args):
        """Query and change the vertical position of the view."""
        res = self.tk.call(self._w, 'yview', *args)
        atual = (self.program.vscrollbar.get()[0]*global_settings.infoLaudo[global_settings.pathpdfatual].len)
        restoemdeslocy = (atual%1.0)*global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh
        flooratual = math.floor(atual)
        #str(p)+global_settings.infoLaudo[relatorio].toc[0]
        tocx = utilities_general.locateToc(flooratual, global_settings.pathpdfatual,\
                                           p0y=restoemdeslocy, init=None, tocpdf=global_settings.infoLaudo[global_settings.pathpdfatual].toc)
        if(tocx!=None):
            tocxx = ''
            if(tocx[1]!=''):
                tocxx = tocx[0]+str(tocx[1])+str(tocx[2])
            tocxitem = str(global_settings.pathpdfatual)+tocxx
            if(self.program.treeviewEqs.item(self.program.treeviewEqs.parent(tocxitem), 'open')):
                self.program.treeviewEqs.selection_set(tocxitem)
            else:
                self.program.treeviewEqs.selection_set(self.program.treeviewEqs.parent(tocxitem))            
        atual = round((self.program.vscrollbar.get()[0]*global_settings.infoLaudo[global_settings.pathpdfatual].len))
        self.program.pagVar.set(atual+1)
        global_settings.infoLaudo[global_settings.pathpdfatual].something_changed = True
        if not args:
            return self._getdoubles(res)        
        
    def yview_moveto(self, fraction, qlpdf=None):
        """Adjusts the view in the window so that FRACTION of the
        total height of the canvas is off-screen to the top."""
        self.tk.call(self._w, 'yview', 'moveto', fraction)
        
        atual = (self.program.vscrollbar.get()[0]*global_settings.infoLaudo[global_settings.pathpdfatual].len)
        restoemdeslocy = (atual%1.0)*global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh
        flooratual = math.floor(atual)
        #str(p)+global_settings.infoLaudo[relatorio].toc[0]
        tocx = utilities_general.locateToc(flooratual, global_settings.pathpdfatual,\
                                           p0y=restoemdeslocy, init=None, tocpdf=global_settings.infoLaudo[global_settings.pathpdfatual].toc)
        if(tocx!=None):
            tocxx = ''
            if(tocx[1]!=''):
                tocxx = tocx[0]+str(tocx[1])+str(tocx[2])
            tocxitem = str(global_settings.pathpdfatual)+tocxx
            if(self.program.treeviewEqs.item(self.program.treeviewEqs.parent(tocxitem), 'open')):
                self.program.treeviewEqs.selection_set(tocxitem)
            else:
                self.program.treeviewEqs.selection_set(self.program.treeviewEqs.parent(tocxitem))                            
        atual = round((self.program.vscrollbar.get()[0]*global_settings.infoLaudo[global_settings.pathpdfatual].len))
        global_settings.infoLaudo[global_settings.pathpdfatual].something_changed = True
        self.program.pagVar.set(atual+1)
        #root.update_idletasks()
        
        

    def yview_scroll(self, number, what):
        """Shift the y-view according to NUMBER which is measured in
        "units" or "pages" (WHAT)."""
        self.tk.call(self._w, 'yview', 'scroll', number, what)
        
        atual = (self.program.vscrollbar.get()[0]*global_settings.infoLaudo[global_settings.pathpdfatual].len)
        restoemdeslocy = (atual%1.0)*global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh
        flooratual = math.floor(atual)
        #str(p)+global_settings.infoLaudo[relatorio].toc[0]
        tocx = utilities_general.locateToc(flooratual, global_settings.pathpdfatual,\
                                           p0y=restoemdeslocy, init=None, tocpdf=global_settings.infoLaudo[global_settings.pathpdfatual].toc)
        if(tocx!=None):
            tocxx = ''
            if(tocx[1]!=''):
                tocxx = tocx[0]+str(tocx[1])+str(tocx[2])
            tocxitem = str(global_settings.pathpdfatual)+tocxx
            if(self.program.treeviewEqs.item(self.program.treeviewEqs.parent(tocxitem), 'open')):
                self.program.treeviewEqs.selection_set(tocxitem)
            else:
                self.program.treeviewEqs.selection_set(self.program.treeviewEqs.parent(tocxitem))                         
        atual = round((self.program.vscrollbar.get()[0]*global_settings.infoLaudo[global_settings.pathpdfatual].len))
        global_settings.infoLaudo[global_settings.pathpdfatual].something_changed = True
        self.program.pagVar.set(atual+1)
        #if(str(atual+1)!=self.program.pagVar.get()):
        #    self.program.pagVar.set(str(atual+1)) 
        #root.update_idletasks()
    def scan_mark(self, x, y):
        """Remember the current X, Y coordinates."""
        self.tk.call(self._w, 'scan', 'mark', x, y)
    def scan_dragto(self, x, y, gain=10):
        """Adjust the view of the canvas to GAIN times the
        difference between X and Y and the coordinates given in
        scan_mark."""
        self.tk.call(self._w, 'scan', 'dragto', x, y, gain)
        
        atual = (self.program.vscrollbar.get()[0]*global_settings.infoLaudo[global_settings.pathpdfatual].len)
        restoemdeslocy = (atual%1.0)*global_settings.infoLaudo[global_settings.pathpdfatual].pixorgh
        flooratual = math.floor(atual)
        #str(p)+global_settings.infoLaudo[relatorio].toc[0]
        tocx = utilities_general.locateToc(flooratual, global_settings.pathpdfatual,\
                                           p0y=restoemdeslocy, init=None, tocpdf=global_settings.infoLaudo[global_settings.pathpdfatual].toc)
        if(tocx!=None):
            tocxx = ''
            if(tocx[1]!=''):
                tocxx = tocx[0]+str(tocx[1])+str(tocx[2])
            tocxitem = str(global_settings.pathpdfatual)+tocxx
            if(self.program.treeviewEqs.item(self.program.treeviewEqs.parent(tocxitem), 'open')):
                self.program.treeviewEqs.selection_set(tocxitem)
            else:
                self.program.treeviewEqs.selection_set(self.program.treeviewEqs.parent(tocxitem)) 
        atual = round((self.program.vscrollbar.get()[0]*global_settings.infoLaudo[global_settings.pathpdfatual].len))
        global_settings.infoLaudo[global_settings.pathpdfatual].something_changed = True
        self.program.pagVar.set(atual+1)  
        #if(str(atual+1)!=self.program.pagVar.get()):
        #    self.program.pagVar.set(str(atual+1))
        #root.update_idletasks()

