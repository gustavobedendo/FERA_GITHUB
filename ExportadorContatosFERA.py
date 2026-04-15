import fitz, traceback, re, indexador_fera
from utilities_general import connectDB
import tkinter as tk
from tkinter import filedialog
import sys, os
def get_rep(relatorio):
    with fitz.open(relatorio) as doc:
        novotexto = doc.get_page_text(0)
        #print(novotexto)
        anexo_ao_laudo = "Anexo\sao\sLaudo\sNº\s([^\n]+)"
        anexo_rep_compiled = re.compile(anexo_ao_laudo)
        results = anexo_rep_compiled.findall(novotexto)[0]
        #print(results)
        return results
    
    

def extractPageText(doc, posicao_init, posicao_fim, compilers=[]):
    hits = {}
    height = None
    for compile_name in compilers:
        hits[compile_name] = []
    pinicial = int(posicao_init[0])
    yinit = float(posicao_init[2])
    pfinal = int(posicao_fim[0])
    yfim = float(posicao_fim[2])
    for pagina in range(pinicial, pfinal+1,1):
        if(pagina == 9921):
            print("teste")
        if(height==None):
            width, height = get_pixmap_height(doc[pagina])
        novotexto = ""
        if(pagina==pinicial and pagina==pfinal):
            novotexto = doc[pagina].get_textbox(fitz.Rect(0, height-yinit, width, height-yfim))
        elif(pagina==pinicial):
            novotexto = doc[pagina].get_textbox(fitz.Rect(0, height-yinit, width, height))
        elif(pagina==pfinal):
            novotexto = doc[pagina].get_textbox(fitz.Rect(0, 0, width, height-yfim))
        else:
            novotexto = doc[pagina].get_text()
        novotexto.replace("\\\n", "\n")
        
        #print(novotexto)
        for compile in compilers:
            for hit in compilers[compile].findall(novotexto):
                hits[compile].append((hit[0].replace("\n", ""), hit[1].replace("\n", "")))
    return hits



def get_pdfs(db_file):
    select_all_pdfs = '''SELECT  P.rel_path_pdf, P.id_pdf, T.toc_unit FROM Anexo_Eletronico_Pdfs P inner 
    join Anexo_Eletronico_Tocs T on P.id_pdf = T.id_pdf where toc_unit like "%Contatos%"
    '''
    sqliteconn = connectDB(db_file)
    #print(2)
    cursor = sqliteconn.cursor()
    cursor.execute(select_all_pdfs)
    relats = cursor.fetchall()
    relatorios = []
    for relat, id_pdf, toc in relats:
        relatorios.append((os.path.normpath(os.path.join(os.path.dirname(db_file), relat)), id_pdf, toc))
        #print(relat)
    return relatorios

def open_file_dialog():
    # Create a file dialog that only allows .db files
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(
        title="Select a database file",
        filetypes=[("Database Files", "*.db"), ("All Files", "*.*")]
    )
    if file_path:
        print(f"Selected database file: {file_path}")
        return file_path
    else:
        print("No file selected")
        sys.exit(1)
        
def process_relatorios(relatorios):
    hits = {}
    for relat, id_pdf, toc in relatorios:
        subsecao = toc.split(" ")[0]
        subsecaofim = str(round(float(subsecao)+0.1, 1))
        doc = fitz.open(relat)
        #tocs = doc.get_toc(simple=False)
        
        nameddests = indexador_fera.grabNamedDestinations(doc)    
        posicao_init = (nameddests[f'subsection.{subsecao}'])
        posicao_fim = (nameddests[f'subsection.{subsecaofim}'])
        regex_telefone = "Nome:\s([^\n]+)\nTelefone:\s([0-9]+)"
        regex_wa = "Nome:\s([^\n]+)\nID de usuário:\s([0-9]+)@s.whatsapp.net"
        #seminstanciacomnumero = "Nome:\s([^\n]+)\nAplicativo:\s([A-Za-z]*WhatsApp[^\n]*)\nConta:\s([^@]+)"
        regex_wacompile = re.compile(regex_wa)
        #seminstanciacomnumerocompile = re.compile(seminstanciacomnumero)
        telcompile = re.compile(regex_telefone)
        hits[relat+"->"+toc] = extractPageText(doc, posicao_init, posicao_fim, {'telefone':telcompile, 'whatsaapp':regex_wacompile})
    return hits

def get_pixmap_height(pagina):
    return pagina.get_pixmap().width, pagina.get_pixmap().height

try:

    db_File = open_file_dialog()
    relats = get_pdfs(db_File)
    print(relats)
    #relats = [(r'D:\Report_2117.pdf', 'x', "1.4 Contatos")]
    #db_File = r"D:\teste.db"
    rep = get_rep(relats[0][0]).replace(".", "")
    #print(relats)
    hits = process_relatorios(relats)

    with open(os.path.join(os.path.dirname(db_File), 'contatos.csv'), 'w', encoding='utf-8', errors='ignore') as arquivo:
        arquivo.write(f"Número REP,Nome,Telefone,Relatório\n")
        for relat in hits:
            for hit in hits[relat]['telefone']:
                arquivo.write(f"{rep},{hit[0]},{hit[1]},{os.path.basename(relat)}\n")
            for hit in hits[relat]['whatsaapp']:
                arquivo.write(f"{rep},{hit[0]},{hit[1]},{os.path.basename(relat)}\n")
    input("OK: Enter para fechar")
except:
    traceback.print_exc()
    input("Erro de execução: Enter para fechar")
#print(hits)

