import fitz  # this is pymupd, pip3 install PyMuPDF
import traceback
import math
# WARNING, this is a bad code, please use it knowing it may break easely
# Author: nah, I'm joking, nobody wants to own this shit XD


def get_rows():
    None
def get_cells_from_rows():
    None

def get_text_blocks(doc, pagina):
    flags = 2+4+64
    page = doc[pagina]
    dictx = page.get_text("rawdict", flags=flags) 
    blocos = dictx['blocks']
    indiceblocos = 0
    while indiceblocos < len(blocos):
        pontosBlocos = blocos[indiceblocos]['bbox']
        if('lines' not in blocos[indiceblocos]):
            indiceblocos += 1
            continue
        if(pontosBlocos[1]<115 or pontosBlocos[1]>810):
            indiceblocos += 1
            continue
        for line in blocos[indiceblocos]['lines']:
            pontosLine = line['bbox']
            
            if(pontosLine[2]<=150):
                print("primeiro")
            elif(pontosLine[2]<=265):
                print("segundo")
            elif(pontosLine[2]<=354):
                print("terceiro")
            elif(pontosLine[2]<=570):
                print("ultimo")
            print(pontosLine)
            print(page.get_textbox(fitz.Rect(pontosLine), textpage=None).replace("\n", " "))

        indiceblocos += 1
              

if __name__ == "__main__":
    pdf_path = r"D:\35584-17\35584-17-Anexo\Anexo\Eq01\Anexo_35584-17_Eq01.pdf"
    doc = fitz.open(pdf_path)
    try:
        get_text_blocks(doc, 8900)
    except:
        traceback.print_exc()
    finally:
        doc.close()
    #print(rows)