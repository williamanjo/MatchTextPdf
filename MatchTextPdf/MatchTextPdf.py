import fitz
from PyPDF2 import PdfFileReader, PdfFileWriter
from tkinter import filedialog , messagebox
from tkinter import *
import os
import pytesseract
from PIL import Image
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files (x86)\Tesseract- OCR\tesseract.exe'

all_files= []
root = Tk()
root.withdraw()

while(True):
    Directory = filedialog.askdirectory(title ="Selecione a Pasta Que Contêm Os Arquivos (*.pdf)")
    if(Directory):
        if(messagebox.askokcancel(title="Pasta Que Contêm Os Arquivos (*.pdf)", message=Directory)):
                break
        else:
            if(messagebox.askquestion (title="Atenção", message="Fechar o Programa ? " ,icon = 'warning') == "yes"):
                exit(0)
    else:
        exit(0)


for dirpath, dirnames, filenames in os.walk(Directory):
    for filename in [f for f in filenames if (f.endswith(".pdf") or f.endswith(".PDF"))]:
        all_files.append(Directory + '/' + filename)

for filename in all_files:

    pdfDANF = PdfFileWriter()
    pdfNFS = PdfFileWriter()
    pdfNull = PdfFileWriter()
    pages = []
    pagesNFS = []
    pagesNull = []
    file_base_name = filename.replace('.pdf', '')
    search_term = "DANF"
    search_term2 = "NFS-e"
    pdf_document = fitz.open(filename);
    pdf = PdfFileReader(filename)

    for current_page in range(len(pdf_document)):
        page = pdf_document.loadPage(current_page)
        if page.searchFor(search_term):
        
            print("%s : %s found on page %i" % (filename,search_term, current_page))
            pages.append(current_page);
        else:
            if page.searchFor(search_term2):
                print("%s : %s found on page %i" % (filename, search_term2, current_page))
                pagesNFS.append(current_page);
            else:
                try: 
                    os.makedirs(Directory + "/Imgs/"+ os.path.basename(file_base_name))
                except:
                    break;
                save_path_img = (Directory + "/Imgs/"+ os.path.basename(file_base_name))
                rotate = int(0)
                zoom_x = 2
                zoom_y = 2

                trans = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate)
                pm = page.getPixmap(matrix=trans, alpha=False)
                pm.writePNG(save_path_img + '/%s.png' % current_page)
                image = Image.open(save_path_img +'/' + str(current_page) + '.png')

                content = pytesseract.image_to_string(image, lang="eng")
                with open(save_path_img +'/' + str(current_page) + '.txt', "w") as f:
                    f.write(content)
                
                if content.find("NF-e") != -1:
                    print("%s : %s found on Image page %i" % (filename,search_term, current_page))
                    pages.append(current_page);
                else:
                    if content.find(search_term2) != -1:
                        print("%s : %s found on Image page %i" % (filename,"NFS-e", current_page))
                        pagesNFS.append(current_page);
                    else:
                        pagesNull.append(current_page);

    # Criando documento DANF
    for page_num in pages:
        pdfDANF.addPage(pdf.getPage(page_num))

    if(len(pages) > 0):
       with open('{0}_{1}.pdf'.format(file_base_name,search_term), 'wb') as f:
        pdfDANF.write(f)
        f.close()

    # Criando documento NFS-e

    for page_num in pagesNFS:
        pdfNFS.addPage(pdf.getPage(page_num))

    if(len(pagesNFS) > 0):
       with open('{0}_{1}.pdf'.format(file_base_name,"NFS-e"), 'wb') as f:
        pdfNFS.write(f)
        f.close()
    # Criando Documento Com as paginas que não foram identificadas
    for page_num in pagesNull:
        pdfNull.addPage(pdf.getPage(page_num))

    if(len(pagesNull) > 0):
       with open('{0}_{1}.pdf'.format(file_base_name,"Não Identificadas"), 'wb') as f:
        pdfNull.write(f)
        f.close()
    pdf_document.close()