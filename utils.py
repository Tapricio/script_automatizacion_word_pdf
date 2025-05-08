import os
import re

def mkdirFolders (path): 
    count = 1
    while os.path.exists(path):
        path = f"{path}({count})"
        count +=1

    pdf_folder = os.path.join(path, "test pdf")
    word_folder = os.path.join(path, "test word")

    os.makedirs(pdf_folder, exist_ok=True)
    os.makedirs(word_folder, exist_ok=True)

#def desdentado():

def limpiar_nombre_archivo(nombre):
    return re.sub(r'[\/:*?"<>|]', '', nombre)

def docx_replace_regex(doc_obj, regex , replace):

    for p in doc_obj.paragraphs:
        if regex.search(p.text):
            inline = p.runs
            # Loop added to work with runs (strings with same style)
            for i in range(len(inline)):
                if regex.search(inline[i].text):
                    text = regex.sub(replace, inline[i].text)
                    inline[i].text = text

    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                docx_replace_regex(cell, regex , replace)




def docx_replace_multiple_regex(doc_obj, replacements):
    # Compila todos los regex
    compiled_replacements = {re.compile(k): v for k, v in replacements.items()}

    def replace_in_paragraphs(paragraphs):
        for p in paragraphs:
            for regex, replace in compiled_replacements.items():
                if regex.search(p.text):
                    inline = p.runs
                    for i in range(len(inline)):
                        if regex.search(inline[i].text):
                            inline[i].text = regex.sub(replace, inline[i].text)

    def replace_in_tables(tables):
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    replace_in_paragraphs(cell.paragraphs)

    replace_in_paragraphs(doc_obj.paragraphs)
    replace_in_tables(doc_obj.tables)


def rut_formateado (data):
    rut_sin_dv = data[:-2]
    dv = data[-1]
    rut = "{:,}".format(int(rut_sin_dv)).replace(",", ".") + "-" + dv
    return rut.upper() 