from fileinput import filename
import os
import re
from turtle import width
import zipfile
from time import sleep
from os import path
import docx
import xml.etree.ElementTree as ET
from PIL import Image

path_docx_file = "./Macro.docx" 


def hasImage(par):
    """get all of the images in a paragraph 
    :param par: a paragraph object from docx
    :return: a list of r:embed 
    """
    ids = []
    root = ET.fromstring(par._p.xml)
    namespace = {
             'a':"http://schemas.openxmlformats.org/drawingml/2006/main", \
             'r':"http://schemas.openxmlformats.org/officeDocument/2006/relationships", \
             'wp':"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"}

    inlines = root.findall('.//wp:inline',namespace)
    for inline in inlines:
        imgs = inline.findall('.//a:blip', namespace)
        for img in imgs:     
            id = img.attrib['{{{0}}}embed'.format(namespace['r'])]
        ids.append(id)
    return ids

def num_sort(test_string):
    return list(map(int, re.findall(r'\d+', test_string)))[0]

def Extract_Image():
    imagenames = []
    i = 0

    #Xóa ảnh cũ trong "Picture/word/media/"
    dir = 'Picture/word/media/'
    for f in os.listdir(dir):
        os.remove(os.path.join(dir, f))

    #Dùng thư viện zipfile extract ảnh 
    z = zipfile.ZipFile(path_docx_file)

    for file in z.filelist:
        if file.filename.startswith('word/media/') :
            imagename = str(file.filename).split('/',2)[2]
            z.extract(path="./Picture", member= file.filename)
            #Xử lỹ các ảnh của đuôi emf
            if imagename.endswith("emf") :
                i = i+1
                newImageNames = "table"+ str(i) + ".jpg"
                Image.open("Picture/word/media/"+imagename).save("Picture/word/media/"+newImageNames)
                imagename = newImageNames

            imagenames.append(imagename)
    imagenames.sort(reverse=False,key=num_sort) 
    # print("-----------------")
    # print(imagenames)
    return imagenames

def Fix_Latex(tex_input, excerpt_detect, paragraph_input):
    if(tex_input == "lời giải" or tex_input == "Lời giải"):
        print("Tồn tại 1 lỗi: \"lời giải\" đã sửa")
        paragraph_input[-1] = paragraph_input[-1].split("\\\\")[0]+"@}\n"
        tex_input = "\loigiai{"
    if(tex_input == excerpt_detect):
        print("Đã sửa " + tex_input +" thành \\doantrich{")
        tex_input = "\doantrich{"
    if(tex_input != "\\begin{ex}" and tex_input != "\\end{ex}" and tex_input != "\\choice" and tex_input != "\\loigiai{" and tex_input != "\\doantrich{" and tex_input !="}" and tex_input !="" and ('@}' not in tex_input)):
        tex_input = tex_input + "\\\\"
    tex_input = tex_input + "\n"
    return tex_input, paragraph_input

def Read_latex_From_Word_To_Tex(excerpt_detect):
    imagenames = Extract_Image()
    try:
        doc = docx.Document(path_docx_file)
    except Exception as err:
        return str(err), "Không tìm thấy file"

    i = 0
    if(path.exists("./Official_word_2_tex.tex") == True):
        os.remove("./Official_word_2_tex.tex")
    with open("./Official_word_2_tex.tex",mode= "a", encoding="utf8") as f:
        paragraph_input=[] 
        for para in doc.paragraphs:
            try:
                isImage = hasImage(para)
                # print("-------------------------")
                if(isImage == []):
                    tex_input = str(para.text)
                    tex_input, paragraph_input = Fix_Latex(tex_input = tex_input, excerpt_detect=excerpt_detect, paragraph_input = paragraph_input)
                else:
                    tex_input = "\\begin{center}\n\\includegraphics[width=250px]{Picture/word/media/"+ imagenames[i] + "}\n\\end{center}\n"
                    i += 1
                paragraph_input.append(tex_input)
            except Exception as err:
                print(para.text)
                return str(err), str(para.text)
        for element in paragraph_input:
            f.write(element)
        f.write("}\n\\end{ex}")    
    print("Số hình ảnh trong file: "+ str(i))
    print(">>> ĐÃ TÁCH ẢNH XONG <<<")
    print("-------------------------")
    return "", ""
    


if __name__ == "__main__":
    Read_latex_From_Word_To_Tex("Đoạn trích")



    