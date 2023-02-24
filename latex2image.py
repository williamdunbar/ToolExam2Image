import math
import os
import shutil
from os import path
from tabnanny import verbose
from time import sleep
from turtle import width
from pdf2image import convert_from_path
from pdflatex import PDFLaTeX
from PIL import Image
import subprocess

from torch import det

ratio = 1160/840
# folder_data = name_file

def add_margin(pil_img, right):
    top = 0
    left = 0
    bottom = 0
    color = (255,255,255)
    width, height = pil_img.size
    new_width = width + right + left
    new_height = height + top + bottom
    result = Image.new(pil_img.mode, (new_width, new_height), color)
    result.paste(pil_img, (left, top))
    return result


def crop(file):
    '''Định dạng kích thước ảnh'''
    # img = Image.open(file)
    # width, height = img.size
    # print(str(width)+" "+ str(height))
    # area = (0,0, width, height) #im.crop((left, top, right, bottom))
    # cropped_img = img.crop(area)
    # cropped_img.save(file)

    img = Image.open(file)
    width, height = img.size
    new_width = math.trunc(width * 840/1160)
    new_height= math.trunc(height * 840/1160)
    new_img = img.resize((new_width, new_height))
    # print(new_width, new_height)
    new_img.save(file)

def question_crop(file):
    # Định dạng kích thước ảnh
    img = Image.open(file)
    width_0, height_0 = img.size
    new_image = add_margin(img,1160-width_0)
    new_image.save(file)

def save_images(images_name,pdf_path,images_path,type):
# Store Pdf with convert_from_path function
    images = convert_from_path(pdf_path)
    if len(images_name)==0:
        print("names is empty")
        return
    for img in images:
        img.save(images_path+"/"+images_name +".jpg", 'JPEG')
        if (type == 0 or type == 1):
            question_crop(images_path+"/"+images_name +".jpg")
        else:
            crop(images_path+"/"+images_name +".jpg")
    # print("Successfully converted")

def create_image_from_latex(image_name,latex,folder_data, folder_exercise, type):
    folder_exercise = os.path.join(folder_data,folder_exercise)
    if(path.exists(folder_exercise) == False):
        os.mkdir(folder_exercise)

    print(latex)
    f=open("Temp.tex","w+", encoding= "utf8")
    f.write("\\documentclass[preview,border=1pt]{standalone}\n\\usepackage{amsmath,amssymb}\n\\usepackage{charter}\n\\usepackage{indentfirst}\n\\usepackage[utf8]{vietnam}\n\\usepackage{longtable}\n\\usepackage{multirow}\n\\usepackage{mathrsfs}\n\\usepackage{tabvar}\n\\usepackage{ifpdf}\n\\usepackage{wasysym}\n\\begin{document}\n")
    f.write(latex+"\n")
    f.write(r"\end{document}")
    f.close()

    subprocess.run(['pdflatex', '-interaction=batchmode', './Temp.tex'], shell= False)

    # pdfl = PDFLaTeX.from_texfile('./Temp.tex')
    # print(1)
    # pdf, log, completed_process = pdfl.create_pdf(keep_pdf_file=True, keep_log_file=True)
    # print(2)

    save_images(image_name,"Temp.pdf",folder_exercise, type)
    # os.remove("temp.pdf")

def Latex2Image(folder_data, List_Excercise ):
    if(path.exists(folder_data) == True):
        shutil.rmtree(folder_data)
        sleep(2)
        os.mkdir(folder_data)
    else:
        os.mkdir(folder_data)
    i = 1
    for excercise in List_Excercise:
        try:
            print(i)
            #Lưu ảnh đoạn trích
            if(excercise.excerpt != ""):
                k = i+1
                excerpt_name = "excerpt_"+ str(k)
                create_image_from_latex(image_name=str(excerpt_name),latex=excercise.excerpt, folder_exercise=str(k),type=0, folder_data=folder_data)
                # print("----------------------------------")
                # print(excercise.excerpt)
            #Lưu ảnh câu hỏi
            if(excercise.question != ""):
                question_name = "question_"+ str(i)
                create_image_from_latex(image_name=str(question_name),latex=excercise.question, folder_exercise=str(i),type=0, folder_data=folder_data)
                # print("----------------------------------")
                # print(excercise.question)
                #Lưu ảnh lời giải
            if(excercise.reference != ""):
                reference_name = "reference_"+ str(i)
                create_image_from_latex(image_name=str(reference_name),latex=excercise.reference, folder_exercise=str(i),type=1, folder_data=folder_data)
                #Lưu ảnh đáp án
            if(excercise.A != ""):
                if "\True" in excercise.A:
                    answer_name = "answer_"+ str(i) + "_A_True"
                    excercise.A = excercise.A.replace("\True","")
                else:
                    answer_name = "answer_"+ str(i) + "_A"
                create_image_from_latex(image_name=str(answer_name),latex=excercise.A, folder_exercise=str(i),type=2, folder_data=folder_data)

                if "\True" in excercise.B:
                    excercise.B = excercise.B.replace("\True","")
                    answer_name = "answer_"+ str(i) + "_B_True"
                else:
                    answer_name = "answer_"+ str(i) + "_B"
                create_image_from_latex(image_name=str(answer_name),latex=excercise.B, folder_exercise=str(i),type=2, folder_data=folder_data)

                if "\True" in excercise.C:
                    excercise.C = excercise.C.replace("\True","")
                    answer_name = "answer_"+ str(i) + "_C_True"
                else:
                    answer_name = "answer_"+ str(i) + "_C"
                create_image_from_latex(image_name=str(answer_name),latex=excercise.C, folder_exercise=str(i),type=2, folder_data=folder_data)

                if "\True" in excercise.D:
                    excercise.D = excercise.D.replace("\True","")
                    answer_name = "answer_"+ str(i) + "_D_True"
                else:
                    answer_name = "answer_"+ str(i) + "_D"
                create_image_from_latex(image_name=str(answer_name),latex=excercise.D, folder_exercise=str(i),type=2, folder_data=folder_data)
            i += 1
        except Exception as err:
            detail = "Câu " + str(i)
            return str(err), detail
    print(">>>ĐÃ CHUYỂN LATEX THÀNH ẢNH XONG<<<")
    print("-------------------------")
    return "", ""