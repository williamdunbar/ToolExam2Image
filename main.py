import os
import subprocess
from time import sleep
from run_macro import RunMacro
from readdocx import Read_latex_From_Word_To_Tex
from readlatex import Read_Latex
from latex2image import Latex2Image


def Excute_Macro(path_docx):
    # macro_run = "W2T_Tool"
    name_docx_file = (path_docx.split(sep="\\")[-1])[:-5] or "Version2"

    if(os.path.exists("./Macro.docx") == True):
        os.remove("./Macro.docx")
    error, detail = RunMacro(path_docx)
    print(">>>ĐÃ CHẠY MACRO XONG<<<")

    return name_docx_file, error, detail

 
def Excute(path_docx,excerpt_detect):
    name_docx_file, error, detail = Excute_Macro(path_docx)
    if(error == "" and detail == ""):
        error1, detail = Read_latex_From_Word_To_Tex(excerpt_detect=excerpt_detect)
        if(error1 == ""):
            List_Excercise, error2, detail = Read_Latex()
            if(error2 == ""):
                error3, detail =Latex2Image(folder_data=name_docx_file, List_Excercise= List_Excercise)
                if(error3 == ""):
                    print("------------------------------------")
                    print("Convert done !!!")
                    return 1, name_docx_file, "", ""
                else:
                    print("------------------------------------")
                    print("Error3: "+ error3 + " at " + detail )
                    return -1, name_docx_file, error3, detail
            else:
                print("------------------------------------")
                print("Error2: "+ error2 + " at " + detail )
                return -1, name_docx_file, error2, detail
        else:
            print("------------------------------------")
            print("Error1: "+ error1 + " at " + detail )
            return -1, name_docx_file, error1, detail
    else:
        print("------------------------------------")
        print("Error0: "+ error + " at " + detail )
        return -1, name_docx_file, error, detail
    
    