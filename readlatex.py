from distutils.log import error
from operator import index
import re
from unicodedata import name


class Excercise:
    def __init__(self, question, A,B,C,D, reference,excerpt):
        self.question = question
        self.A = A
        self.B = B
        self.C = C
        self.D = D
        self.reference = reference
        self.excerpt = excerpt 



def Read_Latex():
    ''' Tạo một list chứa các object Excercise'''
    List_Excercise = []
    index  = 0
    try:
        temp_data = open("./Official_word_2_tex.tex", encoding="utf8").read()
        first_excerpt = (re.findall(r'\\doantrich{(.*?)}\n\\end{ex}', temp_data, re.S) or [""])[0] # đoạn trích đầu tiên (Không nằm trong excercise)
        excercises = re.findall(r'\\begin{ex}(.*?)\\end{ex}', temp_data, re.S)
        # essay_excercises = re.findall(r'\\begin{bt}(.*?)\\end{bt}', temp_data, re.S)

        if first_excerpt != "":
            List_Excercise.append(Excercise(question="",reference="",A="",B="",C= "",D="",excerpt=first_excerpt))

        for excercise in excercises:
                index = index + 1
                '''
                Chia khối bài tập đã thu thập được thành các phần: question, A, B, C, D, answer
                '''
                excercise = "\\begin{ex}"+ excercise + "\\end{ex}"

                ''' Vì findall trả về 1 list mà list này của mình chỉ có 1 phần tử nên dùng question[0] '''
                temp_answer = (re.findall(r'\\choice(.*?)\\loigiai', excercise, re.S) or [""])[0]
                temp_excerpt = (re.findall(r'\\doantrich{(.*?)}\n\\end{ex}', excercise, re.S) or [""])[0] # đoạn trích
                if(temp_excerpt != ""):
                    excerpt = temp_excerpt.split("\n",2)[2]
                    temp_reference = (re.findall(r'\\loigiai{(.*?)}\n\\doantrich', excercise, re.S) or [""])[0]
                else:
                    excerpt = temp_excerpt
                    temp_reference = (re.findall(r'\\loigiai{(.*?)}\n\\end{ex}', excercise, re.S) or [""])[0]
                
                    # print(temp_answer)
                if (temp_answer != ""):
                    question = (re.findall(r'\\begin{ex}\n(.*?)\n\\choice', excercise, re.S) or [""])[0]
                    multi_choice = (re.findall(r'\{\@(.*?)\@\}\n', temp_answer, re.S)) #re.S giúp cho regex đọc được các nội dung có \n
                    # print(multi_choice)
                    A = multi_choice[0]
                    B = multi_choice[1]
                    C = multi_choice[2]
                    D = multi_choice[3]
                    # reference = temp_reference.split("\n",2)[2]
                    reference = temp_reference
                else:
                    question = (re.findall(r'\\begin{ex}\n(.*?)\n\\choice', excercise, re.S) or [""])[0]
                    A,B,C,D = ["","","",""]
                    reference = temp_reference
                    # print (temp_reference)
                List_Excercise.append(Excercise(question,A,B,C,D,reference,excerpt))
    except Exception as err:
        print("Error ReadLatex: "+str(err))
        if(index == 0):
            detail = "Lỗi khi đọc file Official_word_2_tex.tex" 
        else:
            detail = "Câu: "+ str(index)
        return List_Excercise, str(err), detail
    print(">>>ĐÃ PHÂN TÍCH FILE .TEX XONG<<< !!!")
    print("-------------------------")
    return List_Excercise, "", ""

