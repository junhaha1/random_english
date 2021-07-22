import tkinter.ttk as ttk
import tkinter.messagebox as msgbox
from random import sample
from tkinter import *
from docx import Document

root = Tk()
root.title("English Word Exam")
root.geometry("600x300+100+300")

def odd_change_even(num):
    if num % 2 == 0:
        return num
    else:
        return num + 1


def make_word(lst, name):

    doc = Document()
    even = odd_change_even(len(lst))
    row = int(even / 2)
    idx = 0

    table = doc.add_table(rows = row, cols = 4, style = "Table Grid")

    for i in range(row):
        for j in range(4):
            if j % 2 != 0:
                table.rows[i].cells[j].text = ""
            else:
                table.rows[i].cells[j].text = lst[idx]
                idx += 1
    doc.save("C:\\Users\\junha\\OneDrive\\바탕 화면\\파이썬\\준희 영어\\{}.docx".format(name))

def make_exam():
    
    if len(exam_name.get()) == 0:
        msgbox.showwarning("경고", "이름을 입력해주세요")
        return 

    if int(cmb_start.get()) > int(cmb_end.get()):   
        strnum = int(cmb_end.get())
        endnum = int(cmb_start.get())
    else:
        strnum = int(cmb_start.get())
        endnum = int(cmb_end.get())

    test = exam_name.get()
    run_test(test, strnum, endnum)
    
    msgbox.showinfo("알림", "%s가 만들어졌습니다." %test)

def run_test(name, s1, s2):

    outeng = []
    selc_num = int(cmb_num.get())
    a = sample(range(0, 30), selc_num)

    s_path = "C:\\Users\\junha\\OneDrive\\바탕 화면\\파이썬\\준희 영어\\"
    f_path = "C:\\Users\\junha\\OneDrive\\바탕 화면\\파이썬\\준희 영어\\준희 영단어 랜덤\\"

    for j in range(s1, s2 + 1):
        with open("%sword%d.txt" % (f_path, j), 'r') as file:
            lines = file.readlines()
            for i in a:
                outeng.append(lines[i].strip('\n'))

    with open("%s%s.txt" % (s_path, name), 'w') as file:
        for i in outeng:
            file.write(i + '\n')

    make_word(outeng, name)



frame_option = LabelFrame(root, text = "옵션")
frame_option.pack(fill = "x", padx = 5, pady = 5)

#영단어 범위 시작
day_start = Label(frame_option, text = "영단어 시작", width = 10)
day_start.pack(side = "left", padx = 5, pady = 5)

opt_start = [i for i in range(1, 31)]
cmb_start = ttk.Combobox(frame_option, state = "readonly", values = opt_start, width = 10)
cmb_start.current(0)
cmb_start.pack(side = "left", padx = 5, pady = 5)

#영단어 범위 끝
day_end = Label(frame_option, text = "영단어 끝", width = 10)
day_end.pack(side = "left", padx = 5, pady = 5)

opt_end = [i for i in range(1, 31)]
cmb_end = ttk.Combobox(frame_option, state = "readonly", values = opt_end, width = 10)
cmb_end.current(len(opt_end) - 1)
cmb_end.pack(side = "left", padx = 5, pady = 5)

#영단어 갯수
day_count = Label(frame_option, text = "영단어 갯수", width = 10)
day_count.pack(side = "left", padx = 5, pady = 5)

opt_num = [i for i in range(1, 31)]
cmb_num = ttk.Combobox(frame_option, state ="readonly", values = opt_num, width = 10)
cmb_num.current(4)
cmb_num.pack(side = "left", padx = 5, pady = 5)

#파일 이름
name_frame = LabelFrame(root, text = "시험 이름")
name_frame.pack(fill = "x", padx = 5, pady = 5)

day_end = Label(name_frame, text = "이름", width = 10)
day_end.pack(side = "left", padx = 5, pady = 5)

exam_name = Entry(name_frame)
exam_name.pack(side = "left", fill = "x", expand = True, padx = 5, pady = 5, ipady = 4)

#실행 프레임
rand_frame = Frame(root)
rand_frame.pack(padx = 5, pady = 5)


btn_close = Button(rand_frame, padx = 5, pady = 5, text = "닫기", width = 12, command = root.quit)
btn_close.pack(side = "right", padx = 5, pady = 5)

btn_make_exam = Button(rand_frame, padx = 5, pady = 5, text = "단어 시험 만들기", width = 12, command = make_exam)
btn_make_exam.pack(side = "right", padx = 5, pady = 5)

root.resizable(False, False)
root.mainloop()