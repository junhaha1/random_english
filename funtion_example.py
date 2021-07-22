from random import sample
from docx import Document

def odd_change_even(num):
    if num % 2 == 0:
        return num
    else:
        return num + 1


def make_word(lst):

    doc = Document()
    even = odd_change_even(len(lst))
    row = int(even / 4)
    idx = 0

    table = doc.add_table(rows = row, cols = 4, style = "Table Grid")

    for i in range(row):
        for j in range(4):
            if j % 2 != 0:
                table.rows[i].cells[j].text = ""
            else:
                table.rows[i].cells[j].text = lst[idx]
            idx += 1
    doc.save("테스트.docx")

def make_exam(text):
    outeng = []
    a = sample(range(0, 30), 5)

    f_path = "C:\\Users\\junha\\OneDrive\\바탕 화면\\파이썬\\준희 영어\\준희 영단어 랜덤\\"

    for j in range(1, 31):
        with open("%sword%d.txt" % (f_path, j), 'r') as file:
            lines = file.readlines()
            for i in a:
                outeng.append(lines[i].strip('\n'))

    with open("%s%s.txt" % (f_path, text), 'w') as file:
        for i in outeng:
            file.write(i + '\n')
    
    make_word(outeng)

make_exam("test")



