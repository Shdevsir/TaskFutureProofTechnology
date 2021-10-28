from posixpath import abspath
from RPA import PDF
from RPA.PDF import PDF
import os

pdf = PDF()

if not os.path.exists('output'):
    os.mkdir('output')

path = f"{os.path.join(os.getcwd())}/output/"
# pdf.set_convert_settings()



data = pdf.get_text_from_pdf(path+'009-000001410.pdf', 1)
list_of_data = data[1].split()
i = 0
index = []
for c in list_of_data:
    if c == "Investment:":
        index.append(i+1)
    elif c == "Unique":
        index.append(i)
    elif c == "(UII):":
        index.append(i+1)
    i += 1

list_of_data[(index[1]-1)] = list_of_data[(index[1]-1)].replace("2.", "")
name = ' '.join(list_of_data[index[0]:index[1]])
list_of_data[index[2]] = list_of_data[index[2]].replace("Section", "")
name2 = list_of_data[index[2]]