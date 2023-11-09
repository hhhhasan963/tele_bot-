import os
import pandas as pd
import glob
import numpy as np


def convert_to_int(element):
    if isinstance(element, np.int64):
        return element.item()
    elif isinstance(element, np.float64):
        return int(element.item())
    elif element == "صفر":
        return 0
    elif element == "حجب":
        return element
    elif isinstance(element, str):
        return int(element)
    else:
        return element

# قائمة بالمجلدات التي تحتوي على ملفات HTML
directories = r"C:\Users\DELL\Desktop\mark"

# قائمة بجميع الملفات التي تنتهي بـ ".html" في المجلدات
files = []
for dirpath, dirnames, filenames in os.walk(directories):
    files += glob.glob(os.path.join(dirpath, "*.xlsx"))


def get_info_2(df, student_id):
    # البحث عن السطر الذي يحتوي على رقم الجامعي
    student_row = df[df['الرقم الجامعي'] == student_id]
    student_row2 = df[df['الرقم الجامعي.1'] == student_id]

    if not student_row.empty:
        # استخراج اسم الطالب وعلامته
        student_name = student_row['اسم الطالب'].values[0]
        student_grade = student_row['مجموع'].values[0]
        return student_grade, student_name
    
    elif not student_row2.empty:
        # استخراج اسم الطالب وعلامته
        student_name = student_row2['اسم الطالب.1'].values[0]
        student_grade = student_row2['مجموع.1'].values[0]
        return student_grade, student_name
    
    else:
        return None

def get_info_3(df, student_id):
    # البحث عن السطر الذي يحتوي على رقم الجامعي
    student_row = df[df['الرقم الجامعي'] == student_id]
    student_row2 = df[df['الرقم الجامعي.1'] == student_id]
    student_row3 = df[df['الرقم الجامعي.2'] == student_id]

    if not student_row.empty:
        # استخراج اسم الطالب وعلامته
        student_name = student_row['اسم الطالب'].values[0]
        student_grade = student_row['العلامة'].values[0]
        return student_grade, student_name
    
    elif not student_row2.empty:
        # استخراج اسم الطالب وعلامته
        student_name = student_row2['اسم الطالب.1'].values[0]
        student_grade = student_row2['العلامة.1'].values[0]    
        return student_grade, student_name
    
    elif not student_row3.empty:
        # استخراج اسم الطالب وعلامته
        student_name = student_row3['اسم الطالب.2'].values[0]
        student_grade = student_row3['العلامة.2'].values[0]    
        return student_grade, student_name
    
    else:
        return None

student_id = int(input ("inter your id :    "))
grades = []
name = []
subs = []
years = []
deps =[]
num = 0 
for file in files:
    filename = os.path.basename(file)
    filename_without_extension = os.path.splitext(filename)[0]
    listv = filename_without_extension.split("ـ")
    sentence_list_without_spaces = [sentence.strip() for sentence in listv]
    year = sentence_list_without_spaces[0]
    sub = sentence_list_without_spaces[1]
    dep  = sentence_list_without_spaces[2]
    num += 1
    # subs.append(sub)
    # years.append(year)
    # deps.append(dep)
    # نضيفهن في حال وجود علامة فقط

    df = pd.read_excel(file, header=1)
    # الحصول على الرأس
    header = df.columns.tolist()
    #  الرأس كنص
    header_str = ' '.join(header)
    if header_str == "الرقم الجامعي اسم الطالب العلامة الرقم الجامعي.1 اسم الطالب.1 العلامة.1 الرقم الجامعي.2 اسم الطالب.2 العلامة.2":
        if get_info_3(df, student_id) != None:
            x, y = get_info_3(df, student_id)
            x= convert_to_int(x)

            grades.append(x)
            name.append(y)
            subs.append(sub)
            years.append(year)
            deps.append(dep)
    else:
        if get_info_2(df, student_id) != None:
            x, y = get_info_2(df, student_id)
            x= convert_to_int(x)

            grades.append(x)
            name.append(y)
            subs.append(sub)
            years.append(year)
            deps.append(dep)

# print(name)
# print(grades)
# print(subs)
# print(years)
# print(deps)
print(f"""
The student : {name[0]} from {deps[0]} 
have the following grades :
""")
ii = 0
for x in years:
    if x == "س1":
        print(f"grade: {grades[ii]} in {subs[ii]} from{x}")
    elif x == "س2":
        print(f"grade: {grades[ii]} in {subs[ii]} from{x}")
    elif x == "س3":
        print(f"grade: {grades[ii]} in {subs[ii]} from{x}")
    elif x == "س4":
        print(f"grade: {grades[ii]} in {subs[ii]} from{x}")
    else :
        print("erooor")
    ii += 1

print("")
print (f"we have searched in {num} files.")
