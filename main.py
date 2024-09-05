# ----------------------------------- COMPUTER SCIENCE MAJOR PROJECT - SCHEDULER ---------------------------------------
# ----------------------------------- BY: TIM LI, ALEX MELNBARDIS, AND BOB ZHENG ---------------------------------------

# ----------------------------------------------- PROGRAM SETUP --------------------------------------------------------

# Importing required modules

import tkinter as tk
import pandas as pd
import pickle
import csv

from tkinter import ttk, messagebox, NORMAL, DISABLED
from random import randint

# Setting global variables

global data, current_course, current_teacher, current_student, section_dic, current_period, sections_remaining, \
    rooms_remaining, courses, on, index, count_incomplete, period, item_text, back_config, back_page_2, \
    nc, cn, ot, more_info, close

# Main window setup

window = tk.Tk()
window.title('Scheduler')
window.geometry('1440x900')
window.rowconfigure(0, weight=1)

ttk.Style().theme_use('default')
ttk.Style().configure('Treeview', font=('Calibri', 12))
ttk.Style().configure('Treeview.Heading', font=('Calibri', 13))

# Creating a global variable name for the TreeView and "Back" button used for all pages

tree = ttk.Treeview()
tree2 = ttk.Treeview()
tree3 = ttk.Treeview()
back = ttk.Button()


# -------------------------------------------- SCHEDULE GENERATION -----------------------------------------------------


# Defining function "create_schedule" to read data from the "Student Demands" and "Teacher and Course Information" files

def create_schedule():
    global data, window, current_course, current_teacher, current_student, section_dic, current_period, \
        sections_remaining, sections_remaining, rooms_remaining, count_incomplete, period, on, close

    # Nested dictionary for storing data from both files

    data = {1: {}, 2: {}, 3: {}, 4: {}, 5: {}, 6: {}, 7: {}, 8: {}, 9: {}, 10: {}, 11: {}, 12: {}, 13: {}, 14: {},
            15: {}, 16: {}, 17: {}, 18: {}}

    # 1 - Course: Student Name; 2 - Course: Number of Students; 3 - Course: Number of Sections; 4 - Student Name:Grade;
    # 5 - Course: Teacher(s); 6 - Teacher: Teachable Course(s); 7 - Course: Room Number(s); 8 - Teacher: Sections;
    # 9 - Teacher: Periods; 10 - Course: {Course Section: {Teachers: [Students]}}; 11 - Teacher - Room Number(s);
    # 12 - Room Number - Number of Sections; # 13 - Room Number: {Teacher: Sections}; 14 - Student name: Courses;
    # 15 - Course: Periods; # 16 - Student: Course Selection; # 17: Course: {Period: Number of Students};
    # 18 - Teacher - Room number for each section

    # Reading the data from the "Student Demands" file

    file_1 = pd.read_excel('Student Demands.xlsx')

    for row_index, row in file_1.iterrows():
        if file_1.iloc[row_index, 2] not in data[1]:
            data[1][file_1.iloc[row_index, 2]] = []
            data[2][file_1.iloc[row_index, 2]] = 0

        if file_1.iloc[row_index, 0] not in data[16]:
            data[16][file_1.iloc[row_index, 0]] = []

        data[4][file_1.iloc[row_index, 0]] = file_1.iloc[row_index, 1]
        data[14][file_1.iloc[row_index, 0]] = ['', '', '', '', '', '', '', '']
        data[16][file_1.iloc[row_index, 0]].append(file_1.iloc[row_index, 2])
        data[1][file_1.iloc[row_index, 2]].append(file_1.iloc[row_index, 0])
        data[2][file_1.iloc[row_index, 2]] += 1
        data[3][file_1.iloc[row_index, 2]] = int(data[2][file_1.iloc[row_index, 2]] / 17) + \
                                             (data[2][file_1.iloc[row_index, 2]] % 17 > 0)

    del data[3]['CHV2O']
    del data[2]['CHV2O']
    del data[1]['CHV2O']

    courses_list = list(sorted(data[1].keys()))

    # Reading the data from the "Teacher and Course Information" file

    file_2 = pd.read_excel('Teacher and Course Information.xlsx')

    n = 1

    while n < 12:

        for row_index, row in file_2.iterrows():
            if pd.isna(file_2.iloc[row_index, n]):
                continue

            if file_2.iloc[row_index, n].split("-")[0] not in data[5]:
                data[5][file_2.iloc[row_index, n].split("-")[0]] = []
                data[7][file_2.iloc[row_index, n].split("-")[0]] = []
                data[15][file_2.iloc[row_index, n].split("-")[0]] = []
                data[17][file_2.iloc[row_index, n].split("-")[0]] = {}

            if file_2.iloc[row_index, 0] not in data[6]:
                data[6][file_2.iloc[row_index, 0]] = []
                data[11][file_2.iloc[row_index, 0]] = []

            if file_2.iloc[row_index, n].split("-")[-1] not in data[13]:
                data[13][file_2.iloc[row_index, n].split("-")[-1]] = {}
                data[12][file_2.iloc[row_index, n].split("-")[-1]] = 0

            data[5][file_2.iloc[row_index, n].split("-")[0]].append(file_2.iloc[row_index, 0])
            data[6][file_2.iloc[row_index, 0]].append(file_2.iloc[row_index, n].split("-")[0])
            data[7][file_2.iloc[row_index, n].split("-")[0]].append(file_2.iloc[row_index, n].split("-")[-1])
            data[11][file_2.iloc[row_index, 0]].append(file_2.iloc[row_index, n].split("-")[-1])

        n += 1

    for each in data[5].values():
        each[:] = list(set(each))
    for each in data[6].values():
        each[:] = list(set(each))
    for each in data[7].values():
        each[:] = list(set(each))
    for each in data[11].values():
        each[:] = list(set(each))

    teachers_list = list(sorted(data[6].keys()))

    for teacher in data[6].keys():
        data[8][teacher] = []
        data[9][teacher] = []
        data[18][teacher] = []

    for course in courses_list:
        data[10][course] = {}

    section_dic = {}

    for teacher in teachers_list:
        section_dic[teacher] = {}

    sections_remaining = []

    # Assigning course sections to teachers

    for course in dict(sorted(data[3].items(), key=lambda item: item[1])):
        if len(data[5][course]) == 1:
            for section in range(data[3][course]):
                data[8][data[5][course][0]].append(course)

        if course == 'ENG2D' or course == 'CHC2D' or course == 'MHF4U' or course == 'BBI1O':
            for k in range(data[3][course]):
                teacher_index = randint(0, len(data[5][course]) - 1)

                while len(data[8][data[5][course][teacher_index]]) >= 6:
                    teacher_index = randint(0, len(data[5][course]) - 1)

                teacher = data[5][course][teacher_index]
                data[8][teacher].append(course)

    for course in dict(sorted(data[3].items(), key=lambda item: item[1])):
        if len(data[5][course]) > 1 and course != 'ENG2D' and course != 'CHC2D' and course != 'MHF4U' and course != \
                'BBI1O':

            for i in range(data[3][course]):
                loop = 0
                for values in data[5][course]:
                    if len(data[8][values]) >= 6:
                        loop += 1

                teacher_index = randint(0, len(data[5][course]) - 1)

                if loop < len(data[5][course]):
                    while len(data[8][data[5][course][teacher_index]]) >= 6:
                        teacher_index = randint(0, len(data[5][course]) - 1)
                    teacher = data[5][course][teacher_index]
                    data[8][teacher].append(course)
                else:
                    sections_remaining.append(course)

    # Assigning room numbers to teachers

    rooms_remaining = []

    for teacher in dict(data[11].items()):
        if len(data[11][teacher]) == 1:
            data[13][data[11][teacher][0]][teacher] = []
            for section in data[8][teacher]:
                data[13][data[11][teacher][0]][teacher].append(section)
                data[18][teacher].append(data[11][teacher][0])
                data[12][data[11][teacher][0]] += 1

    for teacher in dict(data[11].items()):
        if len(data[11][teacher]) > 1:
            for section in data[8][teacher]:
                random_room = data[7][section][randint(0, len(data[7][section]) - 1)]

                while data[12][random_room] >= 8 and random_room != 'gym':
                    full = True

                    for room in data[7][section]:
                        if data[12][room] < 8:
                            full = False
                    if not full:
                        random_room = data[7][section][randint(0, len(data[7][section]) - 1)]
                    elif full:
                        rooms_remaining.append(section)
                        break

                if teacher not in data[13][random_room]:
                    data[13][random_room][teacher] = []
                    data[13][random_room][teacher].append(section)
                else:
                    data[13][random_room][teacher].append(section)
                data[12][random_room] += 1
                data[18][teacher].append(random_room)

    # Assigning periods to sections

    ap_period = []

    for teacher in dict(data[8].items()):
        if 'CHC2D' in data[8][teacher] or 'GLC2O' in data[8][teacher] or 'ENG2D' in data[8][teacher] \
                or 'MCV4U' in data[8][teacher]:
            for i in range(len(data[8][teacher])):
                period = chr(ord('@') + (randint(1, 8)))

                while period in data[9][teacher]:
                    period = chr(ord('@') + (randint(1, 8)))
                    loop_count = 0

                    while period in data[15][data[8][teacher][i]] and len(data[15][data[8][teacher][i]]) < 8:
                        loop_count += 1
                        if loop_count > 20:
                            break
                        else:
                            period = chr(ord('@') + (randint(1, 8)))

                if period not in data[10][data[8][teacher][i]]:
                    data[10][data[8][teacher][i]][period] = {}
                    data[10][data[8][teacher][i]][period][teacher] = []
                else:
                    data[10][data[8][teacher][i]][period][teacher] = []
                data[9][teacher].append(period)
                data[15][data[8][teacher][i]].append(period)
                data[17][data[8][teacher][i]][period] = 0

    for teacher in dict(data[8].items()):
        if 'CHC2D' not in data[8][teacher] and 'GLC2O' not in data[8][teacher] and 'ENG2D' not in data[8][teacher] \
                and 'MCV4U' not in data[8][teacher]:
            for i in range(len(data[8][teacher])):
                period = chr(ord('@') + (randint(1, 8)))

                while period in data[9][teacher]:
                    period = chr(ord('@') + (randint(1, 8)))
                    loop_count = 0

                    while period in data[15][data[8][teacher][i]] and len(data[15][data[8][teacher][i]]) < 8:
                        loop_count += 1
                        if loop_count > 20:
                            break
                        else:
                            period = chr(ord('@') + (randint(1, 8)))

                    while (data[8][teacher][i] == "APBIO" or data[8][teacher][i] == "APCHE" or
                           data[8][teacher][i] == "APPHY" or data[8][teacher][i] == "APCAL" or
                           data[8][teacher][i] == "ICS4U") and period in ap_period:

                        period = chr(ord('@') + (randint(1, 8)))

                if data[8][teacher][i] == "APBIO" or data[8][teacher][i] == "APCHE" or data[8][teacher][i] == "APPHY" \
                        or data[8][teacher][i] == "APCAL" or data[8][teacher][i] == "ICS4U":
                    ap_period.append(period)

                if period not in data[10][data[8][teacher][i]]:
                    data[10][data[8][teacher][i]][period] = {}
                    data[10][data[8][teacher][i]][period][teacher] = []
                else:
                    data[10][data[8][teacher][i]][period][teacher] = []

                data[9][teacher].append(period)
                data[15][data[8][teacher][i]].append(period)
                data[17][data[8][teacher][i]][period] = 0

    # Adding students to course sections

    period_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
    course_left = []

    def common_period():
        period_avail = []
        for periods in range(len(data[14][student])):
            if data[14][student][periods] == "":
                period_avail.append(period_list[periods])
        period_avail_set = set(period_avail)
        section_avail_set = set(data[15][course])

        if len(period_avail_set.intersection(section_avail_set)) > 0:
            return 1
        else:
            return 0

    con = 0

    for course in sorted(data[1], key=lambda course_temp: len(data[1][course_temp]), reverse=False):
        for student in data[1][course]:
            able = common_period()
            if able == 1:
                course_period = data[15][course][randint(0, len(data[15][course]) - 1)]
                while data[14][student][period_list.index(course_period)] != '':
                    course_period = data[15][course][randint(0, len(data[15][course]) - 1)]
                data[14][student][period_list.index(course_period)] = course
                data[17][course][course_period] += 1

                if len(data[10][course][course_period]) > 1:
                    teacher_list = []
                    for key in data[10][course][course_period]:
                        teacher_list.append(key)
                    data[10][course][course_period][teacher_list[randint(0, len(teacher_list) - 1)]].append(student)

                elif len(data[10][course][course_period]) == 1:
                    data[10][course][course_period][list(data[10][course][course_period].keys())[0]].append(student)

                con += 1

            elif able == 0:
                course_left.append(course)

    count_incomplete = 0

    for student in data[14]:
        if "" in data[14][student]:
            count_incomplete += 1

    for student in data[16]:
        if len(data[16][student]) < 8:
            count_incomplete -= 1

    on = False
    close = False


# ------------------------------------------------ PAGE CREATION -------------------------------------------------------


# Homepage 1 (Course List) ---------------------------------------------------------------------------------------------

def homepage1():
    global tree, back, sections_remaining, courses, tree2, tree3, close
    if on is True and close is False:
        entry_edit.config(state=DISABLED)
        okb.config(state=DISABLED)
        close = True

    delete_page()
    delete_page2()
    delete_page3()

    courses = list(sorted(data[3].keys()))

    # Course list TreeView

    tree = ttk.Treeview(window, height=38, columns=('1', '2', '3', '4'), show='headings')
    tree.heading(1, text='Course')
    tree.heading(2, text='Available')
    tree.heading(3, text='Occupied')
    tree.heading(4, text='Sections')

    # Enabling/disabling the Student List/Course List buttons

    course_list.config(state=DISABLED)
    student_list.config(state=NORMAL)
    # Configuring the scrollbar

    scrollbar = ttk.Scrollbar(window, orient=tk.VERTICAL, command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.grid(row=1, column=2, sticky='ns')

    # Configuring the "Back" button

    back = ttk.Button(window, text='<= Back')
    back.grid(row=0, column=0, sticky='wns')
    back.config(state=DISABLED)

    # "Search" button:

    def search():
        global courses, tree, sections_remaining

        courses = []
        tree_data_1 = []

        search_button.config(state=DISABLED)
        search_entry.config(state=DISABLED)
        reset_button.config(state=NORMAL)

        for each in tree.get_children():
            tree_data_1.append(tree.item(each)['values'])

        delete_page()

        tree = ttk.Treeview(window, height=38, columns=('1', '2', '3', '4'), show='headings')
        tree.heading(1, text='Course')
        tree.heading(2, text='Available')
        tree.heading(3, text='Occupied')
        tree.heading(4, text='Sections')

        for heading_temp in range(5):
            tree.column(heading_temp, width=100)

        tree.grid(row=1, column=0, sticky='nsew')

        count1 = 0

        for row in tree_data_1:
            if search_entry.get().lower() in row[0].lower():
                courses.append(row[0])
                if count1 % 2 == 0:
                    tree.insert('', tk.END, values=(row[0], row[1], row[2], row[3]), tags='even_row')
                else:
                    tree.insert('', tk.END, values=(row[0], row[1], row[2], row[3]))

                count1 += 1

        tree.tag_configure('even_row', background='gray95')

        tree.bind('<Double-1>', double_click_course)

    def search_reset():
        search_button.config(state=NORMAL)
        search_entry.config(state=DISABLED)
        reset_button.config(state=DISABLED)
        search_entry.delete(0, 'end')
        homepage1()

    search_button = ttk.Button(window, text='Search', command=search)
    search_button.place(x=200, y=14)
    search_entry = ttk.Entry(window, width=25, textvariable=tk.StringVar())
    search_entry.place(x=268, y=18)
    reset_button = ttk.Button(window, text='Reset', command=search_reset)
    reset_button.place(x=424, y=14)
    reset_button.config(state=DISABLED)

    # Configuring column size

    for heading in range(5):
        tree.column(heading, width=100)

    tree.grid(row=1, column=0, sticky='nsew')

    # Inserting data

    striping = 0

    for key, value in sorted(data[2].items()):
        if key in sections_remaining:
            tree.insert('', tk.END, values=(key, (17 * data[3][key]), value, data[3][key]), tags=('left',))
        else:
            if striping % 2 == 0:
                tree.insert('', tk.END, values=(key, (17 * data[3][key]), value, data[3][key]), tags='even_row')
            else:
                tree.insert('', tk.END, values=(key, (17 * data[3][key]), value, data[3][key]))

        striping += 1

    tree.tag_configure('even_row', background='gray95')

    # Grid click event (Page 2)

    tree.bind('<Double-1>', double_click_course)


# Homepage 2 (Student List) --------------------------------------------------------------------------------------------

def homepage2():
    global tree, back, back_config, tree2, tree3, close

    if on is True and close is False:
        entry_edit.config(state=DISABLED)
        okb.config(state=DISABLED)
        close = True

    delete_page()
    delete_page2()
    delete_page3()

    # Student list TreeView

    tree = ttk.Treeview(window, height=38, columns=('1', '2'), show='headings')
    tree.heading(1, text='Name')
    tree.heading(2, text='Grade')

    # Enabling/disabling the Course List/Student List buttons

    course_list.config(state=NORMAL)
    student_list.config(state=DISABLED)

    # Configuring the "Back" button

    back_config = "student"

    def back_page():
        homepage1()

    back = ttk.Button(window, text='<= Back', command=back_page)
    back.grid(row=0, column=0, sticky='wns')

    # configuring the scrollbar

    scrollbar = ttk.Scrollbar(window, orient=tk.VERTICAL, command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.grid(row=1, column=2, sticky='ns')

    # "Search" button:

    def search():
        global courses, tree, sections_remaining

        courses = []
        tree_data_2 = []

        search_button.config(state=DISABLED)
        search_entry.config(state=DISABLED)
        reset_button.config(state=NORMAL)

        for each in tree.get_children():
            tree_data_2.append(tree.item(each)['values'])

        delete_page()

        tree = ttk.Treeview(window, height=38, columns=('1', '2'), show='headings')
        tree.heading(1, text='Name')
        tree.heading(2, text='Grade')

        for heading_temp in range(3):
            tree.column(heading_temp, width=150)

        tree.grid(row=1, column=0, sticky='nsew')

        scrollbar_temp = ttk.Scrollbar(window, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar_temp.set)
        scrollbar_temp.grid(row=1, column=2, sticky='ns')

        count1 = 0

        for row in tree_data_2:
            if search_entry.get().lower() in row[0].lower():
                courses.append(row[0])
                if count1 % 2 == 0:
                    tree.insert('', tk.END, values=(row[0], row[1]), tags='even_row')
                else:
                    tree.insert('', tk.END, values=(row[0], row[1]))

                count1 += 1

        tree.tag_configure('even_row', background='gray95')

        tree.bind('<Double-1>', double_click_student_from_list)

    def search_reset():
        search_button.config(state=NORMAL)
        search_entry.config(state=DISABLED)
        reset_button.config(state=DISABLED)
        search_entry.delete(0, 'end')
        homepage2()

    search_button = ttk.Button(window, text='Search', command=search)
    search_button.place(x=200, y=14)
    search_entry = ttk.Entry(window, width=25, textvariable=tk.StringVar())
    search_entry.place(x=268, y=18)
    reset_button = ttk.Button(window, text='Reset', command=search_reset)
    reset_button.place(x=424, y=14)
    reset_button.config(state=DISABLED)

    # Configuring column size

    for heading in range(3):
        tree.column(heading, width=150)

    tree.grid(row=1, column=0, sticky='nsw')

    # Inserting the student list data

    striping = 0

    for key, value in sorted(data[4].items()):
        if striping % 2 == 0:
            tree.insert('', tk.END, values=(key, value), tags='even_row')
        else:
            tree.insert('', tk.END, values=(key, value))

        striping += 1

    tree.tag_configure('even_row', background='gray95')

    # Grid click event (to Student Schedule page)

    tree.bind('<Double-1>', double_click_student_from_list)


# Page 2 (Teacher - Course Section(s) page) ----------------------------------------------------------------------------

def page2():
    global tree, back, tree2, tree3
    entry_edit.config(state=DISABLED)
    okb.config(state=DISABLED)
    delete_page2()
    delete_page3()

    # Enabling both Course List and Student List buttons

    course_list.config(state=NORMAL)
    student_list.config(state=NORMAL)

    # Configuring the "Back" button

    def back_page():
        homepage1()

    back = ttk.Button(window, text='<= Back', command=back_page)
    back.grid(row=0, column=0, sticky='wns')

    # Creating fake disabled search feature

    search_button = ttk.Button(window, text='Search')
    search_button.place(x=200, y=14)
    search_button.config(state=DISABLED)
    search_entry = ttk.Entry(window, width=25, textvariable=tk.StringVar())
    search_entry.place(x=268, y=18)
    search_entry.config(state=DISABLED)
    reset_button = ttk.Button(window, text='Reset')
    reset_button.place(x=424, y=14)
    reset_button.config(state=DISABLED)

    tree2 = ttk.Treeview(window, height=38, columns=('1', '2'), show='headings')
    tree2.heading(1, text='Teachers')
    tree2.heading(2, text='Sections')

    # Configuring column size

    for heading in range(3):
        tree2.column(heading, width=150)

    # Inserting the teacher list data for each course

    striping = 0

    for name in data[5][current_course]:
        if striping % 2 == 0:
            tree2.insert('', tk.END,
                         values=(name, str(data[8][name].count(current_course)) + "         " + current_course),
                         tags="even_row")
        else:
            tree2.insert('', tk.END,
                         values=(name, str(data[8][name].count(current_course)) + "         " + current_course))
        striping += 1

    tree2.tag_configure('even_row', background='gray95')

    tree2.grid(row=1, column=3, sticky='nsw')

    # Grid click event (to Teacher Schedule page)

    tree2.bind('<Double-1>', double_click_teacher)


# Student Schedule page ------------------------------------------------------------------------------------------------

def student_schedule():
    global data, tree, back, back_config, back_page_2, on, more_info, tree2, tree3, close

    delete_page2()

    course_list.config(state=DISABLED)
    student_list.config(state=DISABLED)

    # Configuring the "Back" button

    if back_config == "teacher":
        def back_page_2():
            global back_config, close
            back_config = 0

            student_section_list()
            if on is True and close is False:
                entry_edit.config(state=DISABLED)
                okb.config(state=DISABLED)
                close = True

    if back_config == "student":
        def back_page_2():
            global close

            if on is True and close is False:
                entry_edit.config(state=DISABLED)
                okb.config(state=DISABLED)
                close = True
            homepage2()

    back = ttk.Button(window, text='<= Back', command=back_page_2)
    back.grid(row=0, column=0, sticky='wns')

    tree2 = ttk.Treeview(window, height=38, columns=('1', '2', '3', '4'), show='headings')
    tree2.heading(1, text='Period')
    tree2.heading(2, text='Course')
    tree2.heading(3, text='Teacher')
    tree2.heading(4, text='Room #')

    for heading in range(5):
        tree2.column(heading, width=150)

    tree2.grid(row=1, column=3, sticky='nsew')

    # Create fake disabled search feature

    search_button = ttk.Button(window, text='Search')
    search_button.place(x=200, y=14)
    search_button.config(state=DISABLED)
    search_entry = ttk.Entry(window, width=25, textvariable=tk.StringVar())
    search_entry.place(x=268, y=18)
    search_entry.config(state=DISABLED)
    reset_button = ttk.Button(window, text='Reset')
    reset_button.place(x=424, y=14)
    reset_button.config(state=DISABLED)

    # Course functions (change, add/remove)

    period_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']

    striping = 0

    for course in range(len(data[14][current_student])):
        if data[14][current_student][course] == "":
            if striping % 2 == 0:
                tree2.insert('', tk.END, values=(period_list[course], 'Spare', 'N/A', 'N/A'), tags='even_row')
            else:
                tree2.insert('', tk.END, values=(period_list[course], 'Spare', 'N/A', 'N/A'))

        elif data[14][current_student][course] != "":
            teacher = ''
            room = ''

            if len(data[10][data[14][current_student][course]][period_list[course]]) > 1:
                for teachers in data[10][data[14][current_student][course]][period_list[course]]:
                    if current_student in data[10][data[14][current_student][course]][period_list[course]][teachers]:
                        teacher = teachers

            elif len(data[10][data[14][current_student][course]][period_list[course]]) == 1:
                teacher = list(data[10][data[14][current_student][course]][period_list[course]].keys())[0]

            room = data[18][teacher][data[9][teacher].index(period_list[course])]

            if striping % 2 == 0:
                tree2.insert('', tk.END, values=(period_list[course], data[14][current_student][course], teacher, room),
                             tags='even_row')
            else:
                tree2.insert('', tk.END, values=(period_list[course], data[14][current_student][course], teacher, room))

        striping += 1

    tree2.tag_configure('even_row', background='gray95')

    more_info = False

    def set_cell_value(event):
        global on, data, item_text, entry_edit, nc, ot, more_info, okb, close, cn
        close = False
        for item in tree2.selection():
            item_text = tree2.item(item, "values")
        entry_edit = tk.Entry(window, width=25, textvariable=tk.StringVar())
        nc = str(entry_edit.get())
        on = False

        if on is False:
            period_dict = {1: "A", 2: "B", 3: "C", 4: "D", 5: "E", 6: "F", 7: "G", 8: "H"}
            oc = item_text[1]
            ot = item_text[2]
            column = tree2.identify_column(event.x)
            row = tree2.identify_row(event.y)
            cn = int(str(column).replace('#', ''))
            rn = int(str(row).replace('I', ''))

        def save_edit():
            global on, data, more_info, nc, entry_edit, okb

            if more_info is False:  # enter a course mode
                if str(entry_edit.get()) in set(courses):  # it is a course
                    if period_dict[rn] in data[15][str(entry_edit.get())]:  # period available
                        if data[17][str(entry_edit.get())][period_dict[rn]] < 17:  # student number available

                            if len(list((data[10][str(entry_edit.get())][period_dict[rn]].keys()))) == 1:
                                # only one teacher in that course period
                                on = False

                                # Update the tree

                                teach_insert = str(list((data[10][str(entry_edit.get())][
                                                             period_dict[rn]].keys())))[2:-2]
                                tree2.set(item, column=column, value=str(entry_edit.get()))
                                tree2.set(item, column=3, value=teach_insert)
                                tree2.set(item, column=4,
                                          value=data[18][teach_insert][
                                              data[9][teach_insert].index(period_list[rn - 1])])

                                # Update data

                                if oc != 'Spare':
                                    data[1][oc].remove(current_student)
                                    data[2][oc] -= 1
                                    data[10][oc][period_dict[rn]][ot].remove(current_student)
                                    data[16][current_student].remove(oc)
                                    data[17][oc][period_dict[rn]] -= 1

                                data[1][entry_edit.get()].append(current_student)
                                data[2][entry_edit.get()] += 1
                                data[10][str(entry_edit.get())][period_dict[rn]][str(list(
                                    (data[10][str(entry_edit.get())]
                                     [period_dict[rn]].keys())))[2:-2]].append(current_student)
                                data[14][current_student][rn - 1] = entry_edit.get()
                                data[16][current_student].append(entry_edit.get())
                                data[17][entry_edit.get()][period_dict[rn]] += 1

                                entry_edit.config(state=DISABLED)
                                okb.config(state=DISABLED)

                            else:
                                messagebox.showinfo('Scheduler', "Please specify a teacher.")
                                nc = entry_edit.get()
                                more_info = True
                                entry_edit.delete(0, 'end')

                        else:
                            size_warning = messagebox.askokcancel("Warning", "Class size is larger than 16 students. "
                                                                             "Are you sure?")

                            if size_warning:
                                if len(list(
                                        (data[10][str(entry_edit.get())][
                                            period_dict[rn]].keys()))) == 1:
                                    on = False

                                    # Update the tree

                                    tree2.set(item, column=column, value=str(entry_edit.get()))
                                    tree2.set(item, column=3, value=str(
                                        list((data[10][str(entry_edit.get())]
                                              [period_dict[rn]].keys())))[2:-2])

                                    teach_insert = str(list((data[10][str(entry_edit.get())][
                                                                 period_dict[rn]].keys())))[2:-2]
                                    tree2.set(item, column=4, value=data[18][teach_insert][data[9][teach_insert].index(
                                        period_list[rn - 1])])

                                    # Update data

                                    if oc != 'Spare':
                                        data[1][oc].remove(current_student)
                                        data[2][oc] -= 1
                                        data[10][oc][period_dict[rn]][ot].remove(current_student)
                                        data[16][current_student].remove(oc)
                                        data[17][oc][period_dict[rn]] -= 1

                                    data[1][entry_edit.get()].append(current_student)
                                    data[2][entry_edit.get()] += 1
                                    data[10][str(entry_edit.get())][period_dict[rn]][str(
                                        list((data[10][str(entry_edit.get())][
                                                  period_dict[rn]].keys())))[2:-2]].append(current_student)
                                    data[14][current_student][rn - 1] = entry_edit.get()
                                    data[16][current_student].append(entry_edit.get())
                                    entry_edit.config(state=DISABLED)
                                    okb.config(state=DISABLED)

                                else:
                                    messagebox.showinfo('Scheduler', "Please specify a teacher.")
                                    nc = str(entry_edit.get())
                                    more_info = True
                                    entry_edit.delete(0, 'end')

                    else:
                        messagebox.showinfo("Error", "Course unavailable")
                        entry_edit.config(state=DISABLED)
                        okb.config(state=DISABLED)
                        on = False

                elif str(entry_edit.get()) == '' \
                        or str(entry_edit.get()) == 'None' \
                        or str(entry_edit.get()) == 'none' \
                        or str(entry_edit.get()) == 'Spare' \
                        or str(entry_edit.get()) == 'spare':

                    if oc != 'Spare':
                        remove_warning = messagebox.askokcancel("Warning", "Are sure you want to remove this course?")

                        if remove_warning:
                            tree2.set(item, column=1, value=period_dict[rn])
                            tree2.set(item, column=2, value='Spare')
                            tree2.set(item, column=3, value='N/A')
                            tree2.set(item, column=4, value='N/A')

                            data[1][oc].remove(current_student)
                            data[2][oc] -= 1
                            data[10][oc][period_dict[rn]][ot].remove(current_student)
                            data[14][current_student][rn - 1] = ''
                            data[16][current_student].remove(oc)
                            data[17][oc][period_dict[rn]] -= 1

                            entry_edit.config(state=DISABLED)
                            okb.config(state=DISABLED)
                            on = False

                    else:
                        messagebox.showinfo("Warning", "It is already Spare")
                        entry_edit.config(state=DISABLED)
                        okb.config(state=DISABLED)
                        on = False

                else:
                    messagebox.showinfo("Error", "Not in the Course List!")

            elif str(entry_edit.get()) in set(data[10][nc][period_dict[rn]]):
                on = False

                # Update the tree

                tree2.set(item, column=column, value=nc)
                teach_insert = str(entry_edit.get())
                tree2.set(item, column=3, value=teach_insert)
                tree2.set(item, column=4,
                          value=data[18][teach_insert][data[9][teach_insert].index(period_list[rn - 1])])

                # Update data

                if oc != 'Spare':
                    data[1][oc].remove(current_student)
                    data[2][oc] -= 1
                    data[10][oc][period_dict[rn]][ot].remove(current_student)
                    data[16][current_student].remove(oc)
                    data[17][oc][period_dict[rn]] -= 1

                data[1][nc].append(current_student)
                data[2][nc] += 1
                data[10][nc][period_dict[rn]][entry_edit.get()].append(current_student)
                data[14][current_student][rn - 1] = nc
                data[16][current_student].append(nc)
                entry_edit.config(state=DISABLED)
                okb.config(state=DISABLED)

                more_info = False
                on = False

            elif str(entry_edit.get()) not in set(data[10][nc][period_dict[rn]]):

                if str(entry_edit.get()) == '' or str(entry_edit.get()) == 'None':
                    messagebox.showinfo('Course Change Cancel', "You should specify a teacher name")

                else:
                    messagebox.showinfo('Warning', 'The teacher name does not exist')
                more_info = False
                entry_edit.config(state=DISABLED)
                okb.config(state=DISABLED)
                on = False

        if more_info is False:
            okb = ttk.Button(window, text='Add Course', width=13, command=save_edit)

        if on is False and cn == 2:  # only double-click on the second column can make change
            entry_edit.place(x=858, y=18)
            okb.place(x=1015, y=14)
            on = True

        else:
            entry_edit.config(state=DISABLED)
            okb.config(state=DISABLED)
            on = False

    tree2.bind('<Double-1>', set_cell_value)


# Page 3 (Teacher Schedule page) ---------------------------------------------------------------------------------------


def teacher_schedule():
    global tree, back, section_dic, tree2, tree3

    delete_page3()

    course_list.config(state=NORMAL)
    student_list.config(state=NORMAL)

    # Configuring the "Back" button

    def back_page():
        page2()

    back = ttk.Button(window, text='<= Back', command=back_page)
    back.grid(row=0, column=0, sticky='wns')

    # Create fake disabled search feature

    search_button = ttk.Button(window, text='Search')
    search_button.place(x=200, y=14)
    search_button.config(state=DISABLED)
    search_entry = ttk.Entry(window, width=25, textvariable=tk.StringVar())
    search_entry.place(x=268, y=18)
    search_entry.config(state=DISABLED)
    reset_button = ttk.Button(window, text='Reset')
    reset_button.place(x=424, y=14)
    reset_button.config(state=DISABLED)

    # Setting up the new treeview

    tree3 = ttk.Treeview(window, height=38, columns=('0', '1', '2', '3', '4', '5', '6', '7', '8', '9'), show='headings')
    tree3.heading(0, text="Teacher")
    tree3.heading(1, text='Course')
    tree3.heading(2, text='A')
    tree3.heading(3, text='B')
    tree3.heading(4, text='C')
    tree3.heading(5, text='D')
    tree3.heading(6, text='E')
    tree3.heading(7, text='F')
    tree3.heading(8, text='G')
    tree3.heading(9, text='H')

    tree3.column(1, width=75)
    tree3.column(0, width=150)

    for heading in range(2, 10):
        tree3.column(heading, width=50, anchor="center")

    # Inserting data

    stripe_count_4 = 0
    a = 1

    for stuff in data[6][current_teacher]:
        period_dict = {"A": '', "B": '', "C": '', "D": '', "E": '', "F": '', "G": '', "H": ''}

        count = []

        for number in range(len(data[8][current_teacher])):
            if data[8][current_teacher][number] == stuff:
                count.append(number)

        for p in count:
            period_dict[data[9][current_teacher][p]] = len(
                data[10][stuff][data[9][current_teacher][p]][current_teacher])
            section_dic[current_teacher][data[9][current_teacher][p]] = str(a)
            a += 1

        if stripe_count_4 % 2 == 0:
            tree3.insert('', tk.END, values=(current_teacher, stuff, period_dict["A"], period_dict["B"],
                                             period_dict["C"], period_dict["D"], period_dict["E"],
                                             period_dict["F"], period_dict["G"], period_dict["H"]), tags='even_row')
        else:
            tree3.insert('', tk.END, values=(current_teacher, stuff, period_dict["A"], period_dict["B"],
                                             period_dict["C"], period_dict["D"], period_dict["E"],
                                             period_dict["F"], period_dict["G"], period_dict["H"]))

        stripe_count_4 += 1

    tree3.tag_configure('even_row', background='gray95')

    tree3.grid(row=1, column=4, sticky='nsew')

    # Grid click event (to Section Student List page)

    tree3.bind('<Double-1>', double_click_course2)


# Page 4 (Section Student List page) -----------------------------------------------------------------------------------

def student_section_list():
    global tree, back, back_config, tree2, tree3

    delete_page()
    delete_page2()
    delete_page3()

    course_list.config(state=NORMAL)
    student_list.config(state=NORMAL)

    # Configuring the "Back" button

    back_config = "teacher"

    def back_page():
        homepage1()
        page2()
        teacher_schedule()

    back = ttk.Button(window, text='<= Back', command=back_page)
    back.grid(row=0, column=0, sticky='wns')

    # Create fake disabled search feature

    search_button = ttk.Button(window, text='Search')
    search_button.place(x=200, y=14)
    search_button.config(state=DISABLED)
    search_entry = ttk.Entry(window, width=25, textvariable=tk.StringVar())
    search_entry.place(x=268, y=18)
    search_entry.config(state=DISABLED)
    reset_button = ttk.Button(window, text='Reset')
    reset_button.place(x=424, y=14)
    reset_button.config(state=DISABLED)

    tree = ttk.Treeview(window, height=38, columns=('1', '2'), show='headings')
    tree.heading(1, text='Name')
    tree.heading(2, text='Grade')

    striping = 0

    for name in data[10][current_course][current_period][current_teacher]:
        if striping % 2 == 0:
            tree.insert('', tk.END, values=(name, data[4][name]), tags='even_row')
        else:
            tree.insert('', tk.END, values=(name, data[4][name]))

        striping += 1

    tree.tag_configure('even_row', background='gray95')

    for heading in range(3):
        tree.column(heading, width=100)

    tree.grid(row=1, column=0, sticky='nsew')

    # Grid click event (to Student Schedule page)

    tree.bind('<Double-1>', double_click_student_from_class)


# --------------------------------------------- BUTTONS AND COMMANDS ---------------------------------------------------


# Save configuration to .pickle file feature

def save_configuration():
    pickle_out = open('configuration.pickle', 'wb')
    pickle.dump(data, pickle_out)
    pickle_out.close()
    messagebox.showinfo('Scheduler', 'Configuration saved.')


save = ttk.Button(window, text='Save', command=save_configuration)
save.place(x=100, y=0)


# Load configuration Feature

def load_file():
    global data

    pickle_in = open('configuration.pickle', 'rb')
    data = pickle.load(pickle_in)
    messagebox.showinfo('Scheduler', 'Configuration loaded.')
    homepage1()


load = ttk.Button(window, text='Load', command=load_file)
load.place(x=100, y=26)


# Refresh configuration feature

def refresh_cmd():
    refresh_msg = messagebox.askokcancel("Warning", "Any unsaved changes will be lost. Are you sure?")
    if refresh_msg:
        delete_page()
        create_schedule()
        while sections_remaining != [] or rooms_remaining != []:
            create_schedule()
        homepage1()


refresh_button = ttk.Button(window, text='Refresh', command=refresh_cmd)
refresh_button.place(x=1140, y=0)


# Export student schedules to .csv file feature


def export():
    export_schedules = csv.writer(open('schedules.csv', 'w', newline=''))
    for student, schedule in sorted(data[14].items()):
        export_schedules.writerow([student, ' '.join(schedule)])

    messagebox.showinfo('Scheduler', 'Student schedules exported.')


export_button = ttk.Button(window, text='Export', command=export)
export_button.place(x=1140, y=26)


# "Student List" and "Course List" buttons, entry_edit box and OK button

student_list = ttk.Button(window, text='Student List =>', command=homepage2)
student_list.place(x=1340, y=0, width=100, height=52)
course_list = ttk.Button(window, text='<= Course List', command=homepage1)
course_list.place(x=1240, y=0, width=100, height=52)

entry_edit = tk.Entry(window, width=25, textvariable=tk.StringVar())
okb = ttk.Button(window, text='Add Course', width=13)
okb.place(x=1015, y=14)
okb.config(state=DISABLED)
entry_edit.place(x=858, y=18)
entry_edit.config(state=DISABLED)


# Double click event for Homepage 1

def double_click_course(event):
    global index, courses, current_course

    item = tree.selection()[0]
    index = int((item[1] + item[2] + item[3]), 16)
    current_course = courses[index - 1]

    page2()


# Double click event for Homepage 2

def double_click_student_from_list(event):
    global index, current_student, on, close

    item = tree.selection()[0]
    index = int((item[1] + item[2] + item[3]), 16)
    students = list(sorted(data[4].keys()))
    current_student = students[index - 1]

    if on is True and close is False:
        entry_edit.config(state=DISABLED)
        okb.config(state=DISABLED)
        close = True
    student_schedule()


# Double click event for Teacher - Course Section(s) page

def double_click_teacher(event):
    global index, current_course, current_teacher

    item = tree2.selection()[0]
    index = int((item[1] + item[2] + item[3]), 16)
    current_teacher = data[5][current_course][index - 1]

    teacher_schedule()


# Double click event for Teacher Schedule page

def double_click_course2(event):
    global index, courses, current_course, section_dic, current_period

    item = tree3.identify('item', event.x, event.y)

    if 20 * (len(data[6][current_teacher]) + 1) > event.y > 20:
        index = int((item[1] + item[2] + item[3]), 16)
    current_course = data[6][current_teacher][index - 1]

    if 225 > event.x > 150 and 20 * (len(data[6][current_teacher]) + 1) > event.y > 20:
        page2()

    if event.x > 225 and 20 * (len(data[6][current_teacher]) + 1) > event.y > 20:
        current_period = chr(ord('@') + (event.x - 225) // 50 + 1)
        if current_period in data[10][current_course] and current_teacher in data[10][current_course][current_period]:
            student_section_list()


# Double click event for Section Student List page

def double_click_student_from_class(event):
    global index, current_student, close
    if on is True and close is False:
        entry_edit.config(state=DISABLED)
        okb.config(state=DISABLED)
        close = True
    item = tree.selection()[0]
    index = int((item[1] + item[2] + item[3]), 16)
    students = list(data[10][current_course][current_period][current_teacher])
    current_student = students[index - 1]

    student_schedule()


# Clear page command

def delete_page():
    for col in tree['columns']:
        tree.heading(col, text='')

    tree['columns'] = ()
    tree.delete(*tree.get_children())


def delete_page2():
    for col in tree2['columns']:
        tree2.heading(col, text='')

    tree2['columns'] = ()
    tree2.delete(*tree2.get_children())


def delete_page3():
    for col in tree3['columns']:
        tree3.heading(col, text='')

    tree3['columns'] = ()
    tree3.delete(*tree3.get_children())


# Running the "create_schedule" function

create_schedule()

# Ensuring the program starts only once all sections have been assigned

while sections_remaining != [] or rooms_remaining != [] or count_incomplete >= 50:
    create_schedule()

homepage1()
window.mainloop()
