import xlwt
# import pandas as pd
from pandas import read_excel as pd
import Excel_setting
from os import getcwd
from os import listdir
# import os
# import shutil


def set_Title(sheet, EndThreeNumber):
    # sheet = newStudenfile.add_sheet("Sheet1", cell_overwrite_ok=True)
    sheet.write_merge(0, 0, 0, 1, '學號末3碼:  ' + EndThreeNumber,
                      style=xlwt.easyxf('font: name 標楷體, height 220,  bold on'))
    sheet.write_merge(1, 3, 0, 0, '年級', Chartstyle)
    sheet.write_merge(1, 3, 1, 1, '課程名稱', Chartstyle)
    sheet.write_merge(1, 3, 2, 2, '必/  選修', Chartstyle)
    sheet.write_merge(1, 1, 3, 7, '學分數', Chartstyle)
    sheet.write_merge(
        2, 2, 5, 6, '工程專業課程      (若一課程部分屬於理論，部分屬於設計/實務，分開計算)', Chartstyle)
    sheet.write_merge(2, 3, 3, 3, '數學', Chartstyle)
    sheet.write_merge(2, 3, 4, 4, '基礎  科學', Chartstyle)
    sheet.write(3, 5, '理論', Chartstyle)
    sheet.write(3, 6, '設計', Chartstyle)
    sheet.write_merge(2, 3, 7, 7, '通識 課程', Chartstyle)
    sheet.row(0).height_mismatch = True
    sheet.row(1).height_mismatch = True
    sheet.row(2).height_mismatch = True
    sheet.row(3).height_mismatch = True
    sheet.row(1).height = 256 * 2
    sheet.row(2).height = 256 * 8
    sheet.row(3).height = 256 * 2
    sheet.col(1).width_mismatch = True


def set_End(sheet, index, graduationCredit):  # sheet.write_merge(row1,row2,col1,col2,"",style)
    sheet.write(index, 2, '小計', Chartstyle)
    sheet.write(index + 1, 2, '總計', Chartstyle)
    sheet.write(index, 3, xlwt.Formula('SUM(D5:D' + str(index) + ')'),
                TopStyle)  # Mathematic credits summation
    sheet.write(index, 4, xlwt.Formula('SUM(E5:E' + str(index) + ')'),
                TopRightStyle)  # Science credits summation
    sheet.write(index, 5, xlwt.Formula('SUM(F5:F' + str(index) + ')'),
                TopStyle)  # Theory credits summation
    sheet.write(index, 6, xlwt.Formula('SUM(G5:G' + str(index) + ')'),
                TopRightStyle)  # Design credits summation
    sheet.write_merge(index + 1, index + 1, 3, 4, xlwt.Formula('SUM(D' +
                                                               str(index+1)+':E'+str(index+1)+')'), Chartstyle)  # Math and Science subtotal
    sheet.write_merge(index + 1, index + 1, 5, 6, xlwt.Formula('SUM(F' +
                                                               str(index+1) + ':G' + str(index+1) + ')'), Chartstyle)  # Theory and Design subtotal
    sheet.write_merge(index, index + 1, 7, 7, xlwt.Formula('SUM(H5:H' +
                                                           str(index) + ')'), Chartstyle)  # General course credits summation
    sheet.write_merge(index, index + 1, 0, 1, '修課總學分數',
                      Chartstyle)
    sheet.write_merge(index+2, index+2, 0, 2,
                      "IEET 認證規範 4 課程學分數之要求", Chartstyle)
    sheet.write_merge(index + 2, index + 2, 3, 4,
                      '32學分', Excel_setting.set_GreenBackground())
    sheet.write_merge(index + 2, index + 2, 5, 6,
                      '48學分', Excel_setting.set_GreenBackground())
    sheet.write(index + 2, 7, "", Chartstyle)
    sheet.col(1).width_mismatch = True
    sheet.col(1).width = 256 * 30
    sheet.write_merge(index+3, index+3, 0, 2,
                      "學程最低畢業學分數", Chartstyle)
    sheet.write_merge(index+3, index+3, 3, 7,
                      graduationCredit, Chartstyle)


def write_ETCourse(sheet, row, studentCourse, studentCourseindex, courseTable, courseTableindex):
    sheet.write(row, 0, str(
        studentCourse.iloc[studentCourseindex, 0]), ContentStyle)  # Semester
    # Course Name
    sheet.write(
        row, 1, studentCourse.iloc[studentCourseindex, 1], ContentStyle)
    # Require/Elective subject
    sheet.write(
        row, 2, studentCourse.iloc[studentCourseindex, 2], ThickonRight)
    # Mathematics credits
    sheet.write(row, 3, int(
        courseTable.iloc[courseTableindex, 1]), ContentStyle)
    sheet.write(row, 4, int(
        courseTable.iloc[courseTableindex, 2]), ThickonRight)  # Science credits
    # Theory credits
    sheet.write(row, 5, int(
        courseTable.iloc[courseTableindex, 3]), ContentStyle)
    sheet.write(row, 6, int(
        courseTable.iloc[courseTableindex, 4]), ThickonRight)  # Design credits
    # General credits
    sheet.write(row, 7, 0, ThickonRight)


def write_SpecialCourse(sheet, row, studentCourse, studentCourseindex):
    sheet.write(row, 0, str(
        studentCourse.iloc[studentCourseindex, 0]), ContentStyle)  # Semester
    # Course Name
    sheet.write(
        row, 1, studentCourse.iloc[studentCourseindex, 1], ContentStyle)
    # Require/Elective subject
    sheet.write(
        row, 2, studentCourse.iloc[studentCourseindex, 2], ThickonRight)
    # Mathematics credits
    sheet.write(row, 3, 0, ContentStyle)
    # Science credits
    sheet.write(row, 4, 0, ThickonRight)
    # Theory credits
    sheet.write(row, 5, 0, ContentStyle)
    sheet.write(row, 6, int(
        studentCourse.iloc[studentCourseindex, 3]), ThickonRight)  # Design credits
    # General credits
    sheet.write(row, 7, 0, ThickonRight)


def write_GeneralCourse(sheet, row, studentCourse, studentCourseindex):
    sheet.write(row, 0, str(
        studentCourse.iloc[studentCourseindex, 0]), ContentStyle)  # Semester
    # Course Name
    sheet.write(
        row, 1, studentCourse.iloc[studentCourseindex, 1], ContentStyle)
    # Require/Elective subject
    sheet.write(
        row, 2, studentCourse.iloc[studentCourseindex, 2], ThickonRight)
    # Mathematics credits
    sheet.write(row, 3, 0, ContentStyle)
    # Science credits
    sheet.write(row, 4, 0, ThickonRight)
    # Theory credits
    sheet.write(row, 5, 0, ContentStyle)
    # Design credits
    sheet.write(row, 6, 0, ThickonRight)
    sheet.write(row, 7, int(
        studentCourse.iloc[studentCourseindex, 3]), ThickonRight)  # General credits


def write_Capstone(sheet, row, CapstoneDict):
    for key in CapstoneDict.keys():
        sheet.write(row, 0, str(key[0]),
                    ContentStyle)  # semester
        # course name
        sheet.write(row, 1, str(key[1]), ContentStyle)
        sheet.write(row, 2, CapstoneDict.get(key)[
                    0], ThickonRight)  # require/elective
        # Mathematics credits
        sheet.write(row, 3, 0, ContentStyle)
        # Science credits
        sheet.write(row, 4, 0, ThickonRight)
        # Theory credits
        sheet.write(row, 5, 0, ContentStyle)
        sheet.write(row, 6, CapstoneDict.get(key)[
                    1], ThickonRight)  # Design credits
        sheet.write(row, 7, 0, ThickonRight)
        row = row + 1
    return row


def set_filename(Math, Science, MathSci, TheDesig):
    if (Math >= 9 and Science >= 9 and MathSci >= 32 and TheDesig >= 48):
        return '.xls'
    else:
        return '_Fail.xls'


def adjustWidth(sheet, coursecontent, index):
    maxlength = 0
    if (index == 0):
        maxlength = len(coursecontent.iloc[0, 1])
    elif (len(coursecontent.iloc[index, 1]) > adjustWidth(sheet, coursecontent, index-1)):
        maxlength = len(coursecontent)
        sheet.col(1).width = maxlength * 128
    return maxlength


def Pass(grade):
    pass_threshold = {'A', 'B', 'C'}
    waive_credit = {'抵免'}
    if (grade[0] in pass_threshold or grade in waive_credit):
        return True
    else:
        return False


Chartstyle = Excel_setting.set_ChartStyle()
ContentStyle = Excel_setting.set_ContentStyle()
TopStyle = Excel_setting.set_ThickLineOnTop()
TopRightStyle = Excel_setting.set_ThickLineOnTopRight()
ThickonRight = Excel_setting.set_ThickLineOnRight()


def main():
    path = getcwd()
    Allfile = listdir(path + '\data')

    # Read the student information
    df = pd(path + '\data\\' +
            Allfile[0], usecols='A,C,E,F:I')
    # Read the course table
    ref = pd(path + '\Rules\\' + 'course_table.xlsx', usecols='A:F')
    spc = pd(path + '\Rules\\' + 'Specialcourse.xlsx')
    cap = pd(path + '\Rules\\' + 'Capstone.xlsx')

    newStudenfile = xlwt.Workbook()
    sheet = newStudenfile.add_sheet("Sheet1", cell_overwrite_ok=True)
    Math = 0
    Science = 0
    MathSci = 0
    TheDesig = 0
    Fail = 0
    Success = 0
    Capdict = {}
    l = 0
    k = 4
    for i in range(len(df.iloc[:, 5])):  # Whole courses that students took
        # Read the course name from course_table.xlsx
        if ((df.iloc[i, 1] in ref['course'].tolist()) and Pass(str(df.iloc[i, 4]))):
            write_ETCourse(sheet, k, df, i, ref,
                           ref['course'].tolist().index(df.iloc[i, 1]))
            # MathSci = MathSci + \
            #     sum(ref.iloc[ref['course'].tolist().index(df.iloc[i, 1]), 1:3])
            Math = Math + \
                int(ref.iloc[ref['course'].tolist().index(df.iloc[i, 1]), 1])
            Science = Science + \
                int(ref.iloc[ref['course'].tolist().index(df.iloc[i, 1]), 2])
            MathSci = Math + Science
            TheDesig = TheDesig + \
                sum(ref.iloc[ref['course'].tolist().index(df.iloc[i, 1]), 3:5])
            k = k + 1
        # Read the course name from specialcourse.xlsx
        elif ((df.iloc[i, 1] in spc['course'].tolist()) and Pass(str(df.iloc[i, 4]))):
            write_SpecialCourse(sheet, k, df, i)
            TheDesig = TheDesig + df.iloc[i, 3]
            k = k + 1
        # Directly write the credits on table if course is a general course.
        elif ((df.iloc[i, 1] not in cap['course'].tolist()) and Pass(str(df.iloc[i, 4]))):
            write_GeneralCourse(sheet, k, df, i)
            k = k + 1
        # Store the capstone course in a dictionary
        elif ((df.iloc[i, 1] in cap['course'].tolist()) and Pass(str(df.iloc[i, 4]))):
            # (require/elective, credits)
            Capdict[(df.iloc[i, 0], df.iloc[i, 1])] = (
                str(df.iloc[i, 2]), int(df.iloc[i, 3]))
            TheDesig = TheDesig + df.iloc[i, 3]

        if ((i + 1 < len(df.iloc[:, 5])) and df.iloc[i, 5] != df.iloc[i + 1, 5]):
            # Present student ID is not equal to next student ID.
            # save the excel file
            k = write_Capstone(sheet, k, Capdict)
            set_Title(sheet, df.iloc[i, 5][-3:])
            string = '\\' + df.iloc[i, 5] + \
                set_filename(Math, Science, MathSci, TheDesig)
            set_End(sheet, k, 136)
            if (set_filename(Math, Science, MathSci, TheDesig) == '_Fail.xls'):
                Fail = Fail + 1
            else:
                Success = Success + 1
            MathSci = TheDesig = 0
            Math = 0
            Science = 0
            k = 4
            Capdict.clear()
            newStudenfile.save(
                path + '\out' + string)
            newStudenfile = xlwt.Workbook()
            sheet = newStudenfile.add_sheet("Sheet1", cell_overwrite_ok=True)

        if (i + 1 == len(df.iloc[:, 5])):
            # print out the final student form
            k = write_Capstone(sheet, k, Capdict)
            set_Title(sheet, df.iloc[i, 5][-3:])
            string = '\\' + df.iloc[i, 5] + \
                set_filename(Math, Science, MathSci, TheDesig)
            set_End(sheet, k, 136)
            newStudenfile.save(path + '\out' + string)
            if (set_filename(Math, Science, MathSci, TheDesig) == '_Fail.xls'):
                Fail = Fail + 1
            else:
                Success = Success + 1

    # for i in range(len(df.iloc[:, 5])):
    #     pass

    print("Success:"+str(Success))
    print("Fail:" + str(Fail))
    print("Pass rate:{:.2f}%".format(100*Success/(Success+Fail)))


if __name__ == "__main__":
    print("Welcome NTUST ECE Credits System !!!")
    main()
    input()
