import xlwt
from xlwt import Workbook
import sys
import easygui

# Usage:
# 
# .py file can be anywhere, but will create the excel file 
# in its current directory
# 
# at prompt use gui to select the .txt file with the eprime
# data, and then enter the name you want the excel file to be called 
# (note you do not need to include .xls, that is added automatically)
# https://github.com/sschott20/AAT-txt-parser-to-.xls
# Alex


def main():
# choose file using easygui, output file  
    INPUTFILE = easygui.fileopenbox("Input .txt file with Eprime data")
    OUTPUTFILE = easygui.enterbox("Enter output filename")
    print (INPUTFILE)
    if INPUTFILE == None or OUTPUTFILE == None:
        print ("Must select file location and destination")
        return 1
    
# part 1 reading from textfile with weird utf encoding 

    read_file = open(INPUTFILE, "r", encoding="utf-16-le")
    write_file = open("write_file.txt", "w")

    trials = read_file.read().split("\n")
    # creates a separate .txt file to to simplify transfer to excel file 
    for i in range(1, len(trials)):
        trials[i] = trials[i].lstrip()
        if trials[i] == "Level: 3":
            for j in range(i, len(trials)):
                trials[j]= trials[j].lstrip()
                if trials[j].startswith("StimulusContent:") or trials[j].startswith("CorrectResponse:") or trials[j].startswith("RT:") or trials[j].startswith("ACC:"):
                    write_file.write(trials[j])
                    write_file.write("\n")
                elif trials[j] == "*** LogFrame End ***":
                    write_file.write("\n \n")
                elif trials[j] == "Level: 2":
                    break
            break
    write_file.close()

# part 2 writing to excel

    write_file = open("write_file.txt", "r")
    wb = Workbook()
    sheet1 = wb.add_sheet("Sheet 1")

    boldFont = xlwt.easyxf('font: name Arial, bold on, height 200;')
    regFont = xlwt.easyxf('font: name Arial, bold off, height 200;')

    # collumn titles 
    sheet1.write(0, 0, "VapingPull", boldFont)
    sheet1.write(0, 2, "VapingPush", boldFont)
    sheet1.write(0, 4, "NeutralPull", boldFont)
    sheet1.write(0, 6, "NeutralPush", boldFont)

    for i in range(8):
        sheet1.write(1, i, "RT" if i % 2 == 0 else "ACC", boldFont)
    trials = write_file.read().split()

    # number of elements in each new array
    n = 8

    # separates each chunk of the write file by making 
    # sub arrays of size n, has to be adjusted if the
    # .txt format changes
    trials = [trials[i * n: (i + 1) * n] for i in range((len(trials) + n - 1) //n)]

    col = 0
    row = 0
    bottomRow = [2, 2, 2, 2]
    # writing the error and reaction time to excel file
    for trial in trials: 
        if trial[trial.index("StimulusContent:") + 1] == "regular":
            col += 4
        if trial[trial.index("CorrectResponse:") + 1] == "Push":
            col += 2
        rowInt = int(col/2)
        
        # trial[x] part must be changed if the order 
        # that error/rt/stimulusContent/correctResponse appear
        # in the .txt file change
        # ['ErrorCount:', '1', 'StimulusContent:', 'Neutral', 'CorrectResponse:', 'Pull', 'RT:', '828']
        # could optimize to recognize each class, not just array index
        sheet1.write(bottomRow[rowInt], col, int(trial[trial.index("RT:") + 1]), regFont)
        sheet1.write(bottomRow[rowInt], col + 1, int(trial[trial.index("ACC:") + 1]), regFont)

        bottomRow[rowInt] += 1
        col = 0
        row += 1

    wb.save(OUTPUTFILE + ".xls")

    write_file.close()
    read_file.close()

if __name__ == "__main__":
    main()
