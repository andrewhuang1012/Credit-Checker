from os import getcwd
from os import listdir
import os
import shutil
import pandas as pd
import math
path = getcwd()
data = listdir(path + '\data')
# Read the student information, (StudentID, Class, Rank)
df = pd.read_excel(path + '\data\\' + data[0], usecols='H,J,K')
Repre = {}  # stuentID : rank
presentClass = df.iloc[0, 1]
print("The following are representatives:")
for i in range(len(df.iloc[:, 0])):
    # student without rank or not in ECE
    if (math.isnan(df.iloc[i, 2]) or int(df.iloc[i, 0][4]) != 0):
        continue

    # rank is a value, instead of NaN
    if (df.iloc[i, 1] == presentClass):
        Repre[str(df.iloc[i, 0])] = int(df.iloc[i, 2])

    # The next student's class is not present student's one.
    if (i + 1 > len(df.iloc[:, 0]) or df.iloc[i, 1] != presentClass):
        presentClass = df.iloc[i + 1, 1]
        sort_Repre = sorted(Repre.items(), key=lambda d: d[1])

        if (len(sort_Repre) == 0):
            continue

        print(sort_Repre[0])
        print(sort_Repre[1])
        print(sort_Repre[len(sort_Repre) // 2])
        print(sort_Repre[(len(sort_Repre) // 2) - 1])
        print(sort_Repre[len(sort_Repre) - 1])
        print(sort_Repre[len(sort_Repre) - 2])
        if os.path.isfile(path + '\out\\' + sort_Repre[0][0] + ".xls"):
            shutil.copy(path + '\out\\' + sort_Repre[0][0] + ".xls", path +
                        '\\representatives\\' + sort_Repre[0][0] + "_H1.xls")
        else:
            shutil.copy(path + '\out\\' + sort_Repre[2][0] + ".xls", path +
                        '\\representatives\\' + sort_Repre[2][0] + "_H1.xls")
        if os.path.isfile(path + '\out\\' + sort_Repre[1][0] + ".xls"):
            shutil.copy(path + '\out\\' + sort_Repre[1][0] + ".xls", path +
                        '\\representatives\\' + sort_Repre[1][0] + "_H2.xls")
        else:
            shutil.copy(path + '\out\\' + sort_Repre[3][0] + ".xls", path +
                        '\\representatives\\' + sort_Repre[3][0] + "_H2.xls")
        if os.path.isfile(path + '\out\\' + sort_Repre[len(sort_Repre) // 2][0] + ".xls"):
            shutil.copy(path + '\out\\' + sort_Repre[len(sort_Repre) // 2][0] + ".xls", path +
                        '\\representatives\\' + sort_Repre[len(sort_Repre) // 2][0] + "_M1.xls")
        else:
            shutil.copy(path + '\out\\' + sort_Repre[(len(sort_Repre) // 2)+1][0] + ".xls", path +
                        '\\representatives\\' + sort_Repre[(len(sort_Repre) // 2) + 1][0] + "_M1.xls")
        if os.path.isfile(path + '\out\\' + sort_Repre[(len(sort_Repre) // 2) - 1][0] + ".xls"):
            shutil.copy(path + '\out\\' + sort_Repre[(len(sort_Repre) // 2) - 1][0] + ".xls", path +
                        '\\representatives\\' + sort_Repre[(len(sort_Repre) // 2) - 1][0] + "_M2.xls")
        else:
            shutil.copy(path + '\out\\' + sort_Repre[(len(sort_Repre) // 2)-2][0] + ".xls", path +
                        '\\representatives\\' + sort_Repre[(len(sort_Repre) // 2) - 2][0] + "_M2.xls")
        if os.path.isfile(path + '\out\\' + sort_Repre[len(sort_Repre) - 1][0] + ".xls"):
            shutil.copy(path + '\out\\' + sort_Repre[len(sort_Repre) - 1][0] + ".xls", path +
                        '\\representatives\\' + sort_Repre[len(sort_Repre) - 1][0] + "_B1.xls")
        else:
            shutil.copy(path + '\out\\' + sort_Repre[len(sort_Repre) - 3][0] + ".xls", path +
                        '\\representatives\\' + sort_Repre[len(sort_Repre)-3][0] + "_B1.xls")
        if os.path.isfile(path + '\out\\' + sort_Repre[len(sort_Repre) - 2][0] + ".xls"):
            shutil.copy(path + '\out\\' + sort_Repre[len(sort_Repre) - 2][0] + ".xls", path +
                        '\\representatives\\' + sort_Repre[len(sort_Repre) - 2][0] + "_B2.xls")
        else:
            shutil.copy(path + '\out\\' + sort_Repre[len(sort_Repre) - 4][0] + ".xls", path +
                        '\\representatives\\' + sort_Repre[len(sort_Repre)-4][0] + "_B2.xls")
        Repre.clear()
        Repre[str(df.iloc[i, 0])] = int(df.iloc[i, 2])

print("Finished !!!")
input()
