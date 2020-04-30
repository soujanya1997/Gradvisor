import numpy as np
import pandas as pd
JDWB = pd.read_excel("Jobs.xls")
StudentWB = pd.read_excel("Technical_Consolidation.xls")
Result = np.zeros((JDWB.shape[0]-1, StudentWB.shape[0]-2))
print(Result.shape)
for i in range(JDWB.shape[0]-1):
for j in range(1, StudentWB.shape[0]-1):
if(type(JDWB.iloc[i]['Technologies'])!=float and type(StudentWB.iloc[j]['TechnologiesEx'])!=float):
techs = JDWB.iloc[i]['Technologies'].split(", ")
techsStudent = StudentWB.iloc[j]['TechnologiesEx'].split(", ")
common = [value for value in techsStudent if value in techs]
ratio_technologies = len(common)/len(techs)
else:
ratio_technologies = 0
if(type(JDWB.iloc[i]['Tools'])!=float and type(StudentWB.iloc[j]['ToolsEx'])!=float):
techs = JDWB.iloc[i]['Tools'].split(", ")
techsStudent = StudentWB.iloc[j]['ToolsEx'].split(", ")
common = [value for value in techsStudent if value in techs]
ratio_tools = len(common)/len(techs)
else:
ratio_tools = 0
if(type(JDWB.iloc[i]['Languages'])!=float and type(StudentWB.iloc[j]['Ex'])!=float):
techs = JDWB.iloc[i]['Languages'].split(", ")
techsStudent = StudentWB.iloc[j]['LanguagesEx'].split(", ")
common = [value for value in techsStudent if value in techs]
ratio_languages = len(common)/len(techs)
else:
ratio_languages = 0
ratio = (ratio_technologies+ratio_tools+ratio_languages)/3
Result[i][j-1] = ratio
print(Result)