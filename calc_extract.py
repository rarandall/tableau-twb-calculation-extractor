# -*- coding: utf-8 -*-

# prompt user for twb file

import easygui

file = easygui.fileopenbox(filetypes=['*.twb'])

import xml.etree.ElementTree as ET

tree = ET.parse(file)
root = tree.getroot()

# print(root.tag)
# print(root.attrib)

calcList = []
calcDict = {}

# create a dictionary of name and tableau generated name

for item in root.findall('.//column[@caption]'):
    if item.find(".//calculation") is None:
        continue
    else:
        calcDict[item.attrib['name']] = '[' + item.attrib['caption'] + ']'

# print(calcDict)


# list of calc's name, tableau generated name, and calculation/formula

for item in root.findall('.//column[@caption]'):
    if item.find(".//calculation") is None:
        continue
    else:
        calc_caption = '[' + item.attrib['caption'] + ']'
        calc_name = item.attrib['name']
        calc_formula = item.find(".//calculation").attrib['formula']
        for name, caption in calcDict.items():
            calc_formula = calc_formula.replace(name, caption)
            calc_formula = calc_formula.replace('\r', ' ')
            calc_formula = calc_formula.replace('\n', ' ')

        calc_row = (calc_caption, calc_formula)
        calcList.append(list(calc_row))

# print(calcList)

# convert the list of calcs into a data frame
import pandas as pd

data = calcList

data = pd.DataFrame(data, columns=['Calculation Name', 'Calculation Formula'])

# remove duplicate rows from data frame
data = data.drop_duplicates(subset=None, keep='first', inplace=False)

# print(data)

# export to csv

data.to_csv(r'calculationList.csv')
