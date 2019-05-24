# -*- coding: utf-8 -*-

# This script will extract calculated fields and parameters
# from a Tableau workbook and output a CSV file with four columns:
# Calculation Name, Remote Name, Formula, and Comment
# The comment is from the calculated field when a // is used
# Comments should be placed before the calculation to be extracted properly

#==========================================
# Title:  Tableau TWB Calculation Extractor
# Author: Ray Randall, github.com/rarandall
# Date:   24 May 2019
#==========================================

import easygui
import xml.etree.ElementTree as ET
import os
import pandas as pd

# prompt user for twb file

file = easygui.fileopenbox(filetypes=['*.twb'])

# parse the twb file
tree = ET.parse(file)
root = tree.getroot()

# print(root.tag)
# print(root.attrib)

# create a dictionary of name and tableau generated name

calcDict = {}

for item in root.findall('.//column[@caption]'):
    if item.find(".//calculation") is None:
        continue
    else:
        calcDict[item.attrib['name']] = '[' + item.attrib['caption'] + ']'

# print(calcDict)


# list of calc's name, tableau generated name, and calculation/formula
calcList = []

for item in root.findall('.//column[@caption]'):
    if item.find(".//calculation") is None:
        continue
    else:
        if item.find(".//calculation[@formula]") is None:
            continue
        else:
            calc_caption = '[' + item.attrib['caption'] + ']'
            calc_name = item.attrib['name']
            calc_raw_formula = item.find(".//calculation").attrib['formula']
            calc_comment = ''
            calc_formula = ''
            for line in calc_raw_formula.split('\r\n'):
                if line.startswith('//'):
                    calc_comment = calc_comment + line + ' '
                else:
                    calc_formula = calc_formula + line + ' '
            for name, caption in calcDict.items():
                calc_formula = calc_formula.replace(name, caption)

            calc_row = (calc_caption, calc_name, calc_formula, calc_comment)
            calcList.append(list(calc_row))

# print(calcList)

# convert the list of calcs into a data frame
data = calcList

data = pd.DataFrame(data, columns=['Name', 'Remote Name', 'Formula', 'Comment'])

# remove duplicate rows from data frame
data = data.drop_duplicates(subset=None, keep='first', inplace=False)

print(data)

# export to csv
# get the name of the file

base = os.path.basename(file)
os.path.splitext(base)
filename = os.path.splitext(base)[0]

data.to_csv(filename + '.csv')
