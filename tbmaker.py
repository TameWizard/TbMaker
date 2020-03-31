# -*- coding: utf-8 -*-
"""
Created on Mon Mar 30 10:59:00 2020

@author: TameWizard
"""

import docx;
import pandas as pd;
import os;

numbs = "0123456789"
stop = "),;"

def getText(filename):
    doc = docx.Document(filename);
    fullText = [];
    for para in doc.paragraphs:
        fullText.append(para.text);
    return fullText;

def makeTB(text):
    text.remove(text[0]);
    data = {'Ru': [], 'En': []};
    for i in text:
        isEn = False;
        work = i.split(" ");
        resRu = '';
        resEn = '';
        for e in work:
            if e == '':
                continue
            elif isEn == False and e != "-":
                if e[-1] not in stop and e[0] in numbs:
                    if e == "1.":
                        continue
                    else:
                        break
                elif e[0] != "(":
                    resRu = resRu + e + " ";
                else:
                    continue
            elif isEn == True and e != "-":
                if e[0] == "(":
                    continue
                elif e[-1] in stop and e[0] not in numbs:
                    resEn = resEn + e[:-1] + " ";
                    break    
                elif e[-1] not in stop and e[0] in numbs:
                    if e == "1.":
                        continue
                    else:
                        break
                else:
                    resEn = resEn + e + " ";
            elif e == "-":
                isEn = True;
                continue
        data['Ru'].append(resRu[:-1])
        data['En'].append(resEn[:-1])
    tb = pd.DataFrame.from_dict(data, orient='columns');
    return(tb)

directory = "C:\\Users\\777\\Desktop\\folder\\"
direc2 = "C:\\Users\\777\\Desktop\\folder2\\"

for filename in os.listdir(directory):
    if filename.endswith(".docx"):
        text = getText("{}{}".format(directory, filename))
        tb = makeTB(text)
        tb.dropna()
        tb = tb[tb.Ru != ""]
        tb = tb[tb.En != ""]
        print(tb)
        tb.to_excel("{}{}.xlsx".format(direc2, filename.strip()), index=False)