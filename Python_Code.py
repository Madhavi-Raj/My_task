from pylatex import Document, Section, Subsection,MiniPage
from pylatex.utils import italic, bold
from pylatex import Document, Section, Subsection, Tabular,basic,position,utils
from pylatex import base_classes,MiniPage,Package,NoEscape,NewPage
import pylatex as pl

geometry_options = {"tmargin": "10mm", "lmargin": "10mm","margin": "10mm"}
doc = Document(geometry_options=geometry_options)
doc.packages.append(Package('opensans',options=["default"]))
doc.append(NoEscape(r'\fontsize{12}{15}\selectfont'))

import pandas as pd
df=pd.read_excel('Data for Daily course planner - output template.xlsx')
topics=[df.iloc[0][3],str(df.iloc[0][4])+".0"+" "+"Introduction"]
chapter_name=str(df.iloc[0][4])+" "+(df.iloc[0][5])
d={}
for i in df.iterrows():
    j=list(i[1])
    if(str(j[4])+" "+j[5]!=chapter_name):
        d[chapter_name]=topics
        chapter_name=str(j[4])+" "+j[5]
        topics=[j[3],str(j[4])+"."+str(j[7])+" "+"Introduction"]
    if(str(j[4])+" "+j[5]==chapter_name and str(j[4])+"."+str(j[7])+" "+j[8] not in topics):
        if(j[7]==0):
            continue
        topics.append(str(j[4])+"."+str(j[7])+" "+j[8])
    elif(str(j[4])+" "+j[5]==chapter_name):
        topics.append(str(j[4])+"."+str(j[7])+"."+str(int(j[9]))+" "+j[10])  
d[chapter_name]=topics
for i in d:
    doc.append(NoEscape(r'\begin{center}'))
    doc.append(NoEscape(r'\textbf{Learn Basics - LPS}'+' Course Release Schedule [PT2]'))
    doc.append(NoEscape(r'\end{center}'))
    doc.append(NoEscape(r'\par\noindent\rule{\textwidth}{0.4pt}'))
    h=r'\textbf{'+i+'}'
    doc.append(NoEscape(r'\begin{center}'))
    doc.append(NoEscape(h))
    doc.append(NoEscape(r'\end{center}'))
    subject="Science - Biology"
    class1=str(7)
    chaptertag=d[i][0]
    doc.append(NoEscape(r'\vspace{3mm}'))
    doc.append(NoEscape(r'\renewcommand*{\arraystretch}{2}'))
    with doc.create(Tabular('p{8.5cm}|p{8.5cm}')) as table:
        table.add_hline()
        table.add_row("Subject: "+subject,"Class: "+class1)
        table.add_hline()
        table.add_row("Chapter Tag : "+chaptertag,"Teacher Name : ")
    doc.append(NoEscape(r'\vspace{5mm}'))
    doc.append("\n") 
    doc.append('Start Date:________________________') 
    doc.append(NoEscape(r'\vspace{5mm}'))
    doc.append("\n")
    l=d[i][1:]
    doc.append(NoEscape(r'\renewcommand*{\arraystretch}{2}'))
    doc.append(NoEscape(r'\begin{tabular}{|p{2cm}|p{7cm}|p{2.5cm}|p{2.5cm}|p{3.5cm}|}'))
    doc.append(NoEscape(r'\hline'))
    doc.append(NoEscape(r'\centering{\textbf{S.No}} & \centering{\textbf{Topics}} & \centering{\textbf{Start Date}} & \centering{\textbf{End Date}} & \textbf{\centering{Remarks}}\\'))
    for i in l:
        doc.append(NoEscape(r'\hline'))
        h=i.split(" ",1)
        topic_no=h[0]
        topic_name=h[1]
        s=topic_no+"&"+topic_name+"&&&\\\\"
        raw_s = r'{}'.format(s)
        doc.append(NoEscape(raw_s))
    doc.append(NoEscape(r'\hline'))
    doc.append(NoEscape(r'\end{tabular}'))
    doc.append(NoEscape(r'\vspace{5mm}'))
    doc.append("\n") 
    doc.append('End Date:________________________') 
    doc.append(NoEscape(r'\vspace{5mm}'))
    doc.append("\n")
    doc.append(NewPage())  
doc.generate_pdf("final.sh")
