from pylatex import Document, Section, Subsection,MiniPage
from pylatex.utils import italic, bold
from pylatex import Document, Section, Subsection, Tabular,basic,position,utils
from pylatex import base_classes,MiniPage,Package,NoEscape,NewPage
doc.packages.append(Package('adjustbox'))
geometry_options = {"tmargin": "10mm", "lmargin": "10mm","margin": "10mm"}
doc = Document(geometry_options=geometry_options,font_size =" ")


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
    doc.append("\n")
    doc.append(NoEscape(r'\par\noindent\rule{\textwidth}{0.4pt}'))
    doc.append("\n")
    h=r'\textbf{'+i+'}'
    doc.append(NoEscape(h))
    doc.append("\n")
    subject="Science - Biology"
    class1=str(7)
    chaptertag=d[i][0]
    doc.append(NoEscape(r'\vspace{2mm}'))
    doc.append("\n")
    with doc.create(Tabular('p{8.5cm}|p{8.5cm}')) as table:
            table.add_hline()
            table.add_row("subject: "+subject,"Class: "+class1)
            table.add_empty_row()
            table.add_hline()
            table.add_row("Chapter Tag : "+chaptertag,"Teacher Name : ")
            table.add_empty_row()
    doc.append("\n")
    doc.append(NoEscape(r'\end{center}'))
    doc.append('Start Date:________________________') 
    doc.append(NoEscape(r'\vspace{3mm}'))
    doc.append("\n")
    l=d[i][1:]
    with doc.create(Tabular('|p{1.5cm}|p{7cm}|p{2.5cm}|p{2.5cm}|p{3.5cm}|')) as table:
            table.add_hline()
            table.add_row("S.No","Topics","Start Date","End Date","Remarks")
            table.add_empty_row()
            for i in l:
                table.add_hline()
                h=i.split(" ",1)
                topic_no=h[0]
                topic_name=h[1]
                table.add_row(h[0],h[1],"","","")
                table.add_empty_row()
            table.add_hline()
    doc.append(NewPage())
doc.generate_pdf('finalsh3', clean_tex=False)
