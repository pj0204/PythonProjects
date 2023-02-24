import pandas as pd
from openpyxl import load_workbook
import string
import xlsxwriter
import numpy

df = pd.read_excel('Distrib.xlsx')

df = df.loc[:,df.columns != 'PRODUCTO']
tienda = list(df.columns.values)

#tienda =["cum","qui","cal","dom","esm","ala","ova","lpa","sam","eco","sub","ira","cer","mai","vic","mac","pto","bel","rej"]
for i in tienda [1:]:
    writer = pd.ExcelWriter("G:/Mi unidad/Naturland-Monitor/PythonProjects/OC/oc_"+i+".xlsx",engine="xlsxwriter")
    df1 = df[["CODIGO",i]]
    df2 =df1[df1!=0]
    df3 = df2.dropna(subset=[i])
    df4 = df3.sort_values(by = i, ascending = False)
    df4.to_excel(writer,sheet_name="Hoja1",index=False)
    writer.close()
    




