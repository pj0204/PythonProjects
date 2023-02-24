import pandas as pd
from openpyxl import load_workbook
import string
import xlsxwriter
import numpy

df = pd.read_excel('Distrib_Terra.xlsx')

df1 = df.iloc[:,[0,1,3,6,21,37,53,69,85,101,117,133,149,165,181,197,213,229,245,261,277,293,309]]

df2= df1.fillna(value=0)

df2.columns=['CODIGO', 'PRODUCTO', 'DIST', 'NVAS', 'ECO', 'SUB', 'SAM', 'CUM',
       'QUI', 'LPA', 'ALA', 'ESM', 'DOM', 'OVA', 'CAL', 'IRA', 'CER',
       'MAI', 'VIC', 'MAC', 'PTO', 'BEL', 'REJ']


writer = pd.ExcelWriter("G:/Mi unidad/Naturland-Monitor/PythonProjects/Distribuciones/Distribuciones_TERRA.xlsx",engine="xlsxwriter")
df2.to_excel(writer,sheet_name="Hoja1",index=False)
writer.close()



