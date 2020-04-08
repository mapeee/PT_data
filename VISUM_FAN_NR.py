# -*- coding: utf-8 -*-
"""
Spyder Editor
"""

import xlwt
import re
import time
start_time = time.perf_counter()
import pandas as pd
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 120)

from qgis.core import QgsDistanceArea, QgsPointXY
distance = QgsDistanceArea()

from pyproj import Proj, transform
input_proj = Proj(init='epsg:31467') ##gauss_krueger_coordinate zone 3
output_proj = Proj(init='epsg:25832') ##UTM zone 32N

from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
f = open(path, mode='r')
for i in f: path = i
path = Path.joinpath(Path(r'C:'+path),'PT_data','VISUM_FAN.txt')
f = path.read_text()
f = f.split('\n')

##connection to file
df_FAN = pd.read_excel(r'C:'+f[0], sheet_name=f[1])
df_VISUM = pd.read_excel(r'C:'+f[2], sheet_name=f[3])

wb = xlwt.Workbook()
ws = wb.add_sheet(f[5])
results = 'C:'+f[4]

##functions
def dist_test(XY_FAN,XY_VISUM):
    if len(XY_FAN) == 0:
        return 99999, 0, 0
    X = str(XY_FAN["GK-X"].iloc[0])
    X = X.replace(".","")
    X = X+"00000"
    X = int(X[:7])
    Y = str(XY_FAN["GK-Y"].iloc[0])
    Y = Y.replace(".","")
    Y = Y+"00000"
    Y = int(Y[:7])
    x_out, y_out = transform(input_proj, output_proj, X, Y)
 
    point1 = QgsPointXY(XY_VISUM[6],XY_VISUM[7])
    point2 = QgsPointXY(x_out,y_out)
    dist = distance.measureLine(point1, point2)
    return dist, x_out, y_out

def writer(row,a,dist,hit):
    ws.write(row, a+7, dist[0])
    ws.write(row, a+5, dist[1])
    ws.write(row, a+6, dist[2])
    
    #more values
    No = hit["Master"].iloc[0]
    if hit["Typ"].iloc[0]=="H": No = hit["HST-Nr"].iloc[0]
    if hit["Typ"].iloc[0]=="M": No = hit["Master"].iloc[0]
    if No==0: No = hit["HST-Nr"].iloc[0]
    
    ws.write(row, a+1, int(No))
    ws.write(row, a+2, hit["Name + Ort"].iloc[0])
    ws.write(row, a+3, hit["Alter Name"].iloc[0])
    ws.write(row, a+4, hit["Linien"].iloc[0])
    return

##preparation
#FAN
undrop = ["Typ","HST-Nr","Master","DHID","GK-X","GK-Y","Name","Ortsname","Name + Ort","Alter Name", "Linien"]
fields = list(df_FAN.columns)
for i in undrop: fields.remove(i)
df_FAN = df_FAN.drop(fields,axis= 1)
df_FAN = df_FAN.fillna(0)
df_FAN['DHID_2'] = df_FAN['DHID'].astype(str)
for i in range(5):
    df_FAN["DHID_2"] = df_FAN["DHID_2"].str.replace(':'+str(i+1)+':',str(i+1)+':')
df_FAN['DHID'] = df_FAN['DHID'].astype(str)+':'
df_FAN['Name'] = df_FAN['Name'].astype(str)
df_FAN['Name + Ort'] = df_FAN['Name + Ort'].astype(str)
df_FAN["Name + Ort"] = df_FAN["Name + Ort"].str.replace(',','')

##insert column_names
column = 0
VISUM_columns = len(df_VISUM.columns)
for i in df_VISUM.columns:
    ws.write(0, column, i) 
    column+=1
ws.write(0, column, "FAN_Nr")
ws.write(0, column+1, "FAN_Name_Ort")
ws.write(0, column+2, "FAN_Alter_Name")
ws.write(0, column+3, "FAN_Linien")
ws.write(0, column+4, "FAN_X")
ws.write(0, column+5, "FAN_Y")
ws.write(0, column+6, "Distance")

##insert_values
df_VISUM = df_VISUM.to_numpy()
row = 1
for i in df_VISUM:    
    for a in range(VISUM_columns):
        ws.write(row, a, i[a])
    #hits
    hit = df_FAN[df_FAN['DHID'].str.contains(":"+str(i[2])+":")].head(1)
    dist = dist_test(hit,i)
    if len(hit)>0 and dist[0]<2000:
        writer(row,a,dist,hit)
        row+=1
        continue
    hit = df_FAN[df_FAN['DHID_2'].str.contains(":"+str(i[2])+":")].head(1)
    dist = dist_test(hit,i)
    if len(hit)>0 and dist[0]<2000:
        writer(row,a,dist,hit)
        row+=1
        continue  
    hit = df_FAN[df_FAN['Name + Ort'].str.contains(re.escape(i[1].replace(',','')))].head(1)
    dist = dist_test(hit,i)
    if len(hit)>0 and dist[0]<1000:
        writer(row,a,dist,hit)
        row+=1
        continue
    try: name = i[3].split(", ")[1]
    except:
        name = i[3]
        row+=1
        continue
    name = name.replace("Abzw.","Abzweigung")
    hit = df_FAN[df_FAN['Name'].str.contains(re.escape(name))].head(1)
    dist = dist_test(hit,i)
    if len(hit)>0 and dist[0]<500:
        writer(row,a,dist,hit)
        row+=1
        continue
    row+=1

##output
wb.save(results)
print("Script finished after: "+str(int(time.perf_counter())-int(start_time))+" seconds")