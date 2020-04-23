# -*- coding: utf-8 -*-
"""
Spyder Editor
"""

import xlwt
import time
start_time = time.perf_counter()
import pandas as pd
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 120)

from qgis.core import QgsDistanceArea, QgsPointXY
distance = QgsDistanceArea()
from pyproj import Transformer
transformer = Transformer.from_crs("epsg:31467", "epsg:25832", always_xy=True) ##gauss_krueger_coordinate zone 3 (31467), UTM zone 32N (25832)

from pathlib import Path
f = open(Path.home() / 'python32' / 'python_dir.txt', mode='r')
path = Path.joinpath(Path(r'C:'+f.readline()),'PT_data','VISUM_FAN.txt')
f = path.read_text().split('\n')

##connection to file
df_FAN = pd.read_excel(r'C:'+f[0], sheet_name=f[1])
df_VISUM = pd.read_excel(r'C:'+f[2], sheet_name=f[3])

wb = xlwt.Workbook()
ws = wb.add_sheet(f[5])
results = 'C:'+f[4]

##functions
def project(X,Y):
    X = str(X)
    X = X.replace(".","")
    X = X+"00000"
    X = int(X[:7])
    Y = str(Y)
    Y = Y.replace(".","")
    Y = Y+"00000"
    Y = int(Y[:7])
    x_out, y_out = transformer.transform(X, Y)
    return x_out, y_out

def dist_test(XY_FAN,XY_VISUM):
    if len(XY_FAN) == 0:
        return 99999, 0, 0
    xy_out = project(XY_FAN["GK-X"].iloc[0],XY_FAN["GK-Y"].iloc[0])
    point1 = QgsPointXY(XY_VISUM[6],XY_VISUM[7])
    point2 = QgsPointXY(xy_out[0],xy_out[1])
    dist = distance.measureLine(point1, point2)
    return dist, xy_out[0], xy_out[1]

def dist_next(VISUM,FAN):
    df_FAN = FAN.to_numpy()
    a = [1000,0,"",""]
    for i in df_FAN:
        XY = project(i[4],i[5])
        point1 = QgsPointXY(VISUM[6],VISUM[7])
        point2 = QgsPointXY(XY[0],XY[1])
        dist = distance.measureLine(point1, point2)
        if dist < a[0]:
            No = i[2]
            if i[0]=="H": No = i[1]
            if No==0: No = i[1]
            a = [dist,No,i[8],i[9]]
    return a

def writer(row,a,dist,hit):
    ws.write(row, a+7, dist[0])
    ws.write(row, a+5, dist[1])
    ws.write(row, a+6, dist[2])
    
    #more values
    No = hit["Master"].iloc[0]
    if hit["Typ"].iloc[0]=="H": No = hit["HST-Nr"].iloc[0]
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
    # if i[0]!=6680075:continue
    # hh
    for a in range(VISUM_columns):
        ws.write(row, a, i[a])
    ##part of DHID
    hit = df_FAN[df_FAN['DHID'].str.contains(":"+str(i[2])+":")].head(1)
    dist = dist_test(hit,i)
    if len(hit)>0 and dist[0]<2000:
        writer(row,a,dist,hit)
        row+=1
        continue
    ##subpart of DHID
    hit = df_FAN[df_FAN['DHID_2'].str.contains(":"+str(i[2])+":")].head(1)
    dist = dist_test(hit,i)
    if len(hit)>0 and dist[0]<2000:
        writer(row,a,dist,hit)
        row+=1
        continue
    ##identical name and location
    hit = df_FAN.loc[df_FAN['Name + Ort'] == i[1].replace(',','')].head(1)
    dist = dist_test(hit,i)
    if len(hit)>0 and dist[0]<1000:
        writer(row,a,dist,hit)
        row+=1
        continue
    ##identical name
    try: name = i[3].split(", ")[1]
    except: name = i[3]
    hit = df_FAN.loc[df_FAN['Name'] == name.replace("Abzw.","Abzweigung")].head(1)
    dist = dist_test(hit,i)
    if len(hit)>0 and dist[0]<500:
        writer(row,a,dist,hit)
        row+=1
        continue
    ##closest stop
    next_No = dist_next(i,df_FAN)
    if next_No[1] != 0 and next_No[0]<200:
        ws.write(row, a+1, int(next_No[1]))
        ws.write(row, a+2, next_No[2])
        ws.write(row, a+3, next_No[3])
        ws.write(row, a+7, next_No[0])
        row+=1
        continue
    row+=1

##output
wb.save(results)
print("Script finished after: "+str(int(time.perf_counter())-int(start_time))+" seconds")