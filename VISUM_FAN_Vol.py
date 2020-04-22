# -*- coding: utf-8 -*-
"""
Created on Sat Apr  4 21:32:49 2020

@author: marcus
"""

import xlwt
import time
start_time = time.perf_counter()
import math
import pandas as pd
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 120)

from pathlib import Path
f = open(Path.home() / 'python32' / 'python_dir.txt', mode='r')
path = Path.joinpath(Path(r'C:'+f.readline()),'PT_data','VISUM_Vol.txt')
f = path.read_text().split('\n')

##connection to file
df_Vol = pd.read_excel(r'C:'+f[0], sheet_name = None)
df_VISUM_FAN_Nr = pd.read_excel(r'C:'+f[1], sheet_name='VISUM_FAN').to_numpy()
df_links = pd.read_excel(r'C:'+f[2], sheet_name='Strecken')

wb = xlwt.Workbook()
ws = wb.add_sheet("VISUM_Vol", cell_overwrite_ok=True)
results ='C:'+f[3]

##functions
def string_list(df,column):
    df[column] = df[column].astype(str)
    df[column]= df[column].str.replace(".",",")
    df[column]= df[column].str.split(",").apply(lambda x: [int(i) for i in x])
    return df

##preparation
#VISUM_links
df_links = df_links.fillna(0)
df_links = string_list(df_links,"von")
df_links = string_list(df_links,"nach")
df_links['FAN_von'] = df_links['von']
df_links['FAN_nach'] = df_links['nach']
df_links['HHA_von'] = df_links['von']
df_links['HHA_nach'] = df_links['nach']

##VISUM_FAN
for i in df_VISUM_FAN_Nr:
    if math.isnan(i[8]) == False:
        df_links['FAN_von'] = df_links['FAN_von'].apply(lambda x: [int(i[8]) if e==i[0] else e for e in x])   
        df_links['FAN_nach'] = df_links['FAN_nach'].apply(lambda x: [int(i[8]) if e==i[0] else e for e in x])
    if math.isnan(i[15]) == False:
        df_links['HHA_von'] = df_links['HHA_von'].apply(lambda x: [int(i[15]) if e==i[0] else e for e in x])   
        df_links['HHA_nach'] = df_links['HHA_nach'].apply(lambda x: [int(i[15]) if e==i[0] else e for e in x])


df_links['FAN_von'] = [list(set(b).difference(set(a))) for a, b in zip(df_links.von, df_links.FAN_von)]
df_links['FAN_nach'] = [list(set(b).difference(set(a))) for a, b in zip(df_links.nach, df_links.FAN_nach)]
df_links['HHA_von'] = [list(set(b).difference(set(a))) for a, b in zip(df_links.von, df_links.HHA_von)]
df_links['HHA_nach'] = [list(set(b).difference(set(a))) for a, b in zip(df_links.nach, df_links.HHA_nach)]

print("--join der FAN_Nummern an die Strecken erfolgreich--")

##insert column_names and values to list
ws.write(0, 0, "Nr")
ws.write(0, 1, "VONKNOTNR")
ws.write(0, 2, "NACHKNOTNR")
ws.write(0, 3, "Vol")
ws.write(0, 4, "Vol_Bus")
ws.write(0, 5, "Vol_U")
ws.write(0, 6, "Vol_A")
ws.write(0, 7, "Vol_S")
ws.write(0, 8, "Linien")

t = []  ##new python list to fill in line volumes
for i in df_links.itertuples():
    t.append([i.strecke,i.VONKNOTNR,i.NACHKNOTNR,0,0,0,0,0,""])
 
n = 0
file = open('C:'+f[4],'w')
##Volumes
for name, sheet in df_Vol.items():
    # if "_Bus" in name:continue
    if "Kanten" in name:
        print("--beginne mit: "+name+"--")
        df_Vol_line = df_Vol[name]
        df_Vol_line = df_Vol_line.rename(columns={"Belastung MF": "Belastung_MF", "Von Haltestelle": "VonHst", "Nach Haltestelle":"NachHst"})
        for i in df_Vol_line.itertuples():
            n+=1
            # print("row: "+str(n))
            if n==100:pass
            if "_U_" in name: vol = df_links[(df_links.apply(lambda x: i.Von in x.HHA_von, axis=1))&(df_links.apply(lambda x: i.Nach in x.HHA_nach, axis=1))]
            else: vol = df_links[(df_links.apply(lambda x: i.Von in x.FAN_von, axis=1))&(df_links.apply(lambda x: i.Nach in x.FAN_nach, axis=1))]
            if len(vol)==0:
                file.write(str(i.Von)+"; "+str(i.Nach)+"; "+i.VonHst+"; "+i.NachHst+"; "+str(i.Linien)+"; "+str(i.Belastung_MF)+"\n")
                print (i.Von, i.Nach, i.VonHst, i.NachHst, i.Linien, i.Belastung_MF)
            for row_t in vol.index:
                t[row_t][3] = t[row_t][3]+i.Belastung_MF
                if "_Bus" in name: t[row_t][4] = t[row_t][4]+i.Belastung_MF
                if "_U_" in name: t[row_t][5] = t[row_t][5]+i.Belastung_MF
                if "_AKN" in name: t[row_t][6] = t[row_t][6]+i.Belastung_MF
                if "_S_" in name: t[row_t][7] = t[row_t][7]+i.Belastung_MF
                try: t[row_t][8] = t[row_t][8]+", "+str(i.Linien)
                except:pass
        
        row = 1
        for i in df_links.itertuples():    
            ws.write(row, 0, i.strecke)
            ws.write(row, 1, i.VONKNOTNR)
            ws.write(row, 2, i.NACHKNOTNR)
            ws.write(row, 3, t[row-1][3])
            ws.write(row, 4, t[row-1][4])
            ws.write(row, 5, t[row-1][5])
            ws.write(row, 6, t[row-1][6])
            ws.write(row, 7, t[row-1][7])
            ws.write(row, 8, t[row-1][8])
            row+=1
    
##output
file.close()
wb.save(results)
print("Script finished after: "+str(int(time.perf_counter())-int(start_time))+" seconds")