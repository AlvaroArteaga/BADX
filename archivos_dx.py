import pandas as pd
import glob
import os
from pathlib import Path
from openpyxl import load_workbook
from pandas import ExcelWriter
import numpy as np
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)

hojas = {
'Regulado SPD' : 'regulados',
'Regulado SPD RE' : 'regulados',
'Regulado Reconv' : 'regulados_reconv',
'RE	Regulado Reconv' : 'regulados_reconv',
'Clientes libres propios SPD' : 'clientes',
'Cli peaje otras Dx horario' : 'clientes',
'Clientes libres propios SPD SN' : 'clientes',
'Clientes de peaje Dx horario' : 'clientes',
'Clientes de peaje Dx hora SNOT' : 'clientes',
'Clientes de peaje Dx horario SN' : 'clientes',
'PMGD SPD' : 'pmgds',
'PMGD SPD_SF': 'pmgds',
'FR SPD' : 'pmgds_fr',
'Cabecera Alimentador con PMGD' : 'pmgds_alimentadores',
'Alimentadores' : 'pmgds_alimentadores',
'Detalle Cli libres propios SPD': 'clientes_detalle',
'Detalle cli peaje otras Dx': 'clientes_detalle',
'Detalle clientes de peaje Dx': 'clientes_detalle',
'Detalle PMGD' : 'pmgds_detalle',
'Grupos de Contratos' : 'contratos',
'Grupos de contratos ' : 'contratos',
'Grupo de contratos' : 'contratos',
'Grupo de Contratos' : 'contratos',
'Grupo de Contratos RE' : 'contratos',
'Grupos de Contratos RE' : 'contratos',
'Grupos de contratos' : 'contratos',
'Nuevo Regulado SPD' : 'regulados_nuevos',
'Nuevo Regulado Reconv' : 'regulados_reconv_nuevos',
'Nuevo Cli libres propios SPD' : 'clientes_nuevos',
'Nuevo Cli peaj otras Dx horario' :  'clientes_nuevos',
'Nuevo Detall Cli libr prop SPD' : 'clientes_detalle_nuevos',
'Nuevo Detall cli peaje otras Dx' : 'clientes_detalle_nuevos',
'Nuevo PMGD SPD' : 'pmgds_nuevos',
'PMGD SPD SIN NUEVOS' : 'pmgds_nuevos',
'Nuevo FR SPD' : 'pmgds_fr_nuevos',
'FR SPD SIN NUEVOS' : 'pmgds_fr_nuevos',
'Nuevo Detalle cli peaje Dx' : 'clientes_detalle_nuevos',
'Nuevo Cli peaje Dx horario' : 'clientes_nuevos',
'Medidas de Tx' : 'revisar',
'Nuevo Medidas de Tx' : 'revisar',
'Sheet1' : 'revisar'
}

horarios = {
 'regulados' : 'horario',
 'regulados_reconv' : 'horario',
 'clientes' : 'horario',
 'pmgds' : 'horario',
 'pmgds_fr' : 'horario',
 'pmgds_alimentadores' : 'horario',
 'regulados_nuevos' : 'horario',
 'regulados_reconv_nuevos' : 'horario',
 'clientes_nuevos' : 'horario',
 'clientes_nuevos' : 'horario',
 'pmgds_nuevos' : 'horario',
 'pmgds_fr_nuevos' : 'horario',
 'clientes_nuevos' : 'horario',
 'contratos' : 'nohorario',
 'clientes_detalle' : 'nohorario',
 'pmgds_detalle' : 'nohorario',
 'clientes_detalle_nuevos' : 'nohorario',
 'revisar' :'nohorario'
}


desc = {
    "total [kwh]" : "total",
    "clave transferencia" : "clave_transf",
    "clave coordinado" : "clave_coord",
    "suministrador" : "propietario",
    "zona_balance" : "zona_balance",
    "tipo" : "tipo",
    "alimentador" : "alimentador",
    "descripción" : "descripción",
    "barra" : "barra",
    "hora" : "hora",
    "total" : "total",
    "clave_transf" : "clave_transf",
    "clave_coord" : "clave_coord",
    "propietario" : "propietario"
}

def func_ord(row):
    if row[0] == 'clave_transf':
        val = -12
    elif row[0] == 'origen_dato_archivo':
        val = -11
    elif row[0] == 'origen_dato_hoja':
        val = -10
    elif row[0] == 'zona_balance':
        val = -9
    elif row[0] == 'barra':
        val = -8
    elif row[0] == 'tipo':
        val = -7
    elif row[0] == 'propietario':
        val = -6
    elif row[0] == 'descripción':
        val = -5
    elif row[0] == 'alimentador':
        val = -4
    elif row[0] == 'clave_coord':
        val = -3        
    elif row[0] == 'total':
        val = -2          
    elif row[0] == 'hora':
        val = -1          
    else:
        val = row[0]
    return val



#dir_badx='C:\GitHub\Coordinador_Electrico\BADX'
#dir_mes='C:\GitHub\Coordinador_Electrico\BADX\\2022\Octubre'

dir_badx=r'\\nas-cen1\D.Distribuidoras\Bot\BADX'
dir_mes=r'\\nas-cen1\D.Distribuidoras\Bot\BADX\2022\\Noviembre'

def listar_hojas():
    hojas=[]
    x=os.path.join(str(Path(__file__).parent),"hojas.txt")
    #print(x)          
    archivo_hojas = open(x, "w")
    for path in os.listdir(dir_badx):
        for subpath in os.listdir(dir_badx+"\\"+path):
            for file in glob.glob(dir_badx+"\\"+path+"\\"+subpath + "\*.xlsx"):
                if file.endswith('.xlsx'):
                    excel_file = pd.ExcelFile(file)
                    hojas.extend(excel_file.sheet_names)
                    #print(hojas)
                    #for hoja in excel_file.sheet_names:
                        #archivo_hojas.write(hoja+"\n")
    archivo_hojas.write('\n'.join(set(hojas)))
    #archivo_hojas.write(set(hojas))                    
    archivo_hojas.close()

def revisar_hh():
    reg=[]
    r=''
    regx=''
    x=os.path.join(str(Path(__file__).parent),"reg2.txt")       
    archivo_reg = open(x, "w")
    for path in os.listdir(dir_badx):
        for subpath in os.listdir(dir_badx+"\\"+path):
            for file in glob.glob(dir_badx+"\\"+path+"\\"+subpath + "\*.xlsx"):
                if file.endswith('.xlsx'):
                    #print(file)
                    #reg.append('ARCHIVO: ['+file+']')
                    wb = load_workbook(file)
                    for hoja in wb.get_sheet_names():
                        
                        
                        r=''
                        regx=''
                        if horarios[hojas[hoja]] == 'horario':
                            #print(hoja)
                            #reg.append('HOJA: ['+hoja+']')
                            H=wb.get_sheet_by_name(hoja)
                            A = H['A']
                            for i in range(30):
                                try:
                                    #print(i)
                                    #print(hoja)
                                    #print(file)
                                    #print(A[i].value)
                                    r=str(A[i].value)
                                    r=r.strip()
                                    r=r.lower()
                                    if regx=='hora' and r=='1':
                                        break
                                    elif regx=='1' and r=='2':
                                        reg.pop()
                                        break
                                    elif r != 'None':
                                        regx=r
                                        reg.append(regx)
                                    else:
                                        regx=r
                                except:
                                    pass
                                    
    #archivo_reg.write('\n'.join(reg))
    archivo_reg.write('\n'.join(set(reg)))                 
    archivo_reg.close()










lisreg=['clave_transf','origen_dato_archivo','origen_dato_hoja','zona_balance','barra','tipo','propietario','descripción','alimentador','clave_coord','total','hora']
datos=[]
i=0
j=0
malos=[0]
for file in glob.glob(dir_mes+"\*.xlsx"):
    i=i+1
    if file.endswith('.xlsx'):
        excel_file = pd.ExcelFile(file)
        for hoja in excel_file.sheet_names: #revisa hoja a hoja
            
            if horarios[hojas[hoja]] == 'horario': #accede a las hojas que tienen vectores horarios
                j=j+1
                df1 = pd.read_excel(file, sheet_name=hoja, header=None)
                
                df1=df1.transpose()
                df1=df1.drop_duplicates(ignore_index=True)
                df1=df1.transpose()
                
                
                #print(df1.head()) df['Age'] = df['Age'].astype('string')
                #df1[0] = df1[0].astype('string')
                df1[0] = df1[0].str.lower().fillna(df1[0])
                df1[0]=df1[0].map(desc).fillna(df1[0])
                
                for reg in lisreg:
                    #if df1[0].contains(reg)==0:
                    try:
                        df1[0].value_counts()[reg]
                    except:
                        #df2 = pd.DataFrame(index=df1.index.drop_duplicates(keep='first'))
                        #df2=data.Frame(matrix("", ncol = len(df1.columns), nrow = 1))
                        df2=pd.DataFrame(np.zeros([1,len(df1.columns)]))
                        df2=df2.replace(0, '')
                        df2.loc[0,0]=reg
                        #print(df2)
                
                        df1=df1.append(df2)
                        
                        #df1.set_index(df1[0], drop=False, inplace=True)
                        #df1.set_index(df1[0], drop=True, inplace=True)
                        #df1 = df1.drop(df1.columns[0], axis='columns')
                        
                        #df1 = df1.reindex(lisreg)
                df1['temp'] = df1.apply(func_ord, axis=1)
                df1.set_index(df1['temp'], drop=True, inplace=True)
                df1 = df1.sort_index(ascending=True)
                del(df1['temp'])             
                    #if df1[0].value_counts()[reg] == 0:
                        #print(data['name'].value_counts()['sravan'])
                        #df2 = pd.DataFrame(index=df1.index.drop_duplicates(keep='first'))
                        #df1=df1.append(df2)
                        #print('si')
                #origen_dato_archivo
                os.path.basename(file)
                df1.loc[-11]=str(os.path.basename(file))
                #origen_dato_hoja
                df1.loc[-10]=hoja

                df1.loc[[-11],[0]]='origen_dato_archivo'
                df1.loc[[-10],[0]]='origen_dato_hoja'
         
                #with ExcelWriter(dir_badx+"\\badx_oct22_"+str(i)+"_"+str(j)+".xlsx") as writer:
                #    df1.to_excel(writer, hoja, index=True, header=True)
                if i==1 and j==1:
                    df3=df1.copy()
                else:
                    #print(hoja)
                    df1 = df1.loc[~df1.index.duplicated(keep='first')]
                    temprange=range(df1.shape[1])
                    p=0
                    tempindex=-1
                    for index in temprange:
                        tempindex=tempindex+1
                        #print(index)
                        #print(temprange)
                        #print(tempindex)
                        #print(index)
                        #print('Contenido de la columna: ', df1.iloc[1:1 , index].values)
                        #print(df1._get_value(0, 1, takeable = True))
                        #print(df1.iloc[[0,6,7,8,9],[index]])
                        try:
                            dftemp=df1.iloc[[0,6,7,8,9],[tempindex]]
                            dftemp = dftemp.fillna(0)
                            dftemp = dftemp.replace("", 0)
                        except:
                            #print(index)
                            #print(hoja)
                            #print('malo')
                            #print(index)
                            #print(file)
                            #print(tempindex)
                            df1 = df1.drop(df1.columns[[tempindex]], axis='columns')
                            tempindex=tempindex-1
                            #break
                        #print(hoja)
                        #print('malo')
                        #print(index)
                        
                        #print(file)
                        try:
                            #print(dftemp[1].sum())
                            for (label, content) in dftemp.iteritems():
                               if set(content.values)==set(malos):
                                    #print(index)
                                    df1 = df1.drop(df1.columns[[tempindex]], axis=1)
                                    tempindex=tempindex-1
                                    #print('malo')
                                    #print(index)
                                    #print(hoja)
                                    #print(file)
                        except:
                            pass
                        
                    
                    
                   

                    df3=pd.concat([df3, df1.drop(df1.columns[0], axis='columns')], axis=1)
                    #df1 = df1.drop(df1.columns[0], axis='columns')
                    
                    #df3.join(df1, how='outer')
                    #df3=df3.append(df1)


row = df3.iloc[2].reset_index(drop=True)           # extract first row as series
res=[]
#print(row)
#print(row.size)
res = row.map(hojas)
#print(res.size)
#res=res[res[0]==['pmgds_fr', 'pmgds_fr_nuevos']]
res1 = res.loc[lambda x : (x =='pmgds_fr') | (x =='pmgds_fr_nuevos')]
res2 = res.loc[lambda x : (x =='regulados') | (x =='regulados_reconv') | (x=='regulados_nuevos') | (x=='regulados_reconv_nuevos')]
res3 = res.loc[lambda x : (x =='pmgds') | (x =='pmgds_nuevos')]
res4 = res.loc[lambda x : (x =='pmgds_alimentadores')]


#['pmgds_fr', 'pmgds_fr_nuevos']
#.reset_index().set_index(res.index)
#[lambda x : (x =='pmgds_fr') | (x =='pmgds_fr_nuevos')]
#.isin(['pmgds_fr', 'pmgds_fr_nuevos'])
#filter(items = ['pmgds_fr', 'pmgds_fr_nuevos'])
#data.loc[lambda x : (x < 10) | (x > 20)]
#my_series[my_series.isin([4, 7, 23])]

#.filter(items = ['pmgds_fr', 'pmgds_fr_nuevos'])
#print(res)
#print(res.size)
#df_fr=df3[res]
#print('-------------')
#print(res.index.values.tolist())
df_fr=df3.iloc[:,[0]+res1.index.values.tolist()] 

df5=df3.iloc[:,[0]+res2.index.values.tolist()] 
df6=df3.iloc[:,[0]+res3.index.values.tolist()] 
df7=df3.iloc[:,[0]+res4.index.values.tolist()] 

df4=df3.iloc[:,list(set(row.index.values.tolist())-set(res1.index.values.tolist())-set(res2.index.values.tolist())-set(res3.index.values.tolist())-set(res4.index.values.tolist()))] 

#381


with ExcelWriter(dir_badx+"\\badx_nov22_Medidas_Dx.xlsx") as writer:
                    df_fr.to_excel(writer,index=False, header=False,sheet_name='FR')
                    df4.to_excel(writer,index=False, header=False,sheet_name='05_DX')
                    df5.to_excel(writer,index=False, header=False,sheet_name='06_REG')
                    df6.to_excel(writer,index=False, header=False,sheet_name='07_PMGD')
                    df7.to_excel(writer,index=False, header=False,sheet_name='Alimentadores_PMGD')
                