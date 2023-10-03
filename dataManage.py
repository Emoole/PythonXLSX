import pandas as pd
import datetime

#Diccionario empresas
company_paths={
    890905627:"C:/Users/ASUS/Downloads/Consolidado1.xlsx",
    800041829:"C:/Users/ASUS/Downloads/YOKOMOTOR - CUFES.xlsx",
    900329399:"C:/Users/ASUS/Downloads/ALEMAUTOS - CUFES.xlsx",
    8909250586:"C:/Users/ASUS/Downloads/DISTRIKIA - CUFES.xlsx",
    900188208:"C:/Users/ASUS/Downloads/MUNDOKIA - CUFES.xlsx",
    900470946:"C:/Users/ASUS/Downloads/AUTOZEN - CUFES.xlsx",
    890939998:"C:/Users/ASUS/Downloads/CAR INTEGRADO - CUFES.xlsx",
    900780840:"C:/Users/ASUS/Downloads/SERFINAD - CUFES.xlsx",
    8909240764:"C:/Users/ASUS/Downloads/PLANAUTOS - CUFES.xlsx",
    901091271:"C:/Users/ASUS/Downloads/PLANSEG - CUFES.xlsx"
}

#Archivo que se descarga de la DIAN
file='C:/Users/ASUS/Downloads/data.xlsx'
print('Data leida')
#Cargamos los datos en un dataframe 'data'
data=pd.read_excel(file)
print(data)
#Extraemos el primer registro de la empresa que recibe el documento
companyId=data.loc[0,'EMPRESA']
print('Id company:', companyId)
#Limpiando el companyId
#companyId=companyId.split('-')
#companyId=companyId[0]
#companyId=companyId.strip()
#Validamos si existe el id en el diccionario 'company_paths'
if companyId in company_paths:
    path=company_paths[companyId]
    print ('Result')
    print(companyId)
    print(path)
    #FECHA Y HORA ACTUAL
    now=datetime.datetime.now()
    now=now.strftime("%Y-%m-%d")

    #Agregamos el lote a cada linea del df data
    data['LOTE']=now

    #Cargar Data Consolidado para la empresa solicitada
    dataConsolidate=pd.read_excel(path)
    dataConsolidate=dataConsolidate._append(data, ignore_index=True)
    #Borra los duplicados del dataframe basado en las columnas EMPRESA Y CANTIDAD
    dataConsolidate=dataConsolidate.drop_duplicates(subset=['EMPRESA','CANTIDAD'])
    #Lleva los datos a un excel. Consolidado seleccionado en el diccionario de paths
    dataConsolidate=dataConsolidate.to_excel(path, index=False)

    print(dataConsolidate)
    #print(companyId)
else:
    print("Path no encontrado")

