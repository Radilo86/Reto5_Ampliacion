"""
****** --> https://analisisydecision.es/leer-archivos-excel-con-python/
***** ---> https://programmerclick.com/article/9657707575/
**** ----> https://pandas.pydata.org/docs/reference/api/pandas.Index.values.html
*** -----> https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.to_dict.html
** ------>
* ------->
"""
import pandas as pd

# Crear el diccionario vacío final
df_dict = {}
# Leer archivo Excel
df = pd.read_excel('C:/excelPython.xlsx')

print("\n#######################################################################\nMostramos el contenido del excel: ")
print(df)
print("\n#######################################################################")

# Reemplace las celdas vacías en la tabla de Excel, de lo contrario se informará un error en el siguiente paso
df.fillna("", inplace=True)

df_list = []
for i in df.index.values:
    # loc se indexa por nombre de columna iloc se indexa por posición, usando [[número de fila], [nombre de columna]]
    # en este caso indexamos por nombre
    df_line = df.loc[i, ['email', 'telefono', 'ciudad', 'Nombre']].to_dict()
    # Convierta cada línea en un diccionario y agréguela a la lista
    df_list.append(df_line)
df_dict['datos'] = df_list

print("Mostramos el diccionario: ")
print(df_dict)
print("\n***********************************************************************\nMostramos la lista: ")
print(df_list)
print("\n***********************************************************************")

# guardamos la lista en un Dataframe
guardarCSV = pd.DataFrame(df_list)
# convertimos la lista en un fichero CSV
guardarCSV.to_csv('csvFinal.csv', sep=',')

leerCSV = pd.read_csv('csvFinal.csv')
print("Mostramos el contenido del CSV: ")
print(leerCSV)
