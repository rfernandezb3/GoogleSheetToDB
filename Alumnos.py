# %% [markdown]
# DATOS DE PRODUCCIÓN DE PLANTA - CL

# %% [markdown]
# Imports Globales

# %%
import pandas as pd
import pyodbc 

# %% [markdown]
# Variables de DB

# %%
server = '192.168.1.46'
database = 'DW_STG'
username = 'DWHUSER'
password = '04DM19'
conn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+password+';')
cursor = conn.cursor()


#truncar table 

##conn.execute("Truncate table DW_STG.Produccion.Sheets_Productos_Produccion_CL ")

# %% [markdown]
# 1.Declaración de Variables

# %%
gsheetid  = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"
sheet_productos = "Class%20%Data"

# %% [markdown]
# 2.Descargar archivo y guardar en variable

# %%
gsheet_url_producto  = "https://docs.google.com/spreadsheets/d/{}/gviz/tq?tqx=out:csv&sheet={}".format(gsheetid, sheet_productos)

# %% [markdown]
# 3.Guardar data en DF

# %%
df_productos = pd.read_csv(gsheet_url_producto,skiprows=0,encoding = 'utf-8')

df_productos.columns

# %% [markdown]
# 4.Campos de la primera máquina | 550-1

# %%
df_productos_cabeceras = df_productos.iloc[:,[0,1,2,3,4,5]]
df_productos_cabeceras.columns

# %% [markdown]
# 5.Renombrar el df para la primera máquina

# %%
df_productos_cabeceras.rename(columns={'Student Name':'NombreEstudiante',
                                       'Gender':'Genero',
                                       'Class Level':'Nivel',
                                       'Home State':'EstadoNatal',
                                       'Major':'Grado',
                                       'Extracurricular Activity':'ActividadExtracurricular'
                                       }, inplace=True)

df_productos_cabeceras.columns

# %% [markdown]
# 6.Insert data a la DB

# %%
for row in df_productos_cabeceras.itertuples():
    cursor.execute('''
                    insert into DW_STG.Produccion.Sheets_Alumnos (NombreEstudiante, Genero, Nivel, EstadoNatal,Grado, ActividadExtracurricular )
                    values (?,?,?,?,?,? )
                    ''',
                    str(row.NombreEstudiante) , 
                    str(row.Genero) , 
                    str(row.Nivel) , 
                    str(row.EstadoNatal) , 
                    str(row.Grado)  , 
                    str(row.ActividadExtracurricular )  
               )

conn.commit()

# %%





