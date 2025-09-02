# Importar Librerias
import pandas as pd
import numpy as np
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import json
import os
from openpyxl.utils.cell import get_column_letter, column_index_from_string
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import unicodedata
from datetime import timedelta

#############################################################################################################################
###############                             TARTAMIENTO PARA JORNALES                                 #######################
#############################################################################################################################


#########################   LEER HOJA RESPUESTAS FORMULARIO JORNALES  ##############################

############## VARIABLES POR INTRODUCIR ##########################

spreadsheet_id="10V7_LMf1N9ZFswuoSe0s5c631s55_dMtuZ0T4GsLoQs" #Id del libro de cálculo (Está en URL) EJ: https://docs.google.com/spreadsheets/d/spreadsheetId/edit#gid=0
range_="" #Notación A1 https://developers.google.com/sheets/api/guides/concepts#a1_notation
json_file_path = 'Llave_JSON.json' #Archivo JSON de cuenta de servicio


############## CONFIGURACION DE ACCESO A LAS APIS ##########################

#Cargar JSON 
with open(json_file_path) as f:
    json_file = json.load(f)


# Definir los alcances
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

# Configurar credenciales con google-auth
credentials = Credentials.from_service_account_file(json_file_path, scopes=[
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
])

# Llamar instancia de los servicios de Google
spreadsheet_service = build('sheets', 'v4', credentials=credentials)
drive_service = build('drive', 'v3', credentials=credentials)

###########################################     LEER HOJAS DE DOCUMENTO NECESARIAS      ##################################################


#################################     LEER HOJA DE JORNALES POST TRATAMIENTO    ########################

def read_range(spreadsheet_id, sheet_name="Jornales", range_=None):
  if range_ is not None:
    sheet_name=sheet_name+"!"
  else:
    range_=""

  dict_result = spreadsheet_service.spreadsheets().values().get(
  spreadsheetId=spreadsheet_id, range=sheet_name + range_).execute()
  df = pd.DataFrame(dict_result['values'])
  df.columns=df.iloc[0,:]
  df = df.drop(df.index[0])

  return df

# Leer rango
jornales_historia = read_range(spreadsheet_id=spreadsheet_id)

####################      LEER HOJA DE VALOR MOD         ####################

def read_range(spreadsheet_id, sheet_name="Valor MOF", range_=None):
  if range_ is not None:
    sheet_name=sheet_name+"!"
  else:
    range_=""

  dict_result = spreadsheet_service.spreadsheets().values().get(
  spreadsheetId=spreadsheet_id, range=sheet_name + range_).execute()
  df = pd.DataFrame(dict_result['values'])
  df.columns=df.iloc[0,:]
  df = df.drop(df.index[0])

  return df

# Leer rango
mod = read_range(spreadsheet_id=spreadsheet_id)

mod.columns = [f"{col}_{i//3+1}" for i,col in enumerate(mod.columns)]

# Creamos una lista de DataFrames, uno por cada bloque de 3 columnas:
bloques = []
for i in (1, 2):
    tmp = mod[[f"Pagado a:_{i}", f"Mes_{i}", f"Valor_{i}"]].copy()
    tmp.columns = ["Pagado a:", "Mes", "Valor"]
    bloques.append(tmp)

# Concatenamos verticalmente y reindexamos:
mod = pd.concat(bloques, ignore_index=True)


# Filtrar filas con Empleado vacío, puedes filtrarlas:
mod = mod.dropna(subset=["Pagado a:"])


# Extraer solo mes y año de columna de mes
mod['Mes'] = mod['Mes'].str[-7:]


# Dar formato numerico a columna de valor
mod['Valor'] = mod['Valor'].str.replace('.','')
mod['Valor'] = mod['Valor'].str.replace(',','.')
mod['Valor'] = pd.to_numeric(mod['Valor'])

# Crear columna 'Concepto P&L o Balance General' y Tipo Jornal
mod['Concepto P&L o Balance General'] = 'MDO Fija'
mod['Tipo Jornal'] = 'Directo'


####################      LEER HOJA DE MES PROYECTO         ####################

def read_range(spreadsheet_id, sheet_name="Mes Proyecto", range_=None):
  if range_ is not None:
    sheet_name=sheet_name+"!"
  else:
    range_=""

  dict_result = spreadsheet_service.spreadsheets().values().get(
  spreadsheetId=spreadsheet_id, range=sheet_name + range_).execute()
  df = pd.DataFrame(dict_result['values'])
  df.columns=df.iloc[0,:]
  df = df.drop(df.index[0])

  return df

# Leer rango
mes_proyecto = read_range(spreadsheet_id=spreadsheet_id)

# Extraer solo mes y año de columna de mes
mes_proyecto['Mes'] = mes_proyecto['Mes'].str[-7:]


####################      LEER HOJA DE RESPUEATS DEL FORMULARIO JORNALES         ####################

def read_range(spreadsheet_id, sheet_name="Respuestas de formulario Jornales", range_=None):
  if range_ is not None:
    sheet_name=sheet_name+"!"
  else:
    range_=""

  dict_result = spreadsheet_service.spreadsheets().values().get(
  spreadsheetId=spreadsheet_id, range=sheet_name + range_).execute()
  df = pd.DataFrame(dict_result['values'])
  df.columns=df.iloc[0,:]
  df = df.drop(df.index[0])

  return df

# Leer rango
respuestas_jornales = read_range(spreadsheet_id=spreadsheet_id)

####################      LEER HOJA DE RESPUEATS DEL FORMULARIO COMPRAS         ####################

def read_range(spreadsheet_id, sheet_name="Respuestas de formulario Compras", range_=None):
  if range_ is not None:
    sheet_name=sheet_name+"!"
  else:
    range_=""

  dict_result = spreadsheet_service.spreadsheets().values().get(
  spreadsheetId=spreadsheet_id, range=sheet_name + range_).execute()
  df = pd.DataFrame(dict_result['values'])
  df.columns=df.iloc[0,:]
  df = df.drop(df.index[0])

  return df

# Leer rango
respuestas_compras = read_range(spreadsheet_id=spreadsheet_id)

####################      LEER HOJA DE HISTORIAL DE INSUMOS         ####################

def read_range(spreadsheet_id, sheet_name="Insumos", range_=None):
  if range_ is not None:
    sheet_name=sheet_name+"!"
  else:
    range_=""

  dict_result = spreadsheet_service.spreadsheets().values().get(
  spreadsheetId=spreadsheet_id, range=sheet_name + range_).execute()
  df = pd.DataFrame(dict_result['values'])
  df.columns=df.iloc[0,:]
  df = df.drop(df.index[0])

  return df

# Leer rango
historial_insumos = read_range(spreadsheet_id=spreadsheet_id)


####################      LEER HOJA DE CLASIDICACION DE INSUMOS         ####################

def read_range(spreadsheet_id, sheet_name="Clasificacion Insumos", range_=None):
  if range_ is not None:
    sheet_name=sheet_name+"!"
  else:
    range_=""

  dict_result = spreadsheet_service.spreadsheets().values().get(
  spreadsheetId=spreadsheet_id, range=sheet_name + range_).execute()
  df = pd.DataFrame(dict_result['values'])
  df.columns=df.iloc[0,:]
  df = df.drop(df.index[0])

  return df

# Leer rango
clasificacion_insumos = read_range(spreadsheet_id=spreadsheet_id)


#################################     LEER HOJA DE INVENTARIO    ########################

def read_range(spreadsheet_id, sheet_name="Inventario Inicial", range_=None):
  if range_ is not None:
    sheet_name=sheet_name+"!"
  else:
    range_=""

  dict_result = spreadsheet_service.spreadsheets().values().get(
  spreadsheetId=spreadsheet_id, range=sheet_name + range_).execute()
  df = pd.DataFrame(dict_result['values'])
  df.columns=df.iloc[0,:]
  df = df.drop(df.index[0])

  return df

# Leer rango
inventario = read_range(spreadsheet_id=spreadsheet_id)

####################      LEER HOJA DE RESPUEATS DEL FORMULARIO VENTAS         ####################

def read_range(spreadsheet_id, sheet_name="Respuestas de formulario Ventas", range_=None):
  if range_ is not None:
    sheet_name=sheet_name+"!"
  else:
    range_=""

  dict_result = spreadsheet_service.spreadsheets().values().get(
  spreadsheetId=spreadsheet_id, range=sheet_name + range_).execute()
  df = pd.DataFrame(dict_result['values'])
  df.columns=df.iloc[0,:]
  df = df.drop(df.index[0])

  return df

# Leer rango
respuestas_ventas = read_range(spreadsheet_id=spreadsheet_id)


#################################     LEER HOJA DE VENTAS POST TRATAMIENTO    ########################

def read_range(spreadsheet_id, sheet_name="Ventas", range_=None):
  if range_ is not None:
    sheet_name=sheet_name+"!"
  else:
    range_=""

  dict_result = spreadsheet_service.spreadsheets().values().get(
  spreadsheetId=spreadsheet_id, range=sheet_name + range_).execute()
  df = pd.DataFrame(dict_result['values'])
  df.columns=df.iloc[0,:]
  df = df.drop(df.index[0])

  return df

# Leer rango
ventas_historia = read_range(spreadsheet_id=spreadsheet_id)


#############################################################################################################                                          
#############################          TRATAMIENTO PARA JORNALES             ################################
#############################################################################################################

# Filtrar respuestas por formulario
respuestas_jornales = respuestas_jornales[respuestas_jornales['Marca temporal'] != '']


# Cambiar nombre a columna de unidad de minutos
respuestas_jornales.rename(columns={'Unidad (minutos trabajados)\n\n1 hora = 60 minutos\n1 hora y media = 90 minutos\n2 horas = 120 minutos\n2 horas y media = 150 minutos\n3 horas = 180 minutos\n3 horas y media = 210 minutos\n4 horas = 240 minutos\n4 horas y media = 270 minutos\n5 horas = 300 minutos\n5 horas y media = 330 minutos\n6 horas = 360 minutos\n6 horas y media = 390 minutos\n7 horas = 420 minutos\n7 horas y media = 450 minutos\n8 hora = 480 minutos':
                                    'Unidad (minutos trabajados)', 'Ciclo (solo número)\n':'Ciclo (solo número)', 'Valor Unidad Jornal':'Valor Unidad',
                                    'Item Archivo Financiero Jornales':'Item Archivo Financiero'}, inplace=True)


# Funcion para expandir filas de jornales
def expandir_por_lineas(df, cols_expandir, converters=None):
  
  if converters is None:
      converters = {}

  filas = []
  for _, row in df.iterrows():
      # Para cada columna a expandir, obtener lista procesada
      listas = {}
      for col in cols_expandir:
          raw = str(row.get(col, ""))  # si hay NaN u otro, convertir a str para split
          partes = raw.splitlines()  # mantiene lógica de saltos de línea
          conv = converters.get(col, lambda x: x.strip())
          procesadas = []
          for p in partes:
              p = p.strip()
              try:
                  procesadas.append(conv(p))
              except Exception:
                  procesadas.append(None)
          listas[col] = procesadas

      # Determinar cuántas subfilas generar (mínimo de longitudes)
      if listas:
          longitud = min(len(lst) for lst in listas.values())
      else:
          longitud = 1

      for i in range(longitud):
          nueva = row.copy()
          for col in cols_expandir:
              nueva[col] = listas[col][i]
          filas.append(nueva)

  return pd.DataFrame(filas).reset_index(drop=True)

# Convertidores para las columnas que vas a expandir
converters = {
    'Pagado a:': lambda s: s.strip() if isinstance(s, str) else s,
    'Unidad (minutos trabajados)': lambda s: float(s.replace('.', '').replace(',', '.')) 
        if isinstance(s, str) and s.strip() not in ('', '-', None) else None,
    'Valor Unidad': lambda s: float(s.replace('.', '').replace(',', '.')) 
        if isinstance(s, str) and s.strip() not in ('', '-', None) else None,
}


# Expansión sobre respuestas_jornales
respuestas_jornales = expandir_por_lineas(
    respuestas_jornales,
    ['Pagado a:', 'Unidad (minutos trabajados)', 'Valor Unidad'],
    converters=converters
)

# Reemplazos de nombre felipe perez y paola
respuestas_jornales['Pagado a:'] = respuestas_jornales['Pagado a:'].replace(['Felipe Pérez', 'FELIPE PEREZ', 'felipe pérez', 'felipe perez',
                                                                             'FELIPE PÉREZ', 'Felipe Perez', 'paola', 'Felipe', 'Felipe perez'],
                                                                            
                                                                            ['Felipe Perez', 'Felipe Perez', 'Felipe Perez', 'Felipe Perez',
                                                                             'Felipe Perez','Felipe Perez', 'Paola', 'Felipe Perez', 'Felipe Perez'])

# Crear columna Ciclo
respuestas_jornales['Ciclo'] = 'Ciclo ' + respuestas_jornales['Ciclo (solo número)']

# Crear columna de mes, temporalmente
respuestas_jornales['Mes'] = respuestas_jornales['Fecha Actividad'].str[-7:]

# Obtener el mes del proyecto segun la fecha de actividad
respuestas_jornales = pd.merge(respuestas_jornales, mes_proyecto, how='left', on=['Mes'])

# Obtener Valor de MOD
respuestas_jornales = pd.merge(respuestas_jornales, mod, how='left', on=['Pagado a:', 'Mes'])

# # Modificar valor del jornal a la mano de obra fija
# respuestas_jornales.loc[respuestas_jornales['Pagado a:'] == 'Felipe Perez', 'Valor Unidad'] = respuestas_jornales['Valor']

# Asignar valor por defecto a columna Concepto P&L o Balance General
respuestas_jornales['Concepto P&L o Balance General'][respuestas_jornales['Valor'].isnull()] = 'Jornales'

# Asignar valor por defecto a columna tipo jornal
respuestas_jornales['Tipo Jornal'][respuestas_jornales['Valor'].isnull()] = 'Indirecto'

# Modificar valor del jornal a la mano de obra fija
respuestas_jornales['Valor Unidad'][~respuestas_jornales['Valor'].isnull()] = respuestas_jornales['Valor']

# Eliminar columnas de mes y valor
respuestas_jornales.drop(columns={'Mes', 'Valor'}, inplace=True)

# Crear columna Mes Proyecto
respuestas_jornales['Mes Proyecto'] = 'Mes ' + respuestas_jornales['Mes Proyecto']

respuestas_jornales['Mes del Lote'] = ''

# Dar formato a columna de pagado a:
respuestas_jornales['Pagado a:'] = respuestas_jornales['Pagado a:'].str.title()


# Creaer columna de Unidades
respuestas_jornales['Unidad'] = (respuestas_jornales['Unidad (minutos trabajados)'] / 60) / 8

# Crear columna Total
respuestas_jornales['Total'] = respuestas_jornales['Unidad'] * respuestas_jornales['Valor Unidad']

# Lista de valores que deben excluirse
excluir = [
    'Riego (Fumigación Fitosanitaria)',
    'Nutrientes y mantenimiento (Fertirrigación)',
    'Abono Edáfico (Fertilización)'
]

# Aplicar condición
respuestas_jornales['Item'] = respuestas_jornales.apply(
    lambda row: row['Clasificación/Tipo Actividad'] if row['Item'] == '' and row['Clasificación/Tipo Actividad'] not in excluir else row['Item'],
    axis=1
)


# Dejar columnas necesarias
jornales = respuestas_jornales[['Marca temporal', 'Fecha Actividad', 'Mes Proyecto', 'Mes del Lote', 'Pagado a:', 'Lote',
                                'Concepto P&L o Balance General', 'Clasificación/Tipo Actividad',
                                'Item Archivo Financiero', 'Tipo Jornal', 'Item', 'Unidad', 'Valor Unidad', 'Total',
                                'Ciclo', 'Invernadero', 'Observaciones', 'Cantidad Usada por Item']]


# Descartar respuestas de formulario
jornales_historia = jornales_historia[jornales_historia['Marca temporal'] == '']

# Crear columna dummy de cantidad usada por item igual a 0
jornales_historia['Cantidad Usada por Item'] = '0'

# Crear listado con columnas numericas para dar formato numerico
numericas_2 = ['Valor Unidad', 'Total', 'Unidad']

# Dar formato numerico
for numero2 in numericas_2:
  jornales_historia[numero2] = jornales_historia[numero2].str.replace('.','')
  jornales_historia[numero2] = jornales_historia[numero2].str.replace(',','.')
  jornales_historia[numero2] = pd.to_numeric(jornales_historia[numero2])
  

############################################################################

# ---------- 1) HISTÓRICO ----------
# Fechas y orden
jornales_historia['Fecha Actividad'] = pd.to_datetime(jornales_historia['Fecha Actividad'], format='%d/%m/%Y')
jornales_historia.sort_values(by=['Lote', 'Invernadero', 'Fecha Actividad'], inplace=True)


# ---------- Utilidades ----------
def meses_completos(start, end):
    # Cuenta meses completos entre start y end
    year_diff = end.year - start.year
    month_diff = end.month - start.month
    months = year_diff * 12 + month_diff
    if end.day < start.day:
        months -= 1
    return max(0, months)

def quitar_acentos(texto):
    if not isinstance(texto, str):
        return texto
    nfkd = unicodedata.normalize("NFKD", texto)
    return "".join([c for c in nfkd if not unicodedata.combining(c)])

def extraer_numero_mes(x):
    if pd.isna(x):
        return 0
    if isinstance(x, str):
        import re
        m = re.search(r'(\d+)', x)
        return int(m.group(1)) if m else 0
    try:
        return int(x)
    except:
        return 0

def sumar_meses_preservando_dia(fecha, n_meses):
    # Avanza 'fecha' n_meses conservando el día en lo posible
    year = fecha.year + (fecha.month - 1 + n_meses) // 12
    month = (fecha.month - 1 + n_meses) % 12 + 1
    # Ajusta el día al máximo posible del nuevo mes
    day = min(fecha.day, pd.Timestamp(year, month, 1).days_in_month)
    return pd.Timestamp(year=year, month=month, day=day)


# Normalizaciones auxiliares
jornales_historia['actividad_norm'] = (
    jornales_historia['Clasificación/Tipo Actividad']
    .astype(str).str.strip().str.lower().apply(quitar_acentos)
)
jornales_historia['Mes_Num'] = jornales_historia['Mes del Lote'].apply(extraer_numero_mes)

TARGET_ERRAD = quitar_acentos("Erradicación Plantas").lower()   # "erradicacion plantas"
TARGET_SIEMBRA = quitar_acentos("Siembra plantas").lower()      # "siembra plantas"

def simular_estado_historial(grp):
    """
    Recorre TODO el historial del (Lote, Invernadero) como una máquina de estados
    para obtener el estado final según la nueva lógica:
      - tras erradicación -> esperar primera 'Siembra plantas' (Mes 0 indefinido)
      - primera 'Siembra plantas' fija ancla (Mes 0), desde ahí contar meses completos
      - si se pasaría de Mes 8, reinicia a Mes 0 en esa actividad
    Devuelve:
      - Ultimo Mes
      - Fecha Ultimo Cambio (fecha ancla o la última fecha en que cambió el mes/ancla)
      - Esperando Siembra (bool): True si seguimos en modo "Mes 0 indefinido" (post-erradicación sin siembra)
      - Ancla Siembra (fecha de la primera siembra activa; NaT si ninguna)
    """
    grp = grp.sort_values('Fecha Actividad').copy()

    esperando_siembra = False     # True = modo "Mes 0 indefinido"
    ancla_siembra = pd.NaT        # fecha de la primera siembra válida (después de erradicación)
    ultimo_mes = 0
    fecha_ultimo_cambio = pd.NaT  # fecha ancla o última vez que cambió el mes

    for _, fila in grp.iterrows():
        act = fila['actividad_norm']
        f = fila['Fecha Actividad']

        if act == TARGET_ERRAD:
            # Erradicación: pasar a modo espera de siembra; Mes 0 indefinido
            esperando_siembra = True
            ancla_siembra = pd.NaT
            ultimo_mes = 0
            fecha_ultimo_cambio = f
            continue

        if esperando_siembra:
            # Todo permanece en Mes 0 hasta la PRIMERA "Siembra plantas"
            if act == TARGET_SIEMBRA:
                # Primera siembra posterior a erradicación: fija ancla y arranca ciclo
                ancla_siembra = f
                ultimo_mes = 0
                fecha_ultimo_cambio = f
                esperando_siembra = False
            else:
                # Sigue en Mes 0 indefinido
                ultimo_mes = 0
            continue

        # No estamos esperando siembra: puede o no existir ancla
        if pd.isna(ancla_siembra):
            # Aún no hay ancla (no ha existido una primera siembra "activa" históricamente)
            if act == TARGET_SIEMBRA:
                ancla_siembra = f
                ultimo_mes = 0
                fecha_ultimo_cambio = f
            else:
                # Sin ancla, se mantiene Mes 0
                ultimo_mes = 0
            continue

        # Modo normal con ancla activa: calcular meses desde último cambio
        if pd.notna(fecha_ultimo_cambio):
            m = meses_completos(fecha_ultimo_cambio, f)
            if m >= 1:
                candidato = ultimo_mes + m
                if candidato > 8:
                    # Reinicio a Mes 0 en esta actividad
                    ultimo_mes = 0
                    ancla_siembra = f
                    fecha_ultimo_cambio = f
                else:
                    ultimo_mes = candidato
                    fecha_ultimo_cambio = sumar_meses_preservando_dia(fecha_ultimo_cambio, m)

        # Si llega otra "Siembra plantas" consecutiva, se IGNORA (no cambia el ancla)
        # para mantener la primera siembra como inicio real del ciclo.

    return pd.Series({
        'Ultimo Mes': int(ultimo_mes),
        'Fecha Ultimo Cambio': fecha_ultimo_cambio,
        'Esperando Siembra': bool(esperando_siembra),
        'Ancla Siembra': ancla_siembra
    })

# Resumen por Lote + Invernadero
resultados = jornales_historia.groupby(['Lote', 'Invernadero'], group_keys=False).apply(simular_estado_historial)
resultados.reset_index(inplace=True)
resultados_dict = resultados.set_index(['Lote', 'Invernadero']).to_dict('index')

# Limpieza columnas auxiliares del histórico
jornales_historia.drop(columns={'actividad_norm', 'Mes_Num'}, inplace=True, errors='ignore')

# ---------- 2) NUEVOS REGISTROS ----------
# Fechas y orden
jornales['Fecha Actividad'] = pd.to_datetime(jornales['Fecha Actividad'], format='%d/%m/%Y')
jornales['Marca temporal'] = pd.to_datetime(jornales['Marca temporal'])
jornales.sort_values(by=['Lote', 'Invernadero', 'Fecha Actividad', 'Marca temporal'], inplace=True)

def asignar_mes_del_lote_v2(grupo):
    grupo = grupo.copy()
    key = (grupo['Lote'].iloc[0], grupo['Invernadero'].iloc[0])
    hist = resultados_dict.get(key, None)

    # Estado inicial a partir del histórico simulado
    if hist is None:
        esperando_siembra = False   # nunca hubo erradicación, tampoco siembra
        ancla_siembra = pd.NaT
        ultimo_mes = 0
        fecha_ultimo_cambio = pd.NaT
    else:
        esperando_siembra = bool(hist['Esperando Siembra'])
        ancla_siembra = hist['Ancla Siembra']
        ultimo_mes = int(hist['Ultimo Mes'])
        fecha_ultimo_cambio = hist['Fecha Ultimo Cambio']

    # Normalización para comparar actividades
    grupo['actividad_norm'] = (
        grupo['Clasificación/Tipo Actividad'].astype(str).str.strip().str.lower().apply(quitar_acentos)
    )

    mes_del_lote = []

    for _, fila in grupo.iterrows():
        act = fila['actividad_norm']
        f = fila['Fecha Actividad']

        if act == TARGET_ERRAD:
            # Erradicación: pasar a modo espera siembra; Mes 0 indefinido
            esperando_siembra = True
            ancla_siembra = pd.NaT
            ultimo_mes = 0
            fecha_ultimo_cambio = f
            mes_del_lote.append(f'Mes {ultimo_mes}')
            continue

        if esperando_siembra:
            # Todo queda en Mes 0 hasta la PRIMERA "Siembra plantas"
            if act == TARGET_SIEMBRA:
                ancla_siembra = f
                ultimo_mes = 0
                fecha_ultimo_cambio = f
                esperando_siembra = False
            # (si no es siembra, permanece Mes 0)
            mes_del_lote.append(f'Mes {ultimo_mes}')
            continue

        # No esperando siembra:
        if pd.isna(ancla_siembra):
            # No hay ancla aún: solo la primera "Siembra plantas" fija el inicio real
            if act == TARGET_SIEMBRA:
                ancla_siembra = f
                ultimo_mes = 0
                fecha_ultimo_cambio = f
            mes_del_lote.append(f'Mes {ultimo_mes}')
            continue

        # Modo normal con ancla activa: avanzar por meses completos
        if pd.notna(fecha_ultimo_cambio):
            m = meses_completos(fecha_ultimo_cambio, f)
            if m >= 1:
                candidato = ultimo_mes + m
                if candidato > 8:
                    # Reinicio a Mes 0 en esta actividad
                    ultimo_mes = 0
                    ancla_siembra = f
                    fecha_ultimo_cambio = f
                else:
                    ultimo_mes = candidato
                    fecha_ultimo_cambio = sumar_meses_preservando_dia(fecha_ultimo_cambio, m)

        # Si llega otra "Siembra plantas" luego de tener ancla, se IGNORA (se mantiene la primera)
        mes_del_lote.append(f'Mes {ultimo_mes}')

    grupo['Mes del Lote'] = mes_del_lote
    # limpiar auxiliar
    grupo.drop(columns=['actividad_norm'], inplace=True, errors='ignore')
    return grupo

# Aplicar
jornales_actualizado = jornales.groupby(['Lote', 'Invernadero'], group_keys=False).apply(asignar_mes_del_lote_v2)
jornales_actualizado.reset_index(drop=True, inplace=True)


# Ordenar df jornales actualizado
jornales_actualizado =  jornales_actualizado.sort_values(by=['Fecha Actividad', 'Marca temporal']).reset_index(drop=True)

# Devolver formato a columnas en actualizado
jornales_actualizado['Fecha Actividad'] = jornales_actualizado['Fecha Actividad'].dt.strftime('%d/%m/%Y')
jornales_actualizado['Marca temporal'] = jornales_actualizado['Marca temporal'].dt.strftime('%d/%m/%Y %H:%M:%S')


# Ordenar df jornales hisotria
jornales_historia =  jornales_historia.sort_values(by='Fecha Actividad').reset_index(drop=True)

# Devolver formato a columna de fecha a historia
jornales_historia['Fecha Actividad'] = jornales_historia['Fecha Actividad'].dt.strftime('%d/%m/%Y')

# Concatenar los DataFrames
jornales_completo = pd.concat([jornales_historia, jornales_actualizado]).reset_index(drop=True)

# Dejar columnas necesarias
jornales_consolidado = jornales_completo[['Marca temporal', 'Fecha Actividad', 'Mes Proyecto', 'Mes del Lote', 'Pagado a:', 'Lote',
                                          'Concepto P&L o Balance General', 'Clasificación/Tipo Actividad',
                                          'Item Archivo Financiero', 'Tipo Jornal', 'Item', 'Unidad',
                                          'Valor Unidad', 'Total', 'Ciclo', 'Invernadero', 'Observaciones']]


############################################        ACTUALIZAR SHEET JORNALES     ##############################


############# DEFINICION DE FUNCIONES ##########################

#Limpiar contenido en hoja 
def clear_range(spreadsheet_id, sheet_name="Jornales", range_=None):

    if range_ is not None:
        sheet_name=sheet_name+"!"
    else:
        range_=""

    dict_result = spreadsheet_service.spreadsheets().values().clear(
    spreadsheetId=spreadsheet_id, range=sheet_name + range_).execute()

#Escribir en hoja a partir de dataframe. Empieza a pegar la información a partir de la primera fila vacía que encuentre.
def write_range(spreadsheet_id, dataframe, sheet_name="Jornales", range_=None, include_headers=False):
    
        #rellenar los NaN con vacíos
        dataframe = dataframe.fillna("")
        
        #eliminar saltos de líneas y retornos de carro
        dataframe = dataframe.replace(r"\n","", regex=True).replace(r"\r","", regex=True)
        
        
        if range_ is not None:
          sheet_name=sheet_name+"!"
        else:
          range_=""

        if include_headers==True:
          content_values = dataframe.values.tolist()
          content_values.insert(0, dataframe.columns.tolist())
        else:
          content_values=dataframe.values.tolist()

        body = {
        'values': content_values
        }
        spreadsheet_service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id, body=body, valueInputOption='USER_ENTERED', range=sheet_name + range_).execute()



# #################   CARGAR SHEET Jornales    ##############

# Limpiar Hoja Google sheets
clear_range(spreadsheet_id=spreadsheet_id, sheet_name='Jornales')

# Escribir df en hoja Google Sheets
write_range(spreadsheet_id=spreadsheet_id, sheet_name='Jornales', dataframe=jornales_consolidado, include_headers=True)

# ########################################################################################################################

# print(jornales_consolidado)

#############################################################################################################                                          
#############################          TRATAMIENTO PARA INSUMOS              ################################
#############################################################################################################

##############################################      INSUMOS COMPRADOS     ##############################################

# Crear listado con columnas numericas para dar formato numerico
numericas2 = ['Cantidad Comprada (No usar separador de mil, para decimales usar punto)', 'Valor Unidad']

# Dar formato numerico
for numero in numericas2:
  # respuestas_compras[numero] = respuestas_compras[numero].str.replace('.','')
  # respuestas_compras[numero] = respuestas_compras[numero].str.replace(',','.')
  respuestas_compras[numero] = pd.to_numeric(respuestas_compras[numero])
  
# Crear columna Ciclo
respuestas_compras['Ciclo'] = 'Ciclo ' + respuestas_compras['Ciclo (solo numero)']

# Renombrar columna
respuestas_compras.rename(columns={'Cantidad Comprada (No usar separador de mil, para decimales usar punto)':'Cantidad Comprada/Aplicada',
                                  'Fecha Compra':'Fecha Compra/Aplicacion'},inplace=True)

# Eliminar columna
respuestas_compras.drop(columns={'Ciclo (solo numero)'}, inplace=True)

# Eliminar posibles espacios en blanco en columna de item
clasificacion_insumos['Item'] = clasificacion_insumos['Item'].str.strip()

# Asignar una claisificacion y unidad de medida con base a el item
insumos_compras = pd.merge(respuestas_compras, clasificacion_insumos, on='Item', how='left')

# Llenar Nulos con revisar
insumos_compras[['Clasificación/Tipo Actividad', 'Unidad Medida']] = insumos_compras[['Clasificación/Tipo Actividad', 'Unidad Medida']].fillna('Revisar')

# Crear columna Concepto igual a Aplicación
insumos_compras['Concepto'] = 'COMPRA'

# Crear columna Lote igual a TODOS
insumos_compras['Lote'] = 'TODOS'

# Crear la nueva columna 'Concepto P&L o Balance General'
insumos_compras['Concepto P&L o Balance General'] = insumos_compras['Clasificación/Tipo Actividad'].apply(
    lambda x: 'Plántulas' if 'Plantas' in x else 'Insumos'
)

# Crear columna Total
insumos_compras['Total'] = insumos_compras['Cantidad Comprada/Aplicada'] * insumos_compras['Valor Unidad']

# Crear columna temporal llamada Mes con base a la fecha de aplicacion en df insumos compras
insumos_compras['Mes'] = insumos_compras['Fecha Compra/Aplicacion'].str[-7:]

# Obtener mes del proyecto desde df mes del proyecto
insumos_compras = pd.merge(insumos_compras, mes_proyecto, on='Mes', how='left')

# Añadir la palabra 'MES' en Mes Proyecto
insumos_compras['Mes Proyecto'] = 'MES ' + insumos_compras['Mes Proyecto']

# Crear columna temporal llamada Mes con base a la fecha de actividad en df jornales_consolidado
jornales_consolidado['Mes'] = jornales_consolidado['Fecha Actividad'].str[-7:]

# Agrupar df de jornales consilidado
jornales_ultimos = jornales_consolidado.groupby('Mes', as_index=False).last()

# Obtener mes del lote desde df jornlaes consolidado
insumos_compras = pd.merge(insumos_compras,
                          jornales_ultimos[['Mes', 'Mes del Lote']],
                          on='Mes',
                          how='left'
                          ).drop(columns='Mes')

# Llenar nulos en Mes del lote con 'Pendiente'
insumos_compras['Mes del Lote'] = insumos_compras['Mes del Lote'].fillna('Pendiente')

# Reordenar df de insumos compras
insumos_compras = insumos_compras[['Marca temporal', 'Fecha Compra/Aplicacion', 'Concepto', 'Mes Proyecto',
                                  'Mes del Lote', 'Pagado a:', 'Lote', 'Concepto P&L o Balance General',
                                  'Clasificación/Tipo Actividad', 'Item Archivo Financiero',
                                  'Cantidad Comprada/Aplicada', 'Item', 'Unidad Medida', 'Valor Unidad',
                                  'Total', 'Ciclo', 'Invernadero', 'Observaciones']]

# Crear listado con columnas numericas para dar formato numerico
numericas3 = ['Cantidad Comprada/Aplicada', 'Valor Unidad', 'Total']

# Dar formato numerico
for numero in numericas3:
  historial_insumos[numero] = historial_insumos[numero].str.replace('.','')
  historial_insumos[numero] = historial_insumos[numero].str.replace(',','.')
  historial_insumos[numero] = pd.to_numeric(historial_insumos[numero])

# Descartar resouestas de formulario del historial de insumos
historial_insumos = historial_insumos[historial_insumos['Marca temporal'] == '']

# Unir historial con nuevas compras
historial_insumos = pd.concat([historial_insumos,
                              insumos_compras]
                              )

################################        INVENTARIO      ##################

# Dejar solo columnas necesarias en insumos compras
insumos_compras = insumos_compras[['Concepto P&L o Balance General', 'Clasificación/Tipo Actividad',
                                    'Item', 'Unidad Medida', 'Valor Unidad', 'Cantidad Comprada/Aplicada',
                                    'Total', 'Fecha Compra/Aplicacion']]

# Crear listado con columnas numericas para dar formato numerico
numericas4 = ['Cantidad Comprada/Aplicada', 'Valor Unidad', 'Total']

# Dar formato numerico
for numero in numericas4:
  inventario[numero] = inventario[numero].str.replace('.','')
  inventario[numero] = inventario[numero].str.replace(',','.')
  inventario[numero] = pd.to_numeric(inventario[numero])
  
# Asiganr nuevas compras al inventario
inventario = pd.concat([inventario, insumos_compras])

# Convertir las fechas a formato datetime para ordenarlas correctamente
inventario["Fecha Compra/Aplicacion"] = pd.to_datetime(inventario["Fecha Compra/Aplicacion"], dayfirst=True)

# Ordenar el inventario por fecha de compra ascendente para aplicar FIFO
inventario = inventario.sort_values(by=["Fecha Compra/Aplicacion"]).reset_index(drop=True)

# Dar formato string a columna de fecha
inventario['Fecha Compra/Aplicacion'] = inventario['Fecha Compra/Aplicacion'].dt.strftime('%d/%m/%Y')

# # Dar formato numerico a columna de cnatidad aplicada
# inventario['Cantidad Comprada/Aplicada'] = inventario['Cantidad Comprada/Aplicada'].str.replace('.','')
# inventario['Cantidad Comprada/Aplicada'] = inventario['Cantidad Comprada/Aplicada'].str.replace(',','.')
# inventario['Cantidad Comprada/Aplicada'] = inventario['Cantidad Comprada/Aplicada'].astype(float)

############################################        ACTUALIZAR SHEET INVENTARIO     ##############################

############# DEFINICION DE FUNCIONES ##########################

#Limpiar contenido en hoja 
def clear_range(spreadsheet_id, sheet_name="Inventario", range_=None):

    if range_ is not None:
        sheet_name=sheet_name+"!"
    else:
        range_=""

    dict_result = spreadsheet_service.spreadsheets().values().clear(
    spreadsheetId=spreadsheet_id, range=sheet_name + range_).execute()

#Escribir en hoja a partir de dataframe. Empieza a pegar la información a partir de la primera fila vacía que encuentre.
def write_range(spreadsheet_id, dataframe, sheet_name="Inventario", range_=None, include_headers=False):
    
        #rellenar los NaN con vacíos
        dataframe = dataframe.fillna("")
        
        #eliminar saltos de líneas y retornos de carro
        dataframe = dataframe.replace(r"\n","", regex=True).replace(r"\r","", regex=True)
        
        
        if range_ is not None:
          sheet_name=sheet_name+"!"
        else:
          range_=""

        if include_headers==True:
          content_values = dataframe.values.tolist()
          content_values.insert(0, dataframe.columns.tolist())
        else:
          content_values=dataframe.values.tolist()

        body = {
        'values': content_values
        }
        spreadsheet_service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id, body=body, valueInputOption='USER_ENTERED', range=sheet_name + range_).execute()



# #################   CARGAR SHEET Inventario    ##############

# Limpiar Hoja Google sheets
clear_range(spreadsheet_id=spreadsheet_id, sheet_name='Inventario')

# Escribir df en hoja Google Sheets
write_range(spreadsheet_id=spreadsheet_id, sheet_name='Inventario', dataframe=inventario, include_headers=True)

# # print(inventario)

##############################################      INSUMOS APLICADOS     ##############################################

# Filtrar respuestas por formulario
respuestas_insumos = jornales_completo[jornales_completo['Marca temporal'] != '']

# Crear insumos_nuevos de insumos_aplicacion
insumos_nuevos = respuestas_insumos[respuestas_insumos['Clasificación/Tipo Actividad'].isin(excluir)].reset_index(drop=True)

# Descartar lo que item este vacio
insumos_nuevos = insumos_nuevos[insumos_nuevos['Item'] != '']

# Descartar posibles duplicados en los insumos
insumos_nuevos = insumos_nuevos.drop_duplicates(subset=['Marca temporal', 'Fecha Actividad', 'Mes Proyecto',
                                                        'Mes del Lote', 'Clasificación/Tipo Actividad', 'Item',
                                                        'Ciclo', 'Invernadero', 'Observaciones', 'Cantidad Usada por Item'])

# Validar si no hay nuevas respuestas
if insumos_nuevos.empty:
  
  print('Sin nuevas respuestas')  
  
else:

  
  # Función para dividir las columnas y crear nuevas filas
  def expandir_items(insumos_nuevos, item_col, cantidad_col):
    df = insumos_nuevos.copy()

    # Asegurar que sean strings y evitar None/NaN
    df[item_col] = df[item_col].fillna('').astype(str)
    df[cantidad_col] = df[cantidad_col].fillna('').astype(str)

    # Dividir manejando distintos tipos de salto de línea
    df[item_col] = df[item_col].apply(lambda s: s.splitlines())
    df[cantidad_col] = df[cantidad_col].apply(lambda s: s.replace(',', '.').splitlines())

    # Emparejar y generar lista de tuplas (item, cantidad)
    def emparejar(row):
        items = [it.strip() for it in row[item_col]]
        cantidades_raw = [c.strip() for c in row[cantidad_col]]
        longitud = min(len(items), len(cantidades_raw))
        parejas = []
        for i in range(longitud):
            cantidad = None
            if cantidades_raw[i] != '':
                try:
                    cantidad = float(cantidades_raw[i])
                except ValueError:
                    cantidad = None
            parejas.append((items[i], cantidad))
        return parejas

    df['parejas'] = df.apply(emparejar, axis=1)

    # Explotar las parejas en filas
    df = df.explode('parejas')

    # Separar de nuevo en columnas
    df[[item_col, cantidad_col]] = pd.DataFrame(
        df['parejas'].tolist(), index=df.index
    )

    df = df.drop(columns=['parejas'])

    return df


  # Aplicar la función al DataFrame original
  insumos_aplicacion = expandir_items(insumos_nuevos, 'Item', 'Cantidad Usada por Item')


  # Llenar con No Aplica en pagado a
  insumos_aplicacion['Pagado a:'] = 'No Aplica'

  # Dar formato de mayuscula a columna de item eliminando posibles espacios en blanco
  insumos_aplicacion['Item'] = insumos_aplicacion['Item'].str.upper().str.strip()

  # dejar columnas necesarias
  insumos_aplicacion = insumos_aplicacion[['Marca temporal', 'Fecha Actividad', 'Mes Proyecto', 'Mes del Lote', 'Pagado a:', 'Lote',
                                          'Item Archivo Financiero', 'Item', 'Cantidad Usada por Item', 'Ciclo', 'Invernadero', 'Observaciones']]


  # Crear dicionario para reemplazar palabras para el item
  errores = ['FOSS61', 'CENTAURO', 'HUMUS ALFA', 'CALCINIT', 'SULFATO POTASIO', 'SULFATO DE MAGNESIO', 'SULFATO MAGNESIO',
             'ÁTOMIK', 'REBROTE', 'TRIVIA', 'CARBENDAZIM', 'FÓRUM', 'BELT', 'SUMAG', '10 20 20', 'SABERSOIL', 'SAFER SOIL',
             'HUMUS', 'CABRIOTOP', 'CIROMEX', 'KZYME', 'CIPEMETRINA', 'DIFECOL', 'CYMOCEB']

  # Crear listado con palabras correctas
  aciertos = ['FOSS 61', 'CENTAURO 720', 'HUMUS ALFA 15', 'YARA CALCINIT', 'SULFATO DE POTASIO', 'SULFATO DE MAGNESIO TÉCNICO', 'SULFATO DE MAGNESIO TÉCNICO',
              'ATOMIK', 'REBROTE 10-50-8', 'TRIVIA WP', 'CARBENDAZIN', 'FORUM 500', 'BELT SC', 'SUMAGGRANULADO', '10-20-20', 'SAFERSOIL', 'SAFERSOIL',
              'HUMUS ALFA 15', 'CABRIO TOP', 'CIROMEX BRIO', 'KZIME', 'CIPERMETRINA', 'DIFECOL 250', 'CYMOZEB']


  # Hacer reemplazo de listados
  insumos_aplicacion = insumos_aplicacion.replace(errores, aciertos)

  # Eliminar posibles espacios en blanco en columna de item
  clasificacion_insumos['Item'] = clasificacion_insumos['Item'].str.strip()

  # Asignar una claisificacion y unidad de medida con base a el item
  insumos_aplicacion = pd.merge(insumos_aplicacion, clasificacion_insumos, on='Item', how='left')

  # Llenar Nulos con revisar
  insumos_aplicacion[['Clasificación/Tipo Actividad', 'Unidad Medida']] = insumos_aplicacion[['Clasificación/Tipo Actividad', 'Unidad Medida']].fillna('Revisar')

  # Renomnbrar columnas
  insumos_aplicacion.rename(columns={'Fecha Actividad':'Fecha Compra/Aplicacion', 'Item Archivo Financiero Insumos': 'Item Archivo Financiero',
                                    'Cantidad Usada por Item':'Cantidad Comprada/Aplicada'}, inplace=True)

  # Crear columna Concepto igual a Aplicación
  insumos_aplicacion['Concepto'] = 'APLICACIÓN'

  # Crear la nueva columna 'Concepto P&L o Balance General'
  insumos_aplicacion['Concepto P&L o Balance General'] = insumos_aplicacion['Clasificación/Tipo Actividad'].apply(
      lambda x: 'Plántulas' if 'Plantas' in x else 'Insumos'
  )

  # Convertir columna 'Cantidad Comprada/Aplicada' a numero
  # insumos_aplicacion['Cantidad Comprada/Aplicada'] = insumos_aplicacion['Cantidad Comprada/Aplicada'].str.replace(',', '.')
  # insumos_aplicacion['Cantidad Comprada/Aplicada'] = insumos_aplicacion['Cantidad Comprada/Aplicada'].astype(float)

  # Dar formato de fecha
  insumos_aplicacion["Fecha Compra/Aplicacion"] = pd.to_datetime(insumos_aplicacion["Fecha Compra/Aplicacion"], dayfirst=True)
  
  # Convertir columnas a tipo numérico y rellenar NaN con 0.0
  insumos_aplicacion["Cantidad Comprada/Aplicada"] = pd.to_numeric(
      insumos_aplicacion["Cantidad Comprada/Aplicada"],
      errors="coerce"
  ).fillna(0.0)

  inventario["Cantidad Comprada/Aplicada"] = pd.to_numeric(
      inventario["Cantidad Comprada/Aplicada"],
      errors="coerce"
  ).fillna(0.0)
  inventario["Valor Unidad"] = pd.to_numeric(
      inventario["Valor Unidad"],
      errors="coerce"
  ).fillna(0.0)
  
  # Crear nuevas columnas en la base de aplicación para almacenar el valor unitario y total
  insumos_aplicacion["Valor Unidad"] = 0.0
  insumos_aplicacion["Total"] = 0.0


  # Implementar la lógica FIFO
  for index, row in insumos_aplicacion.iterrows():
    item = row["Item"]
    cantidad_necesaria = row["Cantidad Comprada/Aplicada"]

    if cantidad_necesaria > 0:
        # Filtrar inventario para este ítem
        inventario_disponible = inventario[inventario["Item"] == item].copy()

        total_asignado = 0.0
        valor_unitario_final = 0.0

        for idx, inv_row in inventario_disponible.iterrows():
            cantidad_disponible = inv_row["Cantidad Comprada/Aplicada"]
            valor_unitario    = inv_row["Valor Unidad"]

            if cantidad_disponible <= 0:
                continue

            if cantidad_necesaria <= cantidad_disponible:
                # Cubrir toda la necesidad con este lote
                total_asignado += cantidad_necesaria * valor_unitario
                valor_unitario_final = valor_unitario
                inventario.at[idx, "Cantidad Comprada/Aplicada"] = (
                    cantidad_disponible - cantidad_necesaria
                )
                cantidad_necesaria = 0.0
                break
            else:
                # Consumir todo el lote y seguir
                total_asignado += cantidad_disponible * valor_unitario
                cantidad_necesaria -= cantidad_disponible
                inventario.at[idx, "Cantidad Comprada/Aplicada"] = 0.0
                valor_unitario_final = valor_unitario

        # Guardar resultados en insumos_aplicacion
        insumos_aplicacion.at[index, "Valor Unidad"] = valor_unitario_final
        insumos_aplicacion.at[index, "Total"] = total_asignado

  # Descartar de inventario las cantidades compradas o aplicadas iguales a 0
  inventario = inventario[inventario['Cantidad Comprada/Aplicada'] != 0]
  
  # Modificar columna Total
  inventario['Total'] = inventario['Cantidad Comprada/Aplicada'] * inventario['Valor Unidad']

  # Devolver formato a columnas de fecha
  inventario =  inventario.sort_values(by=['Fecha Compra/Aplicacion']).reset_index(drop=True)
  #inventario['Fecha Compra/Aplicacion'] = inventario['Fecha Compra/Aplicacion'].dt.strftime('%d/%m/%Y')
  insumos_aplicacion['Fecha Compra/Aplicacion'] = insumos_aplicacion['Fecha Compra/Aplicacion'].dt.strftime('%d/%m/%Y')

  # Covertir en negativos la cantidad aplicada y el total
  insumos_aplicacion[['Cantidad Comprada/Aplicada', 'Total']] = -insumos_aplicacion[['Cantidad Comprada/Aplicada', 'Total']]

  # Reorganizar df de insumos aplicacion
  insumos_aplicacion = insumos_aplicacion[['Marca temporal', 'Fecha Compra/Aplicacion', 'Concepto', 'Mes Proyecto',
                                            'Mes del Lote', 'Pagado a:', 'Lote', 'Concepto P&L o Balance General',
                                            'Clasificación/Tipo Actividad', 'Item Archivo Financiero',
                                            'Cantidad Comprada/Aplicada', 'Item', 'Unidad Medida', 'Valor Unidad',
                                            'Total', 'Ciclo', 'Invernadero', 'Observaciones']]
  
  # Unir insumos de aplicacion con insumos compras
  insumos_total = pd.concat([historial_insumos,
                                insumos_aplicacion]
                                ).reset_index(drop=True)

  # Convertir las fechas a formato datetime para ordenarlas correctamente (temporal)
  insumos_total["Fecha Compra/Aplicacion"] = pd.to_datetime(insumos_total["Fecha Compra/Aplicacion"], dayfirst=True)
  insumos_total["Marca temporal"] = pd.to_datetime(insumos_total["Marca temporal"], dayfirst=True)

  # Ordenar df insumos total por fecha
  insumos_total =  insumos_total.sort_values(by=['Fecha Compra/Aplicacion', 'Marca temporal']).reset_index(drop=True)

  # Devolver formato a columnas en actualizado
  insumos_total['Fecha Compra/Aplicacion'] = insumos_total['Fecha Compra/Aplicacion'].dt.strftime('%d/%m/%Y')
  insumos_total['Marca temporal'] = insumos_total['Marca temporal'].dt.strftime('%d/%m/%Y %H:%M:%S')


  ############################################        ACTUALIZAR SHEET INSUMOS     ##############################


  ############# DEFINICION DE FUNCIONES ##########################

  #Limpiar contenido en hoja 
  def clear_range(spreadsheet_id, sheet_name="Insumos", range_=None):

      if range_ is not None:
          sheet_name=sheet_name+"!"
      else:
          range_=""

      dict_result = spreadsheet_service.spreadsheets().values().clear(
      spreadsheetId=spreadsheet_id, range=sheet_name + range_).execute()

  #Escribir en hoja a partir de dataframe. Empieza a pegar la información a partir de la primera fila vacía que encuentre.
  def write_range(spreadsheet_id, dataframe, sheet_name="Insumos", range_=None, include_headers=False):
      
          #rellenar los NaN con vacíos
          dataframe = dataframe.fillna("")
          
          #eliminar saltos de líneas y retornos de carro
          dataframe = dataframe.replace(r"\n","", regex=True).replace(r"\r","", regex=True)
          
          
          if range_ is not None:
            sheet_name=sheet_name+"!"
          else:
            range_=""

          if include_headers==True:
            content_values = dataframe.values.tolist()
            content_values.insert(0, dataframe.columns.tolist())
          else:
            content_values=dataframe.values.tolist()

          body = {
          'values': content_values
          }
          spreadsheet_service.spreadsheets().values().append(
          spreadsheetId=spreadsheet_id, body=body, valueInputOption='USER_ENTERED', range=sheet_name + range_).execute()



  # # #################   CARGAR SHEET Insumos    ##############

  # Limpiar Hoja Google sheets
  clear_range(spreadsheet_id=spreadsheet_id, sheet_name='Insumos')

  # Escribir df en hoja Google Sheets
  write_range(spreadsheet_id=spreadsheet_id, sheet_name='Insumos', dataframe=insumos_total, include_headers=True)

  # ########################################################################################################################

  # print(insumos_total)


  ############################################        ACTUALIZAR SHEET INVENTARIO     ##############################


  ############# DEFINICION DE FUNCIONES ##########################

  #Limpiar contenido en hoja 
  def clear_range(spreadsheet_id, sheet_name="Inventario", range_=None):

      if range_ is not None:
          sheet_name=sheet_name+"!"
      else:
          range_=""

      dict_result = spreadsheet_service.spreadsheets().values().clear(
      spreadsheetId=spreadsheet_id, range=sheet_name + range_).execute()

  #Escribir en hoja a partir de dataframe. Empieza a pegar la información a partir de la primera fila vacía que encuentre.
  def write_range(spreadsheet_id, dataframe, sheet_name="Inventario", range_=None, include_headers=False):
      
          #rellenar los NaN con vacíos
          dataframe = dataframe.fillna("")
          
          #eliminar saltos de líneas y retornos de carro
          dataframe = dataframe.replace(r"\n","", regex=True).replace(r"\r","", regex=True)
          
          
          if range_ is not None:
            sheet_name=sheet_name+"!"
          else:
            range_=""

          if include_headers==True:
            content_values = dataframe.values.tolist()
            content_values.insert(0, dataframe.columns.tolist())
          else:
            content_values=dataframe.values.tolist()

          body = {
          'values': content_values
          }
          spreadsheet_service.spreadsheets().values().append(
          spreadsheetId=spreadsheet_id, body=body, valueInputOption='USER_ENTERED', range=sheet_name + range_).execute()



  #################   CARGAR SHEET Inventario    ##############

  # Limpiar Hoja Google sheets
  clear_range(spreadsheet_id=spreadsheet_id, sheet_name='Inventario')

  # Escribir df en hoja Google Sheets
  write_range(spreadsheet_id=spreadsheet_id, sheet_name='Inventario', dataframe=inventario, include_headers=True)

  # ########################################################################################################################

  # print(inventario)


#############################################################################################################                                          
#############################          TRATAMIENTO PARA  VENTAS              ################################
#############################################################################################################

# Filtrar respuestas por formulario
respuestas_ventas = respuestas_ventas[respuestas_ventas['Marca temporal'] != '']


# Cambiar nombre a columna de unidad de minutos
respuestas_ventas.rename(columns={'Ciclo (solo número)\n':'Ciclo'}, inplace=True)

# Dar formato mayuscula a columna de clasificacion calidad
respuestas_ventas['Clasificación/Calidad'] = respuestas_ventas['Clasificación/Calidad'].str.upper()


# --- Helpers de conversión ---
def _to_number_or(value, empty_to=None):
    """
    Convierte strings tipo '1.234,56' -> 1234.56.
    Si está vacío ('', '-', None) devuelve empty_to.
    """
    if not isinstance(value, str):
        return value if value not in (None,) else empty_to
    s = value.strip()
    if s in ("", "-"):
        return empty_to
    return float(s.replace('.', '').replace(',', '.'))

def _strip_or_none(value):
    return value.strip() if isinstance(value, str) and value.strip() != "" else None


# --- Función para expandir con broadcast y relleno por defecto ---
def expandir_por_lineas(df, cols_expandir, converters=None):
    if converters is None:
        converters = {}

    filas = []
    for _, row in df.iterrows():
        listas = {}
        # 1) Preparar listas por columna (respetando saltos de línea)
        for col in cols_expandir:
            cell = row.get(col, None)
            raw = "" if pd.isna(cell) else str(cell)
            partes = raw.splitlines()
            # Si está vacío, dejamos lista vacía (para luego rellenar)
            conv = converters.get(col, lambda x: x.strip())
            procesadas = []
            for p in partes:
                p = p.strip()
                try:
                    procesadas.append(conv(p))
                except Exception:
                    procesadas.append(None)
            listas[col] = procesadas

        # 2) Largo objetivo = máximo entre columnas (broadcast).
        #    Si todas están vacías, forzamos 1.
        longitudes = [len(lst) for lst in listas.values()] if listas else [0]
        longitud = max(longitudes) if any(longitudes) else 1

        # 3) Generar subfilas con relleno por defecto cuando falte
        for i in range(longitud):
            nueva = row.copy()
            for col in cols_expandir:
                col_list = listas[col]
                if i < len(col_list):
                    nueva[col] = col_list[i]
                else:
                    # Relleno por defecto: conv('') para respetar reglas de la columna
                    conv = converters.get(col, lambda x: x.strip())
                    try:
                        nueva[col] = conv("")
                    except Exception:
                        nueva[col] = None
            filas.append(nueva)

    return pd.DataFrame(filas).reset_index(drop=True)


# --- Convertidores solicitados ---
converters = {
    # Texto
    'Clasificación/Calidad': _strip_or_none,

    # Numéricos
    # Cantidad -> vacío => None
    'Cantidad': lambda s: _to_number_or(s, empty_to=None),

    # Valor Unidad -> vacío => 0.0
    'Valor Unidad': lambda s: _to_number_or(s, empty_to=0.0),

    # (Si aún usas estas columnas previas:)
    'Pagado a:': _strip_or_none,
    'Unidad (minutos trabajados)': lambda s: _to_number_or(s, empty_to=None),
}

# Expansión
cols_expandir = ['Clasificación/Calidad', 'Cantidad', 'Valor Unidad']
respuestas_ventas = expandir_por_lineas(respuestas_ventas, cols_expandir, converters=converters)

# Crear columna Ciclo
respuestas_ventas['Ciclo'] = 'CICLO ' + respuestas_ventas['Ciclo']

# Crear columna de mes, temporalmente
respuestas_ventas['Mes'] = respuestas_ventas['Fecha Cosecha'].str[-7:]

# Obtener el mes del proyecto segun la fecha de actividad
respuestas_ventas = pd.merge(respuestas_ventas, mes_proyecto, how='left', on=['Mes'])

# Eliminar columnas de mes
respuestas_ventas.drop(columns={'Mes'}, inplace=True)

# Crear columna Mes Proyecto
respuestas_ventas['Mes Proyecto'] = 'Mes ' + respuestas_ventas['Mes Proyecto']

# Crear columna semana del ciclo (vacio inicialmente)
respuestas_ventas['Semana del Ciclo Productivo'] = ''

# Crear columna unidades igual a kg
respuestas_ventas['Unidades'] = 'KG'

# Crear columna total con cantidad * valor unidad
respuestas_ventas['Total'] = respuestas_ventas['Cantidad'] * respuestas_ventas['Valor Unidad']

# Dejar columnas necesarias
ventas = respuestas_ventas[['Marca temporal', 'Fecha Cosecha', 'Fecha Venta','Mes Proyecto', 'Semana del Ciclo Productivo', 'Comprador', 'Lote',
                            'Clasificación/Calidad', 'Cantidad', 'Unidades', 'Valor Unidad', 'Total', 'Ciclo', 'Invernadero']]


# Convertir 'Fecha Cosecha' a datetime
ventas['Fecha Cosecha'] = pd.to_datetime(ventas['Fecha Cosecha'], format='%d/%m/%Y')

# Descartar respuestas de formulario
ventas_historia = ventas_historia[ventas_historia['Marca temporal'] == '']

# Crear listado con columnas numericas para dar formato numerico
numericas_3 = ['Valor Unidad', 'Total', 'Cantidad']

# Dar formato numerico
for numero3 in numericas_3:
  ventas_historia[numero3] = ventas_historia[numero3].str.replace('.','')
  ventas_historia[numero3] = ventas_historia[numero3].str.replace(',','.')
  ventas_historia[numero3] = pd.to_numeric(ventas_historia[numero3])
  
# Convertir 'Fecha Cosecha' a datetime
ventas_historia['Fecha Cosecha'] = pd.to_datetime(ventas_historia['Fecha Cosecha'], format='%d/%m/%Y')

# Crear df actividades con base a la fecha y actividades de jornales
actividades = jornales[['Fecha Actividad', 'Lote', 'Invernadero', 'Clasificación/Tipo Actividad']]

# Filtrar las actividades para obtener solo aquellas que corresponden a Erradicación Plantas
actividades = actividades[actividades['Clasificación/Tipo Actividad'] == 'Erradicación Plantas'].drop_duplicates()

# Renombrar columna d efecha d eactividad a fecha de cosecha para el merge
actividades.rename(columns={'Fecha Actividad': 'Fecha Cosecha'}, inplace=True)

# Asignar la actividad de recoleccion a ventas historico
ventas_historia = pd.merge(ventas_historia, actividades, on =['Fecha Cosecha', 'Lote', 'Invernadero'], how='left')

# Llenar nulos de actividad con 'Sin Cosechar'
ventas_historia['Clasificación/Tipo Actividad'] = ventas_historia['Clasificación/Tipo Actividad'].fillna('Sin Cosechar')


# Asignar la actividad de recoleccion a ventas actual
ventas = pd.merge(ventas, actividades, on =['Fecha Cosecha', 'Lote', 'Invernadero'], how='left')

# Llenar nulos de actividad con 'Sin Cosechar'
ventas['Clasificación/Tipo Actividad'] = ventas['Clasificación/Tipo Actividad'].fillna('Sin Cosechar')


######################      OBTENER BASES DEL HISTORICO PARA INICIAR    ########################


# =========================
# Utilidades
# =========================
def lunes_de(fecha: pd.Timestamp) -> pd.Timestamp:
    """Devuelve el lunes (weekday=0) de la semana de 'fecha'."""
    if pd.isna(fecha): 
        return fecha
    return fecha - pd.Timedelta(days=fecha.weekday())

def semanas_transcurridas_lunes_domingo(start: pd.Timestamp, end: pd.Timestamp) -> int:
    """
    Cuántas semanas completas (lunes–domingo) hay entre start y end.
    Si están en la misma semana (mismo lunes de referencia), devuelve 0.
    """
    if pd.isna(start) or pd.isna(end):
        return 0
    start_mon = lunes_de(start)
    end_mon   = lunes_de(end)
    diff_days = (end_mon - start_mon).days
    return max(0, diff_days // 7)

def quitar_acentos(texto):
    if not isinstance(texto, str):
        return texto
    nfkd = unicodedata.normalize("NFKD", texto)
    return "".join([c for c in nfkd if not unicodedata.combining(c)])

TARGET_RECOLECCION = quitar_acentos('Erradicación Plantas').lower()

# =========================
# 1) Historial
# =========================
# Asegurar fechas y ordenar
ventas_historia['Fecha Cosecha'] = pd.to_datetime(
    ventas_historia['Fecha Cosecha'], format='%d/%m/%Y', errors='coerce'
)
ventas_historia.sort_values(by=['Lote', 'Invernadero', 'Fecha Cosecha'], inplace=True)

# Normalizar actividad en historial
ventas_historia['actividad_norm'] = (
    ventas_historia['Clasificación/Tipo Actividad']
    .astype(str).str.strip().str.lower().apply(quitar_acentos)
)

def resumen_anchor(grp: pd.DataFrame) -> pd.Series:
    """
    Devuelve sólo el ANCLA (lunes) del último reset:
      - Si hay recolección en el historial: ancla = lunes(fecha_última_recolección)
      - Si no hay, ancla = lunes(fecha_mínima_del_grupo) para iniciar desde 0
    """
    grp = grp.copy()
    grp.sort_values('Fecha Cosecha', inplace=True)
    reco_fechas = grp.loc[grp['actividad_norm'] == TARGET_RECOLECCION, 'Fecha Cosecha']
    if not reco_fechas.empty:
        anchor = lunes_de(reco_fechas.max())
    else:
        # Sin recolecciones históricas → usamos la primera fecha como referencia
        anchor = lunes_de(grp['Fecha Cosecha'].min())
    return pd.Series({'AnchorMonday': anchor})

resultados = ventas_historia.groupby(['Lote', 'Invernadero'], group_keys=False).apply(resumen_anchor)
resultados.reset_index(inplace=True)
resultados_dict = resultados.set_index(['Lote', 'Invernadero']).to_dict('index')

# Limpieza auxiliar del historial
ventas_historia.drop(columns=['actividad_norm'], inplace=True, errors='ignore')

# =========================
# 2) Nuevos registros "ventas"
# =========================
ventas['Fecha Cosecha']  = pd.to_datetime(ventas['Fecha Cosecha'],  format='%d/%m/%Y', errors='coerce')
ventas['Marca temporal'] = pd.to_datetime(ventas['Marca temporal'], errors='coerce')

ventas.sort_values(
    by=['Lote', 'Invernadero', 'Fecha Cosecha', 'Marca temporal'],
    inplace=True
)

def asignar_semana_del_lote(grupo: pd.DataFrame) -> pd.DataFrame:
    grupo = grupo.copy()
    key = (grupo['Lote'].iloc[0], grupo['Invernadero'].iloc[0])
    hist = resultados_dict.get(key, None)

    # Establecer ancla inicial
    if hist is None or pd.isna(hist.get('AnchorMonday', pd.NaT)):
        # Si no hay ancla histórica válida, arrancamos en el lunes de la primera fecha del grupo
        primera_fecha = pd.to_datetime(grupo['Fecha Cosecha'].iloc[0], errors='coerce')
        anchor = lunes_de(primera_fecha) if pd.notna(primera_fecha) else pd.NaT
    else:
        anchor = hist['AnchorMonday']

    semanas = []
    for _, fila in grupo.iterrows():
        actividad_norm = quitar_acentos(str(fila['Clasificación/Tipo Actividad'])).lower().strip()
        f = fila['Fecha Cosecha']

        if pd.isna(f):
            # Fecha inválida → hereda la última semana calculada
            semanas.append(int(semanas[-1]) if semanas else 0)
            continue

        if actividad_norm == TARGET_RECOLECCION:
            # Reset inmediato en la recolección
            anchor = lunes_de(f)
            semanas.append(0)
            continue

        # Semana = lunes_transcurridos desde el ancla (última recolección o inicial)
        semana_actual = semanas_transcurridas_lunes_domingo(anchor, f)
        semanas.append(int(semana_actual))

    grupo['Semana del Ciclo Productivo'] = pd.to_numeric(semanas, downcast='integer')
    return grupo

ventas_actualizado = ventas.groupby(['Lote', 'Invernadero'], group_keys=False).apply(asignar_semana_del_lote)
ventas_actualizado.reset_index(drop=True, inplace=True)


ventas_actualizado = ventas.groupby(['Lote', 'Invernadero'], group_keys=False).apply(asignar_semana_del_lote)
ventas_actualizado.reset_index(drop=True, inplace=True)

ventas_actualizado.drop(columns={'Clasificación/Tipo Actividad'}, inplace=True)

# Ordenar df ventas actualizado
ventas_actualizado =  ventas_actualizado.sort_values(by=['Fecha Cosecha', 'Marca temporal']).reset_index(drop=True)

# Devolver formato a columnas en actualizado
ventas_actualizado['Fecha Cosecha'] = ventas_actualizado['Fecha Cosecha'].dt.strftime('%d/%m/%Y')
ventas_actualizado['Marca temporal'] = ventas_actualizado['Marca temporal'].dt.strftime('%d/%m/%Y %H:%M:%S')


# Ordenar df ventas hisotria
ventas_historia =  ventas_historia.sort_values(by='Fecha Cosecha').reset_index(drop=True)
ventas_historia.drop(columns={'Clasificación/Tipo Actividad'}, inplace=True)

# Devolver formato a columna de fecha a historia
ventas_historia['Fecha Cosecha'] = ventas_historia['Fecha Cosecha'].dt.strftime('%d/%m/%Y')

# Concatenar los DataFrames
ventas_consolidado = pd.concat([ventas_historia, ventas_actualizado]).reset_index(drop=True)


############################################        ACTUALIZAR SHEET ventas     ##############################


############# DEFINICION DE FUNCIONES ##########################

#Limpiar contenido en hoja 
def clear_range(spreadsheet_id, sheet_name="Ventas", range_=None):

    if range_ is not None:
        sheet_name=sheet_name+"!"
    else:
        range_=""

    dict_result = spreadsheet_service.spreadsheets().values().clear(
    spreadsheetId=spreadsheet_id, range=sheet_name + range_).execute()

#Escribir en hoja a partir de dataframe. Empieza a pegar la información a partir de la primera fila vacía que encuentre.
def write_range(spreadsheet_id, dataframe, sheet_name="Ventas", range_=None, include_headers=False):
    
        #rellenar los NaN con vacíos
        dataframe = dataframe.fillna("")
        
        #eliminar saltos de líneas y retornos de carro
        dataframe = dataframe.replace(r"\n","", regex=True).replace(r"\r","", regex=True)
        
        
        if range_ is not None:
          sheet_name=sheet_name+"!"
        else:
          range_=""

        if include_headers==True:
          content_values = dataframe.values.tolist()
          content_values.insert(0, dataframe.columns.tolist())
        else:
          content_values=dataframe.values.tolist()

        body = {
        'values': content_values
        }
        spreadsheet_service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id, body=body, valueInputOption='USER_ENTERED', range=sheet_name + range_).execute()



#################   CARGAR SHEET VENTAS    ##############

# Limpiar Hoja Google sheets
clear_range(spreadsheet_id=spreadsheet_id, sheet_name='Ventas')

# Escribir df en hoja Google Sheets
write_range(spreadsheet_id=spreadsheet_id, sheet_name='Ventas', dataframe=ventas_consolidado, include_headers=True)

########################################################################################################################

# #print(ventas_consolidado)


