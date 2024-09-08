import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import unicodedata

# Función para normalizar texto (eliminar acentos y convertir a minúsculas)
def normalize_text(text):
    if isinstance(text, str):
        text = text.strip().lower()
        text = ''.join(
            c for c in unicodedata.normalize('NFD', text)
            if unicodedata.category(c) != 'Mn'
        )
    return text

# Configurar la conexión con Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('piola-434923-8017a202381d.json', scope)
client = gspread.authorize(creds)

# Leer el archivo Excel
df = pd.read_excel('movimientoshistoricos.xlsx')

# Renombrar las columnas según el mapeo especificado
df = df.rename(columns={
    'Concert.': 'Fecha',
    'Cant. titulos': 'Cantidad',
    'Precio': 'Precio',
    'Monto': 'Total',
    'Tipo Mov.': 'tipo_mov'
})

# Normalizar la columna 'tipo_mov'
df['tipo_mov'] = df['tipo_mov'].apply(normalize_text)

# Extraer el tipo de movimiento y el símbolo
df[['Transaccion', 'Activo']] = df['tipo_mov'].str.extract(r'([^\(]+)\s*\(?(\w+)?\)?')

# Normalizar la columna 'Transaccion'
df['Transaccion'] = df['Transaccion'].apply(normalize_text)

# Mapear sinónimos a movimientos válidos
def map_transaction(transaction):
    if 'compra' in transaction or 'suscripcion' in transaction:
        return 'compra'
    elif 'deposito' in transaction:
        return 'deposito'
    elif 'dividendo' in transaction or 'dividendos' in transaction:
        return 'dividendo'
    return transaction

df['Transaccion'] = df['Transaccion'].apply(map_transaction)

# Filtrar los movimientos válidos
valid_movements = ['compra', 'venta', 'deposito', 'extraccion', 'interes', 'dividendo']
df_valid = df[df['Transaccion'].isin(valid_movements)]

# Eliminar apóstrofes iniciales en la columna 'Fecha'
df_valid['Fecha'] = df_valid['Fecha'].astype(str).str.lstrip("'")

# Formatear las fechas
def format_date(date):
    try:
        return pd.to_datetime(date, errors='coerce').strftime('%d/%m/%Y')
    except Exception as e:
        return date

df_valid['Fecha'] = df_valid['Fecha'].apply(format_date)

# Reemplazar NaN por cadenas vacías
df_valid = df_valid.fillna('')

# Reemplazar valores vacíos en 'Activo' y 'Cantidad'


# Ordenar por fecha
df_valid = df_valid.sort_values(by='Fecha')

# Formato final
df_valid['Transaccion'] = df_valid['Transaccion'].str.capitalize()
df_valid['Activo'] = df_valid['Activo'].str.upper()
df_valid['Activo'] = df_valid['Activo'].replace('', 'Efectivo')
df_valid['Cantidad'] = df_valid['Cantidad'].replace(0, 1)

# Conectar con la hoja de Google Sheets
spreadsheet = client.open_by_key('1vUobNvISKw0WkX72luynVW35pPUeMI55nLJCqcEEh_E')
worksheet = spreadsheet.worksheet('Transacciones')

# Obtener el número de filas actuales para agregar los datos después de las cabeceras
existing_data = worksheet.get_all_values()
next_row = len(existing_data) + 1

# Preparar los datos para la hoja de Google Sheets en el orden correcto
df_valid = df_valid[['Fecha', 'Activo', 'Transaccion', 'Cantidad', 'Precio', 'Total']]

# Escribir los datos en la hoja de Google Sheets a partir de la fila 3
worksheet.update(f'A3', df_valid.values.tolist())

# Informar sobre los movimientos pasados y no pasados
total_movements = len(df_valid)
invalid_movements = df[~df['Transaccion'].isin(valid_movements)]
total_invalid = len(invalid_movements)

print(f"Movimientos pasados: {total_movements}")
print(f"Movimientos no pasados: {total_invalid}")
print("Movimientos no pasados:")
print(invalid_movements)