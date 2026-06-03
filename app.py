from flask import Flask, render_template, request, redirect, url_for, session, send_file, flash
import pandas as pd
import sqlite3
import os
from functools import wraps
from io import BytesIO
import openpyxl

app = Flask(__name__)
app.secret_key = 'clave_secreta_cambiala_luego'  # Cámbiala por algo seguro

# Base de datos SQLite
DATABASE = 'datos_app.db'

# Lista de usuarios (normalizada, sin duplicados)
USUARIOS = [
    "ANA JULCA", "ANA MAC", "CARLOS ROBLES", "CRISTINA ALBAN", "DARLENE BALTODANO",
    "DIEGO CORTIJO", "GEORGINA HUAMAN", "GUSTAVO VARGAS", "JEAN OLIVARES", "JHONATAN ALAYO",
    "JORGE FEIJOO", "LORENA CORTEZ", "LUCIA ARANA", "LUIS VALDIVIESO", "MARINA SISNIEGAS",
    "MARLENE ARTEAGA", "RONALD SOTO", "ROSA CERIN", "SERGIO BENITES", "VICTOR VARAS", "YANELA MEZA"
]

# Estado de pausa (se guarda en archivo para persistir entre reinicios)
PAUSA_FILE = 'pausa.txt'

def get_pausa():
    if os.path.exists(PAUSA_FILE):
        with open(PAUSA_FILE, 'r') as f:
            return f.read().strip() == 'True'
    return False

def set_pausa(valor):
    with open(PAUSA_FILE, 'w') as f:
        f.write(str(valor))

# Decorador para requerir login
def login_requerido(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'usuario' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

# --- Funciones para manejar la base de datos ---
def init_db():
    """Crea las tablas si no existen, basado en el Excel si ya se subió"""
    conn = sqlite3.connect(DATABASE)
    c = conn.cursor()
    # Tabla para saber qué columnas tiene cada hoja
    c.execute('''CREATE TABLE IF NOT EXISTS metadata (
        hoja TEXT PRIMARY KEY,
        columnas TEXT,
        filas INTEGER
    )''')
    conn.commit()
    conn.close()

def get_columnas(hoja):
    conn = sqlite3.connect(DATABASE)
    c = conn.cursor()
    c.execute("SELECT columnas FROM metadata WHERE hoja = ?", (hoja,))
    row = c.fetchone()
    conn.close()
    if row:
        return row[0].split(',')
    return None

def crear_tabla_dinamica(hoja, df):
    """Crea o reemplaza la tabla para una hoja con las columnas del DataFrame"""
    conn = sqlite3.connect(DATABASE)
    c = conn.cursor()
    # Borrar tabla si existe
    c.execute(f"DROP TABLE IF EXISTS [{hoja}]")
    # Crear tabla con columnas del DataFrame
    columnas_sql = []
    for col in df.columns:
        # Escapar nombres con espacios o caracteres especiales
        col_escapado = f'"{col}"'
        columnas_sql.append(f"{col_escapado} TEXT")
    # Agregar columna id único como PRIMARY KEY (usaremos la columna BF "ID UNICO")
    # Asumimos que existe "ID UNICO" en df. Si no, la creamos.
    if "ID UNICO" not in df.columns:
        # Si no existe, la creamos con valores numéricos
        df["ID UNICO"] = range(1, len(df)+1)
    create_sql = f'CREATE TABLE [{hoja}] ({", ".join(columnas_sql)}, "ID UNICO" TEXT PRIMARY KEY)'
    c.execute(create_sql)
    # Insertar datos
    for _, row in df.iterrows():
        placeholders = ','.join(['?'] * len(df.columns))
        valores = [str(row[col]) if pd.notna(row[col]) else '' for col in df.columns]
        # Asegurar que ID UNICO sea string
        id_valor = str(row["ID UNICO"])
        insert_sql = f'INSERT OR REPLACE INTO [{hoja}] VALUES ({placeholders})'
        c.execute(insert_sql, valores)
    # Guardar metadata
    c.execute("INSERT OR REPLACE INTO metadata (hoja, columnas, filas) VALUES (?, ?, ?)",
              (hoja, ','.join(df.columns), len(df)))
    conn.commit()
    conn.close()

def cargar_excel_en_db(archivo_excel):
    """Lee el Excel subido, procesa las dos hojas y actualiza la base de datos"""
    xls = pd.ExcelFile(archivo_excel)
    hojas_esperadas = ['SCD-2026', 'SAF-2026']
    for hoja in hojas_esperadas:
        if hoja in xls.sheet_names:
            df = pd.read_excel(archivo_excel, sheet_name=hoja, dtype=str)
            df = df.fillna('')  # Reemplazar NaN por cadena vacía
            # Asegurar que las columnas F y G existan (si no, crearlas vacías)
            # Los nombres de columnas según el usuario:
            col_f = "AA ENCARGADO/MATERIALES Y GESTION"
            col_g = "AA ENCARGADO/EQUIPAMIENTO"
            if col_f not in df.columns:
                df[col_f] = ''
            if col_g not in df.columns:
                df[col_g] = ''
            # Asegurar ID UNICO
            if "ID UNICO" not in df.columns:
                df["ID UNICO"] = [f"{hoja}_{i}" for i in range(len(df))]
            crear_tabla_dinamica(hoja, df)
        else:
            # Si falta una hoja, crear tabla vacía con al menos columnas F y G
            df_vacio = pd.DataFrame(columns=["ID UNICO", col_f, col_g])
            crear_tabla_dinamica(hoja, df_vacio)

def exportar_db_a_excel():
    """Exporta todas las tablas a un archivo Excel con dos hojas"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for hoja in ['SCD-2026', 'SAF-2026']:
            conn = sqlite3.connect(DATABASE)
            try:
                df = pd.read_sql_query(f'SELECT * FROM "{hoja}"', conn)
                # Convertir a DataFrame y quitar la columna de índice si existe
                df = df.fillna('')
                df.to_excel(writer, sheet_name=hoja, index=False)
            except:
                # Si no existe la tabla, crear hoja vacía
                pd.DataFrame().to_excel(writer, sheet_name=hoja, index=False)
            conn.close()
    output.seek(0)
    return output

def obtener_filas_usuario(hoja, usuario):
    """Retorna un DataFrame con las filas donde usuario está en F o G"""
    conn = sqlite3.connect(DATABASE)
    col_f = "AA ENCARGADO/MATERIALES Y GESTION"
    col_g = "AA ENCARGADO/EQUIPAMIENTO"
    # La consulta SQL: seleccionar todas las columnas de la tabla hoja
    columnas = get_columnas(hoja)
    if not columnas:
        conn.close()
        return pd.DataFrame()
    # Construir SELECT *
    query = f'SELECT * FROM "{hoja}" WHERE "{col_f}" = ? OR "{col_g}" = ?'
    df = pd.read_sql_query(query, conn, params=(usuario, usuario))
    conn.close()
    return df

# --- Rutas de la web ---
@app.route('/')
def index():
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        usuario = request.form.get('usuario')
        if usuario in USUARIOS:
            session['usuario'] = usuario
            return redirect(url_for('dashboard'))
        else:
            flash('Usuario no válido')
    return render_template('login.html', usuarios=USUARIOS)

@app.route('/logout')
def logout():
    session.pop('usuario', None)
    return redirect(url_for('login'))

@app.route('/dashboard')
@login_requerido
def dashboard():
    hoja_actual = request.args.get('hoja', 'SCD-2026')
    pausa = get_pausa()
    return render_template('dashboard.html', 
                           usuario=session['usuario'],
                           hoja_actual=hoja_actual,
                           pausa=pausa)

@app.route('/get_datos')
@login_requerido
def get_datos():
    hoja = request.args.get('hoja', 'SCD-2026')
    usuario = session['usuario']
    df = obtener_filas_usuario(hoja, usuario)
    # Convertir a lista de diccionarios para JSON
    datos = df.to_dict(orient='records')
    columnas = list(df.columns)
    return {'datos': datos, 'columnas': columnas}

@app.route('/guardar_fila', methods=['POST'])
@login_requerido
def guardar_fila():
    if get_pausa():
        return {'error': 'El sistema está pausado, no se pueden guardar cambios'}, 403
    hoja = request.json.get('hoja')
    fila_id = request.json.get('id_unico')
    campos = request.json.get('campos')  # diccionario {columna: valor}
    conn = sqlite3.connect(DATABASE)
    c = conn.cursor()
    # Construir UPDATE dinámico
    set_clause = ', '.join([f'"{k}" = ?' for k in campos.keys()])
    valores = list(campos.values()) + [fila_id]
    sql = f'UPDATE "{hoja}" SET {set_clause} WHERE "ID UNICO" = ?'
    try:
        c.execute(sql, valores)
        conn.commit()
        return {'ok': True}
    except Exception as e:
        return {'error': str(e)}, 500
    finally:
        conn.close()

@app.route('/subir_excel', methods=['POST'])
@login_requerido
def subir_excel():
    # Solo permite al usuario administrador (puedes cambiar el nombre)
    if session['usuario'] != 'ANA JULCA':  # Cambia por el usuario que será admin
        return {'error': 'No autorizado'}, 403
    if 'archivo' not in request.files:
        return {'error': 'No se envió archivo'}, 400
    archivo = request.files['archivo']
    if archivo.filename == '':
        return {'error': 'Archivo vacío'}, 400
    if not archivo.filename.endswith(('.xlsx', '.xls')):
        return {'error': 'Formato no válido, debe ser .xlsx o .xls'}, 400
    # Guardar temporalmente
    temp_path = 'temp_upload.xlsx'
    archivo.save(temp_path)
    try:
        cargar_excel_en_db(temp_path)
        os.remove(temp_path)
        return {'ok': True, 'mensaje': 'Excel cargado correctamente'}
    except Exception as e:
        os.remove(temp_path)
        return {'error': str(e)}, 500

@app.route('/exportar')
@login_requerido
def exportar():
    # Permiso a cualquier usuario? O solo admin. Lo dejo a todos.
    excel_data = exportar_db_a_excel()
    return send_file(excel_data, as_attachment=True, download_name='datos_exportados.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/toggle_pausa', methods=['POST'])
@login_requerido
def toggle_pausa():
    if session['usuario'] != 'ANA JULCA':
        return {'error': 'No autorizado'}, 403
    nuevo_estado = not get_pausa()
    set_pausa(nuevo_estado)
    return {'pausa': nuevo_estado}

if __name__ == '__main__':
    init_db()
    # Si no hay datos, crear tablas vacías (para que la app arranque)
    for hoja in ['SCD-2026', 'SAF-2026']:
        if not get_columnas(hoja):
            df_vacio = pd.DataFrame(columns=["ID UNICO", "AA ENCARGADO/MATERIALES Y GESTION", "AA ENCARGADO/EQUIPAMIENTO"])
            crear_tabla_dinamica(hoja, df_vacio)
    app.run(debug=True)