from flask import Flask, render_template, request, redirect, url_for, session, send_file, flash, jsonify
import pandas as pd
import sqlite3
import os
from functools import wraps
from io import BytesIO
import openpyxl
import sys

app = Flask(__name__)
app.secret_key = 'clave_secreta_cambiala_luego'

DATABASE = 'datos_app.db'

USUARIOS = [
    "ANA JULCA", "ANA MAC", "CARLOS ROBLES", "CRISTINA ALBAN", "DARLENE BALTODANO",
    "DIEGO CORTIJO", "GEORGINA HUAMAN", "GUSTAVO VARGAS", "JEAN OLIVARES", "JHONATAN ALAYO",
    "JORGE FEIJOO", "LORENA CORTEZ", "LUCIA ARANA", "LUIS VALDIVIESO", "MARINA SISNIEGAS",
    "MARLENE ARTEAGA", "RONALD SOTO", "ROSA CERIN", "SERGIO BENITES", "VICTOR VARAS", "YANELA MEZA"
]

PAUSA_FILE = 'pausa.txt'

# --- Administrador: cámbialo al nombre que quieras ---
ADMIN_NOMBRE = "JEAN OLIVARES"   # <--- CAMBIA AQUÍ EL ADMINISTRADOR

def get_pausa():
    if os.path.exists(PAUSA_FILE):
        with open(PAUSA_FILE, 'r') as f:
            return f.read().strip() == 'True'
    return False

def set_pausa(valor):
    with open(PAUSA_FILE, 'w') as f:
        f.write(str(valor))

def login_requerido(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'usuario' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

# --- Inicialización de la base de datos (se ejecuta al arrancar) ---
def init_db():
    conn = sqlite3.connect(DATABASE)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS metadata (
        hoja TEXT PRIMARY KEY,
        columnas TEXT,
        filas INTEGER
    )''')
    conn.commit()
    conn.close()
    print("Base de datos inicializada (tabla metadata creada)", file=sys.stderr)

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
    conn = sqlite3.connect(DATABASE)
    c = conn.cursor()
    c.execute(f"DROP TABLE IF EXISTS [{hoja}]")
    columnas_sql = []
    for col in df.columns:
        col_escapado = f'"{col}"'
        columnas_sql.append(f"{col_escapado} TEXT")
    create_sql = f'CREATE TABLE [{hoja}] ({", ".join(columnas_sql)})'
    c.execute(create_sql)
    for _, row in df.iterrows():
        placeholders = ','.join(['?'] * len(df.columns))
        valores = [str(row[col]) if pd.notna(row[col]) else '' for col in df.columns]
        insert_sql = f'INSERT INTO [{hoja}] VALUES ({placeholders})'
        c.execute(insert_sql, valores)
    c.execute("INSERT OR REPLACE INTO metadata (hoja, columnas, filas) VALUES (?, ?, ?)",
              (hoja, ','.join(df.columns), len(df)))
    conn.commit()
    conn.close()

def cargar_excel_en_db(archivo_excel):
    xls = pd.ExcelFile(archivo_excel)
    hojas_esperadas = ['SCD-2026', 'SAF-2026']
    col_f = "AA ENCARGADO/MATERIALES Y GESTION"
    col_g = "AA ENCARGADO/EQUIPAMIENTO"
    for hoja in hojas_esperadas:
        if hoja in xls.sheet_names:
            df = pd.read_excel(archivo_excel, sheet_name=hoja, dtype=str)
            df = df.fillna('')
            if col_f not in df.columns:
                df[col_f] = ''
            if col_g not in df.columns:
                df[col_g] = ''
            crear_tabla_dinamica(hoja, df)
        else:
            # Crear tabla vacía
            df_vacio = pd.DataFrame(columns=[col_f, col_g])
            crear_tabla_dinamica(hoja, df_vacio)

def exportar_db_a_excel():
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for hoja in ['SCD-2026', 'SAF-2026']:
            conn = sqlite3.connect(DATABASE)
            try:
                df = pd.read_sql_query(f'SELECT * FROM "{hoja}"', conn)
                df = df.fillna('')
                df.to_excel(writer, sheet_name=hoja, index=False)
            except Exception as e:
                print(f"Error exportando {hoja}: {e}", file=sys.stderr)
                pd.DataFrame().to_excel(writer, sheet_name=hoja, index=False)
            conn.close()
    output.seek(0)
    return output

def obtener_filas_usuario(hoja, usuario):
    conn = sqlite3.connect(DATABASE)
    col_f = "AA ENCARGADO/MATERIALES Y GESTION"
    col_g = "AA ENCARGADO/EQUIPAMIENTO"
    query = f'SELECT rowid, * FROM "{hoja}" WHERE "{col_f}" = ? OR "{col_g}" = ?'
    df = pd.read_sql_query(query, conn, params=(usuario, usuario))
    conn.close()
    return df

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
                           pausa=pausa,
                           admin=ADMIN_NOMBRE)

@app.route('/get_datos')
@login_requerido
def get_datos():
    hoja = request.args.get('hoja', 'SCD-2026')
    usuario = session['usuario']
    df = obtener_filas_usuario(hoja, usuario)
    datos = df.to_dict(orient='records')
    columnas = list(df.columns)
    return jsonify({'datos': datos, 'columnas': columnas})

@app.route('/guardar_fila', methods=['POST'])
@login_requerido
def guardar_fila():
    if get_pausa():
        return jsonify({'error': 'Sistema pausado'}), 403
    hoja = request.json.get('hoja')
    rowid = request.json.get('rowid')
    campos = request.json.get('campos')
    conn = sqlite3.connect(DATABASE)
    c = conn.cursor()
    set_clause = ', '.join([f'"{k}" = ?' for k in campos.keys()])
    valores = list(campos.values()) + [rowid]
    sql = f'UPDATE "{hoja}" SET {set_clause} WHERE rowid = ?'
    try:
        c.execute(sql, valores)
        conn.commit()
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        conn.close()

@app.route('/subir_excel', methods=['POST'])
@login_requerido
def subir_excel():
    if session['usuario'] != ADMIN_NOMBRE:
        return jsonify({'error': 'No autorizado'}), 403
    if 'archivo' not in request.files:
        return jsonify({'error': 'No se envió archivo'}), 400
    archivo = request.files['archivo']
    if archivo.filename == '':
        return jsonify({'error': 'Archivo vacío'}), 400
    if not archivo.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'Formato no válido, debe ser .xlsx o .xls'}), 400
    temp_path = 'temp_upload.xlsx'
    archivo.save(temp_path)
    try:
        cargar_excel_en_db(temp_path)
        os.remove(temp_path)
        return jsonify({'ok': True, 'mensaje': 'Excel cargado correctamente'})
    except Exception as e:
        os.remove(temp_path)
        print(f"Error al cargar Excel: {e}", file=sys.stderr)
        return jsonify({'error': str(e)}), 500

@app.route('/exportar')
@login_requerido
def exportar():
    excel_data = exportar_db_a_excel()
    return send_file(excel_data, as_attachment=True, download_name='datos_exportados.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/toggle_pausa', methods=['POST'])
@login_requerido
def toggle_pausa():
    if session['usuario'] != ADMIN_NOMBRE:
        return jsonify({'error': 'No autorizado'}), 403
    nuevo_estado = not get_pausa()
    set_pausa(nuevo_estado)
    return jsonify({'pausa': nuevo_estado})

# --- Inicialización al arrancar (para gunicorn) ---
init_db()

# Crear tablas vacías si no existen (por si nunca se subió Excel)
for hoja in ['SCD-2026', 'SAF-2026']:
    if not get_columnas(hoja):
        col_f = "AA ENCARGADO/MATERIALES Y GESTION"
        col_g = "AA ENCARGADO/EQUIPAMIENTO"
        df_vacio = pd.DataFrame(columns=[col_f, col_g])
        crear_tabla_dinamica(hoja, df_vacio)

if __name__ == '__main__':
    app.run(debug=True)
