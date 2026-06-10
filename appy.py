# app.py - Aplicación principal con API para sincronización
import streamlit as st
import pandas as pd
import sqlite3
from openpyxl import load_workbook
from flask import Flask, request, jsonify
import threading
import json
import os

st.set_page_config(page_title="Gestión de Cuadrantes", layout="wide")

st.title("📅 Gestión de Cuadrantes - Metrovalencia")
st.caption("Sincronización bidireccional con Excel mediante API")

# ============================================================
# SERVIDOR API (para comunicación con Excel)
# ============================================================

api_app = Flask(__name__)
api_running = False

@api_app.route('/api/exportar', methods=['POST'])
def exportar_desde_excel():
    """Recibe datos desde Excel y los guarda en la BD"""
    try:
        data = request.json
        agentes = data.get('agentes', [])
        
        conn = sqlite3.connect('cuadrantes.db')
        cursor = conn.cursor()
        
        # Limpiar datos existentes
        cursor.execute("DELETE FROM turnos")
        cursor.execute("DELETE FROM agentes")
        cursor.execute("DELETE FROM zonas")
        
        for agente in agentes:
            # Insertar zona
            cursor.execute("INSERT OR IGNORE INTO zonas (nombre) VALUES (?)", (agente['zona'],))
            cursor.execute("SELECT id FROM zonas WHERE nombre = ?", (agente['zona'],))
            zona_id = cursor.fetchone()[0]
            
            # Insertar agente
            cursor.execute(
                "INSERT INTO agentes (codigo, nombre, zona_id) VALUES (?, ?, ?)",
                (agente['codigo'], agente['nombre'], zona_id)
            )
            agente_id = cursor.lastrowid
            
            # Insertar turnos
            for dia, turno in enumerate(agente['turnos'], 1):
                if turno:
                    cursor.execute(
                        "INSERT OR REPLACE INTO turnos (agente_id, dia, turno) VALUES (?, ?, ?)",
                        (agente_id, dia, turno)
                    )
        
        conn.commit()
        conn.close()
        
        return jsonify({"status": "success", "message": f"Importados {len(agentes)} agentes"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

@api_app.route('/api/importar', methods=['GET'])
def importar_a_excel():
    """Devuelve todos los turnos actuales para actualizar Excel"""
    try:
        conn = sqlite3.connect('cuadrantes.db')
        cursor = conn.cursor()
        
        # Obtener todos los turnos
        query = """
            SELECT a.codigo, a.nombre, z.nombre as zona, t.dia, t.turno
            FROM agentes a
            JOIN zonas z ON a.zona_id = z.id
            LEFT JOIN turnos t ON a.id = t.agente_id
            ORDER BY a.id, t.dia
        """
        cursor.execute(query)
        rows = cursor.fetchall()
        
        # Agrupar por agente
        agentes_dict = {}
        for row in rows:
            codigo, nombre, zona, dia, turno = row
            key = f"{codigo}|{nombre}"
            if key not in agentes_dict:
                agentes_dict[key] = {
                    "codigo": codigo,
                    "nombre": nombre,
                    "zona": zona,
                    "turnos": [""] * 31
                }
            if dia and 1 <= dia <= 31:
                agentes_dict[key]["turnos"][dia-1] = turno if turno else ""
        
        conn.close()
        
        return jsonify({
            "status": "success",
            "agentes": list(agentes_dict.values())
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

def iniciar_api():
    """Inicia el servidor API en un hilo separado"""
    global api_running
    if not api_running:
        api_running = True
        threading.Thread(target=lambda: api_app.run(host='127.0.0.1', port=5000, debug=False, use_reloader=False), daemon=True).start()


# ============================================================
# GESTOR DE BASE DE DATOS
# ============================================================

class GestorDB:
    def __init__(self, db_path="cuadrantes.db"):
        self.db_path = db_path
        self.init_db()
    
    def init_db(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS zonas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT UNIQUE NOT NULL
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS agentes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo TEXT NOT NULL,
                nombre TEXT NOT NULL,
                zona_id INTEGER,
                FOREIGN KEY (zona_id) REFERENCES zonas(id)
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS turnos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                agente_id INTEGER NOT NULL,
                dia INTEGER NOT NULL,
                turno TEXT,
                UNIQUE(agente_id, dia),
                FOREIGN KEY (agente_id) REFERENCES agentes(id)
            )
        ''')
        
        conn.commit()
        conn.close()
    
    def get_agentes_por_zona(self):
        conn = sqlite3.connect(self.db_path)
        query = """
            SELECT a.id, a.codigo, a.nombre, z.nombre as zona
            FROM agentes a
            JOIN zonas z ON a.zona_id = z.id
            ORDER BY z.nombre, a.nombre
        """
        df = pd.read_sql(query, conn)
        conn.close()
        return df
    
    def get_turnos_agente(self, agente_id):
        conn = sqlite3.connect(self.db_path)
        query = "SELECT dia, turno FROM turnos WHERE agente_id = ? ORDER BY dia"
        df = pd.read_sql(query, conn, params=(agente_id,))
        conn.close()
        
        turnos = [""] * 31
        for _, row in df.iterrows():
            turnos[row["dia"] - 1] = row["turno"]
        return turnos
    
    def intercambiar_turnos(self, agente1_id, agente2_id, dia):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute("SELECT turno FROM turnos WHERE agente_id = ? AND dia = ?", (agente1_id, dia))
        turno1 = cursor.fetchone()
        turno1 = turno1[0] if turno1 else ""
        
        cursor.execute("SELECT turno FROM turnos WHERE agente_id = ? AND dia = ?", (agente2_id, dia))
        turno2 = cursor.fetchone()
        turno2 = turno2[0] if turno2 else ""
        
        cursor.execute("INSERT OR REPLACE INTO turnos (agente_id, dia, turno) VALUES (?, ?, ?)", (agente1_id, dia, turno2))
        cursor.execute("INSERT OR REPLACE INTO turnos (agente_id, dia, turno) VALUES (?, ?, ?)", (agente2_id, dia, turno1))
        
        conn.commit()
        conn.close()
        return True
    
    def get_estadisticas(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM agentes")
        total_agentes = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM turnos WHERE turno != ''")
        total_turnos = cursor.fetchone()[0]
        cursor.execute("SELECT z.nombre, COUNT(*) FROM zonas z JOIN agentes a ON z.id = a.zona_id GROUP BY z.nombre")
        por_zona = cursor.fetchall()
        conn.close()
        return total_agentes, total_turnos, por_zona


# ============================================================
# INTERFAZ STREAMLIT
# ============================================================

# Iniciar API en segundo plano
iniciar_api()

if 'db' not in st.session_state:
    st.session_state.db = GestorDB()

# Sidebar
with st.sidebar:
    st.header("🔌 Sincronización")
    
    # Mostrar estado de la API
    st.success("🟢 API activa en http://127.0.0.1:5000")
    
    st.markdown("---")
    st.header("📋 Configuración para Excel")
    
    st.markdown("""
    ### Instrucciones para Excel:
    
    1. **Habilitar macros** en Excel
    2. **Insertar el código VBA** (abajo)
    3. **Crear botones** para sincronizar
    """)
    
    with st.expander("📄 Código VBA para Excel"):
        st.code("""
' =====================================================
' MACROS PARA SINCRONIZAR CON LA APP
' =====================================================

Sub EnviarDatosAApp()
    ' Envía todos los turnos desde Excel a la app
    Dim http As Object
    Dim data As String
    Dim agentes As String
    Dim fila As Integer
    Dim col As Integer
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Recorrer agentes (desde fila 12)
    agentes = "["
    fila = 12
    
    Do While Cells(fila, 3).Value <> ""
        codigo = Cells(fila, 3).Value
        nombre = Cells(fila, 5).Value
        zona = Cells(fila, 1).Value
        
        If codigo <> "" And nombre <> "" Then
            ' Recoger turnos (columnas F, H, J...)
            turnos = "["
            For dia = 1 To 31
                col_turno = 6 + (dia - 1) * 2
                turno = Cells(fila, col_turno).Value
                If IsNull(turno) Then turno = ""
                turnos = turnos & """" & turno & """"
                If dia < 31 Then turnos = turnos & ","
            Next
            turnos = turnos & "]"
            
            agente_json = "{""codigo"":""" & codigo & """,""nombre"":""" & nombre & """,""zona"":""" & zona & """,""turnos"":" & turnos & "}"
            agentes = agentes & agente_json & ","
        End If
        fila = fila + 1
    Loop
    
    If Right(agentes, 1) = "," Then agentes = Left(agentes, Len(agentes) - 1)
    agentes = agentes & "]"
    
    data = "{""agentes"":" & agentes & "}"
    
    ' Enviar a la app
    http.Open "POST", "http://127.0.0.1:5000/api/exportar", False
    http.setRequestHeader "Content-Type", "application/json"
    http.send data
    
    MsgBox "Datos enviados a la app", vbInformation
End Sub

Sub RecibirDatosDeApp()
    ' Recibe los turnos modificados desde la app y actualiza Excel
    Dim http As Object
    Dim respuesta As String
    Dim agentes As Object
    Dim agente As Object
    Dim fila As Integer
    Dim dia As Integer
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    http.Open "GET", "http://127.0.0.1:5000/api/importar", False
    http.send
    
    respuesta = http.responseText
    
    ' Parsear JSON (requiere referencia a ScriptControl)
    Dim sc As Object
    Set sc = CreateObject("ScriptControl")
    sc.Language = "JScript"
    
    sc.ExecuteStatement "var datos = " & respuesta
    agentes = sc.Eval("datos.agentes")
    
    ' Actualizar Excel
    For Each agente In agentes
        codigo = agente.codigo
        nombre = agente.nombre
        turnos = agente.turnos
        
        ' Buscar fila del agente
        fila = 12
        Do While Cells(fila, 3).Value <> ""
            If CStr(Cells(fila, 3).Value) = CStr(codigo) And CStr(Cells(fila, 5).Value) = CStr(nombre) Then
                ' Actualizar turnos
                For dia = 1 To 31
                    col_turno = 6 + (dia - 1) * 2
                    Cells(fila, col_turno).Value = turnos(dia - 1)
                Next
                Exit Do
            End If
            fila = fila + 1
        Loop
    Next
    
    MsgBox "Datos actualizados desde la app", vbInformation
End Sub

Sub SincronizarCompleta()
    ' Sincronización bidireccional
    EnviarDatosAApp
    RecibirDatosDeApp
    MsgBox "Sincronización completa", vbInformation
End Sub
        """, language="vb")
    
    st.markdown("---")
    
    # Estadísticas
    total_agentes, total_turnos, por_zona = st.session_state.db.get_estadisticas()
    st.metric("Total agentes en BD", total_agentes)
    st.metric("Turnos asignados", total_turnos)


# Contenido principal
if total_agentes == 0:
    st.info("👈 Usa el botón 'Enviar a la app' desde Excel para cargar los datos")
    
    with st.expander("📖 Cómo configurar Excel"):
        st.markdown("""
        ### Configuración en Excel:
        
        1. **Abrir el Editor de VBA**: `Alt + F11`
        2. **Insertar un módulo**: Insertar → Módulo
        3. **Pegar el código VBA** (mostrado arriba)
        4. **Crear botones** en Excel:
           - Insertar → Formulario → Botón
           - Asignar macros:
             - "Enviar a App" → `EnviarDatosAApp`
             - "Recibir de App" → `RecibirDatosDeApp`
             - "Sincronizar" → `SincronizarCompleta`
        
        ### Uso:
        
        1. **Enviar** - Sube los datos de Excel a la app
        2. **Editar** - Haz cambios en la app (intercambiar turnos)
        3. **Recibir** - Baja los cambios desde la app a Excel
        4. **Sincronizar** - Hace ambos pasos seguidos
        """)
else:
    # Selector de zona
    df_agentes = st.session_state.db.get_agentes_por_zona()
    zonas = df_agentes["zona"].unique()
    zona_sel = st.selectbox("📍 Filtrar por zona", ["TODAS"] + list(zonas))
    
    if zona_sel != "TODAS":
        df_agentes = df_agentes[df_agentes["zona"] == zona_sel]
    
    st.markdown(f"## 📊 Agentes - {zona_sel if zona_sel != 'TODAS' else 'TODAS LAS ZONAS'}")
    
    # Mostrar agentes con turnos
    for _, agente in df_agentes.iterrows():
        with st.expander(f"📌 {agente['codigo']} - {agente['nombre']} ({agente['zona']})"):
            turnos = st.session_state.db.get_turnos_agente(agente["id"])
            
            # Mostrar turnos en una cuadrícula
            cols = st.columns(7)
            for i, turno in enumerate(turnos):
                col_idx = i % 7
                with cols[col_idx]:
                    st.metric(f"Día {i+1}", turno if turno else "—")
    
    # Interfaz de intercambio rápida
    st.markdown("---")
    st.markdown("## 🔄 Intercambio rápido de turnos")
    
    if len(df_agentes) >= 2:
        col1, col2, col3 = st.columns(3)
        
        with col1:
            nombres = [f"{row['codigo']} - {row['nombre']}" for _, row in df_agentes.iterrows()]
            ag1 = st.selectbox("Agente 1", nombres, key="ag1")
            idx1 = df_agentes.iloc[nombres.index(ag1)]
        
        with col2:
            ag2 = st.selectbox("Agente 2", nombres, key="ag2")
            idx2 = df_agentes.iloc[nombres.index(ag2)]
        
        with col3:
            dia = st.selectbox("Día", list(range(1, 32)), key="dia")
        
        if st.button("🔄 Intercambiar turnos", type="primary", use_container_width=True):
            st.session_state.db.intercambiar_turnos(idx1["id"], idx2["id"], dia)
            st.success("✅ Turnos intercambiados. Usa 'Recibir de App' en Excel para actualizar.")
            st.rerun()
    
    st.info("💡 **Recuerda:** Después de hacer cambios, usa la macro **'Recibir de App'** en Excel para actualizar el archivo")
