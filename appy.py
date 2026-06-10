import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import tempfile
import sqlite3

st.set_page_config(page_title="Cuadrantes Metrovalencia", layout="wide")

st.title("📅 Cuadrante de Servicios - Metrovalencia")
st.caption("Carga DIRECTA desde Excel | Sin interpretaciones | Sin inventar datos")

# ============================================================
# CLASE GESTOR DB (simple y directa)
# ============================================================

class GestorDB:
    def __init__(self, db_path="cuadrantes.db"):
        self.db_path = db_path
        self._init_db()
    
    def _init_db(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS agentes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT NOT NULL,
            nombre TEXT NOT NULL,
            zona TEXT,
            fila_excel INTEGER
        )''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS turnos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            agente_id INTEGER NOT NULL,
            dia INTEGER NOT NULL,
            turno TEXT,
            UNIQUE(agente_id, dia)
        )''')
        conn.commit()
        conn.close()
    
    def limpiar(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM turnos")
        cursor.execute("DELETE FROM agentes")
        conn.commit()
        conn.close()
    
    def guardar_agente(self, codigo, nombre, zona, fila_excel):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO agentes (codigo, nombre, zona, fila_excel) VALUES (?, ?, ?, ?)",
            (codigo, nombre, zona, fila_excel)
        )
        agente_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return agente_id
    
    def guardar_turno(self, agente_id, dia, turno):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute(
            "INSERT OR REPLACE INTO turnos (agente_id, dia, turno) VALUES (?, ?, ?)",
            (agente_id, dia, turno)
        )
        conn.commit()
        conn.close()
    
    def get_agentes(self):
        conn = sqlite3.connect(self.db_path)
        df = pd.read_sql("SELECT id, codigo, nombre, zona FROM agentes ORDER BY zona, nombre", conn)
        conn.close()
        return df
    
    def get_turnos(self, agente_id):
        conn = sqlite3.connect(self.db_path)
        df = pd.read_sql("SELECT dia, turno FROM turnos WHERE agente_id = ? ORDER BY dia", conn, params=(agente_id,))
        conn.close()
        turnos = [""] * 31
        for _, row in df.iterrows():
            turnos[row["dia"] - 1] = row["turno"]
        return turnos
    
    def intercambiar_turnos(self, agente1_id, agente2_id, dia):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT turno FROM turnos WHERE agente_id = ? AND dia = ?", (agente1_id, dia))
        t1 = cursor.fetchone()
        t1 = t1[0] if t1 else ""
        cursor.execute("SELECT turno FROM turnos WHERE agente_id = ? AND dia = ?", (agente2_id, dia))
        t2 = cursor.fetchone()
        t2 = t2[0] if t2 else ""
        cursor.execute("INSERT OR REPLACE INTO turnos (agente_id, dia, turno) VALUES (?, ?, ?)", (agente1_id, dia, t2))
        cursor.execute("INSERT OR REPLACE INTO turnos (agente_id, dia, turno) VALUES (?, ?, ?)", (agente2_id, dia, t1))
        conn.commit()
        conn.close()
        return True


# ============================================================
# CARGA DIRECTA DESDE EXCEL - CELDA POR CELDA
# ============================================================

def cargar_desde_excel(archivo_path):
    """Carga los datos DIRECTAMENTE desde el Excel, celda por celda"""
    
    wb = load_workbook(archivo_path, data_only=True)
    sheet = wb["MAYO 2026"]
    
    agentes = []
    
    # RECORREMOS CADA FILA DESDE LA 12 HASTA EL FINAL
    for fila in range(12, 300):
        # Leer valores DIRECTOS de las celdas
        zona = sheet.cell(row=fila, column=1).value      # Columna A
        codigo = sheet.cell(row=fila, column=3).value    # Columna C
        nombre = sheet.cell(row=fila, column=5).value    # Columna E
        
        # Si no hay código, dejamos de buscar
        if not codigo:
            continue
        
        # Saltar si es desplazado o vacante
        nombre_str = str(nombre) if nombre else ""
        if "DESPLAZADO" in nombre_str.upper() or "VACANTE" in nombre_str.upper():
            continue
        
        # Leer turnos - DIRECTOS de cada celda
        turnos = []
        for dia in range(31):
            col_turno = 6 + (dia * 2)  # F=6, H=8, J=10, L=12...
            valor = sheet.cell(row=fila, column=col_turno).value
            turnos.append(str(valor).strip() if valor else "")
        
        agentes.append({
            "fila": fila,
            "zona": str(zona).strip() if zona else "",
            "codigo": str(codigo).strip(),
            "nombre": nombre_str,
            "turnos": turnos
        })
    
    return agentes


# ============================================================
# INTERFAZ
# ============================================================

if 'db' not in st.session_state:
    st.session_state.db = GestorDB()
if 'datos_cargados' not in st.session_state:
    st.session_state.datos_cargados = False


with st.sidebar:
    st.header("📁 Cargar Excel")
    
    archivo = st.file_uploader("Selecciona AÑO 2026 ESTACIONES .xlsx", type=["xlsx"])
    
    if archivo:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(archivo.getvalue())
            archivo_path = tmp.name
        
        if st.button("📥 CARGAR DATOS DIRECTAMENTE", type="primary", use_container_width=True):
            with st.spinner("Leyendo celda por celda..."):
                agentes = cargar_desde_excel(archivo_path)
            
            if agentes:
                # Limpiar BD anterior
                st.session_state.db.limpiar()
                
                # Guardar cada agente y sus turnos
                for ag in agentes:
                    agente_id = st.session_state.db.guardar_agente(
                        ag["codigo"], ag["nombre"], ag["zona"], ag["fila"]
                    )
                    for dia, turno in enumerate(ag["turnos"], 1):
                        if turno:
                            st.session_state.db.guardar_turno(agente_id, dia, turno)
                
                st.session_state.datos_cargados = True
                st.success(f"✅ Cargados {len(agentes)} agentes")
                st.rerun()
            else:
                st.error("No se encontraron agentes. Verifica que los datos empiecen en fila 12")
        
        st.markdown("---")
        st.caption("**Estructura esperada:**")
        st.caption("- Fila 12: primer agente")
        st.caption("- Col A: Zona (JC, AV, AA...)")
        st.caption("- Col C: Código del agente")
        st.caption("- Col E: Nombre del agente")
        st.caption("- Col F, H, J...: Turnos (día 1, 2, 3...)")
    
    if st.session_state.datos_cargados:
        st.markdown("---")
        df_agentes = st.session_state.db.get_agentes()
        zonas = df_agentes["zona"].unique()
        zona_sel = st.selectbox("📍 Filtrar por zona", ["TODAS"] + list(zonas))
        st.session_state.zona_sel = zona_sel
        st.metric("Total agentes", len(df_agentes))


# CONTENIDO PRINCIPAL
if not st.session_state.datos_cargados:
    st.info("👈 **Carga el archivo Excel y haz clic en CARGAR DATOS DIRECTAMENTE**")
    
    with st.expander("📖 ¿Cómo funciona?"):
        st.markdown("""
        ### La app lee DIRECTAMENTE:
        
        | Columna | Contenido |
        |---------|-----------|
        | A | Zona (JC, AV, AA, FO, MS, AE) |
        | C | Código del agente |
        | E | Nombre del agente |
        | F, H, J, L, N... | Turnos (día 1, día 2, día 3...) |
        
        ### No interpreta, no adivina, no inventa:
        - ✅ Lee el valor EXACTO de cada celda
        - ✅ Respeta espacios y mayúsculas
        - ✅ Si la celda está vacía, guarda vacío
        - ✅ Si la celda tiene "2F", guarda "2F"
        
        ### ¿No carga agentes?
        - Verifica que los agentes empiecen en **fila 12**
        - Verifica que la **columna C** tenga códigos
        - Verifica que la **columna E** tenga nombres
        """)

else:
    zona_sel = st.session_state.get('zona_sel', 'TODAS')
    df_agentes = st.session_state.db.get_agentes()
    
    if zona_sel != "TODAS":
        df_agentes = df_agentes[df_agentes["zona"] == zona_sel]
    
    st.markdown(f"## 📊 Agentes - {zona_sel if zona_sel != 'TODAS' else 'TODOS'}")
    st.caption(f"**{len(df_agentes)} agentes** | Clic en cada uno para ver sus turnos")
    
    # Mostrar agentes en columnas
    cols = st.columns(3)
    for idx, (_, agente) in enumerate(df_agentes.iterrows()):
        with cols[idx % 3]:
            with st.expander(f"📌 {agente['codigo']} - {agente['nombre']} ({agente['zona']})"):
                turnos = st.session_state.db.get_turnos(agente["id"])
                for i in range(0, 31, 7):
                    st.write(f"Días {i+1}-{min(i+7,31)}: {', '.join(turnos[i:i+7])}")
    
    # Intercambio de turnos
    st.markdown("---")
    st.markdown("## 🔄 Intercambiar turnos")
    
    if len(df_agentes) >= 2:
        col1, col2, col3 = st.columns(3)
        
        with col1:
            nombres = [f"{row['codigo']} - {row['nombre']}" for _, row in df_agentes.iterrows()]
            ag1 = st.selectbox("Agente 1", nombres, key="ag1")
            ag1_id = df_agentes.iloc[nombres.index(ag1)]["id"]
        
        with col2:
            ag2 = st.selectbox("Agente 2", nombres, key="ag2")
            ag2_id = df_agentes.iloc[nombres.index(ag2)]["id"]
        
        with col3:
            dia = st.selectbox("Día", list(range(1, 32)), key="dia")
        
        # Mostrar turnos actuales
        turnos1 = st.session_state.db.get_turnos(ag1_id)
        turnos2 = st.session_state.db.get_turnos(ag2_id)
        
        st.info(f"📌 **Actual:** {ag1} → `{turnos1[dia-1] if turnos1[dia-1] else '—'}` | {ag2} → `{turnos2[dia-1] if turnos2[dia-1] else '—'}`")
        
        if st.button("🔄 Intercambiar turnos", type="primary", use_container_width=True):
            st.session_state.db.intercambiar_turnos(ag1_id, ag2_id, dia)
            st.success("✅ Turnos intercambiados")
            st.rerun()
    else:
        st.warning("Se necesitan al menos 2 agentes para intercambiar turnos")
