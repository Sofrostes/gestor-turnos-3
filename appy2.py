import streamlit as st
import pandas as pd
import sqlite3

st.set_page_config(page_title="Cuadrantes Metrovalencia", layout="wide")

st.title("📅 Cuadrante de Servicios - Metrovalencia")
st.caption("Carga EXACTA por rangos | Lectura directa de celdas")

# ============================================================
# GESTOR BD
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
            zona TEXT
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
    
    def guardar_agente(self, codigo, nombre, zona):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO agentes (codigo, nombre, zona) VALUES (?, ?, ?)",
            (codigo, nombre, zona)
        )
        agente_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return agente_id
    
    def guardar_turno(self, agente_id, dia, turno):
        if turno:
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
# RANGOS DE FILAS (según lo indicado)
# ============================================================

RANGOS_POR_ZONA = {
    "JC": (12, 127),
    "AEZ6": (140, 181),
    "AEZ7": (196, 245),
    "AEZ8": (262, 309)
}


def cargar_por_rangos(archivo_path):
    """Carga agentes según los rangos de filas predefinidos"""
    
    df = pd.read_excel(archivo_path, sheet_name="MAYO 2026", header=None)
    
    agentes_por_zona = {}
    
    for zona, (fila_inicio, fila_fin) in RANGOS_POR_ZONA.items():
        agentes_zona = []
        
        for fila in range(fila_inicio - 1, min(fila_fin, len(df))):
            # Columna D (índice 3) = código
            # Columna E (índice 4) = nombre
            codigo = df.iloc[fila, 3] if df.shape[1] > 3 else None
            nombre = df.iloc[fila, 4] if df.shape[1] > 4 else None
            
            if pd.isna(codigo) or pd.isna(nombre):
                continue
            
            codigo_str = str(codigo).strip()
            nombre_str = str(nombre).strip()
            
            if codigo_str == "0" or codigo_str == "":
                continue
            if "DESPLAZADO" in nombre_str.upper() or "VACANTE" in nombre_str.upper():
                continue
            
            # Leer turnos (columna F=5, H=7, J=9...)
            turnos = []
            for dia in range(31):
                col_turno = 5 + (dia * 2)  # F=5, H=7, J=9...
                if df.shape[1] > col_turno:
                    valor = df.iloc[fila, col_turno]
                    turnos.append(str(valor).strip() if not pd.isna(valor) else "")
                else:
                    turnos.append("")
            
            agentes_zona.append({
                "codigo": codigo_str,
                "nombre": nombre_str,
                "turnos": turnos,
                "fila_excel": fila + 1
            })
        
        agentes_por_zona[zona] = agentes_zona
    
    return agentes_por_zona


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
        if st.button("📥 CARGAR AGENTES", type="primary", use_container_width=True):
            with st.spinner("Cargando agentes por rangos..."):
                with open("temp.xlsx", "wb") as f:
                    f.write(archivo.getbuffer())
                
                agentes_por_zona = cargar_por_rangos("temp.xlsx")
            
            total = sum(len(ag) for ag in agentes_por_zona.values())
            
            if total > 0:
                st.session_state.db.limpiar()
                
                for zona, agentes in agentes_por_zona.items():
                    for ag in agentes:
                        agente_id = st.session_state.db.guardar_agente(ag["codigo"], ag["nombre"], zona)
                        for dia, turno in enumerate(ag["turnos"], 1):
                            if turno:
                                st.session_state.db.guardar_turno(agente_id, dia, turno)
                
                st.session_state.datos_cargados = True
                st.success(f"✅ Cargados {total} agentes:")
                for zona, agentes in agentes_por_zona.items():
                    st.write(f"   - {zona}: {len(agentes)} agentes")
                st.rerun()
            else:
                st.error("No se encontraron agentes. Verifica los rangos.")
    
    if st.session_state.datos_cargados:
        st.markdown("---")
        df_agentes = st.session_state.db.get_agentes()
        zonas = df_agentes["zona"].unique()
        zona_sel = st.selectbox("📍 Filtrar por zona", ["TODAS"] + list(zonas))
        st.session_state.zona_sel = zona_sel
        st.metric("Total agentes", len(df_agentes))
        
        st.markdown("---")
        st.caption("**Rangos configurados:**")
        for zona, (ini, fin) in RANGOS_POR_ZONA.items():
            st.caption(f"- {zona}: filas {ini}-{fin}")


if not st.session_state.datos_cargados:
    st.info("👈 **Carga el archivo Excel y haz clic en CARGAR AGENTES**")
    
    with st.expander("📖 Rangos configurados"):
        st.markdown("""
        | Zona | Descripción | Filas |
        |------|-------------|-------|
        | JC | Jefes de Circulación | 12 - 127 |
        | AEZ6 | Agentes Zona 6 | 140 - 181 |
        | AEZ7 | Agentes Zona 7 | 196 - 245 |
        | AEZ8 | Agentes Zona 8 | 262 - 309 |
        
        **La app leerá:**
        - **Columna D**: Código del agente
        - **Columna E**: Nombre del agente
        - **Columna F, H, J...**: Turnos (día 1, 2, 3...)
        """)

else:
    zona_sel = st.session_state.get('zona_sel', 'TODAS')
    df_agentes = st.session_state.db.get_agentes()
    
    if zona_sel != "TODAS":
        df_agentes = df_agentes[df_agentes["zona"] == zona_sel]
    
    st.markdown(f"## 📊 Agentes - {zona_sel}")
    st.caption(f"**{len(df_agentes)} agentes** | Clic en cada uno para ver sus turnos")
    
    # Mostrar agentes
    for _, agente in df_agentes.iterrows():
        with st.expander(f"📌 {agente['codigo']} - {agente['nombre']} ({agente['zona']})"):
            turnos = st.session_state.db.get_turnos(agente["id"])
            # Mostrar en filas de 7 días
            for i in range(0, 31, 7):
                fin = min(i+7, 31)
                dias_str = ", ".join([f"D{d+1}:{turnos[d] if turnos[d] else '—'}" for d in range(i, fin)])
                st.write(dias_str)
    
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
        
        turnos1 = st.session_state.db.get_turnos(ag1_id)
        turnos2 = st.session_state.db.get_turnos(ag2_id)
        
        st.info(f"📌 **Turno actual:** {ag1} → `{turnos1[dia-1] if turnos1[dia-1] else '—'}` | {ag2} → `{turnos2[dia-1] if turnos2[dia-1] else '—'}`")
        
        if st.button("🔄 Intercambiar turnos", type="primary", use_container_width=True):
            st.session_state.db.intercambiar_turnos(ag1_id, ag2_id, dia)
            st.success("✅ Turnos intercambiados")
            st.rerun()
    else:
        st.warning("Se necesitan al menos 2 agentes para intercambiar turnos")
    
    # Botón para reiniciar
    if st.button("🔄 Cambiar archivo / recargar", use_container_width=True):
        st.session_state.datos_cargados = False
        st.rerun()
