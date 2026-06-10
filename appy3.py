import streamlit as st
import pandas as pd
import sqlite3

st.set_page_config(page_title="Cuadrantes Metrovalencia", layout="wide")

st.title("📅 Cuadrante de Servicios - Metrovalencia")
st.caption("Vista calendario | Días pares azul / impares amarillo | Carga exacta por rangos")

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
        if turno and turno not in ["0", "0.0", ""]:
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
    
    def get_turno(self, agente_id, dia):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT turno FROM turnos WHERE agente_id = ? AND dia = ?", (agente_id, dia))
        result = cursor.fetchone()
        conn.close()
        return result[0] if result else ""
    
    def actualizar_turno(self, agente_id, dia, turno):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        if turno and turno not in ["0", "0.0", ""]:
            cursor.execute(
                "INSERT OR REPLACE INTO turnos (agente_id, dia, turno) VALUES (?, ?, ?)",
                (agente_id, dia, turno)
            )
        else:
            cursor.execute("DELETE FROM turnos WHERE agente_id = ? AND dia = ?", (agente_id, dia))
        conn.commit()
        conn.close()
    
    def intercambiar_turnos(self, agente1_id, agente2_id, dia):
        t1 = self.get_turno(agente1_id, dia)
        t2 = self.get_turno(agente2_id, dia)
        self.actualizar_turno(agente1_id, dia, t2)
        self.actualizar_turno(agente2_id, dia, t1)
        return True


# ============================================================
# RANGOS DE FILAS
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
            
            turnos = []
            for dia in range(31):
                col_turno = 5 + (dia * 2)
                if df.shape[1] > col_turno:
                    valor = df.iloc[fila, col_turno]
                    if pd.isna(valor):
                        turnos.append("")
                    else:
                        valor_str = str(valor).strip()
                        if valor_str in ["0", "0.0"]:
                            turnos.append("")
                        else:
                            turnos.append(valor_str)
                else:
                    turnos.append("")
            
            agentes_zona.append({
                "codigo": codigo_str,
                "nombre": nombre_str,
                "turnos": turnos
            })
        
        agentes_por_zona[zona] = agentes_zona
    
    return agentes_por_zona


# ============================================================
# VISTA CALENDARIO
# ============================================================

def mostrar_calendario(turnos):
    """Devuelve HTML del calendario para un agente"""
    
    dias_semana = ["LUN", "MAR", "MIÉ", "JUE", "VIE", "SÁB", "DOM"]
    primer_dia_semana = 4  # 1 de mayo 2026 es viernes
    
    html = '<table style="width:100%; border-collapse:collapse; text-align:center;">'
    
    # Cabecera
    html += '<tr>'
    for dia in dias_semana:
        html += f'<th style="background-color:#334155; color:white; padding:6px; border:1px solid #475569; font-size:0.8rem;">{dia}</th>'
    html += '</tr>'
    
    # Calendario (5 semanas)
    for semana in range(5):
        html += '<tr>'
        for dia_semana in range(7):
            dia_num = semana * 7 + dia_semana - primer_dia_semana + 1
            if 1 <= dia_num <= 31:
                turno = turnos[dia_num - 1] if turnos[dia_num - 1] else "—"
                bg_color = "#DBEAFE" if dia_num % 2 == 0 else "#FEF3C7"
                html += f'''
                <td style="background-color:{bg_color}; padding:6px; border:1px solid #cbd5e1;">
                    <div style="font-weight:bold; font-size:0.7rem;">{dia_num}</div>
                    <div style="font-size:0.8rem;">{turno}</div>
                </td>
                '''
            else:
                html += '<td style="background-color:#f1f5f9; border:1px solid #cbd5e1;"></td>'
        html += '</tr>'
    
    html += '</table>'
    return html


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
            with st.spinner("Cargando..."):
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
                st.success(f"✅ Cargados {total} agentes")
                for zona, agentes in agentes_por_zona.items():
                    st.write(f"   - {zona}: {len(agentes)} agentes")
                st.rerun()
            else:
                st.error("No se encontraron agentes")
    
    if st.session_state.datos_cargados:
        st.markdown("---")
        df_agentes = st.session_state.db.get_agentes()
        zonas = df_agentes["zona"].unique()
        zona_sel = st.selectbox("📍 Zona", ["TODAS"] + list(zonas))
        st.session_state.zona_sel = zona_sel
        st.metric("Total agentes", len(df_agentes))


if not st.session_state.datos_cargados:
    st.info("👈 **Carga el archivo Excel y haz clic en CARGAR AGENTES**")
    
    with st.expander("📖 Rangos configurados"):
        st.markdown("""
        | Zona | Filas |
        |------|-------|
        | **JC** | 12 - 127 |
        | **AEZ6** | 140 - 181 |
        | **AEZ7** | 196 - 245 |
        | **AEZ8** | 262 - 309 |
        
        **No se cargan:**
        - Código vacío o "0"
        - DESPLAZADO / VACANTE
        - Turno "0" o vacío
        """)

else:
    zona_sel = st.session_state.get('zona_sel', 'TODAS')
    df_agentes = st.session_state.db.get_agentes()
    
    if zona_sel != "TODAS":
        df_agentes = df_agentes[df_agentes["zona"] == zona_sel]
    
    st.markdown(f"## 📊 {zona_sel}")
    st.caption(f"{len(df_agentes)} agentes | 🔵 Azul = días pares | 🟡 Amarillo = días impares")
    
    # Grid de 2 columnas
    cols = st.columns(2)
    for idx, (_, agente) in enumerate(df_agentes.iterrows()):
        with cols[idx % 2]:
            turnos = st.session_state.db.get_turnos(agente["id"])
            with st.expander(f"📌 {agente['codigo']} - {agente['nombre']}"):
                st.markdown(mostrar_calendario(turnos), unsafe_allow_html=True)
    
    # Intercambio de turnos
    st.markdown("---")
    st.markdown("## 🔄 Intercambiar turnos")
    
    if len(df_agentes) >= 2:
        col1, col2, col3 = st.columns(3)
        
        with col1:
            nombres = [f"{row['codigo']} - {row['nombre']}" for _, row in df_agentes.iterrows()]
            ag1 = st.selectbox("Agente 1", nombres, key="ag1")
            idx1 = nombres.index(ag1)
            ag1_id = df_agentes.iloc[idx1]["id"]
        
        with col2:
            ag2 = st.selectbox("Agente 2", nombres, key="ag2")
            idx2 = nombres.index(ag2)
            ag2_id = df_agentes.iloc[idx2]["id"]
        
        with col3:
            dia = st.selectbox("Día", list(range(1, 32)), key="dia")
        
        turno1 = st.session_state.db.get_turno(ag1_id, dia)
        turno2 = st.session_state.db.get_turno(ag2_id, dia)
        
        st.info(f"📌 **Actual:** {ag1} → `{turno1 if turno1 else '—'}` | {ag2} → `{turno2 if turno2 else '—'}`")
        
        if st.button("🔄 Intercambiar turnos", type="primary", use_container_width=True):
            st.session_state.db.intercambiar_turnos(ag1_id, ag2_id, dia)
            st.success("✅ Turnos intercambiados")
            st.rerun()
    else:
        st.warning("Se necesitan al menos 2 agentes")
    
    if st.button("🔄 Recargar archivo", use_container_width=True):
        st.session_state.datos_cargados = False
        st.rerun()
