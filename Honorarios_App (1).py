import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px

# ============================================================
# 1. CONFIGURACIÓN DE LA PÁGINA
# ============================================================
st.set_page_config(
    page_title="Control de Honorarios",
    layout="wide",
    initial_sidebar_state="expanded"
)
st.title("📊 Panel de Control de Facturación y Honorarios")

# ============================================================
# 2. CONSTANTES
# ============================================================
ARCHIVO_LOCAL = "Honosrario_NM.xlsx"   # ← corregido el nombre (guion bajo)

ESCALAS_MONOTRIBUTO = {
    'A':  6_450_000,
    'B':  9_450_000,
    'C': 13_250_000,
    'D': 16_450_000,
    'E': 19_350_000,
    'F': 24_250_000,
    'G': 29_000_000,
    'H': 44_000_000,
    'I': 49_115_000,
    'J': 56_400_000,
    'K': 68_000_000,
}

# ============================================================
# 3. SESSION STATE (valores iniciales solo si no existen)
# ============================================================
if 'cat_agustin' not in st.session_state:
    st.session_state['cat_agustin'] = 'D'
if 'cat_laura' not in st.session_state:
    st.session_state['cat_laura'] = 'B'

# ============================================================
# 4. FUNCIONES UTILITARIAS
# ============================================================
def formato_abreviado(valor: float) -> str:
    """Formatea un número en millones o miles para mostrar en métricas."""
    if valor >= 1_000_000:
        return f"$ {valor / 1_000_000:.2f} M"
    elif valor >= 1_000:
        return f"$ {valor / 1_000:.1f} k"
    return f"$ {valor:.2f}"


def colorear_clientes(row):
    """Aplica estilos de color a la tabla de clientes según estado y alertas."""
    if row['Estado'] == 'Inactivo':
        return ['background-color: #fee2e2; color: #991b1b; opacity: 0.7;'] * len(row)
    estilos = [''] * len(row)
    if row['Alerta_Revisión'] == 'VENCIDO':
        idx = row.index.get_loc('Alerta_Revisión')
        estilos[idx] = 'background-color: #fef08a; color: #854d0e; font-weight: bold;'
    return estilos


# ============================================================
# 5. CARGA DE DATOS CON CACHÉ INTELIGENTE
#    El caché se invalida automáticamente cuando cambia el archivo
# ============================================================
@st.cache_data(show_spinner="📂 Leyendo y procesando datos...")
def cargar_y_procesar_datos(ruta_archivo: str, _mod_time: float) -> dict:
    """
    Lee el Excel y devuelve un diccionario con todos los DataFrames procesados.
    El parámetro _mod_time (ignorado por pandas, prefijado con _) hace que el
    caché se invalide automáticamente cada vez que el archivo cambia en disco.
    """
    if not os.path.exists(ruta_archivo):
        return {}

    # --- Lectura de hojas ---
    try:
        df_facturas = pd.read_excel(ruta_archivo, sheet_name="Facturas")
        df_clientes = pd.read_excel(ruta_archivo, sheet_name="Clientes")
        df_indices  = pd.read_excel(ruta_archivo, sheet_name="Indices")
    except Exception as e:
        st.error(f"Error al leer el Excel: {e}")
        return {}

    # --- Limpieza de nombres de columna ---
    df_facturas.columns = df_facturas.columns.str.strip()
    df_clientes.columns = df_clientes.columns.str.strip()
    df_indices.columns  = df_indices.columns.str.strip()

    # --- Índices: quitar columnas sin nombre y normalizar ---
    df_indices = df_indices.loc[:, ~df_indices.columns.str.contains('^Unnamed')]
    df_indices = df_indices[['MES', 'IPC  IPIM']].copy()
    df_indices['IPC  IPIM'] = pd.to_numeric(
        df_indices['IPC  IPIM'].astype(str).str.replace(',', '.'), errors='coerce'
    )

    # --- Facturas: columnas opcionales con fallback ---
    df_facturas['Emisor']  = df_facturas['Emisor'].astype(str).str.strip()  if 'Emisor'  in df_facturas.columns else 'Agustín'
    df_facturas['Negocio'] = df_facturas['Negocio'].astype(str).str.strip() if 'Negocio' in df_facturas.columns else 'Estudio'
    df_facturas['Facturacion $'] = pd.to_numeric(
        df_facturas['Facturacion $'].astype(str).str.replace(',', '.'), errors='coerce'
    )

    # --- Clientes: precio numérico ---
    if 'precio' in df_clientes.columns:
        df_clientes['precio'] = pd.to_numeric(
            df_clientes['precio'].astype(str).str.replace(',', '.'), errors='coerce'
        ).fillna(0.0)

    # --- Fechas ---
    df_facturas['Fecha_dt']  = pd.to_datetime(df_facturas['Fecha'], errors='coerce')
    df_indices['MES_dt']     = pd.to_datetime(df_indices['MES'],    errors='coerce')
    df_facturas['Mes_Indice'] = df_facturas['Fecha_dt'].dt.to_period('M').dt.to_timestamp()

    # --- Índice más reciente ---
    df_indices_ord = df_indices.sort_values('MES_dt')
    ultimo_mes     = df_indices_ord.iloc[-1]['MES_dt']
    ultimo_indice  = df_indices_ord.iloc[-1]['IPC  IPIM']

    # --- Merge facturas + índices para deflactar ---
    df_res = pd.merge(
        df_facturas,
        df_indices[['MES_dt', 'IPC  IPIM']],
        left_on='Mes_Indice', right_on='MES_dt',
        how='left'
    )
    df_res['Coeficiente'] = (ultimo_indice / df_res['IPC  IPIM']).fillna(1.0)
    df_res['Facturacion $ Actualizada'] = df_res['Facturacion $'] * df_res['Coeficiente']

    # --- Merge con clientes para enriquecer ---
    df_final = pd.merge(df_res, df_clientes, on='Nro. Doc. Receptor', how='left')

    # --- Historial visual (columnas seleccionadas) ---
    cols_hist = [
        'Fecha_dt', 'Mes_Indice', 'Emisor', 'Negocio', 'Tipo',
        'Punto de Venta', 'Número Desde', 'Nro. Doc. Receptor',
        'Denominación Receptor', 'Facturacion $', 'Facturacion $ Actualizada'
    ]
    df_historial_visual = df_final[[c for c in cols_hist if c in df_final.columns]].copy()

    # --- Procesamiento de clientes: métricas vectorizadas ---
    df_cp = df_clientes.copy()
    df_cp['Actualizacion_dt']       = pd.to_datetime(df_cp['Actualizacion'], errors='coerce')
    df_cp['periodos']               = pd.to_numeric(df_cp['periodos'], errors='coerce').fillna(0).astype(int)
    df_cp['Mes_Actualizacion_dt']   = df_cp['Actualizacion_dt'].dt.to_period('M').dt.to_timestamp()

    df_cp = pd.merge(
        df_cp,
        df_indices[['MES_dt', 'IPC  IPIM']],
        left_on='Mes_Actualizacion_dt', right_on='MES_dt',
        how='left'
    ).rename(columns={'IPC  IPIM': 'IPC_Ult_Actualizacion'}).drop(columns=['MES_dt', 'Mes_Actualizacion_dt'])

    hoy = datetime.now()

    # Meses de antigüedad desde la última actualización (vectorizado)
    df_cp['Meses Desactualizado'] = (
        (hoy.year  - df_cp['Actualizacion_dt'].dt.year)  * 12 +
        (hoy.month - df_cp['Actualizacion_dt'].dt.month)
    ).fillna(0).astype(int)

    # Honorario sugerido por IPC (vectorizado)
    coef_valido = df_cp['IPC_Ult_Actualizacion'].notna() & (df_cp['IPC_Ult_Actualizacion'] > 0)
    df_cp['Honorario Sugerido'] = df_cp['precio']
    df_cp.loc[coef_valido, 'Honorario Sugerido'] = (
        df_cp.loc[coef_valido, 'precio'] *
        (ultimo_indice / df_cp.loc[coef_valido, 'IPC_Ult_Actualizacion'])
    )

    # Alerta (vectorizado)
    df_cp['Alerta_Revisión'] = 'OK'
    mask_vencido = (
        (df_cp['Actualiza'] == 'Si') &
        (df_cp['Estado']    == 'Activo') &
        (df_cp['Meses Desactualizado'] >= df_cp['periodos'])
    )
    df_cp.loc[mask_vencido, 'Alerta_Revisión'] = 'VENCIDO'

    # Ordenar: activos primero, luego por precio descendente
    df_cp['Orden_Estado'] = (df_cp['Estado'] != 'Activo').astype(int)
    df_cp = df_cp.sort_values(['Orden_Estado', 'precio'], ascending=[True, False]).drop(columns=['Orden_Estado'])
    df_cp['Actualizacion_Str'] = df_cp['Actualizacion_dt'].dt.strftime('%m/%Y')

    # --- Índices visual ---
    df_indices_visual = df_indices.copy()
    df_indices_visual['MES'] = df_indices_visual['MES_dt'].dt.strftime('%m/%Y')
    df_indices_visual = df_indices_visual.drop(columns=['MES_dt'])

    return {
        "historial":      df_historial_visual,
        "clientes":       df_cp,
        "indices_vis":    df_indices_visual,
        "ultimo_mes":     ultimo_mes,
        "ultimo_indice":  ultimo_indice,
        "clientes_orig":  df_clientes,
        "facturas_orig":  df_facturas,
        "indices_orig":   df_indices,
        "motor_interno":  df_final,
    }


# ============================================================
# 6. LECTURA CON INVALIDACIÓN AUTOMÁTICA POR CAMBIO DE ARCHIVO
# ============================================================
if not os.path.exists(ARCHIVO_LOCAL):
    st.error(f"⚠️ No se encontró '{ARCHIVO_LOCAL}'. Subí el archivo desde la barra lateral.")
    with st.sidebar:
        archivo_subido = st.file_uploader("Subir Excel Maestro", type=["xlsx"])
        if archivo_subido:
            with open(ARCHIVO_LOCAL, "wb") as f:
                f.write(archivo_subido.getbuffer())
            st.success("¡Archivo cargado!")
            st.rerun()
    st.stop()

mod_time = os.path.getmtime(ARCHIVO_LOCAL)
datos = cargar_y_procesar_datos(ARCHIVO_LOCAL, mod_time)

if not datos:
    st.error("❌ No se pudieron procesar los datos del Excel. Revisá que tenga las hojas 'Facturas', 'Clientes' e 'Indices'.")
    st.stop()

# --- Desempaquetar el diccionario de datos ---
df_historial_base   = datos["historial"]
df_clientes         = datos["clientes"]
df_indices_vis      = datos["indices_vis"]
ult_mes             = datos["ultimo_mes"]
ult_ind             = datos["ultimo_indice"]
df_clientes_orig    = datos["clientes_orig"]
df_facturas_orig    = datos["facturas_orig"]
df_indices_orig     = datos["indices_orig"]
df_motor_interno    = datos["motor_interno"]

# ============================================================
# 7. PRE-CÁLCULO DE SEMÁFOROS
# ============================================================
estado_honorarios = "🟢"
if not df_clientes[df_clientes['Alerta_Revisión'] == 'VENCIDO'].empty:
    estado_honorarios = "🔴"

estado_afip = "🟢"
fecha_max   = df_motor_interno['Fecha_dt'].max()
fecha_12m   = fecha_max - pd.DateOffset(years=1)
fecha_3m    = fecha_max - pd.DateOffset(months=3)

def nivel_alerta_afip(emisor: str, cat_actual: str) -> int:
    """Devuelve 1 (OK), 2 (amarillo) o 3 (rojo) según el margen al tope de categoría."""
    df_e     = df_motor_interno[df_motor_interno['Emisor'] == emisor]
    fact_12m = df_e[(df_e['Fecha_dt'] > fecha_12m) & (df_e['Fecha_dt'] <= fecha_max)]['Facturacion $'].sum()
    fact_3m  = df_e[(df_e['Fecha_dt'] > fecha_3m)  & (df_e['Fecha_dt'] <= fecha_max)]['Facturacion $'].sum()
    promedio = (fact_3m / 3) if fact_3m > 0 else (fact_12m / 12)
    margen   = ESCALAS_MONOTRIBUTO[cat_actual] - fact_12m
    if margen < 0:               return 3   # superó el límite
    elif margen <= promedio:     return 2   # queda 1 mes o menos
    return 1

nivel_ag  = nivel_alerta_afip('Agustín', st.session_state['cat_agustin'])
nivel_la  = nivel_alerta_afip('Laura',   st.session_state['cat_laura'])
max_nivel = max(nivel_ag, nivel_la)
if   max_nivel == 3: estado_afip = "🔴"
elif max_nivel == 2: estado_afip = "🟡"

# ============================================================
# 8. BARRA LATERAL: NAVEGACIÓN Y CONTROLES
# ============================================================
MENU_ESTADISTICAS = "📈 Estudio: Estadísticas"
MENU_CYGNUS       = "🏠 Cygnus Home: KPIs"
MENU_SIMULADOR    = f"{estado_honorarios} Actualización de Honorarios"
MENU_AFIP         = f"{estado_afip} Control Monotributo"
MENU_CLIENTES     = "👥 Maestro de Clientes"
MENU_FACTURAS     = "🧾 Historial de Facturas"
MENU_INDICES      = "📊 Tabla de Índices"

with st.sidebar:
    st.header("Navegación del Sistema")
    seleccion_pantalla = st.radio(
        "Ir a sección:",
        [MENU_ESTADISTICAS, MENU_CYGNUS, MENU_SIMULADOR,
         MENU_AFIP, MENU_CLIENTES, MENU_FACTURAS, MENU_INDICES],
        label_visibility="collapsed"
    )

    st.divider()

    archivo_subido = st.file_uploader("Actualizar Excel Maestro", type=["xlsx"])
    if archivo_subido is not None:
        with open(ARCHIVO_LOCAL, "wb") as f:
            f.write(archivo_subido.getbuffer())
        st.success("¡Archivo cargado y procesado!")
        st.cache_data.clear()   # ← fuerza recarga del caché
        st.rerun()

    st.divider()

    st.header("⏳ Filtro Temporal")
    meses_disponibles = sorted(df_historial_base['Mes_Indice'].dropna().unique())
    if meses_disponibles:
        ops = [m.strftime('%m/%Y') for m in meses_disponibles]
        rango_seleccionado = st.select_slider(
            "Rango de análisis:", options=ops, value=(ops[0], ops[-1]),
            label_visibility="collapsed"
        )
        fecha_inicio_filtro = pd.to_datetime(rango_seleccionado[0], format='%m/%Y')
        fecha_fin_filtro    = pd.to_datetime(rango_seleccionado[1], format='%m/%Y')
    else:
        fecha_inicio_filtro = fecha_fin_filtro = None

# ============================================================
# 9. APLICACIÓN DEL FILTRO TEMPORAL
# ============================================================
if fecha_inicio_filtro and fecha_fin_filtro:
    mask = (
        (df_motor_interno['Mes_Indice'] >= fecha_inicio_filtro) &
        (df_motor_interno['Mes_Indice'] <= fecha_fin_filtro)
    )
    df_motor_filtrado    = df_motor_interno[mask].copy()
    df_historial_filtrado = df_historial_base[
        (df_historial_base['Mes_Indice'] >= fecha_inicio_filtro) &
        (df_historial_base['Mes_Indice'] <= fecha_fin_filtro)
    ].copy()
else:
    df_motor_filtrado     = df_motor_interno.copy()
    df_historial_filtrado = df_historial_base.copy()

# ============================================================
# 10. RENDERIZADO DE PANTALLAS
# ============================================================

# ------------------------------------------------------------------
# 10.1  ESTADÍSTICAS DEL ESTUDIO
# ------------------------------------------------------------------
if seleccion_pantalla == MENU_ESTADISTICAS:
    st.subheader("📊 Estudio Contable: Análisis Evolutivo Real")

    df_est_fil = df_motor_filtrado[df_motor_filtrado['Negocio'] == 'Estudio']
    df_est_int = df_motor_interno[df_motor_interno['Negocio'] == 'Estudio']

    total_nominal     = df_est_fil['Facturacion $'].sum()
    total_actualizado = df_est_fil['Facturacion $ Actualizada'].sum()
    mes_max_real      = df_est_int['Mes_Indice'].max()

    if pd.notna(mes_max_real):
        mes_12m_atras = mes_max_real - pd.DateOffset(months=11)
        fact_ult_mes  = df_est_int[df_est_int['Mes_Indice'] == mes_max_real]['Facturacion $ Actualizada'].sum()
        fact_ult_12m  = df_est_int[df_est_int['Mes_Indice'] >= mes_12m_atras]['Facturacion $ Actualizada'].sum()

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Nominal en Rango",         formato_abreviado(total_nominal))
        col2.metric("Real en Rango",             formato_abreviado(total_actualizado))
        col3.metric("Últimos 12 Meses (Real)",   formato_abreviado(fact_ult_12m))
        col4.metric(f"Último Mes ({mes_max_real.strftime('%m/%Y')})", formato_abreviado(fact_ult_mes))

        st.divider()

        if not df_est_fil.empty:
            df_est_fil = df_est_fil.copy()
            df_est_fil['Año-Mes'] = df_est_fil['Mes_Indice'].dt.strftime('%Y-%m')
            df_evo = df_est_fil.groupby('Año-Mes')[['Facturacion $', 'Facturacion $ Actualizada']].sum().reset_index()
            df_evo.rename(columns={
                'Facturacion $':            'Nominal Histórica',
                'Facturacion $ Actualizada':'Real Indexada'
            }, inplace=True)

            fig_linea = px.line(
                df_evo, x='Año-Mes', y=['Nominal Histórica', 'Real Indexada'],
                color_discrete_sequence=['#636EFA', '#00CC96']
            )
            fig_linea.update_layout(yaxis_tickformat="$.2s", hovermode="x unified")
            fig_linea.update_traces(hovertemplate="%{y:$,.2f}")
            st.plotly_chart(fig_linea, use_container_width=True)

            st.divider()

            df_rank = (
                df_est_fil.groupby('Denominación Receptor')['Facturacion $ Actualizada']
                .sum().reset_index()
                .sort_values('Facturacion $ Actualizada', ascending=True)
            )
            fig_barras = px.bar(
                df_rank, x='Facturacion $ Actualizada', y='Denominación Receptor',
                orientation='h', color_discrete_sequence=['#AB63FA']
            )
            fig_barras.update_layout(xaxis_tickformat="$.2s", height=600)
            fig_barras.update_traces(hovertemplate="%{x:$,.2f}")
            st.plotly_chart(fig_barras, use_container_width=True)

# ------------------------------------------------------------------
# 10.2  CYGNUS HOME KPIs
# ------------------------------------------------------------------
elif seleccion_pantalla == MENU_CYGNUS:
    st.subheader("🏠 Cygnus Home: Rendimiento Comercial")

    df_cy_fil = df_motor_filtrado[df_motor_filtrado['Negocio'] == 'Cygnus Home']
    df_cy_int = df_motor_interno[df_motor_interno['Negocio'] == 'Cygnus Home']

    if df_cy_int.empty:
        st.info("Aún no hay facturas registradas en la base de datos para Cygnus Home.")
    else:
        fecha_max_cy = df_cy_int['Mes_Indice'].max()
        f_12m = fecha_max_cy - pd.DateOffset(months=11)
        f_6m  = fecha_max_cy - pd.DateOffset(months=5)

        fact_ult_mes = df_cy_int[df_cy_int['Mes_Indice'] == fecha_max_cy]['Facturacion $ Actualizada'].sum()
        fact_6m      = df_cy_int[df_cy_int['Mes_Indice'] >= f_6m]['Facturacion $ Actualizada'].sum()
        fact_12m     = df_cy_int[df_cy_int['Mes_Indice'] >= f_12m]['Facturacion $ Actualizada'].sum()

        col1, col2, col3 = st.columns(3)
        col1.metric("Último Mes",       f"$ {fact_ult_mes:,.2f}")
        col2.metric("Últimos 6 Meses",  formato_abreviado(fact_6m))
        col3.metric("Últimos 12 Meses", formato_abreviado(fact_12m))

        st.divider()

        if not df_cy_fil.empty:
            df_cy_fil = df_cy_fil.copy()
            df_cy_fil['Año-Mes'] = df_cy_fil['Mes_Indice'].dt.strftime('%Y-%m')
            df_evo_cy = df_cy_fil.groupby('Año-Mes')[['Facturacion $ Actualizada']].sum().reset_index()
            fig_cy = px.line(
                df_evo_cy, x='Año-Mes', y='Facturacion $ Actualizada',
                color_discrete_sequence=['#FF7F0E']
            )
            fig_cy.update_layout(
                title="Evolución Real de Cygnus Home",
                yaxis_tickformat="$.2s", hovermode="x unified"
            )
            fig_cy.update_traces(hovertemplate="%{y:$,.2f}")
            st.plotly_chart(fig_cy, use_container_width=True)

# ------------------------------------------------------------------
# 10.3  SIMULADOR DE HONORARIOS
# ------------------------------------------------------------------
elif seleccion_pantalla == MENU_SIMULADOR:
    st.subheader("🛠️ Entorno Interactivo de Actualización de Abonos")

    clientes_vencidos = df_clientes[df_clientes['Alerta_Revisión'] == 'VENCIDO']
    if not clientes_vencidos.empty:
        st.error(f"⚠️ Atención: Hay **{len(clientes_vencidos)}** clientes con honorarios atrasados.")
    else:
        st.success("✅ Todos los clientes están con sus honorarios al día.")

    df_sim = df_clientes[df_clientes['Estado'] == 'Activo'].copy()
    df_sim['Nuevo Precio Pactado'] = df_sim['precio']

    cols_sim = ['Nro. Doc. Receptor', 'Denominación Receptor', 'precio',
                'Meses Desactualizado', 'Honorario Sugerido', 'Nuevo Precio Pactado']

    df_editado = st.data_editor(
        df_sim[cols_sim],
        use_container_width=True,
        disabled=['Nro. Doc. Receptor', 'Denominación Receptor', 'precio',
                  'Meses Desactualizado', 'Honorario Sugerido'],
        column_config={
            "precio":               st.column_config.NumberColumn("Precio Actual",        format="$ %.2f"),
            "Honorario Sugerido":   st.column_config.NumberColumn("Sugerido por IPC",      format="$ %.2f"),
            "Nuevo Precio Pactado": st.column_config.NumberColumn("Nuevo Precio Pactado ✏️", format="$ %.2f"),
        },
        key="editor_abonos"
    )

    if st.button("💾 Guardar y Actualizar Base de Datos", type="primary"):
        df_maestro_nuevo = df_clientes_orig.copy()
        cambios = 0
        for _, row in df_editado.iterrows():
            cuit      = row['Nro. Doc. Receptor']
            nuevo_val = row['Nuevo Precio Pactado']
            if nuevo_val != row['precio']:
                df_maestro_nuevo.loc[df_maestro_nuevo['Nro. Doc. Receptor'] == cuit, 'precio']       = nuevo_val
                df_maestro_nuevo.loc[df_maestro_nuevo['Nro. Doc. Receptor'] == cuit, 'Actualizacion'] = datetime.today().strftime('%Y-%m-%d')
                cambios += 1

        if cambios > 0:
            try:
                with pd.ExcelWriter(ARCHIVO_LOCAL, engine='openpyxl') as writer:
                    df_maestro_nuevo.to_excel(writer,   sheet_name='Clientes', index=False)
                    df_facturas_orig.to_excel(writer,   sheet_name='Facturas', index=False)
                    df_indices_orig.to_excel(writer,    sheet_name='Indices',  index=False)
                st.cache_data.clear()   # invalida el caché para forzar recarga
                st.success(f"✅ Se actualizaron {cambios} clientes. Refrescando sistema...")
                st.rerun()
            except Exception as e:
                st.error(f"❌ No se pudo guardar el archivo: {e}")
        else:
            st.info("No se realizaron cambios.")

# ------------------------------------------------------------------
# 10.4  CONTROL MONOTRIBUTO / AFIP
# ------------------------------------------------------------------
elif seleccion_pantalla == MENU_AFIP:
    st.subheader("🏛️ Panel de Recategorización e Inscripción Activa")

    def renderizar_afip(emisor: str, cat_actual: str, col):
        with col:
            st.write(f"### 👤 {emisor}")
            df_e = df_motor_interno[df_motor_interno['Emisor'] == emisor]

            if df_e.empty:
                st.info(f"Sin facturación registrada para {emisor}.")
                return

            facturacion_12m = df_e[
                (df_e['Fecha_dt'] > fecha_12m) & (df_e['Fecha_dt'] <= fecha_max)
            ]['Facturacion $'].sum()

            # Categoría sugerida
            cat_sug = "Excluido"
            for c, lim in ESCALAS_MONOTRIBUTO.items():
                if facturacion_12m <= lim:
                    cat_sug = c
                    break

            st.metric("Facturación Nominal (Últimos 12 meses)", f"$ {facturacion_12m:,.2f}")
            st.info(f"📍 Categoría AFIP Sugerida: **{cat_sug}** (Actual en App: **{cat_actual}**)")

            # --- Barra de progreso hacia el límite de la categoría actual ---
            tope_actual = ESCALAS_MONOTRIBUTO[cat_actual]
            pct = min(facturacion_12m / tope_actual, 1.0)
            color_barra = "🟢" if pct < 0.75 else ("🟡" if pct < 1.0 else "🔴")
            st.write(f"{color_barra} **{pct * 100:.1f}%** del tope de categoría {cat_actual} "
                     f"($ {tope_actual:,.0f})")
            st.progress(pct)

            # --- Proyección de cierre de año ---
            fact_3m_emisor = df_e[
                (df_e['Fecha_dt'] > fecha_3m) & (df_e['Fecha_dt'] <= fecha_max)
            ]['Facturacion $'].sum()
            promedio_mensual = fact_3m_emisor / 3 if fact_3m_emisor > 0 else facturacion_12m / 12
            meses_restantes  = 12 - datetime.now().month
            proyeccion_anual = facturacion_12m + (promedio_mensual * meses_restantes)

            if proyeccion_anual > tope_actual:
                meses_hasta_tope = (tope_actual - facturacion_12m) / promedio_mensual if promedio_mensual > 0 else 0
                st.warning(
                    f"⚡ A este ritmo, superarías el tope en **{meses_hasta_tope:.1f} meses**. "
                    f"Proyección anual: **{formato_abreviado(proyeccion_anual)}**"
                )
            else:
                st.success(
                    f"📅 Proyección anual: **{formato_abreviado(proyeccion_anual)}** "
                    f"— dentro del tope de categoría {cat_actual}."
                )

    col_izq, col_der = st.columns(2)
    renderizar_afip('Agustín', st.session_state['cat_agustin'], col_izq)
    renderizar_afip('Laura',   st.session_state['cat_laura'],   col_der)

    st.divider()
    st.write("### 🛠️ Asistente de Sincronización")
    hizo_tramite = st.checkbox("¿Ya realizaste la recategorización en la web de AFIP?")
    if hizo_tramite:
        nueva_cat_ag = st.selectbox(
            "Nueva Categoría Agustín:",
            options=list(ESCALAS_MONOTRIBUTO.keys()),
            index=list(ESCALAS_MONOTRIBUTO.keys()).index(st.session_state['cat_agustin'])
        )
        nueva_cat_la = st.selectbox(
            "Nueva Categoría Laura:",
            options=list(ESCALAS_MONOTRIBUTO.keys()),
            index=list(ESCALAS_MONOTRIBUTO.keys()).index(st.session_state['cat_laura'])
        )
        if st.button("✅ Confirmar y Aplicar"):
            st.session_state['cat_agustin'] = nueva_cat_ag
            st.session_state['cat_laura']   = nueva_cat_la
            st.success("¡Sincronizado! Los nuevos topes están activos.")
            st.rerun()

# ------------------------------------------------------------------
# 10.5  MAESTRO DE CLIENTES
# ------------------------------------------------------------------
elif seleccion_pantalla == MENU_CLIENTES:
    st.subheader("👥 Maestro de Clientes Completo")

    df_cv = df_clientes.copy()
    df_cv['Actualizacion'] = df_cv['Actualizacion_Str']

    cols_maestro = [
        'Nro. Doc. Receptor', 'Denominación Receptor', 'Formalidad',
        'Periodicidad', 'precio', 'Estado', 'Actualiza', 'Actualizacion',
        'periodos', 'Meses Desactualizado', 'Alerta_Revisión'
    ]
    df_estilado = df_cv[cols_maestro].style.apply(colorear_clientes, axis=1)
    st.dataframe(
        df_estilado,
        use_container_width=True,
        column_config={"precio": st.column_config.NumberColumn("Precio Unitario", format="$ %.2f")}
    )

# ------------------------------------------------------------------
# 10.6  HISTORIAL DE FACTURAS
# ------------------------------------------------------------------
elif seleccion_pantalla == MENU_FACTURAS:
    st.subheader("🧾 Historial de Facturación")

    vista = st.radio(
        "Filtrar visualización de facturas:",
        ["🏢 Estudio Contable", "🏠 Cygnus Home", "🌎 Mostrar Todas"],
        horizontal=True
    )

    if vista == "🏢 Estudio Contable":
        df_render = df_historial_filtrado[df_historial_filtrado['Negocio'] == 'Estudio'].copy()
    elif vista == "🏠 Cygnus Home":
        df_render = df_historial_filtrado[df_historial_filtrado['Negocio'] == 'Cygnus Home'].copy()
    else:
        df_render = df_historial_filtrado.copy()

    df_render['Fecha'] = df_render['Fecha_dt'].dt.strftime('%d/%m/%Y')
    df_render = df_render.sort_values('Fecha_dt', ascending=False)

    cols_tabla = [
        'Fecha', 'Emisor', 'Negocio', 'Tipo', 'Punto de Venta',
        'Número Desde', 'Nro. Doc. Receptor', 'Denominación Receptor',
        'Facturacion $', 'Facturacion $ Actualizada'
    ]
    st.dataframe(
        df_render[[c for c in cols_tabla if c in df_render.columns]],
        use_container_width=True,
        column_config={
            "Facturacion $":            st.column_config.NumberColumn("Original",       format="$ %.2f"),
            "Facturacion $ Actualizada":st.column_config.NumberColumn("A Plata de Hoy", format="$ %.2f"),
        }
    )

# ------------------------------------------------------------------
# 10.7  TABLA DE ÍNDICES
# ------------------------------------------------------------------
elif seleccion_pantalla == MENU_INDICES:
    st.subheader("📊 Índices de Referencia")
    st.dataframe(df_indices_vis, use_container_width=True)