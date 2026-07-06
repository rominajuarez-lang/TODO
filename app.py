# ============================================================
# DASHBOARD STREAMLIT: SIMULACION DE INVENTARIOS + FORECAST + AHORRO
# Para Streamlit Cloud / GitHub
# Archivos necesarios: app.py y requirements.txt
# Ejecutar local:
#   pip install -r requirements.txt
#   streamlit run app.py
# ============================================================

from pathlib import Path
import math
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Inventarios, Forecast y Ahorro", layout="wide")

NOMBRE_EXCEL_DEFAULT = "excel_final.xlsx"
HOJA_DEMANDA = "Demanda"
HOJA_DATOS = "Datos"
HOJA_FORECAST_EMPRESA = "Forescast_Comercial"  # nombre tal como aparece en tu Excel

# ============================================================
# UTILIDADES
# ============================================================
def limpiar_columnas(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def formato_soles(x):
    try:
        x = float(x)
    except Exception:
        return "S/ 0"
    return f"S/ {x:,.2f}"


def formato_num(x):
    try:
        x = float(x)
    except Exception:
        return "0"
    if abs(x) >= 1_000_000:
        return f"{x/1_000_000:.2f}M"
    if abs(x) >= 1_000:
        return f"{x/1_000:.2f}K"
    return f"{x:,.0f}"


def elegir_columna(df, opciones):
    cols_lower = {str(c).strip().lower(): c for c in df.columns}
    for op in opciones:
        if op.lower() in cols_lower:
            return cols_lower[op.lower()]
    return None

# ============================================================
# LECTURA DEL EXCEL
# ============================================================
@st.cache_data(show_spinner="Leyendo Excel...")
def leer_excel(archivo):
    xls = pd.ExcelFile(archivo)
    hojas = xls.sheet_names

    if HOJA_DEMANDA not in hojas:
        raise ValueError(f"No existe la hoja '{HOJA_DEMANDA}'.")
    if HOJA_DATOS not in hojas:
        raise ValueError(f"No existe la hoja '{HOJA_DATOS}'.")
    if HOJA_FORECAST_EMPRESA not in hojas:
        raise ValueError(f"No existe la hoja '{HOJA_FORECAST_EMPRESA}'.")

    demanda = limpiar_columnas(pd.read_excel(xls, sheet_name=HOJA_DEMANDA))
    datos = limpiar_columnas(pd.read_excel(xls, sheet_name=HOJA_DATOS))
    forecast_empresa = limpiar_columnas(pd.read_excel(xls, sheet_name=HOJA_FORECAST_EMPRESA))

    # Demanda
    demanda["date"] = pd.to_datetime(demanda["date"], errors="coerce")
    demanda["product_id"] = demanda["product_id"].astype(str).str.strip()
    demanda["demand_real"] = pd.to_numeric(demanda["demand_real"], errors="coerce").fillna(0)
    demanda = demanda.dropna(subset=["date", "product_id"])

    # Datos
    datos["product_id"] = datos["product_id"].astype(str).str.strip()
    for c in ["initial_stock", "lead_time_months", "review_period_months", "ss", "q_fixed", "lot_size", "cost_order", "cost_holding_month", "cost_stockout", "unit_value"]:
        if c in datos.columns:
            datos[c] = pd.to_numeric(datos[c], errors="coerce").fillna(0)

    # Forecast empresa
    forecast_empresa["date"] = pd.to_datetime(forecast_empresa["date"], errors="coerce")
    forecast_empresa["product_id"] = forecast_empresa["product_id"].astype(str).str.strip()
    forecast_empresa["forecast_company"] = pd.to_numeric(forecast_empresa["forecast_company"], errors="coerce").fillna(0)
    forecast_empresa = forecast_empresa.dropna(subset=["date", "product_id"])

    return demanda, datos, forecast_empresa

# ============================================================
# FORECAST PROPUESTO
# ============================================================
@st.cache_data(show_spinner="Calculando pronóstico propuesto...")
def calcular_forecast_propuesto(demanda, forecast_empresa):
    """
    Forecast propuesto rápido y automático:
    1) Si existe el mismo mes del año anterior, usa demanda del mismo mes del año anterior.
    2) Si no existe, usa promedio móvil de los últimos 3 meses.
    3) Si no existe, usa promedio histórico del SKU.
    """
    dem = demanda.copy().sort_values(["product_id", "date"])
    meses_obj = forecast_empresa[["product_id", "date"]].drop_duplicates().copy()

    historico = dem.set_index(["product_id", "date"])["demand_real"].to_dict()
    promedio_sku = dem.groupby("product_id")["demand_real"].mean().to_dict()

    filas = []
    for row in meses_obj.itertuples(index=False):
        sku = row.product_id
        fecha = pd.Timestamp(row.date)
        fecha_anio_ant = fecha - pd.DateOffset(years=1)

        valor = historico.get((sku, fecha_anio_ant), np.nan)

        if pd.isna(valor):
            ini = fecha - pd.DateOffset(months=3)
            sub = dem[(dem["product_id"] == sku) & (dem["date"] < fecha) & (dem["date"] >= ini)]
            valor = sub["demand_real"].mean() if not sub.empty else np.nan

        if pd.isna(valor):
            valor = promedio_sku.get(sku, 0)

        filas.append({
            "product_id": sku,
            "date": fecha,
            "forecast_propuesto": max(float(valor), 0)
        })

    return pd.DataFrame(filas)

@st.cache_data(show_spinner="Calculando ahorro...")
def preparar_comparacion(demanda, datos, forecast_empresa, forecast_propuesto):
    df = forecast_empresa.merge(forecast_propuesto, on=["product_id", "date"], how="left")
    df = df.merge(demanda[["product_id", "date", "demand_real"]], on=["product_id", "date"], how="left")
    df = df.merge(datos[["product_id", "unit_value"]], on="product_id", how="left")

    df["demand_real"] = pd.to_numeric(df["demand_real"], errors="coerce").fillna(0)
    df["forecast_company"] = pd.to_numeric(df["forecast_company"], errors="coerce").fillna(0)
    df["forecast_propuesto"] = pd.to_numeric(df["forecast_propuesto"], errors="coerce").fillna(0)
    df["unit_value"] = pd.to_numeric(df["unit_value"], errors="coerce").fillna(0)

    df["error_abs_empresa"] = (df["forecast_company"] - df["demand_real"]).abs()
    df["error_abs_propuesto"] = (df["forecast_propuesto"] - df["demand_real"]).abs()

    df["costo_error_empresa"] = df["error_abs_empresa"] * df["unit_value"]
    df["costo_error_propuesto"] = df["error_abs_propuesto"] * df["unit_value"]
    df["ahorro_soles"] = df["costo_error_empresa"] - df["costo_error_propuesto"]

    df["mes"] = df["date"].dt.strftime("%Y-%m")
    return df.sort_values(["product_id", "date"])

# ============================================================
# SIMULACION INVENTARIO
# ============================================================
@st.cache_data(show_spinner="Simulando inventario...")
def simular_sku(demanda, datos, sku):
    dem = demanda[demanda["product_id"] == sku].copy().sort_values("date")
    if dem.empty:
        return pd.DataFrame(), {}

    params = datos[datos["product_id"] == sku].copy()
    if params.empty:
        initial_stock = float(dem["demand_real"].mean() * 2)
        lead_time = 1
        ss = 0
        lot_size = max(float(dem["demand_real"].mean()), 1)
    else:
        p = params.iloc[0]
        initial_stock = float(p.get("initial_stock", 0) or 0)
        lead_time = int(max(1, math.ceil(float(p.get("lead_time_months", 1) or 1))))
        ss = float(p.get("ss", 0) or 0)
        lot_size = float(p.get("lot_size", 0) or 0)
        q_fixed = float(p.get("q_fixed", 0) or 0)
        if lot_size <= 0:
            lot_size = q_fixed
        if lot_size <= 0:
            lot_size = max(float(dem["demand_real"].mean()), 1)

    demanda_promedio = float(dem["demand_real"].mean())
    punto_reorden = demanda_promedio * lead_time + ss

    stock_fisico = initial_stock
    pedidos_pendientes = []
    filas = []
    fechas = list(dem["date"])
    demandas = list(dem["demand_real"])

    for i, (fecha, demanda_mes) in enumerate(zip(fechas, demandas)):
        llegada_mes = 0
        pendientes_nuevos = []
        for pedido in pedidos_pendientes:
            if pedido["fecha_llegada"] <= fecha:
                llegada_mes += pedido["cantidad"]
            else:
                pendientes_nuevos.append(pedido)
        pedidos_pendientes = pendientes_nuevos
        stock_fisico += llegada_mes
        stock_fisico = max(stock_fisico - demanda_mes, 0)

        inventario_posicion = stock_fisico + sum(p["cantidad"] for p in pedidos_pendientes)
        pedido_generado = 0
        if inventario_posicion <= punto_reorden:
            pedido_generado = lot_size
            indice_llegada = min(i + lead_time, len(fechas) - 1)
            pedidos_pendientes.append({"fecha_llegada": fechas[indice_llegada], "cantidad": pedido_generado})

        filas.append({
            "date": fecha,
            "Demanda Real": demanda_mes,
            "Inventario Físico": stock_fisico,
            "Punto de Reorden (s)": punto_reorden,
            "Pedido Generado": pedido_generado,
            "Llegadas de Pedido": llegada_mes,
        })

    info = {
        "stock_inicial": initial_stock,
        "lead_time_meses": lead_time,
        "ss": ss,
        "lote_pedido": lot_size,
        "punto_reorden": punto_reorden,
        "demanda_promedio": demanda_promedio,
    }
    return pd.DataFrame(filas), info

# ============================================================
# GRAFICOS
# ============================================================
def grafico_inventario(df_sim, sku):
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df_sim["date"], y=df_sim["Inventario Físico"], mode="lines+markers", name="Inventario Físico", line=dict(width=2), marker=dict(size=5), yaxis="y1"))
    fig.add_trace(go.Scatter(x=df_sim["date"], y=df_sim["Punto de Reorden (s)"], mode="lines", name="Punto de Reorden (s)", line=dict(width=2, dash="dot", color="red"), yaxis="y1"))
    fig.add_trace(go.Bar(x=df_sim["date"], y=df_sim["Demanda Real"], name="Demanda Real", opacity=0.25, yaxis="y2"))
    df_ped = df_sim[df_sim["Pedido Generado"] > 0]
    fig.add_trace(go.Scatter(x=df_ped["date"], y=df_ped["Pedido Generado"], mode="markers", name="Pedido Generado", marker=dict(symbol="triangle-up", size=12, color="orange", line=dict(width=1, color="black")), yaxis="y2"))
    fig.update_layout(
        title=f"Evolución del stock físico frente a la demanda y generación de órdenes de compra<br><sup>{sku}</sup>",
        height=650,
        hovermode="x unified",
        legend=dict(orientation="h", y=-0.18, x=0.5, xanchor="center"),
        margin=dict(l=60, r=80, t=90, b=90),
        xaxis=dict(title="Fecha", showgrid=False),
        yaxis=dict(title="Unidades en Inventario", rangemode="tozero"),
        yaxis2=dict(title="Demanda / Tamaño de Pedido", overlaying="y", side="right", rangemode="tozero"),
    )
    return fig


def grafico_forecast(df_sku, sku):
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df_sku["date"], y=df_sku["demand_real"], mode="lines+markers", name="Demanda / Ventas reales", line=dict(width=3)))
    fig.add_trace(go.Scatter(x=df_sku["date"], y=df_sku["forecast_company"], mode="lines+markers", name="Pronóstico empresa", line=dict(width=2, dash="dash")))
    fig.add_trace(go.Scatter(x=df_sku["date"], y=df_sku["forecast_propuesto"], mode="lines+markers", name="Pronóstico propuesto", line=dict(width=2)))
    fig.update_layout(
        title=f"Comparación de demanda real vs pronósticos<br><sup>{sku}</sup>",
        height=560,
        hovermode="x unified",
        legend=dict(orientation="h", y=-0.2, x=0.5, xanchor="center"),
        margin=dict(l=60, r=40, t=90, b=90),
        xaxis=dict(title="Fecha"),
        yaxis=dict(title="Unidades", rangemode="tozero"),
    )
    return fig


def grafico_top_ahorro(df_resumen, top_n):
    top = df_resumen.sort_values("ahorro_soles", ascending=False).head(top_n).copy()
    top["SKU corto"] = top["product_id"].str.slice(0, 60)
    fig = px.bar(top.sort_values("ahorro_soles"), x="ahorro_soles", y="SKU corto", orientation="h", title=f"Top {top_n} SKU con mayor ahorro del pronóstico propuesto")
    fig.update_layout(height=max(450, top_n * 28), xaxis_title="Ahorro en soles", yaxis_title="SKU", margin=dict(l=40, r=40, t=70, b=40))
    return fig

# ============================================================
# APP
# ============================================================
st.title("📊 Dashboard de inventarios, pronósticos y ahorro")
st.write("Sube tu Excel y la app mostrará: simulación de inventario, comparación de forecast y ahorro económico en soles.")

archivo_subido = st.file_uploader("Sube tu Excel", type=["xlsx"])

if archivo_subido is not None:
    fuente_excel = archivo_subido
else:
    ruta_default = Path(__file__).parent / NOMBRE_EXCEL_DEFAULT
    if ruta_default.exists():
        fuente_excel = ruta_default
    else:
        st.info("Sube tu Excel para iniciar. En GitHub/Streamlit Cloud no es necesario subir el Excel al repositorio; puedes cargarlo desde esta pantalla.")
        st.stop()

try:
    demanda, datos, forecast_empresa = leer_excel(fuente_excel)
    forecast_propuesto = calcular_forecast_propuesto(demanda, forecast_empresa)
    comparacion = preparar_comparacion(demanda, datos, forecast_empresa, forecast_propuesto)
except Exception as e:
    st.error("No se pudo procesar el Excel. Revisa que tenga las hojas Demanda, Datos y Forescast_Comercial con sus columnas correspondientes.")
    st.exception(e)
    st.stop()

skus = sorted(comparacion["product_id"].dropna().unique())

with st.sidebar:
    st.header("Filtros")
    sku = st.selectbox("SKU", skus)
    top_n = st.slider("Top ahorro", 5, 50, 20)
    mostrar_tablas = st.checkbox("Mostrar tablas", value=True)

# Resumen general
resumen_sku = comparacion.groupby("product_id", as_index=False).agg(
    demanda_total=("demand_real", "sum"),
    forecast_empresa_total=("forecast_company", "sum"),
    forecast_propuesto_total=("forecast_propuesto", "sum"),
    costo_error_empresa=("costo_error_empresa", "sum"),
    costo_error_propuesto=("costo_error_propuesto", "sum"),
    ahorro_soles=("ahorro_soles", "sum"),
)

ahorro_total = comparacion["ahorro_soles"].sum()
costo_emp = comparacion["costo_error_empresa"].sum()
costo_prop = comparacion["costo_error_propuesto"].sum()
reduccion = ((costo_emp - costo_prop) / costo_emp * 100) if costo_emp > 0 else 0

k1, k2, k3, k4 = st.columns(4)
k1.metric("Ahorro total", formato_soles(ahorro_total))
k2.metric("Costo error empresa", formato_soles(costo_emp))
k3.metric("Costo error propuesto", formato_soles(costo_prop))
k4.metric("Reducción del costo de error", f"{reduccion:.1f}%")

# Tabs
tab1, tab2, tab3 = st.tabs(["📈 Comparación forecast", "📦 Simulación inventario", "💰 Ahorro"])

with tab1:
    df_sku = comparacion[comparacion["product_id"] == sku].copy()
    st.plotly_chart(grafico_forecast(df_sku, sku), use_container_width=True)

    ahorro_sku = df_sku["ahorro_soles"].sum()
    c1, c2, c3 = st.columns(3)
    c1.metric("Ahorro del SKU", formato_soles(ahorro_sku))
    c2.metric("Error empresa", formato_num(df_sku["error_abs_empresa"].sum()))
    c3.metric("Error propuesto", formato_num(df_sku["error_abs_propuesto"].sum()))

    if mostrar_tablas:
        st.dataframe(
            df_sku[["mes", "demand_real", "forecast_company", "forecast_propuesto", "costo_error_empresa", "costo_error_propuesto", "ahorro_soles"]],
            use_container_width=True
        )

with tab2:
    df_sim, info = simular_sku(demanda, datos, sku)
    if df_sim.empty:
        st.warning("No hay datos de demanda para simular este SKU.")
    else:
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Stock inicial", formato_num(info["stock_inicial"]))
        m2.metric("Lead time", f'{info["lead_time_meses"]} meses')
        m3.metric("Punto de reorden", formato_num(info["punto_reorden"]))
        m4.metric("Lote de pedido", formato_num(info["lote_pedido"]))
        st.plotly_chart(grafico_inventario(df_sim, sku), use_container_width=True)
        if mostrar_tablas:
            st.dataframe(df_sim, use_container_width=True)

with tab3:
    st.subheader("Ahorro económico del pronóstico propuesto")
    st.write("El ahorro se calcula comparando el costo del error absoluto del pronóstico de la empresa contra el costo del error absoluto del pronóstico propuesto.")
    st.plotly_chart(grafico_top_ahorro(resumen_sku, top_n), use_container_width=True)

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("### Resumen por SKU")
        st.dataframe(resumen_sku.sort_values("ahorro_soles", ascending=False), use_container_width=True)
    with col_b:
        mensual = comparacion.groupby("mes", as_index=False).agg(
            demanda=("demand_real", "sum"),
            forecast_empresa=("forecast_company", "sum"),
            forecast_propuesto=("forecast_propuesto", "sum"),
            ahorro_soles=("ahorro_soles", "sum"),
        )
        st.markdown("### Resumen mensual")
        st.dataframe(mensual, use_container_width=True)

    salida = resumen_sku.sort_values("ahorro_soles", ascending=False).to_csv(index=False).encode("utf-8-sig")
    st.download_button("Descargar resumen de ahorro CSV", data=salida, file_name="resumen_ahorro_forecast.csv", mime="text/csv")
