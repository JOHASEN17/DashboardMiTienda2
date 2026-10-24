import os
from pathlib import Path
import pandas as pd
import streamlit as st
import seaborn as sns
import matplotlib.pyplot as plt

st.set_page_config(page_title="MiTienda | Dashboard", page_icon="📊", layout="wide")
sns.set_theme(style="whitegrid", palette="deep")
BASE_DIR = Path(__file__).parent

@st.cache_data
def cargar_excel():
    candidatos = [BASE_DIR / "LaTiendaCom_convertido.xlsx", BASE_DIR / "LaTiendaCom.xlsx", BASE_DIR / "LaTiendaCom (1)(1).xls", BASE_DIR / "LaTiendaCom (1).xls"]
    archivo = next((p for p in candidatos if p.exists()), None)
    if archivo is None:
        st.error("No se encontró el Excel. Sube al repositorio LaTiendaCom_convertido.xlsx junto a dashboard.py.")
        st.stop()
    xls = pd.ExcelFile(archivo)
    hojas = {nombre: pd.read_excel(archivo, sheet_name=nombre) for nombre in xls.sheet_names}
    for hoja, columna in [("Facturas", "fecha_factura"), ("Pedidos", "fecha_pedido"), ("Clientes", "fecha_creacion"), ("Empleados", "fecha_contratacion"), ("Productos", "fecha_vencimiento"), ("Proveedores", "fecha_servicios")]:
        if hoja in hojas and columna in hojas[hoja].columns:
            hojas[hoja][columna] = pd.to_datetime(hojas[hoja][columna], errors="coerce")
    if "Facturas" in hojas:
        f = hojas["Facturas"]
        for col in ["cantidad_productos", "precio_unitario"]:
            if col in f.columns: f[col] = pd.to_numeric(f[col], errors="coerce")
        if {"cantidad_productos", "precio_unitario"}.issubset(f.columns):
            f["importe_estimado"] = f["cantidad_productos"].fillna(0) * f["precio_unitario"].fillna(0)
        else: f["importe_estimado"] = 0
    if "Productos" in hojas:
        for col in ["precio", "stock"]:
            if col in hojas["Productos"].columns: hojas["Productos"][col] = pd.to_numeric(hojas["Productos"][col], errors="coerce")
    if "Empleados" in hojas and "salario" in hojas["Empleados"].columns:
        hojas["Empleados"]["salario"] = pd.to_numeric(hojas["Empleados"]["salario"], errors="coerce")
    return hojas

def dinero(valor):
    return "$ {:,.2f}".format(float(valor or 0)).replace(",", "X").replace(".", ",").replace("X", ".")

def grafico(titulo, df, kind, x=None, y=None, xlabel=None, ylabel=None, rotation=0, top=None, bins=24):
    if df.empty or (x and x not in df.columns) or (y and y not in df.columns): return
    fig, ax = plt.subplots(figsize=(9, 3.7))
    if kind == "line": sns.lineplot(data=df, x=x, y=y, marker="o", ax=ax)
    elif kind == "bar": sns.barplot(data=df, x=x, y=y, ax=ax)
    elif kind == "barh": sns.barplot(data=df, x=x, y=y, ax=ax)
    elif kind == "hist": sns.histplot(data=df, x=x, bins=bins, kde=True, ax=ax)
    ax.set_title(titulo); ax.set_xlabel(xlabel or x or ""); ax.set_ylabel(ylabel or y or "")
    if rotation: ax.tick_params(axis="x", rotation=rotation)
    fig.tight_layout(); st.pyplot(fig, use_container_width=True); plt.close(fig)

def contar_por_anio(df, col, nombre):
    if col not in df.columns: return pd.DataFrame()
    temp = df.dropna(subset=[col]).copy()
    if temp.empty: return pd.DataFrame()
    temp["año"] = temp[col].dt.year.astype(str)
    return temp.groupby("año", as_index=False).size().rename(columns={"size": nombre}).tail(18)

st.markdown("<div style='padding:24px 28px;border-radius:16px;background:linear-gradient(120deg,#172554,#263fa9 65%,#4773e8);color:white'><div style='letter-spacing:2px;font-size:12px'>ANALÍTICA DE DATOS · MITIENDA</div><h1 style='margin:8px 0'>Panel de control comercial</h1><p style='margin:0'>Resumen de facturas, pedidos, inventario, clientes, empleados y proveedores.</p></div>", unsafe_allow_html=True)
st.write("")
hojas = cargar_excel()
facturas = hojas.get("Facturas", pd.DataFrame()).copy()
productos = hojas.get("Productos", pd.DataFrame()).copy()
pedidos = hojas.get("Pedidos", pd.DataFrame()).copy()
clientes = hojas.get("Clientes", pd.DataFrame()).copy()
empleados = hojas.get("Empleados", pd.DataFrame()).copy()
proveedores = hojas.get("Proveedores", pd.DataFrame()).copy()

st.sidebar.header("Filtros de facturas")
if "fecha_factura" in facturas.columns and facturas["fecha_factura"].notna().any():
    fechas = facturas["fecha_factura"].dropna()
    minimo, maximo = fechas.min().date(), fechas.max().date()
    rango = st.sidebar.date_input("Periodo de facturación", value=(minimo, maximo), min_value=minimo, max_value=maximo)
else: rango = None
pagos = sorted(facturas["forma_de_pago"].dropna().astype(str).unique()) if "forma_de_pago" in facturas else []
tipos = sorted(facturas["tipo_de_venta"].dropna().astype(str).unique()) if "tipo_de_venta" in facturas else []
pago = st.sidebar.selectbox("Forma de pago", ["Todas"] + pagos)
tipo = st.sidebar.selectbox("Tipo de venta", ["Todos"] + tipos)
filtradas = facturas.copy()
if rango and isinstance(rango, (tuple, list)) and len(rango) == 2 and "fecha_factura" in filtradas:
    filtradas = filtradas[filtradas["fecha_factura"].dt.date.between(rango[0], rango[1])]
if pago != "Todas" and "forma_de_pago" in filtradas: filtradas = filtradas[filtradas["forma_de_pago"].astype(str) == pago]
if tipo != "Todos" and "tipo_de_venta" in filtradas: filtradas = filtradas[filtradas["tipo_de_venta"].astype(str) == tipo]

importe = filtradas["importe_estimado"].fillna(0).sum() if "importe_estimado" in filtradas else 0
unidades = int(filtradas["cantidad_productos"].fillna(0).sum()) if "cantidad_productos" in filtradas else 0
n_productos = productos["id_producto"].nunique() if "id_producto" in productos else len(productos)
k1,k2,k3,k4 = st.columns(4)
k1.metric("Facturas filtradas", f"{len(filtradas):,}")
k2.metric("Importe estimado", dinero(importe))
k3.metric("Unidades facturadas", f"{unidades:,}")
k4.metric("Productos en catálogo", f"{n_productos:,}")
st.caption("El importe es una estimación calculada como cantidad × precio unitario; no representa necesariamente el total contable verificado.")
st.divider()
st.header("Gráficos de análisis")
col1,col2 = st.columns(2)
with col1:
    if {"fecha_factura", "importe_estimado"}.issubset(filtradas.columns):
        d=filtradas.dropna(subset=["fecha_factura"]).copy()
        if not d.empty:
            d["mes"] = d["fecha_factura"].dt.to_period("M").astype(str)
            grafico("Evolución mensual del importe estimado", d.groupby("mes",as_index=False)["importe_estimado"].sum().tail(18), "line", "mes", "importe_estimado", "Mes", "Importe estimado", 45)
    if "forma_de_pago" in filtradas:
        d=filtradas["forma_de_pago"].fillna("Sin dato").value_counts().rename_axis("forma_de_pago").reset_index(name="facturas")
        grafico("Facturas por forma de pago",d,"bar","forma_de_pago","facturas","Forma de pago","Número de facturas",20)
    d=contar_por_anio(pedidos,"fecha_pedido","pedidos")
    grafico("Pedidos registrados por año",d,"bar","año","pedidos","Año","Cantidad de pedidos",45)
    d=contar_por_anio(clientes,"fecha_creacion","clientes")
    grafico("Clientes creados por año",d,"line","año","clientes","Año","Clientes",45)
    d=contar_por_anio(proveedores,"fecha_servicios","proveedores")
    grafico("Proveedores por año de registro de servicios",d,"bar","año","proveedores","Año","Proveedores",45)
with col2:
    if {"nombre_producto","stock"}.issubset(productos.columns):
        d=productos.dropna(subset=["nombre_producto","stock"]).nlargest(10,"stock")
        grafico("10 productos con mayor stock",d,"barh","nombre_producto","stock","Producto","Unidades en stock")
    if "precio" in productos.columns:
        grafico("Distribución de precios de productos",productos.dropna(subset=["precio"]),"hist","precio",None,"Precio","Cantidad de productos")
    if "tipo_de_venta" in filtradas:
        d=filtradas["tipo_de_venta"].fillna("Sin dato").value_counts().rename_axis("tipo_de_venta").reset_index(name="facturas")
        grafico("Facturas por tipo de venta",d,"bar","tipo_de_venta","facturas","Tipo de venta","Número de facturas")
    if {"cargo","salario"}.issubset(empleados.columns):
        d=empleados.dropna(subset=["cargo","salario"]).groupby("cargo",as_index=False)["salario"].mean().nlargest(10,"salario")
        grafico("Salario promedio por cargo (máximo 10)",d,"barh","cargo","salario","Salario promedio","Cargo")
st.divider()
st.header("Explorar tablas de datos")
for nombre, df in hojas.items():
    with st.expander(f"{nombre} · {len(df):,} registros"):
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button(f"Descargar {nombre} como CSV", df.to_csv(index=False).encode("utf-8-sig"), file_name=f"{nombre.lower()}.csv", mime="text/csv", key=f"csv_{nombre}")
st.subheader("Detalle de facturas filtradas")
cols = [c for c in ["id_factura","fecha_factura","id_cliente","id_producto","cantidad_productos","precio_unitario","importe_estimado","forma_de_pago","tipo_de_venta"] if c in filtradas.columns]
st.dataframe(filtradas[cols].head(100), use_container_width=True, hide_index=True)
st.info("Los datos son de ejemplo y pueden contener fechas históricas. Confirma en la guía si precio_unitario corresponde al precio por unidad o al importe total de línea.")
st.caption("MiTienda · Dashboard en Streamlit · Datos cargados desde el archivo Excel del repositorio.")
