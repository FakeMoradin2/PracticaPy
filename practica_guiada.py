"""Práctica guiada: limpiar, analizar, graficar, exportar y cargar ventas."""

import os
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from pymongo import MongoClient

ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(exist_ok=True)
CSV_PATH = ROOT / "ventas_tienda_sucia.csv"

# ---------------------------------------------------------------------------
# 2. Cargar y explorar el CSV
# ---------------------------------------------------------------------------
df = pd.read_csv(CSV_PATH)

print("=== FORMA ===")
print(df.shape)
print("\n=== INFO ===")
df.info()
print("\n=== HEAD ===")
print(df.head())
print("\n=== NULOS ===")
print(df.isna().sum())
print("\n=== PRODUCTO ===")
print(df["Producto"].unique())
print("\n=== SUCURSAL ===")
print(df["Sucursal"].unique())
print("\n=== METODO DE PAGO ===")
print(df["Metodo_Pago"].unique())

# ---------------------------------------------------------------------------
# 3. Limpiar los registros
# ---------------------------------------------------------------------------
df_limpio = df.copy()

df_limpio["Producto"] = df_limpio["Producto"].str.strip().str.title()
df_limpio["Sucursal"] = df_limpio["Sucursal"].str.strip().str.title()
df_limpio["Metodo_Pago"] = df_limpio["Metodo_Pago"].str.strip().str.title()
df_limpio["Categoria"] = df_limpio["Categoria"].str.strip().str.title()
df_limpio["Cliente"] = df_limpio["Cliente"].str.strip().str.title()

df_limpio["Fecha"] = pd.to_datetime(df_limpio["Fecha"], errors="coerce")
df_limpio["Precio_Unitario"] = pd.to_numeric(df_limpio["Precio_Unitario"], errors="coerce")
df_limpio["Cantidad"] = pd.to_numeric(df_limpio["Cantidad"], errors="coerce")

negativos = int((df_limpio["Precio_Unitario"] < 0).sum())
df_limpio.loc[df_limpio["Precio_Unitario"] < 0, "Precio_Unitario"] = np.nan

df_limpio["Precio_Unitario"] = df_limpio.groupby("Producto")["Precio_Unitario"].transform(
    lambda s: s.fillna(s.median())
)

df_limpio["Cantidad"] = df_limpio["Cantidad"].fillna(1).astype(int)
df_limpio = df_limpio.dropna(subset=["Fecha"]).reset_index(drop=True)
df_limpio["Total_Venta"] = df_limpio["Precio_Unitario"] * df_limpio["Cantidad"]

print("\n=== PRECIOS NEGATIVOS CONVERTIDOS A NULO ===")
print(negativos)
print("\n=== NULOS DESPUES DE LIMPIAR ===")
print(df_limpio.isna().sum())
print("\n=== DESCRIBE ===")
print(df_limpio.describe())
print(f"\nRegistros originales: {len(df)} | Registros limpios: {len(df_limpio)}")

# ---------------------------------------------------------------------------
# 4. Análisis para toma de decisiones
# ---------------------------------------------------------------------------
ingresos_producto = (
    df_limpio.groupby("Producto")["Total_Venta"].sum().sort_values(ascending=False)
)
ingresos_sucursal = (
    df_limpio.groupby("Sucursal")["Total_Venta"].sum().sort_values(ascending=False)
)

df_limpio["Mes"] = df_limpio["Fecha"].dt.to_period("M")
ventas_mensuales = df_limpio.groupby("Mes")["Total_Venta"].sum()
metodos = df_limpio.groupby("Metodo_Pago")["Total_Venta"].sum().sort_values(ascending=False)
tickets = df_limpio.groupby("Metodo_Pago").size().sort_values(ascending=False)

print("\n=== INGRESOS POR PRODUCTO ===")
print(ingresos_producto.round(2))
print("\n=== INGRESOS POR SUCURSAL ===")
print(ingresos_sucursal.round(2))
print("\n=== VENTAS MENSUALES ===")
print(ventas_mensuales.round(2))
print("\n=== METODOS (INGRESO) ===")
print(metodos.round(2))
print("\n=== METODOS (TICKETS) ===")
print(tickets)

# ---------------------------------------------------------------------------
# 5. Graficar con Matplotlib
# ---------------------------------------------------------------------------
fig, axs = plt.subplots(2, 2, figsize=(14, 9))
fig.suptitle("Análisis de ventas para toma de decisiones", fontsize=16)

axs[0, 0].bar(ingresos_producto.index, ingresos_producto.values, color="steelblue")
axs[0, 0].set_title("Ingresos por producto")
axs[0, 0].tick_params(axis="x", rotation=45)
axs[0, 0].set_ylabel("Ingreso")

axs[0, 1].bar(ingresos_sucursal.index, ingresos_sucursal.values, color="seagreen")
axs[0, 1].set_title("Ingresos por sucursal")
axs[0, 1].tick_params(axis="x", rotation=45)
axs[0, 1].set_ylabel("Ingreso")

axs[1, 0].plot(
    ventas_mensuales.index.astype(str),
    ventas_mensuales.values,
    marker="o",
    color="darkorange",
)
axs[1, 0].set_title("Tendencia mensual de ingresos")
axs[1, 0].tick_params(axis="x", rotation=45)
axs[1, 0].set_ylabel("Ingreso")

axs[1, 1].pie(metodos.values, labels=metodos.index, autopct="%1.1f%%")
axs[1, 1].set_title("Distribución por método de pago")

plt.tight_layout()
plt.savefig(OUTPUT / "analisis_ventas.png", dpi=150)
plt.close()

# ---------------------------------------------------------------------------
# Reto: boxplot de Total_Venta por producto
# ---------------------------------------------------------------------------
orden_productos = ingresos_producto.index.tolist()
datos_box = [
    df_limpio.loc[df_limpio["Producto"] == producto, "Total_Venta"].values
    for producto in orden_productos
]

fig_box, ax_box = plt.subplots(figsize=(10, 6))
ax_box.boxplot(datos_box, tick_labels=orden_productos, showfliers=True)
ax_box.set_title("Total de venta por producto (valores atípicos)")
ax_box.set_ylabel("Total_Venta")
ax_box.tick_params(axis="x", rotation=30)
fig_box.tight_layout()
fig_box.savefig(OUTPUT / "boxplot_total_venta.png", dpi=150)
plt.close(fig_box)


atipicos_partes = []
for producto, grupo in df_limpio.groupby("Producto"):
    q1 = grupo["Total_Venta"].quantile(0.25)
    q3 = grupo["Total_Venta"].quantile(0.75)
    limite = q3 + 1.5 * (q3 - q1)
    parte = grupo.loc[grupo["Total_Venta"] > limite].copy()
    parte["limite_iqr"] = limite
    atipicos_partes.append(parte)

atipicos = pd.concat(atipicos_partes) if atipicos_partes else df_limpio.iloc[0:0]
print("\n=== ATIPICOS POR IQR (Total_Venta, por producto) ===")
print(f"Cantidad: {len(atipicos)} de {len(df_limpio)}")
if len(atipicos):
    resumen_atip = (
        atipicos.groupby("Producto")
        .agg(
            n=("Total_Venta", "size"),
            total_min=("Total_Venta", "min"),
            total_max=("Total_Venta", "max"),
            precio_min=("Precio_Unitario", "min"),
            precio_max=("Precio_Unitario", "max"),
            cantidad_max=("Cantidad", "max"),
        )
        .round(2)
    )
    print(resumen_atip)

precio_mediana = df_limpio.groupby("Producto")["Precio_Unitario"].median()
df_limpio["precio_vs_mediana"] = df_limpio["Precio_Unitario"] / df_limpio["Producto"].map(precio_mediana)
precios_raros = df_limpio.loc[df_limpio["precio_vs_mediana"] > 2, [
    "ID_Venta", "Producto", "Precio_Unitario", "Cantidad", "Total_Venta", "precio_vs_mediana"
]]
print("\n=== PRECIOS UNITARIOS MAYORES AL DOBLE DE LA MEDIANA ===")
print(precios_raros.round(2).to_string(index=False) if len(precios_raros) else "Ninguno")

# Decisión: no se eliminan. Los picos salen de cantidad alta por un precio
# unitario creíble (laptops, monitores, tablets). Quitarlos borraría ventas reales.
df_para_mongo = df_limpio.drop(columns=["precio_vs_mediana"])

# ---------------------------------------------------------------------------
# 6. Exportar a Excel con openpyxl
# ---------------------------------------------------------------------------
wb = Workbook()

ws = wb.active
ws.title = "Datos Limpios"
for row in dataframe_to_rows(df_para_mongo.drop(columns=["Mes"]), index=False, header=True):
    ws.append(row)

for cell in ws[1]:
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="4472C4")

for column in ws.columns:
    width = max(len(str(cell.value)) for cell in column if cell.value is not None)
    ws.column_dimensions[column[0].column_letter].width = min(width + 2, 40)

ws2 = wb.create_sheet("Resumen Producto")
ws2.append(["Producto", "Ingreso Total"])
for producto, total in ingresos_producto.items():
    ws2.append([producto, float(total)])

ws3 = wb.create_sheet("Resumen Sucursal")
ws3.append(["Sucursal", "Ingreso Total"])
for sucursal, total in ingresos_sucursal.items():
    ws3.append([sucursal, float(total)])

excel_path = OUTPUT / "reporte_ventas.xlsx"
wb.save(excel_path)
print(f"\nReporte Excel generado: {excel_path}")

# ---------------------------------------------------------------------------
# 7. Guardar en MongoDB
# Atlas pide MONGO_URI. Si no está, se usa el MongoDB local del equipo.
# ---------------------------------------------------------------------------
mongo_uri = os.environ.get("MONGO_URI", "mongodb://localhost:27017/")
cliente = MongoClient(mongo_uri, serverSelectionTimeoutMS=5000)
cliente.admin.command("ping")

coleccion = cliente["tienda"]["ventas_limpias"]
coleccion.delete_many({})

registros = df_para_mongo.copy()
registros["Fecha"] = registros["Fecha"].dt.strftime("%Y-%m-%d")
registros["Mes"] = registros["Mes"].astype(str)
registros = registros.drop(columns=["Mes"])

resultado = coleccion.insert_many(registros.to_dict(orient="records"))
print(f"\nMongoDB ({mongo_uri}): documentos insertados: {len(resultado.inserted_ids)}")
print("Conteo:", coleccion.count_documents({}))
print("Ejemplo:", coleccion.find_one())

# ---------------------------------------------------------------------------
# 8. Consultas útiles en MongoDB
# ---------------------------------------------------------------------------
print("\n=== INGRESOS POR PRODUCTO (Mongo) ===")
pipeline = [
    {"$group": {"_id": "$Producto", "ingreso": {"$sum": "$Total_Venta"}}},
    {"$sort": {"ingreso": -1}},
]
for doc in coleccion.aggregate(pipeline):
    print(doc)

print("\n=== VENTAS DE ZAPOPAN CON TARJETA (5) ===")
consulta = coleccion.find({"Sucursal": "Zapopan", "Metodo_Pago": "Tarjeta"}).limit(5)
for venta in consulta:
    print(venta)

print("\nListo.")
