import pandas as pd

# Leer el archivo de origen
nombre_archivo_origen = "reporte_alquiler_bicicletas_suscripciones.xlsx"
df_origen = pd.read_excel(nombre_archivo_origen)

# Crear diccionarios para asignar IDs únicos a las dimensiones
dimensiones = {
    "Ciudad": {},
    "Estación": {},
    "Tipo Bicicleta": {},
    "Marca": {},
    "Tipo de Usuario": {},
    "Empleado Responsable": {},
    "Forma de Pago": {},
    "Tipo de Suscripción": {}
}

# Crear DataFrames para cada dimensión
tablas_dimensiones = {}

for nombre, dic in dimensiones.items():
    valores_unicos = df_origen[nombre].dropna().unique()
    dic.update({valor: idx + 1 for idx, valor in enumerate(valores_unicos)})
    tablas_dimensiones[nombre] = pd.DataFrame({
        f"ID_{nombre.replace(' ', '_')}": [dic[valor] for valor in valores_unicos],
        nombre: valores_unicos
    })

# Crear la tabla de hechos
registros = []
for _, row in df_origen.iterrows():
    fecha_alquiler = pd.to_datetime(row["Fecha Alquiler"])
    id_suscripcion = dimensiones["Tipo de Suscripción"].get(row["Tipo de Suscripción"], None)

    registros.append([
        row["ID"],
        row["Fecha Alquiler"], fecha_alquiler.year, fecha_alquiler.month, fecha_alquiler.day,
        dimensiones["Ciudad"][row["Ciudad"]],
        dimensiones["Estación"][row["Estación"]],
        dimensiones["Tipo Bicicleta"][row["Tipo Bicicleta"]],
        dimensiones["Marca"][row["Marca"]],
        row["Duración (min)"], row["Tarifa por Minuto"], row["Total Pago"],
        dimensiones["Tipo de Usuario"][row["Tipo de Usuario"]],
        row["Hora de Inicio"],
        dimensiones["Forma de Pago"][row["Forma de Pago"]],
        row["Descuento Aplicado"],
        dimensiones["Empleado Responsable"][row["Empleado Responsable"]],
        1 if row["Tiene Suscripción"] == "Sí" else 0,
        id_suscripcion,
        row["Inicio de Suscripción"] if pd.notna(row["Inicio de Suscripción"]) else ""
    ])

# Crear DataFrame de hechos
df_hechos = pd.DataFrame(registros, columns=[
    "ID_Alquiler", "Fecha_Alquiler", "Año", "Mes", "Día",
    "ID_Ciudad", "ID_Estación", "ID_Tipo_Bicicleta", "ID_Marca",
    "Duración_Min", "Tarifa_Minuto", "Total_Pago",
    "ID_Tipo_Usuario", "Hora_Inicio",
    "ID_Forma_Pago", "Descuento",
    "ID_Empleado",
    "Tiene_Suscripción", "ID_Tipo_Suscripción", "Inicio_Suscripción"
])

# Guardar todo en un solo archivo Excel (una hoja por tabla)
nombre_archivo_final = "nuevo_modelo_datos_bicicletas.xlsx"
with pd.ExcelWriter(nombre_archivo_final, engine="xlsxwriter") as writer:
    df_hechos.to_excel(writer, sheet_name="Hechos_Alquiler", index=False)
    for nombre_dim, df_dim in tablas_dimensiones.items():
        sheet_name = f"Dim_{nombre_dim.replace(' ', '_')[:31]}"  # Hoja limitada a 31 caracteres
        df_dim.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Archivo Excel generado con hechos y dimensiones: {nombre_archivo_final}")
