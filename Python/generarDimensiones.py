import pandas as pd

# Leer el archivo de origen
nombre_archivo_origen = "reporte_alquiler_bicicletas_suscripciones.xlsx"
df_origen = pd.read_excel(nombre_archivo_origen)

# Definir columnas de dimensiones
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

# Crear IDs únicos para cada valor en cada dimensión
for col in dimensiones.keys():
    valores_unicos = df_origen[col].dropna().unique()
    dimensiones[col] = {valor: idx + 1 for idx, valor in enumerate(valores_unicos)}

# Crear registros de la tabla de hechos
hechos = []
for _, row in df_origen.iterrows():
    fecha = pd.to_datetime(row["Fecha Alquiler"])
    
    id_ciudad = dimensiones["Ciudad"][row["Ciudad"]]
    id_estacion = dimensiones["Estación"][row["Estación"]]
    id_tipo_bici = dimensiones["Tipo Bicicleta"][row["Tipo Bicicleta"]]
    id_marca = dimensiones["Marca"][row["Marca"]]
    id_usuario = dimensiones["Tipo de Usuario"][row["Tipo de Usuario"]]
    id_empleado = dimensiones["Empleado Responsable"][row["Empleado Responsable"]]
    id_pago = dimensiones["Forma de Pago"][row["Forma de Pago"]]
    id_suscripcion = dimensiones["Tipo de Suscripción"].get(row["Tipo de Suscripción"], None)

    hechos.append([
        row["ID"],
        row["Fecha Alquiler"], fecha.year, fecha.month, fecha.day,
        id_ciudad,
        id_estacion,
        id_tipo_bici,
        id_marca,
        row["Duración (min)"], row["Tarifa por Minuto"], row["Total Pago"],
        id_usuario,
        row["Hora de Inicio"],
        id_pago,
        row["Descuento Aplicado"],
        id_empleado,
        row["Tiene Suscripción"],
        id_suscripcion,
        row["Inicio de Suscripción"]
    ])

# Crear DataFrame de hechos
df_hechos = pd.DataFrame(hechos, columns=[
    "ID_Alquiler", "Fecha_Alquiler", "Año", "Mes", "Día",
    "ID_Ciudad",
    "ID_Estación",
    "ID_Tipo_Bicicleta",
    "ID_Marca",
    "Duración_Min", "Tarifa_Minuto", "Total_Pago",
    "ID_Tipo_Usuario",
    "Hora_Inicio",
    "ID_Forma_Pago",
    "Descuento",
    "ID_Empleado",
    "Tiene_Suscripción",
    "ID_Tipo_Suscripción",
    "Inicio_Suscripción"
])

# Guardar tabla de hechos
df_hechos.to_excel("tabla_hechos_alquiler_bicicletas.xlsx", index=False)
print("✔ Tabla de hechos generada: tabla_hechos_alquiler_bicicletas.xlsx")

# Guardar cada dimensión en un archivo separado
for dimension, diccionario in dimensiones.items():
    df_dim = pd.DataFrame(list(diccionario.items()), columns=[dimension, f"ID_{dimension.replace(' ', '_')}"])
    df_dim = df_dim[[f"ID_{dimension.replace(' ', '_')}", dimension]]  # Ordenar columnas
    nombre_archivo_dim = f"dim_{dimension.replace(' ', '_').lower()}.xlsx"
    df_dim.to_excel(nombre_archivo_dim, index=False)
    print(f"✔ Tabla de dimensión generada: {nombre_archivo_dim}")
