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

# Asignar IDs únicos
for col in dimensiones.keys():
    valores_unicos = df_origen[col].dropna().unique()  # dropna para evitar errores con celdas vacías
    dimensiones[col] = {valor: idx + 1 for idx, valor in enumerate(valores_unicos)}

# Crear la tabla de hechos
registros = []
for _, row in df_origen.iterrows():
    fecha_alquiler = pd.to_datetime(row["Fecha Alquiler"])
    
    id_suscripcion = dimensiones["Tipo de Suscripción"].get(row["Tipo de Suscripción"], None)
    
    registros.append([
        row["ID"],
        row["Fecha Alquiler"], fecha_alquiler.year, fecha_alquiler.month, fecha_alquiler.day,
        dimensiones["Ciudad"][row["Ciudad"]], row["Ciudad"],
        dimensiones["Estación"][row["Estación"]], row["Estación"],
        dimensiones["Tipo Bicicleta"][row["Tipo Bicicleta"]], row["Tipo Bicicleta"],
        dimensiones["Marca"][row["Marca"]], row["Marca"],
        row["Duración (min)"], row["Tarifa por Minuto"], row["Total Pago"],
        dimensiones["Tipo de Usuario"][row["Tipo de Usuario"]], row["Tipo de Usuario"],
        row["Hora de Inicio"],
        dimensiones["Forma de Pago"][row["Forma de Pago"]], row["Forma de Pago"],
        row["Descuento Aplicado"],
        dimensiones["Empleado Responsable"][row["Empleado Responsable"]], row["Empleado Responsable"],
        row["Tiene Suscripción"],
        id_suscripcion, row["Tipo de Suscripción"], row["Inicio de Suscripción"]
    ])

# Crear DataFrame de hechos
df_hechos = pd.DataFrame(registros, columns=[
    "ID_Alquiler", "Fecha_Alquiler", "Año", "Mes", "Día",
    "ID_Ciudad", "Ciudad",
    "ID_Estación", "Estación",
    "ID_Tipo_Bicicleta", "Tipo_Bicicleta",
    "ID_Marca", "Marca",
    "Duración_Min", "Tarifa_Minuto", "Total_Pago",
    "ID_Tipo_Usuario", "Tipo_Usuario",
    "Hora_Inicio",
    "ID_Forma_Pago", "Forma_Pago",
    "Descuento",
    "ID_Empleado", "Empleado",
    "Tiene_Suscripción",
    "ID_Tipo_Suscripción", "Tipo_Suscripción", "Inicio_Suscripción"
])

# Guardar en archivo Excel
nombre_archivo_destino = "tabla_hechos_alquiler_bicicletas.xlsx"
df_hechos.to_excel(nombre_archivo_destino, index=False)

print(f"Tabla de hechos generada: {nombre_archivo_destino}")
