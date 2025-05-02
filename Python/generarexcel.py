import pandas as pd
import random
import datetime

# Función para generar fechas aleatorias de 2024
def fecha_aleatoria():
    inicio = datetime.date(2024, 1, 1)
    fin = datetime.date(2024, 12, 31)
    delta = (fin - inicio).days
    return inicio + datetime.timedelta(days=random.randint(0, delta))

# Función para generar fechas de inicio de suscripción entre diciembre 2023 y diciembre 2024
def fecha_suscripcion():
    inicio = datetime.date(2023, 12, 1)
    fin = datetime.date(2024, 12, 31)
    delta = (fin - inicio).days
    return inicio + datetime.timedelta(days=random.randint(0, delta))

# Listas de datos de ejemplo
ciudades = ["Lima", "Arequipa", "Cusco", "Trujillo", "Piura", "Iquitos"]
estaciones = ["Estación Central", "Parque Kennedy", "Plaza de Armas", "Universidad Nacional", "Terminal Sur"]
tipos_bicicleta = ["Urbana", "Eléctrica", "Montaña", "Plegable"]
marcas = ["BikeX", "EcoRide", "CityWheel", "UrbanGo"]
usuarios = ["Usuario Frecuente", "Nuevo Usuario", "Turista", "Estudiante"]
formas_pago = ["Tarjeta", "App Móvil", "Efectivo", "Código QR"]
empleados = ["Luis Ramírez", "Carla Torres", "Diego Mendoza", "Fernanda Soto"]
tipos_suscripcion = ["Básica", "Premium", "Familiar"]

# Número de registros
num_registros = 30000

# Generar datos
registros = []
for i in range(num_registros):
    fecha = fecha_aleatoria()
    ciudad = random.choice(ciudades)
    estacion = random.choice(estaciones)
    tipo_bici = random.choice(tipos_bicicleta)
    marca = random.choice(marcas)
    duracion_min = random.randint(15, 180)  # en minutos
    tiene_suscripcion = random.choice(["Sí", "No"])

    if tiene_suscripcion == "Sí":
        tipo_suscripcion = random.choice(tipos_suscripcion)
        fecha_inicio_sus = fecha_suscripcion().strftime("%Y-%m-%d")
        tarifa_minuto = round(random.uniform(0.05, 0.2), 2)  # tarifa preferencial
    else:
        tipo_suscripcion = ""
        fecha_inicio_sus = ""
        tarifa_minuto = round(random.uniform(0.2, 0.5), 2)

    total_pago = round(duracion_min * tarifa_minuto, 2)
    usuario = random.choice(usuarios)
    hora_inicio = f"{random.randint(6, 21)}:{random.randint(0, 59):02d}"
    forma_pago = random.choice(formas_pago)
    descuento = round(random.uniform(0, 5), 2) if random.random() < 0.2 else 0
    empleado = random.choice(empleados)

    registros.append([
        i + 1, fecha.strftime("%Y-%m-%d"), ciudad, estacion, tipo_bici, marca, duracion_min, tarifa_minuto, total_pago,
        usuario, hora_inicio, forma_pago, descuento, empleado,
        tiene_suscripcion, tipo_suscripcion, fecha_inicio_sus
    ])

# Crear DataFrame con las nuevas columnas
df = pd.DataFrame(registros, columns=[
    "ID", "Fecha Alquiler", "Ciudad", "Estación", "Tipo Bicicleta", "Marca", "Duración (min)", "Tarifa por Minuto", "Total Pago",
    "Tipo de Usuario", "Hora de Inicio", "Forma de Pago", "Descuento Aplicado", "Empleado Responsable",
    "Tiene Suscripción", "Tipo de Suscripción", "Inicio de Suscripción"
])

# Guardar en un archivo Excel
nombre_archivo = "reporte_alquiler_bicicletas_suscripciones.xlsx"
df.to_excel(nombre_archivo, index=False)

print(f"Reporte generado: {nombre_archivo}")
