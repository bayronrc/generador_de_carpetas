import os
import re
import time
from tqdm import tqdm


def validar_ruta(ruta):
    """Verifica si la ruta existe"""
    if not os.path.exists(ruta):
        print(f"❌ La ruta {ruta} no existe.")
        return False
    return True


def renombrar_archivos_zip(ruta):
    """Renombra archivos ZIP en la misma carpeta con barra de progreso"""
    # Obtener lista de archivos ZIP
    print("🔍 Buscando archivos ZIP para renombrar...")
    archivos_zip = [
        archivo for archivo in os.listdir(ruta)
        if os.path.isfile(os.path.join(ruta, archivo)) and archivo.lower().endswith('.zip')
    ]

    if not archivos_zip:
        print("⚠️ No se encontraron archivos ZIP en la ruta.")
        return

    # Configurar la barra de progreso
    print("📈 Iniciando proceso de renombrado...")
    for archivo in tqdm(archivos_zip, desc="Procesando archivos ZIP", unit="archivo"):
        ruta_archivo = os.path.join(ruta, archivo)

        # Separar el nombre base y la extensión
        nombre_base, extension = os.path.splitext(archivo)

        match = re.match(r'^(.*?)_(\d+)$', nombre_base)
        if match:
            # Si tiene un sufijo numérico (por ejemplo, FE258058_2)
            nombre_sin_sufijo = match.group(1)  # FE258058
            numero_sufijo = int(match.group(2))  # 2
            nuevo_numero = numero_sufijo + 1  # 3
            nuevo_nombre_base = f"{nombre_sin_sufijo}_{nuevo_numero}"
        else:
            # Si no tiene sufijo numérico, agregar _1
            nuevo_nombre_base = f"{nombre_base}_1"

        # Crear el nuevo nombre del archivo
        nuevo_nombre = f"{nuevo_nombre_base}{extension}"
        nueva_ruta = os.path.join(ruta, nuevo_nombre)

        # Renombrar el archivo en la misma carpeta
        os.rename(ruta_archivo, nueva_ruta)
        print(f"📄 Archivo {archivo} renombrado a {nuevo_nombre}")

        # Simular tiempo de procesamiento (ajustable)
        time.sleep(0.1)  # Simula 0.1 segundos por archivo


def main():
    """Función principal"""
    # Solicitar la ruta de origen
    ruta = input("📦 Ingrese la ruta de origen: ")

    # Validar la ruta
    if not validar_ruta(ruta):
        return

    # Renombrar los archivos ZIP
    renombrar_archivos_zip(ruta)
    print("✅ Proceso completado.")


if __name__ == '__main__':
    main()