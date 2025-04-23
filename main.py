
import os
import shutil
import re


def crear_directorio(ruta):
    """Crea un directorio si no existe."""
    os.makedirs(ruta, exist_ok=True)
    return ruta


def copiar_archivo(origen, destino):
    """Copia un archivo de una ruta origen a una ruta destino."""
    try:
        shutil.copy2(origen, destino)
        return True
    except Exception as e:
        print(f"Error al copiar archivo: {str(e)}")
        return False


def extraer_numero_factura(nombre_archivo):
    """Extrae el número de factura del nombre del archivo para el JSON."""
    # Ejemplo: AttachedDocument_F-010-200816 -> 200816
    match = re.search(r'(\d+)$', nombre_archivo)
    return match.group(1) if match else ""


def verificar_json_existente(nombre_carpeta, ruta_json_origen):
    """Verifica si el archivo JSON correspondiente existe."""
    numero_factura = extraer_numero_factura(nombre_carpeta)  # Ej: 200816
    if numero_factura:
        nombre_json = f"FE{numero_factura}.json"  # Ej: FE200816.json
        ruta_json = os.path.join(ruta_json_origen, nombre_json)
        if os.path.exists(ruta_json):
            return True, ruta_json, nombre_json
    return False, None, None


def procesar_factura_xml(ruta_base, ruta_destino, archivo):
    """Procesa un archivo XML y crea su carpeta correspondiente."""
    nombre_carpeta = os.path.splitext(archivo)[0]  # Ej: AttachedDocument_F-010-200816
    ruta_carpeta = os.path.join(ruta_destino, nombre_carpeta)

    # Crear carpeta
    crear_directorio(ruta_carpeta)

    # Copiar XML con su nombre original
    ruta_origen = os.path.join(ruta_base, archivo)
    ruta_xml_destino = os.path.join(ruta_carpeta, archivo)  # Mantiene AttachedDocument_F-010-200816.xml
    copiar_archivo(ruta_origen, ruta_xml_destino)

    return nombre_carpeta, ruta_carpeta


def copiar_json_factura(nombre_carpeta, ruta_carpeta, ruta_json_origen):
    """Copia el archivo JSON con el formato FE<numero> a la carpeta."""
    numero_factura = extraer_numero_factura(nombre_carpeta)  # Ej: 200816
    if numero_factura:
        nombre_json = f"FE{numero_factura}.json"  # Ej: FE200816.json
        ruta_json = os.path.join(ruta_json_origen, nombre_json)

        if os.path.exists(ruta_json):
            ruta_json_destino = os.path.join(ruta_carpeta, nombre_json)
            if copiar_archivo(ruta_json, ruta_json_destino):
                print(f"Copiado {nombre_json} a {ruta_carpeta}")
            else:
                print(f"Error al copiar {nombre_json}")
        else:
            print(f"No se encontró {nombre_json} en {ruta_json_origen}")


def procesar_facturas(ruta_base, ruta_destino, ruta_json_origen):
    """Procesa todas las facturas XML y sus JSON correspondientes."""
    # Crear directorio destino principal
    crear_directorio(ruta_destino)

    # Recorrer directorios y procesar archivos
    for carpeta_raiz, _, archivos in os.walk(ruta_base):
        for archivo in archivos:
            if archivo.lower().endswith(".xml"):
                # Extraer nombre de la carpeta
                nombre_carpeta = os.path.splitext(archivo)[0]

                # Verificar si el JSON existe antes de crear la carpeta
                json_existe, ruta_json, nombre_json = verificar_json_existente(nombre_carpeta, ruta_json_origen)
                if json_existe:
                    # Si el JSON existe, procesar el XML y crear la carpeta
                    nombre_carpeta, ruta_carpeta = procesar_factura_xml(carpeta_raiz, ruta_destino, archivo)
                    # Copiar el JSON a la carpeta creada
                    copiar_json_factura(nombre_carpeta, ruta_carpeta, ruta_json_origen)
                else:
                    print(f"No se creó la carpeta para {nombre_carpeta} porque no se encontró el JSON asociado.")

    print(f"Carpetas creadas y procesadas en {ruta_destino}")


# Configuración de rutas
# ruta_base = 'D:/007-Invoices/2025/MAR-2025'
ruta_base = 'D:/007-Invoices/2025/FEB-2025'
ruta_destino = 'D:/007-Invoices/2025-Directories'
ruta_json_origen = 'D:/007-Invoices/json_files'

# Ejecutar el procesamiento
if __name__ == "__main__":
    procesar_facturas(ruta_base, ruta_destino, ruta_json_origen)


# Este script procesa archivos XML y JSON, creando carpetas para cada factura
