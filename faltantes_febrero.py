import os
import shutil
import pandas as pd


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


def leer_numeros_factura(ruta_xlsx):
    """Lee los números de factura desde un archivo .xlsx."""
    try:
        # Leer el archivo Excel
        df = pd.read_excel(ruta_xlsx, engine='openpyxl')
        # Asumiendo que la columna se llama 'Factura', ajusta si es diferente
        if 'Factura' not in df.columns:
            print(f"Error: No se encontró la columna 'Factura' en {ruta_xlsx}. Columnas disponibles: {list(df.columns)}")
            return []
        # Convertir a lista, sin agregar FE todavía
        return [str(numero) for numero in df['Factura'].dropna().astype(str)]
    except Exception as e:
        print(f"Error al leer el archivo .xlsx: {str(e)}")
        return []


def buscar_xml_recursivo(ruta_base, numero_factura):
    """Busca recursivamente un archivo XML que contenga el número de factura."""
    for raiz, _, archivos in os.walk(ruta_base):
        for archivo in archivos:
            if archivo.endswith('.xml') and numero_factura in archivo:
                return os.path.join(raiz, archivo)
    return None


def buscar_json(ruta_json_base, numero_factura):
    """Busca un archivo JSON con el número de factura (FE + número)."""
    try:
        numero_factura_fe = f"FE{numero_factura}"
        for archivo in os.listdir(ruta_json_base):
            if archivo.endswith('.json') and numero_factura_fe in archivo:
                return archivo
        return None
    except Exception as e:
        print(f"Error al buscar JSON: {str(e)}")
        return None


def guardar_facturas_no_encontradas(facturas_no_encontradas, ruta_destino):
    """Guarda los números de factura no encontrados en un archivo .xlsx."""
    try:
        if facturas_no_encontradas:
            df = pd.DataFrame(facturas_no_encontradas, columns=["Factura"])
            ruta_salida = os.path.join(ruta_destino, "facturas_no_encontradas.xlsx")
            df.to_excel(ruta_salida, index=False, engine='openpyxl')
            print(f"Facturas no encontradas guardadas en: {ruta_salida}")
        else:
            print("Todas las facturas tuvieron XML asociado. No se creó archivo de no encontradas.")
    except Exception as e:
        print(f"Error al guardar facturas no encontradas: {str(e)}")


def procesar_factura(numero_factura, ruta_base_xml, ruta_json_base, ruta_destino, facturas_no_encontradas):
    """Procesa una factura: busca XML y JSON, y los copia solo si existe el XML."""
    # Buscar XML primero
    ruta_xml = buscar_xml_recursivo(ruta_base_xml, numero_factura)
    if not ruta_xml:
        print(f"No se encontró XML para factura {numero_factura}. No se creará carpeta.")
        facturas_no_encontradas.append(numero_factura)
        return

    # Extraer el nombre del XML sin extensión para la carpeta
    archivo_xml = os.path.basename(ruta_xml)
    nombre_carpeta = os.path.splitext(archivo_xml)[0]  # Ej: AttachedDocument_F-010-211075
    ruta_carpeta = os.path.join(ruta_destino, nombre_carpeta)

    # Crear carpeta solo si se encontró el XML
    crear_directorio(ruta_carpeta)

    # Copiar XML
    ruta_xml_destino = os.path.join(ruta_carpeta, archivo_xml)
    copiar_archivo(ruta_xml, ruta_xml_destino)

    # Buscar y copiar JSON
    archivo_json = buscar_json(ruta_json_base, numero_factura)
    if archivo_json:
        ruta_json_origen = os.path.join(ruta_json_base, archivo_json)
        ruta_json_destino = os.path.join(ruta_carpeta, archivo_json)
        copiar_archivo(ruta_json_origen, ruta_json_destino)
        print(f"Procesado {nombre_carpeta}: XML y JSON copiados.")
    else:
        print(f"Procesado {nombre_carpeta}: Solo XML copiado.")


def main():
    """Lee números de factura de un .xlsx, busca XML y JSON, y organiza por factura."""
    ruta_xlsx = input("Ingrese la ruta del archivo .xlsx con los números de factura: ")
    ruta_base_xml = input("Ingrese la ruta base de los archivos XML (árbol de carpetas): ")
    ruta_json_base = input("Ingrese la ruta de la carpeta con archivos JSON: ")
    ruta_destino = input("Ingrese la ruta destino para las carpetas: ")

    # Validar rutas
    if not os.path.exists(ruta_xlsx):
        print("El archivo .xlsx no existe. Verifica e intenta nuevamente.")
        return
    if not os.path.exists(ruta_base_xml):
        print("La ruta base de XML no existe. Verifica e intenta nuevamente.")
        return
    if not os.path.exists(ruta_json_base):
        print("La ruta de JSON no existe. Verifica e intenta nuevamente.")
        return
    if not os.path.exists(ruta_destino):
        crear_directorio(ruta_destino)

    # Leer números de factura del .xlsx
    numeros_factura = leer_numeros_factura(ruta_xlsx)
    if not numeros_factura:
        print("No se encontraron números de factura en el archivo .xlsx.")
        return

    # Lista para almacenar facturas no encontradas
    facturas_no_encontradas = []

    # Procesar cada número de factura
    for numero_factura in numeros_factura:
        procesar_factura(numero_factura, ruta_base_xml, ruta_json_base, ruta_destino, facturas_no_encontradas)

    # Guardar facturas no encontradas en un .xlsx
    guardar_facturas_no_encontradas(facturas_no_encontradas, ruta_destino)

    print("Procesamiento completado.")


if __name__ == "__main__":
    main()