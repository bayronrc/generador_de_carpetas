import pandas as pd
import os
import shutil


def verificar_ruta(ruta):
    """Verifica si la ruta existe"""
    return os.path.exists(ruta)


def crear_directorio(ruta):
    """Crea un directorio si no existe"""
    if not os.path.exists(ruta):
        print(f"Creando directorio: {ruta}")
        os.makedirs(ruta)


def procesar_archivo_xlsx(ruta_xlsx, columna):
    """Procesa el archivo .xlsx y devuelve lista de números de factura"""
    try:
        df = pd.read_excel(ruta_xlsx)
        if columna not in df.columns:
            print(f"La columna {columna} no se encuentra en el archivo .xlsx.")
            return []
        return [str(numero) for numero in df[columna].dropna().astype(str)]
    except Exception as e:
        print(f"Error al leer el archivo .xlsx: {e}")
        return []


def validar_factura_soporte(numero_factura, ruta_soportes):
    """Valida si existe carpeta con número de factura en soportes"""
    ruta_carpeta = os.path.join(ruta_soportes, f"FE{numero_factura}")
    if os.path.isdir(ruta_carpeta):
        return ruta_carpeta
    print(f"Carpeta FE{numero_factura} no encontrada en {ruta_soportes}")
    return None


def copiar_carpeta_soporte(ruta_origen, ruta_destino, numero_factura):
    """Copia la carpeta de soporte a la ruta destino"""
    destino_carpeta = os.path.join(ruta_destino, f"FE{numero_factura}")
    try:
        shutil.copytree(ruta_origen, destino_carpeta, dirs_exist_ok=True)
        print(f"Carpeta FE{numero_factura} copiada a {destino_carpeta}")
        return destino_carpeta
    except Exception as e:
        print(f"Error al copiar carpeta FE{numero_factura}: {e}")
        return None


def procesar_factura(ruta_facturas, numero_factura, ruta_carpeta_destino):
    """Agrega contenido de factura a la carpeta copiada"""
    carpeta_factura = f"AttachedDocument_F-010-{numero_factura}"
    ruta_origen_factura = os.path.join(ruta_facturas, carpeta_factura)

    if not os.path.exists(ruta_origen_factura):
        print(f"Carpeta de factura {carpeta_factura} no encontrada")
        return False

    try:
        for item in os.listdir(ruta_origen_factura):
            ruta_item = os.path.join(ruta_origen_factura, item)
            destino_item = os.path.join(ruta_carpeta_destino, item)

            if os.path.isfile(ruta_item):
                shutil.copy2(ruta_item, destino_item)
                print(f"Archivo {item} copiado a {destino_item}")
            elif os.path.isdir(ruta_item):
                shutil.copytree(ruta_item, destino_item, dirs_exist_ok=True)
                print(f"Carpeta {item} copiada a {destino_item}")
        return True
    except Exception as e:
        print(f"Error al procesar factura {numero_factura}: {e}")
        return False


def main():
    """Función principal para procesar facturas y soportes"""
    # Solicitar rutas
    ruta_xlsx = input("Ingrese la ruta del archivo .xlsx: ")
    columna = input("Ingrese el nombre de la columna con números de facturas: ")
    ruta_destino = input("Ingrese la ruta de destino: ")
    ruta_facturas = input("Ingrese la ruta de las facturas: ")
    ruta_soportes = input("Ingrese la ruta de los soportes: ")

    # Validar rutas
    for ruta, nombre in [(ruta_xlsx, "archivo .xlsx"),
                         (ruta_facturas, "facturas"),
                         (ruta_soportes, "soportes")]:
        if not verificar_ruta(ruta):
            print(f"Verifique la ruta de {nombre} y vuelva a intentarlo.")
            return

    # Crear directorio destino
    crear_directorio(ruta_destino)

    # Procesar archivo xlsx
    facturas = procesar_archivo_xlsx(ruta_xlsx, columna)
    if not facturas:
        print("No se encontraron números de factura en el archivo .xlsx")
        return

    print(f"Facturas encontradas: {facturas}")

    # Procesar cada factura
    for factura in facturas:
        # Validar existencia de carpeta en soportes
        ruta_carpeta_soporte = validar_factura_soporte(factura, ruta_soportes)
        if not ruta_carpeta_soporte:
            continue

        # Copiar carpeta de soporte
        ruta_carpeta_copiada = copiar_carpeta_soporte(
            ruta_carpeta_soporte,
            ruta_destino,
            factura
        )
        if not ruta_carpeta_copiada:
            continue

        # Procesar y copiar contenido de factura
        procesar_factura(ruta_facturas, factura, ruta_carpeta_copiada)


if __name__ == '__main__':
    main()