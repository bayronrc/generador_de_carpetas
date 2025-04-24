import pandas as pd
import os
import shutil
import zipfile
import glob


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
    """Agrega contenido de factura a la carpeta copiada y renombra archivo ResultadosMSPS a .json en destino"""
    carpeta_factura = f"AttachedDocument_F-010-{numero_factura}"
    ruta_origen_factura = os.path.join(ruta_facturas, carpeta_factura)

    if not os.path.exists(ruta_origen_factura):
        print(f"Carpeta de factura {carpeta_factura} no encontrada")
        return False

    try:
        # Copiar todos los archivos y carpetas a la carpeta destino
        for item in os.listdir(ruta_origen_factura):
            ruta_item = os.path.join(ruta_origen_factura, item)
            destino_item = os.path.join(ruta_carpeta_destino, item)

            if os.path.isfile(ruta_item):
                shutil.copy2(ruta_item, destino_item)
                print(f"Archivo {item} copiado a {destino_item}")
            elif os.path.isdir(ruta_item):
                shutil.copytree(ruta_item, destino_item, dirs_exist_ok=True)
                print(f"Carpeta {item} copiada a {destino_item}")

        # Buscar archivo ResultadosMSPS_FE{número_factura}_*_A_CUV.txt en la carpeta destino
        patron_resultados = os.path.join(ruta_carpeta_destino, f"ResultadosMSPS_FE{numero_factura}_*_A_CUV.txt")
        archivos_resultados = glob.glob(patron_resultados)


        # Renombrar archivo(s) encontrado(s) a .json en la carpeta destino
        for archivo_txt in archivos_resultados:
            archivo_json = os.path.splitext(archivo_txt)[0] + ".json"
            os.rename(archivo_txt, archivo_json)
            print(f"Renombrado {os.path.basename(archivo_txt)} a {os.path.basename(archivo_json)} en destino")

        patron_rechazo = os.path.join(ruta_carpeta_destino, f"ResultadosMSPS_FE{numero_factura}_ID0_R.txt")
        archivos_rechazo = glob.glob(patron_rechazo)

        for archivo_rechazo in archivos_rechazo:
            try:
                os.remove(archivo_rechazo)
                print(f"Archivo {os.path.basename(archivo_rechazo)} eliminado de destino")
            except Exception as e:
                print(f"Error al eliminar archivo {os.path.basename(archivo_rechazo)} de destino : {e}")

        return True
    except Exception as e:
        print(f"Error al procesar factura {numero_factura}: {e}")
        return False


def comprimir_carpeta(ruta_carpeta, ruta_destino, numero_factura):
    """Comprime los archivos de la carpeta en un archivo ZIP con nombre FE{número_factura}.zip sin subnivel de carpeta"""
    zip_nombre = f"FE{numero_factura}.zip"
    ruta_zip = os.path.join(ruta_destino, zip_nombre)
    try:
        with zipfile.ZipFile(ruta_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(ruta_carpeta):
                for file in files:
                    archivo_ruta = os.path.join(root, file)
                    arcname = os.path.relpath(archivo_ruta, ruta_carpeta)
                    zipf.write(archivo_ruta, arcname)
        print(f"Archivo {zip_nombre} creado en {ruta_destino}")
        return ruta_zip
    except Exception as e:
        print(f"Error al comprimir carpeta FE{numero_factura}: {e}")
        return None


def eliminar_carpeta(ruta_carpeta, numero_factura):
    """Elimina la carpeta especificada"""
    try:
        shutil.rmtree(ruta_carpeta)
        print(f"Carpeta FE{numero_factura} eliminada: {ruta_carpeta}")
        return True
    except Exception as e:
        print(f"Error al eliminar carpeta FE{numero_factura}: {e}")
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
        if not procesar_factura(ruta_facturas, factura, ruta_carpeta_copiada):
            continue

        eliminar_archivo = os.path.join(ruta_carpeta_copiada, f"AttachedDocument_F-010-{factura}.xml")

        # Comprimir la carpeta copiada
        ruta_zip = comprimir_carpeta(ruta_carpeta_copiada, ruta_destino, factura)
        if not ruta_zip:
            continue

        # Eliminar la carpeta copiada
        eliminar_carpeta(ruta_carpeta_copiada, factura)


if __name__ == '__main__':
    main()