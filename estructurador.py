import os
import shutil
import json
import pandas as pd
import re
from datetime import datetime


def validar_ruta(ruta):
    """Verifica si la ruta existe"""
    return os.path.exists(ruta)


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


def copiar_carpetas(facturas, origen, destino):
    """Copia carpetas de facturas desde origen a destino"""
    if not os.path.exists(destino):
        os.makedirs(destino)

    copiadas = []
    no_encontradas = []
    for factura in facturas:
        nombre_carpeta = f"AttachedDocument_F-010-{factura}"
        origen_carpeta = os.path.join(origen, nombre_carpeta)
        destino_carpeta = os.path.join(destino, nombre_carpeta)

        if os.path.exists(origen_carpeta):
            shutil.copytree(origen_carpeta, destino_carpeta, dirs_exist_ok=True)
            copiadas.append(destino_carpeta)
            print(f"✅ Copiada: {nombre_carpeta}")
        else:
            print(f"⚠️ No encontrada: {nombre_carpeta}")
            no_encontradas.append(factura)

    # Guardar facturas no encontradas en un Excel
    if no_encontradas:
        excel_path = os.path.join(destino, "facturas_no_encontradas.xlsx")
        try:
            df = pd.DataFrame({"Factura": no_encontradas})
            df.to_excel(excel_path, index=False)
            print(f"Facturas no encontradas guardadas en {excel_path}")
        except Exception as e:
            print(f"Error al escribir en {excel_path}: {e}")

    return copiadas


def reestructurar_archivo(factura, ruta_carpeta):
    """Reestructura el archivo JSON ResultadosMSPS_FE{factura}_ID0_R.txt"""
    archivo_json = os.path.join(ruta_carpeta, f"ResultadosMSPS_FE{factura}_ID0_R.txt")

    if not os.path.exists(archivo_json):
        print(f"⚠️ Archivo {os.path.basename(archivo_json)} no encontrado en {ruta_carpeta}")
        return False

    try:
        # Leer el JSON original
        with open(archivo_json, 'r', encoding='utf-8') as f:
            data = json.load(f)

        # Extraer ProcesoId y CodigoUnicoValidacion de Observaciones
        proceso_id = 0  # Valor predeterminado si no se encuentra
        codigo_unico = ""  # Valor predeterminado si no se encuentra
        resultados_validacion = []

        if data.get("ResultadosValidacion") and len(data["ResultadosValidacion"]) > 0:
            # Buscar ProcesoId
            observaciones = data["ResultadosValidacion"][0].get("Observaciones", "")
            match_proceso = re.search(r"ProcesoId (\d+)", observaciones)
            if match_proceso:
                proceso_id = int(match_proceso.group(1))

            # Buscar CodigoUnicoValidacion en Observaciones
            for validacion in data["ResultadosValidacion"]:
                print(validacion.get("Observaciones", ""))
                obs = validacion.get("Observaciones", "")
                match_cuv = re.search(r"[0-9a-f]{128}", obs)
                if match_cuv:
                    codigo_unico = match_cuv.group(0)
                    break

            # Mantener solo las validaciones de clase NOTIFICACION
            resultados_validacion = [
                validacion for validacion in data["ResultadosValidacion"]
                if validacion.get("Clase") == "NOTIFICACION"
            ]

        # Crear la nueva estructura
        nueva_data = {
            "ResultState": True,
            "ProcesoId": proceso_id,
            "NumFactura": f"FE{factura}",
            "CodigoUnicoValidacion": codigo_unico if codigo_unico else "No encontrado",
            "FechaRadicacion": datetime.now().strftime("%Y-%m-%dT%H:%M:%S.%f-05:00"),
            "RutaArchivos": ruta_carpeta,
            "ResultadosValidacion": resultados_validacion
        }

        # Escribir el nuevo JSON
        with open(archivo_json, 'w', encoding='utf-8') as f:
            json.dump(nueva_data, f, indent=2, ensure_ascii=False)
        print(f"✅ Reestructurado: {os.path.basename(archivo_json)} en {ruta_carpeta} con ProcesoId {proceso_id}")
        return True
    except Exception as e:
        print(f"Error al reestructurar {os.path.basename(archivo_json)}: {e}")
        return False


def main():
    """Función principal para procesar facturas y reestructurar archivos JSON"""
    try:
        ruta_xlsx = input("Ingrese la ruta del archivo .xlsx con facturas 📄: ")
        columna = input("Ingrese el nombre de la columna con números de facturas 📋: ")
        ruta_carpetas = input("Ingrese la ruta de las carpetas 📂: ")
        ruta_destino = input("Ingrese la ruta de destino 📁: ")
    except EOFError:
        print(
            "❌ Error: No se pudo leer la entrada. Asegúrese de proporcionar las rutas interactivamente o use argumentos de línea de comandos.")
        return

    # Validar rutas
    if not validar_ruta(ruta_xlsx):
        print("⚠️ Verifique la ruta del archivo .xlsx y vuelva a intentarlo.")
        return
    if not validar_ruta(ruta_carpetas):
        print("⚠️ Verifique la ruta de las carpetas y vuelva a intentarlo.")
        return
    if not validar_ruta(ruta_destino):
        print("⚠️ La ruta de destino no existe. Creándola...")
        os.makedirs(ruta_destino)

    # Obtener facturas del Excel
    facturas = procesar_archivo_xlsx(ruta_xlsx, columna)
    if not facturas:
        print("❌ No se encontraron facturas en el archivo .xlsx")
        return

    print(f"Facturas encontradas: {facturas}")

    # Copiar carpetas y obtener las rutas de las carpetas copiadas
    carpetas_copiadas = copiar_carpetas(facturas, ruta_carpetas, ruta_destino)

    # Reestructurar archivos JSON en las carpetas copiadas
    for factura in facturas:
        nombre_carpeta = f"AttachedDocument_F-010-{factura}"
        ruta_carpeta = os.path.join(ruta_destino, nombre_carpeta)
        if ruta_carpeta in carpetas_copiadas:
            reestructurar_archivo(factura, ruta_carpeta)


if __name__ == "__main__":
    main()