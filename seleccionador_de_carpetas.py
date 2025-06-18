import os
import shutil
import zipfile
import pandas as pd

def validar_ruta(ruta):
    return os.path.exists(ruta)

def copiar_carpetas(facturas, ruta_origen, ruta_destino):
    if not os.path.exists(ruta_destino):
        os.makedirs(ruta_destino)

    copiadas = []
    for factura in facturas:
        nombre_carpeta = f"AttachedDocument_F-010-{factura}"
        origen = os.path.join(ruta_origen, nombre_carpeta)
        destino = os.path.join(ruta_destino, nombre_carpeta)
        if os.path.exists(origen):
            shutil.copytree(origen, destino)
            copiadas.append(destino)
            print(f"✅ Copiada: {nombre_carpeta}")
        else:
            print(f"⚠️ No encontrada: {nombre_carpeta}")
    return copiadas

def comprimir_carpetas(ruta_destino, nombre_zip='facturas_comprimidas.zip'):
    zip_path = os.path.join(ruta_destino, nombre_zip)
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for carpeta_raiz, _, archivos in os.walk(ruta_destino):
            for archivo in archivos:
                if archivo != nombre_zip:
                    ruta_completa = os.path.join(carpeta_raiz, archivo)
                    ruta_relativa = os.path.relpath(ruta_completa, ruta_destino)
                    zipf.write(ruta_completa, ruta_relativa)
    print(f"📦 Carpetas comprimidas en: {zip_path}")

def main():
    print("\n=== Copiador de Carpetas de Facturas ===")
    ruta_excel = input("📄 Ruta del archivo .xlsx con los números de factura: ").strip()
    ruta_arbol = input("📁 Ruta del árbol de carpetas de origen: ").strip()
    ruta_destino = input("📂 Ruta de destino para las carpetas copiadas: ").strip()
    nombre_columna = input("📊 Nombre de la columna con los números de factura: ").strip()

    if not validar_ruta(ruta_excel):
        print("❌ La ruta del archivo Excel no existe.")
        return
    if not validar_ruta(ruta_arbol):
        print("❌ La ruta del árbol de carpetas no existe.")
        return

    # Leer facturas del Excel
    try:
        df = pd.read_excel(ruta_excel)
        facturas = df[nombre_columna].dropna().astype(int).tolist()
    except Exception as e:
        print(f"❌ Error al leer el archivo Excel: {e}")
        return

    # Copiar carpetas que coincidan
    carpetas_copiadas = copiar_carpetas(facturas, ruta_arbol, ruta_destino)

    if not carpetas_copiadas:
        print("⚠️ No se encontraron carpetas para copiar.")
        return
    # Comprimir carpetas copiadas
    # comprimir_carpetas(ruta_destino)


if __name__ == '__main__':
    main()
