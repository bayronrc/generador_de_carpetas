# 1. verificar rutas
import os
from typing import List
import pandas as pd
from tqdm import tqdm
import shutil
import json


def validar_ruta(path)->bool:
    """
    Verifica si la ruta existe
    """
    return os.path.exists(path)

def crear_directorio(path)->None:
    """
    Crea el directorio si no existe
    """
    if not validar_ruta(path):
        print(f"📦 Creando directorio: {path}")
        os.makedirs(path)

def validar_directorio(path)-> bool:
     return os.path.isdir(path)
     
        
def procesar_archivo_xlsx(ruta_a_xlsx, columna)->List:
    """_summary_

    Args:
        path_to_xlsx (string): ruta a el archivo xlsx
        column_excel (string): columna en el archivo xlsx
    """
    try:
        df = pd.read_excel(ruta_a_xlsx)
        if columna not in df.columns:
            print(f"❌ La columna {columna} no se encuentra en el archivo .xlsx.")
            return []
        return [str(factura) for factura in df[columna].dropna().astype(str)]
    except Exception as e:
        print(f" ❌ Error al leer el archivo .xlsx: {e}")
        return []

def procesar_soportes(factura, ruta_soportes) -> str :
    ruta_carpeta = os.path.join(ruta_soportes, f"FE{factura}")
    if not ruta_carpeta:
        print(f"📁 ❌ Carpeta FE{factura} no encontrada en {ruta_soportes}")
    return ruta_carpeta
        
def copiar_contenido_carpeta_soporte(carpeta_soportes: str, ruta_destino: str, factura: str)-> str:
    """Copia el contenido de la carpeta origen de soportes de cada factura al el destino

    Args:
        carpeta_soportes (str): Ruta Origen de los soportes
        ruta_destino (str): Ruta destino de los soportes
        factura (str): factura procesada

    Returns:
        str: ruta de la carpeta creada destino
    """
    destino = os.path.join(ruta_destino, f"FE{factura}")
    try:
        return shutil.copytree(carpeta_soportes,destino,dirs_exist_ok=True)
    except shutil.Error as e:
        print(f"Error : {e}")
    
def procesar_factura(ruta_facturas:str, factura:str, ruta_facturas_destino:str)->bool:
    carpeta_factura = f"AttachedDocument_F-010-{factura}"
    ruta_origen_factura = os.path.join(ruta_facturas, carpeta_factura)
    ruta_destino_factura = os.path.join(ruta_facturas_destino,f"FE{factura}")
    
    
    if not os.path.exists(ruta_destino_factura):
        print(f"Ruta : {ruta_destino_factura} no existe se creara la carpeta en el destino")
        crear_directorio(ruta_destino_factura)
    
    try:
        return shutil.copytree(ruta_origen_factura,ruta_destino_factura, dirs_exist_ok=True)
    except Exception as e:
        print(f"Error: {e} ")

    

    


def main():
    """
    Funcion principal para procesar facturas y soportes
    """
    ruta_a_xlsx = input("📄 Ingrese la Ruta del Archivo Excel (xlsx) ::")
    columna_excel = input("📄 Ingrese el nombre de la columna con los numeros de facturas a radicar: ")
    ruta_destino = input("📁 Ingrese la ruta de destino: ")
    ruta_soportes = input(" Ingrese la ruta de los soportes 📦 :")
    ruta_facturas = input("Ingrese la ruta de las facturas 🗄️ :") 
    
    ruta_soportes_destino = f"{ruta_destino}\\Soportes"
    ruta_facturas_destino = f"{ruta_destino}\\Facturas"

    for ruta, _ in [
                    (ruta_a_xlsx,'archivo .xlsx'),
                    (ruta_soportes_destino,'destino soportes'),
                    (ruta_facturas_destino,'destino facturas')
                    ]:
        
        if not validar_ruta(ruta):
            crear_directorio(ruta)
    
    crear_directorio(ruta_destino)
    
    facturas = procesar_archivo_xlsx(ruta_a_xlsx,columna_excel)
    
    if not facturas:
        print(f"❌ No se encontraron números de factura en el archivo .xlsx")
        return
    
    print(f"📊 Total de facturas a procesar: {len(facturas)}")
    print("⏳ Iniciando procesamiento...\n")
    
    # pbar = tqdm(
    #     total=len(facturas),
    #     desc="Procesando facturas",
    #     bar_format='{l_bar}{bar:40}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}]',
    #     ncols=100
    # )
    
    facturas_fallidas = 0
    facturas_exitosas = 0
    for factura in facturas:
        
        # pbar.set_description(f"Processing FE{factura}")
        
        try:
            #validar existencia de la carpeta en soportes
            carpeta_soporte = procesar_soportes(factura,ruta_soportes)
            if not carpeta_soporte:
                facturas_fallidas += 1
                # pbar.update(1)
                continue
            
            if not copiar_contenido_carpeta_soporte(carpeta_soporte, ruta_soportes_destino, factura):
                facturas_fallidas += 1
                #pbar.update(1)
                continue
            
            process = procesar_factura(ruta_facturas, factura,ruta_facturas_destino)
                
            print(process)
            
            
            
        except Exception as e:
            print(f"❌ Error inesperado procesando factura {factura}: {e}")
            facturas_fallidas += 1
        
        # finally:
            # pbar.update(1)
            
    # pbar.close()

if __name__ == '__main__':
    main()
