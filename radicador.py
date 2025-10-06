import os
from typing import List, Optional, Tuple
import pandas as pd
from tqdm import tqdm
import shutil
import glob
import logging
from pathlib import Path

# Configuración de logging con encoding UTF-8
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('procesamiento_facturas.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class ProcesadorFacturas:
    """Clase para procesar facturas y sus soportes"""
    
    def __init__(self, ruta_xlsx: str, columna: str, ruta_destino: str, 
                 ruta_soportes: str, ruta_facturas: str):
        self.ruta_xlsx = Path(ruta_xlsx)
        self.columna = columna
        self.ruta_destino = Path(ruta_destino)
        self.ruta_soportes = Path(ruta_soportes)
        self.ruta_facturas = Path(ruta_facturas)
        
        self.ruta_soportes_destino = self.ruta_destino / "Soportes"
        self.ruta_facturas_destino = self.ruta_destino / "Facturas"
        
        self.facturas_exitosas = 0
        self.facturas_fallidas = 0
        self.errores_detallados = []
    
    def validar_configuracion_inicial(self) -> bool:
        """Valida que todas las rutas de origen existan"""
        if not self.ruta_xlsx.exists():
            logger.error(f"❌ El archivo Excel no existe: {self.ruta_xlsx}")
            return False
        
        if not self.ruta_soportes.exists():
            logger.error(f"❌ La carpeta de soportes no existe: {self.ruta_soportes}")
            return False
        
        if not self.ruta_facturas.exists():
            logger.error(f"❌ La carpeta de facturas no existe: {self.ruta_facturas}")
            return False
        
        return True
    
    def crear_directorios_destino(self) -> None:
        """Crea los directorios de destino si no existen"""
        for directorio in [self.ruta_destino, self.ruta_soportes_destino, 
                          self.ruta_facturas_destino]:
            directorio.mkdir(parents=True, exist_ok=True)
            logger.info(f"📦 Directorio preparado: {directorio}")
    
    def cargar_facturas(self) -> List[str]:
        """Carga los números de factura desde el archivo Excel"""
        try:
            df = pd.read_excel(self.ruta_xlsx)
            
            if self.columna not in df.columns:
                logger.error(f"❌ La columna '{self.columna}' no existe. Columnas disponibles: {list(df.columns)}")
                return []
            
            facturas = [str(factura).strip() for factura in df[self.columna].dropna().astype(str)]
            logger.info(f"📊 Cargadas {len(facturas)} facturas desde el Excel")
            return facturas
            
        except Exception as e:
            logger.error(f"❌ Error al leer el archivo Excel: {e}")
            return []
    
    def procesar_soporte(self, factura: str) -> Optional[Path]:
        """Procesa y copia los soportes de una factura"""
        carpeta_origen = self.ruta_soportes / f"FE{factura}"
        
        if not carpeta_origen.exists():
            logger.warning(f"📁 ❌ Carpeta de soporte no encontrada: FE{factura}")
            return None
        
        carpeta_destino = self.ruta_soportes_destino / f"FE{factura}"
        
        try:
            shutil.copytree(carpeta_origen, carpeta_destino, dirs_exist_ok=True)
            logger.debug(f"✅ Soportes copiados: FE{factura}")
            return carpeta_destino
        except Exception as e:
            logger.error(f"❌ Error al copiar soportes de FE{factura}: {e}")
            return None
    
    def procesar_factura(self, factura: str) -> Optional[Path]:
        """Procesa y copia una factura"""
        carpeta_origen = self.ruta_facturas / f"AttachedDocument_F-010-{factura}"
        
        if not carpeta_origen.exists():
            logger.warning(f"📄 ❌ Carpeta de factura no encontrada: {carpeta_origen.name}")
            return None
        
        carpeta_destino = self.ruta_facturas_destino / f"FE{factura}"
        
        try:
            shutil.copytree(carpeta_origen, carpeta_destino, dirs_exist_ok=True)
            logger.debug(f"✅ Factura copiada: FE{factura}")
            return carpeta_destino
        except Exception as e:
            logger.error(f"❌ Error al copiar factura FE{factura}: {e}")
            return None
    
    def procesar_archivos_factura(self, ruta_factura_destino: Path, factura: str) -> None:
        """Renombra archivos CUV y elimina archivos innecesarios"""
        
        # Renombrar archivos CUV de .txt a .json
        patron_cuv = f"ResultadosMSPS_FE{factura}_*_A_CUV.txt"
        archivos_cuv = list(ruta_factura_destino.glob(patron_cuv))
        
        for archivo in archivos_cuv:
            archivo_json = archivo.with_suffix('.json')
            try:
                # Si el archivo JSON ya existe, eliminarlo primero
                if archivo_json.exists():
                    archivo_json.unlink()
                    logger.debug(f"Archivo JSON existente eliminado: {archivo_json.name}")
                
                archivo.rename(archivo_json)
                logger.debug(f"Renombrado: {archivo.name} -> {archivo_json.name}")
            except Exception as e:
                logger.warning(f"No se pudo renombrar {archivo.name}: {e}")
        
        # Eliminar archivos de rechazo
        patron_rechazo = f"ResultadosMSPS_FE{factura}_ID0_R.txt"
        archivos_rechazo = list(ruta_factura_destino.glob(patron_rechazo))
        
        for archivo in archivos_rechazo:
            try:
                archivo.unlink()
                logger.debug(f"Eliminado archivo de rechazo: {archivo.name}")
            except Exception as e:
                logger.warning(f"No se pudo eliminar {archivo.name}: {e}")
        
        # Eliminar archivos locales
        patron_locales = f"ResultadosLocales_FE{factura}.txt"
        archivos_locales = list(ruta_factura_destino.glob(patron_locales))
        
        for archivo in archivos_locales:
            try:
                archivo.unlink()
                logger.debug(f"Eliminado archivo local: {archivo.name}")
            except Exception as e:
                logger.warning(f"No se pudo eliminar {archivo.name}: {e}")
    
    def procesar_una_factura(self, factura: str) -> bool:
        """Procesa una factura completa (soportes + factura + archivos)"""
        try:
            # Procesar soportes
            if not self.procesar_soporte(factura):
                self.errores_detallados.append(f"FE{factura}: Soportes no encontrados")
                return False
            
            # Procesar factura
            ruta_factura_destino = self.procesar_factura(factura)
            if not ruta_factura_destino:
                self.errores_detallados.append(f"FE{factura}: Factura no encontrada")
                return False
            
            # Procesar archivos
            self.procesar_archivos_factura(ruta_factura_destino, factura)
            
            return True
            
        except Exception as e:
            logger.error(f"❌ Error inesperado procesando FE{factura}: {e}")
            self.errores_detallados.append(f"FE{factura}: {str(e)}")
            return False
    
    def procesar_todas(self) -> Tuple[int, int]:
        """Procesa todas las facturas y retorna (exitosas, fallidas)"""
        
        # Validar configuración
        if not self.validar_configuracion_inicial():
            return 0, 0
        
        # Crear directorios
        self.crear_directorios_destino()
        
        # Cargar facturas
        facturas = self.cargar_facturas()
        if not facturas:
            logger.error("❌ No se pudieron cargar facturas del Excel")
            return 0, 0
        
        logger.info(f"⏳ Iniciando procesamiento de {len(facturas)} facturas...\n")
        
        # Procesar con barra de progreso
        with tqdm(total=len(facturas), desc="Procesando facturas", 
                 bar_format='{l_bar}{bar:40}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}]',
                 ncols=100, colour="green") as pbar:
            
            for factura in facturas:
                pbar.set_description(f"Procesando FE{factura}")
                
                if self.procesar_una_factura(factura):
                    self.facturas_exitosas += 1
                else:
                    self.facturas_fallidas += 1
                
                pbar.update(1)
        
        return self.facturas_exitosas, self.facturas_fallidas
    
    def generar_reporte(self) -> None:
        """Genera un reporte del procesamiento"""
        total = self.facturas_exitosas + self.facturas_fallidas
        
        print("\n" + "="*60)
        print("📊 REPORTE DE PROCESAMIENTO")
        print("="*60)
        print(f"✅ Facturas exitosas: {self.facturas_exitosas}/{total}")
        print(f"❌ Facturas fallidas: {self.facturas_fallidas}/{total}")
        
        if self.facturas_exitosas > 0:
            porcentaje = (self.facturas_exitosas / total) * 100
            print(f"📈 Tasa de éxito: {porcentaje:.1f}%")
        
        if self.errores_detallados:
            print(f"\n⚠️ Detalles de errores ({len(self.errores_detallados)}):")
            for error in self.errores_detallados[:10]:  # Mostrar máximo 10
                print(f"  • {error}")
            
            if len(self.errores_detallados) > 10:
                print(f"  ... y {len(self.errores_detallados) - 10} errores más")
        
        print("="*60)


def solicitar_ruta(mensaje: str, debe_existir: bool = False) -> str:
    """Solicita una ruta al usuario con validación"""
    while True:
        ruta = input(f"{mensaje}: ").strip()
        
        if not ruta:
            print("⚠️ La ruta no puede estar vacía")
            continue
        
        if debe_existir and not os.path.exists(ruta):
            print(f"❌ La ruta no existe: {ruta}")
            continuar = input("¿Desea intentar con otra ruta? (s/n): ").lower()
            if continuar != 's':
                return ""
            continue
        
        return ruta


def main():
    """Función principal para procesar facturas y soportes"""
    print("="*60)
    print("🚀 PROCESADOR DE FACTURAS Y SOPORTES")
    print("="*60 + "\n")
    
    try:
        # Solicitar rutas
        ruta_xlsx = solicitar_ruta("📄 Ingrese la ruta del archivo Excel (xlsx)", debe_existir=True)
        if not ruta_xlsx:
            return
        
        columna_excel = input("📋 Ingrese el nombre de la columna con los números de factura: ").strip()
        
        ruta_soportes = solicitar_ruta("📦 Ingrese la ruta de los soportes", debe_existir=True)
        if not ruta_soportes:
            return
        
        ruta_facturas = solicitar_ruta("🗄️ Ingrese la ruta de las facturas", debe_existir=True)
        if not ruta_facturas:
            return
        
        ruta_destino = solicitar_ruta("📁 Ingrese la ruta de destino")
        
        # Crear procesador y ejecutar
        procesador = ProcesadorFacturas(
            ruta_xlsx=ruta_xlsx,
            columna=columna_excel,
            ruta_destino=ruta_destino,
            ruta_soportes=ruta_soportes,
            ruta_facturas=ruta_facturas
        )
        
        exitosas, fallidas = procesador.procesar_todas()
        procesador.generar_reporte()
        
    except KeyboardInterrupt:
        print("\n\n⚠️ Proceso interrumpido por el usuario")
    except Exception as e:
        logger.error(f"❌ Error fatal en el programa: {e}", exc_info=True)


if __name__ == '__main__':
    main()
