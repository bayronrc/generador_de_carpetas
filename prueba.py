import os
import shutil


def organizar_carpetas(ruta_base):
    # Crear las carpetas destino si no existen
    validados_dir = os.path.join(ruta_base, "validados y aprobados")
    rechazados_locales_dir = os.path.join(ruta_base, "Rechazados Locales")
    rechazados_msps_dir = os.path.join(ruta_base, "Rechazados MSPS")

    os.makedirs(validados_dir, exist_ok=True)
    os.makedirs(rechazados_locales_dir, exist_ok=True)
    os.makedirs(rechazados_msps_dir, exist_ok=True)

    # Iterar sobre todas las carpetas en el directorio base
    for carpeta in os.listdir(ruta_base):
        carpeta_path = os.path.join(ruta_base, carpeta)

        # Verificar que sea un directorio
        if not os.path.isdir(carpeta_path) or carpeta in ["validados y aprobados",
                                                          "Rechazados Locales",
                                                          "Rechazados MSPS"]:
            continue

        # Obtener lista de archivos en la carpeta
        archivos = os.listdir(carpeta_path)

        # 1. Verificar si tiene archivo ResultadosMSPS_FE######_ID??????_A_CUV.txt
        if any(f.startswith('ResultadosMSPS_FE') and f.endswith('_A_CUV.txt') for f in archivos):
            shutil.move(carpeta_path, os.path.join(validados_dir, carpeta))
            print(f"Moviendo {carpeta} a 'validados y aprobados'")
            continue

        # Contar tipos de archivos específicos
        tiene_xml = any(f.endswith('.xml') for f in archivos)
        tiene_json = any(f.endswith('.json') for f in archivos)
        tiene_locales = any(f.startswith('ResultadosLocales_FE') and f.endswith('.txt') for f in archivos)
        tiene_msps_id0 = any(f.startswith('ResultadosMSPS_FE') and f.endswith('_ID0_R.txt') for f in archivos)

        # 2. Verificar si solo tiene xml, json y ResultadosLocales
        if (tiene_xml and tiene_json and tiene_locales and
                len(archivos) == 3 and not tiene_msps_id0):
            shutil.move(carpeta_path, os.path.join(rechazados_locales_dir, carpeta))
            print(f"Moviendo {carpeta} a 'Rechazados Locales'")
            continue

        # 3. Verificar si tiene los 4 archivos específicos
        if (tiene_xml and tiene_json and tiene_locales and tiene_msps_id0 and
                len(archivos) == 4):
            shutil.move(carpeta_path, os.path.join(rechazados_msps_dir, carpeta))
            print(f"Moviendo {carpeta} a 'Rechazados MSPS'")
            continue


def main():
    # Puedes cambiar esta ruta por la que desees procesar
    ruta_base = input("Ingrese la ruta del directorio base: ")
    if os.path.isdir(ruta_base):
        organizar_carpetas(ruta_base)
        print("Organización completada!")
    else:
        print("La ruta especificada no es un directorio válido")


if __name__ == "__main__":
    main()

# Este script organiza las carpetas y archivos según las reglas especificadas.(organiza archivos en carpetas)