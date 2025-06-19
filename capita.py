import os
import pandas as pd
import json

def main():
    # Solicitar la ruta del archivo Excel
    ruta_excel = input("Ingrese la ruta del archivo Excel: ")

    # Validar si la ruta existe
    if not validar_ruta(ruta_excel):
        print("❌ La ruta del archivo Excel no existe.")
        return
    
    # Cargar el archivo Excel
    xls = pd.ExcelFile(ruta_excel)
    
    # Inicializar el diccionario factura
    factura = {
        "numDocumentoIdObligado": "",
        "numFactura": "",
        "tipoNota": None,
        "numNota": None,
        "usuarios": []
    }
    
    # Diccionarios para almacenar servicios por consecutivo
    consultas_por_consecutivo = {}
    procedimientos_por_consecutivo = {}

    # Procesar cada hoja del Excel
    for sheet_name in xls.sheet_names:
        sheet_name_lower = sheet_name.lower()
        
        # Procesar la hoja de transacciones
        if 'transaccion' in sheet_name_lower:
            df = pd.read_excel(xls, sheet_name, dtype=str)
            df = df.fillna('')  # Reemplazar NaN con cadena vacía
            if not df.empty:
                primera_fila = df.iloc[0]
                factura["numDocumentoIdObligado"] = str(primera_fila.get('numDocumentoIdObligado', ''))
                factura["numFactura"] = str(primera_fila.get('numFactura', ''))
                factura["tipoNota"] = None if primera_fila.get('tipoNota', '') == '' else str(primera_fila.get('tipoNota'))
                factura["numNota"] = None if primera_fila.get('numNota', '') == '' else str(primera_fila.get('numNota'))
        
        # Procesar la hoja de usuarios
        elif 'usuarios' in sheet_name_lower:
            df = pd.read_excel(xls, sheet_name, dtype=str)
            df = df.fillna('')  # Reemplazar NaN con cadena vacía
            for _, fila in df.iterrows():
                usuario = {
                    'tipoDocumentoIdentificacion': str(fila.get('tipoDocumentoIdentificacion', '')),
                    'numDocumentoIdentificacion': str(fila.get('numDocumentoIdentificacion', '')),
                    'tipoUsuario': str(fila.get('tipoUsuario', '')).zfill(2),
                    'fechaNacimiento': str(fila.get('fechaNacimiento', '')),
                    'codSexo': str(fila.get('codSexo', '')),
                    'codPaisResidencia': str(fila.get('codPaisResidencia', '')),
                    'codMunicipioResidencia': str(fila.get('codMunicipioResidencia', '')),
                    'codZonaTerritorialResidencia': str(fila.get('codZonaTerritorialResidencia', '')),
                    'incapacidad': str(fila.get('incapacidad', '')),
                    'codPaisOrigen': str(fila.get('codPaisOrigen', '')),
                    'consecutivo': int(fila.get('consecutivo', 0)) if str(fila.get('consecutivo', '')).isdigit() else 0
                    # No incluir 'servicios' aquí, se agrega dinámicamente después
                }
                factura['usuarios'].append(usuario)
        
        # Procesar la hoja de consultas
        elif 'consultas' in sheet_name_lower:
            df = pd.read_excel(xls, sheet_name, dtype=str)
            df = df.fillna('')  # Reemplazar NaN con cadena vacía
            for i, (_, fila) in enumerate(df.iterrows(), 1):  # Iniciar el índice en 1
                consecutivo = int(fila.get('consecutivoUsuario', 0)) if str(fila.get('consecutivoUsuario', '')).isdigit() else 0
                consulta = {
                    'codPrestador': str(fila.get('codPrestador', '')),
                    'fechaInicioAtencion': str(fila.get('fechaInicioAtencion', '')),
                    'numAutorizacion': None if fila.get('numAutorizacion', '') == '' else str(fila.get('numAutorizacion')),
                    'codConsulta': str(fila.get('codConsulta', '')),
                    'modalidadGrupoServicioTecSal': str(fila.get('modalidadGrupoServicioTecSal', '')),
                    'grupoServicios': str(fila.get('grupoServicios', '')),
                    'codServicio': int(fila.get('codServicio', 0)) if str(fila.get('codServicio', '')).isdigit() else 0,
                    'finalidadTecnologiaSalud': str(fila.get('finalidadTecnologiaSalud', '')),
                    'causaMotivoAtencion': str(fila.get('causaMotivoAtencion', '')),
                    'codDiagnosticoPrincipal': str(fila.get('codDiagnosticoPrincipal', '')),
                    'codDiagnosticoRelacionado1': None if fila.get('codDiagnosticoRelacionado1', '') == '' else str(fila.get('codDiagnosticoRelacionado1')),
                    'codDiagnosticoRelacionado2': None if fila.get('codDiagnosticoRelacionado2', '') == '' else str(fila.get('codDiagnosticoRelacionado2')),
                    'codDiagnosticoRelacionado3': None if fila.get('codDiagnosticoRelacionado3', '') == '' else str(fila.get('codDiagnosticoRelacionado3')),
                    'tipoDiagnosticoPrincipal': str(fila.get('tipoDiagnosticoPrincipal', '')),
                    'tipoDocumentoIdentificacion': str(fila.get('tipoDocumentoIdentificacion', '')),
                    'numDocumentoIdentificacion': str(fila.get('numDocumentoIdentificacion', '')),
                    'vrServicio': int(fila.get('vrServicio', 0)) if str(fila.get('vrServicio', '')).replace('.', '').isdigit() else 0,
                    'conceptoRecaudo': str(fila.get('conceptoRecaudo', '')),
                    'valorPagoModerador': int(fila.get('valorPagoModerador', 0)) if str(fila.get('valorPagoModerador', '')).replace('.', '').isdigit() else 0,
                    'numFEVPagoModerador': str(fila.get('numFEVPagoModerador', '')),
                    'consecutivo': i  # Usar el índice del arreglo como consecutivo
                }
                if consecutivo not in consultas_por_consecutivo:
                    consultas_por_consecutivo[consecutivo] = []
                consultas_por_consecutivo[consecutivo].append(consulta)
        
        # Procesar la hoja de procedimientos
        elif 'procedimientos' in sheet_name_lower:
            df = pd.read_excel(xls, sheet_name, dtype=str).fillna('')

            r = df['consecutivoUsuario'].fillna(0)
            print(r)

            # for i, (_, fila) in enumerate(df.iterrows(), 1):  # Iniciar el índice en 1
            #     consecutivo = int(fila.get('consecutivoUsuario', 0)) if str(fila.get('consecutivoUsuario', '')).isdigit() else 0
            #     procedimiento = {
            #         'codPrestador': str(fila.get('codPrestador', '')),
            #         'fechaInicioAtencion': str(fila.get('fechaInicioAtencion', '')),
            #         'idMIPRES': None if fila.get('idMIPRES', '') == '' else str(fila.get('idMIPRES')),
            #         'numAutorizacion': None if fila.get('numAutorizacion', '') == '' else str(fila.get('numAutorizacion')),
            #         'codProcedimiento': str(fila.get('codProcedimiento', '')),
            #         'viaIngresoServicioSalud': str(fila.get('viaIngresoServicioSalud', '')),
            #         'modalidadGrupoServicioTecSal': str(fila.get('modalidadGrupoServicioTecSal', '')),
            #         'grupoServicios': str(fila.get('grupoServicios', '')),
            #         'codServicio': int(fila.get('codServicio', 0)) if str(fila.get('codServicio', '')).isdigit() else 0,
            #         'finalidadTecnologiaSalud': str(fila.get('finalidadTecnologiaSalud', '')),
            #         'tipoDocumentoIdentificacion': str(fila.get('tipoDocumentoIdentificacion', '')),
            #         'numDocumentoIdentificacion': str(fila.get('numDocumentoIdentificacion', '')),
            #         'codDiagnosticoPrincipal': str(fila.get('codDiagnosticoPrincipal', '')),
            #         'codDiagnosticoRelacionado': None if fila.get('codDiagnosticoRelacionado', '') == '' else str(fila.get('codDiagnosticoRelacionado')),
            #         'codComplicacion': None if fila.get('codComplicacion', '') == '' else str(fila.get('codComplicacion')),
            #         'vrServicio': int(fila.get('vrServicio', 0)) if str(fila.get('vrServicio', '')).replace('.', '').isdigit() else 0,
            #         'conceptoRecaudo': str(fila.get('conceptoRecaudo', '')),
            #         'valorPagoModerador': int(fila.get('valorPagoModerador', 0)) if str(fila.get('valorPagoModerador', '')).replace('.', '').isdigit() else 0,
            #         'numFEVPagoModerador': str(fila.get('numFEVPagoModerador', '')),
            #         'consecutivo': i  # Usar el índice del arreglo como consecutivo
            #     }
            #         procedimientos_por_consecutivo[consecutivo] = []
            #     procedimientos_por_consecutivo[consecutivo].append(procedimiento)


    print(json.dumps(factura, indent=4))
    
    # # Vincular servicios a los usuarios y filtrar aquellos sin servicios
    # usuarios_filtrados = []
    # for usuario in factura['usuarios']:
    #     consecutivo = usuario['consecutivo']
    #     # Obtener consultas y procedimientos para este usuario
    #     consultas = consultas_por_consecutivo.get(consecutivo, [])
    #     procedimientos = procedimientos_por_consecutivo.get(consecutivo, [])
        
    #     # Crear el objeto servicios dinámicamente
    #     servicios = {}
    #     if consultas:
    #         servicios['consultas'] = consultas
    #     if procedimientos:
    #         servicios['procedimientos'] = procedimientos
        
    #     # Incluir solo usuarios con servicios no vacíos
    #     if servicios:
    #         usuario['servicios'] = servicios
    #         usuarios_filtrados.append(usuario)
    
    # # Actualizar la lista de usuarios con los filtrados
    # factura['usuarios'] = usuarios_filtrados
    
    # # Generar el nombre del archivo JSON de salida
    # output_file = f"{factura['numFactura']}.json"
    
    # # Guardar en JSON
    # with open(output_file, 'w', encoding='utf-8') as f:
    #     json.dump(factura, f, indent=4, ensure_ascii=False)
    
    # print(f"✅ Archivo JSON creado exitosamente en {output_file}")

def validar_ruta(ruta):
    # Validar si la ruta es un archivo existente
    return os.path.isfile(ruta) and os.path.exists(ruta)

if __name__ == "__main__":
    main()