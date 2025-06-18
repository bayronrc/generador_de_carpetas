import os
import pandas as pd
import json

def main():
    ruta_excel = input("Ingrese la ruta del archivo Excel:")

    if not validar_ruta(ruta_excel):
        print("❌ La ruta del archivo Excel no existe.")
        return
    
    xls = pd.ExcelFile(ruta_excel)

    factura = {}
    for sheet_name in xls.sheet_names:
        if 'transaccion' in sheet_name:
            df = pd.read_excel(xls, sheet_name)
            first_line = df.iloc[0]
            
            factura = {columna : None if pd.isna(valor) else str(valor) for columna, valor in first_line.items()}

        if 'usuarios' in sheet_name:
            factura['usuarios'] = []
            df = pd.read_excel(xls, sheet_name)
            lines = df.values.tolist()
            user = {}
            for line in lines:
                    user = {
                        'tipoDocumentoIdentificacion': str(line[0]),
                        'numDocumentoIdentificacion': str(line[1]),
                        'tipoUsuario': str(line[3]),
                        'fechaNacimiento': str(line[4]),
                        'codSexo': str(line[5]),
                        'codPaisResidencia': str(line[6]),
                        'codMunicipioResidencia': str(line[7]),
                        'codZonaTerritorialResidencia': str(line[8]),
                        'incapacidad': str(line[9]),
                        'codPaisOrigen': str(line[10]),
                        'consecutivo': int(line[11])
                    }
                    print(user)
            factura['usuarios'].append({column : None if pd.isna(valor) else int(valor) for column, valor in user.items()})

    with open(f'{next(pd.read_excel(xls, sheet_name).iloc[0]['numFactura'] for sheet_name in xls.sheet_names if "transaccion" in sheet_name)}.json', 'w') as f:
        json.dump(factura, f, indent=4)

def validar_ruta(ruta):
    return os.path.isfile(ruta) and os.path.exists(ruta)

if __name__ == "__main__":
    main()
