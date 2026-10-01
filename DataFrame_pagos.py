import pandas as pd
import openpyxl
import numpy as np
 
#lectura de excels
pagos_data_base = pd.read_excel("reporte_revision_pbs Julio 2026.xlsm",sheet_name = None, engine="openpyxl",dtype={
        "CEDULA": str,
        "NUMERO_SOLICITUD": str
    })
 
#Data frame de altas
pagos_altas = pd.read_excel("Altas sobrevivencia Julio.xlsx", sheet_name = "Altas", engine="openpyxl", dtype={
        "CEDULA_AFILIADO": str,
        "NUMERO_SOLICITUD": str
    })
 
#Data frame de altas
pagos_nomina = pd.read_excel("Altas sobrevivencia Julio.xlsx", sheet_name = "NOMINA", engine="openpyxl", dtype={
        "CEDULA_AFILIADO": str,
        "NUMERO_SOLICITUD": str
    })
 
 
#DataFrame de la base de datos de pagos
list_df = []
 
for nombre_hoja, df in pagos_data_base.items():
    df["Hoja_origen"] = nombre_hoja
    list_df.append(df)
 
pagos_data_base = pd.concat(
    list_df,
    ignore_index= True
)

#Cortar el nombre de beneficiario en altas
# Crear nombre completo del beneficiario en la base
pagos_data_base["NOMBRE_BENEFICIARIO"] = (
    pagos_data_base[
        [
            "BEN_PRIMER_NOMBRE",
            "BEN_SEGUNDO_NOMBRE",
            "BEN_PRIMER_APELLIDO",
            "BEN_SEGUNDO_APELLIDO"
        ]
    ]
    .fillna("")
    .astype(str)
    .agg(" ".join, axis=1)
    .str.upper()
    .str.strip()
    .str.replace(r"\s+", " ", regex=True)
)


# Normalizar el nombre que viene de Altas
pagos_altas["NOMBRE_BENEFICIARIO"] = (
    pagos_altas["NOMBRE_BENEFICIARIO"]
    .fillna("")
    .astype(str)
    .str.upper()
    .str.strip()
    .str.replace(r"\s+", " ", regex=True)
)

# Normalizar el nombre que viene de Nomina
pagos_nomina["NOMBRE_BENEFICIARIO"] = (
    pagos_nomina["NOMBRE_BENEFICIARIO"]
    .fillna("")
    .astype(str)
    .str.upper()
    .str.strip()
    .str.replace(r"\s+", " ", regex=True)
)


 
#Series de cedulas de la base de datos
cedula_nomina = pagos_nomina["CEDULA_AFILIADO"].drop_duplicates()
 
#Series de cedulas de altas
cedula_altas = pagos_altas["CEDULA_AFILIADO"]
 
 
 
#Creacion de dataframe de altas que esten en la base de datos
#----------------------------------------------------------------------------------------
 
#verificar si pagos de altas estan en la base de datos
verificar_pendientes= cedula_altas.isin(cedula_nomina)
 
#-----------------------------------------------------------------------------------------

#seperar las que estan en la base de datos de las que no estan 
pagos_no_pendientes = cedula_altas[verificar_pendientes]
pagos_pendientes = cedula_altas[~verificar_pendientes]
 
#Creación de dataframe con la información de las altas que no están pendientes
columnas_altas = [
    "CEDULA_AFILIADO",
    "NUMERO_SOLICITUD",
    "NOMBRE_BENEFICIARIO",
    "PRIMER PAGO",
    "MONTO PRIMER PAGO  NOMINA"
]

altas_reducida = pagos_altas.loc[
    pagos_altas["CEDULA_AFILIADO"].isin(pagos_no_pendientes),
    columnas_altas
].copy()


#crear data frame con solo las no pendientes finales 
data_base_no_pendientes = []
columnas = [
    "CEDULA",
    "NUMERO_SOLICITUD",
    "ID_BENEFICIARIO",
    "NOMBRE_BENEFICIARIO",
    "BEN_RELACION",
    "RENTA_BRUTA",
    "FECHA_FALLECIMIENTO",
    "BEN_FECHA_NACIMIENTO"
]

data_base_no_pendientes = (
    pagos_data_base.loc[
        pagos_data_base["CEDULA"].isin(pagos_no_pendientes),
        columnas
    ].copy().drop_duplicates(subset=["CEDULA", "NUMERO_SOLICITUD", "NOMBRE_BENEFICIARIO"])
)
 

pagos_no_pendientes_final = pd.merge(
    altas_reducida,
    data_base_no_pendientes,
    left_on=[
        "CEDULA_AFILIADO",
        "NUMERO_SOLICITUD",
        "NOMBRE_BENEFICIARIO"
    ],
    right_on=[
        "CEDULA",
        "NUMERO_SOLICITUD",
        "NOMBRE_BENEFICIARIO"
    ],
    how="right"
)

#Diferencias
pagos_no_pendientes_final["diferencias"] = pagos_no_pendientes_final["MONTO PRIMER PAGO  NOMINA"] - pagos_no_pendientes_final["RENTA_BRUTA"]


#Estatus
pagos_no_pendientes_final["Estatus"] = np.where(
    pagos_no_pendientes_final["diferencias"].eq(0),
    "OK",
    "Error en pago"
)



print(len(data_base_no_pendientes))
pagos_no_pendientes_final.to_excel("pagos_no_pendientes_final.xlsx", index=False)

