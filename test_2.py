# VERSION AUTOMATIZADA: clasificacion sin usar el estatus manual de Altas
# Archivo: C:\Users\pablo\Documents\pagos_auto\test_2.py

import pandas as pd
import openpyxl
import numpy as np
from difflib import SequenceMatcher
import re
import unicodedata
 
#lectura de excels
pagos_data_base = pd.read_excel("reporte_revision_pbs Julio 2026.xlsm",sheet_name = None, engine="openpyxl",dtype={
        "CEDULA": str,
        "NUMERO_SOLICITUD": str
    })
 
#Data frame de altas
pagos_altas = pd.read_excel("Altas Sobrevivencia Julio.xlsx", sheet_name = "Altas", engine="openpyxl", dtype={
        "CEDULA_AFILIADO": str,
        "NUMERO_SOLICITUD": str
    })
 
#Data frame de altas
pagos_nomina = pd.read_excel("Nomina Julio.xlsm", sheet_name = "NOMINA", engine="openpyxl", dtype={
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

# En la base, las solicitudes de Crecer conservan el prefijo historico SCOT,
# mientras que en Altas usan CREC. Se crea una clave comun solo para el cruce.
pagos_altas["SOLICITUD_CRUCE"] = (
    pagos_altas["NUMERO_SOLICITUD"]
    .astype("string")
    .str.strip()
    .str.upper()
    .str.replace(r"^SCOT", "CREC", regex=True)
)
pagos_data_base["SOLICITUD_CRUCE"] = (
    pagos_data_base["NUMERO_SOLICITUD"]
    .astype("string")
    .str.strip()
    .str.upper()
    .str.replace(r"^SCOT", "CREC", regex=True)
)
def preparar_nombre_comparable(nombre):
    texto = unicodedata.normalize("NFKD", str(nombre).upper())
    texto = "".join(
        caracter
        for caracter in texto
        if not unicodedata.combining(caracter)
    )
    palabras = re.findall(r"[A-Z0-9]+", texto)
    return "".join(palabras), "".join(sorted(palabras))


def similitud_nombres(nombre_altas, nombre_base):
    altas_directo, altas_ordenado = preparar_nombre_comparable(nombre_altas)
    base_directo, base_ordenado = preparar_nombre_comparable(nombre_base)
    return max(
        SequenceMatcher(None, altas_directo, base_directo).ratio(),
        SequenceMatcher(None, altas_ordenado, base_ordenado).ratio(),
    )



 
# Logica original: una fila queda pendiente cuando su cedula no aparece en
# Nomina. La clasificacion no consulta el estatus manual de Altas.
cedula_nomina = pagos_nomina["CEDULA_AFILIADO"].drop_duplicates()
esta_en_nomina = pagos_altas["CEDULA_AFILIADO"].isin(cedula_nomina)
verificar_pendientes = ~esta_en_nomina
pagos_pendientes = pagos_altas.loc[verificar_pendientes].copy()
pagos_no_pendientes = pagos_altas.loc[~verificar_pendientes].copy()

#print(len(pagos_no_pendientes))
#print(len(pagos_pendientes))
#pagos_pendientes.to_excel("pendientes.xlsx", index=False)

#-----------------------------------------------------------------------------------------------
#CreaciÃ³n de dataframe con la informaciÃ³n de las altas que no estÃ¡n pendientes
columnas_altas = [
    "CEDULA_AFILIADO",
    "NUMERO_SOLICITUD",
    "SOLICITUD_CRUCE",
    "NOMBRE_BENEFICIARIO",
    "PRIMER PAGO",
    "MONTO PRIMER PAGO  NOMINA"
]

altas_reducida = pagos_altas.loc[
    ~verificar_pendientes,
    columnas_altas
].copy()


#crear data frame con solo las no pendientes finales 
data_base_no_pendientes = []
columnas = [
    "CEDULA",
    "NUMERO_SOLICITUD",
    "SOLICITUD_CRUCE",
    "ID_BENEFICIARIO",
    "NOMBRE_BENEFICIARIO",
    "BEN_RELACION",
    "RENTA_BRUTA",
    "FECHA_FALLECIMIENTO",
    "BEN_FECHA_NACIMIENTO"
]

data_base_no_pendientes = (
    pagos_data_base.loc[
        pagos_data_base["CEDULA"].isin(pagos_no_pendientes["CEDULA_AFILIADO"]),
        columnas
    ]
    .copy()
    .drop_duplicates(
        subset=["CEDULA", "SOLICITUD_CRUCE", "NOMBRE_BENEFICIARIO"]
    )
    .rename(columns={"NUMERO_SOLICITUD": "NUMERO_SOLICITUD_BASE"})
)
 

pagos_no_pendientes_final = pd.merge(
    altas_reducida,
    data_base_no_pendientes,
    left_on=[
        "CEDULA_AFILIADO",
        "SOLICITUD_CRUCE",
        "NOMBRE_BENEFICIARIO"
    ],
    right_on=[
        "CEDULA",
        "SOLICITUD_CRUCE",
        "NOMBRE_BENEFICIARIO"
    ],
    # Altas controla el resultado: no se incorporan filas que existan
    # solamente en la base de datos.
    how="left",
    indicator="resultado_merge",
    validate="many_to_one"
)


# Segundo intento exclusivo para nombres con diferencias de escritura. Las
# claves de cedula y solicitud siguen siendo obligatorias para evitar cruces
# con otro afiliado o expediente.
columnas_recuperadas = [
    "CEDULA",
    "NUMERO_SOLICITUD_BASE",
    "ID_BENEFICIARIO",
    "BEN_RELACION",
    "RENTA_BRUTA",
    "FECHA_FALLECIMIENTO",
    "BEN_FECHA_NACIMIENTO",
]
pagos_no_pendientes_final["resultado_merge"] = (
    pagos_no_pendientes_final["resultado_merge"].astype(str)
)
candidatos_usados = set()

for indice, fila in pagos_no_pendientes_final.loc[
    pagos_no_pendientes_final["resultado_merge"].eq("left_only")
].iterrows():
    candidatos = data_base_no_pendientes.loc[
        data_base_no_pendientes["CEDULA"].eq(fila["CEDULA_AFILIADO"])
        & data_base_no_pendientes["SOLICITUD_CRUCE"].eq(
            fila["SOLICITUD_CRUCE"]
        )
        & ~data_base_no_pendientes.index.isin(candidatos_usados)
    ].copy()

    if candidatos.empty:
        continue

    candidatos["_similitud"] = candidatos["NOMBRE_BENEFICIARIO"].map(
        lambda nombre: similitud_nombres(fila["NOMBRE_BENEFICIARIO"], nombre)
    )
    candidatos.sort_values("_similitud", ascending=False, inplace=True)
    mejor_indice = candidatos.index[0]
    mejor_puntaje = candidatos.iloc[0]["_similitud"]
    segundo_puntaje = (
        candidatos.iloc[1]["_similitud"] if len(candidatos) > 1 else 0
    )

    coincidencia_confiable = (
        mejor_puntaje >= 0.80
        and (mejor_puntaje >= 0.95 or mejor_puntaje - segundo_puntaje >= 0.08)
    )
    if not coincidencia_confiable:
        continue

    mejor_candidato = candidatos.loc[mejor_indice]
    for columna in columnas_recuperadas:
        pagos_no_pendientes_final.at[indice, columna] = mejor_candidato[columna]
    pagos_no_pendientes_final.at[
        indice, "resultado_merge"
    ] = "coincidencia_aproximada"
    candidatos_usados.add(mejor_indice)

#Diferencias
monto_nomina = pd.to_numeric(
    pagos_no_pendientes_final["MONTO PRIMER PAGO  NOMINA"],
    errors="coerce"
)
renta_bruta = pd.to_numeric(
    pagos_no_pendientes_final["RENTA_BRUTA"],
    errors="coerce"
)

pagos_no_pendientes_final["diferencias"] = (
    monto_nomina - renta_bruta
).round(2)


# Estatus: una renta ausente no se debe presentar como un error de pago.
datos_completos = monto_nomina.notna() & renta_bruta.notna()
pagos_no_pendientes_final["Estatus"] = np.select(
    [
        ~datos_completos,
        pagos_no_pendientes_final["diferencias"].abs().le(0.01),
    ],
    [
        "Sin coincidencia en base",
        "OK",
    ],
    default="Error en pago"
)

pagos_no_pendientes_final.drop(columns="SOLICITUD_CRUCE", inplace=True)

pagos_no_pendientes_tabla_final = pagos_no_pendientes_final[
    [
        "CEDULA_AFILIADO",
        "NUMERO_SOLICITUD",
        "NOMBRE_BENEFICIARIO",
        "PRIMER PAGO",
        "MONTO PRIMER PAGO  NOMINA",
        "ID_BENEFICIARIO",
        "BEN_RELACION",
        "RENTA_BRUTA",
        "FECHA_FALLECIMIENTO",
        "BEN_FECHA_NACIMIENTO",
        "diferencias",
        "Estatus"
    ]
]

# Los pagos correctos no requieren revision. La salida conserva solamente
# errores, casos sin coincidencia y el archivo separado de pendientes.
pagos_no_pendientes_tabla_final = pagos_no_pendientes_tabla_final.loc[
    pagos_no_pendientes_tabla_final["Estatus"].ne("OK")
].copy()

pagos_pendientes.to_excel("pendientes.xlsx", index=False)
pagos_no_pendientes_tabla_final.to_excel("pagos_no_pendientes.xlsx", index=False)
