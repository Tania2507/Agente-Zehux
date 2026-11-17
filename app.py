import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import calendar

# ==================================
# CONFIGURACION
# ==================================

st.set_page_config(
    page_title="Generador Control Gestion",
    layout="wide"
)

st.title("📋 Generador Control Gestión Nuevas y Renovaciones")

st.markdown("""
Cargue el archivo de Zehux y el sistema generará automáticamente
el archivo de control con las reglas definidas.
""")

archivo = st.file_uploader(
    "Seleccione el archivo Excel",
    type=["xlsx"]
)

# ==================================
# FUNCIONES
# ==================================

MESES = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre"
}


def obtener_area(fila):

    area = str(fila.get("Area", "")).strip()

    ini_vigencia = fila.get("Ini Vigencia")

    try:
        ini_vigencia = pd.to_datetime(
            ini_vigencia,
            errors="coerce"
        )
    except:
        ini_vigencia = pd.NaT

    if area.lower() == "facturacion vencida":

        if pd.notnull(ini_vigencia):

            siguiente_mes = (
                ini_vigencia.month % 12
            ) + 1

            return (
                "Facturacion Vencida - "
                + MESES[siguiente_mes]
            )

    if area.lower() == "vigencia futura":

        if pd.notnull(ini_vigencia):

            return (
                "Vigencia Futura - "
                + MESES[
                    ini_vigencia.month
                ]
            )

    if area.lower() == "sucursal":
        return "Comercial"

    return area


def obtener_responsable(fila):

    area = str(
        fila.get("Area", "")
    ).strip()

    if area.lower() == "sucursal":
        return "Comercial"

    return fila.get(
        "Usuario",
        ""
    )


def calcular_estado_ans(
    dias,
    area
):

    area_texto = str(area)

    estados_seguimiento = [

        "Pendiente Blancket",

        "Pendiente Ajuste Blancket",

        "Pendiente Ajuste Blancket – Pre-cobro",

        "Sistemas",

        "Sistemas Lisa",

        "Sistemas Listas Restrictivas"

    ]

    if any(
        x.lower() in area_texto.lower()
        for x in estados_seguimiento
    ):
        return (
            "Fuera de ANS - Seguimiento"
        )

    if dias <= 3:
        return "Dentro de ANS"

    elif dias == 4:
        return "Vence Hoy"

    return "Fuera de ANS"


def obtener_area_cierre(
    fecha_cierre,
    consecutivo
):

    fecha_txt = (
        fecha_cierre
        .strftime("%d%m%Y")
    )

    return (
        f"{consecutivo:02d}. "
        f"Cerrado {fecha_txt}"
    )


# ==================================
# PROCESAMIENTO
# ==================================

if archivo:

    df = pd.read_excel(
        archivo,
        engine="openpyxl"
    )

    st.success(
        "Archivo cargado correctamente"
    )

    # ==========================
    # FECHAS
    # ==========================

    columnas_fecha = [

        "Ini Vigencia",
        "Fecha Radicado",
        "Fecha Asignacion",
        "Fecha cierre",
        "Fecha movimiento SARLAFT",
        "Fecha Ultima Devolucion"

    ]

    for col in columnas_fecha:

        if col in df.columns:

            df[col] = pd.to_datetime(
                df[col],
                errors="coerce"
            )

    # ==========================
    # DIAS OPERACION
    # ==========================

    df["Dias En Operación"] = (
        pd.to_numeric(
            df["Total DC"],
            errors="coerce"
        )
        .fillna(0)
        + 1
    )

    # ==========================
    # VENCIMIENTO RADICADO
    # ==========================

    df["Fecha Vencim. Rad"] = (

        df["Fecha Radicado"]

        + pd.to_timedelta(
            4,
            unit="D"
        )

    )

    # ==========================
    # AREA
    # ==========================

    df["Area Resultado"] = (
        df.apply(
            obtener_area,
            axis=1
        )
    )

    # ==========================
    # RESPONSABLE
    # ==========================

    df["Responsable"] = (
        df.apply(
            obtener_responsable,
            axis=1
        )
    )

    # ==========================
    # POLIZA
    # ==========================

    df["Poliza Final"] = (

        df["Producto"]
        .fillna("")
        .astype(str)

        + "-"

        + df["No. Poliza"]
        .fillna("")
        .astype(str)

    )

    # ==========================
    # ANS
    # ==========================

    df["Estado ANS"] = (

        df.apply(

            lambda x:

            calcular_estado_ans(
                x["Dias En Operación"],
                x["Area Resultado"]
            ),

            axis=1

        )

    )

    # ==========================
# CERRADOS
# ==========================

hoy = datetime.today()

cerrados = df[
    (
        df["Area"]
        .astype(str)
        .str.upper()
        .str.contains(
            "CERRADO",
            na=False
        )
    )
    &
    (
        df["Fecha cierre"].dt.month
        == hoy.month
    )
    &
    (
        df["Fecha cierre"].dt.year
        == hoy.year
    )
].copy()

if len(cerrados) > 0:

    cerrados = (
        cerrados
        .sort_values(
            "Fecha cierre"
        )
    )

    consecutivos = []

    for i, (_, fila) in enumerate(
        cerrados.iterrows(),
        start=1
    ):

        consecutivos.append(

            obtener_area_cierre(
                fila["Fecha cierre"],
                i
            )

        )

    cerrados[
        "Area Resultado"
    ] = consecutivos

    df = df[
        ~(
            df["Area"]
            .astype(str)
            .str.upper()
            .str.contains(
                "CERRADO",
                na=False
            )
        )
    ]

    df = pd.concat(
        [
            df,
            cerrados
        ],
        ignore_index=True
    )
    # ==========================
    # RESULTADO FINAL
    # ==========================

    resultado = pd.DataFrame({

        "Radicado":
            df["Radicado"],

        "Ramo":
            df["Producto"],

        "No. Póliza":
            df["Poliza Final"],

        "Tomador":
            df["Razon Social"],

        "Ini Vigencia":
            df["Ini Vigencia"],

        "Tipo de Radicado":
            df["Tipo de Radicado"],

        "Sucursal":
            df["Sucursal"],

        "Fecha Radicado":
            df["Fecha Radicado"],

        "Fecha Asignación":
            df["Fecha Asignacion"],

        "Ultima fecha OP":
            df["Fecha Ultima Devolucion"],

        "Fecha Vencim. Rad":
            df["Fecha Vencim. Rad"],

        "Dias En Operación":
            df["Dias En Operación"],

        "Estado ANS":
            df["Estado ANS"],

        "Area":
            df["Area Resultado"],

        "Responsable":
            df["Responsable"],

        "Prima Estimada":
            df["Prima Estimada"],

        "Tipo Facturación":
            df["Tipo Poliza"],

        "Usuario Sarlaft":
            df[
                "Ultimo Usuario SARLAFT"
            ],

        "Fecha Sarlaft":
            df[
                "Fecha movimiento SARLAFT"
            ]

    })

    st.subheader(
        "Vista previa"
    )

    st.dataframe(
        resultado.head(20),
        use_container_width=True
    )

    # ==========================
    # DESCARGA
    # ==========================

    salida = BytesIO()

    with pd.ExcelWriter(
        salida,
        engine="openpyxl"
    ) as writer:

        resultado.to_excel(
            writer,
            index=False,
            sheet_name="Control Gestion"
        )

    st.download_button(

        label="📥 Descargar Archivo",

        data=salida.getvalue(),

        file_name=(
            "Control_Gestion.xlsx"
        ),

        mime=(
            "application/vnd."
            "openxmlformats-"
            "officedocument."
            "spreadsheetml.sheet"
        )

    )
