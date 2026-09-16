from __future__ import annotations

import argparse
import hashlib
import re
import sqlite3
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd



# RUTAS DEL PROYECTO


PROJECT_ROOT = Path(__file__).resolve().parents[1]

DEFAULT_INPUT = (
    PROJECT_ROOT
    / "data"
    / "clientes.xlsx"
)

DB_PATH = (
    PROJECT_ROOT
    / "database"
    / "p01_taller_datos.db"
)

CORRECTED_DIR = (
    PROJECT_ROOT
    / "clientes corregidos"
)



# CONFIGURACIÓN


TITLE_COLUMNS = {
    "nombre",
    "ciudad",
    "estado_cliente",
    "canal_origen",
}

NUMERIC_COLUMNS = {
    "edad",
    "compras_12m",
    "valor_compras_12m",
}


CANONICAL = {

    "ciudad": {

        "bogota": "Bogotá",
        "medellin": "Medellín",
        "cali": "Cali",
        "barranquilla": "Barranquilla",
        "bucaramanga": "Bucaramanga",
        "pereira": "Pereira",
        "manizales": "Manizales",

    },

    "estado_cliente": {

        "activo": "Activo",
        "inactivo": "Inactivo",
        "prospecto": "Prospecto",

    },

    "canal_origen": {

        "web": "Web",
        "evento": "Evento",
        "referido": "Referido",
        "tienda": "Tienda",
        "whatsapp": "WhatsApp",

    },
}


EMAIL_RE = re.compile(
    r"^[^\s@]+@[^\s@]+\.[^\s@]+$"
)



# FUNCIONES GENERALES


def normalize_key(value: Any) -> str:

    text = str(value)

    text = text.strip()

    text = text.lower()

    text = unicodedata.normalize(
        "NFD",
        text
    )

    text = "".join(
        ch
        for ch in text
        if unicodedata.category(ch) != "Mn"
    )

    return text


def normalize_column(value: Any) -> str:

    text = normalize_key(value)

    text = re.sub(
        r"[^a-z0-9]+",
        "_",
        text
    )

    return text.strip("_")


def is_blank(value: Any) -> bool:

    if pd.isna(value):
        return True

    if isinstance(value, str):

        if value.strip() == "":
            return True

    return False


def as_text(value: Any) -> str:

    if pd.isna(value):
        return ""

    return str(value)


def file_hash(path: Path) -> str:

    digest = hashlib.sha256()

    with path.open("rb") as file:

        for chunk in iter(
            lambda: file.read(1024 * 1024),
            b"",
        ):

            digest.update(chunk)

    return digest.hexdigest()



# NÚMEROS


def parse_number(
    value: Any
) -> tuple[Any, bool]:

    if is_blank(value):
        return value, False

    if isinstance(value, bool):
        return value, False

    if isinstance(
        value,
        (int, float)
    ):

        if (
            isinstance(value, float)
            and value.is_integer()
        ):

            return int(value), True

        return value, False


    raw = str(value).strip()

    text = raw.replace(
        "$",
        ""
    )

    text = text.replace(
        " ",
        ""
    )


    # -----------------------------------------
    # Ejemplo:
    #
    # $4.447.629
    # 4.447.629
    #
    # pasa a:
    #
    # 4447629
    # -----------------------------------------

    if re.fullmatch(
        r"[-+]?\d{1,3}(?:\.\d{3})+",
        text
    ):

        number = int(
            text.replace(
                ".",
                ""
            )
        )

        return number, True


    # -----------------------------------------
    # Formato:
    #
    # 4,447,629
    # -----------------------------------------

    if re.fullmatch(
        r"[-+]?\d{1,3}(?:,\d{3})+",
        text
    ):

        number = int(
            text.replace(
                ",",
                ""
            )
        )

        return number, True


    # -----------------------------------------
    # Enteros o decimales simples
    # -----------------------------------------

    if re.fullmatch(
        r"[-+]?\d+(?:\.\d+)?",
        text
    ):

        if "." in text:

            number = float(text)

        else:

            number = int(text)


        if (
            isinstance(number, float)
            and number.is_integer()
        ):

            number = int(number)


        changed = (
            as_text(number)
            != raw
        )

        return number, changed


    return value, False



# FECHAS


def parse_date(
    value: Any
) -> str | None:

    if is_blank(value):
        return None


    if isinstance(
        value,
        (pd.Timestamp, datetime)
    ):

        return value.strftime(
            "%Y-%m-%d"
        )


    text = str(value).strip()


    accepted_formats = (

        "%Y-%m-%d",

        "%d/%m/%Y",

        "%d-%m-%Y",

        "%Y/%m/%d",

    )


    for format_date in accepted_formats:

        try:

            parsed = datetime.strptime(
                text,
                format_date
            )

            return parsed.strftime(
                "%Y-%m-%d"
            )

        except ValueError:

            pass


    return None



# CARGAR EXCEL / CSV


def load_dataset(
    path: Path
) -> pd.DataFrame:

    suffix = path.suffix.lower()


    if suffix == ".xlsx":

        df = pd.read_excel(
            path,
            engine="openpyxl"
        )


    elif suffix == ".xls":

        df = pd.read_excel(
            path
        )


    elif suffix == ".csv":

        df = pd.read_csv(
            path,
            sep=None,
            engine="python"
        )


    else:

        raise ValueError(
            "Formato no soportado. "
            "Usa CSV, XLS o XLSX."
        )


    # =====================================================
    # NORMALIZAR NOMBRES DE COLUMNAS
    # =====================================================

    df.columns = [

        normalize_column(col)

        for col in df.columns

    ]


    # =====================================================
    # QUITAR COLUMNA INDEX ARTIFICIAL
    # =====================================================

    for col in list(df.columns):

        if (
            col == "index"
            or col.startswith("unnamed")
        ):

            numeric = pd.to_numeric(
                df[col],
                errors="coerce"
            )

            if (
                len(df) == 0
                or numeric.notna().mean() >= 0.90
            ):

                df = df.drop(
                    columns=[col]
                )


    return df



# MOTOR PRINCIPAL DE LIMPIEZA


def clean_dataset(
    original: pd.DataFrame
) -> tuple[
    pd.DataFrame,
    list[dict[str, Any]],
    list[dict[str, Any]]
]:

    # object permite mezclar temporalmente
    # números con la palabra REVISAR

    df = (
        original
        .copy()
        .astype(object)
    )


    issues: list[
        dict[str, Any]
    ] = []


    changes: list[
        dict[str, Any]
    ] = []


    review_reasons: dict[
        Any,
        list[str]
    ] = {}


    # =====================================================
    # REFERENCIA DEL REGISTRO
    # =====================================================

    def record_reference(
        idx: Any
    ) -> str:

        if (
            "id_cliente"
            in df.columns
        ):

            value = df.at[
                idx,
                "id_cliente"
            ]

            if not is_blank(value):

                return as_text(
                    value
                )


        return str(
            int(idx) + 2
        )


    # =====================================================
    # APLICAR CAMBIO AUTOMÁTICO
    # =====================================================

    def change(
        idx: Any,
        col: str,
        new_value: Any,
        rule: str,
        description: str,
    ) -> None:

        old_value = df.at[
            idx,
            col
        ]


        if (
            as_text(old_value)
            == as_text(new_value)
        ):

            return


        issue = {

            "fila_excel":
                int(idx) + 2,

            "registro":
                record_reference(idx),

            "regla":
                rule,

            "columna":
                col,

            "valor_original":
                as_text(old_value),

            "valor_resultado":
                as_text(new_value),

            "tratamiento":
                "AUTOMATICA_SEGURA",

            "estado":
                "APLICADA",

            "descripcion":
                description,

        }


        issues.append(
            issue
        )


        changes.append({

            "fila_excel":
                int(idx) + 2,

            "registro":
                record_reference(idx),

            "regla":
                rule,

            "columna":
                col,

            "antes":
                as_text(old_value),

            "despues":
                as_text(new_value),

            "origen":
                "AUTOMATICO",

        })


        df.at[
            idx,
            col
        ] = new_value


    # =====================================================
    # MARCAR PARA REVISIÓN HUMANA
    # =====================================================

    def review(
        idx: Any,
        col: str,
        rule: str,
        description: str,
        replace: bool = True,
    ) -> None:

        old_value = df.at[
            idx,
            col
        ]


        if replace:

            result = "REVISAR"

        else:

            result = as_text(
                old_value
            )


        issues.append({

            "fila_excel":
                int(idx) + 2,

            "registro":
                record_reference(idx),

            "regla":
                rule,

            "columna":
                col,

            "valor_original":
                as_text(old_value),

            "valor_resultado":
                result,

            "tratamiento":
                "REVISION_HUMANA",

            "estado":
                "PENDIENTE",

            "descripcion":
                description,

        })


        review_reasons.setdefault(
            idx,
            []
        )


        review_reasons[idx].append(

            f"{col}: {description}"

        )


        if replace:

            df.at[
                idx,
                col
            ] = "REVISAR"


    # =====================================================
    # R001
    # ESPACIOS AL INICIO O FINAL
    #
    # R002
    # ESPACIOS DOBLES
    #
    # R003
    # MAYÚSCULAS / MINÚSCULAS
    # =====================================================

    for col in df.columns:

        for idx in df.index:

            value = df.at[
                idx,
                col
            ]


            if not isinstance(
                value,
                str
            ):

                continue


            if value == "":

                continue


            # ---------------------------------------------
            # Quitar espacios externos
            # ---------------------------------------------

            stripped = (
                value.strip()
            )


            if (
                stripped
                != value
            ):

                change(

                    idx,

                    col,

                    stripped,

                    "R001",

                    "Se eliminaron espacios externos.",

                )


            # ---------------------------------------------
            # Quitar espacios dobles
            # ---------------------------------------------

            current = df.at[
                idx,
                col
            ]


            if isinstance(
                current,
                str
            ):

                collapsed = re.sub(

                    r"\s{2,}",

                    " ",

                    current

                )


                if (
                    collapsed
                    != current
                ):

                    change(

                        idx,

                        col,

                        collapsed,

                        "R002",

                        "Se eliminaron espacios dobles o múltiples.",

                    )


            # ---------------------------------------------
            # Capitalización
            # ---------------------------------------------

            current = df.at[
                idx,
                col
            ]


            if (
                col
                in TITLE_COLUMNS
                and isinstance(
                    current,
                    str
                )
                and current
            ):

                titled = (
                    current.title()
                )


                if (
                    titled
                    != current
                ):

                    change(

                        idx,

                        col,

                        titled,

                        "R003",

                        "Se normalizó la capitalización.",

                    )


    # =====================================================
    # R004
    # NORMALIZAR CORREOS
    # =====================================================

    if "correo" in df.columns:

        for idx in df.index:

            value = df.at[
                idx,
                "correo"
            ]


            if (
                isinstance(
                    value,
                    str
                )
                and value
            ):

                normalized = (
                    value
                    .strip()
                    .lower()
                )


                if (
                    normalized
                    != value
                ):

                    change(

                        idx,

                        "correo",

                        normalized,

                        "R004",

                        "Correo normalizado a minúsculas.",

                    )


    # =====================================================
    # R005
    # NORMALIZAR CATEGORÍAS
    #
    # R020
    # CATEGORÍA DESCONOCIDA
    # =====================================================

    for (
        col,
        mapping
    ) in CANONICAL.items():


        if col not in df.columns:

            continue


        for idx in df.index:

            value = df.at[
                idx,
                col
            ]


            if is_blank(value):

                continue


            if value == "REVISAR":

                continue


            key = normalize_key(
                value
            )


            if key in mapping:

                canonical = mapping[
                    key
                ]


                if (
                    as_text(value)
                    != canonical
                ):

                    change(

                        idx,

                        col,

                        canonical,

                        "R005",

                        "Categoría normalizada.",

                    )


            else:

                review(

                    idx,

                    col,

                    "R020",

                    "Categoría no reconocida.",

                )


    # =====================================================
    # R006
    # CONVERTIR NÚMEROS
    #
    # R007
    # VALOR NO NUMÉRICO
    #
    # R008
    # EDAD FUERA DE RANGO
    #
    # R009
    # VALOR NEGATIVO
    # =====================================================

    numeric_columns = (
        NUMERIC_COLUMNS
        .intersection(
            df.columns
        )
    )


    for col in numeric_columns:

        for idx in df.index:

            value = df.at[
                idx,
                col
            ]


            if is_blank(value):

                continue


            if value == "REVISAR":

                continue


            parsed, changed = (
                parse_number(
                    value
                )
            )


            if (
                changed
                and not isinstance(
                    parsed,
                    str
                )
            ):

                change(

                    idx,

                    col,

                    parsed,

                    "R006",

                    "Número convertido a formato estándar.",

                )


            current = df.at[
                idx,
                col
            ]


            if current == "REVISAR":

                continue


            try:

                number = float(
                    current
                )


            except (
                TypeError,
                ValueError
            ):

                review(

                    idx,

                    col,

                    "R007",

                    "No se puede interpretar como número.",

                )

                continue


            # ---------------------------------------------
            # Edad
            # ---------------------------------------------

            if col == "edad":

                if not (
                    0
                    <= number
                    <= 120
                ):

                    review(

                        idx,

                        col,

                        "R008",

                        "Edad fuera del rango razonable de 0 a 120 años.",

                    )


            # ---------------------------------------------
            # Valores que no deberían ser negativos
            # ---------------------------------------------

            elif col in {

                "compras_12m",

                "valor_compras_12m",

            }:

                if number < 0:

                    review(

                        idx,

                        col,

                        "R009",

                        "No se admite un valor negativo en esta columna.",

                    )


    # =====================================================
    # R010
    # NORMALIZAR FECHA
    #
    # R011
    # FECHA IMPOSIBLE
    # =====================================================

    if (
        "fecha_registro"
        in df.columns
    ):

        for idx in df.index:

            value = df.at[
                idx,
                "fecha_registro"
            ]


            if is_blank(value):

                continue


            parsed = parse_date(
                value
            )


            if parsed is None:

                review(

                    idx,

                    "fecha_registro",

                    "R011",

                    "Fecha inválida o imposible.",

                )


            elif (
                as_text(value)
                != parsed
            ):

                change(

                    idx,

                    "fecha_registro",

                    parsed,

                    "R010",

                    "Fecha normalizada al formato AAAA-MM-DD.",

                )


    # =====================================================
    # R012 / R013
    # VALIDACIÓN DE CORREOS
    # =====================================================

    if "correo" in df.columns:

        for idx in df.index:

            value = df.at[
                idx,
                "correo"
            ]


            if is_blank(value):

                review(

                    idx,

                    "correo",

                    "R012",

                    "Falta el correo electrónico.",

                )


            elif (
                value != "REVISAR"
                and not EMAIL_RE.fullmatch(
                    str(value)
                    .strip()
                    .lower()
                )
            ):

                review(

                    idx,

                    "correo",

                    "R013",

                    "Correo con formato inválido.",

                )


    # =====================================================
    # R014 - R016
    # TELÉFONOS
    # =====================================================

    if "telefono" in df.columns:

        for idx in df.index:

            value = df.at[
                idx,
                "telefono"
            ]


            if is_blank(value):

                review(

                    idx,

                    "telefono",

                    "R014",

                    "Falta el teléfono.",

                )

                continue


            if value == "REVISAR":

                continue


            text = str(
                value
            ).strip()


            # Corrige casos:
            #
            # 3001234567.0
            #
            # producidos por Excel / Pandas

            if re.fullmatch(
                r"\d+\.0",
                text
            ):

                text = text[:-2]


            digits = re.sub(

                r"\D",

                "",

                text

            )


            if (
                7
                <= len(digits)
                <= 15
            ):

                if (
                    digits
                    != text
                ):

                    change(

                        idx,

                        "telefono",

                        digits,

                        "R015",

                        "Teléfono normalizado a solo dígitos.",

                    )


            else:

                review(

                    idx,

                    "telefono",

                    "R016",

                    "Teléfono con longitud o formato sospechoso.",

                )


    # =====================================================
    # R017
    # DATOS IMPORTANTES VACÍOS
    # =====================================================

    important_columns = (

        "id_cliente",

        "nombre",

        "documento",

    )


    for col in important_columns:

        if col not in df.columns:

            continue


        for idx in df.index:

            if is_blank(
                df.at[
                    idx,
                    col
                ]
            ):

                review(

                    idx,

                    col,

                    "R017",

                    "Falta un dato importante.",

                )


    # =====================================================
    # R018
    # DUPLICADOS EXACTOS
    # =====================================================

    duplicate_mask = (
        df.duplicated(
            keep="first"
        )
    )


    duplicate_indices = list(

        df.index[
            duplicate_mask
        ]

    )


    for idx in duplicate_indices:

        issues.append({

            "fila_excel":
                int(idx) + 2,

            "registro":
                record_reference(idx),

            "regla":
                "R018",

            "columna":
                "(fila completa)",

            "valor_original":
                "REGISTRO DUPLICADO",

            "valor_resultado":
                "ELIMINADO",

            "tratamiento":
                "AUTOMATICA_SEGURA",

            "estado":
                "APLICADA",

            "descripcion":
                "Fila exactamente duplicada. "
                "Se conservó la primera aparición.",

        })


        changes.append({

            "fila_excel":
                int(idx) + 2,

            "registro":
                record_reference(idx),

            "regla":
                "R018",

            "columna":
                "(fila completa)",

            "antes":
                "REGISTRO DUPLICADO",

            "despues":
                "ELIMINADO",

            "origen":
                "AUTOMATICO",

        })


    if duplicate_indices:

        df = df.loc[
            ~duplicate_mask
        ].copy()


    # =====================================================
    # R019
    # MISMO DOCUMENTO CON INFORMACIÓN DIFERENTE
    # =====================================================

    if (
        "documento"
        in df.columns
    ):

        docs = (
            df["documento"]
            .astype(str)
        )


        duplicated_docs = (
            docs.duplicated(
                keep=False
            )
        )


        duplicated_data = (
            df.loc[
                duplicated_docs
            ]
        )


        if not duplicated_data.empty:

            groups = (
                duplicated_data
                .groupby(
                    docs[
                        duplicated_docs
                    ],
                    dropna=False
                )
            )


            for _, group in groups:

                if (
                    len(
                        group.drop_duplicates()
                    )
                    > 1
                ):

                    for idx in group.index:

                        review(

                            idx,

                            "documento",

                            "R019",

                            "El mismo documento aparece "
                            "en registros con información diferente.",

                            replace=False,

                        )


    # =====================================================
    # COLUMNA FINAL DE REVISIÓN
    # =====================================================

    df[
        "revision_pendiente"
    ] = [

        " | ".join(
            review_reasons.get(
                idx,
                []
            )
        )

        for idx in df.index

    ]


    return (
        df,
        issues,
        changes
    )



# PERFIL DE COLUMNAS


def profile_dataframe(
    df: pd.DataFrame
) -> pd.DataFrame:

    rows = []


    for col in df.columns:

        series = df[col]

        non_null = (
            series.dropna()
        )


        examples = [

            as_text(value)

            for value
            in non_null
            .head(5)
            .tolist()

        ]


        rows.append({

            "columna":
                col,

            "tipo_original":
                str(series.dtype),

            "nulos_o_vacios":
                int(
                    series
                    .map(is_blank)
                    .sum()
                ),

            "valores_unicos":
                int(
                    series
                    .nunique(
                        dropna=True
                    )
                ),

            "ejemplos":
                ", ".join(
                    examples
                ),

        })


    return pd.DataFrame(
        rows
    )



# CREAR ARCHIVOS FINALES


def write_outputs(

    clean_df: pd.DataFrame,

    original_df: pd.DataFrame,

    issues: list[
        dict[str, Any]
    ],

    changes: list[
        dict[str, Any]
    ],

    input_path: Path,

) -> tuple[
    Path,
    Path
]:


    # =====================================================
    # CREAR CARPETA
    # =====================================================

    CORRECTED_DIR.mkdir(

        parents=True,

        exist_ok=True

    )


    # =====================================================
    # NOMBRES
    # =====================================================

    clean_path = (

        CORRECTED_DIR
        / "clientes_corregidos.xlsx"

    )


    report_path = (

        CORRECTED_DIR
        / "clientes_reporte_limpieza.xlsx"

    )


    # =====================================================
    # EXCEL CORREGIDO
    # =====================================================

    clean_df.to_excel(

        clean_path,

        index=False,

        engine="openpyxl"

    )


    # =====================================================
    # ESTADÍSTICAS
    # =====================================================

    automatic = sum(

        item["tratamiento"]
        == "AUTOMATICA_SEGURA"

        for item
        in issues

    )


    review_count = sum(

        item["tratamiento"]
        == "REVISION_HUMANA"

        for item
        in issues

    )


    # =====================================================
    # RESUMEN
    # =====================================================

    summary = pd.DataFrame(

        [

            [
                "Archivo original",
                input_path.name
            ],

            [
                "Registros originales",
                len(original_df)
            ],

            [
                "Registros finales",
                len(clean_df)
            ],

            [
                "Incidencias detectadas",
                len(issues)
            ],

            [
                "Correcciones automáticas",
                automatic
            ],

            [
                "Pendientes de revisión",
                review_count
            ],

            [
                "Cambios aplicados",
                len(changes)
            ],

        ],

        columns=[
            "Indicador",
            "Valor"
        ]

    )


    # =====================================================
    # REPORTE
    # =====================================================

    with pd.ExcelWriter(

        report_path,

        engine="openpyxl"

    ) as writer:


        summary.to_excel(

            writer,

            sheet_name="Resumen",

            index=False

        )


        pd.DataFrame(
            changes
        ).to_excel(

            writer,

            sheet_name="Cambios_Aplicados",

            index=False

        )


        pd.DataFrame(
            issues
        ).to_excel(

            writer,

            sheet_name="Incidencias",

            index=False

        )


        profile_dataframe(
            original_df
        ).to_excel(

            writer,

            sheet_name="Perfil_Columnas",

            index=False

        )


    return (
        clean_path,
        report_path
    )



# AUDITORÍA SQLITE


def save_audit(

    input_path: Path,

    clean_path: Path,

    report_path: Path,

    issues: list[
        dict[str, Any]
    ],

    changes: list[
        dict[str, Any]
    ],

) -> None:


    DB_PATH.parent.mkdir(

        parents=True,

        exist_ok=True

    )


    conn = sqlite3.connect(
        DB_PATH
    )


    # Se usan nombres nuevos para que no choquen
    # con las tablas de versiones anteriores.

    conn.executescript(
        """

        CREATE TABLE IF NOT EXISTS auditoria_ejecuciones (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            archivo TEXT,

            hash_sha256 TEXT,

            fecha TEXT DEFAULT CURRENT_TIMESTAMP,

            incidencias INTEGER,

            cambios INTEGER,

            pendientes INTEGER,

            archivo_corregido TEXT,

            reporte TEXT

        );


        CREATE TABLE IF NOT EXISTS auditoria_incidencias (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            ejecucion_id INTEGER,

            fila_excel INTEGER,

            registro TEXT,

            regla TEXT,

            columna TEXT,

            valor_original TEXT,

            valor_resultado TEXT,

            tratamiento TEXT,

            estado TEXT,

            descripcion TEXT

        );


        CREATE TABLE IF NOT EXISTS auditoria_cambios (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            ejecucion_id INTEGER,

            fila_excel INTEGER,

            registro TEXT,

            regla TEXT,

            columna TEXT,

            antes TEXT,

            despues TEXT

        );

        """
    )


    pending = sum(

        item["tratamiento"]
        == "REVISION_HUMANA"

        for item
        in issues

    )


    cursor = conn.execute(

        """

        INSERT INTO auditoria_ejecuciones(

            archivo,

            hash_sha256,

            incidencias,

            cambios,

            pendientes,

            archivo_corregido,

            reporte

        )

        VALUES (?, ?, ?, ?, ?, ?, ?)

        """,

        (

            input_path.name,

            file_hash(
                input_path
            ),

            len(issues),

            len(changes),

            pending,

            str(
                clean_path.resolve()
            ),

            str(
                report_path.resolve()
            ),

        )

    )


    run_id = int(
        cursor.lastrowid
    )


    # =====================================================
    # GUARDAR INCIDENCIAS
    # =====================================================

    for item in issues:

        conn.execute(

            """

            INSERT INTO auditoria_incidencias(

                ejecucion_id,

                fila_excel,

                registro,

                regla,

                columna,

                valor_original,

                valor_resultado,

                tratamiento,

                estado,

                descripcion

            )

            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)

            """,

            (

                run_id,

                item[
                    "fila_excel"
                ],

                item[
                    "registro"
                ],

                item[
                    "regla"
                ],

                item[
                    "columna"
                ],

                item[
                    "valor_original"
                ],

                item[
                    "valor_resultado"
                ],

                item[
                    "tratamiento"
                ],

                item[
                    "estado"
                ],

                item[
                    "descripcion"
                ],

            )

        )


    # =====================================================
    # GUARDAR CAMBIOS
    # =====================================================

    for item in changes:

        conn.execute(

            """

            INSERT INTO auditoria_cambios(

                ejecucion_id,

                fila_excel,

                registro,

                regla,

                columna,

                antes,

                despues

            )

            VALUES (?, ?, ?, ?, ?, ?, ?)

            """,

            (

                run_id,

                item[
                    "fila_excel"
                ],

                item[
                    "registro"
                ],

                item[
                    "regla"
                ],

                item[
                    "columna"
                ],

                item[
                    "antes"
                ],

                item[
                    "despues"
                ],

            )

        )


    conn.commit()

    conn.close()



# MAIN


def main() -> int:

    parser = argparse.ArgumentParser(

        description=(
            "P01 - Corrección y limpieza "
            "de base de clientes"
        )

    )


    parser.add_argument(

        "--input",

        type=Path,

        default=DEFAULT_INPUT,

        help=(
            "Ruta de CSV, XLS o XLSX. "
            "Por defecto utiliza "
            "data/clientes.xlsx"
        )

    )


    args = parser.parse_args()


    input_path = (
        args.input.resolve()
    )


    # =====================================================
    # VERIFICAR ARCHIVO
    # =====================================================

    if not input_path.exists():

        print()

        print(
            "[ERROR] "
            "No existe el archivo:"
        )

        print(
            input_path
        )

        print()

        print(
            "Verifica que clientes.xlsx "
            "esté dentro de la carpeta data."
        )

        return 1


    try:

        # =================================================
        # CARGAR DATOS
        # =================================================

        original_df = (
            load_dataset(
                input_path
            )
        )


        # =================================================
        # LIMPIAR
        # =================================================

        clean_df, issues, changes = (
            clean_dataset(
                original_df
            )
        )


        # =================================================
        # EXPORTAR
        # =================================================

        clean_path, report_path = (
            write_outputs(

                clean_df,

                original_df,

                issues,

                changes,

                input_path,

            )
        )


        # =================================================
        # AUDITORÍA
        # =================================================

        save_audit(

            input_path,

            clean_path,

            report_path,

            issues,

            changes,

        )


        # =================================================
        # CONTADORES
        # =================================================

        automatic = sum(

            item["tratamiento"]
            == "AUTOMATICA_SEGURA"

            for item
            in issues

        )


        pending = sum(

            item["tratamiento"]
            == "REVISION_HUMANA"

            for item
            in issues

        )


        # =================================================
        # RESULTADO EN TERMINAL
        # =================================================

        print()

        print(
            "P01 - PROCESO COMPLETADO"
        )

        print(
            "=" * 55
        )


        print(
            f"Archivo analizado:        "
            f"{input_path.name}"
        )


        print(
            f"Registros originales:     "
            f"{len(original_df)}"
        )


        print(
            f"Registros finales:        "
            f"{len(clean_df)}"
        )


        print(
            f"Incidencias detectadas:   "
            f"{len(issues)}"
        )


        print(
            f"Correcciones automáticas: "
            f"{automatic}"
        )


        print(
            f"Pendientes de revisión:   "
            f"{pending}"
        )


        print(
            f"Cambios aplicados:        "
            f"{len(changes)}"
        )


        print(
            "-" * 55
        )


        print(
            "Excel corregido:"
        )

        print(
            clean_path
        )


        print()


        print(
            "Reporte:"
        )

        print(
            report_path
        )


        print()


        print(
            "Base de datos SQLite:"
        )

        print(
            DB_PATH
        )


        print(
            "=" * 55
        )

        print()


        return 0


    except Exception as error:

        print()

        print(
            "[ERROR] "
            "No se pudo completar "
            "el proceso:"
        )

        print(
            error
        )

        print()

        return 1



# EJECUTAR


if __name__ == "__main__":

    raise SystemExit(
        main()
    )