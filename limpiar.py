"""Junta y limpia archivos de ventas capturados a mano (Excel o CSV).

Lee todos los .xlsx y .csv de una carpeta aunque cada uno tenga encabezados
distintos, filas de título arriba, fechas en varios formatos o precios escritos
como texto ("$1,250.00"), y deja todo en un solo Excel ordenado.

Uso:
    python limpiar.py datos_ejemplo
    python limpiar.py mi_carpeta --salida salida/ventas_limpias.xlsx
"""
from __future__ import annotations

import argparse
import logging
import numbers
import re
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

log = logging.getLogger("limpiar")

COLUMNAS = ["fecha", "cliente", "producto", "vendedor", "cantidad", "precio_unitario", "total"]

# Nombres con los que puede aparecer cada columna en los archivos.
# Si llega un archivo con otro nombre, basta con agregarlo aquí.
ALIAS = {
    "fecha": ["fecha", "fecha de venta", "fecha venta"],
    "cliente": ["cliente", "nombre cliente", "nombre del cliente"],
    "producto": ["producto", "articulo", "item", "descripcion"],
    "vendedor": ["vendedor", "vendedora", "vendedor(a)", "agente"],
    "cantidad": ["cantidad", "cant", "cant.", "unidades", "piezas"],
    "precio_unitario": ["precio unitario", "precio", "precio unit", "precio_unitario"],
    "total": ["total", "importe", "monto"],
}

FORMATOS_FECHA = ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y")

NOMBRES_BONITOS = {
    "fecha": "Fecha",
    "cliente": "Cliente",
    "producto": "Producto",
    "vendedor": "Vendedor",
    "cantidad": "Cantidad",
    "precio_unitario": "Precio unitario",
    "total": "Total",
    "archivo_origen": "Archivo de origen",
    "fecha_original": "Fecha como venía",
    "motivo_revision": "Motivo",
}


def normalizar(texto) -> str:
    """'  FECHA de Venta ' -> 'fecha de venta' (sin acentos, mayúsculas ni signos)."""
    sin_acentos = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    solo_letras = re.sub(r"[^a-z0-9 _]", "", sin_acentos.lower())
    return re.sub(r"\s+", " ", solo_letras.replace("_", " ")).strip()


LOOKUP = {normalizar(alias): columna for columna, lista in ALIAS.items() for alias in lista}


# ---------------------------------------------------------------- lectura

def es_vacio(valor) -> bool:
    return valor is None or (not isinstance(valor, str) and bool(pd.isna(valor)))


def vacio_a_none(valor):
    """Celdas vacías, con solo espacios o NaN pasan a None."""
    if es_vacio(valor) or (isinstance(valor, str) and not valor.strip()):
        return None
    return valor


def leer_crudo(ruta: Path) -> pd.DataFrame:
    """Lee el archivo tal cual, sin asumir cuál es la fila de encabezados."""
    if ruta.suffix.lower() == ".xlsx":
        return pd.read_excel(ruta, header=None, dtype=object)
    for codificacion in ("utf-8-sig", "latin-1"):
        try:
            # sep=None deja que pandas detecte si el separador es coma o punto y coma
            return pd.read_csv(ruta, header=None, sep=None, engine="python", encoding=codificacion, dtype=object)
        except UnicodeDecodeError:
            continue
    raise ValueError("no se pudo leer la codificación del CSV")


def encontrar_encabezado(crudo: pd.DataFrame) -> int:
    """Busca la primera fila (entre las 20 primeras) con al menos 3 columnas conocidas."""
    for posicion in range(min(20, len(crudo))):
        reconocidas = sum(1 for celda in crudo.iloc[posicion] if normalizar(celda) in LOOKUP)
        if reconocidas >= 3:
            return posicion
    raise ValueError("no se encontró una fila de encabezados reconocible")


def extraer_tabla(ruta: Path) -> tuple[pd.DataFrame, int]:
    """Devuelve la tabla con las columnas estándar y cuántas filas vacías tenía."""
    crudo = leer_crudo(ruta)
    crudo = crudo.apply(lambda columna: columna.map(vacio_a_none)).astype(object)
    inicio = encontrar_encabezado(crudo)

    nombres = [LOOKUP.get(normalizar(celda)) for celda in crudo.iloc[inicio]]
    tabla = crudo.iloc[inicio + 1 :].copy()
    tabla.columns = nombres
    tabla = tabla.loc[:, [nombre is not None for nombre in nombres]]  # se descartan columnas desconocidas
    tabla = tabla.loc[:, ~tabla.columns.duplicated()].reindex(columns=COLUMNAS)

    vacias = tabla.isna().all(axis=1)
    tabla = tabla[~vacias].reset_index(drop=True)
    tabla["archivo_origen"] = ruta.name
    return tabla, int(vacias.sum())


# ---------------------------------------------------------------- limpieza de valores

def parsear_fecha(valor):
    """Devuelve un Timestamp o NaT. Se asume día/mes/año, que es el formato de México."""
    if es_vacio(valor):
        return pd.NaT
    if isinstance(valor, datetime):
        return pd.Timestamp(valor).normalize()
    if isinstance(valor, numbers.Real) and not isinstance(valor, bool):
        # Fecha guardada como número de Excel (días desde 1899-12-30)
        if 20000 < valor < 80000:
            return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(valor))).normalize()
        return pd.NaT
    texto = str(valor).strip()
    for formato in FORMATOS_FECHA:
        try:
            return pd.Timestamp(datetime.strptime(texto, formato))
        except ValueError:
            continue
    return pd.NaT


def parsear_numero(valor):
    """'$1,250.50' -> 1250.5. Se asume coma para miles y punto para decimales."""
    if es_vacio(valor):
        return None
    if isinstance(valor, numbers.Real) and not isinstance(valor, bool):
        return float(valor)
    texto = re.sub(r"[^\d.,\-]", "", str(valor)).replace(",", "")
    try:
        return float(texto)
    except ValueError:
        return None


def espacios_limpios(valor):
    if es_vacio(valor):
        return None
    texto = re.sub(r"\s+", " ", str(valor)).strip()
    return texto or None


def nombre_propio(valor):
    """'  juan   PÉREZ ' -> 'Juan Pérez'."""
    texto = espacios_limpios(valor)
    return texto.title() if texto else None


def unificar_variantes(serie: pd.Series) -> pd.Series:
    """'hub usb', 'HUB USB' y 'Hub USB' son lo mismo: deja la forma más frecuente."""
    limpia = serie.map(espacios_limpios).astype(object)
    # Según la versión de pandas, un vacío puede venir como None o como NaN
    clave = limpia.map(lambda texto: texto.casefold() if isinstance(texto, str) else None)
    mas_comun = limpia.groupby(clave).agg(lambda grupo: grupo.value_counts().idxmax())
    unificada = clave.map(mas_comun).astype(object)
    return unificada.where(limpia.notna(), None)


def motivo_de_revision(fila) -> str:
    faltan = []
    if pd.isna(fila["fecha"]):
        faltan.append("fecha inválida o vacía")
    if pd.isna(fila["producto"]):
        faltan.append("sin producto")
    if pd.isna(fila["total"]):
        faltan.append("sin total y no se pudo calcular")
    return "; ".join(faltan)


# ---------------------------------------------------------------- limpieza de la tabla completa

def limpiar(tabla: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    """Devuelve (filas buenas, filas para revisar a mano, conteos)."""
    t = tabla.copy()
    t["fecha_original"] = t["fecha"].map(lambda valor: "" if es_vacio(valor) else str(valor))
    t["fecha"] = pd.to_datetime(t["fecha"].map(parsear_fecha))
    t["cliente"] = t["cliente"].map(nombre_propio).astype(object)
    t["vendedor"] = t["vendedor"].map(nombre_propio).astype(object)
    t["producto"] = unificar_variantes(t["producto"])
    for columna in ("cantidad", "precio_unitario", "total"):
        t[columna] = pd.to_numeric(t[columna].map(parsear_numero), errors="coerce")

    # Si falta el total pero hay cantidad y precio, se calcula
    sin_total = t["total"].isna() & t["cantidad"].notna() & t["precio_unitario"].notna()
    t.loc[sin_total, "total"] = (t["cantidad"] * t["precio_unitario"]).round(2)

    t["motivo_revision"] = t.apply(motivo_de_revision, axis=1)
    revisar = t[t["motivo_revision"] != ""]
    buenas = t[t["motivo_revision"] == ""].drop(columns=["motivo_revision", "fecha_original"])

    antes = len(buenas)
    # Duplicado = fila idéntica en todos los campos (sin contar de qué archivo viene)
    buenas = buenas.drop_duplicates(subset=COLUMNAS).sort_values(["fecha", "cliente"]).reset_index(drop=True)
    buenas["cantidad"] = buenas["cantidad"].round().astype("Int64")

    conteos = {"duplicados": antes - len(buenas), "totales_calculados": int(sin_total.sum())}
    return buenas, revisar, conteos


# ---------------------------------------------------------------- salida a Excel

def dar_formato(hoja, filtro: bool = True) -> None:
    azul = PatternFill("solid", fgColor="1F3864")
    for celda in hoja[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = azul
        celda.alignment = Alignment(vertical="center")
        formato = {"Precio unitario": '"$"#,##0.00', "Total": '"$"#,##0.00', "Fecha": "dd/mm/yyyy"}.get(celda.value)
        if formato:
            for fila in hoja.iter_rows(min_row=2, min_col=celda.column, max_col=celda.column):
                fila[0].number_format = formato
    hoja.freeze_panes = "A2"
    if filtro and hoja.max_row > 1:
        hoja.auto_filter.ref = hoja.dimensions
    for columna in hoja.columns:
        largos = [len(str(celda.value)) for celda in columna if celda.value is not None]
        hoja.column_dimensions[get_column_letter(columna[0].column)].width = min(max(largos, default=8) + 2, 45)


def escribir_excel(buenas: pd.DataFrame, revisar: pd.DataFrame, resumen: list[tuple], ruta: Path) -> None:
    ruta.parent.mkdir(parents=True, exist_ok=True)
    columnas_revisar = ["archivo_origen", "fecha_original", "cliente", "producto", "vendedor", "cantidad", "precio_unitario", "total", "motivo_revision"]
    columnas_buenas = COLUMNAS + ["archivo_origen"]

    with pd.ExcelWriter(ruta, engine="openpyxl") as escritor:
        pd.DataFrame(resumen, columns=["Concepto", "Valor"]).to_excel(escritor, sheet_name="Resumen", index=False)
        buenas[columnas_buenas].rename(columns=NOMBRES_BONITOS).to_excel(escritor, sheet_name="Datos", index=False)
        revisar[columnas_revisar].rename(columns=NOMBRES_BONITOS).to_excel(escritor, sheet_name="Revisar", index=False)
        for nombre, hoja in escritor.sheets.items():
            dar_formato(hoja, filtro=nombre != "Resumen")
        for celda in escritor.sheets["Resumen"]["B"][1:]:
            celda.alignment = Alignment(horizontal="left")  # que "ninguno" y los números queden alineados


# ---------------------------------------------------------------- flujo principal

def procesar(carpeta: Path, salida: Path) -> dict:
    archivos = sorted(
        p for p in carpeta.iterdir() if p.suffix.lower() in {".xlsx", ".csv"} and not p.name.startswith("~$")
    )
    if not archivos:
        raise SystemExit(f"No encontré archivos .xlsx ni .csv en {carpeta}")

    tablas, detalle, errores, vacias = [], [], [], 0
    for ruta in archivos:
        try:
            tabla, n_vacias = extraer_tabla(ruta)
        except Exception as error:  # un archivo roto no debe detener a los demás
            log.error("No se pudo leer %s: %s", ruta.name, error)
            errores.append(ruta.name)
            continue
        log.info("%s: %d filas, %d filas vacías", ruta.name, len(tabla), n_vacias)
        tablas.append(tabla)
        detalle.append((ruta.name, len(tabla)))
        vacias += n_vacias
    if not tablas:
        raise SystemExit("Ningún archivo se pudo leer.")

    todo = pd.concat(tablas, ignore_index=True)
    buenas, revisar, conteos = limpiar(todo)

    resumen = [
        ("Archivos leídos", len(detalle)),
        ("Archivos con error", ", ".join(errores) or "ninguno"),
        *[(f"Filas leídas de {nombre}", filas) for nombre, filas in detalle],
        ("Filas vacías descartadas", vacias),
        ("Duplicados eliminados", conteos["duplicados"]),
        ("Totales calculados (cantidad x precio)", conteos["totales_calculados"]),
        ("Filas para revisar a mano (hoja Revisar)", len(revisar)),
        ("Filas limpias (hoja Datos)", len(buenas)),
    ]
    escribir_excel(buenas, revisar, resumen, salida)

    return {
        "archivos": len(detalle),
        "filas_leidas": len(todo),
        "filas_vacias": vacias,
        "duplicados": conteos["duplicados"],
        "totales_calculados": conteos["totales_calculados"],
        "para_revisar": len(revisar),
        "filas_finales": len(buenas),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Junta y limpia archivos de ventas en un solo Excel")
    parser.add_argument("carpeta", type=Path, nargs="?", default=Path("datos_ejemplo"))
    parser.add_argument("--salida", type=Path, default=Path("salida/ventas_limpias.xlsx"))
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(message)s")
    resultado = procesar(args.carpeta, args.salida)
    log.info("")
    log.info("Listo. Archivo creado: %s", args.salida)
    for clave, valor in resultado.items():
        log.info("  %s: %s", clave.replace("_", " "), valor)


if __name__ == "__main__":
    main()
