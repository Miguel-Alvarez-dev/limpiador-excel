"""Genera archivos de ventas intencionalmente desordenados para probar limpiar.py.

Todos los datos son inventados. Cada archivo tiene problemas distintos, como los
que aparecen en la vida real cuando varias personas capturan en Excel.

Uso:
    python generar_datos.py                  # crea la carpeta datos_ejemplo/
    python generar_datos.py --carpeta otra
"""
from __future__ import annotations

import argparse
import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

PRODUCTOS = {
    "Teclado mecánico": 850,
    "Mouse inalámbrico": 320,
    'Monitor 24"': 2900,
    "Audífonos": 540,
    "Webcam HD": 690,
    "Hub USB": 280,
}
CLIENTES = ["Juan Pérez", "María López", "Carlos Ramírez", "Ana Torres", "Luis Hernández", "Sofía Díaz", "Pedro Castillo"]
VENDEDORES = ["Laura", "Diego", "Marcos"]


def pedidos(cantidad: int, inicio: date, dias: int, rng: random.Random) -> pd.DataFrame:
    """Crea pedidos inventados, todos distintos entre sí."""
    vistos, filas = set(), []
    while len(filas) < cantidad:
        producto = rng.choice(list(PRODUCTOS))
        fila = {
            "fecha": inicio + timedelta(days=rng.randrange(dias)),
            "cliente": rng.choice(CLIENTES),
            "producto": producto,
            "vendedor": rng.choice(VENDEDORES),
            "cantidad": rng.randint(1, 5),
            "precio_unitario": PRODUCTOS[producto],
        }
        fila["total"] = fila["cantidad"] * fila["precio_unitario"]
        llave = tuple(fila.values())
        if llave not in vistos:
            vistos.add(llave)
            filas.append(fila)
    return pd.DataFrame(filas)


def ensuciar_texto(serie: pd.Series, rng: random.Random) -> pd.Series:
    """Mezcla mayúsculas, minúsculas y espacios de más, como pasa al capturar a mano."""

    def una(valor: str) -> str:
        azar = rng.random()
        if azar < 0.20:
            return valor.lower()
        if azar < 0.35:
            return f"  {valor} "
        if azar < 0.45:
            return valor.upper()
        return valor

    return serie.map(una)


def con_duplicados(df: pd.DataFrame, cuantos: int, semilla: int) -> pd.DataFrame:
    return pd.concat([df, df.sample(cuantos, random_state=semilla)], ignore_index=True).sample(
        frac=1, random_state=semilla
    ).reset_index(drop=True)


def moneda(valor: float) -> str:
    return f"${valor:,.2f}"


def archivo_enero(carpeta: Path, rng: random.Random, semilla: int) -> None:
    """Título arriba, fechas como texto, precios como "$1,250.00", duplicados y filas vacías."""
    df = con_duplicados(pedidos(40, date(2026, 1, 1), 31, rng), 2, semilla)
    df["fecha"] = df["fecha"].map(lambda d: d.strftime("%d/%m/%Y"))
    for col in ("cliente", "producto", "vendedor"):
        df[col] = ensuciar_texto(df[col], rng)
    df["precio_unitario"] = df["precio_unitario"].map(moneda)
    df["total"] = df["total"].map(moneda)
    df.columns = ["Fecha", "Cliente", "Producto", "Vendedor", "Cantidad", "Precio unitario", "Total"]

    vacia = pd.DataFrame([{}])
    df = pd.concat([df.iloc[:10], vacia, df.iloc[10:25], vacia, df.iloc[25:]], ignore_index=True)

    with pd.ExcelWriter(carpeta / "ventas_enero.xlsx", engine="openpyxl") as escritor:
        df.to_excel(escritor, sheet_name="Enero", index=False, startrow=3)
        hoja = escritor.sheets["Enero"]
        hoja["A1"] = "REPORTE DE VENTAS - ENERO 2026"
        hoja["A2"] = "Capturó: Laura"


def archivo_febrero(carpeta: Path, rng: random.Random, semilla: int) -> None:
    """Otros nombres de columna, una columna que sobra y tres totales sin llenar."""
    df = con_duplicados(pedidos(35, date(2026, 2, 1), 28, rng), 1, semilla)
    df["fecha"] = pd.to_datetime(df["fecha"])
    for col in ("cliente", "producto", "vendedor"):
        df[col] = ensuciar_texto(df[col], rng)
    df["total"] = df["total"].astype(object)
    for posicion in (4, 17, 30):
        df.loc[posicion, "total"] = None
    df["observaciones"] = [rng.choice(["", "", "", "entrega pendiente", "pagó en efectivo"]) for _ in range(len(df))]
    df.columns = ["FECHA DE VENTA", "NOMBRE CLIENTE", "ARTICULO", "VENDEDOR(A)", "CANT.", "PRECIO", "IMPORTE", "OBSERVACIONES"]

    with pd.ExcelWriter(carpeta / "ventas_febrero.xlsx", engine="openpyxl") as escritor:
        df.to_excel(escritor, sheet_name="Hoja1", index=False)


def archivo_marzo(carpeta: Path, rng: random.Random) -> None:
    """CSV con punto y coma y acentos en latin-1, con dos fechas que no se pueden leer."""
    df = pedidos(30, date(2026, 3, 1), 31, rng)
    df["fecha"] = df["fecha"].map(lambda d: d.isoformat())
    df.loc[6, "fecha"] = "sin fecha"
    df.loc[21, "fecha"] = ""
    for col in ("cliente", "producto", "vendedor"):
        df[col] = ensuciar_texto(df[col], rng)
    df.to_csv(carpeta / "ventas_marzo.csv", sep=";", encoding="latin-1", index=False)


def generar(carpeta: Path, semilla: int = 42) -> dict:
    """Crea los tres archivos y devuelve cuántos problemas se metieron a propósito."""
    carpeta.mkdir(parents=True, exist_ok=True)
    rng = random.Random(semilla)
    archivo_enero(carpeta, rng, semilla)
    archivo_febrero(carpeta, rng, semilla)
    archivo_marzo(carpeta, rng)
    return {
        "filas_validas": 40 + 35 + 30 - 2,
        "duplicados": 3,
        "filas_vacias": 2,
        "fechas_invalidas": 2,
        "totales_vacios": 3,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Genera datos de ventas desordenados")
    parser.add_argument("--carpeta", type=Path, default=Path("datos_ejemplo"))
    args = parser.parse_args()
    esperado = generar(args.carpeta)
    print(f"Listo: archivos creados en {args.carpeta}/")
    print(f"Problemas metidos a propósito: {esperado}")


if __name__ == "__main__":
    main()
