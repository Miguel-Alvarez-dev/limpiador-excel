"""Pruebas de limpiar.py. Se corren con:  pytest -q"""
import pandas as pd
import pytest

import generar_datos
import limpiar


def test_normalizar_encabezados():
    assert limpiar.normalizar("  FECHA de Venta ") == "fecha de venta"
    assert limpiar.normalizar("VENDEDOR(A)") == "vendedora"
    assert limpiar.normalizar("Precio_unitario") == "precio unitario"
    assert limpiar.normalizar("CANT.") == "cant"


@pytest.mark.parametrize(
    "texto, esperado",
    [
        ("15/01/2026", "2026-01-15"),
        ("2026-03-14", "2026-03-14"),
        ("05-02-2026", "2026-02-05"),
        ("  07/01/26 ", "2026-01-07"),
    ],
)
def test_fechas_en_texto(texto, esperado):
    assert limpiar.parsear_fecha(texto) == pd.Timestamp(esperado)


def test_fecha_como_numero_de_excel():
    serial = (pd.Timestamp("2026-01-05") - pd.Timestamp("1899-12-30")).days
    assert limpiar.parsear_fecha(serial) == pd.Timestamp("2026-01-05")


@pytest.mark.parametrize("valor", ["sin fecha", "", None, "31/02/2026", 12])
def test_fechas_invalidas_dan_nat(valor):
    assert pd.isna(limpiar.parsear_fecha(valor))


@pytest.mark.parametrize(
    "valor, esperado",
    [("$1,250.50", 1250.5), ("850", 850.0), (320, 320.0), (" $ 2,900.00 ", 2900.0), ("abc", None), (None, None)],
)
def test_numeros(valor, esperado):
    assert limpiar.parsear_numero(valor) == esperado


def test_nombres_propios():
    assert limpiar.nombre_propio("  juan   PÉREZ ") == "Juan Pérez"
    assert limpiar.nombre_propio("   ") is None


def test_unificar_variantes_deja_la_forma_mas_comun():
    serie = pd.Series(["Hub USB", "hub usb", "Hub USB", "HUB USB ", "Mouse", None])
    resultado = limpiar.unificar_variantes(serie)
    assert list(resultado[:4]) == ["Hub USB"] * 4
    assert resultado[4] == "Mouse"
    assert pd.isna(resultado[5])


def test_encabezado_con_filas_de_titulo_arriba():
    crudo = pd.DataFrame(
        [
            ["REPORTE DE VENTAS", None, None, None],
            [None, None, None, None],
            ["Fecha", "Cliente", "Producto", "Total"],
            ["01/01/2026", "Ana", "Mouse", 320],
        ]
    )
    assert limpiar.encontrar_encabezado(crudo) == 2


def test_sin_encabezados_reconocibles_falla_con_mensaje_claro():
    crudo = pd.DataFrame([["a", "b"], ["c", "d"]])
    with pytest.raises(ValueError, match="encabezados"):
        limpiar.encontrar_encabezado(crudo)


def test_flujo_completo_con_datos_sucios(tmp_path):
    esperado = generar_datos.generar(tmp_path / "entrada")
    salida = tmp_path / "salida" / "ventas_limpias.xlsx"

    resultado = limpiar.procesar(tmp_path / "entrada", salida)

    # Cada problema que se metió a propósito se detecta
    assert resultado["archivos"] == 3
    assert resultado["filas_vacias"] == esperado["filas_vacias"]
    assert resultado["duplicados"] == esperado["duplicados"]
    assert resultado["totales_calculados"] == esperado["totales_vacios"]
    assert resultado["para_revisar"] == esperado["fechas_invalidas"]
    assert resultado["filas_finales"] == esperado["filas_validas"]

    datos = pd.read_excel(salida, sheet_name="Datos")
    assert datos["Fecha"].notna().all()
    assert set(datos["Cliente"]) <= set(generar_datos.CLIENTES)
    assert set(datos["Vendedor"]) <= set(generar_datos.VENDEDORES)
    assert set(datos["Producto"]) == set(generar_datos.PRODUCTOS)
    assert (datos["Cantidad"] * datos["Precio unitario"] == datos["Total"]).all()

    revisar = pd.read_excel(salida, sheet_name="Revisar")
    assert len(revisar) == esperado["fechas_invalidas"]
    assert revisar["Motivo"].str.contains("fecha").all()
