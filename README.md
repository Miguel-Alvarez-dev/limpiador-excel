# Limpiador de archivos de ventas (Excel y CSV)

Script en Python que junta varios archivos de ventas capturados a mano, cada uno con su propio desorden, y los deja en un solo Excel limpio y ordenado.

Los datos de ejemplo son inventados. Los generé yo para tener algo con qué probar.

## El problema

Cuando varias personas capturan ventas en Excel, cada archivo sale distinto: uno trae títulos arriba, otro escribe los precios como texto (`$1,250.00`), otro llama a la columna `IMPORTE` en lugar de `Total`, y siempre hay filas repetidas, nombres en mayúsculas o con espacios de más y totales que faltan. Juntar todo a mano toma horas y es fácil equivocarse.

## Qué hace

- Encuentra la fila de encabezados aunque haya títulos o filas vacías arriba.
- Reconoce columnas con nombres distintos (`Fecha` / `FECHA DE VENTA`, `Cantidad` / `CANT.`, `Total` / `IMPORTE`) e ignora las que sobran.
- Lee `.xlsx` y `.csv`, detectando el separador y los acentos.
- Convierte las fechas, vengan como texto en varios formatos, como fecha de Excel o como número.
- Convierte precios escritos como texto (`$1,250.00`) en números.
- Unifica nombres y productos: `juan   PÉREZ` y `JUAN PÉREZ` pasan a `Juan Pérez`.
- Calcula el total cuando falta (cantidad por precio).
- Quita filas vacías y duplicados.
- Lo que no puede arreglar no lo borra ni lo inventa: lo manda a la hoja **Revisar**, con el motivo.

## Cómo usarlo

```bash
pip install -r requirements.txt
python generar_datos.py          # crea la carpeta datos_ejemplo/ con 3 archivos sucios
python limpiar.py datos_ejemplo  # junta y limpia todo
```

Con tus propios archivos: `python limpiar.py mi_carpeta --salida salida/resultado.xlsx`.

El resultado es `salida/ventas_limpias.xlsx`, con tres hojas:

| Hoja | Qué contiene |
| --- | --- |
| Resumen | Cuántas filas se leyeron, cuántas se limpiaron y qué se corrigió |
| Datos | Las ventas limpias, ordenadas por fecha |
| Revisar | Las filas que necesitan revisión humana, con el motivo |

## Resultado con los datos de ejemplo

Tres archivos (enero, febrero y marzo), cada uno con problemas distintos:

| Concepto | Valor |
| --- | --- |
| Filas leídas | 108 |
| Filas vacías descartadas | 2 |
| Duplicados eliminados | 3 |
| Totales calculados | 3 |
| Filas para revisar (fecha inválida) | 2 |
| Filas limpias | 103 |

## Supuestos que conviene saber

- Las fechas se leen como **día/mes/año**, que es el formato de México. Una fecha como `03/04/2026` se interpreta como 3 de abril.
- En los números, la coma separa miles y el punto los decimales.
- Un duplicado es una fila idéntica en todos los campos. Si en un negocio es normal que el mismo cliente compre lo mismo el mismo día, hay que revisar esta regla.

## Adaptarlo a otros archivos

Si llega un archivo con otro nombre de columna (por ejemplo `Artículo vendido`), se agrega a la lista `ALIAS` al inicio de `limpiar.py`. No hay que tocar nada más.

## Pruebas

```bash
pytest -q
```

Son 22 pruebas: las funciones de limpieza una por una y un flujo completo con datos sucios que verifica que cada problema metido a propósito se detecte. Probado con pandas 2.3 y pandas 3.0.

## Herramientas

Python, pandas, openpyxl y pytest.
