# Excel/CSV Data Cleaner & Report Generator

Script en Python que toma un archivo CSV o Excel con datos "sucios" (típico export
de un sistema de ventas, CRM o formulario) y entrega un archivo limpio, listo para
usar, más un reporte resumen.

## Qué resuelve

Datos reales casi nunca vienen limpios. Este script corrige automáticamente:

- **Texto inconsistente**: nombres en mayúsculas/minúsculas mezcladas, espacios extra.
- **Fechas en múltiples formatos** (`15/03/2024`, `2024-03-16`, `2024.03.17`...) → se
  normalizan todas a `YYYY-MM-DD`.
- **Filas duplicadas** (mismo cliente/fecha/producto registrado más de una vez).
- **Valores faltantes**: precios vacíos se rellenan con el promedio del mismo
  producto; ciudades vacías se marcan como "Desconocido" en vez de quedar en blanco.

Y genera un Excel con dos hojas:
1. **Datos Limpios** — la tabla corregida, con encabezados formateados.
2. **Resumen** — ventas totales, ciudad con más ventas, producto más vendido,
   desglose por ciudad y por producto.

## Uso

```bash
pip install -r requirements.txt
python clean_data.py sample_messy_data.csv reporte_limpio.xlsx
```

El script imprime en consola un log de lo que corrigió (útil para mostrarle al
cliente exactamente qué se hizo):

```
=== Reporte de limpieza ===
Filas originales:        12
Duplicados eliminados:   4
Precios faltantes rellenados (con promedio del producto): 1
Ciudades desconocidas marcadas: 2
Filas finales:           8
```

## Adaptable a otros casos

La estructura (limpiar texto → normalizar fechas → quitar duplicados → rellenar
faltantes → generar resumen) se puede adaptar a cualquier dataset tabular:
inventarios, leads de ventas, registros de asistencia, listas de contactos, etc.
Solo cambian los nombres de columnas y las reglas de negocio del resumen.

## Nota de portafolio

Este proyecto fue construido como pieza de portafolio para freelancing
(automatización de datos / limpieza de Excel), usando datos de ejemplo
ficticios — no contiene información real de ningún cliente.
