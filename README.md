# Conciliador Bancario

> **Estado: prototipo / MVP.** Es una primera versión funcional para validar la idea. Puede tener errores con formatos de extracto poco comunes; revisá siempre los resultados antes de usarlos en registros contables.

Herramienta web para conciliar el libro mayor contra el extracto bancario. Funciona completamente en el navegador: los archivos no se suben a ningún servidor.

**Demo:** https://conciliador-banck.netlify.app

## Qué hace

- Lee archivos `.xlsx`, `.xls` y `.csv`, incluso si el extracto trae filas de título antes del encabezado.
- Detecta automáticamente las columnas de fecha, descripción y monto (o débito/crédito por separado), con opción de corregirlas a mano.
- Normaliza montos en formato argentino (`1.234,56`) y de otros formatos, y fechas `dd/mm/aaaa`, `aaaa-mm-dd` o seriales de Excel.
- Concilia con tolerancia configurable de monto y de días, emparejando cada movimiento una sola vez con la fecha más cercana.
- Muestra tres resultados: conciliados, solo en el mayor y solo en el banco, con la diferencia neta.
- Exporta un Excel con la conciliación completa y una hoja de pendientes.

## Uso local

No requiere instalación. Abrí `index.html` en el navegador o serví la carpeta con cualquier servidor estático.

## Deploy

Sitio estático publicado en Netlify (sin build command, publish directory `/`).

## Créditos

Basado en una idea de [LuMarka/ComparadorExcel](https://github.com/LuMarka/ComparadorExcel).
