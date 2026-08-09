# Generador Masivo de Diccionarios de Datos desde Shapefiles

Script en Python que escanea recursivamente una carpeta y genera, por cada
shapefile (`.shp`) que encuentra, un archivo Excel con su **diccionario de
datos**: estructura de campos, dominio de valores, tipo de dato y longitud
real de cada campo.

Pensado originalmente para capas de **Planes de Acondicionamiento
Territorial y Desarrollo Urbano** (D.S. N.º 012-2022-VIVIENDA), pero
funciona con cualquier shapefile.

## Características

- Recorrido recursivo de carpetas: procesa todos los `.shp` que encuentre,
  sin importar en qué subcarpeta estén.
- La **longitud real** de los campos de texto se obtiene directamente del
  esquema nativo del `.dbf` (vía [`pyshp`](https://pypi.org/project/pyshp/)),
  por lo que coincide con lo que se ve en QGIS
  (*Propiedades de la capa → Campos → Longitud*).
- Genera automáticamente el dominio de valores de cada campo (lista de
  valores únicos para campos categóricos, o un valor de ejemplo para
  campos numéricos o de texto libre).
- Detecta el tipo de geometría de la capa (`PUNTO`, `POLILINEA`,
  `POLIGONO`).
- Arma el título del reporte a partir del campo `DESCR_PLAN`, abreviando
  automáticamente el prefijo "PLAN ESPECIFICO" a "PE".
- Salida en Excel con formato corporativo (colores, bordes, encabezados)
  usando `openpyxl`.

## Instalación

```bash
git clone https://github.com/<tu-usuario>/shapefile-data-dictionary.git
cd shapefile-data-dictionary
pip install -r requirements.txt
```

> Requiere Python 3.9+. `geopandas` depende de librerías geoespaciales del
> sistema (GDAL/GEOS/PROJ); si la instalación con `pip` falla, se recomienda
> usar `conda`/`mamba`:
> ```bash
> conda install -c conda-forge geopandas pyshp openpyxl pandas
> ```

## Uso

```bash
python generar_diccionarios.py /ruta/a/la/carpeta/raiz
```

Por defecto, cada Excel (`DD_<nombre_capa>.xlsx`) se guarda en la carpeta
**padre** de la carpeta donde está el shapefile correspondiente (mismo
comportamiento que la versión original del script).

Para guardar todos los Excel generados en una sola carpeta de salida:

```bash
python generar_diccionarios.py /ruta/a/la/carpeta/raiz --salida /ruta/de/salida
```

### Ejemplo

```bash
python generar_diccionarios.py "C:\SIG\Proyectos\PDU_2024" --salida "C:\SIG\Diccionarios"
```

## Estructura del Excel generado

Cada hoja incluye:

1. **Título**: `DICCIONARIO DE DATOS - <DESCR_PLAN o nombre de la capa>`
2. **Metadatos**: objeto, geometría, grupo de objetos, feature class,
   tabla, fuente primaria/secundaria y uso normativo.
3. **Tabla de estructura de campos** con las columnas:
   `NOMBRE DE CAMPO`, `DOMINIO O VALORES POSIBLES (TIPOLOGÍA)`,
   `DESCRIPCIÓN`, `TIPO DE DATO`, `EXTENSION`.

## Notas

- El script usa la fuente `Swis721 Cn BT` para el Excel. Si no está
  instalada en tu sistema, Excel la sustituirá automáticamente al abrir el
  archivo (no afecta la generación del reporte). Puedes cambiarla editando
  la constante `FUENTE_EXCEL` al inicio del script.
- Si un campo de texto no aparece en el esquema del `.dbf` (caso poco
  común), se usa como respaldo el diccionario `LONGITUDES_RESPALDO`.

## Licencia

Este proyecto se distribuye bajo la licencia MIT. Ver [LICENSE](LICENSE).
