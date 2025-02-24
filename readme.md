# Proyecto CSV a XLSX y PDF

Este proyecto convierte un archivo CSV de asistentes en un archivo XLSX, ordenando los nombres y apellidos de diferentes maneras y luego convierte el archivo XLSX a PDF.

## Archivos

- `asistentes.csv`: Contiene la lista de asistentes con nombres y apellidos.
- `csvToXlsxToPDF.py`: Convierte el archivo CSV a XLSX y luego a PDF.

## Uso

1. Asegúrate de tener `openpyxl`, `pandas`, y `pdfkit` instalados:
    ```sh
    pip install openpyxl pandas pdfkit
    ```

2. Instala `wkhtmltopdf`:
    - Descarga `wkhtmltopdf` desde [wkhtmltopdf download page](https://wkhtmltopdf.org/downloads.html).
    - Sigue las instrucciones de instalación para tu sistema operativo.
    - Asegúrate de que el ejecutable `wkhtmltopdf` esté en tu PATH o especifica la ruta en el script.

3. Coloca el archivo `asistentes.csv` en el mismo directorio que los scripts de Python.

4. Ejecuta el script de Python para convertir el archivo CSV a XLSX y luego a PDF:
    ```sh
    python csvToXlsxToPDF.py
    ```

5. El archivo `asistencia_YYYY-MM-DD.xlsx` y `asistencia_YYYY-MM-DD.pdf` se generarán en el mismo directorio.

## Estructura del CSV

El archivo `asistentes.csv` debe tener el siguiente formato:
```
Nombre,Apellido
Salvador,Delgado Bolancé
Juan Antonio, León Cobos
Héctor,Garcia Alba
Lorenzo,Perez Muñoz
...
```

## Notas

- Asegúrate de que el archivo CSV esté codificado en UTF-8.
- El script ordena los asistentes alfabéticamente por apellidos y luego por nombres en caso de apellidos iguales.
- Los nombres de `asistentes.csv` han sido generados por [ChatGPT](chatgpt.com).

## Autores

- [slvdr510](https://github.com/slvdr510)
- [julexo](https://github.com/julexo): Realiza aportaciones sobre la lógica de iteración.
