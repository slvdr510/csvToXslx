# Proyecto CSV a XLSX

Este proyecto convierte un archivo CSV de asistentes en un archivo XLSX, ordenando los nombres y apellidos de diferentes maneras.


## Archivos

- `asistentes.csv`: Contiene la lista de asistentes con nombres y apellidos.
- `alumnosCsvToXslx_alfabeticamenteDerecha.py`: Ordena los asistentes alfabéticamente por apellido y nombre y los coloca en el archivo XLSX de izquierda a derecha.
- `alumnosCsvToXslx_alfabeticamenteAbajo.py`: Ordena los asistentes alfabéticamente por apellido y nombre y los coloca en el archivo XLSX de arriba hacia abajo.

## Uso

1. Asegúrate de tener `openpyxl` instalado:
    ```sh
    pip install openpyxl
    ```

2. Coloca el archivo `asistentes.csv` en el mismo directorio que los scripts de Python.

3. Ejecuta uno de los scripts de Python según el orden deseado:

    - Para ordenar de izquierda a derecha:
        ```sh
        python alumnosCsvToXslx_alfabeticamenteDerecha.py
        ```

    - Para ordenar de arriba hacia abajo:
        ```sh
        python alumnosCsvToXslx_alfabeticamenteAbajo.py
        ```

4. El archivo `asistentes.xlsx` se generará en el mismo directorio.

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
- Los scripts ordenan los asistentes alfabéticamente por apellidos y luego por nombres en caso de apellidos iguales.
- Los nombres de `asistentes.csv` han sido generados por [ChatGPT](chatgpt.com).

## Autores

- [slvdr510](https://github.com/slvdr510)
- [julexo](https://github.com/julexo): Realiza aportaciones sobre la lógica de iteración.