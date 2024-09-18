# Procesador de coincidencias difusas en Excel

Este programa en Python procesa dos columnas de un archivo Excel, buscando coincidencias difusas entre ellas y generando un nuevo archivo con los resultados. El algoritmo utiliza la biblioteca fuzzywuzzy para comparar nombres en dos columnas, determinar la coincidencia más cercana y obtener un identificador asociado a esa coincidencia.

El programa lee un archivo de Excel (con formato .xls o .xlsx) que contiene dos columnas relevantes para la comparación:

- Columna B: Lista de nombres a ser buscados.
- Columna D: Lista de nombres contra los cuales se buscarán coincidencias.
- Columna C: Contiene un identificador que se extrae una vez se encuentra la coincidencia en la columna D.

El programa realiza las siguientes tareas:

- Lectura del archivo Excel: El archivo se carga utilizando pandas, detectando si es un formato .xls o .xlsx para elegir el motor de lectura adecuado.
- Comparación difusa: Utilizando fuzzywuzzy, se realiza una comparación difusa entre los valores de la columna B y la columna D, determinando cuál es la coincidencia más cercana.
- Recuperación del identificador: Una vez encontrada la mejor coincidencia en la columna D, se extrae el identificador correspondiente de la columna C.
- Generación de resultados: El programa genera un nuevo archivo Excel con tres columnas:
  - Nombre: El nombre original de la columna B.
  - Identificador: El identificador correspondiente a la coincidencia encontrada en la columna D.
  - Coincidencia: El nombre en la columna D que más se asemeja al de la columna B.
  - Porcentaje: El porcentaje de similitud entre los nombres comparados.




## Cómo ejecutar este programa

Este programa solo se ha comprobado con las versiones de Python > 3.10.x
- Acceder a la web y descarga la última versión de Python3: https://www.python.org/downloads/
- Crea un directorio en tu escritorio.
- Dentro de este directorio copiaremos y pegaremos ****TODOS**** los archivos dentro de este repositorio.
- Abrimos la terminal de tu PC y nos moveremos hasta posicionarte en el directorio del proyecto.
- Crearemos un entorno virtual (opcional pero recomendado)
  - Introducimos: ***python -m venv env***
  - Activamos el entorno virtual:
    - Si es Windows: ***.\env\Scripts\activate***
    - Si es MacOS / Linux: ***source env/bin/activate***
  - Instalamos las dependencias necesarias (encontradas en el archivo ***requirements.txt***):
    - Escribimos el comando: ***pip install -r requirements.txt***
- Finalmente, ejecutamos el programa con el comando:
  - Si el archivo es old_file.xls: ***"python main.py*** 
  - Si el archivo tiene diferente formato: ***"python main.py [nombre]"***.

Para que esto funcione es necesario que dentro del directorio que trabajemos estén los archivos main.py, old_file.xls o .xlsx y requirements.txt.




## Requisitos

Para ejecutar este programa es necesario tener en el mismo directorio el archivo que queremos tratar.

- Debe de tener como nombre: ***old_file***
- Solo se permiten extensiones ***".xls" y ".xlsx"***





## Resultados

Se generará como resultado un archivo con nombre ***file_trained con la extensión .xlsx*** con las siguientes columnas:

- **Nombre**: Valor de la columna con la que vas a trabajar.
- **Identificador**: Valor encontrado en la coincidencia.
- **Coincidencia**: Valor con el match más exacto entre la fila que usarás y con la que quieres coincidir.
- **Procentaje**: Valor numérico en formato de porcentaje de match obtenido.





### Copyright

© 2024 ***[QN Digital Solutions]***. All Right Reserved.