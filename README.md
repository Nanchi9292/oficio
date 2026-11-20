# Oficios de exoneración

Aplicación web para generar oficios de la Municipalidad de Pococí, editar el anexo de Excel y guardar los documentos de cada expediente.

## Abrir la aplicación

Sirva la carpeta con un servidor HTTP, por ejemplo con `python -m http.server 8000`, y abra [oficio.html](http://localhost:8000/oficio.html). La aplicación carga SheetJS, JSZip y html2pdf desde un CDN. Si dispone de una base de contribuyentes, puede importarla desde la aplicación o colocar `contribuyentes.xlsx` junto al HTML.

## Editor de Excel

- Abra un libro `.xlsx`, seleccione una hoja y edite las celdas o use la barra de fórmulas.
- Use Tab o Enter para desplazarse, y Escape para cancelar la edición. Puede pegar un rango copiado desde Excel y recorrer las hojas largas con los controles de filas.
- El editor recalcula fórmulas compatibles, incluidas referencias entre hojas y funciones de búsqueda. El libro descargado conserva las fórmulas y solicita un recálculo al abrirse en Excel.
- El editor admite un subconjunto de Excel. Valide en Excel los libros con fórmulas matriciales, referencias estructuradas u otras funciones avanzadas.

## Explorador de expedientes

- Cree o seleccione la carpeta de una cédula. Puede buscar carpetas por nombre o cédula, y filtrar y ordenar sus archivos.
- Agregue archivos con el selector o arrástrelos al explorador. Los adjuntos con nombres repetidos reciben un sufijo para conservar ambas copias.
- Puede renombrar carpetas, descargar, renombrar o eliminar archivos, y exportar la carpeta como ZIP. Imprimir permite guardar los PDF en el expediente correspondiente a la cédula del oficio.
- Los expedientes se guardan en el almacenamiento del navegador. Descargue el ZIP para conservar una copia fuera del navegador.
