#PowerfullPHPSpreadsheetGeneratorTESTS

## Descripción

El código presentado en este repositorio corresponde a pruebas realizadas con el afán de encontrar solución definitiva a la generación de archivos XLSX, XML de Office usando PHP solamente. Hay muchas librerías interesantes que ofrecen la posibilidad de manejo de Spreadsheets pero la gran mayoría tienen un consumo de memoria MUY ALTO.

Una forma efectiva de generar archivos de más de 1M de filas sin consumir todos los recursos del servidor es generar un archivo XML al cual se hace append de los datos con el formato XML de Office 2003. Sin embargo, ¿existe la posibilidad de generar archivos XLSX para la descarga del usuario normal? Los XML son efectivos pero no más cómodos para el usuario porque pueden llegar a abrirse con cualquier otro programa y no de manera predeterminada con Excel y el peso es superior al de un XLSX.

Según las pruebas realizadas si es posible. Al ser XLSX un archivo comprimido que contiene varios XML con información sobre el Spreadsheet se puede armar manualmente y comprimirlo con la extensión XLSX. Sin embargo el consumo de memoria para comprimir archivos de más de 200k registros con PHP es muy alto. Si se ejecuta la compresión del lado del SO. No con PHP sino utilizando alguna librería de ZIP ejecutándola con shell_exec o de alguna manera por consola. Solo basta con comprimir el archivo y asignarle la extensión XLSX. De esta forma es muy posible crear archivos con millones de registros de manera efectiva usando PHP.

Esta librería experimental, que aún no se encuentra completa busca solucionar este problema de manera efectiva permitiendo optimizar el uso de memoria y el uso de CPU para exportar archivos a partir de una base de datos con gran cantidad de registros.

## Liecencia
El código se ofrece con la licencia MIT sin ningún tipo de garantía y cada quien es responsable de cómo lo utiliza.