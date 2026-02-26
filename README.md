💸 Actualizador Financiero & Agro (Argentina)

Un script de automatización en Python que se ejecuta en segundo plano para la extracción, procesamiento y volcado de cotizaciones financieras y datos agropecuarios de Argentina directamente en una planilla de Excel local.

✨ Características Principales

Ejecución Silenciosa: Diseñado para correr en terminal o en segundo plano sin interfaces pesadas, ideal para tareas programadas.

Extracción de Divisas: Realiza scraping y consultas a APIs para obtener el Dólar Oficial (BNA), Dólar MEP y Dólar Libre (Blue).

Índices Macroeconómicos: Obtiene valores históricos actualizados de UVA, Índice CAC (Cámara Argentina de la Construcción), Salario Mínimo Vital y Móvil (SMVyM) e IPC.

Agro / Pizarra Rosario: Descarga los precios diarios de cereales (Trigo, Maíz, Sorgo, Girasol, Soja) aplicando formato condicional en Excel para resaltar valores estimativos.

Gestión Inteligente de Excel: Crea automáticamente las hojas faltantes, rellena fechas sin cotización (arrastrando el último valor válido) y aplica estilos y anchos de columna.

Protección contra Bloqueos: Detecta si el archivo Excel está siendo utilizado por otro usuario o programa abortando el proceso para evitar corrupciones de datos.

🛠️ Tecnologías Utilizadas

Python 3

Pandas: Procesamiento, limpieza y reestructuración de datos (DataFrames).

BeautifulSoup4 & Requests: Web scraping y consumo de APIs REST.

Openpyxl: Lectura, escritura y estilizado de archivos Excel.

🚀 Uso e Instalación

Instala las dependencias necesarias:
Abre tu terminal o símbolo del sistema e instala las librerías:

pip install requests pandas beautifulsoup4 openpyxl urllib3


Configuración de la Ruta del Excel:
Abre el archivo Consulta_cotizaciones_auto.py con cualquier editor de texto (como el Bloc de notas o VS Code) y busca la variable EXCEL_FILE_PATH en la parte superior. Modifícala con la ruta exacta donde deseas que se guarde o actualice tu planilla de Excel.

Ejemplo:

EXCEL_FILE_PATH = r"C:\Mis documentos\Cotizaciones y datos macro.xlsx"


Prueba Manual:
Haz doble clic en el archivo ejecutar_actualizador.bat o ejecuta en consola:

python Consulta_cotizaciones_auto.py


Si el archivo Excel no existe en la ruta que especificaste, el script creará uno nuevo automáticamente, siempre y cuando la carpeta contenedora exista.

⏰ Configurar Actualización Automática cada 24hs (Windows)

Para que el script funcione de forma 100% autónoma todos los días, utilizaremos el Programador de Tareas de Windows:

Presiona la tecla Windows, escribe Programador de tareas (Task Scheduler) y ábrelo.

En el panel derecho, haz clic en Crear tarea básica...

Nombre: Ponle un nombre fácil, ej. "Actualizador Financiero". Clic en Siguiente.

Desencadenador: Selecciona Diariamente y haz clic en Siguiente. Configura la hora a la que quieres que se ejecute (ej. 10:00 AM).

Acción: Selecciona Iniciar un programa y haz clic en Siguiente.

Programa o script: Haz clic en Examinar... y busca el archivo ejecutar_actualizador.bat que descargaste con este repositorio.

Iniciar en (Opcional pero Recomendado): Pega la ruta de la carpeta donde está guardado el .bat (sin comillas y sin el nombre del archivo).

Haz clic en Finalizar.

¡Listo! Tu computadora abrirá el script automáticamente cada día a la hora fijada, descargará los datos y actualizará el Excel de forma autónoma. Si por casualidad tienes el Excel abierto en ese momento, el script lo detectará, te avisará en la consola y se cerrará por seguridad.
