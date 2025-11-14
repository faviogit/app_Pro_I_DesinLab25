AUTOMATIZACIÓN DE INFORMES: EXCEL A WORD CON PYTHON
🚀 Resumen del Proyecto
Este repositorio aloja una solución de automatización diseñada para transcribir datos de un archivo de Excel directamente a un documento de Word, generando un informe personalizado.

El objetivo principal es eliminar la tediosa tarea de transcripción manual, enfocándose específicamente en la creación automática de informes individuales de calificaciones para estudiantes. Es una herramienta fundamental para profesionales de oficina, docentes y administradores que buscan eficiencia en el manejo de documentos y datos.

# Tecnologías Clave
Este proyecto está construido principalmente con Python y aprovecha la potencia de las siguientes librerías:

Python 3.13: Entorno de ejecución requerido.

Pandas: Utilizado para la lectura, manipulación y análisis eficiente de los datos contenidos en el archivo de Excel.

python-docx (Asumido/Sugerido): Librería clave para interactuar y modificar el documento de Word (plantilla).

openpyxl: Dependencia utilizada por Pandas para leer archivos .xlsx modernos.

# Requisitos y Configuración
Para ejecutar este proyecto localmente, debes tener instalado Python 3.13 o superior.

1. Clonar el Repositorio
Bash

git clone https://github.com/faviogit/app_Pro_I_DesinLab25.git
2. Instalación de Dependencias
Se recomienda usar un entorno virtual. Luego, instala las librerías necesarias mediante pip:


# Instalación de librerías esenciales
pip install pandas
pip install openpyxl
pip install python-docx   (Requerida para la manipulación de Word)
Nota: Las dependencias completas deberían estar listadas en un archivo requirements.txt si el proyecto fuera a crecer.

📁 Estructura del Repositorio
La estructura del proyecto está diseñada para una clara separación entre el código, los datos de entrada y las plantillas:

.
├── datos/
│   ├── plantilla_informe.docx  # Plantilla base de Word con marcadores
│   └── datos_estudiantes.xlsx   # Archivo de Excel con las calificaciones
├── main.py                    # Script principal de automatización
└── README.md                  # Este archivo
main.py: Contiene la lógica central del programa: lee Excel, procesa datos y genera los archivos Word.

datos/: Carpeta que almacena los archivos de entrada (plantilla de Word y fuente de datos en Excel).

# 💡 Modo de Uso
El script main.py está configurado para leer los datos del archivo Excel y, basándose en la plantilla de Word, generar automáticamente el informe de calificaciones de cada estudiante.

Pasos para la Ejecución:
Asegúrate de que los archivos datos_estudiantes.xlsx y plantilla_informe.docx estén ubicados correctamente dentro de la carpeta datos/.

Ejecuta el script principal desde la línea de comandos:


python main.py
El script procesará los datos y los informes generados se guardarán en una carpeta de salida (se sugiere crear una carpeta output/ para alojar los informes finales, como Informe_Juan_Perez.docx).

# 🤝 Contribuciones y Contacto
Las contribuciones son bienvenidas, especialmente en la mejora de la eficiencia del procesamiento de datos o la optimización de la manipulación de documentos de Word.
