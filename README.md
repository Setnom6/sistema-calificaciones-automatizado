# Evaluación Automatizada con Google Apps Script

Este proyecto automatiza la creación y actualización de hojas de calificaciones en Google Sheets, 
a partir de listados de alumnos, instrumentos de evaluación y criterios.

---

## Qué hace
- Crea automáticamente una hoja de calificaciones por clase (`calificaciones_<nombre>`).
- Genera hojas individuales de **desglose** y **media** para cada alumno.
- Inserta fórmulas en la hoja `general` con los promedios trimestrales.
- Detecta y añade nuevos alumnos desde la hoja `listado`.

---

## Estructura esperada en Drive

Debes tener una carpeta de clase con esta estructura:

/carpetaGeneralCalificaciones/
├── proyectoAppsScript
└── tuCarpetaDeClase/
    ├── listado
    ├── criteriosDeEvaluacion
    └── instrumentos


Donde los tres archivos correspondientes a cada clase son hojas de cálculo con una estructura fija.

Puedes ver un ejemplo público aquí:  

[Carpeta de ejemplo en Drive](https://drive.google.com/drive/folders/1toUH6mrPhyBBRJA9DyIzLdK8Iwr8HJUU?usp=sharing)

*(solo lectura, sin datos personales)*

También hay una copia de esos archivos de ejemplo en este mismo repositorio. Se pueden tener tantas carpetas para clases distintas como se quiera, siempre que mantengan esa estructura interna.

---

## Instalación

1. Crea una **copia** de la carpeta de ejemplo en tu Google Drive.
2. Renombra la carpeta de `claseEjemplo` como quieras y rellena los google sheets que contiene con tus datos (ver README.md de ejemplo para saber formato).
3. Abre el proyecto de Apps Script `creador clase`.
4. Una vez dentro del proyecto de Apps Script, abre el archivo `buildProject.gs`, cambia la variable const `CLASS_FOLDER_NAME = "claseEjemplo"` por el nombre que tenga tu carpeta de clase y pulsa `Ejecutar`.
5. Espera a que la ejecución termine (unos 2 minutos) y vuelve al Drive. Se habrá creado una nueva hoja de calificaciones lista para usar.

---

## Uso

Una vez se ha construido el archivo de calificaciones para una clase concreta se puede usar manualmente. En la hoja 'general' aparecerá el listado de alumnos de esa clase junto con casillas para las calificaciones de cada trimestre, las cuales se actualizarán automáticamente. Pulsando en el enlace 'Abrir desglose' se abrirá la hoja de calificaciones desglosadas del alumno. Aquí se podrá introducir por cada instrumento de evaluación, las notas de los criterios de evaluación afectados. A la derecha se muestra una media orientativa del instrumento de evaluación en cuestión. También se puede acceder a la hoja 'media' de cada alumno para visualizar la mediad de las competencias hasta el momento. Allí se calculará una media del trimestre que es la que aparece automáticamente en la hoja 'general'.

## Actualizar datos

Se puede añadir manualmente hasta un instrumento de evaluación extra en cada alumno y en cada trimestre. Sin embargo, si se desa añadir un alumno nuevo o algún instrumento a todos los alumnos, hay que ejecutar scripts desde Apps Scripts. 

### Añadir instrumento de evaluación

1. Primero se añade el criterio en el google sheet original 'intrumentos' de esa clase.
2. Abrimos `creador clase`y buscamos el archivo `actualizarInstrumentos.gs`. Una vez en ese archivo, escribimos la clase que queremos actualizar (nombre de la carpeta) y lo Ejecutamos.
3. Al acabar, se habrá actualizado la hoja de 'desglose' de todos los alumnos con el nuevo instrumento de evaluación

### Añadir nuevo alumno

1. Primero añadimos el nuevo alumno al google sheet 'listado'.
2. Abrimos `creador clase`y buscamos el archivo `actualizarListado.gs`. Una vez en ese archivo, escribimos la clase que queremos actualizar (nombre de la carpeta) y lo Ejecutamos.
3. Al acabar, se habrá añadido el alumno en la hoja 'general' de su clase, así como sus archivos 'desglose' y 'media'

## Créditos
Creado por José Manuel Montes Armenteros.  
Inspirado en un sistema modular de evaluación en Google Sheets para docentes.

Licencia: BSD 3-Clause License
