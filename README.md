# Meeting Generator - Automatización del Programa S-140

Aplicación Python que automatiza el llenado del documento oficial **S-140** (Programa para la reunión de entre semana) utilizando como fuente de datos un documento de programa mensual.

---

## 🚀 Instalación

### Requisitos previos

- Python 3.12 o superior
- pip (gestor de paquetes de Python)

### Pasos

1. Clonar o copiar la carpeta `meeting_generator` a su equipo.

2. Instalar las dependencias:

```bash
cd meeting_generator
pip install -r requirements.txt
```

La única dependencia externa es `python-docx`. Los demás módulos (`tkinter`, `pathlib`, `re`, `typing`, `logging`) son parte de la biblioteca estándar de Python.

---

## ▶️ Ejecución

```bash
python main.py
```

Aparecerá una interfaz gráfica donde deberá:

1. **Seleccionar el documento del programa** (ej: `JULIO-AGOSTO-REVISADO 2.docx`).
2. **Seleccionar la plantilla S-140** (ej: `S-140_S.docx`).
3. **Presionar "Generar S-140_COMPLETADO.docx"**.

El archivo generado se guardará como `S-140_COMPLETADO.docx` en el mismo directorio de la plantilla.

---

## 📁 Estructura del proyecto

```
meeting_generator/
│
├── main.py               # Punto de entrada, GUI con tkinter
├── parser.py             # Analizador del documento fuente (.docx)
├── template_writer.py    # Escritor de la plantilla S-140
├── models.py             # Modelos de datos (dataclasses)
├── utils.py              # Utilidades (logging, limpieza de texto, etc.)
├── config.py             # Configuración centralizada (regex, constantes)
├── requirements.txt      # Dependencias del proyecto
└── README.md             # Este archivo
```

### Responsabilidad de cada módulo

| Módulo | Responsabilidad |
|---|---|
| `main.py` | Interfaz gráfica (tkinter), selección de archivos, orquestación del flujo |
| `parser.py` | Lee el documento fuente, detecta semanas, extrae datos con regex |
| `template_writer.py` | Carga la plantilla, localiza marcadores por contexto y reemplaza los textos preservando formato |
| `models.py` | Define las clases de datos: `MeetingWeek`, `BibleTreasures`, `MinistryPoint`, `StudentAssignment`, `LivingAsChristians` |
| `utils.py` | Funciones auxiliares: limpieza de texto, separación de nombres, logging |
| `config.py` | Todas las regex compiladas, constantes, marcadores de plantilla, tiempos por defecto |

---

## 🔍 Cómo funciona el parser

El parser analiza el documento fuente párrafo por párrafo:

1. **Detección de semanas**: Busca el patrón `SEMANA DEL ...` para marcar el inicio de cada semana.
2. **Extracción de cabecera**: Obtiene fecha, lectura semanal, presidente y canción inicial mediante regex.
3. **Identificación de secciones**: Detecta los encabezados `TESOROS DE LA BIBLIA`, `SEAMOS MEJORES MAESTROS` y `NUESTRA VIDA CRISTIANA`.
4. **Extracción de puntos**: Cada punto se identifica por su número (1-10), seguido de un título y uno o más nombres de persona.
5. **Separación de nombres**: El separador `//` divide automáticamente estudiante y ayudante (o conductor y lector).
6. **Validación**: Cada semana debe tener al menos fecha, presidente y los 3 oradores de Tesoros.

### Tolerancia a variaciones

- Acepta `PRESIDENTE:` y `PRESIDENTE.` (con punto)
- Acepta `CANCIÓN` y `CANCION` (con/sin tilde)
- Tolera espacios múltiples y saltos de línea
- Soporta variaciones en el formato de fechas (`SEMANA DEL X AL Y`, `SEMANA DEL X DE MES AL Y DE MES`)

---

## 📝 Cómo funciona el template writer

El writer utiliza una estrategia de **reemplazo contextual**:

1. **Detección de formularios**: Encuentra cada copia del formulario en la plantilla buscando el marcador `[FECHA]`.
2. **Recorrido secuencial con estado**: Mantiene el contexto de la sección actual (cabecera, Tesoros, Maestros, Vida Cristiana) y del punto actual.
3. **Reemplazo por etiquetas**: Localiza etiquetas como `Presidente:`, `Estudiante/Ayudante:`, `Conductor/Lector:`, `[Título]`, `[Nombre]`, `[Nombre/Nombre]`, `Canción [Número]` y reemplaza el marcador adyacente.
4. **Preservación de formato**: Opera a nivel de `runs` de python-docx, modificando solo el texto y conservando tipografía, tamaño, color, negrita, cursiva, etc.

### Lo que NO se modifica

- Tipografía (fuente, tamaño, color)
- Márgenes
- Tablas
- Alineación
- Espaciado
- Bordes
- Estilos de párrafo
- Encabezados y pies de página

---

## 🛠️ Cómo agregar nuevas reglas

### Agregar un nuevo patrón regex

Editar `config.py` y añadir la nueva expresión regular compilada:

```python
RE_MI_NUEVO_PATRON: re.Pattern = re.compile(
    r"mi_patron_regex",
    re.IGNORECASE
)
```

Luego usarla en `parser.py` donde corresponda.

### Modificar la detección de secciones

En `config.py`, ajustar las regex de sección:

```python
RE_SECTION_TREASURES = re.compile(r"...")
RE_SECTION_MINISTRY = re.compile(r"...")
RE_SECTION_CHRISTIAN_LIFE = re.compile(r"...")
```

### Agregar un nuevo tipo de punto

1. Definir el campo en `models.py` (en la dataclass correspondiente).
2. Agregar la lógica de extracción en `parser.py`.
3. Agregar el mapeo en `template_writer.py`.

### Ajustar la tolerancia del parser

Modificar los métodos `_looks_like_title()` y `_looks_like_person_name()` en `parser.py` para ajustar las heurísticas de detección.

---

## 📊 Logging

El programa genera automáticamente un archivo `meeting_generator.log` en el directorio actual que registra:

- Archivo fuente leído
- Cantidad de semanas encontradas
- Cada semana procesada (fecha y estado)
- Advertencias (datos faltantes)
- Errores (semanas que no se pudieron procesar)
- Archivo generado

---

## ⚠️ Limitaciones conocidas

- La plantilla S-140 debe conservar los marcadores originales (`[FECHA]`, `[Nombre]`, `[Título]`, etc.) para que el reemplazo funcione correctamente.
- Si una semana tiene más puntos de los que admite la plantilla (por ejemplo, 5 puntos en "Seamos Mejores Maestros" cuando la plantilla solo tiene 4), los puntos sobrantes no se reflejarán.
- Nombres sin separador `//` entre dos personas (ej: `Lorena Cabrera Isabella Interiano`) no se dividirán automáticamente; se tratarán como un solo nombre.
- El parser asume que el formato del documento fuente sigue la estructura de reuniones de los Testigos de Jehová (con secciones Tesoros, Maestros, Vida Cristiana).

---

## 📄 Licencia

Uso interno. Desarrollado para automatizar tareas administrativas de congregación.