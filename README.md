# Meeting Generator - Automatización del Programa S-140

Aplicación Python que automatiza el llenado del documento oficial **S-140** (Programa para la reunión de entre semana) utilizando como fuente de datos un documento de programa mensual.

---

## 🚀 Instalación

### Requisitos previos

- Python 3.10 o superior
- pip (gestor de paquetes de Python)
- tkinter disponible en la instalación de Python; algunas distribuciones de
  Linux lo ofrecen como un paquete separado

### Pasos

1. Clonar o copiar el proyecto completo y entrar en su directorio raíz.

2. Crear y activar un entorno virtual:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

En Windows PowerShell, la activación equivalente es:

```powershell
.venv\Scripts\Activate.ps1
```

3. Instalar el proyecto y sus dependencias:

```bash
python -m pip install -e .
```

`pyproject.toml` es la fuente canónica de metadatos y dependencias. La única
dependencia externa de ejecución es `python-docx`; `tkinter`, `pathlib`, `re`,
`typing` y `logging` pertenecen a la biblioteca estándar de Python.

### Nombre de la congregación

El nombre no se guarda en el repositorio. Antes de iniciar la aplicación,
defínalo en el entorno local:

```bash
export MEETING_GENERATOR_CONGREGATION_NAME="Su congregación"
```

Si no se define, el encabezado usa `Congregación Loarque`.

---

## ▶️ Ejecución

Con el entorno virtual activo, puede iniciar la aplicación mediante el comando
instalado:

```bash
meeting-generator
```

También puede ejecutarla como módulo desde la raíz del proyecto:

```bash
python -m meeting_generator
```

En Linux, el launcher incluido busca un intérprete ejecutable junto al proyecto
en este orden: `.venv-app/bin/python`, `.venv/bin/python` y
`.venv/Scripts/python.exe`. Si ninguno está disponible, intenta usar `python3`
del `PATH`:

```bash
./launch_meeting_generator.sh
```

Aparecerá una interfaz gráfica donde deberá:

1. **Seleccionar el documento del programa** (ej: `programa-mensual.docx`).
2. **Seleccionar la plantilla S-140**.
3. **Presionar "Generar y guardar como..."**.
4. Elegir el nombre y la ubicación del archivo de salida.

Cancelar la selección de la fuente o la plantilla conserva el archivo y su
etiqueta anteriores. El botón Generar se habilita cuando ambos están
seleccionados y utiliza esas mismas rutas.

El diálogo propone `S-140_COMPLETADO.docx` junto a la plantilla, pero permite
elegir otra ruta. Antes de reemplazar un archivo existente se solicita
confirmación, y la fuente o la plantilla nunca pueden usarse como salida.

Durante la generación, la ventana sigue respondiendo y los controles de
selección y generación permanecen bloqueados. Se restauran al terminar,
tanto en éxito como en error. Si cierra la ventana durante la generación,
el cierre espera a que termine la escritura en curso.

---

## 📁 Estructura del proyecto

```
program_automation/
├── .github/workflows/quality.yml  # CI de análisis estático
├── pyproject.toml                 # Empaquetado, dependencias y entry point
├── launch_meeting_generator.sh    # Launcher portable para Linux
├── README.md
├── meeting_generator/
│   ├── __main__.py                # Entrada para python -m meeting_generator
│   ├── main.py                    # GUI con tkinter
│   ├── generator.py               # Orquestación fail-closed
│   ├── results.py                 # Resultado estructurado de generación
│   ├── parser.py                  # Analizador del documento fuente
│   ├── template_writer.py         # Escritor de la plantilla S-140
│   ├── models.py                  # Modelos de datos (dataclasses)
│   ├── validation.py              # Error agregado de validación
│   ├── utils.py                   # Utilidades compartidas
│   ├── config.py                  # Patrones y constantes compartidos
│   └── requirements.txt           # Compatibilidad con la instalación anterior
└── tests/
    ├── __init__.py                  # Habilita discovery desde la raíz
    ├── source_fixtures.py           # Fuentes DOCX sintéticas
    ├── template_fixture.py          # Plantilla DOCX sintética
    └── test_meeting_schedule.py
```

### Responsabilidad de cada módulo

| Módulo | Responsabilidad |
|---|---|
| `__main__.py` | Permite ejecutar el paquete con `python -m meeting_generator` |
| `main.py` | Interfaz gráfica (tkinter), selección de archivos y presentación del resultado |
| `generator.py` | Coordina parser y writer y conserva el estado de publicación |
| `results.py` | Define `GenerationResult`: publicación, semanas escritas, omitidas, errores y ruta de salida |
| `parser.py` | Lee el documento fuente, detecta semanas, extrae datos con regex |
| `template_writer.py` | Carga la plantilla, localiza marcadores por contexto y reemplaza los textos preservando formato |
| `models.py` | Define `Participant`, `Assignment`, `MeetingWeek` y su contrato de validación |
| `validation.py` | Define el error agregado que presenta todos los problemas encontrados |
| `utils.py` | Funciones auxiliares: limpieza de texto, separación de nombres, logging |
| `config.py` | Patrones regex, constantes, marcadores de plantilla y tiempos por defecto |

La configuración compartida efectiva vive en `meeting_generator/config.py`:
salida sugerida, logging, nombre de congregación, patrones utilizados y reglas
de horario. También contiene la única definición de `TEMPLATE_MARKERS`, una
tupla que el escritor importa para limpiar y validar contenido pendiente.
`PLACEHOLDERS` es el subconjunto utilizado para limpiar filas sobrantes; los
rótulos fijos del formulario permanecen en el documento.

La detección de secciones de la fuente se implementa en `parser.py`; los
encabezados que reconoce la plantilla se definen en `template_writer.py`.
La capacidad se descubre en la plantilla y los campos requeridos se validan
en `models.py`, sin límites de configuración que esos módulos no utilicen.

---

## 🔍 Cómo funciona el parser

El parser obtiene las cabeceras desde los párrafos y las asignaciones desde
tablas de tres columnas:

1. **Detección de semanas**: Busca el patrón `SEMANA DEL ...` para marcar el inicio de cada semana.
2. **Extracción de cabecera**: Obtiene fecha, lectura semanal, presidente y canción inicial mediante regex.
3. **Identificación de secciones**: Detecta los encabezados `TESOROS DE LA BIBLIA`, `SEAMOS MEJORES MAESTROS` y `NUESTRA VIDA CRISTIANA`.
4. **Extracción de puntos**: Cada punto conserva su número, título, duración y participantes; también se separan filas que contienen varias asignaciones.
   En filas combinadas debe haber un párrafo no vacío de título y un grupo de
   participantes por cada número. Si las cantidades no coinciden, se informa
   la semana y la fila y se rechaza el documento sin reemplazar la salida previa.
5. **Separación de nombres**: El separador `//` divide automáticamente estudiante y ayudante (o conductor y lector).
   Se acepta un nombre o dos nombres separados por un único `//`. Se rechazan
   grupos vacíos, grupos que solo contienen un sufijo, más de dos participantes
   y barras que no formen ese separador. El error indica semana y fila y evita
   reemplazar la salida previa. Presidente y oraciones requieren una sola persona.
6. **Inferencias específicas**: Completa duraciones y números omitidos únicamente cuando las reglas existentes permiten deducirlos sin ambigüedad.
   Las asignaciones de Tesoros y Ministerio sin número se rechazan indicando
   la semana y la fila, incluso cuando comparten celda con un encabezado.
   Los encabezados, canciones y filas vacías siguen siendo válidos; los números
   de Vida Cristiana conservan su regla de inferencia existente.
   Las filas no vacías con columnas omitidas o desplazadas, y las fusiones que
   impiden distinguir número, contenido y participantes, se rechazan con semana
   y fila. También se rechazan las asignaciones incompletas y las mezclas de
   oración o conclusión con otra actividad. Las filas combinadas deben conservar
   un número y un grupo de participantes por actividad; una única asignación no
   puede absorber varias duraciones o grupos. La conclusión puede compartir fila
   con una canción cuando no contiene otra asignación.
7. **Validación completa**: Rechaza el documento entero si falta un campo,
   canción, participante, duración o número requerido; nunca devuelve semanas
   parciales.

### Tolerancia a variaciones

- Acepta `PRESIDENTE:` y `PRESIDENTE.` (con punto)
- Acepta `CANCIÓN` y `CANCION` (con/sin tilde)
- Tolera espacios múltiples y saltos de línea
- Soporta variaciones en el formato de fechas (`SEMANA DEL X AL Y`, `SEMANA DEL X DE MES AL Y DE MES`)
- Reconoce los 66 libros bíblicos, incluidos nombres numerados y compuestos;
  la posición de la lectura delimita la fecha de la semana.

### Normalización y correcciones de nombres

La normalización limpia espacios, convierte nombres completamente en mayúsculas
a formato título y conserva los sufijos de participantes. Mantiene las letras,
la puntuación y todos los componentes del nombre: no corrige apellidos ni
selecciona personas por fecha o número de asignación.

La interfaz usa los nombres del documento fuente. Para corregirlos, actualice
la fuente o declare correcciones explícitas al usar la API. Cada `NameCorrection`
requiere el nombre original completo, el reemplazo y un motivo no vacío:

```python
from pathlib import Path

from meeting_generator.generator import generate_document
from meeting_generator.models import NameCorrection

result = generate_document(
    Path("programa-mensual.docx"),
    Path("S-140_S.docx"),
    Path("S-140_COMPLETADO.docx"),
    name_corrections=(
        NameCorrection(
            original_name="Ana Ejemploo",
            corrected_name="Ana Ejemplo",
            reason="Nombre confirmado por la persona responsable del programa",
        ),
    ),
)
```

El ejemplo usa nombres ficticios; el motivo debe indicar la justificación real.
Las correcciones comparan nombres completos tras normalizar espacios y mayúsculas,
incluidos los sufijos cuando estén presentes. Solo se aplican al documento de esa
ejecución, mantienen intacta la fuente y registran en el log el archivo, la semana,
el campo, el nombre anterior, el nuevo y el motivo. Las reglas que repitan un mismo
nombre original normalizado se rechazan antes de escribir la salida.
Los nombres originales y corregidos deben identificar a una sola persona:
no admiten separadores ni grupos vacíos. La misma validación se aplica a las
semanas proporcionadas directamente a `fill_template`.

Se retiraron los reemplazos automáticos de `Cerrano`, `Sussy` y `O”Connor`, así como
la sustitución de participante por fecha y punto: los documentos contienen
variantes, pero el historial no documenta una autorización para aplicarlas a toda
fuente. Cualquier corrección necesaria debe declararse con su justificación.

---

## 📝 Cómo funciona el template writer

El writer utiliza una estrategia de **reemplazo contextual**:

1. **Detección de formularios**: Encuentra cada copia del formulario en la plantilla buscando el marcador `[FECHA]`.
2. **Validación previa**: Comprueba la estructura y la capacidad de cada sección
   antes de modificar o guardar el resultado.
3. **Recorrido secuencial con estado**: Mantiene el contexto de la sección actual (cabecera, Tesoros, Maestros, Vida Cristiana) y del punto actual.
4. **Reemplazo por etiquetas**: Localiza etiquetas como `Presidente:`, `Estudiante/Ayudante:`, `Conductor/Lector:`, `[Título]`, `[Nombre]`, `[Nombre/Nombre]`, `Canción [Número]` y reemplaza el marcador adyacente.
5. **Verificación de salida**: Comprueba los formularios usados antes de limpiar
   los no utilizados; cualquier marcador pendiente impide guardar el archivo.
6. **Publicación atómica**: Guarda primero en un archivo temporal del mismo
   directorio, comprueba que sea un DOCX legible y solo entonces reemplaza la
   ruta de salida. Resuelve el directorio antes del reemplazo y verifica el
   archivo publicado; el temporal solo se elimina cuando no llegó a publicarse.
7. **Resultado estructurado**: Solo informa éxito cuando la ruta creada existe;
   un fallo anterior al reemplazo devuelve cero semanas escritas, todas las
   semanas detectadas como omitidas y ninguna ruta de salida. Si el reemplazo
   ya ocurrió, conserva las semanas publicadas, cero omitidas, la ruta y el
   error posterior. `published` indica que la publicación ocurrió, aunque el
   archivo deje de estar disponible; `succeeded` requiere que no haya errores
   y que el archivo siga siendo accesible y no esté vacío.
8. **Preservación de formato**: Al sustituir texto, opera a nivel de `runs` de python-docx y conserva tipografía, tamaño, color, negrita, cursiva, etc.

La interfaz distingue los fallos anteriores y posteriores a la publicación.
En un fallo posterior indica la ruta, las semanas escritas y si esta ejecución
creó el archivo o reemplazó uno anterior; el reemplazo no se revierte.

Al descubrir filas y escribir asignaciones, procesa una sola vez el contenido de
cada celda fusionada, en su fila de origen. Las continuaciones verticales no crean
filas adicionales ni borran horas o participantes; la limpieza auxiliar respeta
la identidad de la celda principal. Una celda de participantes compartida entre
asignaciones admite los mismos nombres en todas ellas; si se requieren nombres
distintos, se informa la semana y la fila y se conserva cualquier salida anterior.

Se escribe y verifica la hora de inicio de cada actividad, incluso cuando su
celda horaria está vacía, conservando el formato. Si una celda fusionada debe
mostrar horas distintas o una hora no queda escrita correctamente, se rechaza
la generación indicando semana, tabla y fila, sin reemplazar la salida anterior.
La comprobación se repite después de eliminar formularios y filas sobrantes y
al reabrir el DOCX temporal, antes de publicarlo. También se rechaza una fusión
si la limpieza altera la hora de una actividad que permanece en el documento.
Si una fila omite la columna horaria, se rechaza la plantilla en lugar de
escribir y validar la hora en la siguiente columna.
Las fusiones de horas con títulos o etiquetas se rechazan para conservar el
contenido, y las horas antiguas de filas sin actividad se limpian.

La **conclusión** dura tres minutos y requiere exactamente un participante.
Siempre se escribe el nombre indicado en la fuente y se verifica en la fila de
conclusión. Se utiliza su marcador de participante o, si falta, la celda final
vacía del lado derecho, conservando el formato. Si falta el nombre, no hay una
celda adecuada o la escritura falla, se rechaza la generación y se conserva
cualquier salida anterior.

### Formato que se conserva

En el documento generado se eliminan los formularios no utilizados y las filas
de asignación sobrantes. También se eliminan las tablas cuyos formularios quedan
todos sin utilizar. En el contenido que permanece se conserva:

- Tipografía (fuente, tamaño, color)
- Márgenes
- Formato de las tablas y celdas conservadas
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

La detección de los tres encabezados de sección se encuentra en
`ProgramParser._detect_section_from_lines()` dentro de `parser.py`. Cualquier
cambio debe acompañarse con casos de prueba para las variantes aceptadas.

Los patrones de cabecera, canción, nombres y cierre utilizados por los módulos
se mantienen en `config.py`. La expresión que extrae duraciones de los títulos
está en `ProgramParser._create_assignment()`.

### Agregar un nuevo tipo de punto

1. Definir el campo en `models.py` (en la dataclass correspondiente).
2. Agregar la lógica de extracción en `parser.py`.
3. Agregar el mapeo en `template_writer.py`.

### Ajustar la tolerancia del parser

Según el dato afectado, revisar en `ProgramParser` los métodos
`_create_assignment()`, `_parse_participants()`, `_detect_section_from_lines()`
y las funciones `_infer_missing_*()`. Las nuevas variaciones deben cubrirse con
una prueba de regresión.

---

## ✅ Pruebas

Desde la raíz del proyecto y con el entorno virtual activo:

```bash
python -m unittest discover -v
```

La suite genera fuentes y plantilla DOCX sintéticas, anonimizadas y temporales.
No necesita documentos operativos ni archivos binarios versionados y elimina
los fixtures al terminar.

### Medición de generación

Para repetir la medición de rendimiento:

```bash
python -m tests.benchmark_generation
```

El benchmark usa la fuente sintética de nueve semanas y la plantilla de las
pruebas, sin modificar documentos operativos. Mide la generación completa,
incluidos análisis, validación y guardado; excluye la creación de fixtures y la
interfaz gráfica. Ejecuta un calentamiento y cinco generaciones medidas, y
devuelve sus tiempos, la mediana, las versiones y una huella de las partes del
DOCX que permite comparar contenido y formato sin las fechas del contenedor ZIP.

La comparación local del punto 8, con Python 3.12.3 y `python-docx` 1.2.0, dio:

| Versión | Mediana para nueve semanas |
| --- | ---: |
| Antes de reutilizar filas y celdas | 1,69 s |
| Después | 0,95 s |

La reducción fue del 43,8 %, con las partes del DOCX idénticas. Los aproximadamente
2,23 s del análisis anterior quedan como referencia histórica: no se reprodujo
esa medición con los mismos documentos y condiciones. La mejora indicada arriba
compara los mismos fixtures y entorno antes y después del cambio.

### Calidad estática

Cree un entorno de desarrollo separado con Python 3.12 para que la instalación,
las pruebas y las herramientas usen el mismo intérprete:

```bash
python3.12 -m venv .venv-dev
source .venv-dev/bin/activate
python -m pip install -e ".[dev]"
python -m pip check
python -m unittest discover -v
python -m ruff check meeting_generator tests
python -m ruff format --check meeting_generator tests
python -m mypy meeting_generator
```

Si la instalación de Python no incluye `ensurepip`, puede crear el entorno con
`python3.12 -m virtualenv .venv-dev` cuando `virtualenv` esté disponible en el
sistema. En Windows, use `py -3.12 -m venv .venv-dev` y active el entorno con
`.venv-dev\Scripts\Activate.ps1`.

Los métodos del parser y del escritor usan los tipos de `python-docx` para tablas,
filas, celdas y párrafos. Mypy rechaza firmas incompletas; no se omiten errores de
estos módulos. Mantiene Python 3.10 como objetivo de compatibilidad.

El código de la aplicación y las pruebas usan Ruff Formatter. Para aplicar el
formato antes de revisar el diff:

```bash
python -m ruff format meeting_generator tests
```

Use nombres descriptivos para índices, participantes y rangos de texto. Los
comentarios deben explicar reglas o decisiones que el código no haga evidentes.

El workflow `.github/workflows/quality.yml` ejecuta las pruebas, Ruff, la
comprobación de formato y Mypy con Python 3.10 en cada `push` y `pull_request`,
usando también `python -m`.

---

## 📊 Logging

El programa escribe en consola y en `meeting_generator.log`, en el directorio
actual. El archivo rota al alcanzar 5 MiB y conserva tres copias:
`meeting_generator.log.1`, `.2` y `.3`. Los límites se definen en `config.py`
mediante `LOG_MAX_BYTES` y `LOG_BACKUP_COUNT`.

Si no se puede abrir el archivo de log o resolver el directorio actual, la
aplicación continúa con logging en consola y muestra una advertencia en la
interfaz. Si ya había handlers operativos, se conservan cuando falla la apertura
del nuevo archivo.

La configuración usa el logger `meeting_generator`. Al reconfigurar con éxito,
`setup_logging` sustituye y cierra sus handlers propios anteriores; el logger
global conserva su configuración y los mensajes del paquete no se propagan a él.

Los registros habituales contienen el número de semana, las cantidades de
asignaciones, inferencias, advertencias y resultados. No incluyen nombres de
participantes, rutas de archivos ni fragmentos del contenido fuente.
Los errores de validación conservan la semana, fila o punto afectado, y las
excepciones inesperadas conservan su traza para diagnosticar el fallo.

Las correcciones explícitas mantienen su registro de auditoría: archivo,
semana, campo, nombre original, nombre corregido y motivo. Estos datos permiten
comprobar las sustituciones solicitadas; no se registran durante la normalización.

---

## ⚠️ Limitaciones conocidas

- La plantilla S-140 debe conservar los marcadores originales (`[FECHA]`, `[Nombre]`, `[Título]`, etc.) para que el reemplazo funcione correctamente.
- Si una semana o el conjunto completo excede la capacidad de la plantilla, la
  generación se detiene y muestra el problema; no se truncan asignaciones.
- Nombres sin separador `//` entre dos personas (ej: `Ana Ejemplo Beatriz Prueba`) no se dividirán automáticamente; se tratarán como un solo nombre.
- El parser asume que el formato del documento fuente sigue la estructura de reuniones de los Testigos de Jehová (con secciones Tesoros, Maestros, Vida Cristiana).

---

## 📄 Licencia

Uso interno. Desarrollado para automatizar tareas administrativas de congregación.
