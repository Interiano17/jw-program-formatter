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

Si no se define, el encabezado usa el valor genérico `Congregación`.

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

En Linux, el launcher incluido detecta automáticamente `.venv` junto al
proyecto y, si no existe, intenta usar `python3`:

```bash
./launch_meeting_generator.sh
```

Aparecerá una interfaz gráfica donde deberá:

1. **Seleccionar el documento del programa** (ej: `programa-mensual.docx`).
2. **Seleccionar la plantilla S-140**.
3. **Presionar "Generar y guardar como..."**.
4. Elegir el nombre y la ubicación del archivo de salida.

El diálogo propone `S-140_COMPLETADO.docx` junto a la plantilla, pero permite
elegir otra ruta. Antes de reemplazar un archivo existente se solicita
confirmación, y la fuente o la plantilla nunca pueden usarse como salida.

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
| `generator.py` | Coordina parser y writer con semántica de todo o nada |
| `results.py` | Define `GenerationResult`: semanas escritas, omitidas, errores y ruta confirmada |
| `parser.py` | Lee el documento fuente, detecta semanas, extrae datos con regex |
| `template_writer.py` | Carga la plantilla, localiza marcadores por contexto y reemplaza los textos preservando formato |
| `models.py` | Define `Participant`, `Assignment`, `MeetingWeek` y su contrato de validación |
| `validation.py` | Define el error agregado que presenta todos los problemas encontrados |
| `utils.py` | Funciones auxiliares: limpieza de texto, separación de nombres, logging |
| `config.py` | Patrones regex, constantes, marcadores de plantilla y tiempos por defecto |

---

## 🔍 Cómo funciona el parser

El parser obtiene las cabeceras desde los párrafos y las asignaciones desde
tablas de tres columnas:

1. **Detección de semanas**: Busca el patrón `SEMANA DEL ...` para marcar el inicio de cada semana.
2. **Extracción de cabecera**: Obtiene fecha, lectura semanal, presidente y canción inicial mediante regex.
3. **Identificación de secciones**: Detecta los encabezados `TESOROS DE LA BIBLIA`, `SEAMOS MEJORES MAESTROS` y `NUESTRA VIDA CRISTIANA`.
4. **Extracción de puntos**: Cada punto conserva su número, título, duración y participantes; también se separan filas que contienen varias asignaciones.
5. **Separación de nombres**: El separador `//` divide automáticamente estudiante y ayudante (o conductor y lector).
6. **Inferencias específicas**: Completa duraciones y números omitidos únicamente cuando las reglas existentes permiten deducirlos sin ambigüedad.
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
   ruta de salida.
7. **Resultado estructurado**: Solo informa éxito cuando la ruta creada existe;
   en cualquier fallo devuelve cero semanas escritas, todas las semanas
   detectadas como omitidas y los errores encontrados.
8. **Preservación de formato**: Opera a nivel de `runs` de python-docx, modificando solo el texto y conservando tipografía, tamaño, color, negrita, cursiva, etc.

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

La detección de los tres encabezados de sección se encuentra en
`ProgramParser._detect_section_from_lines()` dentro de `parser.py`. Cualquier
cambio debe acompañarse con casos de prueba para las variantes aceptadas.

Los patrones de cabecera, canción, duración y cierre que sí son compartidos se
mantienen en `config.py`.

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

### Calidad estática

Instale las herramientas de desarrollo y ejecute Ruff y Mypy con:

```bash
python -m pip install -e ".[dev]"
ruff check meeting_generator tests
mypy meeting_generator
```

El workflow `.github/workflows/quality.yml` ejecuta las pruebas, Ruff y Mypy en
cada `push` y `pull_request`.

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
- Si una semana o el conjunto completo excede la capacidad de la plantilla, la
  generación se detiene y muestra el problema; no se truncan asignaciones.
- Nombres sin separador `//` entre dos personas (ej: `Ana Ejemplo Beatriz Prueba`) no se dividirán automáticamente; se tratarán como un solo nombre.
- El parser asume que el formato del documento fuente sigue la estructura de reuniones de los Testigos de Jehová (con secciones Tesoros, Maestros, Vida Cristiana).

---

## 📄 Licencia

Uso interno. Desarrollado para automatizar tareas administrativas de congregación.
