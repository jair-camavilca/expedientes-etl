# Expedientes ETL

[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![Excel](https://img.shields.io/badge/Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/microsoft-365/excel)
[![Word](https://img.shields.io/badge/Word-2B579A?style=for-the-badge&logo=microsoftword&logoColor=white)](https://www.microsoft.com/microsoft-365/word)
![ETL](https://img.shields.io/badge/Proceso-ETL-6A1B9A?style=for-the-badge)
![Version](https://img.shields.io/badge/Version-1.0.0-2EA44F?style=for-the-badge)

Herramienta de procesamiento de datos judiciales que transforma archivos Excel y Word en libros Excel uniformes para alimentar sistemas de consulta judicial.

El sistema extrae y estandariza dos campos de salida:

- `codigo`: identificador judicial del expediente, normalizado.
- `parte`: parte procesal; contiene los dos primeros apellidos de la persona natural o la razón social de la entidad.

El proyecto está diseñado para ejecutarse localmente. Los documentos judiciales permanecen en el equipo del usuario y no forman parte del repositorio.

## Funcionalidades

- Procesamiento de archivos `.xlsx`, `.docx` y `.doc`.
- Un Excel independiente por cada documento de entrada.
- Normalización del correlativo inicial del expediente a cinco dígitos.
- Conservación de expedientes válidos con formatos atípicos o abreviados.
- Expansión de subexpedientes expresados como `0/1/2`.
- Extracción conservadora de dos apellidos, descartando nombres de pila y texto procesal secundario.
- Conservación de razones sociales y entidades públicas.
- Conservación de partes distintas cuando un mismo código presenta un conflicto.
- Registro visible de expedientes faltantes como `NO DEFINIDO`.
- Reporte de calidad con revisión manual, conflictos y resumen.
- Estado persistente `PENDIENTE`/`VERIFICADO` para alertas MEDIA.
- Formato corporativo en los archivos Excel generados.

## Requisitos

- Windows 10/11.
- Python 3.10 o superior.
- Microsoft Word instalado únicamente si se procesarán archivos `.doc` antiguos.

Dependencias Python:

```text
pandas>=2.0
openpyxl>=3.1
```

## Instalación

Desde la carpeta del proyecto:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

Si ya existe una instalación funcional de Python, también puede ejecutarse con el intérprete disponible en el equipo.

## Uso recomendado

1. Coloque los archivos `.xlsx`, `.docx` o `.doc` en `entrada`.
2. Ejecute `ejecutar_expedientes_etl.bat` con doble clic.
3. Revise los archivos generados en `salida`.
4. Consulte los controles en `notas`.

El programa no combina documentos distintos: crea un Excel de salida por cada archivo de entrada.

También puede ejecutarse desde PowerShell:

```powershell
cd C:\Work\proyectos\expedientes-etl
python expedientes_etl.py
```

Para seleccionar archivos concretos:

```powershell
python expedientes_etl.py `
  --excel C:\ruta\casos.xlsx `
  --word C:\ruta\resumen.docx
```

Si los encabezados del Excel no son detectados automáticamente:

```powershell
python expedientes_etl.py `
  --excel C:\ruta\casos.xlsx `
  --columna-codigo "ID Expediente" `
  --columna-parte "Litigantes"
```

Los parámetros anteriores se refieren a los encabezados originales del Excel de entrada:

- `ID Expediente`: identificador del expediente en la fuente.
- `Litigantes`: parte procesal o litigante registrado en la fuente.

No son los nombres de las columnas finales. La salida siempre contiene exactamente:
`codigo` y `parte`.

## Estructura

```text
expedientes-etl/
├── entrada/                    # Archivos judiciales locales; no se versiona su contenido
├── salida/                     # Excels procesados; no se versiona su contenido
├── notas/                      # Reportes auxiliares; no se versionan
├── expedientes_etl.py           # Aplicación principal
├── entidades_juridicas.txt      # Catálogo editable de términos institucionales
├── requirements.txt             # Dependencias Python
├── README.md                    # Documentación profesional
├── README_DESARROLLO.md         # Notas internas, ignoradas por Git
├── tests/                       # Pruebas unitarias e integración
├── ejecutar_expedientes_etl.bat # Lanzador para Windows
└── .gitignore
```

## Archivos de salida

Cada entrada produce un libro con una hoja principal que contiene exactamente `codigo` y `parte`.

Además, se genera un libro de control de calidad con:

- `revision_manual`: alertas, nivel, motivo y estado editable.
- `conflictos`: códigos asociados a partes diferentes.
- `resumen`: cantidades y explicación de los niveles de revisión.

### Niveles de revisión

- `ALTA`: requiere completar o corregir información en la fuente. Incluye `NO DEFINIDO`.
- `MEDIA`: requiere confirmar que la extracción sea correcta.
- `VERIFICADO`: se aplica a una alerta MEDIA confirmada y evita que vuelva a mostrarse como pendiente en la siguiente ejecución.

Marcar un caso como `VERIFICADO` no modifica el Excel principal ni elimina datos judiciales.

## Reglas de transformación

### Expedientes

- El primer bloque numérico se completa a cinco dígitos con ceros a la izquierda.
- Se conserva la estructura completa del código.
- Los duplicados exactos de `codigo + parte` se eliminan.
- Si un código tiene partes distintas, se conservan todas y se informa el conflicto.
- Un expediente ausente no se inventa: se registra como `NO DEFINIDO`.

### Partes

- Para personas naturales se conservan los dos primeros apellidos principales.
- Se descartan nombres de pila, apodos, contrapartes y texto procesal secundario.
- Se conservan apellidos compuestos cuando forman parte del nombre principal.
- Para entidades se conserva la razón social detectada.
- En Word, cada expediente hereda el titular correspondiente de su bloque o sección.

El archivo `entidades_juridicas.txt` permite agregar términos institucionales sin editar el código. Debe contener un término por línea.

## Pruebas

```powershell
cd C:\Work\proyectos\expedientes-etl
python -m unittest discover -s tests -v
```

Las pruebas cubren normalización de expedientes, extracción de partes, entidades, conflictos, alertas, archivos reales y estados verificados.

## Solución de problemas

### El programa indica que un archivo está abierto

Cierre el Excel de salida o el reporte de control de calidad indicado en el mensaje y vuelva a ejecutar el `.bat`.

### Falta una dependencia Python

Ejecute:

```powershell
python -m pip install -r requirements.txt
```

### Se procesa un archivo `.doc` antiguo

Debe estar instalado Microsoft Word. El sistema lo convierte temporalmente a `.docx`, lo procesa y elimina la copia temporal.

### Aparece `NO DEFINIDO`

Revise el Word original. Significa que se encontró la etiqueta de expediente, pero no un número. El sistema no inventa ese dato.

## Protección de información

Los archivos judiciales pueden contener información sensible. No suba documentos de entrada, resultados ni reportes al repositorio. El `.gitignore` incluido excluye esas carpetas; verifique siempre los archivos seleccionados antes de hacer `git push`.

## Contexto

* Proyecto en empresa
