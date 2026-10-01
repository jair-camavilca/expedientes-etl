# Guía de uso de Expedientes ETL

## Uso rápido

1. Cierre los archivos Excel o Word que vaya a procesar.
2. Ejecute `ejecutar_expedientes_etl.bat`.
3. Revise los archivos generados en `salida` y el control de calidad en `salida/control_calidad`.
4. Cuando termine la revisión, responda `S` a la pregunta del mismo `.bat` para guardar el trabajo, o `N` para conservarlo en las carpetas activas.
5. Si responde `S`, el programa moverá cada archivo procesado a su propia carpeta `trabajos/AAAA/AAAA-MM-DD_nombre/`.

## Uso desde la interfaz web local

Para una operación más sencilla, cierre Excel o Word y ejecute `ejecutar_expedientes_web.bat`. Se abrirá una página local en el navegador. Seleccione uno o varios archivos, pulse **Procesar archivos** y descargue el resultado Excel o el archivo ZIP con el control de calidad.

La interfaz funciona en la computadora del usuario. Los archivos no se cargan a Internet. Para utilizarla por primera vez, instale las dependencias indicadas en `requirements.txt`.

## Archivos aceptados

El programa procesa archivos:

- Excel: `.xlsx` y `.xls` según las dependencias disponibles.
- Word: `.docx`.
- Word antiguo: `.doc`, si Microsoft Word está instalado.

Los Excel pueden conservar títulos, columnas vacías, filas vacías y encabezados repetidos. El sistema busca automáticamente columnas relacionadas con expediente y parte procesal. No es necesario reorganizar manualmente el archivo.

## Carpetas principales

- `entrada`: archivos pendientes del trabajo actual.
- `salida`: resultados Excel del trabajo actual.
- `salida/control_calidad`: reportes `revision_manual`, `conflictos` y `resumen` del trabajo actual.
- `trabajos/AAAA/AAAA-MM-DD_nombre`: trabajos anteriores ya revisados.
- `config`: catálogos editables de entidades jurídicas y nombres frecuentes.
- `tests`: pruebas técnicas; no es necesario abrirla para ejecutar el programa.

## Resultado

La salida principal contiene únicamente:

- `codigo`: expediente normalizado cuando corresponde.
- `parte`: nombre o entidad principal extraída.

Los nombres se simplifican conservando la información útil. Los conflictos con un mismo código y distintas partes se mantienen como registros separados.

## Control de calidad

- `revision_manual`: casos que requieren confirmación en la fuente.
- `conflictos`: códigos relacionados con más de una parte.
- `resumen`: cantidades generales del procesamiento.

`NO DEFINIDO` significa que el documento no contenía un número de expediente detectable. El programa no inventa ese dato.

Los códigos incompletos o atípicos se conservan tal como aparecen y se muestran como alertas para que una persona los confirme.

## Cómo guardar y organizar varios trabajos

Después de revisar, responda `S` en el mismo `ejecutar_expedientes_etl.bat`. La carpeta resultante tendrá esta forma:

```text
trabajos/2026/2026-09-23_clientes_lima/
├── entrada/
└── salida/
    └── control_calidad/
```

El nombre se genera automáticamente con la fecha y el nombre de cada archivo original. Si se procesan varios archivos, se crean varias carpetas independientes. Después de guardar, las carpetas activas quedan listas para el siguiente trabajo.

## Si aparece un error

- Si indica que un archivo está abierto, cierre Excel o Word y vuelva a ejecutar el `.bat`.
- Si aparecen alertas, revise la fuente original antes de corregir la salida.
- No coloque resultados anteriores en la `entrada` de otro trabajo.
- No elimine información de la fuente para evitar una alerta.
