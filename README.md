# documentos-a-markdown

Convierte carpetas de documentos de Office en Markdown con estructura jerárquica consistente
—`#`, `##`, `###`— y valida que esa estructura sirva para alimentar un sistema RAG.

Nació de un problema concreto: carpetas completas de documentos en Word, PDF, PowerPoint y
Excel debían alimentar un RAG. Un RAG no lee documentos completos: los fragmenta, y la calidad
de lo que recupera depende de la estructura del texto. Sin jerarquía consistente, el
recuperador devuelve fragmentos sin contexto y las respuestas empeoran sin que aparezca ningún
error.

El proceso manual tomaba **alrededor de 30 minutos por documento**. El detalle está en
[`docs/proceso.md`](docs/proceso.md).

> **Esta es una versión de demostración**, reconstruida desde cero con datos ficticios.
> No incluye documentos, código ni datos de ninguna empresa.

## Qué convierte

| Formato | Resultado |
|---|---|
| `.docx` (Word) | Títulos, párrafos, listas, tablas, citas y énfasis |
| `.pptx` (PowerPoint) | Una sección por diapositiva, con su título, viñetas y tablas |
| `.xlsx` (Excel) | Una sección por hoja, con cada hoja como tabla |
| `.pdf` | ❌ Se detecta y se informa que no está soportado |

## Qué conserva y qué se pierde

### Word (`.docx`)

| Elemento | Resultado |
|---|---|
| Títulos, subtítulos y el título del documento | ✅ Encabezados `#`, `##`, `###` |
| Párrafos | ✅ Texto plano |
| Negrita y cursiva | ✅ `**texto**`, `*texto*` |
| Listas con viñetas y numeradas | ✅ Listas de Markdown, con sangría por nivel |
| Tablas | ✅ Tablas de Markdown |
| Citas y notas de texto | ✅ Blockquote `>` |
| Imágenes | ⚠️ Se marcan como pendientes, no se extraen |
| Encabezados y pies de página | ❌ Se pierden |
| Colores, fuentes, saltos de página | ❌ Se pierden a propósito |

### PowerPoint (`.pptx`)

El título de la primera diapositiva se convierte en el `#` del documento. Cada diapositiva es
un `##`, con sus cuadros de texto como viñetas y sus tablas como tablas.

### Excel (`.xlsx`)

Cada hoja se convierte en un `## Hoja: nombre` seguido de una tabla. La primera fila con
contenido se toma como encabezado. Las celdas vacías quedan en blanco.

La pérdida de formato visual es intencional en los tres casos: el objetivo es quedarse con la
**estructura**, que es lo que sirve para fragmentar y para leer en el navegador.

## Requisitos

- Python 3.10 o superior
- `python-docx`, `python-pptx` y `openpyxl`

```bash
pip install -r requisitos.txt
```

Las librerías se importan solo cuando hacen falta: si únicamente vas a convertir `.docx`, no
necesitas tener instalado `openpyxl`.

## Uso

Convertir una carpeta completa:

```bash
python convertir.py datos/ejemplo/ --salida salida/
```

Convertir un archivo suelto:

```bash
python convertir.py datos/ejemplo/informe-ejemplo.docx
```

Los `.md` se escriben junto al original, o en la carpeta que indiques con `--salida`. Los
archivos de formatos no soportados se informan y se omiten, sin detener el proceso.

## Validación de estructura

Convertir no basta: el `.md` tiene que servir para que un parser lo fragmente bien. Después de
cada conversión, el script revisa:

- que el primer encabezado sea `#` (H1)
- que no haya saltos de nivel (`#` → `###`)
- que la sangría de las listas sea múltiplo de 2
- que el documento tenga al menos un encabezado

Si algo falla, lo informa:

```
informe-ejemplo.docx -> salida/informe-ejemplo.md
  aviso: línea 3: el primer encabezado es H3, debería ser H1
  aviso: línea 6: sangría de 3 espacios (se esperan múltiplos de 2)
```

Si todo está bien, imprime `estructura OK`. Los avisos no bloquean la conversión: en el flujo
original eran exactamente los puntos que se revisaban a mano.

## Ejemplo incluido

En `datos/ejemplo/` hay tres documentos inventados, generados por `datos/crear-ejemplos.py`,
que contienen a propósito todos los elementos que el conversor debe manejar:

| Archivo | Qué prueba |
|---|---|
| `informe-ejemplo.docx` | Título, encabezados, énfasis, listas, tabla y cita |
| `presentacion-ejemplo.pptx` | Portada, viñetas y tabla dentro de una diapositiva |
| `inventario-ejemplo.xlsx` | Dos hojas con encabezados y cifras |

Sus conversiones `.md` están al lado, como resultado esperado: si el resultado cambia de forma
inesperada, algo se rompió.

## Estructura

```
documentos-a-markdown/
├── convertir.py                  El conversor
├── datos/
│   ├── ejemplo/                  Documentos ficticios y sus .md esperados
│   └── crear-ejemplos.py         Regenera los tres documentos
├── docs/
│   ├── proceso.md                Cómo se hizo, incluida la parte manual
│   ├── decisiones.md             Por qué está hecho así
│   ├── arquitectura.mermaid      Las piezas y cómo se hablan
│   └── flujo.mermaid             El recorrido de un dato
└── requisitos.txt
```

## Decisiones de diseño

Cuatro que explican el comportamiento, con el detalle en
[`docs/decisiones.md`](docs/decisiones.md):

**Los títulos bajan un nivel.** El `Title` de Word se convierte en `#` y `Heading 1` en `##`.
Así el documento siempre tiene un solo H1, que es la convención de Markdown.

**Las tablas respetan su posición.** Se recorre el cuerpo del documento en orden real, no la
lista de párrafos; de lo contrario todas las tablas aparecerían al final.

**Cada bloque se separa con una línea en blanco.** No es cosmético: en Markdown, una lista
pegada a una cita se absorbe dentro de la cita y deja de ser una lista.

**Las librerías se importan por formato.** Quien solo convierte Word no arrastra la dependencia
de Excel.

**No usa MarkItDown.** La herramienta original se apoyaba en esa librería, que cubre más
formatos en una sola llamada. Esta demo escribe la extracción por formato a propósito: sin
control sobre cómo se traduce la jerarquía no se puede validarla, y esa validación era el punto.
El detalle está en [`docs/decisiones.md`](docs/decisiones.md).

## Limitaciones conocidas

- **No convierte PDF.** Se detecta y se informa. Un PDF guarda la posición visual del texto, no
  su estructura, así que cualquier extracción pierde los títulos y las tablas. Es un problema
  distinto y merece su propio desarrollo.
- **No extrae imágenes.** Solo deja la marca de su posición.
- **Las celdas combinadas** de Word se repiten como celdas independientes.
- **Excel sin límite de filas.** Una hoja de miles de filas produce una tabla enorme; conviene
  revisar el resultado antes de entregarlo.

## Verificación

La conversión se probó contra los tres documentos de ejemplo: cada elemento de las tablas de
arriba se genera como se describe, y el validador responde `estructura OK` en los tres casos.

El validador se probó además con un Markdown mal formado (primer encabezado en `H3` y una
lista con sangría de tres espacios): reporta los dos problemas.
