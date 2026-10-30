# Decisiones de diseño

Una línea por decisión, con el porqué. Es el documento que responde "¿y por qué así?" en una
entrevista.

---

## 2026-10-29 — La salida es Markdown, no HTML ni texto plano

**Contexto:** los documentos debían alimentar un RAG y, de paso, poder compararse entre
versiones y leerse sin abrir Word.
**Opciones:** `.txt`, `.html`, `.md`.
**Decisión:** `.md`. Se versiona con git línea a línea, GitHub lo renderiza, y conserva la
estructura semántica (títulos, listas, tablas) sin arrastrar estilos.
**Consecuencia:** se pierde el formato visual; se aceptó porque nadie lo usaba para imprimir.

---

## 2026-10-29 — Conversor propio en vez de pandoc

**Contexto:** pandoc ya convierte `.docx` a Markdown y es más completo.
**Opciones:** invocar pandoc, escribir el conversor.
**Decisión:** escribirlo, con una sola dependencia (`python-docx`).
**Consecuencia:** menos funciones, pero se entiende cada línea y no hay que instalar un binario
externo. Para el alcance de la herramienta, pandoc era una caja negra innecesaria.

---

## 2026-10-29 — Los títulos bajan un nivel

**Contexto:** Word suele tener un `Title` y varios `Heading 1`. Traducidos literalmente
producen varios H1 seguidos.
**Opciones:** mapear uno a uno, o desplazar un nivel.
**Decisión:** `Title` → `#` y `Heading N` → `#` × (N+1), con tope en 6.
**Consecuencia:** el documento conserva un solo H1 y la jerarquía relativa se mantiene. Un
`Heading 6` comparte nivel con `Heading 5`, que es la limitación de Markdown.

---

## 2026-10-29 — Se recorre el cuerpo del documento, no la lista de párrafos

**Contexto:** `python-docx` ofrece `doc.paragraphs`, pero esa vista no incluye las tablas.
**Opciones:** usar `doc.paragraphs` y procesar las tablas aparte; recorrer `doc.element.body`.
**Decisión:** recorrer el cuerpo y despachar según el tipo de elemento XML.
**Consecuencia:** las tablas aparecen donde estaban. Con la otra opción quedaban todas al
final del archivo, que es un error silencioso y difícil de notar.

---

## 2026-10-29 — Cada bloque se separa con una línea en blanco

**Contexto:** la primera versión pegaba los elementos unos debajo de otros.
**Hallazgo:** una lista numerada justo debajo de una cita (`>`) se absorbía **dentro** de la
cita, y dejaba de ser una lista al renderizar.
**Decisión:** agrupar las líneas en bloques y unirlos con una línea en blanco; los ítems de
una misma lista se unen con un solo salto.
**Consecuencia:** el Markdown renderiza igual que el documento original. La separación no es
estética: cambia el significado.

---

## 2026-10-29 — Reimplementar la extracción en vez de usar MarkItDown

**Contexto:** la herramienta original se apoyaba en `MarkItDown`, que resuelve la extracción de
Word, PDF, PowerPoint, Excel e imágenes en una sola llamada.
**Opciones:** usar la misma librería en la demo; escribir la extracción por formato.
**Decisión:** escribirla por formato, con `python-docx`, `python-pptx` y `openpyxl`.
**Consecuencia:** más código propio y menos formatos cubiertos, pero control sobre cómo se
traduce la jerarquía —y sin ese control no se puede **validar** la estructura, que era
justamente la parte que se revisaba a mano. Con MarkItDown se gana cobertura y se pierde la
posibilidad de corregir el resultado. Para esta demo, el control valía más que la cobertura.

---

## 2026-10-29 — Un conversor por formato, con importación diferida

**Contexto:** tres formatos de entrada, y cada uno necesita una librería distinta y pesada.
**Opciones:** exigir las tres siempre; importar solo la que hace falta.
**Decisión:** cada conversor importa su librería dentro de la función, y si falta, avisa con el
comando exacto para instalarla.
**Consecuencia:** quien solo convierte Word no necesita tener `openpyxl` ni `python-pptx`
instalados. El error, cuando ocurre, dice qué hacer en vez de mostrar un `ModuleNotFoundError`.

---

## 2026-10-29 — El PDF se detecta pero no se convierte

**Contexto:** el pedido original incluía `.pdf`, y en las carpetas había PDF.
**Decisión:** no convertirlo, pero **avisar** cuando aparece uno en la carpeta.
**Consecuencia:** un PDF guarda la posición visual del texto, no su estructura. Cualquier
extracción devuelve párrafos sueltos sin saber qué era un título ni dónde empezaba una tabla.
Es mejor decirlo que entregar un resultado silenciosamente malo.
