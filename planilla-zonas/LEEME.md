# Separar planilla por zonas (LibreOffice Calc)

Un botón para LibreOffice que deja la planilla del mail lista para imprimir
separada por zonas, sin hacerlo a mano.

> Es independiente de la app de Precintos.

## Qué hace al apretar el botón

1. Ordena todas las filas por la columna **zona** (la busca por el título;
   si no la encuentra usa la columna G). Ordena por el número de zona, así
   "ZONA - 10" va después de "ZONA - 9". Dentro de cada zona respeta el orden
   original de la planilla.
2. Pone toda la planilla en **letra 14** y saca lo que se repite, para que
   la letra salga más grande en el papel:
   - En la columna de **lugar**, el código repetido entre paréntesis y el
     final que es igual en todas las filas. Por ejemplo,
     `10068 - TERMINAL 4 (10068) ( BS.AS.(CAPITAL) (001))` queda como
     `10068 - TERMINAL 4`. Si el final cambia entre filas, lo deja.
   - En las **fechas**, si todas son del mismo año, las muestra sin el año
     (`09/10 19:00`).
3. Hoja **horizontal**, márgenes chicos (0,7 cm a los costados) y **sin encabezado ni pie de
   página**: cada hoja muestra solo la tabla.
4. **Cada fila en una sola línea**: las columnas toman el ancho justo de sus
   datos (los títulos largos van en dos líneas) y, si no entra todo a lo
   ancho, la impresión se achica solo lo necesario.
5. Hace que **cada zona empiece en una hoja nueva** y **repite la fila de
   títulos** arriba de cada hoja.
6. Abre la **vista previa** para que imprimas todo de una sola vez.

## Instalación (una sola vez)

Hay dos formas; con una alcanza.

### Forma A: instalador (recomendada)

1. Descargá el archivo **`SepararPorZonas.oxt`** de esta carpeta.
2. Abrí LibreOffice y andá a **Herramientas → Gestor de extensiones**.
3. Tocá **Añadir…**, elegí `SepararPorZonas.oxt` y aceptá.
4. Cerrá LibreOffice por completo y volvé a abrirlo.

(En Windows también podés hacer doble clic en el `.oxt`).

### Forma B: pegar la macro a mano

1. Abrí LibreOffice Calc → **Herramientas → Macros → Editar macros…**
2. A la izquierda, en **Mis macros y diálogos → Standard**, hacé doble clic
   en **Module1** (si no existe: clic derecho en *Standard* → **Insertar →
   Módulo BASIC**).
3. Borrá lo que haya en la ventana, pegá todo el contenido de
   **`src/Zonas.bas`** y guardá (Ctrl+G o el disquete).
4. Para usarla: **Herramientas → Macros → Ejecutar macro… → Mis macros →
   Standard → Module1 → SepararPorZonas → Ejecutar**.
   Para tenerla a mano, agregale un botón o un atajo (ver abajo; con esta
   forma la macro aparece en **Mis macros → Standard → Module1**).

## Uso

1. Abrí el `.xls` que te mandan por mail.
2. Tocá el botón **Separar por zonas**. Está en una barra de herramientas nueva
   o en el menú **Herramientas → Complementos → Separar por zonas**.
3. Se abre la vista previa → **Imprimir**.

Si no ves el botón: **Ver → Barras de herramientas** y activá la barra nueva
que agregó el complemento.

### Atajo de teclado (opcional)

**Herramientas → Personalizar → Teclado**. En *Categoría* buscá
**Macros de LibreOffice → Macros de la aplicación → PlanillaZonas → Zonas**,
elegí **SepararPorZonas**, seleccioná una tecla libre (por ejemplo F9 con
Ctrl) y tocá **Modificar**.

## Notas

- El botón cambia la planilla abierta (el orden, la letra y los saltos de
  página), no el mail. Si querés conservarla así, guardala con
  **Guardar como**.
- Trabaja sobre la hoja que tengas abierta en pantalla en ese momento.

## Para modificarlo

El código está en `src/Zonas.bas`. Arriba de todo están los valores que se
pueden cambiar fácil (tamaño de letra, columna de zona por defecto).
Después de editarlo, generá de nuevo el instalador:

```
python3 construir.py
```
