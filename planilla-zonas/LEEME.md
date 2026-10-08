# Separar planilla por zonas (LibreOffice Calc)

Un botón para LibreOffice que deja la planilla del mail lista para imprimir
separada por zonas, sin hacerlo a mano.

> Es independiente de la app de Precintos.

## Qué hace al apretar el botón

1. Ordena todas las filas por la columna **zona** (la busca por el título;
   si no la encuentra usa la columna G). Dentro de cada zona respeta el orden
   original de la planilla.
2. Pone toda la planilla en **letra 14** y ajusta el ancho de las columnas
   para que no se corte nada.
3. Hace que **cada zona empiece en una hoja nueva** al imprimir.
4. **Repite la fila de títulos** arriba de cada hoja.
5. Si las columnas no entran en vertical, pone la hoja en **horizontal**.
   Si tampoco entran así, achica apenas lo necesario para que no se corte.
6. Abre la **vista previa** para que imprimas todo de una sola vez.

## Instalación (una sola vez)

1. Descargá el archivo **`SepararPorZonas.oxt`** de esta carpeta.
2. Abrí LibreOffice y andá a **Herramientas → Gestor de extensiones**.
3. Tocá **Añadir…**, elegí `SepararPorZonas.oxt` y aceptá.
4. Cerrá LibreOffice por completo y volvé a abrirlo.

(En Windows también podés hacer doble clic en el `.oxt`).

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
