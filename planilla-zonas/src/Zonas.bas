REM ===================================================================
REM  Planilla por zonas  -  macro para LibreOffice Calc
REM
REM  Deja la planilla abierta lista para imprimir separada por zona:
REM   1. Ordena las filas por la columna "zona".
REM   2. Pone la letra en 14 y ajusta el ancho de las columnas.
REM   3. Inserta un salto de pagina antes de cada zona nueva.
REM   4. Repite la fila de titulos arriba de cada hoja impresa.
REM   5. Deja la hoja en horizontal y sin encabezado ni pie de pagina;
REM      si las columnas no entran, achica lo justo para que no se corte nada.
REM ===================================================================

Option Explicit

Const TAMANO_LETRA = 14
Const COLUMNA_ZONA_POR_DEFECTO = 6   ' 0 = A, 1 = B ... 6 = G
Const MARGEN_COLUMNA = 200           ' aire extra por columna (centesimas de mm)

Sub SepararPorZonas
	Dim oDoc As Object
	Dim nZonas As Long

	oDoc = ThisComponent
	If IsNull(oDoc) Then
		MsgBox "Abrí primero la planilla que querés separar.", 48, "Separar por zonas"
		Exit Sub
	End If
	If Not oDoc.supportsService("com.sun.star.sheet.SpreadsheetDocument") Then
		MsgBox "Abrí primero la planilla que querés separar.", 48, "Separar por zonas"
		Exit Sub
	End If

	nZonas = PrepararZonas(oDoc, oDoc.CurrentController.ActiveSheet)

	Select Case nZonas
	Case -1
		MsgBox "La hoja está vacía.", 48, "Separar por zonas"
	Case -2
		MsgBox "No encontré la columna ""zona"" (ni en los títulos ni en la columna G).", 48, "Separar por zonas"
	Case Else
		MsgBox "Listo: " & nZonas & " zona(s), cada una empieza en una hoja nueva." & Chr(10) & _
		       "Ahora se abre la vista previa; desde ahí podés imprimir.", 64, "Separar por zonas"
		Dim oDisp As Object
		oDisp = createUnoService("com.sun.star.frame.DispatchHelper")
		oDisp.executeDispatch(oDoc.CurrentController.Frame, ".uno:PrintPreview", "", 0, Array())
	End Select
End Sub

REM Devuelve la cantidad de zonas, -1 si la hoja esta vacia
REM o -2 si no hay columna de zona.
Function PrepararZonas(oDoc As Object, oHoja As Object) As Long
	Dim oCursor As Object, oRango As Object, oCol As Object, oEstilo As Object
	Dim ultFila As Long, ultCol As Long, colZona As Long, primeraFila As Long
	Dim tieneTitulos As Boolean
	Dim f As Long, c As Long, nZonas As Long, anchoTotal As Long

	' --- Rango con datos (sin filas vacias al final) ---
	oCursor = oHoja.createCursor()
	oCursor.gotoEndOfUsedArea(False)
	ultFila = oCursor.RangeAddress.EndRow
	ultCol = oCursor.RangeAddress.EndColumn
	Do While ultFila >= 0
		If Not FilaVacia(oHoja, ultFila, ultCol) Then Exit Do
		ultFila = ultFila - 1
	Loop
	If ultFila < 0 Then
		PrepararZonas = -1
		Exit Function
	End If

	colZona = BuscarColumnaZona(oHoja, ultCol)
	If colZona < 0 Then
		PrepararZonas = -2
		Exit Function
	End If

	' Si la primera fila tiene texto en la columna de zona, son los titulos
	Dim sPrimera As String
	sPrimera = Trim(oHoja.getCellByPosition(colZona, 0).getString())
	tieneTitulos = (sPrimera <> "" And Not IsNumeric(sPrimera))
	If tieneTitulos Then primeraFila = 1 Else primeraFila = 0
	If ultFila < primeraFila Then
		PrepararZonas = -1
		Exit Function
	End If

	oRango = oHoja.getCellRangeByPosition(0, 0, ultCol, ultFila)

	' --- 1. Ordenar por zona ---
	Dim aCampos(0) As New com.sun.star.table.TableSortField
	aCampos(0).Field = colZona
	aCampos(0).IsAscending = True
	Dim aOrden(1) As New com.sun.star.beans.PropertyValue
	aOrden(0).Name = "SortFields"
	aOrden(0).Value = aCampos()
	aOrden(1).Name = "ContainsHeader"
	aOrden(1).Value = tieneTitulos
	oRango.sort(aOrden())

	' --- 2. Letra 14 y columnas a medida ---
	oRango.CharHeight = TAMANO_LETRA
	oRango.CharHeightAsian = TAMANO_LETRA
	oRango.CharHeightComplex = TAMANO_LETRA
	oRango.Rows.OptimalHeight = True
	oRango.Columns.OptimalWidth = True
	anchoTotal = 0
	For c = 0 To ultCol
		oCol = oHoja.Columns.getByIndex(c)
		oCol.Width = oCol.Width + MARGEN_COLUMNA
		anchoTotal = anchoTotal + oCol.Width
	Next c

	' --- 3. Salto de pagina en cada cambio de zona ---
	oHoja.removeAllManualPageBreaks()
	nZonas = 1
	For f = primeraFila + 1 To ultFila
		If Trim(oHoja.getCellByPosition(colZona, f).getString()) <> _
		   Trim(oHoja.getCellByPosition(colZona, f - 1).getString()) Then
			oHoja.Rows.getByIndex(f).IsStartOfNewPage = True
			nZonas = nZonas + 1
		End If
	Next f

	' --- 4. Area de impresion y titulos repetidos ---
	Dim aAreas(0) As New com.sun.star.table.CellRangeAddress
	aAreas(0) = oRango.RangeAddress
	oHoja.setPrintAreas(aAreas())
	If tieneTitulos Then
		Dim oTitulos As New com.sun.star.table.CellRangeAddress
		oTitulos.Sheet = oRango.RangeAddress.Sheet
		oTitulos.StartColumn = 0
		oTitulos.EndColumn = ultCol
		oTitulos.StartRow = 0
		oTitulos.EndRow = 0
		oHoja.setTitleRows(oTitulos)
	End If
	oHoja.setPrintTitleRows(tieneTitulos)

	' --- 5. Hoja horizontal, solo la tabla (sin encabezado ni pie) ---
	Dim ladoCorto As Long, ladoLargo As Long
	oEstilo = oDoc.StyleFamilies.getByName("PageStyles").getByName(oHoja.PageStyle)
	ladoCorto = oEstilo.Width
	ladoLargo = oEstilo.Height
	If ladoCorto > ladoLargo Then
		ladoCorto = oEstilo.Height
		ladoLargo = oEstilo.Width
	End If
	oEstilo.IsLandscape = True
	oEstilo.Width = ladoLargo
	oEstilo.Height = ladoCorto
	oEstilo.HeaderIsOn = False
	oEstilo.FooterIsOn = False

	oEstilo.ScaleToPages = 0
	oEstilo.ScaleToPagesX = 0
	oEstilo.ScaleToPagesY = 0
	oEstilo.PageScale = 100
	If anchoTotal > ladoLargo - oEstilo.LeftMargin - oEstilo.RightMargin Then
		' No entra a lo ancho: achicar solo lo necesario para que no se corte
		oEstilo.ScaleToPagesX = 1
		oEstilo.ScaleToPagesY = 0
	End If

	PrepararZonas = nZonas
End Function

Function BuscarColumnaZona(oHoja As Object, ultCol As Long) As Long
	Dim c As Long
	For c = 0 To ultCol
		If LCase(Trim(oHoja.getCellByPosition(c, 0).getString())) = "zona" Then
			BuscarColumnaZona = c
			Exit Function
		End If
	Next c
	If ultCol >= COLUMNA_ZONA_POR_DEFECTO Then
		BuscarColumnaZona = COLUMNA_ZONA_POR_DEFECTO
	Else
		BuscarColumnaZona = -1
	End If
End Function

Function FilaVacia(oHoja As Object, f As Long, ultCol As Long) As Boolean
	Dim c As Long
	For c = 0 To ultCol
		If oHoja.getCellByPosition(c, f).getType() <> com.sun.star.table.CellContentType.EMPTY Then
			FilaVacia = False
			Exit Function
		End If
	Next c
	FilaVacia = True
End Function
