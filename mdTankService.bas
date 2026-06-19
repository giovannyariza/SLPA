Attribute VB_Name = "mdTankService"
Option Explicit

' ---------------------------------------------------------------------------------------------------------
' MÓDULO: mdTankService
' DESCRIPCIÓN:
'   Servicio de acceso a Tablas de Aforo (ListObject) con caché en RAM.
'   Responsable exclusivo de: cargar tablas de calibración, mantener caché, y retornar volúmenes TOV.
'   Soporta múltiples tablas con niveles de fila variables por tanque.
' ---------------------------------------------------------------------------------------------------------

' Variables de módulo para persistencia en RAM
Private m_CachedData As Variant        ' Matriz con los volúmenes (DataBodyRange)
Private m_CachedHeaders As Variant     ' Matriz con los Tags (HeaderRowRange)
Private m_CurrentTableName As String   ' Identificador de la tabla en caché
Private m_TankRanges As Object         ' Scripting.Dictionary: Key=Tag, Value=Array(levelMin, levelMax)

''' <summary>
''' Recupera el volumen TOV de una Tabla de Aforo (ListObject) mediante indexación directa.
''' Si la tabla no está en caché, la carga automáticamente.
''' Valida el nivel contra el rango mínimo y máximo efectivo de cada tanque.
''' </summary>

Public Function GetVolumeFromTable(ByVal TargetTable As ListObject, ByVal TankTag As String, ByVal Level_mm As Long) As Double
  On Error GoTo ErrHandler

  ' Nivel negativo no es válido
  If Level_mm < 0 Then Exit Function

  If m_CurrentTableName <> TargetTable.Name Or IsEmpty(m_CachedData) Then
    LoadTableToCache TargetTable
  End If

  ' Si el caché quedó vacío (tabla sin datos), salir
  If IsEmpty(m_CachedData) Then Exit Function

  Dim colIndex As Variant
  colIndex = Application.Match(TankTag, m_CachedHeaders, 0)

  If IsError(colIndex) Then
    Debug.Print "TankService: Tag '" & TankTag & "' no encontrado en la tabla " & TargetTable.Name
    Exit Function
  End If

  ' VALIDACIÓN DE RANGO POR TANQUE
  Dim lvlMin As Long, lvlMax As Long
  lvlMin = 0
  lvlMax = UBound(m_CachedData, 1) - 1

  If Not m_TankRanges Is Nothing And m_TankRanges.Exists(TankTag) Then
    lvlMin = m_TankRanges(TankTag)(0)
    lvlMax = m_TankRanges(TankTag)(1)
  End If

  If Level_mm < lvlMin Or Level_mm > lvlMax Then
    Debug.Print "TankService: Nivel " & Level_mm & "mm fuera de rango [" & lvlMin & "-" & lvlMax & "] para " & TankTag
    Exit Function
  End If

  Dim targetRow As Long
  targetRow = Level_mm + 1

  ' Validación de límites físicos de la tabla
  If targetRow < 1 Or targetRow > UBound(m_CachedData, 1) Then
    Debug.Print "TankService: Nivel " & Level_mm & "mm fuera de rango para " & TankTag
    Exit Function
  End If

  GetVolumeFromTable = m_CachedData(targetRow, colIndex)

  Exit Function
ErrHandler:
  Debug.Print "Error en mdTankService.GetVolumeFromTable: " & Err.Description
  GetVolumeFromTable = 0
End Function

''' <summary>
''' Carga el ListObject completo a matrices de memoria (RAM).
''' Adicionalmente, escanea cada columna para determinar el nivel mínimo y máximo
''' con datos efectivos para cada tanque.
''' </summary>

Private Sub LoadTableToCache(ByVal targetTable As ListObject)
  m_CurrentTableName = targetTable.Name
  m_CachedHeaders = targetTable.HeaderRowRange.Value

  If targetTable.ListRows.Count > 0 Then
    m_CachedData = targetTable.DataBodyRange.Value
    Debug.Print "TankService: Cache actualizado - Tabla '" & targetTable.Name & "' (" & UBound(m_CachedData, 1) & " filas)."
  Else
    m_CachedData = Empty
    Debug.Print "TankService: Cache actualizado - Tabla '" & targetTable.Name & "' (sin datos)."
  End If

  ' ESCANEO DE RANGOS POR TANQUE
  ' Para cada columna (tanque), encontrar la primera y última fila con datos
  Set m_TankRanges = CreateObject("Scripting.Dictionary")
  m_TankRanges.CompareMode = 1 ' TextCompare

  Dim totalRows As Long
  totalRows = UBound(m_CachedData, 1)

  Dim totalCols As Long
  totalCols = UBound(m_CachedData, 2)

  Dim columnIndex As Long, rowIndex As Long, rowIndexFromBottom As Long
  Dim tagName As String
  Dim foundMin As Long, foundMax As Long
  Dim hasData As Boolean

  For columnIndex = 1 To totalCols
    foundMin = -1
    foundMax = -1
    hasData = False

    ' Obtener el nombre del tanque desde los encabezados
    If IsArray(m_CachedHeaders) Then
      tagName = CStr(m_CachedHeaders(1, columnIndex))
    Else
      tagName = CStr(m_CachedHeaders)
    End If

    If Trim(tagName) = "" Then GoTo NextCol

    ' Recorrer filas de arriba hacia abajo para encontrar el primer dato
    For rowIndex = 1 To totalRows
      If Not IsEmpty(m_CachedData(rowIndex, columnIndex)) And IsNumeric(m_CachedData(rowIndex, columnIndex)) And m_CachedData(rowIndex, columnIndex) <> 0 Then
        foundMin = rowIndex - 1  ' Convertir índice de matriz (base 1) a nivel mm (base 0)
        hasData = True
        Exit For
      End If
    Next rowIndex

    ' Recorrer filas de abajo hacia arriba para encontrar el último dato
    If hasData Then
      For rowIndexFromBottom = totalRows To 1 Step -1
        If Not IsEmpty(m_CachedData(rowIndexFromBottom, columnIndex)) And IsNumeric(m_CachedData(rowIndexFromBottom, columnIndex)) And m_CachedData(rowIndexFromBottom, columnIndex) <> 0 Then
          foundMax = rowIndexFromBottom - 1
          Exit For
        End If
      Next rowIndexFromBottom
    End If

    ' Registrar el rango si se encontraron datos
    If hasData And foundMin >= 0 And foundMax >= 0 Then
      m_TankRanges.Add tagName, Array(foundMin, foundMax)
    End If

NextCol:
  Next columnIndex

  Debug.Print "TankService: Cache actualizado - Tabla '" & m_CurrentTableName & "' (" & totalRows & " filas, " & m_TankRanges.Count & " tanques con datos)."
End Sub

''' <summary>
''' Libera la memoria ocupada por el cache. Util al cerrar el proceso.
''' </summary>

Public Sub ClearCache()
  m_CachedData = Empty
  m_CachedHeaders = Empty
  m_CurrentTableName = ""
End Sub

' =========================================================================================================
' FUNCIONES DE CARGA DE CONFIGURACION DE TANQUES
' =========================================================================================================

''' <summary>
''' Carga un objeto clsTank desde tbl_Tanques por su Tag.
''' Usa los valores de la tabla para configurar todas las propiedades del tanque,
''' incluyendo el fluido contenido.
''' </summary>

Public Function LoadTankFromTable(ByVal TankTag As String) As clsTank
  On Error GoTo ErrHandler

  Dim tankTable As ListObject
  On Error Resume Next
  Set tankTable = ThisWorkbook.Sheets("cnf_Tanques").ListObjects("tbl_Tanques")
  On Error GoTo 0

  If tankTable Is Nothing Then
    Debug.Print "ERROR: No se encontro " & "tbl_Tanques"
    Exit Function
  End If

  If tankTable.ListRows.Count = 0 Then
    Debug.Print "ADVERTENCIA: " & "tbl_Tanques" & " esta vacia."
    Exit Function
  End If

  Dim colMap As Object
  Set colMap = CreateObject("Scripting.Dictionary")
  colMap.CompareMode = 1

  Dim tableColumn As ListColumn
  For Each tableColumn In tankTable.ListColumns
    colMap(tableColumn.Name) = tableColumn.Index
  Next tableColumn

  Dim tableRow As ListRow
  Dim tagKey As String
  For Each tableRow In tankTable.ListRows
    tagKey = GetCellValue(tableRow, colMap, "Tag", "")
    If UCase(Trim(tagKey)) = UCase(Trim(TankTag)) Then
      Set LoadTankFromTable = BuildTankFromRow(tableRow, colMap, tagKey)
      Exit Function
    End If
  Next tableRow

  Debug.Print "ADVERTENCIA: Tanque '" & TankTag & "' no encontrado en " & "tbl_Tanques"
  Exit Function

ErrHandler:
  Debug.Print "Error en LoadTankFromTable: " & Err.Description
  Set LoadTankFromTable = Nothing
End Function

''' <summary>
''' Carga la configuracion de tanques desde tbl_Tanques como Dictionary de objetos clsTank.
''' Si la tabla no existe, retorna un diccionario vacio con un aviso.
''' </summary>

Public Function LoadTankConfigs() As Object
  Set LoadTankConfigs = CreateObject("Scripting.Dictionary")
  LoadTankConfigs.CompareMode = 1

  Dim tankTable As ListObject
  On Error Resume Next
  Set tankTable = ThisWorkbook.Sheets("cnf_Tanques").ListObjects("tbl_Tanques")
  On Error GoTo 0

  If tankTable Is Nothing Then
    Debug.Print "ADVERTENCIA: No se encontro " & "tbl_Tanques" & ". Se usaran valores por defecto."
    Exit Function
  End If

  If tankTable.ListRows.Count = 0 Then Exit Function

  Dim colMap As Object
  Set colMap = CreateObject("Scripting.Dictionary")
  colMap.CompareMode = 1

  Dim tableColumn As ListColumn
  For Each tableColumn In tankTable.ListColumns
    colMap(tableColumn.Name) = tableColumn.Index
  Next tableColumn

  Dim tableRow As ListRow
  Dim tagKey As String

  For Each tableRow In tankTable.ListRows
    tagKey = GetCellValue(tableRow, colMap, "Tag", "")
    If tagKey = "" Then GoTo NextTank

    Set LoadTankConfigs(tagKey) = BuildTankFromRow(tableRow, colMap, tagKey)

NextTank:
  Next tableRow
End Function

''' <summary>
''' Helper interno: Construye un clsTank desde una fila de tabla y mapa de columnas.
''' </summary>

Private Function BuildTankFromRow(ByRef tankRow As ListRow, ByRef colMap As Object, ByVal tagKey As String) As clsTank
  Dim newTank As New clsTank
  newTank.Tag = tagKey
  newTank.Description = GetCellValue(tankRow, colMap, "Description", "")
  newTank.System = GetCellValue(tankRow, colMap, "System", "")
  newTank.Service = GetCellValue(tankRow, colMap, "Service", "")
  newTank.Status = GetCellValue(tankRow, colMap, "Status", "OP")

  Set newTank.Fluid = New clsFluid
  newTank.Fluid.Name = GetCellValue(tankRow, colMap, "FluidName", tagKey & "_Fluido")
  newTank.Fluid.Description = GetCellValue(tankRow, colMap, "FluidDescription", "")
  newTank.Fluid.FluidType = GetCellFluidType(tankRow, colMap, "FluidType", CRD)
  newTank.Fluid.TypicalAPI = GetCellNumeric(tankRow, colMap, "API60", 35)
  newTank.Fluid.TypicalViscosity = GetCellNumeric(tankRow, colMap, "Viscosity", 0)
  newTank.Fluid.ReferenceTemperature = GetCellNumeric(tankRow, colMap, "ReferenceTemperature", 60)
  newTank.Fluid.ReferencePressure = GetCellNumeric(tankRow, colMap, "ReferencePressure", 0)

  newTank.Material = GetCellMaterial(tankRow, colMap, "Material", MCrbn)
  newTank.ShellThickness = GetCellNumeric(tankRow, colMap, "ShellThickness", 0)
  newTank.RoofType = GetCellValue(tankRow, colMap, "RoofType", "")
  newTank.FloorType = GetCellValue(tankRow, colMap, "FloorType", "")

  newTank.NominalCapacity = GetCellNumeric(tankRow, colMap, "NominalCapacity", 0)
  newTank.Diameter = GetCellNumeric(tankRow, colMap, "Diameter", 0)
  newTank.EffectiveHeight = GetCellNumeric(tankRow, colMap, "EffectiveHeight", 0)
  newTank.ReferenceHeight = GetCellNumeric(tankRow, colMap, "ReferenceHeight", 0)
  newTank.SafeFillLevel = GetCellNumeric(tankRow, colMap, "SafeFillLevel", 0)
  newTank.SafePumpLevel = GetCellNumeric(tankRow, colMap, "SafePumpLevel", 0)

  newTank.IsThermalInsulated = GetCellBool(tankRow, colMap, "IsThermalInsulated", False)
  newTank.HasFloatingRoof = GetCellBool(tankRow, colMap, "HasFloatingRoof", False)
  newTank.IsTableNetOfRoof = GetCellBool(tankRow, colMap, "IsTableNetOfRoof", False)

  newTank.CriticalZoneLower = GetCellNumeric(tankRow, colMap, "CriticalZoneLower", 0)
  newTank.CriticalZoneUpper = GetCellNumeric(tankRow, colMap, "CriticalZoneUpper", 0)
  newTank.RoofWeight = GetCellNumeric(tankRow, colMap, "RoofWeight", 0)
  newTank.BaseDeduction = GetCellNumeric(tankRow, colMap, "BaseDeduction", 0)
  newTank.MinLevelDeduction = GetCellNumeric(tankRow, colMap, "MinLevelDeduction", 0)
  newTank.MaxLevelDeduction = GetCellNumeric(tankRow, colMap, "MaxLevelDeduction", 0)
  newTank.APICorrectionLT = GetCellNumeric(tankRow, colMap, "APICorrectionLT", 0)
  newTank.APICorrectionGT = GetCellNumeric(tankRow, colMap, "APICorrectionGT", 0)

  Set BuildTankFromRow = newTank
End Function

''' <summary>
''' Carga todas las tablas de aforo en el cache.
''' </summary>

Public Sub LoadStrappingTablesToCache()
  Dim targetWorksheet As Worksheet
  On Error Resume Next
  Set targetWorksheet = ThisWorkbook.Sheets("cnf_Tanques")
  On Error GoTo 0
  If targetWorksheet Is Nothing Then Exit Sub

  Dim tableObject As ListObject
  For Each tableObject In targetWorksheet.ListObjects
    If tableObject.ListRows.Count > 0 Then
      On Error Resume Next
      GetVolumeFromTable tableObject, tableObject.Name, 0
      On Error GoTo 0
    End If
  Next tableObject
  Debug.Print "Tablas de aforo cargadas en cache."
End Sub

''' <summary>
''' Busca la tabla de aforo cacheada para un tanque por su Tag.
''' </summary>

Public Function GetCachedStrappingTable(ByVal tag As String) As Object
  Dim targetWorksheet As Worksheet
  On Error Resume Next
  Set targetWorksheet = ThisWorkbook.Sheets("cnf_Tanques")
  On Error GoTo 0
  If targetWorksheet Is Nothing Then Exit Function

  Dim tableObject As ListObject
  For Each tableObject In targetWorksheet.ListObjects
    Dim listColumn As ListColumn
    For Each listColumn In tableObject.ListColumns
      If listColumn.Name = tag Or listColumn.Name = UCase(tag) Then
        Set GetCachedStrappingTable = tableObject
        Exit Function
      End If
    Next listColumn
  Next tableObject
End Function

' =========================================================================================================
' HELPERS DE LECTURA DE CELDA (privados al modulo)
' =========================================================================================================

Private Function GetCellValue(ByRef tableRow As ListRow, ByRef colMap As Object, ByVal colName As String, ByVal defaultVal As Variant) As Variant
  If colMap.Exists(colName) Then
    Dim cellValue As Variant
    cellValue = tableRow.Range.Cells(1, colMap(colName) - tableRow.Range.Column + 1).Value
    If Not IsEmpty(cellValue) And Not IsNull(cellValue) Then
      GetCellValue = cellValue
      Exit Function
    End If
  End If
  GetCellValue = defaultVal
End Function

Private Function GetCellNumeric(ByRef tableRow As ListRow, ByRef colMap As Object, ByVal colName As String, ByVal defaultVal As Double) As Double
  Dim cellValue As Variant
  cellValue = GetCellValue(tableRow, colMap, colName, defaultVal)
  If IsNumeric(cellValue) Then
    GetCellNumeric = CDbl(cellValue)
  Else
    GetCellNumeric = defaultVal
  End If
End Function

Private Function GetCellBool(ByRef tableRow As ListRow, ByRef colMap As Object, ByVal colName As String, ByVal defaultVal As Boolean) As Boolean
  Dim cellValue As Variant
  cellValue = GetCellValue(tableRow, colMap, colName, defaultVal)
  If VarType(cellValue) = vbBoolean Then
    GetCellBool = CBool(cellValue)
  ElseIf VarType(cellValue) = vbString Then
    GetCellBool = (LCase(Trim(CStr(cellValue))) = "si" Or LCase(Trim(CStr(cellValue))) = "true" Or LCase(Trim(CStr(cellValue))) = "verdadero" Or cellValue = 1)
  ElseIf IsNumeric(cellValue) Then
    GetCellBool = (CDbl(cellValue) <> 0)
  Else
    GetCellBool = defaultVal
  End If
End Function

Private Function GetCellFluidType(ByRef tableRow As ListRow, ByRef colMap As Object, ByVal colName As String, ByVal defaultVal As eTypeLiq) As eTypeLiq
  Dim cellValue As Variant
  cellValue = GetCellValue(tableRow, colMap, colName, defaultVal)
  If IsNumeric(cellValue) Then
    GetCellFluidType = CLng(cellValue)
  Else
    Dim sanitizedValue As String
    sanitizedValue = UCase(Trim(CStr(cellValue)))
    Select Case sanitizedValue
      Case "CRD", "CRUDO": GetCellFluidType = CRD
      Case "REF", "REFINADO": GetCellFluidType = REF
      Case "LUB", "LUBRICANTE": GetCellFluidType = LUB
      Case Else: GetCellFluidType = defaultVal
    End Select
  End If
End Function

Private Function GetCellMaterial(ByRef tableRow As ListRow, ByRef colMap As Object, ByVal colName As String, ByVal defaultVal As eMtrl) As eMtrl
  Dim cellValue As Variant
  cellValue = GetCellValue(tableRow, colMap, colName, defaultVal)
  If IsNumeric(cellValue) Then
    GetCellMaterial = CLng(cellValue)
  Else
    Dim sanitizedValue As String
    sanitizedValue = UCase(Trim(CStr(cellValue)))
    Select Case sanitizedValue
      Case "MCBN", "ACERO CARBON", "CARBON": GetCellMaterial = MCrbn
      Case "ST304", "304": GetCellMaterial = St304
      Case "ST316", "316": GetCellMaterial = St316
      Case "ST4PH", "4PH": GetCellMaterial = St4PH
      Case Else: GetCellMaterial = defaultVal
    End Select
  End If
End Function
