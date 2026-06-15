Attribute VB_Name = "mdTankService"
Option Explicit

' ---------------------------------------------------------------------------------------------------------
' MODULO: mdTankService
' DESCRIPCION:
'   Servicio de acceso a Tablas de Aforo (ListObject) con cache en RAM.
'   Responsable exclusivo de: cargar tablas de calibracion, mantener cache, retornar volumenes TOV,
'   y cargar configuracion de tanques desde tbl_Tanques.
'   Soporta multiples tablas con niveles de fila variables por tanque.
' ---------------------------------------------------------------------------------------------------------

Private Const TANQUES_SHEET As String = "cnf_Tanques"
Private Const TANQUES_TABLE As String = "tbl_Tanques"

' Variables de modulo para persistencia en RAM
Private m_CachedData As Variant        ' Matriz con los volumenes (DataBodyRange)
Private m_CachedHeaders As Variant     ' Matriz con los Tags (HeaderRowRange)
Private m_CurrentTableName As String   ' Identificador de la tabla en cache
Private m_TankRanges As Object         ' Scripting.Dictionary: Key=Tag, Value=Array(levelMin, levelMax)

' Estructura para mantener configuracion de un tanque
Public Type TankConfig
  Fluid As clsFluid
  Material As eMtrl
  Diameter As Double
  ShellThickness As Double
  HasFloatingRoof As Boolean
  RoofWeight As Double
  IsTableNetOfRoof As Boolean
  APICorrectionGT As Double
  APICorrectionLT As Double
  MinLevelDeduction As Double
  MaxLevelDeduction As Double
  BaseDeduction As Double
  IsLoaded As Boolean
End Type

''' <summary>
''' Recupera el volumen TOV de una Tabla de Aforo (ListObject) mediante indexacion directa.
''' Si la tabla no esta en cache, la carga automaticamente.
''' Valida el nivel contra el rango minimo y maximo efectivo de cada tanque.
''' </summary>

Public Function GetVolumeFromTable(ByVal TargetTable As ListObject, ByVal TankTag As String, ByVal Level_mm As Long) As Double
  On Error GoTo ErrHandler

  If Level_mm < 0 Then Exit Function

  If m_CurrentTableName <> TargetTable.Name Or IsEmpty(m_CachedData) Then
    LoadTableToCache TargetTable
  End If

  If IsEmpty(m_CachedData) Then Exit Function

  Dim colIndex As Variant
  colIndex = Application.Match(TankTag, m_CachedHeaders, 0)

  If IsError(colIndex) Then
    Debug.Print "TankService: Tag '" & TankTag & "' no encontrado en la tabla " & TargetTable.Name
    Exit Function
  End If

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

  GetVolumeFromTable = m_CachedData(targetRow, colIndex)

  Exit Function
ErrHandler:
  Debug.Print "Error en mdTankService.GetVolumeFromTable: " & Err.Description
  GetVolumeFromTable = 0
End Function

''' <summary>
''' Carga el ListObject completo a matrices de memoria (RAM).
''' Adicionalmente, escanea cada columna para determinar el nivel minimo y maximo
''' con datos efectivos para cada tanque.
''' </summary>

Private Sub LoadTableToCache(ByVal lo As ListObject)
  m_CurrentTableName = lo.Name
  m_CachedHeaders = lo.HeaderRowRange.Value

  If lo.ListRows.Count > 0 Then
    m_CachedData = lo.DataBodyRange.Value
  Else
    m_CachedData = Empty
    Debug.Print "TankService: Cache actualizado - Tabla '" & m_CurrentTableName & "' (sin datos)."
    Exit Sub
  End If

  ' ESCANEO DE RANGOS POR TANQUE
  Set m_TankRanges = CreateObject("Scripting.Dictionary")
  m_TankRanges.CompareMode = 1

  Dim totalRows As Long
  totalRows = UBound(m_CachedData, 1)
  Dim totalCols As Long
  totalCols = UBound(m_CachedData, 2)

  Dim c As Long, r As Long
  Dim tagName As String
  Dim foundMin As Long, foundMax As Long
  Dim hasData As Boolean

  For c = 1 To totalCols
    foundMin = -1
    foundMax = -1
    hasData = False

    If IsArray(m_CachedHeaders) Then
      tagName = CStr(m_CachedHeaders(1, c))
    Else
      tagName = CStr(m_CachedHeaders)
    End If

    If Trim(tagName) = "" Then GoTo NextCol

    For r = 1 To totalRows
      If Not IsEmpty(m_CachedData(r, c)) And IsNumeric(m_CachedData(r, c)) And m_CachedData(r, c) <> 0 Then
        foundMin = r - 1
        hasData = True
        Exit For
      End If
    Next r

    If hasData Then
      For r = totalRows To 1 Step -1
        If Not IsEmpty(m_CachedData(r, c)) And IsNumeric(m_CachedData(r, c)) And m_CachedData(r, c) <> 0 Then
          foundMax = r - 1
          Exit For
        End If
      Next r
    End If

    If hasData And foundMin >= 0 And foundMax >= 0 Then
      m_TankRanges.Add tagName, Array(foundMin, foundMax)
    End If

NextCol:
  Next c

  Debug.Print "TankService: Cache actualizado - Tabla '" & m_CurrentTableName & "' (" & totalRows & " filas, " & m_TankRanges.Count & " tanques con datos)."
End Sub

''' <summary>
''' Libera la memoria ocupada por el cache. Util al cerrar el proceso.
''' </summary>

Public Sub ClearCache()
  m_CachedData = Empty
  m_CachedHeaders = Empty
  m_CurrentTableName = ""
  If Not m_TankRanges Is Nothing Then
    m_TankRanges.RemoveAll
    Set m_TankRanges = Nothing
  End If
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

  Dim lo As ListObject
  On Error Resume Next
  Set lo = ThisWorkbook.Sheets(TANQUES_SHEET).ListObjects(TANQUES_TABLE)
  On Error GoTo 0

  If lo Is Nothing Then
    Debug.Print "ERROR: No se encontro " & TANQUES_TABLE
    Exit Function
  End If

  If lo.ListRows.Count = 0 Then
    Debug.Print "ADVERTENCIA: " & TANQUES_TABLE & " esta vacia."
    Exit Function
  End If

  Dim colMap As Object
  Set colMap = CreateObject("Scripting.Dictionary")
  colMap.CompareMode = 1

  Dim lc As ListColumn
  For Each lc In lo.ListColumns
    colMap(lc.Name) = lc.Index
  Next lc

  Dim lr As ListRow
  Dim tagKey As String
  For Each lr In lo.ListRows
    tagKey = GetCellValue(lr, colMap, "Tag", "")
    If UCase(Trim(tagKey)) = UCase(Trim(TankTag)) Then
      Dim t As New clsTank
      t.Tag = tagKey
      t.Description = GetCellValue(lr, colMap, "Description", "")
      t.System = GetCellValue(lr, colMap, "System", "")
      t.Service = GetCellValue(lr, colMap, "Service", "")
      t.Status = GetCellValue(lr, colMap, "Status", "OP")

      Set t.Fluid = New clsFluid
      t.Fluid.Name = GetCellValue(lr, colMap, "FluidName", tagKey & "_Fluido")
      t.Fluid.Description = GetCellValue(lr, colMap, "FluidDescription", "")
      t.Fluid.FluidType = GetCellFluidType(lr, colMap, "FluidType", CRD)
      t.Fluid.TypicalAPI = GetCellNumeric(lr, colMap, "API60", 35)
      t.Fluid.TypicalViscosity = GetCellNumeric(lr, colMap, "Viscosity", 0)
      t.Fluid.ReferenceTemperature = GetCellNumeric(lr, colMap, "ReferenceTemperature", 60)
      t.Fluid.ReferencePressure = GetCellNumeric(lr, colMap, "ReferencePressure", 0)

      t.Material = GetCellMaterial(lr, colMap, "Material", MCrbn)
      t.ShellThickness = GetCellNumeric(lr, colMap, "ShellThickness", 0)
      t.RoofType = GetCellValue(lr, colMap, "RoofType", "")
      t.FloorType = GetCellValue(lr, colMap, "FloorType", "")

      t.NominalCapacity = GetCellNumeric(lr, colMap, "NominalCapacity", 0)
      t.Diameter = GetCellNumeric(lr, colMap, "Diameter", 0)
      t.EffectiveHeight = GetCellNumeric(lr, colMap, "EffectiveHeight", 0)
      t.ReferenceHeight = GetCellNumeric(lr, colMap, "ReferenceHeight", 0)
      t.SafeFillLevel = GetCellNumeric(lr, colMap, "SafeFillLevel", 0)
      t.SafePumpLevel = GetCellNumeric(lr, colMap, "SafePumpLevel", 0)

      t.IsThermalInsulated = GetCellBool(lr, colMap, "IsThermalInsulated", False)
      t.HasFloatingRoof = GetCellBool(lr, colMap, "HasFloatingRoof", False)
      t.IsTableNetOfRoof = GetCellBool(lr, colMap, "IsTableNetOfRoof", False)

      t.CriticalZoneLower = GetCellNumeric(lr, colMap, "CriticalZoneLower", 0)
      t.CriticalZoneUpper = GetCellNumeric(lr, colMap, "CriticalZoneUpper", 0)
      t.RoofWeight = GetCellNumeric(lr, colMap, "RoofWeight", 0)
      t.BaseDeduction = GetCellNumeric(lr, colMap, "BaseDeduction", 0)
      t.MinLevelDeduction = GetCellNumeric(lr, colMap, "MinLevelDeduction", 0)
      t.MaxLevelDeduction = GetCellNumeric(lr, colMap, "MaxLevelDeduction", 0)
      t.APICorrectionLT = GetCellNumeric(lr, colMap, "APICorrectionLT", 0)
      t.APICorrectionGT = GetCellNumeric(lr, colMap, "APICorrectionGT", 0)

      t.StrappingTable = GetCachedStrappingTable(tagKey)

      Set LoadTankFromTable = t
      Exit Function
    End If
  Next lr

  Debug.Print "ADVERTENCIA: Tanque '" & TankTag & "' no encontrado en " & TANQUES_TABLE
  Exit Function

ErrHandler:
  Debug.Print "Error en LoadTankFromTable: " & Err.Description
  Set LoadTankFromTable = Nothing
End Function

''' <summary>
''' Carga la configuracion de tanques desde tbl_Tanques como Dictionary de TankConfig.
''' Si la tabla no existe, retorna un diccionario vacio con un aviso.
''' </summary>

Public Function LoadTankConfigs() As Object
  Set LoadTankConfigs = CreateObject("Scripting.Dictionary")
  LoadTankConfigs.CompareMode = 1

  Dim lo As ListObject
  On Error Resume Next
  Set lo = ThisWorkbook.Sheets(TANQUES_SHEET).ListObjects(TANQUES_TABLE)
  On Error GoTo 0

  If lo Is Nothing Then
    Debug.Print "ADVERTENCIA: No se encontro " & TANQUES_TABLE & ". Se usaran valores por defecto."
    Exit Function
  End If

  If lo.ListRows.Count = 0 Then Exit Function

  Dim colMap As Object
  Set colMap = CreateObject("Scripting.Dictionary")
  colMap.CompareMode = 1

  Dim lc As ListColumn
  For Each lc In lo.ListColumns
    colMap(lc.Name) = lc.Index
  Next lc

  Dim lr As ListRow
  Dim cfg As TankConfig
  Dim tagKey As String

  For Each lr In lo.ListRows
    tagKey = GetCellValue(lr, colMap, "Tag", "")
    If tagKey = "" Then GoTo NextTank

    cfg.IsLoaded = True

    Set cfg.Fluid = New clsFluid
    cfg.Fluid.FluidType = GetCellFluidType(lr, colMap, "FluidType", CRD)
    cfg.Fluid.TypicalAPI = GetCellNumeric(lr, colMap, "API60", 35)
    cfg.Fluid.ReferenceTemperature = GetCellNumeric(lr, colMap, "ReferenceTemperature", 60)
    cfg.Fluid.ReferencePressure = GetCellNumeric(lr, colMap, "ReferencePressure", 0)

    cfg.Material = GetCellMaterial(lr, colMap, "Material", MCrbn)
    cfg.Diameter = GetCellNumeric(lr, colMap, "Diameter", 0)
    cfg.ShellThickness = GetCellNumeric(lr, colMap, "ShellThickness", 0)
    cfg.HasFloatingRoof = GetCellBool(lr, colMap, "HasFloatingRoof", False)
    cfg.RoofWeight = GetCellNumeric(lr, colMap, "RoofWeight", 0)
    cfg.IsTableNetOfRoof = GetCellBool(lr, colMap, "IsTableNetOfRoof", False)
    cfg.APICorrectionGT = GetCellNumeric(lr, colMap, "APICorrectionGT", 0)
    cfg.APICorrectionLT = GetCellNumeric(lr, colMap, "APICorrectionLT", 0)
    cfg.MinLevelDeduction = GetCellNumeric(lr, colMap, "MinLevelDeduction", 0)
    cfg.MaxLevelDeduction = GetCellNumeric(lr, colMap, "MaxLevelDeduction", 0)
    cfg.BaseDeduction = GetCellNumeric(lr, colMap, "BaseDeduction", 0)

    LoadTankConfigs(tagKey) = cfg

NextTank:
  Next lr
End Function

''' <summary>
''' Carga todas las tablas de aforo en el cache.
''' </summary>

Public Sub LoadStrappingTablesToCache()
  Dim ws As Worksheet
  On Error Resume Next
  Set ws = ThisWorkbook.Sheets(TANQUES_SHEET)
  On Error GoTo 0
  If ws Is Nothing Then Exit Sub

  Dim lo As ListObject
  For Each lo In ws.ListObjects
    If lo.ListRows.Count > 0 Then
      On Error Resume Next
      GetVolumeFromTable lo, lo.Name, 0
      On Error GoTo 0
    End If
  Next lo
  Debug.Print "Tablas de aforo cargadas en cache."
End Sub

''' <summary>
''' Busca la tabla de aforo cacheada para un tanque por su Tag.
''' </summary>

Public Function GetCachedStrappingTable(ByVal tag As String) As Object
  Dim ws As Worksheet
  On Error Resume Next
  Set ws = ThisWorkbook.Sheets(TANQUES_SHEET)
  On Error GoTo 0
  If ws Is Nothing Then Exit Function

  Dim lo As ListObject
  For Each lo In ws.ListObjects
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
      If lc.Name = tag Or lc.Name = UCase(tag) Then
        Set GetCachedStrappingTable = lo
        Exit Function
      End If
    Next lc
  Next lo
End Function

' =========================================================================================================
' HELPERS DE LECTURA DE CELDA (privados al modulo)
' =========================================================================================================

Private Function GetCellValue(ByRef lr As ListRow, ByRef colMap As Object, ByVal colName As String, ByVal defaultVal As Variant) As Variant
  If colMap.Exists(colName) Then
    Dim v As Variant
    v = lr.Range.Cells(1, colMap(colName) - lr.Range.Column + 1).Value
    If Not IsEmpty(v) And Not IsNull(v) Then
      GetCellValue = v
      Exit Function
    End If
  End If
  GetCellValue = defaultVal
End Function

Private Function GetCellNumeric(ByRef lr As ListRow, ByRef colMap As Object, ByVal colName As String, ByVal defaultVal As Double) As Double
  Dim v As Variant
  v = GetCellValue(lr, colMap, colName, defaultVal)
  If IsNumeric(v) Then
    GetCellNumeric = CDbl(v)
  Else
    GetCellNumeric = defaultVal
  End If
End Function

Private Function GetCellBool(ByRef lr As ListRow, ByRef colMap As Object, ByVal colName As String, ByVal defaultVal As Boolean) As Boolean
  Dim v As Variant
  v = GetCellValue(lr, colMap, colName, defaultVal)
  If VarType(v) = vbBoolean Then
    GetCellBool = CBool(v)
  ElseIf VarType(v) = vbString Then
    GetCellBool = (LCase(Trim(CStr(v))) = "si" Or LCase(Trim(CStr(v))) = "true" Or LCase(Trim(CStr(v))) = "verdadero" Or v = 1)
  ElseIf IsNumeric(v) Then
    GetCellBool = (CDbl(v) <> 0)
  Else
    GetCellBool = defaultVal
  End If
End Function

Private Function GetCellFluidType(ByRef lr As ListRow, ByRef colMap As Object, ByVal colName As String, ByVal defaultVal As eTypeLiq) As eTypeLiq
  Dim v As Variant
  v = GetCellValue(lr, colMap, colName, defaultVal)
  If IsNumeric(v) Then
    GetCellFluidType = CLng(v)
  Else
    Dim s As String
    s = UCase(Trim(CStr(v)))
    Select Case s
      Case "CRD", "CRUDO": GetCellFluidType = CRD
      Case "REF", "REFINADO": GetCellFluidType = REF
      Case "LUB", "LUBRICANTE": GetCellFluidType = LUB
      Case Else: GetCellFluidType = defaultVal
    End Select
  End If
End Function

Private Function GetCellMaterial(ByRef lr As ListRow, ByRef colMap As Object, ByVal colName As String, ByVal defaultVal As eMtrl) As eMtrl
  Dim v As Variant
  v = GetCellValue(lr, colMap, colName, defaultVal)
  If IsNumeric(v) Then
    GetCellMaterial = CLng(v)
  Else
    Dim s As String
    s = UCase(Trim(CStr(v)))
    Select Case s
      Case "MCBN", "ACERO CARBON", "CARBON": GetCellMaterial = MCrbn
      Case "ST304", "304": GetCellMaterial = St304
      Case "ST316", "316": GetCellMaterial = St316
      Case "ST4PH", "4PH": GetCellMaterial = St4PH
      Case Else: GetCellMaterial = defaultVal
    End Select
  End If
End Function
