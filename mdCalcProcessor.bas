Attribute VB_Name = "mdCalcProcessor"
Option Explicit

' ---------------------------------------------------------------------------------------------------------
' MODULO: mdCalcProcessor
' DESCRIPCION:
'   Orquestador del flujo de liquidacion de tanques.
'   Lee datos de tbl_Registro, calcula factores API usando clsTank y mdAPICalcs,
'   y escribe resultados en tbl_Motor respetando el historico.
' ---------------------------------------------------------------------------------------------------------

Private Const REGISTRO_SHEET As String = "ent_Registro"
Private Const REGISTRO_TABLE As String = "tbl_Registro"
Private Const MOTOR_SHEET As String = "calc_Motor"
Private Const MOTOR_TABLE As String = "tbl_Motor"
Private Const TANQUES_SHEET As String = "cnf_Tanques"
Private Const TANQUES_TABLE As String = "tbl_Tanques"

' Indices de columnas en tbl_Registro (se resuelven dinamicamente)
Private idx_Fecha As Long
Private idx_Tag As Long
Private idx_Nivel As Long
Private idx_NivelAgua As Long
Private idx_TempLiq As Long
Private idx_TempAmb As Long
Private idx_Responsable As Long
Private idx_Observaciones As Long

' Estructura para mantener configuracion de un tanque
Private Type TankConfig
  API60 As Double
  FluidType As eTypeLiq
  Material As eMtrl
  Diameter As Double
  ShellThickness As Double
  HasFloatingRoof As Boolean
  RoofWeight As Double
  IsTableNetOfRoof As Boolean
  ReferenceAPIObs As Double
  APICorrectionGT As Double
  APICorrectionLT As Double
  MinLevelDeduction As Double
  MaxLevelDeduction As Double
  BaseDeduction As Double
  IsLoaded As Boolean
End Type

''' <summary>
''' Funcion principal: procesa registros pendientes de tbl_Registro y escribe en tbl_Motor.
''' </summary>
''' <param name="ForceMode">Si True, recalcula sobrescribiendo existentes. Si False (default), solo procesa nuevos.</param>

Public Sub ProcesarRegistros(Optional ByVal ForceMode As Boolean = False)
  On Error GoTo ErrHandler

  Dim startTime As Double
  startTime = Timer

  Debug.Print "======================================================================"
  Debug.Print "INICIANDO PROCESAMIENTO DE REGISTROS"
  Debug.Print "Modo: " & IIf(ForceMode, "Forzado (recalculo)", "Incremental (solo nuevos)")
  Debug.Print "======================================================================"

  ' 1. Cargar tabla de configuracion de tanques
  Dim tankConfigs As Object ' Dictionary of TankConfig
  Set tankConfigs = LoadTankConfigs()
  Debug.Print "Configuracion cargada: " & tankConfigs.Count & " tanques"

  ' 2. Cargar registros de entrada
  Dim regData As Variant
  regData = GetTableData(REGISTRO_SHEET, REGISTRO_TABLE)
  If IsEmpty(regData) Then
    Debug.Print "ERROR: No se encontraron datos en " & REGISTRO_TABLE
    Exit Sub
  End If
  Dim regRows As Long
  regRows = UBound(regData, 1)
  Debug.Print "Registros en " & REGISTRO_TABLE & ": " & regRows

  ' 3. Resolver indices de columnas de tbl_Registro
  If Not ResolveRegistroColumns(regData) Then
    Debug.Print "ERROR: No se pudieron resolver las columnas de " & REGISTRO_TABLE
    Exit Sub
  End If

  ' 4. Cargar diccionario de claves existentes en tbl_Motor (Fecha+Tag)
  Dim existingKeys As Object ' Scripting.Dictionary
  Set existingKeys = LoadExistingMotorKeys()
  Debug.Print "Registros existentes en " & MOTOR_TABLE & ": " & existingKeys.Count

  ' 5. Filtrar registros pendientes (no duplicados)
  Dim pendingRows As Collection
  Set pendingRows = FilterPendingRecords(regData, existingKeys, ForceMode)
  Debug.Print "Registros pendientes de calcular: " & pendingRows.Count

  If pendingRows.Count = 0 Then
    Debug.Print "No hay registros pendientes. Proceso finalizado."
    Exit Sub
  End If

  ' 6. Cargar tablas de aforo en cache
  mdTankService.ClearCache
  LoadStrappingTablesToCache

  ' 7. Procesar cada registro pendiente
  Dim results As Collection
  Set results = New Collection

  Dim i As Long
  For i = 1 To pendingRows.Count
    Dim rowIdx As Long
    rowIdx = pendingRows(i)
    Dim resultRow As Variant
    resultRow = ProcessSingleRecord(regData, rowIdx, tankConfigs)
    If Not IsEmpty(resultRow) Then
      results.Add resultRow
    End If
  Next i

  Debug.Print "Registros calculados exitosamente: " & results.Count

  ' 8. Escribir resultados en bloque a tbl_Motor
  If results.Count > 0 Then
    WriteResultsToMotor results
  End If

  ' 9. Limpieza
  mdTankService.ClearCache

  Dim elapsed As Double
  elapsed = Timer - startTime
  Debug.Print "======================================================================"
  Debug.Print "PROCESAMIENTO COMPLETADO en " & Format(elapsed, "0.00") & " segundos"
  Debug.Print "======================================================================"

  Exit Sub

ErrHandler:
  Debug.Print "Error en ProcesarRegistros: " & Err.Description
  mdTankService.ClearCache
End Sub

' ---------------------------------------------------------------------------------------------------------
' FUNCIONES AUXILIARES
' ---------------------------------------------------------------------------------------------------------

''' <summary>
''' Carga la configuracion de tanques desde tbl_Tanques.
''' Si la tabla no existe, retorna un diccionario vacio con un aviso.
''' </summary>

Private Function LoadTankConfigs() As Object
  Set LoadTankConfigs = CreateObject("Scripting.Dictionary")
  LoadTankConfigs.CompareMode = 1 ' TextCompare

  Dim lo As ListObject
  On Error Resume Next
  Set lo = ThisWorkbook.Sheets(TANQUES_SHEET).ListObjects(TANQUES_TABLE)
  On Error GoTo 0

  If lo Is Nothing Then
    Debug.Print "ADVERTENCIA: No se encontro " & TANQUES_TABLE & ". Se usaran valores por defecto."
    Exit Function
  End If

  If lo.ListRows.Count = 0 Then Exit Function

  ' Mapear nombres de columnas a indices
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

    cfg.API60 = GetCellNumeric(lr, colMap, "API60", 35)
    cfg.FluidType = GetCellFluidType(lr, colMap, "FluidType", CRD)
    cfg.Material = GetCellMaterial(lr, colMap, "Material", MCrbn)
    cfg.Diameter = GetCellNumeric(lr, colMap, "Diameter", 0)
    cfg.ShellThickness = GetCellNumeric(lr, colMap, "ShellThickness", 0)
    cfg.HasFloatingRoof = GetCellBool(lr, colMap, "HasFloatingRoof", False)
    cfg.RoofWeight = GetCellNumeric(lr, colMap, "RoofWeight", 0)
    cfg.IsTableNetOfRoof = GetCellBool(lr, colMap, "IsTableNetOfRoof", False)
    cfg.ReferenceAPIObs = GetCellNumeric(lr, colMap, "ReferenceAPIObs", cfg.API60)
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
''' Obtiene todos los datos de una tabla como Variant array.
''' </summary>

Private Function GetTableData(ByVal sheetName As String, ByVal tableName As String) As Variant
  Dim lo As ListObject
  On Error Resume Next
  Set lo = ThisWorkbook.Sheets(sheetName).ListObjects(tableName)
  On Error GoTo 0

  If lo Is Nothing Then
    GetTableData = Empty
    Exit Function
  End If

  If lo.ListRows.Count = 0 Then
    ' Solo headers, retornar array de 1 fila con los nombres
    Dim hdr() As Variant
    hdr = lo.HeaderRowRange.Value
    ReDim result(1 To 1, 1 To UBound(hdr, 2))
    Dim c As Long
    For c = 1 To UBound(hdr, 2)
      result(1, c) = hdr(1, c)
    Next c
    GetTableData = result
    Exit Function
  End If

  GetTableData = lo.DataBodyRange.Value
End Function

''' <summary>
''' Resuelve los indices de columnas de tbl_Registro buscando por nombre.
''' </summary>

Private Function ResolveRegistroColumns(ByRef data As Variant) As Boolean
  Dim headerRow As Variant
  ' Si hay mas de 1 fila, la primera es header (DataBodyRange no incluye headers)
  ' tbl_Registro tiene headers en la tabla, data en DataBodyRange
  ' Asumimos que la primera fila del array es la primera fila de datos
  ' Los headers estan en HeaderRowRange, no en DataBodyRange

  Dim lo As ListObject
  On Error Resume Next
  Set lo = ThisWorkbook.Sheets(REGISTRO_SHEET).ListObjects(REGISTRO_TABLE)
  On Error GoTo 0
  If lo Is Nothing Then ResolveRegistroColumns = False: Exit Function

  Dim hdr As Variant
  hdr = lo.HeaderRowRange.Value

  idx_Fecha = FindColumnIndex(hdr, "Fecha")
  idx_Tag = FindColumnIndex(hdr, "Tag")
  idx_Nivel = FindColumnIndex(hdr, "Nivel")
  idx_NivelAgua = FindColumnIndex(hdr, "Nivel_Agua_Libre")
  idx_TempLiq = FindColumnIndex(hdr, "Temp_Liq")
  idx_TempAmb = FindColumnIndex(hdr, "Temp_Amb")
  idx_Responsable = FindColumnIndex(hdr, "Responsable")
  idx_Observaciones = FindColumnIndex(hdr, "Observaciones")

  ' Columnas obligatorias
  If idx_Fecha = 0 Or idx_Tag = 0 Or idx_Nivel = 0 Or idx_TempLiq = 0 Or idx_TempAmb = 0 Then
    Debug.Print "Columnas obligatorias faltantes: Fecha=" & idx_Fecha & ", Tag=" & idx_Tag & _
      ", Nivel=" & idx_Nivel & ", TempLiq=" & idx_TempLiq & ", TempAmb=" & idx_TempAmb
    ResolveRegistroColumns = False
    Exit Function
  End If

  ' Columnas opcionales con valores por defecto
  If idx_NivelAgua = 0 Then idx_NivelAgua = 0 ' se manejara como 0
  If idx_Responsable = 0 Then idx_Responsable = 0
  If idx_Observaciones = 0 Then idx_Observaciones = 0

  ResolveRegistroColumns = True
End Function

''' <summary>
''' Busca el indice de una columna en un array de headers.
''' </summary>

Private Function FindColumnIndex(ByRef hdr As Variant, ByVal name As String) As Long
  If Not IsArray(hdr) Then
    FindColumnIndex = 0
    Exit Function
  End If
  Dim c As Long
  For c = 1 To UBound(hdr, 2)
    If InStr(1, CStr(hdr(1, c)), name, vbTextCompare) > 0 Then
      FindColumnIndex = c
      Exit Function
    End If
  Next c
  FindColumnIndex = 0
End Function

''' <summary>
''' Carga las claves existentes en tbl_Motor como diccionario.
''' La clave es "YYYY-MM-DD|TAG".
''' </summary>

Private Function LoadExistingMotorKeys() As Object
  Set LoadExistingMotorKeys = CreateObject("Scripting.Dictionary")
  LoadExistingMotorKeys.CompareMode = 1

  Dim data As Variant
  data = GetTableData(MOTOR_SHEET, MOTOR_TABLE)
  If IsEmpty(data) Then Exit Function

  Dim lo As ListObject
  On Error Resume Next
  Set lo = ThisWorkbook.Sheets(MOTOR_SHEET).ListObjects(MOTOR_TABLE)
  On Error GoTo 0
  If lo Is Nothing Then Exit Function

  Dim hdr As Variant
  hdr = lo.HeaderRowRange.Value

  Dim idxF As Long, idxT As Long
  idxF = FindColumnIndex(hdr, "Fecha")
  idxT = FindColumnIndex(hdr, "Tag")

  If idxF = 0 Or idxT = 0 Then Exit Function

  Dim r As Long
  For r = 1 To UBound(data, 1)
    Dim fVal As Variant, tVal As Variant
    fVal = data(r, idxF)
    tVal = data(r, idxT)
    If Not IsEmpty(fVal) And Not IsEmpty(tVal) Then
      Dim key As String
      key = MakeKey(fVal, CStr(tVal))
      LoadExistingMotorKeys(key) = True
    End If
  Next r
End Function

''' <summary>
''' Filtra los registros de tbl_Registro que no existen en tbl_Motor.
''' Retorna una Collection de indices de fila pendientes.
''' </summary>

Private Function FilterPendingRecords(ByRef data As Variant, ByRef existingKeys As Object, ByVal ForceMode As Boolean) As Collection
  Set FilterPendingRecords = New Collection

  Dim lo As ListObject
  On Error Resume Next
  Set lo = ThisWorkbook.Sheets(REGISTRO_SHEET).ListObjects(REGISTRO_TABLE)
  On Error GoTo 0
  If lo Is Nothing Then Exit Function

  Dim hdr As Variant
  hdr = lo.HeaderRowRange.Value
  Dim idxF As Long, idxT As Long
  idxF = FindColumnIndex(hdr, "Fecha")
  idxT = FindColumnIndex(hdr, "Tag")

  Dim r As Long
  For r = 1 To UBound(data, 1)
    Dim fVal As Variant, tVal As Variant
    fVal = data(r, idxF)
    tVal = data(r, idxT)

    If IsEmpty(fVal) Or IsEmpty(tVal) Then GoTo NextRow
    If Len(Trim(CStr(fVal))) = 0 Or Len(Trim(CStr(tVal))) = 0 Then GoTo NextRow

    Dim key As String
    key = MakeKey(fVal, CStr(tVal))

    If ForceMode Then
      ' En modo forzado, se procesan todos
      FilterPendingRecords.Add r
    Else
      ' Modo incremental: solo si no existe
      If Not existingKeys.Exists(key) Then
        FilterPendingRecords.Add r
      End If
    End If

NextRow:
  Next r
End Function

''' <summary>
''' Arma la clave unica para un registro: "YYYY-MM-DD|TAG".
''' </summary>

Private Function MakeKey(ByVal fecha As Variant, ByVal tag As String) As String
  Dim d As Date
  On Error Resume Next
  d = CDate(fecha)
  On Error GoTo 0
  If d = 0 Then d = DateValue(CStr(fecha))
  MakeKey = Format(d, "yyyy-mm-dd") & "|" & UCase(Trim(tag))
End Function

''' <summary>
''' Carga todas las tablas de aforo en el cache de mdTankService.
''' </summary>

Private Sub LoadStrappingTablesToCache()
  Dim ws As Worksheet
  On Error Resume Next
  Set ws = ThisWorkbook.Sheets(TANQUES_SHEET)
  On Error GoTo 0
  If ws Is Nothing Then Exit Sub

  Dim lo As ListObject
  For Each lo In ws.ListObjects
    If lo.ListRows.Count > 0 Then
      ' Forzar carga al cache llamando con un tag dummy
      ' (GetVolumeFromTable carga la tabla al cache si no esta)
      On Error Resume Next
      mdTankService.GetVolumeFromTable lo, lo.Name, 0
      On Error GoTo 0
    End If
  Next lo
  Debug.Print "Tablas de aforo cargadas en cache."
End Sub

''' <summary>
''' Procesa un solo registro de tbl_Registro y retorna un array con los resultados.
''' </summary>

Private Function ProcessSingleRecord(ByRef regData As Variant, ByVal rowIdx As Long, ByRef tankConfigs As Object) As Variant
  On Error GoTo ErrHandler

  Dim tag As String

  ' Extraer datos del registro
  Dim fecha As Date
  fecha = CDate(regData(rowIdx, idx_Fecha))
  tag = Trim(CStr(regData(rowIdx, idx_Tag)))
  Dim nivel As Long
  nivel = CLng(regData(rowIdx, idx_Nivel))
  Dim tempLiq As Double
  tempLiq = CDbl(regData(rowIdx, idx_TempLiq))
  Dim tempAmb As Double
  If idx_TempAmb > 0 And Not IsEmpty(regData(rowIdx, idx_TempAmb)) Then
    tempAmb = CDbl(regData(rowIdx, idx_TempAmb))
  Else
    tempAmb = tempLiq ' Si no hay temp ambiente, usar temp liquido
  End If
  Dim nivelAgua As Double
  If idx_NivelAgua > 0 And Not IsEmpty(regData(rowIdx, idx_NivelAgua)) Then
    nivelAgua = CDbl(regData(rowIdx, idx_NivelAgua))
  Else
    nivelAgua = 0
  End If
  Dim responsable As String
  If idx_Responsable > 0 And Not IsEmpty(regData(rowIdx, idx_Responsable)) Then
    responsable = CStr(regData(rowIdx, idx_Responsable))
  Else
    responsable = ""
  End If
  Dim observaciones As String
  If idx_Observaciones > 0 And Not IsEmpty(regData(rowIdx, idx_Observaciones)) Then
    observaciones = CStr(regData(rowIdx, idx_Observaciones))
  Else
    observaciones = ""
  End If

  ' Obtener configuracion del tanque
  Dim cfg As TankConfig
  If tankConfigs.Exists(tag) Then
    cfg = tankConfigs(tag)
  Else
    ' Configuracion por defecto si no existe en tbl_Tanques
    cfg.IsLoaded = True
    cfg.API60 = 35
    cfg.FluidType = CRD
    cfg.Material = MCrbn
    cfg.Diameter = 0
    cfg.ShellThickness = 0
    cfg.HasFloatingRoof = False
    Debug.Print "ADVERTENCIA: Tanque '" & tag & "' no encontrado en configuracion. Usando valores por defecto."
  End If

  ' Crear tanque temporal y configurar
  Dim t As New clsTank
  t.Tag = tag
  t.FluidType = cfg.FluidType
  t.ReferenceAPI60F = cfg.API60
  t.ReferenceAPIObs = cfg.ReferenceAPIObs
  t.Material = cfg.Material
  t.Diameter = cfg.Diameter
  t.ShellThickness = cfg.ShellThickness
  t.HasFloatingRoof = cfg.HasFloatingRoof
  t.RoofWeight = cfg.RoofWeight
  t.IsTableNetOfRoof = cfg.IsTableNetOfRoof
  t.MinLevelDeduction = cfg.MinLevelDeduction
  t.MaxLevelDeduction = cfg.MaxLevelDeduction
  t.BaseDeduction = cfg.BaseDeduction
  t.APICorrectionGT = cfg.APICorrectionGT
  t.APICorrectionLT = cfg.APICorrectionLT

  ' Asignar tabla de aforo si esta en cache
  t.StrappingTable = GetCachedStrappingTable(tag)

  ' Calcular factores
  Dim tov As Double, ctl As Double, ctsh As Double, fra As Double
  Dim gov As Double, gsv As Double, nsv As Double, csw As Double

  tov = t.GetTOV(nivel)
  If tov <= 0 Then
    tov = 0 ' Sin tabla de aforo, no se puede calcular
  End If

  ctl = t.CalculateCTL(cfg.API60, tempLiq)
  ctsh = t.CalculateCTSH(tempLiq, tempAmb)
  fra = t.CalculateFRA(nivel, cfg.ReferenceAPIObs)
  csw = t.CalculateCSW(nivelAgua)

  If tov > 0 Then
    gov = t.CalculateGOV(nivel, tempLiq, tempAmb, cfg.API60)
    gsv = t.CalculateGSV(nivel, tempLiq, tempAmb, cfg.API60)
    nsv = t.CalculateNSV(nivel, tempLiq, tempAmb, cfg.API60, nivelAgua)
  Else
    gov = 0
    gsv = 0
    nsv = 0
  End If

  ' Armar fila de resultado
  Dim result(1 To 17) As Variant
  result(1) = fecha               ' Fecha
  result(2) = tag                 ' Tag_Tanque
  result(3) = nivel               ' Nivel_mm
  result(4) = nivelAgua           ' Nivel_Agua_mm
  result(5) = tempLiq             ' Temp_Liq_F
  result(6) = tempAmb             ' Temp_Amb_F
  result(7) = cfg.API60           ' API60
  result(8) = tov                 ' TOV_Bbl
  result(9) = ctl                 ' CTL
  result(10) = ctsh               ' CTSH
  result(11) = fra                ' FRA_Bbl
  result(12) = gov                ' GOV_Bbl
  result(13) = gsv                ' GSV_Bbl
  result(14) = csw                ' CSW
  result(15) = nsv                ' NSV_Bbl
  result(16) = responsable        ' Responsable
  result(17) = observaciones      ' Observaciones

  ProcessSingleRecord = result
  Set t = Nothing
  Exit Function

ErrHandler:
  Debug.Print "Error procesando registro fila " & rowIdx & " (Tag=" & tag & "): " & Err.Description
  ProcessSingleRecord = Empty
  Set t = Nothing
End Function

''' <summary>
''' Busca la tabla de aforo cacheada para un tanque por su Tag.
''' </summary>

Private Function GetCachedStrappingTable(ByVal tag As String) As Object
  Dim ws As Worksheet
  On Error Resume Next
  Set ws = ThisWorkbook.Sheets(TANQUES_SHEET)
  On Error GoTo 0
  If ws Is Nothing Then Exit Function

  Dim lo As ListObject
  For Each lo In ws.ListObjects
    ' Verificar si alguna columna tiene el Tag del tanque
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
      If lc.Name = tag Or lc.Name = UCase(tag) Then
        Set GetCachedStrappingTable = lo
        Exit Function
      End If
    Next lc
  Next lo
End Function

''' <summary>
''' Escribe los resultados en bloque a tbl_Motor.
''' </summary>

Private Sub WriteResultsToMotor(ByRef results As Collection)
  Dim lo As ListObject
  On Error Resume Next
  Set lo = ThisWorkbook.Sheets(MOTOR_SHEET).ListObjects(MOTOR_TABLE)
  On Error GoTo 0

  If lo Is Nothing Then
    Debug.Print "ERROR: No se encontro " & MOTOR_TABLE & " en la hoja " & MOTOR_SHEET
    Exit Sub
  End If

  Dim n As Long
  n = results.Count
  If n = 0 Then Exit Sub

  Dim ncols As Long
  ncols = 17 ' Numero de columnas de resultado

  ' Crear matriz de resultados
  Dim outData() As Variant
  ReDim outData(1 To n, 1 To ncols)

  Dim i As Long, j As Long
  For i = 1 To n
    Dim row As Variant
    row = results(i)
    For j = 1 To ncols
      outData(i, j) = row(j)
    Next j
  Next i

  ' Escribir en bloque
  If lo.ListRows.Count = 0 Then
    ' Tabla vacia: escribir directamente
    Dim targetRange As Range
    Set targetRange = lo.HeaderRowRange.Offset(1, 0).Resize(n, ncols)
    targetRange.Value = outData
  Else
    ' Tabla con datos: agregar filas
    Dim firstRow As Long
    firstRow = lo.ListRows.Count + 1
    lo.Resize lo.Range.Resize(lo.Range.Rows.Count + n, ncols)
    lo.DataBodyRange.Offset(lo.DataBodyRange.Rows.Count - n, 0).Resize(n, ncols).Value = outData
  End If

  Debug.Print "Escritos " & n & " registros en " & MOTOR_TABLE
End Sub

' ---------------------------------------------------------------------------------------------------------
' HELPERS DE LECTURA DE CELDA
' ---------------------------------------------------------------------------------------------------------

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
