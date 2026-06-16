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

' Indices de columnas en tbl_Registro (se resuelven dinamicamente)
Private idx_Fecha As Long
Private idx_Tag As Long
Private idx_Nivel As Long
Private idx_NivelAgua As Long
Private idx_TempLiq As Long
Private idx_TempAmb As Long
Private idx_Responsable As Long
Private idx_Observaciones As Long

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

  ' 1. Cargar configuracion de tanques (desde mdTankService)
  Dim tankConfigs As Object
  Set tankConfigs = mdTankService.LoadTankConfigs()
  Debug.Print "Configuracion cargada: " & tankConfigs.Count & " tanques"

  ' 2. Cargar registros de entrada
  Dim regData As Variant
  regData = GetTableData(REGISTRO_SHEET, REGISTRO_TABLE)
  If IsEmpty(regData) Then
    Debug.Print "ERROR: No se encontraron datos en " & REGISTRO_TABLE
    Exit Sub
  End If
  Debug.Print "Registros en " & REGISTRO_TABLE & ": " & UBound(regData, 1)

  ' 3. Resolver indices de columnas de tbl_Registro
  If Not ResolveRegistroColumns() Then
    Debug.Print "ERROR: No se pudieron resolver las columnas de " & REGISTRO_TABLE
    Exit Sub
  End If

  ' 4. Cargar diccionario de claves existentes en tbl_Motor (Fecha+Tag)
  Dim existingKeys As Object
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

  ' 6. Cargar tablas de aforo en cache (desde mdTankService)
  mdTankService.ClearCache
  mdTankService.LoadStrappingTablesToCache

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
' FUNCIONES AUXILIARES DEL PROCESO
' ---------------------------------------------------------------------------------------------------------

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

Private Function ResolveRegistroColumns() As Boolean
  Dim lo As ListObject
  On Error Resume Next
  Set lo = ThisWorkbook.Sheets(REGISTRO_SHEET).ListObjects(REGISTRO_TABLE)
  On Error GoTo 0
  If lo Is Nothing Then ResolveRegistroColumns = False: Exit Function

  Dim hdr As Variant
  hdr = lo.HeaderRowRange.Value

  idx_Fecha = mdHelpers.FindColumnIndex(hdr, "Fecha")
  idx_Tag = mdHelpers.FindColumnIndex(hdr, "Tag")
  idx_Nivel = mdHelpers.FindColumnIndex(hdr, "Nivel")
  idx_NivelAgua = mdHelpers.FindColumnIndex(hdr, "Nivel_Agua_Libre")
  idx_TempLiq = mdHelpers.FindColumnIndex(hdr, "Temp_Liq")
  idx_TempAmb = mdHelpers.FindColumnIndex(hdr, "Temp_Amb")
  idx_Responsable = mdHelpers.FindColumnIndex(hdr, "Responsable")
  idx_Observaciones = mdHelpers.FindColumnIndex(hdr, "Observaciones")

  If idx_Fecha = 0 Or idx_Tag = 0 Or idx_Nivel = 0 Or idx_TempLiq = 0 Or idx_TempAmb = 0 Then
    Debug.Print "Columnas obligatorias faltantes: Fecha=" & idx_Fecha & ", Tag=" & idx_Tag & _
      ", Nivel=" & idx_Nivel & ", TempLiq=" & idx_TempLiq & ", TempAmb=" & idx_TempAmb
    ResolveRegistroColumns = False
    Exit Function
  End If

  If idx_NivelAgua = 0 Then idx_NivelAgua = 0
  If idx_Responsable = 0 Then idx_Responsable = 0
  If idx_Observaciones = 0 Then idx_Observaciones = 0

  ResolveRegistroColumns = True
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
  idxF = mdHelpers.FindColumnIndex(hdr, "Fecha")
  idxT = mdHelpers.FindColumnIndex(hdr, "Tag")

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
  idxF = mdHelpers.FindColumnIndex(hdr, "Fecha")
  idxT = mdHelpers.FindColumnIndex(hdr, "Tag")

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
      FilterPendingRecords.Add r
    Else
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
''' Procesa un solo registro de tbl_Registro y retorna un array con los resultados.
''' </summary>

Private Function ProcessSingleRecord(ByRef regData As Variant, ByVal rowIdx As Long, ByRef tankConfigs As Object) As Variant
  On Error GoTo ErrHandler

  Dim tag As String

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
    tempAmb = tempLiq
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

  ' Obtener configuracion del tanque (ahora es un objeto clsTank)
  Dim t As clsTank
  If tankConfigs.Exists(tag) Then
    Set t = tankConfigs(tag)
  Else
    ' Configuracion por defecto si no existe en tbl_Tanques
    Set t = New clsTank
    t.Tag = tag
    Set t.Fluid = New clsFluid
    t.Fluid.FluidType = CRD
    t.Fluid.TypicalAPI = 35
    t.Material = MCrbn
    t.Diameter = 0
    t.ShellThickness = 0
    t.HasFloatingRoof = False
    Debug.Print "ADVERTENCIA: Tanque '" & tag & "' no encontrado en configuracion. Usando valores por defecto."
  End If

  ' Asignar tabla de aforo si esta en cache
  t.StrappingTable = mdTankService.GetCachedStrappingTable(tag)

  ' Calcular factores
  Dim tov As Double, ctsh As Double, fra As Double
  Dim gov As Double, gsv As Double, nsv As Double, csw As Double

  tov = t.GetTOV(nivel)
  If tov <= 0 Then tov = 0

  ctsh = t.CalculateCTSH(tempLiq, tempAmb)
  fra = t.CalculateFRA(nivel, t.Fluid.TypicalAPI)
  csw = t.CalculateCSW(nivelAgua)

  If tov > 0 Then
    gov = t.CalculateGOV(nivel, tempLiq, tempAmb)
    gsv = t.CalculateGSV(nivel, tempLiq, tempAmb)
    nsv = t.CalculateNSV(nivel, tempLiq, tempAmb, bsw:=nivelAgua)
  Else
    gov = 0
    gsv = 0
    nsv = 0
  End If

  Dim result(1 To 17) As Variant
  result(1) = fecha
  result(2) = tag
  result(3) = nivel
  result(4) = nivelAgua
  result(5) = tempLiq
  result(6) = tempAmb
  result(7) = t.Fluid.TypicalAPI
  result(8) = tov
  result(9) = t.CalculateCTL(tempLiq)
  result(10) = ctsh
  result(11) = fra
  result(12) = gov
  result(13) = gsv
  result(14) = csw
  result(15) = nsv
  result(16) = responsable
  result(17) = observaciones

  ProcessSingleRecord = result
  Set t = Nothing
  Exit Function

ErrHandler:
  Debug.Print "Error procesando registro fila " & rowIdx & " (Tag=" & tag & "): " & Err.Description
  ProcessSingleRecord = Empty
  Set t = Nothing
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
  ncols = 17

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

  If lo.ListRows.Count = 0 Then
    Dim targetRange As Range
    Set targetRange = lo.HeaderRowRange.Offset(1, 0).Resize(n, ncols)
    targetRange.Value = outData
  Else
    lo.Resize lo.Range.Resize(lo.Range.Rows.Count + n, ncols)
    lo.DataBodyRange.Offset(lo.DataBodyRange.Rows.Count - n, 0).Resize(n, ncols).Value = outData
  End If

  Debug.Print "Escritos " & n & " registros en " & MOTOR_TABLE
End Sub
