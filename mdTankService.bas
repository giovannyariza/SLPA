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
''' <param name="TargetTable">Objeto ListObject que contiene las tablas.</param>
''' <param name="TankTag">Identificador del tanque (encabezado de columna).</param>
''' <param name="Level_mm">Nivel medido en milímetros (corresponde al índice de fila).</param>
''' <returns>Volumen TOV (Double). Retorna 0 si hay error o nivel fuera de rango.</returns>

Public Function GetVolumeFromTable(ByVal TargetTable As ListObject, ByVal TankTag As String, ByVal Level_mm As Long) As Double
  On Error GoTo ErrHandler

  ' Nivel negativo no es válido
  If Level_mm < 0 Then Exit Function

  ' GESTIÓN DE CACHÉ: ¿Es la misma tabla que tenemos en memoria?
  If m_CurrentTableName <> TargetTable.Name Or IsEmpty(m_CachedData) Then
    LoadTableToCache TargetTable
  End If

  ' Si el caché quedó vacío (tabla sin datos), salir
  If IsEmpty(m_CachedData) Then Exit Function

  ' Localizar la columna usando Application.Match sobre los encabezados cacheados
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

  ' Indexación Directa (O(1))
  ' En un ListObject, la fila 1 del DataBodyRange corresponde al nivel 0mm.
  Dim targetRow As Long
  targetRow = Level_mm + 1

  ' Recuperación del valor (Velocidad de memoria RAM)
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

Private Sub LoadTableToCache(ByVal lo As ListObject)
  m_CurrentTableName = lo.Name

  ' Encabezados siempre (necesarios para Match)
  m_CachedHeaders = lo.HeaderRowRange.Value

  ' Datos: solo si hay filas
  If lo.ListRows.Count > 0 Then
    m_CachedData = lo.DataBodyRange.Value
  Else
    m_CachedData = Empty
    Debug.Print "TankService: Cache actualizado - Tabla '" & m_CurrentTableName & "' (sin datos)."
    Exit Sub
  End If

  ' ESCANEO DE RANGOS POR TANQUE
  ' Para cada columna (tanque), encontrar la primera y última fila con datos
  Set m_TankRanges = CreateObject("Scripting.Dictionary")
  m_TankRanges.CompareMode = 1 ' TextCompare

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

    ' Obtener el nombre del tanque desde los encabezados
    If IsArray(m_CachedHeaders) Then
      tagName = CStr(m_CachedHeaders(1, c))
    Else
      tagName = CStr(m_CachedHeaders)
    End If

    If Trim(tagName) = "" Then GoTo NextCol

    ' Recorrer filas de arriba hacia abajo para encontrar el primer dato
    For r = 1 To totalRows
      If Not IsEmpty(m_CachedData(r, c)) And IsNumeric(m_CachedData(r, c)) And m_CachedData(r, c) <> 0 Then
        foundMin = r - 1  ' Convertir índice de matriz (base 1) a nivel mm (base 0)
        hasData = True
        Exit For
      End If
    Next r

    ' Recorrer filas de abajo hacia arriba para encontrar el último dato
    If hasData Then
      For r = totalRows To 1 Step -1
        If Not IsEmpty(m_CachedData(r, c)) And IsNumeric(m_CachedData(r, c)) And m_CachedData(r, c) <> 0 Then
          foundMax = r - 1
          Exit For
        End If
      Next r
    End If

    ' Registrar el rango si se encontraron datos
    If hasData And foundMin >= 0 And foundMax >= 0 Then
      m_TankRanges.Add tagName, Array(foundMin, foundMax)
    End If

NextCol:
  Next c

  Debug.Print "TankService: Cache actualizado - Tabla '" & m_CurrentTableName & "' (" & totalRows & " filas, " & m_TankRanges.Count & " tanques con datos)."
End Sub

''' <summary>
''' Libera la memoria ocupada por el caché. Útil al cerrar el proceso.
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
