Attribute VB_Name = "mdTankService"
Option Explicit

' ---------------------------------------------------------------------------------------------------------
' MÓDULO: mdTankService
' DESCRIPCIÓN:
'   Servicio de acceso a Tablas de Aforo (ListObject) con caché en RAM.
'   Responsable exclusivo de: cargar tablas de calibración, mantener caché, y retornar volúmenes TOV.
' ---------------------------------------------------------------------------------------------------------

' Variables de módulo para persistencia en RAM
Private m_CachedData As Variant      ' Matriz con los volúmenes (DataBodyRange)
Private m_CachedHeaders As Variant   ' Matriz con los Tags (HeaderRowRange)
Private m_CurrentTableName As String ' Identificador de la tabla en caché

''' <summary>
''' Recupera el volumen TOV de una Tabla de Aforo (ListObject) mediante indexación directa.
''' Si la tabla no está en caché, la carga automáticamente.
''' </summary>
''' <param name="TargetTable">Objeto ListObject que contiene las tablas.</param>
''' <param name="TankTag">Identificador del tanque (encabezado de columna).</param>
''' <param name="Level_mm">Nivel medido en milímetros (corresponde al índice de fila).</param>
''' <returns>Volumen TOV (Double). Retorna 0 si hay error o nivel fuera de rango.</returns>

Public Function GetVolumeFromTable(ByVal TargetTable As ListObject, ByVal TankTag As String, ByVal Level_mm As Long) As Double
  On Error GoTo ErrHandler

  ' GESTIÓN DE CACHÉ: ¿Es la misma tabla que tenemos en memoria?
  If m_CurrentTableName <> TargetTable.Name Or IsEmpty(m_CachedData) Then
    LoadTableToCache TargetTable
  End If

  ' Localizar la columna usando Application.Match sobre los encabezados cacheados
  Dim colIndex As Variant
  colIndex = Application.Match(TankTag, m_CachedHeaders, 0)

  If IsError(colIndex) Then
    Debug.Print "TankService: Tag '" & TankTag & "' no encontrado en la tabla " & TargetTable.Name
    Exit Function
  End If

  ' Indexación Directa (O(1))
  ' En un ListObject, la fila 1 del DataBodyRange corresponde al nivel 0mm.
  Dim targetRow As Long
  targetRow = Level_mm + 1

  ' Validación de límites físicos de la tabla
  If targetRow < 1 Or targetRow > UBound(m_CachedData, 1) Then
    Debug.Print "TankService: Nivel " & Level_mm & "mm fuera de rango para " & TankTag
    Exit Function
  End If

  ' Recuperación del valor (Velocidad de memoria RAM)
  GetVolumeFromTable = m_CachedData(targetRow, colIndex)

  Exit Function
ErrHandler:
  Debug.Print "Error en mdTankService.GetVolumeFromTable: " & Err.Description
  GetVolumeFromTable = 0
End Function

''' <summary>
''' Carga el ListObject completo a matrices de memoria (RAM).
''' </summary>

Private Sub LoadTableToCache(ByVal lo As ListObject)
  m_CurrentTableName = lo.Name

  ' Encabezados siempre (necesarios para Match)
  m_CachedHeaders = lo.HeaderRowRange.Value

  ' Datos: solo si hay filas
  If lo.ListRows.Count > 0 Then
    m_CachedData = lo.DataBodyRange.Value
    Debug.Print "TankService: Cache actualizado - Tabla '" & lo.Name & "' (" & UBound(m_CachedData, 1) & " filas)."
  Else
    m_CachedData = Empty
    Debug.Print "TankService: Cache actualizado - Tabla '" & lo.Name & "' (sin datos)."
  End If
End Sub

''' <summary>
''' Libera la memoria ocupada por el caché. Útil al cerrar el proceso.
''' </summary>

Public Sub ClearCache()
  m_CachedData = Empty
  m_CachedHeaders = Empty
  m_CurrentTableName = ""
End Sub
