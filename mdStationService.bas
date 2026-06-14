Attribute VB_Name = "mdStationService"
Option Explicit

' ---------------------------------------------------------------------------------------------------------
' MÓDULO: mdStationService
' DESCRIPCIÓN:
'   Servicio de acceso a la tabla de estaciones (tbl_Estaciones).
'   Responsable exclusivo de: leer la tabla, mapear filas a objetos clsStation, y proveer lookup por ID.
' ---------------------------------------------------------------------------------------------------------

Private Const SHEET_NAME As String = "cnf_Maestro"
Private Const TABLE_NAME As String = "tbl_Estaciones"

''' <summary>
''' Obtiene todas las estaciones de la tabla y las devuelve como una colección de objetos clsStation.
''' </summary>

Public Function GetAllStations() As Collection
  On Error GoTo ErrHandler

  Dim lo As ListObject: Set lo = GetStationsTable()
  Dim lr As ListRow
  Dim stationsColl As New Collection
  Dim objStation As clsStation

  If lo.ListRows.Count = 0 Then
    Set GetAllStations = stationsColl
    Exit Function
  End If

  For Each lr In lo.ListRows
    Set objStation = MapRowToStation(lr)
    ' Si el Tag ya existe (duplicado), se salta sin romper la colección
    On Error Resume Next
    stationsColl.Add objStation, objStation.Tag
    On Error GoTo ErrHandler
  Next lr

  Set GetAllStations = stationsColl
  Exit Function

ErrHandler:
  Debug.Print "Error en GetAllStations: " & Err.Description
  Set GetAllStations = New Collection
End Function

''' <summary>
''' Busca una estación específica por su ID.
''' </summary>

Public Function GetStationByID(ByVal StationID As String) As clsStation
  On Error GoTo ErrHandler

  Dim lo As ListObject: Set lo = GetStationsTable()
  Dim lr As ListRow
  Dim idColIndex As Long

  idColIndex = lo.ListColumns("ID_Estación").Index

  For Each lr In lo.ListRows
    If lr.Range.Cells(1, idColIndex).Value = StationID Then
      Set GetStationByID = MapRowToStation(lr)
      Exit Function
    End If
  Next lr

  Exit Function
ErrHandler:
  Set GetStationByID = Nothing
End Function

''' <summary>
''' Función de Mapeo: Convierte una fila de tabla en un objeto clsStation.
''' Es extensible: si agregas una columna que coincida con una propiedad de la clase, se asignará sola.
''' </summary>

Private Function MapRowToStation(ByRef row As ListRow) As clsStation
  Dim obj As New clsStation
  Dim lo As ListObject: Set lo = row.Parent
  Dim lc As ListColumn
  Dim headerName As String
  Dim cellValue As Variant

  For Each lc In lo.ListColumns
    headerName = lc.Name
    cellValue = row.Range.Cells(1, lc.Index).Value

    On Error Resume Next
    Select Case headerName
      Case "ID_Estación": obj.Tag = CStr(cellValue)
      Case "Nombre":      obj.Name = CStr(cellValue)
      Case "Descripción": obj.Description = CStr(cellValue)
      Case "Estado":      obj.Status = CStr(cellValue)
      Case "Ubicación":   obj.Location = CStr(cellValue)
      Case Else
        CallByName obj, headerName, VbLet, cellValue
    End Select
    On Error GoTo 0
  Next lc

  Set MapRowToStation = obj
End Function

''' <summary>
''' Helper para obtener el objeto Tabla de forma segura.
''' </summary>

Private Function GetStationsTable() As ListObject
  On Error Resume Next

  Set GetStationsTable = ThisWorkbook.Sheets(SHEET_NAME).ListObjects(TABLE_NAME)

  If GetStationsTable Is Nothing Then
    Err.Raise eErrors.errServiceTableNotFound, "mdStationService", _
      "No se encontro la tabla '" & TABLE_NAME & "' en la hoja '" & SHEET_NAME & "'"
  End If

End Function
