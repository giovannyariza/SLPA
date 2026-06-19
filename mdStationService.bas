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

  Dim stationTable As ListObject: Set stationTable = GetStationsTable()
  Dim rowItem As ListRow
  Dim stationsColl As New Collection
  Dim objStation As clsStation

  If stationTable.ListRows.Count = 0 Then
    Set GetAllStations = stationsColl
    Exit Function
  End If

  For Each rowItem In stationTable.ListRows
    Set objStation = MapRowToStation(rowItem)
    ' Si el Tag ya existe (duplicado), se salta sin romper la colección
    On Error Resume Next
    stationsColl.Add objStation, objStation.Tag
    On Error GoTo ErrHandler
  Next rowItem

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

  Dim stationTable As ListObject: Set stationTable = GetStationsTable()
  Dim rowItem As ListRow
  Dim idColIndex As Long

  idColIndex = stationTable.ListColumns("ID_Estación").Index

  For Each rowItem In stationTable.ListRows
    If rowItem.Range.Cells(1, idColIndex).Value = StationID Then
      Set GetStationByID = MapRowToStation(rowItem)
      Exit Function
    End If
  Next rowItem

  Exit Function
ErrHandler:
  Set GetStationByID = Nothing
End Function

''' <summary>
''' Función de Mapeo: Convierte una fila de tabla en un objeto clsStation.
''' Es extensible: si agregas una columna que coincida con una propiedad de la clase, se asignará sola.
''' </summary>

Private Function MapRowToStation(ByRef row As ListRow) As clsStation
  Dim objStation As New clsStation
  Dim tableObject As ListObject: Set tableObject = row.Parent
  Dim listColumn As ListColumn
  Dim headerName As String
  Dim cellValue As Variant

  For Each listColumn In tableObject.ListColumns
    headerName = listColumn.Name
    cellValue = row.Range.Cells(1, listColumn.Index).Value

    On Error Resume Next
    Select Case headerName
      Case "ID_Estación": objStation.Tag = CStr(cellValue)
      Case "Nombre":      objStation.Name = CStr(cellValue)
      Case "Descripción": objStation.Description = CStr(cellValue)
      Case "Estado":      objStation.Status = CStr(cellValue)
      Case "Ubicación":   objStation.Location = CStr(cellValue)
      Case Else
        CallByName objStation, headerName, VbLet, cellValue
    End Select
    On Error GoTo 0
  Next listColumn

  Set MapRowToStation = objStation
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
