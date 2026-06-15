VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ---------------------------------------------------------------------------------------------------------
' CLASE: clsStation
' DESCRIPCIÓN:
' Representa una Estación como un punto central en la instalación.
' Contiene una colección de componentes (tanques, pozos, líneas, etc.) que pertenecen a esa estación.
'
' DEPENDENCIAS:
'   - IComponent.cls: Para la colección de componentes (permite almacenar diferentes tipos).
'
' PROPIEDADES PRINCIPALES:
'   - Tag (String): Identificador único de la estación. (ej: EC2, EC3, EA1)
'   - Name (String): Nombre de la estación (puede ser mas descriptivo que el Tag).
'   - Description (String): Descripción detallada de la estación.
'   - Status (String): Estado operativo actual de la estación (ej: OP, FS, MT, etc).
'   - Location (String): Información de ubicación geográfica.
'   - Components (Collection): Colección de componentes (objetos que implementan IComponent) en esta estación. (Solo lectura)
'
' METODOS:
'   - AddComponent (Component As IComponent): Agrega un componente a la estación.
'   - RemoveComponent (Component As IComponent): Remueve un componente de la estación.
'   - FindComponentByTag (Tag As String): Busca un componente por su Tag.
'   - ShowProperties: Muestra las propiedades de la Estación.
'
' ---------------------------------------------------------------------------------------------------------
' PROPIEDADES PRIVADAS (BACKING FIELDS)
' ---------------------------------------------------------------------------------------------------------
Private tag_ As String
Private name_ As String
Private description_ As String
Private status_ As String
Private location_ As String
Private components_ As Collection ' Colección para almacenar los componentes

' ---------------------------------------------------------------------------------------------------------
' MÉTODOS DE INSTANCIA DE CLASE
' ---------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
  Set components_ = New Collection ' Inicializar la colección al crear la clase
End Sub

Private Sub Class_Terminate()
  Set components_ = Nothing
End Sub

' ---------------------------------------------------------------------------------------------------------
' PROPIEDADES PÚBLICAS GET/LET CON VALIDACIÓN
' ---------------------------------------------------------------------------------------------------------
Public Property Get Tag() As String
  Tag = tag_
End Property

Public Property Let Tag(ByVal value As String)
  If Trim(value) = "" Then
    Err.Raise eErrors.errStationEmptyTag, "clsStation", "El Tag no puede estar vacío."
  End If
  tag_ = value
End Property

Public Property Get Name() As String
  Name = name_
End Property

Public Property Let Name(ByVal value As String)
  If Trim(value) = "" Then
    Err.Raise eErrors.errStationEmptyName, "clsStation", "El Nombre no puede estar vacío."
  End If
  name_ = value
End Property

Public Property Get Description() As String
  Description = description_
End Property

Public Property Let Description(ByVal value As String)
  description_ = value
End Property

Public Property Get Status() As String
  Status = status_
End Property

Public Property Let Status(ByVal value As String)
  If Trim(value) = "" Then
    Err.Raise eErrors.errStationEmptyStatus, "clsStation", "El Estado no puede estar vacío."
  End If
  status_ = value
End Property

Public Property Get Location() As String
  Location = location_
End Property

Public Property Let Location(ByVal value As String)
  location_ = value
End Property

Public Property Get Components() As Collection
  ' Acceso a la colección de componentes (Sólo lectura)
  Set Components = components_
End Property

' ---------------------------------------------------------------------------------------------------------
' MÉTODOS PÚBLICOS
' ---------------------------------------------------------------------------------------------------------
Public Sub ShowProperties()
' Metodo para mostrar propiedades de la estación (útil para depuracion)
  Debug.Print "--- Propiedades de la Estación ---"
  Debug.Print " Tag: " & Me.Tag
  Debug.Print " Nombre: " & Me.Name
  Debug.Print " Descripción: " & Me.Description
  Debug.Print " Estado Operativo: " & Me.Status
  Debug.Print " Ubicación: " & Me.Location
  Debug.Print " Componentes: " & Me.Components.Count
  Debug.Print "----------------------------------"
End Sub

Public Sub AddComponent(ByVal Component As IComponent)
' Metodo que agrega un componente a la colección de la estación y lo vincula bidireccionalmente.
  On Error GoTo ErrHandler
  
  If Component Is Nothing Then
    Err.Raise eErrors.errStationNullComponent, "clsStation", "El componente a agregar es nulo (Nothing)."
  End If

  ' Intentar agregar a la colección usando el Tag como clave única
  ' Si el Tag ya existe, generará el error 457
  components_.Add Component, Component.Tag
  
  ' Establecer la referencia de la estación en el componente
  On Error Resume Next
  Set Component.Station = Me
  On Error GoTo ErrHandler
  
  Exit Sub

ErrHandler:
  If Err.Number = 457 Then
    Debug.Print "Error: Ya existe un componente con el Tag '" & Component.Tag & "' en esta estación."
  Else
    Debug.Print "Error en AddComponent: " & Err.Description
  End If
End Sub

Public Sub RemoveComponent(ByVal Identifier As Variant)
' Remueve un componente de la colección de la estación por su Tag. Desasocia la estación del componente.
  On Error GoTo ErrHandler
  
  ' Buscar el componente para desvincular la estación antes de borrarlo
  Dim comp As IComponent
  Set comp = FindComponentByTag(CStr(Identifier))
  
  If Not comp Is Nothing Then
    On Error Resume Next
    Set comp.Station = Nothing ' Quitar vínculo
    On Error GoTo ErrHandler
    
    components_.Remove Identifier
  Else
    Err.Raise 5, "clsStation", "Componente no encontrado para eliminar."
  End If
  
  Exit Sub

ErrHandler:
  Debug.Print "Error en RemoveComponent: " & Err.Description
End Sub

Public Function FindComponentByTag(ByVal Tag As String) As IComponent
' Busca y recupera un componente especifico en la colección de la estación por su Tag.
  On Error Resume Next
  
  Set FindComponentByTag = components_.Item(Tag)
  
  If Err.Number <> 0 Then
    Set FindComponentByTag = Nothing
    Err.Clear
  End If
End Function