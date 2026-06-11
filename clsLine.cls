VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ---------------------------------------------------------------------------------------------------------
' CLASE: clsLine
' DESCRIPCION:
' Representa una Linea de Flujo / Tuberia.
' Utiliza composicion con clsComponent para heredar atributos comunes y aplica la interfaz IComponent para
' permitir el polimorfismo.
'
' DEPENDENCIAS:
'   - IComponent.cls: Para la coleccion de componentes (permite almacenar diferentes tipos).
'   - clsComponent.cls: Para la base de propiedades comunes.
'
' PROPIEDADES PRINCIPALES:
'   - Length (Double): Longitud de la linea (metros o pies).
'   - Diameter (Double): Diametro nominal o interno (pulgadas).
'   - Material (String): Material de construccion (ej: Acero al Carbono, HDPE).
'
' METODOS:
'   - ShowProperties: Muestra las propiedades del Pozo.
'
Implements IComponent

' Componente base para simular herencia
Private baseComponent_ As clsComponent

' ---------------------------------------------------------------------------------------------------------
' PROPIEDADES PRIVADAS (BACKING FIELDS)
' ---------------------------------------------------------------------------------------------------------

Private length_ As Double
Private diameter_ As Double
Private material_ As eMtrl

' ---------------------------------------------------------------------------------------------------------
' METODOS DE INSTANCIA DE CLASE
' ---------------------------------------------------------------------------------------------------------

Private Sub Class_Initialize()
  Set baseComponent_ = New clsComponent
  baseComponent_.ComponentType = "Línea" ' Inicializar el componente base y definir el tipo
End Sub

Private Sub Class_Terminate()
  Set baseComponent_ = Nothing ' Libera la memoria ocupada por la composición
End Sub

' ---------------------------------------------------------------------------------------------------------
' PROPIEDADES DE LA INTERFAZ ICOMPONENT (DELEGADAS AL BASE)
' ---------------------------------------------------------------------------------------------------------

Public Property Get Tag() As String
  Tag = baseComponent_.Tag
End Property

Public Property Let Tag(ByVal value As String)
  If Trim(value) = "" Then
    Err.Raise vbObjectError + 550, "clsLine", "El Tag de la línea no puede estar vacío."
  End If
  baseComponent_.Tag = UCase(Trim(value))
End Property

Public Property Get Description() As String
  Description = baseComponent_.Description
End Property

Public Property Let Description(ByVal value As String)
  baseComponent_.Description = value
End Property

Public Property Get ComponentType() As String
  ComponentType = baseComponent_.ComponentType
End Property

Public Property Get System() As String
  System = baseComponent_.System
End Property

Public Property Let System(ByVal value As String)
  If Trim(value) = "" Then
    Err.Raise vbObjectError + 551, "clsLine", "El Sistema de la línea no puede estar vacío."
  End If
  baseComponent_.System = value
End Property

Public Property Get Service() As String
  Service = baseComponent_.Service
End Property

Public Property Let Service(ByVal value As String)
  If Trim(value) = "" Then
    Err.Raise vbObjectError + 552, "clsLine", "El Servicio de la línea no puede estar vacío."
  End If
  baseComponent_.Service = value
End Property

Public Property Get Status() As String
  Status = baseComponent_.Status
End Property

Public Property Let Status(ByVal value As String)
  If Trim(value) = "" Then
    Err.Raise vbObjectError + 553, "clsLine", "El Estado de la línea no puede estar vacío."
  End If
  baseComponent_.Status = value
End Property

Public Property Get Station() As clsStation
  Set Station = baseComponent_.Station
End Property

Public Property Set Station(ByVal value As clsStation)
  Set baseComponent_.Station = value
End Property

' ---------------------------------------------------------------------------------------------------------
' PROPIEDADES ESPECIFICAS DE LA LINEA
' ---------------------------------------------------------------------------------------------------------

Public Property Get Length() As Double
  Length = length_
End Property

Public Property Let Length(ByVal value As Double)
  If value <= 0 Then
    Err.Raise vbObjectError + 554, "clsLine", "La longitud de la línea debe ser mayor a cero."
  End If
  length_ = value
End Property

Public Property Get Diameter() As Double
  Diameter = diameter_
End Property

Public Property Let Diameter(ByVal value As Double)
  If value <= 0 Then
    Err.Raise vbObjectError + 555, "clsLine", "El diámetro de la línea debe ser mayor a cero."
  End If
  diameter_ = value
End Property

Public Property Get Material() As eMtrl
  Material = material_
End Property

Public Property Let Material(ByVal value As eMtrl)
  material_ = value
End Property

' ---------------------------------------------------------------------------------------------------------
' METODOS PUBLICOS
' ---------------------------------------------------------------------------------------------------------

Public Sub ShowProperties()
  On Error Resume Next
  Debug.Print "--- Detalle de Línea ---"
  Debug.Print "Tag de la línea: " & Me.Tag
  Debug.Print "Tipo de Servicio: " & Me.Service
  Debug.Print "Longitud: " & Me.Length & " m/ft"
  Debug.Print "Diámetro: " & Me.Diameter & " in"
  Debug.Print "Material: " & Me.Material
  If Not Me.Station Is Nothing Then
    Debug.Print "Estación Asignada: " & Me.Station.Tag
  End If
  Debug.Print "------------------------"
End Sub

' ---------------------------------------------------------------------------------------------------------
' IMPLEMENTACION EXPLICITA Y MANDATORIA DE LA INTERFAZ ICOMPONENT
' ---------------------------------------------------------------------------------------------------------

Private Property Get IComponent_Tag() As String
  IComponent_Tag = Me.Tag
End Property

Private Property Let IComponent_Tag(ByVal rhs As String)
  Me.Tag = rhs
End Property

Private Property Get IComponent_Description() As String
  IComponent_Description = Me.Description
End Property

Private Property Let IComponent_Description(ByVal rhs As String)
  Me.Description = rhs
End Property

Private Property Get IComponent_ComponentType() As String
  IComponent_ComponentType = Me.ComponentType
End Property

Private Property Let IComponent_ComponentType(ByVal rhs As String)
  ' Propiedad de solo lectura programática
End Property

Private Property Get IComponent_System() As String
  IComponent_System = Me.System
End Property

Private Property Let IComponent_System(ByVal rhs As String)
  Me.System = rhs
End Property

Private Property Get IComponent_Service() As String
  IComponent_Service = Me.Service
End Property

Private Property Let IComponent_Service(ByVal rhs As String)
  Me.Service = rhs
End Property

Private Property Get IComponent_Status() As String
  IComponent_Status = Me.Status
End Property

Private Property Let IComponent_Status(ByVal rhs As String)
  Me.Status = rhs
End Property

Private Property Get IComponent_Station() As clsStation
  Set IComponent_Station = Me.Station
End Property

Private Property Set IComponent_Station(ByVal rhs As clsStation)
  Set Me.Station = rhs
End Property

Private Sub IComponent_ShowProperties()
  Me.ShowProperties
End Sub