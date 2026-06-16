VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ---------------------------------------------------------------------------------------------------------
' CLASE: clsWell
' DESCRIPCION:
' Representa un Tanque de Pozo (Productor o Inyector).
' Utiliza composicion con clsComponent para heredar atributos comunes y aplica la interfaz IComponent para
' permitir el polimorfismo.
'
' DEPENDENCIAS:
'   - IComponent.cls: Para la coleccion de componentes (permite almacenar diferentes tipos).
'   - clsComponent.cls: Para la base de propiedades comunes.
'
' PROPIEDADES PRINCIPALES:
'   - Depth (Double): Profundidad total del pozo (pies o metros).
'   - FlowRate (Double): Tasa de flujo actual (Producción / Inyeccion).
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

Private depth_ As Double
Private flowRate_ As Double
Private fluid_ As clsFluid

' ---------------------------------------------------------------------------------------------------------
' METODOS DE INSTANCIA DE CLASE
' ---------------------------------------------------------------------------------------------------------

Private Sub Class_Initialize()
  Set baseComponent_ = New clsComponent
  baseComponent_.ComponentType = "Pozo" ' Inicializar el componente base y definir el tipo
End Sub

Private Sub Class_Terminate()
  Set baseComponent_ = Nothing ' Libera la memoria ocupada por la composicion
End Sub

' ---------------------------------------------------------------------------------------------------------
' PROPIEDADES DE LA INTERFAZ ICOMPONENT (DELEGADAS AL BASE)
' ---------------------------------------------------------------------------------------------------------

Public Property Get Tag() As String
  Tag = baseComponent_.Tag
End Property

Public Property Let Tag(ByVal value As String)
  If Trim(value) = "" Then
    Err.Raise eErrors.errWellEmptyTag, "clsWell", "El Tag del pozo no puede estar vacío."
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
    Err.Raise eErrors.errWellEmptySystem, "clsWell", "El Sistema del pozo no puede estar vacío."
  End If
  baseComponent_.System = value
End Property

Public Property Get Service() As String
  Service = baseComponent_.Service
End Property

Public Property Let Service(ByVal value As String)
  If Trim(value) = "" Then 
    Err.Raise eErrors.errWellEmptyService, "clsWell", "El Servicio del pozo no puede estar vacío."
  End If
  baseComponent_.Service = value
End Property

Public Property Get Status() As String
  Status = baseComponent_.Status
End Property

Public Property Let Status(ByVal value As String)
  If Trim(value) = "" Then 
    Err.Raise eErrors.errWellEmptyStatus, "clsWell", "El Estado del pozo no puede estar vacío."
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
' PROPIEDADES ESPECÍFICAS DEL POZO
' ---------------------------------------------------------------------------------------------------------

Public Property Get Depth() As Double
  Depth = depth_
End Property

Public Property Let Depth(ByVal value As Double)
  If value <= 0 Then 
    Err.Raise eErrors.errWellInvalidDepth, "clsWell", "La profundidad del pozo debe ser mayor a cero."
  End If
  depth_ = value
End Property

Public Property Get FlowRate() As Double
  FlowRate = flowRate_
End Property

Public Property Let FlowRate(ByVal value As Double)
  ' Permite valores negativos únicamente si se modela conceptualmente un pozo inyector
  flowRate_ = value
End Property

' -- Fluido producido por el pozo --
Public Property Get Fluid() As clsFluid
  If fluid_ Is Nothing Then Set fluid_ = New clsFluid
  Set Fluid = fluid_
End Property

Public Property Set Fluid(ByVal value As clsFluid)
  Set fluid_ = value
End Property

' ---------------------------------------------------------------------------------------------------------
' METODOS PUBLICOS
' ---------------------------------------------------------------------------------------------------------

Public Sub ShowProperties()
  On Error Resume Next
  Debug.Print "--- Detalle de Pozo ---"
  Debug.Print "Tag del Pozo: " & Me.Tag
  Debug.Print "Funcion o Servicio: " & Me.Service
  Debug.Print "Profundidad: " & Me.Depth & " ft/m"
  Debug.Print "Tasa de Flujo: " & Me.FlowRate & " BFPD/MSCFD"
  If Not fluid_ Is Nothing Then
    Debug.Print "Fluido: " & fluid_.Name & " | API: " & FormatNumber(fluid_.TypicalAPI, 1)
  End If
  If Not Me.Station Is Nothing Then
    Debug.Print "Estacion Asignada: " & Me.Station.Tag
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