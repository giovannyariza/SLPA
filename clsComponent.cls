VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ---------------------------------------------------------------------------------------------------------
' CLASE: clsComponent
' DESCRIPCI�N:
' Representa un componente de un sistema.
'
' DEPENDENCIAS:
'   - IComponent.cls: Para la colecci�n de componentes (permite almacenar diferentes tipos).
'
' PROPIEDADES PRINCIPALES:
'   - Tag (String): Identificador �nico del componente. (ej: ATK-7213)
'   - Description (String): Descripci�n detallada del componente.
'   - ComponentType (String): Tipo de componente (ej: Tanque, Pozo, L�nea).
'   - System (String): Sistema al que pertenece el componente (ej: Almacenamiento, Tratamiento)
'   - Service (String): Funci�n o servicio dentro del proceso (ej: Surgencia, Fiscalizaci�n, Recibo, Despacho)
'   - Status (String): Estado actual del componente (ej: OP, OF, MT, MC, etc)
'   - Station (clsStation): Estaci�n a la cual pertenece.
'
' M�TODOS:
'   - ShowProperties: Muestra las propiedades del componente.
'
Implements IComponent

' ---------------------------------------------------------------------------------------------------------
' PROPIEDADES PRIVADAS (BACKING FIELDS)
' ---------------------------------------------------------------------------------------------------------
Private tag_ As String
Private description_ As String
Private type_ As String
Private system_ As String
Private service_ As String
Private status_ As String
Private station_ As clsStation ' Referencia a la estaci�n padre

' ---------------------------------------------------------------------------------------------------------
' PROPIEDADES P�BLICAS GET/LET CON VALIDACI�N
' ---------------------------------------------------------------------------------------------------------
Public Property Get Tag() As String
  Tag = tag_
End Property

Public Property Let Tag(ByVal value As String)
  If Trim(value) = "" Then
    Err.Raise eErrors.errComponentEmptyTag, "clsComponent", "El Tag del componente no puede estar vac�o."
  End If
  tag_ = UCase(Trim(value))
End Property

Public Property Get Description() As String
  Description = description_
End Property

Public Property Let Description(ByVal value As String)
  description_ = value
End Property

Public Property Get ComponentType() As String
  ComponentType = type_
End Property

Public Property Let ComponentType(ByVal value As String)
  If Trim(value) = "" Then
    Err.Raise eErrors.errComponentEmptyType, "clsComponent", "El Tipo de Componente no puede estar vac�o."
  End If
  type_ = value
End Property

Public Property Get System() As String
  System = system_
End Property

Public Property Let System(ByVal value As String)
  If Trim(value) = "" Then
    Err.Raise eErrors.errComponentEmptySystem, "clsComponent", "El Sistema no puede estar vac�o."
  End If
  system_ = value
End Property

Public Property Get Service() As String
  Service = service_
End Property

Public Property Let Service(ByVal value As String)
  If Trim(value) = "" Then
    Err.Raise eErrors.errComponentEmptyService, "clsComponent", "El Servicio no puede estar vac�o."
  End If
  service_ = value
End Property

Public Property Get Status() As String
  Status = status_
End Property

Public Property Let Status(ByVal value As String)
  If Trim(value) = "" Then
    Err.Raise eErrors.errComponentEmptyStatus, "clsComponent", "El Estado no puede estar vac�o."
  End If
  status_ = value
End Property

Public Property Get Station() As clsStation
' Referencia a la estaci�n (v�nculo bidireccional)
  Set Station = station_
End Property

Public Property Set Station(ByVal value As clsStation)
  Set station_ = value
End Property

' ---------------------------------------------------------------------------------------------------------
' M�TODOS P�BLICOS
' ---------------------------------------------------------------------------------------------------------
Public Sub ShowProperties()
' M�todo para mostrar propiedades del componente (�til para depuraci�n)
  Debug.Print "Tag del Componente: " & tag_
  Debug.Print "Descripci�n Detallada: " & description_
  Debug.Print "Tipo de Componente: " & type_
  Debug.Print "Sistema al que pertenece: " & system_
  Debug.Print "Funci�n o Servicio: " & service_
  Debug.Print "Estado del Componente: " & status_
  If Not station_ Is Nothing Then
    Debug.Print "Pertenece a Estaci�n: " & station_.Tag
  End If
End Sub

' ---------------------------------------------------------------------------------------------------------
' IMPLEMENTACI�N DE LA INTERFAZ ICOMPONENT
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
    Me.ComponentType = rhs
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
