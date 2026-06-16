VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ---------------------------------------------------------------------------------------------------------
' INTERFACE: IComponent
' DESCRIPCI�N:
'   Define el contrato para todos los componentes de una instalaci�n.
'   Las clases como clsTank, clsWell, clsLine pueden implementar esta interfaz
'   para permitir que sean tratadas de forma polim�rfica a trav�s de una
'   variable declarada As IComponent.
'
' MIEMBROS:
'   - Tag (String): Identificador �nico.
'   - Description (String): Descripci�n.
'   - ComponentType (String): Tipo espec�fico (ej: "Tanque", "Pozo").
'   - System (String): Sistema dentro del proceso al que pertenece.
'   - Service (String): Funci�n en el proceso.
'   - Status (String): Estado operativo.
'   - Station (clsStation): Referencia a la estaci�n a la que pertenece el componente.
'   - ShowProperties: M�todo para mostrar informaci�n.
' ---------------------------------------------------------------------------------------------------------

' ---------------------------------------------------------------------------------------------------------
' PROPIEDADES P�BLICAS GET/LET
' ---------------------------------------------------------------------------------------------------------
Public Property Get Tag() As String
End Property
Public Property Let Tag(ByVal value As String)
End Property

Public Property Get Description() As String
End Property
Public Property Let Description(ByVal value As String)
End Property

Public Property Get ComponentType() As String
End Property
Public Property Let ComponentType(ByVal value As String)
End Property

Public Property Get System() As String
End Property
Public Property Let System(ByVal value As String)
End Property

Public Property Get Service() As String
End Property
Public Property Let Service(ByVal value As String)
End Property

Public Property Get Status() As String
End Property
Public Property Let Status(ByVal value As String)
End Property

Public Property Get Station() As clsStation
' Referencia a la estaci�n a la que pertenece el componente
End Property
Public Property Set Station(ByVal value As clsStation)
End Property

' ---------------------------------------------------------------------------------------------------------
' M�TODOS P�BLICOS
' ---------------------------------------------------------------------------------------------------------
Public Sub ShowProperties()
End Sub
