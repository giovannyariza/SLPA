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
' DESCRIPCIÓN:
'   Define el contrato para todos los componentes de una instalación.
'   Las clases como clsTank, clsWell, clsLine pueden implementar esta interfaz
'   para permitir que sean tratadas de forma polimórfica a través de una
'   variable declarada As IComponent.
'
' MIEMBROS:
'   - Tag (String): Identificador único.
'   - Description (String): Descripción.
'   - ComponentType (String): Tipo específico (ej: "Tanque", "Pozo").
'   - System (String): Sistema dentro del proceso al que pertenece.
'   - Service (String): Función en el proceso.
'   - Status (String): Estado operativo.
'   - Station (clsStation): Referencia a la estación a la que pertenece el componente.
'   - ShowProperties: Método para mostrar información.
'
' ---------------------------------------------------------------------------------------------------------
' PROPIEDADES PÚBLICAS GET/LET
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
' Referencia a la estación a la que pertenece el componente
End Property

Public Property Set Station(ByVal value As clsStation)
End Property

' ---------------------------------------------------------------------------------------------------------
' MÉTODOS PÚBLICOS
' ---------------------------------------------------------------------------------------------------------
Public Sub ShowProperties()
End Sub
