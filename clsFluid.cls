VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFluid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ---------------------------------------------------------------------------------------------------------
' CLASE: clsFluid
' DESCRIPCIÓN:
'   Representa las propiedades termofísicas y la caracterización de un fluido (Crudo, Refinados, etc.).
'   Incluye lógica de auto-sincronización entre Gravedad API y Densidad Relativa.
'
' DEPENDENCIAS:
'   - mdGlobals: Para los Enums eTypeLiq.
'   - mdConversion: Para el cálculo de interconversión de densidades.
'
' PROPIEDADES PRINCIPALES:
'   - Name (String): Nombre del fluido (ej: Nafta, Gasolina, Crudo).
'   - Description (String): Descripcion detallada del fluido.
'   - FluidType (eTypeLiq): Código de tipo de fluido (CRD, REF, LUB).
'   - TypicalDensity (Double): Densidad Relativa Tipica (SGU).
'   - TypicalAPI (Double): Gravedad API Tipica.
'   - TypicalViscosity (Double): Viscosidad Cinemática (cSt).
'   - ReferenceTemperature (Double): Temperatura de referencia.
'   - ReferencePressure (Double): Presión de referencia.
'
' METODOS:
'   - ShowProperties: Muestra las propiedades del Fluido.
'   - GetFluidTypeName: Retorna el nombre legible del tipo de líquido basado en el Enum eTypeLiq.
'
' ---------------------------------------------------------------------------------------------------------
' PROPIEDADES PRIVADAS (BACKING FIELDS)
' ---------------------------------------------------------------------------------------------------------
Private name_ As String
Private description_ As String
Private fluidType_ As eTypeLiq
Private typicalDensity_ As Double
Private typicalAPI_ As Double
Private typicalViscosity_ As Double
Private referenceTemperature_ As Double
Private referencePressure_ As Double

' ---------------------------------------------------------------------------------------------------------
' PROPIEDADES PÚBLICAS GET/LET CON VALIDACION
' ---------------------------------------------------------------------------------------------------------
Public Property Get Name() As String
  Name = name_
End Property

Public Property Let Name(ByVal value As String)
  name_ = Trim(value)
End Property

Public Property Get Description() As String
  Description = description_
End Property

Public Property Let Description(ByVal value As String)
  description_ = value
End Property

Public Property Get FluidType() As eTypeLiq
  FluidType = fluidType_
End Property

Public Property Let FluidType(ByVal value As eTypeLiq)
  fluidType_ = value
End Property

Public Property Get TypicalAPI() As Double
' Gravedad API del fluido. Al establecer este valor, se actualiza automáticamente la Densidad Típica.
  TypicalAPI = typicalAPI_
End Property

Public Property Let TypicalAPI(ByVal value As Double)
  ' Validación: El API no puede ser menor a -131.5 (límite matemático de la fórmula)
  If value <= -131.5 Then
    Err.Raise eErrors.errFluidInvalidAPI, "clsFluid", "La gravedad API debe ser superior a -131.5."
  End If  
  typicalAPI_ = value
  ' Sincronización automática usando mdConversion
  typicalDensity_ = mdConversion.CONVDENS(value, A2S, True)
End Property

Public Property Get TypicalDensity() As Double
' Densidad Relativa (SGU). Al establecer este valor, se actualiza automáticamente la Gravedad API.
  TypicalDensity = typicalDensity_
End Property

Public Property Let TypicalDensity(ByVal value As Double)
  If value <= 0 Then
    Err.Raise eErrors.errFluidInvalidDensity, "clsFluid", "La densidad relativa debe ser mayor a cero."
  End If  
  typicalDensity_ = value
  ' Sincronización automática usando mdConversion
  typicalAPI_ = mdConversion.CONVDENS(value, S2A, True)
End Property

Public Property Get TypicalViscosity() As Double
  TypicalViscosity = typicalViscosity_
End Property

Public Property Let TypicalViscosity(ByVal value As Double)
  If value < 0 Then
    Err.Raise eErrors.errFluidNegativeViscosity, "clsFluid", "La viscosidad no puede ser un valor negativo."
  End If
  typicalViscosity_ = value
End Property

Public Property Get ReferenceTemperature() As Double
  ReferenceTemperature = referenceTemperature_
End Property

Public Property Let ReferenceTemperature(ByVal value As Double)
  referenceTemperature_ = value
End Property

Public Property Get ReferencePressure() As Double
  ReferencePressure = referencePressure_
End Property

Public Property Let ReferencePressure(ByVal value As Double)
  referencePressure_ = value
End Property

' ---------------------------------------------------------------------------------------------------------
' METODOS PUBLICOS
' ---------------------------------------------------------------------------------------------------------
Public Sub ShowProperties()
' Método para mostrar propiedades del fluido (útil para depuración)
  Debug.Print "--- Propiedades del Fluido ---"
  Debug.Print " Nombre: " & name_
  Debug.Print " Tipo API: " & Me.GetFluidTypeName
  Debug.Print " Gravedad API: " & FormatNumber(typicalAPI_, 2) & " °API"
  Debug.Print " Densidad Rel.: " & FormatNumber(typicalDensity_, 4) & " SGU"
  Debug.Print " Viscosidad: " & typicalViscosity_ & " cSt"
  Debug.Print " Ref. Temp/Pres: " & referenceTemperature_ & " °F / " & referencePressure_ & " psi"
  Debug.Print "------------------------------"
End Sub

Public Function GetFluidTypeName() As String
' Retorna el nombre legible del tipo de líquido basado en el Enum eTypeLiq.
  Select Case fluidType_
    Case CRD: GetFluidTypeName = "Crudo (Crude Oil)"
    Case REF: GetFluidTypeName = "Producto Refinado"
    Case LUB: GetFluidTypeName = "Lubricante"
    Case Else: GetFluidTypeName = "No definido"
  End Select
End Function