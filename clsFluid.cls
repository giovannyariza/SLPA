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
' ---------------------------------------------------------------------------------------------------------
' PROPIEDADES PRIVADAS (BACKING FIELDS)
' ---------------------------------------------------------------------------------------------------------
Private name_ As String
Private description_ As String
Private type_ As eTypeLiq
Private refTemperature_ As Double
Private refPressure_ As Double
Private typDensity_ As Double
Private typAPI_ As Double
Private typViscosity_ As Double

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
  FluidType = type_
End Property

Public Property Let FluidType(ByVal value As eTypeLiq)
  type_ = value
End Property

''' <summary>
''' Gravedad API del fluido. Al establecer este valor, se actualiza automáticamente la Densidad Típica.
''' </summary>

Public Property Get TypicalAPI() As Double
  TypicalAPI = typAPI_
End Property

Public Property Let TypicalAPI(ByVal value As Double)
  ' Validación: El API no puede ser menor a -131.5 (límite matemático de la fórmula)
  If value <= -131.5 Then
    Err.Raise vbObjectError + 560, "clsFluid", "La gravedad API debe ser superior a -131.5."
  End If  
  typAPI_ = value
  ' Sincronización automática usando mdConversion
  typDensity_ = mdConversion.CONVDENS(value, A2S, True)
End Property

''' <summary>
''' Densidad Relativa (SGU). Al establecer este valor, se actualiza automáticamente la Gravedad API.
''' </summary>

Public Property Get TypicalDensity() As Double
  TypicalDensity = typDensity_
End Property

Public Property Let TypicalDensity(ByVal value As Double)
  If value <= 0 Then
    Err.Raise vbObjectError + 561, "clsFluid", "La densidad relativa debe ser mayor a cero."
  End If  
  typDensity_ = value
  ' Sincronización automática usando mdConversion
  typAPI_ = mdConversion.CONVDENS(value, S2A, True)
End Property

Public Property Get TypicalViscosity() As Double
  TypicalViscosity = typViscosity_
End Property

Public Property Let TypicalViscosity(ByVal value As Double)
  If value < 0 Then
    Err.Raise vbObjectError + 562, "clsFluid", "La viscosidad no puede ser un valor negativo."
  End If
  typViscosity_ = value
End Property

Public Property Get ReferenceTemperature() As Double
  ReferenceTemperature = refTemperature_
End Property

Public Property Let ReferenceTemperature(ByVal value As Double)
  refTemperature_ = value
End Property

Public Property Get ReferencePressure() As Double
  ReferencePressure = refPressure_
End Property

Public Property Let ReferencePressure(ByVal value As Double)
  refPressure_ = value
End Property

' ---------------------------------------------------------------------------------------------------------
' METODOS PUBLICOS
' ---------------------------------------------------------------------------------------------------------

''' <summary>
''' Imprime en la ventana de Inmediato las propiedades actuales del fluido para depuración.
''' </summary>

Public Sub ShowProperties()
  Debug.Print "--- Propiedades del Fluido ---"
  Debug.Print " Nombre: " & name_
  Debug.Print " Tipo API: " & Me.GetFluidTypeName
  Debug.Print " Gravedad API: " & FormatNumber(typAPI_, 2) & " °API"
  Debug.Print " Densidad Rel.: " & FormatNumber(typDensity_, 4) & " SGU"
  Debug.Print " Viscosidad: " & typViscosity_ & " cSt"
  Debug.Print " Ref. Temp/Pres: " & refTemperature_ & " °F / " & refPressure_ & " psi"
  Debug.Print "------------------------------"
End Sub

''' <summary>
''' Retorna el nombre legible del tipo de líquido basado en el Enum eTypeLiq.
''' </summary>

Public Function GetFluidTypeName() As String
  Select Case type_
    Case CRD: GetFluidTypeName = "Crudo (Crude Oil)"
    Case REF: GetFluidTypeName = "Producto Refinado"
    Case LUB: GetFluidTypeName = "Lubricante"
    Case Else: GetFluidTypeName = "No definido"
  End Select
End Function