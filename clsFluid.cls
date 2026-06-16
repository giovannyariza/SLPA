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
' DESCRIPCION:
'   Representa las propiedades termofisicas y la caracterizacion de un fluido.
'   Incluye logica de auto-sincronizacion entre Gravedad API y Densidad Relativa,
'   y metodos de calculo que envuelven las formulas API MPMS de mdAPICalcs.
'
' DEPENDENCIAS:
'   - mdGlobals: Para los Enums eTypeLiq y constantes.
'   - mdConversion: Para interconversion de densidades.
'   - mdAPICalcs: Para formulas API MPMS (CTL, CPL, ALPHA60, FP, APIOBS).
'   - mdHelpers: Para validaciones de finitud.
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

' Cache de calculos (evita recalcular si los inputs no cambiaron)
Private cachedAlpha60_ As Double
Private cachedAlpha60Valid_ As Boolean

' ---------------------------------------------------------------------------------------------------------
' PROPIEDADES PUBLICAS GET/LET CON VALIDACION
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
  cachedAlpha60Valid_ = False
End Property

''' <summary>
''' Gravedad API del fluido. Al establecer este valor, se actualiza automáticamente la Densidad Típica.
''' </summary>

' -- Gravedad API del fluido. Al establecer, se actualiza automaticamente la Densidad Tipica. --
Public Property Get TypicalAPI() As Double
  TypicalAPI = typicalAPI_
End Property

Public Property Let TypicalAPI(ByVal value As Double)
  If Not mdHelpers.IsValidAPI(value, fluidType_) Then
    Err.Raise eErrors.errFluidInvalidAPI, "clsFluid", "La gravedad API debe ser superior a -131.5."
  End If
  typicalAPI_ = value
  typicalDensity_ = mdConversion.CONVDENS(value, A2S, True)
  cachedAlpha60Valid_ = False
End Property

''' <summary>
''' Densidad Relativa (SGU). Al establecer este valor, se actualiza automáticamente la Gravedad API.
''' </summary>

' -- Densidad Relativa (SGU). Al establecer, se actualiza automaticamente la Gravedad API. --
Public Property Get TypicalDensity() As Double
  TypicalDensity = typicalDensity_
End Property

Public Property Let TypicalDensity(ByVal value As Double)
  If value <= 0 Then
    Err.Raise eErrors.errFluidInvalidDensity, "clsFluid", "La densidad relativa debe ser mayor a cero."
  End If
  typicalDensity_ = value
  typicalAPI_ = mdConversion.CONVDENS(value, S2A, True)
  cachedAlpha60Valid_ = False
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

' -- Propiedades calculadas (solo lectura) --

''' <summary>
''' Densidad a 60F en SGU. Sincronizada con TypicalAPI.
''' </summary>
Public Property Get Density60F() As Double
  If typicalAPI_ > -cAPI_B Then
    Density60F = mdConversion.CONVDENS(typicalAPI_, A2S, True)
  Else
    Density60F = typicalDensity_
  End If
End Property

''' <summary>
''' API Observado equivalente a una temperatura dada.
''' </summary>
Public Property Get APIAtTemp(ByVal tempF As Double) As Double
  On Error Resume Next
  APIAtTemp = mdAPICalcs.APIOBS(typicalAPI_, tempF, fluidType_)
  If Err.Number <> 0 Then APIAtTemp = 0
End Property

''' <summary>
''' Densidad observada a una temperatura dada (SGU).
''' </summary>
Public Property Get DensityAtTemp(ByVal tempF As Double) As Double
  On Error Resume Next
  Dim ctl As Double
  ctl = mdAPICalcs.CTL(typicalAPI_, tempF, fluidType_)
  DensityAtTemp = Me.Density60F * ctl
  If Err.Number <> 0 Then DensityAtTemp = 0
End Property

''' <summary>
''' Coeficiente de expansion termica Alfa60 (1/F). Cacheado internamente.
''' </summary>
Public Property Get Alpha60() As Double
  If Not cachedAlpha60Valid_ Then
    cachedAlpha60_ = mdAPICalcs.ALPHA60(typicalAPI_, fluidType_)
    cachedAlpha60Valid_ = True
  End If
  Alpha60 = cachedAlpha60_
End Property

''' <summary>
''' Factor de compresibilidad escalado (Fp) a una temperatura dada.
''' </summary>
Public Property Get Compressibility(ByVal tempF As Double) As Double
  On Error Resume Next
  Compressibility = mdAPICalcs.FP(typicalAPI_, tempF, fluidType_)
  If Err.Number <> 0 Then Compressibility = 0
End Property

' ---------------------------------------------------------------------------------------------------------
' METODOS DE CALCULO (envolturas de mdAPICalcs)
' ---------------------------------------------------------------------------------------------------------

''' <summary>
''' Calcula el Factor de Correccion por Temperatura del Liquido (CTL).
''' </summary>
Public Function CalculateCTL(ByVal tempF As Double) As Double
  On Error GoTo ErrHandler
  CalculateCTL = mdAPICalcs.CTL(typicalAPI_, tempF, fluidType_)
  Exit Function
ErrHandler:
  Debug.Print "Error en Fluid.CalculateCTL [" & name_ & "]: " & Err.Description
  CalculateCTL = 1
End Function

''' <summary>
''' Calcula el Factor de Correccion por Presion del Liquido (CPL).
''' Retorna 1 si no hay presion.
''' </summary>
Public Function CalculateCPL(ByVal tempF As Double, Optional ByVal presObs As Double = 0, _
  Optional ByVal pe As Double = 0) As Double
  On Error GoTo ErrHandler
  If presObs <= 0 Then
    CalculateCPL = 1
    Exit Function
  End If
  CalculateCPL = mdAPICalcs.CPL(typicalAPI_, tempF, presObs, fluidType_, pe)
  Exit Function
ErrHandler:
  Debug.Print "Error en Fluid.CalculateCPL [" & name_ & "]: " & Err.Description
  CalculateCPL = 1
End Function

''' <summary>
''' Calcula el Factor de Correccion por Agua y Sedimento (CSW).
''' CSW = 1 - (BSW / 100). BSW en porcentaje.
''' </summary>
Public Function CalculateCSW(ByVal bsw As Double) As Double
  On Error GoTo ErrHandler
  If bsw < 0 Or bsw > 100 Then
    CalculateCSW = 1
    Exit Function
  End If
  CalculateCSW = 1 - (bsw / 100)
  Exit Function
ErrHandler:
  Debug.Print "Error en Fluid.CalculateCSW [" & name_ & "]: " & Err.Description
  CalculateCSW = 1
End Function

''' <summary>
''' Calcula el Factor de Compresibilidad Escalado (Fp).
''' </summary>
Public Function CalculateFP(ByVal tempF As Double) As Double
  On Error GoTo ErrHandler
  CalculateFP = mdAPICalcs.FP(typicalAPI_, tempF, fluidType_)
  Exit Function
ErrHandler:
  Debug.Print "Error en Fluid.CalculateFP [" & name_ & "]: " & Err.Description
  CalculateFP = 0
End Function

''' <summary>
''' Calcula el API Observado a una temperatura dada.
''' </summary>
Public Function CalculateAPIObs(ByVal tempF As Double) As Double
  On Error GoTo ErrHandler
  CalculateAPIObs = mdAPICalcs.APIOBS(typicalAPI_, tempF, fluidType_)
  Exit Function
ErrHandler:
  Debug.Print "Error en Fluid.CalculateAPIObs [" & name_ & "]: " & Err.Description
  CalculateAPIObs = 0
End Function

''' <summary>
''' Calculo completo de propiedades volumetricas del fluido.
''' Retorna una estructura FluidCalcResult con todos los factores.
''' </summary>
Public Function CalculateVolumetricProperties(ByVal tempF As Double, _
  Optional ByVal presObs As Double = 0) As FluidCalcResult
  Dim result As FluidCalcResult
  On Error GoTo ErrHandler

  result.IsValid = False

  ' Validar inputs
  If Not mdHelpers.IsFinite(typicalAPI_) Or Not mdHelpers.IsFinite(tempF) Then GoTo ExitHere

  result.Alpha60 = Me.Alpha60
  result.CTL = Me.CalculateCTL(tempF)
  result.CPL = Me.CalculateCPL(tempF, presObs)
  result.CTPL = result.CTL * result.CPL
  result.Compressibility = Me.CalculateFP(tempF)
  result.Density60F = Me.Density60F
  result.DensityObs = result.Density60F * result.CTL
  result.APIObs = Me.CalculateAPIObs(tempF)
  result.IsValid = True

ExitHere:
  CalculateVolumetricProperties = result
  Exit Function

ErrHandler:
  Debug.Print "Error en Fluid.CalculateVolumetricProperties [" & name_ & "]: " & Err.Description
  GoTo ExitHere
End Function

' ---------------------------------------------------------------------------------------------------------
' METODOS DE VALIDACION
' ---------------------------------------------------------------------------------------------------------

''' <summary>
''' Valida si el fluido tiene propiedades consistentes para calculo.
''' </summary>
Public Function IsValid() As Boolean
  IsValid = (typicalAPI_ > -cAPI_B) And _
            (typicalDensity_ > 0) And _
            (fluidType_ <> 0) And _
            mdHelpers.IsFinite(typicalAPI_)
End Function

''' <summary>
''' Valida si el fluido cumple los rangos normativos API MPMS para transferencia de custodia.
''' </summary>
Public Function ValidateForCustodyTransfer(Optional ByRef errorMsg As String) As Boolean
  ValidateForCustodyTransfer = False
  errorMsg = ""

  If Not mdHelpers.IsFinite(typicalAPI_) Then
    errorMsg = "API no es un valor finito."
    Exit Function
  End If

  If typicalAPI_ <= -cAPI_B Then
    errorMsg = "API debe ser mayor a -131.5."
    Exit Function
  End If

  ' Validar rango normativo segun tipo
  If Not mdHelpers.IsValidAPI(typicalAPI_, fluidType_, CheckNormative:=True) Then
    Select Case fluidType_
      Case CRD, REF
        errorMsg = "API " & typicalAPI_ & " fuera de rango normativo [" & _
          cAPI60_RangeCrudeRefined_MIN & "-" & cAPI60_RangeCrudeRefined_MAX & "] para crudo/refinado."
      Case LUB
        errorMsg = "API " & typicalAPI_ & " fuera de rango normativo [" & _
          cAPI60_RangeLubricant_MIN & "-" & cAPI60_RangeLubricant_MAX & "] para lubricante."
      Case Else
        errorMsg = "Tipo de fluido no definido."
    End Select
    Exit Function
  End If

  If typicalDensity_ <= 0 Then
    errorMsg = "Densidad debe ser mayor a cero."
    Exit Function
  End If

  ValidateForCustodyTransfer = True
End Function

' ---------------------------------------------------------------------------------------------------------
' METODOS DE VISUALIZACION
' ---------------------------------------------------------------------------------------------------------
Public Sub ShowProperties()
' Método para mostrar propiedades del fluido (útil para depuración)
  Debug.Print "--- Propiedades del Fluido ---"
  Debug.Print " Nombre: " & name_
  Debug.Print " Tipo: " & Me.GetFluidTypeName
  Debug.Print " Gravedad API (60F): " & FormatNumber(typicalAPI_, 2) & " API"
  Debug.Print " Densidad Relativa: " & FormatNumber(typicalDensity_, 4) & " SGU"
  Debug.Print " Densidad a 60F: " & FormatNumber(Me.Density60F, 4) & " SGU"
  Debug.Print " Viscosidad: " & typicalViscosity_ & " cSt"
  Debug.Print " Alfa60: " & FormatNumber(Me.Alpha60, 8) & " 1/F"
  Debug.Print " Ref. Temp/Pres: " & referenceTemperature_ & " F / " & referencePressure_ & " psi"
  ' Debug.Print " Valido para custodia: " & IIf(Me.ValidateForCustodyTransfer(errM), "Si", "No - " & errM)
  Debug.Print "------------------------------"
End Sub

''' <summary>
''' Retorna el nombre legible del tipo de líquido basado en el Enum eTypeLiq.
''' </summary>
Public Function GetFluidTypeName() As String
  Select Case fluidType_
    Case CRD: GetFluidTypeName = "Crudo (Crude Oil)"
    Case REF: GetFluidTypeName = "Producto Refinado"
    Case LUB: GetFluidTypeName = "Lubricante"
    Case Else: GetFluidTypeName = "No definido"
  End Select
End Function
