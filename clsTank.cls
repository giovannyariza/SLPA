Option Explicit

' ---------------------------------------------------------------------------------------------------------
' CLASE: clsTank
' DESCRIPCION:
' Representa un Tanque de Almacenamiento.
' Utiliza composicion con clsComponent para heredar atributos comunes y aplica la interfaz IComponent para
' permitir el polimorfismo.
'
' DEPENDENCIAS:
'   - IComponent.cls: Para la coleccion de componentes (permite almacenar diferentes tipos).
'   - clsComponent.cls: Para la base de propiedades comunes.
'
' METODOS:
'   - ShowProperties: Muestra las propiedades del Tanque.
'
Implements IComponent

' Componente base para simular herencia
Private baseComponent_ As clsComponent

' ---------------------------------------------------------------------------------------------------------
' PROPIEDADES PRIVADAS (BACKING FIELDS)
' ---------------------------------------------------------------------------------------------------------
' -- Fluido contenido --
Private fluid_ As clsFluid

' -- Calibracion --
Private calibrationDate_ As Date
Private calibrationMethod_ As String
Private calibrationAgency_ As String
Private calibrationCertificate_ As String

' -- Incertidumbre --
Private coverageFactor_ As Double
Private expandedUncertainty_ As Double
Private confidenceLevel_ As Double

' -- Construccion --
Private material_ As eMtrl
Private shellThickness_ As Double
Private roofType_ As String
Private floorType_ As String

' -- Dimensiones --
Private nominalCapacity_ As Double
Private diameter_ As Double
Private effectiveHeight_ As Double
Private referenceHeight_ As Double
Private safeFillLevel_ As Double
Private safePumpLevel_ As Double

' -- Temperatura y API de referencia --
Private referenceAPI60F_ As Double
Private referenceAPIObs_ As Double
Private calibrationTemperature_ As Double
Private shellTemperature_ As Double

' -- Caracteristicas booleanas --
Private isThermalInsulated_ As Boolean
Private hasFloatingRoof_ As Boolean
Private isTableNetOfRoof_ As Boolean

' -- Deducciones y niveles criticos --
Private criticalZoneLower_ As Double
Private criticalZoneUpper_ As Double
Private roofWeight_ As Double
Private baseDeduction_ As Double
Private minLevelDeduction_ As Double
Private maxLevelDeduction_ As Double
Private apiCorrectionLT_ As Double
Private apiCorrectionGT_ As Double

' -- Referencia externa --
Private strappingTable_ As Object ' ListObject

' ---------------------------------------------------------------------------------------------------------
' METODOS DE INSTANCIA DE CLASE
' ---------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
  Set baseComponent_ = New clsComponent
  baseComponent_.ComponentType = "Tanque"
End Sub

Private Sub Class_Terminate()
  Set baseComponent_ = Nothing
End Sub

' ---------------------------------------------------------------------------------------------------------
' PROPIEDADES DE LA INTERFAZ ICOMPONENT (DELEGADAS AL BASE)
' ---------------------------------------------------------------------------------------------------------
Public Property Get Tag() As String
  Tag = baseComponent_.Tag
End Property

Public Property Let Tag(ByVal value As String)
  If Trim(value) = "" Then
    Err.Raise eErrors.errTankEmptyTag, "clsTank", "El Tag del tanque no puede estar vacio."
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

Public Property Let ComponentType(ByVal value As String)
  baseComponent_.ComponentType = value
End Property

Public Property Get System() As String
  System = baseComponent_.System
End Property

Public Property Let System(ByVal value As String)
  baseComponent_.System = value
End Property

Public Property Get Service() As String
  Service = baseComponent_.Service
End Property

Public Property Let Service(ByVal value As String)
  baseComponent_.Service = value
End Property

Public Property Get Status() As String
  Status = baseComponent_.Status
End Property

Public Property Let Status(ByVal value As String)
  baseComponent_.Status = value
End Property

Public Property Get Station() As clsStation
  Set Station = baseComponent_.Station
End Property

Public Property Set Station(ByVal value As clsStation)
  Set baseComponent_.Station = value
End Property

' ---------------------------------------------------------------------------------------------------------
' PROPIEDADES ESPECIFICAS DEL TANQUE (en orden de backing fields)
' ---------------------------------------------------------------------------------------------------------
' -- Fluido --
Public Property Get Fluid() As clsFluid
  If fluid_ Is Nothing Then Set fluid_ = New clsFluid
  Set Fluid = fluid_
End Property

Public Property Set Fluid(ByVal value As clsFluid)
  Set fluid_ = value
End Property

' -- Calibracion --
Public Property Get CalibrationDate() As Date
  CalibrationDate = calibrationDate_
End Property

Public Property Let CalibrationDate(ByVal value As Date)
  calibrationDate_ = value
End Property

Public Property Get CalibrationMethod() As String
  CalibrationMethod = calibrationMethod_
End Property

Public Property Let CalibrationMethod(ByVal value As String)
  calibrationMethod_ = value
End Property

Public Property Get CalibrationAgency() As String
  CalibrationAgency = calibrationAgency_
End Property

Public Property Let CalibrationAgency(ByVal value As String)
  calibrationAgency_ = value
End Property

Public Property Get CalibrationCertificate() As String
  CalibrationCertificate = calibrationCertificate_
End Property

Public Property Let CalibrationCertificate(ByVal value As String)
  calibrationCertificate_ = value
End Property

' -- Incertidumbre --
Public Property Get CoverageFactor() As Double
  CoverageFactor = coverageFactor_
End Property

Public Property Let CoverageFactor(ByVal value As Double)
  If value <= 0 Then
    Err.Raise eErrors.errTankInvalidCoverageFactor, "clsTank", "El factor de cobertura debe ser mayor a cero."
  End If
  coverageFactor_ = value
End Property

Public Property Get ExpandedUncertainty() As Double
  ExpandedUncertainty = expandedUncertainty_
End Property

Public Property Let ExpandedUncertainty(ByVal value As Double)
  If value < 0 Or value > 100 Then
    Err.Raise eErrors.errTankInvalidUncertainty, "clsTank", "La incertidumbre expandida debe estar entre 0 y 100."
  End If
  expandedUncertainty_ = value
End Property

Public Property Get ConfidenceLevel() As Double
  ConfidenceLevel = confidenceLevel_
End Property

Public Property Let ConfidenceLevel(ByVal value As Double)
  If value < 0 Or value > 100 Then
    Err.Raise eErrors.errTankInvalidConfidence, "clsTank", "El nivel de confianza debe estar entre 0 y 100."
  End If
  confidenceLevel_ = value
End Property

' -- Construccion --
Public Property Get Material() As eMtrl
  Material = material_
End Property

Public Property Let Material(ByVal value As eMtrl)
  material_ = value
End Property

Public Property Get ShellThickness() As Double
  ShellThickness = shellThickness_
End Property

Public Property Let ShellThickness(ByVal value As Double)
  If value <= 0 Then
    Err.Raise eErrors.errTankInvalidShellThickness, "clsTank", "El espesor del casco debe ser mayor a cero."
  End If
  shellThickness_ = value
End Property

Public Property Get RoofType() As String
  RoofType = roofType_
End Property

Public Property Let RoofType(ByVal value As String)
  roofType_ = value
End Property

Public Property Get FloorType() As String
  FloorType = floorType_
End Property

Public Property Let FloorType(ByVal value As String)
  floorType_ = value
End Property

' -- Dimensiones --
Public Property Get NominalCapacity() As Double
  NominalCapacity = nominalCapacity_
End Property

Public Property Let NominalCapacity(ByVal value As Double)
  If value < 0 Then
    Err.Raise vbObjectError + 528, "clsTank", "La capacidad nominal no puede ser negativa."
  End If
  nominalCapacity_ = value
End Property

Public Property Get Diameter() As Double
  Diameter = diameter_
End Property

Public Property Let Diameter(ByVal value As Double)
  If value < 0 Then
    Err.Raise vbObjectError + 529, "clsTank", "El diametro no puede ser negativo."
  End If
  diameter_ = value
End Property

Public Property Get EffectiveHeight() As Double
  EffectiveHeight = effectiveHeight_
End Property

Public Property Let EffectiveHeight(ByVal value As Double)
  If value < 0 Then
    Err.Raise vbObjectError + 530, "clsTank", "La altura efectiva no puede ser negativa."
  End If
  effectiveHeight_ = value
End Property

Public Property Get ReferenceHeight() As Double
  ReferenceHeight = referenceHeight_
End Property

Public Property Let ReferenceHeight(ByVal value As Double)
  referenceHeight_ = value
End Property

Public Property Get SafeFillLevel() As Double
  SafeFillLevel = safeFillLevel_
End Property

Public Property Let SafeFillLevel(ByVal value As Double)
  If value < 0 Then
    Err.Raise vbObjectError + 531, "clsTank", "El nivel de llenado seguro no puede ser negativo."
  End If
  safeFillLevel_ = value
End Property

Public Property Get SafePumpLevel() As Double
  SafePumpLevel = safePumpLevel_
End Property

Public Property Let SafePumpLevel(ByVal value As Double)
  If value < 0 Then
    Err.Raise vbObjectError + 532, "clsTank", "El nivel de bombeo seguro no puede ser negativo."
  End If
  safePumpLevel_ = value
End Property

' -- Temperatura y API de referencia --
Public Property Get ReferenceAPI60F() As Double
  ReferenceAPI60F = referenceAPI60F_
End Property

Public Property Let ReferenceAPI60F(ByVal value As Double)
  If value <= -cAPI_B Then
    Err.Raise vbObjectError + 533, "clsTank", "API60 de referencia debe ser mayor a -131.5."
  End If
  referenceAPI60F_ = value
End Property

Public Property Get ReferenceAPIObs() As Double
  ReferenceAPIObs = referenceAPIObs_
End Property

Public Property Let ReferenceAPIObs(ByVal value As Double)
  If value <= -cAPI_B Then
    Err.Raise vbObjectError + 534, "clsTank", "API Observado de referencia debe ser mayor a -131.5."
  End If
  referenceAPIObs_ = value
End Property

Public Property Get CalibrationTemperature() As Double
  CalibrationTemperature = calibrationTemperature_
End Property

Public Property Let CalibrationTemperature(ByVal value As Double)
  calibrationTemperature_ = value
End Property

Public Property Get ShellTemperature() As Double
  ShellTemperature = shellTemperature_
End Property

Public Property Let ShellTemperature(ByVal value As Double)
  shellTemperature_ = value
End Property

' -- Caracteristicas booleanas --
Public Property Get IsThermalInsulated() As Boolean
  IsThermalInsulated = isThermalInsulated_
End Property

Public Property Let IsThermalInsulated(ByVal value As Boolean)
  isThermalInsulated_ = value
End Property

Public Property Get HasFloatingRoof() As Boolean
  HasFloatingRoof = hasFloatingRoof_
End Property

Public Property Let HasFloatingRoof(ByVal value As Boolean)
  hasFloatingRoof_ = value
End Property

Public Property Get IsTableNetOfRoof() As Boolean
  IsTableNetOfRoof = isTableNetOfRoof_
End Property

Public Property Let IsTableNetOfRoof(ByVal value As Boolean)
  isTableNetOfRoof_ = value
End Property

' -- Deducciones y niveles criticos --
Public Property Get CriticalZoneLower() As Double
  CriticalZoneLower = criticalZoneLower_
End Property

Public Property Let CriticalZoneLower(ByVal value As Double)
  criticalZoneLower_ = value
End Property

Public Property Get CriticalZoneUpper() As Double
  CriticalZoneUpper = criticalZoneUpper_
End Property

Public Property Let CriticalZoneUpper(ByVal value As Double)
  criticalZoneUpper_ = value
End Property

Public Property Get RoofWeight() As Double
  RoofWeight = roofWeight_
End Property

Public Property Let RoofWeight(ByVal value As Double)
  If value < 0 Then
    Err.Raise vbObjectError + 535, "clsTank", "El peso del techo no puede ser negativo."
  End If
  roofWeight_ = value
End Property

Public Property Get BaseDeduction() As Double
  BaseDeduction = baseDeduction_
End Property

Public Property Let BaseDeduction(ByVal value As Double)
  baseDeduction_ = value
End Property

Public Property Get MinLevelDeduction() As Double
  MinLevelDeduction = minLevelDeduction_
End Property

Public Property Let MinLevelDeduction(ByVal value As Double)
  minLevelDeduction_ = value
End Property

Public Property Get MaxLevelDeduction() As Double
  MaxLevelDeduction = maxLevelDeduction_
End Property

Public Property Let MaxLevelDeduction(ByVal value As Double)
  maxLevelDeduction_ = value
End Property

Public Property Get APICorrectionLT() As Double
  APICorrectionLT = apiCorrectionLT_
End Property

Public Property Let APICorrectionLT(ByVal value As Double)
  apiCorrectionLT_ = value
End Property

Public Property Get APICorrectionGT() As Double
  APICorrectionGT = apiCorrectionGT_
End Property

Public Property Let APICorrectionGT(ByVal value As Double)
  apiCorrectionGT_ = value
End Property

' -- Referencia externa --
Public Property Set StrappingTable(ByVal value As Object) ' ListObject
  Set strappingTable_ = value
End Property

Public Property Get StrappingTable() As Object ' ListObject
  Set StrappingTable = strappingTable_
End Property

' ---------------------------------------------------------------------------------------------------------
' METODOS PUBLICOS
' ---------------------------------------------------------------------------------------------------------

Public Sub ShowProperties()
  Debug.Print "--- Detalle de Tanque ---"
  Debug.Print "Tag del Tanque: " & Me.Tag
  Debug.Print "Producto: " & Me.Fluid.Name
  Debug.Print "Sistema: " & Me.System
  Debug.Print "Funcion o Servicio: " & Me.Service
  Debug.Print "Estado: " & Me.Status
  Debug.Print "Capacidad Nominal: " & Me.NominalCapacity & " Bbls"
  Debug.Print "Diametro: " & Me.Diameter
  Debug.Print "Aislado Termicamente?: " & IIf(Me.IsThermalInsulated, "Si", "No")
  Debug.Print "Tiene Techo Flotante?: " & IIf(Me.HasFloatingRoof, "Si", "No")
  If Not Me.Station Is Nothing Then
    Debug.Print "Pertenece a Estacion: " & Me.Station.tag
  End If
  Debug.Print "-------------------------"
End Sub

' ---------------------------------------------------------------------------------------------------------
' METODOS DE CONVENIENCIA PARA LIQUIDACION (tbl_Motor)
' ---------------------------------------------------------------------------------------------------------
''' <summary>
''' Obtiene el Volumen Total Observado (TOV) desde la tabla de aforo cacheada.
''' </summary>
Public Function GetTOV(ByVal Level_mm As Long) As Double
  On Error GoTo ErrHandler
  If strappingTable_ Is Nothing Then Exit Function
  GetTOV = mdTankService.GetVolumeFromTable(strappingTable_, Me.Tag, Level_mm)
  Exit Function
ErrHandler:
  Debug.Print "Error en GetTOV [" & Me.Tag & "]: " & Err.Description
  GetTOV = 0
End Function

''' <summary>
''' Calcula el Factor de Correccion por Temperatura del Liquido (CTL).
''' Si API60F se omite, usa el API del fluido contenido.
''' </summary>
Public Function CalculateCTL(ByVal tempObsF As Double, ByVal API60F As Double) As Double
  On Error GoTo ErrHandler
  If API60F <= -cAPI_B Then
    CalculateCTL = 1
    Exit Function
  End If
  CalculateCTL = mdAPICalcs.CTL(API60F, tempObsF, Me.Fluid.FluidType)
  Exit Function
ErrHandler:
  Debug.Print "Error en CalculateCTL [" & Me.Tag & "]: " & Err.Description
  CalculateCTL = 1
End Function

''' <summary>
''' Calcula el Factor de Correccion por Presion del Liquido (CPL).
''' Si no hay presion, retorna 1 (sin correccion).
''' </summary>
Public Function CalculateCPL(ByVal tempObsF As Double, ByVal API60F As Double, _
  Optional ByVal presObs As Double = 0, Optional ByVal pe As Double = 0) As Double
  On Error GoTo ErrHandler
  If presObs <= 0 Then
    CalculateCPL = 1
    Exit Function
  End If
  If API60F <= -cAPI_B Then
    CalculateCPL = 1
    Exit Function
  End If
  CalculateCPL = mdAPICalcs.CPL(API60F, tempObsF, presObs, Me.Fluid.FluidType, pe)
  Exit Function
ErrHandler:
  Debug.Print "Error en CalculateCPL [" & Me.Tag & "]: " & Err.Description
  CalculateCPL = 1
End Function

''' <summary>
''' Calcula el Factor de Correccion por Temperatura de la Coraza (CTSH).
''' </summary>
Public Function CalculateCTSH(ByVal tempLiqF As Double, ByVal tempAmbF As Double) As Double
  On Error GoTo ErrHandler
  CalculateCTSH = mdAPICalcs.ctsh(tempLiqF, tempAmbF, Me.Material, f, cTEMPBASE_F)
  Exit Function
ErrHandler:
  Debug.Print "Error en CalculateCTSH [" & Me.Tag & "]: " & Err.Description
  CalculateCTSH = 1
End Function

''' <summary>
''' Calcula el Factor de Corrección por Agua y Sedimento (CSW).
''' </summary>
Public Function CalculateCSW(ByVal bsw As Double) As Double
  On Error GoTo ErrHandler
  CalculateCSW = 1 - (bsw / 100)
  Exit Function
ErrHandler:
  Debug.Print "Error en CalculateCSW [" & Me.Tag & "]: " & Err.Description
  CalculateCSW = 1
End Function

''' <summary>
''' Calcula el Volumen Bruto Observado (GOV).
''' GOV = TOV * CTSH - FRA
''' Si API60F se omite, usa el API del fluido contenido.
''' </summary>
Public Function CalculateGOV(ByVal Level_mm As Long, ByVal tempLiqF As Double, ByVal tempAmbF As Double, _
  ByVal API60F As Double) As Double
  On Error GoTo ErrHandler
  Dim tov As Double
  tov = Me.GetTOV(Level_mm)
  If tov <= 0 Then
    CalculateGOV = 0
    Exit Function
  End If
  Dim ctsh As Double
  ctsh = Me.CalculateCTSH(tempLiqF, tempAmbF)
  Dim fra As Double, apiObs As Double
  apiObs = mdAPICalcs.APIOBS(API60F, tempLiqF, Me.Fluid.FluidType)
  If apiObs > 0 Then
    fra = Me.CalculateFRA(Level_mm, apiObs)
  Else
    fra = Me.CalculateFRA(Level_mm, Me.ReferenceAPIObs)
  End If
  CalculateGOV = tov * ctsh - fra
  Exit Function
ErrHandler:
  Debug.Print "Error en CalculateGOV [" & Me.Tag & "]: " & Err.Description
  CalculateGOV = 0
End Function

''' <summary>
''' Calcula el Volumen Estandar Bruto (GSV).
''' GSV = GOV * CTL * CPL
''' Si API60F se omite, usa el API del fluido contenido.
''' </summary>
Public Function CalculateGSV(ByVal Level_mm As Long, ByVal tempLiqF As Double, ByVal tempAmbF As Double, _
  ByVal API60F As Double, Optional ByVal presObs As Double = 0, Optional ByVal pe As Double = 0) As Double
  On Error GoTo ErrHandler
  Dim gov As Double
  gov = Me.CalculateGOV(Level_mm, tempLiqF, tempAmbF, API60F)
  If gov <= 0 Then
    CalculateGSV = 0
    Exit Function
  End If
  Dim CTL As Double
  CTL = Me.CalculateCTL(tempLiqF, API60F)
  Dim CPL As Double
  CPL = Me.CalculateCPL(tempLiqF, API60F, presObs, pe)
  CalculateGSV = gov * CTL * CPL
  Exit Function
ErrHandler:
  Debug.Print "Error en CalculateGSV [" & Me.Tag & "]: " & Err.Description
  CalculateGSV = 0
End Function

''' <summary>
''' Calcula el Volumen Estandar Neto (NSV).
''' NSV = GSV * CSW
''' Si API60F se omite, usa el API del fluido contenido.
''' </summary>
Public Function CalculateNSV(ByVal Level_mm As Long, ByVal tempLiqF As Double, ByVal tempAmbF As Double, _
  ByVal API60F As Double, Optional ByVal bsw As Double = 0) As Double
  On Error GoTo ErrHandler
  Dim gsv As Double
  gsv = Me.CalculateGSV(Level_mm, tempLiqF, tempAmbF, API60F)
  If gsv <= 0 Then
    CalculateNSV = 0
    Exit Function
  End If
  Dim csw As Double
  csw = Me.CalculateCSW(bsw)
  CalculateNSV = gsv * csw
  Exit Function
ErrHandler:
  Debug.Print "Error en CalculateNSV [" & Me.Tag & "]: " & Err.Description
  CalculateNSV = 0
End Function

''' <summary>
''' Resuelve el API60 a usar: si el parametro es 0, usa el del fluido contenido.
''' </summary>
Private Function ResolveAPI60(ByVal API60F As Double) As Double
  If API60F > -cAPI_B And API60F <> 0 Then
    ResolveAPI60 = API60F
  Else
    ' Usar el fluido contenido si esta configurado
    If Not fluid_ Is Nothing And fluid_.IsValid Then
      ResolveAPI60 = fluid_.TypicalAPI
    ElseIf Me.ReferenceAPI60F > -cAPI_B Then
      ResolveAPI60 = Me.ReferenceAPI60F
    ElseIf Me.ReferenceAPIObs > -cAPI_B Then
      ResolveAPI60 = Me.ReferenceAPIObs
    Else
      ResolveAPI60 = 0 ' Sin API valido disponible
    End If
  End If
End Function

''' <summary>
''' Calcula el Ajuste por Techo Flotante (FRA) para liquidacion.
''' </summary>
Public Function CalculateFRA(ByVal currentLevel As Double, ByVal actualAPIObs As Double) As Double
  On Error GoTo ErrHandler

  If Not Me.HasFloatingRoof Then Exit Function
  If currentLevel <= Me.MinLevelDeduction Then Exit Function

  Dim totalPotentialDeduction As Double

  If Not Me.IsTableNetOfRoof Then
    Dim DensObs As Double
    DensObs = mdConversion.CONVDENS(actualAPIObs, A2S, True)
    totalPotentialDeduction = mdAPICalcs.FRA(Me.RoofWeight, DensObs, 1)
  Else
    Dim diffAPI As Double
    diffAPI = actualAPIObs - Me.ReferenceAPIObs
    If diffAPI > 0 Then
      totalPotentialDeduction = diffAPI * Me.APICorrectionGT
    Else
      totalPotentialDeduction = Abs(diffAPI) * Me.APICorrectionLT
    End If
  End If

  If currentLevel >= Me.MaxLevelDeduction Then
    CalculateFRA = totalPotentialDeduction
  Else
    Dim partialFactor As Double
    partialFactor = (currentLevel - Me.MinLevelDeduction) / (Me.MaxLevelDeduction - Me.MinLevelDeduction)
    CalculateFRA = totalPotentialDeduction * partialFactor
  End If

  Exit Function

ErrHandler:
  Debug.Print "Error en CalculateFRA [" & Me.Tag & "]: " & Err.Description
  CalculateFRA = 0
End Function

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