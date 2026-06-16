VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' PROPIEDADES PRINCIPALES:
'   - Capacity (Double): Volumen Nominal (Bbls).
'   - Diameter (Double): Diámetro del tanque (pies/metros).
'   - Material (Long): Tipo de material (para factor de expansión).
'   - HasFloatingRoof (Boolean): Indica si tiene techo flotante.
'   - RoofWeight (Double): Peso del techo flotante (libras/kilos).
'   - IsInsulated (Boolean): Indica si tiene recubrimiento térmico.
'   - RestingLevel (Double): Nivel donde el techo toca las patas. (Zona Crítica Mínima).
'   - FloatingLevel (Double): Nivel donde el techo flota libremente. (Zona Crítica Máxima).
'   - IsTableNetOfRoof (Boolean): Indica si la tabla ya descuenta el volumen desplazado por el techo.
'   - TableRefAPI (Double): API Observado de referencia.
'   - TableBaseDeduction (Double): Deducción volumen por desplazamiento de la Membrana.
'   - APICorrectionGT (Double): Volumen corregido por cada grado API mayor al API Observado de referencia.
'   - APICorrectionLT (Double): Volumen corregido por cada grado API menor al API Observado de referencia.
'   - MinLevelDeduction (Double): Nivel mínimo de deducción.
'   - MaxLevelDeduction (Double): Nivel máximo de deducción.
'
' MÉTODOS:
'   - ShowProperties: Muestra las propiedades del Tanque.
'
Implements IComponent

' Componente base para simular herencia
Private baseComponent_ As clsComponent

' ---------------------------------------------------------------------------------------------------------
' PROPIEDADES PRIVADAS (BACKING FIELDS)
' ---------------------------------------------------------------------------------------------------------
Private capacity_ As Double
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

' Referencia a la estacion a la que pertenece
Public Property Get Station() As clsStation
  Set Station = baseComponent_.Station
End Property

Public Property Set Station(ByVal value As clsStation)
  Set baseComponent_.Station = value
End Property

' ---------------------------------------------------------------------------------------------------------
' PROPIEDADES ESPECIFICAS DEL TANQUE (en orden de backing fields)
' ---------------------------------------------------------------------------------------------------------
Public Property Get Capacity() As Double
  Capacity = capacity_
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
' MÉTODOS PÚBLICOS
' ---------------------------------------------------------------------------------------------------------

Public Sub ShowProperties()
' Metodo para mostrar propiedades del tanque (util para depuracion)
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
    Debug.Print "Pertenece a Estacion: " & Me.Station.Tag
  End If
  Debug.Print "-------------------------"
End Sub

Public Function GetDynamicFRA(ByVal currentLevel As Double, ByVal actualAPIObs As Double) As Double
' Orquesta la deducción de volumen según el tipo de tabla y condiciones de campo.
  On Error GoTo ErrHandler

  If Not Me.HasFloatingRoof Then Exit Function
  If currentLevel <= Me.MinLevelDeduction Then Exit Function

  Dim totalPotentialDeduction As Double

  If Not Me.IsTableNetOfRoof Then
    Dim densObs As Double
    densObs = mdConversion.CONVDENS(actualAPIObs, A2S, True)
    totalPotentialDeduction = mdAPICalcs.FRA(Me.RoofWeight, densObs, 1)
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
  IComponent_ComponentType = "Tanque"
End Property

Private Property Let IComponent_ComponentType(ByVal rhs As String)
' Propiedad de solo lectura programatica
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
