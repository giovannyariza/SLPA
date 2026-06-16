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
' DESCRIPCIÓN:
' Representa un Tanque de Almacenamiento.
' Utiliza composición con clsComponent para heredar atributos comunes y aplica la interfaz IComponent para
' permitir el polimorfismo.
'
' DEPENDENCIAS:
'   - IComponent.cls: Para la colección de componentes (permite almacenar diferentes tipos).
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
Private material_ As eMtrl
Private hasFloatingRoof_ As Boolean
Private roofWeight_ As Double
Private isInsulated_ As Boolean
Private restingLevel_ As Double
Private floatingLevel_ As Double
Private isTableNetOfRoof_ As Boolean
Private tableRefAPI_ As Double
Private tableBaseDeduction_ As Double
Private apiCorrectionGT_ As Double
Private apiCorrectionLT_ As Double
Private minLevelDeduction_ As Double
Private maxLevelDeduction_ As Double

' ---------------------------------------------------------------------------------------------------------
' MÉTODOS DE INSTANCIA DE CLASE
' ---------------------------------------------------------------------------------------------------------

Private Sub Class_Initialize()
  Set baseComponent_ = New clsComponent ' Inicializar el componente base y definir el tipo
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
    Err.Raise eErrors.errTankEmptyTag, "clsTank", "El Tag del tanque no puede estar vacío."
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
' Referencia a la estación a la que pertenece
  Set Station = baseComponent_.Station
End Property

Public Property Set Station(ByVal value As clsStation)
  Set baseComponent_.Station = value
End Property

' ---------------------------------------------------------------------------------------------------------
' PROPIEDADES ESPECIFICAS DEL TANQUE
' ---------------------------------------------------------------------------------------------------------
' -- Fluido --
Public Property Get FluidType() As eTypeLiq
  FluidType = fluidType_
End Property

Public Property Let FluidType(ByVal value As eTypeLiq)
  fluidType_ = value
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
    Err.Raise vbObjectError + 524, "clsTank", "El factor de cobertura debe ser mayor a cero."
  End If
  coverageFactor_ = value
End Property

Public Property Get ExpandedUncertainty() As Double
  ExpandedUncertainty = expandedUncertainty_
End Property

Public Property Let ExpandedUncertainty(ByVal value As Double)
  If value < 0 Or value > 100 Then
    Err.Raise vbObjectError + 525, "clsTank", "La incertidumbre expandida debe estar entre 0 y 100."
  End If
  expandedUncertainty_ = value
End Property

Public Property Get ConfidenceLevel() As Double
  ConfidenceLevel = confidenceLevel_
End Property

Public Property Let ConfidenceLevel(ByVal value As Double)
  If value < 0 Or value > 100 Then
    Err.Raise vbObjectError + 526, "clsTank", "El nivel de confianza debe estar entre 0 y 100."
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
    Err.Raise vbObjectError + 527, "clsTank", "El espesor del casco debe ser mayor a cero."
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
    Err.Raise eErrors.errTankNegativeCapacity, "clsTank", "La capacidad nominal no puede ser negativa."
  End If
  capacity_ = value
End Property

Public Property Get Diameter() As Double
  Diameter = diameter_
End Property

Public Property Let Diameter(ByVal value As Double)
  If value < 0 Then
    Err.Raise eErrors.errTankNegativeDiameter, "clsTank", "El diámetro no puede ser negativo."
  End If
  diameter_ = value
End Property

Public Property Get Material() As eMtrl
  Material = material_
End Property

Public Property Let Material(ByVal value As eMtrl)
  material_ = value
End Property

Public Property Get HasFloatingRoof() As Boolean
  HasFloatingRoof = hasFloatingRoof_
End Property

Public Property Let HasFloatingRoof(ByVal value As Boolean)
  hasFloatingRoof_ = value
End Property

Public Property Get RoofWeight() As Double
  RoofWeight = roofWeight_
End Property

Public Property Let RoofWeight(ByVal value As Double)
  If value < 0 Then
    Err.Raise eErrors.errTankNegativeRoofWeight, "clsTank", "El peso del techo no puede ser negativo."
  End If
  roofWeight_ = value
End Property

Public Property Get IsInsulated() As Boolean
  IsInsulated = isInsulated_
End Property

Public Property Let IsInsulated(ByVal value As Boolean)
  isInsulated_ = value
End Property

Public Property Let RestingLevel(ByVal value As Double)
  restingLevel_ = value
End Property

Public Property Get RestingLevel() As Double
  RestingLevel = restingLevel_
End Property

Public Property Let FloatingLevel(ByVal value As Double)
  floatingLevel_ = value
End Property

Public Property Get FloatingLevel() As Double
  FloatingLevel = floatingLevel_
End Property

Public Property Let IsTableNetOfRoof(ByVal value As Boolean)
  isTableNetOfRoof_ = value
End Property

Public Property Get IsTableNetOfRoof() As Boolean
  IsTableNetOfRoof = isTableNetOfRoof_
End Property

Public Property Let TableRefAPI(ByVal value As Double)
  tableRefAPI_ = value
End Property

Public Property Get TableRefAPI() As Double
  TableRefAPI = tableRefAPI_
End Property

Public Property Let TableBaseDeduction(ByVal value As Double)
  tableBaseDeduction_ = value
End Property

Public Property Get TableBaseDeduction() As Double
  TableBaseDeduction = tableBaseDeduction_
End Property

Public Property Let APICorrectionGT(ByVal value As Double)
  apiCorrectionGT_ = value
End Property

Public Property Get APICorrectionGT() As Double
  APICorrectionGT = apiCorrectionGT_
End Property

Public Property Let APICorrectionLT(ByVal value As Double)
  apiCorrectionLT_ = value
End Property

Public Property Get APICorrectionLT() As Double
  APICorrectionLT = apiCorrectionLT_
End Property

Public Property Let MinLevelDeduction(ByVal value As Double)
  minLevelDeduction_ = value
End Property

Public Property Get MinLevelDeduction() As Double
  MinLevelDeduction = minLevelDeduction_
End Property

Public Property Let MaxLevelDeduction(ByVal value As Double)
  maxLevelDeduction_ = value
End Property

Public Property Get MaxLevelDeduction() As Double
  MaxLevelDeduction = maxLevelDeduction_
End Property

' ---------------------------------------------------------------------------------------------------------
' MÉTODOS PÚBLICOS
' ---------------------------------------------------------------------------------------------------------
Public Sub ShowProperties()
' Metodo para mostrar propiedades del tanque (util para depuracion)
  Debug.Print "--- Detalle de Tanque ---"
  Debug.Print "Tag del Tanque: " & Me.Tag
  Debug.Print "Sistema: " & Me.System
  Debug.Print "Funcion o Servicio: " & Me.Service
  Debug.Print "Estado: " & Me.Status
  Debug.Print "Capacidad Nominal: " & Me.NominalCapacity & " Bbls"
  Debug.Print "Diametro: " & Me.Diameter
  Debug.Print "Aislado Termicamente?: " & IIf(Me.IsThermalInsulated, "Si", "No")
  Debug.Print "Tiene Techo Flotante?: " & IIf(Me.HasFloatingRoof, "Si", "No")
  If Not Me.Station Is Nothing Then
    Debug.Print "Pertenece a Estación: " & Me.Station.Tag
  End If
  Debug.Print "-------------------------"
End Sub

Public Function GetDynamicFRA(ByVal currentLevel As Double, ByVal actualAPIObs As Double) As Double
' Orquesta la deducción de volumen según el tipo de tabla y condiciones de campo.
  On Error GoTo ErrHandler
  
' Verificación de hardware (Crucial: Debe estar en True)
  If Not hasFloatingRoof_ Then Exit Function
  
  ' Verificación de nivel (mm)
  If currentLevel <= minLevelDeduction_ Then Exit Function

  Dim totalPotentialDeduction As Double
  
  ' DECISIÓN DE LÓGICA: ¿Tabla Bruta o Tabla Neta?
  If Not isTableNetOfRoof_ Then
    ' --- CASO A: TABLA BRUTA (Arquímedes) ---
    ' Calculamos la densidad observada a partir del API medido
    Dim densObs As Double
    densObs = mdConversion.CONVDENS(actualAPIObs, A2S, True)
    
    ' FRA = Peso / Densidad (Usando el motor de mdAPICalcs)
    totalPotentialDeduction = mdAPICalcs.FRA(roofWeight_, densObs, 1)
      
  Else
    ' --- CASO B: TABLA NETA (Ajuste por Coeficientes API) ---
    ' Implementamos la lógica de tu tabla de aforo ATK-7210
    Dim diffAPI As Double: diffAPI = actualAPIObs - tableRefAPI_
    
    If diffAPI > 0 Then
      ' API Real > Referencia (Líquido más liviano, más deducción)
      totalPotentialDeduction = diffAPI * apiCorrectionGT_
    Else
      ' API Real < Referencia (Líquido más pesado, menos deducción)
      totalPotentialDeduction = Abs(diffAPI) * apiCorrectionLT_
    End If
  End If

  ' AJUSTE POR NIVEL (Zona de Flotación Parcial)
  ' Si el nivel está entre Min y Max, la membrana no desplaza su peso total
  If currentLevel >= maxLevelDeduction_ Then
    ' Flotación Total
    GetDynamicFRA = totalPotentialDeduction
  Else
    ' Interpolación Lineal en la transición
    Dim partialFactor As Double
    partialFactor = (currentLevel - minLevelDeduction_) / (maxLevelDeduction_ - minLevelDeduction_)
    GetDynamicFRA = totalPotentialDeduction * partialFactor
  End If

  Exit Function

ErrHandler:
  Debug.Print "Error en GetDynamicFRA [" & Me.Tag & "]: " & Err.Description
  GetDynamicFRA = 0
End Function

' ---------------------------------------------------------------------------------------------------------
' IMPLEMENTACIÓN EXPLÍCITA Y MANDATORIA DE LA INTERFAZ ICOMPONENT
' ---------------------------------------------------------------------------------------------------------

Private Property Get IComponent_Tag() As String
  IComponent_Tag = Me.Tag
End Property

Private Property Let IComponent_Tag(ByVal rhs As String)
  Me.Tag = rhs
End Property

Private Property Get IComponent_Description() As String
  IComponent_Description = baseComponent_.Description
End Property

Private Property Get IComponent_ComponentType() As String
  IComponent_ComponentType = "Tanque"
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

Private Property Let IComponent_Description(ByVal rhs As String)
  baseComponent_.Description = rhs
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
