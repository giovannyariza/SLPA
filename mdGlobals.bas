Attribute VB_Name = "mdGlobals"
Option Explicit

' ------------------------------------------------------------------------------
' MÓDULO CENTRAL: mdGlobals
' DESCRIPCIÓN:
'   Este módulo contiene definiciones de enumeraciones (Enums) y constantes
'   globales utilizadas en todo el proyecto, particularmente en los cálculos
'   basados en las normas API MPMS. Centralizar estas definiciones ayuda a
'   mantener la coherencia y facilitar su gestión.
'
' DEPENDENCIAS:
'   - Ninguna dependencia directa de otros módulos o clases definidas en este
'     proyecto.
'
' CONTENIDO:
'   - Enumeraciones (Enums): Definen conjuntos de constantes relacionadas con
'     unidades de medida, tipos de líquido, etc.
'   - Constantes: Valores fijos utilizados en fórmulas y cálculos.
' ------------------------------------------------------------------------------

' ------------------------------------------------------------------------------
' ENUMERACIONES DE FLUIDOS Y PROPIEDADES TERMOFÍSICAS
' ------------------------------------------------------------------------------

' Selección del tipo de líquido para cálculos volumétricos y tablas de
' corrección API
Public Enum eTypeLiq
  CRD = 1 ' Crudo / Petróleo (Equivalente a Tabla API 6A)
  REF = 2 ' Productos Refinados (Equivalente a Tablas API 6B, 6C, 6E)
  LUB = 3 ' Aceites Lubricantes (Equivalente a Tabla API 6D)
End Enum

' Selección de unidades de temperatura estándar de ingeniería
Public Enum eTmpUnits
  F = 1 ' Fahrenheit
  C = 2 ' Celsius
  K = 3 ' Kelvin
  R = 4 ' Rankine
End Enum

' Selección de unidades de presión para cálculos hidrodinámicos
Public Enum ePrsUnits
  PSI = 1 ' Libras por pulgada cuadrada (Pounds per Square Inch)
  BAR = 2 ' Bares (Bar)
  KPA = 3 ' Kilopascales (Kilopascal)
End Enum

' Selección de unidades de densidad de fluidos
Public Enum eDnsUnits
  API = 1 ' Gravedad API (Grados API)
  KGM = 2 ' Kilogramos por Metro Cúbico (Kg/m³)
  SGU = 3 ' Unidades de Gravedad Específica (Specific Gravity Units)
End Enum

' Selección de unidades de volumen operativo
Public Enum eVolUnits
  BBL = 1 ' Barriles (Barrels)
  GAL = 2 ' Galones Americanos (US Gallons)
  M3 = 3  ' Metros Cúbicos (Cubic Meters)
  Lt = 4  ' Litros (Liters)
End Enum

' Selección de unidades de masa operativa
Public Enum eMassUnits
  KG = 1  ' Kilogramos (Kilograms)
  LB = 2  ' Libras (Pounds)
  TON = 3 ' Toneladas Métricas (Metric Tons)
End Enum

' Selección de unidades de longitud
Public Enum eLengthUnits
  MR = 1  ' Metros
  MM = 2  ' Milímetros
  CM = 3  ' Centímetros
  FT = 4  ' Pies
  IC = 5  ' Pulgadas
End Enum

' ------------------------------------------------------------------------------
' ENUMERACIONES DE ESPECIFICACIONES MATERIALES
' ------------------------------------------------------------------------------

' Selección normalizada del tipo de material para Tuberías y Tanques de
' almacenamiento
Public Enum eMtrl
  MCrbn = 1 ' Acero al Carbono (Carbon Steel)
  St304 = 2 ' Stainless Steel 304
  St316 = 3 ' Stainless Steel 316
  St4PH = 4 ' Stainless Steel 4PH
End Enum

' ------------------------------------------------------------------------------
' ENUMERACIONES DE ESTADOS OPERATIVOS DE COMPONENTES
' ------------------------------------------------------------------------------

' Estados operativos válidos para estaciones, tanques, pozos y líneas
Public Enum eComponentStatus
  OP = 1 ' Operativo
  OF = 2 ' Operativo con Fallas
  MT = 3 ' Mantenimiento
  MC = 4 ' Fuera de Servicio Mecánico
  FS = 5 ' Fuera de Servicio
End Enum

' ------------------------------------------------------------------------------
' ENUMERACIÓN CENTRALIZADA DE CÓDIGOS DE ERROR
' ------------------------------------------------------------------------------

' ------------------------------------------------------------------------------
' TIPOS DEFINIDOS POR EL USUARIO (UDT)
' ------------------------------------------------------------------------------

' Estructura de retorno para calculos volumetricos de fluidos
Public Type FluidCalcResult
  Alpha60 As Double          ' Coeficiente de expansion termica (1/F)
  CTL As Double              ' Factor correccion temperatura
  CPL As Double              ' Factor correccion presion
  CTPL As Double             ' Factor combinado (CTL * CPL)
  Compressibility As Double  ' Factor de compresibilidad escalado (Fp)
  DensityObs As Double       ' Densidad observada a tempF (SGU)
  APIObs As Double           ' API observado a tempF
  Density60F As Double       ' Densidad a 60F (SGU)
  IsValid As Boolean         ' Si el calculo fue exitoso
End Type

' Códigos de error para Err.Raise en todas las clases del proyecto.
' Elimina la dependencia de números mágicos (vbObjectError + 500, etc.)
Public Enum eErrors
  ' clsComponent
  errComponentEmptyTag = vbObjectError + 500
  errComponentEmptyType = vbObjectError + 501
  errComponentEmptySystem = vbObjectError + 502
  errComponentEmptyService = vbObjectError + 503
  errComponentEmptyStatus = vbObjectError + 504
  ' clsStation
  errStationEmptyTag = vbObjectError + 513
  errStationEmptyName = vbObjectError + 514
  errStationEmptyStatus = vbObjectError + 515
  errStationNullComponent = vbObjectError + 516
  ' clsTank
  errTankEmptyTag = vbObjectError + 520
  errTankNegativeCapacity = vbObjectError + 521
  errTankNegativeDiameter = vbObjectError + 522
  errTankNegativeRoofWeight = vbObjectError + 523
  errTankInvalidCoverageFactor = vbObjectError + 524
  errTankInvalidUncertainty = vbObjectError + 525
  errTankInvalidConfidence = vbObjectError + 526
  errTankInvalidShellThickness = vbObjectError + 527
  errTankNegativeNominalCapacity = vbObjectError + 528
  ' clsWell
  errWellEmptyTag = vbObjectError + 530
  errWellEmptySystem = vbObjectError + 531
  errWellEmptyService = vbObjectError + 532
  errWellEmptyStatus = vbObjectError + 533
  errWellInvalidDepth = vbObjectError + 534
  ' clsLine
  errLineEmptyTag = vbObjectError + 550
  errLineEmptySystem = vbObjectError + 551
  errLineEmptyService = vbObjectError + 552
  errLineEmptyStatus = vbObjectError + 553
  errLineInvalidLength = vbObjectError + 554
  errLineInvalidDiameter = vbObjectError + 555
  ' clsFluid
  errFluidInvalidAPI = vbObjectError + 560
  errFluidInvalidDensity = vbObjectError + 561
  errFluidNegativeViscosity = vbObjectError + 562
  ' mdStationService
  errServiceTableNotFound = vbObjectError + 1000
End Enum

' ------------------------------------------------------------------------------
' VARIABLES PUBLICAS
' ------------------------------------------------------------------------------

Public vReRaiseValidationErrors As Boolean ' Por defecto es False

' ------------------------------------------------------------------------------
' CONSTANTES PUBLICAS
' ------------------------------------------------------------------------------

Public Const cTEMPBASE_F As Double = 60
Public Const cTEMPBASE_C As Double = 15.56

' Temperatura base de referencia de 60 grados Fahrenheit (60 °F) expresada en
' la escala de temperatura IPTS-68.
Public Const cTEMPBASE_F68 As Double = 60.0068748977356
Public Const cTEMPBASE_C68 As Double = 15.5555555555556

Public Const cTEMPDIFF_F68 As Double = 0.0068748977356

' Factor de shift de temperatura utilizado en algunas fórmulas de factores de
' corrección API MPMS 11.1 (2007).
Public Const cTEMPSHIFT As Double = 1.37497954711989E-02

' Densidad del agua a 60 grados Fahrenheit (15.56 °C) en Kilogramos por Metro
' Cúbico (Kg/m³).
Public Const cWATERDENSKG_60F As Double = 999.016

' Rangos de validación de temperatura utilizado en algunas fórmulas de factores
' de corrección API MPMS 11.1 (2019).
Public Const cTEMPVALIDRANGE_MIN As Double = -350
Public Const cTEMPVALIDRANGE_MAX As Double = 600

' Constantes de Escala de Temperatura para Corrección ITS-90 a IPTS-68
Public Const cTEMPSCALE_MIN As Double = -183
Public Const cTEMPSCALE_MAX As Double = 630

' Rangos de validación de presión utilizado en algunas fórmulas de factores de
' corrección API MPMS 11.1 (2019).
Public Const cPRESSVALIDRANGE_MIN As Double = -100
Public Const cPRESSVALIDRANGE_MAX As Double = 20000

' Constantes normativas API MPMS 11.1
Public Const cK0_CRUDE As Double = 341.0957
Public Const cK1_LUBRICANT As Double = 0.34878

' Fronteras de Densidad (Tabla 6B)
Public Const cRHO_FUEL_OIL As Double = 838.3127
Public Const cRHO_JET_FUEL As Double = 787.5195
Public Const cRHO_TRANSITION As Double = 770.352
Public Const cRHO_GASOLINE As Double = 610.6

' Coeficientes Da
Public Const cDA_CRUDE As Double = 2
Public Const cDA_FUEL_OIL As Double = 1.3
Public Const cDA_JET_FUEL As Double = 2
Public Const cDA_TRANSITION As Double = 8.5
Public Const cDA_GASOLINES As Double = 1.5
Public Const cDA_LUBRICANTS As Double = 1

' Constantes de Conversión de Densidad API
Public Const cAPI_A As Double = 141.5
Public Const cAPI_B As Double = 131.5

' Factor utilizado en el término de la derivada de CTL (Dt) para el cálculo
' iterativo de API_60
Public Const cAPI_F16 As Double = 1.6

' Coeficientes FP: Compresibilidad
Public Const cFP_A As Double = -1.9947
Public Const cFP_B As Double = 0.00013427
Public Const cFP_C As Long = 793920
Public Const cFP_D As Integer = 2326

' Constantes Derivada FP
Public Const cFPDERIVATIVE_A As Double = 7.9392
Public Const cFPDERIVATIVE_B As Double = 0.02326

Public Const cPRESS_SCALING_API As Double = 0.00001

' Coeficientes Expansión Térmica del Casco
Public Const cCTSH_MCRBN_F As Double = 0.0000062
Public Const cCTSH_MCRBN_C As Double = 0.0000112
Public Const cCTSH_ST304_F As Double = 0.0000096
Public Const cCTSH_ST304_C As Double = 0.0000173
Public Const cCTSH_ST316_F As Double = 0.00000883
Public Const cCTSH_ST316_C As Double = 0.0000159
Public Const cCTSH_ST4PH_F As Double = 0.000006
Public Const cCTSH_ST4PH_C As Double = 0.0000108

' Factor para calcular la Densidad a 60F en la escala IPTS-68
Public Const cTAYLOR As Double = 0.8

Public Const cEPSILON As Double = 0.000001

' Tolerancia para redondeo bancario (Banker's Rounding) en ROUNDAPI
Public Const cEPSILON_ROUNDING As Double = 0.0000000001

' Coeficientes de corrección del hidrómetro por expansión del vidrio (API MPMS
' Cap. 9.3)
' Escala Fahrenheit
Public Const cHYCF_A As Double = 0.00001278
Public Const cHYCF_B As Double = 0.0000000062
' Escala Celsius
Public Const cHYCC_C As Double = 0.000233
Public Const cHYCC_D As Double = 0.00000023

' Propiedades de Materiales (API 650 / 12.1.1)
Public Const cYOUNG_MODULUS_STEEL_IMP As Double = 30000000
Public Const cYOUNG_MODULUS_STEEL_MET As Double = 206842718900

' ------------------------------------------------------------------------------
' PARA ORGANIZAR
' ------------------------------------------------------------------------------

' Rango de Gravedad API a 60°F para Crudos y Productos Refinados
Public Const cAPI60_RangeCrudeRefined_MIN As Double = 0
Public Const cAPI60_RangeCrudeRefined_MAX As Double = 130

' Rango de Gravedad API a 60°F para Aceites Lubricantes
Public Const cAPI60_RangeLubricant_MIN As Double = 0
Public Const cAPI60_RangeLubricant_MAX As Double = 40
