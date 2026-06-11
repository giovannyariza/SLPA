Attribute VB_Name = "mdGlobals"
Option Explicit

' ---------------------------------------------------------------------------------------------------------
' MÓDULO CENTRAL: mdGlobals
' DESCRIPCIÓN:
'   Este módulo contiene definiciones de enumeraciones (Enums) y constantes globales utilizadas en todo el 
'   proyecto, particularmente en los cálculos basados en las normas API MPMS. Centralizar estas 
'   definiciones ayuda a mantener la coherencia y facilitar su gestión.
'
' DEPENDENCIAS:
'   - Ninguna dependencia directa de otros módulos o clases definidas en este proyecto.
'
' CONTENIDO:
'   - Enumeraciones (Enums): Definen conjuntos de constantes relacionadas con unidades de medida, tipos de
'     líquido, etc.
'   - Constantes: Valores fijos utilizados en fórmulas y cálculos.
' ---------------------------------------------------------------------------------------------------------

' ---------------------------------------------------------------------------------------------------------
' ENUMERACIONES DE FLUIDOS Y PROPIEDADES TERMOFÍSICAS
' ---------------------------------------------------------------------------------------------------------

' Selección del tipo de líquido para cálculos volumétricos y tablas de corrección API
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
  SGU = 3 ' Unidades de Gravedad Específica (Specific Gravity Units - adimensional)
End Enum

' ---------------------------------------------------------------------------------------------------------
' ENUMERACIONES PARA OPERACIONES DE CONVERSIÓN DE UNIDADES
' ---------------------------------------------------------------------------------------------------------

' Factores de conversión de unidades de presión
Public Enum eConvPrs
  P2B = 1 ' Convertir de PSI a Bares
  P2K = 2 ' Convertir de PSI a Kilopascales
  B2P = 3 ' Convertir de Bares a PSI
  B2K = 4 ' Convertir de Bares a Kilopascales
  K2P = 5 ' Convertir de Kilopascales a PSI
  K2B = 6 ' Convertir de Kilopascales a Bares
End Enum

' Factores de conversión de unidades de temperatura
Public Enum eConvTmp
  C2F = 1  ' Convertir de Celsius a Fahrenheit
  C2K = 2  ' Convertir de Celsius a Kelvin
  C2R = 3  ' Convertir de Celsius a Rankine
  F2C = 4  ' Convertir de Fahrenheit a Celsius
  F2K = 5  ' Convertir de Fahrenheit a Kelvin
  F2R = 6  ' Convertir de Fahrenheit a Rankine
  K2C = 7  ' Convertir de Kelvin a Celsius
  K2F = 8  ' Convertir de Kelvin a Fahrenheit
  K2R = 9  ' Convertir de Kelvin a Rankine
  R2C = 10 ' Convertir de Rankine a Celsius
  R2F = 11 ' Convertir de Rankine a Fahrenheit
  R2K = 12 ' Convertir de Rankine a Kelvin
End Enum

' Factores de conversión de unidades de densidad / gravedad
Public Enum eConvDns
  A2S = 1 ' Convertir de Gravedad API a Specific Gravity (SGU)
  A2K = 2 ' Convertir de Gravedad API a Densidad (Kg/m³)
  S2A = 3 ' Convertir de Specific Gravity (SGU) a Gravedad API
  S2K = 4 ' Convertir de Specific Gravity (SGU) a Densidad (Kg/m³)
  K2A = 5 ' Convertir de Densidad (Kg/m³) a Gravedad API
  K2S = 6 ' Convertir de Densidad (Kg/m³) a Specific Gravity (SGU)
End Enum

' ---------------------------------------------------------------------------------------------------------
' ENUMERACIONES DE ESPECIFICACIONES MATERIALES
' ---------------------------------------------------------------------------------------------------------

' Selección normalizada del tipo de material para Tuberías y Tanques de almacenamiento
Public Enum eMtrl
  MCrbn = 1 ' Acero al Carbono (Carbon Steel)
  St304 = 2 ' Stainless Steel 304
  St316 = 3 ' Stainless Steel 316
  St4PH = 4 ' Stainless Steel 4PH
End Enum

' -------------------------------------------------------------------------------------------------
' VARIABLES PUBLICAS
' -------------------------------------------------------------------------------------------------
Public vReRaiseValidationErrors As Boolean ' Por defecto es False

' ---------------------------------------------------------------------------------------------------------
' CONSTANTES PUBLICAS
' ---------------------------------------------------------------------------------------------------------

Public Const cBaseTempF As Double = 60
Public Const cBaseTempC As Double = 15.56

' Factor de shift de temperatura utilizado en algunas fórmulas de factores de corrección API MPMS 11.1 (2007).
Public Const cTmpShift As Double = 1.37497954711989E-02

' Densidad del agua a 60 grados Fahrenheit (15.56 °C) en Kilogramos por Metro Cúbico (Kg/m³).
Public Const cWtrDensKgM3_60F As Double = 999.016

' Rangos de validación de temperatura utilizado en algunas fórmulas de factores de corrección API MPMS 11.1 (2019).
Public Const cTempFValidRangeMin As Double = -350
Public Const cTempFValidRangeMax As Double = 600

' Temperatura base de referencia de 60 grados Fahrenheit (60 °F) expresada en la escala de temperatura IPTS-68.
Public Const cTmpBase68 As Double = 60.0068748977356
Public Const cBaseTempC68 As Double = 15.5555555555556

' Constantes de Escala de Temperatura para Corrección ITS-90 a IPTS-68
Public Const cTempScale_Min As Double = -183
Public Const cTempScale_Max As Double = 630

' 
Public Const c60F_TempDiff_IPTS68_ITS90_F As Double = 0.0068748977356
' -------------------------------------------------------------------------------------------------

' Rango de Gravedad API a 60°F para Crudos y Productos Refinados
Public Const cAPI60_RangeCrudeRefined_Min As Double = 0
Public Const cAPI60_RangeCrudeRefined_Max As Double = 130

' Rango de Gravedad API a 60°F para Aceites Lubricantes
Public Const cAPI60_RangeLubricant_Min As Double = 0
Public Const cAPI60_RangeLubricant_Max As Double = 40

' Coeficientes TSH
Public Const cTSH_CoefLiq As Double = 7
Public Const cTSH_Divisor As Double = 8

' Coeficientes Expansión Térmica del Casco
Public Const cCTSH_CoefMCrbn_F As Double = 0.0000062
Public Const cCTSH_CoefMCrbn_C As Double = 0.0000112
Public Const cCTSH_CoefSt304_F As Double = 0.0000096
Public Const cCTSH_CoefSt304_C As Double = 0.0000173
Public Const cCTSH_CoefSt316_F As Double = 0.00000883
Public Const cCTSH_CoefSt316_C As Double = 0.0000159
Public Const cCTSH_CoefSt4PH_F As Double = 0.000006
Public Const cCTSH_CoefSt4PH_C As Double = 0.0000108

'
Public Const cShrinkageFactorCU_Const As Double = 4.86E-8
Public Const cShrinkageFactorSI_Const As Double = 2.69E-04
Public Const cShrinkageFactor_Exp1 As Double = 0.819
Public Const cShrinkageFactor_Exp2 As Double = 2.28

'
Public Const cFloatComparison_Epsilon As Double = 0.000000000001

'
Public Const cErrorSentinel As Double = -999999