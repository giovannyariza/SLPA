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

Public Const cTEMPBASE_F As Double = 60
Public Const cTEMPBASE_C As Double = 15.56

' Temperatura base de referencia de 60 grados Fahrenheit (60 °F) expresada en la escala de temperatura IPTS-68.
Public Const cTEMPBASE_F68 As Double = 60.0068748977356
Public Const cTEMPBASE_C68 As Double = 15.5555555555556

Public Const cTEMPDIFF_F68 As Double = 0.0068748977356

' Factor de shift de temperatura utilizado en algunas fórmulas de factores de corrección API MPMS 11.1 (2007).
Public Const cTEMPSHIFT As Double = 1.37497954711989E-02

' Densidad del agua a 60 grados Fahrenheit (15.56 °C) en Kilogramos por Metro Cúbico (Kg/m³).
Public Const cWATERDENSKG_60F As Double = 999.016

' Rangos de validación de temperatura utilizado en algunas fórmulas de factores de corrección API MPMS 11.1 (2019).
Public Const cTEMPVALIDRANGE_MIN As Double = -350
Public Const cTEMPVALIDRANGE_MAX As Double = 600

' Constantes de Escala de Temperatura para Corrección ITS-90 a IPTS-68
Public Const cTEMPSCALE_MIN As Double = -183
Public Const cTEMPSCALE_MAX As Double = 630

' Rangos de validación de presión utilizado en algunas fórmulas de factores de corrección API MPMS 11.1 (2019).
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

' Coeficientes FP: Compresibilidad
Public Const cFP_A As Double = -1.9947
Public Const cFP_B As Double = 0.00013427
Public Const cFP_C As Long = 793920
Public Const cFP_D As Integer = 2326
Public Const cPRESS_SCALING_API As Integer = 2326

' Coeficientes Expansión Térmica del Casco
Public Const cCTSH_MCRBN_F As Double = 0.0000062
Public Const cCTSH_MCRBN_C As Double = 0.0000112
Public Const cCTSH_ST304_F As Double = 0.0000096
Public Const cCTSH_ST304_C As Double = 0.0000173
Public Const cCTSH_ST316_F As Double = 0.00000883
Public Const cCTSH_ST316_C As Double = 0.0000159
Public Const cCTSH_ST4PH_F As Double = 0.000006
Public Const cCTSH_ST4PH_C As Double = 0.0000108

' Propiedades de Materiales (API 650 / 12.1.1)
Public Const cYOUNG_MODULUS_STEEL_IMP As Double = 30000000
Public Const cYOUNG_MODULUS_STEEL_MET As Double = 206842718900
' -------------------------------------------------------------------------------------------------
' PARA ORGANIZAR
' -------------------------------------------------------------------------------------------------

' Rango de Gravedad API a 60°F para Crudos y Productos Refinados
Public Const cAPI60_RangeCrudeRefined_MIN As Double = 0
Public Const cAPI60_RangeCrudeRefined_MAX As Double = 130

' Rango de Gravedad API a 60°F para Aceites Lubricantes
Public Const cAPI60_RangeLubricant_MIN As Double = 0
Public Const cAPI60_RangeLubricant_MAX As Double = 40

'
Public Const cShrinkageFactorCU_Const As Double = 4.86E-8
Public Const cShrinkageFactorSI_Const As Double = 2.69E-04
Public Const cShrinkageFactor_Exp1 As Double = 0.819
Public Const cShrinkageFactor_Exp2 As Double = 2.28

'
Public Const cFloatComparison_Epsilon As Double = 0.000000000001

'
Public Const cErrorSentinel As Double = -999999

-------------------------------------------------------------------------------------------------
' Constantes de Desplazamiento de Temperatura
Public Const cC2K_Offset As Double = 273.15
Public Const cC2R_Offset As Double = 491.67
Public Const cF2R_Offset As Double = 459.67
' -------------------------------------------------------------------------------------------------
' Constantes de Escala de Temperatura para Corrección ITS-90 a IPTS-68
Public Const cTempScale_Min As Double = -183
Public Const cTempScale_Max As Double = 630
' -------------------------------------------------------------------------------------------------
' Constantes de Conversión de Temperatura
Public Const cC2F_Factor As Double = 9 / 5
Public Const cF2C_Factor As Double = 5 / 9
' -------------------------------------------------------------------------------------------------
' Constantes de Conversión de Presión
Public Const cPSI2KPA As Double = 6.8947590868
Public Const cPSI2BAR As Double = 6.8947590868E-2
' -------------------------------------------------------------------------------------------------
' Constantes de Conversión de Volumen
' https://www.convertworld.com/es/volumen/
Public Const cBBL2GAL As Double = 42.000008585
Public Const cBBL2M3 As Double = 0.15898723857
Public Const cBBL2LT As Double = 158.98723857
' -------------------------------------------------------------------------------------------------
' Constantes de Conversión de Masa
' https://www.convertworld.com/es/masa/
Public Const cKG2TON As Double = 0.001
Public Const cKG2LB As Double = 2.2046226218
' -------------------------------------------------------------------------------------------------
' Constantes de Conversión de Longitud
' https://www.convertworld.com/es/longitud/
Public Const cMT2MM As Double = 1000
Public Const cMT2CM As Double = 100
Public Const cMT2FT As Double = 3.280839895
Public Const cMT2IN As Double = 39.37007874
' -------------------------------------------------------------------------------------------------
' Constantes de Conversión de Densidad API
Public Const cAPI_Factor1 As Double = 141.5
Public Const cAPI_Factor2 As Double = 131.5
' -------------------------------------------------------------------------------------------------
' Factor utilizado en el término exponencial para calcular la Densidad a 60F en la escala IPTS-68
Public Const cAPI_Factor08 As Double = 0.8
' -------------------------------------------------------------------------------------------------
' Factor utilizado en el término de la derivada de CTL (Dt) para el cálculo iterativo de API_60
Public Const cAPI_Factor16 As Double = 1.6
' -------------------------------------------------------------------------------------------------
' Rango de Gravedad API a 60°F para Crudos y Productos Refinados
Public Const cAPI60_RangeCrudeRefined_Min As Double = 0
Public Const cAPI60_RangeCrudeRefined_Max As Double = 130
' -------------------------------------------------------------------------------------------------
' Rango de Gravedad API a 60°F para Aceites Lubricantes
Public Const cAPI60_RangeLubricant_Min As Double = 0
Public Const cAPI60_RangeLubricant_Max As Double = 40
' -------------------------------------------------------------------------------------------------
' Rangos de Densidad Relativa para Crudos, Refinados y Lubricantes
Public Const cRho60_RangeCrude_Min As Double = 610.6
Public Const cRho60_RangeCrude_Max As Double = 1163.5
Public Const cRho60_RangeFuelOils_Min As Double = 838.3127
Public Const cRho60_RangeFuelOils_Max As Double = 1163.5
Public Const cRho60_RangeJetFuels_Min As Double = 787.5195
Public Const cRho60_RangeJetFuels_Max As Double = 838.3127
Public Const cRho60_RangeTransition_Min As Double = 770.352
Public Const cRho60_RangeTransition_Max As Double = 787.5195
Public Const cRho60_RangeGasolines_Min As Double = 610.6
Public Const cRho60_RangeGasolines_Max As Double = 770.352
Public Const cRho60_RangeLubricant_Min As Double = 800.9
Public Const cRho60_RangeLubricant_Max As Double = 1163.5
' -------------------------------------------------------------------------------------------------
' Coeficientes FP
Public Const cFP_CoefA As Double = -1.9947
Public Const cFP_CoefB As Double = 0.00013427
Public Const cFP_CoefC As Long = 793920
Public Const cFP_CoefD As Integer = 2326
' -------------------------------------------------------------------------------------------------
' Constantes Derivada FP
Public Const cFPDerivative_Const1 As Double = 7.9392
Public Const cFPDerivative_Const2 As Double = 0.02326
' -------------------------------------------------------------------------------------------------
' Coeficientes HYC
Public Const cHYCF_Coef1 As Double = 0.00001278
Public Const cHYCF_Coef2 As Double = 0.0000000062
Public Const cHYCC_Coef1 As Double = 0.000233
Public Const cHYCC_Coef2 As Double = 0.00000023
' -------------------------------------------------------------------------------------------------
' Coeficientes TSH
Public Const cTSH_CoefLiq As Double = 7
Public Const cTSH_Divisor As Double = 8
' -------------------------------------------------------------------------------------------------
' Coeficientes Expansión Térmica del Casco
Public Const cCTSH_CoefMCrbn_F As Double = 0.0000062
Public Const cCTSH_CoefMCrbn_C As Double = 0.0000112
Public Const cCTSH_CoefSt304_F As Double = 0.0000096
Public Const cCTSH_CoefSt304_C As Double = 0.0000173
Public Const cCTSH_CoefSt316_F As Double = 0.00000883
Public Const cCTSH_CoefSt316_C As Double = 0.0000159
Public Const cCTSH_CoefSt4PH_F As Double = 0.000006
Public Const cCTSH_CoefSt4PH_C As Double = 0.0000108
' -------------------------------------------------------------------------------------------------
' Coeficientes Da
Public Const cDA_CoefCrude As Double = 2
Public Const cDA_CoefRefined1 As Double = 1.3
Public Const cDA_CoefRefined2 As Double = 2
Public Const cDA_CoefRefined3 As Double = 8.5
Public Const cDA_CoefRefined4 As Double = 1.5
Public Const cDA_CoefLubricant As Double = 1