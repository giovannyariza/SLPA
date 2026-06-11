Attribute VB_Name = "mdAPIFluidCalcs"
Option Explicit

' ---------------------------------------------------------------------------------------------------------
' MÓDULO: mdAPIFluidCalcs
' DESCRIPCIÓN:
'   Este módulo contiene funciones para calcular varios factores y propiedades de hidrocarburos líquidos
'   basados en las normas API MPMS, principalmente el Capítulo 11.1 (Factores de Corrección de Volumen por
'   Temperatura y Presión) y el Capítulo 9.3 (Determinación de Densidad por Hidrómetro).
'   Las funciones aquí contenidas dependen principalmente de las PROPIEDADES DEL FLUIDO y las condiciones
'   de operación (temperatura, presión), no de las características del equipo de medición (tanque, medidor).
'
' DEPENDENCIAS:
'   - mdGlobals.bas: Para definiciones de enumeraciones (Enums) y constantes API MPMS.
'   - mdConversion.bas: Para funciones de conversión de unidades (CONVTEMP, CONVDENS, CONVTEMP68).
'
' FUNCIONES FUNDAMENTALES (API MPMS Cap. 9 y 11):
'   - HYC: Factor de Corrección por Temperatura del Hidrómetro (API MPMS 9.3).
'   - DENSHYC: Densidad/API corregida por Temperatura del Hidrómetro (API MPMS 9.3).
'   - ALPHA60: Coeficiente de Expansión Térmica a 60F (API MPMS 11.1 Apéndice E/F).
'   - FP: Factor de Compresibilidad Escalado (API MPMS 11.1).
'   - CTL: Factor de Corrección por Temperatura del Líquido (API MPMS 11.1).
'   - CPL: Factor de Corrección por Presión del Líquido (API MPMS 11.1).
'   - API60F: Calcula API a 60F y Presión de Equilibrio (Iterativo API MPMS 11.1 Apéndice E).
'   - APIOBS: Calcula API Observada (a presión base) a partir de API60 (Derivado de API MPMS 11.1).
'   - ROUNDAPI: Redondeo de valores según reglas específicas API MPMS 11.1.
' ---------------------------------------------------------------------------------------------------------

''' <summary>
''' Calcula el Factor de Corrección por Temperatura del Hidrómetro (HYC).
''' Basado en API MPMS Capítulo 9.3 (Corrección por expansión del vidrio).
''' </summary>
''' <param name="TmpObs">Temperatura observada del líquido durante la lectura.</param>
''' <param name="TmpUnits">Unidad de la temperatura (F, C, K, R).</param>
''' <returns>Factor de corrección (Double). Retorna 1.0 si hay error (factor neutro).</returns>

Public Function HYC(ByVal TmpObs As Double, Optional TmpUnits As eTmpUnits = F) As Double
  On Error GoTo ErrHandler ' Manejo de errores
  
  ' Validar finitud del parámetro de entrada inicial
  If Not mdHelpers.IsFinite(TmpObs) Then GoTo ErrHandler

  ' Coeficientes HYC
  Const cHYCF_Coef1 As Double = 0.00001278
  Const cHYCF_Coef2 As Double = 0.0000000062 ' Escala Fahrenheit
  Const cHYCC_Coef1 As Double = 0.000233
  Const cHYCC_Coef2 As Double = 0.00000023 ' Escala Celsius

  Dim tempObsF As Double ' Temperatura observada convertida a Fahrenheit
  Dim tempObsC As Double ' Temperatura observada convertida a Celsius
  Dim deltaT As Double ' Temperatura usada en la fórmula (F o C)
  
  ' Convertir temperatura observada a Fahrenheit y Celsius para su uso posterior
  Select Case TmpUnits
    Case F, R:
      tempObsF = mdConversion.CONVTEMP(TmpObs, IIF(TmpUnits = F, F, R2F))
      deltaT = tempObsF - cBaseTempF
      HYC = 1 - (cHYCF_Coef1 * deltaT) - (cHYCF_Coef2 * (deltaT ^ 2))
    Case C, K:
      tempObsC = mdConversion.CONVTEMP(TmpObs, IIF(TmpUnits = C, C, K2C))
      deltaT = tempObsC - cBaseTempC
      HYC = 1 - (cHYCC_Coef1 * deltaT) - (cHYCC_Coef2 * (deltaT ^ 2))
    Case Else:
      HYC = 1 ' Unidad no reconocida
  End Select
  
  ' Validar finitud del resultado final
  If Not mdHelpers.IsFinite(HYC) Then HYC = 1
  Exit Function

ErrHandler:
  Debug.Print "Error en HYC: " & Err.Description  
  HYC = 1 ' Retorno de seguridad sin correccion
End Function

''' <summary>
''' Calcula la densidad o API corregido por hidrómetro.
''' Convierte la lectura física a densidad absoluta para aplicar el factor HYC.
''' </summary>
''' <param name="DensObs">Densidad observada del líquido durante la lectura.</param>
''' <param name="TmpObs">Temperatura observada del líquido durante la lectura.</param>
''' <param name="DensUnits">Unidad de densidad (API, KGM, SGU).</param>
''' <param name="TmpUnits">Unidad de la temperatura (F, C, K, R).</param>
''' <returns>Densidad corregida por temperatura del hidrometro (Double). Retorna 0.0 si hay error (factor neutro).</returns>

Public Function DENSHYC(ByVal DensObs As Double, ByVal TmpObs As Double, Optional DensUnits As eDnsUnits = API, Optional TmpUnits As eTmpUnits = F) As Double
  On Error GoTo ErrHandler ' Manejo de errores

  ' Validar finitud del parámetro de entrada inicial
  If Not mdHelpers.IsFinite(DensObs) Or Not mdHelpers.IsFinite(TmpObs) Then GoTo ErrHandler

  Dim CorrHyd As Double ' Factor de corrección por hidrómetro
  Dim DensObsInKGM As Double ' Densidad observada convertida a Kg/m³ (base intermedia)
  Dim DensCorrInKGM As Double ' Densidad corregida en Kg/m³
  
  ' Convertir la densidad observada a Kg/m³ como base para la corrección
  Select Case DensUnits
    Case API
        If DensObs <= -131.5 Then GoTo ErrHandler ' Límite físico API
        DensObsInKGM = mdConversion.CONVDENS(DensObs, A2K, True)
    Case KGM
        If DensObs <= 0 Then GoTo ErrHandler
        DensObsInKGM = DensObs
    Case SGU
        If DensObs <= 0 Then GoTo ErrHandler
        DensObsInKGM = mdConversion.CONVDENS(DensObs, S2K, True)
  End Select

  ' Obtener el factor de corrección por hidrómetro (HYC)
  CorrHyd = HYC(TmpObs, TmpUnits)
  
  ' Aplicar la corrección multiplicando la densidad en Kg/m³ por el factor HYC
  DensCorrInKGM = DensObsInKGM * CorrHyd
  
  ' Convertir la densidad corregida de Kg/m³ de nuevo a las unidades originales de salida
  Select Case DensUnits
    Case API
      DENSHYC = mdConversion.CONVDENS(DensCorrInKGM, K2A, True)
    Case KGM
      DENSHYC = DensCorrInKGM
    Case SGU
      DENSHYC = mdConversion.CONVDENS(DensCorrInKGM, K2S, True)
  End Select
  Exit Function

ErrHandler:
  Debug.Print "Error en DENSHYC: " & Err.Description
  DENSHYC = 0 ' Retorno de seguridad sin correccion
End Function

''' <summary>
''' Calcula el Coeficiente de Expansión Térmica a 60°F (Alfa60) y la Densidad en escala IPTS-68.
''' Estándar: API MPMS 11.1.
''' </summary>
''' <param name="API60">Gravedad API a 60°F (ITS-90).</param>
''' <param name="TypeLiq">Tipo de líquido (Crudo, Refinado, Lubricante).</param>
''' <param name="Rho68_OUT">Opcional ByRef: Retorna la densidad ajustada a IPTS-68 [Kg/m³].</param>
''' <returns>Coeficiente Alfa60 [1/°F]. Retorna 0 si hay error.</returns>

Public Function ALPHA60(ByVal API60 As Double, Optional TypeLiq As eTypeLiq = CRD, Optional ByRef Rho68 As Double) As Double
  On Error GoTo ErrHandler ' Manejo de errores

  ' Validar finitud del parámetro de entrada inicial
  If Not mdHelpers.IsFinite(API60) Then GoTo ErrHandler
  If API60 <= -131.5 Then GoTo ErrHandler

  Dim Rho60 As Double ' Densidad a 60F (ITS-90) en Kg/m³ relativo al agua
  Dim K0 As Double, K1 As Double, K2 As Double ' Constantes de las tablas API
  
  ' Convertir la Gravedad API60 (ITS-90)
  Rho60 = mdConversion.CONVDENS(API60, A2K, True)
  
  ' Selección de Constantes K (API MPMS 11.1 Tablas 6A, 6B, 6D)
  Select Case TypeLiq
    Case CRD ' Tabla 6A: Crudos
      K0 = 341.0957: K1 = 0: K2 = 0
        
    Case LUB ' Tabla 6D: Lubricantes
      K0 = 0: K1 = 0.34878: K2 = 0
        
    Case REF ' Tabla 6B: Productos Refinados (Selección por rangos de densidad)
      If Rho60 >= 838.3127 Then       ' Fuel Oils
        K0 = 103.872: K1 = 0.2701: K2 = 0
      ElseIf Rho60 >= 787.5195 Then   ' Jet Fuels
        K0 = 330.301: K1 = 0: K2 = 0
      ElseIf Rho60 >= 770.352 Then    ' Transition Zone
        K0 = 1489.067: K1 = 0: K2 = -0.0018684
      ElseIf Rho60 >= 610.6 Then      ' Gasolines
        K0 = 192.4571: K1 = 0.2438: K2 = 0
      Else
        ' Fuera de rango para Refinados según estándar
        GoTo ErrHandler
      End If
  End Select

  ' Calcular Alfa60 (en 1/°F) usando las constantes K y Rho60. Fórmula de API MPMS 11.1 (2007) Apéndice E
  ' Alfa60 = K0/Rho60^2 + K1/Rho60 + K2
  ALPHA60 = (K0 / (Rho60 ^ 2)) + (K1 / Rho60) + K2
  
  ' Calcular Rho68 (densidad a 60F IPTS-68 en Kg/m³) a partir de Rho60 (ITS-90). API MPMS 11.1 (2007) Apéndice F.  
  Rho68 = Rho60 * Exp(ALPHA60 * cTmpShift)
  Exit Function

ErrHandler:
  Debug.Print "Error en ALPHA60: " & Err.Description
  ALPHA60 = 0 ' Retorno de seguridad sin correccion
  Rho68 = 0
End Function

''' <summary>
''' Calcula el Coeficiente de Compresibilidad Escalado (Fp).
''' Basado en API MPMS 11.1 (2007).
''' </summary>
''' <param name="API60">Gravedad API a 60°F (ITS-90).</param>
''' <param name="TempF">Temperatura observada en Fahrenheit (ITS-90).</param>
''' <param name="TypeLiq">Tipo de líquido (CRD, REF, LUB).</param>
''' <returns>Factor de compresibilidad escalado. Retorna 0 si hay error.</returns>

Public Function FP(ByVal API60 As Double, ByVal TempF As Double, Optional TypeLiq As eTypeLiq = CRD) As Double
  On Error GoTo ErrHandler ' Manejo de errores

  ' Validaciones de Integridad y Finitud (Uso de mdHelpers)
  If Not mdHelpers.IsFinite(API60) Or Not mdHelpers.IsFinite(TempF) Then GoTo ErrHandler

  ' Validaciones de Rangos Físicos (Límites estándar API 11.1)
  If API60 <= -131.5 Then GoTo ErrHandler

  ' Coeficientes FP: Estos valores están escalados para trabajar con densidad en Kg/m³
  Const A As Double = -1.9947
  Const B As Double = 0.00013427
  Const C As Long = 793920
  Const D As Integer = 2326

  ' Temperatura observada convertida a IPTS-68 Celsius
  Dim TempC68 As Double

  TempC68 = mdConversion.CONVTEMP68(TempF, C)

  ' Densidad a 60F ajustada a IPTS-68 en Kg/m³
  Dim Rho68 As Double
  ' Llamamos a ALPHA60 solo para obtener Rho68.  
  Call ALPHA60(API60, TypeLiq, Rho68) ' Llama a ALPHA60 y pasa Rho68 ByRef

  ' Validación de seguridad física
  If rho68 <= 0 GoTo ErrHandler

  ' Calcular el Coeficiente de Compresibilidad Escalado (FP)
  ' Fórmula: exp[A + B*t_c68 + (C + D*t_c68)/rho68^2]
  Dim Term1 As Double, Term2 As Double
  
  Term1 = A + B * TempC68
  Term2 = (C + D * TempC68) / (Rho68 ^ 2)

  Dim exponente As Double
  
  exponente = Term1 + Term2

  ' Prevención de desbordamiento (Overflow)
  If exponente > 700 Then 
    FP = 0
    Exit Function
  End If

  ' Resultado Final
  FP = Exp(exponente)

  ' Validaciones de Finitud
  If Not mdHelpers.IsFinite(FP) Then FP = 0
  Exit Function

ErrHandler:
  Debug.Print "Error en FP: " & Err.Description
  FP = 0 ' Retorno de seguridad sin correccion
End Function

''' <summary>
''' Calcula el Factor de Corrección por Temperatura del Líquido (CTL / VCF).
''' Basado en API MPMS Capítulo 11.1 (2007/2019).
''' </summary>
''' <param name="API60">Gravedad API a 60°F (ITS-90).</param>
''' <param name="TempF">Temperatura observada en Fahrenheit (ITS-90).</param>
''' <param name="TypeLiq">Tipo de líquido (CRD, REF, LUB).</param>
''' <returns>Factor CTL (Double). Retorna 1.0 (factor neutro) si hay error.</returns>

Public Function CTL(ByVal API60 As Double, ByVal TempF As Double, Optional TypeLiq As eTypeLiq = CRD) As Double
  On Error GoTo ErrHandler ' Manejo de errores

  ' Validaciones de Integridad y Finitud (mdHelpers)
  If Not mdHelpers.IsFinite(API60) Or Not mdHelpers.IsFinite(TempF) Then GoTo ErrHandler

  ' Validaciones de Rangos Físicos (Límites estándar API 11.1)
  If API60 <= -131.5 Or TempF < -350 Or TempF > 600 Then GoTo ErrHandler

  Dim Alfa60 As Double      
  ' Coeficiente de expansión térmica a 60F (en 1/°F)
  Alfa60 = ALPHA60(API60, TypeLiq)

  ' Si Alfa60 es 0 o inválido, no podemos calcular CTL
  If Alfa60 <= 0 Then GoTo ErrHandler

  Dim TempF68 As Double
  ' Convertir temperatura observada de ITS-90 F a IPTS-68 F
  TempF68 = mdConversion.CONVTEMP68(TempF, F)

  Dim DeltaT As Double
  ' Calcular la diferencia de temperatura con la temperatura base 60F (en IPTS-68 Fahrenheit)
  DeltaT = TempF68 - 60  
  
  ' Calcular el Factor de Corrección por Temperatura (CTL)
  ' Fórmula (API MPMS 11.1 2007 Pág 15): exp[-Alpha_60 * Delta_t * (1 + 0.8 * Alpha_60 * Delta_t)]
  Dim exponente As Double
  Const cTaylor As Double = 0.8
  
  exponente = -Alfa60 * DeltaT * (1 + cTaylor * Alfa60 * DeltaT)

  ' Prevención de desbordamiento (Overflow)
  If exponente > 700 Then 
    CTL = 1
    Exit Function
  ElseIf exponente < -700 Then
    CTL = 0 ' El volumen se vuelve virtualmente cero (temperatura extrema)
    Exit Function
  End If

  ' Resultado Final
  CTL = Exp(exponente)

  ' Validaciones de Integridad y Finitud
  If Not mdHelpers.IsFinite(CTL) Or CTL <= 0 Then CTL = 1  
  Exit Function

ErrHandler:
  Debug.Print "Error crítico en CTL (API60: " & API60 & ", Temp: " & TempF & "): " & Err.Description
  CTL = 1 ' Retorno de seguridad sin correccion
End Function

''' <summary>
''' Calcula el Factor de Corrección por Presión del Líquido (CPL).
''' Basado en API MPMS Capítulo 11.1.
''' </summary>
''' <param name="API60">Gravedad API a 60°F (ITS-90).</param>
''' <param name="TempF">Temperatura observada en Fahrenheit (ITS-90).</param>
''' <param name="Pres">Presión observada en PSI.</param>
''' <param name="TypeLiq">Tipo de líquido (CRD, REF, LUB).</param>
''' <param name="Pe">Presión de equilibrio en PSI (Presión de vapor). Por defecto 0.</param>
''' <returns>Factor CPL (Double). Retorna 1.0 si hay error o el factor es inválido.</returns>

Public Function CPL(ByVal API60 As Double, ByVal TempF As Double, ByVal Pres As Double, Optional TypeLiq As eTypeLiq = CRD, Optional ByVal Pe As Double = 0) As Double
  On Error GoTo ErrHandler ' Manejo de errores

  ' 1. Validaciones de Integridad y Finitud (Uso de mdHelpers)
  If Not mdHelpers.IsFinite(API60) Or Not mdHelpers.IsFinite(TempF) Or _
     Not mdHelpers.IsFinite(Pres) Or Not mdHelpers.IsFinite(Pe) Then GoTo ErrHandler
  
  ' Validaciones de Rangos Físicos API 11.1
  If Pres < Pe Or Pres > 20000 Or API60 <= -131.5 Then GoTo ErrHandler

  Dim Fcp As Double
  ' Hallar el Factor de Compresibilidad Escalado (FP)
  Fcp = FP(API60, TempF, TypeLiq)

  ' Si Fp es 0 o inválido, no podemos calcular CPL (retorno neutro)
  If Fcp <= 0 Then GoTo ErrHandler

  Dim DeltaP As Double
  ' Calcular la diferencia de presión (Pres - Pe) en PSI
  DeltaP = Pres - Pe

  Dim Denominador As Double ' Denominador en la fórmula CPL
  ' Calcular el Factor de Corrección por Presión (CPL)
  ' Fórmula: 1 / (1 - Fcp * (Pres - Pe) * 10 ^ -5)  
  Denominador = 1 - (Fcp * DeltaP * 0.00001) ' 10^-5 es 0.00001
  
  ' Protección contra indeterminación o resultados físicamente imposibles
  If Abs(Denominador) < 0.1 Then
    CPL = 1
    Exit Function
  End If

  ' Resultado Final
  CPL = 1 / Denominador

  ' Validaciones de Integridad y Finitud
  If Not mdHelpers.IsFinite(CPL) Or CPL < 1# Then CPL = 1
  Exit Function

ErrHandler:
  Debug.Print "Error en CPL (API60: " & API60 & ", P: " & Pres & "): " & Err.Description
  CPL = 1 ' Retorno de seguridad sin correccion
End Function

''' <summary>
''' Calcula la Gravedad API a 60°F y Presión de Equilibrio mediante el método de Newton-Raphson.
''' Basado en API MPMS Capítulo 11.1 (Apéndice E).
''' </summary>
''' <param name="APIOBS">Gravedad API observada (corregida por hidrómetro).</param>
''' <param name="TempObs">Temperatura observada en Fahrenheit (ITS-90).</param>
''' <param name="TypeLiq">Tipo de líquido (CRD, REF, LUB).</param>
''' <param name="PresObs">Presión observada en PSI. Por defecto 0.</param>
''' <param name="Pe">Presión de equilibrio en PSI. Por defecto 0.</param>
''' <returns>Gravedad API a 60°F. Retorna 0 si no hay convergencia o error.</returns>

Function API60F(ByVal APIOBS As Double, ByVal TempObs As Double, Optional TypeLiq As eTypeLiq = CRD, Optional ByVal PresObs As Double = 0, Optional ByVal Pe As Double = 0) As Double
  On Error GoTo ErrHandler ' Manejo de errores

  ' Validaciones de Integridad y Finitud (mdHelpers)
  If Not mdHelpers.IsFinite(APIOBS) Or Not mdHelpers.IsFinite(TempObs) Then Exit Function
  If APIOBS <= -131.5 Then Exit Function
  
  ' Convertir temperatura observada de ITS-90 F a IPTS-68 F y C UNA VEZ fuera del bucle
  Dim TempObs_F68 As Double
  TempObs_F68 = mdConversion.CONVTEMP68(TempObs, F)
  
  Dim TempObs_C68 As Double
  TempObs_C68 = mdConversion.CONVTEMP68(TempObs, C)
  
  ' Calcular Delta_t en Fahrenheit IPTS-68
  Dim DeltaT_F68 As Double
  DeltaT_F68 = TempObs_F68 - 60

  ' Convertir APIOBS (ITS-90) a Densidad observada en Kg/m³ (RhoObs)
  Dim RhoObs As Double
  RhoObs = mdConversion.CONVDENS(APIOBS, A2K, True)

  Dim Rho60 As Double ' Densidad a 60F (ITS-90) en Kg/m³ (Valor iterado)
  ' Inicializar el valor de Rho60 (ITS-90) para la iteración. Usar RhoObs como punto de partida.
  Rho60 = RhoObs
  
  ' Iniciar el proceso iterativo (Método de Newton)
  Const maxIterations As Long = 20 ' Número máximo de iteraciones
  Const tolerance As Double = 0.00001 ' Tolerancia para la diferencia de densidad (Kg/m³)
  Dim m As Byte ' Contador de iteraciones
  Dim API60_Iter As Double  ' API60 (ITS-90) correspondiente a Rho60_Iter
  Dim Alfa60_Iter As Double ' Alfa60 (1/°F) para API60_Iter
  Dim Rho68_Iter As Double  ' Densidad a 60F IPTS-68 (Kg/m³) para API60_Iter (necesaria para FP)
  Dim CTL_Iter As Double    ' CTL para API60_Iter y TempObs
  Dim CPL_Iter As Double    ' CPL para API60_Iter, TempObs, PresObs, Pe
  Dim CTPL_Iter As Double   ' CTPL para API60_Iter, TempObs, PresObs, Pe
  Dim Em as Double          ' Función de error
  Dim Da As Double          ' Coeficiente Da para la derivada
  Dim Dt As Double          ' Derivada correspondiente a la temperatura
  Dim FP_Iter As Double     ' FP para API60_Iter y TempObs
  Dim Dp As Double
  Dim Dp_Num As Double
  Dim Dp_Denom As Double  
  Dim DRho60 as Double      ' Corrección (DRho60) usando la derivada (Método de Newton)

  For m = 1 To maxIterations
    ' Convertir el Rho60_Iter (supuesto, ITS-90) a API60_Iter (ITS-90)
    API60_Iter = mdConversion.CONVDENS(Rho60, K2A, True)

    ' Obtener Alfa60 (1/°F) y Rho68 (Kg/m³) para el API60_Iter actual
    Alfa60_Iter = ALPHA60(API60_Iter, TypeLiq, Rho68_Iter)

    ' Calcular CTL para el API60_Iter actual
    CTL_Iter = CTL(API60_Iter, TempObs, TypeLiq)

    ' Calcular CPL para el API60_Iter actual
    CPL_Iter = CPL(API60_Iter, TempObs, PresObs, TypeLiq, Pe)

    ' Calcular CTPL_Iter como el producto de CTL_Iter y CPL_Iter
    CTPL_Iter = CTL_Iter * CPL_Iter

    ' Calcular Em (Error)
    Em = (RhoObs / CTPL_Iter) - Rho60

    ' Criterio de parada
    If Abs(Em) < tolerance Then
      API60F = API60_Iter
      Exit Function
    End If

    ' Cálculo de la Derivada para el ajuste (Newton Step)
    ' Obtener coeficiente Da (ajuste de expansión diferencial)
    Da = GetDaCoefficient(TypeLiq, Rho60)

    ' Calcular el término Dt (API MPMS 11.1 2007 Apéndice E E.3)
    ' Dt = Da * Alpha_60 * Delta_t * (1 + 1.6 * Alpha_60 * Delta_t)
    Dt = Da * Alfa60_Iter * DeltaT_F68 * (1 + 1.6 * Alfa60_Iter * DeltaT_F68)

    ' Obtener FP_Iter para el API60_Iter actual y TempObs (ITS-90 F). CPL llamó a FP, pero la derivada necesita FP_Iter explícitamente.
    FP_Iter = FP(API60_Iter, TempObs, TypeLiq)

    ' Calcular el término Dp (API MPMS 11.1 2007 Apéndice E E.5)
    ' Dp = -(2 * CPL_m * P_obs * F_cp_m * (7.9392 + 0.02326 * t_obs_F68)) / (rho_m^2 * Alpha_60_m)
    ' P_obs es PresObs (en PSI), t_obs_F68 es TempObs_F68, rho_m es Rho60, Alpha_60_m es Alfa60_Iter
    Dp_Num = -(2 * CPL_Iter * PresObs * FP_Iter * (7.9392 + 0.02326 * TempObs_F68))
    Dp_Denom = (Rho60 ^ 2 * Alfa60_Iter)
    Dp = Dp_Num / Dp_Denom

    ' Calcular el siguiente paso de corrección (DRho60)
    DRho60 = Em / (1 + Dt + Dp)
    
    ' Valor de Rho60 (ITS-90) para la siguiente iteración
    Rho60 = Rho60 + DRho60

    ' Seguridad: Si la densidad se vuelve negativa o irreal, abortar
    If Rho60 <= 0 Then Exit For
  Next m ' Siguiente iteración

  ' Si llega aquí, no hubo convergencia
  Debug.Print "API60F: No se alcanzó convergencia en " & maxIterations & " iteraciones."
  API60F = 0
  Exit Function

ErrHandler:
  Debug.Print "Error en API60F: " & Err.Description
  API60F = 0 ' Retorno de seguridad sin correccion
End Function

''' <summary>
''' Helper privado para obtener el coeficiente Da según el tipo de líquido.
''' Basado en API MPMS 11.1 Apéndice E.
''' </summary>
''' <param name="TypeLiq">Tipo de líquido (CRD, REF, LUB).</param>
''' <param name="Rho60">Densidad a 60F (ITS-90) en Kg/m³ relativo al agua</param>
''' <returns>Coeficiente para el cálculo del factor de corrección en la derivada (Da) por tipo de producto.</returns>

Private Function GetDaCoefficient(ByVal TypeLiq As eTypeLiq, ByVal Rho60 As Double) As Double
  Select Case TypeLiq
    Case CRD: GetDaCoefficient = 2   ' Crudos (Tabla 6A)
    Case LUB: GetDaCoefficient = 1   ' Lubricantes (Tabla 6D)
    Case REF ' Productos Refinados (Tabla 6B)
      If Rho60 >= 838.3127 Then
        GetDaCoefficient = 1.3       ' Fuel Oils
      ElseIf Rho60 >= 787.5195 Then
        GetDaCoefficient = 2         ' Jet Fuels
      ElseIf Rho60 >= 770.352 Then
        GetDaCoefficient = 8.5       ' Transition
      Else
        GetDaCoefficient = 1.5       ' Gasolines
      End If
    Case Else
      GetDaCoefficient = 0
  End Select
End Function

''' <summary>
''' Calcula la Gravedad API observada a una temperatura dada partiendo del API a 60°F.
''' Asume condiciones de presión base (atmosférica).
''' </summary>
''' <param name="API60">Gravedad API base a 60°F (ITS-90).</param>
''' <param name="TempF">Temperatura a la cual se desea proyectar el API [°F].</param>
''' <param name="TypeLiq">Tipo de líquido (CRD, REF, LUB).</param>
''' <returns>Gravedad API proyectada (Double). Retorna 0 si hay error.</returns>

Public Function APIOBS(ByVal API60 As Double, ByVal TempF As Double, Optional TypeLiq As eTypeLiq = CRD) As Double
  On Error GoTo ErrHandler ' Manejo de errores

  ' 1. Validaciones de Integridad (mdHelpers)
  If Not mdHelpers.IsFinite(API60) Or Not mdHelpers.IsFinite(TempF) Then GoTo ErrHandler
  
  ' Límites físicos de la escala API
  If API60 <= -131.5 Then GoTo ErrHandler
  
  Dim Rho60 As Double
  ' Convertir API60 (BASE) a Densidad (ITS-90) en Kg/m³ a 60F
  Rho60 = mdConversion.CONVDENS(API60, A2K, True)
  
  Dim Ftl As Double
  ' Calcular el Factor de Corrección por Temperatura del Líquido (CTL)
  Ftl = CTL(API60, TempF, TypeLiq)
  
  Dim RhoObs As Double ' Densidad en Kg/m³ a la Temperatura Observada
  ' Calcular la Densidad en Kg/m³ a la Temperatura Observada (a presión base/atmosférica)
  ' Fórmula: Densidad@Obs = Densidad@Base * CTL (Esto es una simplificación para APIOBS)
  ' Nota: Para correcciones completas con presión, sería Densidad@Obs = Densidad@Base * CTL * CPL
  RhoObs = Rho60 * Ftl

  ' Convertir la densidad resultante de nuevo a escala API
  If RhoObs <= 0 Then GoTo ErrHandler
  
  ' Convertir la Densidad a la Temperatura Observada (en Kg/m³) a Gravedad API
  APIOBS = mdConversion.CONVDENS(RhoObs, K2A, True)

  ' Validación de Finitud
  If Not mdHelpers.IsFinite(APIOBS) Then APIOBS = 0
  Exit Function

ErrHandler:
  Debug.Print "Error en APIOBS: " & Err.Description
  APIOBS = 0 ' Retorno de seguridad sin correccion
End Function

''' <summary>
''' Redondea un valor según las reglas API MPMS 11.1 (Redondeo al par más cercano).
''' Permite incrementos personalizados (ej. redondear al 0.05 más cercano).
''' </summary>
''' <param name="Val">Valor numérico a redondear.</param>
''' <param name="PosDec">Número de posiciones decimales.</param>
''' <param name="EndNum">Último dígito del incremento (ej. 1 para 0.1, 5 para 0.05). Defecto 1.</param>
''' <returns>Valor redondeado (Double). Retorna el valor original si hay error.</returns>

Public Function ROUNDAPI(ByVal Val As Double, ByVal PosDec As Long, Optional ByVal EndNum As Byte = 1) As Double
  On Error GoTo ErrHandler ' Manejo de errores

  ' Validaciones de Integridad (mdHelpers)
  If Not mdHelpers.IsFinite(Val) Then 
    ROUNDAPI = Val
    Exit Function
  End If
  
  ' PosDec debe ser positivo y EndNum mayor a 0
  If PosDec < 0 Or EndNum <= 0 Then 
    ROUNDAPI = Val
    Exit Function
  End If

  Dim Inc As Double
  ' Determinar incremento de redondeo
  Inc = (10 ^ -PosDec) * EndNum
  
  Dim NormVal As Double
  ' Normalizar el valor de acuerdo al incremento y obtener su parte decimal
  NormVal = Abs(Val) / Inc

  Dim IntPart As Long
  ' Obtener la parte entera del valor normalizado
  IntPart = Fix(NormVal)
  
  Dim DecPart As Double
  ' Obtener la parte decimal del valor normalizado
  DecPart = NormVal - IntPart
  
  Const EPSILON As Double = 0.0000000001 ' Pequeña tolerancia para comparación de punto flotante
  Dim ValRounded as Double ' Valor truncado para el redondeo especial de 0.5  
  
  ' Truncar el valor normalizado para el redondeo especial de 0.5 Regla API:
  ' Si la parte decimal es EXACTAMENTE 0.5, redondear al ENTERO PAR más cercano.
  ' Si la parte decimal es > 0.5, redondear hacia arriba (IntPart + 0.5) si no fuera por la regla API.
  ' Si la parte decimal es < 0.5, redondear hacia abajo IntPart.
  If Abs(DecPart - 0.5) < EPSILON Then ' Si la parte fraccionaria es muy cercana a 0.5
    ' La parte fraccionaria es 0.5. Verificar si el ENTERO de NormVal es par o impar.
    If (IntPart Mod 2) <> 0 Then ' Si el entero es IMPAR
      ValRounded = IntPart + 1 ' Redondear hacia arriba para hacerlo PAR
    Else ' Si el entero es PAR
      ValRounded = IntPart ' Truncar (redondear hacia abajo) para mantenerlo PAR
    End If
  Else
    ' La parte fraccionaria NO es 0.5. Redondeo estándar al entero más cercano.
    ValRounded = Int(NormVal + 0.5)
  End If

  ' Mantener el signo del valor original y aplicar el incremento
  ROUNDAPI = Sgn(Val) * Inc * ValRounded

  ' Verificación de Finitud
  If Not mdHelpers.IsFinite(ROUNDAPI) Then ROUNDAPI = Val
  Exit Function

ErrHandler:
  Debug.Print "Error en ROUNDAPI: " & Err.Description
  ROUNDAPI = Val
End Function