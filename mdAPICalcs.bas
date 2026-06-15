Attribute VB_Name = "mdAPICalcs"
Option Explicit

' ---------------------------------------------------------------------------------------------------------
' MÓDULO: mdAPICalcs
' DESCRIPCIÓN:
'   Este módulo contiene funciones para calcular cantidades de hidrocarburos líquidos, basado en las normas 
'   API MPMS. Implementa los estándares internacionales para la transferencia de custodia, permitiendo la 
'   normalización de volúmenes observados a condiciones estándar (60°F / 14.696 psi).
'
' DEPENDENCIAS:
'   - mdGlobals: Para definiciones de enumeraciones (Enums) y constantes API MPMS.
'   - mdConversion: Para funciones de conversión de unidades (CONVTEMP, CONVDENS, CONVTEMP68).
'   - mdHelpers: Validaciones de integridad matemática y finitud (Failsafe).
'
' FUNCIONES FUNDAMENTALES:
' ---------------------------------------------------------------------------------------------------------
' CORRECCIÓN INSTRUMENTAL (API MPMS Cap. 9.3)
'   - HYC: Factor de Corrección de la Escala del Hidrometro por Temperatura.
'   - DENSHYC: Densidad/API corregida por Temperatura del Hidrómetro.
' ---------------------------------------------------------------------------------------------------------
' DINÁMICA DEL FLUIDO (API MPMS Cap. 11.1)
'   - ALPHA60MED: Coeficiente de Expansión Térmica Promedio (Alfa 60 Medio) a partir de una serie de mediciones de densidad y temperatura.
'   - ALPHA60: Coeficiente de Expansión Térmica a 60°F (Alfa60) y la Densidad en escala IPTS-68. (Apéndice E/F).
'   - FP: Factor de Compresibilidad Escalado.
'   - CTL: Factor de Corrección por Temperatura del Líquido.
'   - CPL: Factor de Corrección por Presión del Líquido.
'   - API60F: Calcula API a 60F y Presión de Equilibrio (Iterativo Newton-Raphson, Apéndice E).
'   - APIOBS: Calcula API Observada (a presión base) a partir de API60.
'   - ROUNDAPI: Redondeo de valores según reglas específicas (Banker's Rounding)
' ---------------------------------------------------------------------------------------------------------
' AJUSTES EN TANQUES (API MPMS Cap. 12.1.1): Correcciones geométricas y físicas de la infraestructura de almacenamiento.
'   - TSH: Cálculo de la temperatura ponderada de la coraza del tanque.
'   - CTSH: Factor de corrección por expansión térmica del acero en el cuerpo del tanque.
'   - FRA: Ajuste volumétrico por desplazamiento estático (Principio de Arquímedes) del techo flotante.
' ---------------------------------------------------------------------------------------------------------

' ---------------------------------------------------------------------------------------------------------
' FUNCIONES DE CORRECCIÓN INSTRUMENTAL (API MPMS Cap. 9.3)
' Ajustes por la expansión térmica del vidrio en hidrómetros.
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

  Dim tempObsF As Double ' Temperatura observada convertida a Fahrenheit
  Dim tempObsC As Double ' Temperatura observada convertida a Celsius
  Dim deltaT As Double ' Temperatura usada en la fórmula (F o C)
  
  ' Convertir temperatura observada a Fahrenheit y Celsius para su uso posterior
  Select Case TmpUnits
    Case F, R:
      tempObsF = mdConversion.CONVTEMP(TmpObs, IIF(TmpUnits = F, F, R2F))
      deltaT = tempObsF - cTEMPBASE_F
      HYC = 1 - (cHYCF_A * deltaT) - (cHYCF_B * (deltaT ^ 2))
    Case C, K:
      tempObsC = mdConversion.CONVTEMP(TmpObs, IIF(TmpUnits = C, C, K2C))
      deltaT = tempObsC - cTEMPBASE_C
      HYC = 1 - (cHYCC_C * deltaT) - (cHYCC_D * (deltaT ^ 2))
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
        If DensObs <= -cAPI_B Then GoTo ErrHandler ' Límite físico API
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

' ---------------------------------------------------------------------------------------------------------
' FUNCIONES ASOCIADAS A LA DINÁMICA DEL FLUIDO (API MPMS Cap. 11.1)
' Cálculo de factores de corrección volumétrica basados en la ecuación de estado API.
' ---------------------------------------------------------------------------------------------------------

''' <summary>
''' Calcula el Coeficiente de Expansión Térmica Promedio (Alfa 60 Medio) a partir de una serie de mediciones de densidad y temperatura.
''' Basado en API MPMS 11.1 (2007) - Aplicaciones Especiales.
''' </summary>
''' <param name="DensRange">Rango o Arreglo con lecturas de densidad observada.</param>
''' <param name="TempRange">Rango o Arreglo con lecturas de temperatura observada.</param>
''' <param name="TypeLiq">Tipo de fluido (CRD, REF, LUB).</param>
''' <returns>Promedio de Alfa 60 (Double). Retorna 0 si hay error o datos insuficientes.</returns>

Public Function ALPHA60MED(ByVal DensRange As Variant, ByVal TempRange As Variant, Optional ByVal TypeLiq As eTypeLiq = CRD) As Variant
  On Error GoTo ErrHandler ' Manejador de errores
  
  ' Convertir entradas a arreglos de Double (Handles Ranges and Arrays)
  Dim ArrDens() As Double
  ArrDens = mdHelpers.ConvertToDoubleArray(DensRange)

  Dim ArrTmps() As Double
  ArrTmps = mdHelpers.ConvertToDoubleArray(TempRange)
  
  ' Validar que ambos arreglos tengan la misma dimensión y mínimo 10 datos
  Dim n As Long
  n = UBound(ArrDens)

  If n <> UBound(ArrTmps) Or n < 10 Then
    ALPHA60MED = CVErr(xlErrNum) ' Retorna #NUM! en Excel si los datos son insuficientes
    Exit Function
  End If

  ' 3. Procesamiento Matricial
  Dim i As Long
  Dim SumAlpha As Double
  Dim API60_Iter As Double
  Dim Alfa60_Iter As Double
  
  Dim ValidCounts As Long
  ValidCounts = 0
  
  For i = 1 To n
    ' Solo procesar si ambos datos son finitos (mdHelpers)
    If mdHelpers.IsFinite(ArrDens(i)) And mdHelpers.IsFinite(ArrTmps(i)) Then      
      ' a. Hallar API 60 base para este punto (Newton-Raphson)
      API60_Iter = API60F(ArrDens(i), ArrTmps(i), TypeLiq, 0, 0)
      
      If API60_Iter > 0 Then
        ' b. Hallar Alfa 60 para este API 60
        Alfa60_Iter = ALPHA60(API60_Iter, TypeLiq)
        
        If Alfa60_Iter > 0 Then
          SumAlpha = SumAlpha + Alfa60_Iter
          ValidCounts = ValidCounts + 1
        End If
      End If
    End If
  Next i
  
  ' 4. Calcular Promedio
  If ValidCounts >= 10 Then
    ALPHA60MED = SumAlpha / ValidCounts
  Else
    ALPHA60MED = 0
  End If
  Exit Function

ErrHandler:
  Debug.Print "Error en ALPHA60MED: " & Err.Description
  ALPHA60MED = 0
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

  ' Validación de Finitud (mdHelpers)
  If Not mdHelpers.IsFinite(API60) Then GoTo ErrHandler

  ' Validación de Integridad
  If Not mdHelpers.IsValidAPI(API60, TypeLiq, CheckNormative:=False) Then GoTo ErrHandler

  Dim Rho60 As Double
  ' Convertir la Gravedad API60 (ITS-90) en Kg/m³ relativo al agua
  Rho60 = mdConversion.CONVDENS(API60, A2K, True)

  ' Validación de Integridad
  If Rho60 <= 0 Then GoTo ErrHandler
  
  ' Constantes de las tablas API
  Dim K0 As Double, K1 As Double, K2 As Double  
  ' Selección de Constantes K (API MPMS 11.1 Tablas 6A, 6B, 6D)
  Call GetKConstants(TypeLiq, Rho60, K0, K1, K2)

  ' Verificación de seguridad: si no se encontró un rango válido, abortar
  If K0 = 0 And K1 = 0 And K2 = 0 And TypeLiq <> LUB Then GoTo ErrHandler

  ' Calcular Alfa60 (en 1/°F) usando las constantes K y Rho60. Fórmula de API MPMS 11.1 (2007) Apéndice E
  ' Alfa60 = K0/Rho60^2 + K1/Rho60 + K2
  ALPHA60 = (K0 / (Rho60 ^ 2)) + (K1 / Rho60) + K2
  
  ' Calcular Rho68 (densidad a 60F IPTS-68 en Kg/m³) a partir de Rho60 (ITS-90). API MPMS 11.1 (2007) Apéndice F.  
  Rho68 = Rho60 * Exp(ALPHA60 * cTEMPSHIFT)

  ' Validación de Finitud
  If Not mdHelpers.IsFinite(ALPHA60) Then ALPHA60 = 0
  Exit Function

ErrHandler:
  Debug.Print "Error en ALPHA60: " & Err.Description
  ALPHA60 = 0 ' Retorno de seguridad sin correccion
  Rho68 = 0
End Function

''' <summary>
''' Provee las constantes K0, K1 y K2 según el tipo de líquido y su densidad.
''' Basado en API MPMS 11.1 (Tablas 6A, 6B y 6D).
''' </summary>
''' <param name="TypeLiq">Tipo de líquido (CRD, REF, LUB).</param>
''' <param name="Rho60">Densidad a 60F (ITS-90) en Kg/m³ relativo al agua</param>
''' <returns>Constantes para el cálculo del factor Alfa a 60°F por tipo de producto.</returns>
Public Sub GetKConstants(ByVal TypeLiq As eTypeLiq, ByVal Rho60 As Double, ByRef K0 As Double, ByRef K1 As Double, ByRef K2 As Double)  
  ' Inicializamos en cero por seguridad
  K0 = 0: K1 = 0: K2 = 0

  Select Case TypeLiq
    Case CRD ' Tabla 6A: Petróleo Crudo
      K0 = cK0_CRUDE: K1 = 0: K2 = 0
        
    Case LUB ' Tabla 6D: Aceites Lubricantes
      K0 = 0: K1 = cK1_LUBRICANT: K2 = 0
        
    Case REF ' Tabla 6B: Productos Refinados
      ' Selección por rangos de densidad ITS-90
      If Rho60 >= cRHO_FUEL_OIL Then ' Fuel Oils
          K0 = 103.872: K1 = 0.2701: K2 = 0
      ElseIf Rho60 >= cRHO_JET_FUEL Then ' Jet Fuels
          K0 = 330.301: K1 = 0: K2 = 0
      ElseIf Rho60 >= cRHO_TRANSITION Then ' Transition Zone
          K0 = 1489.067: K1 = 0: K2 = -0.0018684
      ElseIf Rho60 >= cRHO_GASOLINE Then ' Gasolines
          K0 = 192.4571: K1 = 0.2438: K2 = 0
      End If

    Case Else
      ' Tipo de líquido no soportado
  End Select
End Sub

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
  If Not mdHelpers.IsValidAPI(API60, TypeLiq, CheckNormative:=False) Then GoTo ErrHandler

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
  
  Term1 = cFP_A + cFP_B * TempC68
  Term2 = (cFP_C + cFP_D * TempC68) / (Rho68 ^ 2)

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
  If Not mdHelpers.IsValidAPI(API60, TypeLiq, CheckNormative:=False) Or Not mdHelpers.IsValidTemperature(TempF) Then GoTo ErrHandler

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
  DeltaT = TempF68 - cTEMPBASE_F  
  
  ' Calcular el Factor de Corrección por Temperatura (CTL)
  ' Fórmula (API MPMS 11.1 2007 Pág 15): exp[-Alpha_60 * Delta_t * (1 + 0.8 * Alpha_60 * Delta_t)]
  Dim exponente As Double
  
  exponente = -Alfa60 * DeltaT * (1 + cTAYLOR * Alfa60 * DeltaT)

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
  If Pres < Pe Or Pres > cPRESSVALIDRANGE_MAX Or Not mdHelpers.IsValidAPI(API60, TypeLiq, CheckNormative:=False) Then GoTo ErrHandler

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
  Denominador = 1 - (Fcp * DeltaP * cPRESS_SCALING_API) ' 10^-5 es 0.00001
  
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

Public Function API60F(ByVal APIOBS As Double, ByVal TempObs As Double, Optional TypeLiq As eTypeLiq = CRD, Optional ByVal PresObs As Double = 0, Optional ByVal Pe As Double = 0) As Double
  On Error GoTo ErrHandler ' Manejo de errores

  ' Validaciones de Integridad y Finitud (mdHelpers)
  If Not mdHelpers.IsFinite(APIOBS) Or Not mdHelpers.IsFinite(TempObs) Then Exit Function
  If APIOBS <= -cAPI_B Then Exit Function
  
  ' Convertir temperatura observada de ITS-90 F a IPTS-68 F y C UNA VEZ fuera del bucle
  Dim TempObs_F68 As Double
  TempObs_F68 = mdConversion.CONVTEMP68(TempObs, F)
  
  Dim TempObs_C68 As Double
  TempObs_C68 = mdConversion.CONVTEMP(TempObs_F68, F2C)
  
  ' Calcular Delta_t en Fahrenheit IPTS-68
  Dim DeltaT_F68 As Double
  DeltaT_F68 = TempObs_F68 - cTEMPBASE_F

  ' Convertir APIOBS (ITS-90) a Densidad observada en Kg/m³ (RhoObs)
  Dim RhoObs As Double
  RhoObs = mdConversion.CONVDENS(APIOBS, A2K, True)

  Dim Rho60 As Double ' Densidad a 60F (ITS-90) en Kg/m³ (Valor iterado)
  ' Inicializar el valor de Rho60 (ITS-90) para la iteración. Usar RhoObs como punto de partida.
  Rho60 = RhoObs
  
  ' Iniciar el proceso iterativo (Método de Newton)
  Const MAXITERATIONS As Long = 25 ' Número máximo de iteraciones
  Dim m As Byte ' Contador de iteraciones
  Dim API60_Iter As Double  ' API60 (ITS-90) correspondiente a Rho60_Iter
  Dim Alfa60_Iter As Double ' Alfa60 (1/°F) para API60_Iter
  Dim Rho68_Iter As Double  ' Densidad a 60F IPTS-68 (Kg/m³) para API60_Iter (necesaria para FP)
  Dim CTL_Iter As Double    ' CTL para API60_Iter y TempObs
  Dim CPL_Iter As Double    ' CPL para API60_Iter, TempObs, PresObs, Pe
  Dim CTPL_Iter As Double   ' CTPL para API60_Iter, TempObs, PresObs, Pe
  Dim Em As Double          ' Función de error
  Dim Da As Double          ' Coeficiente Da para la derivada
  Dim Dt As Double          ' Derivada correspondiente a la temperatura
  Dim FP_Iter As Double     ' FP para API60_Iter y TempObs
  Dim Dp As Double
  Dim Dp_Num As Double
  Dim Dp_Denom As Double  
  Dim DRho60 As Double      ' Corrección (DRho60) usando la derivada (Método de Newton)

  For m = 1 To MAXITERATIONS
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
    If Abs(Em) < cEPSILON Then
      API60F = API60_Iter
      Exit Function
    End If

    ' Cálculo de la Derivada para el ajuste (Newton Step)
    ' Obtener coeficiente Da (ajuste de expansión diferencial)
    Da = GetDaCoefficient(TypeLiq, Rho60)

    ' Calcular el término Dt (API MPMS 11.1 2007 Apéndice E E.3)
    ' Dt = Da * Alpha_60 * Delta_t * (1 + 1.6 * Alpha_60 * Delta_t)
    Dt = Da * Alfa60_Iter * DeltaT_F68 * (1 + cAPI_F16 * Alfa60_Iter * DeltaT_F68)

    ' Obtener FP_Iter para el API60_Iter actual y TempObs (ITS-90 F). CPL llamó a FP, pero la derivada necesita FP_Iter explícitamente.
    FP_Iter = FP(API60_Iter, TempObs, TypeLiq)

    ' Calcular el término Dp (API MPMS 11.1 2007 Apéndice E E.5)
    ' Dp = -(2 * CPL_m * P_obs * F_cp_m * (7.9392 + 0.02326 * TempObs_C68)) / (rho_m^2 * Alpha_60_m)
    Dp_Num = -(2 * CPL_Iter * (PresObs - Pe) * (FP_Iter * cEPSILON) * (cFPDERIVATIVE_A + cFPDERIVATIVE_B * TempObs_C68))
    Dp_Denom = (Rho60 ^ 2 * Alfa60_Iter)

    If Abs(Dp_Denom) > cEPSILON Then
      Dp = Dp_Num / Dp_Denom
    Else
      Dp = 0
    End If

    ' Calcular el siguiente paso de corrección (DRho60)
    DRho60 = Em / (1 + Dt + Dp)
    
    ' Valor de Rho60 (ITS-90) para la siguiente iteración
    Rho60 = Rho60 + DRho60

    ' Seguridad: Si la densidad se vuelve negativa o irreal, abortar
    If Rho60 <= 0 Then Exit For
  Next m ' Siguiente iteración

  ' Si llega aquí, no hubo convergencia
  Debug.Print "API60F: No se alcanzó convergencia en " & MAXITERATIONS & " iteraciones."
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
    Case CRD: GetDaCoefficient = cDA_CRUDE      ' Crudos (Tabla 6A)
    Case LUB: GetDaCoefficient = cDA_LUBRICANTS ' Lubricantes (Tabla 6D)
    Case REF ' Productos Refinados (Tabla 6B)
      If Rho60 >= cRHO_FUEL_OIL Then
        GetDaCoefficient = cDA_FUEL_OIL       ' Fuel Oils
      ElseIf Rho60 >= cRHO_JET_FUEL Then
        GetDaCoefficient = cDA_JET_FUEL       ' Jet Fuels
      ElseIf Rho60 >= cRHO_TRANSITION Then
        GetDaCoefficient = cDA_TRANSITION     ' Transition
      Else
        GetDaCoefficient = cDA_GASOLINES      ' Gasolines
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
  If Not mdHelpers.IsValidAPI(API60, TypeLiq, CheckNormative:=False) Then GoTo ErrHandler
  
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

  Dim ValRounded As Double ' Valor truncado para el redondeo especial de 0.5

  ' Truncar el valor normalizado para el redondeo especial de 0.5 Regla API:
  ' Si la parte decimal es EXACTAMENTE 0.5, redondear al ENTERO PAR mas cercano.
  ' Si la parte decimal es > 0.5, redondear hacia arriba (IntPart + 0.5) si no fuera por la regla API.
  ' Si la parte decimal es < 0.5, redondear hacia abajo IntPart.
  If Abs(DecPart - 0.5) < cEPSILON_ROUNDING Then ' Si la parte fraccionaria es muy cercana a 0.5
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
  ROUNDAPI = Val ' Retorno de seguridad sin correccion
End Function

' ---------------------------------------------------------------------------------------------------------
' FUNCIONES PARA AJUSTES EN TANQUES (API MPMS Cap. 12.1.1)
' Correcciones geométricas y físicas de la infraestructura de almacenamiento.
' ---------------------------------------------------------------------------------------------------------

''' <summary>
''' Calcula la temperatura de la pared (coraza) del tanque.
''' API MPMS 12.1.1, Sección 5.1.
''' </summary>
''' <param name="TempLiq">Temperatura del producto.</param>
''' <param name="TempAmb">Temperatura ambiente.</param>
''' <returns>Temperatura de la coraza del tanque.</returns>

Public Function TSH(ByVal TmpLiq As Double, ByVal TmpAmb As Double) As Double
  On Error GoTo ErrHandler ' Manejador de errores
  
  ' Validación de Finitud
  If Not mdHelpers.IsFinite(TmpLiq) Or Not mdHelpers.IsFinite(TmpAmb) Then GoTo ErrHandler
  
  ' Fórmula: ((7 * T_liq) + T_amb) / 8
  TSH = (7 * TmpLiq + TmpAmb) / 8

  ' Validación de Finitud
  If Not mdHelpers.IsFinite(TSH) Then GoTo ErrHandler
  Exit Function

ErrHandler:
  Debug.Print "Error en TSH: " & Err.Description
  TSH = TmpLiq ' Retorno de seguridad sin correccion
End Function

''' <summary>
''' Calcula el Factor de Corrección por Temperatura de la Coraza (CTSH).
''' </summary>
''' <param name="TmpLiq">Temperatura del producto.</param>
''' <param name="TmpAmb">Temperatura ambiente.</param>
''' <param name="Mtrl">Tipo de material (eMtrl).</param>
''' <param name="TmpUnits">Unidad térmica (F o C).</param>
''' <param name="TmpBase">Temperatura de referencia (Defecto 60°F o 15°C).</param>
''' <returns>Factor de Corrección por Temperatura de la Coraza.</returns>

Public Function CTSH(ByVal TmpLiq As Double, ByVal TmpAmb As Double, Optional ByVal Mtrl As eMtrl = MCrbn, Optional ByVal TmpUnits As eTmpUnits = F, Optional ByVal TmpBase As Double = 60) As Double
  On Error GoTo ErrHandler ' Manejador de Errores
  
  ' Validaciones de Finitud
  If Not mdHelpers.IsFinite(TmpLiq) Or Not mdHelpers.IsFinite(TmpAmb) Then GoTo ErrHandler
  
  Dim Tcfl As Double
  ' Obtener Coeficiente de Expansión Lineal (Tcfl)
  Tcfl = GetLinearExpansionCoefficient(Mtrl, TmpUnits)
  
  If Tcfl = 0 Then GoTo ErrHandler

  Dim Tshl As Double
  ' 3. Calcular Temperatura de Coraza
  Tshl = TSH(TmpLiq, TmpAmb)

  Dim DeltaTmp As Double
  ' Calcular DeltaTmp    
  DeltaTmp = Tshl - TmpBase
  
  ' 4. Cálculo del CTSH (Expansión de área del cilindro)
  ' La norma utiliza el binomio al cuadrado: (1 + Tcfl * DeltaTmp)^2
  ' lo que expandido es: 1 + 2 * Tcfl * dT + Tcfl^2 * dT^2
  CTSH = (1 + (Tcfl * DeltaTmp)) ^ 2
  
  ' Validación de seguridad
  If Not mdHelpers.IsFinite(CTSH) Then CTSH = 1
  Exit Function

ErrHandler:
  CTSH = 1 ' Factor neutro en caso de error
End Function

''' <summary>
''' Calcula el Ajuste por Techo Flotante (FRA) en volumen.
''' El resultado debe RESTARSE del volumen observado.
''' </summary>
''' <param name="RoofWeight">Peso del techo (lb o kg).</param>
''' <param name="Dens60">Densidad a temperatura base (SGU o Kg/m3).</param>
''' <param name="CTL">Factor de corrección por temperatura del líquido.</param>

Public Function FRA(ByVal RoofWeight As Double, ByVal Dens60 As Double, ByVal CTL As Double) As Double
    On Error GoTo ErrHandler ' Manejador de Errores
    
  ' Validación de Finitud
  If Not mdHelpers.IsFinite(RoofWeight) Or Not mdHelpers.IsFinite(Dens60) Or Not mdHelpers.IsFinite(CTL) Then GoTo ErrHandler
  
  Dim DensObs As Double
  ' Densidad en condiciones observadas = Densidad_Std * CTL
  DensObs = Dens60 * CTL
  
  If DensObs > 0.001 Then
    ' FRA = Peso / Densidad_Observada
    FRA = RoofWeight / (DensObs * 350.507)
  Else
    FRA = 0
  End If

  ' Validación de Finitud
  If Not mdHelpers.IsFinite(FRA) GoTo ErrHandler
  Exit Function

ErrHandler:
  Debug.Print "Error en FRA (Peso: " & RoofWeight & "): " & Err.Description
  FRA = 0
End Function

''' <summary>
''' Retorna el coeficiente de expansión lineal (alpha) según API 12.1.1.
''' </summary>

Private Function GetLinearExpansionCoefficient(ByVal Mtrl As eMtrl, ByVal TmpUnits As eTmpUnits) As Double
  Select Case Mtrl
    Case MCrbn ' Acero al Carbono
      GetLinearExpansionCoefficient = IIf(TmpUnits = F, cCTSH_MCRBN_F, cCTSH_MCRBN_C)
    Case St304 ' Acero Inoxidable 304
      GetLinearExpansionCoefficient = IIf(TmpUnits = F, cCTSH_ST304_F, cCTSH_ST304_C)
    Case St316 ' Acero Inoxidable 316
      GetLinearExpansionCoefficient = IIf(TmpUnits = F, cCTSH_ST316_F, cCTSH_ST316_C)
    Case St4PH ' Acero Inoxidable 17-4 PH
      GetLinearExpansionCoefficient = IIf(TmpUnits = F, cCTSH_ST4PH_F, cCTSH_ST4PH_C)
    Case Else
      GetLinearExpansionCoefficient = 0
  End Select
End Function

''' <summary>
''' Calcula el Factor de Corrección por Carga Hidrostática (CBhp).
''' Compensa la expansión elástica de la coraza del tanque por el peso del líquido.
''' Basado en API MPMS 12.1.1, Sección 5.3.
''' Solo aplica CBhp si la tabla fue generada por métodos geométricos (medición con cinta, láser o triangulación), 
''' que es lo más común en tanques verticales (Upright Cylindrical Tanks).
''' </summary>
''' <param name="Height">Nivel del líquido observado (ft o m).</param>
''' <param name="DensObs">Densidad relativa del líquido a temp. de operación (SGU).</param>
''' <param name="Diameter">Diámetro nominal del tanque (ft o m).</param>
''' <param name="AvgThickness">Espesor promedio de la coraza (in o mm).</param>
''' <param name="IsMetric">True para unidades Métricas (m, mm), False para Imperiales (ft, in).</param>
''' <returns>Factor CBhp (Double). Retorna 1.0 si hay error.</returns>

Public Function CBHP(ByVal Height As Double, ByVal DensObs As Double, ByVal Diameter As Double, ByVal AvgThickness As Double, Optional ByVal IsMetric As Boolean = False) As Double
  On Error GoTo ErrHandler
  
  ' Validaciones de Integridad
  If Not mdHelpers.IsFinite(Height) Or Not mdHelpers.IsFinite(DensObs) Or Not mdHelpers.IsFinite(Diameter) Or Not mdHelpers.IsFinite(AvgThickness) Then GoTo ErrHandler
  
  If AvgThickness <= 0 Or Diameter <= 0 Then GoTo ErrHandler

  ' Definición de Constantes Físicas (Acero al Carbono API 650)
  ' Módulo de Elasticidad (Young's Modulus)
  ' Imperial: 30,000,000 psi | Métrico: 206,842,700,000 Pa (207 GPa)
  Dim E As Double
  Dim GravityConstant As Double
  
  If Not IsMetric Then
    E = cYOUNG_MODULUS_STEEL_IMP ' psi
    GravityConstant = 0.4335 ' psi/ft (presión del agua por pie de altura)
  Else
    E = cYOUNG_MODULUS_STEEL_MET ' Pa
    GravityConstant = 9806.65 ' Pa/m (presión del agua por metro de altura)
  End If

  ' Cálculo de la Presión Hidrostática Promedio (P)
  ' P = Densidad_Relativa * Constante_Gravedad * Altura_Liquido
  ' Nota: La presión se evalúa usualmente a la mitad de la columna de líquido 
  ' para obtener una deformación promedio en la coraza.
  Dim Pressure As Double
  Pressure = DensObs * GravityConstant * (Height / 2)
  
  ' Cálculo de la deformación radial (Hoop Stress principle)
  ' La fórmula simplificada de expansión de volumen para un cilindro delgado:
  ' Delta_V / V = (P * D) / (2 * E * t) 
  ' CBhp = 1 + (Delta_V / V)  
  Dim dValue As Double
  dValue = Diameter
  
  Dim tValue As Double
  tValue = AvgThickness
  
  ' Ajuste de unidades de espesor si es necesario
  ' Si es Imperial, convertimos Diámetro a pulgadas para que coincida con E (psi) y espesor (in)
  If Not IsMetric Then
    dValue = Diameter * 12
  End If
  
  Dim expansionFactor As Double
  expansionFactor = (Pressure * dValue) / (2 * E * tValue)
  
  CBHP = 1 + expansionFactor
  
  ' Validación final
  If Not mdHelpers.IsFinite(CBHP) Or CBHP < 1 Then CBHP = 1
  Exit Function

ErrHandler:
  Debug.Print "Error en CBHP: " & Err.Description
  CBHP = 1
End Function
