Attribute VB_Name = "mdAPICalcs"
Option Explicit

' ------------------------------------------------------------------------------
' MÓDULO: mdAPICalcs
' DESCRIPCIÓN:
'   Este módulo contiene funciones para calcular cantidades de hidrocarburos 
'   líquidos, basado en las normas API MPMS. Implementa los estándares 
'   internacionales para la transferencia de custodia, permitiendo la 
'   normalización de volúmenes observados a condiciones estándar 
'   (60°F / 14.696 psi).
'
' DEPENDENCIAS:
'   - mdGlobals: Para definiciones de enumeraciones (Enums) y constantes 
'                API MPMS.
'   - mdConversion: Para funciones de conversión de unidades (CONVTEMP, 
'                   CONVDENS, CONVTEMP68).
'   - mdHelpers: Validaciones de integridad matemática y finitud (Failsafe).
' ------------------------------------------------------------------------------

' ------------------------------------------------------------------------------
' FUNCIONES DE CORRECCIÓN INSTRUMENTAL (API MPMS Cap. 9.3)
' Ajustes por la expansión térmica del vidrio en hidrómetros.
' ------------------------------------------------------------------------------

Public Function HYC(ByVal TempObs As Double, _
                    Optional TempUnits As eTmpUnits = FHR) As Double
' Calcula el Factor de Corrección por Temperatura del Hidrómetro (HYC).
' Basado en API MPMS Capítulo 9.3 (Corrección por expansión del vidrio).
  
  On Error GoTo ErrHandler
  If Not mdHelpers.IsFinite(TempObs) Then GoTo ErrHandler

  Dim tempObsF As Double ' Temperatura observada convertida a Fahrenheit
  Dim tempObsC As Double ' Temperatura observada convertida a Celsius
  Dim deltaT As Double   ' Temperatura usada en la fórmula (F o C)

  Select Case TempUnits
    Case FHR, RNK:
      If TempUnits = FHR Then
        tempObsF = TempObs
      Else
        tempObsF = mdConversion.CONVTEMP(TempObs, TempUnits, FHR)
      End If
      deltaT = tempObsF - cTEMPBASE_F
      HYC = 1 - (cHYCF_A * deltaT) - (cHYCF_B * (deltaT ^ 2))
    Case CLS, KLV:
      If TempUnits = CLS Then
        tempObsC = TempObs
      Else
        tempObsC = mdConversion.CONVTEMP(TempObs, TempUnits, CLS)
      End If      
      deltaT = tempObsC - cTEMPBASE_C
      HYC = 1 - (cHYCC_C * deltaT) - (cHYCC_D * (deltaT ^ 2))
    Case Else:
      HYC = 1 ' Unidad no reconocida
  End Select

  If Not mdHelpers.IsFinite(HYC) Then HYC = 1
  Exit Function

ErrHandler:
  Debug.Print "Error en HYC: " & Err.Description  
  HYC = 1 ' Retorno de seguridad sin correccion
End Function

' ------------------------------------------------------------------------------

Public Function DENSHYC(ByVal DensObs As Double, _
                        ByVal TempObs As Double, _
                        Optional DensUnits As eDnsUnits = API, _
                        Optional TempUnits As eTmpUnits = FHR) As Double
' Calcula la densidad o API corregido por hidrómetro.
' Convierte la lectura física a densidad absoluta para aplicar el factor HYC
  
  On Error GoTo ErrHandler
  If Not mdHelpers.IsFinite(DensObs) Or _
     Not mdHelpers.IsFinite(TempObs) Then GoTo ErrHandler

  Dim corrHyd As Double       ' Factor de corrección por hidrómetro
  Dim densObsInKGM As Double  ' Densidad observada convertida a Kg/m³
  Dim densCorrInKGM As Double ' Densidad corregida en Kg/m³

  Select Case DensUnits
    Case API
        If DensObs <= -cAPI_B Then GoTo ErrHandler
        densObsInKGM = mdConversion.CONVDENS(DensObs, API, KGM, True)
    Case KGM
        If DensObs <= 0 Then GoTo ErrHandler
        densObsInKGM = DensObs
    Case SGU
        If DensObs <= 0 Then GoTo ErrHandler
        densObsInKGM = mdConversion.CONVDENS(DensObs, SGU, KGM, True)
  End Select

  ' Obtener el factor de corrección por hidrómetro (HYC)
  corrHyd = HYC(TempObs, TempUnits)  
  ' Aplicar la corrección multiplicando la densidad en Kg/m³ por el factor HYC
  densCorrInKGM = densObsInKGM * corrHyd

  Select Case DensUnits
    Case API
      DENSHYC = mdConversion.CONVDENS(densCorrInKGM, KGM, API, True)
    Case KGM
      DENSHYC = densCorrInKGM
    Case SGU
      DENSHYC = mdConversion.CONVDENS(densCorrInKGM, KGM, SGU, True)
  End Select
  Exit Function

ErrHandler:
  Debug.Print "Error en DENSHYC: " & Err.Description
  DENSHYC = 0 ' Retorno de seguridad sin correccion
End Function

' ------------------------------------------------------------------------------
' FUNCIONES ASOCIADAS A LA DINÁMICA DEL FLUIDO (API MPMS Cap. 11.1)
' Cálculo de factores de corrección volumétrica basados en la ecuación de estado
' API.
' ------------------------------------------------------------------------------

Public Function ALPHA60MED(ByVal DensRange As Variant, _
                           ByVal TempRange As Variant, _
                           Optional ByVal TypeLiq As eTypeLiq = CRD) As Variant
' Calcula el Coeficiente de Expansión Térmica Promedio (Alfa 60 Medio) a partir 
' de una serie de mediciones de densidad y temperatura. Basado en API MPMS 11.1 
' (2007) - Aplicaciones Especiales.
  
  On Error GoTo ErrHandler
  
  ' Convertir entradas a arreglos de Double (Handles Ranges and Arrays)
  Dim arrDens() As Double
  arrDens = mdHelpers.ConvertToDoubleArray(DensRange)
  
  Dim arrTmps() As Double
  arrTmps = mdHelpers.ConvertToDoubleArray(TempRange)
  
  ' Validar que ambos arreglos tengan la misma dimensión y mínimo 10 datos
  Dim sampleCount As Long
  sampleCount = UBound(arrDens)

  If sampleCount <> UBound(arrTmps) Or sampleCount < 10 Then
    ' Retorna #NUM! en Excel si los datos son insuficientes
    ALPHA60MED = CVErr(xlErrNum)
    Exit Function
  End If

  ' Procesamiento Matricial
  Dim arrayIndex As Long
  Dim sumAlpha As Double
  Dim api60_Iter As Double
  Dim alfa60_Iter As Double
  
  Dim validCounts As Long
  validCounts = 0
  
  For arrayIndex = 1 To sampleCount
    If mdHelpers.IsFinite(arrDens(arrayIndex)) And mdHelpers.IsFinite(arrTmps(arrayIndex)) Then      
      ' Hallar API 60 base para este punto (Newton-Raphson)
      api60_Iter = API60F(arrDens(arrayIndex), arrTmps(arrayIndex), TypeLiq, 0, 0)
      
      If api60_Iter > 0 Then
        ' Hallar Alfa 60 para este API 60
        alfa60_Iter = ALPHA60(api60_Iter, TypeLiq)
        
        If alfa60_Iter > 0 Then
          sumAlpha = sumAlpha + alfa60_Iter
          validCounts = validCounts + 1
        End If
      End If
    End If
  Next arrayIndex
  
  ' Calcular Promedio
  If validCounts >= 10 Then
    ALPHA60MED = sumAlpha / validCounts
  Else
    ALPHA60MED = 0
  End If
  Exit Function

ErrHandler:
  Debug.Print "Error en ALPHA60MED: " & Err.Description
  ALPHA60MED = 0
End Function

' ------------------------------------------------------------------------------

Public Function ALPHA60(ByVal API60 As Double, _
                        Optional TypeLiq As eTypeLiq = CRD, _
                        Optional ByRef Rho68 As Double) As Double
' Calcula el Coeficiente de Expansión Térmica a 60°F (Alfa60) y la Densidad en
' escala IPTS-68. Estándar: API MPMS 11.1.
  
  On Error GoTo ErrHandler
  If Not mdHelpers.IsFinite(API60) Then GoTo ErrHandler

  If Not mdHelpers.IsValidAPI(API60, TypeLiq, CheckNormative:=False) Then _
     GoTo ErrHandler

  Dim rho60 As Double
  rho60 = mdConversion.CONVDENS(API60, API, KGM, True)
  If rho60 <= 0 Then GoTo ErrHandler
  
  ' Selección de Constantes K (API MPMS 11.1 Tablas 6A, 6B, 6D)
  Dim k0 As Double, k1 As Double, k2 As Double  
  Call GetKConstants(TypeLiq, rho60, k0, k1, k2)
  ' Verificación de seguridad: si no se encontró un rango válido, abortar
  If k0 = 0 And k1 = 0 And k2 = 0 And TypeLiq <> LUB Then GoTo ErrHandler

  ' Calcular Alfa60 (en 1/°F) usando las constantes K y rho60.
  ' Alfa60 = k0/rho60^2 + k1/rho60 + k2
  ALPHA60 = (k0 / (rho60 ^ 2)) + (k1 / rho60) + k2  
  ' Calcular Rho68 (densidad a 60F IPTS-68 en Kg/m³) a partir de rho60 (ITS-90)
  Rho68 = rho60 * Exp(ALPHA60 * cTEMPSHIFT)

  If Not mdHelpers.IsFinite(ALPHA60) Then ALPHA60 = 0
  Exit Function

ErrHandler:
  Debug.Print "Error en ALPHA60: " & Err.Description
  ALPHA60 = 0 ' Retorno de seguridad sin correccion
  Rho68 = 0
End Function

' ------------------------------------------------------------------------------

Public Sub GetKConstants(ByVal TypeLiq As eTypeLiq, _
                         ByVal rho60 As Double, _
                         ByRef k0 As Double, _
                         ByRef k1 As Double, _
                         ByRef k2 As Double)
' Provee las constantes k0, k1 y k2 según el tipo de líquido y su densidad.
' Basado en API MPMS 11.1 (Tablas 6A, 6B y 6D).

  k0 = 0: k1 = 0: k2 = 0

  Select Case TypeLiq
    Case CRD ' Tabla 6A: Petróleo Crudo
      k0 = ck0_CRUDE: k1 = 0: k2 = 0
        
    Case LUB ' Tabla 6D: Aceites Lubricantes
      k0 = 0: k1 = ck1_LUBRICANT: k2 = 0
        
    Case REF ' Tabla 6B: Productos Refinados
      ' Selección por rangos de densidad ITS-90
      If rho60 >= cRHO_FUEL_OIL Then ' Fuel Oils
          k0 = 103.872: k1 = 0.2701: k2 = 0
      ElseIf rho60 >= cRHO_JET_FUEL Then ' Jet Fuels
          k0 = 330.301: k1 = 0: k2 = 0
      ElseIf rho60 >= cRHO_TRANSITION Then ' Transition Zone
          k0 = 1489.067: k1 = 0: k2 = -0.0018684
      ElseIf rho60 >= cRHO_GASOLINE Then ' Gasolines
          k0 = 192.4571: k1 = 0.2438: k2 = 0
      End If

    Case Else
      ' Tipo de líquido no soportado
  End Select
End Sub

' ------------------------------------------------------------------------------

Public Function FP(ByVal API60 As Double, _
                   ByVal TempF As Double, _
                   Optional TypeLiq As eTypeLiq = CRD) As Double
' Calcula el Coeficiente de Compresibilidad Escalado (Fp).
' Basado en API MPMS 11.1 (2007).
  
  On Error GoTo ErrHandler
  If Not mdHelpers.IsFinite(API60) Or _
     Not mdHelpers.IsFinite(TempF) Then GoTo ErrHandler

  If Not mdHelpers.IsValidAPI(API60, TypeLiq, CheckNormative:=False) Then _
     GoTo ErrHandler

  ' Temperatura observada convertida a IPTS-68 Celsius
  Dim TempC68 As Double
  TempC68 = mdConversion.CONVTEMP68(TempF, C)

  ' Densidad a 60F ajustada a IPTS-68 en Kg/m³
  Dim Rho68 As Double
  Call ALPHA60(API60, TypeLiq, Rho68)
  If rho68 <= 0 Then GoTo ErrHandler

  ' Calcular el Coeficiente de Compresibilidad Escalado (FP)
  ' Fórmula: exp[A + B*t_c68 + (C + D*t_c68)/rho68^2]
  Dim Term1 As Double, Term2 As Double, exponente As Double  
  Term1 = cFP_A + cFP_B * TempC68
  Term2 = (cFP_C + cFP_D * TempC68) / (Rho68 ^ 2)
  exponente = Term1 + Term2
  ' Prevención de desbordamiento (Overflow)
  If exponente > 700 Then 
    FP = 0
    Exit Function
  End If
  ' Resultado Final
  FP = Exp(exponente)

  If Not mdHelpers.IsFinite(FP) Then FP = 0
  Exit Function

ErrHandler:
  Debug.Print "Error en FP: " & Err.Description
  FP = 0 ' Retorno de seguridad sin correccion
End Function

' ------------------------------------------------------------------------------

Public Function CTL(ByVal API60 As Double, _
                    ByVal TempF As Double, _
                    Optional TypeLiq As eTypeLiq = CRD) As Double
' Calcula el Factor de Corrección por Temperatura del Líquido (CTL / VCF).
' Basado en API MPMS Capítulo 11.1 (2007/2019).
  
  On Error GoTo ErrHandler

  If Not mdHelpers.IsFinite(API60) Or _
     Not mdHelpers.IsFinite(TempF) Then GoTo ErrHandler

  If Not mdHelpers.IsValidAPI(API60, TypeLiq, CheckNormative:=False) Or _
     Not mdHelpers.IsValidTemperature(TempF) Then GoTo ErrHandler

  ' Coeficiente de expansión térmica a 60F (en 1/°F)
  Dim Alfa60 As Double
  Alfa60 = ALPHA60(API60, TypeLiq)
  If Alfa60 <= 0 Then GoTo ErrHandler

  ' Convertir temperatura observada de ITS-90 F a IPTS-68 F
  Dim TempF68 As Double
  TempF68 = mdConversion.CONVTEMP68(TempF, F)

  ' Calcular la diferencia de temperatura con la temperatura base 60F
  Dim DeltaT As Double
  DeltaT = TempF68 - cTEMPBASE_F
  
  ' Calcular el Factor de Corrección por Temperatura (CTL)
  ' Fórmula: exp[-Alpha_60 * Delta_t * (1 + 0.8 * Alpha_60 * Delta_t)]
  Dim exponente As Double  
  exponente = -Alfa60 * DeltaT * (1 + cTAYLOR * Alfa60 * DeltaT)
  ' Prevención de desbordamiento (Overflow)
  If exponente > 700 Then 
    CTL = 1
    Exit Function
  ElseIf exponente < -700 Then
    CTL = 0
    Exit Function
  End If
  ' Resultado Final
  CTL = Exp(exponente)

  If Not mdHelpers.IsFinite(CTL) Or CTL <= 0 Then CTL = 1  
  Exit Function

ErrHandler:
  Debug.Print "Error crítico en CTL (API60: " & API60 & _
              ", Temp: " & TempF & "): " & Err.Description
  CTL = 1 ' Retorno de seguridad sin correccion
End Function

' ------------------------------------------------------------------------------

Public Function CPL(ByVal API60 As Double, _
                    ByVal TempF As Double, _
                    ByVal Pres As Double, _
                    Optional TypeLiq As eTypeLiq = CRD, _
                    Optional ByVal equilibriumPressure As Double = 0) As Double
' Calcula el Factor de Corrección por Presión del Líquido (CPL).
' Basado en API MPMS Capítulo 11.1.
  
  On Error GoTo ErrHandler

  If Not mdHelpers.IsFinite(API60) Or Not mdHelpers.IsFinite(TempF) Or _
     Not mdHelpers.IsFinite(Pres) Or Not mdHelpers.IsFinite(equilibriumPressure) Then _
     GoTo ErrHandler
  
  ' Validaciones de Rangos Físicos API 11.1
  If Pres < equilibriumPressure Or Pres > cPRESSVALIDRANGE_MAX Or _
     Not mdHelpers.IsValidAPI(API60, TypeLiq, CheckNormative:=False) Then _
     GoTo ErrHandler

  ' Hallar el Factor de Compresibilidad Escalado (FP)
  Dim Fcp As Double
  Fcp = FP(API60, TempF, TypeLiq)
  If Fcp <= 0 Then GoTo ErrHandler

  ' Calcular la diferencia de presión (Pres - Pe) en PSI
  Dim DeltaP As Double
  DeltaP = Pres - equilibriumPressure

  ' Calcular el Factor de Corrección por Presión (CPL)
  ' Fórmula: 1 / (1 - Fcp * (Pres - Pe) * 10 ^ -5)
  Dim Denominador As Double
  Denominador = 1 - (Fcp * DeltaP * cPRESS_SCALING_API)  
  ' Protección contra indeterminación
  If Abs(Denominador) < 0.1 Then
    CPL = 1
    Exit Function
  End If
  ' Resultado Final
  CPL = 1 / Denominador

  If Not mdHelpers.IsFinite(CPL) Or CPL < 1 Then CPL = 1
  Exit Function

ErrHandler:
  Debug.Print "Error en CPL (API60: " & API60 & _
              ", P: " & Pres & "): " & Err.Description
  CPL = 1 ' Retorno de seguridad sin correccion
End Function

' ------------------------------------------------------------------------------

Public Function API60F(ByVal APIOBS As Double, _
                       ByVal TempObs As Double, _
                       Optional TypeLiq As eTypeLiq = CRD, _
                       Optional ByVal PresObs As Double = 0, _
                       Optional ByVal equilibriumPressure As Double = 0) As Double
' Calcula la Gravedad API a 60°F y Presión de Equilibrio mediante el método de
' Newton-Raphson. Basado en API MPMS Capítulo 11.1 (Apéndice E).
  
  On Error GoTo ErrHandler

  If Not mdHelpers.IsFinite(APIOBS) Or _
     Not mdHelpers.IsFinite(TempObs) Then Exit Function
  
  If APIOBS <= -cAPI_B Then Exit Function
  
  ' Convertir temperatura observada de ITS-90 F a IPTS-68 F y C
  Dim TempObs_F68 As Double
  TempObs_F68 = mdConversion.CONVTEMP68(TempObs, F)
  
  Dim TempObs_C68 As Double
  TempObs_C68 = mdConversion.CONVTEMP(TempObs_F68, F, C)
  
  ' Calcular Delta_t en Fahrenheit IPTS-68
  Dim DeltaT_F68 As Double
  DeltaT_F68 = TempObs_F68 - cTEMPBASE_F

  ' Convertir APIOBS (ITS-90) a Densidad observada en Kg/m³ (RhoObs)
  Dim RhoObs As Double
  RhoObs = mdConversion.CONVDENS(APIOBS, API, KGM, True)

  ' Inicializar el valor de rho60 (ITS-90) para la iteración.
  ' Usar RhoObs como punto de partida.
  Dim rho60 As Double
  rho60 = RhoObs
  
  ' Iniciar el proceso iterativo (Método de Newton)
  Const MAXITERATIONS As Long = 25 ' Número máximo de iteraciones
  Dim iterationIndex As Byte        ' Contador de iteraciones
  Dim api60_Iter As Double          ' API60 (ITS-90) correspondiente a rho60_Iter
  Dim alfa60_Iter As Double         ' Alfa60 (1/°F) para api60_Iter
  Dim Rho68_Iter As Double          ' Densidad a 60F IPTS-68 (Kg/m³) para api60_Iter
  Dim CTL_Iter As Double            ' CTL para api60_Iter y TempObs
  Dim CPL_Iter As Double            ' CPL para api60_Iter, TempObs, PresObs, Pe
  Dim CTPL_Iter As Double           ' CTPL para api60_Iter, TempObs, PresObs, Pe
  Dim densityError As Double        ' Función de error
  Dim derivativeCoefficient As Double ' Coeficiente Da para la derivada
  Dim temperatureDerivativeTerm As Double ' Derivada correspondiente a la temperatura
  Dim FP_Iter As Double             ' FP para api60_Iter y TempObs
  Dim pressureDerivativeTerm As Double
  Dim pressureDerivativeNumerator As Double
  Dim pressureDerivativeDenominator As Double
  Dim Drho60 As Double              ' Corrección (Drho60) usando la derivada

  For iterationIndex = 1 To MAXITERATIONS
    ' Convertir el rho60_Iter (supuesto, ITS-90) a api60_Iter (ITS-90)
    api60_Iter = mdConversion.CONVDENS(rho60, KGM, API, True)
    ' Obtener Alfa60 (1/°F) y Rho68 (Kg/m³) para el api60_Iter actual
    alfa60_Iter = ALPHA60(api60_Iter, TypeLiq, Rho68_Iter)
    ' Calcular CTL para el api60_Iter actual
    CTL_Iter = CTL(api60_Iter, TempObs, TypeLiq)
    ' Calcular CPL para el api60_Iter actual
    CPL_Iter = CPL(api60_Iter, TempObs, PresObs, TypeLiq, equilibriumPressure)
    ' Calcular CTPL_Iter como el producto de CTL_Iter y CPL_Iter
    CTPL_Iter = CTL_Iter * CPL_Iter
    ' Calcular densityError (Error)
    densityError = (RhoObs / CTPL_Iter) - rho60
    ' Criterio de parada
    If Abs(densityError) < cEPSILON Then
      API60F = api60_Iter
      Exit Function
    End If

    ' Cálculo de la Derivada para el ajuste (Newton Step)
    ' Obtener coeficiente derivativeCoefficient (ajuste de expansión diferencial)
    derivativeCoefficient = GetDaCoefficient(TypeLiq, rho60)
    ' Calcular el término temperatureDerivativeTerm (API MPMS 11.1 2007 Apéndice E E.3)
    ' temperatureDerivativeTerm = derivativeCoefficient * Alpha_60 * Delta_t * (1 + 1.6 * Alpha_60 * Delta_t)
    temperatureDerivativeTerm = derivativeCoefficient * alfa60_Iter * DeltaT_F68 * _
         (1 + cAPI_F16 * alfa60_Iter * DeltaT_F68)
    ' Obtener FP_Iter para el api60_Iter actual y TempObs (ITS-90 F).
    ' CPL llamó a FP, pero la derivada necesita FP_Iter explícitamente.
    FP_Iter = FP(api60_Iter, TempObs, TypeLiq)
    ' Calcular el término pressureDerivativeTerm (API MPMS 11.1 2007 Apéndice E E.5)
    ' pressureDerivativeTerm = -(2 * CPL_m * P_obs * F_cp_m * (7.9392 + 0.02326 * TempObs_C68)) / (rho_m^2 * Alpha_60_m)
    pressureDerivativeNumerator = -(2 * CPL_Iter * (PresObs - equilibriumPressure) * (FP_Iter * cEPSILON) * _
              (cFPDERIVATIVE_A + cFPDERIVATIVE_B * TempObs_C68))
    pressureDerivativeDenominator = (rho60 ^ 2 * alfa60_Iter)

    If Abs(pressureDerivativeDenominator) > cEPSILON Then
      pressureDerivativeTerm = pressureDerivativeNumerator / pressureDerivativeDenominator
    Else
      pressureDerivativeTerm = 0
    End If

    ' Calcular el siguiente paso de corrección (Drho60)
    Drho60 = densityError / (1 + temperatureDerivativeTerm + pressureDerivativeTerm)    
    ' Valor de rho60 (ITS-90) para la siguiente iteración
    rho60 = rho60 + Drho60
    ' Seguridad: Si la densidad se vuelve negativa o irreal, abortar
    If rho60 <= 0 Then Exit For
  Next iterationIndex ' Siguiente iteración

  ' Si llega aquí, no hubo convergencia
  Debug.Print "API60F: No se alcanzó convergencia en " & _
              MAXITERATIONS & " iteraciones."
  API60F = 0
  Exit Function

ErrHandler:
  Debug.Print "Error en API60F: " & Err.Description
  API60F = 0 ' Retorno de seguridad sin correccion
End Function

' ------------------------------------------------------------------------------

Private Function GetDaCoefficient(ByVal TypeLiq As eTypeLiq, _
                                  ByVal rho60 As Double) As Double
' Helper privado para obtener el coeficiente Da según el tipo de líquido.
' Basado en API MPMS 11.1 Apéndice E.

  Select Case TypeLiq
    Case CRD: GetDaCoefficient = cDA_CRUDE      ' Crudos (Tabla 6A)
    Case LUB: GetDaCoefficient = cDA_LUBRICANTS ' Lubricantes (Tabla 6D)
    Case REF                                    ' Productos Refinados (Tabla 6B)
      If rho60 >= cRHO_FUEL_OIL Then
        GetDaCoefficient = cDA_FUEL_OIL         ' Fuel Oils
      ElseIf rho60 >= cRHO_JET_FUEL Then
        GetDaCoefficient = cDA_JET_FUEL         ' Jet Fuels
      ElseIf rho60 >= cRHO_TRANSITION Then
        GetDaCoefficient = cDA_TRANSITION       ' Transition
      Else
        GetDaCoefficient = cDA_GASOLINES        ' Gasolines
      End If
    Case Else
      GetDaCoefficient = 0
  End Select
End Function

' ------------------------------------------------------------------------------

Public Function APIOBS(ByVal API60 As Double, _
                       ByVal TempF As Double, _
                       Optional TypeLiq As eTypeLiq = CRD) As Double
' Calcula la Gravedad API observada a una temperatura dada partiendo del API
' a 60°F. Asume condiciones de presión base (atmosférica).

  On Error GoTo ErrHandler

  If Not mdHelpers.IsFinite(API60) Or _
     Not mdHelpers.IsFinite(TempF) Then GoTo ErrHandler

  If Not mdHelpers.IsValidAPI(API60, TypeLiq, CheckNormative:=False) Then _
     GoTo ErrHandler
  
  ' Convertir API60 (BASE) a Densidad (ITS-90) en Kg/m³ a 60F
  Dim rho60 As Double
  rho60 = mdConversion.CONVDENS(API60, API, KGM, True)
  
  ' Calcular el Factor de Corrección por Temperatura del Líquido (CTL)
  Dim Ftl As Double
  Ftl = CTL(API60, TempF, TypeLiq)
  
  ' Calcular la Densidad en Kg/m³ a la Temperatura Observada (a presión base)
  ' Fórmula: Densidad@Obs = Densidad@Base * CTL
  Dim RhoObs As Double
  RhoObs = rho60 * Ftl
  If RhoObs <= 0 Then GoTo ErrHandler
  
  ' Convertir la Densidad a la Temperatura Observada (en Kg/m³) a Gravedad API
  APIOBS = mdConversion.CONVDENS(RhoObs, KGM, API, True)

  If Not mdHelpers.IsFinite(APIOBS) Then APIOBS = 0
  Exit Function

ErrHandler:
  Debug.Print "Error en APIOBS: " & Err.Description
  APIOBS = 0 ' Retorno de seguridad sin correccion
End Function

' ------------------------------------------------------------------------------

Public Function ROUNDAPI(ByVal Val As Double, _
                         ByVal PosDec As Long, _
                         Optional ByVal EndNum As Byte = 1) As Double
' Redondea un valor según las reglas API MPMS 11.1 (Bankers's Rounding).
' Permite incrementos personalizados (ej. redondear al 0.05 más cercano).
  
  On Error GoTo ErrHandler

  If Not mdHelpers.IsFinite(Val) Or PosDec < 0 Or EndNum <= 0 Then 
    ROUNDAPI = Val
    Exit Function
  End If

  ' Determinar incremento de redondeo
  Dim Inc As Double
  Inc = (10 ^ -PosDec) * EndNum
  
  ' Normalizar el valor de acuerdo al incremento y obtener su parte decimal
  Dim NormVal As Double
  NormVal = Abs(Val) / Inc

  ' Obtener la parte entera del valor normalizado
  Dim IntPart As Long
  IntPart = Fix(NormVal)
  
  ' Obtener la parte decimal del valor normalizado
  Dim DecPart As Double
  DecPart = NormVal - IntPart

  Dim ValRounded As Double ' Valor truncado para el redondeo especial de 0.5

  ' Truncar el valor normalizado para el redondeo especial de 0.5 Regla API:
  ' Si la parte decimal es EXACTAMENTE 0.5, redondear al ENTERO PAR mas cercano.
  ' Si la parte decimal es > 0.5, redondear hacia arriba (IntPart + 0.5).
  ' Si la parte decimal es < 0.5, redondear hacia abajo IntPart.
  If Abs(DecPart - 0.5) < cEPSILON_ROUNDING Then
    ' Verificar si el ENTERO de NormVal es par o impar.
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

  If Not mdHelpers.IsFinite(ROUNDAPI) Then ROUNDAPI = Val
  Exit Function

ErrHandler:
  Debug.Print "Error en ROUNDAPI: " & Err.Description
  ROUNDAPI = Val ' Retorno de seguridad sin correccion
End Function

' ------------------------------------------------------------------------------

Public Function SHRINK(ByVal lightBlendPercent As Double, _
                       ByVal APILight As Double, _
                       ByVal APIHeavy As Double) As Double
' Calcula el porcentaje de encogimiento (shrinkage) en mezclas de hidrocarburos
' livianos con crudos pesados.
' Basado en correlaciones empíricas de la industria petrolera (Ecopetrol ICPET).
  
  On Error GoTo ErrHandler

  If Not mdHelpers.IsFinite(lightBlendPercent) Or Not mdHelpers.IsFinite(APILight) Or _
     Not mdHelpers.IsFinite(APIHeavy) Then GoTo ErrHandler

  If lightBlendPercent < 0 Or lightBlendPercent > 100 Then GoTo ErrHandler
  If APILight <= 0 Or APIHeavy <= 0 Then GoTo ErrHandler

  ' Variables para coeficientes
  Dim shrinkCoefficient As Double, shrinkExponentY As Double, shrinkExponentZ As Double

  ' Selección de coeficientes según API del componente pesado
  If APIHeavy <= 10.8 Then
    shrinkCoefficient = 1.532E-08
    shrinkExponentY = 1.6769
    shrinkExponentZ = 1.7841
  ElseIf APIHeavy <= 12# Then
    shrinkCoefficient = 2.8822E-07
    shrinkExponentY = 1.222
    shrinkExponentZ = 1.4731
  Else ' APIHeavy > 12.0
    shrinkCoefficient = 4.2178E-07
    shrinkExponentY = 1.1812
    shrinkExponentZ = 1.4151
  End If

  ' Cálculo del encogimiento
  ' Formula: encogimiento = shrinkCoefficient * lightBlendPercent * (100 - lightBlendPercent)^shrinkExponentY * (API_liviano - API_pesado)^shrinkExponentZ
  Dim term1 As Double, term2 As Double, term3 As Double
  
  term1 = shrinkCoefficient * lightBlendPercent
  term2 = (100 - lightBlendPercent) ^ shrinkExponentY
  term3 = (APILight - APIHeavy) ^ shrinkExponentZ

  SHRINK = term1 * term2 * term3

  If Not mdHelpers.IsFinite(SHRINK) Then SHRINK = 0
  Exit Function

ErrHandler:
  Debug.Print "Error en SHRINK (Valor de mezcla: " & lightBlendPercent & ", APILight: " & APILight & _
              ", APIHeavy: " & APIHeavy & "): " & Err.Description
  SHRINK = 0 ' Retorno de seguridad sin corrección
End Function

' ------------------------------------------------------------------------------
' FUNCIONES PARA AJUSTES EN TANQUES (API MPMS Cap. 12.1.1)
' Correcciones geométricas y físicas de la infraestructura de almacenamiento.
' ------------------------------------------------------------------------------

Public Function TSH(ByVal TmpLiq As Double, _
                    ByVal TmpAmb As Double) As Double
' Calcula la temperatura de la pared (coraza) del tanque.
' API MPMS 12.1.1, Sección 5.1.
  
  On Error GoTo ErrHandler

  If Not mdHelpers.IsFinite(TmpLiq) Or _
     Not mdHelpers.IsFinite(TmpAmb) Then GoTo ErrHandler
  
  ' Fórmula: ((7 * T_liq) + T_amb) / 8
  TSH = (7 * TmpLiq + TmpAmb) / 8

  If Not mdHelpers.IsFinite(TSH) Then GoTo ErrHandler
  Exit Function

ErrHandler:
  Debug.Print "Error en TSH: " & Err.Description
  TSH = TmpLiq ' Retorno de seguridad sin correccion
End Function

' ------------------------------------------------------------------------------

Public Function CTSH(ByVal TmpLiq As Double, _
                     ByVal TmpAmb As Double, _
                     Optional ByVal Mtrl As eMtrl = MCrbn, _
                     Optional ByVal TempUnits As eTmpUnits = F, _
                     Optional ByVal TmpBase As Double = cTEMPBASE_F) As Double
' Calcula el Factor de Corrección por Temperatura de la Coraza (CTSH).
  
  On Error GoTo ErrHandler

  If Not mdHelpers.IsFinite(TmpLiq) Or _
     Not mdHelpers.IsFinite(TmpAmb) Then GoTo ErrHandler
  
  ' Obtener Coeficiente de Expansión Lineal (Tcfl)
  Dim Tcfl As Double
  Tcfl = GetLinearExpansionCoefficient(Mtrl, TempUnits)  
  If Tcfl = 0 Then GoTo ErrHandler

  ' Calcular Temperatura de Coraza
  Dim Tshl As Double
  Tshl = TSH(TmpLiq, TmpAmb)

  ' Calcular DeltaTmp
  Dim DeltaTmp As Double
  DeltaTmp = Tshl - TmpBase
  
  ' Cálculo del CTSH (Expansión de área del cilindro)
  ' La norma utiliza el binomio al cuadrado: (1 + Tcfl * DeltaTmp)^2
  ' lo que expandido es: 1 + 2 * Tcfl * dT + Tcfl^2 * dT^2
  CTSH = (1 + (Tcfl * DeltaTmp)) ^ 2

  If Not mdHelpers.IsFinite(CTSH) Then CTSH = 1
  Exit Function

ErrHandler:
  Debug.Print "Error en CTSH: " & Err.Description
  CTSH = 1 ' Factor neutro en caso de error
End Function

' ------------------------------------------------------------------------------

Public Function FRA(ByVal RoofWeight As Double, _
                    ByVal Dens60 As Double, _
                    ByVal CTL As Double) As Double
' Calcula el Ajuste por Techo Flotante (FRA) en volumen.
' El resultado debe RESTARSE del volumen observado.
  
  On Error GoTo ErrHandler

  If Not mdHelpers.IsFinite(RoofWeight) Or Not mdHelpers.IsFinite(Dens60) Or _
     Not mdHelpers.IsFinite(CTL) Then GoTo ErrHandler
  
  ' Densidad en condiciones observadas = Densidad_Std * CTL
  Dim DensObs As Double
  DensObs = Dens60 * CTL
  
  ' FRA = Peso / Densidad_Observada
  If DensObs > 0.001 Then
    FRA = RoofWeight / DensObs
  Else
    FRA = 0
  End If

  If Not mdHelpers.IsFinite(FRA) Then GoTo ErrHandler
  Exit Function

ErrHandler:
  Debug.Print "Error en FRA (Peso: " & RoofWeight & "): " & Err.Description
  FRA = 0
End Function

' ------------------------------------------------------------------------------

Private Function GetLinearExpansionCoefficient( _
                                          ByVal Mtrl As eMtrl, _
                                          ByVal TempUnits As eTmpUnits) As Double
' Retorna el coeficiente de expansión lineal (alpha) según API 12.1.1.

  Select Case Mtrl
    Case MCrbn ' Acero al Carbono
      GetLinearExpansionCoefficient = _
                                IIf(TempUnits = F, cCTSH_MCRBN_F, cCTSH_MCRBN_C)
    Case St304 ' Acero Inoxidable 304
      GetLinearExpansionCoefficient = _
                                IIf(TempUnits = F, cCTSH_ST304_F, cCTSH_ST304_C)
    Case St316 ' Acero Inoxidable 316
      GetLinearExpansionCoefficient = _
                                IIf(TempUnits = F, cCTSH_ST316_F, cCTSH_ST316_C)
    Case St4PH ' Acero Inoxidable 17-4 PH
      GetLinearExpansionCoefficient = _
                                IIf(TempUnits = F, cCTSH_ST4PH_F, cCTSH_ST4PH_C)
    Case Else
      GetLinearExpansionCoefficient = 0
  End Select
End Function

' ------------------------------------------------------------------------------

Public Function CBHP(ByVal Height As Double, _
                     ByVal DensObs As Double, _
                     ByVal Diameter As Double, _
                     ByVal AvgThickness As Double, _
                     Optional ByVal IsMetric As Boolean = False) As Double
' Calcula el Factor de Corrección por Carga Hidrostática (CBhp).
' Compensa la expansión elástica de la coraza del tanque por el peso del
' líquido. Basado en API MPMS 12.1.1, Sección 5.3.
' Solo aplica si la tabla fue generada por métodos geométricos (medición con
' cinta, láser o triangulación).
  
  On Error GoTo ErrHandler

  If Not mdHelpers.IsFinite(Height) Or Not mdHelpers.IsFinite(DensObs) Or _
     Not mdHelpers.IsFinite(Diameter) Or Not mdHelpers.IsFinite(AvgThickness) _
     Then GoTo ErrHandler
  
  If AvgThickness <= 0 Or Diameter <= 0 Then GoTo ErrHandler

  ' Definición de Constantes Físicas (Acero al Carbono API 650)
  ' Módulo de Elasticidad (Young's Modulus)
  ' Imperial: 30,000,000 psi | Métrico: 206,842,700,000 Pa (207 GPa)
  Dim youngsModulus As Double
  Dim gravityConstant As Double
  
  If Not IsMetric Then
    youngsModulus = cYOUNG_MODULUS_STEEL_IMP ' psi
    gravityConstant = 0.4335 ' psi/ft (presión del agua por pie de altura)
  Else
    youngsModulus = cYOUNG_MODULUS_STEEL_MET ' Pa
    gravityConstant = 9806.65 ' Pa/m (presión del agua por metro de altura)
  End If

  ' Cálculo de la Presión Hidrostática Promedio (P)
  ' P = Densidad_Relativa * Constante_Gravedad * Altura_Liquido
  ' Nota: La presión se evalúa usualmente a la mitad de la columna de líquido 
  ' para obtener una deformación promedio en la coraza.
  Dim Pressure As Double
  Pressure = DensObs * gravityConstant * (Height / 2)
  
  ' Cálculo de la deformación radial (Hoop Stress principle)
  ' La fórmula simplificada de expansión de volumen para un cilindro delgado:
  ' Delta_V / V = (P * D) / (2 * E * t) 
  ' CBhp = 1 + (Delta_V / V)  
  Dim diameterForCalculation As Double
  diameterForCalculation = Diameter
  
  Dim thicknessForCalculation As Double
  thicknessForCalculation = AvgThickness
  
  ' Ajuste de unidades de espesor si es necesario
  ' Si es Imperial, convertimos Diámetro a pulgadas para que coincida con
  ' E (psi) y espesor (in)
  If Not IsMetric Then
    diameterForCalculation = Diameter * 12
  End If
  
  Dim expansionFactor As Double
  expansionFactor = (Pressure * diameterForCalculation) / (2 * youngsModulus * thicknessForCalculation)
  
  CBHP = 1 + expansionFactor

  If Not mdHelpers.IsFinite(CBHP) Or CBHP < 1 Then CBHP = 1
  Exit Function

ErrHandler:
  Debug.Print "Error en CBHP: " & Err.Description
  CBHP = 1
End Function