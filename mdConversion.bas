Attribute VB_Name = "mdConversion"
Option Explicit

' ------------------------------------------------------------------------------
' MÓDULO CENTRAL: mdConversion
' DESCRIPCIÓN:
' Implementa los algoritmos de conversión de unidades operacionales y los
' coeficientes de corrección térmica de fluidos bajo el estándar API MPMS
' Sección 11.1.
'
' DEPENDENCIAS:
'   - mdGlobals: Para definiciones de enumeraciones (Enums) y constantes 
'                API MPMS.
' ------------------------------------------------------------------------------

' ------------------------------------------------------------------------------
' FUNCIONES DE CONVERSION DE UNIDADES
' ------------------------------------------------------------------------------

Public Function CONVPRES(ByVal Pressure As Double, _
                         ByVal SourceUnits As ePrsUnits, _
                         ByVal TargetUnits As ePrsUnits) As Double
' Realiza la conversión de unidades de presión (PSI, Bares, Kilopascales) bajo
' factores de conversión y usando PSI como pivote matemático.

  On Error GoTo ErrHandler

  If SourceUnits = TargetUnits Then
    CONVPRES = Pressure
    Exit Function
  End If

  ' API MPMS 11.1 (Pag 22)
  Const cPSI_TO_KPA As Double = 6.8947590868
  Const cPSI_TO_BAR As Double = 6.8947590868E-2

  ' Convertir entrada a PSI
  Dim psiPivot As Double  
  Select Case SourceUnits
    Case PSI: psiPivot = Pressure
    Case BAR: psiPivot = Pressure / cPSI_TO_BAR
    Case KPA: psiPivot = Pressure / cPSI_TO_KPA
    Case Else: GoTo ErrHandler
  End Select

  ' Convertir de PSI a Destino
  Select Case TargetUnits
    Case PSI: CONVPRES = psiPivot
    Case BAR: CONVPRES = psiPivot * cPSI_TO_BAR
    Case KPA: CONVPRES = psiPivot * cPSI_TO_KPA
    Case Else: GoTo ErrHandler
  End Select
  Exit Function

ErrHandler:
  Debug.Print "Error en CONVPRES: " & Err.Description
  CONVPRES = 0 ' Retorno de seguridad sin correccion
End Function

' ------------------------------------------------------------------------------

Public Function CONVTEMP(ByVal Temperature As Double, _
                         ByVal SourceUnits As eTempUnits, _
                         ByVal TargetUnits As eTempUnits) As Double
' Realiza la conversión de unidades de temperatura (Celsius, Fahrenheit, Kelvin,
' Rankine) bajo factores de conversión y usando Celsius como pivote matemático.
  
  On Error GoTo ErrHandler

  If SourceUnits = TargetUnits Then
    CONVTEMP = Temperature
    Exit Function
  End If

  ' Factores de Conversión API MPMS 11.1 (Pag 22)
  'Const cFAHRENHEIT_ABSOLUTE_ZERO As Double = 459.67
  Const cKELVIN_OFFSET As Double = 273.15
  Const cRANKINE_OFFSET As Double = 491.67
  Const cC_TO_F_MULTIPLIER As Double = 9 / 5
  Const cF_TO_C_MULTIPLIER As Double = 5 / 9
  Const cFAHRENHEIT_OFFSET As Double = 32

  ' Convertir entrada a Celsius
  Dim celsiusPivot As Double
  Select Case SourceUnits
    Case CLS: celsiusPivot = Temperature
    Case FHR: celsiusPivot = (Temperature - cFAHRENHEIT_OFFSET) * cF_TO_C_MULTIPLIER
    Case KLV: celsiusPivot = Temperature - cKELVIN_OFFSET
    Case RNK: celsiusPivot = (Temperature - cRANKINE_OFFSET) * cF_TO_C_MULTIPLIER
    Case Else: GoTo ErrHandler
  End Select

  ' Convertir de Celsius a Destino
  Select Case TargetUnits
    Case CLS: CONVTEMP = celsiusPivot
    Case FHR: CONVTEMP = celsiusPivot * cC_TO_F_MULTIPLIER + cFAHRENHEIT_OFFSET
    Case KLV: CONVTEMP = celsiusPivot + cKELVIN_OFFSET
    Case RNK: CONVTEMP = celsiusPivot * cC_TO_F_MULTIPLIER + cRANKINE_OFFSET
    Case Else: GoTo ErrHandler
  End Select
  Exit Function

ErrHandler:
  Debug.Print "Error en CONVTEMP: " & Err.Description
  CONVTEMP = 0 ' Retorno de seguridad sin correccion
End Function

' ------------------------------------------------------------------------------

Public Function CONVDENS(ByVal Density As Double, _
                         ByVal SourceUnits As eDnsUnits, _
                         ByVal TargetUnits As eDnsUnits, _
                         Optional ByVal WaterRel As Boolean = True) As Double
' Realiza la conversión de unidades de densidad entre Grados API, Gravedad
' Específica (SGU) y Densidad Absoluta (Kg/m³), aplicando las ecuaciones de
' calibración del estándar API MPMS Capítulo 11.1 y usando Kg/m³ como pivote.
  
  On Error GoTo ErrHandler

  If SourceUnits = TargetUnits Then
    CONVDENS = Density
    Exit Function
  End If
  
  ' Ajuste de densidad del agua según API MPMS 11.1
  Dim waterDensityKgM3 As Double  
  waterDensityKgM3 = IIf(WaterRel, cWATERDENSKG_60F, 1000)

  ' Convertir Entrada a Kg/m3 (Pivote)
  Dim kgm3Pivot As Double
  Select Case SourceUnits
    Case KGM: kgm3Pivot = Density
    Case SGU: kgm3Pivot = Density * waterDensityKgM3
    Case API
      ' Evitar división por cero
      If (Density + cAPI_B) = 0 Then GoTo ErrHandler
      kgm3Pivot = (cAPI_A * waterDensityKgM3) / (Density + cAPI_B)
    Case Else: GoTo ErrHandler
  End Select

  ' Convertir Kg/m3 a Destino
  Select Case TargetUnits
    Case KGM: CONVDENS = kgm3Pivot
    Case SGU: CONVDENS = kgm3Pivot / waterDensityKgM3
    Case API
      If kgm3Pivot = 0 Then GoTo ErrHandler
      CONVDENS = (cAPI_A * waterDensityKgM3 / kgm3Pivot) - cAPI_B
    Case Else: GoTo ErrHandler
  End Select
  Exit Function

ErrHandler:
  Debug.Print "Error en CONVDENS: " & Err.Description
  CONVDENS = 0 ' Retorno de seguridad sin correccion
End Function

' ------------------------------------------------------------------------------

Public Function CONVVOL(ByVal Volume As Double, _
                        ByVal SourceUnits As eVolUnits, _
                        ByVal TargetUnits As eVolUnits) As Double
' Realiza la conversión de unidades de volumen (Barriles, Metros Cúbicos,
' Galones, Litros) bajo factores de conversión y usando Barriles como pivote 
' matemático.

  On Error GoTo ErrHandler
  
  If SourceUnits = TargetUnits Then
    CONVVOL = Volume
    Exit Function
  End If
  
  ' Factores de conversión API (Pivote: BBL)
  Const BBL_TO_MT3 As Double = 0.1589872386
  Const BBL_TO_GAL As Double = 42.000008585
  Const BBL_TO_LTR As Double = 158.98723857
  
  ' Convertir Entrada a Barriles (Pivote)
  Dim bblPivot As Double
  Select Case SourceUnits
    Case BBL: bblPivot = Volume
    Case MT3: bblPivot = Volume / BBL_TO_MT3
    Case GAL: bblPivot = Volume / BBL_TO_GAL
    Case LTR: bblPivot = Volume / BBL_TO_LTR
  End Select

  ' Convertir Barriles a Destino
  Select Case TargetUnits
    Case BBL: CONVVOL = bblPivot
    Case MT3: CONVVOL = bblPivot * BBL_TO_MT3
    Case GAL: CONVVOL = bblPivot * BBL_TO_GAL
    Case LTR: CONVVOL = bblPivot * BBL_TO_LTR
  End Select
  Exit Function

ErrHandler:
  Debug.Print "Error en CONVVOL: " & Err.Description
  CONVVOL = 0 ' Retorno de seguridad sin correccion
End Function

' ------------------------------------------------------------------------------

Public Function CONVMASS(ByVal Mass As Double, _
                         ByVal SourceUnits As eMassUnits, _
                         ByVal TargetUnits As eMassUnits) As Double
' Realiza la conversión de unidades de masa (Kilogramos, Libras, Toneladas) bajo
' factores de conversión y usando Kilogramos como pivote matemático.

  On Error GoTo ErrHandler
  
  If SourceUnits = TargetUnits Then
    CONVMASS = Mass
    Exit Function
  End If

  ' Factores de conversión (Pivote: KGR)
  Const KGR_TO_LBR As Double = 2.2046226218
  Const KGR_TO_TON As Double = 0.001 ' Tonelada Métrica
  
  ' Convertir entrada a Kilogramo (Pivote)
  Dim kgrPivot As Double
  Select Case SourceUnits
      Case KGR: kgrPivot = Mass
      Case LBR: kgrPivot = Mass / KGR_TO_LBR
      Case TON: kgrPivot = Mass / KGR_TO_TON
      Case Else: GoTo ErrHandler
  End Select
  
  ' Convertir de Kilogramo a Destino
  Select Case TargetUnits
      Case KGR: CONVMASS = kgrPivot
      Case LBR: CONVMASS = kgrPivot * KGR_TO_LBR
      Case TON: CONVMASS = kgrPivot * KGR_TO_TON
      Case Else: GoTo ErrHandler
  End Select
  Exit Function

ErrHandler:
  Debug.Print "Error en CONVMASS: " & Err.Description
  CONVMASS = 0 ' Retorno de seguridad sin correccion
End Function

' ------------------------------------------------------------------------------

Public Function CONVLENGTH(ByVal Length As Double, _
                           ByVal SourceUnits As eLengthUnits, _
                           ByVal TargetUnits As eLengthUnits) As Double
' Realiza la conversión de unidades de Longitud (Metros, Pies, Pulgadas) bajo 
' factores de conversión y usando Metros como pivote matemático.

  On Error GoTo ErrHandler
  
  If SourceUnits = TargetUnits Then
    CONVLENGTH = Length
    Exit Function
  End If

  ' Factores de conversión oficiales API/ASTM (Pivote: MTR)
  Const MTR_TO_FTS As Double = 3.280839895
  Const MTR_TO_INC As Double = 39.37007874  
  
  ' Convertir entrada a Metros (Pivote)
  Dim mtrPivot As Double
  Select Case SourceUnits
    Case MTR: mtrPivot = Length
    Case FTS: mtrPivot = Length / MTR_TO_FTS
    Case INC: mtrPivot = Length / MTR_TO_INC
    Case Else: GoTo ErrHandler
  End Select
  
  ' 2. Convertir de Metros a Destino
  Select Case TargetUnits
    Case MTR: CONVLENGTH = mtrPivot
    Case FTS: CONVLENGTH = mtrPivot * MTR_TO_FTS
    Case INC: CONVLENGTH = mtrPivot * MTR_TO_INC
    Case Else: GoTo ErrHandler
  End Select
  Exit Function

ErrHandler:
  Debug.Print "Error en CONVLENGTH: " & Err.Description
  CONVLENGTH = 0 ' Retorno de seguridad sin correccion
End Function

' ------------------------------------------------------------------------------

Public Function CONVTEMP68(ByVal Temperature As Double, _
                           Optional ByVal TempUnits As eTempUnits = FHR)
' Convierte un valor de temperatura desde la escala moderna ITS-90
' (International Temperature Scale of 1990) hacia la escala previa IPTS-68
' (International Practical Temperature Scale of 1968)

  On Error GoTo ErrHandler

  ' Normalización a Celsius (Pivote CLS)
  Dim temp90 As Double
  If TempUnits = CLS Then
    temp90 = Temperature
  Else
    temp90 = CONVTEMP(Temperature, TempUnits, CLS)
  End If
' Validación de rango operativo (API 11.1: -183°C a 630°C)
  If temp90 < cTEMPSCALE_MIN Or temp90 > cTEMPSCALE_MAX Then
    ' Fuera de rango para el polinomio oficial, se retorna el valor original
    CONVTEMP68 = Temperature
    Exit Function
  End If

  ' Cálculo del Delta mediante Helper Polinomial
  Dim deltaTemp As Double
  deltaTemp = GetDeltaIPTS68(temp90)
  
  ' Aplicación de la corrección
  Dim temp68 As Double
  temp68 = temp90 - deltaTemp

  ' Desnormalización a la unidad de entrada
  If TempUnits = CLS Then
    CONVTEMP68 = temp68
  Else
    CONVTEMP68 = CONVTEMP(temp68, CLS, TempUnits)
  End If
  Exit Function

ErrHandler:
  Debug.Print "Error en CONVTEMP68: " & Err.Description
  CONVTEMP68 = 0 ' Retorno de seguridad sin correccion
End Function

' ------------------------------------------------------------------------------

Private Function GetDeltaIPTS68(ByVal Temp90C As Double) As Double
' Implementa la sumatoria asintótica de potencias de la escala API

  ' Factor de normalización adimensional (tau)
  Dim tau As Double
  tau = Temp90C / 630

  ' Coeficientes oficiales API MPMS 11.1 / IPTS-68
  Dim coeffs As Variant
  coeffs = Array(0, -0.148759, -0.267408, 1.08076, 1.269056, _
                    -4.089591, -1.871251, 7.438081, -3.536296)

  ' Cálculo eficiente del polinomio (Serie de potencias)
  Dim delta As Double
  delta = 0
  
  Dim counter As Integer
  For counter = 1 To 8
    delta = delta - coeffs(counter) * (tau ^ counter)
  Next counter

  GetDeltaIPTS68 = delta
End Function

' ------------------------------------------------------------------------------

Public Function CONVDENS68(ByVal Density60 As Double, _
                           Optional ByVal DnsUnits As eDnsUnits = API, _
                           Optional ByVal TypeLiq As eTypeLiq = CRD, _
                           Optional ByVal Alfa60 As Double = 0) As Double
' Ajusta la densidad base de un fluido convertida de grados API a 60°F
' (calculada originalmente en la escala térmica moderna ITS-90) hacia la escala
' práctica previa IPTS-68, aplicando el algoritmo iterativo oficial del estándar
' API MPMS Capítulo 11.1.

  On Error GoTo ErrHandler

  ' Conversion de la densidad de entrada a Gravedad API60 (Pivote para cálculos)
  Dim API60 As Double
  If DnsUnits = KGM Then
    API60 = Density60
  Else
    API60 = mdConversion.CONVDENS(Density60, DnsUnits, API, True)
  End If
  
  Dim rho60 as Double
  rho60 = mdConversion.CONVDENS(API60, API, KGM, True)

  If API60 <= 0 Then GoTo ErrHandler
  
  Dim rho68 As Double
  Dim alfa60_Calc As Double
  Dim rho68_Calc as Double
  Dim deltaRho As Double

  ' Obtención del Coeficiente de Expansión Térmica (Alpha60)
  If Alfa60 = 0 Then
    ' Calcular Alfa60 con la misma lógica usada en mdAPICalcs.bas
    alfa60_Calc = mdAPICalcs.ALPHA60(API60, TypeLiq, rho68_Calc)
    If alfa60_Calc = 0 Or rho68_Calc <= 0 Then GoTo ErrHandler
    rho68 = rho68_Calc
  Else
    ' Ajuste simplificado directo si el usuario inyecta de forma explícita el
    ' coeficiente Alfa60
    deltaRho = Exp(0.5 * Alfa60 * cTEMPSHIFT * (1 + 0.4 * Alfa60 * cTEMPSHIFT))
    ' Aplicación y Desnormalización a la unidad de entrada
    rho68 = rho60 * deltaRho
  End If

  If DnsUnits = KGM Then
    CONVDENS68 = rho68
  Else
    CONVDENS68 = mdConversion.CONVDENS(rho68, KGM, DnsUnits, True)
  End If
  Exit Function

ErrHandler:
  Debug.Print "Error en CONVDENS68: " & Err.Description
  CONVDENS68 = Density60 ' Retorno de seguridad sin correccion
End Function