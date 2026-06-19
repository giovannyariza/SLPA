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
' factores de conversión API y usando PSI como pivote matemático.

  On Error GoTo ErrHandler

  ' API MPMS 11.1 (Pag 22)
  Const cPSI_TO_KPA As Double = 6.8947590868
  Const cPSI_TO_BAR As Double = 6.8947590868E-2
  
  If SourceUnits = TargetUnits Then
    CONVPRES = Pressure
    Exit Function
  End If

  Dim psiPivote As Double
  Select Case SourceUnits
    Case PSI
      psiPivote = Pressure
    Case BAR
      psiPivote = Pressure / cPSI_TO_BAR
    Case KPA
      psiPivote = Pressure / cPSI_TO_KPA
    Case Else
      GoTo ErrHandler
  End Select

  Select Case TargetUnits
    Case PSI
      CONVPRES = psiPivote
    Case BAR
      CONVPRES = psiPivote * cPSI_TO_BAR
    Case KPA
      CONVPRES = psiPivote * cPSI_TO_KPA
    Case Else
      GoTo ErrHandler
  End Select
  Exit Function

ErrHandler:
  Debug.Print "Error en CONVPRES: " & Err.Description
  CONVPRES = 0 ' Retorno de seguridad sin correccion
End Function

' ------------------------------------------------------------------------------

Public Function CONVTEMP(ByVal Temperature As Double, _
                         ByVal SourceUnits As eTmpUnits, _
                         ByVal TargetUnits As eTmpUnits) As Double
' Realiza la conversión de unidades de temperatura (Celsius, Fahrenheit, Kelvin,
' Rankine) usando Celsius como pivote matemático.
  
  On Error GoTo ErrHandler

  ' API MPMS 11.1 (Pag 22)
  Const cKELVIN_OFFSET As Double = 273.15
  Const cRANKINE_OFFSET As Double = 491.67
  Const cFAHRENHEIT_ABSOLUTE_ZERO As Double = 459.67
  Const cC_TO_F_MULTIPLIER As Double = 9 / 5
  Const cF_TO_C_MULTIPLIER As Double = 5 / 9
  Const cFAHRENHEIT_OFFSET As Double = 32

  If SourceUnits = TargetUnits Then
    CONVTEMP = Temperature
    Exit Function
  End If

  Dim celsiusPivot As Double
  Select Case SourceUnits
    Case C
      celsiusPivot = Temperature
    Case F
      celsiusPivot = (Temperature - cFAHRENHEIT_OFFSET) * cF_TO_C_MULTIPLIER
    Case K
      celsiusPivot = Temperature - cKELVIN_OFFSET
    Case R
      celsiusPivot = Temperature * cF_TO_C_MULTIPLIER
    Case Else
      GoTo ErrHandler
  End Select

  Select Case TargetUnits
    Case C
      CONVTEMP = celsiusPivot
    Case F
      CONVTEMP = celsiusPivot * cC_TO_F_MULTIPLIER + cFAHRENHEIT_OFFSET
    Case K
      CONVTEMP = celsiusPivot + cKELVIN_OFFSET
    Case R
      CONVTEMP = celsiusPivot * cC_TO_F_MULTIPLIER + cRANKINE_OFFSET
    Case Else
      GoTo ErrHandler
  End Select
  Exit Function

ErrHandler:
  Debug.Print "Error en CONVTEMP: " & Err.Description
  CONVTEMP = 0 ' Retorno de seguridad sin correccion
End Function

' ------------------------------------------------------------------------------

Public Function CONVVOL(ByVal Volume As Double, _
                        ByVal SourceUnits As eVolUnits, _
                        ByVal TargetUnits As eVolUnits) As Double
' Convierte un valor de volumen entre diferentes unidades operacionales
' utilizando el Barril (BBL) como pivote matemático.

  On Error GoTo ErrHandler

  ' Constantes de Conversión de Volumen
  ' https://www.convertworld.com/es/volumen/
  Const cBBL2GAL As Double = 42.000008585
  Const cBBL2M3 As Double = 0.15898723857
  Const cBBL2LT As Double = 158.98723857
    
  ' 1. VALIDACIÓN FILTRADA DE INTEGRIDAD
  If Not mdHelpers.IsFinite(Volume) Or Volume < 0 Then GoTo ErrHandler
  
  ' 2. ATRAVIESE DIRECTO: Si las unidades coinciden se retorna el valor de forma inmediata
  If SourceUnits = TargetUnits Then
    CONVVOL = Volume
    Exit Function
  End If
  
  ' 3. REDUCCIÓN A UNIDAD PIVOTE (Convertir de Origen a Barriles - BBL)
  Dim bblPivote As Double
  Select Case SourceUnits
    Case BBL: bblPivote = Volume
    Case GAL: bblPivote = Volume / cBBL2GAL
    Case M3:  bblPivote = Volume / cBBL2M3
    Case Lt:  bblPivote = Volume / cBBL2LT
    Case Else: GoTo ErrHandler
  End Select
  
  ' Double-check de estabilidad numérica en el pivote
  If Not mdHelpers.IsFinite(bblPivote) Then GoTo ErrHandler
  
  ' 4. DESPLIEGUE DESDE PIVOTE (Convertir de Barriles a Destino)
  Select Case TargetUnits
    Case BBL: CONVVOL = bblPivote
    Case GAL: CONVVOL = bblPivote * cBBL2GAL
    Case M3:  CONVVOL = bblPivote * cBBL2M3
    Case Lt:  CONVVOL = bblPivote * cBBL2LT
    Case Else: GoTo ErrHandler
  End Select
  Exit Function

ErrHandler:
  ' Captura y reporte pasivo en consola inmediata, protegiendo las celdas de caídas imprevistas
  Debug.Print "Error en CONVVOL: " & Err.Description
  CONVVOL = 0 ' Retorno de seguridad sin correccion
End Function

' ------------------------------------------------------------------------------

Public Function CONVMASS(ByVal Mass As Double, _
                         ByVal SourceUnits As eMassUnits, _
                         ByVal TargetUnits As eMassUnits) As Double
' Convierte un valor de masa entre diferentes unidades utilizando el Kilogramo
' (KG) como pivote matemático
  On Error GoTo ErrHandler

  ' Constantes de Conversión de Masa
  ' https://www.convertworld.com/es/masa/
  Const cKG2TON As Double = 0.001
  Const cKG2LB As Double = 2.2046226218
  
  ' 1. VALIDACIÓN FILTRADA DE INTEGRIDAD
  If Not mdHelpers.IsFinite(Mass) Or Mass < 0 Then GoTo ErrHandler
  
  ' 2. ATRAVIESE DIRECTO
  If SourceUnits = TargetUnits Then
    CONVMASS = Mass
    Exit Function
  End If
  
  ' 3. REDUCCIÓN A UNIDAD PIVOTE (Convertir de Origen a Kilogramos - KG)
  Dim kgPivote As Double
  Select Case SourceUnits
    Case KG:  kgPivote = Mass
    Case LB:  kgPivote = Mass / cKG2LB
    Case TON: kgPivote = Mass / cKG2TON
    Case Else: GoTo ErrHandler
  End Select
  
  If Not mdHelpers.IsFinite(kgPivote) Then GoTo ErrHandler
  
  ' 4. DESPLIEGUE DESDE PIVOTE (Convertir de Kilogramos a Destino)
  Select Case TargetUnits
    Case KG:  CONVMASS = kgPivote
    Case LB:  CONVMASS = kgPivote * cKG2LB
    Case TON: CONVMASS = kgPivote * cKG2TON
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
' Convierte un valor de longitud geométrica utilizando el Metro (MT) como pivote
' matemático.

  On Error GoTo ErrHandler

  ' Constantes de Conversión de Longitud
  ' https://www.convertworld.com/es/longitud/
  Const cMT2MM As Double = 1000
  Const cMT2CM As Double = 100
  Const cMT2FT As Double = 3.280839895
  Const cMT2IN As Double = 39.37007874
    
  ' 1. VALIDACIÓN FILTRADA DE INTEGRIDAD FÍSICA
  If Not mdHelpers.IsFinite(Length) Or Length < 0 Then GoTo ErrHandler
  
  ' 2. ATRAVIESE DIRECTO
  If SourceUnits = TargetUnits Then
    CONVLENGTH = Length
    Exit Function
  End If
  
  ' 3. REDUCCIÓN A UNIDAD PIVOTE (Convertir de Origen a Metros - MT)
  Dim mtPivote As Double
  Select Case SourceUnits
    Case MR: mtPivote = Length
    Case MM: mtPivote = Length / cMT2MM
    Case CM: mtPivote = Length / cMT2CM
    Case FT: mtPivote = Length / cMT2FT
    Case IC: mtPivote = Length / cMT2IN
    Case Else: GoTo ErrHandler
  End Select
  
  If Not mdHelpers.IsFinite(mtPivote) Then GoTo ErrHandler
  
  ' 4. DESPLIEGUE DESDE PIVOTE (Convertir de Metros a Destino)
  Select Case TargetUnits
    Case MR: CONVLENGTH = mtPivote
    Case MM: CONVLENGTH = mtPivote * cMT2MM
    Case CM: CONVLENGTH = mtPivote * cMT2CM
    Case FT: CONVLENGTH = mtPivote * cMT2FT
    Case IC: CONVLENGTH = mtPivote * cMT2IN
    Case Else: GoTo ErrHandler
  End Select
  Exit Function

ErrHandler:
  Debug.Print "Error en CONVLENGTH: " & Err.Description
  CONVLENGTH = 0 ' Retorno de seguridad sin correccion
End Function

' ------------------------------------------------------------------------------
' FUNCIONES DE CONVERSION (API MPMS Cap. 11.1)
' ------------------------------------------------------------------------------

Public Function CONVDENS(ByVal Density As Double, _
                         ByVal SourceUnits As eDnsUnits, _
                         ByVal TargetUnits As eDnsUnits, _
                         Optional ByVal WaterRel As Boolean = True) As Double
' Realiza la conversión de unidades de densidad entre Grados API, Gravedad
' Específica (SGU) y Densidad Absoluta (Kg/m³), aplicando las ecuaciones de
' calibración del estándar API MPMS Capítulo 11.1 y usando Kg/m³ como pivote.
  
  On Error GoTo ErrHandler
  
  ' Densidad del agua a 60 grados Fahrenheit (15.56 °C) en Kg/m³.
  Const cWATER_DENSITY_60F_KGM3 As Double = cWATERDENSKG_60F

  Dim waterDensityKgM3 As Double
  Dim kgPivote As Double
  
  ' Valores base por defecto (Agua pura teórica)
  waterDensityKgM3 = 1000 ' Densidad del agua en Kg/m3
  
  ' Ajuste hidrodinámico según API MPMS Capítulo 11.1 (Pág. 212)
  If WaterRel = True Then
    waterDensityKgM3 = cWATER_DENSITY_60F_KGM3 ' Densidad del agua en Kg/m3
  End If

  If SourceUnits = TargetUnits Then
    CONVDENS = Density
    Exit Function
  End If
  
  Select Case SourceUnits
    Case API
      If Density <= -cAPI_B Then GoTo ErrHandler
      kgPivote = (cAPI_A / (Density + cAPI_B)) * waterDensityKgM3
    Case KGM
      If Density <= 0 Then GoTo ErrHandler
      kgPivote = Density
    Case SGU
      If Density <= 0 Then GoTo ErrHandler
      kgPivote = Density * waterDensityKgM3
    Case Else
      GoTo ErrHandler
  End Select

  Select Case TargetUnits
    Case API
      CONVDENS = cAPI_A / (kgPivote / waterDensityKgM3) - cAPI_B
    Case KGM
      CONVDENS = kgPivote
    Case SGU
      CONVDENS = kgPivote / waterDensityKgM3
    Case Else
      GoTo ErrHandler
  End Select
  Exit Function

ErrHandler:
  Debug.Print "Error en CONVDENS: " & Err.Description
  CONVDENS = 0 ' Retorno de seguridad sin correccion
End Function

' ------------------------------------------------------------------------------

Public Function CONVTEMP68(ByVal Tmp90 As Double, _
                           Optional ByVal TmpUnits As eTmpUnits = F)
' Convierte un valor de temperatura desde la escala moderna ITS-90
' (International Temperature Scale of 1990) hacia la escala previa IPTS-68
' (International Practical Temperature Scale of 1968)

  On Error GoTo ErrHandler

  Dim tauFactor As Double, Tmp68 As Double, deltaTemp As Double
  Dim coeff(1 To 8) As Double
  Dim coeffIndex As Byte
  
  ' NORMALIZACIÓN TÉRMICA: Convertir de Fahrenheit a Celsius
  If TmpUnits = F Then
    Tmp90 = CONVTEMP(Tmp90, F, C)
  End If
  
  ' Calculo de la temperatura escalada (tauFactor)
  ' Parámetro de normalización adimensional según API MPMS Capítulo 11.1
  ' (Apéndice A, Pág. 210)
  tauFactor = Tmp90 / 630
  
  ' Coeficientes oficiales del polinomio de ajuste API / IPTS-68
  coeff(1) = -0.148759: coeff(2) = -0.267408: coeff(3) = 1.08076:  coeff(4) = 1.269056
  coeff(5) = -4.089591: coeff(6) = -1.871251: coeff(7) = 7.438081: coeff(8) = -3.536296
  
  ' Acumulacion asintotica por serie de potencias
  deltaTemp = 0
  For coeffIndex = 1 To 8
    deltaTemp = deltaTemp + coeff(coeffIndex) * (tauFactor ^ coeffIndex)
  Next coeffIndex
  
  ' Aplicacion de la correccion a escala
  Tmp68 = Tmp90 - deltaTemp
  
  ' DESNORMALIZACIÓN: Retorna el resultado en la misma unidad fisica de entrada
  If TmpUnits = F Then
    CONVTEMP68 = CONVTEMP(Tmp68, C, F)
  Else
    CONVTEMP68 = Tmp68
  End If
  Exit Function

ErrHandler:
  Debug.Print "Error en CONVTEMP68: " & Err.Description
  CONVTEMP68 = 0 ' Retorno de seguridad sin correccion
End Function

' ------------------------------------------------------------------------------

Public Function CONVDENS68(ByVal API60 As Double, _
                           Optional ByVal TypeLiq As eTypeLiq = CRD, _
                           Optional ByVal Alfa60 As Double) As Double
' Ajusta la densidad base de un fluido convertida de grados API a 60°F
' (calculada originalmente en la escala térmica moderna ITS-90) hacia la escala
' práctica previa IPTS-68, aplicando el algoritmo iterativo oficial del estándar
' API MPMS Capítulo 11.1.

  On Error GoTo ErrHandler
  
  Dim K0 As Double, K1 As Double, K2 As Double
  Dim coeffA As Double, coeffB As Double, Rho60 As Double
  
  ' Conversion de la Gravedad API60 a Densidad en Kg/m3 Relativa a la Densidad
  ' del Agua a 60 F.
  Rho60 = CONVDENS(API60, API, KGM, True)
  
  If Alfa60 = 0 Then
    ' Selección de Constantes K (API MPMS 11.1 Tablas 6A, 6B, 6D)
    Call GetKConstants(TypeLiq, Rho60, K0, K1, K2)

    ' Blindaje analítico secundario: Si la densidad no cuadró en ningún rango,
    ' se evita el cálculo hidrodinámico
    If K0 = 0 And K1 = 0 And K2 = 0 Then GoTo ErrHandler

    ' Ejecucion de ecuaciones polinomiales de ajuste de escala
    coeffA = (cTEMPSHIFT / 2) * (((K0 / Rho60) + K1) * (1 / Rho60) + K2)    
    coeffB = ((2 * K0) + (K1 * Rho60)) / (K0 + (K1 + (K2 * Rho60) * Rho60))
    
    CONVDENS68 = Rho60 * (1 + ((Exp(coeffA * (1 + 0.8 * coeffA)) - 1) / _
                 (1 + coeffA * (1 + 1.6 * coeffA) * coeffB)))
  Else
    ' Ajuste simplificado directo si el usuario inyecta de forma explícita el
    ' coeficiente Alfa60
    CONVDENS68 = Rho60 * Exp(0.5 * Alfa60 * cTEMPSHIFT * _
                 (1 + 0.4 * Alfa60 * cTEMPSHIFT))
  End If
  Exit Function

ErrHandler:
  Debug.Print "Error en CONVDENS68: " & Err.Description
  CONVDENS68 = 0 ' Retorno de seguridad sin correccion
End Function