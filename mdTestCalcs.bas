Attribute VB_Name = "mdTestCalcs"
Option Explicit

' ---------------------------------------------------------------------------------------------------------
' MÓDULO DE PRUEBAS: mdTestCalcs
' DESCRIPCIÓN:
'   Contiene procedimientos de prueba para validar todas las funciones del módulo mdConversion.
'   Cada función tiene un procedimiento de prueba asociado con valores de ejemplo.
' ---------------------------------------------------------------------------------------------------------

' =========================================================================================================
' PRUEBAS DE CONVERSIÓN DE PRESIÓN (CONVPRES)
' =========================================================================================================

Sub Test_CONVPRES()
  Debug.Print "=========================================="
  Debug.Print "PRUEBAS: CONVPRES (Presión)"
  Debug.Print "=========================================="
  
  Dim psiValue As Double: psiValue = 1000
  Dim barValue As Double
  Dim kpaValue As Double
  
  ' PSI a BAR
  barValue = CONVPRES(psiValue, PSI, BAR)
  Debug.Print "CONVPRES(" & psiValue & " PSI, PSI, BAR) = " & Round(barValue, 4) & " BAR"
  
  ' PSI a KPA
  kpaValue = CONVPRES(psiValue, PSI, KPA)
  Debug.Print "CONVPRES(" & psiValue & " PSI, PSI, KPA) = " & Round(kpaValue, 4) & " KPA"
  
  ' BAR a PSI
  psiValue = CONVPRES(10, BAR, PSI)
  Debug.Print "CONVPRES(10 BAR, BAR, PSI) = " & Round(psiValue, 4) & " PSI"
  
  ' KPA a BAR
  barValue = CONVPRES(100, KPA, BAR)
  Debug.Print "CONVPRES(100 KPA, KPA, BAR) = " & Round(barValue, 4) & " BAR"
  
  ' Mismo valor (no debe convertir)
  psiValue = CONVPRES(1000, PSI, PSI)
  Debug.Print "CONVPRES(1000 PSI, PSI, PSI) = " & psiValue & " PSI (sin conversión)"
  
  Debug.Print ""
End Sub

' =========================================================================================================
' PRUEBAS DE CONVERSIÓN DE TEMPERATURA (CONVTEMP)
' =========================================================================================================

Sub Test_CONVTEMP()
  Debug.Print "=========================================="
  Debug.Print "PRUEBAS: CONVTEMP (Temperatura)"
  Debug.Print "=========================================="
  
  ' Fahrenheit a Celsius
  Dim tempC As Double
  tempC = CONVTEMP(68, FHR, CLS)
  Debug.Print "CONVTEMP(68°F, FHR, CLS) = " & Round(tempC, 2) & "°C (esperado: 20°C)"
  
  ' Celsius a Fahrenheit
  Dim tempF As Double
  tempF = CONVTEMP(100, CLS, FHR)
  Debug.Print "CONVTEMP(100°C, CLS, FHR) = " & Round(tempF, 2) & "°F (esperado: 212°F)"
  
  ' Celsius a Kelvin
  Dim tempK As Double
  tempK = CONVTEMP(0, CLS, KLV)
  Debug.Print "CONVTEMP(0°C, CLS, KLV) = " & Round(tempK, 2) & "K (esperado: 273.15K)"
  
  ' Rankine a Celsius
  tempC = CONVTEMP(491.67, RNK, CLS)
  Debug.Print "CONVTEMP(491.67°R, RNK, CLS) = " & Round(tempC, 2) & "°C (esperado: 0°C)"
  
  ' Fahrenheit a Rankine
  Dim tempR As Double
  tempR = CONVTEMP(32, FHR, RNK)
  Debug.Print "CONVTEMP(32°F, FHR, RNK) = " & Round(tempR, 2) & "°R (esperado: 491.67°R)"
  
  Debug.Print ""
End Sub

' =========================================================================================================
' PRUEBAS DE CONVERSIÓN DE DENSIDAD (CONVDENS)
' =========================================================================================================

Sub Test_CONVDENS()
  Debug.Print "=========================================="
  Debug.Print "PRUEBAS: CONVDENS (Densidad)"
  Debug.Print "=========================================="
  
  Dim apiValue As Double
  Dim sgValue As Double
  Dim kgValue As Double
  
  ' API a SGU
  apiValue = 35
  sgValue = CONVDENS(apiValue, API, SGU, True)
  Debug.Print "CONVDENS(" & apiValue & " API, API, SGU) = " & Round(sgValue, 6) & " SGU"
  
  ' SGU a API
  sgValue = 0.85
  apiValue = CONVDENS(sgValue, SGU, API, True)
  Debug.Print "CONVDENS(" & sgValue & " SGU, SGU, API) = " & Round(apiValue, 4) & " API"
  
  ' API a Kg/m³
  kgValue = CONVDENS(40, API, KGM, True)
  Debug.Print "CONVDENS(40 API, API, KGM) = " & Round(kgValue, 2) & " Kg/m³"
  
  ' Kg/m³ a SGU
  sgValue = CONVDENS(850, KGM, SGU, True)
  Debug.Print "CONVDENS(850 Kg/m³, KGM, SGU) = " & Round(sgValue, 4) & " SGU"
  
  ' SGU a Kg/m³ (agua pura = 1000)
  kgValue = CONVDENS(1.0, SGU, KGM, False)
  Debug.Print "CONVDENS(1.0 SGU, SGU, KGM, WaterRel=False) = " & kgValue & " Kg/m³"
  
  Debug.Print ""
End Sub

' =========================================================================================================
' PRUEBAS DE CONVERSIÓN DE VOLUMEN (CONVVOL)
' =========================================================================================================

Sub Test_CONVVOL()
  Debug.Print "=========================================="
  Debug.Print "PRUEBAS: CONVVOL (Volumen)"
  Debug.Print "=========================================="
  
  ' Barriles a Metros Cúbicos
  Dim mt3s As Double
  mt3s = CONVVOL(100, BBL, MT3)
  Debug.Print "CONVVOL(100 BBL, BBL, MT3) = " & Round(mt3s, 4) & " m³"
  
  ' Metros Cúbicos a Barriles
  Dim bbls As Double
  bbls = CONVVOL(10, MT3, BBL)
  Debug.Print "CONVVOL(10 m³, MT3, BBL) = " & Round(bbls, 4) & " BBL"
  
  ' Barriles a Galones
  Dim gals As Double
  gals = CONVVOL(50, BBL, GAL)
  Debug.Print "CONVVOL(50 BBL, BBL, GAL) = " & gals & " galones"
  
  ' Galones a Litros
  Dim ltrs As Double
  ltrs = CONVVOL(100, GAL, LTR)
  Debug.Print "CONVVOL(100 galones, GAL, LTR) = " & Round(ltrs, 2) & " litros"
  
  ' Litros a Metros Cúbicos
  mt3s = CONVVOL(1000, LTR, MT3)
  Debug.Print "CONVVOL(1000 litros, LTR, MT3) = " & Round(mt3s, 4) & " m³"
  
  Debug.Print ""
End Sub

' =========================================================================================================
' PRUEBAS DE CONVERSIÓN DE MASA (CONVMASS)
' =========================================================================================================

Sub Test_CONVMASS()
  Debug.Print "=========================================="
  Debug.Print "PRUEBAS: CONVMASS (Masa)"
  Debug.Print "=========================================="
  
  ' Kilograms a Libras
  Dim lbs As Double
  lbs = CONVMASS(1000, KGR, LBR)
  Debug.Print "CONVMASS(1000 kg, KGR, LBR) = " & Round(lbs, 2) & " lbs"
  
  ' Libras a Kilograms
  Dim kg As Double
  kg = CONVMASS(500, LBR, KGR)
  Debug.Print "CONVMASS(500 lbs, LBR, KGR) = " & Round(kg, 2) & " kg"
  
  ' Kilograms a Toneladas
  Dim tons As Double
  tons = CONVMASS(5000, KGR, TON)
  Debug.Print "CONVMASS(5000 kg, KGR, TON) = " & tons & " toneladas"
  
  ' Toneladas a Kilograms
  kg = CONVMASS(5, TON, KGR)
  Debug.Print "CONVMASS(5 toneladas, TON, KGR) = " & kg & " kg"
  
  Debug.Print ""
End Sub

' =========================================================================================================
' PRUEBAS DE CONVERSIÓN DE LONGITUD (CONVLENGTH)
' =========================================================================================================

Sub Test_CONVLENGTH()
  Debug.Print "=========================================="
  Debug.Print "PRUEBAS: CONVLENGTH (Longitud)"
  Debug.Print "=========================================="
  
  ' Metros a Pies
  Dim ft As Double
  ft = CONVLENGTH(100, MTR, FTS)
  Debug.Print "CONVLENGTH(100 m, MTR, FTS) = " & Round(ft, 4) & " ft"
  
  ' Pies a Metros
  Dim m As Double
  m = CONVLENGTH(50, FTS, MTR)
  Debug.Print "CONVLENGTH(50 ft, FTS, MTR) = " & Round(m, 4) & " m"
  
  ' Metros a Pulgadas
  Dim inch As Double
  inch = CONVLENGTH(1, MTR, INC)
  Debug.Print "CONVLENGTH(1 m, MTR, INC) = " & Round(inch, 4) & " in"
  
  ' Pulgadas a Pies
  ft = CONVLENGTH(24, INC, FTS)
  Debug.Print "CONVLENGTH(24 in, INC, FTS) = " & ft & " ft"
  
  Debug.Print ""
End Sub

' =========================================================================================================
' PRUEBAS DE CONVERSIÓN DE TEMPERATURA CON CORRECCIÓN IPTS-68 (CONVTEMP68)
' =========================================================================================================

Sub Test_CONVTEMP68()
  Debug.Print "=========================================="
  Debug.Print "PRUEBAS: CONVTEMP68 (Temperatura ITS-90 → IPTS-68)"
  Debug.Print "=========================================="
  
  ' Celsius ITS-90 a IPTS-68
  Dim temp68 As Double
  temp68 = CONVTEMP68(100, CLS)
  Debug.Print "CONVTEMP68(100°C ITS-90, CLS) = " & Round(temp68, 6) & "°C IPTS-68"
  
  ' Fahrenheit ITS-90 a IPTS-68
  Dim tempF68 As Double
  tempF68 = CONVTEMP68(212, FHR)
  Debug.Print "CONVTEMP68(212°F ITS-90, FHR) = " & Round(tempF68, 6) & "°F IPTS-68"
  
  ' Punto de ebullición del agua (100°C ITS-90 ≈ 99.973°C IPTS-68)
  temp68 = CONVTEMP68(100, CLS)
  Debug.Print "Punto ebullición agua: " & Round(temp68, 4) & "°C IPTS-68 (esperado ~99.97°C)"
  
  Debug.Print ""
End Sub

' =========================================================================================================
' PRUEBAS DE CONVERSIÓN DE DENSIDAD CON CORRECCIÓN IPTS-68 (CONVDENS68)
' =========================================================================================================

Sub Test_CONVDENS68()
  Debug.Print "=========================================="
  Debug.Print "PRUEBAS: CONVDENS68 (Densidad API + Corrección IPTS-68)"
  Debug.Print "=========================================="
  
  Dim rho68 As Double
  
  ' API a densidad IPTS-68 (crudo)
  rho68 = CONVDENS68(35, API, CRD)
  Debug.Print "CONVDENS68(35 API, API, CRD) = " & Round(rho68, 4) & " Kg/m³ IPTS-68"
  
  ' API a densidad IPTS-68 (refinado)
  rho68 = CONVDENS68(45, API, REF)
  Debug.Print "CONVDENS68(45 API, API, REF) = " & Round(rho68, 4) & " Kg/m³ IPTS-68"
  
  ' API a densidad IPTS-68 (lubricante)
  rho68 = CONVDENS68(25, API, LUB)
  Debug.Print "CONVDENS68(25 API, API, LUB) = " & Round(rho68, 4) & " Kg/m³ IPTS-68"
  
  Debug.Print ""
End Sub

' =========================================================================================================
' PRUEBAS DE VALIDACIÓN CRUZADA
' =========================================================================================================

Sub Test_ValidationCrossCheck()
  Debug.Print "=========================================="
  Debug.Print "PRUEBAS: Validación Cruzada (Redondeo)"
  Debug.Print "=========================================="
  
  ' Prueba redondeo API (debe ser ~11.5)
  Dim api1 As Double: api1 = 11.5
  Dim sg1 As Double: sg1 = CONVDENS(api1, API, SGU)
  Dim api2 As Double: api2 = CONVDENS(sg1, SGU, API)
  Debug.Print "API original: " & api1 & " → SGU: " & Round(sg1, 6) & " → API regresado: " & Round(api2, 6)
  
  ' Prueba redondeo temperatura
  Dim temp1 As Double: temp1 = 68
  Dim tempC As Double: tempC = CONVTEMP(temp1, FHR, CLS)
  Dim temp2 As Double: temp2 = CONVTEMP(tempC, CLS, FHR)
  Debug.Print "Temp original: " & temp1 & "°F → °C: " & Round(tempC, 4) & " → °F regresado: " & Round(temp2, 4)
  
  Debug.Print ""
End Sub

' =========================================================================================================
' EJECUTAR TODAS LAS PRUEBAS
' =========================================================================================================

Sub RunAllConversionTests()
  Debug.Print "================================================================"
  Debug.Print "EJECUTANDO PRUEBAS COMPLETAS DEL MÓDULO mdConversion"
  Debug.Print "================================================================"
  Debug.Print ""
  
  Call Test_CONVPRES
  Call Test_CONVTEMP
  Call Test_CONVDENS
  Call Test_CONVVOL
  Call Test_CONVMASS
  Call Test_CONVLENGTH
  Call Test_CONVTEMP68
  Call Test_CONVDENS68
  Call Test_ValidationCrossCheck
  
  Debug.Print "================================================================"
  Debug.Print "TODAS LAS PRUEBAS SE COMPLETARON"
  Debug.Print "================================================================"
End Sub

' =========================================================================================================
' PRUEBA INDIVIDUAL RÁPIDA (Macro asignado a botón)
' =========================================================================================================

Sub Pruebas()
' Método de prueba original simplificado
  Dim Api60 As Double: Api60 = 11.5
  Dim Volume As Double: Volume = 36265.55
  Dim Rho60 As Double: Rho60 = CONVDENS(Api60, API, SGU)
  Dim FactorMassBbl As Double: FactorMassBbl = CONVVOL(1, BBL, MT3) * Rho60
  Dim MassVol As Double: MassVol = Volume * FactorMassBbl

  Debug.Print "Densidad (SGU): " & Round(Rho60, 5)
  Debug.Print "Factor masa (ton/BBL): " & Round(FactorMassBbl, 5)
  Debug.Print "Masas total: " & Round(MassVol, 2) & " toneladas"
End Sub
