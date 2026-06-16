Attribute VB_Name = "mdTestCalcs"
Option Explicit

''' <summary>
''' Ejecuta una bater�a de pruebas para validar la precisi�n del motor API.
''' Los resultados se imprimen en la Ventana de Inmediato (Ctrl + G).
''' </summary>
Public Sub RunAPIEngineTests()
    Debug.Print "======================================================"
    Debug.Print "INICIANDO TEST DE MOTOR API MPMS 11.1"
    Debug.Print "======================================================"
    
    Test_RoundingLogic
    Test_HydrometerCorrection
    Test_NewtonRaphson_Crude
    
    Debug.Print "======================================================"
    Debug.Print "PRUEBAS FINALIZADAS"
    Debug.Print "======================================================"
End Sub

''' <summary>
''' Escenario 1: Validaci�n del Redondeo al Par (Banker's Rounding).
''' </summary>
Private Sub Test_RoundingLogic()
    Debug.Print "[TEST 1] Redondeo al Par (Regla API):"
    
    ' El .5 debe ir al par m�s cercano
    Debug.Print "  2.5 -> " & mdAPICalcs.ROUNDAPI(2.5, 0) & " (Esperado: 2)"
    Debug.Print "  3.5 -> " & mdAPICalcs.ROUNDAPI(3.5, 0) & " (Esperado: 4)"
    
    ' Prueba de incremento (al 0.05 m�s cercano)
    Dim valCustom As Double: valCustom = 10.126
    Debug.Print "  " & valCustom & " al 0.05 -> " & mdAPICalcs.ROUNDAPI(valCustom, 2, 5) & " (Esperado: 10.15)"
    Debug.Print "------------------------------------------------------"
End Sub

''' <summary>
''' Escenario 2: Correcci�n de Hidr�metro (API 9.3).
''' </summary>
Private Sub Test_HydrometerCorrection()
    Debug.Print "[TEST 2] Correcci�n de Hidr�metro (HYC):"
    
    Dim apiObs As Double: apiObs = 30#
    Dim tempObs As Double: tempObs = 100# ' Mucho m�s caliente que la base 60�F
    
    Dim apiCorr As Double
    apiCorr = mdAPICalcs.DENSHYC(apiObs, tempObs, API, F)
    
    ' A 100�F el vidrio se expande, el hidr�metro flota m�s y la lectura es ligeramente err�nea.
    Debug.Print "  API Obs: " & apiObs & " @ " & tempObs & "F"
    Debug.Print "  API Corr (Vidrio): " & Round(apiCorr, 4)
    Debug.Print "------------------------------------------------------"
End Sub

''' <summary>
''' Escenario 3: Motor Completo Newton-Raphson (API60F).
''' Caso: Crudo a 95�F y 200 PSI.
''' </summary>
Private Sub Test_NewtonRaphson_Crude()
    Debug.Print "[TEST 3] Motor Newton-Raphson (Conversi�n a 60�F):"
    
    Dim apiObs As Double: apiObs = 35.4
    Dim tempObs As Double: tempObs = 95.5
    Dim presObs As Double: presObs = 215.3 ' Presi�n significativa
    
    ' 1. Corregir Hidr�metro
    Dim apiHyd As Double: apiHyd = mdAPICalcs.DENSHYC(apiObs, tempObs, API, F)
    
    ' 2. Hallar API60 (Iterativo)
    Dim api60 As Double
    api60 = mdAPICalcs.API60F(apiHyd, tempObs, CRD, presObs, 0)
    
    ' 3. Obtener factores finales para auditor�a
    Dim fCTL As Double: fCTL = mdAPICalcs.CTL(api60, tempObs, CRD)
    Dim fCPL As Double: fCPL = mdAPICalcs.CPL(api60, tempObs, presObs, CRD, 0)
    
    Debug.Print "  Datos Entrada: API " & apiObs & " / T " & tempObs & "F / P " & presObs & "psi"
    Debug.Print "  >>> RESULTADO API60: " & mdAPICalcs.ROUNDAPI(api60, 1)
    Debug.Print "  Factor CTL: " & Round(fCTL, 6)
    Debug.Print "  Factor CPL: " & Round(fCPL, 6)
    Debug.Print "  Factor Combinado (CTPL): " & Round(fCTL * fCPL, 6)
    Debug.Print "------------------------------------------------------"
End Sub