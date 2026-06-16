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
    Test_ShrinkageCalculation

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

''' <summary>
''' Escenario 4: Cálculo de Encogimiento (Shrinkage) en mezclas de hidrocarburos.
''' </summary>
Private Sub Test_ShrinkageCalculation()
    Debug.Print "[TEST] Cálculo de Encogimiento (SHRINK):"

    ' Caso 1: API pesado 11.37
    Dim X1 As Double: X1 = 6.32 ' Concentración del componente liviano [%]
    Dim APILight1 As Double: APILight1 = 62.1 ' API del componente liviano
    Dim APIHeavy1 As Double: APIHeavy1 = 11.37 ' API del componente pesado
    Dim shrink1 As Double
    shrink1 = mdAPICalcs.SHRINK(X1, APILight1, APIHeavy1)
    
    Debug.Print "  Caso 1:"
    Debug.Print "    Concentración liviana: " & X1 & "%"
    Debug.Print "    API liviano: " & APILight1
    Debug.Print "    API pesado: " & APIHeavy1
    Debug.Print "    Encogimiento: " & mdAPICalcs.ROUNDAPI(shrink1, 5, 1) & "%"
    
    ' Caso 2:
    Dim X2 As Double: X2 = 6.17 ' Concentración del componente liviano [%]
    Dim APILight2 As Double: APILight2 = 62.1 ' API del componente liviano
    Dim APIHeavy2 As Double: APIHeavy2 = 11.37 ' API del componente pesado
    Dim shrink2 As Double
    shrink2 = mdAPICalcs.SHRINK(X2, APILight2, APIHeavy2)
    
    Debug.Print "  Caso 2:"
    Debug.Print "    Concentración liviana: " & X2 & "%"
    Debug.Print "    API liviano: " & APILight2
    Debug.Print "    API pesado: " & APIHeavy2
    Debug.Print "    Encogimiento: " & mdAPICalcs.ROUNDAPI(shrink2, 5, 1) & "%"
    
    ' Caso 3: API pesado > 12.0
    Dim X3 As Double: X3 = 6.22 ' Concentración del componente liviano [%]
    Dim APILight3 As Double: APILight3 = 64.1 ' API del componente liviano
    Dim APIHeavy3 As Double: APIHeavy3 = 12.4 ' API del componente pesado
    Dim shrink3 As Double
    shrink3 = mdAPICalcs.SHRINK(X3, APILight3, APIHeavy3)
    
    Debug.Print "  Caso 3:"
    Debug.Print "    Concentración liviana: " & X3 & "%"
    Debug.Print "    API liviano: " & APILight3
    Debug.Print "    API pesado: " & APIHeavy3
    Debug.Print "    Encogimiento: " & mdAPICalcs.ROUNDAPI(shrink3, 5, 1) & "%"
    
    ' Caso 4:
    Dim X4 As Double: X4 = 6.07 ' Concentración del componente liviano [%]
    Dim APILight4 As Double: APILight4 = 64.1 ' API del componente liviano
    Dim APIHeavy4 As Double: APIHeavy4 = 12.4 ' API del componente pesado
    Dim shrink4 As Double
    shrink4 = mdAPICalcs.SHRINK(X4, APILight4, APIHeavy4)
    
    Debug.Print "  Caso 4:"
    Debug.Print "    Concentración liviana: " & X4 & "%"
    Debug.Print "    API liviano: " & APILight4
    Debug.Print "    API pesado: " & APIHeavy4
    Debug.Print "    Encogimiento: " & mdAPICalcs.ROUNDAPI(shrink4, 5, 1) & "%"
    
    ' Caso 5:
    Dim X5 As Double: X5 = 5.88 ' Concentración del componente liviano [%]
    Dim APILight5 As Double: APILight5 = 64.1 ' API del componente liviano
    Dim APIHeavy5 As Double: APIHeavy5 = 12.36 ' API del componente pesado
    Dim shrink5 As Double
    shrink5 = mdAPICalcs.SHRINK(X5, APILight5, APIHeavy5)
    
    Debug.Print "  Caso 5:"
    Debug.Print "    Concentración liviana: " & X4 & "%"
    Debug.Print "    API liviano: " & APILight5
    Debug.Print "    API pesado: " & APIHeavy5
    Debug.Print "    Encogimiento: " & mdAPICalcs.ROUNDAPI(shrink5, 5, 1) & "%"
    
    ' Caso 7:
    Dim X6 As Double: X6 = 5.74 ' Concentración del componente liviano [%]
    Dim APILight6 As Double: APILight6 = 64.1 ' API del componente liviano
    Dim APIHeavy6 As Double: APIHeavy6 = 12.36 ' API del componente pesado
    Dim shrink6 As Double
    shrink6 = mdAPICalcs.SHRINK(X6, APILight6, APIHeavy6)
    
    Debug.Print "  Caso 6:"
    Debug.Print "    Concentración liviana: " & X6 & "%"
    Debug.Print "    API liviano: " & APILight6
    Debug.Print "    API pesado: " & APIHeavy6
    Debug.Print "    Encogimiento: " & mdAPICalcs.ROUNDAPI(shrink6, 5, 1) & "%"
    
    Debug.Print "------------------------------------------------------"
End Sub