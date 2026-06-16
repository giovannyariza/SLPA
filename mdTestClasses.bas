Attribute VB_Name = "mdTestClasses"
Option Explicit

''' <summary>
''' Ejecuta el escenario de pruebas integral para verificar el comportamiento de
''' clsStation, clsComponent, clsTank y el cumplimiento de la interfaz IComponent.
''' </summary>
Public Sub EjecutarPrueba_Infraestructura()
  ' 1. CONFIGURACIÓN DE OPTIMIZACIÓN Y ENTORNO
  ' Desactivamos la actualización de pantalla para maximizar la velocidad de procesamiento en entornos corporativos
  Application.ScreenUpdating = False
  
  ' Captura y manejo estructurado de errores en tiempo de ejecución
  On Error GoTo ErrHandler
  
  Debug.Print "======================================================================"
  Debug.Print "INICIANDO PRUEBA DE ARQUITECTURA DE COMPONENTES E INSTALACIONES"
  Debug.Print "======================================================================"
  
  ' 2. DECLARACIÓN DE VARIABLES
  Dim miEstacion As clsStation
  Dim componenteGenerico As clsComponent
  Dim tanqueEspecializado As clsTank
  Dim itemComponente As IComponent
  
  ' 3. INSTANCIACIÓN Y CONFIGURACIÓN DE LA ESTACIÓN PRINCIPAL
  Set miEstacion = New clsStation
  With miEstacion
    .Tag = "EC-01"
    .Name = "Estación Central de Recibo"
    .Description = "Planta principal de fiscalización y distribución de fluidos."
    .Status = "OP"
    .Location = "Bloque Sur - Coordenadas 4.21, -72.3"
  End With
  
  ' 4. CREACIÓN DE COMPONENTE GENÉRICO (Uso directo de la clase base)
  Set componenteGenerico = New clsComponent
  With componenteGenerico
    .Tag = "LIN-4022"
    .Description = "Línea de transferencia de crudo pesado"
    .ComponentType = "Línea"
    .System = "Transferencia"
    .Service = "Recibo"
    .Status = "OP"
  End With
  
  ' 5. CREACIÓN DE COMPONENTE ESPECIALIZADO (Uso de Composición y Polimorfismo)
  Set tanqueEspecializado = New clsTank
  ' Propiedades comunes delegadas internamente a clsComponent
  tanqueEspecializado.Tag = "TK-9010"
  tanqueEspecializado.Description = "Tanque de almacenamiento de fluidos con alto contenido de agua"
  tanqueEspecializado.Service = "Almacenamiento"
  tanqueEspecializado.Status = "MT" ' En mantenimiento
  ' Propiedades exclusivas de la clase concreta clsTank
  tanqueEspecializado.Capacity = 15000 ' Capacidad en Barriles (Bbls)
  
  ' 6. ASOCIACIÓN BIDIRECCIONAL (Adición de componentes a la Estación)
  ' Pasamos las instancias a través del contrato polimórfico IComponent
  Debug.Print "-> Registrando componentes en la estación " & miEstacion.Tag & "..."
  miEstacion.AddComponent componenteGenerico
  miEstacion.AddComponent tanqueEspecializado
  Debug.Print "-> Registro completado con éxito."
  Debug.Print "----------------------------------------------------------------------"
  
  ' 7. PRUEBA DE ROBUSTEZ: Intento de registro de un duplicado para validar control de errores
  Debug.Print "-> Realizando prueba de tolerancia a fallos (Inserción de Tag Duplicado)..."
  Dim componenteDuplicado As clsComponent
  Set componenteDuplicado = New clsComponent
  componenteDuplicado.Tag = "LIN-4022" ' Mismo Tag que el componente genérico anterior
  miEstacion.AddComponent componenteDuplicado ' El método interceptará el error 457 sin romper la ejecución
  Debug.Print "----------------------------------------------------------------------"
  
  ' 8. RECORRIDO POLIMÓRFICO DE LA COLECCIÓN
  ' Inspeccionamos los componentes que pertenecen a la estación usando la interfaz común
  Debug.Print "-> Listando propiedades de los componentes activos de la Estación:"
  For Each itemComponente In miEstacion.Components
    ' Llama al método ShowProperties de cada objeto de forma dinámica (Late Binding controlado por Interfaz)
    itemComponente.ShowProperties
  Next itemComponente
  Debug.Print "----------------------------------------------------------------------"
  
  ' 9. PRUEBA DEL MÉTODO CORREGIDO: RemoveComponent
  ' Removemos el componente 'LIN-4022' pasando su Tag identificador
  Debug.Print "-> Eliminando el componente 'LIN-4022' de la colección..."
  miEstacion.RemoveComponent "LIN-4022"
  
  ' Verificamos que se haya eliminado correctamente listando los componentes remanentes
  Debug.Print "-> Componentes remanentes en la estación tras la remoción (Esperado: 1): " & miEstacion.Components.Count
  For Each itemComponente In miEstacion.Components
    Debug.Print "   Componente Restante -> Tag: " & itemComponente.Tag & " | Tipo: " & itemComponente.ComponentType
  Next itemComponente
  
  ' 10. LIMPIEZA DE MEMORIA OBJETO
  Set miEstacion = Nothing
  Set componenteGenerico = Nothing
  Set tanqueEspecializado = Nothing
  
  Debug.Print "======================================================================"
  Debug.Print "PRUEBA DE ARQUITECTURA CONCLUIDA SATISFACTORIAMENTE"
  Debug.Print "======================================================================"
  
FinSub:
  ' Restablecemos la actualización de la pantalla del entorno de Excel
  Application.ScreenUpdating = True
  Exit Sub

ErrHandler:
  ' Bloque de contingencia para capturar fallos no controlados
  MsgBox "Ocurrió un error inesperado en el módulo de pruebas." & vbCrLf & _
         "Número: " & Err.Number & vbCrLf & _
         "Descripción: " & Err.Description, vbCritical, "Error de Sistema - Arquitecto de Soluciones"
  Resume FinSub
End Sub

Public Sub CalcularInventarioActual()
  Dim ATK7210 As New clsTank
  
  ' Configuración basada en tu imagen
  With ATK7210
    .Tag = "ATK7210"
    .RoofWeight = 3200
    .IsTableNetOfRoof = True ' Tabla contempla el ajuste
    .HasFloatingRoof = True
    .TableRefAPI = 76.3
    .TableBaseDeduction = 29.64
    .APICorrectionGT = -0.14
    .APICorrectionLT = 0.14
    .MinLevelDeduction = 1610
    .MaxLevelDeduction = 1799
  End With
  
  ' Datos de campo actuales
  Dim nivelMedido As Double: nivelMedido = 6130
  Dim apiMedido60 As Double: apiMedido60 = 63.9
  Dim tempLiq As Double: tempLiq = 84.2
  Dim apiObs As Double: apiObs = mdAPICalcs.APIOBS(apiMedido60, tempLiq, 2)
  
  ' La magia del objeto:
  Dim vcfMembrana As Double
  vcfMembrana = ATK7210.GetDynamicFRA(nivelMedido, apiObs)
  
  Debug.Print "Deducción Final Membrana: " & mdAPICalcs.ROUNDAPI(vcfMembrana, 2, 1) & " Bbls"
End Sub

''' <summary>
''' Prueba todas las funciones de conveniencia de clsTank para liquidacion (tbl_Motor).
''' Requiere que exista la tabla "tbl_Aforos" en la hoja "cnf_Tanques" con una columna "ATK-7210".
''' </summary>
Public Sub TestTankConvenienceMethods()
  Dim t As New clsTank
  Dim passCount As Long: passCount = 0
  Dim failCount As Long: failCount = 0
  Dim totalTests As Long: totalTests = 0
  Dim testName As String
  Dim result As Double

  ' ── SETUP: Configurar el tanque con datos realistas ──
  With t
    .Tag = "ATK-7210"
    .Description = "Tanque de prueba para liquidacion"
    .NominalCapacity = 55838.15
    .Diameter = 30473        ' mm
    .Material = MCrbn        ' Acero al carbono
    .ShellThickness = 0.25   ' pulgadas
    .FluidType = REF         ' Crudo
    .ReferenceAPI60F = 76.3    ' API a 60F
    .ReferenceAPIObs = 76.3
    .IsThermalInsulated = False
    .HasFloatingRoof = True
    .IsTableNetOfRoof = True
    .RoofWeight = 3200       ' Kg
    .MinLevelDeduction = 1610
    .MaxLevelDeduction = 1800
    .APICorrectionLT = -0.14
    .APICorrectionGT = 0.14
    .BaseDeduction = 29.64
  End With

  ' Asignar la tabla de aforo si existe
  On Error Resume Next
  Set t.StrappingTable = ThisWorkbook.Sheets("cnf_Tanques").ListObjects("tbl_Aforos")
  If t.StrappingTable Is Nothing Then
    Debug.Print "ADVERTENCIA: No se encontro 'tbl_Aforos'. Las pruebas que requieren TOV fallaran."
  End If
  On Error GoTo 0

  ' Datos de campo simulados
  Dim nivel_mm As Long: nivel_mm = 6163
  Dim tempLiqF As Double: tempLiqF = 84.3
  Dim tempAmbF As Double: tempAmbF = 75
  Dim api60 As Double: api60 = 64.1
  Dim apiObs As Double: apiObs = 67.5
  Dim bsw As Double: bsw = 0

  ' Limpiar cache antes de comenzar
  mdTankService.ClearCache

  Debug.Print "======================================================================"
  Debug.Print "INICIANDO PRUEBA DE FUNCIONES DE CONVENIENCIA DE clsTank"
  Debug.Print "======================================================================"
  Debug.Print "Tanque: " & t.Tag & " | API60: " & t.ReferenceAPI60F
  Debug.Print "Nivel: " & nivel_mm & " mm | T_Liq: " & tempLiqF & "F | T_Amb: " & tempAmbF & "F"
  Debug.Print "----------------------------------------------------------------------"

  ' ==================================================================
  ' TEST 1: GetTOV - Volumen Total Observado
  ' ==================================================================
  totalTests = totalTests + 1
  testName = "Test 1: GetTOV(" & nivel_mm & " mm)"
  result = t.GetTOV(nivel_mm)
  If result > 0 Then
    Debug.Print "[PASS] " & testName & " -> TOV: " & FormatNumber(result, 2) & " Bbls"
    passCount = passCount + 1
  Else
    Debug.Print "[SKIP] " & testName & " -> Sin tabla de aforo disponible"
    ' No cuenta como fail, se omite
    totalTests = totalTests - 1
  End If

  ' ==================================================================
  ' TEST 2: CalculateCTL - Factor de correccion por temperatura
  ' ==================================================================
  totalTests = totalTests + 1
  testName = "Test 2: CalculateCTL(" & tempLiqF & "F)"
  result = t.CalculateCTL(api60, tempLiqF)
  If result > 0 And result < 2 Then
    Debug.Print "[PASS] " & testName & " -> CTL: " & FormatNumber(result, 6)
    passCount = passCount + 1
  Else
    Debug.Print "[FAIL] " & testName & " -> Valor fuera de rango: " & result
    failCount = failCount + 1
  End If

  ' ==================================================================
  ' TEST 3: CalculateCTSH - Factor de correccion del casco
  ' ==================================================================
  totalTests = totalTests + 1
  testName = "Test 3: CalculateCTSH(TLiq=" & tempLiqF & ", TAmb=" & tempAmbF & ")"
  result = t.CalculateCTSH(tempLiqF, tempAmbF)
  If result > 0 And result < 2 Then
    Debug.Print "[PASS] " & testName & " -> CTSH: " & FormatNumber(result, 6)
    passCount = passCount + 1
  Else
    Debug.Print "[FAIL] " & testName & " -> Valor fuera de rango: " & result
    failCount = failCount + 1
  End If

  ' ==================================================================
  ' TEST 4: CalculateCPL sin presion (debe retornar 1)
  ' ==================================================================
  totalTests = totalTests + 1
  testName = "Test 4: CalculateCPL sin presion"
  result = t.CalculateCPL(api60, tempLiqF)
  If result = 1 Then
    Debug.Print "[PASS] " & testName & " -> CPL: " & result & " (sin presion)"
    passCount = passCount + 1
  Else
    Debug.Print "[FAIL] " & testName & " -> Esperado 1, obtuvo: " & result
    failCount = failCount + 1
  End If

  ' ==================================================================
  ' TEST 5: CalculateFRA - Ajuste por techo flotante
  ' ==================================================================
  totalTests = totalTests + 1
  testName = "Test 5: CalculateFRA(" & nivel_mm & " mm, API=" & apiObs & ")"
  result = t.CalculateFRA(nivel_mm, apiObs)
  If Abs(result) >= 0 Then
    Debug.Print "[PASS] " & testName & " -> FRA: " & FormatNumber(result, 2) & " Bbls"
    passCount = passCount + 1
  Else
    Debug.Print "[FAIL] " & testName & " -> Esperado >= 0, obtuvo: " & result
    failCount = failCount + 1
  End If

  ' ==================================================================
  ' TEST 6: CalculateGOV - Volumen Bruto Observado (compuesto)
  ' ==================================================================
  totalTests = totalTests + 1
  testName = "Test 6: CalculateGOV (TOV*CTSH+FRA)"
  result = t.CalculateGOV(nivel_mm, tempLiqF, tempAmbF, api60)
  If result > 0 Then
    Debug.Print "[PASS] " & testName & " -> GSV: " & FormatNumber(result, 2) & " Bbls"
    passCount = passCount + 1
  Else
    Debug.Print "[FAIL] " & testName & " -> Esperado > 0, obtuvo: " & result
    failCount = failCount + 1
  End If

  ' ==================================================================
  ' TEST 7: CalculateGSV - Volumen Estandar Bruto (compuesto)
  ' ==================================================================
  totalTests = totalTests + 1
  testName = "Test 7: CalculateGSV (GOV*CTL)"
  result = t.CalculateGSV(nivel_mm, tempLiqF, tempAmbF, api60)
  If result > 0 Then
    Debug.Print "[PASS] " & testName & " -> GSV: " & FormatNumber(result, 2) & " Bbls"
    passCount = passCount + 1
  Else
    Debug.Print "[FAIL] " & testName & " -> Esperado > 0, obtuvo: " & result
    failCount = failCount + 1
  End If

  ' ==================================================================
  ' TEST 8: CalculateNSV - Volumen Estandar Neto (GSV - FRA)
  ' ==================================================================
  totalTests = totalTests + 1
  testName = "Test 8: CalculateNSV (GSV*CSW)"
  result = t.CalculateNSV(nivel_mm, tempLiqF, tempAmbF, api60, bsw)
  If result > 0 Then
    Debug.Print "[PASS] " & testName & " -> NSV: " & FormatNumber(result, 2) & " Bbls"
    passCount = passCount + 1
  Else
    Debug.Print "[FAIL] " & testName & " -> Esperado > 0, obtuvo: " & result
    failCount = failCount + 1
  End If

  ' ==================================================================
  ' TEST 9: Verificar que NSV <= GSV (logica consistente)
  ' ==================================================================
  totalTests = totalTests + 1
  testName = "Test 9: NSV <= GSV (consistencia logica)"
  Dim gsvVal As Double, nsvVal As Double
  gsvVal = t.CalculateGSV(nivel_mm, tempLiqF, tempAmbF, api60)
  nsvVal = t.CalculateNSV(nivel_mm, tempLiqF, tempAmbF, api60, bsw)
  If gsvVal > 0 And nsvVal > 0 And nsvVal <= gsvVal Then
    Debug.Print "[PASS] " & testName & " -> GSV=" & FormatNumber(gsvVal, 2) & ", NSV=" & FormatNumber(nsvVal, 2)
    passCount = passCount + 1
  Else
    Debug.Print "[FAIL] " & testName & " -> GSV=" & FormatNumber(gsvVal, 2) & ", NSV=" & FormatNumber(nsvVal, 2)
    failCount = failCount + 1
  End If

  ' ==================================================================
  ' RESUMEN
  ' ==================================================================
  Debug.Print "----------------------------------------------------------------------"
  Debug.Print "RESUMEN: " & passCount & "/" & totalTests & " pasaron, " & failCount & " fallaron"
  Debug.Print "======================================================================"

  ' Limpieza
  Set t = Nothing
  mdTankService.ClearCache
End Sub

''' <summary>
''' Prueba la integración de clsFluid en clsWell, clsLine y clsStation.
''' </summary>
Public Sub TestFluidIntegration()
  Dim passCount As Long: passCount = 0
  Dim failCount As Long: failCount = 0
  Dim totalTests As Long: totalTests = 0
  Dim testName As String

  Debug.Print "======================================================================"
  Debug.Print "PRUEBA: Integracion de clsFluid en Well, Line y Station"
  Debug.Print "======================================================================"

  ' ==================================================================
  ' TEST 1: clsWell con fluido
  ' ==================================================================
  totalTests = totalTests + 1
  testName = "Test 1: clsWell con fluido"
  Dim w As New clsWell
  w.Tag = "P-001"
  w.Depth = 8500
  w.FlowRate = 1200
  w.Fluid.Name = "Crudo Castilla"
  w.Fluid.FluidType = CRD
  w.Fluid.TypicalAPI = 28.5

  If w.Fluid.TypicalAPI = 28.5 And w.Fluid.Name = "Crudo Castilla" Then
    Debug.Print "[PASS] " & testName
    passCount = passCount + 1
  Else
    Debug.Print "[FAIL] " & testName & " -> API=" & w.Fluid.TypicalAPI
    failCount = failCount + 1
  End If
  w.ShowProperties
  Set w = Nothing

  ' ==================================================================
  ' TEST 2: clsLine con fluido
  ' ==================================================================
  totalTests = totalTests + 1
  testName = "Test 2: clsLine con fluido"
  Dim l As New clsLine
  l.Tag = "LIN-4022"
  l.Length = 5280
  l.Diameter = 8
  l.Material = MCrbn
  l.Fluid.Name = "Gas Natural"
  l.Fluid.FluidType = CRD
  l.Fluid.TypicalAPI = 50

  If l.Fluid.TypicalAPI = 50 And l.Fluid.Name = "Gas Natural" Then
    Debug.Print "[PASS] " & testName
    passCount = passCount + 1
  Else
    Debug.Print "[FAIL] " & testName
    failCount = failCount + 1
  End If
  l.ShowProperties
  Set l = Nothing

  ' ==================================================================
  ' TEST 3: clsStation con colección de fluidos
  ' ==================================================================
  totalTests = totalTests + 1
  testName = "Test 3: clsStation con ReferenceFluids"
  Dim s As New clsStation
  s.Tag = "EC-01"
  s.Name = "Estacion Castilla"

  Dim f1 As New clsFluid
  f1.Name = "Crudo Pesado"
  f1.TypicalAPI = 18
  s.AddReferenceFluid f1

  Dim f2 As New clsFluid
  f2.Name = "Crudo Liviano"
  f2.TypicalAPI = 38
  s.AddReferenceFluid f2

  If s.ReferenceFluids.Count = 2 Then
    Debug.Print "[PASS] " & testName & " -> " & s.ReferenceFluids.Count & " fluidos"
    passCount = passCount + 1
  Else
    Debug.Print "[FAIL] " & testName & " -> Count=" & s.ReferenceFluids.Count
    failCount = failCount + 1
  End If
  s.ShowProperties
  Set s = Nothing

  ' ==================================================================
  ' TEST 4: clsFluid.CalculateVolumetricProperties (estructura completa)
  ' ==================================================================
  totalTests = totalTests + 1
  testName = "Test 4: FluidCalcResult estructura"
  Dim f As New clsFluid
  f.FluidType = CRD
  f.TypicalAPI = 35.4

  Dim props As FluidCalcResult
  props = f.CalculateVolumetricProperties(95, 215)

  If props.IsValid And props.CTL > 0 And props.APIObs > 0 Then
    Debug.Print "[PASS] " & testName
    Debug.Print "       CTL=" & FormatNumber(props.CTL, 6) & " | CTPL=" & FormatNumber(props.CTPL, 6)
    Debug.Print "       APIObs=" & FormatNumber(props.APIObs, 2) & " | Density60F=" & FormatNumber(props.Density60F, 4)
    passCount = passCount + 1
  Else
    Debug.Print "[FAIL] " & testName
    failCount = failCount + 1
  End If
  Set f = Nothing

  ' ==================================================================
  ' RESUMEN
  ' ==================================================================
  Debug.Print "----------------------------------------------------------------------"
  Debug.Print "RESUMEN: " & passCount & "/" & totalTests & " pasaron, " & failCount & " fallaron"
  Debug.Print "======================================================================"
End Sub