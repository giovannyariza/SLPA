Attribute VB_Name = "mdTestClasses"
Option Explicit

''' <summary>
''' Ejecuta el escenario de pruebas integral para verificar el comportamiento de
''' clsStation, clsComponent, clsTank y el cumplimiento de la interfaz IComponent.
''' </summary>
Public Sub EjecutarPrueba_Infraestructura()
  ' 1. CONFIGURACI�N DE OPTIMIZACI�N Y ENTORNO
  ' Desactivamos la actualizaci�n de pantalla para maximizar la velocidad de procesamiento en entornos corporativos
  Application.ScreenUpdating = False
  
  ' Captura y manejo estructurado de errores en tiempo de ejecuci�n
  On Error GoTo ErrHandler
  
  Debug.Print "======================================================================"
  Debug.Print "INICIANDO PRUEBA DE ARQUITECTURA DE COMPONENTES E INSTALACIONES"
  Debug.Print "======================================================================"
  
  ' 2. DECLARACI�N DE VARIABLES
  Dim miEstacion As clsStation
  Dim componenteGenerico As clsComponent
  Dim tanqueEspecializado As clsTank
  Dim itemComponente As IComponent
  
  ' 3. INSTANCIACI�N Y CONFIGURACI�N DE LA ESTACI�N PRINCIPAL
  Set miEstacion = New clsStation
  With miEstacion
    .Tag = "EC-01"
    .Name = "Estaci�n Central de Recibo"
    .Description = "Planta principal de fiscalizaci�n y distribuci�n de fluidos."
    .Status = "OP"
    .Location = "Bloque Sur - Coordenadas 4.21, -72.3"
  End With
  
  ' 4. CREACI�N DE COMPONENTE GEN�RICO (Uso directo de la clase base)
  Set componenteGenerico = New clsComponent
  With componenteGenerico
    .Tag = "LIN-4022"
    .Description = "L�nea de transferencia de crudo pesado"
    .ComponentType = "L�nea"
    .System = "Transferencia"
    .Service = "Recibo"
    .Status = "OP"
  End With
  
  ' 5. CREACI�N DE COMPONENTE ESPECIALIZADO (Uso de Composici�n y Polimorfismo)
  Set tanqueEspecializado = New clsTank
  ' Propiedades comunes delegadas internamente a clsComponent
  tanqueEspecializado.Tag = "TK-9010"
  tanqueEspecializado.Description = "Tanque de almacenamiento de fluidos con alto contenido de agua"
  tanqueEspecializado.Service = "Almacenamiento"
  tanqueEspecializado.Status = "MT" ' En mantenimiento
  ' Propiedades exclusivas de la clase concreta clsTank
  tanqueEspecializado.Capacity = 15000 ' Capacidad en Barriles (Bbls)
  
  ' 6. ASOCIACI�N BIDIRECCIONAL (Adici�n de componentes a la Estaci�n)
  ' Pasamos las instancias a trav�s del contrato polim�rfico IComponent
  Debug.Print "-> Registrando componentes en la estaci�n " & miEstacion.Tag & "..."
  miEstacion.AddComponent componenteGenerico
  miEstacion.AddComponent tanqueEspecializado
  Debug.Print "-> Registro completado con �xito."
  Debug.Print "----------------------------------------------------------------------"
  
  ' 7. PRUEBA DE ROBUSTEZ: Intento de registro de un duplicado para validar control de errores
  Debug.Print "-> Realizando prueba de tolerancia a fallos (Inserci�n de Tag Duplicado)..."
  Dim componenteDuplicado As clsComponent
  Set componenteDuplicado = New clsComponent
  componenteDuplicado.Tag = "LIN-4022" ' Mismo Tag que el componente gen�rico anterior
  miEstacion.AddComponent componenteDuplicado ' El m�todo interceptar� el error 457 sin romper la ejecuci�n
  Debug.Print "----------------------------------------------------------------------"
  
  ' 8. RECORRIDO POLIM�RFICO DE LA COLECCI�N
  ' Inspeccionamos los componentes que pertenecen a la estaci�n usando la interfaz com�n
  Debug.Print "-> Listando propiedades de los componentes activos de la Estaci�n:"
  For Each itemComponente In miEstacion.Components
    ' Llama al m�todo ShowProperties de cada objeto de forma din�mica (Late Binding controlado por Interfaz)
    itemComponente.ShowProperties
  Next itemComponente
  Debug.Print "----------------------------------------------------------------------"
  
  ' 9. PRUEBA DEL M�TODO CORREGIDO: RemoveComponent
  ' Removemos el componente 'LIN-4022' pasando su Tag identificador
  Debug.Print "-> Eliminando el componente 'LIN-4022' de la colecci�n..."
  miEstacion.RemoveComponent "LIN-4022"
  
  ' Verificamos que se haya eliminado correctamente listando los componentes remanentes
  Debug.Print "-> Componentes remanentes en la estaci�n tras la remoci�n (Esperado: 1): " & miEstacion.Components.Count
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
  ' Restablecemos la actualizaci�n de la pantalla del entorno de Excel
  Application.ScreenUpdating = True
  Exit Sub

ErrHandler:
  ' Bloque de contingencia para capturar fallos no controlados
  MsgBox "Ocurri� un error inesperado en el m�dulo de pruebas." & vbCrLf & _
         "N�mero: " & Err.Number & vbCrLf & _
         "Descripci�n: " & Err.Description, vbCritical, "Error de Sistema - Arquitecto de Soluciones"
  Resume FinSub
End Sub





Sub CalcularInventarioActual()
    Dim ATK7210 As New clsTank
    
    ' Configuración basada en tu imagen
    With ATK7210
        .Tag = "ATK7210"
        .RoofWeight = 3200
        .IsTableNetOfRoof = False ' Tabla contempla el ajuste
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