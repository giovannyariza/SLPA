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

