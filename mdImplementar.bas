Attribute VB_Name = "mdImplementar"
' Manejo de Errores Centralizado: Actualmente usas constantes como vbObjectError + 500. 
' Considera crear una Enum de errores global para que el mantenimiento de los códigos 
' de error sea más sencillo.

' Factory Pattern: Si el proyecto crece, podrías implementar un módulo mdFactory para 
' instanciar componentes. En lugar de hacer Set x = New clsTank, llamarías a 
' CreateComponent("Tank").

' Diccionarios vs Colecciones: Si planeas manejar miles de componentes y necesitas 
' verificar la existencia de un Tag frecuentemente sin generar errores, el objeto Scripting.
' Dictionary es ligeramente más versátil que Collection, aunque este último es más "nativo" 
' en VBA.

' Formulario de gestión para estos componentes.
' Herramienta de importación/exportación de datos vía Power Query?

' Integremos este objeto clsFluid como una propiedad dentro de la clase clsTank para automatizar cálculos de inventario?

' Estrategia Avanzada: Si el rendimiento se vuelve un problema, considera pre-calcular una Matriz de CTL/CPL en una tabla 
' oculta de Excel o en un modelo de Power Pivot y usar interpolación lineal. Sin embargo, para transferencias de custodia 
' legales, tu implementación actual basada en el estándar API MPMS 11.1 es la única aceptable por auditoría.

' Implementar lógica de integración con clsFluid para que un objeto fluido pueda auto-calcular su API60 simplemente 
' pasándole las condiciones observadas?

' Las funciones de conveniencia, donde aplicarlas?
' Las funciones que orquestan el cálculo para un fluido específico deberían estar dentro de la clase clsFluid.
' No definas las funciones de conveniencia en las clases si planeas procesar millones de filas desde Excel. Instanciar 
' un objeto clsFluid por cada fila de una tabla de 100,000 registros matará el rendimiento. Para procesos masivos, usa 
' siempre las funciones del módulo estándar (mdAPIFluidCalcs) directamente o vía Power Query (si es posible traducir la 
' lógica a M).

' Diferencia entre el calculo de la función FRA y el ajuste por volumen entregado en tablas de aforo?

' Integremos estas funciones en un flujo de trabajo práctico (como procesar una tabla de datos observados)?

' Tipo de Lógica	            Ubicación Sugerida	      ¿Por qué?
' Algoritmos API Puros      	mdAPIFluidCalcs	          Reutilización y rigor normativo.
' Lógica de Objeto	          clsFluid	                Encapsula el comportamiento del fluido.
' Orquestación para Excel	    mdAPIService	            Permite usar fórmulas potentes en las celdas.
' Validaciones de Equipo	    clsTank / clsLine	        El equipo sabe sus límites de presión/temperatura.

' Implementemos el módulo de servicio mdAPIService para conectar todo lo que hemos optimizado con la interfaz de usuario de Excel?

' ¿Te gustaría que probemos un escenario específico de "Productos Refinados" (Tabla 6B) para validar la zona de transición, 
' o prefieres que montemos una interfaz en Excel para ver estos resultados en una tabla?

' Integrar toda esta lógica de "Tanques" en la clase clsTank?

' ¿Deseas que movamos ahora estas funciones a un Módulo de Clase para empezar a crear objetos "Tanque" y "Fluido" que utilicen estos cálculos de forma automática?

' ¿Deseas que integremos esta función en un reporte de fiscalización automático o prefieres revisar alguna otra función de la norma API 12.1.1?