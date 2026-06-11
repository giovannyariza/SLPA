Attribute VB_Name = "mdHelpers"
Option Explicit

' ---------------------------------------------------------------------------------------------------------
' MODULO AUXILIAR: mdHelpers
' DESCRIPCION:
' Contiene funciones utilitarias y algoritmicas de soporte matematico y de validacion para 
' garantizar la integridad de los datos en los modulos de calculo.
' ---------------------------------------------------------------------------------------------------------

''' <summary>
''' Verifica si un numero de punto flotante (Double) es finito.
''' Un numero finito no es infinito positivo o negativo, ni es una indeterminacion matematica (NaN).
''' </summary>
''' <param name="Value">El valor numerico de tipo Double a evaluar.</param>
''' <returns>True si el numero es completamente computable; False si es NaN o Infinito.</returns>

Public Function IsFinite(ByVal Value As Double) As Boolean
  ' Manejo estructural de errores para evitar caidas imprevistas por desbordamientos de memoria de bajo nivel
  On Error GoTo ErrHandler

  ' Logica matematica optima y de alta velocidad:
  ' 1. (Value = Value) evalua False unicamente si el valor es NaN (Not a Number).
  ' 2. ((Value * 0) = 0) evalua False si el valor original es Infinito positivo o negativo.
  IsFinite = (Value = Value) And ((Value * 0) = 0)
  
  Exit Function

ErrHandler:
  ' En caso de un desbordamiento extremo (Overflow Error 6), se asume con seguridad que no es finito
  IsFinite = False
End Function

''' <summary>
''' Funcion extendida corporativa: Valida si una dimension fisica (diámetro, longitud, etc.)
''' es un numero finito, real y estrictamente mayor que cero.
''' </summary>
''' <param name="Value">El valor fisico a evaluar.</param>
''' <returns>True si la dimension es matemáticamente válida para cálculos de ingenieria.</returns>

Public Function IsValidDimension(ByVal Value As Double) As Boolean
  On Error GoTo ErrHandler

  ' Combina la validacion de finitud matemática con la restriccion fisica del negocio
  If IsFinite(Value) Then
    IsValidDimension = (Value > 0)
  Else
    IsValidDimension = False
  End If
  
  Exit Function

ErrHandler:
  IsValidDimension = False
End Function