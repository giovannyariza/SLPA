Attribute VB_Name = "mdHelpers"
Option Explicit

' ---------------------------------------------------------------------------------------------------------
' MÓDULO AUXILIAR: mdHelpers
' DESCRIPCIÓN:
' Contiene funciones utilitarias y algorítmicas de soporte matemático y de validación para
' garantizar la integridad de los datos en los módulos de cálculo.
' ---------------------------------------------------------------------------------------------------------

''' <summary>
''' Verifica si un número de punto flotante (Double) es finito.
''' Un número finito no es infinito positivo o negativo, ni es una indeterminacion matematica (NaN).
''' </summary>
''' <param name="Value">El valor numérico de tipo Double a evaluar.</param>
''' <returns>True si el número es completamente computable; False si es NaN o Infinito.</returns>

Public Function IsFinite(ByVal Value As Double) As Boolean
  ' Manejo estructural de errores para evitar caídas imprevistas por desbordamientos de memoria de bajo nivel
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
''' Función extendida corporativa: Valida si una dimensión fisica (diametro, longitud, etc.)
''' es un número finito, real y estrictamente mayor que cero.
''' </summary>
''' <param name="Value">El valor físico a evaluar.</param>
''' <returns>True si la dimensión es matemáticamente válida para cálculos de ingeniería.</returns>

Public Function IsValidDimension(ByVal Value As Double) As Boolean
  On Error GoTo ErrHandler

  ' Combina la validación de finitud matemática con la restricción física del negocio.
  If IsFinite(Value) Then
    IsValidDimension = (Value > 0)
  Else
    IsValidDimension = False
  End If

  Exit Function

ErrHandler:
  IsValidDimension = False
End Function

''' <summary>
''' Verifica si un valor se encuentra dentro de un rango cerrado [Min, Max].
''' </summary>
''' <param name="Value">Valor a evaluar.</param>
''' <param name="Min">Límite inferior del rango.</param>
''' <param name="Max">Límite superior del rango.</param>
''' <returns>True si Min <= Value <= Max.</returns>

Public Function IsInRange(ByVal Value As Double, ByVal Min As Double, ByVal Max As Double) As Boolean
  On Error GoTo ErrHandler

  If Not IsFinite(Value) Or Not IsFinite(Min) Or Not IsFinite(Max) Then
    IsInRange = False
    Exit Function
  End If

  IsInRange = (Value >= Min) And (Value <= Max)
  Exit Function

ErrHandler:
  IsInRange = False
End Function

''' <summary>
''' Valida si una gravedad API es físicamente posible ( > -131.5 ) y opcionalmente
''' está dentro del rango normativo definido en mdGlobals para crudo/refinados o lubricantes.
''' </summary>
''' <param name="API">Valor de gravedad API a validar.</param>
''' <param name="TypeLiq">Tipo de liquido (CRD, REF, LUB) para elegir el rango normativo.</param>
''' <param name="CheckNormative">Si True (default), también valida el rango normativo API MPMS.</param>
''' <returns>True si el API es válido; False en caso contrario.</returns>

Public Function IsValidAPI(ByVal API As Double, Optional ByVal TypeLiq As eTypeLiq = CRD, Optional ByVal CheckNormative As Boolean = True) As Boolean
  On Error GoTo ErrHandler

  If Not IsFinite(API) Then
    IsValidAPI = False
    Exit Function
  End If

  ' Límite matemático de la fórmula API
  If API <= -cAPI_B Then
    IsValidAPI = False
    Exit Function
  End If

  ' Rango normativo según tipo de líquido (opcional)
  If CheckNormative Then
    Select Case TypeLiq
      Case CRD, REF
        IsValidAPI = IsInRange(API, cAPI60_RangeCrudeRefined_MIN, cAPI60_RangeCrudeRefined_MAX)
      Case LUB
        IsValidAPI = IsInRange(API, cAPI60_RangeLubricant_MIN, cAPI60_RangeLubricant_MAX)
      Case Else
        IsValidAPI = False
    End Select
  Else
    IsValidAPI = True
  End If

  Exit Function

ErrHandler:
  IsValidAPI = False
End Function

''' <summary>
''' Valida si una temperatura está dentro de los rangos aceptables por la norma API MPMS 11.1.
''' Asume Fahrenheit como unidad de entrada.
''' </summary>
''' <param name="TempF">Temperatura en Fahrenheit a validar.</param>
''' <returns>True si la temperatura está dentro del rango válido.</returns>

Public Function IsValidTemperature(ByVal TempF As Double) As Boolean
  On Error GoTo ErrHandler

  If Not IsFinite(TempF) Then
    IsValidTemperature = False
    Exit Function
  End If

  IsValidTemperature = IsInRange(TempF, cTEMPVALIDRANGE_MIN, cTEMPVALIDRANGE_MAX)
  Exit Function

ErrHandler:
  IsValidTemperature = False
End Function

''' <summary>
''' Convierte un Variant (Rango de Excel o Array) en un arreglo unidimensional de Doubles base 1.
''' Funcion genérica para procesar datos provenientes de hojas de calculo.
''' </summary>
''' <param name="InputVar">Rango de Excel o Variant/Array de entrada.</param>
''' <returns>Arreglo de Double base 1. Arreglo vacío si no hay datos numéricos.</returns>

Public Function ConvertToDoubleArray(ByVal InputVar As Variant) As Double()
  Dim res() As Double
  Dim cell As Variant
  Dim count As Long: count = 0

  ' Determinar tamaño
  If TypeName(InputVar) = "Range" Or IsArray(InputVar) Then
    ' Redimensionar temporalmente
    ReDim res(1 To 10000) ' Límite arbitrario para lotes de fiscalización

    For Each cell In InputVar
      If IsNumeric(cell) And Not IsEmpty(cell) Then
        count = count + 1
        res(count) = CDbl(cell)
      End If
    Next cell

    ' Ajustar al tamaño real
    If count > 0 Then
      ReDim Preserve res(1 To count)
    Else
      ReDim res(0 To 0)
    End If
  End If

  ConvertToDoubleArray = res
End Function

''' <summary>
''' Busca el indice de una columna en un array de headers (1-based, 2D array).
''' Retorna 0 si no se encuentra.
''' </summary>
''' <param name="hdr">Array de headers (1 a N, 1 a M).</param>
''' <param name="name">Nombre de la columna a buscar (case-insensitive, partial match).</param>
''' <returns>Indice de la columna (1-based) o 0 si no se encuentra.</returns>

Public Function FindColumnIndex(ByRef hdr As Variant, ByVal name As String) As Long
  If Not IsArray(hdr) Then
    FindColumnIndex = 0
    Exit Function
  End If
  Dim c As Long
  For c = 1 To UBound(hdr, 2)
    If InStr(1, CStr(hdr(1, c)), name, vbTextCompare) > 0 Then
      FindColumnIndex = c
      Exit Function
    End If
  Next c
  FindColumnIndex = 0
End Function
