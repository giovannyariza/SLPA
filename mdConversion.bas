Attribute VB_Name = "mdConversion"
Option Explicit

' ---------------------------------------------------------------------------------------------------------
' MÓDULO CENTRAL: mdConversion
' DESCRIPCIÓN:
' Implementa los algoritmos de conversión de unidades operacionales y los coeficientes de
' corrección térmica de fluidos bajo el estándar API MPMS Sección 11.1.
' ---------------------------------------------------------------------------------------------------------

''' <summary>
''' Realiza la conversión de unidades de presión (PSI, Bares, Kilopascales) bajo factores de conversión API.
''' </summary>
''' <param name="Pressure">Valor numérico de la presión a convertir.</param>
''' <param name="ConvPrs">Tipo de conversión solicitada (Enum de mdGlobals).</param>
''' <returns>Valor de la presión en la nueva unidad. Retorna 0 si ocurre un error.</returns>

Public Function CONVPRES(ByVal Pressure As Double, ConvPrs As eConvPrs) As Double
  On Error GoTo ErrHandler

  ' API MPMS 11.1 (Pag 22)
  Const A As Double = 6.8947590868
  Const B As Double = 100
  Const C As Double = 6.8947590868E-2
  
  Select Case ConvPrs
    Case P2B  ' PSI a Bar.
      CONVPRES = Pressure * C
    Case P2K  ' PSI a KPa.
      CONVPRES = Pressure * A
    Case B2P  ' Bar a PSI.
      CONVPRES = Pressure / C
    Case B2K  ' Bar a KPa.
      CONVPRES = Pressure * B
    Case K2P  ' KPa a PSI.
      CONVPRES = Pressure / A
    Case K2B  ' KPa a Bar.
      CONVPRES = Pressure / B
    Case Else
      CONVPRES = 0
  End Select
  Exit Function

ErrHandler:
  Debug.Print "Error en CONVPRES: " & Err.Description
  CONVPRES = 0 ' Retorno de seguridad sin correccion
End Function

''' <summary>
''' Realiza la conversión de unidades de temperatura (Celsius, Fahrenheit, Kelvin, Rankine).
''' </summary>
''' <param name="Temperature">Valor numérico de la temperatura base.</param>
''' <param name="ConvTmp">Tipo de conversión térmica solicitada (Enum de mdGlobals).</param>
''' <returns>Valor térmico convertido.</returns>

Public Function CONVTEMP(ByVal Temperature As Double, ConvTmp As eConvTmp) As Double
  On Error GoTo ErrHandler

  ' API MPMS 11.1 (Pag 22)
  Const A As Double = 273.15
  Const B As Double = 491.67
  Const C As Double = 459.67
  Const D As Double = 9 / 5
  Const E As Double = 5 / 9
  Const F As Double = 32

  Select Case ConvTmp
    Case C2F  ' Celsius a Fahrenheit.
      CONVTEMP = Temperature * (D) + F
    Case C2K  ' Celsius a Kelvin.
      CONVTEMP = Temperature + A
    Case C2R  ' Celsius a Rankine.
      CONVTEMP = Temperature * (D) + B
    Case F2C  ' Fahrenheit a Celsius.
      CONVTEMP = (Temperature - F) * (E)
    Case F2K  ' Fahrenheit a Kelvin.
      CONVTEMP = (Temperature - F) * (E) + A
    Case F2R  ' Fahrenheit a Rankine.
      CONVTEMP = Temperature + C
    Case K2C  ' Kelvin a Celsius.
      CONVTEMP = Temperature - A
    Case K2F  ' Kelvin a Fahrenheit.
      CONVTEMP = (Temperature - A) * (D) + F
    Case K2R  ' Kelvin a Rankine.
      CONVTEMP = Temperature * (D)
    Case R2C  ' Rankine a Celsius.
      CONVTEMP = (Temperature - B) * (E)
    Case R2F  ' Rankine a Fahrenheit.
      CONVTEMP = Temperature - C
    Case R2K  ' Rankine a Kelvin.
      CONVTEMP = Temperature * (E)
    Case Else
      CONVTEMP = 0
  End Select
  Exit Function

ErrHandler:
  Debug.Print "Error en CONVTEMP: " & Err.Description
  CONVTEMP = 0 ' Retorno de seguridad sin correccion
End Function

''' <summary>
''' Realiza la conversión de unidades de densidad entre Grados API, Gravedad Específica (SGU) 
''' y Densidad Absoluta (Kg/m³), aplicando las ecuaciones de calibración del estándar API MPMS Capítulo 11.1.
''' </summary>
''' <param name="Density">Valor numérico de la densidad base a convertir.</param>
''' <param name="ConvDns">Tipo de conversión de densidad solicitada (Miembro de la enumeración eConvDns en mdGlobals).</param>
''' <param name="WaterRel">Opcional (Por defecto True). Si es True, utiliza la densidad del agua calibrada a 60°F (cWtrDensKgM3_60F). Si es False, utiliza la base teórica de 1000 Kg/m³.</param>
''' <returns>Valor de la densidad convertido a la nueva unidad física. Retorna 0 si ocurre una división por cero o un dato de entrada inválido.</returns>
''' <remarks>
''' Requisito de Arquitectura: Requiere que la constante corporativa [cWtrDensKgM3_60F] esté declarada de forma global en mdGlobals. 
''' El método incluye protección implícita contra indeterminaciones matemáticas en los límites físicos del API (-131.5).
''' </remarks>

Public Function CONVDENS(ByVal Density As Double, ConvDns As eConvDns, Optional WaterRel As Boolean = True) As Double
  On Error GoTo ErrHandler

  ' Constantes de Conversión de Densidad API
  Const A As Double = 141.5
  Const B As Double = 131.5
  Const WDK60 As Double = 999.016 ' Densidad del agua a 60 grados Fahrenheit (15.56 °C) en Kg/m³.

  Dim WDK As Double, WDS As Double
  
  ' Valores base por defecto (Agua pura teórica)
  WDK = 1000 ' Densidad del agua en Kg/m3
  WDS = 1 ' Densidad del agua en SGU
  
  ' Ajuste hidrodinámico según API MPMS Capítulo 11.1 (Pág. 212)
  If WaterRel = True Then
    WDK = WDK60 ' Densidad del agua en Kg/m3
  End If
  
  Select Case ConvDns
    Case A2S  ' API a SGU
      CONVDENS = (A / (Density + B)) * WDS
    Case A2K  ' API a Kg/m3
      CONVDENS = (A / (Density + B)) * WDK
    Case S2A  ' SGU a API
      CONVDENS = (A * WDS / Density) - B
    Case S2K  ' SGU a Kg/m3
      CONVDENS = Density * WDK
    Case K2A  ' Kg/m3 a API
      CONVDENS = A / (Density / WDK) - B
    Case K2S  ' Kg/m3 a SGU
      CONVDENS = Density / WDK
    Case Else
      CONVDENS = 0
  End Select
  Exit Function

ErrHandler:
  Debug.Print "Error crítico en CONVDENS [Parámetros Inválidos]: " & Err.Description
  CONVDENS = 0 ' Retorno de seguridad sin correccion
End Function

''' <summary>
''' Convierte un valor de temperatura desde la escala moderna ITS-90 (International Temperature Scale of 1990) 
''' hacia la escala previa IPTS-68 (International Practical Temperature Scale of 1968), requerida mandatoriamente 
''' para los algoritmos de cálculo de densidad y volumen del estándar API MPMS Capítulo 11.1.
''' </summary>
''' <param name="Tmp90">Valor de la temperatura en la escala base ITS-90.</param>
''' <param name="TmpUnits">Opcional (Por defecto F). Unidad física de la temperatura ingresada (Miembro de la enumeración eTmpUnits en mdGlobals, ej: F o C).</param>
''' <returns>Valor numérico de la temperatura corregida en la escala IPTS-68, expresada en la misma unidad de entrada. Retorna 0 si ocurre una excepción.</returns>
''' <remarks>
''' El estándar API MPMS Capítulo 11.1 se formuló originalmente utilizando la escala IPTS-68. 
''' Dado que los termómetros industriales modernos reportan en ITS-90, esta conversión de transiciones de escala 
''' es un paso intermedio crítico y obligatorio antes de calcular los coeficientes térmicos K0 y K1.
''' </remarks>

Public Function CONVTEMP68(ByVal Tmp90 As Double, Optional TmpUnits As eTmpUnits = F)
  On Error GoTo ErrHandler

  Dim Tao As Double, Tmp68 As Double, DeltaTmp As Double
  Dim A(1 To 8) As Double
  Dim i As Byte
  
  ' NORMALIZACIÓN TÉRMICA: Convertir de Fahrenheit a Celsius
  If TmpUnits = F Then
    Tmp90 = CONVTEMP(Tmp90, F2C)
  End If
  
  ' Calculo de la temperatura escalada (Tao)
  ' Parámetro de normalización adimensional según API MPMS Capítulo 11.1 (Apéndice A, Pág. 210)
  Tao = Tmp90 / 630
  
  ' Coeficientes oficiales del polinomio de ajuste API / IPTS-68
  A(1) = -0.148759: A(2) = -0.267408: A(3) = 1.08076:  A(4) = 1.269056
  A(5) = -4.089591: A(6) = -1.871251: A(7) = 7.438081: A(8) = -3.536296
  
  ' Acumulacion asintotica por serie de potencias
  DeltaTmp = 0
  For i = 1 To 8
    DeltaTmp = DeltaTmp + A(i) * (Tao ^ i)
  Next i
  
  ' Aplicacion de la correccion a escala
  Tmp68 = Tmp90 - DeltaTmp
  
  ' DESNORMALIZACIÓN: Retorna el resultado en la misma unidad fisica de entrada
  If TmpUnits = F Then
    CONVTEMP68 = CONVTEMP(Tmp68, C2F)
  Else
    CONVTEMP68 = Tmp68
  End If
  Exit Function

ErrHandler:
  Debug.Print "Error crítico en algoritmo CONVTEMP68: " & Err.Description
  CONVTEMP68 = 0 ' Retorno de seguridad sin correccion
End Function

''' <summary>
''' Ajusta la densidad base de un fluido convertida de grados API a 60°F (calculada originalmente en la escala térmica moderna ITS-90) 
''' hacia la escala práctica previa IPTS-68, aplicando el algoritmo iterativo oficial del estándar API MPMS Capítulo 11.1.
''' </summary>
''' <param name="API60">Valor numérico de la gravedad API observada o corregida a 60°F.</param>
''' <param name="TypeLiq">Opcional (Por defecto CRD). Tipo de hidrocarburo bajo evaluación (Miembro de eTypeLiq en mdGlobals: CRD, REF, LUB).</param>
''' <param name="Alfa60">Opcional (Por defecto 0). Coeficiente térmico de expansión a 60°F. Si se provee un valor diferente de cero, el algoritmo omitirá la búsqueda matricial de K0, K1 y K2.</param>
''' <returns>Densidad absoluta ajustada en la escala IPTS-68 expresada en unidades de masa/volumen [Kg/m³]. Retorna 0 si ocurre un fallo matemático.</returns>
''' <remarks>
''' El manual de estándares de medición de petróleo de la API exige este ajuste debido a que las tablas base de 1980 
''' se definieron sobre IPTS-68, mientras que las calibraciones de laboratorio actuales se realizan bajo ITS-90.
''' </remarks>

Function CONVDENS68(ByVal API60 As Double, Optional TypeLiq As eTypeLiq = CRD, Optional ByVal Alfa60 As Double) As Double
  ' Control de errores estructurado para mitigar divisiones por cero accidentales en celdas vacías o datos corruptos
  On Error GoTo ErrHandler

  Dim RngTbl(1 To 6, 1 To 5) As Double
  Dim LimInf As Double, LimSup As Double
  
  Dim K0 As Double, K1 As Double, K2 As Double
  Dim A As Double, B As Double, Rho60 As Double
  Dim i As Byte
  
  ' Conversion de la Gravedad API60 a Densidad en Kg/m3 Relativa a la Densidad del Agua a 60 F.
  Rho60 = CONVDENS(API60, A2K, True)
  
  If Alfa60 = 0 Then
    ' Matriz de Rangos de densidad y coeficientes API Capítulo 11.1.
    RngTbl(1, 1) = 610.6:     RngTbl(1, 2) = 1163.5:    RngTbl(1, 3) = 341.0957:  RngTbl(1, 4) = 0:        RngTbl(1, 5) = 0
    RngTbl(2, 1) = 838.3127:  RngTbl(2, 2) = 1163.5:    RngTbl(2, 3) = 103.872:   RngTbl(2, 4) = 0.2701:   RngTbl(2, 5) = 0
    RngTbl(3, 1) = 787.5195:  RngTbl(3, 2) = 838.3127:  RngTbl(3, 3) = 330.301:   RngTbl(3, 4) = 0:        RngTbl(3, 5) = 0
    RngTbl(4, 1) = 770.352:   RngTbl(4, 2) = 787.5195:  RngTbl(4, 3) = 1489.067:  RngTbl(4, 4) = 0:        RngTbl(4, 5) = -0.0018684
    RngTbl(5, 1) = 610.6:     RngTbl(5, 2) = 770.352:   RngTbl(5, 3) = 192.4571:  RngTbl(5, 4) = 0.2438:   RngTbl(5, 5) = 0
    RngTbl(6, 1) = 800.9:     RngTbl(6, 2) = 1163.5:    RngTbl(6, 3) = 0:         RngTbl(6, 4) = 0.34878:  RngTbl(6, 5) = 0
    
    ' Valores de contingencia iniciales por seguridad analítica
    K0 = 0: K1 = 0: K2 = 0

    ' Seleccion de las constantes dependiendo del tipo de liquido y su densidad.
    Select Case TypeLiq
      Case CRD
        LimInf = RngTbl(1, 1): LimSup = RngTbl(1, 2): K0 = RngTbl(1, 3): K1 = RngTbl(1, 4):   K2 = RngTbl(1, 5)  ' Tabla 6A
      Case REF
        For i = 2 To 5
          If Rho60 >= RngTbl(i, 1) And Rho60 < RngTbl(i, 2) Then
            LimInf = RngTbl(i, 1): LimSup = RngTbl(i, 2): K0 = RngTbl(i, 3):   K1 = RngTbl(i, 4):   K2 = RngTbl(i, 5)
          End If
        Next i  ' Tabla 6B
      Case LUB
        LimInf = RngTbl(6, 1): LimSup = RngTbl(6, 2): K0 = RngTbl(6, 3):   K1 = RngTbl(6, 4):   K2 = RngTbl(6, 5)  ' Tabla 6D
    End Select

    ' Blindaje analítico secundario: Si la densidad no cuadró en ningún rango, se evita el cálculo hidrodinámico
    If K0 = 0 And K1 = 0 And K2 = 0 Then GoTo ErrHandler

    ' Ejecucion de ecuaciones polinomiales de ajuste de escala
    A = (cTEMPSHIFT / 2) * (((K0 / Rho60) + K1) * (1 / Rho60) + K2)    
    B = ((2 * K0) + (K1 * Rho60)) / (K0 + (K1 + (K2 * Rho60) * Rho60))
    
    CONVDENS68 = Rho60 * (1 + ((Exp(A * (1 + 0.8 * A)) - 1) / (1 + A * (1 + 1.6 * A) * B)))
  Else
    ' Ajuste simplificado directo si el usuario inyecta de forma explícita el coeficiente Alfa60
    CONVDENS68 = Rho60 * Exp(0.5 * Alfa60 * cTEMPSHIFT * (1 + 0.4 * Alfa60 * cTEMPSHIFT))
  End If
  Exit Function

ErrHandler:
  Debug.Print "Error crítico detectado en algoritmo CONVDENS68: " & Err.Description
  CONVDENS68 = 0 ' Retorno de seguridad sin correccion
End Function
