Attribute VB_Name = "mdTestCalcs"
Option Explicit

Sub Pruebas()
  Dim Api60 As Double: Api60 = 11.5
  Dim Volume As Double: Volume = 36265.55
  Dim Rho60 As Double: Rho60 = CONVDENS(Api60, API, SGU)
  Dim FactorMassBbl As Double: FactorMassBbl =  CONVVOL(1, BBL, MT3) * Rho60
  Dim MassVol As Double: MassVol = Volume * FactorMassBbl

  Debug.Print Round(Rho60, 5)
  Debug.Print Round(FactorMassBbl, 5)
  Debug.Print Round(MassVol, 2) ' En Toneladas

  ' Dim Visc as Double: Visc = 12000
End Sub