Attribute VB_Name = "modDerivedMath"
Option Explicit
'//Author: Microsoft Corporation
'//Source: Derived Math Functions - MSN Library Visual Studio 6.0

Public Function Sec(ByVal X As Variant) As Double
    Sec = 1 / Cos(X)
End Function

Public Function Cosec(ByVal X As Variant) As Double
    Cosec = 1 / Sin(X)
End Function

Public Function Cotan(ByVal X As Variant) As Double
    Cotan = 1 / Tan(X)
End Function

Public Function Arcsin(ByVal X As Variant) As Double
    Arcsin = Atn(X / Sqr(-X * X + 1))
End Function

Public Function Arccos(ByVal X As Variant) As Double
    Arccos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
End Function

Public Function Arcsec(ByVal X As Variant) As Double
    Arcsec = Atn(X / Sqr(X * X - 1)) + Sgn((X) - 1) * (2 * Atn(1))
End Function

Public Function Arccosec(ByVal X As Variant) As Double
    Arccosec = Atn(X / Sqr(X * X - 1)) + (Sgn(X) - 1) * (2 * Atn(1))
End Function

Public Function Arccotan(ByVal X As Variant) As Double
    Arccotan = Atn(X) + 2 * Atn(1)
End Function

Public Function HSin(ByVal X As Variant) As Double
    HSin = (Exp(X) - Exp(-X)) / 2
End Function

Public Function HCos(ByVal X As Variant) As Double
    HCos = (Exp(X) + Exp(-X)) / 2
End Function

Public Function HTan(ByVal X As Variant) As Double
    HTan = (Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X))
End Function

Public Function HSec(ByVal X As Variant) As Double
    HSec = 2 / (Exp(X) + Exp(-X))
End Function

Public Function HCosec(ByVal X As Variant) As Double
    HCosec = 2 / (Exp(X) - Exp(-X))
End Function

Public Function HCotan(ByVal X As Variant) As Double
    HCotan = (Exp(X) + Exp(-X)) / (Exp(X) - Exp(-X))
End Function

Public Function HArcsin(ByVal X As Variant) As Double
    HArcsin = Log(X + Sqr(X * X + 1))
End Function

Public Function HArccos(ByVal X As Variant) As Double
    HArccos = Log(X + Sqr(X * X - 1))
End Function

Public Function HArctan(ByVal X As Variant) As Double
    HArctan = Log((1 + X) / (1 - X)) / 2
End Function

Public Function HArcsec(ByVal X As Variant) As Double
    HArcsec = Log((Sqr(-X * X + 1) + 1) / X)
End Function

Public Function HArccosec(ByVal X As Variant) As Double
    HArccosec = Log((Sgn(X) * Sqr(X * X + 1) + 1) / X)
End Function

Public Function HArccotan(ByVal X As Variant) As Double
    HArccotan = Log((X + 1) / (X - 1)) / 2
End Function

Public Function LogN(ByVal X As Variant, ByVal N As Variant) As Double
    LogN = Log(X) / Log(N)
End Function
