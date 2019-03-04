
'Math functions not provided with VBScript

'The native math functions are Sin, Cos, Tan, Atn, Log

'Adapted from the Script56.chm. See also the <a href="https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/3ca8tfek(v%3dvs.84)"> online docs </a>
'
Class MathFunctions

    'Property Sec
    'Parameter: Angle in radians
    'Returns Secant
    'Remarks: Sec = 1 / Cos(X)
    Function Sec(X)
       Sec = 1 / Cos(X)
    End Function

    'Property Cosec
    'Parameter: Angle in radians
    'Returns Cosecant
    'Remarks: Cosec = 1 / Sin(X)
    Function Cosec(X)
       Cosec = 1 / Sin(X)
    End Function

    'Property Cotan
    'Parameter: Angle in radians
    'Returns Cotangent
    'Remarks: Cotan = 1 / Tan(X)
    Function Cotan(X)
       Cotan = 1 / Tan(X)
    End Function

    'Property Arcsin
    'Parameter: A ratio
    'Returns Arcsine
    'Remarks: Arcsin = Atn(X / Sqr(-X * X + 1))
    Function Arcsin(X)
       Arcsin = Atn(X / Sqr(-X * X + 1))
    End Function

    'Property Arccos
    'Parameter: A ratio
    'Returns Inverse Cosine
    'Remarks: Arccos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
    Function Arccos(X)
       Arccos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
    End Function

    'Property Arcsec
    'Parameter: A ratio
    'Returns Inverse Secant
    'Remarks: Arcsec = Atn(X / Sqr(X * X - 1)) + Sgn((X) -1) * (2 * Atn(1))
    Function Arcsec(X)
       Arcsec = Atn(X / Sqr(X * X - 1)) + Sgn((X) -1) * (2 * Atn(1))
    End Function

    'Property Arccosec
    'Parameter: A ratio
    'Returns Inverse Cosecant
    'Remarks: Arccosec = Atn(X / Sqr(X * X - 1)) + (Sgn(X) - 1) * (2 * Atn(1))
    Function Arccosec(X)
       Arccosec = Atn(X / Sqr(X * X - 1)) + (Sgn(X) - 1) * (2 * Atn(1))
    End Function

    'Property Arccotan
    'Parameter: A ratio
    'Returns Inverse Cotangent
    'Remarks: Arccotan = Atn(X) + 2 * Atn(1)
    Function Arccotan(X)
       Arccotan = Atn(X) + 2 * Atn(1)
    End Function

    'Property HSin
    'Parameter: Hyperbolic angle
    'Returns Hyperbolic Sine
     'Remarks: HSin = (Exp(X) - Exp(-X)) / 2
   Function HSin(X)
       HSin = (Exp(X) - Exp(-X)) / 2
    End Function

    'Property HCos
    'Parameter: Hyperbolic angle
    'Returns Hyperbolic Cosine
    'Remarks: HCos = (Exp(X) + Exp(-X)) / 2
    Function HCos(X)
       HCos = (Exp(X) + Exp(-X)) / 2
    End Function

    'Property HTan
    'Parameter: Hyperbolic angle
    'Returns Hyperbolic Tangent
    'Remarks: HTan = (Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X))
    Function HTan(X)
       HTan = (Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X))
    End Function

    'Property HSec
    'Parameter: Hyperbolic angle
    'Returns Hyperbolic Secant
    'Remarks: HSec = 2 / (Exp(X) + Exp(-X))
    Function HSec(X)
       HSec = 2 / (Exp(X) + Exp(-X))
    End Function

    'Property HCosec
    'Parameter: Hyperbolic angle
    'Returns Hyperbolic Cosecant
    'Remarks: HCosec = 2 / (Exp(X) - Exp(-X))
    Function HCosec(X)
       HCosec = 2 / (Exp(X) - Exp(-X))
    End Function

    'Property HCotan
    'Parameter: Hyperbolic angle
    'Returns Hyperbolic Cotangent
    'Remarks: HCotan = (Exp(X) + Exp(-X)) / (Exp(X) - Exp(-X))
    Function HCotan(X)
       HCotan = (Exp(X) + Exp(-X)) / (Exp(X) - Exp(-X))
    End Function

    'Property HArcsin
    'Parameter: X
    'Returns Inverse Hyperbolic Sine of X
    'Remarks: HArcsin = Log(X + Sqr(X * X + 1))
    Function HArcsin(X)
       HArcsin = Log(X + Sqr(X * X + 1))
    End Function

    'Property HArccos
    'Parameter: X
    'Returns Inverse Hyperbolic Cosine of X
    'Remarks: HArccos = Log(X + Sqr(X * X - 1))
    Function HArccos(X)
       HArccos = Log(X + Sqr(X * X - 1))
    End Function

    'Property HArctan
    'Parameter: X
    'Returns Inverse Hyperbolic Tangent of X
    'Remarks: HArctan = Log((1 + X) / (1 - X)) / 2
    Function HArctan(X)
       
    End Function

    'Property HArcsec
    'Parameter: X
    'Remarks: 
    'Returns Inverse Hyperbolic Secant of X
    'Remarks: HArcsec = Log((Sqr(-X * X + 1) + 1) / X)
    Function HArcsec(X)
       HArcsec = Log((Sqr(-X * X + 1) + 1) / X)
    End Function

    'Property HArccosec
    'Parameter: X
    'Returns Inverse Hyperbolic Cosecant of X
    'Remarks: HArccosec = Log((Sgn(X) * Sqr(X * X + 1) +1) / X)
    Function HArccosec(X)
       HArccosec = Log((Sgn(X) * Sqr(X * X + 1) +1) / X)
    End Function

    'Property HArccotan
    'Parameter: X
    'Returns Inverse Hyperbolic Cotangent of X
    'Remarks: HArccotan = Log((X + 1) / (X - 1)) / 2
    Function HArccotan(X)
       HArccotan = Log((X + 1) / (X - 1)) / 2
    End Function

    'Property LogN
    'Parameters: X, N
    'Returns Logarithm of X to base N
    'Remarks: LogN = Log(X) / Log(N)
    Function LogN(X, N)
       LogN = Log(X) / Log(N)
    End Function

End Class

