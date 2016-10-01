
'Math functions not provided with VBScript

'The native math functions are Sin, Cos, Tan, Atn, Log

Class MathFunctions

    'Property Sec
    'Parameter: Angle in radians
    'Returns Secant
    Function Sec(X)
       Sec = 1 / Cos(X)
    End Function

    'Property Cosec
    'Parameter: Angle in radians
    'Returns Cosecant
    Function Cosec(X)
       Cosec = 1 / Sin(X)
    End Function

    'Property Cotan
    'Parameter: Angle in radians
    'Returns Cotangent
    Function Cotan(X)
       Cotan = 1 / Tan(X)
    End Function

    'Property Arcsin
    'Parameter: A ratio
    'Returns Arcsine
    Function Arcsin(X)
       Arcsin = Atn(X / Sqr(-X * X + 1))
    End Function

    'Property Arccos
    'Parameter: A ratio
    'Returns Inverse Cosine
    Function Arccos(X)
       Arccos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
    End Function

    'Property Arcsec
    'Parameter: A ratio
    'Returns Inverse Secant
    Function Arcsec(X)
       Arcsec = Atn(X / Sqr(X * X - 1)) + Sgn((X) -1) * (2 * Atn(1))
    End Function

    'Property Arccosec
    'Parameter: A ratio
    'Returns Inverse Cosecant
    Function Arccosec(X)
       Arccosec = Atn(X / Sqr(X * X - 1)) + (Sgn(X) - 1) * (2 * Atn(1))
    End Function

    'Property Arccotan
    'Parameter: A ratio
    'Returns Inverse Cotangent
    Function Arccotan(X)
       Arccotan = Atn(X) + 2 * Atn(1)
    End Function

    'Property HSin
    'Parameter: Hyperbolic angle
    'Returns Hyperbolic Sine
    Function HSin(X)
       HSin = (Exp(X) - Exp(-X)) / 2
    End Function

    'Property HCos
    'Parameter: Hyperbolic angle
    'Returns Hyperbolic Cosine
    Function HCos(X)
       HCos = (Exp(X) + Exp(-X)) / 2
    End Function

    'Property HTan
    'Parameter: Hyperbolic angle
    'Returns Hyperbolic Tangent
    Function HTan(X)
       HTan = (Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X))
    End Function

    'Property HSec
    'Parameter: Hyperbolic angle
    'Returns Hyperbolic Secant
    Function HSec(X)
       HSec = 2 / (Exp(X) + Exp(-X))
    End Function

    'Property HCosec
    'Parameter: Hyperbolic angle
    'Returns Hyperbolic Cosecant
    Function HCosec(X)
       HCosec = 2 / (Exp(X) - Exp(-X))
    End Function

    'Property HCotan
    'Parameter: Hyperbolic angle
    'Returns Hyperbolic Cotangent
    Function HCotan(X)
       HCotan = (Exp(X) + Exp(-X)) / (Exp(X) - Exp(-X))
    End Function

    'Property HArcsin
    'Parameter: X
    'Returns Inverse Hyperbolic Sine of X
    Function HArcsin(X)
       HArcsin = Log(X + Sqr(X * X + 1))
    End Function

    'Property HArccos
    'Parameter: X
    'Returns Inverse Hyperbolic Cosine of X
    Function HArccos(X)
       HArccos = Log(X + Sqr(X * X - 1))
    End Function

    'Property HArctan
    'Parameter: X
    'Returns Inverse Hyperbolic Tangent of X
    Function HArctan(X)
       HArctan = Log((1 + X) / (1 - X)) / 2
    End Function

    'Property HArcsec
    'Parameter: X
    'Returns Inverse Hyperbolic Secant of X
    Function HArcsec(X)
       HArcsec = Log((Sqr(-X * X + 1) + 1) / X)
    End Function

    'Property HArccosec
    'Parameter: X
    'Returns Inverse Hyperbolic Cosecant of X
    Function HArccosec(X)
       HArccosec = Log((Sgn(X) * Sqr(X * X + 1) +1) / X)
    End Function

    'Property HArccotan
    'Parameter: X
    'Returns Inverse Hyperbolic Cotangent of X
    Function HArccotan(X)
       HArccotan = Log((X + 1) / (X - 1)) / 2
    End Function

    'Property LogN
    'Parameters: X, N
    'Returns Logarithm of X to base N
    Function LogN(X, N)
       LogN = Log(X) / Log(N)
    End Function

End Class