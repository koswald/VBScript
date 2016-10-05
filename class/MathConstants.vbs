
Class MathConstants

    'Property: Pi
    'Returns: 3.14159...

    Property Get Pi : Pi = 4 * Atn(1) : End Property '3.1415926535897932384626433832795

    'Property: DEGRAD
    'Returns pi/180
    'Remark: Used to convert degrees to radians

    Property Get DEGRAD : DEGRAD = pi/180 : End Property

    'Property: RADEG
    'Returns: 180/pi
    'Remark: Used to convert radians to degrees

    Property Get RADEG : RADEG = 180/pi : End Property

End Class