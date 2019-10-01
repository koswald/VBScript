
Class MathConstants

    'Property: Pi
    'Returns: 3.14159265358979
    Property Get Pi : Pi = 4 * Atn(1) : End Property

    'Property: DEGRAD
    'Returns Pi/180
    'Remark: Used to convert degrees to radians
    Property Get DEGRAD : DEGRAD = Pi/180 : End Property

    'Property: RADEG
    'Returns: 180/Pi
    'Remark: Used to convert radians to degrees
    Property Get RADEG : RADEG = 180/Pi : End Property

End Class

