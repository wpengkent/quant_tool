Option Explicit

Public Enum ApplicationStateType
    Current = 1
    Optimized
End Enum

Public Enum OptionDirection
    CallOpt = 1
    PutOpt = -1
End Enum

Public Enum EuropeanPayoff
    Standard = 1
    Digital_CoN
    Digital_AoN
End Enum

Public Enum NormalCDFMethod
    Excel = 1
    Abram
End Enum

Public Enum InterpAxis
    Keys = 1
    Values
End Enum