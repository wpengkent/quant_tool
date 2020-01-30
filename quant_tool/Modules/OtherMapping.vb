Option Explicit

Public Function ReadSetting_LegA_Index(str_RawCode As String) As String
    ReadSetting_LegA_Index = Split(UCase(str_RawCode), "|")(0)
End Function

Public Function ReadSetting_LegB_Index(str_RawCode As String) As String
    ReadSetting_LegB_Index = Split(UCase(str_RawCode), "|")(1)
End Function

Public Function ReadSetting_LegA_PmtFreq(str_RawCode As String) As String
    ReadSetting_LegA_PmtFreq = Split(UCase(str_RawCode), "|")(2)
End Function

Public Function ReadSetting_LegB_PmtFreq(str_RawCode As String) As String
    ReadSetting_LegB_PmtFreq = Split(UCase(str_RawCode), "|")(3)
End Function

Public Function ReadSetting_LegA(str_RawCode As String) As String
    ReadSetting_LegA = Split(UCase(str_RawCode), "|")(0)
End Function

Public Function ReadSetting_LegB(str_RawCode As String) As String
    ReadSetting_LegB = Split(UCase(str_RawCode), "|")(1)
End Function