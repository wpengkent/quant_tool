Option Explicit

' ## MEMBER DATA
Private wks_Location As Worksheet
Private dcp_FXVolCodes As DictParams, dic_FXVolCodes As Dictionary
Private dcp_DataTypes As DictParams, dic_DataTypes As Dictionary
Private dcp_CcyCalendars As DictParams, dic_CcyCalendars As Dictionary
Private dcp_CcySpotDays As DictParams, dic_CcySpotDays As Dictionary
Private dcp_FXCurveNames As DictParams, dic_FXCurveNames As Dictionary
Private dcp_ThetaModes As DictParams, dic_ThetaModes As Dictionary
Private dcp_SourceTables As DictParams, dic_SourceTables As Dictionary
Private dcp_PillarSets As DictParams, dic_PillarSets As Dictionary


' ## INITIALIZATION
Public Sub Initialize(wks_LocationInput)
    Set wks_Location = wks_LocationInput
    Dim rng_DataTopLeft As Range: Set rng_DataTopLeft = wks_Location.Range("A3")
    Dim int_ColOffset As Integer: int_ColOffset = 0

    ' Define location of top cells for dictionary keys and values
    With rng_DataTopLeft
        Set dcp_FXVolCodes.KeysTopLeft = .Offset(0, int_ColOffset)

        int_ColOffset = int_ColOffset + 1
        Set dcp_FXVolCodes.ValuesTopLeft = .Offset(0, int_ColOffset)

        int_ColOffset = int_ColOffset + 2
        Set dcp_DataTypes.KeysTopLeft = .Offset(0, int_ColOffset)

        int_ColOffset = int_ColOffset + 1
        Set dcp_DataTypes.ValuesTopLeft = .Offset(0, int_ColOffset)

        int_ColOffset = int_ColOffset + 2
        Set dcp_CcyCalendars.KeysTopLeft = .Offset(0, int_ColOffset)
        Set dcp_CcySpotDays.KeysTopLeft = .Offset(0, int_ColOffset)
        Set dcp_FXCurveNames.KeysTopLeft = .Offset(0, int_ColOffset)

        int_ColOffset = int_ColOffset + 1
        Set dcp_CcyCalendars.ValuesTopLeft = .Offset(0, int_ColOffset)

        int_ColOffset = int_ColOffset + 1
        Set dcp_CcySpotDays.ValuesTopLeft = .Offset(0, int_ColOffset)

        int_ColOffset = int_ColOffset + 1
        Set dcp_FXCurveNames.ValuesTopLeft = .Offset(0, int_ColOffset)

        int_ColOffset = int_ColOffset + 2
        Set dcp_PillarSets.KeysTopLeft = .Offset(0, int_ColOffset)

        int_ColOffset = int_ColOffset + 1
        Set dcp_PillarSets.ValuesTopLeft = .Offset(0, int_ColOffset)

        int_ColOffset = int_ColOffset + 2
        Set dcp_ThetaModes.KeysTopLeft = .Offset(0, int_ColOffset)

        int_ColOffset = int_ColOffset + 1
        Set dcp_ThetaModes.ValuesTopLeft = .Offset(0, int_ColOffset)

        int_ColOffset = int_ColOffset + 2
        Set dcp_SourceTables.KeysTopLeft = .Offset(0, int_ColOffset)

        int_ColOffset = int_ColOffset + 1
        Set dcp_SourceTables.ValuesTopLeft = .Offset(0, int_ColOffset)
    End With

    ' Create dictionaries
    Set dic_FXVolCodes = Gather_Dictionary(dcp_FXVolCodes)
    Set dic_DataTypes = Gather_Dictionary(dcp_DataTypes)
    Set dic_CcyCalendars = Gather_Dictionary(dcp_CcyCalendars)
    Set dic_CcySpotDays = Gather_Dictionary(dcp_CcySpotDays)
    Set dic_FXCurveNames = Gather_Dictionary(dcp_FXCurveNames)
    Set dic_ThetaModes = Gather_Dictionary(dcp_ThetaModes)
    Set dic_SourceTables = Gather_Dictionary(dcp_SourceTables)
    Set dic_PillarSets = Gather_Dictionary(dcp_PillarSets)
End Sub


' ## PROPERTIES
Public Property Get Dict_FXVolCodes() As Dictionary
    Set Dict_FXVolCodes = dic_FXVolCodes
End Property

Public Property Get Dict_DataTypes() As Dictionary
    Set Dict_DataTypes = dic_DataTypes
End Property

Public Property Get Dict_CcyCalendars() As Dictionary
    Set Dict_CcyCalendars = dic_CcyCalendars
End Property

Public Property Get Dict_FXCurveNames() As Dictionary
    Set Dict_FXCurveNames = dic_FXCurveNames
End Property

Public Property Get Dict_ThetaModes() As Dictionary
    Set Dict_ThetaModes = dic_ThetaModes
End Property

Public Property Get Dict_SourceTables() As Dictionary
    Set Dict_SourceTables = dic_SourceTables
End Property

Public Property Get Dict_PillarSets() As Dictionary
    Set Dict_PillarSets = dic_PillarSets
End Property


' ## METHODS - LOOKUP
Public Function Lookup_MappedFXVolPair(str_Fgn As String, str_Dom As String) As String
    ' ## Return the vol pair with the correct quotation if a mapping rule is specified, otherwise return the original quotation
    Dim str_Output As String
    Dim str_RawPair As String: str_RawPair = str_Fgn & str_Dom

    If dic_FXVolCodes.Exists(str_RawPair) Then
        str_Output = dic_FXVolCodes(str_RawPair)
    Else
        str_Output = str_RawPair
    End If

    Lookup_MappedFXVolPair = str_Output
End Function

Public Function Lookup_CCYCalendar(str_CCY As String) As String
    ' ## Return the default calendar for the specified currency
    Debug.Assert dic_CcyCalendars.Exists(str_CCY)
    Lookup_CCYCalendar = dic_CcyCalendars(str_CCY)
End Function

Public Function Lookup_CCYSpotDays(str_CCY As String) As Integer
    ' ## Return the number of days to spot for FX rates vs USD
    Debug.Assert dic_CcySpotDays.Exists(str_CCY)
    Lookup_CCYSpotDays = dic_CcySpotDays(str_CCY)
End Function