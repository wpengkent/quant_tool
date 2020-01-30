Option Explicit

Public Type ApplicationState
    ScreenUpdating As Boolean
    CalculationMode As XlCalculation
    EventsEnabled As Boolean
    StatusBarMsg As Variant
    DisplayAlerts As Boolean
End Type

' ## General sybase connection parameters
Public Type SybaseParams
    IPAddress As String
    Port As Long
    UserID As String
    Password As String
End Type


' ## MUREX QUERY PARAMETERS

' ## Result query
Public Type MxQP_Result
    ResultTable As String
    SystemDate As Long
    ResultType As String
    IsSplitByDate As Boolean
    ScenMin As Integer
    ScenMax As Integer
    TradeSet As Collection
    OutputForm As String
    IncExcl As String
    TradeSetFormula As String
End Type

' ## Rate curve query
Public Type MxQP_RateCurve
    Curve As String
    SystemDate As Long
    DataSet As String
End Type

' ## Security spot market data query
Public Type MxQP_EqSpot
    StartDate As Long
    EndDate As Long
    DataSet As String
End Type

' ## Holiday query
Public Type MxQP_Hols
    Calendar As String
    MinDate As Long
    MaxDate As Long
End Type

' ## Creating dictionaries from mapping tables in sheet
Public Type DictParams
    KeysTopLeft As Range
    ValuesTopLeft As Range
End Type

' ## Contains all information required for business day calculations
Public Type Calendar
    HolDates As Range
    Weekends As String
End Type