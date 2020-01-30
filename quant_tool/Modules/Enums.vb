Option Explicit

Public Enum ApplicationStateType
    Current = 1
    Optimized = 2
End Enum

Public Enum InstType
    All = 0
    IRS = 1
    CFL = 2
    SWT = 3
    FXF = 4
    DEP = 5
    FRA = 6
    FVN = 7
    BND = 8
    FBR = 9
    FTB = 10
    FBN = 11
    FRE = 12
    ECS = 13
    EQO = 14
    EQF = 15
    EQS = 16
    BA = 17
    NID = 18
    FXFut = 19 '#Matt
    RngAcc = 20 '#Alvin
End Enum

Public Enum CurveType
    EQSPT = 1
    FXSPT
    FXV
    IRC
    cvl
    SVL
    EQVOL
    EVL '##Matt edit
End Enum

Public Enum BookingAttribute
    ' ## Keys of the instrument booking dictionary
    Sheet = 1
    IDSelection = 2
    Params = 3
    Outputs = 4
    BaseChg = 5
End Enum

Public Enum InstAction
    Select_All = 1
    Select_None = 2
    Rebase = 3
    DefineTarget = 4
    Calc_PnL = 101
    Calc_DV01 = 102
    Calc_DV02 = 103
    Calc_Vega = 104
    Calc_Delta = 105
    Calc_Gamma = 106

    'QJK added 26102016
    Calc_FlatVega = 107

    Display_Flows = 201
    Calc_Yield = 5
End Enum

Public Enum ResultType
    ' ## Must match index of columns in 'Target Trades'
    PnL = 1
    PnLChg = 2
    DV01 = 3
    DV02 = 4
    Vega = 5
End Enum

Public Enum RevalType
    ' ## Used by instrument cache to decide which values to recalculate
    All = 1
    PnL = 2
    DV01 = 3
    DV02 = 4
    Vega = 5
    Yield = 6
    Delta = 7
    Gamma = 8

    'QJK added 26102016
    Flat_Vega = 9
End Enum

Public Enum StaticInfoType
    ConfigSheet = 1
    CalendarSet
    IRGeneratorSet
    YieldGeneratorSet
    DateShifterSet
    IRQuerySet
    MappingRules
    'RatesDB
    'ScenarioDB
End Enum

Public Enum ValType
    PnL = 1
    MV = 2
    Cash = 3
End Enum

Public Enum ShockType
    Absolute = 1
    Relative = 2
End Enum

Public Enum VolPair
    XY = 1  ' Fgn/Dom
    XQ = 2  ' Fgn/Quanto
    YQ = 3  ' Dom/Quanto
End Enum