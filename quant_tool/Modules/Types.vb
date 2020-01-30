Option Explicit

' ## Information about where single trade information is stored in the workbook
Public Type SCFParams
    TradeID As String
    Amount As Double
    CCY As String
    PmtDate As Long
    Curve_Disc As String
    Curve_SpotDisc As String
End Type


Public Type BondParams
    TradeID As String
    ValueDate As Long
    Swapstart As Long
    StubType As String
    IsLongCpn As Boolean
    TradeDate As Long
    FirstAccDate As Long
    MatDate As Long
    RollDate As Long
    PaymentSchType As String
    GenerationLimitPoint As Long
    IsFwdGeneration As Boolean
    PExch_Start As Boolean
    PExch_Intermediate As Boolean
    PExch_End As Boolean
    FloatEst As Boolean
    ForceToMV As Boolean
    AmortSchedule As Dictionary
    IsEOM2830 As Boolean
    IsRoundFlow As Boolean
    Notional As Double
    CCY As String
    index As String
    RateOrMargin As Double
    PmtFreq As String
    Daycount As String
    IsUniformPeriods As Boolean
    BDC As String
    EOM As Boolean
    PmtCal As String
    estcal As String
    Curve_Disc As String
    Curve_Est As String
    Fixings As Dictionary
    ModStarts As Dictionary
    DiscSpread As Double
    CalType As String
    BondMarketPrice As Double

    YieldCalc As String
    RateComputingMode As String
    DaycountConv As String
    Periodicity As String
    Yield As Double
    Duration As Double
    MacaulayDuration As Double
    ModifiedDuration As Double
    FixingCurve As String
    Fix_AI As Double
    Fix_IT As Double
    YieldSchedule As String

End Type


Public Type IRLegParams
    TradeID As String
    ValueDate As Long
    Swapstart As Long
    GenerationRefPoint As Long
    GenerationLimitPoint As Long
    IsFwdGeneration As Boolean
    Term As String
    PExch_Start As Boolean
    PExch_Intermediate As Boolean
    PExch_End As Boolean
    FloatEst As Boolean
    ForceToMV As Boolean
    AmortSchedule As Dictionary

    Notional As Double
    CCY As String
    index As String
    RateOrMargin As Double
    PmtFreq As String
    Daycount As String
    IsUniformPeriods As Boolean
    BDC As String
    EOM As Boolean
    PmtCal As String
    estcal As String
    Curve_Disc As String
    Curve_Est As String
    Fixings As Dictionary
    ModStarts As Dictionary
    StubInterpolate As Boolean
    FixInArrears As Boolean
    DisableConvAdj As Boolean
    '#Alvin during IRDigi validation
    IsDigital As Boolean
    '#Alvin during IRDigi validation

    '#Alvin 21/08/2018 Range Accrual validation
    FixedFloat As String
    Rate As Double
    ExoticType As String
    Schedule As String
    NbofDays As String
    ApplyTo As String
    RangeIndex As String
    Correl As Double
    AboveUpper As String
    AboveLower As String
    RangeType As String
    Upper As Double
    Lower As Double
    FirstnLastDay As String
    dic_PeriodStart As Dictionary
    dic_PeriodEnd As Dictionary  '#Alvin 21/08/2018 Range Accrual validation
    PerDayShifter As String    '#Added Alvin 29/08/2018
    GlobalShifter As String
    RateFactor As Double
    Lockout As Variant
    Lockoutmode As String '#Added Alvin 29/08/2018
    FixingsDigi As Dictionary '#Added Alvin 02/10/2018
    ModStartsDigi As Dictionary '#Added Alvin 02/10/2018
    VariableRate As Dictionary '#Added KL 13/02/2019
    VariableRange1 As Dictionary '#Added KL 13/02/2019
    VariableRange2 As Dictionary '#Added KL 13/02/2019
    VariableRange3 As Dictionary '#Added KL 13/02/2019
    VariableRange4 As Dictionary '#Added KL 13/02/2019

End Type

Public Type InstParams_IRS
    TradeID As String
    Pay_LegA As Boolean
    Pay_LegB As Boolean
    CCY_PnL As String
    LegA As IRLegParams
    LegB As IRLegParams
End Type

Public Type InstParams_RngAcc
    TradeID As String
    Pay_LegA As Boolean
    CCY_PnL As String
    LegA As IRLegParams
    LegB As IRLegParams
    LegA2Digi As IRLegParams
    LegB2Digi As IRLegParams  '#Alvin
    IsCallable As Boolean 'KL 201901 for HW1F
    Callable_LegA As Boolean 'KL 201902 for HW1F
    VolCurve As String 'KL 201901 for HW1F
    SpotStep As Integer 'KL 201901 for HW1F
    TimeStep As Integer 'KL 201901 for HW1F
    MeanRev As Double 'KL 201901 for HW1F
    GeneratorA As String 'KL 201901 for HW1F
    GeneratorB As String 'KL 201901 for HW1F
    CallDate As Dictionary 'KL 201901 for HW1F
    Swapstart As Dictionary 'KL 201901 for HW1F

End Type

Public Type InstParams_CFL
    TradeID As String
    Underlying As IRLegParams
    VolCurve As String
    strike As Double
    Direction As OptionDirection
    CCY_PnL As String
    BuySell As String
    Premium As SCFParams
    IsDigital As Boolean
End Type

Public Type InstParams_SWT
    TradeID As String
    ValueDate As Long
    Pay_LegA As Boolean
    CCY_PnL As String
    IsSmile As Boolean
    VolCurve As String
    BuySell As String
    Exercise As String
    OptionMat As Long
    LegA As IRLegParams
    LegB As IRLegParams
    Premium As SCFParams
    SpotStep As Integer 'KL 201812 for HW1F - BSWAP
    TimeStep As Integer 'KL 201812 for HW1F - BSWAP
    MeanRev As Double 'KL 201812 for HW1F - BSWAP
    GeneratorA As String 'KL 201812 for HW1F - BSWAP
    GeneratorB As String 'KL 201812 for HW1F - BSWAP
    CallDate As Dictionary '#Alvin 20181213 for HW1F - BSWAP
    Swapstart As Dictionary '#Alvin 20181213 for HW1F - BSWAP

End Type

Public Type InstParams_FXF
    TradeID As String
    ValueDate As Long
    Pay_FlowA As Boolean
    CCY_PnL As String
    FlowA As SCFParams
    FlowB As SCFParams
End Type

Public Type InstParams_DEP
    TradeID As String
    ValueDate As Long
    StartDate As Long
    MatDate As Long
    IsLoan As Boolean
    Principal As Double
    CCY_Principal As String
    PExch As Boolean
    Rate As Double
    Daycount As String
    CCY_PnL As String
    Curve_Disc As String
End Type

Public Type InstParams_FRA
    TradeID As String
    ValueDate As Long
    StartDate As Long
    IsBuy As Boolean
    Notional As Double
    Generator As String
    Rate As Double
    CCY_PnL As String
    Fixing As Variant
End Type

Public Type InstParams_FVN
    TradeID As String
    ValueDate As Long
    MatDate As Long
    DelivDate As Long
    IsBuy As Boolean
    CCY_Fgn As String
    CCY_Dom As String
    Notional_Fgn As Double
    ExerciseType As String
    Direction As OptionDirection
    IsDigital As Boolean
    strike As Double
    CCY_Payout As String
    QuantoFactor As Double
    CSFactor As Double
    CCY_PnL As String
    IsSmile As Boolean
    IsRescaling As Boolean
    Curve_Disc As String
    Curve_SpotDisc As String
    Premium As SCFParams
End Type

Public Type InstParams_BND
    TradeID As String
    ValueDate As Long
    SpotDays As Integer
    StubType As String
    IsLongCpn As Boolean
    TradeDate As Long
    FirstAccDate As Long
    MatDate As Long
    RollDate As Long
    PaymentSchType As String
    IsFwdGeneration As Boolean
    EOM2830 As Boolean
    IsRoundFlow As Boolean
    IsBuy As Boolean
    Principal As Double
    CCY_Principal As String
    index As String
    RateOrMargin As Double
    PmtFreq As String
    Daycount As String
    BDC As String
    IsUniformPeriods As Boolean
    EOM As Boolean
    CCY_PnL As String
    PmtCal As String
    estcal As String
    Curve_Disc As String
    Curve_Est As String
    Curve_SpotDisc As String
    Fixings As Dictionary
    ModStarts As Dictionary
    AmortSchedule As Dictionary
    DiscSpread As Double
    CalType As String
    BondMarketPrice As Double
    PurchaseCost As SCFParams
    YieldGenerator As String
End Type

Public Type InstParams_FBR
    TradeID As String
    ValueDate As Long
    MatDate As Long
    DelivDate As Long
    IsBuy As Boolean
    CCY_Fgn As String
    CCY_Dom As String
    Notional_Fgn As Double
    OptDirection As OptionDirection
    strike As Double
    IsKnockOut As Boolean
    LowerBar As Double
    UpperBar As Double
    WindowStart As Long
    WindowEnd As Long
    CCY_Payout As String
    QuantoFactor As Double
    CCY_PnL As String
    IsSmile_Orig As Boolean
    IsSmile_IfKnocked As Boolean
    IsRescaling As Boolean
    Curve_Disc As String
    Curve_SpotDisc As String
    Premium As SCFParams
End Type

Public Type InstParams_FTB
    TradeID As String
    ValueDate As Long
    FutMat As Long
    IsBuy As Boolean
    Notional As Double
    CCY_Notional As String
    UndTerm As String
    Daycount As String
    Price_Mkt As Double
    AdjUndMatDate As Long
    SettleDays As Integer
    IsSpreadOn_PnL As Boolean
    IsSpreadOn_DV01 As Boolean
    MeasureMode As String
    CCY_PnL As String
    PmtCal As String
    Curve_Est As String
    Price_Orig As String
End Type

Public Type InstParams_FBN
    TradeID As String
    ValueDate As Long
    FutMat As Long
    UndMat As Long
    SettleDays As Integer
    IsBuy As Boolean
    Notional As Double
    Generator As String
    Coupon As Double
    IsUniformPeriods As Boolean
    BDC_Accrual As String
    Price_Mkt As Double
    ConvFac As Double
    PriceType As String
    IsSpreadOn_PnL As Boolean
    IsSpreadOn_DV01 As Boolean
    CCY_PnL As String
    Price_Orig As String
End Type

Public Type InstParams_FRE
    TradeID As String
    ValueDate As Long
    MatDate As Long
    DelivDate As Long
    IsBuy As Boolean
    CCY_Fgn As String
    CCY_Dom As String
    RebateAmt As Double
    CCY_Rebate As String
    IsKnockOut As Boolean
    LowerBar As Double
    UpperBar As Double
    WindowStart As Long
    WindowEnd As Long
    IsInstantRebate As Boolean
    CCY_PnL As String
    IsSmile As Boolean
    Curve_Disc As String
    Curve_SpotDisc As String
    Premium As SCFParams
End Type

Public Type InstParams_ECS
    TradeID As String
    ValueDate As Long
    SpotDays As Integer
    Security As String
    Quantity As Long
    CCY_Sec As String
    IsBuy As Boolean
    CCY_PnL As String
    SpotCal As String
    Curve_SpotDisc As String
    OutputType As String
    PurchaseCost As SCFParams
End Type

Public Type DateShifterParams
    ShifterName As String
    BaseShifter As DateShifter
    Calendar As String
    DaysToShift As Integer
    IsBusDays As Boolean
    BDC As String
    Algorithm As String
End Type

Public Type InstParams_EQO
    TradeID As String
    ValueDate As Long
    SpotDays As Integer
    MatDate As Long
    DelivDate As Long
    IsFutures As Boolean
    FuturesContract As String
    FutMat_Date As Long
    ExerciseType As String
    IsCall As Boolean
    IsBuy As Boolean
    Quantity As Long
    LotSize As Long
    strike As Double
    Parity As Double
    Security As String
    VolCode As String
    CCY_Sec As String
    CCY_PnL As String
    SpotCal As String
    Curve_SpotDisc As String
    Curve_MarketDisc As String
    Settlement As String
    SettlementDate As Long
    OptionSpot As Integer
    DividendSpot As Integer
    Div_Amount As Double
    DivPayment_Date As Long
    DivEx_Date As Long
    Div_Type As String
    'VolSpread As Double
    VolSpread As Variant '##Matt edit
    PLType As String
    PurchaseCost As SCFParams
    Description As String
    OriValueDate As Long

End Type

Public Type InstParams_EQF
    TradeID As String
    ValueDate As Long
    SpotDays As Integer
    MatDate As Long
    IsBuy As Boolean
    Quantity As Long
    LotSize As Long
    Futures As String
    Fut_ContractPrice As Double
    Fut_MktPrice As Double
    Security As String
    CCY_Sec As String
    CCY_PnL As String
    SpotCal As String
    Curve_Disc As String
    Curve_Div As String
    Div_Yield As Double
    PLType As String
    OriValueDate As Long
End Type

Public Type InstParams_EQS
    TradeID As String
    ValueDate As Long
    Swapstart As Long
    Term As String
    IsCalcEqualFix As Boolean
    IsEqPayer As Boolean
    IsTotalRet As Boolean
    EqSpotDays As Integer
    ConstantType As String
    Quantity As Long
    Notional As Double
    CCY_Fix As String
    CCY_PnL As String
    PLType As String
    CustomSheet As Worksheet
    CustomDate As Dictionary

    Security As String
    CCY_Eq As String
    EqFreq As String
    EqFixCal As String
    EqPmtCal As String
    EqDivType As String
    EqDiv As New Dictionary
    EqFixing As New Dictionary
    FxFixing As New Dictionary
    EqEstCurve As String
    EqDiscCurve As String
    EqBDC As String
    EqIsEOM As Boolean


    CCY_Id As String
    IdSpotDays As Integer
    IdFreq As String
    IdFixCal As String
    IdPmtCal As String
    IdIsFix As Boolean
    IdEstCurve As String
    IdDiscCurve As String
    IdDayCnt As String
    IdBDC As String
    IdIsEOM As Boolean
    RateOrMargin As Double
    IdFixing As New Dictionary
End Type

Public Type InstParams_FXFut
    TradeID As String
    ValueDate As Long
    SpotDays As Integer
    SpotCal As String
    MatDate As Long
    PayCal As String
    IsBuy As Boolean
    Quantity As Long
    LotSize As Long
    LotSizeCCY As String
    Futures As String
    Fut_ContractPrice As Double
    Fut_MktPrice As Double
    Underlying As String
    CCY_PnL As String
    FlowA As SCFParams
    FlowB As SCFParams
End Type