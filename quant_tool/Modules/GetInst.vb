Option Explicit
' ## Functions to return instrument objects based on the specified input parameters

Public Function GetInst_RngAcc(fld_Params As InstParams_RngAcc, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing) As Inst_RngAcc

    Dim rngacc_Output As New Inst_RngAcc
    Call rngacc_Output.Initialize(fld_Params, dic_CurveSet, dic_StaticInfoInput)
    Set GetInst_RngAcc = rngacc_Output
End Function
Public Function GetInst_IRS(fld_Params As InstParams_IRS, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing) As Inst_IRSwap

    Dim irs_Output As New Inst_IRSwap
    Call irs_Output.Initialize(fld_Params, dic_CurveSet, dic_StaticInfoInput)
    Set GetInst_IRS = irs_Output
End Function

Public Function GetInst_CFL(fld_Params As InstParams_CFL, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing) As Inst_CapFloor

    Dim cfl_Output As New Inst_CapFloor
    Call cfl_Output.Initialize(fld_Params, dic_CurveSet, dic_StaticInfoInput)
    Set GetInst_CFL = cfl_Output
End Function

Public Function GetInst_SWT(fld_Params As InstParams_SWT, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing) As Inst_IRSwaption

    Dim swt_Output As New Inst_IRSwaption
    Call swt_Output.Initialize(fld_Params, dic_CurveSet, dic_StaticInfoInput)
    Set GetInst_SWT = swt_Output
End Function

Public Function GetInst_FXF(fld_Params As InstParams_FXF, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing) As Inst_FXFwd

    Dim fxf_Output As New Inst_FXFwd
    Call fxf_Output.Initialize(fld_Params, dic_CurveSet, dic_StaticInfoInput)
    Set GetInst_FXF = fxf_Output
End Function

Public Function GetInst_DEP(fld_Params As InstParams_DEP, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing) As Inst_SimpleDep

    Dim dep_Output As New Inst_SimpleDep
    Call dep_Output.Initialize(fld_Params, dic_CurveSet, dic_StaticInfoInput)
    Set GetInst_DEP = dep_Output
End Function

Public Function GetInst_FRA(fld_Params As InstParams_FRA, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing) As Inst_FRA

    Dim fra_Output As New Inst_FRA
    Call fra_Output.Initialize(fld_Params, dic_CurveSet, dic_StaticInfoInput)
    Set GetInst_FRA = fra_Output
End Function

Public Function GetInst_FVN(fld_Params As InstParams_FVN, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing) As Inst_Opt_FXVan

    Dim fvn_Output As New Inst_Opt_FXVan
    Call fvn_Output.Initialize(fld_Params, dic_CurveSet, dic_StaticInfoInput)
    Set GetInst_FVN = fvn_Output
End Function

Public Function GetInst_BND(fld_Params As InstParams_BND, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing) As Inst_Bond

    Dim bnd_Output As New Inst_Bond
    Call bnd_Output.Initialize(fld_Params, dic_CurveSet, dic_StaticInfoInput)
    Set GetInst_BND = bnd_Output
End Function
Public Function GetInst_NID(fld_Params As InstParams_BND, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing) As Inst_NID  'QJK code 03102014

    Dim bnd_Output As New Inst_NID
    Call bnd_Output.Initialize(fld_Params, dic_CurveSet, dic_StaticInfoInput)
    Set GetInst_NID = bnd_Output
End Function
Public Function GetInst_BA(fld_Params As InstParams_BND, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing) As Inst_BA

    Dim bnd_Output As New Inst_BA
    Call bnd_Output.Initialize(fld_Params, dic_CurveSet, dic_StaticInfoInput)
    Set GetInst_BA = bnd_Output
End Function

Public Function GetInst_FBR(fld_Params As InstParams_FBR, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing) As Inst_Opt_FXBar

    Dim fbr_Output As New Inst_Opt_FXBar
    Call fbr_Output.Initialize(fld_Params, dic_CurveSet, dic_StaticInfoInput)
    Set GetInst_FBR = fbr_Output
End Function

Public Function GetInst_FTB(fld_Params As InstParams_FTB, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing) As Inst_Fut_IRBill

    Dim ftb_Output As New Inst_Fut_IRBill
    Call ftb_Output.Initialize(fld_Params, dic_CurveSet, dic_StaticInfoInput)
    Set GetInst_FTB = ftb_Output
End Function

Public Function GetInst_FBN(fld_Params As InstParams_FBN, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing) As Inst_Fut_Bond

    Dim fbn_Output As New Inst_Fut_Bond
    Call fbn_Output.Initialize(fld_Params, dic_CurveSet, dic_StaticInfoInput)
    Set GetInst_FBN = fbn_Output
End Function

Public Function GetInst_FRE(fld_Params As InstParams_FRE, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing) As Inst_FXRebate

    Dim fre_Output As New Inst_FXRebate
    Call fre_Output.Initialize(fld_Params, dic_CurveSet, dic_StaticInfoInput)
    Set GetInst_FRE = fre_Output
End Function

Public Function GetInst_ECS(fld_Params As InstParams_ECS, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing) As Inst_EqCash

    Dim ecs_Output As New Inst_EqCash
    Call ecs_Output.Initialize(fld_Params, dic_CurveSet, dic_StaticInfoInput)
    Set GetInst_ECS = ecs_Output
End Function

Public Function GetInst_EQO(fld_Params As InstParams_EQO, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing) As Inst_EQOptions

    Dim eqo_Output As New Inst_EQOptions
    Call eqo_Output.Initialize(fld_Params, dic_CurveSet, dic_StaticInfoInput)
    Set GetInst_EQO = eqo_Output
End Function

Public Function GetInst_EQF(fld_Params As InstParams_EQF, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing) As Inst_EQFut

    Dim eqf_Output As New Inst_EQFut
    Call eqf_Output.Initialize(fld_Params, dic_CurveSet, dic_StaticInfoInput)
    Set GetInst_EQF = eqf_Output
End Function

Public Function GetInst_EQS(fld_Params As InstParams_EQS, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing) As Inst_EqSwap

    Dim eqs_Output As New Inst_EqSwap
    Call eqs_Output.Initialize(fld_Params, dic_CurveSet, dic_StaticInfoInput)
    Set GetInst_EQS = eqs_Output
End Function

Public Function GetInst_FXFut(fld_Params As InstParams_FXFut, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing) As Inst_FXFut

    Dim fxfut_Output As New Inst_FXFut
    Call fxfut_Output.Initialize(fld_Params, dic_CurveSet, dic_StaticInfoInput)
    Set GetInst_FXFut = fxfut_Output
End Function