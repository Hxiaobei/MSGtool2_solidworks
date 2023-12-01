Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Imports System.Windows.Form
Imports System.IO
Imports System.Xml.Serialization

Module swToosl

    Public Const PI = 3.14159265358979
    Public Errors(5) As String
    '钣金输出 主函数入口
    Function ExportToCAD(Model As ModelDoc2, Quantity As Integer, ByRef DraDoc As DrawingDoc) As Object
        Dim Erorrs(5) As String
        Dim Feat As Feature
        Dim p, ii, nn As Integer, Bip As Byte '只有一个特征的位置
        nn = Model.FeatureManager.GetFeatureCount(True)
        For ii = 0 To nn - 20 Step 1
            Feat = Model.FeatureByPositionReverse(ii)
            If Feat.GetTypeName = "FlatPattern" Then p = p + 1 : If p = 1 Then Bip = ii
            If p > 1 Then Exit For
        Next

        If p = 0 Then Erorrs(4) = "此零件没有钣金实体" : GoTo Ea
        'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx 
#Region "读取零件的自定义属性值"
        'modelDocExtension _ CustomPropertyManager(自定义属性接口)
        'CustomtompropertyManager(特定配置接口)
        Dim Prop(4) As String
        Dim config As Configuration = Model.GetActiveConfiguration
        Dim cusPropMgr As CustomPropertyManager = config.CustomPropertyManager
        Prop(0) = config.Name

        If Quantity = 0 Then
            Dim ResolvedValOut As String
            cusPropMgr.Get5(SzmStr(14), False, Nothing, ResolvedValOut, Nothing)
            If ResolvedValOut = "" Then Erorrs(4) = "请添加数量属性值" : GoTo Ea
            Quantity = ResolvedValOut
        End If

        Prop(1) = Model.MaterialIdName
        If bRet(7) Then
            If Prop(1) = "" And IsNothing(DraDoc) Then Erorrs(4) = "未添加材质" : GoTo Ea
        Else
            If Prop(1) = "" Then Prop(1) = "未指定|"
        End If
        Try
            Dim Mob As Object = Split(Prop(1), "|")
            Prop(1) = Mob(1)
        Catch ex As Exception
            Prop(1) = "未指定"
        End Try

        FacDat(Prop(1))

#End Region

        Dim Thickess As Double
        If p = 1 Then
            Feat = Model.FeatureByPositionReverse(Bip)
            Dim FlatPatt As FlatPatternFeatureData = Feat.GetDefinition
            Try
                If FlatPatt.ShowSlitInCornerRelief = False Then FlatPatt.ShowSlitInCornerRelief = True
                If FlatPatt.MergeFace = False Then FlatPatt.MergeFace = True
                If FlatPatt.SimplifyBends = False Then FlatPatt.SimplifyBends = True
                If FlatPatt.CornerTreatment Then FlatPatt.CornerTreatment = False
            Catch ex As Exception

            End Try

            Feat.ModifyDefinition(FlatPatt, Model, Nothing)
            For ii = 0 To UBound(Mater)
                If Mater(ii) = "" Then Exit For
                If InStr(Prop(1), Mater(ii)) > 0 Then
                    If GetFCr(FlatPatt, Model) = False Then Erorrs(4) = "请确认固定面0,1,1" : GoTo Ea
                End If
            Next

            Dim OFeat = Feat.GetParents
            Dim suFeat As Feature
            For Each suFeat In OFeat
                If suFeat.GetTypeName = "SheetMetal" Then Exit For
            Next

            Thickess = suFeat.GetDefinition.Thickness * 1000 : Thickess = Math.Round(Thickess, 2)
            Erorrs(4) = BendDf(Feat, Model, Thickess) : If Erorrs(4) <> "" Then GoTo Ea
            Prop(2) = Thickess
            Prop(3) = Quantity
            If DraDoc Is Nothing Then
                'If DxfOutD(Feat, Model, Prop) = False Then Erorrs(4) = "展开错误" : GoTo Ea
                If DxfOutDll(Feat, Model, Prop, DraDoc) = False Then Erorrs(4) = "展开错误" : GoTo Ea
            Else
                If DxfOutDll(Feat, Model, Prop, DraDoc) = False Then Erorrs(4) = "展开错误" : GoTo Ea
            End If

            '多实体输出、、、、、、、、、、、、、、、
        ElseIf p > 1 Then
            Dim Ib As Byte = 0
            Dim ForFeat() As Feature, PO As Boolean
            For ii = nn - 5 To 1 Step -1
                Dim Feat5 As Feature = Model.FeatureByPositionReverse(ii)
                Dim ie As Integer
                Dim FeatType As String = Feat5.GetTypeName
                'Dim FeatName As String = Feat5.Name
                Select Case FeatType

                    Case "CutListFolder"
                        ReDim Preserve ForFeat(ie)
                        ForFeat(ie) = Feat5
                        ie = ie + 1
                    Case "SolidBodyFolder"
                        If Not PO Then
                            Feat5.GetSpecificFeature2.UpdateCutList()
                            PO = True
                            ii = nn - 10
                        End If

                    Case "ProfileFeature"
                        If bRet(8) = "1" Or bRet(9) = "1" Then
                            If Feat5.Visible = 2 Then
                                Feat5.Select2(False, 0)
                                Model.BlankSketch()
                            End If
                        End If
                    Case Else
                        If bRet(8) = "1" Or bRet(9) = "1" Then
                            Dim subFeat As Feature = Feat5.GetFirstSubFeature
                            Do Until subFeat Is Nothing
                                If subFeat.GetTypeName = "ProfileFeature" Then
                                    If subFeat.Visible = 2 Then
                                        subFeat.Select2(False, 0)
                                        Model.BlankSketch()
                                    End If
                                End If
                                subFeat = subFeat.GetNextSubFeature
                            Loop
                        End If
                End Select

            Next ii

            If ForFeat Is Nothing Then Erorrs(4) = "请更新切割清单" : GoTo Ea

            For Each suFeat As Feature In ForFeat
                If suFeat Is Nothing Then Exit For
                Dim BodyFolder As BodyFolder = suFeat.GetSpecificFeature2
                If suFeat.GetSpecificFeature2.GetCutListType = 2 Then

                    Dim vBodies As Object = BodyFolder.GetBodies

                    If Not vBodies Is Nothing Then
                        For i = LBound(vBodies) To UBound(vBodies)
                            Dim BodyM As Body2 = vBodies(i)
                            BodyM.Select2(False, Nothing)
                            Model.FeatureManager.ShowBodies()
                            Dim PartM As String = BodyM.GetMaterialIdName2

                            If PartM <> "" Then
                                Dim arrs() As String = Split(PartM, "|")
                                PartM = arrs(1)
                            Else
                                PartM = Prop(1)
                            End If

                            Feat = BodyM.GetFeatures(BodyM.GetFeatureCount - 1)

                            If Feat.GetTypeName2 = "FlatPattern" Then
                                Ib = Ib + 1
                                Dim Aname As String = "-A" & Ib
                                If suFeat.Name <> Aname Then suFeat.Name = "-A" & Ib
                                If Feat.Name <> Aname & "-FP" Then Feat.Name = Aname & "-FP"

                                Dim Cn As Byte = BodyFolder.GetBodyCount
                                Dim Number1 As Integer = Quantity * Cn

                                Dim FlatPatt As FlatPatternFeatureData = Feat.GetDefinition
                                If FlatPatt.ShowSlitInCornerRelief = False Then FlatPatt.ShowSlitInCornerRelief = True
                                If FlatPatt.MergeFace = False Then FlatPatt.MergeFace = True
                                If FlatPatt.SimplifyBends = False Then FlatPatt.SimplifyBends = True
                                If FlatPatt.CornerTreatment Then FlatPatt.CornerTreatment = False
                                Feat.ModifyDefinition(FlatPatt, Model, Nothing)
                                For n = 0 To UBound(Mater)
                                    If Mater(n) = "" Then Exit For
                                    If InStr(PartM, Mater(n)) > 0 Then
                                        If GetFCr(FlatPatt, Model) = False Then Erorrs(4) = "请确认固定面0,1,1" : GoTo Ea
                                    End If
                                Next

                                Dim OFeat As Object = Feat.GetParents
                                Dim sFeat As Feature
                                For Each sFeat In OFeat
                                    If sFeat.GetTypeName = "SheetMetal" Then Exit For
                                Next

                                Thickess = sFeat.GetDefinition.Thickness * 1000 : Thickess = Math.Round(Thickess, 2)
                                Erorrs(4) = BendDf(Feat, Model, Thickess) : If Erorrs(4) <> "" Then GoTo Ea
                                Prop(2) = Thickess
                                Prop(3) = Number1
                                Prop(4) = Aname
                                'Dim OutputPath As String = PathName & Aname & "  " & PartM & "-" & Thickess & "T" & Nstring
                                If DraDoc Is Nothing Then
                                    'If DxfOutD(Feat, Model, Prop) = False Then Erorrs(4) = "展开错误" : GoTo Ea
                                    If DxfOutDll(Feat, Model, Prop, DraDoc) = False Then Erorrs(4) = "展开错误" : GoTo Ea
                                Else
                                    If DxfOutDll(Feat, Model, Prop, DraDoc) = False Then Erorrs(4) = "展开错误" : GoTo Ea
                                End If

                            End If
                        Next
                    End If
                End If
            Next

            Erorrs(5) = "多个平板式，输出后请检查"
        End If

Ea:     ExportToCAD = Erorrs
    End Function

    '系数纠错（全局模式）"
    Function GlobalCorrection(Model As ModelDoc2) As Boolean


        Dim SheetMetal As SheetMetalFeatureData
        Dim FeatCounts(50) As Integer
        Dim Thickess As Double
        Dim P As Byte

        Dim Material As String = Model.MaterialIdName
        Try
            Dim Mob As Object = Split(Material, "|")
            Material = Mob(1)
        Catch ex As Exception

        End Try

        FacDat(Material)

        Dim FeatureMgr As FeatureManager = Model.FeatureManager

        Dim nn As Integer = FeatureMgr.GetFeatureCount(True)
        Dim Feature As Feature
        For ii = nn - 15 To 0 Step -1
            Feature = Model.FeatureByPositionReverse(ii)
            If Feature.GetTypeName = "SheetMetal" Then FeatCounts(P) = ii : P = P + 1
        Next

        Dim ShMeFolder As SheetMetalFolder = FeatureMgr.GetSheetMetalFolder

        If P = 0 Then Exit Function

        If P = 1 And ShMeFolder IsNot Nothing Then

            Dim OFeat As Object = ShMeFolder.GetSheetMetals '获取钣金特征
            Feature = OFeat(0)
            SheetMetal = Feature.GetDefinition
            Thickess = Math.Round(SheetMetal.Thickness, 8)

            Dim bRet As Boolean
            SheetMetal.GetOverrideDefaultParameter2(0, bRet)
            If bRet Then SheetMetal.SetOverrideDefaultParameter2(0, False)

            SheetMetal.GetOverrideDefaultParameter2(1, bRet)
            If bRet Then SheetMetal.SetOverrideDefaultParameter2(1, False)

            SheetMetal.GetOverrideDefaultParameter2(2, bRet)
            If bRet Then SheetMetal.SetOverrideDefaultParameter2(2, False)

            Feature.ModifyDefinition(SheetMetal, Model, Nothing)

            '修改钣金文件夹系数
            Feature = ShMeFolder.GetFeature
            SheetMetal = Feature.GetDefinition
            If BendAllowance(SheetMetal, Thickess * 1000.0#) Then Feature.ModifyDefinition(SheetMetal, Model, Nothing)
        End If

        For ii = 0 To UBound(FeatCounts)
            If FeatCounts(ii) = 0 Then Exit For
            Feature = Model.FeatureByPositionReverse(FeatCounts(ii))

            SheetMetal = Feature.GetDefinition
            Thickess = Math.Round(SheetMetal.Thickness, 8)
            'bRet = SheetMetal.AccessSelections(Model, Nothing)
            If BendAllowance(SheetMetal, Thickess * 1000.0#) Then Feature.ModifyDefinition(SheetMetal, Model, Nothing)
            Dim K(1) As Double
            For i = 0 To UBound(Factor)
                If Factor(i, 0) = 0 Then Exit For
                If Factor(i, 0) = Thickess * 1000.0# Then
                    K(0) = Factor(i, 2)
                    K(1) = Factor(i, 3) / 1000.0#
                    Exit For
                End If
            Next

            Dim Childs As Object = Feature.GetChildren
            For Each Feat As Feature In Childs
                Select Case Feat.GetTypeName
                    '绘制的折弯
                    Case "SM3dBend" : Process_SM3dBend(Feat, Model)
                    '斜接法兰
                    Case "SMMiteredFlange" : Process_SMMiteredFlange(Feat, Model)
                    '边线-法兰
                    Case "EdgeFlange" : Process_EdgeFlange(Feat, Model)
                    '褶边
                    Case "Hem" : Process_Hem(Thickess, Feat, Model)
                    '转折
                    Case "Jog" : Process_Jog(Feat, Model)
                End Select

                ' 加工钣金特征的子特征
                Dim subFeat As Feature = Feat.GetFirstSubFeature
                Do While Not subFeat Is Nothing
                    Dim subFeatName As String = subFeat.GetTypeName
                    Dim bRet As Boolean = subFeatName = "OneBend" Or subFeatName = "SketchBend" Or subFeatName = "ToroidalBend"
                    If bRet Then Process_OneBend(subFeat, Model, Thickess, K)
                    subFeat = subFeat.GetNextSubFeature()
                Loop
            Next
        Next
        Return True
    End Function

#Region "加工特征"

    Private Sub Process_SM3dBend(Feat As Feature, Model As ModelDoc2)
        Dim SketchBend As SketchedBendFeatureData
        SketchBend = Feat.GetDefinition
        SketchBend.UseDefaultBendAllowance = True
        SketchBend.UseDefaultBendRadius = True
        Feat.ModifyDefinition(SketchBend, Model, Nothing)
    End Sub

    Private Sub Process_SMMiteredFlange(Feat As Feature, Model As ModelDoc2)
        Dim bRet1, bRet2 As Boolean
        Dim MiterFlange As MiterFlangeFeatureData
        MiterFlange = Feat.GetDefinition
        bRet1 = MiterFlange.UseDefaultBendAllowance
        bRet2 = MiterFlange.UseDefaultBendRadius
        If bRet1 = False Then MiterFlange.UseDefaultBendAllowance = True
        If bRet2 = False Then MiterFlange.UseDefaultBendRadius = True
        If bRet1 = False Or bRet2 = False Then Feat.ModifyDefinition(MiterFlange, Model, Nothing)
    End Sub

    Private Sub Process_EdgeFlange(Feat As Feature, Model As ModelDoc2)
        Dim bRet1, bRet2 As Boolean
        Dim EdgeFlange As EdgeFlangeFeatureData
        EdgeFlange = Feat.GetDefinition
        bRet1 = EdgeFlange.UseDefaultBendAllowance
        bRet2 = EdgeFlange.UseDefaultBendRadius

        If bRet1 = False Then EdgeFlange.UseDefaultBendAllowance = True
        If bRet2 = False Then EdgeFlange.UseDefaultBendRadius = True

        If bRet1 = False Or bRet2 = False Then Feat.ModifyDefinition(EdgeFlange, Model, Nothing)
    End Sub

    Private Sub Process_Hem(Thickess As Double, Feat As Feature, Model As ModelDoc2)
        If Thickess = 0 Then
            Dim ObFeat = Feat.GetParents
            For Each sFeat As Feature In ObFeat
                If sFeat.GetTypeName2 = "SheetMetal" Then
                    Dim SheetMetal As SheetMetalFeatureData = sFeat.GetDefinition
                    Thickess = SheetMetal.Thickness
                End If
            Next
        End If

        If SzmStr(5) = "1" Then
            For i = 1 To UBound(Factor)
                If Factor(i, 0) = 0 Then Exit For
                If Factor(i, 0) = Thickess * 1000.0# Then
                    Dim Hem As HemFeatureData = Feat.GetDefinition
                    Hem.UseDefaultBendAllowance = False
                    Dim BendAce As CustomBendAllowance = Hem.GetCustomBendAllowance
                    BendAce.Type = 2
                    BendAce.KFactor = Factor(i, 4)
                    Hem.SetCustomBendAllowance(BendAce)
                    Feat.ModifyDefinition(Hem, Model, Nothing)
                    Exit For
                End If
            Next
        End If

    End Sub

    Private Sub Process_Jog(Feat As Feature, Model As ModelDoc2)
        Dim bRet1, bRet2 As Boolean
        Dim Jog As JogFeatureData
        Jog = Feat.GetDefinition
        bRet1 = Jog.UseDefaultBendAllowance
        bRet2 = Jog.UseDefaultBendRadius

        If bRet1 = False Then Jog.UseDefaultBendAllowance = True
        If bRet2 = False Then Jog.UseDefaultBendRadius = True
        If bRet1 = False Or bRet2 = False Then Feat.ModifyDefinition(Jog, Model, Nothing)
    End Sub

    Private Sub Process_OneBend(ByRef Feat As Feature, ByRef Model As ModelDoc2, ByRef T As Double, ByRef K() As Double)
        Dim OneBend As OneBendFeatureData = Feat.GetDefinition
        Dim bRet1 As Boolean = OneBend.UseDefaultBendAllowance
        Dim R As Double = OneBend.BendRadius

        If bRet(0) Then
            Dim Angle As Double = Math.Round(OneBend.BendAngle * 180 / PI, 1)
            If R / T > 1.5 Then
                If bRet1 Then OneBend.UseDefaultBendAllowance = False
                OneBend.KFactor = K(0)
                Feat.ModifyDefinition(OneBend, Model, Nothing)
                Exit Sub
            End If
            If Angle <> 90 Or R / T < 1.5 Then
                If bRet1 Then OneBend.UseDefaultBendAllowance = False
                OneBend.KFactor = (2 * (2 * (T + R) - K(1)) / PI - R) / T
                Feat.ModifyDefinition(OneBend, Model, Nothing)
                Exit Sub
            End If
        Else
            If bRet1 = False Then
                OneBend.UseDefaultBendAllowance = True
                Feat.ModifyDefinition(OneBend, Model, Nothing)
            End If
        End If
    End Sub
#End Region

    '系数纠错（仅钣金特征纠错）
    Function LocalCorrection(Model As ModelDoc2) As Boolean

        Dim Material As String = Model.MaterialIdName
        Try
            Dim Mob As Object = Split(Material, "|")
            Material = Mob(1)
        Catch ex As Exception

        End Try

        FacDat(Material)

        Dim Feature As Feature
        Dim SheetMetal As SheetMetalFeatureData
        Dim Thickess As Double
        Dim P As Integer

        Dim FeatureMgr As FeatureManager = Model.FeatureManager
        Dim SheetMetalFolder As SheetMetalFolder = FeatureMgr.GetSheetMetalFolder

        If SheetMetalFolder IsNot Nothing Then
            P = SheetMetalFolder.GetSheetMetalCount()
            If P = 0 Then Exit Function

            If P = 1 Then
                Dim OFeat As Object = SheetMetalFolder.GetSheetMetals '获取钣金特征
                Feature = OFeat(0)
                SheetMetal = Feature.GetDefinition
                Thickess = Math.Round(SheetMetal.Thickness, 8)

                Dim bRet As Boolean
                SheetMetal.GetOverrideDefaultParameter2(0, bRet)
                If bRet Then SheetMetal.SetOverrideDefaultParameter2(0, False)

                SheetMetal.GetOverrideDefaultParameter2(1, bRet)
                If bRet Then SheetMetal.SetOverrideDefaultParameter2(1, False)

                SheetMetal.GetOverrideDefaultParameter2(2, bRet)
                If bRet Then SheetMetal.SetOverrideDefaultParameter2(2, False)

                Feature.ModifyDefinition(SheetMetal, Model, Nothing)

                End If

            '修改钣金文件夹系数
            Feature = SheetMetalFolder.GetFeature
            SheetMetal = Feature.GetDefinition
            'bRet = SheetMetal.AccessSelections(Model, Nothing)
            If Thickess <> SheetMetal.Thickness Or BendAllowance(SheetMetal, Thickess * 1000.0#) Then SheetMetal.Thickness = Thickess
            Feature.ModifyDefinition(SheetMetal, Model, Nothing)
        End If

        Dim nn As Integer = FeatureMgr.GetFeatureCount(True)
        For ii = nn - 15 To 0 Step -1
            Feature = Model.FeatureByPositionReverse(ii)
            Dim FeatN As String = Feature.GetTypeName
            If P <> 1 Then
                If FeatN = "SheetMetal" Then
                    SheetMetal = Feature.GetDefinition
                    Thickess = Math.Round(SheetMetal.Thickness, 8)
                    'bRet = SheetMetal.AccessSelections(Model, Nothing)
                    If BendAllowance(SheetMetal, Thickess * 1000.0#) Then Feature.ModifyDefinition(SheetMetal, Model, Nothing)
                End If
            End If

            If P = 1 Then
                If FeatN = "Hem" Then Process_Hem(Thickess, Feature, Model)
            Else
                If FeatN = "Hem" Then Process_Hem(0, Feature, Model)
            End If

        Next

        Return True
    End Function

    '清除特定配置
    Function DeletesCustprop(Model As ModelDoc2) As Boolean

        Dim CurCFGname = Model.GetConfigurationNames
        Dim CurCFGnameCount As Integer = Model.GetConfigurationCount

        For ii = 0 To CurCFGnameCount - 1
            Dim cusPropMgr As CustomPropertyManager = Model.Extension.CustomPropertyManager(CurCFGname(ii))
            Dim vCustomPropNames = cusPropMgr.GetNames
            If Not IsNothing(vCustomPropNames) Then
                For Each CusPropName As String In vCustomPropNames
                    cusPropMgr.Delete2(CusPropName)
                Next
            End If
        Next

    End Function

    '清除自定义属性
    Function DeletesCus(Model As ModelDoc2) As Boolean

        Dim cusPropMgr As CustomPropertyManager = Model.Extension.CustomPropertyManager("")
        Dim vCustomPropNames = cusPropMgr.GetNames
        If IsNothing(vCustomPropNames) Then Exit Function

        For ii = LBound(vCustomPropNames) To UBound(vCustomPropNames)
            Dim CusPropName As String = vCustomPropNames(ii)
            cusPropMgr.Delete2(CusPropName)
        Next

    End Function

    '固定面着色
    Function FixedFaceColour(ByRef Model As ModelDoc2) As Boolean
        Dim Part() As ModelDoc2
        If Model.GetType = 2 Then

            Dim SelMgr As SelectionMgr = Model.SelectionManager
            Dim count As Integer = SelMgr.GetSelectedObjectCount2(0)
            ReDim Part(count - 1)
            For i = 0 To count - 1
                Dim comp As Component2 = SelMgr.GetSelectedObjectsComponent4(i + 1, 0)
                Part(i) = comp.GetModelDoc2

            Next

        ElseIf Model.GetType = 1 Then
            ReDim Part(0)
            Part(0) = Model
        Else

            Return False
        End If


        For i = 0 To UBound(Part)

            Dim nn As Integer = Part(i).FeatureManager.GetFeatureCount(True)

            For ii = 0 To nn - 18
                Dim Feat As Feature = Part(i).FeatureByPositionReverse(ii)
                If Feat.GetTypeName = "FlatPattern" Then
                    Dim FlatPatt As FlatPatternFeatureData = Feat.GetDefinition
                    ' bRet = FlatPatt.AccessSelections(Model, Nothing)
                    Try
                        Dim FixedFace As Face2 = FlatPatt.FixedFace2
                        Dim vFaceProp As Object = FlatPatt.FixedFace2.MaterialPropertyValues
                        Dim Entity As Entity = FixedFace : Entity.Select4(False, Nothing)
                        If vFaceProp Is Nothing Then
                            Part(i).SelectedFaceProperties(RGB(0, 255, 255), 0, 0, 0, 0, 0, 0, False, "")
                        Else
                            Part(i).SelectedFaceProperties(0, 0, 0, 0, 0, 0, 0, True, "")
                        End If
                    Catch ex As Exception
                        FixedFaceColour = False
                    End Try
                End If
            Next
        Next
        Return True
    End Function

    '更改固定面
    Function setFixFace(Model As ModelDoc2) As Boolean

        Dim swSelMgr As SelectionMgr = Model.SelectionManager

        Try
            Dim Face As Face2 = swSelMgr.GetSelectedObject6(1, -1)
            Dim Assem As AssemblyDoc
            Dim EditModel As ModelDoc2 = Model
            If Model.GetType = 2 Then
                Assem = Model
                Assem.EditPart()
                EditModel = Assem.GetEditTarget()
            ElseIf Model.GetType <> 1 Then
                Return False
            End If

            Dim Body As Body2 = Face.GetBody
            Dim Features = Body.GetFeatures

            For i = 0 To UBound(Features)
                Dim Feat As Feature = Features(i)
                If Feat.GetTypeName = "FlatPattern" Then
                    DellFCr(Feat, EditModel)
                    Dim FlatPatt As FlatPatternFeatureData = Feat.GetDefinition
                    FlatPatt.AccessSelections(EditModel, Nothing)
                    Try
                        Dim FixedFace As Face2 = FlatPatt.FixedFace2
                        Dim Entity As Entity = FixedFace : Entity.Select4(False, Nothing)
                        EditModel.SelectedFaceProperties(0, 0, 0, 0, 0, 0, 0, True, "")
                        FlatPatt.FixedFace2 = Face
                        Feat.ModifyDefinition(FlatPatt, EditModel, Nothing)
                        FlatPatt.ReleaseSelectionAccess()
                    Catch ex As Exception

                    End Try
                    Exit For
                End If
            Next
            If Not IsNothing(Assem) Then Assem.EditAssembly()
        Catch ex As Exception
            MsgBox("请选择正确的面")
            Exit Function
        End Try
    End Function

    '草图定位
    Sub 草图定位(Part As ModelDoc2)
        Dim EditModel As ModelDoc2
        Dim Point As SketchPoint
        Dim SelMgr As SelectionMgr
        Dim SelData As SelectData
        Dim Pot(3) As Double
        Dim Bool As Boolean

        SelMgr = Part.SelectionManager
        Point = SelMgr.GetSelectedObject6(1, -1)
        If Point Is Nothing Then Exit Sub
        Pot(0) = Point.X : Pot(1) = Point.Y : Pot(2) = Point.Z
        Part.ClearSelection2(True)
        If Part.GetType = 2 Then EditModel = Part.GetEditTarget
        If EditModel Is Nothing Then
            Part.Extension.SelectAll()
            Part.Extension.MoveOrCopy(False, 1, True, Pot(0), Pot(1), 0, 0, 0, 0)
            Part.ClearSelection2(True)

            Bool = Point.Select4(False, SelData)
            Bool = Part.Extension.SelectByID2("Point1@原点", "EXTSKETCHPOINT", 0, 0, 0, True, 0, Nothing, 0)
        Else
            Bool = Part.Extension.SketchBoxSelect("-5000000", "-5000000", Nothing, "5000000", "5000000", Nothing)

            Part.Extension.MoveOrCopy(False, 1, True, Pot(0), Pot(1), 0, 0, 0, 0)

        End If

        'swEditModel.SketchAddConstraints "sgCOINCIDENT"
        Part.SketchAddConstraints("sgCOINCIDENT")
    End Sub

    '实体阵列xxx
    Sub 实体阵列(Model As ModelDoc2)
        Dim Feature As Feature， Feat(1) As Feature， SName(1) As String

        Dim swSelMgr As SelectionMgr = Model.SelectionManager
        If swSelMgr.GetSelectedObjectCount2(-1) <> 2 Then Exit Sub

        For i = 1 To 2
            If swSelMgr.GetSelectedObjectType3(i, -1) <> 22 Then Exit Sub
            Feature = Model.SelectionManager.GetSelectedObject6(i, -1)
            If Feature.GetTypeName2 = "HoleWzd" Then
                Feat(1) = Feature.GetFirstSubFeature
                SName(1) = Feat(1).Name
            End If
            If Feature.GetTypeName2 = "Stock" Then Feat(0) = Feature
        Next

        Dim Vface = Feat(0).GetFaces
        For Each Bface As Face2 In Vface
            Dim Body As Body2 = Bface.GetBody
            If Not Body Is Nothing Then
                SName(0) = Body.Name
                Dim status As Boolean = Feat(1).Select2(False, 0)
                status = Model.Extension.SelectByID2(SName(0), "SOLIDBODY", 0, 0, 0, True, 256, Nothing, 0)
                Model.FeatureManager.FeatureSketchDrivenPattern(True, False)
                Exit Sub
            End If
        Next

    End Sub


    '展开折叠
    Sub 展开折叠(Part As ModelDoc2)

        Dim boolstatus As Boolean
        Dim Feat As Feature
        Dim vSuppStateArr As Object
        Dim nn As Integer

        nn = Part.FeatureManager.GetFeatureCount(True)
        For ii = nn To 1 Step -1
            Feat = Part.FeatureByPositionReverse(nn - ii)
            If Feat.GetTypeName = "FlatPattern" Then Exit For
        Next

        boolstatus = Feat.Select2(False, 0)
        vSuppStateArr = Feat.IsSuppressed2(swInConfigurationOpts_e.swThisConfiguration, 0)
        If vSuppStateArr(0) = False Then
            Part.EditSuppress()
        ElseIf vSuppStateArr(0) = True Then
            Part.EditUnsuppress()
        End If

        Part.ClearSelection2(True)
    End Sub

    '钣金特征命名
    Sub 钣金特征命名(Part As ModelDoc2)

        Dim sheetMetalFolder As SheetMetalFolder = Part.FeatureManager.GetSheetMetalFolder
        If sheetMetalFolder Is Nothing Then
            Dim nn As Integer = Part.FeatureManager.GetFeatureCount(True)
            Dim p As Integer = 0
            For ii = nn - 18 To 1 Step -1
                Dim Feat As Feature = Part.FeatureByPositionReverse(ii)
                If Feat.GetTypeName = "SheetMetal" Then
                    p = p + 1
                    If p > 1 Then
                        If Feat.Name <> "钣金" & p Then Feat.Name = "钣金" & p
                    Else
                        If Feat.Name <> "钣金" Then Feat.Name = "钣金"
                    End If

                End If
            Next
        End If
    End Sub

    '钣金法兰辅助尺寸
    Sub DimFlange(Model As ModelDoc2)

        Dim SelMgr As SelectionMgr = Model.SelectionManager
        If SelMgr.GetSelectedObjectCount2(-1) <> 1 Then Exit Sub
        Debug.Print(SelMgr.GetSelectedObjectType3(1, -1))
        Dim Feat As Feature
        Select Case SelMgr.GetSelectedObjectType3(1, -1)
            Case 22
                Feat = Model.SelectionManager.GetSelectedObject6(1, -1)
                Exit Select
            Case 2
                Dim Face1 As Face2 = Model.SelectionManager.GetSelectedObject6(1, -1)
                Feat = Face1.GetFeature
            Case Else
                MsgBox("请选择正确的对象")
                Exit Sub
        End Select

        If Feat.GetTypeName <> "EdgeFlange" Then Exit Sub

        Dim subFeat As Feature = Feat.GetFirstSubFeature

        Do While Not IsNothing(subFeat)
            If subFeat.GetTypeName = "ProfileFeature" Then

#Region "法兰调整"
                '法兰调整
                Dim MathUtil As MathUtility = MySwApp.GetMathUtility
                Dim Pt1, Pt2
                Dim Abl As Boolean = False
                '''If SzmStr(9) = 1 Then Abl = True

                Try
                    subFeat.Select2(False, Nothing)
                Catch ex As Exception

                End Try

                Model.EditSketch() ' 编辑草图
                Dim Sketch As Sketch = Model.GetActiveSketch2

                Dim MathTrans As MathTransform = Sketch.ModelToSketchTransform.Inverse '设置矩阵对象MathTrans
                Dim SkLine(3) As SketchLine

                Dim SkArr As Object = Sketch.GetSketchSegments
                Dim SkSeg As SketchSegment
                For Each SkSeg In SkArr
                    Dim SkRelArr As Object = SkSeg.GetRelations()
                    If IsNothing(SkRelArr) Then GoTo Ex1
                    For Each SkRel As SketchRelation In SkRelArr
                        Select Case SkRel.GetRelationType
                            Case 5
                                If IsNothing(SkLine(1)) Then
                                    SkLine(1) = SkSeg
                                Else
                                    SkLine(2) = SkSeg
                                End If
                            Case 32
                                SkLine(0) = SkSeg
                            Case 4
                                SkLine(3) = SkSeg
                        End Select
                    Next
Ex1:            Next

                Dim Pot(3) As SketchPoint, PtDle1(2), PtDle2(2), Parsms(6) As Double
                Pot(0) = SkLine(1).GetStartPoint2
                If Math.Round(Pot(0).Y, 10) <> 0 Then
                    Pot(1) = Pot(0) : Pot(0) = SkLine(1).GetEndPoint2
                Else
                    Pot(1) = SkLine(1).GetEndPoint2
                End If
                Pot(2) = SkLine(2).GetStartPoint2
                If Math.Round(Pot(2).Y, 10) <> 0 Then
                    Pot(3) = Pot(2) : Pot(2) = SkLine(2).GetEndPoint2
                Else
                    Pot(3) = SkLine(2).GetEndPoint2
                End If

                PtDle1(0) = (Pot(0).X + Pot(2).X) / 2 : PtDle1(1) = -0.1 / 1000
                PtDle2(0) = PtDle1(0) : PtDle2(1) = PtDle1(1) : PtDle2(2) = 1

                Dim MathP As MathPoint = MathUtil.CreatePoint(PtDle1).MultiplyTransform(MathTrans)
                Pt1 = MathP.ArrayData
                MathP = MathUtil.CreatePoint(PtDle2).MultiplyTransform(MathTrans)
                Pt2 = MathP.ArrayData
                Parsms(0) = Pt1(0) : Parsms(1) = Pt1(1) : Parsms(2) = Pt1(2) : Parsms(3) = Pt2(0) : Parsms(4) = Pt2(1) : Parsms(5) = Pt2(2)
                Dim SelParams As Object = Parsms
                Dim bRet As Boolean = Model.SelectByRay(SelParams, swSelectType_e.swSelFACES) '选取面
                Dim Face As Face2 = Model.SelectionManager.GetSelectedObject6(1, -1)
                '获得法兰边线草图所在的面
                bRet = Pot(0).X > Pot(2).X
                If bRet Then
                    Pot(0).Select4(False, Nothing)
                    Model.Extension.MoveOrCopy(False, 1, True, 0, 0, 0, -10 / 1000, 0, 0)
                    Pot(2).Select4(False, Nothing)
                    Model.Extension.MoveOrCopy(False, 1, True, 0, 0, 0, 10 / 1000, 0, 0)
                Else
                    Pot(0).Select4(False, Nothing)
                    Model.Extension.MoveOrCopy(False, 1, True, 0, 0, 0, 10 / 1000, 0, 0)
                    Pot(2).Select4(False, Nothing)
                    Model.Extension.MoveOrCopy(False, 1, True, 0, 0, 0, -10 / 1000, 0, 0)
                End If

                Dim Pit(1, 2) As Double
                PtDle1(0) = (Pot(0).X + Pot(1).X) / 2 : PtDle1(1) = (Pot(0).Y + Pot(1).Y) / 2
                MathP = MathUtil.CreatePoint(PtDle1).MultiplyTransform(MathTrans)
                Pt1 = MathP.ArrayData
                PtDle2(0) = (Pot(2).X + Pot(3).X) / 2 : PtDle2(1) = (Pot(2).Y + Pot(3).Y) / 2
                MathP = MathUtil.CreatePoint(PtDle2).MultiplyTransform(MathTrans)
                Pt2 = MathP.ArrayData
                Pit(0, 0) = Pt1(0) : Pit(0, 1) = Pt1(1) : Pit(0, 2) = Pt1(2)
                Pit(1, 0) = Pt2(0) : Pit(1, 1) = Pt2(1) : Pit(1, 2) = Pt2(2)

                PtDle1(0) = Pot(0).X : PtDle1(1) = Pot(0).Y : PtDle1(1) = Pot(0).Z
                MathP = MathUtil.CreatePoint(PtDle1).MultiplyTransform(MathTrans)
                Pt1 = MathP.ArrayData
                PtDle2(0) = Pot(2).X : PtDle2(1) = Pot(2).Y : PtDle2(1) = Pot(2).Z
                MathP = MathUtil.CreatePoint(PtDle2).MultiplyTransform(MathTrans)
                Pt2 = MathP.ArrayData


                Dim Edges As Object = Face.GetEdges
                For Each Edge As Edge In Edges
                    Dim hs(1) As Double
                    hs(0) = Model.ClosestDistance(Edge, Pot(0), Nothing, Nothing)
                    hs(0) = Math.Round(hs(0), 10)
                    hs(1) = Model.ClosestDistance(Edge, Pot(2), Nothing, Nothing)
                    hs(1) = Math.Round(hs(1), 10)
                    Dim Led As Double = hs(1) - hs(0)

                    If Led = 0 And hs(0) <> 0 Then
                        Dim Ent As Entity = Edge : Ent.Select4(False, Nothing)
                        SkSeg = SkLine(3) : SkSeg.Select4(True, Nothing)
                        Model.AddDiameterDimension2((Pit(0, 0) + Pit(1, 0)) / 2, (Pit(0, 1) + Pit(1, 1)) / 2, (Pit(0, 2) + Pit(1, 2)) / 2)

                    ElseIf Led > 0 Then
                        Dim Ent As Entity = Edge : Ent.Select4(False, Nothing)
                        Pot(0).Select4(True, Nothing)
                        Dim swDyDim As DisplayDimension = Model.AddDiameterDimension2(Pit(0, 0), Pit(0, 1), Pit(0, 2))
                        Dim Dimen As Dimension = swDyDim.GetDimension2(0)
                        Dimen.SystemValue = 0
                    ElseIf Led < 0 Then
                        Dim Ent As Entity = Edge : Ent.Select4(False, Nothing)
                        Pot(2).Select4(True, Nothing)
                        Dim swDyDim As DisplayDimension = Model.AddDiameterDimension2(Pit(1, 0), Pit(1, 1), Pit(1, 2))
                        Dim Dimen As Dimension = swDyDim.GetDimension2(0)
                        Dimen.SystemValue = 0
                    End If

                Next

                If Abl Then
                    If bRet Then
                        PtDle2(0) = Pot(0).X + 15 / 1000 : PtDle1(1) = Pot(0).Y + 15 / 1000
                        PtDle1(0) = Pot(2).X - 15 / 1000 : PtDle2(1) = Pot(2).Y + 15 / 1000
                    Else
                        PtDle1(0) = Pot(0).X + 15 / 1000 : PtDle1(1) = Pot(0).Y + 15 / 1000
                        PtDle2(0) = Pot(2).X - 15 / 1000 : PtDle2(1) = Pot(2).Y + 15 / 1000
                    End If
                    MathP = MathUtil.CreatePoint(PtDle1).MultiplyTransform(MathTrans)
                    Pt1 = MathP.ArrayData
                    MathP = MathUtil.CreatePoint(PtDle2).MultiplyTransform(MathTrans)
                    Pt2 = MathP.ArrayData

                    SkSeg = SkLine(1) : SkSeg.Select4(False, Nothing)
                    Model.SketchConstraintsDel(0, "sgVERTICAL2D")
                    SkSeg = SkLine(0) : SkSeg.Select4(True, Nothing)
                    Model.AddDiameterDimension2(Pt1(0), Pt1(1), Pt1(2))
                    SkSeg = SkLine(2) : SkSeg.Select4(False, Nothing)
                    Model.SketchConstraintsDel(0, "sgVERTICAL2D")
                    SkSeg = SkLine(0) : SkSeg.Select4(True, Nothing)
                    Model.AddDiameterDimension2(Pt2(0), Pt2(1), Pt2(2))
                End If
                Model.SketchManager.InsertSketch(True)
#End Region
            End If
            subFeat = subFeat.GetNextSubFeature
        Loop



    End Sub

    '隐藏所有草图
    Sub HideSk(Model As ModelDoc2)

        Dim nn As Integer = Model.FeatureManager.GetFeatureCount(True)
        For ii = nn - 18 To 1 Step -1
            Dim Feat As Feature = Model.FeatureByPositionReverse(ii)

            Select Case Feat.GetTypeName
                Case "ProfileFeature"
                    Dim CdFeat As Object = Feat.GetChildren
                    If CdFeat Is Nothing Then
                        If Feat.Visible = 1 Then
                            Feat.Select2(False, 0)
                            Model.UnblankSketch()
                        End If
                    Else
                        If Feat.visble <> 1 Then
                            Feat.Select2(False, 0)
                            Model.BlankSketch()
                        End If
                    End If

                Case Else
                    Dim subFeat As Feature = Feat.GetFirstSubFeature
                    Do Until subFeat Is Nothing
                        If subFeat.GetTypeName = "ProfileFeature" Then
                            If subFeat.Visible <> 1 Then
                                subFeat.Select2(False, 0)
                                Model.BlankSketch()
                            End If
                        End If
                        subFeat = subFeat.GetNextSubFeature
                    Loop
            End Select
        Next

    End Sub

    '草图约束
    Sub SketchConstrainAll(Model As ModelDoc2)

        Dim Sketch As Sketch = Model.GetActiveSketch
        Sketch.ConstrainAll()

    End Sub

    '零件同步工程图命名
    Sub RenamePart(Model As ModelDoc2)
        Dim MoType As Integer = Model.GetType
        If MoType <> 1 And MoType <> 2 Then Exit Sub
        Dim swSelMgr As SelectionMgr = Model.SelectionManager
        If swSelMgr.GetSelectedObjectCount2(-1) <> 1 Then Exit Sub
        If swSelMgr.GetSelectedObjectType3(1, -1) <> 20 Then
            MsgBox("请在模型树中选择一个零件")
            Exit Sub
        End If

        Dim NewName As String = InputBox("输入新名称")
        If NewName = "" Then Exit Sub
        Dim ModelPth1 As String

        If MoType = 2 Then
            Dim swComp As Component2 = swSelMgr.GetSelectedObjectsComponent3(1, -1)
            'swComp.Name2 = NewName
            ModelPth1 = swComp.GetPathName
        Else
            ModelPth1 = Model.GetPathName
        End If

        Try
            Model.Extension.RenameDocument(NewName)
            'Model.ClearSelection2(True)
            Dim bRet As Boolean = Model.Save3(2, 0, 0)
            If Not bRet Then MsgBox("名称有误") : Exit Sub
        Catch ex As Exception
            MsgBox("名称有误")
            Exit Sub
        End Try

        Dim PathStr As String = Left(ModelPth1, InStrRev(ModelPth1, "\"))

        Dim PathName As String = Left(ModelPth1, Len(ModelPth1) - 7) & ".SLDDRW"
        If File.Exists(PathName) Then
            Dim ModelPth2 As String = PathStr & NewName & ".sldprt"
            MySwApp.ReplaceReferencedDocument(PathName, ModelPth1, ModelPth2)
            NewName = PathStr & NewName & ".SLDDRW"
            File.Move(PathName, NewName)
        End If
        UpdateComponentName(Model)
        Model.Save3(2, 0, 0)
    End Sub

    '零件同步工程图命名
    Sub setModelname(Model As ModelDoc2)

        Dim swSelMgr As SelectionMgr = Model.SelectionManager
        If swSelMgr.GetSelectedObjectCount2(-1) <> 1 Then Exit Sub
        If swSelMgr.GetSelectedObjectType3(1, -1) <> 20 Then Exit Sub

        Dim NewName As String = InputBox("输入新名称")
        If NewName = "" Then Exit Sub
        Dim ModelPth1 As String

        Try
            Model.Extension.RenameDocument(NewName)
            'Model.ClearSelection2(True)
            Dim bRet As Boolean = Model.Save3(2, 0, 0)
            If Not bRet Then MsgBox("名称有误") : Exit Sub
        Catch ex As Exception
            MsgBox("名称有误")
            Exit Sub
        End Try

        Dim PathStr As String = Left(ModelPth1, InStrRev(ModelPth1, "\"))

        Dim PathName As String = Left(ModelPth1, Len(ModelPth1) - 7) & ".SLDDRW"
        If File.Exists(PathName) Then
            Dim ModelPth2 As String = PathStr & NewName & ".sldprt"
            MySwApp.ReplaceReferencedDocument(PathName, ModelPth1, ModelPth2)
            NewName = PathStr & NewName & ".SLDDRW"
            File.Move(PathName, NewName)
        End If
        UpdateComponentName(Model)
        Model.Save3(2, 0, 0)
    End Sub


    '粘贴来自cad的文件
    Sub pasteCAD()
        Dim TPme As String = dllPath & "Sm.drwdot" 'SzmStr(0)
        Dim DraDoc As DrawingDoc = MySwApp.NewDocument(TPme, swDwgPaperSizes_e.swDwgPaperAsize, 0, 0)
        Dim ModelDoc As ModelDoc2 = DraDoc
        ModelDoc.Paste()
        ModelDoc.Extension.SelectAll()
        ModelDoc.EditCopy()
        MySwApp.CloseDoc(ModelDoc.GetTitle())
    End Sub

#Region "子函数"
    '钣金输出
    Private Function DxfOut(ByRef Feat As Feature, ByRef Model As ModelDoc2, ByRef Spath As String, ByRef ModelPath As String) As Boolean
        DxfOut = False
        Dim Optin As Integer = 1
        Feat.Select2(False, 0) : Model.EditUnsuppress2() '解压缩
        If Feat.GetErrorCode2(Nothing) <> "0" Then Feat.Select2(False, 0) : Model.EditSuppress() : Exit Function
        Feat.Select2(False, 0) : Model.EditSuppress() '压缩

        If SzmStr(7) = "1" Then
            Optin = Optin + 8
        End If
        If SzmStr(8) = "1" Then
            Optin = Optin + 4
        End If

        Feat.Select2(False, 0)
        Model.ExportToDWG2(Spath & ".dxf", ModelPath, 1, True, Nothing, Nothing, False, Optin, Nothing)
        Return True
    End Function
    Private Function DxfOut2(ByRef Feat As Feature, ByRef Model As ModelDoc2, ByRef Spath As String, ByRef ModelPath As String) As Boolean

        Dim XDirection As Boolean = False '是否x轴翻转
        Feat.Select2(False, 0) : Model.EditUnsuppress2() '解压缩

        If Feat.GetErrorCode2(Nothing) <> "0" Then
            Feat.Select2(False, 0) : Model.EditSuppress() '压缩
            Return False
        End If
        MySwApp.RunCommand(167, "")

        Dim SubFeat As Feature = Feat.GetFirstSubFeature
        SubFeat.Select2(False, 0)

        MySwApp.SetUserPreferenceToggle(8, False)
        If SzmStr(8) = "1" Then
            Model.UnblankSketch()
        Else
            Model.BlankSketch() '隐藏
        End If
        SubFeat = SubFeat.GetNextSubFeature
        SubFeat.Select2(False, 0)
        MySwApp.RunCommand(169, "")
        Model.BlankSketch() '隐藏

        Dim dataViews As String() = {"*当前"}
        Dim Part As PartDoc = Model
        Part.ExportToDWG2(Spath & ".dxf", ModelPath, 3, False, Nothing, XDirection, False, 0, dataViews)

        If File.Exists(Spath & ".dxf") Then File.Delete(Spath & ".dxf")
        File.Move(Spath & " (当前).dxf", Spath & ".dxf")

        SubFeat = Feat.GetFirstSubFeature
        SubFeat.Select2(False, 0)
        Model.UnblankSketch() '显示
        SubFeat = SubFeat.GetNextSubFeature
        SubFeat.Select2(False, 0)
        MySwApp.RunCommand(169, "")
        Model.UnblankSketch() '显示
        Feat.Select2(False, 0) : Model.EditSuppress() '压缩
        MySwApp.SetUserPreferenceToggle(8, True)
        Return True

    End Function
    Private Function DxfOut3(ByRef Feat As Feature, ByRef Model As ModelDoc2, ByRef Spath As String, ByRef ModelPath As String) As Boolean
        Feat.Select2(False, 0) : Model.EditUnsuppress2() '解压缩

        If Feat.GetErrorCode2(Nothing) <> "0" Then
            Feat.Select2(False, 0) : Model.EditSuppress() '压缩
            Return False
        End If

        Dim swFlatPat As FlatPatternFeatureData = Feat.GetDefinition
        Dim swFixFace As Face2
        Try
            swFixFace = Feat.GetDefinition.FixedFace2
        Catch ex As Exception
            Dim swFixEdge As Edge = Feat.GetDefinition.FixedFace2
            Dim ObFace As Object = swFixEdge.GetTwoAdjacentFaces2
            swFixFace = ObFace(0)
        End Try

        Dim Eity As Entity = swFixFace
        Eity.Select4(False, Nothing)

        Dim Part As PartDoc = Model
        Part.ExportToDWG2(Spath & ".dxf", ModelPath, 2, False, Nothing, Nothing, False, 1, Nothing)

        If File.Exists(Spath & ".dxf") Then File.Delete(Spath & ".dxf")
        File.Move(Spath & " (Entity_0).dxf", Spath & ".dxf")
        Feat.Select2(False, 0) : Model.EditSuppress() '压缩
        Return True
    End Function
    Private Function DxfOutD(ByRef Feat As Feature, ByRef Model As ModelDoc2, Prop() As String) As Boolean
        Feat.Select2(False, 0) : Model.EditUnsuppress2() '解压缩
        If Feat.GetErrorCode2(Nothing) <> "0" Then
            Feat.Select2(False, 0) : Model.EditSuppress() '压缩
            Return False
        End If
        Dim Spath As String = Model.GetPathName
        'MySwApp.RunCommand(167, "")
        Model.ActiveView.RollBy(0.001)
        Dim SubFeat As Feature = Feat.GetFirstSubFeature
        SubFeat.GetNextSubFeature.Select2(False, 0) : MySwApp.RunCommand(169, "")

        Dim DraDoc As DrawingDoc = MySwApp.NewDocument(dllPath & "Sm.drwdot", 12, 0, 0)
        Dim myView As View = DraDoc.CreateDrawViewFromModelView3(Spath, "", 0, 0, 0)
        Dim Enty As Entity = Feat : myView.SelectEntity(Enty, False) : DraDoc.ChangeComponentLayer("钣金展开", True)
        'DraDoc.SetupSheet6("图纸1", Nothing, Nothing, 1, 1, True, "", 0, 0, "", True, 0, 0, 0, 0, 0, 0)
        Dim DraModel As ModelDoc2 = DraDoc
        If SzmStr(8) = "1" Then
            Dim Sketch As Sketch = SubFeat.GetSpecificFeature2
            Dim OSketchSeg = Sketch.GetSketchSegments
            If OSketchSeg Is Nothing Then GoTo Rx
            DraModel.ClearSelection2(True) : DraDoc.SetCurrentLayer("折弯")
            Dim Names As String = Model.GetTitle : Names = Left(Names, Len(Names) - 7)
            For i = 0 To UBound(OSketchSeg)
                Dim SketchSeg As SketchSegment = OSketchSeg(i)
                Names = Strings.Replace(SketchSeg.GetName(), "直线", "Line") & "@" & SubFeat.Name & "@" & Names & "@" & myView.Name 'Line
                DraModel.Extension.SelectByID2(Names, "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0)
            Next i
            DraModel.SketchManager.SketchUseEdge3(False, False)
        End If
        Enty = SubFeat : myView.SelectEntity(Enty, False) : DraDoc.BlankSketch()

Rx:     Spath = Strings.Left(Spath, Len(Spath) - 7)
        If Prop(0) <> "" Then
            If Prop(0) = "默认" Then
                Prop(0) = ""
            Else
                Prop(0) = "." & Prop(0)
            End If
        End If

        Prop(2) = "-" & Prop(2) & "T"
        If Prop(3) = "1" Then
            Prop(3) = ""
        Else
            Prop(3) = "(数量" & Prop(3) & "件)"
        End If
        Spath = Spath & Prop(0) & Prop(4) & " " & Prop(1) & Prop(2) & Prop(3)
        If SzmStr(10) = "1" Then
            DraModel.SaveAs3(Spath & ".DWG", 0, 2)
        Else
            DraModel.SaveAs3(Spath & ".dxf", 0, 2)
        End If

        MySwApp.CloseDoc("")
        Feat.Select2(False, 0) : Model.EditSuppress() '压缩

        Return True

    End Function
    Private Function DxfOutDll(ByRef Feat As Feature, ByRef Model As ModelDoc2, Prop() As String, ByRef DraDoc As DrawingDoc) As Boolean

        Feat.Select2(False, 0) : Model.EditUnsuppress2() '解压缩
        If Feat.GetErrorCode2(Nothing) <> "0" Then
            Feat.Select2(False, 0) : Model.EditSuppress() '压缩
            Return False
        End If

        'MySwApp.RunCommand(167, "")
        Model.ActiveView.RollBy(0.001)
        Dim SubFeat As Feature = Feat.GetFirstSubFeature
        SubFeat.GetNextSubFeature.Select2(False, 0) : MySwApp.RunCommand(169, "")
        Dim DraModel As ModelDoc2
        Dim bRets As Boolean = False
        If DraDoc Is Nothing Then bRets = True : DraDoc = MySwApp.NewDocument(dllPath & "Sm.drwdot", 12, 0, 0)

        DraModel = DraDoc
        MySwApp.ActivateDoc3(DraModel.GetTitle, False, 2, Nothing)


        Dim myView As View = DraDoc.CreateDrawViewFromModelView3(Model.GetPathName, "", 0, 0, 0)
        Dim Enty As Entity = Feat : myView.SelectEntity(Enty, False) : DraDoc.ChangeComponentLayer("钣金展开", True)

        If bRet(2) = "1" Then
            Dim Sketch As Sketch = SubFeat.GetSpecificFeature2
            Dim OSketchSeg = Sketch.GetSketchSegments
            If OSketchSeg Is Nothing Then GoTo Rx
            DraModel.ClearSelection2(True) : DraDoc.SetCurrentLayer("折弯")
            Dim Names As String = Model.GetTitle : Names = Left(Names, Len(Names) - 7)
            For i = 0 To UBound(OSketchSeg)
                Dim SketchSeg As SketchSegment = OSketchSeg(i)
                Names = Strings.Replace(SketchSeg.GetName(), "直线", "Line") & "@" & SubFeat.Name & "@" & Names & "@" & myView.Name 'Line
                DraModel.Extension.SelectByID2(Names, "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0)
            Next i
            DraModel.SketchManager.SketchUseEdge3(False, False)
        End If
        Enty = SubFeat : myView.SelectEntity(Enty, False) : DraDoc.BlankSketch()

Rx:     Dim Spath As String = Model.GetTitle

        Spath = Strings.Left(Spath, Len(Spath) - 7)
        If Prop(0) <> "" Then
            If Prop(0) = "默认" Then
                Prop(0) = ""
            Else
                Prop(0) = "." & Prop(0)
            End If
        End If
        Spath = Spath & Prop(0) & Prop(4)

        If bRets = False Then
            DraModel.Extension.SelectByID2(myView.Name, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
            myView.ReplaceViewWithSketch()
            myView = DraDoc.ActiveDrawingView
            myView.SetName2(Spath)
            Dim Xb() As Double = myView.GetOutline

            If Xo <> 0 Then
                PitX = PitX + (Xb(2) - Xb(0)) / 2 + Xo + 0.125
            End If
            myView.Position = {PitX, Fio}
            DraDoc.SetCurrentLayer("文字")
            DraModel.FontUnits(6.5)
            DraModel.FontFace("Txt")
            Dim Note As Note = DraModel.InsertNote("NAME=" & Spath + Chr(13) + Chr(10) +
                                                    "M=" & Prop(1) + Chr(13) + Chr(10) +
                                                    "T=" & Prop(2) + Chr(13) + Chr(10) +
                                                    "Q=" & Prop(3))
            Note.Angle = 0
            Note.SetBalloon(0, 0)
            Note.LockPosition = False
            Dim Annotation As Annotation = Note.GetAnnotation()
            Annotation.SetLeader3(0, 0, False, False, False, False)
            Annotation.SetPosition2(PitX, Fio, 0)
            Annotation.LeaderLineStyle = 1
            Dim outline() As Double = myView.GetOutline
            Dim Pot() As Double = myView.Position

            MySwApp.ActivateDoc3(Model.GetTitle, False, 1, Nothing)
            Feat.Select2(False, 0) : Model.EditSuppress() '压缩
            Iox = Iox + 1
            Xo = (Xb(2) - Xb(0)) / 2
            If Iox > 25 Then Iox = 0 : Fio = Fio - 2 : PitX = 2 : Xo = 0
            Return True
        Else

            DraDoc.SetCurrentLayer("文字")
            DraModel.FontUnits(6.5)
            DraModel.FontFace("Txt")
            Dim Note As Note = DraModel.InsertNote("NAME=" & Spath + Chr(13) + Chr(10) +
                                                    "M=" & Prop(1) + Chr(13) + Chr(10) +
                                                    "T=" & Prop(2) + Chr(13) + Chr(10) +
                                                    "Q=" & Prop(3))
            Note.Angle = 0
            Note.SetBalloon(0, 0)
            Note.LockPosition = False
            Dim Annotation As Annotation = Note.GetAnnotation()
            Annotation.SetLeader3(0, 0, False, False, False, False)
            Annotation.SetPosition2(0, 0, 0)
            Annotation.LeaderLineStyle = 1

            Spath = Model.GetPathName
            Spath = Strings.Left(Spath, Len(Spath) - 7) & Prop(0) & Prop(4)
            If SzmStr(10) = "1" Then
                DraModel.SaveAs3(Spath & ".DWG", 0, 2)
            Else
                DraModel.SaveAs3(Spath & ".dxf", 0, 2)
            End If
            MySwApp.CloseDoc("")
            DraDoc = Nothing
            Feat.Select2(False, 0) : Model.EditSuppress() '压缩
            Return True
        End If

    End Function

    '子特征系数检查
    Private Function BendDf(ByRef Feat As Feature, ByRef Model As ModelDoc2, ByRef Thickess As Double) As String

        BendDf = ""
        Dim dea As Double
        Try
            dea = SzmStr(1)
        Catch ex As Exception
            dea = 0
        End Try

        Dim kc, ks, ki As Double
        Dim OneBend As OneBendFeatureData
        Dim SubFeat As Feature = Feat.GetFirstSubFeature

        For i = 0 To UBound(Factor)
            If Factor(i, 0) = 0 Then BendDf = "此厚度没有录入" : Exit Function
            If Factor(i, 0) = Thickess Then
                kc = 2 * (Factor(i, 0) + Factor(i, 1)) - 0.5 * PI * (Factor(i, 1) + Factor(i, 0) * Factor(i, 2))
                ks = Factor(i, 4)
                ki = Factor(i, 2)
                Exit For
            End If

        Next

        While Not SubFeat Is Nothing
            If SubFeat.GetTypeName = "UiBend" Then

                Dim vSuppStateArr As Object = SubFeat.IsSuppressed2(swInConfigurationOpts_e.swThisConfiguration, 0)
                If vSuppStateArr(0) = True Then
                    SubFeat.Select2(False, 0)
                    Model.EditUnsuppress2()
                End If

                Dim FeatName As String = SubFeat.Name
                If InStr(FeatName, "Derived") > 0 Then GoTo Ex

                Dim Parents As Object = SubFeat.GetParents
                Dim BendFeat As Feature

                For Each BendFeat In Parents
                    FeatName = BendFeat.GetTypeName
                    If InStr(FeatName, "Bend") > 0 Then Exit For
                Next

                vSuppStateArr = BendFeat.IsSuppressed2(swInConfigurationOpts_e.swThisConfiguration, 0)
                If vSuppStateArr(0) = True Then
                    BendFeat.Select2(False, 0)
                    Model.EditUnsuppress2()
                End If
                OneBend = BendFeat.GetDefinition
                FeatName = BendFeat.Name & " "

                If OneBend.BendAllowanceType = -1 Then

                    Dim CustBend As CustomBendAllowance = OneBend.GetCustomBendAllowance

                    Dim k(2) As Double
                    k(0) = Math.Round(CustBend.BendDeduction * 1000, 2)
                    k(1) = Math.Round(OneBend.BendAngle * 180 / PI, 1)
                    k(2) = Math.Round(OneBend.BendRadius * 1000, 2)

                    If k(2) / Thickess < 1.5 Or k(2) <> 90 Then
                        If dea < Math.Abs(k(0) - kc) Then BendDf = FeatName & "系数错误" : Exit Function
                    Else
                        BendDf = FeatName & "折弯半径超出扣除范围"
                        Exit Function
                    End If

                ElseIf OneBend.BendAllowanceType = 2 Then

                    Dim k(1) As Double
                    k(0) = Math.Round(OneBend.KFactor, 2)
                    k(1) = Math.Round(OneBend.BendRadius * 1000, 2)


                    If k(1) / Thickess < 1.5 Then
                        If Math.Round(OneBend.BendAngle * 180 / PI, 1) <> 180 Then
                            k(0) = 2 * (Thickess + k(1)) - 0.5 * PI * (k(1) + Thickess * k(0))
                            If dea < Math.Abs(k(0) - kc) Then BendDf = FeatName & "系数错误" : Exit Function
                        Else

                            If k(0) <> ks And SzmStr(5) = "1" Then BendDf = FeatName & "系数错误" : Exit Function

                        End If

                    Else
                        If ki <> k(0) Then BendDf = FeatName & "系数错误" : Exit Function
                    End If

                Else
                    BendDf = FeatName & "系数类型不正确请检查"
                    Exit Function

                End If

            End If
Ex:         SubFeat = SubFeat.GetNextSubFeature
        End While

    End Function
    '读取固定面颜色
    Private Function GetFCr(FlatPatt As FlatPatternFeatureData, Model As ModelDoc2) As Boolean

        GetFCr = True

        Try
            Dim FixedFace As Face2 = FlatPatt.FixedFace2
            Dim Entity As Entity = FixedFace
            Entity.Select4(False, Nothing)
            Dim vFaceProp = FixedFace.MaterialPropertyValues

            If vFaceProp Is Nothing Then
                GetFCr = False : Exit Function
            ElseIf vFaceProp(0) <> "0" And vFaceProp(1) <> "1" And vFaceProp(2) <> "1" Then
                GetFCr = False : Exit Function
            End If
        Catch ex As Exception
            GetFCr = True
        End Try

    End Function
    '特征系数修改
    Private Function BendAllowance(ByRef SheetMetal As SheetMetalFeatureData, ByRef Thickess As Double) As Boolean
        Dim boole As Boolean = False
        For i = 0 To UBound(Factor)
            If Factor(i, 0) = 0 Then Exit Function
            If Factor(i, 0) = Thickess Then
                Dim PBendData As CustomBendAllowance = SheetMetal.GetCustomBendAllowance

                If SzmStr(4) = "1" Then SheetMetal.BendRadius = Factor(i, 1) / 1000.0# : boole = True
                If SzmStr(3) = "1" Then
                    If PBendData.Type <> 3 Then PBendData.Type = 3 : boole = True
                    If PBendData.BendDeduction <> Factor(i, 3) / 1000.0# Then PBendData.BendDeduction = Factor(i, 3) / 1000.0# : boole = True
                    SheetMetal.SetCustomBendAllowance(PBendData)
                    Return boole
                Else
                    If SheetMetal.KFactor <> Factor(i, 2) Then SheetMetal.KFactor = Factor(i, 2) : boole = True
                    Return boole
                End If

                Exit For
            End If
        Next

    End Function

    Private Sub TraSubFeat(ByRef Feat As Feature)
        Dim subFeat As Feature = Feat.GetFirstSubFeature
        While IsNothing(subFeat) = False
            TraSubFeat(subFeat)
            subFeat = subFeat.GetNextSubFeature
            subFeat.Select2(True, Nothing)
        End While
    End Sub

    Private Function DellFCr(Fl As Feature, Model As ModelDoc2) As Boolean
        Dim FlatPatt As FlatPatternFeatureData = Fl.GetDefinition
        ' bRet = FlatPatt.AccessSelections(Model, Nothing)
        Try
            Dim FixedFace As Face2 = FlatPatt.FixedFace2
            Dim vFaceProp As Object = FlatPatt.FixedFace2.MaterialPropertyValues
            Dim Entity As Entity = FixedFace : Entity.Select4(False, Nothing)
            Model.SelectedFaceProperties(RGB(0, 255, 255), 0, 0, 0, 0, 0, 0, True, "")
        Catch ex As Exception
            DellFCr = False
        End Try
        ' FlatPatt.ReleaseSelectionAccess()
        Return True
    End Function

    Private Function AddFCr(Fl As Feature, Model As ModelDoc2) As Boolean

        Dim FlatPatt As FlatPatternFeatureData = Fl.GetDefinition
        ' bRet = FlatPatt.AccessSelections(Model, Nothing)
        Try
            Dim FixedFace As Face2 = FlatPatt.FixedFace2
            Dim vFaceProp As Object = FlatPatt.FixedFace2.MaterialPropertyValues
            Dim Entity As Entity = FixedFace : Entity.Select4(False, Nothing)
            Model.SelectedFaceProperties(RGB(0, 255, 255), 0, 0, 0, 0, 0, 0, False, "")
        Catch ex As Exception
            AddFCr = False
        End Try
        ' FlatPatt.ReleaseSelectionAccess()
        Return True
    End Function

#End Region

End Module
'钣金输出模块支持的功能：
'1返回钣金导出过程的错误给用户
'读取钣金系数文件 如果系数文件不存在则自动生成 并且告知用户
'超级选择工具
