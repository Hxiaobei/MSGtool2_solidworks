Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Imports System.IO

Module DraDoc
    Structure FeatC   '创建用户bai自定义的类du型.
        Dim Agle As Double
        Dim Cface As Face2   '定义元素的数据类型。
        Dim SkPot As SketchPoint
        Dim Aface() As Face2
        Dim Abyte As Byte
    End Structure

    Const PI As Double = 3.14159265358979
    Dim View As View, DrawingA As ModelDoc2
    Dim swMathUtil As MathUtility = MySwApp.GetMathUtility
    Dim ViewMathTransform As MathTransform
    Dim Vec(), Thickness, BendR As Double
    '法兰标注主函数入口
    Sub BendMark(Drawing As DrawingDoc)
        View = Drawing.ActiveDrawingView
        If View Is Nothing Then Exit Sub
        DrawingA = Drawing
#Region "获取视图z轴转换向量"
        Dim zVector As Object
        ViewMathTransform = View.ModelToViewTransform.Inverse()
        ViewMathTransform.GetData2(Nothing, Nothing, zVector, Nothing, 0)
        Vec = zVector.ArrayData
        Vec(0) = Math.Round(Vec(0), 10)
        Vec(1) = Math.Round(Vec(1), 10)
        Vec(2) = Math.Round(Vec(2), 10)
        ViewMathTransform = View.ModelToViewTransform
#End Region
        Dim Fe() As FeatC, Fi As Byte
        Dim Model As ModelDoc2 = View.ReferencedDocument : If Model.GetType <> 1 Then Exit Sub
        Dim Bodys As Object = View.Bodies : If IsNothing(Bodys) Then Bodys = Model.GetBodies2(0, True)
        For Each Body As Body2 In Bodys
            ReDim Fe(0)
            Dim Feats = Body.GetFeatures
            For Each Feat As Feature In Feats
                If Feat.GetTypeName = "FlatPattern" Then
                    Dim ObFeat = Feat.GetParents
                    For Each smFeat As Feature In ObFeat
                        If smFeat.GetTypeName = "SheetMetal" Then
                            Dim SheetMetal As SheetMetalFeatureData = smFeat.GetDefinition
                            Thickness = SheetMetal.Thickness
                        End If
                    Next
                    Dim subFeat As Feature = Feat.GetFirstSubFeature
                    Try

                    Catch ex As Exception

                    End Try
                    Do While Not subFeat Is Nothing
                        If subFeat.GetTypeName = "UiBend" Then
                            Dim PFeat = subFeat.GetParents
                            Dim BendFeat As Feature = PFeat(0)
                            Dim BendData As OneBendFeatureData = BendFeat.GetDefinition
                            BendR = BendData.BendRadius
                            'If BendR > Thickness * 10.0# Then GoTo Ex
                            Fe(Fi).Agle = 180.0# - Math.Round(BendData.BendAngle * 180.0# / PI, 3)
                            Dim bRet As Boolean = GetBendFace(subFeat, Fe(Fi).Agle, Fe(Fi).Cface, Fe(Fi).Aface, Fe(Fi).SkPot)
                            If bRet Then Fi = Fi + 1 : ReDim Preserve Fe(Fi)
                        End If
Ex:                     subFeat = subFeat.GetNextSubFeature()
                    Loop
                End If
            Next


#Region "标注"
            Dim PointD() As Double, E1, E2 As Byte, IndxOP As Integer

            For ia = 0 To UBound(Fe) - 2
                For ib = ia + 1 To UBound(Fe) - 1
                    Drawing.ClearSelection2(True)
                    Select Case True
                        Case IsFace(Fe(ia).Aface(0), Fe(ib).Aface(0))
                            Fe(ia).Abyte = Fe(ia).Abyte + 2 : Fe(ib).Abyte = Fe(ib).Abyte + 2
                            E1 = getEntityForView(Fe(ia), 1) : E2 = getEntityForView(Fe(ib), 1)
                            PointD = getMidFacePinot(Fe(ia).Aface(0))

                        Case IsFace(Fe(ia).Aface(0), Fe(ib).Aface(1))
                            Fe(ia).Abyte = Fe(ia).Abyte + 2 : Fe(ib).Abyte = Fe(ib).Abyte + 1
                            E1 = getEntityForView(Fe(ia), 1) : E2 = getEntityForView(Fe(ib), 0)
                            PointD = getMidFacePinot(Fe(ia).Aface(0))

                        Case IsFace(Fe(ia).Aface(1), Fe(ib).Aface(0))
                            Fe(ia).Abyte = Fe(ia).Abyte + 1 : Fe(ib).Abyte = Fe(ib).Abyte + 2
                            E1 = getEntityForView(Fe(ia), 0) : E2 = getEntityForView(Fe(ib), 1)
                            PointD = getMidFacePinot(Fe(ia).Aface(1))

                        Case IsFace(Fe(ia).Aface(1), Fe(ib).Aface(1))
                            Fe(ia).Abyte = Fe(ia).Abyte + 1 : Fe(ib).Abyte = Fe(ib).Abyte + 1
                            E1 = getEntityForView(Fe(ia), 0) : E2 = getEntityForView(Fe(ib), 0)
                            PointD = getMidFacePinot(Fe(ia).Aface(1))

                        Case Else
                            E1 = 0 : E2 = 0
                    End Select

                    If E1 = 2 Then Fe(ia).SkPot.Select4(True, Nothing)
                    If E2 = 2 Then Fe(ib).SkPot.Select4(True, Nothing)
                    If E1 <> 0 And E2 <> 0 Then
                        Dim swDyDim As DisplayDimension = DrawingA.AddDimension2(PointD(0), PointD(1), 0#)
                        Dim swDim As Dimension = swDyDim.GetDimension2(0)
                        If E1 = 3 Or E2 = 3 Then
                            IndxOP = swDim.SetArcEndCondition(1, 3)
                            IndxOP = swDim.SetArcEndCondition(2, 3)
                        End If

                    End If
                Next
            Next

            '标注不匹配对象
            For ia = 0 To UBound(Fe) - 1
                Drawing.ClearSelection2(True)
                Dim B As Byte = Fe(ia).Abyte
                Select Case B
                    Case 0
                        E1 = getEntityForView(Fe(ia), 0)
                        E2 = getEntityEdgeForView(Fe(ia).Aface(1))
                        PointD = getMidFacePinot(Fe(ia).Aface(1))
                        If E1 = 2 Then Fe(ia).SkPot.Select4(True, Nothing)
                        If E1 <> 0 And E2 <> 0 Then
                            Dim swDyDim As DisplayDimension = DrawingA.AddDimension2(PointD(0), PointD(1), 0#)
                            Dim swDim As Dimension = swDyDim.GetDimension2(0)
                            If E1 = 3 Then
                                IndxOP = swDim.SetArcEndCondition(1, 3)
                                IndxOP = swDim.SetArcEndCondition(2, 3)
                            End If

                        End If

                        E1 = getEntityForView(Fe(ia), 1)
                        E2 = getEntityEdgeForView(Fe(ia).Aface(0))
                        PointD = getMidFacePinot(Fe(ia).Aface(0))
                        If E1 = 2 Then Fe(ia).SkPot.Select4(True, Nothing)
                        If E1 <> 0 And E2 <> 0 Then
                            Dim swDyDim As DisplayDimension = DrawingA.AddDimension2(PointD(0), PointD(1), 0#)
                            Dim swDim As Dimension = swDyDim.GetDimension2(0)
                            If E1 = 3 Then
                                IndxOP = swDim.SetArcEndCondition(1, 3)
                                IndxOP = swDim.SetArcEndCondition(2, 3)
                            End If
                        End If

                    Case 1
                        E1 = getEntityForView(Fe(ia), 1)
                        E2 = getEntityEdgeForView(Fe(ia).Aface(0))
                        PointD = getMidFacePinot(Fe(ia).Aface(0))
                        If E1 = 2 Then Fe(ia).SkPot.Select4(True, Nothing)
                        If E1 <> 0 And E2 <> 0 Then
                            Dim swDyDim As DisplayDimension = DrawingA.AddDimension2(PointD(0), PointD(1), 0#)
                            Dim swDim As Dimension = swDyDim.GetDimension2(0)
                            If E1 = 3 Then
                                IndxOP = swDim.SetArcEndCondition(1, 3)
                                IndxOP = swDim.SetArcEndCondition(2, 3)
                            End If
                        End If
                    Case 2
                        E1 = getEntityForView(Fe(ia), 0)
                        E2 = getEntityEdgeForView(Fe(ia).Aface(1))
                        PointD = getMidFacePinot(Fe(ia).Aface(1))
                        If E1 = 2 Then Fe(ia).SkPot.Select4(True, Nothing)
                        If E1 <> 0 And E2 <> 0 Then
                            Dim swDyDim As DisplayDimension = DrawingA.AddDimension2(PointD(0), PointD(1), 0#)
                            Dim swDim As Dimension = swDyDim.GetDimension2(0)
                            If E1 = 3 Then
                                IndxOP = swDim.SetArcEndCondition(1, 3)
                                IndxOP = swDim.SetArcEndCondition(2, 3)
                            End If
                        End If
                End Select
            Next
#End Region
            DrawingA.ClearSelection2(True)
Ex2:    Next
    End Sub
    '视图旋转
    Sub ViewRotate(Model As ModelDoc2)
        Dim Draw As DrawingDoc
        Dim View As View
        Dim Age As Double
        Draw = Model
        View = Draw.ActiveDrawingView
        If IsNothing(View) Then Exit Sub
        Age = View.Angle
        View.Angle = Age + PI / 2
    End Sub
    '为零件创建工程图
    Sub CreateDrw(Model As ModelDoc2, Optional Liew As Byte = 0)
        ReaderData()
        Dim TPme As String = SzmStr(0)
        If File.Exists(TPme) = False Then MsgBox("路径错误") : Exit Sub

        Dim Draw As DrawingDoc, DrawModel As ModelDoc2
        If Model.GetType <> swDocumentTypes_e.swDocPART Then Exit Sub
        Dim flieName As String = Model.GetPathName
        flieName = Left(flieName, Len(flieName) - 7)
        flieName = flieName & ".SLDDRW"

        If File.Exists(flieName) = False Then
CC:
            Draw = MySwApp.NewDocument(TPme, swDwgPaperSizes_e.swDwgPaperAsize, 0, 0)
            DrawModel = Draw
            Draw.Create1stAngleViews2(Model.GetPathName)
            DrawModel.SaveAs3(flieName, 0, 0)
        Else
            If Liew = 0 Then
                Dim Buttons As Integer = MsgBox("工程图存在文件中是否打开？", vbYesNo)
                If Buttons = 7 Then
                    If MsgBox("是否覆盖原有文件"， vbYesNo) = 6 Then
                        GoTo CC
                    Else
                        Exit Sub
                    End If
                End If
            End If

            Draw = MySwApp.OpenDoc6(flieName, 3, 0, "", Nothing, Nothing)
            DrawModel = Draw
            Dim Title As String = DrawModel.GetTitle
            MySwApp.ActivateDoc3(Title, False, 1, Nothing)

        End If

    End Sub
    '调整视图比例
    Sub ViewRatio（）

    End Sub
    '导出图纸到pdf
    Sub DrwToPdf()

    End Sub
    '删除悬空标注
    Sub DelDimensions()

    End Sub


#Region "法兰标注子函数"
    '判断bend是否与视图法向一致
    Function GetBendFace(Feat As Feature, Agle As Double, ByRef Cface As Face2, ByRef Aface() As Face2, ByRef skPoint As SketchPoint) As Boolean
        Dim By As Byte = 0
        ReDim Aface(1)
        Dim FaceArr = Feat.GetFaces
        For Each Face As Face2 In FaceArr
            Dim Surface As Surface = Face.GetSurface
            If Surface.IsCylinder Then
                Dim CyP = Surface.CylinderParams
                Dim BR As Double = CyP(6)
                If BR > (BendR + Thickness / 2) Then
                    '判断向量是否一致
                    Dim Edges = Face.GetEdges
                    For Each Edge As Edge In Edges
                        Dim CurveL As Curve = Edge.GetCurve
                        If CurveL.IsLine Then
                            If By = 0 Then
                                Dim Vle() As Double = CurveL.LineParams 'VarType(Vline) = 8197
                                Dim LVe() As Double = {Vle(3), Vle(4), Vle(5)}
                                If VectorsAreEqual(LVe) = True Then
                                    By = 1
                                Else
                                    Return False
                                End If
                                If Agle < 90.0# Then Cface = Face
                                By = 1
                            End If

                            Try
                                Dim ObFace = Edge.GetTwoAdjacentFaces2
                                For Each Face2 As Face2 In ObFace
                                    Dim SurF As Surface = Face2.GetSurface
                                    If SurF.IsPlane Then
                                        If Aface(0) Is Nothing Then
                                            Aface(0) = Face2
                                        Else
                                            Aface(1) = Face2
                                        End If
                                    End If
                                Next
                            Catch ex As Exception
                                Return False
                            End Try
                        End If
                    Next
                    If Aface(1) Is Nothing Then Aface(0) = Nothing : Return False
                    Select Case Agle
                        Case 0#
                            Return True
                        Case < 90.0#
                            DimensionAgle(Aface(0), Aface(1), 0)
                        Case > 90.0#
                            skPoint = DimensionAgle(Aface(0), Aface(1), 1)
                    End Select
                    Return True
                End If
            End If
        Next
    End Function

    Function GetPot(Face As Face2) As Object

        Dim dirArr(1, 1) As Double
        Dim Edges = Face.GetEdges
        For Each Edge As Edge In Edges
            Dim CurveL As Curve = Edge.GetCurve
            If CurveL.IsLine Then
                Dim Vle() As Double = CurveL.LineParams
                Dim LVe() As Double = {Vle(3), Vle(4), Vle(5)}
                If dirArr(0, 0) <> dirArr(1, 0) Or dirArr(0, 1) <> dirArr(1, 1) Then
                    Dim Dpot(1) As Double
                    Dpot(0) = (dirArr(0, 0) + dirArr(1, 0)) / 2
                    Dpot(1) = (dirArr(0, 1) + dirArr(1, 1)) / 2
                    Return Dpot

                    Dim swEnt As Entity = Edge
                    Dim bRet As Boolean = View.SelectEntity(swEnt, True)
                    Exit Function
                End If
            End If
        Next
    End Function

    Function GetPotTFx(Aface1 As Face2) As Object
        Dim Fx As Face2
        Dim Dpot(1) As Double
        Dim Edges As Object
        Dim Vedge As Object
        Dim Vline As Object
        Dim Edge As Edge

        Dim bRet As Boolean
        Dim Cl As Curve

        Dim startV As Object
        Dim endV As Object
        Dim endV1 As Object
        Dim startPt As Object
        Dim endPt As Object

        Dim Ls(1) As Double
        Dim L(2) As Double
        Dim Arr(1) As Double

        Dim vFaceP As Object
        Dim FArray As Object
        Dim PpF As Face2

        Fx = Aface1
        Edges = Fx.GetEdges
        For Each Vedge In Edges
            Edge = Vedge
            Cl = Edge.GetCurve
            Vline = Cl.LineParams
            If VarType(Vline) = 8197 Then

                vFaceP = Edge.GetTwoAdjacentFaces2
                PpF = vFaceP(0)
                FArray = PpF.Normal
                If FArray(0) = "0" And FArray(1) = "0" And FArray(2) = "0" Then GoTo nnx
                PpF = vFaceP(1)
                FArray = PpF.Normal
                If FArray(0) = "0" And FArray(1) = "0" And FArray(2) = "0" Then GoTo nnx

                startV = Edge.GetStartVertex
                startPt = startV.GetPoint
                endV = Edge.GetEndVertex
                endPt = endV.GetPoint
                Ls(0) = Math.Abs(startPt(0) - endPt(0)) + Math.Abs(startPt(1) - endPt(1)) + Math.Abs(startPt(2) - endPt(2))
                If Ls(0) > Ls(1) Then Ls(1) = Ls(0) : endV1 = endV
            End If
nnx:
        Next
        If endV1 Is Nothing Then Exit Function
        Dim swEnt As Entity = endV1
        bRet = View.SelectEntity(swEnt, True)

    End Function
    '标注角度
    Function DimensionAgle(F1 As Face2, F2 As Face2, BOpt As Byte) As SketchPoint
        Dim swEnt1, swEnt2 As Entity
        Dim Pot1(), Pot2() As Double

        Dim Edges = F1.GetEdges
        For Each Edge As Edge In Edges
            Dim CurveL As Curve = Edge.GetCurve
            If CurveL.IsLine Then
                Dim Vle() As Double = CurveL.LineParams
                Dim LVe() As Double = {Vle(3), Vle(4), Vle(5)}
                If VectorsAreEqual(LVe) = False Then
                    swEnt1 = Edge
                    Dim P0() As Double = Edge.GetStartVertex.GetPoint
                    Dim P1() As Double = Edge.GetEndVertex.GetPoint
                    Pot1 = {(P0(0) + P1(0)) / 2, (P0(1) + P1(1)) / 2, (P0(2) + P1(2)) / 2}
                    Exit For
                End If
            End If
        Next
        Edges = F2.GetEdges
        For Each Edge As Edge In Edges
            Dim CurveL As Curve = Edge.GetCurve
            If CurveL.IsLine Then
                Dim Vle() As Double = CurveL.LineParams
                Dim LVe() As Double = {Vle(3), Vle(4), Vle(5)}
                If VectorsAreEqual(LVe) = False Then
                    swEnt2 = Edge
                    Dim P0() As Double = Edge.GetStartVertex.GetPoint
                    Dim P1() As Double = Edge.GetEndVertex.GetPoint
                    Pot2 = {(P0(0) + P1(0)) / 2, (P0(1) + P1(1)) / 2, (P0(2) + P1(2)) / 2}
                    Exit For
                End If
            End If
        Next

        Dim Point() As Double = {(Pot1(0) + Pot2(0)) / 2, (Pot1(1) + Pot2(1)) / 2, (Pot1(2) + Pot2(2)) / 2}
        Dim MathPoint As MathPoint = swMathUtil.CreatePoint(Point).MultiplyTransform(ViewMathTransform)
        Dim Pot() As Double = MathPoint.ArrayData

        View.SelectEntity(swEnt1, False) : View.SelectEntity(swEnt2, True)
        Dim DisplayDim = DrawingA.AddDimension2(Pot(0), Pot(1), 0#)
        View.SelectEntity(swEnt2, True)
        If BOpt = 1 Then
            DrawingA.ClearSelection2(True)
            Dim skPoint As SketchPoint = DrawingA.SketchManager.CreatePoint(0, 0, 0)
            skPoint.Select4(False, Nothing)
            View.SelectEntity(swEnt1, True) : View.SelectEntity(swEnt2, True)
            DrawingA.SketchAddConstraints("sgATINTERSECT")
            Return skPoint
        End If
        Return Nothing
    End Function
    '返回视图标注对象
    Function getEntityForView(ByRef FeB As FeatC, ByRef B As Byte) As Byte

        Select Case FeB.Agle
            Case = 90
                Dim Face As Face2 = FeB.Aface(B)
                Dim Edges = Face.GetEdges
                For Each Edge As Edge In Edges
                    Dim Curve As Curve = Edge.GetCurve
                    If Curve.IsLine Then
                        Dim Vle() As Double = Curve.LineParams
                        Dim LVe() As Double = {Vle(3), Vle(4), Vle(5)}
                        If VectorsAreEqual(LVe) = False Then
                            Dim Enty As Entity = Edge
                            View.SelectEntity(Enty, True)
                            Return 1
                        End If
                    End If
                Next
            Case < 90
                Dim Face As Face2 = FeB.Cface
                Dim Edges = Face.GetEdges
                For Each Edge As Edge In Edges
                    Dim Curve As Curve = Edge.GetCurve
                    If Not Curve.IsLine Then
                        Dim Enty As Entity = Edge
                        View.SelectEntity(Enty, True)
                        Return 3
                    End If
                Next
                Return 0
            Case > 90
                Return 2
        End Select
        Return 0
    End Function
    '返回视图标注点（面的中点）
    Function getMidFacePinot(Face As Face2) As Object
        Dim Bl As Boolean = False
        Dim P0() As Double

        Dim Edges = Face.GetEdges
        For Each Edge As Edge In Edges
            Dim Curve As Curve = Edge.GetCurve
            If Curve.IsLine Then
                Dim Vle() As Double = Curve.LineParams
                Dim LVe() As Double = {Vle(3), Vle(4), Vle(5)}
                If VectorsAreEqual(LVe) = True Then
                    If Bl Then
                        Dim P1() As Double = Edge.GetStartVertex.GetPoint
                        Dim Pot() As Double = {(P0(0) + P1(0)) / 2, (P0(1) + P1(1)) / 2, (P0(2) + P1(2)) / 2}
                        Dim MathPoint As MathPoint = swMathUtil.CreatePoint(Pot).MultiplyTransform(ViewMathTransform)
                        Pot = MathPoint.ArrayData
                        Return Pot
                    Else
                        P0 = Edge.GetStartVertex.GetPoint
                    End If

                    Bl = True
                End If
            End If
        Next
        Return Nothing
    End Function
    '确定两个向量在1.0e-10的容差内是否相等
    Function VectorsAreEqual(Vle() As Double) As Boolean

        Dim db As Double = Vec(0) ^ 2 + Vec(1) ^ 2 + Vec(2) ^ 2 * Math.Sqrt(Vle(0) ^ 2 + Vle(1) ^ 2 + Vle(2) ^ 2)

        Dim dblDot As Double = (Vec(0) * Vle(0) + Vec(1) * Vle(1) + Vec(2) * Vle(2)) / db
        db = Math.Abs(Math.Abs(dblDot) - 1.0#)
        If db < 0.0000000001 Then '1.0e-10
            Return True
        Else
            Return False
        End If
    End Function
    '判断面是否匹配
    Function IsFace(F1 As Face2, F2 As Face2) As Boolean

        If F1.IsSame(F2) Then Return True
        Dim Norm1() As Double = F1.Normal : Dim Norm2() As Double = F2.Normal
        Dim bRet1 As Boolean = (Math.Abs(Norm1(0) - Norm2(0)) + Math.Abs(Norm1(1) - Norm2(1)) + Math.Abs(Norm1(2) - Norm2(2))) < 0.000000001
        Dim bRet2 As Boolean = (Math.Abs(Norm1(0) + Norm2(0)) + Math.Abs(Norm1(1) + Norm2(1)) + Math.Abs(Norm1(2) + Norm2(2))) < 0.000000001
        If Not (bRet2 Or bRet2) Then Return False

        Dim Edges = F1.GetEdges
        Dim Edge As Edge
        For Each Edge In Edges
            Dim CurveL As Curve = Edge.GetCurve
            If CurveL.IsLine Then Exit For
        Next

        Dim StartV As Vertex = Edge.GetStartVertex
        Edges = StartV.GetEdges

        For Each Edge In Edges
            Dim CurveL As Curve = Edge.GetCurve
            If CurveL.IsLine Then
                Dim Vle() As Double = CurveL.LineParams 'VarType(Vline) = 8197
                If Math.Abs(Vle(3) * Norm1(0) + Vle(4) * Norm1(1) + Vle(5) * Norm1(2)) > 0.000000001 Then
                    Dim Vext As Vertex = Edge.GetStartVertex
                    If MySwApp.IsSame(Vext, StartV) = 1 Then Vext = Edge.GetEndVertex

                    'Dim Pt1() As Double = Edge.GetStartVertex.GetPoint
                    'Dim Pt2() As Double = Edge.GetEndVertex.GetPoint
                    'Dim Lp As Double = (Pt1(0) - Pt2(0) + Pt1(1) - Pt2(1) + Pt1(2) - Pt2(2)) * 1000

                    Edges = Vext.GetEdges
                    For Each Edge2 As Edge In Edges
                        If MySwApp.IsSame(Edge2, Edge) = 0 Then
                            Dim twoFace = Edge2.GetTwoAdjacentFaces2
                            If twoFace(0).IsSame(F2) Or twoFace(1).IsSame(F2) Then Return True
                        End If
                    Next

                End If
            End If
        Next

    End Function
    '返回边缘法兰对象
    Function getEntityEdgeForView(Face As Face2) As Byte

        Dim Edget As Edge, Ls As Double

        Dim Edges = Face.GetEdges
        For Each Edge As Edge In Edges
            Dim CurveL As Curve = Edge.GetCurve
            If CurveL.IsLine Then
                Dim Vle() As Double = CurveL.LineParams : Dim LVe() As Double = {Vle(3), Vle(4), Vle(5)}
                If VectorsAreEqual(LVe) = True Then
                    Dim TwoFace = Edge.GetTwoAdjacentFaces2
                    For Each FtoF As Face2 In TwoFace
                        Dim Surface As Surface = FtoF.GetSurface
                        If Surface.IsCylinder Then GoTo EX
                    Next

                    Dim P1() As Double = Edge.GetStartVertex.Getpoint : Dim P2() As Double = Edge.GetEndVertex.Getpoint
                    Dim dot As Double = Math.Abs(P1(0) - P2(0)) + Math.Abs(P1(1) - P2(1)) + Math.Abs(P1(2) - P2(2))
                    If dot > Ls Then Ls = dot : Edget = Edge

                End If
            End If
EX:     Next

        If Not Edget Is Nothing Then
            Dim TwoFace = Edget.GetTwoAdjacentFaces2
            For Each Ft As Face2 In TwoFace
                If Not Ft.IsSame(Face) Then
                    Edges = Ft.GetEdges
                    For Each Edge As Edge In Edges
                        Dim CurveL As Curve = Edge.GetCurve
                        If CurveL.IsLine Then
                            Dim Vle() As Double = CurveL.LineParams : Dim LVe() As Double = {Vle(3), Vle(4), Vle(5)}
                            If VectorsAreEqual(LVe) = False Then
                                'Dim Vex As Vertex = Edget.GetStartVertex
                                Dim Enty As Entity = Edge
                                View.SelectEntity(Enty, True)
                                Return 1
                            End If
                        End If

                    Next
                End If
            Next
        End If

        Return 0
    End Function
#End Region
End Module
'批量替换模板  setupsheet6 方法
'调整视图
'文字替换