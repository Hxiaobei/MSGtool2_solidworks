
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Imports Microsoft.Office.Interop.Excel


Module Assem

    '按照到所选组件的距离选择其他组件
    Sub RangeSelection(Model As ModelDoc2)
        If Model.GetType <> 2 Then
            MySwApp.SendMsgToUser("请在装配体中执行")
        End If

        Dim SelMgr As SelectionMgr = Model.SelectionManager
        Dim SelData As SelectData = SelMgr.CreateSelectData


        Dim Entityf As Entity = SelMgr.GetSelectedObject6(1, -1)
        If Entityf Is Nothing Then MsgBox("请在设计树中选取一个组件") : Exit Sub
        Dim Comp1 As Component2 = Entityf.GetComponent

        If Comp1 Is Nothing Then MsgBox("请在设计树中选取一个组件") : Exit Sub

        Dim DistA As Double = InputBox("输入最小距离") / 1000.0#
        Dim AssyDoc As AssemblyDoc = Model
        Dim Components As Object = AssyDoc.GetComponents(True)
        For I = 0 To UBound(Components)
            Dim Comp2 As Component2 = Components(I)
            If Comp2.Name2 <> Comp1.Name2 Then
                Dim nDist As Double = Model.ClosestDistance(Comp1, Comp2, Nothing, Nothing)
                If DistA >= nDist Then Comp2.Select4(True, SelData, False)
            End If
        Next I
    End Sub
    '装配体打包改名
    Sub PackRename()

    End Sub
    '装配体设计树排序
    Sub SortByName(Model As ModelDoc2)
        If Model.GetType <> 2 Then MySwApp.SendMsgToUser("请打开装配体")
        Dim AssyDoc As AssemblyDoc = Model
        Dim SelMgr As SelectionMgr = Model.SelectionManager
        Dim Compos As Object = AssyDoc.GetComponents(True)
        Dim nCom As Integer = UBound(Compos)
        Dim sArray(nCom) As String, iArray(nCom) As Integer
        Dim Comp As Component2
        For i = 0 To nCom
            Comp = Compos(i)
            sArray(i) = Comp.Name2
            iArray(i) = i
        Next

        For a = 0 To (nCom - 1)
            For b = (a + 1) To nCom
                If sArray(a) > sArray(b) Then
                    Dim sTe As String = sArray(a)
                    Dim iTe As Integer = iArray(a)
                    sArray(a) = sArray(b)
                    iArray(a) = iArray(b)
                    sArray(b) = sTe
                    iArray(b) = iTe
                End If
            Next
        Next
        For i = (nCom - 1) To 0 Step -1
            AssyDoc.ReorderComponents(Compos(iArray(i)), Compos(iArray(i + 1)), 2)
        Next
    End Sub
    '装配体中虚拟零件保存命名
    Sub VirtualPartSave()

    End Sub
    '装配体改名（Excel）
    Sub AmRenameExl(Model As ModelDoc2)
        MySwApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swExtRefUpdateCompNames, False)
        Dim swAssyDoc As AssemblyDoc = Model
        Dim Components = swAssyDoc.GetComponents(False) '顶层true所有false
        Dim Name(3000, 2) As String
        Try
            For Each swComp As Component2 In Components
                Dim swModel As ModelDoc2 = swComp.GetModelDoc2
                If Not swModel Is Nothing Then
                    Dim TileName As String = swModel.GetTitle
                    TileName = Left(TileName, Len(TileName) - 7)
                    Dim PropMgr As CustomPropertyManager = swModel.Extension.CustomPropertyManager("")
                    Dim ReName As String
                    PropMgr.Get5("ReName", False, Nothing, ReName, Nothing)

                    If ReName <> "ReName" And ReName <> "" And InStr(swComp.Name2, ReName) = 0 Then
                        Dim Str0 As String = swModel.GetPathName
                        swComp.Select4(False, Nothing, False)
                        Model.Extension.RenameDocument(ReName) : swComp.Name2 = TileName
                        Dim Str1 As String = swModel.GetPathName
                        PropMgr.Add3("ReName", 30, "ReName", 2)
                        For i = 0 To UBound(Name)
                            If Name(i, 0) = "" Then
                                Name(i, 0) = Left(Str0, Len(Str0) - 7) & ".SLDDRW"
                                Name(i, 1) = Str0
                                Name(i, 2) = Str1
                                Exit For
                            End If
                        Next
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox("出错")
            Exit Sub
        End Try
        Model.Save3(2, 0, 0)

        For i = 0 To UBound(Name)
            If Name(i, 0) = "" Then Exit For
            Dim PathName As String = Name(i, 0)
            If IO.File.Exists(PathName) Then
                MySwApp.ReplaceReferencedDocument(PathName, Name(i, 1), Name(i, 2))
                Dim NewName As String = Left(Name(i, 2), Len(Name(i, 2)) - 7) & ".SLDDRW"
                IO.File.Move(PathName, NewName)
            End If

        Next
        UpdateComponentName(Model)
    End Sub
    '创建 子组件Boom表
    Sub CreateBoomA(Model As ModelDoc2, Template As String)

        Dim Comp As Component2 = Model.ConfigurationManager.ActiveConfiguration.GetRootComponent3(False)
        Dim RCtName As String = Comp.Name2
        Model.Extension.SelectByID2(RCtName, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        Model.HideComponent2()
        Model.ClearSelection2(True)

        Dim xlApp As Application = CreateObject("Excel.Application")
        Dim xlBook As Workbook = xlApp.Workbooks.Add(Template)
        Dim xlSheet As Worksheet = xlBook.Sheets(1)
        xlApp.Visible = False '设置可见

        Dim ChildComp As Object
        Dim CdComp As Component2
        Dim CName, CtName, Arr(2000) As String
        Dim i, ii As Integer
        Dim u, v, w, x As Integer
        Dim ia(2000) As Integer

        ChildComp = Comp.GetChildren
        For Each CdComp In ChildComp
            CName = CdComp.Name2
            CtName = CName & "@" & RCtName
            CName = Left(CName, InStrRev(CName, "-") - 1)
            For i = 0 To UBound(Arr)
                If Arr(i) = "" Then Exit For
                If CName = Arr(i) Then
                    ia(i) = ia(i) + 1
                    xlSheet.Cells(i + 4, 3) = ia(i) + 1
                    GoTo E10
                End If
            Next

            Arr(ii) = CName : ii = ii + 1

            Model.Extension.SelectByID2(CtName, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
            Model.ShowComponent2()
            Model.ClearSelection2(True)
            Model.ShowNamedView2("*等轴测", 7)
            Model.ViewZoomtofit2()

            Model.SaveAs3("C:\A0.JPG", 0, 2)

            Model.Extension.SelectByID2(CtName, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
            Model.HideComponent2()
            Model.ClearSelection2(True)

            xlSheet.Cells(ii + 3, 1) = "A" & ii
            xlSheet.Cells(ii + 3, 2) = CName
            xlSheet.Cells(ii + 3, 3) = 1
            u = xlSheet.Cells(ii + 3, 4).Top + 4
            x = xlSheet.Cells(ii + 3, 4).Height - 4
            v = xlSheet.Cells(ii + 3, 4).Left + 20
            w = xlSheet.Cells(ii + 3, 4).Width - 40
            xlApp.ActiveSheet.Shapes.AddPicture("C:\A0.JPG", True, True, v, u, w, x)
            Kill("C:\A0.JPG")
E10:    Next
        Dim Pth As String = "D:\works\excel 统计\" & RCtName & "清单.xlsx"
        xlApp.ActiveWorkbook.SaveAs(Pth)
        xlApp.ActiveWorkbook.Close()

        Model.Extension.SelectByID2(RCtName, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        Model.ShowComponent2()
        Model.ClearSelection2(True)
    End Sub
    '创建 装配体第一层级boom表
    Sub CreateBoomB(model As ModelDoc2, Template As String)
        Dim xlApp As Application = CreateObject("Excel.Application")
        Dim xlBook As Workbook = xlApp.Workbooks.Add(Template)
        Dim xlSheet As Worksheet = xlBook.Sheets(1)
        xlApp.Visible = False '设置可见
    End Sub
    '刷新组件名称
    Sub UpdateComponentName(Model As ModelDoc2)
        MySwApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swExtRefUpdateCompNames, False)
        Dim swAssyDoc As AssemblyDoc = Model
        Dim Components = swAssyDoc.GetComponents(False) '顶层true所有false
        For Each swComp As Component2 In Components
            Dim swModel As ModelDoc2 = swComp.GetModelDoc2
            If Not swModel Is Nothing Then
                Dim TileName As String = swModel.GetTitle
                TileName = Left(TileName, Len(TileName) - 7)
                If InStr(swComp.Name2, TileName) = 0 Then
                    swComp.Select4(False, Nothing, False)
                    swComp.Name2 = TileName
                End If
            End If
        Next
    End Sub
    '钣金参数修改
    Sub setsmt(Model As ModelDoc2)
        Dim SeleMgr As SelectionMgr = Model.SelectionManager
        'Dim selObj As SelectData = SelectionMgr.GetSelectedObject6(1, -1)
        If SeleMgr.GetSelectedObjectCount2(0) <> 1 Then Exit Sub
        Try
            Dim Face As Face2 = SeleMgr.GetSelectedObject6(1, -1)
            Dim Ent As Entity = Face
            Dim Comp As Component2 = Ent.GetComponent
            Dim part As ModelDoc2 = Comp.GetModelDoc2
            Dim Body As Body2 = Face.GetBody
            Dim feats As Object = Body.GetFeatures

            For i = 0 To UBound(feats)
                Dim Feat As Feature = feats(i)
                'Debug.Print(Feat.GetTypeName)
                If Feat.GetTypeName = "SheetMetal" Then

C1:                 Dim Dispdim As DisplayDimension = Feat.GetFirstDisplayDimension
                    Do While Not Dispdim Is Nothing
                        Dim DimAnn As Annotation = Dispdim.GetAnnotation
                        Dim DimName As String = DimAnn.GetName
                        If DimName = "D3" Or DimName = "D7" Or DimName = "Thickness" Or DimName = "厚度" Then
                            Dim Thickess As Double = InputBox("输入厚度") / 1000.0#
                            Dim Dimen As Dimension = Dispdim.GetDimension2(Nothing)
                            Dimen.SystemValue = Thickess
                            'Model.EditRebuild3()
                            LocalCorrection(part)
                            Model.EditRebuild3()
                            Exit Sub
                        End If
                        Dispdim = Feat.GetNextDisplayDimension(Dispdim)
                    Loop


                    Dim Parents As Object = Feat.GetParents
                    For Each Feat In Parents
                        'Debug.Print(Feat.GetTypeName)
                        If Feat.GetTypeName = "TemplateSheetMetal" Then
                            GoTo C1
                        End If
                    Next

                End If
            Next

        Catch ex As Exception

        End Try

    End Sub

End Module
