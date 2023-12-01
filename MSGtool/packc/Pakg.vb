Imports SolidWorks.Interop.sldworks
Imports System.Windows.Forms
Imports System.IO
Imports SolidWorks.Interop.swconst
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Public Class Pakg


    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim swModel As ModelDoc2 = MySwApp.ActiveDoc
        If swModel Is Nothing Then
            MySwApp.SendMsgToUser("当前没有任何文档打开，该程序必须在装配体中运行！")
            Exit Sub
        ElseIf swModel.GetType <> 2 Then
            MySwApp.SendMsgToUser("当前打开的文档不是一个装配体，请打开装配体后再试！")
            Exit Sub
        End If

        Dim swAssyDoc As AssemblyDoc = swModel
        Dim Components = swAssyDoc.GetComponents(False) '顶层true所有false

        For Each swComp As Component2 In Components
            Dim swPart As ModelDoc2 = swComp.GetModelDoc2
            If swPart Is Nothing Then GoTo Ec
            If swPart.GetType <> 1 Then GoTo Ec
            钣金特征命名(swPart)
Ec:     Next

    End Sub

    Private Sub 刷新表格_Click(sender As Object, e As EventArgs) Handles 刷新表格.Click
        Dim swModel As ModelDoc2 = MySwApp.ActiveDoc
        If swModel Is Nothing Then
            MySwApp.SendMsgToUser("当前没有任何文档打开，该程序必须在装配体中运行！")
            Exit Sub
        ElseIf swModel.GetType <> 2 Then
            MySwApp.SendMsgToUser("当前打开的文档不是一个装配体，请打开装配体后再试！")
            Exit Sub
        End If

        Dim swAssyDoc As AssemblyDoc = swModel
        Dim Components = swAssyDoc.GetComponents(False) '顶层true所有false


        ListView1.BeginUpdate()
        ListView1.Clear()
        Dim Colus0 As ColumnHeader = ListView1.Columns.Add("序号")
        Colus0.Width = 30
        Dim Colus1 As ColumnHeader = ListView1.Columns.Add("零件名称")
        Colus1.Width = 200
        Dim Colus2 As ColumnHeader = ListView1.Columns.Add("工程图")
        Colus2.Width = 25
        'Dim Colus3 As ColumnHeader = ListView1.Columns.Add("厚度")
        'Colus3.Width = 40
        Dim Colus4 As ColumnHeader = ListView1.Columns.Add("数量")
        Colus4.Width = 35
        'Dim Colus5 As ColumnHeader = ListView1.Columns.Add("折弯刀数")
        'Colus5.Width = 35

        Dim i As Integer
        For Each swComp As Component2 In Components
            Dim swPart As ModelDoc2 = swComp.GetModelDoc2
            If swPart Is Nothing Then GoTo Ec
            If swPart.GetType <> 1 Then GoTo Ec


            Dim Spath As String = swPart.GetPathName
            Spath = Strings.Left(Spath, Len(Spath) - 7)
            Dim STitle As String = swPart.GetTitle & " * " & swComp.ReferencedConfiguration

            'Columns Items
            For ii = 0 To ListView1.Items.Count - 1
                If ListView1.Items(ii).SubItems(1).Text = STitle Then
                    Dim ir As Integer
                    ir = ListView1.Items(ii).SubItems(3).Text
                    ListView1.Items(ii).SubItems(3).Text = ir + 1
                    GoTo Ec
                End If
            Next

            Dim Item As ListViewItem = ListView1.Items.Add(i)
            Item.SubItems.Add(swPart.GetTitle & " * " & swComp.ReferencedConfiguration)

            If File.Exists(Spath & ".SLDDRW") Then
                Item.SubItems.Add("存在")
            Else
                Item.SubItems.Add("")
            End If

            Item.SubItems.Add(1)
            i = i + 1
Ec:     Next

        ListView1.EndUpdate()
    End Sub

    Private Sub 打开零件_Click(sender As Object, e As EventArgs) Handles 打开零件.Click

        If ListView1.SelectedIndices.Count <> 1 Then Exit Sub
        Dim STitle As String = ListView1.SelectedItems(0).SubItems(1).Text
        Dim Sarr() As String = Split(STitle, " * ")
        Try
            Dim swPart As ModelDoc2 = MySwApp.ActivateDoc3(Sarr(0), False, 1, Nothing)
            swPart.ShowConfiguration2(Sarr(1))
        Catch ex As Exception
            MsgBox("零件没有再sw中打开")
        End Try

    End Sub

    Private Sub 打开工程图_Click(sender As Object, e As EventArgs) Handles 打开工程图.Click

        Dim Count As Integer = ListView1.SelectedIndices.Count
        For i = 0 To Count - 1
            Dim STitle As String = ListView1.SelectedItems(i).SubItems(1).Text
            Dim Sarr() As String = Split(STitle, " * ")
            STitle = Sarr(0)
            Try
                Dim ModelDoc As ModelDoc2 = MySwApp.GetOpenDocument(STitle)
                If Not ModelDoc Is Nothing Then CreateDrw(ModelDoc, 1)
            Catch ex As Exception
                MsgBox("零件没有再sw中打开")
            End Try

        Next

    End Sub

    Private Sub 选择_Click(sender As Object, e As EventArgs) Handles 选择.Click
        Dim swModel As ModelDoc2 = MySwApp.ActiveDoc
        If swModel Is Nothing Then
            MySwApp.SendMsgToUser("当前没有任何文档打开，该程序必须在装配体中运行！")
            Exit Sub
        ElseIf swModel.GetType <> 2 Then
            MySwApp.SendMsgToUser("当前打开的文档不是一个装配体，请打开装配体后再试！")
            Exit Sub
        End If

        Dim swAssyDoc As AssemblyDoc = swModel
        Dim Components = swAssyDoc.GetComponents(False) '顶层true所有false

        Dim Count As Integer = ListView1.SelectedIndices.Count
        For i = 0 To Count - 1
            Dim STitle As String = ListView1.SelectedItems(i).SubItems(1).Text

            For Each swComp As Component2 In Components
                Dim swPart As ModelDoc2 = swComp.GetModelDoc2
                If swPart Is Nothing Then GoTo Ea
                If swPart.GetType <> 1 Then GoTo Ea
                If STitle = swPart.GetTitle & " * " & swComp.ReferencedConfiguration Then
                    Try
                        swComp.Select4(True, Nothing, False)
                    Catch ex As Exception

                    End Try
                End If
Ea:         Next
        Next
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Dim swModel As ModelDoc2 = MySwApp.ActiveDoc
        If swModel Is Nothing Then
            MySwApp.SendMsgToUser("当前没有任何文档打开，该程序必须在装配体中运行！")
            Exit Sub
        ElseIf swModel.GetType <> 2 Then
            MySwApp.SendMsgToUser("当前打开的文档不是一个装配体，请打开装配体后再试！")
            Exit Sub
        End If

        Dim swAssyDoc As AssemblyDoc = swModel
        Dim Components = swAssyDoc.GetComponents(False)

        For Each ChildComp As Component2 In Components
            Dim PartSub As ModelDoc2 = ChildComp.GetModelDoc2
            If PartSub Is Nothing Then GoTo E10
            swToosl.DeletesCustprop(PartSub)
            swToosl.DeletesCus(PartSub)
E10:    Next
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click

        Dim swModel As ModelDoc2 = MySwApp.ActiveDoc
        If swModel Is Nothing Then
            MySwApp.SendMsgToUser("当前没有任何文档打开，该程序必须在装配体中运行！")
            Exit Sub
        ElseIf swModel.GetType <> 2 Then
            MySwApp.SendMsgToUser("当前打开的文档不是一个装配体，请打开装配体后再试！")
            Exit Sub
        End If


        Dim arr(1, 3000) As String
        Dim swAssyDoc As AssemblyDoc = swModel
        Dim Components = swAssyDoc.GetComponents(False)

        '提取组件
        For Each ChildComp As Component2 In Components
            Dim PartSub As ModelDoc2 = ChildComp.GetModelDoc2
            If PartSub Is Nothing Then GoTo E10
            Dim PropMgr As CustomPropertyManager = PartSub.Extension.CustomPropertyManager("")
            PropMgr.Add3("ReName", 30, "ReName", 2)

            If PartSub.GetType <> swDocumentTypes_e.swDocPART Then GoTo E10

            Dim ConfigName As String = ChildComp.ReferencedConfiguration
            Dim Srtname As String = PartSub.GetPathName & "*" & ConfigName
            Dim cusPropMgr As CustomPropertyManager = PartSub.Extension.CustomPropertyManager(ConfigName)
            '写属性
            For ii = 0 To UBound(arr, 2) Step 1
                If arr(0, ii) = "" Then Exit For
                If Srtname = arr(0, ii) Then
                    arr(1, ii) = arr(1, ii) + 1
                    cusPropMgr.Add3("Amount", 30, arr(1, ii), 2)
                    GoTo E10
                End If
            Next
            For ii = 0 To UBound(arr, 2)
                If arr(0, ii) = "" Then
                    arr(0, ii) = Srtname
                    arr(1, ii) = 1
                    cusPropMgr.Add3("Amount", 30, arr(1, ii), 2)
                    cusPropMgr.Add3("Material", 30, Chr(34) & "SW-Material@零件.SLDPRT" & Chr(34), 2)
                    cusPropMgr.Add3("Thickness", 30, Chr(34) & "厚度@钣金@零件.SLDPRT" & Chr(34), 2)
                    'cusPropMgr.Add3("Option", 30, 1, 2) '1 草图 
                    'cusPropMgr.Add3("L3", 30, Chr(34) & " LS@零件.SLDPRT" & Chr(34), 2)
                    Exit For
                End If
            Next

E10:    Next

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim SwModel As ModelDoc2 = MySwApp.ActiveDoc
        If swModel Is Nothing Then
            MySwApp.SendMsgToUser("当前没有任何文档打开，该程序必须在装配体中运行！")
            Exit Sub
        ElseIf swModel.GetType <> 2 Then
            MySwApp.SendMsgToUser("当前打开的文档不是一个装配体，请打开装配体后再试！")
            Exit Sub
        End If

    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        Dim myForm As batch = New batch
        myForm.Show()
    End Sub

    Private Sub 更新展开图ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 更新展开图ToolStripMenuItem.Click

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) 
        Dim swModel As ModelDoc2 = MySwApp.ActiveDoc
        If swModel Is Nothing Then
            MySwApp.SendMsgToUser("当前没有任何文档打开，该程序必须在装配体中运行！")
            Exit Sub
        ElseIf swModel.GetType <> 2 Then
            MySwApp.SendMsgToUser("当前打开的文档不是一个装配体，请打开装配体后再试！")
            Exit Sub
        End If

        Dim swAssyDoc As AssemblyDoc = swModel
        Dim Components = swAssyDoc.GetComponents(False) '顶层true所有false

        For Each swComp As Component2 In Components
            Dim swPart As ModelDoc2 = swComp.GetModelDoc2
            If swPart Is Nothing Then GoTo Ec
            If swPart.GetType <> 1 Then GoTo Ec
            LocalCorrection(swPart)
Ec:     Next
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) 
        Dim swModel As ModelDoc2 = MySwApp.ActiveDoc
        If swModel Is Nothing Then
            MySwApp.SendMsgToUser("当前没有任何文档打开，该程序必须在装配体中运行！")
            Exit Sub
        ElseIf swModel.GetType <> 2 Then
            MySwApp.SendMsgToUser("当前打开的文档不是一个装配体，请打开装配体后再试！")
            Exit Sub
        End If

        Dim swAssyDoc As AssemblyDoc = swModel
        Dim Components = swAssyDoc.GetComponents(False) '顶层true所有false

        For Each swComp As Component2 In Components
            Dim swPart As ModelDoc2 = swComp.GetModelDoc2
            If swPart Is Nothing Then GoTo Ec
            If swPart.GetType <> 1 Then GoTo Ec
            GlobalCorrection(swPart)
Ec:     Next
    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        pasteCAD()
        'CreateBody()
    End Sub
    '导出属性
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim Model As ModelDoc2 = MySwApp.ActiveDoc
        If Model Is Nothing Then
            MySwApp.SendMsgToUser("当前没有任何文档打开，该程序必须在装配体中运行！")
            Exit Sub
        ElseIf Model.GetType <> 2 Then
            MySwApp.SendMsgToUser("当前打开的文档不是一个装配体，请打开装配体后再试！")
            Exit Sub
        End If

        Dim AssyDoc As AssemblyDoc = Model
        Dim Components = AssyDoc.GetComponents(False) '顶层true所有false

        For Each swComp As Component2 In Components
            Dim swPart As ModelDoc2 = swComp.GetModelDoc2
            If swPart Is Nothing Then GoTo Ec
            If swPart.GetType <> 1 Then GoTo Ec
            LocalCorrection(swPart)
Ec:     Next
    End Sub
    Private Function getChildComp(Comp As Component2, nLeve As Long)
        Dim sPadStr As String
        For i = 0 To nLeve
            sPadStr = sPadStr & "  "
        Next




        Dim vCC As Object = Comp.GetChildren
        For Each childComp As Component2 In vCC
            getChildComp(childComp, nLeve + 1)
        Next

    End Function

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub 钣金零件ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 钣金零件ToolStripMenuItem.Click
        Dim swModel As ModelDoc2 = MySwApp.ActiveDoc
        If swModel Is Nothing Then
            MySwApp.SendMsgToUser("当前没有任何文档打开，该程序必须在装配体中运行！")
            Exit Sub
        ElseIf swModel.GetType <> 2 Then
            MySwApp.SendMsgToUser("当前打开的文档不是一个装配体，请打开装配体后再试！")
            Exit Sub
        End If

        Dim swAssyDoc As AssemblyDoc = swModel
        Dim Components = swAssyDoc.GetComponents(False) '顶层true所有false


        ListView1.BeginUpdate()
        ListView1.Clear()
        Dim Colus0 As ColumnHeader = ListView1.Columns.Add("序号")
        Colus0.Width = 30
        Dim Colus1 As ColumnHeader = ListView1.Columns.Add("零件名称")
        Colus1.Width = 200
        Dim Colus2 As ColumnHeader = ListView1.Columns.Add("工程图")
        Colus2.Width = 25
        Dim Colus3 As ColumnHeader = ListView1.Columns.Add("厚度")
        Colus3.Width = 40
        Dim Colus4 As ColumnHeader = ListView1.Columns.Add("数量")
        Colus4.Width = 35
        Dim Colus5 As ColumnHeader = ListView1.Columns.Add("折弯刀数")
        Colus5.Width = 35

        Dim i As Integer
        For Each swComp As Component2 In Components
            Dim swPart As ModelDoc2 = swComp.GetModelDoc2
            If swPart Is Nothing Then GoTo Ec
            If swPart.GetType <> 1 Then GoTo Ec

            Dim Feat As Feature
            Dim nn As Integer = swPart.FeatureManager.GetFeatureCount(True)
            For ii = nn - 18 To 1 Step -1
                Feat = swPart.FeatureByPositionReverse(ii)
                If Feat.GetTypeName = "SheetMetal" Then GoTo Eb
            Next
            GoTo Ec

Eb:         Dim Spath As String = swPart.GetPathName
            Spath = Strings.Left(Spath, Len(Spath) - 7)
            Dim STitle As String = swPart.GetTitle & " * " & swComp.ReferencedConfiguration

            'Columns Items
            For ii = 0 To ListView1.Items.Count - 1
                If ListView1.Items(ii).SubItems(1).Text = STitle Then
                    Dim ir As Integer
                    ir = ListView1.Items(ii).SubItems(4).Text
                    ListView1.Items(ii).SubItems(4).Text = ir + 1
                    GoTo Ec
                End If
            Next

            Dim Item As ListViewItem = ListView1.Items.Add(i)
            Item.SubItems.Add(swPart.GetTitle & " * " & swComp.ReferencedConfiguration)

            If File.Exists(Spath & ".SLDDRW") Then
                Item.SubItems.Add("存在")
            Else
                Item.SubItems.Add("")
            End If

            Dim SheetMetal As SheetMetalFeatureData = Feat.GetDefinition
            Dim Thickess As Single = Math.Round(SheetMetal.Thickness, 8) * 1000.0#
            Item.SubItems.Add(Thickess)

            Item.SubItems.Add(1)

            Dim BendCount As Integer = 0
            Dim ChFeat = Feat.GetChildren
            For Each flFeat As Feature In ChFeat
                If flFeat.GetTypeName = "FlatPattern" Then
                    Dim subFeat As Feature = flFeat.GetFirstSubFeature
                    While Not subFeat Is Nothing
                        If subFeat.GetTypeName = "UiBend" Then BendCount = BendCount + 1
                        subFeat = subFeat.GetNextSubFeature
                    End While
                    Exit For
                End If
            Next
            Item.SubItems.Add(BendCount)
            i = i + 1
Ec:     Next

        ListView1.EndUpdate()
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Dim Model As ModelDoc2 = MySwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub
        swToosl.草图定位(Model)

    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Dim Model As ModelDoc2 = MySwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub
        If Model.GetType <> 1 Then Exit Sub
        swToosl.SketchConstrainAll(Model)
    End Sub


    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        '可以集成输出面的功能
        Dim Model As ModelDoc2 = MySwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub
        Try
            Dim Spath As String = Model.GetPathName
            Spath = Strings.Left(Spath, Len(Spath) - 7)
            Dim ss As Integer = Model.GetType
            Select Case Model.GetType

                Case 1 To 2
                    Dim DraDoc As DrawingDoc = MySwApp.NewDocument(dllPath & "Sm.drwdot", 12, 0, 0)
                    Dim myView As SolidWorks.Interop.sldworks.View = DraDoc.CreateDrawViewFromModelView3(Model.GetPathName, "", 0, 0, 0)
                    DraDoc.SaveAs3(Spath & "__.DWG", 0, 2)
                    MySwApp.CloseDoc("")
                    DraDoc = Nothing
                Case = 3
                    Model.SaveAs3(Spath & ".DWG", 0, 2)
            End Select
        Catch ex As Exception
            MsgBox("请保存文件后重试")
        End Try

    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click

        Dim ia As Integer = 0
        Do While File.Exists("C:\Body\" & ia & "Bs.body")
            File.Delete("C:\Body\" & ia & "Bs.body")
            ia += 1
        Loop

        Dim Model As ModelDoc2 = MySwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub
        If Model.GetType <> 1 Then Exit Sub

        Dim Part As PartDoc = Model
        Dim SelMgr As SelectionMgr = Model.SelectionManager
        Dim Count As Integer = SelMgr.GetSelectedObjectCount2(-1)

        If Count > 0 Then
            For i = 1 To Count
                Dim Face As Face2
                Try
                    Face = SelMgr.GetSelectedObject6(i, -1)
                Catch ex As Exception
                    GoTo Ext
                End Try
                Dim Bodyi As Body2 = Face.GetBody
                'Bodyi = Bodyi.Copy2(False)
                Dim stream As IStream = Nothing
                CreateStreamOnHGlobal(IntPtr.Zero, True, stream)

                Bodyi.Save(stream)
                Dim comStream = New ComStream(stream, False, False)
                Directory.CreateDirectory("C:\Body")
                Using fileStream = File.Create("C:\Body\" & i - 1 & "Bs.body")
                    comStream.Seek(0, SeekOrigin.Begin)
                    comStream.CopyTo(fileStream)
                End Using
            Next
        Else
Ext:        Dim vBodies As Object = Part.GetBodies2(-1, False)
            For i = 0 To UBound(vBodies)
                Dim Bodyi As Body2 = vBodies(i)
                'Bodyi = Bodyi.Copy2(false)
                Dim stream As IStream = Nothing
                CreateStreamOnHGlobal(IntPtr.Zero, True, stream)

                Bodyi.Save(stream)
                Dim comStream = New ComStream(stream, False, False)
                Directory.CreateDirectory("C:\Body")
                Using fileStream = File.Create("C:\Body\" & i & "Bs.body")
                    comStream.Seek(0, SeekOrigin.Begin)
                    comStream.CopyTo(fileStream)
                End Using
            Next
        End If

        '提取模型中的实体

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

        Dim Model As ModelDoc2 = MySwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub
        If Model.GetType <> 1 Then Exit Sub

        Try
            Model.Save3(1, Nothing, Nothing)
        Catch ex As Exception
            MsgBox("请保存文件后重试")
            Exit Sub
        End Try

        Dim Part As PartDoc = Model
        Dim vBodies As Object = Part.GetBodies2(-1, False)
        Dim Bodys() As Body2
        ReDim Bodys(UBound(vBodies))

        For i = 0 To UBound(vBodies)
            Dim Bodyi As Body2 = vBodies(i)
            Bodys(i) = Bodyi.Copy2(True)
        Next

        Dim Feature As Feature = Model.FirstFeature()
        While Feature IsNot Nothing
            Feature.Select2(True, 1)
            Feature = Feature.GetNextFeature()
        End While
        Model.Extension.DeleteSelection2(1)
        Model.Save3(1, Nothing, Nothing)
        For i = 0 To UBound(Bodys)
            Model.CreateFeatureFromBody3(Bodys(i), False, 2)
        Next
        '提取模型中的实体
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click

        Dim Model As ModelDoc2 = MySwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub
        If Model.GetType <> 1 Then Exit Sub

        Dim i As Integer = 0

        Do While File.Exists("C:\Body\" & i & "Bs.body")
            ' = BodiesA(i)
            Dim stream As IStream = Nothing
            CreateStreamOnHGlobal(IntPtr.Zero, True, stream)
            Dim comStream = New ComStream(stream, True, True)
            Using fileStream = File.OpenRead("C:\Body\" & i & "Bs.body")
                fileStream.CopyTo(comStream)
                comStream.Seek(0, SeekOrigin.Begin)
            End Using
            Dim Modeler As IModeler = MySwApp.IGetModeler()
            Dim Bodyi As Body2 = modeler.Restore(stream)

            Model.CreateFeatureFromBody3(Bodyi, False, 2)
            i = i + 1
        Loop


    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Dim Model As ModelDoc2 = MySwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub
        'If Model.GetType <> 1 Then Exit Sub
        Try
            Dim Spath As String = Model.GetPathName
            Spath = Strings.Left(Spath, Len(Spath) - 7) & ".pdf"
            Model.SaveAs3(Spath, 0, 2)
        Catch ex As Exception
            MsgBox("请保存文件后重试")
        End Try

    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        Dim Model As ModelDoc2 = MySwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub
        If Model.GetType <> 1 Then Exit Sub
        DimFlange(model)
    End Sub


    Private Sub Clickzz(sender As Object, e As EventArgs) Handles TreeView1.NodeMouseDoubleClick

        'If Not getSwApp() Then Exit Sub


        'mainDS()
        '孤立编辑()
    End Sub
    '设计树显示入口
    Private Sub InitTree(myModel As ModelDoc2) 'TreeView1

        Dim FeatureMgr As FeatureManager = myModel.FeatureManager
        Dim rootNode As TreeControlItem = FeatureMgr.GetFeatureTreeRootItem2(1)

        TreeView1.BeginUpdate()
        TreeView1.Nodes.Clear()

        Dim Node1 As TreeNode = getTreeView(rootNode)
        For i = 0 To Node1.Nodes.Count - 1
            TreeView1.Nodes.Add(Node1.Nodes.Item(i))
        Next

        TreeView1.EndUpdate()
    End Sub
    Private Function getTreeView(TrreNode As TreeControlItem) As TreeNode


        Dim nodes As TreeNode = New TreeNode()
        nodes.Text = TrreNode.Text

        'Dim indents As TreeNodeCollection = TreeView1.Nodes().Add(TrreNode.Text)

        Dim childNode As TreeControlItem

        childNode = TrreNode.GetFirstChild()

        While Not childNode Is Nothing
            Dim tree As TreeNode = getTreeView(childNode)
            nodes.Nodes.Add(tree)
            'For i = 0 To tree.Nodes.Count - 1
            '    nodes.Nodes.Add(tree.Nodes.Item(i))
            'Next

            childNode = childNode.GetNext
        End While

        Return nodes
    End Function

    Private Sub Button32_Click(sender As Object, e As EventArgs) Handles Button32.Click
        '添加外部辅助工具 替换模板，格式转换，文件命名，文件处理（随文件层级复制某类型文件），替换文字（sw工程图）
    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles Button8.Click
        '合适视图到图纸最大范围
    End Sub

    Private Sub Clickzz(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseDoubleClick

    End Sub

    Private Sub 刷新ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 刷新ToolStripMenuItem.Click
        Dim Model As ModelDoc2 = MySwApp.ActiveDoc
        Dim aModel As ModelDoc2

        Select Case Model.GetType
            Case 1
                aModel = Model
            Case 2

                Dim SelMgr As SelectionMgr = Model.SelectionManager

            Case Else
                Exit Sub
        End Select

        InitTree(aModel)
    End Sub

    Private Sub Button33_Click(sender As Object, e As EventArgs) Handles Button33.Click
        If Not getSwApp() Then Exit Sub
        Dim myModel As ModelDoc2 = MySwApp.ActiveDoc
        Dim Flis As String = ".X_T"
        If Ra1.Checked = True Then
            Flis = ".STP"
        End If

        myModel.Extension.SaveAs2("C:\Body\OutData\000" & Flis, 0, 2, Nothing, "", False, Nothing, Nothing)

        File.Copy("C:\Body\OutData\000" & Flis, "C:\Body\OutData\000.txt")
        Kill("C:\Body\OutData\000" & Flis)
        Rename("C:\Body\OutData\000.txt", "C:\Body\OutData\000" & Flis)

    End Sub

    Private Sub Button34_Click(sender As Object, e As EventArgs) Handles Button34.Click
        'If Not getSwApp() Then Exit Sub
        'MySwApp.Visible = False
        Dim myModel As ModelDoc2 = MySwApp.ActiveDoc
        If myModel.GetType <> 1 Then Exit Sub

        Dim SelBool As Boolean = True
        Dim SelFeat As Feature
        Try
            Dim SelMgr As SelectionMgr = myModel.SelectionManager
            SelFeat = SelMgr.GetSelectedObject6(1, -1)
            If SelFeat.GetTypeName <> "BaseBody" Then SelBool = False
        Catch ex As Exception
            SelBool = False
        End Try

        If SelBool Then
            If MsgBox("是否执行替换体"， vbYesNo) = 6 Then
                SelBool = True
            Else
                SelBool = False
            End If
        End If

        Dim Flis As String = ".X_T"
        If Ra1.Checked = True Then
            Flis = ".STP"
        End If
        Dim ImportStepData As ImportStepData = MySwApp.GetImportFileData(Flis)
        Dim ImModel As ModelDoc2 = MySwApp.LoadFile4("C:\Body\OutData\000" & Flis, "", ImportStepData, Nothing)
        Dim ImPart As PartDoc = ImModel
        Dim vBodies As Object = ImPart.GetBodies2(-1, False)

        If vBodies Is Nothing Then GoTo Er1
        If UBound(vBodies) > 1 And SelBool Then
            MsgBox("输入文件有多个体")
            GoTo Er1
        End If

        If SelBool Then

            Dim Selbody As Body2 = vBodies(0)
            Selbody = Selbody.Copy
            SelFeat.ISetBody3(Selbody, False)

        Else
            For Each bodyi As Body2 In vBodies
                Dim Part As PartDoc = myModel
                'Part.InsertImportedFeature("C:\Body\OutData\000" & Flis, Nothing)
                myModel.CreateFeatureFromBody3(bodyi.Copy2(False), False, 2)
            Next
        End If

Er1:    MySwApp.CloseDoc(ImModel.GetTitle)

    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        Dim Model As ModelDoc2 = MySwApp.ActiveDoc
        If Model Is Nothing Then Exit Sub
        If Model.GetType <> 2 Then Exit Sub

        Dim SelMgr As SelectionMgr = Model.SelectionManager
        Dim count As Integer = SelMgr.GetSelectedObjectCount2(0)

        For i = 0 To count - 1
            Dim comp As Component2 = SelMgr.GetSelectedObjectsComponent4(i + 1, 0)
            Dim Models As ModelDoc2 = comp.GetModelDoc2
            If Models IsNot Nothing Then

            End If
        Next



    End Sub

    Private Sub Button35_Click(sender As Object, e As EventArgs) Handles Button35.Click
        getSwApp()
        CavityTriadManipulatorHandler.CTriadManipulator()
    End Sub
End Class

'MySwApp.OpenDoc6("", 1, 0, "", Nothing, Nothing)
'.EPRT