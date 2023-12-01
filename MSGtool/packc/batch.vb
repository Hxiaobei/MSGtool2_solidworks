Imports SolidWorks.Interop.sldworks
Imports System.Windows.Forms
Imports System.Windows
Imports System.IO
Imports SolidWorks.Interop.swconst
Public Class batch
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
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
        Dim Colus0 As ColumnHeader = ListView1.Columns.Add("文件名")
        Colus0.Width = 260
        Dim Colus1 As ColumnHeader = ListView1.Columns.Add("数量")
        Colus1.Width = 30
        Dim Colus2 As ColumnHeader = ListView1.Columns.Add("处理结果")
        Colus2.Width = 100

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

Eb:         Dim STitle As String = swPart.GetPathName & " * " & swComp.ReferencedConfiguration

            'Columns Items
            For ii = 0 To ListView1.Items.Count - 1
                If ListView1.Items(ii).SubItems(0).Text = STitle Then
                    Dim ir As Integer = ListView1.Items(ii).SubItems(1).Text
                    ListView1.Items(ii).SubItems(1).Text = ir + 1
                    GoTo Ec
                End If
            Next
            Dim File As String = swPart.GetPathName & " * " & swComp.ReferencedConfiguration
            Dim Item As ListViewItem = ListView1.Items.Add(File)
            Item.SubItems.Add(1)
            Item.SubItems.Add("")

Ec:     Next


        ListView1.EndUpdate()
        MySwApp.CloseDoc(swModel.GetTitle)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        OpenFileDialog1.Filter = "零件(*.SLDPRT)|*.SLDPRT"

        If OpenFileDialog1.ShowDialog() = Forms.DialogResult.OK Then
            Dim FileNames() As String = OpenFileDialog1.FileNames
            ListView1.BeginUpdate()
            ListView1.Clear()
            Dim Colus0 As ColumnHeader = ListView1.Columns.Add("文件名")
            Colus0.Width = 260
            Dim Colus1 As ColumnHeader = ListView1.Columns.Add("处理结果")
            Colus1.Width = 300
            For Each file As String In FileNames
                Dim Item As ListViewItem = ListView1.Items.Add(file)
                Item.SubItems.Add("")
            Next

            ListView1.EndUpdate()
        End If
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click


        Dim swModel As ModelDoc2 = MySwApp.ActiveDoc
        If Not swModel Is Nothing Then MySwApp.SendMsgToUser("为了程序稳定运行请关闭所有打开文件！") : Exit Sub
        Dim bRet As Boolean = ListView1.Columns.Count > 2
        Dim path As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop)
        Dim bRet1(1) As Boolean
        bRet1(0) = MySwApp.GetUserPreferenceToggle(8) : bRet1(1) = MySwApp.GetUserPreferenceToggle(86)
        MySwApp.SetUserPreferenceToggle(8, False) : MySwApp.SetUserPreferenceToggle(86, False)

        Using swr As StreamWriter = New StreamWriter(path & "\MSGtool 导出dxf处理结果.txt")
            'swModel = Nothing
            ListView1.BeginUpdate()
            For ii = 0 To ListView1.Items.Count - 1
                Dim STitle As String = ListView1.Items(ii).SubItems(0).Text
                Dim ir As Integer
                If bRet Then ir = ListView1.Items(ii).SubItems(1).Text

                Dim SrOb = Split(STitle, " * ")
                STitle = SrOb(0)
                Dim swDocSpecification As DocumentSpecification = MySwApp.GetOpenDocSpec(STitle)
                If UBound(SrOb) > 1 Then
                    Dim Config As String = SrOb(1)
                    swDocSpecification.ConfigurationName = Config
                End If
                swModel = MySwApp.OpenDoc7(swDocSpecification)
                'DraDoc
                If swModel Is Nothing Then
                    GoTo ER
                Else

                    Dim Err() As String = ExportToCAD(swModel, ir, Nothing)
                    If bRet Then
                        ListView1.Items(ii).SubItems(2).Text = Err(2) & " " & Err(3) & " " & Err(4) & " " & Err(5)
                    Else
                        ListView1.Items(ii).SubItems(1).Text = Err(2) & " " & Err(3) & " " & Err(4) & " " & Err(5)
                    End If
                    Dim R As String = Err(1) & " " & Err(2) & " " & Err(3) & " " & Err(4) & " " & Err(5)
                    swr.WriteLine(swModel.GetTitle & "  " & R)
                    MySwApp.CloseDoc(swModel.GetTitle)
                    swModel = Nothing
                End If
                ListView1.EndUpdate()
ER:         Next
            'MySwApp.Visible = True
        End Using

        MySwApp.SetUserPreferenceToggle(8, bRet1(0))
        MySwApp.SetUserPreferenceToggle(86, bRet1(1))
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
    End Sub

    Private Sub 打开文件ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 打开文件ToolStripMenuItem.Click
        Dim Count As Integer = ListView1.SelectedIndices.Count
        For i = 0 To Count - 1
            Dim STitle As String = ListView1.SelectedItems(i).SubItems(0).Text
            Dim Sarr() As String = Split(STitle, " * ")
            STitle = Sarr(0)
            Try
                Dim SrOb = Split(STitle, " * ")
                STitle = SrOb(0)
                Dim swDocSpecification As DocumentSpecification = MySwApp.GetOpenDocSpec(STitle)
                If UBound(SrOb) > 1 Then
                    Dim Config As String = SrOb(1)
                    swDocSpecification.ConfigurationName = Config
                End If
                MySwApp.OpenDoc7(swDocSpecification)
            Catch ex As Exception
                MsgBox("错误")
            End Try

        Next
    End Sub

    Private Sub 从列表中删除ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 从列表中删除ToolStripMenuItem.Click
        Dim Count As Integer = ListView1.SelectedIndices.Count
        ListView1.BeginUpdate()
        For i = Count - 1 To 0 Step -1
            ListView1.SelectedItems(i).Remove()
        Next
        ListView1.EndUpdate()
    End Sub

    Private Sub 清空列表ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 清空列表ToolStripMenuItem.Click
        ListView1.Clear()
        Dim Colus0 As ColumnHeader = ListView1.Columns.Add("文件名")
        Colus0.Width = 300
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim swModel As ModelDoc2 = MySwApp.ActiveDoc
        If Not swModel Is Nothing Then MySwApp.SendMsgToUser("为了程序稳定运行请关闭所有打开文件！") : Exit Sub
        Dim DraDoc As DrawingDoc = MySwApp.NewDocument(dllPath & "Sm.drwdot", 12, 0, 0)
        Dim path As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop)
        Dim bRet1(1) As Boolean
        bRet1(0) = MySwApp.GetUserPreferenceToggle(8) : bRet1(1) = MySwApp.GetUserPreferenceToggle(86)
        MySwApp.SetUserPreferenceToggle(8, False) : MySwApp.SetUserPreferenceToggle(86, False)

        Using swr As StreamWriter = New StreamWriter(path & "\MSGtool 导出展开处理结果.txt")
            'swModel = Nothing
            ListView1.BeginUpdate()
            For ii = 0 To ListView1.Items.Count - 1
                Dim STitle As String = ListView1.Items(ii).SubItems(0).Text
                Dim SrOb = Split(STitle, " * ") ： STitle = SrOb(0)

                Dim swDocSpecification As DocumentSpecification = MySwApp.GetOpenDocSpec(STitle)
                If UBound(SrOb) > 1 Then
                    Dim Config As String = SrOb(1)
                    swDocSpecification.ConfigurationName = Config
                End If
                swModel = MySwApp.OpenDoc7(swDocSpecification)
                'DraDoc
                If swModel Is Nothing Then
                    GoTo ER
                Else
                    Dim Err() As String = ExportToCAD(swModel, 0, DraDoc)

                    ListView1.Items(ii).SubItems(1).Text = Err(2) & " " & Err(3) & " " & Err(4) & " " & Err(5)

                    Dim R As String = Err(1) & " " & Err(2) & " " & Err(3) & " " & Err(4) & " " & Err(5)
                    swr.WriteLine(swModel.GetTitle & "  " & R)
                    MySwApp.CloseDoc(swModel.GetTitle)
                    swModel = Nothing
                End If
                ListView1.EndUpdate()
ER:         Next
            'MySwApp.Visible = True
        End Using
        MySwApp.SetUserPreferenceToggle(8, bRet1(0)) : MySwApp.SetUserPreferenceToggle(86, bRet1(1))
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

    End Sub
End Class