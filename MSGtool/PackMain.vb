Imports SolidWorks.Interop.sldworks
Imports System.Windows.Forms
Imports System.Windows
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq


Public Class PackMain
    Dim swModel As ModelDoc2

    Public TreeItemList As List(Of TreeListViewItem) '记录每一个多列树节点的列表
    Public newModelNameDict As Dictionary(Of String, String) '记录每一个源模型变换之后的模型文件名称

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        swModel = myswApp.ActiveDoc

        If swModel Is Nothing Then
            MsgBox("请打开模型")
            Exit Sub
        End If

        If swModel.GetType = 1 Or swModel.GetType = 2 Then

            If swModel.GetPathName = "" Then
                MsgBox("请先保存模型")
                Exit Sub
            End If

            LoadModel(swModel)
        Else
            MsgBox("只能在零件或者装配体中使用")
            Exit Sub
        End If

    End Sub

    Private Sub LoadModel(ByVal swModel As ModelDoc2)
        Dim swConfMgr As ConfigurationManager
        Dim swConf As Configuration
        Dim swRootComp As Component2

        swConfMgr = swModel.ConfigurationManager
        swConf = swConfMgr.ActiveConfiguration
        swRootComp = swConf.GetRootComponent3(True)

        tvTaskList.Items.Clear()

        Dim curItem As TreeListViewItem

        Dim FilePath As String

        FilePath = swModel.GetPathName

        Dim FileName As String

        FileName = Path.GetFileName(FilePath)

        SetMsg("开始处理装配体：" & FilePath)
        TreeItemList = New List(Of TreeListViewItem)
        newModelNameDict = New Dictionary(Of String, String)
        curItem = tvTaskList.Items.Add(FileName, 0)

        curItem.SubItems.Add(FilePath)
        curItem.SubItems.Add(FileName)
        curItem.SubItems.Item(1).Tag = FilePath
        curItem.SubItems.Item(2).Tag = FileName

        TreeItemList.Add(curItem)
        newModelNameDict.Add(FilePath, FileName)

        If swModel.GetType = 2 Then
            SetMsg("切换装配体到还原状态：" & FilePath)
            swModel.ResolveAllLightWeightComponents(False)
        End If


        TraCom(swRootComp, curItem)
        curItem.IsExpanded = True
        SetMsg("模型树加载完成")

    End Sub
    Private Sub SetMsg(ByVal msg As String)

        ToolStripStatusLabel1.Text = Now

        ToolStripStatusLabel2.Text = msg

        StatusStrip1.Refresh()
    End Sub

    Public Sub TraCom(ByVal swCom As Component2, ByVal ParTreeViewNode As TreeListViewItem)

        Dim myComDict As Dictionary(Of String, String)
        myComDict = New Dictionary(Of String, String)
        Dim swComps As Object

        swComps = swCom.GetChildren

        If IsNothing(swComps) = False Then

            For i = 0 To UBound(swComps)

                Dim curComP As Component2

                curComP = swComps(i)

                If curComP.IsSuppressed Then

                    Continue For
                End If
                If curComP.IsVirtual Then

                    Continue For
                End If

                Dim compPath As String

                compPath = curComP.GetPathName

                Dim FileName As String

                FileName = Path.GetFileName(compPath)

                compPath = compPath.ToUpper

                If myComDict.ContainsKey(compPath) = False Then

                    myComDict.Add(compPath, compPath)

                    Dim curItem As TreeListViewItem
                    curItem = ParTreeViewNode.Items.Add(FileName, 1)
                    SetMsg("添加子件：" & compPath)
                    curItem.SubItems.Add(compPath)
                    curItem.SubItems.Add(FileName)

                    curItem.SubItems.Item(1).Tag = compPath
                    curItem.SubItems.Item(2).Tag = FileName

                    TreeItemList.Add(curItem)

                    If newModelNameDict.ContainsKey(compPath) = False Then

                        newModelNameDict.Add(compPath, FileName)
                    End If

                    TraCom(curComP, curItem)
                End If


            Next

        End If

        '后续调试
        Dim vModelPathName As Object
        Dim vComponentPathName As Object

        Dim vFeature As Object
        Dim vDataType As Object
        Dim vStatus As Object
        Dim vRefEntity As Object
        Dim vFeatComp As Object
        Dim nConfigOpt As Long
        Dim sConfigName As String
        Dim nRefCount As Long
        Dim swModel As ModelDoc2
        Dim swModDocExt As ModelDocExtension

        swModel = swCom.GetModelDoc2

        If swModel Is Nothing Then
            Exit Sub
        End If
        swModDocExt = swModel.Extension
        nRefCount = swModDocExt.ListExternalFileReferencesCount

        swModDocExt.ListExternalFileReferences(vModelPathName, vComponentPathName,
                                           vFeature, vDataType, vStatus, vRefEntity,
                                           vFeatComp, nConfigOpt, sConfigName)
        Dim refList As List(Of String)

        refList = New List(Of String)

        For i = 0 To nRefCount - 1

            '20180824 不考虑对虚拟件的参考引用
            If refList.IndexOf(vModelPathName(i)) < 0 And vModelPathName(i) <> "" _
                And FileIO.FileSystem.FileExists(vModelPathName(i)) And myComDict.ContainsKey(UCase(vModelPathName(i))) = False Then

                refList.Add(vModelPathName(i))

                Dim curItem As TreeListViewItem

                Dim FileName As String

                FileName = Path.GetFileName(vModelPathName(i))
                curItem = ParTreeViewNode.Items.Add(FileName, 1)
                SetMsg("添加参考引用：" & vModelPathName(i))
                curItem.SubItems.Add(vModelPathName(i))
                curItem.SubItems.Add(FileName)
                curItem.SubItems.Item(1).Tag = vModelPathName(i)
                curItem.SubItems.Item(2).Tag = FileName

                TreeItemList.Add(curItem)
                If newModelNameDict.ContainsKey(vModelPathName(i)) = False Then

                    newModelNameDict.Add(vModelPathName(i), FileName)
                End If
            End If


        Next i





    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        If FolderBrowserDialog1.ShowDialog = Forms.DialogResult.OK Then

            TextBox1.Text = FolderBrowserDialog1.SelectedPath

            If TextBox1.Text.EndsWith("\") = False Then

                TextBox1.Text = TextBox1.Text & "\"
            End If

        End If
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click

        '文件夹
        Dim disPath As String

        disPath = TextBox1.Text
        If My.Computer.FileSystem.DirectoryExists(disPath) = False Then
            MsgBox("目标文件夹不存在")
            Exit Sub
        End If

        If disPath.EndsWith("\") = False Then

            TextBox1.Text = disPath & "\"
            disPath = disPath & "\"
        End If


        '获取所有需要复制的模型的文件路径
        If TreeItemList Is Nothing Then
            MsgBox("请先加载模型树")
            Exit Sub
        End If

        If TreeItemList.Count = 0 Then
            MsgBox("请先加载模型树")
            Exit Sub
        End If

        'Dim modelPathT = From t As TreeListViewItem In TreeItemList Select CStr(t.SubItems.Item(1).Tag)

        'Dim ModelPath As List(Of String)

        'ModelPath = modelPathT.ToList

        'ModelPath = ModelPath.Distinct.ToList

        ''复制模型到新文件夹
        'SetMsg("开始复制文件")
        'For i = 0 To ModelPath.Count - 1

        '    '源模型路径
        '    Dim rp As String
        '    rp = ModelPath(i)

        '    '新模型luj
        '    Dim np As String

        '    np = disPath & newModelNameDict.Item(rp)
        '    SetMsg("开始复制文件：" & rp)
        '    My.Computer.FileSystem.CopyFile(rp, np)


        'Next
        'SetMsg("文件复制完成")

        '复制模型到新文件夹
        SetMsg("开始复制文件")


        For i = 0 To TreeItemList.Count - 1

            Dim curTreeItem As TreeListViewItem

            curTreeItem = TreeItemList(i)

            '源模型路径
            Dim rp As String
            rp = curTreeItem.SubItems.Item(1).Tag

            '新模型路径
            Dim np As String

            np = disPath & curTreeItem.SubItems.Item(2).Text
            SetMsg("开始复制文件：" & rp)
            If My.Computer.FileSystem.FileExists(np) = False Then

                My.Computer.FileSystem.CopyFile(rp, np)

            End If
            '判断是否有附加的文件的依据：附加文件的存储路径和模型一致，仅后缀名不同

            CopyExt(CheckBox1.Checked, rp, ".slddrw", disPath) 'CopyExt(CheckBox1.Checked, rp,CheckBox1.text, disPath)
            CopyExt(CheckBox2.Checked, rp, ".dwg", disPath)
            CopyExt(CheckBox2.Checked, rp, ".pdf", disPath)
            'If CheckBox1.Checked Then

            '    Dim slddrwT As String
            '    slddrwT = Path.GetDirectoryName(rp) & "\" & Path.GetFileNameWithoutExtension(rp) & ".slddrw"

            '    If My.Computer.FileSystem.FileExists(slddrwT) Then
            '        Dim newSlddrwT As String
            '        newSlddrwT = disPath & Path.GetFileNameWithoutExtension(rp) & ".slddrw"

            '        If My.Computer.FileSystem.FileExists(newSlddrwT) = False Then

            '            My.Computer.FileSystem.CopyFile(slddrwT, newSlddrwT)


            '        End If

            '    End If

            'End If



        Next i

        '更新参考引用

        SetMsg("更新参考文件")

        myswApp.CloseAllDocuments(True)

        Dim modelT = From t As TreeListViewItem In TreeItemList Where t.ChildrenCount > 0 Select t

        For n = 0 To modelT.Count - 1

            Dim ReferencingDocument As String

            Dim curTreeLitem As TreeListViewItem
            curTreeLitem = modelT(n)

            ReferencingDocument = disPath & curTreeLitem.SubItems.Item(2).Text

            Dim modelT1 = From t1 As TreeListViewItem In TreeItemList Where t1.Parent Is curTreeLitem Select t1

            For n1 = 0 To modelT1.Count - 1

                Dim ReferencedDocument As String
                Dim NewReference As String
                Dim curTreeLitem1 As TreeListViewItem
                curTreeLitem1 = modelT1(n1)

                ReferencedDocument = curTreeLitem1.SubItems.Item(1).Tag
                NewReference = disPath & curTreeLitem1.SubItems.Item(2).Text

                '  ReferencingDocument & vbCrLf & ReferencedDocument & vbCrLf & NewReference 

                myswApp.ReplaceReferencedDocument(ReferencingDocument, ReferencedDocument, NewReference)

            Next

        Next

        '如果有工程图复制的情况，要更新工程图的参考模型
        If CheckBox1.Checked = False Then
            For i = 0 To TreeItemList.Count - 1

                Dim ReferencedDocument As String
                Dim NewReferenceDocument As String
                Dim curTreeLitem As TreeListViewItem
                curTreeLitem = TreeItemList(i)

                ReferencedDocument = curTreeLitem.SubItems.Item(1).Tag
                NewReferenceDocument = disPath & curTreeLitem.SubItems.Item(2).Text '新的模型的路径

                Dim NewSlddrwPath As String
                NewSlddrwPath = disPath & Path.GetFileNameWithoutExtension(NewReferenceDocument) & ".slddrw"

                If My.Computer.FileSystem.FileExists(NewSlddrwPath) Then

                    myswApp.ReplaceReferencedDocument(NewSlddrwPath, ReferencedDocument, NewReferenceDocument)

                End If


            Next
        End If



        SetMsg("参考文件更新完成")




    End Sub
    Public Sub CopyExt(ByVal b As Boolean, ByVal rp As String, ByVal ext As String, ByVal disPath As String)


        If b Then
            Dim r As String
            r = Path.GetDirectoryName(rp) & "\" & Path.GetFileNameWithoutExtension(rp) & ext

            If My.Computer.FileSystem.FileExists(r) Then
                Dim n As String
                n = disPath & Path.GetFileNameWithoutExtension(rp) & ext

                If My.Computer.FileSystem.FileExists(n) = False Then

                    My.Computer.FileSystem.CopyFile(r, n)

                End If

            End If

        End If
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        '全选
        For i = 0 To TreeItemList.Count - 1

            Dim curTreeItem As TreeListViewItem

            curTreeItem = TreeItemList(i)

            curTreeItem.Checked = True


            '需要继续处理相同名称的节点
            Dim modelT1 = From t1 As TreeListViewItem In TreeItemList Where t1.SubItems.Item(0).Text = curTreeItem.SubItems.Item(0).Text Select t1

            For n1 = 0 To modelT1.Count - 1
                Dim curTreeLitem1 As TreeListViewItem
                curTreeLitem1 = modelT1(n1)

                curTreeLitem1.Checked = curTreeItem.Checked
            Next


        Next i
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        '全选
        For i = 0 To TreeItemList.Count - 1

            Dim curTreeItem As TreeListViewItem

            curTreeItem = TreeItemList(i)

            curTreeItem.Checked = Not curTreeItem.Checked

            '需要继续处理相同名称的节点
            Dim modelT1 = From t1 As TreeListViewItem In TreeItemList Where t1.SubItems.Item(0).Text = curTreeItem.SubItems.Item(0).Text Select t1

            For n1 = 0 To modelT1.Count - 1
                Dim curTreeLitem1 As TreeListViewItem
                curTreeLitem1 = modelT1(n1)

                curTreeLitem1.Checked = curTreeItem.Checked
            Next


        Next i
    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        If TextBox2.Text = "" Then
            Exit Sub
        End If

        Dim unChar() As Char

        unChar = Path.GetInvalidFileNameChars()

        For i = 0 To TextBox2.Text.Length - 1

            For n = 0 To unChar.Count - 1

                If unChar(n) = TextBox2.Text.Substring(i, 1) Then
                    MsgBox("前缀不能含有非法字符")
                    Exit Sub

                End If
            Next


        Next

        Dim modelT = From t As TreeListViewItem In TreeItemList Where t.Checked = True Select t

        For n = 0 To modelT.Count - 1

            Dim curTreeLitem As TreeListViewItem
            curTreeLitem = modelT(n)

            curTreeLitem.SubItems.Item(2).Text = TextBox2.Text & curTreeLitem.SubItems.Item(2).Text

        Next

    End Sub

    Private Sub Button7_Click(sender As System.Object, e As System.EventArgs) Handles Button7.Click
        If TextBox3.Text = "" Then
            Exit Sub
        End If

        Dim unChar() As Char

        unChar = Path.GetInvalidFileNameChars()

        For i = 0 To TextBox3.Text.Length - 1

            For n = 0 To unChar.Count - 1

                If unChar(n) = TextBox2.Text.Substring(i, 1) Then
                    MsgBox("后缀不能含有非法字符")
                    Exit Sub

                End If
            Next


        Next

        Dim modelT = From t As TreeListViewItem In TreeItemList Where t.Checked = True Select t

        For n = 0 To modelT.Count - 1

            Dim curTreeLitem As TreeListViewItem
            curTreeLitem = modelT(n)

            Dim tt() As String

            tt = Split(curTreeLitem.SubItems.Item(2).Text, ".")

            Dim t1 As String
            t1 = tt(tt.Count - 1)

            ReDim Preserve tt(tt.Count - 2)


            curTreeLitem.SubItems.Item(2).Text = Join(tt, ".") & TextBox3.Text & "." & t1

        Next
    End Sub



    Private Sub tvTaskList_ItemChecked(sender As Object, e As System.Windows.Forms.ItemCheckedEventArgs) Handles tvTaskList.ItemChecked

        For i = 0 To TreeItemList.Count - 1

            Dim curTreeItem As TreeListViewItem

            curTreeItem = e.Item

            '需要继续处理相同名称的节点
            Dim modelT1 = From t1 As TreeListViewItem In TreeItemList Where t1.SubItems.Item(0).Text = curTreeItem.SubItems.Item(0).Text Select t1

            For n1 = 0 To modelT1.Count - 1
                Dim curTreeLitem1 As TreeListViewItem
                curTreeLitem1 = modelT1(n1)

                curTreeLitem1.Checked = curTreeItem.Checked
            Next


        Next i


    End Sub


    Private Sub Button8_Click(sender As System.Object, e As System.EventArgs) Handles Button8.Click
        If TextBox5.Text = "" Then
            Exit Sub
        End If

        Dim unChar() As Char

        unChar = Path.GetInvalidFileNameChars()

        For i = 0 To TextBox5.Text.Length - 1

            For n = 0 To unChar.Count - 1

                If unChar(n) = TextBox5.Text.Substring(i, 1) Then
                    MsgBox("后缀不能含有非法字符")
                    Exit Sub

                End If
            Next


        Next


        For i = 0 To TextBox4.Text.Length - 1

            For n = 0 To unChar.Count - 1

                If unChar(n) = TextBox4.Text.Substring(i, 1) Then
                    MsgBox("后缀不能含有非法字符")
                    Exit Sub

                End If
            Next


        Next

        Dim modelT = From t As TreeListViewItem In TreeItemList Where t.Checked = True Select t

        For n = 0 To modelT.Count - 1

            Dim curTreeLitem As TreeListViewItem
            curTreeLitem = modelT(n)

            Dim tt() As String

            tt = Split(curTreeLitem.SubItems.Item(2).Text, ".")

            Dim t1 As String
            t1 = tt(tt.Count - 1)

            ReDim Preserve tt(tt.Count - 2)

            curTreeLitem.SubItems.Item(2).Text = Replace(Join(tt, "."), TextBox5.Text, TextBox4.Text) & "." & t1

        Next

    End Sub
End Class

'进行参考引用的更新的时候：要更新的模型路径、原路径，新路径