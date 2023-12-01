Imports System.Windows.Forms
Imports System.Windows
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Public Class SZM
    Dim cc() As Control
    Dim ic As Integer
    Private Sub SZM_Activated(sender As Object, e As EventArgs) Handles Me.HandleCreated
        If cc Is Nothing Then IterateThroughControls(Me)

        ReaderData()
        Dim ia, ib As Integer

        For Each Cv As Object In cc
            If Cv Is Nothing Then Exit For
            Select Case TypeName(Cv)
                Case "TextBox"
                    Dim Tb As TextBox = Cv
                    Tb.Text = SzmStr(ia)
                    ia = ia + 1
                Case "CheckBox"
                    Dim Cb As CheckBox = Cv
                    Cb.Checked = bRet(ib)
                    ib = ib + 1
                Case "RadioButton"
                    Dim Rb As RadioButton = Cv
                    Rb.Checked = bRet(ib)
                    ib = ib + 1
            End Select
        Next

        If SzmStr(ia) = "////" Then ia += 1



    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If cc(0) Is Nothing Then ic = 0 : IterateThroughControls(Me)
        Dim ia, ib As Integer
        For Each Cv As Object In cc
            If Cv Is Nothing Then Exit For
            Select Case TypeName(Cv)

                Case "CheckBox"
                    Dim Cb As CheckBox = Cv
                    bRet(ib) = Cb.Checked
                    ib = ib + 1
                Case "RadioButton"
                    Dim Rb As RadioButton = Cv
                    bRet(ib) = Rb.Checked
                    ib = ib + 1

                Case "TextBox"
                    Dim Tb As TextBox = Cv
                    SzmStr(ia) = Tb.Text
                    ia = ia + 1

            End Select
        Next


        WriterData()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, Button7.Click, Button6.Click
        OpenFileDialog1.Filter = "工程图模板(*.DRWDOT)|*.DRWDOT|All Files(*.*)|*.*"
        If OpenFileDialog1.ShowDialog() = Forms.DialogResult.OK Then 工程图路径.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        Dim lic As New lic
        lic.ShowDialog()
        lic.Dispose()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles 检查精度.TextChanged

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs)
        If Not getSwApp() Then Exit Sub
        mainDS()
        '孤立编辑()

    End Sub

    Sub IterateThroughControls(ByVal parent As Control)
        cc = {
        工程图路径, 检查精度, 材质, 释放槽比率, TextBox7, TextBox8, 导出面路径, 展开图文件, TextBox4, TextBox1, TextBox2, TextBox3,
        自定义属性， 特定属性, TextBox5, RadioButtonK, RadioButton扣, RadioButton矩形, RadioButton圆, RadioButton撕裂,
        CheckBox1, CheckBox2, CheckBox材质, CheckBox草图, CheckBox折弯线, CheckBoxDWG, CheckBox8， CheckBox4， CheckBox6}

    End Sub

    Private Sub 删除ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 删除ToolStripMenuItem.Click
        '单击进入编辑状态，删除活动的dataview选中的表格
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        '打开安装目录
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        '修改系数文件
    End Sub
End Class