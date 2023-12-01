<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class batch
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。  
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.打开文件ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.从列表中删除ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.清空列表ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.Button9 = New System.Windows.Forms.Button()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(12, 12)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(63, 23)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "添加文件"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(81, 12)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(64, 23)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "当前装配体"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'ListView1
        '
        Me.ListView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2})
        Me.ListView1.GridLines = True
        Me.ListView1.HideSelection = False
        Me.ListView1.Location = New System.Drawing.Point(-2, 41)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(725, 549)
        Me.ListView1.TabIndex = 4
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "文件名称"
        Me.ColumnHeader1.Width = 263
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "数量"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.Multiselect = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(164, 12)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(75, 23)
        Me.Button4.TabIndex = 5
        Me.Button4.Text = "导出DXF"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(393, 12)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(74, 23)
        Me.Button5.TabIndex = 5
        Me.Button5.Text = "纠错*特征"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.打开文件ToolStripMenuItem, Me.从列表中删除ToolStripMenuItem, Me.清空列表ToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.ShowImageMargin = False
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(124, 70)
        '
        '打开文件ToolStripMenuItem
        '
        Me.打开文件ToolStripMenuItem.Name = "打开文件ToolStripMenuItem"
        Me.打开文件ToolStripMenuItem.Size = New System.Drawing.Size(123, 22)
        Me.打开文件ToolStripMenuItem.Text = "打开文件"
        '
        '从列表中删除ToolStripMenuItem
        '
        Me.从列表中删除ToolStripMenuItem.Name = "从列表中删除ToolStripMenuItem"
        Me.从列表中删除ToolStripMenuItem.Size = New System.Drawing.Size(123, 22)
        Me.从列表中删除ToolStripMenuItem.Text = "从列表中删除"
        '
        '清空列表ToolStripMenuItem
        '
        Me.清空列表ToolStripMenuItem.Name = "清空列表ToolStripMenuItem"
        Me.清空列表ToolStripMenuItem.Size = New System.Drawing.Size(123, 22)
        Me.清空列表ToolStripMenuItem.Text = "清空列表"
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(473, 12)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(74, 23)
        Me.Button6.TabIndex = 5
        Me.Button6.Text = "纠错*全局"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(553, 12)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(65, 23)
        Me.Button7.TabIndex = 5
        Me.Button7.Text = "转换钣金"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'Button8
        '
        Me.Button8.Location = New System.Drawing.Point(245, 12)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(68, 23)
        Me.Button8.TabIndex = 5
        Me.Button8.Text = "导出展开"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'Button9
        '
        Me.Button9.Location = New System.Drawing.Point(319, 12)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(68, 23)
        Me.Button9.TabIndex = 5
        Me.Button9.Text = "更新展开"
        Me.Button9.UseVisualStyleBackColor = True
        '
        'batch
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(724, 592)
        Me.ContextMenuStrip = Me.ContextMenuStrip1
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button9)
        Me.Controls.Add(Me.Button8)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.ListView1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Name = "batch"
        Me.Text = "MSG批处理"
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button2 As Windows.Forms.Button
    Friend WithEvents Button3 As Windows.Forms.Button
    Friend WithEvents ListView1 As Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As Windows.Forms.ColumnHeader
    Friend WithEvents OpenFileDialog1 As Windows.Forms.OpenFileDialog
    Friend WithEvents Button4 As Windows.Forms.Button
    Friend WithEvents Button5 As Windows.Forms.Button
    Friend WithEvents ContextMenuStrip1 As Windows.Forms.ContextMenuStrip
    Friend WithEvents 打开文件ToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents 从列表中删除ToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents 清空列表ToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents Button6 As Windows.Forms.Button
    Friend WithEvents Button7 As Windows.Forms.Button
    Friend WithEvents Button8 As Windows.Forms.Button
    Friend WithEvents Button9 As Windows.Forms.Button
End Class
