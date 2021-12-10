<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDrawingComp
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
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

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.grpSolidSolid = New System.Windows.Forms.GroupBox()
        Me.btnSldGo = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmbCompText2 = New System.Windows.Forms.ComboBox()
        Me.cmbCompText1 = New System.Windows.Forms.ComboBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnSansyo = New System.Windows.Forms.Button()
        Me.txtPath = New System.Windows.Forms.TextBox()
        Me.grpSolidSolid.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(25, 123)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(0, 12)
        Me.Label6.TabIndex = 0
        '
        'grpSolidSolid
        '
        Me.grpSolidSolid.Controls.Add(Me.btnSldGo)
        Me.grpSolidSolid.Controls.Add(Me.Label2)
        Me.grpSolidSolid.Controls.Add(Me.Label1)
        Me.grpSolidSolid.Controls.Add(Me.cmbCompText2)
        Me.grpSolidSolid.Controls.Add(Me.cmbCompText1)
        Me.grpSolidSolid.Location = New System.Drawing.Point(12, 69)
        Me.grpSolidSolid.Name = "grpSolidSolid"
        Me.grpSolidSolid.Size = New System.Drawing.Size(462, 171)
        Me.grpSolidSolid.TabIndex = 9
        Me.grpSolidSolid.TabStop = False
        Me.grpSolidSolid.Text = "SolidWorks図面ファイルの比較"
        '
        'btnSldGo
        '
        Me.btnSldGo.Location = New System.Drawing.Point(381, 139)
        Me.btnSldGo.Name = "btnSldGo"
        Me.btnSldGo.Size = New System.Drawing.Size(75, 23)
        Me.btnSldGo.TabIndex = 2
        Me.btnSldGo.Text = "実行"
        Me.btnSldGo.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 79)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(35, 12)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "図面2"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(35, 12)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "図面1"
        '
        'cmbCompText2
        '
        Me.cmbCompText2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCompText2.FormattingEnabled = True
        Me.cmbCompText2.Location = New System.Drawing.Point(12, 101)
        Me.cmbCompText2.Name = "cmbCompText2"
        Me.cmbCompText2.Size = New System.Drawing.Size(362, 20)
        Me.cmbCompText2.TabIndex = 0
        '
        'cmbCompText1
        '
        Me.cmbCompText1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCompText1.FormattingEnabled = True
        Me.cmbCompText1.Location = New System.Drawing.Point(12, 48)
        Me.cmbCompText1.Name = "cmbCompText1"
        Me.cmbCompText1.Size = New System.Drawing.Size(362, 20)
        Me.cmbCompText1.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnSansyo)
        Me.GroupBox1.Controls.Add(Me.txtPath)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(462, 46)
        Me.GroupBox1.TabIndex = 10
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "比較結果出力先フォルダパス"
        '
        'btnSansyo
        '
        Me.btnSansyo.Location = New System.Drawing.Point(385, 15)
        Me.btnSansyo.Name = "btnSansyo"
        Me.btnSansyo.Size = New System.Drawing.Size(71, 23)
        Me.btnSansyo.TabIndex = 3
        Me.btnSansyo.Text = "参照"
        Me.btnSansyo.UseVisualStyleBackColor = True
        '
        'txtPath
        '
        Me.txtPath.Enabled = False
        Me.txtPath.Location = New System.Drawing.Point(6, 17)
        Me.txtPath.Name = "txtPath"
        Me.txtPath.Size = New System.Drawing.Size(368, 19)
        Me.txtPath.TabIndex = 2
        '
        'frmDrawingComp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(486, 251)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.grpSolidSolid)
        Me.Controls.Add(Me.Label6)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmDrawingComp"
        Me.Text = "図面比較"
        Me.grpSolidSolid.ResumeLayout(False)
        Me.grpSolidSolid.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents grpSolidSolid As System.Windows.Forms.GroupBox
    Friend WithEvents btnSldGo As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmbCompText2 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbCompText1 As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents btnSansyo As Windows.Forms.Button
    Friend WithEvents txtPath As Windows.Forms.TextBox
End Class
