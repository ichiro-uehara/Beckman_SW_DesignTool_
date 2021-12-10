<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCompInfo
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.lblModel = New System.Windows.Forms.Label()
        Me.lblHikaku = New System.Windows.Forms.Label()
        Me.lblCSVName = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(205, 12)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "一致したアイテムフォントを変更しています。"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 36)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(199, 12)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "SolidWorks側の図面を確認してください。"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(410, 516)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "OK"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.HorizontalScrollbar = True
        Me.ListBox1.ItemHeight = 12
        Me.ListBox1.Location = New System.Drawing.Point(11, 177)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(474, 328)
        Me.ListBox1.TabIndex = 2
        '
        'lblModel
        '
        Me.lblModel.AutoSize = True
        Me.lblModel.Location = New System.Drawing.Point(12, 156)
        Me.lblModel.Name = "lblModel"
        Me.lblModel.Size = New System.Drawing.Size(137, 12)
        Me.lblModel.TabIndex = 3
        Me.lblModel.Text = "流用先に無いアイテム一覧 "
        '
        'lblHikaku
        '
        Me.lblHikaku.AutoSize = True
        Me.lblHikaku.Location = New System.Drawing.Point(12, 64)
        Me.lblHikaku.Name = "lblHikaku"
        Me.lblHikaku.Size = New System.Drawing.Size(53, 12)
        Me.lblHikaku.TabIndex = 3
        Me.lblHikaku.Text = "比較文言"
        '
        'lblCSVName
        '
        Me.lblCSVName.AutoSize = True
        Me.lblCSVName.Location = New System.Drawing.Point(12, 102)
        Me.lblCSVName.Name = "lblCSVName"
        Me.lblCSVName.Size = New System.Drawing.Size(59, 12)
        Me.lblCSVName.TabIndex = 3
        Me.lblCSVName.Text = "出力csv名"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(284, 516)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(120, 23)
        Me.Button2.TabIndex = 4
        Me.Button2.Text = "印刷ページ設定"
        Me.Button2.UseVisualStyleBackColor = True
        Me.Button2.Visible = False
        '
        'frmCompInfo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(498, 551)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.lblCSVName)
        Me.Controls.Add(Me.lblHikaku)
        Me.Controls.Add(Me.lblModel)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmCompInfo"
        Me.Text = "図面比較"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents lblModel As System.Windows.Forms.Label
    Friend WithEvents lblHikaku As System.Windows.Forms.Label
    Friend WithEvents lblCSVName As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
End Class
