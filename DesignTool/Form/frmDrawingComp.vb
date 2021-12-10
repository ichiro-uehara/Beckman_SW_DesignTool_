Imports System.Windows.Forms
Imports System.Collections.Generic
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst

Public Class frmDrawingComp

    Public m_DirectoryPath As String = ""
    Public m_AutoPath As String = ""
    Private m_HelpPath As String = ""
    Private m_activeDrawList As List(Of String) = Nothing

    Public m_newDraw As String = ""
    Public m_oldDraw As String = ""
    Private m_ModelPathList As List(Of String) = New List(Of String)
    Public m_CompMode As Integer = 0 ' 0=SolidとAuto 1=Solid同士

    Public Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.FormBorderStyle = FormBorderStyle.FixedSingle

        ' 現在開かれている図面一覧を取得
        m_activeDrawList = New List(Of String)
        Dim swApp As SldWorks = clsDesignTool.GetSWApp
        Dim swModels As Object = swApp.GetDocuments
        For i As Integer = 0 To swApp.GetDocumentCount - 1
            Dim tmpModel As ModelDoc2 = swModels(i)
            If tmpModel.GetType = 3 Then
                m_activeDrawList.Add(System.IO.Path.GetFileNameWithoutExtension(tmpModel.GetPathName))
                m_ModelPathList.Add(tmpModel.GetPathName)
            End If
        Next

        If m_activeDrawList.Count < 2 Then
            Me.btnSldGo.Enabled = False
        End If

    End Sub

    Private Sub frmDrawingComp_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown

        If m_activeDrawList.Count > 1 Then
            If cmbCompText1.Items.Count = 0 Then
                For i As Integer = 0 To m_activeDrawList.Count - 1
                    cmbCompText1.Items.Add(m_activeDrawList(i))
                    cmbCompText2.Items.Add(m_activeDrawList(i))
                Next
            End If
            If cmbCompText1.Items.Count >= 2 Then
                cmbCompText1.SelectedIndex = 0
                cmbCompText2.SelectedIndex = 1
            End If
        End If

    End Sub

    Private Sub btnSldGo_Click(sender As System.Object, e As System.EventArgs) Handles btnSldGo.Click

        If Me.cmbCompText1.Text = Me.cmbCompText2.Text Then
            MsgBox("同じ図面を比較することは出来ません。", MsgBoxStyle.ApplicationModal, "図面比較")
            Return
        End If

        If MsgBox("図面比較を行うと現在作業中の図面が上書き保存されます。" & vbCrLf & "図面比較を実施しますか。", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "図面比較") <> MsgBoxResult.Yes Then
            Return
        End If

        Dim newPath As String = ""
        Dim oldPath As String = ""

        For i As Integer = 0 To m_activeDrawList.Count - 1
            If m_activeDrawList(i) = Me.cmbCompText1.Text Then
                oldPath = m_ModelPathList(i)
            End If
            If m_activeDrawList(i) = Me.cmbCompText2.Text Then
                newPath = m_ModelPathList(i)
            End If
        Next

        Dim swApp As SldWorks = clsDesignTool.GetSWApp
        Dim swModel As ModelDoc2 = swApp.ActivateDoc(newPath)
        Dim swModel2 As ModelDoc2 = swApp.ActivateDoc(oldPath)

        swModel.Save()
        Dim tempName As String = System.IO.Path.GetFileName(newPath)
        Dim tempName2 As String = System.IO.Path.GetFileName(oldPath)
        Dim tempSpt() As String = Split(tempName, ".")
        Dim tempSpt2() As String = Split(tempName2, ".")
        'If swModel.SaveAs(System.IO.Path.GetDirectoryName(newPath) + "\" + tempSpt(0) + "_DrawComp.SLDDRW") Then
        If swModel.SaveAs(System.IO.Path.GetDirectoryName(newPath) + "\" + tempSpt(0) + "_DrawComp.SLDDRW") And
            swModel2.SaveAs(System.IO.Path.GetDirectoryName(oldPath) + "\" + tempSpt2(0) + "_DrawComp.SLDDRW") Then
            m_oldDraw = Me.cmbCompText1.Text + "_DrawComp"
            m_newDraw = Me.cmbCompText2.Text + "_DrawComp"

            m_CompMode = 1
            Me.DialogResult = vbOK
            Me.Close()
        Else
            MsgBox("図面バックアップの作成に失敗しました。", MsgBoxStyle.ApplicationModal, "図面比較")
        End If

    End Sub

    Private Sub btnSansyo_Click(sender As Object, e As EventArgs) Handles btnSansyo.Click

        Dim fbd As New FolderBrowserDialog
        fbd.Description = "出力先フォルダを指定してください。"
        fbd.RootFolder = System.Environment.SpecialFolder.Desktop
        fbd.SelectedPath = "C:\Windows"
        fbd.ShowNewFolderButton = True

        If fbd.ShowDialog(Me) = DialogResult.OK Then
            Me.txtPath.Text = fbd.SelectedPath
            m_DirectoryPath = fbd.SelectedPath
        End If

    End Sub
End Class