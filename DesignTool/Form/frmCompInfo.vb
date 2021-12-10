Imports SolidWorks.Interop.sldworks
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Public Class frmCompInfo

    Private m_AnotationList As List(Of Annotation)
    Private m_AnotationColorList As List(Of Integer)
    Private m_AnnotationItaric As List(Of Boolean)
    Private m_AnnotationFontName As List(Of String)
    Private m_AnnotationDocumentFlag As List(Of Boolean)
    Private m_AnnotationSize As List(Of Double)

    Private m_DirectoryPath As String = ""
    Private WithEvents m_DrawingDoc As DrawingDoc = Nothing

    Public Sub New(ByVal annotations As List(Of Annotation), _
                    ByVal annotationColors As List(Of Integer), _
                    ByVal annotationItaric As List(Of Boolean), _
                    ByVal annotationFontName As List(Of String), _
                    ByVal AnnotationDocumentFlag As List(Of Boolean), _
                    ByVal AnnotationSize As List(Of Double), _
                    ByVal directoryPath As String, _
                    ByVal hikakumoto As String, _
                    ByVal hikakusaki As String, _
                    ByVal viewFlag As Boolean, _
                    ByVal viewList As List(Of String))

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        Me.MaximizeBox = Not Me.MaximizeBox
        Me.MinimizeBox = Not Me.MinimizeBox
        Me.ControlBox = Not Me.ControlBox

        m_AnotationList = annotations
        m_AnotationColorList = annotationColors
        m_AnnotationItaric = annotationItaric
        m_AnnotationFontName = annotationFontName
        m_DirectoryPath = directoryPath
        m_AnnotationSize = AnnotationSize

        If viewFlag = False Then
            Me.lblHikaku.Text = hikakumoto + ".dwg （流用元） と" + vbCrLf + hikakusaki + ".SLDDRW （流用先） の比較"
            Me.lblCSVName.Text = "出力ファイル名：" + vbCrLf + "AC_" + hikakumoto + ".csv" + vbCrLf + "SW_" + hikakusaki + ".csv"
        Else
            Me.Label2.Text = "流用先に指定した図面を確認してください。"
            Me.lblHikaku.Text = hikakumoto + ".SLDDRW （流用元） と " + vbCrLf + hikakusaki + ".SLDDRW （流用先） の比較"
            Me.lblCSVName.Text = ""
        End If

        If viewList.Count > 0 Then
            For i As Integer = 0 To viewList.Count - 1
                Me.ListBox1.Items.Add(viewList(i))
            Next
        Else
            Me.ListBox1.Enabled = False
            'Me.ListBox1.BackColor = Drawing.Color.Gray
        End If

        Dim swApp As SldWorks = clsDesignTool.GetSWApp

        m_DrawingDoc = swApp.ActiveDoc

        ' バージョン取得
        Dim ver As System.Diagnostics.FileVersionInfo
        ver = _
            System.Diagnostics.FileVersionInfo.GetVersionInfo( _
            System.Reflection.Assembly.GetExecutingAssembly().Location)

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        ' 元の色に戻さないようにせーと。
        '' 元の色に戻す
        'For i As Integer = 0 To m_AnotationList.Count - 1
        '    If m_AnotationList(i) Is Nothing Then
        '        Continue For
        '    End If
        '    Try
        '        m_AnotationList(i).Color = m_AnotationColorList(i)

        '        Dim tempTextFormat As TextFormat = m_AnotationList(i).GetTextFormat(0)
        '        If tempTextFormat IsNot Nothing Then
        '            If tempTextFormat.TypeFaceName = "NONE" Then
        '                m_AnotationList(i).ISetTextFormat(i, True, Nothing)
        '            Else
        '                tempTextFormat.Italic = m_AnnotationItaric(i)
        '                tempTextFormat.TypeFaceName = m_AnnotationFontName(i)
        '                tempTextFormat.CharHeight = m_AnnotationSize(i)
        '                m_AnotationList(i).ISetTextFormat(i, False, tempTextFormat)
        '            End If
        '        Else
        '            m_AnotationList(i).ISetTextFormat(i, True, Nothing)
        '        End If

        '        ' ドキュメントフォントを使用していた場合ドキュメントフォントを保存するに戻す
        '        If m_AnnotationDocumentFlag(i) = True Then
        '            m_AnotationList(i).ISetTextFormat(i, True, Nothing)
        '        End If

        '    Catch
        '        Continue For
        '    End Try
        'Next

        If m_DirectoryPath <> "" Then
            MsgBox(m_DirectoryPath + vbCrLf + "上記場所に比較結果のcsvファイルを出力しました。", MsgBoxStyle.ApplicationModal, "図面比較")
        End If

        MsgBox("図面比較が終了しました。", MsgBoxStyle.ApplicationModal, "図面比較")

        Me.Close()
    End Sub

    ' 印刷処理案が出てたころの残骸処理（印刷処理の参考にでも）
    'Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

    '    Dim swApp As SldWorks = clsGetSWApp()
    '    Dim swModel As ModelDoc2 = swApp.ActiveDoc
    '    swModel.Extension.RunCommand(SolidWorks.Interop.swcommands.swCommands_e.swCommands_Page_Setup, "印刷")

    'End Sub

    Private Sub frmCompInfo_FormClosing(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        'Dim swApp As SldWorks = clsGetSWApp
        'swApp.EnableFileMenu = True

        ' 保存ボタン無効
        '' 選択モード解除
        'RemoveHandler m_DrawingDoc.FileSaveAsNotify, AddressOf Me.DrawingDocFileSaveAsNotify
        'RemoveHandler m_DrawingDoc.FileSaveNotify, AddressOf Me.DrawingDocFileSaveNotify

    End Sub

    ' 保存ボタン無効
    'Private Function DrawingDocFileSaveAsNotify(ByVal FileName As String) As Integer
    '    MsgBox("図面比較中は保存出来ません。", vbExclamation)
    '    Return -1
    'End Function

    'Private Function DrawingDocFileSaveNotify(ByVal FileName As String) As Integer
    '    MsgBox("図面比較中は保存出来ません。", vbExclamation)
    '    Return -1
    'End Function

End Class