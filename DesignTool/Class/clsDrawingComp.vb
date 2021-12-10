Imports System.Collections.Generic
Imports SolidWorks.Interop.sldworks
Imports System.IO
Imports System.Collections
Imports System
Imports SolidWorks.Interop.swconst

Public Class clsDrawingComp

    ' ************************************************
    '
    ' メイン処理
    '
    ' ************************************************

    ' SolidWorks同士の比較
    Public Shared Function DrawingCompSolid(ByVal directoryPath As String, ByVal oldSolidName As String, ByVal newSolidName As String) As Boolean

        DrawingCompSolid = False

        Try
            Dim solidCSVData As New List(Of String)
            Dim acadData As New List(Of clsCadData)
            Dim acadCSVData As New List(Of String)
            Dim acadDocName As String = ""
            Dim compCSVData As New List(Of String)
            Dim noCompCSVData As New List(Of String)

            ' SolidWorksからデータ抽出
            Dim solidControl As New clsSolidControl(Nothing)
            If solidControl.DrawingCompDataSolid(directoryPath, oldSolidName, newSolidName, compCSVData, noCompCSVData) = False Then
                Return False
            End If
            solidControl = Nothing

            If solidCSVData Is Nothing Then
                MsgBox("SolidWorksの図面情報が取得出来ませんでした。", MsgBoxStyle.ApplicationModal, "図面比較")
                Return False
            End If

            If compCSVData Is Nothing Or compCSVData.Count = 0 Then
                MsgBox("一致するアイテムが存在しませんでした。", MsgBoxStyle.ApplicationModal, "図面比較")
                Return False
            End If

            If directoryPath <> "" Then

                ' 不一致リストにソートありでヘッダーを付け加える
                clsDCCommon.SortCsvData(True, 2, noCompCSVData)
                clsDCCommon.OutputCSVFile(directoryPath + "\不一致リスト.csv", noCompCSVData)

                ' 一致リストにソートありでヘッダーを付け加える
                clsDCCommon.SortCsvData(True, 3, compCSVData)
                clsDCCommon.OutputCSVFile(directoryPath + "\一致リスト.csv", compCSVData)

            End If

            MsgBox("図面比較が完了しました。", MsgBoxStyle.ApplicationModal, "図面比較")

            Return True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        Finally
        End Try

    End Function

End Class
