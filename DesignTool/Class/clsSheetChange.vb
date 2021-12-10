Imports System.Collections.Generic
Imports SolidWorks.Interop.sldworks
Imports System.IO
Imports System.Collections
Imports System
Imports SolidWorks.Interop.swconst

Public Class clsSheetChange

    Public Shared Function InsertSheet() As Boolean


        Try

            Dim swApp As SldWorks = clsDesignTool.GetSWApp
            Dim swModel As ModelDoc2 = swApp.ActiveDoc

            If swModel.GetType <> 3 Then Return False

            Dim swDraw As DrawingDoc = swModel

            Dim sheetCount As Integer = swDraw.GetSheetCount

            ' 今アクティブなシートを記憶
            Dim swSheet As Sheet = swDraw.GetCurrentSheet()
            Dim curSheetName As String = swSheet.GetName
            Dim changeSheetName As String = swDraw.GetSheetNames(0)

            ' シート1をアクティブにして斜線入りに変更
            swDraw.ActivateSheet(changeSheetName)
            swSheet = swDraw.GetCurrentSheet()

            Dim ary As ArrayList = LoadIniData("[SheetFormatFolder]")
            If ary Is Nothing Then Return False

            Dim tempLateName As String = System.IO.Path.GetFileName(swSheet.GetTemplateName())
            Dim tempSpt() As String = Split(tempLateName, ".")

            Dim slashName As String = ""
            Dim slashName2 As String = ""
            Dim noChangeFlag As Boolean = False
            If InStr(tempLateName, "-P1_2") <> 0 Then
                slashName = tempLateName
                noChangeFlag = True
            Else
                slashName = tempSpt(0) + " -P1_2." + tempSpt(1)
            End If

            If tempLateName.Contains("assy") Then
                slashName2 = slashName
            Else
                If noChangeFlag = True Then
                    slashName2 = slashName.Replace("-P1_2", "-P2_2-")
                Else
                    slashName2 = tempSpt(0) + " -P2_2-." + tempSpt(1)
                End If
            End If

            Dim path As String = ary(0) + "\" + slashName
            Dim path2 As String = ary(0) + "\" + slashName2
            Dim outStr As String = ""

            If System.IO.File.Exists(path) Then

                ' 今が斜線なしシートの場合、斜線ありに変更
                If noChangeFlag = False Then
                    swDraw.SetupSheet5(swSheet.GetName, 12, 12, 1, 1, False, path, 0, 0, "ﾃﾞﾌｫﾙﾄ", True)
                    outStr = swSheet.GetName + " を斜線入りに変更しました。"
                End If

                Dim newSheetName As String = swSheet.GetName

                If swSheet.GetName.LastIndexOf(1) = swSheet.GetName.Length - 1 Then
                    newSheetName = swSheet.GetName.Remove(swSheet.GetName.Length - 1)
                End If

                Dim loopCount = 2

                Do While (loopCount < 100)
                    If swDraw.NewSheet3(newSheetName + loopCount.ToString, 12, 12, 1, 1, False, path2, 0, 0, "ﾃﾞﾌｫﾙﾄ") = True Then
                        If outStr <> "" Then outStr += vbCrLf
                        outStr += "斜線入り " + newSheetName + loopCount.ToString + " を追加しました。"
                        Exit Do
                    End If
                    loopCount = loopCount + 1
                Loop

                MsgBox(outStr)
            Else
                MsgBox("シートフォーマット " + path + " が見つかりませんでした。")
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Function

    Public Shared Function ChangeSheet() As Boolean

        Try

            Dim swApp As SldWorks = clsDesignTool.GetSWApp
            Dim swModel As ModelDoc2 = swApp.ActiveDoc

            If swModel.GetType <> 3 Then Return False

            Dim swDraw As DrawingDoc = swModel
            Dim swSheet As Sheet = swDraw.GetCurrentSheet()
            Dim ary As ArrayList = LoadIniData("[SheetFormatFolder]")
            Dim vSheetName As Object = swDraw.GetSheetNames

            If ary Is Nothing Then Return False

            Dim tempLateName As String = System.IO.Path.GetFileName(swSheet.GetTemplateName())
            Dim tempSpt() As String = Split(tempLateName, ".")
            Dim slashName As String = ""
            'If InStr(tempLateName, "_SLASH") <> 0 Then
            '    MsgBox("すでに斜線入りシートフォーマットが使用されています。")
            '    Return True
            'Else
            If InStr(tempLateName, "-P1_2") <> 0 Or InStr(tempLateName, "-P2_2-") <> 0 Then
                MsgBox("すでに斜線入りシートフォーマットが使用されています。")
                Return True
            Else
                If Array.IndexOf(vSheetName, swSheet.GetName) = 0 Then
                    slashName = tempSpt(0) + " -P1_2." + tempSpt(1)
                Else
                    If tempLateName.Contains("assy") Then
                        slashName = tempSpt(0) + " -P1_2." + tempSpt(1)
                    Else
                        slashName = tempSpt(0) + " -P2_2-." + tempSpt(1)
                    End If
                End If
            End If

            Dim path As String = ary(0) + "\" + slashName

            If System.IO.File.Exists(path) Then
                swDraw.SetupSheet5(swSheet.GetName, 12, 12, 1, 1, False, path, 0, 0, "ﾃﾞﾌｫﾙﾄ", True)
                MsgBox(swSheet.GetName + " を斜線入りに変更しました。")
                Return True
            Else
                MsgBox("シートフォーマット " + path + " が見つかりませんでした。")
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function

    Public Overloads Shared Function LoadIniData(ByVal KeyName As String) As ArrayList

        LoadIniData = Nothing
        If KeyName = "" Then Return Nothing

        If InStr(KeyName, "[") = 0 Or InStr(KeyName, "]") = 0 Then
            MsgBox("iniファイルの取得に失敗しました。" & vbCrLf & KeyName & "の値を確認してください。", MsgBoxStyle.ApplicationModal, "iniファイル読込")
            Return Nothing
        End If

        Dim LoadArrayList As New ArrayList
        Dim ProgramArrayList As New ArrayList
        Dim prIni As System.IO.StreamReader

        Dim appPath As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)

        Try
            prIni = New System.IO.StreamReader(appPath + "\Setting.ini")
        Catch
            MsgBox(appPath + "\Setting.iniの取得に失敗しました。" & vbCrLf & "上記の場所にファイルが存在するか確認してください。", MsgBoxStyle.ApplicationModal, "iniファイル読込")
            Return Nothing
        End Try

        Dim startFlag As Boolean = False
        While prIni.Peek() > -1
            Dim ReadLine As String = prIni.ReadLine()

            If ReadLine = KeyName Then
                startFlag = True
                Continue While
            End If

            If startFlag = False Then Continue While

            ' 空白、改行は飛ばす
            If ReadLine = "" Or ReadLine = vbCrLf Then
                Continue While
            End If

            ' #はコメント行として扱う
            If InStr(ReadLine, "#") <> 0 Then
                Continue While
            End If

            If InStr(ReadLine, "[") <> 0 Or InStr(ReadLine, "]") <> 0 Then
                Exit While
            End If

            LoadArrayList.Add(ReadLine)
        End While

        'ファイルを解放します。
        prIni.Close()

        LoadIniData = LoadArrayList
        LoadArrayList = Nothing

        ProgramArrayList = Nothing
        prIni = Nothing


    End Function

    ' Private Shared Function 


End Class
