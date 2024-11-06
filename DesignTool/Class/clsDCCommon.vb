' ************************************************
'
' コモンクラス
'
' ************************************************

Imports SolidWorks.Interop.sldworks
Imports System.Runtime.InteropServices
Imports System.Collections
Imports System.Collections.Generic

Public Class clsDCCommon

    ' iniファイル関連
    <DllImport("KERNEL32.DLL")>
    Public Shared Function WritePrivateProfileString(
        ByVal lpAppName As String,
        ByVal lpKeyName As String,
        ByVal lpString As String,
        ByVal lpFileName As String) As Integer
    End Function

    <DllImport("KERNEL32.DLL", CharSet:=CharSet.Auto)>
    Public Shared Function GetPrivateProfileString(
        ByVal lpAppName As String,
        ByVal lpKeyName As String,
        ByVal lpDefault As String,
        ByVal lpReturnedString As String,
        ByVal nSize As Integer,
        ByVal lpFileName As String) As Integer
    End Function

    <DllImport("KERNEL32.DLL", CharSet:=CharSet.Auto)>
    Public Shared Function GetPrivateProfileInt(
        ByVal lpAppName As String,
        ByVal lpKeyName As String,
        ByVal nDefault As Integer,
        ByVal lpFileName As String) As Integer
    End Function

    ' SolidWorksApplicationの取得
    Public Shared Function GetSldWorksApp() As SldWorks

        GetSldWorksApp = Nothing

        Try

            GetSldWorksApp = GetObject(, "SldWorks.Application.31")
            If (GetSldWorksApp Is Nothing) Then
                GetSldWorksApp = CreateObject("SldWorks.Application.31")
            End If

        Catch
            MsgBox("GetSlWorksApp" & "関数にて予期せぬエラーが発生しました。", MsgBoxStyle.Critical)
        Finally
        End Try

    End Function

    ' ************************************************
    '
    ' CSV出力処理
    '
    ' ************************************************

    Public Shared Function OutputCSVFile(ByVal savePath As String, ByVal outputValues As List(Of String)) As Boolean

        OutputCSVFile = False

        Try

            Dim sw As New IO.StreamWriter(savePath, False, System.Text.Encoding.GetEncoding("shift_jis"))
            Dim outputStr As String = ""

            For i As Integer = 0 To outputValues.Count - 1
                If outputValues(i) = "" Or outputValues(i) = vbCrLf Then
                    Continue For
                End If
                outputStr += outputValues(i)
                outputStr += vbCrLf
            Next

            ' csv出力
            sw.WriteLine(outputStr)
            sw.Close()

            OutputCSVFile = True

        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox("OutputCSVFile" & "関数にて予期せぬエラーが発生しました。", MsgBoxStyle.Critical)
        Finally
        End Try

    End Function

    ' ************************************************
    '
    ' CSV入力処理
    '
    ' ************************************************

    Public Shared Function InputCSVFile(ByVal loadPath As String,
                                        ByRef inputCadDatas As List(Of clsCadData),
                                        ByRef acadCSVData As List(Of String)) As Boolean

        InputCSVFile = False

        Try

            Dim sw As New IO.StreamReader(loadPath, System.Text.Encoding.GetEncoding("shift_jis"))
            'ファイルの最後までループ
            Do Until sw.EndOfStream
                Dim tempStr As String = sw.ReadLine()
                Dim tempStrs() As String = Split(tempStr, clsDesignTool.m_SepValue)

                If tempStrs.Length <> 8 Then
                    Continue Do
                End If

                ' csv形式として格納
                acadCSVData.Add(tempStr)

                ' データとして格納
                Dim tempCadData As New clsCadData
                tempCadData.cadType = tempStrs(0)
                tempCadData.itemType = tempStrs(1)
                tempCadData.itemPrefix = tempStrs(2)
                tempCadData.itemValue = tempStrs(3)
                tempCadData.itemSuffix = tempStrs(4)
                tempCadData.itemSymbol = tempStrs(5)
                tempCadData.X = tempStrs(6)
                tempCadData.Y = tempStrs(7)

                inputCadDatas.Add(tempCadData)

            Loop
            sw.Close()     'ファイルを閉じる

            InputCSVFile = True

        Catch ex As Exception
            MsgBox(ex.Message)
            'MsgBox("InputCSVFile" & "関数にて予期せぬエラーが発生しました。", MsgBoxStyle.Critical)
        Finally
        End Try

    End Function

    ' ************************************************
    '
    ' 出力CADデータをソートする
    '
    ' ************************************************
    Public Shared Sub SortCADList(ByRef sortListData As List(Of String))

        Try

            If sortListData Is Nothing Or sortListData.Count = 0 Then
                Return
            End If

            sortListData.Sort()

            'Dim sortCompCSVData As New List(Of String)
            'Dim sortValue As New List(Of String)
            'Dim sortValueType As New List(Of String)

            'Dim sortDouble As New List(Of Double)
            'Dim sortString As New List(Of String)
            'Dim sortDoubleType As New List(Of String)
            'Dim sortStringType As New List(Of String)

            'Dim memo As String = ""
            'Dim memo2 As String = ""
            'Dim usedNo As New List(Of Integer)
            'Dim tryDouble As Double = 0.0F

            'For i As Integer = 0 To sortListData.Count - 1
            '    Dim tempSpt() As String = Split(sortListData(i), clsDesignTool.m_SepValue)
            '    If i = 0 Then
            '        memo = tempSpt(0) ' swLinearDimensionなど
            '        If Double.TryParse(tempSpt(3), tryDouble) Then
            '            sortDouble.Add(Double.Parse(tempSpt(3)))
            '            sortDoubleType.Add(tempSpt(1))
            '        Else
            '            sortString.Add(tempSpt(3))
            '            sortStringType.Add(tempSpt(1))
            '        End If
            '        ' sortValue.Add(sortListData(i))
            '    Else
            '        If memo <> tempSpt(0) Or i = sortListData.Count - 1 Then

            '            ' 最後は追加してから処理を行う
            '            If i = sortListData.Count - 1 Then
            '                If Double.TryParse(tempSpt(3), tryDouble) Then
            '                    sortDouble.Add(Double.Parse(tempSpt(3)))
            '                    sortDoubleType.Add(tempSpt(1))
            '                Else
            '                    sortString.Add(tempSpt(3))
            '                    sortStringType.Add(tempSpt(1))
            '                End If
            '                'sortValue.Add(sortListData(i))
            '            End If

            '            ' ソート
            '            sortDouble.Sort()
            '            sortString.Sort()

            '            ' カテゴリを取得
            '            Dim tempTypeList As New List(Of String)
            '            For j As Integer = 0 To sortListData.Count - 1
            '                Dim tempSpt2() As String = Split(sortListData(j), clsDesignTool.m_SepValue)
            '                Dim isSafe As Boolean = False
            '                For k As Integer = 0 To tempTypeList.Count - 1
            '                    If tempTypeList(k) = tempSpt2(1) Then
            '                        isSafe = True
            '                    End If
            '                Next
            '                If isSafe = True Then Continue For
            '                tempTypeList.Add(tempSpt2(1))
            '            Next

            '            ' カテゴリ別にソート
            '            For j As Integer = 0 To tempTypeList.Count - 1
            '                ' 数値
            '                For k As Integer = 0 To sortDoubleType.Count - 1
            '                    If tempTypeList(j) = sortDoubleType(k) Then
            '                        sortValue.Add(sortDouble(k))
            '                        sortValueType.Add(sortDoubleType(k))
            '                    End If
            '                Next
            '                ' 文字列
            '                For k As Integer = 0 To sortStringType.Count - 1
            '                    If tempTypeList(j) = sortStringType(k) Then
            '                        sortValue.Add(sortString(k))
            '                        sortValueType.Add(sortStringType(k))
            '                    End If
            '                Next
            '            Next

            '            ' ソートした内容に入れなおす
            '            For j As Integer = 0 To sortListData.Count - 1
            '                Dim tempSpt2() As String = Split(sortListData(j), clsDesignTool.m_SepValue)
            '                For k As Integer = 0 To sortValue.Count - 1
            '                    If Double.TryParse(tempSpt2(3), tryDouble) = True And
            '                        Double.TryParse(sortValue(k), tryDouble) = True Then
            '                        If Math.Abs(Double.Parse(tempSpt2(3)) - Double.Parse(sortValue(k))) < 0.00001 Then
            '                            Continue For
            '                        End If
            '                    Else
            '                        If tempSpt2(1) <> sortValueType(k) Or tempSpt2(3) <> sortValue(k) Then
            '                            Continue For
            '                        End If
            '                    End If

            '                    Dim safeFlag2 As Boolean = False
            '                    For l As Integer = 0 To usedNo.Count - 1
            '                        If j = usedNo(l) Then
            '                            safeFlag2 = True
            '                            Exit For
            '                        End If
            '                    Next
            '                    If safeFlag2 = True Then Continue For

            '                    sortCompCSVData.Add(sortListData(j))
            '                    usedNo.Add(k)
            '                    Exit For

            '                Next
            '            Next

            '            '' すべて数値の場合
            '            'If sortString.Count = sortDouble.Count Then
            '            '    ' 値順でソート
            '            '    sortDouble.Sort()
            '            '    For j As Integer = 0 To sortDouble.Count - 1
            '            '        For k As Integer = 0 To sortValue.Count - 1

            '            '            Dim safeFlag2 As Boolean = False
            '            '            For l As Integer = 0 To usedNo.Count - 1
            '            '                If k = usedNo(l) Then
            '            '                    safeFlag2 = True
            '            '                    Exit For
            '            '                End If
            '            '            Next
            '            '            If safeFlag2 = True Then Continue For

            '            '            Dim tempSpt2() As String = Split(sortValue(k), clsDesignTool.m_SepValue)
            '            '            If Math.Abs(Double.Parse(tempSpt2(3)) - sortDouble(j)) < 0.00001 Then
            '            '                sortCompCSVData.Add(sortValue(k))
            '            '                usedNo.Add(k)
            '            '            End If
            '            '        Next
            '            '    Next
            '            'Else
            '            '    sortString.Sort()
            '            '    For j As Integer = 0 To sortString.Count - 1
            '            '        For k As Integer = 0 To sortValue.Count - 1

            '            '            Dim safeFlag2 As Boolean = False
            '            '            For l As Integer = 0 To usedNo.Count - 1
            '            '                If k = usedNo(l) Then
            '            '                    safeFlag2 = True
            '            '                    Exit For
            '            '                End If
            '            '            Next
            '            '            If safeFlag2 = True Then Continue For

            '            '            Dim tempSpt2() As String = Split(sortValue(k), clsDesignTool.m_SepValue)
            '            '            If tempSpt2(3) = sortString(j) Then
            '            '                sortCompCSVData.Add(sortValue(k))
            '            '                usedNo.Add(k)
            '            '            End If
            '            '        Next
            '            '    Next
            '            'End If

            '            sortValue.Clear()
            '            sortDouble.Clear()
            '            sortString.Clear()
            '            usedNo.Clear()

            '            'memo = tempSpt(0)
            '            'If Double.TryParse(tempSpt(3), tryDouble) Then
            '            '    sortDouble.Add(Double.Parse(tempSpt(3)))
            '            '    sortDoubleType.Add(tempSpt(1))
            '            'Else
            '            '    sortString.Add(tempSpt(3))
            '            '    sortStringType.Add(tempSpt(1))
            '            'End If
            '            'sortValue.Add(sortListData(i))
            '        Else
            '            If Double.TryParse(tempSpt(3), tryDouble) Then
            '                sortDouble.Add(Double.Parse(tempSpt(3)))
            '                sortDoubleType.Add(tempSpt(1))
            '            Else
            '                sortString.Add(tempSpt(3))
            '                sortStringType.Add(tempSpt(1))
            '            End If
            '            'sortValue.Add(sortListData(i))
            '        End If
            '    End If
            'Next
            'sortListData.Clear()
            'sortListData = sortCompCSVData

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try
    End Sub

    ' ************************************************
    '
    ' csv一行分の形式にして返す
    '
    ' ************************************************
    Public Shared Function GetCsvLine(ByVal cadType As String,
                                      ByVal itemType As String,
                                      ByVal itemSuffix As String,
                                      ByVal itemValue As String,
                                      ByVal itemPrefix As String,
                                      ByVal itemSymbol As String,
                                      ByVal X As String,
                                      ByVal Y As String) As String

        Try

            Return cadType + clsDesignTool.m_SepValue _
                    + itemType + clsDesignTool.m_SepValue _
                    + itemSuffix + clsDesignTool.m_SepValue _
                    + itemValue + clsDesignTool.m_SepValue _
                    + itemPrefix + clsDesignTool.m_SepValue _
                    + itemSymbol + clsDesignTool.m_SepValue _
                    + X + clsDesignTool.m_SepValue _
                    + Y

        Catch ex As Exception
            MsgBox(ex.Message)
            Return ""
        End Try

    End Function

    ' ************************************************
    '
    ' Solidworksアイテムの日本語名を返す
    '
    ' ************************************************
    Public Shared Function GetSWNihongoType(ByVal itemType As String) As String

        GetSWNihongoType = Nothing
        Dim nihongo As String = " "

        Try
            Select Case itemType
                Case "swAngularDimension"
                    nihongo = "角度寸法"
                Case "swDiameterDimension"
                    nihongo = "直径寸法"
                Case "swRadialDimension"
                    nihongo = "半径寸法"
                Case "swLinearDimension"
                    nihongo = "直線寸法"
                Case "swChamferDimension"
                    nihongo = "面取り寸法"
                Case "swDatumTag"
                    nihongo = "データム記号"
                Case "swGTol"
                    nihongo = "幾何交差"
                Case "swNote"
                    nihongo = "注記・バルーン"
                Case "swSFSymbol"
                    nihongo = "表面粗さ"
                Case "swWeldSymbol"
                    nihongo = "溶接シンボル"
                Case "swHatching"
                    nihongo = "ハッチング"
                Case "swFaceHatching"
                    nihongo = "断面図ハッチング"
            End Select

            Return nihongo

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Function

    ' ************************************************
    '
    ' AutoCADアイテムの日本語名を返す
    '
    ' ************************************************
    Public Shared Function GetACADNihongoType(ByVal itemType As String) As String

        GetACADNihongoType = Nothing
        Dim nihongo As String = " "

        Try
            Select Case itemType
                Case "AcDb2LineAngularDimension"
                    nihongo = "角度寸法"
                Case "AcDbDiametricDimension"
                    nihongo = "直径寸法"
                Case "AcDbRadialDimension"
                    nihongo = "半径寸法"
                Case "AcDbRotatedDimension"
                    nihongo = "回転寸法"
                Case "AcDbText"
                    nihongo = "テキスト"
                Case "AcDbMText"
                    nihongo = "マルチテキスト"
                Case "AcDbFcf"
                    nihongo = "幾何交差"
            End Select

            Return nihongo

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Function

    Public Shared Function GetReplaceText(ByVal text As String) As String

        Try

            text = text.Replace(vbCrLf, "")
            text = text.Replace(" ", "")
            text = text.Replace("　", "")
            text = text.Replace(vbLf, "")

            Return text

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
            Return text
        End Try

    End Function

    ' ************************************************
    '
    ' 形式を整える
    ' sortMode ソート無し：false ソートあり：true
    ' headerType なし：-1 シングル：0 ダブル(SW-AC)：1 シングル(流用先)：2 シングル(流用元)：3 ダブル(SW-SW)：4
    '
    ' ************************************************
    Public Shared Sub SortCsvData(ByVal sortMode As Boolean,
                                  ByVal headerType As Integer,
                                  ByRef dataList As List(Of String))

        Try

            ' ソート処理
            If sortMode = True Then
                dataList.Sort()
            End If

            ' ヘッダー無し
            If headerType < 0 Then
                Return
            End If

            Dim tempData As New List(Of String)
            For i As Integer = 0 To dataList.Count - 1
                Dim value As String = dataList(i)
                tempData.Add(value)
            Next
            dataList.Clear()

            Select Case headerType
                Case 0
                    ' ヘッダーをセット
                    dataList.Add("SW A/C" + clsDesignTool.m_SepValue _
                                + "アイテム" + clsDesignTool.m_SepValue _
                                + "接頭" + clsDesignTool.m_SepValue _
                                + "値" + clsDesignTool.m_SepValue _
                                + "接尾" + clsDesignTool.m_SepValue _
                                + "記号" + clsDesignTool.m_SepValue _
                                + "座標X" + clsDesignTool.m_SepValue _
                                + "座標Y" + clsDesignTool.m_SepValue)
                Case 1
                    dataList.Add("SW" + clsDesignTool.m_SepValue _
                                + "アイテム" + clsDesignTool.m_SepValue _
                                + "接頭" + clsDesignTool.m_SepValue _
                                + "値" + clsDesignTool.m_SepValue _
                                + "接尾" + clsDesignTool.m_SepValue _
                                + "記号" + clsDesignTool.m_SepValue _
                                + "座標X" + clsDesignTool.m_SepValue _
                                + "座標Y" + clsDesignTool.m_SepValue _
                                + "SW A/C" + clsDesignTool.m_SepValue _
                                + "アイテム" + clsDesignTool.m_SepValue _
                                + "接頭" + clsDesignTool.m_SepValue _
                                + "値" + clsDesignTool.m_SepValue _
                                + "接尾" + clsDesignTool.m_SepValue _
                                + "記号" + clsDesignTool.m_SepValue _
                                + "座標X" + clsDesignTool.m_SepValue _
                                + "座標Y" + clsDesignTool.m_SepValue)
                Case 2
                    dataList.Add("モデル名" + clsDesignTool.m_SepValue + " " _
                                + "アイテム" + clsDesignTool.m_SepValue + " " _
                                + "座標X" + clsDesignTool.m_SepValue + " " _
                                + "座標Y" + clsDesignTool.m_SepValue + " " _
                                + "接頭" + clsDesignTool.m_SepValue + " " _
                                + "値" + clsDesignTool.m_SepValue + " " _
                                + "接尾" + clsDesignTool.m_SepValue + " " _
                                + "記号" + clsDesignTool.m_SepValue)
                Case 3
                    dataList.Add("図面1モデル名" + clsDesignTool.m_SepValue + " " _
                                + "アイテム" + clsDesignTool.m_SepValue + " " _
                                + "座標X" + clsDesignTool.m_SepValue + " " _
                                + "座標Y" + clsDesignTool.m_SepValue + " " _
                                + "接頭" + clsDesignTool.m_SepValue + " " _
                                + "値" + clsDesignTool.m_SepValue + " " _
                                + "接尾" + clsDesignTool.m_SepValue + " " _
                                + "記号" + clsDesignTool.m_SepValue + " " _
                                + "図面2モデル名" + clsDesignTool.m_SepValue + " " _
                                + "アイテム" + clsDesignTool.m_SepValue + " " _
                                + "座標X" + clsDesignTool.m_SepValue + " " _
                                + "座標Y" + clsDesignTool.m_SepValue + " " _
                                + "接頭" + clsDesignTool.m_SepValue + " " _
                                + "値" + clsDesignTool.m_SepValue + " " _
                                + "接尾" + clsDesignTool.m_SepValue + " " _
                                + "記号" + clsDesignTool.m_SepValue)
            End Select

            For i As Integer = 0 To tempData.Count - 1
                Dim value As String = tempData(i)
                dataList.Add(value)
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub

    ' ********************************************
    '
    ' 接頭表記
    '
    ' ********************************************

    Public Shared Function ChangePrefix(ByVal mode As String, ByVal prefix As String) As String

        ChangePrefix = prefix

        Try

            Dim tempPrefix As String = prefix
            Dim answerWord As String = ""
            Dim stockFlag As Boolean = False
            Dim stockValue As String = ""

            If prefix = "" Then
                Return ""
            End If

            If mode = "AC" Then

                'If String.Compare(prefix, "%%d", True) = 0 Then
                '    tempPrefix = "°"
                'ElseIf String.Compare(prefix, "%%p", True) = 0 Then
                '    tempPrefix = "±"
                'ElseIf String.Compare(prefix, "%%c", True) = 0 Then
                '    tempPrefix = "φ"
                'End If

                ' 一文字ずつ調べる
                For i As Integer = 0 To prefix.Length - 1

                    If prefix(i) = "{" Or prefix(i) = "}" Then
                        Continue For
                    End If

                    ' %%～チェック
                    If prefix(i) = "%" Then
                        If i < prefix.Length - 2 Then
                            If prefix(i + 1) = "%" Then
                                If String.Compare(prefix(i + 2), "d", True) = 0 Then
                                    answerWord += "°"
                                    i += 2
                                ElseIf String.Compare(prefix(i + 2), "p", True) = 0 Then
                                    answerWord += "±"
                                    i += 2
                                ElseIf String.Compare(prefix(i + 2), "c", True) = 0 Then
                                    answerWord += "φ"
                                    i += 2
                                Else
                                    ' %%v
                                    i += 2
                                End If
                                Continue For
                            End If
                        End If
                    End If

                    ' 不要文字
                    If prefix(i) = vbLf Then
                        Continue For
                    End If

                    ' 文字格納
                    answerWord += prefix(i)

                Next

            Else

                'If String.Compare(prefix, "<MOD-DEG>", True) = 0 Then
                '    tempPrefix = "°"
                'ElseIf String.Compare(prefix, "<MOD-PM>", True) = 0 Then
                '    tempPrefix = "±"
                'ElseIf String.Compare(prefix, "<MOD-DIAM>", True) = 0 Then
                '    tempPrefix = "φ"
                'End If

                ' 一文字ずつ調べる
                For i As Integer = 0 To tempPrefix.Length - 1

                    If tempPrefix(i) = "<" Then
                        stockFlag = True
                        Continue For
                    End If

                    If tempPrefix(i) = ">" Then
                        stockFlag = False
                        Select Case stockValue

                            Case "MOD-DEG"
                                answerWord += "°"
                            Case "MOD-DIAM"
                                answerWord += "φ"
                            Case "MOD-PM"
                                answerWord += "±"
                            Case Else
                                answerWord += stockValue

                        End Select
                        stockValue = ""
                        Continue For
                    End If

                    If stockFlag Then
                        stockValue += tempPrefix(i)
                        Continue For
                    End If

                    answerWord += tempPrefix(i)

                Next
            End If

            ChangePrefix = answerWord

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Function

    ' ********************************************
    '
    ' 接尾表記
    '
    ' ********************************************

    Public Shared Function ChangeAcadDimSuffix(ByVal suffix As String) As String

        ChangeAcadDimSuffix = suffix

        Try

            If suffix = "" Then Exit Function

            Dim answerWord As String = ""
            Dim up As String = ""
            Dim down As String = ""
            Dim noLookFlag As Boolean = False
            Dim upFlag As Boolean = False
            Dim downFlag As Boolean = False

            ' 一文字ずつ調べる
            For i As Integer = 0 To suffix.Length - 1

                If suffix(i) = "{" Or suffix(i) = "}" Then
                    If suffix(i) = "}" Then
                        noLookFlag = False ' 初期化
                        upFlag = False
                        downFlag = False
                    End If
                    Continue For
                End If

                ' \Hチェック
                If suffix(i) = "\" Then
                    If i < suffix.Length - 1 Then
                        If suffix(i + 1) = "H" Then
                            noLookFlag = True
                            Continue For
                        End If
                    End If
                End If

                ' \A2or\A0チェック
                If suffix(i) = "\" Then
                    If i < suffix.Length - 4 Then
                        If suffix(i + 1) = "A" Then
                            If suffix(i + 2) = "2" Then
                                up = suffix(i + 4)
                            ElseIf suffix(i + 2) = "0" Then
                                down = suffix(i + 4)
                            End If
                            i += 4
                            noLookFlag = True
                            Continue For
                        End If
                    End If
                End If

                ' \Sチェック
                If suffix(i) = "\" Then
                    If i < suffix.Length - 1 Then
                        If suffix(i + 1) = "S" Then
                            i += 1
                            upFlag = True
                            Continue For
                        End If
                    End If
                End If

                ' ^チェック
                If suffix(i) = "^" Then
                    upFlag = False
                    downFlag = True
                    Continue For
                End If

                ' %%～チェック
                If suffix(i) = "%" Then
                    If i < suffix.Length - 2 Then
                        If suffix(i + 1) = "%" Then
                            Dim tempStr As String = suffix(i) + suffix(i + 1) + suffix(i + 2)
                            answerWord += clsDCCommon.ChangePrefix("AC", tempStr)
                            i += 2
                            Continue For
                        End If
                    End If
                End If

                If suffix(i) = ";" Then
                    If noLookFlag Then
                        noLookFlag = False
                    End If
                    If downFlag = True Then
                        downFlag = False
                    End If
                    Continue For
                End If

                If noLookFlag = False Then
                    If upFlag = True Then
                        answerWord += up + suffix(i)
                        up = ""
                    ElseIf downFlag = True Then
                        answerWord += down + suffix(i)
                        down = ""
                    Else
                        answerWord += suffix(i)
                    End If
                End If

            Next
            'For i As Integer = 0 To suffix.Length - 1

            '    If suffix(i) = "{" Or suffix(i) = "}" Then
            '        If suffix(i) = "}" Then
            '            noLookFlag = False ' 初期化
            '            upFlag = False
            '            downFlag = False
            '        End If
            '        Continue For
            '    End If

            '    ' \Hチェック
            '    If suffix(i) = "\" Then
            '        If i < suffix.Length - 1 Then
            '            If String.Compare(suffix(i + 1), "H", True) = 0 Then
            '                noLookFlag = True
            '                Continue For
            '            End If
            '        End If
            '    End If

            '    ' \A2or\A0チェック
            '    If suffix(i) = "\" Then
            '        If i < suffix.Length - 4 Then
            '            If String.Compare(suffix(i + 1), "A", True) Then
            '                If suffix(i + 2) = "2" Then
            '                    up = suffix(i + 4)
            '                ElseIf suffix(i + 2) = "0" Then
            '                    down = suffix(i + 4)
            '                End If
            '                i += 4
            '                noLookFlag = True
            '                Continue For
            '            End If
            '        End If
            '    End If

            '    ' \Sチェック
            '    If suffix(i) = "\" Then
            '        If i < suffix.Length - 1 Then
            '            If String.Compare(suffix(i + 1), "S", True) Then
            '                i += 1
            '                upFlag = True
            '                Continue For
            '            End If
            '        End If
            '    End If

            '    ' ^チェック
            '    If suffix(i) = "^" Then
            '        upFlag = False
            '        downFlag = True
            '        Continue For
            '    End If

            '    ' %%～チェック
            '    If suffix(i) = "%" Then
            '        If i < suffix.Length - 2 Then
            '            If suffix(i + 1) = "%" Then
            '                Dim tempStr As String = suffix(i) + suffix(i + 1) + suffix(i + 2)
            '                answerWord += clsDCCommon.ChangePrefix("AC", tempStr)
            '                i += 2
            '                Continue For
            '            End If
            '        End If
            '    End If

            '    If suffix(i) = ";" Then
            '        If noLookFlag Then
            '            noLookFlag = False
            '        End If
            '        If downFlag = True Then
            '            downFlag = False
            '        End If
            '        Continue For
            '    End If

            '    If noLookFlag = False Then
            '        If upFlag = True Then
            '            answerWord += up + suffix(i)
            '            up = ""
            '        ElseIf downFlag = True Then
            '            answerWord += down + suffix(i)
            '            down = ""
            '        Else
            '            answerWord += suffix(i)
            '        End If
            '    End If

            'Next

            ChangeAcadDimSuffix = answerWord

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Function

    ' ********************************************
    '
    ' 接尾表記
    '
    ' ********************************************

    Public Shared Function ChangeAcadGtol(ByVal gtol As String) As String

        ChangeAcadGtol = gtol

        Try

            If gtol = "" Then Exit Function

            Dim answerWord As String = ""

            ' 一文字ずつ調べる
            For i As Integer = 0 To gtol.Length - 1

                If gtol(i) = "{" Or gtol(i) = "}" Then
                    Continue For
                End If

                ' \Hチェック
                If gtol(i) = "\" Then
                    If i < gtol.Length - 6 Then
                        If gtol(i + 1) = "F" Then

                            If String.Compare(gtol(i + 6), "a", True) = 0 Then
                                answerWord += "ANGULAR"
                            ElseIf String.Compare(gtol(i + 6), "b", True) = 0 Then
                                answerWord += "PERP"
                            ElseIf String.Compare(gtol(i + 6), "c", True) = 0 Then
                                answerWord += "FLAT"
                            ElseIf String.Compare(gtol(i + 6), "d", True) = 0 Then
                                answerWord += "SPROF"
                            ElseIf String.Compare(gtol(i + 6), "e", True) = 0 Then
                                answerWord += "CIRC"
                            ElseIf String.Compare(gtol(i + 6), "f", True) = 0 Then
                                answerWord += "PARA"
                            ElseIf String.Compare(gtol(i + 6), "g", True) = 0 Then
                                answerWord += "CYL"
                            ElseIf String.Compare(gtol(i + 6), "h", True) = 0 Then
                                answerWord += "SRUN"
                            ElseIf String.Compare(gtol(i + 6), "j", True) = 0 Then
                                answerWord += "POSI"
                            ElseIf String.Compare(gtol(i + 6), "k", True) = 0 Then
                                answerWord += "LPROF"
                            ElseIf String.Compare(gtol(i + 6), "t", True) = 0 Then
                                answerWord += "TRUN"
                            ElseIf String.Compare(gtol(i + 6), "r", True) = 0 Then
                                answerWord += "CONC"
                            ElseIf String.Compare(gtol(i + 6), "u", True) = 0 Then
                                answerWord += "STRAIGHT"
                            ElseIf String.Compare(gtol(i + 6), "i", True) = 0 Then
                                answerWord += "SYMMETRY"
                            ElseIf String.Compare(gtol(i + 6), "m", True) = 0 Then
                                answerWord += "MMC"
                            ElseIf String.Compare(gtol(i + 6), "s", True) = 0 Then
                                answerWord += "FMC"
                            ElseIf String.Compare(gtol(i + 6), "l", True) = 0 Then
                                answerWord += "LMC"
                            End If

                            i += 6
                            Continue For
                        End If
                    End If
                End If

                ' %%～チェック
                If gtol(i) = "%" Then
                    If i < gtol.Length - 2 Then
                        If gtol(i + 1) = "%" Then
                            If String.Compare(gtol(i + 2), "d", True) = 0 Then
                                answerWord += "°"
                                i += 2
                            ElseIf String.Compare(gtol(i + 2), "p", True) = 0 Then
                                answerWord += "±"
                                i += 2
                            ElseIf String.Compare(gtol(i + 2), "c", True) = 0 Then
                                answerWord += "φ"
                                i += 2
                            Else
                                ' %%v
                                i += 2
                            End If
                            Continue For
                        End If
                    End If
                End If

                ' 不要文字
                If gtol(i) = vbLf Then
                    Continue For
                End If

                ' 文字格納
                answerWord += gtol(i)

            Next

            ChangeAcadGtol = answerWord

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Function

    Public Shared Function ChangeSolidGtol(ByVal gtol As String) As String

        ChangeSolidGtol = gtol

        Try

            If gtol = "" Then Exit Function

            Dim answerWord As String = ""
            Dim stockFlag As Boolean = False
            Dim stockValue As String = ""

            ' 一文字ずつ調べる
            For i As Integer = 0 To gtol.Length - 1

                If gtol(i) = "<" Then
                    stockFlag = True
                    Continue For
                End If

                If gtol(i) = ">" Then
                    stockFlag = False
                    Select Case stockValue

                        Case "IGTOL-CIRC"
                            answerWord += "CIRC"
                        Case "IGTOL-CONC"
                            answerWord += "CONC"
                        Case "IGTOL-PARA"
                            answerWord += "PARA"
                        Case "IGTOL-FLAT"
                            answerWord += "FLAT"
                        Case "IGTOL-PERP"
                            answerWord += "PERP"
                        Case "IGTOL-POSI"
                            answerWord += "POSI"
                        Case "IGTOL-STRAIGHT"
                            answerWord += "STRAIGHT"
                        Case "IGTOL-CYL"
                            answerWord += "CYL"
                        Case "IGTOL-LPROF"
                            answerWord += "LPROF"
                        Case "IGTOL-SPROF"
                            answerWord += "SPROF"
                        Case "IGTOL-RERP"
                            answerWord += "RERP"
                        Case "IGTOL-ANGULAR"
                            answerWord += "ANGULAR"
                        Case "IGTOL-SRUN"
                            answerWord += "SRUN"
                        Case "IGTOL-TRUN"
                            answerWord += "TRUN"
                        Case "IGTOL-SYMMETRY"
                            answerWord += "SYMMETRY"

                        Case "MOD-DEG"
                            answerWord += "°"
                        Case "MOD-DIAM"
                            answerWord += "φ"
                        Case "MOD-PM"
                            answerWord += "±"
                        Case "MOD-MMC"
                            answerWord += "MMC"
                        Case "MOD-FMC"
                            answerWord += "FMC"
                        Case "MOD-LMC"
                            answerWord += "LMC"
                        Case Else
                            answerWord += stockValue

                    End Select
                    stockValue = ""
                    Continue For
                End If

                If stockFlag Then
                    stockValue += gtol(i)
                    Continue For
                End If

                answerWord += gtol(i)

            Next

            ChangeSolidGtol = answerWord

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Function

End Class
