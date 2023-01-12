' ************************************************
'
' SolidWorks関係処理クラス
'
' ************************************************

Imports System.Collections.Generic
Imports SolidWorks.Interop.sldworks
Imports System.IO
Imports System.Collections
Imports System
Imports SolidWorks.Interop.swconst

Public Class clsSolidControl

    Private csvDirectoryPath As String = ""
    Private csvSeparate As String = ""
    Private iniFilePath As String = ""
    Private iniFileName As String = "MaterialExtraction.ini"
    Dim m_AutoCadData As List(Of clsCadData) = Nothing
    Dim m_AutoCadDataPath As String = ""
    ' ************************************************
    '
    ' コンストラクタ
    '
    ' ************************************************

    Public Sub New(ByVal autoCadData As List(Of clsCadData))
        m_AutoCadData = autoCadData
    End Sub

    ' ************************************************
    '
    ' デストラクタ
    '
    ' ************************************************

    Protected Overrides Sub Finalize()

        MyBase.Finalize()

    End Sub

    ' ************************************************
    '
    ' SolidWorksの図面から文字列情報をすべて取得して新規csvファイルに出力する
    ' (SolidWorks同士の比較)
    '
    ' ************************************************
    Public Function DrawingCompDataSolid(ByVal directoryPath As String,
                                         ByVal oldSolidName As String,
                                         ByVal newSolidName As String,
                                         ByRef compCSVData As List(Of String),
                                         ByRef noCompCSVData As List(Of String)) As Boolean

        DrawingCompDataSolid = False

        Try

            Dim swApp As SldWorks = Nothing
            'Dim swModel As ModelDoc2 = Nothing
            Dim swDraw As DrawingDoc = Nothing
            Dim swView As View = Nothing
            Dim sNoteText As String = 0
            Dim sNoteCount As Integer = 0
            Dim annotations As New List(Of Annotation)
            Dim annotations2 As New List(Of Annotation)

            'Dim oldHatching As Integer = 0
            'Dim newHatching As Integer = 0

            'Dim oldAnnotations As New List(Of Annotation)
            Dim newAnnotations As New List(Of Annotation)
            Dim newAnnotations2 As New List(Of Annotation)

            Dim annotationColors As New List(Of Integer)
            Dim annotationItarics As New List(Of Boolean)
            Dim annotationFontNames As New List(Of String)
            Dim annotationDocumentFlag As New List(Of Boolean)

            ' フォントの大きさ
            Dim annotationSize As New List(Of Double)

            swApp = clsDCCommon.GetSldWorksApp()
            Dim swModels As Object = swApp.GetDocuments()
            Dim swOldModel As ModelDoc2 = Nothing
            Dim swNewModel As ModelDoc2 = Nothing

            For ii As Integer = 0 To swApp.GetDocumentCount - 1
                Dim tmpModel As ModelDoc2 = swModels(ii)
                If tmpModel.GetType = 3 Then
                    'Dim title As String = tmpModel.GetTitle()
                    Dim modelName As String = System.IO.Path.GetFileNameWithoutExtension(tmpModel.GetPathName)
                    If oldSolidName = modelName Then
                        swOldModel = tmpModel
                    End If
                    If newSolidName = modelName Then
                        swNewModel = tmpModel
                    End If
                End If
            Next

            If oldSolidName Is Nothing Or swNewModel Is Nothing Then
                MsgBox("比較対象図面が取得出来ませんでした。", MsgBoxStyle.ApplicationModal, "図面比較")
                Return False
            End If

            ' 一致リスト初期化
            If compCSVData Is Nothing Then
                compCSVData = New List(Of String)
            Else
                compCSVData.Clear()
            End If

            ' 不一致リスト初期化
            If noCompCSVData Is Nothing Then
                noCompCSVData = New List(Of String)
            Else
                noCompCSVData.Clear()
            End If

            Dim oldCsvData As New List(Of String)
            Dim newCsvData As New List(Of String)

            Dim oldHatchList As New List(Of SketchHatch)
            Dim newHatchList As New List(Of SketchHatch)

            Dim oldHatchListf As New List(Of FaceHatch)
            Dim newHatchListf As New List(Of FaceHatch)

            Dim swAnnotations As Object = Nothing                       ' アノテートアイテム
            Dim swDisplayDimensions() As DisplayDimension = Nothing     ' 寸法
            Dim swSFSymbols() As SFSymbol = Nothing                     ' シンボル
            Dim swNotes() As Note = Nothing                             ' 注記・バルーン
            Dim swDatumTags() As DatumTag = Nothing                     ' データム記号
            Dim usedNo As New List(Of Integer)

            Dim swLayerMgr As LayerMgr = Nothing
            Dim layerList As Object = Nothing
            Dim oldLayerName As List(Of String) = New List(Of String)
            Dim newLayerName As List(Of String) = New List(Of String)

            Dim oldLayerNamef As List(Of String) = New List(Of String)
            Dim newLayerNamef As List(Of String) = New List(Of String)

            Dim oldfaceLayerName As String = ""
            Dim newfaceLayerName As String = ""

            Dim oldLayerTName As List(Of String) = New List(Of String)
            Dim newLayerTName As List(Of String) = New List(Of String)
            Dim oldLayerT As New List(Of TableAnnotation)
            Dim newLayerT As New List(Of TableAnnotation)

            Dim modelNameOld As String = ""
            Dim modelNameNew As String = ""

            Dim color1 As Integer = RGB(255, 0, 0)
            Dim color2 As Integer = RGB(0, 255, 0)

            Dim thred As Integer = -1

            'sketchhatch number
            Dim scount As Integer = 0
            Dim scount2 As Integer = 0

            Dim facehatchflg As Integer = 0

            Dim OldfacehatchingData As New List(Of String)
            Dim NewfacehatchingData As New List(Of String)

            Dim SheetChange As Object
            SheetChange = New clsSheetChange()
            Dim ary As ArrayList = SheetChange.LoadIniData("[SheetColor1]")
            If ary IsNot Nothing Then
                Dim arr() As String = ary(0).split(",")
                If arr.Length = 3 Then
                    color1 = RGB(CInt(arr(0)), CInt(arr(1)), CInt(arr(2)))
                End If
            End If

            Dim ary2 As ArrayList = SheetChange.LoadIniData("[SheetColor2]")
            If ary2 IsNot Nothing Then
                Dim arr() As String = ary2(0).split(",")
                If arr.Length = 3 Then
                    color2 = RGB(CInt(arr(0)), CInt(arr(1)), CInt(arr(2)))
                End If
            End If

            ' しきい値を取得
            Dim ary3 As ArrayList = SheetChange.LoadIniData("[Threshold]")
            If ary3 IsNot Nothing Then
                thred = CInt(ary3(0))
            End If

            swApp.ArrangeWindows(2)

            Dim sheetNames As Object = Nothing

            'ブロック分解処理
            For i As Integer = 0 To 1

                If i = 0 Then
                    swDraw = swOldModel
                Else
                    swDraw = swNewModel
                End If

                For j As Integer = 0 To 2
                    sheetNames = swDraw.GetSheetNames()
                    Dim sketchMgr As SketchManager = swDraw.SketchManager

                    Dim skBlocks As Object = sketchMgr.GetSketchBlockDefinitions
                    If Not skBlocks Is Nothing Then

                        For Each skblockDef As SketchBlockDefinition In skBlocks

                            Dim skInsList As Object = skblockDef.GetInstances()
                            If Not skInsList Is Nothing Then

                                For Each skIns As SketchBlockInstance In skInsList
                                    sketchMgr.ExplodeSketchBlockInstance(skIns)
                                Next

                            End If
                        Next
                    End If
                Next

            Next

            For i As Integer = 0 To 1

                usedNo = Nothing
                usedNo = New List(Of Integer)

                If i = 0 Then
                    swDraw = swOldModel
                    modelNameOld = System.IO.Path.GetFileNameWithoutExtension(swOldModel.GetPathName)
                    swLayerMgr = swOldModel.GetLayerManager()
                    layerList = swLayerMgr.GetLayerList
                    swOldModel.ViewZoomtofit2()
                Else
                    swDraw = swNewModel
                    modelNameNew = System.IO.Path.GetFileNameWithoutExtension(swNewModel.GetPathName)
                    swLayerMgr = swNewModel.GetLayerManager()
                    layerList = swLayerMgr.GetLayerList
                    swNewModel.ViewZoomtofit2()
                    'Dim blocks As Object = swDraw.GetBlockDefinitions()
                    'For Each block In blocks

                    'Next
                    'swDraw.ExplodeBlockInstance()


                End If

                Dim solidCsvData As New List(Of String)
                Dim hatchingCsvData As New List(Of String)
                Dim hatchingCsvData2 As New List(Of String)
                Dim hatchList As New List(Of SketchHatch)
                Dim hatchList2 As New List(Of FaceHatch)
                Dim tableList As New List(Of TableAnnotation)

                sheetNames = Nothing
                sheetNames = swDraw.GetSheetNames()
                Dim ansStr As String
                Dim posStr As String
                Dim modelName As String = ""

                'For j As Integer = 0 To swDraw.GetSheetCount - 1

                '    swDraw.ActivateSheet(sheetNames(j))

                ' ビューの内容を順番に取得 
                swView = swDraw.GetFirstView

                ' ビューの数分ループ
                While Not swView Is Nothing

                    swAnnotations = swView.GetAnnotations
                    If Not swAnnotations Is Nothing Then
                        For Each swAnnotation As Annotation In swAnnotations

                            Dim itemType As String = ""
                            Dim itemText As String = ""
                            Dim itemSuffix As String = ""
                            Dim itemPrefix As String = ""
                            Dim itemSymbol As String = ""

                            If swAnnotation Is Nothing Then
                                Continue For
                            End If

                            Dim tempObj As Object = swAnnotation.GetPosition()
                            If tempObj Is Nothing Then
                                Continue For
                            End If
                            If Double.Parse(tempObj(0)) = 0.0 Or Double.Parse(tempObj(1)) = 0.0 Then
                                Continue For
                            End If
                            Dim LinearFlag As Integer = 0

                            Select Case swAnnotation.GetType

                                Case swAnnotationType_e.swNote              ' 注記・バルーン

                                    Try
                                        Dim tempNote As Note = swAnnotation.GetSpecificAnnotation
                                        itemType = "swNote"
                                        itemText = tempNote.GetText()
                                    Catch ex As Exception
                                        MsgBox(ex.Message)
                                    End Try

                                Case swAnnotationType_e.swSFSymbol          ' 表面粗さ

                                    Try
                                        itemType = "swSFSymbol"
                                        Dim tempSymbol As SFSymbol = swAnnotation.GetSpecificAnnotation
                                        ' 後で変える
                                        'itemText = tempSymbol.GetText(swDimensionTextParts_e.swDimensionTextPrefix)

                                        For k As Integer = 0 To tempSymbol.GetTextCount - 1
                                            itemText += tempSymbol.GetTextAtIndex(k)
                                        Next

                                    Catch ex As Exception
                                        MsgBox(ex.Message)
                                    End Try
                                Case swAnnotationType_e.swDisplayDimension  ' 寸法線

                                    Dim swModel As ModelDoc2 = Nothing
                                    If i = 0 Then
                                        swModel = swOldModel
                                    Else
                                        swModel = swNewModel
                                    End If
                                    If GetDimensionData(swModel, swAnnotation,
                                                            itemType, itemText, itemPrefix, itemSuffix, itemSymbol) = False Then
                                        Continue For
                                    End If

                                Case swAnnotationType_e.swDatumTag          ' データム記号
                                    Try
                                        itemType = "swDatumTag"
                                        Dim tempDatumTag As DatumTag = swAnnotation.GetSpecificAnnotation
                                        For k As Integer = 0 To tempDatumTag.GetTextCount - 1
                                            itemText += tempDatumTag.GetTextAtIndex(k)
                                        Next
                                    Catch ex As Exception
                                        MsgBox(ex.Message)
                                    End Try
                                Case swAnnotationType_e.swGTol              ' 幾何交差
                                    Try
                                        itemType = "swGTol"
                                        Dim tempGTol As Gtol = swAnnotation.GetSpecificAnnotation
                                        ' itemText = tempGTol.GetText(swDimensionTextParts_e.swDimensionTextAll)
                                        For k As Integer = 0 To tempGTol.GetTextCount - 1
                                            If k = 0 Then

                                                Try
                                                    itemText += tempGTol.GetFrameSymbols2(1)(0)
                                                Catch ex As Exception
                                                    itemText += tempGTol.GetTextAtIndex(k)
                                                End Try
                                            Else
                                                itemText += tempGTol.GetTextAtIndex(k)
                                            End If
                                        Next

                                        itemText = clsDCCommon.ChangeSolidGtol(itemText)

                                    Catch ex As Exception
                                        MsgBox(ex.Message)
                                    End Try

                                Case swAnnotationType_e.swWeldSymbol ' 溶接記号
                                    Try
                                        itemType = "swWeldSymbol"
                                        Dim tempWeldSymbol As WeldSymbol = swAnnotation.GetSpecificAnnotation
                                        ' itemText = tempGTol.GetText(swDimensionTextParts_e.swDimensionTextAll)
                                        For k As Integer = 0 To tempWeldSymbol.GetTextCount - 1
                                            itemText += tempWeldSymbol.GetTextAtIndex(k)
                                        Next

                                        itemText = clsDCCommon.ChangeSolidGtol(itemText)

                                    Catch ex As Exception
                                        MsgBox(ex.Message)
                                    End Try

                                Case Else
                                    Continue For
                            End Select

                            Try
                                ' 空文字列しか無いものは飛ばす
                                If itemText = "" Then
                                    Continue For
                                End If

                                ' 改行文字削除
                                itemPrefix = clsDCCommon.GetReplaceText(itemPrefix)
                                itemText = clsDCCommon.GetReplaceText(itemText)
                                itemSuffix = clsDCCommon.GetReplaceText(itemSuffix)
                                itemSymbol = clsDCCommon.GetReplaceText(itemSymbol)

                                ' 全角文字をすべて半角にすべて変換
                                itemPrefix = StrConv(itemPrefix, VbStrConv.Narrow)
                                itemText = StrConv(itemText, VbStrConv.Narrow)
                                itemSuffix = StrConv(itemSuffix, VbStrConv.Narrow)
                                itemSymbol = StrConv(itemSymbol, VbStrConv.Narrow)

                                ' SolidWorks側のCSV出力用
                                If Not swAnnotation Is Nothing Then
                                    tempObj = Nothing
                                    tempObj = swAnnotation.GetPosition()
                                    ' SolidWorks側
                                    Dim nihongo As String = clsDCCommon.GetSWNihongoType(itemType)
                                    ansStr = ""
                                    posStr = ""
                                    tempObj = Nothing
                                    tempObj = swAnnotation.GetPosition()

                                    If i = 0 Then
                                        modelName = swOldModel.GetPathName
                                    Else
                                        modelName = swNewModel.GetPathName
                                    End If
                                    modelName = System.IO.Path.GetFileNameWithoutExtension(modelName)
                                    'modelName = modelName.Replace(" - Model", "")

                                    If Not tempObj Is Nothing Then
                                        posStr = (Math.Round(Double.Parse(tempObj(0) * 1000.0), 3).ToString _
                                                       + clsDesignTool.m_SepValue + " " +
                                                        Math.Round(Double.Parse(tempObj(1) * 1000.0), 3).ToString) + clsDesignTool.m_SepValue + " "
                                    Else
                                        posStr = "0.0" + clsDesignTool.m_SepValue + " " + "0.0" + clsDesignTool.m_SepValue + " "
                                    End If

                                    ansStr += modelName + clsDesignTool.m_SepValue + " " +
                                                nihongo + clsDesignTool.m_SepValue + " " +
                                                posStr _
                                                + clsDCCommon.ChangePrefix("SW", itemPrefix) + clsDesignTool.m_SepValue + " " +
                                                itemText + clsDesignTool.m_SepValue + " " +
                                                itemSuffix + clsDesignTool.m_SepValue + " " +
                                                itemSymbol + clsDesignTool.m_SepValue + " "

                                    'If Not tempObj Is Nothing Then
                                    '    ansStr += (Math.Round(Double.Parse(tempObj(0) * 1000.0), 3).ToString _
                                    '               + clsDesignTool.m_SepValue +
                                    '                Math.Round(Double.Parse(tempObj(1) * 1000.0), 3).ToString)
                                    'Else
                                    '    ansStr += "0.0" + clsDesignTool.m_SepValue + "0.0"
                                    'End If

                                    solidCsvData.Add(ansStr)
                                    ' 新しい方の図面のアノテート情報を保持
                                    If i = 1 Then
                                        annotations.Add(swAnnotation)
                                    End If

                                    If i = 0 Then
                                        annotations2.Add(swAnnotation)
                                    End If

                                End If
                            Catch ex As Exception
                                MsgBox(ex.Message)
                            End Try
                        Next
                    End If

                    Dim ansStrList2 As List(Of String) = New List(Of String)

                    ' ハッチング情報取得
                    Dim hatchData As List(Of String) = GetHatching(swView, hatchList, hatchList2, ansStrList2)
                    If hatchData.Count <> 0 Then
                        For k As Integer = 0 To hatchData.Count - 1
                            ansStr = modelName + clsDesignTool.m_SepValue + " " _
                                + clsDCCommon.GetSWNihongoType("swHatching") + " " _
                                + "" + clsDesignTool.m_SepValue _
                                + "" + clsDesignTool.m_SepValue _
                                + "" + clsDesignTool.m_SepValue _
                                + hatchData(k) + clsDesignTool.m_SepValue _
                                + "" + clsDesignTool.m_SepValue _
                                + "" + clsDesignTool.m_SepValue

                            hatchingCsvData.Add(ansStr)
                        Next
                    End If
                    If ansStrList2.Count <> 0 Then
                        For k As Integer = 0 To ansStrList2.Count - 1
                            ansStr = modelName + clsDesignTool.m_SepValue + " " _
                                + clsDCCommon.GetSWNihongoType("swFaceHatching") + " " _
                                + "" + clsDesignTool.m_SepValue _
                                + "" + clsDesignTool.m_SepValue _
                                + "" + clsDesignTool.m_SepValue _
                                + ansStrList2(k) + clsDesignTool.m_SepValue _
                                + "" + clsDesignTool.m_SepValue _
                                + "" + clsDesignTool.m_SepValue

                            hatchingCsvData2.Add(ansStr)
                        Next
                    End If

                    Dim table As TableAnnotation = swView.GetFirstTableAnnotation()
                    'Dim instance As IAnnotation

                    If table IsNot Nothing Then
                        Do While (1)
                            If table Is Nothing Then Exit Do
                            tableList.Add(table)
                            table = table.GetNext
                        Loop
                    End If

                    'Dim objFeature As Feature = swView.FirstFeature()
                    'If objFeature IsNot Nothing Then
                    '    Do While (1)
                    '        If objFeature Is Nothing Then Exit Do

                    '        objFeature = objFeature.GetNextFeature
                    '    Loop
                    'End If

                    swView = swView.GetNextView
                End While

                'Next
                swDraw.ActivateSheet(sheetNames(0))
                swDraw.EditRebuild()

                If hatchList.Count <> 0 Then
                    Do
                        For Each hatch In hatchList
                            Dim layflg As Integer = 0
                            For Each lay In layerList
                                Dim currentLayer As Layer = swLayerMgr.GetLayer(lay)
                                If currentLayer.Name = hatch.Layer.ToString() Then
                                    If i = 0 Then
                                        swDraw.CreateLayer2("COMP" + currentLayer.Name, "", CInt(color1), currentLayer.Style, currentLayer.Width, True, True)
                                        oldLayerName.Add("COMP" + currentLayer.Name)
                                        layflg = 1
                                        'Exit Do
                                    Else
                                        swDraw.CreateLayer2("COMP" + currentLayer.Name, "", CInt(color2), currentLayer.Style, currentLayer.Width, True, True)
                                        newLayerName.Add("COMP" + currentLayer.Name)
                                        layflg = 1
                                        'Exit Do
                                    End If

                                    Exit For
                                End If
                            Next

                            If i = 0 Then
                                If layflg = 0 Then
                                    For Each lay In layerList
                                        Dim currentLayer As Layer = swLayerMgr.GetLayer(lay)
                                        swDraw.CreateLayer2("COMP" + currentLayer.Name, "", CInt(color1), currentLayer.Style, currentLayer.Width, True, True)
                                        oldLayerName.Add("COMP" + currentLayer.Name)
                                        Exit For
                                    Next
                                End If
                            Else
                                If layflg = 0 Then
                                    For Each lay In layerList
                                        Dim currentLayer As Layer = swLayerMgr.GetLayer(lay)
                                        swDraw.CreateLayer2("COMP" + currentLayer.Name, "", CInt(color1), currentLayer.Style, currentLayer.Width, True, True)
                                        newLayerName.Add("COMP" + currentLayer.Name)
                                        Exit For
                                    Next
                                End If
                            End If
                        Next

                        Exit Do
                    Loop
                End If

                If hatchList2.Count <> 0 Then
                    Do
                        For Each hatch In hatchList2
                            Dim layflg As Integer = 0
                            For Each lay In layerList
                                Dim currentLayer As Layer = swLayerMgr.GetLayer(lay)
                                If currentLayer.Name = hatch.Layer.ToString() Then
                                    If i = 0 Then
                                        swDraw.CreateLayer2("COMP" + currentLayer.Name, "", CInt(color1), currentLayer.Style, currentLayer.Width, True, True)
                                        oldLayerNamef.Add("COMP" + currentLayer.Name)
                                        layflg = 1
                                        'oldfaceLayerName = "COMP" + currentLayer.Name
                                        'Exit Do
                                    Else
                                        swDraw.CreateLayer2("COMP" + currentLayer.Name, "", CInt(color2), currentLayer.Style, currentLayer.Width, True, True)
                                        newLayerNamef.Add("COMP" + currentLayer.Name)
                                        layflg = 1
                                        'newfaceLayerName = "COMP" + currentLayer.Name
                                        'Exit Do
                                    End If

                                    Exit For
                                End If
                            Next

                            If i = 0 Then
                                'If (oldfaceLayerName = "") Then
                                If layflg = 0 Then
                                    For Each lay In layerList
                                        Dim currentLayer As Layer = swLayerMgr.GetLayer(lay)
                                        swDraw.CreateLayer2("COMP" + currentLayer.Name, "", CInt(color1), currentLayer.Style, currentLayer.Width, True, True)
                                        oldLayerNamef.Add("COMP" + currentLayer.Name)
                                        'oldfaceLayerName = "COMP" + currentLayer.Name
                                        Exit For
                                    Next
                                End If
                            Else
                                'If (newfaceLayerName = "") Then
                                If layflg = 0 Then
                                    For Each lay In layerList
                                        Dim currentLayer As Layer = swLayerMgr.GetLayer(lay)
                                        swDraw.CreateLayer2("COMP" + currentLayer.Name, "", CInt(color1), currentLayer.Style, currentLayer.Width, True, True)
                                        newLayerNamef.Add("COMP" + currentLayer.Name)
                                        'newfaceLayerName = "COMP" + currentLayer.Name
                                        Exit For
                                    Next
                                End If
                            End If
                        Next

                        Exit Do
                    Loop
                End If

                If tableList.Count <> 0 Then
                    For Each table In tableList
                        For Each lay In layerList
                            Dim currentLayer As Layer = swLayerMgr.GetLayer(lay)
                            If currentLayer.Name = table.GetAnnotation.Layer Then
                                If i = 0 Then
                                    swDraw.CreateLayer2("COMPT" + currentLayer.Name, "", CInt(color1), currentLayer.Style, currentLayer.Width, True, True)
                                    oldLayerTName.Add("COMPT" + currentLayer.Name)
                                Else
                                    swDraw.CreateLayer2("COMPT" + currentLayer.Name, "", CInt(color2), currentLayer.Style, currentLayer.Width, True, True)
                                    newLayerTName.Add("COMPT" + currentLayer.Name)
                                End If
                            End If
                        Next
                    Next
                End If

                If (newLayerName.Count > 0 Or oldLayerName.Count > 0 Or newfaceLayerName IsNot "" Or oldfaceLayerName IsNot "") Then
                    For j As Integer = 0 To hatchingCsvData.Count - 1
                        solidCsvData.Add(hatchingCsvData(j))
                    Next
                    For j As Integer = 0 To hatchingCsvData2.Count - 1
                        solidCsvData.Add(hatchingCsvData2(j))
                    Next

                    If i = 0 Then
                        scount = hatchingCsvData.Count
                    Else
                        scount2 = hatchingCsvData.Count
                    End If

                    If (hatchList.Count > 0) Then
                        If i = 0 Then
                            oldHatchList = hatchList
                        Else
                            newHatchList = hatchList
                        End If
                    End If

                    If (hatchList2.Count > 0) Then
                        If i = 0 Then
                            oldHatchListf = hatchList2
                        Else
                            newHatchListf = hatchList2
                        End If
                    End If
                End If


                If i = 0 Then
                    oldLayerT = tableList
                Else
                    newLayerT = tableList
                End If



                ' データ格納
                For j As Integer = 0 To solidCsvData.Count - 1

                    Dim tempStr As String = solidCsvData(j)
                    If i = 0 Then
                        oldCsvData.Add(tempStr)
                    Else
                        newCsvData.Add(tempStr)
                    End If
                Next

            Next

            ' 表示用
            Dim viewOldData As New List(Of String)
            Dim viewNewData As New List(Of String)

            Dim oldHatchList2 As New List(Of SketchHatch)
            Dim newHatchList2 As New List(Of SketchHatch)

            Dim oldHatchListf2 As New List(Of FaceHatch)
            Dim newHatchListf2 As New List(Of FaceHatch)

            Dim oldLayerName2 As New List(Of String)
            Dim newLayerName2 As New List(Of String)

            Dim oldLayerNamef2 As New List(Of String)
            Dim newLayerNamef2 As New List(Of String)

            Dim compCSVData2 As List(Of String) = New List(Of String)

            Dim oldTableList As New List(Of TableAnnotation)
            Dim newTableList As New List(Of TableAnnotation)

            If oldLayerT.Count > newLayerT.Count Then
                oldTableList = oldLayerT
            End If

            If oldLayerT.Count < newLayerT.Count Then
                newTableList = newLayerT
            End If

            'For i As Integer = 0 To oldLayerT.Count - 1
            '    Dim existTable As Boolean = False
            '    For j As Integer = 0 To newLayerT.Count - 1
            '        If oldLayerT(i).Title = newLayerT(j).Title Then
            '            existTable = True
            '        End If
            '    Next
            '    If existTable = False Then
            '        oldTableList.Add(oldLayerT(i))
            '        oldLayerTName2.Add(oldLayerTName(i))
            '    End If
            'Next

            'For i As Integer = 0 To newLayerT.Count - 1
            '    Dim existTable As Boolean = False
            '    For j As Integer = 0 To oldLayerT.Count - 1
            '        If newLayerT(i).Title = oldLayerT(j).Title Then
            '            existTable = True
            '        End If
            '    Next
            '    If existTable = False Then
            '        newTableList.Add(newLayerT(i))
            '        newLayerTName2.Add(newLayerTName(i))
            '    End If
            'Next

            ' 比較処理（流用先基準）
            ' noCompCSVData(new)を格納
            Dim CompCsv2 As String = ""
            usedNo = Nothing
            usedNo = New List(Of Integer)
            Dim ilist As List(Of Integer) = New List(Of Integer)
            For i As Integer = 0 To newCsvData.Count - 1
                Dim safeFlag As Boolean = False
                Dim safeFlag2 As Boolean = False
                Dim newSpt() As String = Split(newCsvData(i), clsDesignTool.m_SepValue)
                'Dim used As Integer = 0
                Dim ct As Integer = 0
                Dim xp As Double = thred
                'Dim yp As Double = 200

                For j As Integer = 0 To oldCsvData.Count - 1
                    Dim isSafe = False
                    For k As Integer = 0 To usedNo.Count - 1
                        If j = usedNo(k) Then
                            isSafe = True
                            Exit For
                        End If
                    Next
                    If isSafe = True Then
                        Continue For
                    End If

                    Dim oldSpt() As String = Split(oldCsvData(j), clsDesignTool.m_SepValue)
                    'If annotations.Count - 1 >= i Then

                    If oldSpt(1) = newSpt(1) And
                        oldSpt(2) = newSpt(2) And
                        oldSpt(3) = newSpt(3) And
                        oldSpt(4) = newSpt(4) And
                        oldSpt(5) = newSpt(5) And
                        oldSpt(6) = newSpt(6) And
                        oldSpt(7) = newSpt(7) Then

                        ' 一致（new→oldの順に並べる）
                        'compCSVData2.Add(newCsvData(j) + clsDesignTool.m_SepValue + oldCsvData(i))

                        ' 使用済no格納
                        usedNo.Add(j)
                        ' 一致フラグON
                        safeFlag = True

                        If safeFlag2 = True Then
                            safeFlag2 = False
                        End If

                        Exit For

                    End If

                    If oldSpt(2) = newSpt(2) And
                           oldSpt(3) = newSpt(3) And
                           oldSpt(5) IsNot newSpt(5) Then
                        ' 使用済no格納
                        usedNo.Add(j)
                        safeFlag = False
                        safeFlag2 = False

                        Exit For

                    End If

                    If oldSpt(1) = newSpt(1) And
                           oldSpt(4) = newSpt(4) And
                           oldSpt(5) = newSpt(5) And
                           oldSpt(6) = newSpt(6) And
                           oldSpt(7) = newSpt(7) Then

                        'If Math.Abs(CDbl(oldSpt(2)) - CDbl(newSpt(2))) < 50 And Math.Abs(CDbl(oldSpt(3)) - CDbl(newSpt(3))) < 50 Then
                        'CompCsv2 = newCsvData(i) + clsDesignTool.m_SepValue + oldCsvData(j)
                        ' 使用済no格納
                        'usedNo.Add(j)

                        If (thred < -1 And ct = 0) Then
                            CompCsv2 = newCsvData(i) + clsDesignTool.m_SepValue + oldCsvData(j)
                            xp = Math.Abs(CDbl(oldSpt(2)) - CDbl(newSpt(2))) + Math.Abs(CDbl(oldSpt(3)) - CDbl(newSpt(3)))
                            ct = 1
                        Else
                            If (Math.Abs(CDbl(oldSpt(2)) - CDbl(newSpt(2))) + Math.Abs(CDbl(oldSpt(3)) - CDbl(newSpt(3))) <= xp) Then
                                'used = j
                                CompCsv2 = newCsvData(i) + clsDesignTool.m_SepValue + oldCsvData(j)
                                xp = Math.Abs(CDbl(oldSpt(2)) - CDbl(newSpt(2))) + Math.Abs(CDbl(oldSpt(3)) - CDbl(newSpt(3)))
                            End If
                        End If

                        ' 一致フラグON
                        safeFlag2 = True
                        'Exit For
                        'End If

                        ' 使用済no格納
                        'usedNo.Add(j)
                        '    ' 一致フラグON
                        '    safeFlag = True
                        '    Exit For

                    End If
                Next

                If safeFlag2 = True Then
                    'usedNo.Add(used)
                    compCSVData2.Add(CompCsv2)
                    ilist.Add(i)
                End If

                ' 不一致の場合リストに格納
                If safeFlag = False And safeFlag2 = False Then
                    If annotations.Count - 1 >= i Then
                        Dim tempAnno As Annotation = annotations(i)
                        newAnnotations.Add(tempAnno)
                    Else
                        If (newHatchList.Count > 0 And newHatchList.Count > i - annotations.Count) Then
                            newHatchList2.Add(newHatchList(i - annotations.Count))
                            newLayerName2.Add(newLayerName(i - annotations.Count))
                        ElseIf (newHatchListf.Count > 0 And newHatchListf.Count > i - annotations.Count - scount2) Then
                            newHatchListf2.Add(newHatchListf(i - annotations.Count - scount2))
                            newLayerNamef2.Add(newLayerNamef(i - annotations.Count - scount2))
                        End If

                    End If

                        viewNewData.Add(newCsvData(i))
                    noCompCSVData.Add(newCsvData(i))
                End If

            Next

            ' 比較処理（流用元基準）
            ' noCompCSVData(old)とその他出力データを格納
            Dim compCSVData3 As List(Of String) = New List(Of String)
            Dim CompCsv As String = ""
            usedNo = Nothing
            usedNo = New List(Of Integer)
            Dim ilist2 As List(Of Integer) = New List(Of Integer)

            For i As Integer = 0 To oldCsvData.Count - 1
                Dim safeFlag As Boolean = False
                Dim safeFlag2 As Boolean = False
                Dim oldSpt() As String = Split(oldCsvData(i), clsDesignTool.m_SepValue)
                'Dim used As Integer = 0
                Dim ct As Integer = 0
                Dim xp As Double = thred

                For j As Integer = 0 To newCsvData.Count - 1
                    Dim isSafe = False
                    For k As Integer = 0 To usedNo.Count - 1
                        If j = usedNo(k) Then
                            isSafe = True
                            Exit For
                        End If
                    Next
                    If isSafe = True Then
                        Continue For
                    End If

                    Dim newSpt() As String = Split(newCsvData(j), clsDesignTool.m_SepValue)
                    'If annotations.Count - 1 >= i Then

                    If oldSpt(1) = newSpt(1) And
                           oldSpt(2) = newSpt(2) And
                           oldSpt(3) = newSpt(3) And
                           oldSpt(4) = newSpt(4) And
                           oldSpt(5) = newSpt(5) And
                           oldSpt(6) = newSpt(6) And
                           oldSpt(7) = newSpt(7) Then

                        ' 一致（new→oldの順に並べる）
                        compCSVData.Add(newCsvData(j) + clsDesignTool.m_SepValue + oldCsvData(i))
                        ' 使用済no格納
                        usedNo.Add(j)
                        ' 一致フラグON
                        safeFlag = True

                        If safeFlag2 = True Then
                            safeFlag2 = False
                        End If
                        Exit For

                    End If

                    If oldSpt(2) = newSpt(2) And
                           oldSpt(3) = newSpt(3) And
                           oldSpt(5) IsNot newSpt(5) Then
                        ' 使用済no格納
                        usedNo.Add(j)
                        safeFlag = False
                        safeFlag2 = False

                        Exit For

                    End If

                    If oldSpt(1) = newSpt(1) And
                           oldSpt(4) = newSpt(4) And
                           oldSpt(5) = newSpt(5) And
                           oldSpt(6) = newSpt(6) And
                           oldSpt(7) = newSpt(7) Then

                        'If Math.Abs(CDbl(oldSpt(2)) - CDbl(newSpt(2))) < 50 And Math.Abs(CDbl(oldSpt(3)) - CDbl(newSpt(3))) < 50 Then
                        ' 一致（new→oldの順に並べる）
                        'CompCsv = newCsvData(j) + clsDesignTool.m_SepValue + oldCsvData(i)
                        'compCSVData.Add(newCsvData(j) + clsDesignTool.m_SepValue + oldCsvData(i))
                        ' 使用済no格納
                        'usedNo.Add(j)
                        'If (Math.Abs(CDbl(oldSpt(2)) - CDbl(newSpt(2))) <= xp And Math.Abs(CDbl(oldSpt(3)) - CDbl(newSpt(3))) <= yp) Then

                        If (thred < -1 And ct = 0) Then
                            CompCsv = newCsvData(j) + clsDesignTool.m_SepValue + oldCsvData(i)
                            xp = Math.Abs(CDbl(oldSpt(2)) - CDbl(newSpt(2))) + Math.Abs(CDbl(oldSpt(3)) - CDbl(newSpt(3)))
                            ct = 1
                        Else
                            If (Math.Abs(CDbl(oldSpt(2)) - CDbl(newSpt(2))) + Math.Abs(CDbl(oldSpt(3)) - CDbl(newSpt(3))) <= xp) Then
                                'used = j
                                CompCsv = newCsvData(j) + clsDesignTool.m_SepValue + oldCsvData(i)
                                xp = Math.Abs(CDbl(oldSpt(2)) - CDbl(newSpt(2))) + Math.Abs(CDbl(oldSpt(3)) - CDbl(newSpt(3)))
                            End If
                        End If

                        ' 一致フラグON
                        safeFlag2 = True

                        'Exit For
                        'End If

                        '' 一致（new→oldの順に並べる）
                        'compCSVData.Add(newCsvData(j) + clsDesignTool.m_SepValue + oldCsvData(i))
                        '' 使用済no格納
                        'usedNo.Add(j)
                        '' 一致フラグON
                        'safeFlag = True

                        'Exit For

                    End If
                Next

                If safeFlag2 = True Then
                    'usedNo.Add(used)
                    'compCSVData.Add(CompCsv)
                    compCSVData3.Add(CompCsv)
                    ilist2.Add(i)
                End If

                ' 不一致の場合リストに格納
                If safeFlag = False And safeFlag2 = False Then
                    If annotations2.Count - 1 >= i Then
                        Dim tempAnno As Annotation = annotations2(i)
                        'Dim tempFormat As TextFormat = Nothing
                        newAnnotations2.Add(tempAnno)
                    Else
                        'oldHatchList2.Add(oldHatchList(i - annotations2.Count))
                        'oldLayerName2.Add(oldLayerName(i - annotations2.Count))
                        If (oldHatchList.Count > 0 And oldHatchList.Count > i - annotations2.Count) Then
                            oldHatchList2.Add(oldHatchList(i - annotations2.Count))
                            'oldLayerName2.Add(oldLayerName(i - annotations2.Count))
                            oldLayerName2.Add(oldLayerName(i - annotations2.Count))
                        ElseIf (oldHatchListf.Count > 0 And oldHatchListf.Count > i - annotations2.Count - scount) Then
                            oldHatchListf2.Add(oldHatchListf(i - annotations2.Count - scount))
                            oldLayerNamef2.Add(oldLayerNamef(i - annotations2.Count - scount))
                        End If
                    End If

                    viewOldData.Add(oldCsvData(i))
                    noCompCSVData.Add(oldCsvData(i))

                End If

            Next

            For i As Integer = 0 To compCSVData2.Count - 1
                For j As Integer = 0 To compCSVData3.Count - 1
                    If compCSVData2(i) = compCSVData3(j) Then
                        compCSVData.Add(compCSVData2(i))
                        compCSVData2(i) = "0"
                        compCSVData3(j) = "0"
                        Exit For
                    End If
                Next
            Next

            Dim oldCsvData2 As New List(Of String)
            Dim newCsvData2 As New List(Of String)
            Dim ilist1_2 As List(Of Integer) = New List(Of Integer)
            Dim ilist2_2 As List(Of Integer) = New List(Of Integer)

            For i As Integer = 0 To compCSVData2.Count - 1
                If compCSVData2(i) IsNot "0" Then
                    newCsvData2.Add(newCsvData(ilist(i)))
                    ilist1_2.Add(ilist(i))
                End If
            Next

            For i As Integer = 0 To compCSVData3.Count - 1
                If compCSVData3(i) IsNot "0" Then
                    oldCsvData2.Add(oldCsvData(ilist2(i)))
                    ilist2_2.Add(ilist2(i))
                End If
            Next

            For i As Integer = 0 To newCsvData2.Count - 1
                If newCsvData2(i) = "0" Then
                    Continue For
                End If
                Dim newSpt() As String = Split(newCsvData2(i), clsDesignTool.m_SepValue)
                Dim ct As Integer = 0
                Dim xp As Integer = thred
                Dim newd As Integer = 0
                Dim oldd As Integer = 0
                CompCsv2 = Nothing


                For j As Integer = 0 To oldCsvData2.Count - 1
                    If oldCsvData2(j) = "0" Then
                        Continue For
                    End If
                    Dim oldSpt() As String = Split(oldCsvData2(j), clsDesignTool.m_SepValue)

                    If oldSpt(1) = newSpt(1) And
                           oldSpt(4) = newSpt(4) And
                           oldSpt(5) = newSpt(5) And
                           oldSpt(6) = newSpt(6) And
                           oldSpt(7) = newSpt(7) Then

                        If (thred < 0 And ct = 0) Then
                            CompCsv2 = newCsvData2(i) + clsDesignTool.m_SepValue + oldCsvData2(j)
                            newd = i
                            oldd = j
                            xp = Math.Abs(CDbl(oldSpt(2)) - CDbl(newSpt(2))) + Math.Abs(CDbl(oldSpt(3)) - CDbl(newSpt(3)))
                            ct = 1
                        Else
                            If (Math.Abs(CDbl(oldSpt(2)) - CDbl(newSpt(2))) + Math.Abs(CDbl(oldSpt(3)) - CDbl(newSpt(3))) <= xp) Then
                                newd = i
                                oldd = j
                                CompCsv2 = newCsvData2(i) + clsDesignTool.m_SepValue + oldCsvData2(j)
                                xp = Math.Abs(CDbl(oldSpt(2)) - CDbl(newSpt(2))) + Math.Abs(CDbl(oldSpt(3)) - CDbl(newSpt(3)))
                            End If
                        End If
                    End If
                Next

                If CompCsv2 IsNot Nothing Then
                    compCSVData.Add(CompCsv2)
                    newCsvData2(newd) = "0"
                    oldCsvData2(oldd) = "0"
                End If
            Next

            For i As Integer = 0 To newCsvData2.Count - 1
                If newCsvData2(i) IsNot "0" Then
                    newAnnotations.Add(annotations(ilist1_2(i)))
                    viewNewData.Add(newCsvData(ilist1_2(i)))
                    noCompCSVData.Add(newCsvData(ilist1_2(i)))
                End If
            Next

            For i As Integer = 0 To oldCsvData2.Count - 1
                If oldCsvData2(i) IsNot "0" Then
                    newAnnotations2.Add(annotations2(ilist2_2(i)))
                    viewOldData.Add(oldCsvData(ilist2_2(i)))
                    noCompCSVData.Add(oldCsvData(ilist2_2(i)))
                End If
            Next

            If oldLayerT.Count <> 0 Or newLayerT.Count <> 0 Then
                If oldTableList.Count <> 0 Then
                    noCompCSVData.Add(modelNameOld + clsDesignTool.m_SepValue + " " + "部品表")
                ElseIf newTableList.Count <> 0 Then
                    noCompCSVData.Add(modelNameNew + clsDesignTool.m_SepValue + " " + "部品表")
                Else
                    noCompCSVData.Add(modelNameNew + clsDesignTool.m_SepValue + " " + "部品表" + clsDesignTool.m_SepValue + " " + modelNameOld + clsDesignTool.m_SepValue + " " + "部品表")
                End If
            End If


            swApp.ActivateDoc(swOldModel.GetPathName)

            If viewOldData.Count > 0 Then

                ' 書式変更
                For j As Integer = 0 To newAnnotations2.Count - 1

                    If newAnnotations2(j) Is Nothing Then
                        Continue For
                    End If

                    Dim formatCount As Integer = newAnnotations2(j).GetTextFormatCount
                    For k As Integer = 0 To formatCount - 1

                        Try

                            Dim tempFormat As TextFormat = newAnnotations2(j).GetTextFormat(k)
                            newAnnotations2(j).SetTextFormat(k, False, tempFormat)
                            newAnnotations2(j).Color = CInt(color1)
                        Catch ex As Exception
                            Dim Str As String = ex.Message
                        End Try
                    Next
                Next
            End If

            If oldHatchList2.Count > 0 Then
                For j As Integer = 0 To oldHatchList2.Count - 1
                    oldHatchList2(j).Layer = oldLayerName2(j)
                Next
            End If

            If oldHatchListf2.Count > 0 Then
                facehatchflg = 1
                For j As Integer = 0 To oldHatchListf2.Count - 1
                    oldHatchListf2(j).Layer = oldLayerNamef2(j)
                    'oldHatchListf2(j).Layer = oldfaceLayerName
                Next
            End If

            If oldTableList.Count > 0 Then
                For j As Integer = 0 To oldTableList.Count - 1
                    oldTableList(j).GetAnnotation.Layer = oldLayerTName(j)
                Next
            End If

            swApp.ActivateDoc(swNewModel.GetPathName)

            If viewNewData.Count > 0 Then

                ' 書式変更
                For j As Integer = 0 To newAnnotations.Count - 1

                    If newAnnotations(j) Is Nothing Then
                        Continue For
                    End If

                    Dim formatCount As Integer = newAnnotations(j).GetTextFormatCount
                    For k As Integer = 0 To formatCount - 1

                        Try

                            Dim tempFormat As TextFormat = newAnnotations(j).GetTextFormat(k)
                            newAnnotations(j).SetTextFormat(k, False, tempFormat)
                            newAnnotations(j).Color = CInt(color2)
                        Catch ex As Exception
                            Dim Str As String = ex.Message
                        End Try
                    Next
                Next


            End If

            If newHatchList2.Count > 0 Then
                For j As Integer = 0 To newHatchList2.Count - 1
                    newHatchList2(j).Layer = newLayerName2(j)
                Next
            End If

            If newHatchListf2.Count > 0 Then
                facehatchflg = 1
                For j As Integer = 0 To newHatchListf2.Count - 1
                    newHatchListf2(j).Layer = newLayerNamef2(j)
                    'newHatchListf2(j).Layer = newfaceLayerName
                Next
            End If

            If newTableList.Count > 0 Then
                For j As Integer = 0 To newTableList.Count - 1
                    newTableList(j).GetAnnotation.Layer = newLayerTName(j)
                Next
            End If

            swOldModel.ForceRebuild3(False)
            swNewModel.ForceRebuild3(False)

            swOldModel.Save()
            swNewModel.Save()

            If facehatchflg = 1 Then
                Dim oldName As String = swOldModel.GetPathName
                Dim newName As String = swNewModel.GetPathName

                swApp.CloseAllDocuments(True)
                swApp.OpenDoc(oldName, swDocumentTypes_e.swDocDRAWING)
                swApp.OpenDoc(newName, swDocumentTypes_e.swDocDRAWING)

                swApp.ActivateDoc(oldName)
                swApp.ActivateDoc(newName)

                swModels = swApp.GetDocuments()

                For ii As Integer = 0 To swApp.GetDocumentCount - 1
                    Dim tmpModel As ModelDoc2 = swModels(ii)
                    If tmpModel.GetType = 3 Then
                        'Dim title As String = tmpModel.GetTitle()
                        Dim modelName2 As String = System.IO.Path.GetFileNameWithoutExtension(tmpModel.GetPathName)
                        If oldSolidName = modelName2 Then
                            swOldModel = tmpModel
                        End If
                        If newSolidName = modelName2 Then
                            swNewModel = tmpModel
                        End If
                    End If
                Next

                swOldModel.ForceRebuild3(False)
                swNewModel.ForceRebuild3(False)

                swOldModel.Save()
                swNewModel.Save()

                swApp.ArrangeWindows(2)

                swOldModel.ViewZoomtofit2()
                swNewModel.ViewZoomtofit2()

                System.Windows.Forms.MessageBox.Show("断面図のハッチングが存在します。" & vbCrLf &
                                                     "SW APIの問題で、色が変わらない可能性があるので、" & vbCrLf &
                                                     "色が変わっていないハッチングを目視で確認してください。")

            End If

            Return True

        Catch ex As Exception
            MsgBox("DrawingCompData" & "関数にて予期せぬエラーが発生しました。" + vbCrLf + ex.Message, MsgBoxStyle.Critical)
        Finally
        End Try

    End Function

    ' ******************************************
    '
    ' 0を指定桁数まで足す関数
    '
    ' ******************************************
    Private Function MakeZero(ByVal targetStr As String, ByVal precision As Integer) As String

        MakeZero = targetStr

        Try
            Dim tempCount As Integer = 0
            Dim tempFlag As Boolean = False
            For k = 0 To targetStr.Length - 1
                If tempFlag Then
                    tempCount += 1
                End If
                If targetStr(k) = "." Then
                    tempFlag = True
                End If
            Next
            If tempFlag Then
                For k = 1 To precision
                    If tempCount < k Then
                        targetStr += "0"
                    End If
                Next
            End If

            MakeZero = targetStr

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try

    End Function

    '寸法データ取得
    Private Function GetDimensionData(ByVal swModel As ModelDoc2,
                                        ByVal annotation As Annotation,
                                        ByRef itemType As String,
                                        ByRef itemText As String,
                                        ByRef itemPrefix As String,
                                        ByRef itemSuffix As String,
                                        ByRef itemSymbol As String) As Boolean

        GetDimensionData = False

        Try
            'Try
            Dim tempDimension As DisplayDimension = annotation.GetSpecificAnnotation
            'itemText = tempDimension.GetText(swDimensionTextParts_e.swDimensionTextAll)

            Dim dimen As Dimension = tempDimension.GetDimension
            Dim tempValues() As Double = dimen.GetValue3(1, "")


            'Add by touga
            Dim Precision As Integer = tempDimension.GetPrimaryPrecision2
            If Precision = swDimensionPrecisionSettings_e.swPrecisionFollowsDocumentSetting Then
                Precision = swModel.Extension.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingLinearDimPrecision, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
            End If

            Dim torePrecision As Integer = tempDimension.GetPrimaryTolPrecision2
            If torePrecision < 0 Then
                torePrecision = swModel.Extension.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swDetailingLinearTolPrecision, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
                If torePrecision < 0 Then
                    ' 寸法値と同じ
                    torePrecision = Precision
                End If
            End If

            If Precision >= 0 And Precision <= 15 Then
                For k As Integer = 0 To tempValues.Length - 1
                    'Add by touga
                    Dim RoundVal As Double = Math.Round(tempValues(k), Precision)
                    'itemText += tempValues(k).ToString + " "
                    itemText += RoundVal.ToString
                Next
            End If

            ' itemText = tempDimension.GetDimension.Getvalue3(1, "")
            Select Case tempDimension.Type2
                Case swDimensionType_e.swDimensionTypeUnknown
                    itemType = "swDimensionTypeUnknown"
                Case swDimensionType_e.swOrdinateDimension
                    itemType = "swOrdinateDimension"
                Case swDimensionType_e.swLinearDimension
                    itemType = "swLinearDimension"
                Case swDimensionType_e.swAngularDimension
                    itemType = "swAngularDimension"
                Case swDimensionType_e.swArcLengthDimension
                    itemType = "swArcLengthDimension"
                Case swDimensionType_e.swRadialDimension
                    itemType = "swRadialDimension"
                Case swDimensionType_e.swDiameterDimension
                    itemType = "swDiameterDimension"
                Case swDimensionType_e.swHorOrdinateDimension
                    itemType = "swHorOrdinateDimension"
                Case swDimensionType_e.swVertOrdinateDimension
                    itemType = "swVertOrdinateDimension"
                Case swDimensionType_e.swZAxisDimension
                    itemType = "swZAxisDimension"
                Case swDimensionType_e.swChamferDimension
                    itemType = "swChamferDimension"

                    ' C面取り
                    Dim length As Double = 0
                    Dim angle As Double = 0
                    dimen.GetSystemChamferValues(length, angle)
                    Dim ansValue As Double = Math.Sin(angle) * length * 1000.0
                    itemText = Math.Round(ansValue, Precision).ToString

                Case swDimensionType_e.swHorLinearDimension
                    itemType = "swHorLinearDimension"
                Case swDimensionType_e.swVertLinearDimension
                    itemType = "swVertLinearDimension"
                Case swDimensionType_e.swScalarDimension
                    itemType = "swScalarDimension"
                Case Else
                    Return False
            End Select

            ' 接頭接尾表記取得
            itemPrefix = tempDimension.GetText(swDimensionTextParts_e.swDimensionTextPrefix)

            ' 接尾表記取得
            Dim tempDblValues() As Double = dimen.GetToleranceValues
            If tempDblValues IsNot Nothing Then
                If tempDblValues.Length > 1 Then
                    'If dimen.GetToleranceType <> 0 Then
                    If dimen.GetToleranceType = 2 Or dimen.GetToleranceType = 8 Then
                        ' 上下寸法
                        itemSuffix = dimen.Tolerance.GetShaftFitValue
                        itemSuffix += dimen.Tolerance.GetHoleFitValue
                        itemSuffix += clsDCCommon.ChangePrefix("SW", tempDimension.GetText(swDimensionTextParts_e.swDimensionTextSuffix))

                        If itemType = "swAngularDimension" Then
                            itemSuffix += "°"
                        End If

                        Dim tempStr1 As String = Math.Round((tempDblValues(1) * 1000.0), 3).ToString()
                        If tempStr1 <> "" And tempStr1 <> "0" Then
                            tempStr1 = MakeZero(tempStr1, torePrecision)
                            If Double.Parse(tempStr1) > 0.0 Then
                                tempStr1 = "+" + tempStr1
                            End If
                        End If

                        Dim tempStr2 As String = Math.Round((tempDblValues(0) * 1000.0), 3).ToString()
                        If tempStr2 <> "" And tempStr2 <> "0" Then
                            tempStr2 = MakeZero(tempStr2, torePrecision)
                            If Double.Parse(tempStr2) > 0.0 Then
                                tempStr2 = "+" + tempStr2
                            End If
                        End If

                        itemSuffix += tempStr1
                        itemSuffix += tempStr2
                    Else
                        If dimen.GetToleranceType <> 0 Then
                            itemSuffix = dimen.Tolerance.GetShaftFitValue
                            itemSuffix += dimen.Tolerance.GetHoleFitValue
                            itemSuffix += clsDCCommon.ChangePrefix("SW", tempDimension.GetText(swDimensionTextParts_e.swDimensionTextSuffix))

                            If itemType = "swAngularDimension" Then
                                itemSuffix += "°"
                            End If

                            Dim tempStr1 As String = Math.Round((tempDblValues(1) * 1000.0), 3).ToString
                            Dim tempStr2 As String = Math.Round((tempDblValues(0) * 1000.0), 3).ToString

                            If dimen.GetToleranceType = 4 Then
                                If tempStr1 <> "0" Then
                                    ' 普通許容差
                                    itemSuffix += "±"
                                    tempStr1 = MakeZero(tempStr1, torePrecision)
                                    itemSuffix += tempStr1
                                End If
                            ElseIf dimen.GetToleranceType = 5 Then
                                ' MIN
                                itemSuffix += "min."
                            ElseIf dimen.GetToleranceType = 6 Then
                                ' MAX
                                itemSuffix += "max."
                            Else

                                ' C面取り
                                If itemType = "swChamferDimension" Then
                                    If tempDimension.ChamferTextStyle = 4 Then
                                        itemPrefix += "C"
                                    End If
                                End If

                                If tempStr1 <> "0" Then
                                    tempStr1 = MakeZero(tempStr1, torePrecision)
                                    itemSuffix += tempStr1
                                End If
                                If tempStr2 <> "0" Then
                                    tempStr2 = MakeZero(tempStr2, torePrecision)
                                    itemSuffix += tempStr2
                                End If
                            End If
                        Else
                            'itemSuffix = dimen.Tolerance.GetShaftFitValue
                            'itemSuffix += dimen.Tolerance.GetHoleFitValue
                            'itemSuffix += clsDCCommon.ChangePrefix("SW", tempDimension.GetText(swDimensionTextParts_e.swDimensionTextSuffix))

                            itemSuffix = ""
                            If itemType = "swAngularDimension" Then
                                itemSuffix += "°"
                            End If

                            If itemType = "swChamferDimension" Then
                                If tempDimension.ChamferTextStyle = 4 Then
                                    itemPrefix += "C"
                                End If
                            End If

                        End If
                    End If
                End If
            Else
                If itemType = "swAngularDimension" Then
                    itemSuffix += "°"
                End If
            End If

            If tempDimension.ShowParenthesis Then
                itemSymbol = "()"
            End If

            If dimen.GetToleranceType = 1 Then
                itemSymbol += "□"
            End If

            GetDimensionData = True

        Catch ex As Exception
            MsgBox(ex.Message, vbExclamation)
        End Try


    End Function

    '穴寸法テキスト取得
    Public Function GetHoleVariables(ByVal swModel As ModelDoc2, swDisplayDimension As DisplayDimension) As String

        GetHoleVariables = ""

        Dim swCalloutLengthVariable As CalloutLengthVariable
        Dim swCalloutAngleVariable As CalloutAngleVariable
        Dim swCalloutStringVariable As CalloutStringVariable
        Dim swCalloutVariable As CalloutVariable
        Dim i As Integer
        Dim vType As Integer
        Dim holeVariables As Object
        Dim str1 As String = ""
        Dim str2 As String = ""
        Dim ansStr As String = ""

        Try
            holeVariables = swDisplayDimension.GetHoleCalloutVariables
            ansStr += (UBound(holeVariables) + 1).ToString
            Debug.Print("")

            'Determine type of hole callout variable and get and set some values
            For i = 0 To UBound(holeVariables)
                swCalloutVariable = holeVariables(i)
                ansStr += " " + swCalloutVariable.VariableName
                ansStr += " " + swCalloutVariable.UserReadableVariableName
                ansStr += " " + swCalloutVariable.Type.ToString
                vType = swCalloutVariable.Type
                If vType = swCalloutVariableType_e.swCalloutVariableType_Length Then
                    swCalloutLengthVariable = swCalloutVariable
                    ansStr += " " + i.ToString
                    ansStr += " " + str1
                    ansStr += " " + str2
                    ansStr += " " + swCalloutLengthVariable.Length.ToString
                    ansStr += " " + swCalloutLengthVariable.Precision.ToString
                    ansStr += " " + swCalloutLengthVariable.TolerancePrecision.ToString
                    swCalloutLengthVariable.Precision = swCalloutLengthVariable.Precision - 1 - i
                    ansStr += " " + swCalloutLengthVariable.Precision.ToString
                    swCalloutVariable.ToleranceType = swTolType_e.swTolBILAT
                ElseIf vType = swCalloutVariableType_e.swCalloutVariableType_Angle Then
                    swCalloutAngleVariable = swCalloutVariable
                    ansStr += " " + i.ToString
                    ansStr += " " + str1
                    ansStr += " " + str2
                    ansStr += " " + swCalloutAngleVariable.Angle.ToString
                ElseIf vType = swCalloutVariableType_e.swCalloutVariableType_String Then
                    swCalloutStringVariable = swCalloutVariable
                    ansStr += " " + i.ToString
                    ansStr += " " + str1
                    ansStr += " " + str2
                    ansStr += " " + swCalloutStringVariable.String
                End If
            Next

            Return ansStr

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function

    ' ハッチング取得
    Public Function GetHatching(swView As View, ByRef hatchList As List(Of SketchHatch), ByRef hatchList2 As List(Of FaceHatch), ByRef ansStrList2 As List(Of String)) As List(Of String)

        Dim swSketch As Sketch
        Dim vFaceHatch As Object
        Dim swFaceHatch As FaceHatch
        Dim vSketchHatch As Object
        Dim swSketchHatch As SketchHatch
        Dim ansStr As String = ""
        Dim ansStrList As List(Of String) = New List(Of String)
        Dim swFace As SolidWorks.Interop.sldworks.Face2

        swSketch = swView.GetSketch
        vFaceHatch = swView.GetFaceHatches

        If Not vFaceHatch Is Nothing Then

            For Each swFaceHatch In vFaceHatch
                swFace = swFaceHatch.Face
                ansStr += swFace.GetArea.ToString
                'ansStr += (swFaceHatch.Angle * 57.3).ToString
                'ansStr += " " + swFaceHatch.Color.ToString
                'ansStr += " " + swFaceHatch.Definition.ToString
                ansStr += " " + swFaceHatch.Layer.ToString
                ansStr += " " + swFaceHatch.Pattern.ToString
                ansStr += " " + swFaceHatch.Scale2.ToString
                ansStr += " " + swFaceHatch.SolidFill.ToString

                hatchList2.Add(swFaceHatch)
                ansStrList2.Add(ansStr)
                swFaceHatch.UseMaterialHatch = False
                'swFaceHatch.UseMaterialHatch = False
                'swFaceHatch.Color = CInt(RGB(255, 0, 0))

                ansStr = ""
                'Exit For
            Next

        End If

        vSketchHatch = swSketch.GetSketchHatches

        If Not vSketchHatch Is Nothing Then

            For Each swSketchHatch In vSketchHatch
                swFace = swSketchHatch.IGetFace2
                ansStr += swFace.GetArea.ToString
                ansStr += " " + (swSketchHatch.Angle * 57.3).ToString
                'ansStr += " " + swSketchHatch.Color.ToString
                ansStr += " " + swSketchHatch.GetID(0).ToString + swSketchHatch.GetID(1).ToString
                ansStr += " " + swSketchHatch.Layer.ToString
                ansStr += " " + swSketchHatch.LayerOverride.ToString
                ansStr += " " + swSketchHatch.Pattern.ToString
                ansStr += " " + swSketchHatch.Scale2.ToString
                ansStr += " " + swSketchHatch.SolidFill.ToString

                hatchList.Add(swSketchHatch)
                ansStrList.Add(ansStr)

                ansStr = ""
                'Exit For
            Next
        End If

        Return ansStrList

    End Function

End Class
