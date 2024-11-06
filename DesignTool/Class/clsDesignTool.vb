Imports SolidWorks.Interop.sldworks
Imports System.Runtime.InteropServices
Imports System.Collections
Imports System.Collections.Generic

Public Class clsCadData

    'Public itemType As String   ' アイテムタイプ
    'Public itemValue As String  ' 値
    'Public X As String          ' X位置
    'Public Y As String          ' Y位置
    'Public Z As String          ' Z位置

    Public cadType As String    ' CADタイプ(SW or A/C)
    Public itemType As String   ' アイテムタイプ
    Public itemSuffix As String ' 接頭表記
    Public itemValue As String  ' 値
    Public itemPrefix As String ' 接尾表記
    Public itemSymbol As String ' 記号
    Public X As String          ' X位置
    Public Y As String          ' Y位置

End Class

Public Class clsDesignTool

    Public Const m_SepValue As String = Chr(9) ' タブ

    ' @(f)
    ' 機能      : SolidWorksアプリケーションクラスを取得
    ' 引き数    : なし
    ' 返り値    : なし
    ' 機能説明  :
    ' 備考      : IMAIZUMI ADD 2011/08/24
    Public Shared Function GetSWApp() As SldWorks
        Dim swApp As SldWorks
        swApp = GetObject(, "SldWorks.application.31")
        'GetObject(, "SldWorks.Application")
        Return swApp
    End Function

End Class
