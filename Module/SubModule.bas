Attribute VB_Name = "SubModule"
Option Explicit



#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private g_ExportWsName_ As String
Private g_ExportTempChartName_ As String
Private g_ExportOutputPath_ As String
Private g_ExportBusy_ As Boolean
Private g_ExportDone_ As Boolean
Private g_ExportLastError_ As String

Private Type PictureItem
    ShapeName As String
    TopPos As Double
    LeftPos As Double
End Type


Public Enum BrowserType
    BrowserType_Edge = 1
    BrowserType_Chrome = 2
End Enum

' UTF-8 テキストを保存する
Public Sub SaveTextUtf8(ByVal filePath As String, ByVal text As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2 'text
    st.Charset = "utf-8"
    st.Open
    st.WriteText text
    st.SaveToFile filePath, 2 'overwrite
    st.Close
End Sub

' UTF-8 テキストを読み込む
Public Function LoadTextUtf8(ByVal filePath As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2 'text
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile filePath
    LoadTextUtf8 = st.ReadText
    st.Close
End Function

' ダウンロード完了後に指定名へ変更して保存先パスを返す
Public Function DownloadAndSaveAs( _
    drv As IWebDriver, _
    elem As IWebElement, _
    downloadFolder As String, _
    ByVal fileName As String, _
    Optional waitSec As Long = 120 _
) As String

    EnsureFolderExists downloadFolder

    ' allowAndName=True でダウンロード監視を開始する
    Call drv.DownloadWatchStart(downloadFolder, True)

    elem.Click
    Call drv.DownloadWatchWaitCompleted(waitSec)

    Dim suggested As String
    suggested = drv.DownloadWatchSuggestedFilename()

    Dim actualPath As String
    actualPath = drv.DownloadWatchFilePath()

    Dim folderPath As String
    Dim newName As String
    Dim ext As String
    Dim newPath As String

    folderPath = left$(actualPath, InStrRev(actualPath, "\"))

    newName = fileName
    If InStrRev(newName, ".") = 0 Then
        ext = Mid$(suggested, InStrRev(suggested, "."))
        If ext = "" Or left$(ext, 1) <> "." Then
            ext = Mid$(actualPath, InStrRev(actualPath, "."))
        End If
        newName = newName & ext
    End If

    newPath = folderPath & newName

    If Dir$(newPath, vbNormal) <> "" Then
        Err.Raise vbObjectError + 517, "DownloadAndSaveAs", "既に存在します: " & newPath
    End If

    If StrComp(actualPath, newPath, vbTextCompare) <> 0 Then
        Name actualPath As newPath
    End If

    DownloadAndSaveAs = newPath
End Function

' 指定フォルダが存在しない場合は階層ごとに作成する
Public Sub EnsureFolderExists(ByVal folderPath As String)
    Dim p As String, cur As String
    Dim parts() As String
    Dim i As Long

    p = Replace(folderPath, "/", "\")
    If Right$(p, 1) = "\" Then p = left$(p, Len(p) - 1)

    ' UNC パス(\\server\share\...) に対応する
    If left$(p, 2) = "\\" Then
        parts = Split(Mid$(p, 3), "\")
        cur = "\\" & parts(0) & "\" & parts(1)  ' \\server\share
        i = 2
    Else
        parts = Split(p, "\")
        cur = parts(0)                           ' C: など
        i = 1
    End If

    Do While i <= UBound(parts)
        cur = cur & "\" & parts(i)
        If Dir$(cur, vbDirectory) = "" Then MkDir cur
        i = i + 1
    Loop
End Sub

' フルパスからファイル名だけを返す
Public Function FileNameFromFullPath(ByVal fullPath As String) As String
    Dim p As String
    Dim pos As Long

    p = Replace(fullPath, "/", "\")        ' 区切りを \ に統一
    pos = InStrRev(p, "\")                 ' 最後の \ の位置

    If pos = 0 Then
        FileNameFromFullPath = p           ' 区切りが無い＝入力全体がファイル名
    ElseIf pos = Len(p) Then
        FileNameFromFullPath = ""          ' 末尾が \ で終わる（フォルダ指定など）
    Else
        FileNameFromFullPath = Mid$(p, pos + 1)
    End If
End Function

' 要素コレクションから行数相当を返す
Function GetElemRow(elems As IWebElements) As Long
Dim elem As IWebElement
Dim cnt As Long
cnt = 0
For Each elem In elems
    cnt = cnt + 1
Next
GetElemRow = cnt / 8
End Function

' 要素コレクション件数を返す
Function GetElemCount(elems As IWebElements) As Long
    Dim elem As IWebElement
    Dim cnt As Long
    cnt = 0
    For Each elem In elems
        cnt = cnt + 1
    Next
    GetElemCount = cnt
End Function

' 末尾 1 文字を除いた文字列を返す
Function RemoveRightMostChar(ByVal s As String) As String
    If Len(s) > 0 Then
        RemoveRightMostChar = left(s, Len(s) - 1)
    Else
        RemoveRightMostChar = s
    End If
End Function

'========================================================
' table 要素を二次元配列へ変換する
' 前提:
' - 単純な table のみ対象
' - rowspan / colspan / ネスト table は未対応
' - 各セルの文字列は GetTextContent で取得する
' - 各 tr には少なくとも1つの th または td がある前提
'========================================================
Public Function TableElementToArray2D(tbl As IWebElement) As Variant
    Dim rows As IWebElements
    Set rows = tbl.FindElementsByTag("tr")
    
    Dim rowElem As IWebElement
    Dim cellElem As IWebElement
    Dim cells As IWebElements
    
    Dim rowCount As Long
    Dim maxColCount As Long
    Dim colCount As Long
    
    ' 1回目の走査:
    ' 行数と最大列数を数える
    For Each rowElem In rows
        rowCount = rowCount + 1
        colCount = 0
        
        ' th / td を DOM順でまとめて取得する
        Set cells = rowElem.FindElementsByCss("th,td")
        
        For Each cellElem In cells
            colCount = colCount + 1
        Next
        
        If colCount > maxColCount Then
            maxColCount = colCount
        End If
    Next
    
    ' 行がない場合は Empty を返す
    If rowCount = 0 Or maxColCount = 0 Then
        TableElementToArray2D = Empty
        Exit Function
    End If
    
    Dim arr() As Variant
    ReDim arr(1 To rowCount, 1 To maxColCount)
    
    Dim r As Long
    Dim c As Long
    
    ' 2回目の走査:
    ' 配列へ格納する
    r = 0
    For Each rowElem In rows
        r = r + 1
        c = 0
        
        Set cells = rowElem.FindElementsByCss("th,td")
        
        For Each cellElem In cells
            c = c + 1
            arr(r, c) = cellElem.GetTextContent
        Next
        
        ' 足りない列は空文字で埋める
        Do While c < maxColCount
            c = c + 1
            arr(r, c) = ""
        Loop
    Next
    
    TableElementToArray2D = arr
End Function
'========================================================
' 要素コレクションの textContent を 1 次元配列へ変換する
' 前提:
' - elems が 0 件の場合は Empty を返す
'========================================================
Public Function WebElementsToTextArray1D(elems As IWebElements) As Variant
    Dim elem As IWebElement
    Dim cnt As Long
    
    cnt = 0
    For Each elem In elems
        cnt = cnt + 1
    Next
    
    If cnt = 0 Then
        WebElementsToTextArray1D = Empty
        Exit Function
    End If
    
    Dim arr() As Variant
    ReDim arr(1 To cnt)
    
    Dim i As Long
    i = 0
    For Each elem In elems
        i = i + 1
        arr(i) = elem.GetTextContent
    Next
    
    WebElementsToTextArray1D = arr
End Function

'========================================================
' 要素コレクションの指定属性値を 1 次元配列へ変換する
' 前提:
' - elems が 0 件の場合は Empty を返す
' - 属性が存在しない場合は空文字が入る
'========================================================
Public Function WebElementsToAttributeArray1D(elems As IWebElements, attrName As String) As Variant
    Dim elem As IWebElement
    Dim cnt As Long
    
    cnt = 0
    For Each elem In elems
        cnt = cnt + 1
    Next
    
    If cnt = 0 Then
        WebElementsToAttributeArray1D = Empty
        Exit Function
    End If
    
    Dim arr() As Variant
    ReDim arr(1 To cnt)
    
    Dim i As Long
    i = 0
    For Each elem In elems
        i = i + 1
        arr(i) = elem.GetAttribute(attrName)
    Next
    
    WebElementsToAttributeArray1D = arr
End Function

'========================================================
' 要素コレクションの textContent を区切り文字で連結して返す
' 前提:
' - elems が 0 件の場合は空文字を返す
'========================================================
Public Function JoinWebElementsText(elems As IWebElements, Optional delimiter As String = vbLf) As String
    Dim elem As IWebElement
    Dim result As String
    Dim isFirst As Boolean
    
    isFirst = True
    
    For Each elem In elems
        If isFirst Then
            result = elem.GetTextContent
            isFirst = False
        Else
            result = result & delimiter & elem.GetTextContent
        End If
    Next
    
    JoinWebElementsText = result
End Function

'========================================================
' table 要素を TSV 文字列へ変換する
' 前提:
' - TableElementToArray2D の結果をそのまま TSV 化する
' - 行区切りは vbCrLf
' - 列区切りは delimiter
'========================================================
Public Function TableElementToTSV(tbl As IWebElement, Optional delimiter As String = vbTab) As String
    Dim arr As Variant
    arr = TableElementToArray2D(tbl)
    
    If IsEmpty(arr) Then
        TableElementToTSV = ""
        Exit Function
    End If
    
    Dim r As Long
    Dim c As Long
    Dim lineText As String
    Dim result As String
    
    For r = 1 To UBound(arr, 1)
        lineText = ""
        
        For c = 1 To UBound(arr, 2)
            If c = 1 Then
                lineText = CStr(arr(r, c))
            Else
                lineText = lineText & delimiter & CStr(arr(r, c))
            End If
        Next
        
        If r = 1 Then
            result = lineText
        Else
            result = result & vbCrLf & lineText
        End If
    Next
    
    TableElementToTSV = result
End Function

'========================================================
' 2 次元配列を指定 Range に一括出力する
' 前提:
' - arr は 1 始まり 2 次元配列を想定
' - Empty の場合は何もしない
'========================================================
Public Sub Array2DToRange(arr As Variant, topLeft As Range)
    If IsEmpty(arr) Then Exit Sub
    
    Dim rowCount As Long
    Dim colCount As Long
    
    rowCount = UBound(arr, 1) - LBound(arr, 1) + 1
    colCount = UBound(arr, 2) - LBound(arr, 2) + 1
    
    topLeft.Resize(rowCount, colCount).value = arr
End Sub


'========================================================
' 既存ブラウザ attach + スクリーンショット貼り付け
' 何ができるか:
' - 今開いている Edge / Chrome に attach する
' - スクリーンショットを撮影する
' - 指定シートの最下部に、指定行数ぶん空けて貼り付ける
'
' 引数:
' - browserType : BrowserType_Edge / BrowserType_Chrome
' - sheetName   : 貼り付け先シート名
' - blankRows   : 最終位置から空ける行数（省略時 1）
'
' このサンプルの確認ポイント:
' - 既存ブラウザへ attach できること
' - 既存セルの下、または既存画像の下に追記されること
' - 2回目以降も同じ場所ではなく下へ積み上がること
'========================================================
Public Sub AttachBrowserAndPasteScreenshot( _
    ByVal browserType_ As BrowserType, _
    ByVal sheetName As String, _
    Optional ByVal blankRows As Long = 1)

    Dim drv As IWebDriver
    Dim ws As Worksheet
    Dim targetLeft As Single
    Dim targetTop As Single

    If blankRows < 0 Then
        Err.Raise vbObjectError + 2501, "SubModule.AttachBrowserAndPasteScreenshot", _
                  "blankRows には 0 以上を指定してください。"
    End If

    Set ws = ThisWorkbook.Worksheets(sheetName)

    Select Case browserType_
        Case BrowserType.BrowserType_Edge
            Set drv = NewAttachedEdgeDriver()
        Case BrowserType.BrowserType_Chrome
            Set drv = NewAttachedChromeDriver()
        Case Else
            Err.Raise vbObjectError + 2502, "SubModule.AttachBrowserAndPasteScreenshot", _
                      "browserType が不正です。"
    End Select

    targetLeft = ws.cells(1, 1).left
    targetTop = GetNextPasteTop_(ws, blankRows)

    Call drv.ScreenShotPasteToSheet( _
        left:=targetLeft, _
        top:=targetTop, _
        sheetName:=ws.name)

    Set drv = Nothing
End Sub
Private Function GetNextPasteTop_(ByVal ws As Worksheet, ByVal blankRows As Long) As Single
    Dim lastCell As Range
    Dim lastDataRow As Long
    Dim lastShapeBottomRow As Long
    Dim baseRow As Long
    Dim targetRow As Long
    Dim shp As Shape

    lastDataRow = 0
    lastShapeBottomRow = 0

    ' 将来のデータ増加時の計算量劣化を防ぐため、
    ' 全セル走査ではなく Find で最終使用セルを一度だけ取得する
    Set lastCell = ws.cells.Find(What:="*", _
                                 After:=ws.cells(1, 1), _
                                 LookIn:=xlFormulas, _
                                 LookAt:=xlPart, _
                                 SearchOrder:=xlByRows, _
                                 SearchDirection:=xlPrevious, _
                                 MatchCase:=False)

    If Not lastCell Is Nothing Then
        lastDataRow = lastCell.Row
    End If

    ' 既存画像の最下行を取得する
    For Each shp In ws.Shapes
        If shp.BottomRightCell.Row > lastShapeBottomRow Then
            lastShapeBottomRow = shp.BottomRightCell.Row
        End If
    Next shp

    If lastDataRow > lastShapeBottomRow Then
        baseRow = lastDataRow
    Else
        baseRow = lastShapeBottomRow
    End If

    If baseRow = 0 Then
        targetRow = 1
    ElseIf baseRow = 1 Then
        ' 1行目しか使っていない場合は空行を入れない
        targetRow = 2
    Else
        targetRow = baseRow + blankRows + 1
    End If

    GetNextPasteTop_ = ws.cells(targetRow, 1).top
    Debug.Print lastDataRow, lastShapeBottomRow, baseRow, targetRow
End Function


Public Sub ExportPicturesInSheetByOrder( _
    ByVal targetSheetName As String, _
    ByVal outputFolder As String, _
    Optional ByVal filePrefix As String = "img_")

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(targetSheetName)

    EnsureFolderExists outputFolder

    Dim items() As PictureItem
    Dim picCount As Long
    picCount = CollectPictureShapes(ws, items)

    If picCount = 0 Then
        Debug.Print "画像がありません。"
        Exit Sub
    End If

    ' 将来の画像数増加時に、毎回最小位置を探す O(n^2) を避けるため、
    ' 収集後に一度だけクイックソートで並べ替える
    QuickSortPictureItems items, LBound(items), UBound(items)

    Dim i As Long
    Dim filePath As String

    For i = LBound(items) To UBound(items)
        filePath = outputFolder & "\" & filePrefix & Format$(i, "000") & ".png"
    
        ExportShapeAsPng ws.Shapes(items(i).ShapeName), filePath
        WaitForExportShapeAsPng 15
    
        Debug.Print filePath
    Next i
End Sub

Private Function CollectPictureShapes(ByVal ws As Worksheet, ByRef items() As PictureItem) As Long
    Dim shp As Shape
    Dim count As Long
    count = 0

    For Each shp In ws.Shapes
        If IsExportTargetShape(shp) Then
            count = count + 1
            ReDim Preserve items(1 To count)

            items(count).ShapeName = shp.name
            items(count).TopPos = shp.top
            items(count).LeftPos = shp.left
        End If
    Next shp

    CollectPictureShapes = count
End Function

Private Function IsExportTargetShape(ByVal shp As Shape) As Boolean
    Select Case shp.Type
        Case msoPicture, msoLinkedPicture
            IsExportTargetShape = True
        Case Else
            IsExportTargetShape = False
    End Select
End Function

Private Sub QuickSortPictureItems(ByRef arr() As PictureItem, ByVal first As Long, ByVal last As Long)
    Dim i As Long
    Dim j As Long
    Dim pivotTop As Double
    Dim pivotLeft As Double
    Dim tmp As PictureItem

    i = first
    j = last
    pivotTop = arr((first + last) \ 2).TopPos
    pivotLeft = arr((first + last) \ 2).LeftPos

    Do While i <= j
        Do While ComparePictureItem(arr(i).TopPos, arr(i).LeftPos, pivotTop, pivotLeft) < 0
            i = i + 1
        Loop

        Do While ComparePictureItem(arr(j).TopPos, arr(j).LeftPos, pivotTop, pivotLeft) > 0
            j = j - 1
        Loop

        If i <= j Then
            tmp = arr(i)
            arr(i) = arr(j)
            arr(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop

    If first < j Then QuickSortPictureItems arr, first, j
    If i < last Then QuickSortPictureItems arr, i, last
End Sub

Private Function ComparePictureItem( _
    ByVal top1 As Double, ByVal left1 As Double, _
    ByVal top2 As Double, ByVal left2 As Double) As Long

    If top1 < top2 Then
        ComparePictureItem = -1
    ElseIf top1 > top2 Then
        ComparePictureItem = 1
    Else
        If left1 < left2 Then
            ComparePictureItem = -1
        ElseIf left1 > left2 Then
            ComparePictureItem = 1
        Else
            ComparePictureItem = 0
        End If
    End If
End Function



Public Sub ExportShapeAsPng(ByVal shp As Shape, ByVal outputPath As String)
    Dim ws As Worksheet
    Dim chObj As ChartObject

    If g_ExportBusy_ Then
        Err.Raise vbObjectError + 2601, "ExportShapeAsPng", _
                  "前回の画像エクスポート処理がまだ完了していません。"
    End If

    Set ws = shp.Parent

    g_ExportWsName_ = ws.name
    g_ExportTempChartName_ = "TempChartForExport_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CLng(Timer * 1000)
    g_ExportOutputPath_ = outputPath
    g_ExportBusy_ = True
    g_ExportDone_ = False
    g_ExportLastError_ = ""

    shp.CopyPicture Appearance:=xlScreen, Format:=xlPicture

    Set chObj = ws.ChartObjects.Add(left:=0, top:=0, width:=shp.width, height:=shp.height)
    chObj.name = g_ExportTempChartName_

    Application.OnTime Now + TimeValue("00:00:01"), "'" & ThisWorkbook.name & "'!PasteAndExportTempChartPng"
End Sub

Public Sub PasteAndExportTempChartPng()
    Dim ws As Worksheet
    Dim chObj As ChartObject

    On Error GoTo EH

    Set ws = ThisWorkbook.Worksheets(g_ExportWsName_)
    Set chObj = ws.ChartObjects(g_ExportTempChartName_)

    With chObj.Chart
        .Paste
        .Export fileName:=g_ExportOutputPath_, FilterName:="PNG"
    End With

    chObj.Delete
    Application.CutCopyMode = False

    g_ExportBusy_ = False
    g_ExportDone_ = True
    Exit Sub

EH:
    On Error Resume Next
    If Not chObj Is Nothing Then chObj.Delete
    Application.CutCopyMode = False

    g_ExportLastError_ = Err.Description
    g_ExportBusy_ = False
    g_ExportDone_ = True
End Sub

Public Sub WaitForExportShapeAsPng(Optional ByVal timeoutSec As Long = 15)
    Dim startT As Single

    startT = Timer

    Do While g_ExportDone_ = False
        DoEvents
        Sleep 50

        If Timer - startT > timeoutSec Then
            g_ExportBusy_ = False
            Err.Raise vbObjectError + 2602, "WaitForExportShapeAsPng", _
                      "画像エクスポートがタイムアウトしました。"
        End If
    Loop

    If Len(g_ExportLastError_) > 0 Then
        Err.Raise vbObjectError + 2603, "WaitForExportShapeAsPng", g_ExportLastError_
    End If
End Sub

