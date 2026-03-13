VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a4_ExecuteCDP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)

Private events_ As a5_CDPEventHandler
Private MsgGenerate_ As a4x1_CDPMessageGenerator
Private MsgHandler_ As a4x2_CDPMessageHandler
Private json_ As a2_JSONHandler

' 依存オブジェクトを初期化する
Public Sub Init(MsgGenerate As a4x1_CDPMessageGenerator, _
                Handler As a4x2_CDPMessageHandler, _
                json As a2_JSONHandler, _
                Optional events As a5_CDPEventHandler)
    Set MsgGenerate_ = MsgGenerate
    Set MsgHandler_ = Handler
    Set json_ = json
    Set events_ = events
End Sub

' Network ドメインを有効化する
Public Function NetworkEnable() As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.NetworkEnable()
    NetworkEnable = MsgHandler_.SendAndReceive(sendMsg)
End Function

' Network ドメインを無効化する
Public Function NetworkDisable() As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.NetworkDisable()
    NetworkDisable = MsgHandler_.SendAndReceive(sendMsg)
End Function

' requestId に対応するレスポンス本文を取得する
Public Function GetResponseBody(requestId As String) As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.NetworkGetResponseBody(requestId)
    GetResponseBody = MsgHandler_.SendAndReceive(sendMsg)
End Function

' ライフサイクルイベント通知を設定する
Public Function SetLifecycleEventsEnabled(enabled As Boolean) As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.SetLifecycleEventsEnabled(enabled)
    SetLifecycleEventsEnabled = MsgHandler_.SendAndReceive(sendMsg)
End Function
' Browser系************************************************
Public Function SessionClose() As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.SessionClose
    SessionClose = MsgHandler_.SendAndReceive(sendMsg)
End Function

' ************************************************

' DOM(ElementId)系************************************************
' CDPメソッドがないため JavaScript を実行する
Public Function GetDocumentObjectID() As String
    Dim resMsg As String
    resMsg = ExecuteJavaScript("document;")
    GetDocumentObjectID = json_.GetValue(resMsg, "result", "result", "objectId")
End Function

' document の nodeId を取得する
Public Function GetDocumentNodeID() As Long
    Dim sendMsg As String, resMsg As String
    sendMsg = MsgGenerate_.GetDocumentID
    resMsg = MsgHandler_.SendAndReceive(sendMsg)
    GetDocumentNodeID = json_.GetValue(resMsg, "result", "root", "nodeId")
End Function

' objectId を nodeId へ変換する
Private Function ChangeObjectIdToNodeId(objectId As String) As String
    Dim sendMsg As String, resMsg As String
    sendMsg = MsgGenerate_.RequestNode(objectId)
    resMsg = MsgHandler_.SendAndReceive(sendMsg)
    ChangeObjectIdToNodeId = json_.GetValue(resMsg, "result", "nodeId")
End Function

' nodeId を objectId へ変換する
Private Function ChangeNodeIdToObjectId(nodeId As Long) As String
    Dim sendMsg As String, resMsg As String
    sendMsg = MsgGenerate_.ResolveNode(nodeId)
    resMsg = MsgHandler_.SendAndReceive(sendMsg)
    ChangeNodeIdToObjectId = json_.GetValue(resMsg, "result", "objectId")
End Function
' ************************************************

' DOM(Element情報)系************************************************
Private Function GetAttributes(nodeId As Long) As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.GetAttributes(nodeId)
    GetAttributes = MsgHandler_.SendAndReceive(sendMsg)
End Function

' ノード情報を取得する
Public Function DescribeNode(Id As Variant) As Variant
    Dim sendMsg As String
    sendMsg = MsgGenerate_.DescribeNode(Id)
    DescribeNode = MsgHandler_.SendAndReceive(sendMsg)
End Function
' ************************************************

' Runtime************************************************
' executionContextCreated イベントによる実行コンテキスト生成通知を有効にする
' 有効化後は既存コンテキスト分のイベントも即時に通知される
Public Function RuntimeEnable() As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.RunTime_Enable()
    RuntimeEnable = MsgHandler_.SendAndReceive(sendMsg)
End Function
' JavaScriptコードを実行する
Public Function ExecuteJavaScript(script As String) As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.JavaScritpEvaluate(script)
    ExecuteJavaScript = MsgHandler_.SendAndReceive(sendMsg)
End Function

' JavaScript関数を実行する
Public Function CallJavaScriptFunction(funcName As String, Id As Variant, args As Variant) As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.CallJavaScriptFunction(funcName, Id, args)
    CallJavaScriptFunction = MsgHandler_.SendAndReceive(sendMsg)
End Function

' JavaScript関数を複数引数で実行する
Public Function CallJavaScriptFunction_ArgMultiple(funcName As String, Id As Variant, ParamArray args() As Variant) As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.CallJavaScriptFunction_ArgMultiple(funcName, Id, args)
    CallJavaScriptFunction_ArgMultiple = MsgHandler_.SendAndReceive(sendMsg)
End Function

' ************************************************
' Input Domain*******************************************************************
' キーイベントをページへディスパッチする
Public Function DispatchKeyEvent(action As String, modifiers As Long, windowsVirtualKeyCode As Long) As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.DispatchKeyEvent(action, modifiers, windowsVirtualKeyCode)
    DispatchKeyEvent = MsgHandler_.SendAndReceive(sendMsg)
End Function

' マウスイベントを指定座標へディスパッチする
Public Function DispatchMouseEvent(eventType As String, x As Long, y As Long, _
                                                        button As String, clickCount As Long, _
                                                        Optional modifiers As Long = 0, _
                                                        Optional deltaX As Long = 0, _
                                                        Optional deltaY As Long = 0) As String

    Dim sendMsg As String
    sendMsg = MsgGenerate_.DispatchMouseEvent( _
                        eventType, x, y, button, clickCount, _
                        modifiers, deltaX, deltaY)
    DispatchMouseEvent = MsgHandler_.SendAndReceive(sendMsg)
End Function

' *******************************************************************

' DOM操作系************************************************
Public Function DOMEnable() As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.DOM_Enable()
    DOMEnable = MsgHandler_.SendAndReceive(sendMsg)
End Function

Public Function Focus(objectId As String) As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.Focus(objectId)
    Focus = MsgHandler_.SendAndReceive(sendMsg)
End Function

Public Function GetBoxModel(objectId As String) As Variant
    Dim sendMsg As String
    sendMsg = MsgGenerate_.GetBoxModel(objectId)
    GetBoxModel = MsgHandler_.SendAndReceive(sendMsg)
End Function

Public Function SetFileInputFiles(objectId As Variant, files As String) As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.SetFileInputFiles(objectId, files)
    SetFileInputFiles = MsgHandler_.SendAndReceive(sendMsg)
End Function

Public Sub FindElementByQuerySelector(Selector As String)
    Dim sendMsg As String, resMsg As String
    sendMsg = MsgGenerate_.querySelector(Selector)
    resMsg = MsgHandler_.SendAndReceive(sendMsg)

    Dim nodeId As Long
    nodeId = json_.GetValue(resMsg, "result", "nodeId")
    Dim nodeInfoJson As String
    nodeInfoJson = DescribeNode(nodeId)
End Sub

' ************************************************

' Input系************************************************
Public Function Click(objId As String) As String
    CallJavaScriptFunction "function(value){this.click();}", objId, ""
End Function
' ************************************************

' Target:Page系************************************************
Public Function OpenURL(targetURL As String) As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.URLOpen(targetURL)
    OpenURL = MsgHandler_.SendAndReceive(sendMsg)
End Function

' Page ドメインを有効化する
Public Function PageEnable() As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.PageEnable()
    PageEnable = MsgHandler_.SendAndReceive(sendMsg)
End Function

' フレームツリーを取得する
Public Function GetFrameTree() As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.GetFrameTree()
    GetFrameTree = MsgHandler_.SendAndReceive(sendMsg)
End Function

' フレーム用の IsolatedWorld を生成する
Public Function CreateIsolatedFrameWorld(frameId As String) As Long
    Dim sendMsg As String, resMsg As String
    sendMsg = MsgGenerate_.CreateIsolatedFrameWorld(frameId)
    resMsg = MsgHandler_.SendAndReceive(sendMsg)
    If InStr(resMsg, "error") > 0 Then
        CreateIsolatedFrameWorld = 0
    Else
        CreateIsolatedFrameWorld = json_.GetValue(resMsg, "result", "executionContextId")
        Sleep 200
    End If
End Function

' Target 検出を有効化する
Public Function SetDiscoverTargets() As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.SetDiscoverTargets()
    SetDiscoverTargets = MsgHandler_.SendAndReceive(sendMsg)
End Function

' targetId にアタッチして SessionId を反映する
Public Function AttachToTarget(targetId As String) As String
    Dim sendMsg As String, res As String, SessionId As String
    sendMsg = MsgGenerate_.AttachToTarget(targetId)
    res = MsgHandler_.SendAndReceive(sendMsg)
    SessionId = json_.GetValue(res, "result", "sessionId")
    MsgGenerate_.SessionId = SessionId
    AttachToTarget = res
End Function

' targetId をアクティブ化する
Public Function ActivateTarget(targetId As String) As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.ActivateTarget(targetId)
    ActivateTarget = MsgHandler_.SendAndReceive(sendMsg)
End Function

' targetId からデタッチする
Public Function DetachFromTarget(targetId As String) As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.DetachFromTarget(targetId)
    DetachFromTarget = MsgHandler_.SendAndReceive(sendMsg)
End Function

' targetId を閉じる
Public Function TargetClose(targetId As String) As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.TargetClose(targetId)
    TargetClose = MsgHandler_.SendAndReceive(sendMsg)
End Function

' Target 一覧を取得する
Public Function GetTargetInfo() As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.GetTargets()
    GetTargetInfo = MsgHandler_.SendAndReceive(sendMsg)
End Function

' JavaScript ダイアログを処理する
Public Function HandleAlertDialog() As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.HandleJavaScriptDialog()
    HandleAlertDialog = MsgHandler_.SendAndReceive(sendMsg)
End Function

' 自作Function系************************************************

' ページ表示完了まで待機する
Public Function WaitingForWebPageDisplay()
    Dim javaScriptCode As String, responseMsg As String, state As String
    Do
        Sleep 800
        javaScriptCode = "document.readyState;"
        responseMsg = ExecuteJavaScript(javaScriptCode)
        state = json_.GetValue(responseMsg, "result", "result", "value")
        DoEvents
    Loop Until (state = "complete")
End Function

' 指定条件で要素を探索し、ObjectId を返す
Public Function GetObjectIdOfElementBy(way As String, basicInfo As a3_BasicInfos, Codes As Variant) As String

    Dim startTime As Date: startTime = Timer()
    Dim res As String, objectId As String
    Do
        If basicInfo.ExecutionContextId = 0 Then
            res = ExecuteJavaScript((Codes(0)))
        Else
            res = CallJavascriptFunctionWrapper((Codes(1)), basicInfo, way)
        End If

        objectId = json_.GetValue(res, "result", "result", "objectId")

        If objectId <> "" Then Exit Do

        If CInt((Timer() - startTime)) > 4 Then Exit Do

        DoEvents:  Sleep 500
    Loop

    GetObjectIdOfElementBy = objectId
End Function

' 実行コンテキスト消失時の再取得を含めて関数呼び出しを行う
Public Function CallJavascriptFunctionWrapper(funcCode As String, basicInfo As a3_BasicInfos, args As Variant) As String
    Dim res As String
    res = CallJavaScriptFunction(funcCode, basicInfo.ExecutionContextId, args)
    If InStr(res, "Cannot find context with specified id") > 0 Then
        basicInfo.UpdateExecutionContextId Me
        res = CallJavaScriptFunction(funcCode, basicInfo.ExecutionContextId, args)
    End If
    CallJavascriptFunctionWrapper = res
End Function

' 要素情報から IWebElement を生成して返す
' ObjectId・NodeId・属性辞書・frameId を必要に応じて設定する
Public Function MakeElementObject( _
        objectId As String, _
        findInfo As Collection, _
        basicInfo As a3_BasicInfos, _
        events As a5_CDPEventHandler, _
        Helper As a6_ExecuteHelperFunction) As IWebElement

    Dim info As String
    info = DescribeNode(objectId)
    If InStr(info, "attributes") > 0 Then
        Dim attributes As Object
        Set attributes = MakeAttributesDic(json_.GetValue(info, "result", "node", "attributes"))
    End If
    If InStr(info, "frameId") > 0 Then
        Dim frameId As String
        frameId = json_.GetValue(info, "result", "node", "frameId")
    End If

    Dim element As b0x0_WebElement: Set element = New b0x0_WebElement
    element.Init json_, basicInfo, Me, events, Helper, _
                       objectId, findInfo, attributes, frameId
    Set MakeElementObject = element
End Function
' 属性配列を辞書へ変換して返す
Private Function MakeAttributesDic(attributes As String) As Object
    Dim ary As Variant: ary = Split(attributes, ",")
    Dim attributesDic As Object: Set attributesDic = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(ary) To UBound(ary) - 1 Step 2
        If attributesDic.Exists(ary(i)) = False Then
            attributesDic.Add ary(i), ary(i + 1)
        End If
    Next
    Set MakeAttributesDic = attributesDic
End Function

' 指定条件に一致する要素数を取得する
Public Function GetElementsCountBy(way As String, basicInfo As a3_BasicInfos, Codes As Variant) As Long
    Dim responseMsg As String
    If basicInfo.ExecutionContextId = 0 Then
        responseMsg = ExecuteJavaScript((Codes(0)))
    Else
        responseMsg = CallJavascriptFunctionWrapper((Codes(1)), basicInfo, way)
    End If
    GetElementsCountBy = json_.GetValue(responseMsg, "result", "result", "value")
End Function

' 要素存在判定の共通処理
Public Function IsExistsElementBy(way As String, basicInfo As a3_BasicInfos, Codes As Variant) As Boolean

    Dim responseMsg As String
    If basicInfo.ExecutionContextId = 0 Then
        responseMsg = ExecuteJavaScript((Codes(0)))
    Else
        responseMsg = CallJavascriptFunctionWrapper((Codes(1)), basicInfo, way)
    End If
    Dim objectId As String
    objectId = json_.GetValue(responseMsg, "result", "result", "objectId")

    If objectId = "" Then
        IsExistsElementBy = False
    Else
        IsExistsElementBy = True
    End If
End Function

' 要素出現待機の共通処理
Public Function WaitForElementToAppearBy( _
    way As String, maxWaitTime_sec As Long, basicInfo As a3_BasicInfos, Codes As Variant) As Boolean

    Dim startTime As Date: startTime = Timer()
    Do
        If IsExistsElementBy(way, basicInfo, Codes) Then
            WaitForElementToAppearBy = True
            Exit Function
        ElseIf CInt((Timer() - startTime)) > maxWaitTime_sec Then
            WaitForElementToAppearBy = False
            Exit Function
        End If
        DoEvents: Sleep 500
    Loop
End Function

' Download
Public Function SetDownloadBehavior(downloadPath As String, Optional allowAndName As Boolean = False, Optional eventsEnabled As Boolean = True) As String
     Dim sendMsg As String
     sendMsg = MsgGenerate_.BrowserSetDownloadBehavior(downloadPath, allowAndName, eventsEnabled)

     SetDownloadBehavior = MsgHandler_.SendAndReceive(sendMsg)
End Function

' Screenshot / PDF / Hooks
Public Function CaptureScreenshot(formatName As String, Optional quality As Long = -1, Optional clipJson As String = "", Optional fromSurface As Boolean = True, Optional captureBeyondViewport As Boolean = False) As String
    Dim sendMsg As String, resMsg As String
    sendMsg = MsgGenerate_.PageCaptureScreenshot(formatName, quality, clipJson, fromSurface, captureBeyondViewport)
    resMsg = MsgHandler_.SendAndReceive(sendMsg)
    CaptureScreenshot = json_.GetValue(resMsg, "result", "data")
End Function

' PDF を生成して Base64 データを返す
Public Function PrintToPDF(Optional printBackground As Boolean = True, Optional landscape As Boolean = False, Optional paperWidth As Double = 8.27, Optional paperHeight As Double = 11.69, Optional marginTop As Double = 0, Optional marginBottom As Double = 0, Optional marginLeft As Double = 0, Optional marginRight As Double = 0, Optional scale_ As Double = 1, Optional pageRanges As String = "") As String
    Dim sendMsg As String, resMsg As String
    sendMsg = MsgGenerate_.PagePrintToPDF(printBackground, landscape, paperWidth, paperHeight, marginTop, marginBottom, marginLeft, marginRight, scale_, pageRanges)
    resMsg = MsgHandler_.SendAndReceive(sendMsg)
    PrintToPDF = json_.GetValue(resMsg, "result", "data")
End Function

' 新規ドキュメント評価用スクリプトを登録する
Public Function AddScriptToEvaluateOnNewDocument(source As String, Optional worldName As String = "", Optional includeCommandLineAPI As Boolean = False, Optional runImmediately As Boolean = False) As String
    Dim sendMsg As String, resMsg As String
    sendMsg = MsgGenerate_.PageAddScriptToEvaluateOnNewDocument(source, worldName, includeCommandLineAPI, runImmediately)
    resMsg = MsgHandler_.SendAndReceive(sendMsg)
    AddScriptToEvaluateOnNewDocument = json_.GetValue(resMsg, "result", "identifier")
End Function

' 登録済みスクリプトを削除する
Public Function RemoveScriptToEvaluateOnNewDocument(identifier As String) As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.PageRemoveScriptToEvaluateOnNewDocument(identifier)
    RemoveScriptToEvaluateOnNewDocument = MsgHandler_.SendAndReceive(sendMsg)
End Function

' ファイル選択ダイアログの割り込みを設定する
Public Function SetInterceptFileChooserDialog(enabled As Boolean) As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.PageSetInterceptFileChooserDialog(enabled)
    SetInterceptFileChooserDialog = MsgHandler_.SendAndReceive(sendMsg)
End Function

' backendNodeId 指定でファイル入力値を設定する
Public Function SetFileInputFilesByBackendNodeId(backendNodeId As Long, filePath As String) As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.SetFileInputFilesByBackendNodeId(backendNodeId, filePath)
    SetFileInputFilesByBackendNodeId = MsgHandler_.SendAndReceive(sendMsg)
End Function

' 通信が一定時間アイドルになるまで待機する
Public Function WaitForNetworkIdle(timeoutSec As Long, Optional idleMillis As Long = 500) As Boolean
    If events_ Is Nothing Then Err.Raise 10001, , "events_ is Nothing"
    Dim startT As Double: startT = Timer()
    Do
        DoEvents
        Sleep 50
        If events_.NetworkInFlight = 0 Then
            If (Timer() - events_.NetworkLastActivity) * 1000# >= idleMillis Then
                WaitForNetworkIdle = True
                Exit Function
            End If
        End If
    Loop While (Timer() - startT) <= timeoutSec
    WaitForNetworkIdle = False
End Function

' DOM 更新イベントが発生するまで待機する
Public Function WaitForDomMutation(timeoutSec As Long) As Boolean
    If events_ Is Nothing Then Err.Raise 10001, , "events_ is Nothing"
    Dim startT As Double: startT = Timer()
    events_.DomUpdated = False
    Do
        DoEvents
        Sleep 50
        If events_.DomUpdated Then
            WaitForDomMutation = True
            Exit Function
        End If
    Loop While (Timer() - startT) <= timeoutSec
    WaitForDomMutation = False
End Function

' Storage domain
Public Function StorageGetCookies() As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.StorageGetCookies()
    StorageGetCookies = MsgHandler_.SendAndReceive(sendMsg)
End Function

' Storage の Cookie を設定する
Public Function StorageSetCookies(cookieParamsJsonArray As String) As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.StorageSetCookies(cookieParamsJsonArray)
    StorageSetCookies = MsgHandler_.SendAndReceive(sendMsg)
End Function

' Storage の Cookie を削除する
Public Function StorageClearCookies() As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.StorageClearCookies()
    StorageClearCookies = MsgHandler_.SendAndReceive(sendMsg)
End Function

' StorageKey を取得する
Public Function StorageGetStorageKey(Optional frameId As String = "") As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.StorageGetStorageKey(frameId)
    StorageGetStorageKey = MsgHandler_.SendAndReceive(sendMsg)
End Function

' DOMStorage domain
Public Function DOMStorageEnable() As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.DOMStorageEnable()
    DOMStorageEnable = MsgHandler_.SendAndReceive(sendMsg)
End Function

' DOMStorage ドメインを無効化する
Public Function DOMStorageDisable() As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.DOMStorageDisable()
    DOMStorageDisable = MsgHandler_.SendAndReceive(sendMsg)
End Function

' DOMStorage の項目一覧を取得する
Public Function DOMStorageGetItems(securityOrigin As String, storageKey As String, isLocal As Boolean) As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.DOMStorageGetItems(securityOrigin, storageKey, isLocal)
    DOMStorageGetItems = MsgHandler_.SendAndReceive(sendMsg)
End Function

' DOMStorage の項目を設定する
Public Function DOMStorageSetItem(securityOrigin As String, storageKey As String, isLocal As Boolean, key As String, value As String) As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.DOMStorageSetItem(securityOrigin, storageKey, isLocal, key, value)
    DOMStorageSetItem = MsgHandler_.SendAndReceive(sendMsg)
End Function

' DOMStorage の項目を削除する
Public Function DOMStorageRemoveItem(securityOrigin As String, storageKey As String, isLocal As Boolean, key As String) As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.DOMStorageRemoveItem(securityOrigin, storageKey, isLocal, key)
    DOMStorageRemoveItem = MsgHandler_.SendAndReceive(sendMsg)
End Function

' DOMStorage をクリアする
Public Function DOMStorageClear(securityOrigin As String, storageKey As String, isLocal As Boolean) As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.DOMStorageClear(securityOrigin, storageKey, isLocal)
    DOMStorageClear = MsgHandler_.SendAndReceive(sendMsg)
End Function

' アラートダイアログを受諾する
Public Function AcceptAlertDialog(Optional promptText As String = "") As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.HandleJavaScriptDialogEx(True, promptText)
    AcceptAlertDialog = MsgHandler_.SendAndReceive(sendMsg)
End Function

' アラートダイアログを閉じる
Public Function DismissAlertDialog(Optional promptText As String = "") As String
    Dim sendMsg As String
    sendMsg = MsgGenerate_.HandleJavaScriptDialogEx(False, promptText)
    DismissAlertDialog = MsgHandler_.SendAndReceive(sendMsg)
End Function

'第1引数のCDPメソッドのコードの実行
Public Function ExecuteDirect(CDPMethodJson As String) As String
        Dim sendMsg  As String
        ExecuteDirect = MsgHandler_.SendAndReceive(CDPMethodJson)
End Function

