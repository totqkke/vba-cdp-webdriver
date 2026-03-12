VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a4x1_CDPMessageGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private id_ As Long
Private sessionId_ As String

' 現在のメッセージ ID を返す
Public Property Get Id() As Long
    Id = id_
End Property

' SessionId を保持する
Public Property Let SessionId(value As String)
    sessionId_ = value
End Property

' TargetDomain*******************************************************************
' Page ドメインを有効化する
Public Function PageEnable() As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Page.enable""}"
    PageEnable = message
End Function

' Target 検出を有効化するメッセージを生成する
Public Function SetDiscoverTargets() As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Target.setDiscoverTargets""," & _
                      """params"": {""discover"": true}}"
    SetDiscoverTargets = message
End Function

' targetId へアタッチするメッセージを生成する
Public Function AttachToTarget(targetId As String) As String
    id_ = id_ + 1

    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Target.attachToTarget"",""params"": {" & _
              """targetId"": " & JsonString(targetId) & "}}"

    AttachToTarget = message
End Function

' targetId をアクティブ化するメッセージを生成する
Public Function ActivateTarget(targetId As String) As String
    id_ = id_ + 1

    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Target.activateTarget"",""params"": {" & _
              """targetId"": " & JsonString(targetId) & "}}"

    ActivateTarget = message
End Function

' targetId からデタッチするメッセージを生成する
Public Function DetachFromTarget(targetId As String) As String
    id_ = id_ + 1

    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Target.detachFromTarget"",""params"": {" & _
              """targetId"": " & JsonString(targetId) & "}}"

    DetachFromTarget = message
End Function

' targetId を閉じるメッセージを生成する
Public Function TargetClose(targetId As String) As String
    id_ = id_ + 1

    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Target.closeTarget"",""params"": {" & _
              """targetId"": " & JsonString(targetId) & "}}"

    TargetClose = message
End Function

' 利用可能なターゲット一覧を取得する
Public Function GetTargets() As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Target.getTargets""}"
    GetTargets = message
End Function
' *******************************************************************

' BrowserDomain*******************************************************************

' Browser.close のメッセージを生成する
Public Function SessionClose() As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Browser.close""}"
    SessionClose = message
End Function
' *******************************************************************

' PageDomain*******************************************************************

' URL を開くメッセージを生成する
Public Function URLOpen(URL As String) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Page.navigate""," & _
              """params"": {""url"": """ & EscapeJsonString(URL) & """}}"
    URLOpen = message
End Function

' JavaScript ダイアログを受諾するメッセージを生成する
Public Function HandleJavaScriptDialog() As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Page.handleJavaScriptDialog""," & _
                     """params"": {""accept"": true}}"
    HandleJavaScriptDialog = message
End Function

' フレームツリー取得メッセージを生成する
Public Function GetFrameTree() As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Page.getFrameTree""}"
    GetFrameTree = message
End Function

' IsolatedWorld 作成メッセージを生成する
Public Function CreateIsolatedFrameWorld(frameId As String) As String
    id_ = id_ + 1

    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Page.createIsolatedWorld"",""params"": {" & _
              """frameId"": " & JsonString(frameId) & ",""grantUniveralAccess"": true}}"

    CreateIsolatedFrameWorld = message
End Function

' *******************************************************************
' Input Domain*******************************************************************
' キーイベントをページへディスパッチする
Public Function DispatchKeyEvent(action As String, modifiers As Long, windowsVirtualKeyCode As Long) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Input.dispatchKeyEvent""," & _
                      """params"": {""type"":  """ & action & """,""modifiers"": " & modifiers & ",""windowsVirtualKeyCode"": " & windowsVirtualKeyCode & "}}"
    DispatchKeyEvent = message

End Function

' マウスイベントを指定座標へディスパッチする
Public Function DispatchMouseEvent(eventType As String, x As Long, y As Long, _
                                                        button As String, clickCount As Long, _
                                                        Optional modifiers As Long = 0, _
                                                        Optional deltaX As Long = 0, Optional deltaY As Long = 0) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Input.dispatchMouseEvent""," & _
                      """params"": {""type"":  """ & eventType & """,""x"": " & x & ",""y"": " & y & "," & _
                                            """button"": """ & button & """,""clickCount"": " & clickCount
                                            ' button は none / left / middle / right / back / forward を指定する

    ' modifiers は Alt=1 / Ctrl=2 / Meta(Command)=4 / Shift=8 を表す
    If modifiers <> 0 Then
        message = message & ",""modifiers"": " & modifiers
    End If

    If deltaX <> 0 Or deltaY <> 0 Then
        message = message & ",""deltaX"": " & deltaX & ",""deltaY"": " & deltaY
    End If

    message = message & "}}"
    DispatchMouseEvent = message
End Function
' *******************************************************************
' DOM Domain*******************************************************************
' DOM ドメインを有効化する
Public Function DOM_Enable()
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.enable""}"
    DOM_Enable = message
End Function

' 指定要素へフォーカスを設定する
Public Function Focus(objectId As String) As String
    id_ = id_ + 1

    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.focus"",""params"": {" & _
              """objectId"": " & JsonString(objectId) & "}}"

    Focus = message
End Function

' Document の nodeId を取得する
Public Function GetDocumentID() As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.getDocument""," & _
                     """params"": {""depth"": " & 1 & "}}"
    GetDocumentID = message
End Function

' CSS セレクタで nodeId を取得する
Public Function querySelector(Selector As String, Optional nodeId As Long = 1) As String
    id_ = id_ + 1

    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.querySelector"",""params"": {" & _
              """nodeId"": " & nodeId & ",""selector"": " & JsonString(Selector) & "}}"

    querySelector = message
End Function

' objectId から nodeId を取得する
Public Function RequestNode(objectId As String) As String
    id_ = id_ + 1

    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.requestNode"",""params"": {" & _
              """objectId"": " & JsonString(objectId) & "}}"

    RequestNode = message
End Function

' nodeId から objectId を取得する
Public Function ResolveNode(nodeId As Long)
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.resolveNode""," & _
                      """params"": {""nodeId"":  " & nodeId & "}}"
    ResolveNode = message
End Function
' 指定 nodeId の属性情報を取得する
Public Function GetAttributes(nodeId As Long)
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.getAttributes""," & _
                      """params"":{""nodeId"":  " & nodeId & "}}"
    GetAttributes = message
End Function

' 指定 ID のノード情報を取得する
Public Function DescribeNode(Id As Variant) As Variant
    id_ = id_ + 1

    Dim message As String
    If IsNumeric(Id) Then
        message = "{""id"": " & id_ & ",""method"": ""DOM.describeNode"",""params"": {" & _
                  """nodeId"": " & Id & ",""depth"": 1,""pierce"": false}}"
    Else
        message = "{""id"": " & id_ & ",""method"": ""DOM.describeNode"",""params"": {" & _
                  """objectId"": " & JsonString(CStr(Id)) & ",""depth"": 1,""pierce"": false}}"
    End If

    DescribeNode = message
End Function

Public Function RequestChildNodes() As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.RequestChildNodes""," & _
                      """params"": {""nodeId"":  " & id_ & ",""depth"":  -1}}"
    RequestChildNodes = message
End Function

' 指定要素へフォーカスを設定する
Public Function GetBoxModel(objectId As String) As String
    id_ = id_ + 1

    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.getBoxModel"",""params"": {" & _
              """objectId"": " & JsonString(objectId) & "}}"

    GetBoxModel = message
End Function

' file input へファイルパスを設定する
Public Function SetFileInputFiles(objectId As Variant, files As String) As String
    id_ = id_ + 1

    Dim filesJson As String
    filesJson = "[" & JsonString(files) & "]"

    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.setFileInputFiles"",""params"": {" & _
              """files"": " & filesJson & ",""objectId"": " & JsonString(CStr(objectId)) & "}}"

    SetFileInputFiles = message
End Function
' backendNodeId 指定で file input へファイルパスを設定する
Public Function SetFileInputFilesByBackendNodeId(backendNodeId As Long, filePath As String) As String
    id_ = id_ + 1

    Dim filesJson As String
    filesJson = "[" & JsonString(filePath) & "]"

    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.setFileInputFiles"",""params"": {" & _
              """files"": " & filesJson & ",""backendNodeId"": " & backendNodeId & "}}"

    SetFileInputFilesByBackendNodeId = message
End Function

' RunTime Domain*******************************************************************

' Runtime.enable のメッセージを生成する
Public Function RunTime_Enable()
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Runtime.enable""}"
    RunTime_Enable = message
End Function

' JavaScript 評価メッセージを生成する
Public Function JavaScritpEvaluate(script As String)
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Runtime.evaluate""," & _
              """params"": {""expression"": """ & EscapeJsonString(script) & """,""returnByValue"": false}}"
    JavaScritpEvaluate = message
End Function

' JavaScript 関数呼び出しメッセージを生成する
Public Function CallJavaScriptFunction(funcDeclear As String, objectId As Variant, argments As Variant)
    id_ = id_ + 1

    Dim JavaScriptObjectArgs As String

    If argments = "true" Or argments = "false" Then
        JavaScriptObjectArgs = "[{""value"": " & argments & "}]"
    Else
        JavaScriptObjectArgs = "[{""value"": " & JsonString(CStr(argments)) & "}]"
    End If

    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Runtime.callFunctionOn"",""params"": {" & _
              """functionDeclaration"": " & JsonString(funcDeclear) & ",""arguments"": " & JavaScriptObjectArgs

    If IsNumeric(objectId) Then
        message = message & ",""executionContextId"": " & objectId & "}}"
    Else
        message = message & ",""objectId"": " & JsonString(CStr(objectId)) & "}}"
    End If

    CallJavaScriptFunction = message
End Function

' 複数引数の JavaScript 関数呼び出しメッセージを生成する
Public Function CallJavaScriptFunction_ArgMultiple(funcDeclear As String, objectId As Variant, ParamArray args() As Variant)
    id_ = id_ + 1

    Dim arg_arr As Variant
    arg_arr = args(0)(0)

    Dim parts As String
    Dim i As Long

    For i = LBound(arg_arr) To UBound(arg_arr)
        If i Mod 2 = 0 Then
            parts = parts & JsonString(CStr(arg_arr(i))) & ": "
        Else
            If arg_arr(i) = "true" Or arg_arr(i) = "false" Or IsNumeric(arg_arr(i)) Then
                parts = parts & CStr(arg_arr(i))
            Else
                parts = parts & JsonString(CStr(arg_arr(i)))
            End If

            If i < UBound(arg_arr) Then
                parts = parts & ","
            End If
        End If
    Next

    Dim JavaScriptObjectArgs As String
    JavaScriptObjectArgs = "[{""value"":{" & parts & "}}]"

    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Runtime.callFunctionOn"",""params"": {" & _
              """functionDeclaration"": " & JsonString(funcDeclear) & ",""arguments"": " & JavaScriptObjectArgs

    If IsNumeric(objectId) Then
        message = message & ",""executionContextId"": " & objectId & "}}"
    Else
        message = message & ",""objectId"": " & JsonString(CStr(objectId)) & "}}"
    End If

    CallJavaScriptFunction_ArgMultiple = message
End Function
' *******************************************************************

' Browser download 制御
' Browser.setDownloadBehavior のメッセージを生成する
Public Function BrowserSetDownloadBehavior(downloadPath As String, Optional allowAndName As Boolean = False, Optional eventsEnabled As Boolean = True) As String
    id_ = id_ + 1

    Dim behavior As String
    If allowAndName Then
        behavior = "allowAndName"
    Else
        behavior = "allow"
    End If

    Dim pathEsc As String
    pathEsc = downloadPath
    pathEsc = Replace(pathEsc, "\", "\\")
    pathEsc = Replace(pathEsc, """", "\""")  ' ダブルクォート対策

    Dim msg As String
    msg = "{""id"": " & id_ & ",""method"": ""Browser.setDownloadBehavior"",""params"": {""behavior"": """ & behavior & """,""downloadPath"": """ & pathEsc & """,""eventsEnabled"": " & LCase$(CStr(eventsEnabled)) & "}}"

    BrowserSetDownloadBehavior = msg
End Function

' Page.captureScreenshot のメッセージを生成する
Public Function PageCaptureScreenshot(formatName As String, Optional quality As Long = -1, Optional clipJson As String = "", Optional fromSurface As Boolean = True, Optional captureBeyondViewport As Boolean = False) As String
    id_ = id_ + 1

    Dim params As String
    params = """format"": """ & formatName & """"
    If quality >= 0 Then params = params & ",""quality"": " & quality
    If clipJson <> "" Then params = params & ",""clip"": " & clipJson
    params = params & ",""fromSurface"": " & LCase$(CStr(fromSurface))
    params = params & ",""captureBeyondViewport"": " & LCase$(CStr(captureBeyondViewport))

    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Page.captureScreenshot"",""params"": {" & params & "}}"
    PageCaptureScreenshot = message
End Function

' ファイル選択ダイアログ割り込み設定メッセージを生成する
Public Function SetInterceptFileChooserDialog(enabled As Boolean) As String
    id_ = id_ + 1
    Dim enabledStr As String
    If enabled Then
        enabledStr = "true"
    Else
        enabledStr = "false"
    End If

    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Page.setInterceptFileChooserDialog"",""params"": {""enabled"": " & enabledStr & "}}"
    SetInterceptFileChooserDialog = message
End Function
' Page.printToPDF のメッセージを生成する
Public Function PagePrintToPDF(Optional printBackground As Boolean = True, Optional landscape As Boolean = False, Optional paperWidth As Double = 8.27, Optional paperHeight As Double = 11.69, Optional marginTop As Double = 0, Optional marginBottom As Double = 0, Optional marginLeft As Double = 0, Optional marginRight As Double = 0, Optional scale_ As Double = 1, Optional pageRanges As String = "") As String

    id_ = id_ + 1

    Dim params As String
    params = """printBackground"": " & LCase$(CStr(printBackground)) & ",""landscape"": " & LCase$(CStr(landscape))
    params = params & ",""paperWidth"": " & Replace(CStr(paperWidth), ",", ".")
    params = params & ",""paperHeight"": " & Replace(CStr(paperHeight), ",", ".")
    params = params & ",""marginTop"": " & Replace(CStr(marginTop), ",", ".")
    params = params & ",""marginBottom"": " & Replace(CStr(marginBottom), ",", ".")
    params = params & ",""marginLeft"": " & Replace(CStr(marginLeft), ",", ".")
    params = params & ",""marginRight"": " & Replace(CStr(marginRight), ",", ".")
    params = params & ",""scale"": " & Replace(CStr(scale_), ",", ".")
    If pageRanges <> "" Then params = params & ",""pageRanges"": """ & pageRanges & """"

    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Page.printToPDF"",""params"": {" & params & "}}"
    PagePrintToPDF = message
End Function

' Page.addScriptToEvaluateOnNewDocument のメッセージを生成する
Public Function PageAddScriptToEvaluateOnNewDocument(source As String, Optional worldName As String = "", Optional includeCommandLineAPI As Boolean = False, Optional runImmediately As Boolean = False) As String
    id_ = id_ + 1

    Dim params As String
    params = """source"": """ & EscapeJsonString(source) & """"
    If worldName <> "" Then params = params & ",""worldName"": """ & EscapeJsonString(worldName) & """"
    params = params & ",""includeCommandLineAPI"": " & LCase$(CStr(includeCommandLineAPI))
    params = params & ",""runImmediately"": " & LCase$(CStr(runImmediately))

    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Page.addScriptToEvaluateOnNewDocument"",""params"": {" & params & "}}"
    PageAddScriptToEvaluateOnNewDocument = message
End Function

' Page.removeScriptToEvaluateOnNewDocument のメッセージを生成する
Public Function PageRemoveScriptToEvaluateOnNewDocument(identifier As String) As String
    id_ = id_ + 1

    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Page.removeScriptToEvaluateOnNewDocument"",""params"": {" & _
              """identifier"": " & JsonString(identifier) & "}}"

    PageRemoveScriptToEvaluateOnNewDocument = message
End Function

' Page.setLifecycleEventsEnabled のメッセージを生成する
Public Function PageSetLifecycleEventsEnabled(enabled As Boolean) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Page.setLifecycleEventsEnabled"",""params"": {""enabled"": " & LCase$(CStr(enabled)) & "}}"
    PageSetLifecycleEventsEnabled = message
End Function

' Page.setInterceptFileChooserDialog のメッセージを生成する
Public Function PageSetInterceptFileChooserDialog(enabled As Boolean) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Page.setInterceptFileChooserDialog"",""params"": {""enabled"": " & LCase$(CStr(enabled)) & "}}"
    PageSetInterceptFileChooserDialog = message
End Function

' Network Domain
Public Function NetworkEnable() As String
    id_ = id_ + 1
    NetworkEnable = "{""id"": " & id_ & ",""method"": ""Network.enable""}"
End Function
' Network Domain
Public Function NetworkDisable() As String
    id_ = id_ + 1
    NetworkDisable = "{""id"": " & id_ & ",""method"": ""Network.disable""}"
End Function

' レスポンス本文取得メッセージを生成する
Public Function NetworkGetResponseBody(requestId As String) As String
    id_ = id_ + 1

    NetworkGetResponseBody = "{""id"": " & id_ & ",""method"": ""Network.getResponseBody"",""params"": {" & _
                             """requestId"": " & JsonString(requestId) & "}}"
End Function

' Page Domain
Public Function SetLifecycleEventsEnabled(enabled As Boolean) As String
    id_ = id_ + 1
    Dim v As String
    If enabled Then v = "true" Else v = "false"
    SetLifecycleEventsEnabled = "{""id"": " & id_ & ",""method"": ""Page.setLifecycleEventsEnabled"",""params"": {""enabled"": " & v & "}}"
End Function

'========================
' Storage Domain (Cookies / StorageKey)
'========================

' Storage.getCookies のメッセージを生成する
Public Function StorageGetCookies(Optional browserContextId As String = "") As String
    id_ = id_ + 1

    If browserContextId <> "" Then
        StorageGetCookies = "{""id"": " & id_ & ",""method"": ""Storage.getCookies"",""params"": {" & _
                            """browserContextId"": " & JsonString(browserContextId) & "}}"
    Else
        StorageGetCookies = "{""id"": " & id_ & ",""method"": ""Storage.getCookies""}"
    End If
End Function

' Storage.setCookies のメッセージを生成する
Public Function StorageSetCookies(cookieParamsJsonArray As String, Optional browserContextId As String = "") As String
    ' cookieParamsJsonArray は Network.CookieParam の JSON 配列を受け取る
    id_ = id_ + 1
    If browserContextId <> "" Then
        StorageSetCookies = "{""id"": " & id_ & ",""method"": ""Storage.setCookies"",""params"":{""cookies"": " & cookieParamsJsonArray & ",""browserContextId"":""" & EscapeJsonString(browserContextId) & """}}"
    Else
        StorageSetCookies = "{""id"": " & id_ & ",""method"": ""Storage.setCookies"",""params"":{""cookies"": " & cookieParamsJsonArray & "}}"
    End If
End Function

' Storage.clearCookies のメッセージを生成する
Public Function StorageClearCookies(Optional browserContextId As String = "") As String
    id_ = id_ + 1

    If browserContextId <> "" Then
        StorageClearCookies = "{""id"": " & id_ & ",""method"": ""Storage.clearCookies"",""params"": {" & _
                              """browserContextId"": " & JsonString(browserContextId) & "}}"
    Else
        StorageClearCookies = "{""id"": " & id_ & ",""method"": ""Storage.clearCookies""}"
    End If
End Function

' Storage.getStorageKey のメッセージを生成する
Public Function StorageGetStorageKey(Optional frameId As String = "") As String
    id_ = id_ + 1

    If frameId <> "" Then
        StorageGetStorageKey = "{""id"": " & id_ & ",""method"": ""Storage.getStorageKey"",""params"": {" & _
                               """frameId"": " & JsonString(frameId) & "}}"
    Else
        StorageGetStorageKey = "{""id"": " & id_ & ",""method"": ""Storage.getStorageKey""}"
    End If
End Function

'========================
' DOMStorage Domain (LocalStorage / SessionStorage)
'========================

' DOMStorage 識別子 JSON を生成する
Private Function MakeDOMStorageId(securityOrigin As String, storageKey As String, isLocal As Boolean) As String
    Dim b As String
    b = LCase$(CStr(isLocal))
    MakeDOMStorageId = "{""securityOrigin"":""" & EscapeJsonString(securityOrigin) & """,""storageKey"":""" & EscapeJsonString(storageKey) & """,""isLocalStorage"":" & b & "}"
End Function

' DOMStorage.enable のメッセージを生成する
Public Function DOMStorageEnable() As String
    id_ = id_ + 1
    DOMStorageEnable = "{""id"": " & id_ & ",""method"": ""DOMStorage.enable""}"
End Function

' DOMStorage.disable のメッセージを生成する
Public Function DOMStorageDisable() As String
    id_ = id_ + 1
    DOMStorageDisable = "{""id"": " & id_ & ",""method"": ""DOMStorage.disable""}"
End Function

' DOMStorage 項目設定メッセージを生成する
Public Function DOMStorageSetItem(securityOrigin As String, storageKey As String, isLocal As Boolean, key As String, value As String) As String
    id_ = id_ + 1
    DOMStorageSetItem = _
        "{""id"": " & id_ & _
        ",""method"": ""DOMStorage.setDOMStorageItem"",""params"":{""storageId"":" & _
        MakeDOMStorageId(securityOrigin, storageKey, isLocal) & _
        ",""key"":""" & EscapeJsonString(key) & """,""value"":""" & EscapeJsonString(value) & """}}"
End Function

' DOMStorage 項目削除メッセージを生成する
Public Function DOMStorageRemoveItem(securityOrigin As String, storageKey As String, isLocal As Boolean, key As String) As String
    id_ = id_ + 1
    DOMStorageRemoveItem = _
        "{""id"": " & id_ & _
        ",""method"": ""DOMStorage.removeDOMStorageItem"",""params"":{""storageId"":" & _
        MakeDOMStorageId(securityOrigin, storageKey, isLocal) & _
        ",""key"":""" & EscapeJsonString(key) & """}}"
End Function

' DOMStorage 項目取得メッセージを生成する
Public Function DOMStorageGetItems(securityOrigin As String, storageKey As String, isLocal As Boolean) As String
    id_ = id_ + 1
    DOMStorageGetItems = _
        "{""id"": " & id_ & _
        ",""method"": ""DOMStorage.getDOMStorageItems"",""params"":{""storageId"":" & _
        MakeDOMStorageId(securityOrigin, storageKey, isLocal) & _
        "}}"
End Function

' DOMStorage クリアメッセージを生成する
Public Function DOMStorageClear(securityOrigin As String, storageKey As String, isLocal As Boolean) As String
    id_ = id_ + 1
    DOMStorageClear = _
        "{""id"": " & id_ & _
        ",""method"": ""DOMStorage.clear"",""params"":{""storageId"":" & _
        MakeDOMStorageId(securityOrigin, storageKey, isLocal) & _
        "}}"
End Function

' JavaScript ダイアログ詳細制御メッセージを生成する
Public Function HandleJavaScriptDialogEx(accept As Boolean, Optional promptText As String = "") As String
    id_ = id_ + 1

    Dim params As String
    params = """accept"": " & LCase$(CStr(accept))
    If promptText <> "" Then
        params = params & ",""promptText"": """ & EscapeJsonString(promptText) & """"
    End If

    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Page.handleJavaScriptDialog"",""params"": {" & params & "}}"

    HandleJavaScriptDialogEx = message
End Function

' JSON 文字列をエスケープする
Private Function EscapeJsonString(value As String) As String
    Dim s As String
    s = Replace(value, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    EscapeJsonString = s
End Function

' JSON 文字列リテラルへ変換する
Private Function JsonString(ByVal value As String) As String
    JsonString = """" & EscapeJsonString(value) & """"
End Function


