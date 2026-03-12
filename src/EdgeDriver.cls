VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EdgeDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Implements IWebDriver

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
' Private Const CLASS_NAME As String = "Chrome_WidgetWin_1"
'Private Const CAPTION_NAME As String = "新しいタブ - Google Chrome" '"data:, - Google Chrome"

'Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
'        ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
'

Private Declare PtrSafe Function GetWindowThreadProcessId Lib "USER32.DLL" ( _
        ByVal hWnd As LongPtr, ByRef ProcessId As Long) As Long

Private Declare PtrSafe Sub SetForeGroundWindow Lib "user32" Alias "SetForegroundWindow" (ByVal ms As LongPtr)

Private webSocket_ As a1_WebSocketCommunicator
Private json_ As a2_JSONHandler
Private basicInfo_ As a3_BasicInfos
Private CDP_ As a4_ExecuteCDP
Private events_ As a5_CDPEventHandler
Private Helper_ As a6_ExecuteHelperFunction

Private processId_ As Long
'Private chromeHwnd_ As LongPtr

Private Sub Class_Initialize()
    On Error GoTo EH

    Dim userDataDir As String
    Dim needsBootstrap As Boolean

    userDataDir = CreateObject("Shell.Application").Namespace("shell:Downloads").Self.path & "\マクロ実行専用UserDataフォルダForEdge\"

    EnsureUserDataFolder userDataDir
    needsBootstrap = (IsFirstRunCompleted(userDataDir) = False)

    If needsBootstrap Then
        DeleteFirstRunMarker userDataDir
        WarmUpFirstRun userDataDir
    End If

    processId_ = StartEdge(userDataDir)
    New_InnerObjects

    CDP_.SetDiscoverTargets
    CDP_.AttachToTarget basicInfo_.targetId

    CDP_.PageEnable
    CDP_.DOMEnable
    CDP_.RuntimeEnable
    CDP_.NetworkEnable
    CDP_.SetLifecycleEventsEnabled True

    WaitForActivation

    If needsBootstrap Then
        MarkFirstRunCompleted userDataDir
    End If

    Exit Sub

EH:
    Dim errNum As Long
    Dim errSrc As String
    Dim errDesc As String

    errNum = Err.Number
    errSrc = Err.source
    errDesc = Err.Description

    On Error Resume Next
    CloseInternalSession
    On Error GoTo 0

    If needsBootstrap Then
        DeleteFirstRunMarker userDataDir
        Err.Raise vbObjectError + 2601, "EdgeDriver.Class_Initialize", _
                  "初回セットアップを完了できませんでした。専用UserDataフォルダを削除して再実行してください。"
    Else
        Err.Raise errNum, errSrc, errDesc
    End If
End Sub

' Edge をリモートデバッグ有効で起動する
Private Function StartEdge(userDataDir As String) As Long
    Const EDGE_EXE_REG_KEY As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe\"
    Const DEBUG_MODE_OPTION As String = " --remote-debugging-Port=9222"
    Const USER_DATA_DIR_OPTION As String = " --user-data-dir="
    Const ENABLE_AUTOMATION_OPTION As String = " --enable-automation"

    Dim edgeExePath As String
    Dim objExec As Object
    With CreateObject("WScript.Shell")
        edgeExePath = """" & .RegRead(EDGE_EXE_REG_KEY) & """"
        Set objExec = .Exec( _
                            edgeExePath & DEBUG_MODE_OPTION & ENABLE_AUTOMATION_OPTION & " --disable-popup-blocking --no-first-run --disable-fre" _
                            & USER_DATA_DIR_OPTION & "" & userDataDir & "")
        StartEdge = objExec.ProcessId
        Set objExec = Nothing
    End With
End Function

' 専用 UserData フォルダを作成する
Private Sub EnsureUserDataFolder(ByVal userDataDir As String)
    If Dir$(userDataDir, vbDirectory) = "" Then
        MkDir userDataDir
    End If
End Sub

' 初回実行完了マーカーファイルのパスを返す
Private Function GetFirstRunMarkerPath(ByVal userDataDir As String) As String
    If Right$(userDataDir, 1) <> "\" Then userDataDir = userDataDir & "\"
    GetFirstRunMarkerPath = userDataDir & "init.ok"
End Function

' 初回実行完了マーカーの有無を判定する
Private Function IsFirstRunCompleted(ByVal userDataDir As String) As Boolean
    IsFirstRunCompleted = (Dir$(GetFirstRunMarkerPath(userDataDir)) <> "")
End Function

' 初回実行完了マーカーを作成する
Private Sub MarkFirstRunCompleted(ByVal userDataDir As String)
    Dim markerPath As String
    Dim ff As Integer

    markerPath = GetFirstRunMarkerPath(userDataDir)
    ff = FreeFile

    Open markerPath For Output As #ff
    Print #ff, "ok"
    Close #ff
End Sub

' 初回実行完了マーカーを削除する
Private Sub DeleteFirstRunMarker(ByVal userDataDir As String)
    Dim markerPath As String
    markerPath = GetFirstRunMarkerPath(userDataDir)

    If Dir$(markerPath) <> "" Then
        Kill markerPath
    End If
End Sub

' 内部セッションを終了して参照を解放する
Private Sub CloseInternalSession()
    On Error Resume Next

    If Not CDP_ Is Nothing Then
        CDP_.SessionClose
    End If

    If Not webSocket_ Is Nothing Then
        webSocket_.CloseWebSocket
    End If

    Set CDP_ = Nothing
    Set webSocket_ = Nothing
    Set events_ = Nothing
    Set basicInfo_ = Nothing
    Set json_ = Nothing
    Set Helper_ = Nothing

    On Error GoTo 0
End Sub

' 初回起動安定化のためウォームアップを行う
Private Sub WarmUpFirstRun(ByVal userDataDir As String)
    Const FIRST_RUN_EXTRA_LAUNCHES As Long = 2
    Dim i As Long

    For i = 1 To FIRST_RUN_EXTRA_LAUNCHES
        processId_ = StartEdge(userDataDir)
        New_InnerObjects

        Helper_.SleepByWinAPI 500
        CloseInternalSession
    Next
End Sub
' ブラウザ起動完了を待機し、前面化にも使用する
Private Sub WaitForActivation()
On Error Resume Next

    Dim counter As Long
recheck:
    AppActivate processId_
    If Err.Number <> 0 Then
        counter = counter + 1
        If counter > 20 Then
            On Error GoTo 0
            Err.Raise vbObjectError + 2401, typeName(Me) & ".WaitForActivation", _
                      "10秒経過しましたがブラウザが有効になりませんでした。マクロで開いたブラウザが起動したままであれば、落としてから再実行してください。"
        End If
        Helper_.SleepByWinAPI 500
        Err.Number = 0
        GoTo recheck
    End If

On Error GoTo 0
End Sub

' targetId に対応する WebSocket パスを取得する
Private Function FindWebSocketPathByTargetId(tabs As Collection, ByVal targetId As String) As String
    Dim i As Long
    For i = 1 To tabs.count
        If tabs(i)("targetId") = targetId Then
            FindWebSocketPathByTargetId = Replace(tabs(i)("webSocketDebuggerUrl"), "ws://127.0.0.1:9222", "")
            Exit Function
        End If
    Next
End Function

' 指定 targetId へ再接続する
Private Sub ReconnectToTarget(ByVal targetId As String)
    Dim tabs As Collection: Set tabs = GetPageTabs()
    Dim webSocketURLPath As String
    webSocketURLPath = FindWebSocketPathByTargetId(tabs, targetId)

    If webSocketURLPath = "" Then
        Err.Raise vbObjectError + 2403, typeName(Me) & ".ReconnectToTarget", _
            "再接続先の targetId が見つかりませんでした。"
    End If

    On Error Resume Next
    If Not webSocket_ Is Nothing Then webSocket_.CloseWebSocket
    On Error GoTo 0

    Set webSocket_ = New a1_WebSocketCommunicator
    If webSocket_.Init(webSocketURLPath) = False Then
        Err.Raise vbObjectError + 2404, typeName(Me) & ".ReconnectToTarget", _
            "WebSocket再接続に失敗しました。"
    End If

    Dim msgGenerator As a4x1_CDPMessageGenerator
    Set msgGenerator = New a4x1_CDPMessageGenerator

    Dim Handler As a4x2_CDPMessageHandler
    Set Handler = New a4x2_CDPMessageHandler
    Handler.Init webSocket_, json_, events_

    Set CDP_ = New a4_ExecuteCDP
    CDP_.Init msgGenerator, Handler, json_, events_

    basicInfo_.Update targetId, json_, True
    CDP_.AttachToTarget targetId
    CDP_.ActivateTarget targetId
    CDP_.PageEnable
    CDP_.DOMEnable
    CDP_.RuntimeEnable
    CDP_.NetworkEnable
    CDP_.SetLifecycleEventsEnabled True

    basicInfo_.DocumentObjectId = CDP_.GetDocumentObjectID
    basicInfo_.DocumentNodeId = CDP_.GetDocumentNodeID
End Sub

' 内部オブジェクトを生成する
Private Sub New_InnerObjects()
    Set json_ = New a2_JSONHandler
    Set basicInfo_ = New a3_BasicInfos: basicInfo_.Init json_

    Set webSocket_ = New a1_WebSocketCommunicator
    If webSocket_.Init(basicInfo_.webSocketURLPath) = False Then
        Err.Raise vbObjectError + 2402, typeName(Me) & ".New_InnerObjects", "WebSocket初期化に失敗しました。"
    End If

    Set events_ = New a5_CDPEventHandler
    events_.Init json_, basicInfo_

    Dim msgGenerator As a4x1_CDPMessageGenerator
    Set msgGenerator = New a4x1_CDPMessageGenerator

    Dim Handler As a4x2_CDPMessageHandler
    Set Handler = New a4x2_CDPMessageHandler
    Handler.Init webSocket_, json_, events_

    Set CDP_ = New a4_ExecuteCDP
    CDP_.Init msgGenerator, Handler, json_, events_

    Set Helper_ = New a6_ExecuteHelperFunction
End Sub

' page タブ一覧を取得する
Private Function GetPageTabs() As Collection
    Dim client As Object: Set client = CreateObject("MSXML2.ServerXMLHTTP")
    client.Open "GET", "http://127.0.0.1:9222/json"
    client.setRequestHeader "Content-Type", "application/json"
    client.send

    Dim res As String: res = client.responseText
    Dim count As Long: count = json_.GetJsonLength(res)

    Dim tmpCollection As Collection: Set tmpCollection = New Collection
    Dim jsObjStr As String, typeName As String
    Dim tempDic As Object
    Dim i As Long

    For i = 0 To count - 1
        jsObjStr = json_.GetObject(res, i)
        typeName = json_.GetValue(jsObjStr, "type")
        If typeName = "page" Then
            Set tempDic = CreateObject("Scripting.Dictionary")
            tempDic.Add "title", json_.GetValue(jsObjStr, "title")
            tempDic.Add "targetId", json_.GetValue(jsObjStr, "id")
            tempDic.Add "webSocketDebuggerUrl", json_.GetValue(jsObjStr, "webSocketDebuggerUrl")
            tmpCollection.Add tempDic
        End If
    Next

    Dim tabs As Collection: Set tabs = New Collection
    For i = tmpCollection.count To 1 Step -1
        tabs.Add tmpCollection(i)
    Next

    Set GetPageTabs = tabs
End Function

' 以下、IWebDriver 公開メソッド
'************************************************
' Public
' 現在のページ URL を返す
Private Property Get IWebDriver_URL() As String
    CDP_.SetDiscoverTargets  'これを使ってPage情報最新化
    Dim i As Long
    For i = 1 To events_.PageInfos.count
        If events_.PageInfos(i)("targetId") = basicInfo_.targetId Then
            IWebDriver_URL = events_.PageInfos(i)("url")
            Exit Property
        End If
    Next
End Property

' Public
' 現在のページの HTML ソースを取得する
Private Function IWebDriver_PageSource(target As TargetIs) As String
    Dim res As String
    Select Case target
        Case TargetIs.html_outerhtml
            res = CDP_.ExecuteJavaScript("document.all[0].outerHTML;")
        Case TargetIs.html_innerhtml
            res = CDP_.ExecuteJavaScript("document.all[0].innerHTML;")
        Case TargetIs.body_outerhtml
            res = CDP_.ExecuteJavaScript("document.body.outerHTML;")
        Case TargetIs.body_innerhtml
            res = CDP_.ExecuteJavaScript("document.body.innerHTML;")
        Case TargetIs.body_innertext
            res = CDP_.ExecuteJavaScript("document.body.innerText;")
    End Select

    Dim htmlSource As String
    htmlSource = json_.GetValue(res, "result", "result", "value")
    IWebDriver_PageSource = htmlSource
End Function

' Public
' 指定 URL を開く
Private Sub IWebDriver_OpenURL(targetURL As String)
    Dim res As String
    res = CDP_.OpenURL(targetURL)
    If InStr(res, "invalid URL") > 0 Then
        Helper_.DisplayErrorAndEnd _
            "＜OpenURL＞" & vbCrLf & _
            "指定された以下URLは無効なURLです。" & vbCrLf & _
            targetURL & vbCrLf
    End If

    CDP_.WaitingForWebPageDisplay

    CDP_.SetDiscoverTargets
    CDP_.PageEnable
'    CDP_.DOMEnable
    basicInfo_.DocumentObjectId = CDP_.GetDocumentObjectID
    basicInfo_.DocumentNodeId = CDP_.GetDocumentNodeID
End Sub

' Public
' JavaScript コードを実行する
Private Function IWebDriver_ExecuteScript(script As String) As String
    Dim responseMsg As String
    responseMsg = CDP_.ExecuteJavaScript(script)
    IWebDriver_ExecuteScript = json_.GetValue(responseMsg, "result", "result", "value")
End Function

' Public
' XPath 指定で要素を取得する
Private Function IWebDriver_FindElementByXpath(xPath As String) As IWebElement

    ' 属性指定で日本語を含む場合に備えて URL エンコードする
    xPath = json_.EncodeURL(xPath)

    Dim Codes As Variant: Codes = Helper_.GetXpathCodes(xPath)
    Dim objectId As String: objectId = CDP_.GetObjectIdOfElementBy(xPath, basicInfo_, Codes)
    If objectId = "" Then GoTo quit
    Dim findInfo As Collection: Set findInfo = Helper_.MakeFindInfo("xpath", xPath)
    Set IWebDriver_FindElementByXpath = _
        CDP_.MakeElementObject(objectId, findInfo, basicInfo_, events_, Helper_)
Exit Function

quit:
'Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & xPath
'Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("XPath", xPath)
End Function

' Public
' Id 指定で要素を取得する
Private Function IWebDriver_FindElementById(Id As String) As IWebElement
    Dim Codes As Variant: Codes = Helper_.GetIdCodes(Id)
    Dim objectId As String: objectId = CDP_.GetObjectIdOfElementBy(Id, basicInfo_, Codes)
    If objectId = "" Then GoTo quit
    Dim findInfo As Collection: Set findInfo = Helper_.MakeFindInfo("id", Id)
    Set IWebDriver_FindElementById = _
        CDP_.MakeElementObject(objectId, findInfo, basicInfo_, events_, Helper_)
Exit Function

quit:
'Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & Id
'Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("Id", Id)
End Function

' Public
' Name 指定で要素を取得する
Private Function IWebDriver_FindElementByName(name As String) As IWebElement

    Dim Codes As Variant: Codes = Helper_.GetNameCodes(name)
    Dim objectId As String: objectId = CDP_.GetObjectIdOfElementBy(name, basicInfo_, Codes)
    If objectId = "" Then GoTo quit
    Dim findInfo As Collection: Set findInfo = Helper_.MakeFindInfo("name", name)
    Set IWebDriver_FindElementByName = _
        CDP_.MakeElementObject(objectId, findInfo, basicInfo_, events_, Helper_)
Exit Function

quit:
'Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & name
'Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("Name", name)
End Function

' Public
' タグ名指定で要素コレクションを取得する
Private Function IWebDriver_FindElementsByTag(tagName As String) As IWebElements
    Dim countCodes As Variant: countCodes = Helper_.GetTagCountCodes(tagName)
    Dim elemsCount As Long: elemsCount = CDP_.GetElementsCountBy(tagName, basicInfo_, countCodes)
    If elemsCount = 0 Then GoTo quit

    Dim elements As Collection: Set elements = New Collection

    Dim objectId As String
    Dim Codes As Variant
    Dim index As Long
    For index = 0 To elemsCount - 1
        Codes = Helper_.GetTagCodes(tagName, index)
        objectId = CDP_.GetObjectIdOfElementBy(tagName, basicInfo_, Codes)
        Dim findInfo As Collection: Set findInfo = Helper_.MakeFindInfo("tag", tagName, index)
        elements.Add _
            CDP_.MakeElementObject(objectId, findInfo, basicInfo_, events_, Helper_)
    Next

    Dim elems As c0x0_WebElements: Set elems = New c0x0_WebElements
    elems.Init elements

    Set IWebDriver_FindElementsByTag = elems
Exit Function

quit:
'Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & tagName
'Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("Tag", tagName, "○○")
End Function

' Public
' クラス名指定で要素コレクションを取得する
Private Function IWebDriver_FindElementsByClassName(className As String) As IWebElements
    Dim countCodes As Variant: countCodes = Helper_.GetClassNameCountCodes(className)
    Dim elemsCount As Long: elemsCount = CDP_.GetElementsCountBy(className, basicInfo_, countCodes)
    If elemsCount = 0 Then GoTo quit

    Dim elements As Collection: Set elements = New Collection

    Dim objectId As String
    Dim Codes As Variant
    Dim index As Long
    For index = 0 To elemsCount - 1
        Codes = Helper_.GetClassNameCodes(className, index)
        objectId = CDP_.GetObjectIdOfElementBy(className, basicInfo_, Codes)
        Dim findInfo As Collection: Set findInfo = Helper_.MakeFindInfo("className", className, index)
        elements.Add _
            CDP_.MakeElementObject(objectId, findInfo, basicInfo_, events_, Helper_)
    Next

    Dim elems As c0x0_WebElements: Set elems = New c0x0_WebElements
    elems.Init elements

    Set IWebDriver_FindElementsByClassName = elems
Exit Function

quit:
'Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & className
'Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("className", className, "○○")
End Function

' Public
' XPath 指定で要素コレクションを取得する
Private Function IWebDriver_FindElementsByXPath(xPath As String) As IWebElements
    Dim elements As Collection: Set elements = New Collection
    Dim res As String
    Dim arrayObjectId As String
    Dim elemsCount As Long
    Dim index As Long
    Dim itemRes As String
    Dim objectId As String
    Dim findInfo As Collection
    Dim elems As c0x0_WebElements
    Dim buildArrayCode As String
    Dim buildArrayExpr As String
    Dim countCodes As Variant

    xPath = json_.EncodeURL(xPath)

    ' 件数は既存の安定経路を使う
    countCodes = Helper_.GetXPathCountCodes(xPath)
    elemsCount = CDP_.GetElementsCountBy(xPath, basicInfo_, countCodes)
    If elemsCount = 0 Then GoTo quit

    ' 将来のデータ増加時の計算量劣化を防ぐため、
    ' indexごとに先頭から再走査せず、XPath結果を1回だけ配列化して添字参照する
    buildArrayCode = "function(xpath){" & _
                     "var result=document.evaluate(decodeURIComponent(xpath), document, null, XPathResult.ORDERED_NODE_ITERATOR_TYPE, null);" & _
                     "var ary=[];" & _
                     "var n=result.iterateNext();" & _
                     "while(n){ary.push(n);n=result.iterateNext();}" & _
                     "return ary;" & _
                     "}"

    buildArrayExpr = "(function(){" & _
                     "var result=document.evaluate(decodeURIComponent('" & xPath & "'), document, null, XPathResult.ORDERED_NODE_ITERATOR_TYPE, null);" & _
                     "var ary=[];" & _
                     "var n=result.iterateNext();" & _
                     "while(n){ary.push(n);n=result.iterateNext();}" & _
                     "return ary;" & _
                     "})();"

    If basicInfo_.ExecutionContextId = 0 Then
        res = CDP_.ExecuteJavaScript(buildArrayExpr)
    Else
        res = CDP_.CallJavascriptFunctionWrapper(buildArrayCode, basicInfo_, xPath)
    End If

    arrayObjectId = json_.GetValue(res, "result", "result", "objectId")
    If arrayObjectId = "" Then
        Err.Raise vbObjectError + 2405, typeName(Me) & ".FindElementsByXPath", "XPath結果配列の生成に失敗しました。"
    End If

    For index = 0 To elemsCount - 1
        itemRes = CDP_.CallJavaScriptFunction("function(value){return this[Number(value)];}", arrayObjectId, CStr(index))
        objectId = json_.GetValue(itemRes, "result", "result", "objectId")

        If objectId = "" Then
            Err.Raise vbObjectError + 2406, typeName(Me) & ".FindElementsByXPath", "XPath結果の要素取得に失敗しました。index=" & CStr(index)
        End If

        Set findInfo = Helper_.MakeFindInfo("xpath", xPath, index)
        elements.Add CDP_.MakeElementObject(objectId, findInfo, basicInfo_, events_, Helper_)
    Next

    Set elems = New c0x0_WebElements
    elems.Init elements
    Set IWebDriver_FindElementsByXPath = elems
    Exit Function

quit:
    'Beep
    Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("XPath", xPath, "○○")
End Function

' Public
' 指定値に一致するフレームへ切り替える
Private Sub IWebDriver_SwitchFrameByNameOrUrlOrIndex(value As Variant)
    Dim frameInfos As Collection
    Dim frameId As String
    Dim count As Long

    Do 'frame関連イベント発生タイミングによって情報取得に失敗するケースあり。念のため3回までループ。
        Set frameInfos = GetFrameInfos
        If IsNumeric(value) Then
            If frameInfos.count < value Then GoTo quit
            frameId = json_.GetValue(frameInfos(value), "id")
        Else
            frameId = GetTargetFrameId(frameInfos, value)
        End If
        count = count + 1
    Loop While frameId = "" And count < 3

    If frameId = "" Then GoTo quit

    Dim ExecutionContextId As Long
    ExecutionContextId = CDP_.CreateIsolatedFrameWorld(frameId)
    basicInfo_.ExecutionContextId = ExecutionContextId
    basicInfo_.IsolatedFrameId = frameId

Exit Sub

quit:
Helper_.DisplayErrorAndEnd "＜SwitchFrameByNameOrUrlOrIndex＞" & vbCrLf & _
                                "指定された以下のframeは存在しません。" & vbCrLf & _
                                value & vbCrLf & vbCrLf & _
                                "※iframeにName属性があれば、Nameでの指定を推奨します。" & vbCrLf & vbCrLf & _
                                "※フレームのIndex番号、Name、URLを確認したい場合、" & vbCrLf & _
                                "　 DebugPrintAccessibleFrameInfoを使用してください。" & vbCrLf & _
                                "　 クロームから通知されるフレーム情報がイミディエイトに出力されます。" & vbCrLf & vbCrLf & _
                                "※このメソッドでうまくいかない場合、" & vbCrLf & _
                                "　 SwitchFrameByIframeElementの" & vbCrLf & _
                                "　 使用を検討してください。" & vbCrLf
End Sub

' frame ツリーから各フレーム情報をコレクションへ集約する
' frame ツリーは result → frameTree 配下に存在する
' frame と childFrames の再帰構造で構成される
Private Function GetFrameInfos() As Collection
    Dim collect As Collection: Set collect = New Collection
    Dim json As String: json = CDP_.GetFrameTree
'    collect.add json_.GetObject(json, "result", "frameTree", "frame"))

    Dim frmInfo As String, frminfo2 As String
    Dim count As Long
    Do
        ' childFrames から 1 件の frame 情報を取得する
        frmInfo = json_.GetObject(json, "result", "frameTree", "childFrames", count)
        If frmInfo = "" Then Exit Do
        ' frame プロパティを取得してコレクションへ追加する
        frminfo2 = json_.GetObject(frmInfo, "frame")
        collect.Add frminfo2
        ' childFrames がある場合は再帰処理で同様に集約する
        If InStr(frmInfo, "childFrames") > 0 Then
            GetChildFrameInfo collect, frmInfo
        End If
        count = count + 1
    Loop
    Set GetFrameInfos = collect
End Function
' childFrame がある場合は再帰処理する
Private Sub GetChildFrameInfo(collect As Collection, json As String)
    Dim frmInfo As String, frminfo2 As String
    Dim count As Long
    Do
        frmInfo = json_.GetObject(json, "childFrames", count)
        If frmInfo = "" Then Exit Do
        frminfo2 = json_.GetObject(frmInfo, "frame")
        collect.Add frminfo2
        If InStr(frmInfo, "childFrames") > 0 Then
            GetChildFrameInfo collect, frmInfo
        End If
        count = count + 1
    Loop
End Sub
' name または url が一致するフレーム ID を返す
Private Function GetTargetFrameId(frameInfos As Collection, value As Variant) As String
    Dim frameInfo As Variant
    Dim name As String, URL As String
    For Each frameInfo In frameInfos
        name = json_.GetValue(frameInfo, "name")
        URL = json_.GetValue(frameInfo, "url")
        If name = value Or URL = value Then
            GetTargetFrameId = json_.GetValue(frameInfo, "id")
            Exit Function
        End If
    Next
End Function

' Public
' 切り替え可能なフレーム情報をデバッグ出力する
Private Sub IWebDriver_DebugPrintAccessibleFrameInfo()
    CDP_.PageEnable
    CDP_.RuntimeEnable
    Dim frameInfos As Collection: Set frameInfos = GetFrameInfos
    Dim i As Long
    For i = 1 To frameInfos.count
        'Debug.Print "index：" & i & _
                           "　name：" & json_.GetValue(frameInfos(i), "name") & _
                           "　url：" & json_.GetValue(frameInfos(i), "url")
    Next
End Sub

' Public

' iframe 要素を基準にフレームを切り替える
Private Sub IWebDriver_SwitchFrameByIframeElement(IframeElement As IWebElement)
    Dim frameInfos As Collection: Set frameInfos = GetFrameInfos
    Dim elem As b0x0_WebElement: Set elem = IframeElement
    If elem.frameId = "" Then
        Helper_.DisplayErrorAndEnd "＜SwitchFrameByIframeElement＞" & vbCrLf & _
                                      "指定されたElementがIframeではありません" & vbCrLf & vbCrLf & _
                                      "Iframe要素を引数に渡して下さい" & vbCrLf
    End If

    CDP_.GetTargetInfo
    CDP_.SetDiscoverTargets

    'frameElementのframeIdが正当なframeIdならIsolatedWorldで対応
    Dim ExecutionContextId As Long
    ExecutionContextId = CDP_.CreateIsolatedFrameWorld(elem.frameId)
    If ExecutionContextId <> 0 Then
        basicInfo_.ExecutionContextId = ExecutionContextId
        basicInfo_.IsolatedFrameId = elem.frameId
        Exit Sub
    Else
        Helper_.DisplayErrorAndEnd "＜SwitchFrameByIframeElement＞" & vbCrLf & _
                                       "このIframeはこのメソッドでは切り替えができません。" & vbCrLf & vbCrLf & _
                                      "※SwitchFrameByNameOrUrlOrIndexの" & vbCrLf & _
                                      "　 使用を検討してください。" & vbCrLf
    End If

End Sub

' Public
' デフォルトフレームへ戻す
Private Sub IWebDriver_SwitchFrameToDefault()
    If basicInfo_.IsolatedFrameId <> "" Then
        basicInfo_.IsolatedFrameId = ""
        basicInfo_.ExecutionContextId = 0
    End If

'    If basicInfo_.IsTargetPageDefault = False Then
'        SwitchTargetToPageDefault
'    End If
End Sub

' Public
' XPath 指定要素の存在有無を判定する
Private Function IWebDriver_IsExistsElementByXpath(xPath As String) As Boolean
    ' 属性指定で日本語を含む場合に備えて URL エンコードする
    xPath = json_.EncodeURL(xPath)

    IWebDriver_IsExistsElementByXpath = _
        CDP_.IsExistsElementBy(xPath, basicInfo_, Helper_.GetXpathCodes(xPath))
End Function
' Public
' Id 指定要素の存在有無を判定する
Private Function IWebDriver_IsExistsElementById(Id As String) As Boolean
    IWebDriver_IsExistsElementById = _
        CDP_.IsExistsElementBy(Id, basicInfo_, Helper_.GetIdCodes(Id))
End Function
' Public
' Name 指定要素の存在有無を判定する
Private Function IWebDriver_IsExistsElementByName(name As String) As Boolean
    IWebDriver_IsExistsElementByName = _
        CDP_.IsExistsElementBy(name, basicInfo_, Helper_.GetNameCodes(name))
End Function

' Public
' XPath 指定要素の出現を待機する
Private Sub IWebDriver_WaitForElementToAppearByXpath(xPath As String, maxWaitTime_sec As Long)
    ' 属性指定で日本語を含む場合に備えて URL エンコードする
    xPath = json_.EncodeURL(xPath)

    If CDP_.WaitForElementToAppearBy(xPath, _
        maxWaitTime_sec, basicInfo_, Helper_.GetXpathCodes(xPath)) = False Then

        Helper_.DisplayErrorAndEnd _
            "<WaitForElementToAppearByXPath>" & vbCrLf & _
             maxWaitTime_sec & "秒経過しても、指定の以下要素は現れませんでした。" & vbCrLf & _
             xPath
    End If
End Sub
' Public
' Id 指定要素の出現を待機する
Private Sub IWebDriver_WaitForElementToAppearById(Id As String, maxWaitTime_sec As Long)
    If CDP_.WaitForElementToAppearBy(Id, _
        maxWaitTime_sec, basicInfo_, Helper_.GetIdCodes(Id)) = False Then

        Helper_.DisplayErrorAndEnd _
            "<WaitForElementToAppearById>" & vbCrLf & _
             maxWaitTime_sec & "秒経過しても、指定の以下要素は現れませんでした。" & vbCrLf & _
             Id
    End If
End Sub
' Public
' Name 指定要素の出現を待機する
Private Sub IWebDriver_WaitForElementToAppearByName(name As String, maxWaitTime_sec As Long)
    If CDP_.WaitForElementToAppearBy(name, _
        maxWaitTime_sec, basicInfo_, Helper_.GetNameCodes(name)) = False Then

        Helper_.DisplayErrorAndEnd _
            "<WaitForElementToAppearByName>" & vbCrLf & _
             maxWaitTime_sec & "秒経過しても、指定の以下要素は現れませんでした。" & vbCrLf & _
             name
    End If
End Sub

' Public
' Alert ダイアログを受諾して閉じる
Private Sub IWebDriver_HandleAlertDialog()
    Dim res As String
    Dim fileSafeCounter As Long

    Do
        Helper_.SleepByWinAPI 500
        res = CDP_.HandleAlertDialog
        If InStr(res, "No dialog is showing") = 0 Then
            Helper_.SleepByWinAPI 500
            Exit Sub
        End If

        fileSafeCounter = fileSafeCounter + 1
    Loop While fileSafeCounter < 6

quit:
    Helper_.DisplayErrorAndEnd "<HandleAlertDialog>" & vbCrLf & _
                                   "ダイアログが表示されていません。" & vbCrLf & _
                                   "※このメソッドはAlertが表示される場合に使用します。" & vbCrLf
End Sub

' WinAPI の Sleep を呼び出す
Private Sub IWebDriver_SleepByWinAPI(milliSecond As Long)
    Helper_.SleepByWinAPI milliSecond
End Sub

' Index 指定でタブを切り替える
Private Sub IWebDriver_SwitchTabByIndex(index As Long)
    Dim tabs As Collection
    Set tabs = GetPageTabs()

    If index < 1 Or index > tabs.count Then
        Err.Raise vbObjectError + 2405, typeName(Me) & ".SwitchTabByIndex", _
            "指定されたIndexがタブ数の範囲外です。"
    End If

    ReconnectToTarget CStr(tabs(index)("targetId"))
End Sub

' 現在開いているタブタイトル一覧を取得する
Public Function IWebDriver_GetTabList() As String
    Dim client As Object: Set client = CreateObject("MSXML2.ServerXMLHTTP")
    client.Open "GET", "http://127.0.0.1:9222/json"
    client.setRequestHeader "Content-Type", "application/json"
    client.send

    Dim res As String: res = client.responseText
    Dim count As Long: count = json_.GetJsonLength(res)

    Dim tmpCollection As Collection: Set tmpCollection = New Collection
    Dim jsObjStr As String, typeName As String
    Dim tempDic As Object
    Dim i As Long

    ' page タイプのタブだけを抽出する
    For i = 0 To count - 1
        jsObjStr = json_.GetObject(res, i)
        typeName = json_.GetValue(jsObjStr, "type")
        If typeName = "page" Then
            Set tempDic = CreateObject("Scripting.Dictionary")
            tempDic.Add "title", json_.GetValue(jsObjStr, "title")
            tempDic.Add "targetId", json_.GetValue(jsObjStr, "id")
            tempDic.Add "webSocketDebuggerUrl", json_.GetValue(jsObjStr, "webSocketDebuggerUrl")
            tmpCollection.Add tempDic
        End If
    Next

    ' JSON 返却順が右→左のため逆順へ並べ替える
    Dim tabCollection As Collection: Set tabCollection = New Collection
    For i = tmpCollection.count To 1 Step -1
        tabCollection.Add tmpCollection(i)
    Next

    ' タイトルをカンマ区切りで連結する
    Dim result As String
    For i = 1 To tabCollection.count
        result = result & tabCollection(i)("title")
        If i < tabCollection.count Then result = result & ","
    Next

    IWebDriver_GetTabList = result
End Function
' タイトル一致でタブを切り替える
Public Sub IWebDriver_SwitchTabByTitle(tabTitle As String)
    Dim client As Object: Set client = CreateObject("MSXML2.ServerXMLHTTP")
    client.Open "GET", "http://127.0.0.1:9222/json"
    client.setRequestHeader "Content-Type", "application/json"
    client.send

    Dim res As String: res = client.responseText
    Dim count As Long: count = json_.GetJsonLength(res)

    Dim tmpCollection As Collection: Set tmpCollection = New Collection
    Dim jsObjStr As String, typeName As String
    Dim tempDic As Object
    Dim i As Long

    ' page タイプのタブだけを抽出する
    For i = 0 To count - 1
        jsObjStr = json_.GetObject(res, i)
        typeName = json_.GetValue(jsObjStr, "type")
        If typeName = "page" Then
            Set tempDic = CreateObject("Scripting.Dictionary")
            tempDic.Add "title", json_.GetValue(jsObjStr, "title")
            tempDic.Add "targetId", json_.GetValue(jsObjStr, "id")
            tempDic.Add "webSocketDebuggerUrl", json_.GetValue(jsObjStr, "webSocketDebuggerUrl")
            tmpCollection.Add tempDic
        End If
    Next

    ' 右→左の順序を左→右へ並べ替える
    Dim tabCollection As Collection: Set tabCollection = New Collection
    For i = tmpCollection.count To 1 Step -1
        tabCollection.Add tmpCollection(i)
    Next

    ' 一致するタイトルを探索する
    Dim idx As Long: idx = -1
    For i = 1 To tabCollection.count
        If tabCollection(i)("title") = tabTitle Then
            idx = i
            Exit For
        End If
    Next

    If idx = -1 Then
        Err.Raise Number:=10001, Description:="指定タイトルのタブが見つかりません: " & tabTitle
    End If

    ' 見つかった index で既存ロジックを利用する
    Call IWebDriver_SwitchTabByIndex(idx)
End Sub

' 現在セッションのウィンドウを閉じる
Private Sub IWebDriver_CloseWindow()
    CDP_.SessionClose
End Sub

' ウィンドウを指定量スクロールする
Private Sub IWebDriver_ScrollWindow(yPixel As Long, Optional xPixel As Long = 0&)
    CDP_.DispatchMouseEvent "mouseWheel", 0, 0, "none", 0, 0, xPixel, yPixel
End Sub

' ブラウザを前面化する
Private Sub IWebDriver_SetForeGroundWindow()
    WaitForActivation
End Sub

'Private Sub IWebDriver_AllowPopUpWindowForURL(targetURL As String)
'    Dim chromeExe As String
'    Dim objExec As Object
'    With CreateObject("WScript.Shell")
'        .RegWrite name:="HKEY_LOCAL_MACHINE\SOFTWARE\Policies\{Google}", _
'                        value:=targetURL, _
'                        Type:="REG_SZ"
'    End With
'End Sub

Private Sub IWebDriver_Back()
    Dim result As String
    result = CDP_.DispatchMouseEvent("mousePressed", 0, 0, "back", 1)
    result = CDP_.DispatchMouseEvent("mouseReleased", 0, 0, "back", 1)
End Sub

' ブラウザを進む操作する
Private Sub IWebDriver_Forward()
    Dim result As String
    result = CDP_.DispatchMouseEvent("mousePressed", 0, 0, "forward", 1)
    result = CDP_.DispatchMouseEvent("mouseReleased", 0, 0, "forward", 1)
End Sub

' DownloadWatch 関連処理
Private Sub IWebDriver_DownloadWatchStart(downloadDir As String, Optional allowAndName As Boolean = False)
 events_.ClearDownloadInfo
 CDP_.SetDownloadBehavior downloadDir, allowAndName, True
End Sub

' ダウンロード完了まで待機する
Private Function IWebDriver_DownloadWatchWaitCompleted(maxWaitTime_sec As Long) As Boolean
 Dim startT As Single
 startT = Timer
 Do
  ' イベント受信を進める
  CDP_.ExecuteJavaScript "0;"
  If events_.DownloadInfoDic("State") = "completed" Then
   IWebDriver_DownloadWatchWaitCompleted = True
   Exit Function
  End If
  If CLng(Timer - startT) > maxWaitTime_sec Then Exit Function
  DoEvents
  Helper_.SleepByWinAPI 200
 Loop
End Function

' 提案ファイル名を返す
Private Function IWebDriver_DownloadWatchSuggestedFilename() As String
 IWebDriver_DownloadWatchSuggestedFilename = CStr(events_.DownloadInfoDic("SuggestedFilename"))
End Function

' ダウンロードファイルパスを返す
Private Function IWebDriver_DownloadWatchFilePath() As String
 IWebDriver_DownloadWatchFilePath = CStr(events_.DownloadInfoDic("FilePath"))
End Function

' 現在ページを PDF 保存し、保存先パスを返す
Private Function IWebDriver_PrintToPDFSaveAsFile( _
    Optional folderPath As String = "", _
    Optional fileName As String = "", _
    Optional printBackground As Boolean = True, _
    Optional landscape As Boolean = False) As String

    If folderPath <> "" Then
        If Dir(folderPath, vbDirectory) = "" Then Helper_.DisplayErrorAndEnd "指定されたフォルダが存在しません"
    End If

    If fileName <> "" Then
        If Helper_.IncludeNGCharInFileName(fileName) Then Helper_.DisplayErrorAndEnd "指定されたファイル名に使用不可の文字が含まれています"
    End If

    Dim base64Pdf As String
    base64Pdf = CDP_.PrintToPDF(printBackground, landscape)

    Dim decodedPdf As Variant
    decodedPdf = Helper_.DecodeBase64(base64Pdf)

    If folderPath = "" Then
        folderPath = ThisWorkbook.path
    Else
        If Right(folderPath, 1) = "\" Then folderPath = left(folderPath, Len(folderPath) - 1)
    End If
    If fileName = "" Then fileName = "page" & Helper_.GetFileNumber(folderPath, "pdf")

    Dim fullPath As String: fullPath = folderPath & "\" & fileName & ".pdf"

    With CreateObject("ADODB.Stream")
        .Type = 1
        .Open
        .Write decodedPdf
        .SaveToFile fullPath, 2
        .Close
    End With

    IWebDriver_PrintToPDFSaveAsFile = fullPath
End Function

' 新規ドキュメント評価用スクリプトを登録する
Private Function IWebDriver_AddScriptToEvaluateOnNewDocument(source As String, Optional worldName As String = "", Optional includeCommandLineAPI As Boolean = False, Optional runImmediately As Boolean = False) As String
    IWebDriver_AddScriptToEvaluateOnNewDocument = CDP_.AddScriptToEvaluateOnNewDocument(source, worldName, includeCommandLineAPI, runImmediately)
End Function

' 登録済みスクリプトを削除する
Private Sub IWebDriver_RemoveScriptToEvaluateOnNewDocument(identifier As String)
    CDP_.RemoveScriptToEvaluateOnNewDocument identifier
End Sub

' ライフサイクルイベント通知を設定する
Private Sub IWebDriver_SetLifecycleEventsEnabled(enabled As Boolean)
    CDP_.SetLifecycleEventsEnabled enabled
End Sub

' ファイル選択ダイアログ割り込みを設定する
Private Sub IWebDriver_SetInterceptFileChooserDialog(enabled As Boolean)
    CDP_.SetInterceptFileChooserDialog enabled
End Sub

' LifecycleEvents / FileChooser (公開取得API)
Private Function IWebDriver_GetLifecycleEvents() As Collection
 Dim c As Collection
 If events_.LifecycleEvents Is Nothing Then
  Set c = New Collection
 Else
  Set c = events_.LifecycleEvents
 End If
 Set IWebDriver_GetLifecycleEvents = c
End Function

' ライフサイクルイベント一覧をクリアする
Private Sub IWebDriver_ClearLifecycleEvents()
 events_.ClearLifecycleEvents
End Sub

' FileChooser の BackendNodeId を返す
Private Function IWebDriver_GetFileChooserBackendNodeId() As Long
 On Error Resume Next
 IWebDriver_GetFileChooserBackendNodeId = CLng(events_.FileChooserInfoDic("BackendNodeId"))
 On Error GoTo 0
End Function

' FileChooser の mode を返す
Private Function IWebDriver_GetFileChooserMode() As String
 On Error Resume Next
 IWebDriver_GetFileChooserMode = CStr(events_.FileChooserInfoDic("Mode"))
 On Error GoTo 0
End Function

' FileChooser 情報をクリアする
Private Sub IWebDriver_ClearFileChooserInfo()
 events_.ClearFileChooserInfo
End Sub

' Screenshotを取り、画像ファイルとして保存し、保存先パスを返す。
Private Function IWebDriver_ScreenShotSaveAsFile( _
    Optional folderPath As String = "", _
    Optional fileName As String = "", _
    Optional imageFormat As Image = Image.png) As String

    If folderPath <> "" Then
        If Dir(folderPath, vbDirectory) = "" Then Helper_.DisplayErrorAndEnd "指定されたフォルダが存在しません"
    End If

    If fileName <> "" Then
        If Helper_.IncludeNGCharInFileName(fileName) Then Helper_.DisplayErrorAndEnd "指定されたファイル名に使用不可の文字が含まれています"
    End If

    Dim formatName As String: formatName = Helper_.GetFormatName(imageFormat)
    Dim base64EncodedImageData As String, decodedImageData As Variant
    base64EncodedImageData = CDP_.CaptureScreenshot(formatName)
    decodedImageData = Helper_.DecodeBase64(base64EncodedImageData)

    If folderPath = "" Then
        folderPath = ThisWorkbook.path
    Else
        If Right(folderPath, 1) = "\" Then folderPath = left(folderPath, Len(folderPath) - 1)
    End If
    If fileName = "" Then fileName = "screenshot" & Helper_.GetFileNumber(folderPath, formatName)

    Dim fullPath As String: fullPath = folderPath & "\" & fileName & "." & formatName

    With CreateObject("ADODB.Stream")
        .Type = 1
        .Open
        .Write decodedImageData
        .SaveToFile fullPath, 2
        .Close
    End With

    IWebDriver_ScreenShotSaveAsFile = fullPath
End Function

' FileChooser 情報を使ってファイルを設定する
Public Sub IWebDriver_FileChooserSetFile(filePath As String, Optional maxWaitTime_sec As Long = 10)
    Dim startT As Single: startT = Timer

    Do
        CDP_.ExecuteJavaScript "0;"   'イベント受信を回す

        If CLng(events_.FileChooserInfoDic("BackendNodeId")) <> 0 Then Exit Do
        If CLng(Timer - startT) > maxWaitTime_sec Then Exit Do

        DoEvents
        Helper_.SleepByWinAPI 100
    Loop

    If CLng(events_.FileChooserInfoDic("BackendNodeId")) = 0 Then
        Helper_.DisplayErrorAndEnd "<FileChooserSetFile> Page.fileChooserOpenedが来ませんでした"
    End If

    CDP_.SetFileInputFilesByBackendNodeId CLng(events_.FileChooserInfoDic("BackendNodeId")), filePath
End Sub

' DOM 変化が発生するまで待機する
Private Sub IWebDriver_WaitForDomMutation(maxWaitTime_sec As Long)
    Dim js As String
    Dim res As String
    Dim startCount As Long, currCount As Long
    Dim startTime As Double

    js = ""
    js = js & "(function(){"
    js = js & " if(window.__vbaDomMutationObserverInstalled) return true;"
    js = js & " window.__vbaDomMutationCount=0;"
    js = js & " var obs=new MutationObserver(function(muts){"
    js = js & "  window.__vbaDomMutationCount += muts.length;"
    js = js & " });"
    js = js & " var root=document.documentElement||document;"
    js = js & " obs.observe(root,{subtree:true,childList:true,attributes:true,characterData:true});"
    js = js & " window.__vbaDomMutationObserverInstalled=true;"
    js = js & " return true;"
    js = js & "})();"

    Call CDP_.ExecuteJavaScript(js)

    res = CDP_.ExecuteJavaScript("window.__vbaDomMutationCount||0;")
    startCount = CLng(json_.GetValue(res, "result", "result", "value"))

    startTime = Timer()
    Do
        res = CDP_.ExecuteJavaScript("window.__vbaDomMutationCount||0;")
        currCount = CLng(json_.GetValue(res, "result", "result", "value"))
        If currCount > startCount Then Exit Sub

        If (Timer() - startTime) > CDbl(maxWaitTime_sec) Then Exit Do
        DoEvents
        Helper_.SleepByWinAPI 200
    Loop

    Helper_.DisplayErrorAndEnd "<WaitForDomMutation>" & vbCrLf & _
        maxWaitTime_sec & "秒待機しましたがDOM変化が発生しませんでした。"
End Sub

' Network idle になるまで待機する
Private Sub IWebDriver_WaitForNetworkIdle(maxWaitTime_sec As Long)
    Dim js As String
    Dim res As String, isIdle As String
    Dim startTime As Double

    js = ""
    js = js & "(function(){"
    js = js & " if(window.__vbaNetHookInstalled) return true;"
    js = js & " window.__vbaNetInflight=0; window.__vbaNetLastActivity=Date.now();"
    js = js & " if(window.fetch){"
    js = js & "  var _f=window.fetch;"
    js = js & "  window.fetch=function(){"
    js = js & "   window.__vbaNetInflight++; window.__vbaNetLastActivity=Date.now();"
    js = js & "   return _f.apply(this,arguments).finally(function(){"
    js = js & "    window.__vbaNetInflight--; window.__vbaNetLastActivity=Date.now();"
    js = js & "   });"
    js = js & "  };"
    js = js & " }"
    js = js & " var _open=XMLHttpRequest.prototype.open;"
    js = js & " var _send=XMLHttpRequest.prototype.send;"
    js = js & " XMLHttpRequest.prototype.open=function(){ this.__vbaTracked=true; return _open.apply(this,arguments); };"
    js = js & " XMLHttpRequest.prototype.send=function(){"
    js = js & "  if(this.__vbaTracked){"
    js = js & "   window.__vbaNetInflight++; window.__vbaNetLastActivity=Date.now();"
    js = js & "   this.addEventListener('loadend',function(){"
    js = js & "    window.__vbaNetInflight--; window.__vbaNetLastActivity=Date.now();"
    js = js & "   },{once:true});"
    js = js & "  }"
    js = js & "  return _send.apply(this,arguments);"
    js = js & " };"
    js = js & " window.__vbaNetHookInstalled=true; return true;"
    js = js & "})();"

    Call CDP_.ExecuteJavaScript(js)

    startTime = Timer()
    Do
        res = CDP_.ExecuteJavaScript("((window.__vbaNetInflight||0)==0 && (Date.now()-(window.__vbaNetLastActivity||0))>=500);")
        isIdle = CStr(json_.GetValue(res, "result", "result", "value"))
        If isIdle = "True" Then Exit Sub

        If (Timer() - startTime) > CDbl(maxWaitTime_sec) Then Exit Do
        DoEvents
        Helper_.SleepByWinAPI 200
    Loop

    Helper_.DisplayErrorAndEnd "<WaitForNetworkIdle>" & vbCrLf & _
        maxWaitTime_sec & "秒待機しましたがNetworkIdleになりませんでした。"
End Sub

' Network ドメインを無効化する
Private Sub IWebDriver_NetworkDisable()
    CDP_.NetworkDisable
End Sub

' 保持中のレスポンス一覧をクリアする
Private Sub IWebDriver_ClearNetworkResponses()
    events_.ClearNetworkResponses
End Sub

' 指定 URL を含むレスポンス到着を待機する
Private Function IWebDriver_WaitForResponse(urlContains As String, timeoutSec As Long) As String
    events_.ClearNetworkResponses

    Dim startT As Single: startT = Timer

    Do
        'イベント受信を進める（DownloadWatchWaitCompleted と同じ考え方）
        CDP_.ExecuteJavaScript "0;"  ' ←イベントをevents_へ流す [3](https://toshibaap-my.sharepoint.com/personal/taketo_ikeda_r45_mail_toshiba/Documents/20260303_153736/ChromeDriver.cls)[4](https://toshibaap-my.sharepoint.com/personal/taketo_ikeda_r45_mail_toshiba/Documents/20260303_153736/EdgeDriver.cls)[1](https://toshibaap-my.sharepoint.com/personal/taketo_ikeda_r45_mail_toshiba/Documents/20260303_153736/a4x2_CDPMessageHandler.cls)

        Dim i As Long
        For i = events_.NetworkResponses.count To 1 Step -1
            If InStr(1, CStr(events_.NetworkResponses(i)("Url")), urlContains, vbTextCompare) > 0 Then
                IWebDriver_WaitForResponse = CStr(events_.NetworkResponses(i)("RequestId"))
                Exit Function
            End If
        Next

        If CLng(Timer - startT) > timeoutSec Then Exit Do
        DoEvents
        Helper_.SleepByWinAPI 100
    Loop

    Helper_.DisplayErrorAndEnd "<WaitForResponse>" & vbCrLf & _
        timeoutSec & "秒待機しましたが指定URLを含むレスポンスが来ませんでした。" & vbCrLf & _
        urlContains
End Function

' requestId に対応するレスポンス本文を返す
Private Function IWebDriver_GetResponseBody(requestId As String) As String
    Dim res As String
    res = CDP_.GetResponseBody(requestId)
    IWebDriver_GetResponseBody = json_.GetObject(res, "result")
End Function

'----------------------------
' internal helpers
'----------------------------
Private Function GetCurrentSecurityOrigin() As String
    Dim res As String
    res = CDP_.ExecuteJavaScript("location.origin;")
    GetCurrentSecurityOrigin = json_.GetValue(res, "result", "result", "value")
End Function

' 現在ページの storageKey を取得する
Private Function GetCurrentStorageKey() As String
    ' frameId 省略で current target の storageKey を取る [1](https://chromedevtools.github.io/devtools-protocol/tot/Storage/)
    Dim res As String
    res = CDP_.StorageGetStorageKey()
    GetCurrentStorageKey = json_.GetValue(res, "result", "storageKey")
End Function

' DOMStorage ドメインを有効化する
Private Sub EnsureDOMStorageEnabled()
    Static enabled As Boolean
    If enabled = False Then
        CDP_.DOMStorageEnable
        enabled = True
    End If
End Sub

' DOMStorage の全項目を取得する
Private Function DOMStorageGetAll(isLocal As Boolean) As String
    EnsureDOMStorageEnabled
    Dim origin As String: origin = GetCurrentSecurityOrigin()
    Dim sk As String: sk = GetCurrentStorageKey()
    Dim res As String
    res = CDP_.DOMStorageGetItems(origin, sk, isLocal)
    ' entries 配列だけを返す
    DOMStorageGetAll = json_.GetObject(res, "result", "entries")
End Function

' DOMStorage から指定キーの値を取得する
Private Function DOMStorageGetItemValue(isLocal As Boolean, key As String) As String
    Dim entriesJson As String
    entriesJson = DOMStorageGetAll(isLocal)

    Dim n As Long: n = json_.GetJsonLength(entriesJson)
    Dim i As Long
    For i = 0 To n - 1
        Dim pairJson As String
        pairJson = json_.GetObject(entriesJson, i) ' ["k","v"]
        If json_.GetValue(pairJson, 0) = key Then
            DOMStorageGetItemValue = json_.GetValue(pairJson, 1)
            Exit Function
        End If
    Next
    DOMStorageGetItemValue = ""
End Function

' DOMStorage へ項目を設定する
Private Sub DOMStorageSetItemCore(isLocal As Boolean, key As String, value As String)
    EnsureDOMStorageEnabled
    Dim origin As String: origin = GetCurrentSecurityOrigin()
    Dim sk As String: sk = GetCurrentStorageKey()
    Call CDP_.DOMStorageSetItem(origin, sk, isLocal, key, value)
End Sub

' DOMStorage から項目を削除する
Private Sub DOMStorageRemoveItemCore(isLocal As Boolean, key As String)
    EnsureDOMStorageEnabled
    Dim origin As String: origin = GetCurrentSecurityOrigin()
    Dim sk As String: sk = GetCurrentStorageKey()
    Call CDP_.DOMStorageRemoveItem(origin, sk, isLocal, key)
End Sub

' DOMStorage をクリアする
Private Sub DOMStorageClearCore(isLocal As Boolean)
    EnsureDOMStorageEnabled
    Dim origin As String: origin = GetCurrentSecurityOrigin()
    Dim sk As String: sk = GetCurrentStorageKey()
    Call CDP_.DOMStorageClear(origin, sk, isLocal)
End Sub

'----------------------------
' Cookies (Storage domain)
'----------------------------
Private Function IWebDriver_GetCookies() As String
    Dim res As String
    res = CDP_.StorageGetCookies()
    IWebDriver_GetCookies = json_.GetObject(res, "result", "cookies")
End Function

' Cookie を削除する
Private Sub IWebDriver_ClearCookies()
    Call CDP_.StorageClearCookies
End Sub

' Cookie 配列を設定する
Private Function IWebDriver_SetCookies(cookieParamsJsonArray As String) As Boolean
    Dim res As String
    res = CDP_.StorageSetCookies(cookieParamsJsonArray)
    ' SendAndReceive 側で CDP error を処理する前提で True を返す
    IWebDriver_SetCookies = True
End Function

' Cookie 1 件を設定する
Private Function IWebDriver_SetCookie( _
    name As String, _
    value As String, _
    Optional URL As String = "", _
    Optional domain As String = "", _
    Optional path As String = "/", _
    Optional secure As Boolean = False, _
    Optional httpOnly As Boolean = False, _
    Optional sameSite As String = "", _
    Optional expires As Double = 0 _
) As Boolean

    ' CookieParam を1件だけ組み立てて Storage.setCookies に渡す [1](https://chromedevtools.github.io/devtools-protocol/tot/Storage/)
    Dim j As String
    j = "{""name"":""" & EscapeJsonStringLocal(name) & """,""value"":""" & EscapeJsonStringLocal(value) & """"
    If URL <> "" Then j = j & ",""url"":""" & EscapeJsonStringLocal(URL) & """"
    If domain <> "" Then j = j & ",""domain"":""" & EscapeJsonStringLocal(domain) & """"
    If path <> "" Then j = j & ",""path"":""" & EscapeJsonStringLocal(path) & """"
    If secure Then j = j & ",""secure"":true" Else j = j & ",""secure"":false"
    If httpOnly Then j = j & ",""httpOnly"":true" Else j = j & ",""httpOnly"":false"
    If sameSite <> "" Then j = j & ",""sameSite"":""" & EscapeJsonStringLocal(sameSite) & """"
    If expires <> 0 Then j = j & ",""expires"":" & CStr(expires)
    j = j & "}"
    Call CDP_.StorageSetCookies("[" & j & "]")
    IWebDriver_SetCookie = True
End Function

'----------------------------
' LocalStorage (DOMStorage domain)
'----------------------------
Private Function IWebDriver_LocalStorageGetAll() As String
    IWebDriver_LocalStorageGetAll = DOMStorageGetAll(True)
End Function

' LocalStorage の指定値を取得する
Private Function IWebDriver_LocalStorageGetItem(key As String) As String
    IWebDriver_LocalStorageGetItem = DOMStorageGetItemValue(True, key)
End Function

' LocalStorage へ項目を設定する
Private Sub IWebDriver_LocalStorageSetItem(key As String, value As String)
    DOMStorageSetItemCore True, key, value
End Sub

' LocalStorage から項目を削除する
Private Sub IWebDriver_LocalStorageRemoveItem(key As String)
    DOMStorageRemoveItemCore True, key
End Sub

' LocalStorage をクリアする
Private Sub IWebDriver_LocalStorageClear()
    DOMStorageClearCore True
End Sub

'----------------------------
' SessionStorage (DOMStorage domain)
'----------------------------
Private Function IWebDriver_SessionStorageGetAll() As String
    IWebDriver_SessionStorageGetAll = DOMStorageGetAll(False)
End Function

' SessionStorage の指定値を取得する
Private Function IWebDriver_SessionStorageGetItem(key As String) As String
    IWebDriver_SessionStorageGetItem = DOMStorageGetItemValue(False, key)
End Function

' SessionStorage へ項目を設定する
Private Sub IWebDriver_SessionStorageSetItem(key As String, value As String)
    DOMStorageSetItemCore False, key, value
End Sub

' SessionStorage から項目を削除する
Private Sub IWebDriver_SessionStorageRemoveItem(key As String)
    DOMStorageRemoveItemCore False, key
End Sub

' SessionStorage をクリアする
Private Sub IWebDriver_SessionStorageClear()
    DOMStorageClearCore False
End Sub

' 指定 index のタブを閉じる
Private Sub IWebDriver_CloseTab(index As Long)
    Dim tabs As Collection: Set tabs = GetPageTabs()

    If index < 1 Or index > tabs.count Then
        Err.Raise vbObjectError + 2406, typeName(Me) & ".CloseTab", _
            "指定されたIndexがタブ数の範囲外です。"
    End If

    Dim targetId As String
    Dim nextTargetId As String
    Dim wasCurrent As Boolean

    targetId = tabs(index)("targetId")
    wasCurrent = (targetId = basicInfo_.targetId)

    If wasCurrent And tabs.count > 1 Then
        If index < tabs.count Then
            nextTargetId = tabs(index + 1)("targetId")
        Else
            nextTargetId = tabs(index - 1)("targetId")
        End If
    End If

    Call CDP_.TargetClose(targetId)

    If wasCurrent Then
        If nextTargetId <> "" Then
            Helper_.SleepByWinAPI 200
            ReconnectToTarget nextTargetId
        Else
            On Error Resume Next
            If Not webSocket_ Is Nothing Then webSocket_.CloseWebSocket
            On Error GoTo 0
            basicInfo_.targetId = ""
            basicInfo_.PageDefaultTargetId = ""
            basicInfo_.DocumentObjectId = ""
            basicInfo_.DocumentNodeId = 0
            basicInfo_.ExecutionContextId = 0
            basicInfo_.IsolatedFrameId = ""
        End If
    End If
End Sub

' 指定タイトルのタブを閉じる
Private Sub IWebDriver_CloseTabByTitle(tabTitle As String)
    Dim tabs As Collection: Set tabs = GetPageTabs()
    Dim i As Long

    For i = 1 To tabs.count
        If tabs(i)("title") = tabTitle Then
            Call IWebDriver_CloseTab(i)
            Exit Sub
        End If
    Next

    Err.Raise vbObjectError + 2407, typeName(Me) & ".CloseTabByTitle", _
        "指定タイトルのタブが見つかりません: " & tabTitle
End Sub

' CSS セレクタ指定で要素を取得する
Private Function IWebDriver_FindElementByCss(cssSelector As String) As IWebElement
    cssSelector = json_.EncodeURL(cssSelector)

    Dim Codes As Variant
    Codes = Helper_.GetCssCodes(cssSelector)

    Dim objectId As String
    objectId = CDP_.GetObjectIdOfElementBy(cssSelector, basicInfo_, Codes)
    If objectId = "" Then GoTo quit

    Dim findInfo As Collection
    Set findInfo = Helper_.MakeFindInfo("css", cssSelector)

    Set IWebDriver_FindElementByCss = _
        CDP_.MakeElementObject(objectId, findInfo, basicInfo_, events_, Helper_)
    Exit Function

quit:
    Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("Css", cssSelector)
End Function

' CSS セレクタ指定で要素コレクションを取得する
Private Function IWebDriver_FindElementsByCss(cssSelector As String) As IWebElements
    cssSelector = json_.EncodeURL(cssSelector)

    Dim countCodes As Variant
    countCodes = Helper_.GetCssCountCodes(cssSelector)

    Dim elemsCount As Long
    elemsCount = CDP_.GetElementsCountBy(cssSelector, basicInfo_, countCodes)
    If elemsCount = 0 Then GoTo quit

    Dim elements As Collection
    Set elements = New Collection

    Dim objectId As String
    Dim Codes As Variant
    Dim index As Long

    For index = 0 To elemsCount - 1
        Codes = Helper_.GetCssResultCodes(cssSelector, index)
        objectId = CDP_.GetObjectIdOfElementBy(cssSelector, basicInfo_, Codes)

        Dim findInfo As Collection
        Set findInfo = Helper_.MakeFindInfo("css", cssSelector, index)

        elements.Add _
            CDP_.MakeElementObject(objectId, findInfo, basicInfo_, events_, Helper_)
    Next

    Dim elems As c0x0_WebElements
    Set elems = New c0x0_WebElements
    elems.Init elements

    Set IWebDriver_FindElementsByCss = elems
    Exit Function

quit:
    Helper_.DisplayErrorAndEnd _
        "<FindElementsByCss>" & vbCrLf & _
        "指定された以下CSSセレクタの要素は存在しませんでした。" & vbCrLf & _
        cssSelector & vbCrLf & _
        "(イミディエイトウィンドウで確認可能)"
End Function

' テキスト指定で要素を取得する
Private Function IWebDriver_FindElementByText(text As String, Optional partialMatch As Boolean = False) As IWebElement
    text = json_.EncodeURL(text)

    Dim Codes As Variant
    Codes = Helper_.GetTextCodes(text, partialMatch)

    Dim objectId As String
    objectId = CDP_.GetObjectIdOfElementBy(text, basicInfo_, Codes)
    If objectId = "" Then GoTo quit

    Dim findKind As String
    If partialMatch Then
        findKind = "text_partial"
    Else
        findKind = "text"
    End If

    Dim findInfo As Collection
    Set findInfo = Helper_.MakeFindInfo(findKind, text)

    Set IWebDriver_FindElementByText = _
        CDP_.MakeElementObject(objectId, findInfo, basicInfo_, events_, Helper_)
    Exit Function

quit:
    Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("Text", text)
End Function

' テキスト指定で要素コレクションを取得する
Private Function IWebDriver_FindElementsByText(text As String, Optional partialMatch As Boolean = False) As IWebElements
    text = json_.EncodeURL(text)

    Dim countCodes As Variant
    countCodes = Helper_.GetTextCountCodes(text, partialMatch)

    Dim elemsCount As Long
    elemsCount = CDP_.GetElementsCountBy(text, basicInfo_, countCodes)
    If elemsCount = 0 Then GoTo quit

    Dim elements As Collection
    Set elements = New Collection

    Dim objectId As String
    Dim Codes As Variant
    Dim index As Long
    Dim findKind As String

    If partialMatch Then
        findKind = "text_partial"
    Else
        findKind = "text"
    End If

    For index = 0 To elemsCount - 1
        Codes = Helper_.GetTextResultCodes(text, index, partialMatch)
        objectId = CDP_.GetObjectIdOfElementBy(text, basicInfo_, Codes)

        Dim findInfo As Collection
        Set findInfo = Helper_.MakeFindInfo(findKind, text, index)

        elements.Add _
            CDP_.MakeElementObject(objectId, findInfo, basicInfo_, events_, Helper_)
    Next

    Dim elems As c0x0_WebElements
    Set elems = New c0x0_WebElements
    elems.Init elements

    Set IWebDriver_FindElementsByText = elems
    Exit Function

quit:
    Helper_.DisplayErrorAndEnd _
        "<FindElementsByText>" & vbCrLf & _
        "指定された以下テキストの要素は存在しませんでした。" & vbCrLf & _
        text & vbCrLf & _
        "(イミディエイトウィンドウで確認可能)"
End Function

' Shadow DOM を含めて CSS セレクタ指定で要素を取得する
Private Function IWebDriver_FindElementByCssDeep(cssSelector As String) As IWebElement
    cssSelector = json_.EncodeURL(cssSelector)

    Dim Codes As Variant
    Codes = Helper_.GetDeepCssCodes(cssSelector)

    Dim objectId As String
    objectId = CDP_.GetObjectIdOfElementBy(cssSelector, basicInfo_, Codes)
    If objectId = "" Then GoTo quit

    Dim findInfo As Collection
    Set findInfo = Helper_.MakeFindInfo("css_deep", cssSelector)

    Set IWebDriver_FindElementByCssDeep = _
        CDP_.MakeElementObject(objectId, findInfo, basicInfo_, events_, Helper_)
    Exit Function

quit:
    Helper_.DisplayErrorAndEnd _
        "<FindElementByCssDeep>" & vbCrLf & _
        "指定された以下CSSセレクタ(Deep)の要素は存在しませんでした。" & vbCrLf & _
        cssSelector & vbCrLf & _
        "(イミディエイトウィンドウで確認可能)"
End Function

' Shadow DOM を含めて CSS セレクタ指定で要素コレクションを取得する
Private Function IWebDriver_FindElementsByCssDeep(cssSelector As String) As IWebElements
    cssSelector = json_.EncodeURL(cssSelector)

    Dim countCodes As Variant
    countCodes = Helper_.GetDeepCssCountCodes(cssSelector)

    Dim elemsCount As Long
    elemsCount = CDP_.GetElementsCountBy(cssSelector, basicInfo_, countCodes)
    If elemsCount = 0 Then GoTo quit

    Dim elements As Collection
    Set elements = New Collection

    Dim objectId As String
    Dim Codes As Variant
    Dim index As Long

    For index = 0 To elemsCount - 1
        Codes = Helper_.GetDeepCssResultCodes(cssSelector, index)
        objectId = CDP_.GetObjectIdOfElementBy(cssSelector, basicInfo_, Codes)

        Dim findInfo As Collection
        Set findInfo = Helper_.MakeFindInfo("css_deep", cssSelector, index)

        elements.Add _
            CDP_.MakeElementObject(objectId, findInfo, basicInfo_, events_, Helper_)
    Next

    Dim elems As c0x0_WebElements
    Set elems = New c0x0_WebElements
    elems.Init elements

    Set IWebDriver_FindElementsByCssDeep = elems
    Exit Function

quit:
    Helper_.DisplayErrorAndEnd _
        "<FindElementsByCssDeep>" & vbCrLf & _
        "指定された以下CSSセレクタ(Deep)の要素は存在しませんでした。" & vbCrLf & _
        cssSelector & vbCrLf & _
        "(イミディエイトウィンドウで確認可能)"
End Function

' Shadow DOM を含めてテキスト指定で要素を取得する
Private Function IWebDriver_FindElementByTextDeep(text As String, Optional partialMatch As Boolean = False) As IWebElement
    text = json_.EncodeURL(text)

    Dim Codes As Variant
    Codes = Helper_.GetDeepTextCodes(text, partialMatch)

    Dim objectId As String
    objectId = CDP_.GetObjectIdOfElementBy(text, basicInfo_, Codes)
    If objectId = "" Then GoTo quit

    Dim findKind As String
    If partialMatch Then
        findKind = "text_deep_partial"
    Else
        findKind = "text_deep"
    End If

    Dim findInfo As Collection
    Set findInfo = Helper_.MakeFindInfo(findKind, text)

    Set IWebDriver_FindElementByTextDeep = _
        CDP_.MakeElementObject(objectId, findInfo, basicInfo_, events_, Helper_)
    Exit Function

quit:
    Helper_.DisplayErrorAndEnd _
        "<FindElementByTextDeep>" & vbCrLf & _
        "指定された以下テキスト(Deep)の要素は存在しませんでした。" & vbCrLf & _
        text & vbCrLf & _
        "(イミディエイトウィンドウで確認可能)"
End Function

' Shadow DOM を含めてテキスト指定で要素コレクションを取得する
Private Function IWebDriver_FindElementsByTextDeep(text As String, Optional partialMatch As Boolean = False) As IWebElements
    text = json_.EncodeURL(text)

    Dim countCodes As Variant
    countCodes = Helper_.GetDeepTextCountCodes(text, partialMatch)

    Dim elemsCount As Long
    elemsCount = CDP_.GetElementsCountBy(text, basicInfo_, countCodes)
    If elemsCount = 0 Then GoTo quit

    Dim elements As Collection
    Set elements = New Collection

    Dim objectId As String
    Dim Codes As Variant
    Dim index As Long
    Dim findKind As String

    If partialMatch Then
        findKind = "text_deep_partial"
    Else
        findKind = "text_deep"
    End If

    For index = 0 To elemsCount - 1
        Codes = Helper_.GetDeepTextResultCodes(text, index, partialMatch)
        objectId = CDP_.GetObjectIdOfElementBy(text, basicInfo_, Codes)

        Dim findInfo As Collection
        Set findInfo = Helper_.MakeFindInfo(findKind, text, index)

        elements.Add _
            CDP_.MakeElementObject(objectId, findInfo, basicInfo_, events_, Helper_)
    Next

    Dim elems As c0x0_WebElements
    Set elems = New c0x0_WebElements
    elems.Init elements

    Set IWebDriver_FindElementsByTextDeep = elems
    Exit Function

quit:
    Helper_.DisplayErrorAndEnd _
        "<FindElementsByTextDeep>" & vbCrLf & _
        "指定された以下テキスト(Deep)の要素は存在しませんでした。" & vbCrLf & _
        text & vbCrLf & _
        "(イミディエイトウィンドウで確認可能)"
End Function

' Alert / Confirm / Prompt を受諾する
Private Function IWebDriver_AcceptAlertDialog(Optional promptText As String = "") As String
    Dim res As String
    Dim failSafeCounter As Long

    Do
        Helper_.SleepByWinAPI 500
        res = CDP_.AcceptAlertDialog(promptText)

        If InStr(res, "No dialog is showing") = 0 Then
            Helper_.SleepByWinAPI 500
            IWebDriver_AcceptAlertDialog = res
            Exit Function
        End If

        failSafeCounter = failSafeCounter + 1
    Loop While failSafeCounter < 6

    Helper_.DisplayErrorAndEnd _
        "<AcceptAlertDialog>" & vbCrLf & _
        "ダイアログが表示されていません。" & vbCrLf & _
        "※このメソッドはAlert/Confirm/Promptが表示される場合に使用します。"
End Function

' Alert / Confirm / Prompt を却下する
Private Function IWebDriver_DismissAlertDialog(Optional promptText As String = "") As String
    Dim res As String
    Dim failSafeCounter As Long

    Do
        Helper_.SleepByWinAPI 500
        res = CDP_.DismissAlertDialog(promptText)

        If InStr(res, "No dialog is showing") = 0 Then
            Helper_.SleepByWinAPI 500
            IWebDriver_DismissAlertDialog = res
            Exit Function
        End If

        failSafeCounter = failSafeCounter + 1
    Loop While failSafeCounter < 6

    Helper_.DisplayErrorAndEnd _
        "<DismissAlertDialog>" & vbCrLf & _
        "ダイアログが表示されていません。" & vbCrLf & _
        "※このメソッドはAlert/Confirm/Promptが表示される場合に使用します。"
End Function

' ネットワーク割り込みスクリプトを有効化する
Private Sub IWebDriver_EnableNetworkInterception()
    Static scriptId As String

    Dim script As String
    script = ""
    script = script & "(function(){"
    script = script & "if(window.__vbaNetworkInterceptionInstalled){return true;}"
    script = script & "window.__vbaNetworkInterceptionInstalled=true;"
    script = script & "window.__vbaBlockedURLPatterns=window.__vbaBlockedURLPatterns||[];"
    script = script & "window.__vbaMockResponses=window.__vbaMockResponses||[];"
    script = script & "window.__vbaMatchUrl=function(url,pattern){try{return String(url).indexOf(String(pattern))>=0;}catch(e){return false;}};"
    script = script & "window.__vbaGetMockResponse=function(url){var i,m;for(i=0;i<window.__vbaMockResponses.length;i++){m=window.__vbaMockResponses[i];if(window.__vbaMatchUrl(url,m.urlPattern)){return m;}}return null;};"
    script = script & "window.__vbaIsBlockedUrl=function(url){var i;for(i=0;i<window.__vbaBlockedURLPatterns.length;i++){if(window.__vbaMatchUrl(url,window.__vbaBlockedURLPatterns[i])){return true;}}return false;};"

    script = script & "if(window.fetch && !window.__vbaOriginalFetch){"
    script = script & "window.__vbaOriginalFetch=window.fetch.bind(window);"
    script = script & "window.fetch=function(input,init){"
    script = script & "var url=(typeof input==='string')?input:((input&&input.url)?input.url:String(input));"
    script = script & "if(window.__vbaIsBlockedUrl(url)){return Promise.reject(new Error('Blocked by VBA interceptor'));}"
    script = script & "var mock=window.__vbaGetMockResponse(url);"
    script = script & "if(mock){return Promise.resolve(new Response(mock.body,{status:mock.statusCode,headers:{'Content-Type':mock.contentType}}));}"
    script = script & "return window.__vbaOriginalFetch(input,init);"
    script = script & "};"
    script = script & "}"

    script = script & "if(window.XMLHttpRequest && !window.__vbaOriginalXHROpen){"
    script = script & "window.__vbaOriginalXHROpen=XMLHttpRequest.prototype.open;"
    script = script & "window.__vbaOriginalXHRSend=XMLHttpRequest.prototype.send;"
    script = script & "XMLHttpRequest.prototype.open=function(method,url){this.__vbaUrl=url;return window.__vbaOriginalXHROpen.apply(this,arguments);};"
    script = script & "XMLHttpRequest.prototype.send=function(){"
    script = script & "var url=this.__vbaUrl||'';"
    script = script & "if(window.__vbaIsBlockedUrl(url)){throw new Error('Blocked by VBA interceptor');}"
    script = script & "var mock=window.__vbaGetMockResponse(url);"
    script = script & "if(mock){"
    script = script & "var xhr=this;"
    script = script & "Object.defineProperty(xhr,'readyState',{value:4,configurable:true});"
    script = script & "Object.defineProperty(xhr,'status',{value:mock.statusCode,configurable:true});"
    script = script & "Object.defineProperty(xhr,'responseText',{value:String(mock.body),configurable:true});"
    script = script & "Object.defineProperty(xhr,'response',{value:String(mock.body),configurable:true});"
    script = script & "setTimeout(function(){if(xhr.onreadystatechange){xhr.onreadystatechange();} if(xhr.onload){xhr.onload();}},0);"
    script = script & "return;"
    script = script & "}"
    script = script & "return window.__vbaOriginalXHRSend.apply(this,arguments);"
    script = script & "};"
    script = script & "}"

    script = script & "return true;"
    script = script & "})();"

    If scriptId = "" Then
        scriptId = CDP_.AddScriptToEvaluateOnNewDocument(script)
    End If

    Call CDP_.ExecuteJavaScript(script)
End Sub

' ネットワーク割り込みスクリプトを無効化する
Private Sub IWebDriver_DisableNetworkInterception()
    Dim script As String

    script = ""
    script = script & "(function(){"
    script = script & "try{"
    script = script & "if(window.__vbaOriginalFetch){"
    script = script & "window.fetch=window.__vbaOriginalFetch;"
    script = script & "window.__vbaOriginalFetch=null;"
    script = script & "}"
    script = script & "if(window.__vbaOriginalXHROpen){"
    script = script & "XMLHttpRequest.prototype.open=window.__vbaOriginalXHROpen;"
    script = script & "window.__vbaOriginalXHROpen=null;"
    script = script & "}"
    script = script & "if(window.__vbaOriginalXHRSend){"
    script = script & "XMLHttpRequest.prototype.send=window.__vbaOriginalXHRSend;"
    script = script & "window.__vbaOriginalXHRSend=null;"
    script = script & "}"
    script = script & "window.__vbaBlockedURLPatterns=[];"
    script = script & "window.__vbaMockResponses=[];"
    script = script & "window.__vbaNetworkInterceptionInstalled=false;"
    script = script & "return true;"
    script = script & "}catch(e){return false;}"
    script = script & "})();"

    Call CDP_.ExecuteJavaScript(script)
End Sub

' 遮断対象 URL パターンを追加する
Private Sub IWebDriver_AddBlockedURLPattern(urlPattern As String)
    Dim encodedPattern As String
    encodedPattern = json_.EncodeURL(urlPattern)

    Call IWebDriver_EnableNetworkInterception

    Dim script As String
    script = ""
    script = script & "(function(){"
    script = script & "window.__vbaBlockedURLPatterns=window.__vbaBlockedURLPatterns||[];"
    script = script & "var p=decodeURIComponent('" & encodedPattern & "');"
    script = script & "if(window.__vbaBlockedURLPatterns.indexOf(p) < 0){"
    script = script & "window.__vbaBlockedURLPatterns.push(p);"
    script = script & "}"
    script = script & "return true;"
    script = script & "})();"

    Call CDP_.ExecuteJavaScript(script)
End Sub

' 遮断対象 URL パターンをクリアする
Private Sub IWebDriver_ClearBlockedURLPatterns()
    Call IWebDriver_EnableNetworkInterception

    Dim script As String
    script = ""
    script = script & "(function(){"
    script = script & "window.__vbaBlockedURLPatterns=[];"
    script = script & "return true;"
    script = script & "})();"

    Call CDP_.ExecuteJavaScript(script)
End Sub

' モックレスポンスを追加する
Private Sub IWebDriver_AddMockResponse( _
    urlPattern As String, _
    statusCode As Long, _
    body As String, _
    Optional contentType As String = "text/plain" _
)
    Dim encodedPattern As String
    Dim encodedBody As String
    Dim encodedContentType As String

    encodedPattern = json_.EncodeURL(urlPattern)
    encodedBody = json_.EncodeURL(body)
    encodedContentType = json_.EncodeURL(contentType)

    Call IWebDriver_EnableNetworkInterception

    Dim script As String
    script = ""
    script = script & "(function(){"
    script = script & "window.__vbaMockResponses=window.__vbaMockResponses||[];"
    script = script & "window.__vbaMockResponses.push({"
    script = script & "urlPattern:decodeURIComponent('" & encodedPattern & "'),"
    script = script & "statusCode:" & statusCode & ","
    script = script & "body:decodeURIComponent('" & encodedBody & "'),"
    script = script & "contentType:decodeURIComponent('" & encodedContentType & "')"
    script = script & "});"
    script = script & "return true;"
    script = script & "})();"

    Call CDP_.ExecuteJavaScript(script)
End Sub

' モックレスポンス一覧をクリアする
Private Sub IWebDriver_ClearMockResponses()
    Call IWebDriver_EnableNetworkInterception

    Dim script As String
    script = ""
    script = script & "(function(){"
    script = script & "window.__vbaMockResponses=[];"
    script = script & "return true;"
    script = script & "})();"

    Call CDP_.ExecuteJavaScript(script)
End Sub

' JSON 文字列をエスケープする
Private Function EscapeJsonStringLocal(ByVal value As String) As String
    Dim s As String
    s = Replace(value, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    EscapeJsonStringLocal = s
End Function


