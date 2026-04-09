VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "b0x0_WebElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IWebElement

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
        ByVal hWndParent As LongPtr, ByVal hwndChildAfter As LongPtr, _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
        ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Any) As LongPtr

Private json_ As a2_JSONHandler
Private basicInfo_ As a3_BasicInfos
Private CDP_ As a4_ExecuteCDP
Private events_ As a5_CDPEventHandler
Private Helper_ As a6_ExecuteHelperFunction

Private objectId_ As String
Private findInfo_ As Collection
Private nodeId_ As Long
Private attributes_ As Object
Private frameId_ As String

Private xpath_ As String
'

Public Sub Init(json As a2_JSONHandler, _
                       basicInfo As a3_BasicInfos, _
                       ExecCDP As a4_ExecuteCDP, _
                       events As a5_CDPEventHandler, _
                       execHelper As a6_ExecuteHelperFunction, _
                       objectId As String, _
                       findInfo As Collection, _
                       Optional attributes As Object = Nothing, _
                       Optional frameId As String)

    Set json_ = json
    Set basicInfo_ = basicInfo
    Set CDP_ = ExecCDP
    Set events_ = events
    Set Helper_ = execHelper

    objectId_ = objectId
    Set findInfo_ = findInfo
    Set attributes_ = attributes
    frameId_ = frameId

End Sub

' 現在の objectId を返す
Public Property Get objectId() As String
    objectId = objectId_
End Property

' 現在の frameId を返す
Public Property Get frameId() As String
    frameId = frameId_
End Property

' CallJavaScriptFunction呼出しの共通関数
' 戻り値が Cannot find context with specified id の場合に備えて ObjectId を再取得する
Private Function CallFunction(functionCode As String, args As Variant) As String
    Dim res As String
    res = CDP_.CallJavaScriptFunction(functionCode, objectId_, args)
    If InStr(res, "Cannot find context with specified id") > 0 Then
        basicInfo_.UpdateExecutionContextId CDP_
        ReGetObjectId
        res = CDP_.CallJavaScriptFunction(functionCode, objectId_, args)
    End If
    CallFunction = res
End Function

' 複数引数版の JavaScript 関数呼び出しを共通処理化する
Private Function CallFunction_ArgMultiple(functionCode As String, ParamArray args() As Variant) As String
    Dim res As String
    res = CDP_.CallJavaScriptFunction_ArgMultiple(functionCode, objectId_, args)
    If InStr(res, "Cannot find context with specified id") > 0 Then
        basicInfo_.UpdateExecutionContextId CDP_
        ReGetObjectId
        res = CDP_.CallJavaScriptFunction_ArgMultiple(functionCode, objectId_, args)
    End If
    CallFunction_ArgMultiple = res
End Function

' 保持中の探索条件から objectId を再取得する
Public Function ReGetObjectId()
    Dim Codes As Variant
    Dim findIndex As Long

    findIndex = CLng(findInfo_(3))

    Select Case findInfo_(1)
      Case "id"
        Codes = Helper_.GetIdCodes(findInfo_(2))

      Case "xpath"
        If findIndex >= 0 Then
            Codes = Helper_.GetXPathResultCodes(findInfo_(2), findIndex)
        Else
            Codes = Helper_.GetXpathCodes(findInfo_(2))
        End If

      Case "name"
        Codes = Helper_.GetNameCodes(findInfo_(2))

      Case "tag"
        Codes = Helper_.GetTagCodes(findInfo_(2), findIndex)

      Case "className"
        Codes = Helper_.GetClassNameCodes(findInfo_(2), findIndex)

      Case "css"
        If findIndex >= 0 Then
            Codes = Helper_.GetCssResultCodes(findInfo_(2), findIndex)
        Else
            Codes = Helper_.GetCssCodes(findInfo_(2))
        End If

      Case "text"
        If findIndex >= 0 Then
            Codes = Helper_.GetTextResultCodes(findInfo_(2), findIndex, False)
        Else
            Codes = Helper_.GetTextCodes(findInfo_(2), False)
        End If

      Case "text_partial"
        If findIndex >= 0 Then
            Codes = Helper_.GetTextResultCodes(findInfo_(2), findIndex, True)
        Else
            Codes = Helper_.GetTextCodes(findInfo_(2), True)
        End If

      Case "css_deep"
        If findIndex >= 0 Then
            Codes = Helper_.GetDeepCssResultCodes(findInfo_(2), findIndex)
        Else
            Codes = Helper_.GetDeepCssCodes(findInfo_(2))
        End If

      Case "text_deep"
        If findIndex >= 0 Then
            Codes = Helper_.GetDeepTextResultCodes(findInfo_(2), findIndex, False)
        Else
            Codes = Helper_.GetDeepTextCodes(findInfo_(2), False)
        End If

      Case "text_deep_partial"
        If findIndex >= 0 Then
            Codes = Helper_.GetDeepTextResultCodes(findInfo_(2), findIndex, True)
        Else
            Codes = Helper_.GetDeepTextCodes(findInfo_(2), True)
        End If

      Case Else
        Err.Raise vbObjectError + 2111, "b0x0_WebElement.ReGetObjectId", _
            "未対応の findInfo 種別です: " & CStr(findInfo_(1))
    End Select

    objectId_ = CDP_.GetObjectIdOfElementBy(findInfo_(2), basicInfo_, Codes)
End Function

'***********************************************
' Public API
' チェックボックス要素の Checked 属性を設定する
Private Property Let IWebElement_Checked(RHS As Boolean)
    Dim res As String
    'checkedプロパティを持っているか事前にチェック。持っていない場合はじく。
    res = CallFunction("function(){return this.checked;}", "")
    If InStr(res, """type"":""boolean""") = 0 Then GoTo quit

    Dim boolString As String
    If RHS Then
        boolString = "true"
    Else
        boolString = "false"
    End If

    res = CallFunction("function(value){ this.checked = value;}", boolString)
Exit Property

quit:
Helper_.DisplayErrorAndEnd _
    "<WebElement：Checkedプロパティ>" & vbCrLf & _
    "このElementはCheckedプロパティを持ちません。" & vbCrLf & _
    "※CheckedはCheckBoxで使用可能" & vbCrLf
End Property
' Public API
' チェックボックス要素のチェック状態を取得する
Private Property Get IWebElement_Checked() As Boolean
    Dim res As String
    res = CallFunction("function(){return this.checked;}", "")
    If InStr(res, """type"":""boolean""") = 0 Then GoTo quit

    Dim Checked As Boolean
    Checked = json_.GetValue(res, "result", "result", "value")
    IWebElement_Checked = Checked
Exit Property

quit:
Helper_.DisplayErrorAndEnd "このElementはCheckedプロパティを持ちません。" & vbCrLf & _
                                            "※CheckedはCheckBoxで使用可能" & vbCrLf
End Property

' Public API
' 要素の有効状態を取得する
Private Property Get IWebElement_IsEnabled() As Boolean
    Dim res As String
    res = CallFunction("function(){return this.disabled;}", "")
    If InStr(res, """type"":""boolean""") = 0 Then GoTo quit

    Dim enabled As Boolean
    enabled = json_.GetValue(res, "result", "result", "value")
    IWebElement_IsEnabled = Not enabled
Exit Property

quit:
Helper_.DisplayErrorAndEnd "このElementはEnabledプロパティを持ちません。" & vbCrLf
End Property

' 要素を表示領域へ収めたうえで CDP マウスイベントを発行する
Private Sub IWebElement_DispatchCDPMouseEvent(action As CDP_Mouse, Optional point As ClickPoint = ClickPoint.Center)

    ' 画面外にいるとクリックできないため、要素の位置を確認
    Dim contentBox As Object: Set contentBox = GetContentBox()
    Dim viewPortHeight As Long: viewPortHeight = GetViewPortHeight()
    '画面の中心に要素がくるように、Scroll
    CDP_.DispatchMouseEvent "mouseWheel", 0, 0, "none", 0, 0, 0, contentBox("TopLeftY") - (viewPortHeight / 2)
    Sleep 200
    'そのうえで、画面内の位置として再度、位置を取得。
    Set contentBox = GetContentBox()

    Dim x As Long, y As Long
    Select Case point
        Case ClickPoint.Center
            x = contentBox("CenterX")
            y = contentBox("CenterY")
        Case ClickPoint.topLeft
            x = contentBox("TopLeftX") + 2
            y = contentBox("TopLeftY") + 2
        Case ClickPoint.TopRight
            x = contentBox("TopRightX") - 2
            y = contentBox("TopRightY") + 2
        Case ClickPoint.BottomLeft
            x = contentBox("BottomLeftX") + 2
            y = contentBox("BottomLeftY") - 2
        Case ClickPoint.BottomRight
            x = contentBox("BottomRightX") - 2
            y = contentBox("BottomRightY") - 2
    End Select

    Select Case action
        Case CDP_Mouse.LeftClick_XY
            CDPClick "left", x, y, 1
        Case CDP_Mouse.LeftDoubleClick_XY
            CDPClick "left", x, y, 2
        Case CDP_Mouse.RightClick_XY
            CDPClick "right", x, y, 1
        Case CDP_Mouse.ScrollToThisElement
            'この要素位置まで移動するだけで何もしない
    End Select
End Sub

'Private Function IsRequireScrollingToThisElement(action As CDP_Mouse, Optional x_pixel As Long = 0, Optional y_pixel As Long = 0) As Boolean
'    If action = CDP_Mouse.Back Or action = CDP_Mouse.Forward _
'    Or action = CDP_Mouse.Moved Or action = CDP_Mouse.MouseWheel Then
'        IsRequireScrollingToThisElement = False
'    Else
'        If x_pixel = 0 And y_pixel = 0 Then
'            IsRequireScrollingToThisElement = True
'        Else
'            IsRequireScrollingToThisElement = False
'        End If
'    End If
'End Function

' ObjectId を引数に取る CDP メソッド呼び出しを共通処理化する
' 戻り値が Cannot find context with specified id の場合に備えて ObjectId を再取得する
' 現在は Focus と GetBoxModel の呼び出しで使用する
Private Function CallObjectIdUsingMethod(methodName As String) As String
    Dim res As String
    res = CallByName(CDP_, methodName, VbMethod, objectId_)

    If InStr(res, "Cannot find context with specified id") > 0 Then
        basicInfo_.UpdateExecutionContextId CDP_
        ReGetObjectId
        res = CallByName(CDP_, methodName, VbMethod, objectId_)
    End If
    CallObjectIdUsingMethod = res
End Function

' BoxModel の content 領域座標を辞書として取得する
Private Function GetContentBox(Optional ByVal res As String = vbNullString) As Object
    If Len(res) = 0 Then
        res = CallObjectIdUsingMethod("GetBoxModel")
    End If

    Dim quadXY As Variant: quadXY = json_.GetValue(res, "result", "model", "content")
    quadXY = Split(quadXY, ",")

    Dim posName As Variant
    posName = Split("TopLeftX,TopLeftY,TopRightX,TopRightY,BottomRightX,BottomRightY,BottomLeftX,BottomLeftY", ",")

    Dim contentBox As Object: Set contentBox = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = LBound(posName) To UBound(posName)
        contentBox.Add posName(i), quadXY(i)
    Next

    Dim width As Long:  width = json_.GetValue(res, "result", "model", "width")
    Dim height As Long: height = json_.GetValue(res, "result", "model", "height")
    contentBox.Add "Width", width
    contentBox.Add "Height", height
    contentBox.Add "CenterX", quadXY(0) + (width / 2)
    contentBox.Add "CenterY", quadXY(1) + (height / 2)

    Set GetContentBox = contentBox
End Function

' viewport の高さを取得する
Private Function GetViewPortHeight() As Long
    Dim res As String
    res = CallFunction("function(){return window.innerHeight;}", "")
    res = json_.GetValue(res, "result", "result", "value")
    GetViewPortHeight = Int(res)
End Function

' 指定座標へ CDP クリックイベントを発行する
Private Sub CDPClick(button As String, x As Long, y As Long, clickNum As Long)
    Dim result As String
    Dim i As Long
    For i = 1 To clickNum
        result = CDP_.DispatchMouseEvent("mousePressed", x, y, button, 1)
        result = CDP_.DispatchMouseEvent("mouseReleased", x, y, button, 1)
    Next
End Sub

' Public API
' CDP によるキーボードイベントを発行する
Private Sub IWebElement_DispatchCDPKeyEvent(action As CDP_Keyboard, KeyCode As key, Optional alt As Boolean = False, Optional shift As Boolean = False, Optional ctrl As Boolean = False)
    Dim actionStr As String
    Select Case action
        Case CDP_Keyboard.KeyDown_
            actionStr = "keyDown"
        Case CDP_Keyboard.KeyUp_
            actionStr = "keyUp"
        Case CDP_Keyboard.RawKeyDown
            actionStr = "rawKeyDown"
        Case CDP_Keyboard.Char
            actionStr = "char"
    End Select

    Dim modifiers As Long
    If alt Then modifiers = 1
    If ctrl Then modifiers = modifiers + 2
    If shift Then modifiers = modifiers + 8

    Dim result As String
    result = CDP_.DispatchKeyEvent(actionStr, modifiers, KeyCode)
End Sub

' Public API
' JavaScript によるクリックを実行する
Private Sub IWebElement_Click()
    CallObjectIdUsingMethod "Focus"

    Dim res As String
    res = CallFunction("function(){this.click();}", "")

    If InStr(res, "this.click is not a function") Then
        res = CallFunction("function(){this.dispatchEvent(new MouseEvent('click', {bubbles: true}));}", "")
    End If

    If InStr(res, "this.dispatchEvent is not a function") Then
        FocusAndEnterKey
    End If

    ' ダイアログが開いた場合は、ここでページ/フレーム変化待ちに入らない
    ' 呼び出し元が AcceptAlertDialog / DismissAlertDialog を行う前に
    ' DOM.enable / Page.enable を投げて足止めされるのを防ぐ
    If events_.DialogInfoDic("IsExistDialog") Then Exit Sub

    If IsPageChanging Then WaitPageChangeToComplete
    If IsIsolatedFrameChanging Then WaitIsolatedFrameChangeToComplete
End Sub

' 要素へフォーカスして Enter キーを送る
Private Sub FocusAndEnterKey()
    CallObjectIdUsingMethod "Focus"
    CDP_.DispatchKeyEvent "keyDown", 0, 13
    CDP_.DispatchKeyEvent "keyUp", 0, 13
End Sub

' 要素へフォーカスして Enter キーを送る
Private Sub IWebElement_FocusAndEnterKey()
    CallObjectIdUsingMethod "Focus"
    CDP_.DispatchKeyEvent "keyDown", 0, 13
    CDP_.DispatchKeyEvent "keyUp", 0, 13
End Sub

' 現在ページで遷移が進行中かを判定する
Private Function IsPageChanging() As Boolean
    CDP_.PageEnable '一回無害なメソッドを投げて、イベント発生状況更新

    If events_.IsCurrentTargetScheduledNavigation _
    Or events_.IsCurrentTargetStartedLoading Then
        IsPageChanging = True
    Else
        IsPageChanging = False
    End If
End Function

' 現在ページの遷移完了を待機する
Private Sub WaitPageChangeToComplete()

    ' TargetNavigated が返らないケースに備えて反復回数へ上限を設ける
    Dim failSafeCounter As Long
    Dim isEndChange As Boolean
    Do
        isEndChange = events_.IsCurrentTargetNavigated
'                                Or　chrome_.CDPEvents.IsCurrentTargetClearedScheduledNavigation
        If isEndChange Then Exit Do

        CDP_.RuntimeEnable
        Sleep 1000
        failSafeCounter = failSafeCounter + 1
    Loop While (isEndChange = False And failSafeCounter <= 4)

    events_.IsCurrentTargetScheduledNavigation = False
    events_.IsCurrentTargetStartedLoading = False
    events_.IsCurrentTargetClearedScheduledNavigation = False
    events_.IsCurrentTargetNavigated = False
End Sub

' IsolatedFrame で遷移が進行中かを判定する
Private Function IsIsolatedFrameChanging() As Boolean
    CDP_.DOMEnable '一回無害なメソッド(frameなのでDOM系)を投げて、イベント発生状況更新

    If events_.IsIsolatedFrameScheduledNavigation _
    Or events_.IsIsolatedFrameStartedLoading Then
        IsIsolatedFrameChanging = True
    Else
        IsIsolatedFrameChanging = False
    End If
End Function

' IsolatedFrame の遷移完了を待機する
Private Sub WaitIsolatedFrameChangeToComplete()

    ' FrameNavigated が返らないケースに備えて反復回数へ上限を設ける
    Dim failSafeCounter As Long
    Dim isEndChange As Boolean
    Do
        isEndChange = events_.IsIsolatedFrameNavigated
'                               Or chrome_.CDPEvents.IsIsolatedFrameClearedScheduledNavigation
        If isEndChange Then Exit Do

        CDP_.DOMEnable
        CDP_.RuntimeEnable
        Sleep 1000
        failSafeCounter = failSafeCounter + 1
    Loop While (isEndChange = False And failSafeCounter <= 4)

    events_.IsIsolatedFrameScheduledNavigation = False
    events_.IsIsolatedFrameStartedLoading = False
    events_.IsIsolatedFrameClearedScheduledNavigation = False
    events_.IsIsolatedFrameNavigated = False
End Sub

' value が期待値へ反映されるまで待機する
Private Sub WaitUntilValueEquals(expected As String, Optional timeoutSec As Long = 5)
    Dim startT As Single: startT = Timer

    Do
        Dim res As String
        res = CallFunction("function(){return this.value;}", "")

        Dim cur As String
        cur = json_.GetValue(res, "result", "result", "value")

        If cur = expected Then Exit Sub

        Sleep 50
        DoEvents
    Loop While (Timer - startT) < timeoutSec

    Helper_.DisplayErrorAndEnd _
        "＜SetValue/SetText＞" & vbCrLf & _
        "valueが反映されませんでした。"
End Sub

' Public API
' 要素の value 属性へ文字列を設定する
' 過去バージョン互換のため SetValue と同等処理を残す
Private Sub IWebElement_SetText(text As String, Optional timeoutSec As Long = 5)
    CallObjectIdUsingMethod "Focus"
    Dim encodingText As String: encodingText = json_.EncodeURL(text)
    CallFunction "function(value){value = decodeURI(value);this.value = value;}", encodingText

    CallFunction Helper_.GetDispatchInputEventCode, encodingText
    CallFunction Helper_.GetDispatchChangeEventCode, encodingText

    WaitUntilValueEquals text, timeoutSec
End Sub

' Public API
' 要素の value 属性へ文字列を設定する
Private Sub IWebElement_SetValue(value As String, Optional timeoutSec As Long = 5)
    CallObjectIdUsingMethod "Focus"
    Dim encodingText As String: encodingText = json_.EncodeURL(value)
    CallFunction "function(value){value = decodeURI(value);this.value = value;}", encodingText

    CallFunction Helper_.GetDispatchInputEventCode, encodingText
    CallFunction Helper_.GetDispatchChangeEventCode, encodingText

    WaitUntilValueEquals value, timeoutSec
End Sub

' Public API
' 要素の textContent へ文字列を設定する
Private Sub IWebElement_SetTextContent(text As String)
    CallObjectIdUsingMethod "Focus"
    Dim encodingText As String: encodingText = json_.EncodeURL(text)
    CallFunction "function(value){value = decodeURI(value);this.textContent  = value;}", encodingText

    CallFunction Helper_.GetDispatchInputEventCode, encodingText
    CallFunction Helper_.GetDispatchChangeEventCode, encodingText
End Sub

' Public API
' 要素の value 属性へ PropertySetter 経由で文字列を設定する
Private Function IWebElement_SetValueViaPropertySetter(value As String) As String
    FocusAndEnterKey
    Dim encodingText As String: encodingText = json_.EncodeURL(value)
    CallFunction "function(value){value = decodeURI(value);Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(this, value);}", encodingText

    CallFunction Helper_.GetDispatchInputEventCode, ""
    CallFunction Helper_.GetDispatchChangeEventCode, ""
End Function

' Public API
' SelectBox の選択肢を文字列で選択する
Private Sub IWebElement_SelectItemInSelectBoxByText(text As String)
    Dim encodingText As String: encodingText = json_.EncodeURL(text)

    Dim res As String: res = CallFunction(Helper_.GetSelectItemCode, encodingText)

    Dim IsSelect As Boolean: IsSelect = json_.GetValue(res, "result", "result", "value")
    If IsSelect = False Then
        Err.Raise vbObjectError + 2110, "b0x0_WebElement.SelectItemInSelectBoxByText", _
                    "<SelectItemInSelectBoxByText>" & vbCrLf & _
                    "「" & text & "」は選択肢の中に存在しません。"
    End If

'    CallFunction "function(value){this.onchange();}", encodingText
    CallFunction Helper_.GetDispatchChangeEventCode, encodingText
End Sub

' Public API
Private Function IWebElement_FindElementByName(name As String) As IWebElement
    Dim funcCode As String, res As String
    funcCode = "function(name){return  this.getElementsByName(name)[0];}"
    res = CallFunction(funcCode, name)

    Dim objectId As String
    objectId = json_.GetValue(res, "result", "result", "objectId")
    If objectId = "" Then GoTo quit

    Dim findInfoInElem As Collection
    Set findInfoInElem = Helper_.MakeFindInfo("name", name)

    Set IWebElement_FindElementByName = _
        CDP_.MakeElementObject(objectId, findInfoInElem, basicInfo_, events_, Helper_)
Exit Function

quit:
'Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & name
'Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("Name", name)
End Function

' Public API
Private Function IWebElement_FindElementsByTag(tagName As String) As IWebElements
    Dim elemsCount As Long
    elemsCount = GetElementsCount(tagName)
    If elemsCount = 0 Then GoTo quit

    Dim elements As Collection: Set elements = New Collection
    Dim index As Long
    For index = 0 To elemsCount - 1
        elements.Add CreateElementsItemByIndex(tagName, index)
    Next

    Dim elems As c0x0_WebElements: Set elems = New c0x0_WebElements
    elems.Init elements

    Set IWebElement_FindElementsByTag = elems
Exit Function
quit:
'Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & tagName
'Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("Tag", tagName, "○○")
End Function

' 子要素から CSS セレクタ指定で 1 件取得する
Private Function IWebElement_FindElementByCss(cssSelector As String) As IWebElement
    Dim encodingCssSelector As String
    encodingCssSelector = json_.EncodeURL(cssSelector)

    Dim funcCode As String
    Dim res As String

    funcCode = "function(cssSelector){return this.querySelector(decodeURIComponent(cssSelector));}"
    res = CallFunction(funcCode, encodingCssSelector)

    Dim objectId As String
    objectId = json_.GetValue(res, "result", "result", "objectId")
    If objectId = "" Then GoTo quit

    Dim findInfoInElem As Collection
    Set findInfoInElem = Helper_.MakeFindInfo("css", encodingCssSelector)

    Set IWebElement_FindElementByCss = _
        CDP_.MakeElementObject(objectId, findInfoInElem, basicInfo_, events_, Helper_)
    Exit Function

quit:
    Helper_.DisplayErrorAndEnd _
        "<FindElementByCss>" & vbCrLf & _
        "指定された以下CSSセレクタの要素は存在しませんでした。" & vbCrLf & _
        cssSelector & vbCrLf & _
        "(イミディエイトウィンドウで確認可能)"
End Function

' 子要素から CSS セレクタ指定で複数取得する
Private Function IWebElement_FindElementsByCss(cssSelector As String) As IWebElements
    Dim encodingCssSelector As String
    encodingCssSelector = json_.EncodeURL(cssSelector)

    Dim funcCode As String
    Dim res As String
    Dim elemsCount As Long

    funcCode = "function(cssSelector){return this.querySelectorAll(decodeURIComponent(cssSelector)).length;}"
    res = CallFunction(funcCode, encodingCssSelector)
    elemsCount = CLng(json_.GetValue(res, "result", "result", "value"))
    If elemsCount = 0 Then GoTo quit

    Dim elements As Collection
    Set elements = New Collection

    Dim index As Long
    Dim objectId As String
    Dim findInfoInElem As Collection

    For index = 0 To elemsCount - 1
        funcCode = "function(cssSelector){return this.querySelectorAll(decodeURIComponent(cssSelector))[" & index & "];}"
        res = CallFunction(funcCode, encodingCssSelector)

        objectId = json_.GetValue(res, "result", "result", "objectId")
        If objectId = "" Then
            Err.Raise vbObjectError + 2112, "b0x0_WebElement.FindElementsByCss", _
                "CSSセレクタ結果の要素取得に失敗しました。index=" & CStr(index)
        End If

        Set findInfoInElem = Helper_.MakeFindInfo("css", encodingCssSelector, index)
        elements.Add CDP_.MakeElementObject(objectId, findInfoInElem, basicInfo_, events_, Helper_)
    Next

    Dim elems As c0x0_WebElements
    Set elems = New c0x0_WebElements
    elems.Init elements

    Set IWebElement_FindElementsByCss = elems
    Exit Function

quit:
    Helper_.DisplayErrorAndEnd _
        "<FindElementsByCss>" & vbCrLf & _
        "指定された以下CSSセレクタの要素は存在しませんでした。" & vbCrLf & _
        cssSelector & vbCrLf & _
        "(イミディエイトウィンドウで確認可能)"
End Function

' 子要素からテキスト指定で 1 件取得する
Private Function IWebElement_FindElementByText(text As String, Optional partialMatch As Boolean = False) As IWebElement
    Dim encodingText As String
    encodingText = json_.EncodeURL(text)

    Dim partial As String
    If partialMatch Then
        partial = "true"
    Else
        partial = "false"
    End If

    Dim funcCode As String
    Dim res As String

    funcCode = "function(text){" & _
               "var target=decodeURIComponent(text);" & _
               "var partial=" & partial & ";" & _
               "var elems=this.getElementsByTagName('*');" & _
               "var i,cur;" & _
               "for(i=0;i<elems.length;i++){" & _
               "cur=(elems[i].textContent||'').replace(/\s+/g,' ').trim();" & _
               "if((partial && cur.indexOf(target) >= 0) || (!partial && cur === target)){" & _
               "return elems[i];" & _
               "}" & _
               "}" & _
               "return null;" & _
               "}"

    res = CallFunction(funcCode, encodingText)

    Dim objectId As String
    objectId = json_.GetValue(res, "result", "result", "objectId")
    If objectId = "" Then GoTo quit

    Dim findKind As String
    If partialMatch Then
        findKind = "text_partial"
    Else
        findKind = "text"
    End If

    Dim findInfoInElem As Collection
    Set findInfoInElem = Helper_.MakeFindInfo(findKind, encodingText)

    Set IWebElement_FindElementByText = _
        CDP_.MakeElementObject(objectId, findInfoInElem, basicInfo_, events_, Helper_)
    Exit Function

quit:
    Helper_.DisplayErrorAndEnd _
        "<FindElementByText>" & vbCrLf & _
        "指定された以下テキストの要素は存在しませんでした。" & vbCrLf & _
        text & vbCrLf & _
        "(イミディエイトウィンドウで確認可能)"
End Function

' 子要素からテキスト指定で複数取得する
Private Function IWebElement_FindElementsByText(text As String, Optional partialMatch As Boolean = False) As IWebElements
    Dim encodingText As String
    encodingText = json_.EncodeURL(text)

    Dim partial As String
    If partialMatch Then
        partial = "true"
    Else
        partial = "false"
    End If

    Dim funcCode As String
    Dim res As String
    Dim elemsCount As Long

    funcCode = "function(text){" & _
               "var target=decodeURIComponent(text);" & _
               "var partial=" & partial & ";" & _
               "var elems=this.getElementsByTagName('*');" & _
               "var i,cur,cnt=0;" & _
               "for(i=0;i<elems.length;i++){" & _
               "cur=(elems[i].textContent||'').replace(/\s+/g,' ').trim();" & _
               "if((partial && cur.indexOf(target) >= 0) || (!partial && cur === target)){" & _
               "cnt++;" & _
               "}" & _
               "}" & _
               "return cnt;" & _
               "}"

    res = CallFunction(funcCode, encodingText)
    elemsCount = CLng(json_.GetValue(res, "result", "result", "value"))
    If elemsCount = 0 Then GoTo quit

    Dim elements As Collection
    Set elements = New Collection

    Dim index As Long
    Dim objectId As String
    Dim findInfoInElem As Collection
    Dim findKind As String

    If partialMatch Then
        findKind = "text_partial"
    Else
        findKind = "text"
    End If

    For index = 0 To elemsCount - 1
        funcCode = "function(text){" & _
                   "var target=decodeURIComponent(text);" & _
                   "var partial=" & partial & ";" & _
                   "var elems=this.getElementsByTagName('*');" & _
                   "var i,cur,cnt=0;" & _
                   "for(i=0;i<elems.length;i++){" & _
                   "cur=(elems[i].textContent||'').replace(/\s+/g,' ').trim();" & _
                   "if((partial && cur.indexOf(target) >= 0) || (!partial && cur === target)){" & _
                   "if(cnt === " & index & "){" & _
                   "return elems[i];" & _
                   "}" & _
                   "cnt++;" & _
                   "}" & _
                   "}" & _
                   "return null;" & _
                   "}"

        res = CallFunction(funcCode, encodingText)
        objectId = json_.GetValue(res, "result", "result", "objectId")
        If objectId = "" Then
            Err.Raise vbObjectError + 2113, "b0x0_WebElement.FindElementsByText", _
                "Text locator 結果の要素取得に失敗しました。index=" & CStr(index)
        End If

        Set findInfoInElem = Helper_.MakeFindInfo(findKind, encodingText, index)
        elements.Add CDP_.MakeElementObject(objectId, findInfoInElem, basicInfo_, events_, Helper_)
    Next

    Dim elems As c0x0_WebElements
    Set elems = New c0x0_WebElements
    elems.Init elements

    Set IWebElement_FindElementsByText = elems
    Exit Function

quit:
    Helper_.DisplayErrorAndEnd _
        "<FindElementsByText>" & vbCrLf & _
        "指定された以下テキストの要素は存在しませんでした。" & vbCrLf & _
        text & vbCrLf & _
        "(イミディエイトウィンドウで確認可能)"
End Function

' Shadow DOM を含めて CSS セレクタ指定で 1 件取得する
Private Function IWebElement_FindElementByCssDeep(cssSelector As String) As IWebElement
    Dim encodingCssSelector As String
    encodingCssSelector = json_.EncodeURL(cssSelector)

    Dim funcCode As String
    Dim res As String

    funcCode = "function(cssSelector){" & _
               "var selector=decodeURIComponent(cssSelector);" & _
               "function findIn(root, selector){" & _
               "if(!root){return null;}" & _
               "if(root.querySelector){" & _
               "var hit=root.querySelector(selector);" & _
               "if(hit){return hit;}" & _
               "}" & _
               "if(!root.querySelectorAll){return null;}" & _
               "var nodes=root.querySelectorAll('*');" & _
               "for(var i=0;i<nodes.length;i++){" & _
               "if(nodes[i].shadowRoot){" & _
               "var subHit=findIn(nodes[i].shadowRoot, selector);" & _
               "if(subHit){return subHit;}" & _
               "}" & _
               "}" & _
               "return null;" & _
               "}" & _
               "if(this.shadowRoot){" & _
               "var selfShadowHit=findIn(this.shadowRoot, selector);" & _
               "if(selfShadowHit){return selfShadowHit;}" & _
               "}" & _
               "return findIn(this, selector);" & _
               "}"

    res = CallFunction(funcCode, encodingCssSelector)

    Dim objectId As String
    objectId = json_.GetValue(res, "result", "result", "objectId")
    If objectId = "" Then GoTo quit

    Dim findInfoInElem As Collection
    Set findInfoInElem = Helper_.MakeFindInfo("css_deep", encodingCssSelector)

    Set IWebElement_FindElementByCssDeep = _
        CDP_.MakeElementObject(objectId, findInfoInElem, basicInfo_, events_, Helper_)
    Exit Function

quit:
    Helper_.DisplayErrorAndEnd _
        "<FindElementByCssDeep>" & vbCrLf & _
        "指定された以下CSSセレクタ(Deep)の要素は存在しませんでした。" & vbCrLf & _
        cssSelector & vbCrLf & _
        "(イミディエイトウィンドウで確認可能)"
End Function

' Shadow DOM を含めて CSS セレクタ指定で複数取得する
Private Function IWebElement_FindElementsByCssDeep(cssSelector As String) As IWebElements
    Dim encodingCssSelector As String
    encodingCssSelector = json_.EncodeURL(cssSelector)

    Dim funcCode As String
    Dim res As String
    Dim elemsCount As Long

    funcCode = "function(cssSelector){" & _
               "var selector=decodeURIComponent(cssSelector);" & _
               "function collectIn(root, selector, out){" & _
               "if(!root){return;}" & _
               "if(root.querySelectorAll){" & _
               "var hits=root.querySelectorAll(selector);" & _
               "for(var i=0;i<hits.length;i++){" & _
               "out.push(hits[i]);" & _
               "}" & _
               "var nodes=root.querySelectorAll('*');" & _
               "for(var j=0;j<nodes.length;j++){" & _
               "if(nodes[j].shadowRoot){" & _
               "collectIn(nodes[j].shadowRoot, selector, out);" & _
               "}" & _
               "}" & _
               "}" & _
               "}" & _
               "var result=[];" & _
               "if(this.shadowRoot){" & _
               "collectIn(this.shadowRoot, selector, result);" & _
               "}" & _
               "collectIn(this, selector, result);" & _
               "return result.length;" & _
               "}"

    res = CallFunction(funcCode, encodingCssSelector)
    elemsCount = CLng(json_.GetValue(res, "result", "result", "value"))
    If elemsCount = 0 Then GoTo quit

    Dim elements As Collection
    Set elements = New Collection

    Dim index As Long
    Dim objectId As String
    Dim findInfoInElem As Collection

    For index = 0 To elemsCount - 1
        funcCode = "function(cssSelector){" & _
                   "var selector=decodeURIComponent(cssSelector);" & _
                   "function collectIn(root, selector, out){" & _
                   "if(!root){return;}" & _
                   "if(root.querySelectorAll){" & _
                   "var hits=root.querySelectorAll(selector);" & _
                   "for(var i=0;i<hits.length;i++){" & _
                   "out.push(hits[i]);" & _
                   "}" & _
                   "var nodes=root.querySelectorAll('*');" & _
                   "for(var j=0;j<nodes.length;j++){" & _
                   "if(nodes[j].shadowRoot){" & _
                   "collectIn(nodes[j].shadowRoot, selector, out);" & _
                   "}" & _
                   "}" & _
                   "}" & _
                   "}" & _
                   "var result=[];" & _
                   "if(this.shadowRoot){" & _
                   "collectIn(this.shadowRoot, selector, result);" & _
                   "}" & _
                   "collectIn(this, selector, result);" & _
                   "return (result.length > " & index & ") ? result[" & index & "] : null;" & _
                   "}"

        res = CallFunction(funcCode, encodingCssSelector)
        objectId = json_.GetValue(res, "result", "result", "objectId")
        If objectId = "" Then
            Err.Raise vbObjectError + 2114, "b0x0_WebElement.FindElementsByCssDeep", _
                "CSSセレクタ(Deep)結果の要素取得に失敗しました。index=" & CStr(index)
        End If

        Set findInfoInElem = Helper_.MakeFindInfo("css_deep", encodingCssSelector, index)
        elements.Add CDP_.MakeElementObject(objectId, findInfoInElem, basicInfo_, events_, Helper_)
    Next

    Dim elems As c0x0_WebElements
    Set elems = New c0x0_WebElements
    elems.Init elements

    Set IWebElement_FindElementsByCssDeep = elems
    Exit Function

quit:
    Helper_.DisplayErrorAndEnd _
        "<FindElementsByCssDeep>" & vbCrLf & _
        "指定された以下CSSセレクタ(Deep)の要素は存在しませんでした。" & vbCrLf & _
        cssSelector & vbCrLf & _
        "(イミディエイトウィンドウで確認可能)"
End Function

' Shadow DOM を含めてテキスト指定で 1 件取得する
Private Function IWebElement_FindElementByTextDeep(text As String, Optional partialMatch As Boolean = False) As IWebElement
    Dim encodingText As String
    encodingText = json_.EncodeURL(text)

    Dim partial As String
    partial = IIf(partialMatch, "true", "false")

    Dim jsCore As String
    jsCore = "var target=decodeURIComponent(text);" & _
             "var partial=" & partial & ";" & _
             "function normalizeText(s){return (s||'').replace(/\s+/g,' ').trim();}" & _
             "function findIn(root){" & _
             "if(!root||!root.querySelectorAll){return null;}" & _
             "var nodes=root.querySelectorAll('*');" & _
             "for(var i=0;i<nodes.length;i++){" & _
             "var cur=normalizeText(nodes[i].textContent);" & _
             "if((partial && cur.indexOf(target) >= 0) || (!partial && cur === target)){return nodes[i];}" & _
             "if(nodes[i].shadowRoot){var subHit=findIn(nodes[i].shadowRoot);if(subHit){return subHit;}}" & _
             "}" & _
             "return null;" & _
             "}" & _
             "if(this.shadowRoot){var selfShadowHit=findIn(this.shadowRoot);if(selfShadowHit){return selfShadowHit;}}" & _
             "return findIn(this);"

    Dim funcCode As String
    funcCode = "function(text){" & jsCore & "}"

    Dim res As String
    res = CallFunction(funcCode, encodingText)

    Dim objectId As String
    objectId = json_.GetValue(res, "result", "result", "objectId")
    If objectId = "" Then GoTo quit

    Dim findKind As String
    findKind = IIf(partialMatch, "text_deep_partial", "text_deep")

    Dim findInfoInElem As Collection
    Set findInfoInElem = Helper_.MakeFindInfo(findKind, encodingText)

    Set IWebElement_FindElementByTextDeep = _
        CDP_.MakeElementObject(objectId, findInfoInElem, basicInfo_, events_, Helper_)
    Exit Function

quit:
    Helper_.DisplayErrorAndEnd _
        "<FindElementByTextDeep>" & vbCrLf & _
        "指定された以下テキスト(Deep)の要素は存在しませんでした。" & vbCrLf & _
        text & vbCrLf & _
        "(イミディエイトウィンドウで確認可能)"
End Function

' Shadow DOM を含めてテキスト指定で複数取得する
Private Function IWebElement_FindElementsByTextDeep(text As String, Optional partialMatch As Boolean = False) As IWebElements
    Dim encodingText As String: encodingText = json_.EncodeURL(text)
    Dim partial As String: partial = IIf(partialMatch, "true", "false")

    Dim funcCode As String, res As String, elemsCount As Long

    funcCode = "function(text){var t=decodeURIComponent(text),p=" & partial & ";function n(s){return (s||'').replace(/\s+/g,' ').trim();}" & _
               "function c(r){if(!r||!r.querySelectorAll)return 0;var a=r.querySelectorAll('*'),i,x=0;" & _
               "for(i=0;i<a.length;i++){var v=n(a[i].textContent);" & _
               "if((p&&v.indexOf(t)>=0)||(!p&&v===t))x++;if(a[i].shadowRoot)x+=c(a[i].shadowRoot);}return x;}" & _
               "var k=0;if(this.shadowRoot)k+=c(this.shadowRoot);return k+c(this);}"

    res = CallFunction(funcCode, encodingText)
    elemsCount = CLng(json_.GetValue(res, "result", "result", "value"))
    If elemsCount = 0 Then GoTo quit

    Dim elements As New Collection
    Dim index As Long, objectId As String
    Dim findKind As String: findKind = IIf(partialMatch, "text_deep_partial", "text_deep")

    For index = 0 To elemsCount - 1
        funcCode = "function(text){var t=decodeURIComponent(text),p=" & partial & ";function n(s){return (s||'').replace(/\s+/g,' ').trim();}" & _
                   "function f(r,i){if(!r||!r.querySelectorAll)return{c:0,n:null};var a=r.querySelectorAll('*'),x=0,j;" & _
                   "for(j=0;j<a.length;j++){var v=n(a[j].textContent);" & _
                   "if((p&&v.indexOf(t)>=0)||(!p&&v===t)){if(x===i)return{c:x,n:a[j]};x++;}" & _
                   "if(a[j].shadowRoot){var s=f(a[j].shadowRoot,i-x);if(s.n)return s;x+=s.c;}}" & _
                   "return{c:x,n:null};}" & _
                   "var o=" & index & ";if(this.shadowRoot){var z=f(this.shadowRoot,o);if(z.n)return z.n;o-=z.c;}return f(this,o).n;}"

        res = CallFunction(funcCode, encodingText)
        objectId = json_.GetValue(res, "result", "result", "objectId")
        If objectId = "" Then Err.Raise vbObjectError + 2115, "FindElementsByTextDeep", "取得失敗 index=" & index

        elements.Add CDP_.MakeElementObject(objectId, Helper_.MakeFindInfo(findKind, encodingText, index), basicInfo_, events_, Helper_)
    Next

    Dim elems As New c0x0_WebElements
    elems.Init elements
    Set IWebElement_FindElementsByTextDeep = elems
    Exit Function

quit:
    Helper_.DisplayErrorAndEnd "<FindElementsByTextDeep>" & vbCrLf & _
        "指定テキスト(Deep)の要素は存在しませんでした。" & vbCrLf & text
End Function

' 子要素のタグ名一致件数を取得する
Private Function GetElementsCount(tagName As String) As Long
    Dim funcCode As String, res As String
    funcCode = "function(tagName){return  this.getElementsByTagName('" & tagName & "').length;}"
    res = CallFunction(funcCode, tagName)
    GetElementsCount = json_.GetValue(res, "result", "result", "value")
End Function

' タグ名検索結果から index 指定要素を生成する
Private Function CreateElementsItemByIndex(tagName As String, index As Long) As IWebElement
    Dim funcCode As String, res As String
    funcCode = "function(tagName){return  this.getElementsByTagName(tagName)[" & index & "];}"
    res = CallFunction(funcCode, tagName)

    Dim objectId As String
    objectId = json_.GetValue(res, "result", "result", "objectId")
    Dim findInfoInElem As Collection: Set findInfoInElem = Helper_.MakeFindInfo("tag", tagName, index)

    Set CreateElementsItemByIndex = _
        CDP_.MakeElementObject(objectId, findInfoInElem, basicInfo_, events_, Helper_)
End Function

' Public API
' 要素の textContent を取得する
Private Function IWebElement_GetTextContent() As String
    Dim funcCode As String, res As String
    funcCode = "function(){return this.textContent.replace(/\\s/g, ' ');}"
    res = CallFunction(funcCode, "")

    IWebElement_GetTextContent = json_.GetValue(res, "result", "result", "value")
End Function

' Public API
' セレクトボックスの選択中 value を取得する
Private Function IWebElement_GetSelectedValue() As String
    Dim funcCode As String, res As String
    funcCode = "function(){return this.value.replace(/\\s/g, ' ');}"
    res = CallFunction(funcCode, "")

    IWebElement_GetSelectedValue = json_.GetValue(res, "result", "result", "value")
End Function

' Public API
' セレクトボックスの選択中 textContent を取得する
Private Function IWebElement_GetSelectedTextContent() As String
    Dim funcCode As String, res As String
    funcCode = Helper_.GetSelectedTextContentCode
    res = CallFunction(funcCode, "")

    IWebElement_GetSelectedTextContent = json_.GetValue(res, "result", "result", "value")
End Function

' Public API
' type=file の input 要素へファイルを設定する
Private Sub IWebElement_SetFileToFileSelectInputButton(filePath As String)
'    filePath = Replace(filePath, "\", "\\")
'    filePath = Json_.EncodeURL(filePath)
'    filePath = Replace(filePath,    "%5C", "\")
'    chrome_.SetFileInputFiles objectId_,
'CDPMethodで設定しようとしても、ファイルパスに日本語が含まれる場合、
'エンコードして送信はできるが、クローム側でデコードしてくれない為、没
    Dim res As String: res = CallObjectIdUsingMethod("GetBoxModel")
    Dim contentBox As Object: Set contentBox = GetContentBox(res)

    CDP_.DispatchMouseEvent "mouseWheel", 0, 0, "none", 0, 0, 0, CLng(contentBox("CenterY")) - 50

    IWebElement_DispatchCDPMouseEvent CDP_Mouse.LeftClick_XY
    Sleep 1000
    ControlFileSelectDialog filePath
    Sleep 1000
End Sub
' ファイル選択ダイアログを操作する
Private Sub ControlFileSelectDialog(filePath As String)
    Dim DialogHwnd As LongPtr: DialogHwnd = FindWindow(vbNullString, "開く")
    Dim InputHwnd As Long
    InputHwnd = FindWindowEx(DialogHwnd, 0&, "ComboBoxEx32", "")
    InputHwnd = FindWindowEx(InputHwnd, 0&, "ComboBox", "")
    InputHwnd = FindWindowEx(InputHwnd, 0&, "Edit", "")
    Call SendMessage(InputHwnd, &HC, 0, ByVal filePath)

    Dim ButtonHwnd As Long
    ButtonHwnd = FindWindowEx(DialogHwnd, 0&, "Button", "開く(&O)")
    Call SendMessage(ButtonHwnd, &HF5, 0, 0&)
End Sub

' Public API
' クリック後に制御を返さないアラート出現時の処理を行う
Private Sub IWebElement_ClickAndThenAlertDialogErase()
    CallObjectIdUsingMethod "Focus"
    Dim res As String
    res = CallFunction("function(){this.click();}", "")
    Sleep 1000

    CDP_.HandleAlertDialog
End Sub

' Public API
' 要素の属性値を取得する
Private Function IWebElement_GetAttribute(attributeName As String) As String
    Dim res As String
    res = CallFunction("function(name){return this.getAttribute(name);}", attributeName)
    If InStr(res, """value"":null") > 0 Then
        IWebElement_GetAttribute = ""
    Else
        IWebElement_GetAttribute = json_.GetValue(res, "result", "result", "value")
    End If
End Function

' Public API
' 要素の class 属性へ値を設定する
Private Sub IWebElement_SetClassName(className As String)
    Dim funcCode As String
    Dim res As String
    Dim encodingClassName As String

    encodingClassName = json_.EncodeURL(className)
    funcCode = "function(value){this.className = decodeURI(value);}"
    res = CallFunction(funcCode, encodingClassName)
    res = json_.GetValue(res, "result", "result", "value")
End Sub

Private Function IWebElement_SetAttribute(attributeName As String, value As Variant) As String
    Dim funcCode As String, res As String
    ' 引数を 1 つしか渡せないためオブジェクトとしてまとめて渡す
    funcCode = "function(value){let a = value.arg1;let b = value.arg2;console.log(a,b); this.setAttribute( a , b);}"
    ' 呼び出し時はプロパティ名と値を対で渡す
    res = CallFunction_ArgMultiple(funcCode, "arg1", attributeName, "arg2", value)
    IWebElement_SetAttribute = json_.GetValue(res, "result", "result")

    ' res = CallFunction(funcCode, "{""arg1"", " & attributeName & ", ""arg2""," & value & "}")
    ' JS オブジェクト文字列を直接渡す形式なら CallFunction_ArgMultiple は不要となる
    ' 既存実装互換のため現状のままとする
End Function

' Public API
Private Sub IWebElement_DispatchJSChangeEvent()
    CallFunction Helper_.GetDispatchChangeEventCode, ""
End Sub

' Public API
Private Sub IWebElement_DispatchJSInputEvent()
    CallFunction Helper_.GetDispatchInputEventCode, ""
End Sub

' Public API
Private Sub IWebElement_DispatchJSFocusEvent()
    CallFunction Helper_.GetDispatchFocusEventCode, ""
End Sub

' Public API
Private Sub IWebElement_DispatchJSMouseEvent(action As JS_Mouse)
    Dim jsCode As String
    jsCode = Helper_.GetDispatchMouseEventCode(action)
    CallFunction jsCode, ""
End Sub

' 要素が表示状態になるまで待機する
Private Sub IWebElement_WaitUntilVisible(Optional timeoutSec As Long = 5)
    Dim startT As Single: startT = Timer

    Do
        Dim res As String
        res = CallFunction( _
            "function(){ " & _
            "  const s = window.getComputedStyle(this);" & _
            "  const t = (this.getAttribute && this.getAttribute('type')) ? this.getAttribute('type') : '';" & _
            "  const ok = " & _
            "    t !== 'hidden' && " & _
            "    s.display !== 'none' && " & _
            "    s.visibility !== 'hidden' && " & _
            "    s.opacity !== '0' && " & _
            "    (this.offsetWidth > 0 || this.offsetHeight > 0 || this.getClientRects().length > 0);" & _
            "  return ok;" & _
            "}", _
            "" _
        )

        Dim ok As String
        ok = json_.GetValue(res, "result", "result", "value")
        ok = Replace(ok, """", "")

        If LCase$(ok) = "true" Then Exit Sub

        Sleep 50
        DoEvents
    Loop While (Timer - startT) < timeoutSec

    Helper_.DisplayErrorAndEnd _
        "＜WaitUntilVisible＞" & vbCrLf & _
        timeoutSec & "秒経過しても要素が表示状態になりませんでした。"
End Sub

' 要素が有効状態になるまで待機する
Private Sub IWebElement_WaitUntilEnabled(Optional timeoutSec As Long = 5)
    Dim startT As Single: startT = Timer

    Do
        Dim res As String
        res = CallFunction( _
            "function(){ " & _
            "  const s = window.getComputedStyle(this);" & _
            "  const aria = (this.getAttribute && this.getAttribute('aria-disabled')) ? this.getAttribute('aria-disabled') : '';" & _
            "  const disProp = (typeof this.disabled === 'boolean') ? this.disabled : null;" & _
            "  const disAttr = (this.hasAttribute && this.hasAttribute('disabled')) ? this.hasAttribute('disabled') : false;" & _
            "  const ok = (disProp === false) && (disAttr === false) && (aria !== 'true');" & _
            "  return ok;" & _
            "}", _
            "" _
        )

        Dim ok As String
        ok = json_.GetValue(res, "result", "result", "value")
        ok = Replace(ok, """", "")

        If LCase$(ok) = "true" Then Exit Sub

        Sleep 50
        DoEvents
    Loop While (Timer - startT) < timeoutSec

    Helper_.DisplayErrorAndEnd _
        "＜WaitUntilEnabled＞" & vbCrLf & _
        timeoutSec & "秒経過しても要素が有効状態になりませんでした。"
End Sub

' 要素が非表示または消滅するまで待機する
Private Sub IWebElement_WaitUntilHidden(Optional timeoutSec As Long = 5)
    Dim startT As Single: startT = Timer

    Do
        Dim res As String
        res = CallFunction( _
            "function(){ " & _
            "  const s = window.getComputedStyle(this);" & _
            "  const connected = (this.isConnected !== undefined) ? this.isConnected : true;" & _
            "  const rects = (this.getClientRects) ? this.getClientRects().length : 0;" & _
            "  const ow = (this.offsetWidth !== undefined) ? this.offsetWidth : 0;" & _
            "  const oh = (this.offsetHeight !== undefined) ? this.offsetHeight : 0;" & _
            "  const hidden = " & _
            "    (!connected) || " & _
            "    (s.display === 'none') || " & _
            "    (s.visibility === 'hidden') || " & _
            "    (s.opacity === '0') || " & _
            "    (ow === 0 && oh === 0 && rects === 0);" & _
            "  console.log('[WaitUntilHidden]', {id:this.id, connected:connected, display:s.display, visibility:s.visibility, opacity:s.opacity, ow:ow, oh:oh, rects:rects, ok:hidden});" & _
            "  return hidden;" & _
            "}", _
            "" _
        )

        Dim ok As String
        ok = json_.GetValue(res, "result", "result", "value")
        ok = Replace(ok, """", "")

        If LCase$(ok) = "true" Then Exit Sub

        Sleep 50
        DoEvents
    Loop While (Timer - startT) < timeoutSec

    Helper_.DisplayErrorAndEnd _
        "＜WaitUntilHidden＞" & vbCrLf & _
        timeoutSec & "秒経過しても要素が非表示/消滅しませんでした。"
End Sub
' keyCode は押下キーのコード番号を表す
' altKey / shiftKey / ctrlKey は同時押下状態を表す
Private Sub IWebElement_DispatchJSKeyboardEvent(action As JS_Keyboard, KeyCode As key, Optional alt As Boolean = False, Optional shift As Boolean = False, Optional ctrl As Boolean = False)
    Dim jsCode As String
    jsCode = Helper_.GetDispatchKeyboardEventCode(action, KeyCode, alt, shift, ctrl)
    CallFunction jsCode, ""
End Sub

' 要素プロパティ値を取得する
Private Function IWebElement_GetProperty(PropertyName As String) As Variant
    Dim funcCode As String, res As String
    funcCode = "function(value){return this['" & PropertyName & "'];}"
    res = CallFunction(funcCode, "")
    Dim returnValue As Variant
    returnValue = json_.GetValue(res, "result", "result", "value")
    If IsNumeric(returnValue) Then
        IWebElement_GetProperty = CLng(returnValue)
    ElseIf IsDate(returnValue) Then
         IWebElement_GetProperty = CDate(returnValue)
    ElseIf returnValue = "true" Then
        IWebElement_GetProperty = True
    ElseIf returnValue = "false" Then
        IWebElement_GetProperty = False
    Else
        IWebElement_GetProperty = returnValue
    End If
End Function

' 要素プロパティ値を設定する
Private Function IWebElement_SetProperty(PropertyName As String, value As Variant) As String
    Dim funcCode As String, res As String
    funcCode = "function(value){return this['" & PropertyName & "'] = value ;}"
    res = CallFunction(funcCode, value)
    IWebElement_SetProperty = json_.GetValue(res, "result", "result")
End Function


