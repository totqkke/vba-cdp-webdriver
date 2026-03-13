VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a6_ExecuteHelperFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)

Private Const MOUSE_EVENT_NAME_NUM = 17
Private Const MOUSE_EVENT_NAMES = _
"mousedown,mousepress,mouseup,mouseenter,mouseleave, mousemove,mouseover,mouseout," & _
"click,dblclick," & _
"drag,dragstart,dragend,dragenter,dragexit,dragleave,dragover,drop"

Private Const KEYBOARD_EVENT_KEY_NUM = 51
' キーコード区分ごとの個数メモ
Private Const KEY_CODES = _
"9,13,16,17,18," & _
"27,32,33,34," & _
"37,38,39,40," & _
"44,46," & _
"48,49,50,51,52,53,54,55,56,57," & _
"65,66,67,68,69,70,71," & _
"72,73,74,75,76,77,78," & _
"79,80,81,82,83,84,85,86,87,88,89,90"

Private Const KEY_NAMES = _
"Tab,Enter,Shift,Control,Alt," & _
"Escape,Space,PageUp,PageDown," & _
"ArrowLeft,ArrowUp,ArrowRight,ArrowDown," & _
"PrintScreen,Delete," & _
"Digit0,Digit1,Digit2,Digit3,Digit4,Digit5,Digit6,Digit7,Digit8,Digit9," & _
"KeyA,KeyB,KeyC,KeyD,KeyE,KeyF,KeyG," & _
"KeyH,KeyI,KeyJ,KeyK,KeyL,KeyM,KeyN," & _
"KeyO,KeyP,KeyQ,KeyR,KeyS,KeyT,KeyU,KeyV,KeyW,KeyX,KeyY,KeyZ"

Private mouseEventNames_ As Object
Private keyboardEventNames_ As Object
Private keyNames_ As Object

' 内部辞書を初期化する
Private Sub Class_Initialize()
    Set mouseEventNames_ = CreateObject("Scripting.Dictionary")
    Dim mouseEventNamesString As Variant: mouseEventNamesString = Split(MOUSE_EVENT_NAMES, ",")
    Dim num As Long
    For num = 1 To MOUSE_EVENT_NAME_NUM
        mouseEventNames_.Add num, mouseEventNamesString(num - 1)
    Next

    Set keyboardEventNames_ = CreateObject("Scripting.Dictionary")
    keyboardEventNames_.Add 1, "keydown"
    keyboardEventNames_.Add 2, "keypress"
    keyboardEventNames_.Add 3, "keyup"

    Set keyNames_ = CreateObject("Scripting.Dictionary")
    Dim keyboardEventCodes As Variant: keyboardEventCodes = Split(KEY_CODES, ",")
    Dim KeyNames As Variant: KeyNames = Split(KEY_NAMES, ",")
    Dim i As Long, KeyCode As Long
    Dim KeyName As String
    For i = 0 To KEYBOARD_EVENT_KEY_NUM - 1
        KeyCode = keyboardEventCodes(i)
        KeyName = KeyNames(i)
        keyNames_.Add KeyCode, KeyName
    Next
End Sub

' WinAPI の Sleep を呼び出す
Public Sub SleepByWinAPI(ms As Long)
    Sleep ms
End Sub

' XPath 指定で要素取得用の JavaScript コードを返す
Public Function GetXpathCodes(xPath As String) As Variant
    Dim Codes(1) As String
    Codes(0) = "document.evaluate(decodeURIComponent('" & xPath & "'), document, null, " & _
                        "XPathResult.FIRST_ORDERED_NODE_TYPE, null ).singleNodeValue;"
    Codes(1) = "function(xpath){return  document.evaluate(decodeURIComponent(xpath), document, null, " & _
                            "XPathResult.FIRST_ORDERED_NODE_TYPE, null ).singleNodeValue;}"
    GetXpathCodes = Codes
End Function

' ID 指定で要素取得用の JavaScript コードを返す
Public Function GetIdCodes(Id As String) As Variant
    Dim Codes(1) As String
    Codes(0) = "document.getElementById('" & Id & "');"
    Codes(1) = "function(id){return  document.getElementById(id);}"
    GetIdCodes = Codes
End Function

' name 指定で要素取得用の JavaScript コードを返す
Public Function GetNameCodes(name As String) As Variant
    Dim Codes(1) As String
    Codes(0) = "document.getElementsByName('" & name & "')[0];"
    Codes(1) = "function(name){return  document.getElementsByName(name)[0];}"
    GetNameCodes = Codes
End Function

' タグ名指定で要素取得用の JavaScript コードを返す
Public Function GetTagCodes(tagName As String, index As Long) As Variant
    Dim Codes(1) As String
    Codes(0) = "document.getElementsByTagName('" & tagName & "')[" & index & "];"
    Codes(1) = "function(tagName){return  document.getElementsByTagName(tagName)[" & index & "];}"
    GetTagCodes = Codes
End Function

' クラス名指定で要素取得用の JavaScript コードを返す
Public Function GetClassNameCodes(className As String, index As Long) As Variant
    Dim Codes(1) As String
    Codes(0) = "document.getElementsByClassName('" & className & "')[" & index & "];"
    Codes(1) = "function(className){return document.getElementsByClassName(className)[" & index & "];}"
    GetClassNameCodes = Codes
End Function

' タグ名指定で要素数取得用の JavaScript コードを返す
Public Function GetTagCountCodes(tagName As String) As Variant
    Dim countCodes(1) As String
    countCodes(0) = "document.getElementsByTagName('" & tagName & "').length;"
    countCodes(1) = "function(tagName){return  document.getElementsByTagName('" & tagName & "').length;}"
    GetTagCountCodes = countCodes
End Function

' クラス名指定で要素数取得用の JavaScript コードを返す
Public Function GetClassNameCountCodes(className As String) As Variant
    Dim countCodes(1) As String
    countCodes(0) = "document.getElementsByClassName('" & className & "').length;"
    countCodes(1) = "function(className){return  document.getElementsByClassName('" & className & "').length;}"
    GetClassNameCountCodes = countCodes
End Function

' XPath 指定で要素数取得用の JavaScript コードを返す

Public Function GetXPathCountCodes(xPath As String)
    Dim Codes(1) As String
    Codes(0) = "result = document.evaluate(decodeURIComponent('" & xPath & "'), document, null, " & _
                        "XPathResult.ORDERED_NODE_ITERATOR_TYPE, null );" & _
                     "ary = [];" & _
                     "thisNode = result.iterateNext();" & _
                     "while (thisNode) {" & _
                         "ary.push(thisNode);" & _
                         "thisNode = result.iterateNext();" & _
                     "}" & _
                     "ary.length;"

    Codes(1) = "function(xpath){" & _
                            "result = document.evaluate(decodeURIComponent(xpath), document, null, " & _
                                "XPathResult.ORDERED_NODE_ITERATOR_TYPE, null );" & _
                            "ary = [];" & _
                            "thisNode = result.iterateNext();" & _
                            "while (thisNode) {" & _
                                "ary.push(thisNode);" & _
                                "thisNode = result.iterateNext();" & _
                            "}" & _
                            "return ary.length;" & _
                        "}"
    GetXPathCountCodes = Codes
End Function

' XPath 指定で index 要素取得用の JavaScript コードを返す

Public Function GetXPathResultCodes(xPath As String, index As Long)
    Dim Codes(1) As String
    Codes(0) = "result = document.evaluate(decodeURIComponent('" & xPath & "'), document, null, " & _
                        "XPathResult.ORDERED_NODE_ITERATOR_TYPE, null );" & _
                     "count = 0;" & _
                     "thisNode = result.iterateNext();" & _
                     "while (thisNode) {" & _
                         "if(count === " & index & "){" & _
                             "break;" & _
                         "}" & _
                         "thisNode = result.iterateNext();" & _
                         "count ++;" & _
                     "}" & _
                     "thisNode;"

    Codes(1) = "function(xpath){" & _
                            "result = document.evaluate(decodeURIComponent(xpath), document, null, " & _
                                "XPathResult.ORDERED_NODE_ITERATOR_TYPE, null );" & _
                            "count = 0;" & _
                            "thisNode = result.iterateNext();" & _
                            "while (thisNode) {" & _
                                "if(count === " & index & "){" & _
                                    "break;" & _
                                "}" & _
                                "thisNode = result.iterateNext();" & _
                                "count ++;" & _
                            "}" & _
                            "return thisNode;" & _
                        "}"
    GetXPathResultCodes = Codes
End Function

' 共通エラーを送出して処理終了する
' デバッグ中のハンドル未解放を避けるため強制終了とする
Public Sub DisplayErrorAndEnd(displayText As String)
    displayText = displayText & vbCrLf & "処理を終了します。"
    Err.Raise vbObjectError + 2048, "a6_ExecuteHelperFunction.DisplayErrorAndEnd", displayText
End Sub

' select 要素の選択変更用 JavaScript コードを返す
Public Function GetSelectItemCode() As String
    GetSelectItemCode = _
        "function(value){value = decodeURI(value);" & _
            "for(let i = 0; i < this.options.length; i++){" & _
                "let val1 = this.options[i].value.replace(/\r?\n/g,'');" & _
                "let val2 = this.options[i].textContent.replace(/\r?\n/g,'');" & _
                "if(val1 === value || val2 === value){" & _
                    "this.options[i].selected = true;" & _
                    "return true;" & _
                "}" & _
            "}" & _
        "return false;}"
End Function

' select 要素の選択テキスト取得用 JavaScript コードを返す
Public Function GetSelectedTextContentCode() As String
    GetSelectedTextContentCode = _
    "function(){" & _
        "for(let i = 0; i < this.options.length; i++){" & _
            "if(this.options[i].selected){" & _
                "return this.options[i].textContent.replace(/\\s/g, ' ');" & _
            "}" & _
        "}" & _
    "}"
End Function

' change イベント発火用 JavaScript コードを返す
Public Function GetDispatchChangeEventCode() As String
    GetDispatchChangeEventCode = _
    "function(){" & _
        "this.dispatchEvent(new CustomEvent('change', {bubbles: true}));" & _
    "}"
End Function

' input イベント発火用 JavaScript コードを返す
Public Function GetDispatchInputEventCode() As String
    GetDispatchInputEventCode = _
    "function(){" & _
        "this.dispatchEvent(new CustomEvent('input', {bubbles: true}));" & _
    "}"
End Function

' focus イベント発火用 JavaScript コードを返す
Public Function GetDispatchFocusEventCode() As String
    GetDispatchFocusEventCode = _
    "function(){" & _
        "this.dispatchEvent(new CustomEvent('focus', {bubbles: true}));" & _
    "}"
End Function

' MouseEvent 発火用 JavaScript コードを返す
Public Function GetDispatchMouseEventCode(action As JS_Mouse) As String
    Dim eventName As String: eventName = mouseEventNames_(action)

    GetDispatchMouseEventCode = _
    "function(){" & _
        "this.dispatchEvent(new MouseEvent('" & eventName & "', {bubbles: true}));" & _
    "}"
End Function

' KeyboardEvent 発火用 JavaScript コードを返す
Public Function GetDispatchKeyboardEventCode(action As JS_Keyboard, KeyCode As key, Optional alt As Boolean = False, Optional shift As Boolean = False, Optional ctrl As Boolean = False) As String
    Dim eventActionName As String: eventActionName = keyboardEventNames_(action)

    Dim KeyName As String: KeyName = keyNames_(KeyCode)

    Dim optionArgs As String
    optionArgs = "{keyCode: " & KeyCode & ",key: '" & KeyName & "' ,bubbles: true"
    If alt Then optionArgs = optionArgs & ",altKey: true"
    If shift Then optionArgs = optionArgs & ",shiftKey: true"
    If ctrl Then optionArgs = optionArgs & ",ctrlKey: true "
    optionArgs = optionArgs & "}"

    GetDispatchKeyboardEventCode = _
    "function(){" & _
        "this.dispatchEvent(new KeyboardEvent('" & eventActionName & "'," & optionArgs & "));" & _
    "}"
End Function

'Public Function GetDispatchKeyArrowDownEventCode() As String
'    GetDispatchKeyArrowDownEventCode = _
'    "function(){" & _
'        "this.dispatchEvent(new KeyboardEvent('keydown', {keyCode: 40,key: 'ArrowDown' ,bubbles: true}));" & _
'    "}"
'End Function

' 要素探索条件を Collection として返す
Public Function MakeFindInfo(way_string As String, way As String, Optional index As Long = -1) As Collection
    Dim findInfo As Collection: Set findInfo = New Collection
    findInfo.Add way_string
    findInfo.Add way
    findInfo.Add index
    Set MakeFindInfo = findInfo
End Function

' Find 系メソッド用のエラーメッセージを生成する
Public Function MakeFindErrorMessage(wayString As String, way As String, Optional wait As String = "") As String
    Dim errMsg As String
    errMsg = _
    "＜FindElementBy" & wayString & "＞" & vbCrLf & _
    "指定された以下" & wayString & "の要素は存在しませんでした。" & vbCrLf & _
    way & vbCrLf & "(イミディエイトウィンドウで確認可能)" & vbCrLf & vbCrLf

    If wayString = "XPath" Then
        errMsg = errMsg & _
        "※引数のXPathにはChromeの検証ツールで「完全なXPath」が" & vbCrLf & _
        "　 コピペできるので、それを指定してください。" & vbCrLf & vbCrLf
    End If

    errMsg = errMsg & _
        "※もし取得したい要素がiframe内の要素である場合、SwitchFrameを" & vbCrLf & _
        "　 使用して、フレームを切り替えてからでないと取得できません。" & vbCrLf & vbCrLf

    errMsg = errMsg & _
    "※上記に沿っていても、要素が画面に出現する前だと失敗します。" & vbCrLf

    If wait = "" Then
        errMsg = errMsg & _
        "　 必要であればWaitForElementToAppearBy" & wayString & "で" & vbCrLf
    Else
        errMsg = errMsg & _
        "　 必要であればWaitForElementToAppearBy" & wait & "で" & vbCrLf
    End If

    errMsg = errMsg & _
    "　 要素の出現を確認する処理を入れましょう。" & vbCrLf

    MakeFindErrorMessage = errMsg
End Function

' ファイル名に使用不可文字が含まれるかを判定する
Public Function IncludeNGCharInFileName(ByVal fileName As String) As Boolean
    Dim ng As String
    ng = "\/:*?""<>|"
    Dim i As Long
    For i = 1 To Len(ng)
        If InStr(fileName, Mid$(ng, i, 1)) > 0 Then
            IncludeNGCharInFileName = True
            Exit Function
        End If
    Next
    IncludeNGCharInFileName = False
End Function

' Image Enum を拡張子文字列へ変換する
Public Function GetFormatName(ByVal imageFormat As Long) As String
    Select Case imageFormat
        Case 2: GetFormatName = "jpeg"
        Case 3: GetFormatName = "webp"
        Case Else: GetFormatName = "png"
    End Select
End Function

' folderPath 内で未使用の連番を返す
Public Function GetFileNumber(ByVal folderPath As String, ByVal ext As String) As String
    Dim n As Long: n = 1
    Do While Dir(folderPath & "\*" & n & "." & ext) <> ""
        n = n + 1
    Loop
    GetFileNumber = CStr(n)
End Function

' Base64 文字列を Byte 配列へ復号する
Public Function DecodeBase64(ByVal base64Text As String) As Variant
    Dim xml As Object
    Dim node As Object
    Set xml = CreateObject("MSXML2.DOMDocument")
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.text = base64Text
    DecodeBase64 = node.nodeTypedValue
End Function

' CSS セレクタ指定で要素取得用の JavaScript コードを返す
Public Function GetCssCodes(cssSelector As String) As Variant
    Dim Codes(1) As String

    Codes(0) = "document.querySelector(decodeURIComponent('" & cssSelector & "'));"
    Codes(1) = "function(cssSelector){return document.querySelector(decodeURIComponent(cssSelector));}"

    GetCssCodes = Codes
End Function

' CSS セレクタ指定で要素数取得用の JavaScript コードを返す
Public Function GetCssCountCodes(cssSelector As String) As Variant
    Dim countCodes(1) As String

    countCodes(0) = "document.querySelectorAll(decodeURIComponent('" & cssSelector & "')).length;"
    countCodes(1) = "function(cssSelector){return document.querySelectorAll(decodeURIComponent(cssSelector)).length;}"

    GetCssCountCodes = countCodes
End Function

' CSS セレクタ指定で index 要素取得用の JavaScript コードを返す
Public Function GetCssResultCodes(cssSelector As String, index As Long) As Variant
    Dim Codes(1) As String

    Codes(0) = "document.querySelectorAll(decodeURIComponent('" & cssSelector & "'))[" & index & "];"
    Codes(1) = "function(cssSelector){return document.querySelectorAll(decodeURIComponent(cssSelector))[" & index & "];}"

    GetCssResultCodes = Codes
End Function

' テキスト指定で要素取得用の JavaScript コードを返す
Public Function GetTextCodes(text As String, Optional partialMatch As Boolean = False) As Variant
    Dim Codes(1) As String
    Dim partial As String

    If partialMatch Then
        partial = "true"
    Else
        partial = "false"
    End If

    Codes(0) = "(function(){" & _
               "var target=decodeURIComponent('" & text & "');" & _
               "var partial=" & partial & ";" & _
               "var elems=document.getElementsByTagName('*');" & _
               "var i,cur;" & _
               "for(i=0;i<elems.length;i++){" & _
               "cur=(elems[i].textContent||'').replace(/\s+/g,' ').trim();" & _
               "if((partial && cur.indexOf(target) >= 0) || (!partial && cur === target)){" & _
               "return elems[i];" & _
               "}" & _
               "}" & _
               "return null;" & _
               "})();"

    Codes(1) = "function(text){" & _
               "var target=decodeURIComponent(text);" & _
               "var partial=" & partial & ";" & _
               "var elems=document.getElementsByTagName('*');" & _
               "var i,cur;" & _
               "for(i=0;i<elems.length;i++){" & _
               "cur=(elems[i].textContent||'').replace(/\s+/g,' ').trim();" & _
               "if((partial && cur.indexOf(target) >= 0) || (!partial && cur === target)){" & _
               "return elems[i];" & _
               "}" & _
               "}" & _
               "return null;" & _
               "}"

    GetTextCodes = Codes
End Function

' テキスト指定で要素数取得用の JavaScript コードを返す
Public Function GetTextCountCodes(text As String, Optional partialMatch As Boolean = False) As Variant
    Dim Codes(1) As String
    Dim partial As String

    If partialMatch Then
        partial = "true"
    Else
        partial = "false"
    End If

    Codes(0) = "(function(){" & _
               "var target=decodeURIComponent('" & text & "');" & _
               "var partial=" & partial & ";" & _
               "var elems=document.getElementsByTagName('*');" & _
               "var i,cur,cnt=0;" & _
               "for(i=0;i<elems.length;i++){" & _
               "cur=(elems[i].textContent||'').replace(/\s+/g,' ').trim();" & _
               "if((partial && cur.indexOf(target) >= 0) || (!partial && cur === target)){" & _
               "cnt++;" & _
               "}" & _
               "}" & _
               "return cnt;" & _
               "})();"

    Codes(1) = "function(text){" & _
               "var target=decodeURIComponent(text);" & _
               "var partial=" & partial & ";" & _
               "var elems=document.getElementsByTagName('*');" & _
               "var i,cur,cnt=0;" & _
               "for(i=0;i<elems.length;i++){" & _
               "cur=(elems[i].textContent||'').replace(/\s+/g,' ').trim();" & _
               "if((partial && cur.indexOf(target) >= 0) || (!partial && cur === target)){" & _
               "cnt++;" & _
               "}" & _
               "}" & _
               "return cnt;" & _
               "}"

    GetTextCountCodes = Codes
End Function

' テキスト指定で index 要素取得用の JavaScript コードを返す
Public Function GetTextResultCodes(text As String, index As Long, Optional partialMatch As Boolean = False) As Variant
    Dim Codes(1) As String
    Dim partial As String

    If partialMatch Then
        partial = "true"
    Else
        partial = "false"
    End If

    Codes(0) = "(function(){" & _
               "var target=decodeURIComponent('" & text & "');" & _
               "var partial=" & partial & ";" & _
               "var elems=document.getElementsByTagName('*');" & _
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
               "})();"

    Codes(1) = "function(text){" & _
               "var target=decodeURIComponent(text);" & _
               "var partial=" & partial & ";" & _
               "var elems=document.getElementsByTagName('*');" & _
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

    GetTextResultCodes = Codes
End Function

' Shadow DOM を含む CSS セレクタ検索用 JavaScript コードを返す
Public Function GetDeepCssCodes(cssSelector As String) As Variant
    Dim Codes(1) As String

    Codes(0) = "(function(){" & _
               "var selector=decodeURIComponent('" & cssSelector & "');" & _
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
               "return findIn(document, selector);" & _
               "})();"

    Codes(1) = "function(cssSelector){" & _
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
               "return findIn(document, selector);" & _
               "}"

    GetDeepCssCodes = Codes
End Function

' Shadow DOM を含む CSS セレクタ件数取得用 JavaScript コードを返す
Public Function GetDeepCssCountCodes(cssSelector As String) As Variant
    Dim Codes(1) As String

    Codes(0) = "(function(){" & _
               "var selector=decodeURIComponent('" & cssSelector & "');" & _
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
               "collectIn(document, selector, result);" & _
               "return result.length;" & _
               "})();"

    Codes(1) = "function(cssSelector){" & _
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
               "collectIn(document, selector, result);" & _
               "return result.length;" & _
               "}"

    GetDeepCssCountCodes = Codes
End Function

' Shadow DOM を含む CSS セレクタ index 要素取得用 JavaScript コードを返す
Public Function GetDeepCssResultCodes(cssSelector As String, index As Long) As Variant
    Dim Codes(1) As String

    Codes(0) = "(function(){" & _
               "var selector=decodeURIComponent('" & cssSelector & "');" & _
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
               "collectIn(document, selector, result);" & _
               "return (result.length > " & index & ") ? result[" & index & "] : null;" & _
               "})();"

    Codes(1) = "function(cssSelector){" & _
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
               "collectIn(document, selector, result);" & _
               "return (result.length > " & index & ") ? result[" & index & "] : null;" & _
               "}"

    GetDeepCssResultCodes = Codes
End Function

' Shadow DOM を含むテキスト検索用 JavaScript コードを返す
Public Function GetDeepTextCodes(text As String, Optional partialMatch As Boolean = False) As Variant
    Dim Codes(1) As String
    Dim partial As String

    If partialMatch Then
        partial = "true"
    Else
        partial = "false"
    End If

    Codes(0) = "(function(){" & _
               "var target=decodeURIComponent('" & text & "');" & _
               "var partial=" & partial & ";" & _
               "function normalizeText(s){" & _
               "return (s||'').replace(/\s+/g,' ').trim();" & _
               "}" & _
               "function findIn(root){" & _
               "if(!root||!root.querySelectorAll){return null;}" & _
               "var nodes=root.querySelectorAll('*');" & _
               "for(var i=0;i<nodes.length;i++){" & _
               "var cur=normalizeText(nodes[i].textContent);" & _
               "if((partial && cur.indexOf(target) >= 0) || (!partial && cur === target)){" & _
               "return nodes[i];" & _
               "}" & _
               "if(nodes[i].shadowRoot){" & _
               "var subHit=findIn(nodes[i].shadowRoot);" & _
               "if(subHit){return subHit;}" & _
               "}" & _
               "}" & _
               "return null;" & _
               "}" & _
               "return findIn(document);" & _
               "})();"

    Codes(1) = "function(text){" & _
               "var target=decodeURIComponent(text);" & _
               "var partial=" & partial & ";" & _
               "function normalizeText(s){" & _
               "return (s||'').replace(/\s+/g,' ').trim();" & _
               "}" & _
               "function findIn(root){" & _
               "if(!root||!root.querySelectorAll){return null;}" & _
               "var nodes=root.querySelectorAll('*');" & _
               "for(var i=0;i<nodes.length;i++){" & _
               "var cur=normalizeText(nodes[i].textContent);" & _
               "if((partial && cur.indexOf(target) >= 0) || (!partial && cur === target)){" & _
               "return nodes[i];" & _
               "}" & _
               "if(nodes[i].shadowRoot){" & _
               "var subHit=findIn(nodes[i].shadowRoot);" & _
               "if(subHit){return subHit;}" & _
               "}" & _
               "}" & _
               "return null;" & _
               "}" & _
               "return findIn(document);" & _
               "}"

    GetDeepTextCodes = Codes
End Function

' Shadow DOM を含むテキスト件数取得用 JavaScript コードを返す
Public Function GetDeepTextCountCodes(text As String, Optional partialMatch As Boolean = False) As Variant
    Dim Codes(1) As String
    Dim partial As String

    If partialMatch Then
        partial = "true"
    Else
        partial = "false"
    End If

    Codes(0) = "(function(){" & _
               "var target=decodeURIComponent('" & text & "');" & _
               "var partial=" & partial & ";" & _
               "function normalizeText(s){" & _
               "return (s||'').replace(/\s+/g,' ').trim();" & _
               "}" & _
               "function countIn(root){" & _
               "if(!root||!root.querySelectorAll){return 0;}" & _
               "var nodes=root.querySelectorAll('*');" & _
               "var i,cnt=0;" & _
               "for(i=0;i<nodes.length;i++){" & _
               "var cur=normalizeText(nodes[i].textContent);" & _
               "if((partial && cur.indexOf(target) >= 0) || (!partial && cur === target)){" & _
               "cnt++;" & _
               "}" & _
               "if(nodes[i].shadowRoot){" & _
               "cnt += countIn(nodes[i].shadowRoot);" & _
               "}" & _
               "}" & _
               "return cnt;" & _
               "}" & _
               "return countIn(document);" & _
               "})();"

    Codes(1) = "function(text){" & _
               "var target=decodeURIComponent(text);" & _
               "var partial=" & partial & ";" & _
               "function normalizeText(s){" & _
               "return (s||'').replace(/\s+/g,' ').trim();" & _
               "}" & _
               "function countIn(root){" & _
               "if(!root||!root.querySelectorAll){return 0;}" & _
               "var nodes=root.querySelectorAll('*');" & _
               "var i,cnt=0;" & _
               "for(i=0;i<nodes.length;i++){" & _
               "var cur=normalizeText(nodes[i].textContent);" & _
               "if((partial && cur.indexOf(target) >= 0) || (!partial && cur === target)){" & _
               "cnt++;" & _
               "}" & _
               "if(nodes[i].shadowRoot){" & _
               "cnt += countIn(nodes[i].shadowRoot);" & _
               "}" & _
               "}" & _
               "return cnt;" & _
               "}" & _
               "return countIn(document);" & _
               "}"

    GetDeepTextCountCodes = Codes
End Function

' Shadow DOM を含むテキスト index 要素取得用 JavaScript コードを返す
Public Function GetDeepTextResultCodes(text As String, index As Long, Optional partialMatch As Boolean = False) As Variant
    Dim Codes(1) As String
    Dim partial As String

    partial = IIf(partialMatch, "true", "false")

    ' 共通 JS 本体を 1 つにまとめる
    Dim jsCore As String
    jsCore = "var target=decodeURIComponent(ARG_TEXT);" & _
             "var partial=" & partial & ";" & _
             "function normalizeText(s){return (s||'').replace(/\s+/g,' ').trim();}" & _
             "function findAt(root, targetIndex){" & _
             "if(!root||!root.querySelectorAll){return {count:0,node:null};}" & _
             "var nodes=root.querySelectorAll('*');" & _
             "var i,cnt=0;" & _
             "for(i=0;i<nodes.length;i++){" & _
             "var cur=normalizeText(nodes[i].textContent);" & _
             "if((partial && cur.indexOf(target) >= 0) || (!partial && cur === target)){" & _
             "if(cnt === targetIndex){return {count:cnt,node:nodes[i]};}" & _
             "cnt++;" & _
             "}" & _
             "if(nodes[i].shadowRoot){" & _
             "var sub=findAt(nodes[i].shadowRoot, targetIndex - cnt);" & _
             "if(sub.node){return sub;}" & _
             "cnt += sub.count;" & _
             "}" & _
             "}" & _
             "return {count:cnt,node:null};" & _
             "}"

    ' Codes(0): 即時実行IIFE
    Codes(0) = "(function(){" & _
               Replace(jsCore, "ARG_TEXT", "'" & text & "'") & _
               "return findAt(document, " & index & ").node;" & _
               "})();"

    ' Codes(1): 関数文字列（引数textを受け取る）
    Codes(1) = "function(text){" & _
               Replace(jsCore, "ARG_TEXT", "text") & _
               "return findAt(document, " & index & ").node;" & _
               "}"

    GetDeepTextResultCodes = Codes
End Function



