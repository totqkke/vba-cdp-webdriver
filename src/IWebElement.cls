VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWebElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' JavaScript MouseEvent 用の列挙体
Public Enum JS_Mouse
: LeftDown = 1: LeftPress: Leftup
: Enter: Leave: Move
: Over: out
: LeftClick: LeftDoublelClick
': Drag: DragStart: DragEnd
': DragEnter: DragExit: DragLeave: DragOver
': Drop
End Enum

' auxclick / Scroll / wheel は別扱いとする
' JavaScript KeyboardEvent 用の列挙体
Public Enum JS_Keyboard
    KeyDown = 1
    KeyPress
    KeyUp
End Enum

' キーコードを表す列挙体
Public Enum key
: Tab_Key = 9: Enter_Key = 13: shift_Key = 16: Control_Key = 17: Alt_Key = 18
: Escape_Key = 27: Space_Key = 32
: PageUp_Key = 33: PageDown_Key = 34
: ArrowLeft_Key = 37: ArrowUp_Key = 38: ArrowRight_Key = 39: ArrowDown_Key = 40
: PrintScreen_Key = 44: Delete_Key = 46
: Digit0 = 48: Digit1: Digit2: Digit3: Digit4:
: Digit5: Digit6: Digit7: Digit8: Digit9
: KeyA = 65: KeyB: KeyC: KeyD: KeyE: KeyF: KeyG
: KeyH: KeyI: KeyJ: KeyK: KeyL: KeyM: KeyN: KeyO
: KeyP: KeyQ: KeyR: KeyS: KeyT: KeyU: KeyV: KeyW: KeyX: KeyY: KeyZ
End Enum

' 参考となる他イベント
' Select / submit / toggle などは別扱いとする
Public Enum CDP_Mouse
    LeftClick_XY = 1
    LeftDoubleClick_XY
    RightClick_XY
    ScrollToThisElement
End Enum

' CDP キーボードイベント種別を表す列挙体
Public Enum CDP_Keyboard
    KeyDown_ = 1
    KeyUp_
    RawKeyDown
    Char
End Enum

' クリック位置を表す列挙体
Public Enum ClickPoint
    Center = 1
    topLeft
    TopRight
    BottomLeft
    BottomRight
End Enum

' click は press と release の組み合わせで表す
' middle / Moved / MouseWheel / Back / Forward は別扱いとする
Public Property Get Checked() As Boolean
End Property

Public Property Let Checked(bool As Boolean)
End Property

Public Property Get IsEnabled() As Boolean
End Property

' 互換性維持のため SetValue と同等処理を残す
Public Sub SetText(text As String, Optional timeoutSec As Long = 5)
End Sub

' 要素の value 属性へ文字列を設定する
Public Sub SetValue(value As String, Optional timeoutSec As Long = 5)
End Sub

' 要素の textContent へ文字列を設定する
Public Sub SetTextContent(text As String)
End Sub

' 要素へフォーカスして Enter キーを送る
Public Sub FocusAndEnterKey()
End Sub

' PropertySetter 経由で value を設定する
Public Function SetValueViaPropertySetter(value As String) As String
End Function

' 要素をクリックする
Public Sub Click()
End Sub

' SelectBox の選択肢を文字列で選択する
Public Sub SelectItemInSelectBoxByText(text As String)
End Sub

' 子要素から Name 指定で 1 件取得する
Public Function FindElementByName(name As String) As IWebElement
End Function

' 子要素からタグ名指定で複数取得する
Public Function FindElementsByTag(tagName As String) As IWebElements
End Function

' 要素の textContent を取得する
Public Function GetTextContent() As String
End Function

' セレクトボックスの選択中 value を取得する
Public Function GetSelectedValue() As String
End Function

' セレクトボックスの選択中 textContent を取得する
Public Function GetSelectedTextContent() As String
End Function

' type=file の input 要素へファイルを設定する
Public Sub SetFileToFileSelectInputButton(filePath As String)
End Sub

' クリック後に制御を返さないアラート出現時の処理を行う
Public Sub ClickAndThenAlertDialogErase()
End Sub

' 要素の属性値を取得する
Public Function GetAttribute(attributeName As String) As String
End Function

' 要素の属性値を設定する
Public Function SetAttribute(attributeName As String, value As Variant) As String
'
End Function

' 要素の class 属性へ値を設定する
Public Sub SetClassName(className As String)
End Sub

' CDP マウスイベントを発行する
Public Sub DispatchCDPMouseEvent(action As CDP_Mouse, Optional point As ClickPoint = ClickPoint.Center)
End Sub

' CDP キーボードイベントを発行する
Public Sub DispatchCDPKeyEvent(action As CDP_Keyboard, KeyCode As key, Optional alt As Boolean = False, Optional shift As Boolean = False, Optional ctrl As Boolean = False)
End Sub

' JavaScript input イベントを発火する
Public Sub DispatchJSInputEvent()
End Sub

' JavaScript change イベントを発火する
Public Sub DispatchJSChangeEvent()
End Sub

' JavaScript focus イベントを発火する
Public Sub DispatchJSFocusEvent()
End Sub

' JavaScript MouseEvent を発火する
Public Sub DispatchJSMouseEvent(action As JS_Mouse)
End Sub

' JavaScript KeyboardEvent を発火する
Public Sub DispatchJSKeyboardEvent(action As JS_Keyboard, KeyCode As key, Optional alt As Boolean = False, Optional shift As Boolean = False, Optional ctrl As Boolean = False)
End Sub

' 要素が表示状態になるまで待機する
Public Sub WaitUntilVisible(Optional timeoutSec As Long = 5)
End Sub

' 要素が有効状態になるまで待機する
Public Sub WaitUntilEnabled(Optional timeoutSec As Long = 5)
End Sub

' 要素が非表示または消滅するまで待機する
Public Sub WaitUntilHidden(Optional timeoutSec As Long = 5)
End Sub

' 要素プロパティ値を取得する
Public Function GetProperty(PropertyName As String) As Variant
End Function

' 要素プロパティ値を設定する
Public Function SetProperty(PropertyName As String, value As Variant) As String
'
End Function

' CSS セレクタ指定で要素を取得する
Public Function FindElementByCss(cssSelector As String) As IWebElement
End Function

' CSS セレクタ指定で要素コレクションを取得する
Public Function FindElementsByCss(cssSelector As String) As IWebElements
End Function

' テキスト指定で要素を取得する
Public Function FindElementByText(text As String, Optional partialMatch As Boolean = False) As IWebElement
End Function

' テキスト指定で要素コレクションを取得する
Public Function FindElementsByText(text As String, Optional partialMatch As Boolean = False) As IWebElements
End Function

' Shadow DOM を含めて CSS セレクタ指定で要素を取得する
Public Function FindElementByCssDeep(cssSelector As String) As IWebElement
End Function

' Shadow DOM を含めて CSS セレクタ指定で要素コレクションを取得する
Public Function FindElementsByCssDeep(cssSelector As String) As IWebElements
End Function

' Shadow DOM を含めてテキスト指定で要素を取得する
Public Function FindElementByTextDeep(text As String, Optional partialMatch As Boolean = False) As IWebElement
End Function

' Shadow DOM を含めてテキスト指定で要素コレクションを取得する
Public Function FindElementsByTextDeep(text As String, Optional partialMatch As Boolean = False) As IWebElements
End Function

