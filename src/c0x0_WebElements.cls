VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c0x0_WebElements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWebElements

Private elements_ As Collection

' 要素コレクションを初期化する
Public Sub Init(elements As Collection)
    Set elements_ = elements
End Sub

' 保持中の要素数を返す
Private Function IWebElements_Count() As Long
    IWebElements_Count = elements_.count
End Function

' 指定 index の要素を返す
Private Function IWebElements_Item(index As Variant) As IWebElement
    Set IWebElements_Item = elements_.Item(index + 1)
End Function

' 列挙用の Enumerator を返す
Private Function IWebElements_NewEnum() As stdole.IUnknown
    Set IWebElements_NewEnum = elements_.[_NewEnum]
End Function

