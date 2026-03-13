VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWebElements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 保持中の要素数を返す
Public Function count() As Long
End Function

' 指定 index の要素を返す
Public Function Item(index) As IWebElement
Attribute Item.VB_UserMemId = 0
End Function

' 列挙用の Enumerator を返す
Public Function NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
End Function

