VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a2_JSONHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private html As Object

Private Sub Class_Initialize()
    Set html = CreateObject("htmlfile")
    html.Write "<meta http-equiv='X-UA-Compatible' content='IE=edge' />"

    ' URLエンコード関数を初期化する
    html.Write _
        "<script>" & vbCrLf & _
            "document.EncodeURL = " & vbCrLf & _
            "function(s){" & vbCrLf & _
            "return encodeURI(s);" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"

    ' URLデコード関数を初期化する
    html.Write _
        "<script>" & vbCrLf & _
            "document.DecodeURL = " & vbCrLf & _
            "function(s){" & vbCrLf & _
            "return decodeURI(s);" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"

    ' JSON操作関数を初期化する
    '============================

    ' 配列からオブジェクトを取り出す関数を初期化する
    html.Write _
        "<script>" & vbCrLf & _
            "document.getObjectFromArrayBySetIndex = " & vbCrLf & _
            "function(targetJson,index){" & vbCrLf & _
            "var vals = JSON.parse(targetJson);" & vbCrLf & _
            "return JSON.stringify(vals[index]);" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"

    ' オブジェクトからオブジェクトを取り出す関数を初期化する
    html.Write _
        "<script>" & vbCrLf & _
            "document.getObjectFromObjectBySetKey = " & vbCrLf & _
            "function(targetJson,key){" & vbCrLf & _
            "var vals = JSON.parse(targetJson);" & vbCrLf & _
            "return JSON.stringify(vals[key]);" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"

    ' 配列から値を取り出す関数を初期化する
    html.Write _
        "<script>" & vbCrLf & _
            "document.getValueFromArrayBySetIndex = " & vbCrLf & _
            "function(targetJson,index){" & vbCrLf & _
            "var vals = JSON.parse(targetJson);" & vbCrLf & _
            "return vals[index];" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"

    ' オブジェクトから値を取り出す関数を初期化する
    html.Write _
        "<script>" & vbCrLf & _
            "document.getValueFromObjectBySetKey = " & vbCrLf & _
            "function(targetJson,key){" & vbCrLf & _
            "var vals = JSON.parse(targetJson);" & vbCrLf & _
            "return vals[key];" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"

    ' 配列長を取得する関数を初期化する
    html.Write _
        "<script>" & vbCrLf & _
            "document.getJsonLength = " & vbCrLf & _
            "function(targetJson){" & vbCrLf & _
            "var vals = JSON.parse(targetJson);" & vbCrLf & _
            "return vals.length;" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"
End Sub

' URLエンコード結果を返す
Public Function EncodeURL(json As String) As String
    EncodeURL = html.EncodeURL(json)
End Function

' URLデコード結果を返す
Public Function DecodeURL(json As String) As String
    DecodeURL = html.DecodeURL(json)
End Function

' JSONから末端の値を取り出す
Public Function GetValue(ByVal json As String, ParamArray argsArray() As Variant) As String
    Dim i As Long
    Dim result As String

    For i = LBound(argsArray) To UBound(argsArray)
        If json = "" Then
            GetValue = result
            Exit Function
        End If

        If i <> UBound(argsArray) Then
            If IsNumeric(argsArray(i)) Then
                json = html.getObjectFromArrayBySetIndex(json, argsArray(i))
            Else
                json = html.getObjectFromObjectBySetKey(json, argsArray(i))
            End If
        Else
            If IsNumeric(argsArray(i)) Then
                result = html.getValueFromArrayBySetIndex(json, argsArray(i))
            Else
                result = html.getValueFromObjectBySetKey(json, argsArray(i))
            End If
        End If
    Next

    GetValue = result
End Function

' JSONから途中または末端のオブジェクトを取り出す
Public Function GetObject(ByVal json As String, ParamArray argsArray() As Variant) As String
    Dim i As Long
    Dim result As String

    For i = LBound(argsArray) To UBound(argsArray)
        If json = "" Then
            GetObject = json
            Exit Function
        End If

        If IsNumeric(argsArray(i)) Then
            json = html.getObjectFromArrayBySetIndex(json, argsArray(i))
        Else
            json = html.getObjectFromObjectBySetKey(json, argsArray(i))
        End If
    Next

    GetObject = json
End Function

' JSON配列の長さを返す
Public Function GetJsonLength(json As String) As Long
    GetJsonLength = html.GetJsonLength(json)
End Function

