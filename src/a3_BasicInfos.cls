VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a3_BasicInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)

Private Const HOST_ As String = "127.0.0.1"
Private Const DEBUGGING_PORT As Long = 9222
Private Const RESOURCE_PATH As String = "/json" 'タブ一覧を表す。
Private Const HTTP_SCHEMA As String = "http://"
Private Const WEBSOCKET_SCHEMA As String = "ws://"

Private Const KEY_WEBSOCKETURL As String = "webSocketDebuggerUrl"

Public PageDefaultTargetId As String

Public TargetInfoJSON As String
Public webSocketURL As String
Public targetId As String
Public webSocketURLPath As String
Public Context As String

Public DocumentNodeId As Long
Public DocumentObjectId As String

Public IsolatedFrameId As String
Public ExecutionContextId As Long

Public Property Get Host() As String
    Host = HOST_
End Property

Public Property Get DebuggingPort() As String
    DebuggingPort = DEBUGGING_PORT
End Property

Public Property Get ResourcePath() As String
    ResourcePath = RESOURCE_PATH
End Property

' 現在のターゲット情報を初期化する
Public Sub Init(json As a2_JSONHandler)
    TargetInfoJSON = GetTargetInfoJSON
    
    Dim i As Long
    For i = 0 To json.GetJsonLength(TargetInfoJSON)
        If json.GetValue(TargetInfoJSON, i, "type") = "page" _
        And json.GetValue(TargetInfoJSON, i, "title") = "新しいタブ" Then
            webSocketURL = json.GetValue(TargetInfoJSON, i, KEY_WEBSOCKETURL)
            Exit For
        End If
    Next
    
    targetId = json.GetValue(TargetInfoJSON, 0, "id")
    PageDefaultTargetId = targetId
    webSocketURLPath = Replace(webSocketURL, _
    WEBSOCKET_SCHEMA & HOST_ & ":" & DEBUGGING_PORT, "")
    Context = Replace(webSocketURLPath, targetId, "")
End Sub

' DevToolsのターゲット一覧JSONを取得する
Private Function GetTargetInfoJSON() As String
    Const TIMEOUT_SEC As Long = 15

    Dim client As Object
    Dim startedAt As Date

    startedAt = Now

Retry_:
    On Error GoTo RetryWait_

    Set client = Nothing
    Set client = CreateObject("MSXML2.ServerXMLHTTP")
    client.Open "GET", HTTP_SCHEMA & HOST_ & ":" & DEBUGGING_PORT & RESOURCE_PATH
    client.setRequestHeader "Content-Type", "application/json"
    client.send

    GetTargetInfoJSON = client.responseText
    Exit Function

RetryWait_:
    On Error GoTo 0
    DoEvents
    Sleep 500

    ' ブラウザ起動直後でも一定時間は再試行して接続確立を待機する
    If DateDiff("s", startedAt, Now) >= TIMEOUT_SEC Then
        Err.Raise vbObjectError + 2201, "a3_BasicInfos.GetTargetInfoJSON", _
                  "デバッグ対象ブラウザ(" & HOST_ & ":" & DEBUGGING_PORT & ")への接続が " & TIMEOUT_SEC & " 秒以内に確立できませんでした。"
    End If

    GoTo Retry_
End Function

' ターゲット切り替え後の基本情報を更新する
Public Sub Update(newTargetId As String, json As a2_JSONHandler, _
    isPageSwitch As Boolean)

    If isPageSwitch Then PageDefaultTargetId = newTargetId
    targetId = newTargetId
    TargetInfoJSON = GetTargetInfoJSON
    webSocketURL = GetWebSocketDebuggerUrl(json)
    webSocketURLPath = Replace(webSocketURL, WEBSOCKET_SCHEMA & HOST_ & ":" & DEBUGGING_PORT, "")
    Context = Replace(webSocketURLPath, targetId, "")
End Sub

' 対象targetIdに対応するWebSocket URLを取得する
Private Function GetWebSocketDebuggerUrl(json As a2_JSONHandler) As String
    Dim collect As Collection: Set collect = New Collection
    Dim targetInfo As String
    Dim count As Long

    Do
        targetInfo = json.GetObject(TargetInfoJSON, count)
        If targetInfo = "" Then Exit Do
        collect.Add targetInfo
        count = count + 1
    Loop

    Dim i As Long
    Dim targetType As String, Id As String

    For i = 1 To collect.count
        targetType = json.GetValue(collect(i), "type")
        If targetType <> "page" Then GoTo next_

        Id = json.GetValue(collect(i), "id")
        If Id <> targetId Then GoTo next_

        GetWebSocketDebuggerUrl = json.GetValue(collect(i), KEY_WEBSOCKETURL)
        Exit Function
next_:
    Next
End Function

' 既定ターゲットを指しているかを判定する
Public Function IsTargetPageDefault() As Boolean
    If targetId = PageDefaultTargetId Then
        IsTargetPageDefault = True
    Else
        IsTargetPageDefault = False
    End If
End Function

' IsolatedWorld再作成後のExecutionContextIdを更新する
Public Sub UpdateExecutionContextId(CDP As a4_ExecuteCDP)
    If IsolatedFrameId <> "" Then
        ExecutionContextId = CDP.CreateIsolatedFrameWorld(IsolatedFrameId)
        DoEvents
    End If
End Sub

