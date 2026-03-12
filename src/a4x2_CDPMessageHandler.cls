VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a4x2_CDPMessageHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private webSocket_ As a1_WebSocketCommunicator
Private json_ As a2_JSONHandler
Private events_ As a5_CDPEventHandler

#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function GetTickCount64 Lib "kernel32" () As LongLong
    #Else
        Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    #End If
#Else
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

' ミリ秒単位の経過時間を取得する
Private Function TickMs() As Double
#If VBA7 Then
    #If Win64 Then
        TickMs = CDbl(GetTickCount64())
    #Else
        TickMs = CDbl(GetTickCount())
    #End If
#Else
    TickMs = CDbl(GetTickCount())
#End If
End Function

' 依存オブジェクトを初期化する
Public Sub Init(WebSocket As a1_WebSocketCommunicator, _
                json As a2_JSONHandler, _
                events As a5_CDPEventHandler)
    Set webSocket_ = WebSocket
    Set json_ = json
    Set events_ = events
End Sub

' メッセージ送信後に対応する応答または関連イベントを待機する
Public Function SendAndReceive(sendMsg As String, Optional timeoutSec As Long = 30) As String
    Dim sendId As Long: sendId = CLng(json_.GetValue(sendMsg, "id"))
    Debug.Print vbCrLf & ">>>>>>送信内容：" & sendMsg

    If webSocket_.SendMessage(sendMsg) = False Then GoTo quit

    Dim startTick As Double: startTick = TickMs()
    Dim res As String, resId As Variant
    Dim resIdVal As Long

    Do
        If (TickMs() - startTick) >= (CDbl(timeoutSec) * 1000#) Then
            res = ""
            GoTo quit
        End If

        res = webSocket_.GetMessage
        If res = "受信エラー発生" Then GoTo quit

        resId = json_.GetValue(res, "id")

        If IsNumeric(resId) Then
            resIdVal = CLng(resId)
        Else
            resIdVal = -1
        End If

        If resIdVal = -1 Then
            If InStr(res, "method") > 0 Then
                events_.GetInfo res

                ' ダイアログ出現時はイベント処理結果を優先して呼び出し元へ返す
                If events_.DialogInfoDic("IsExistDialog") Then Exit Function
            ElseIf InStr(res, """error""") > 0 Then
                SendAndReceive = res
                Exit Function
            End If
        Else
            If InStr(res, """error""") > 0 Then
                SendAndReceive = res
                Exit Function
            End If
        End If

        Debug.Print "受信：" & res
    Loop While (sendId <> resIdVal)

    SendAndReceive = res
    Exit Function

quit:
    Debug.Print "エラー:" & sendMsg

    Dim errMsg As String
    If res = "" Then
        errMsg = "タイムアウト/送信エラー発生：イミディエイト参照"
    ElseIf res = "受信エラー発生" Then
        errMsg = "受信エラー発生：イミディエイト参照"
    Else
        errMsg = "クロームからのエラー通知あり：イミディエイト参照"
    End If

    webSocket_.http.CloseHInternetHandles
    Err.Raise vbObjectError + 2301, "a4x2_CDPMessageHandler.SendAndReceive", errMsg
End Function

