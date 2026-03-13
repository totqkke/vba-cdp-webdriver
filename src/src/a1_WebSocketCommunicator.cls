VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a1_WebSocketCommunicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const NO_ERROR As Long = 0

Private Declare PtrSafe Function WinHttpWebSocketSend Lib "WinHttp" ( _
    ByVal hWebSocket As LongPtr, _
    ByVal eBufferType As Long, _
    ByVal pvBuffer As LongPtr, _
    ByVal dwBufferLength As Long _
    ) As Long

Private Const WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE = 0 'バイナリメッセージ全体または最終フラグメントを表す。
Private Const WINHTTP_WEB_SOCKET_BINARY_FRAGMENT_BUFFER_TYPE = 1 'バイナリメッセージの途中フラグメントを表す。
Private Const WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE = 2 'UTF-8メッセージ全体または最終フラグメントを表す。
Private Const WINHTTP_WEB_SOCKET_UTF8_FRAGMENT_BUFFER_TYPE = 3 'UTF-8メッセージの途中フラグメントを表す。
Private Const WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE = 4 'CloseFrameを表す。

Private Declare PtrSafe Function WinHttpWebSocketReceive Lib "WinHttp" ( _
    ByVal hWebSocket As LongPtr, _
    ByRef pvBuffer As Any, _
    ByVal dwBufferLength As Long, _
    ByRef pdwBytesRead As Long, _
    ByRef peBufferType As Long _
    ) As Long

Private Declare PtrSafe Function WinHttpWebSocketClose Lib "WinHttp" ( _
    ByVal hWebSocket As LongPtr, _
    ByVal usStatus As Integer, _
    ByVal pvReason As LongPtr, _
    ByVal dwReasonLength As Long _
    ) As Long

Private Const WINHTTP_WEB_SOCKET_SUCCESS_CLOSE_STATUS = 1000
Private Const WINHTTP_WEB_SOCKET_ENDPOINT_TERMINATED_CLOSE_STATUS = 1001
Private Const WINHTTP_WEB_SOCKET_PROTOCOL_ERROR_CLOSE_STATUS = 1002
Private Const WINHTTP_WEB_SOCKET_INVALID_DATA_TYPE_CLOSE_STATUS = 1003
Private Const WINHTTP_WEB_SOCKET_EMPTY_CLOSE_STATUS = 1005
Private Const WINHTTP_WEB_SOCKET_ABORTED_CLOSE_STATUS = 1006
Private Const WINHTTP_WEB_SOCKET_INVALID_PAYLOAD_CLOSE_STATUS = 1007
Private Const WINHTTP_WEB_SOCKET_POLICY_VIOLATION_CLOSE_STATUS = 1008
Private Const WINHTTP_WEB_SOCKET_MESSAGE_TOO_BIG_CLOSE_STATUS = 1009
Private Const WINHTTP_WEB_SOCKET_UNSUPPORTED_EXTENSIONS_CLOSE_STATUS = 1010
Private Const WINHTTP_WEB_SOCKET_SERVER_ERROR_CLOSE_STATUS = 1011
Private Const WINHTTP_WEB_SOCKET_SECURE_HANDSHAKE_ERROR_CLOSE_STATUS = 1015

Private Declare PtrSafe Function WinHttpWebSocketQueryCloseStatus Lib "WinHttp" ( _
    ByVal hWebSocket As LongPtr, _
    ByRef usStatus As Integer, _
    ByRef pvReason As Any, _
    ByVal dwReasonLength As Long, _
    ByRef pdwReasonLengthConsumed As Long _
    ) As Long

' WebSocket受信に使用する自作構造体
Private Type response
    Buffer(4095) As Byte      '第1引数
    BufferLength As Long      '第2引数
    ReceiveBytes As Long      '第3引数
    status As Long            '第4引数
    CurrentPointer As Long    '第5引数
    result As Long            '戻り値
    collect As Collection
End Type

Public http As a1x1_HTTPCommunicator

' WebSocket接続を初期化する
Public Function Init(path As String) As Boolean
    Set http = New a1x1_HTTPCommunicator
    If http.Connect("127.0.0.1", path, 9222) = False Then GoTo quit
    If http.UpgradeToWebSocket() = False Then GoTo quit

    Init = True
Exit Function
quit:
    Init = False
End Function

' Chromeへメッセージを送信する
Public Function SendMessage(message As String) As Boolean
    Dim b() As Byte
    b = ToUtf8Bytes_NoBOM(message)

    Dim result As Long
    result = WinHttpWebSocketSend( _
                http.websockethandle, _
                WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE, _
                VarPtr(b(0)), _
                UBound(b) + 1)

    SendMessage = (result = NO_ERROR)
End Function

' 文字列をBOMなしUTF-8バイト列へ変換する
Private Function ToUtf8Bytes_NoBOM(ByVal s As String) As Byte()
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 2          'テキストとして扱う
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText s

    stm.Position = 0
    stm.Type = 1          'バイナリとして扱う
    Dim b() As Byte
    b = stm.Read
    stm.Close

    ' BOMを除去して送信データをUTF-8本体のみにする
    If UBound(b) >= 2 Then
        If b(0) = CByte(&HEF) And b(1) = CByte(&HBB) And b(2) = CByte(&HBF) Then
            Dim b2() As Byte, i As Long
            ReDim b2(UBound(b) - 3)
            For i = 3 To UBound(b)
                b2(i - 3) = b(i)
            Next
            b = b2
        End If
    End If

    ToUtf8Bytes_NoBOM = b
End Function

' 受信したフラグメントを結合用コレクションへ退避する
Private Sub CollectBytes(col As Collection, buf() As Byte, count As Long)
    If count <= 0 Then Exit Sub

    Dim tmp() As Byte
    Dim i As Long
    ReDim tmp(count - 1)

    For i = 0 To count - 1
        tmp(i) = buf(i)
    Next

    col.Add tmp
End Sub

' Chromeからのレスポンス(JSON)を受信する
Public Function GetMessage() As String
    Dim res As response
    Dim bufLen As Long
    bufLen = UBound(res.Buffer) + 1

    Set res.collect = New Collection

    Do
        res.result = WinHttpWebSocketReceive( _
                        http.websockethandle, res.Buffer(0), bufLen, _
                        res.ReceiveBytes, res.status)
        If res.result <> NO_ERROR Then GoTo quit

        If (res.status = WINHTTP_WEB_SOCKET_UTF8_FRAGMENT_BUFFER_TYPE) Or _
           (res.status = WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE) Then
            CollectBytes res.collect, res.Buffer, res.ReceiveBytes

            ' 最終フラグメント受信時にメッセージ全体を復元する
            If res.status = WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE Then
                GetMessage = MakeResponseMessage(res.collect)
                Exit Function
            End If
        Else
            GoTo quit
        End If
    Loop

quit:
    GetMessage = "受信エラー発生"
End Function

' バッファ管理情報を更新する
Private Sub UpdateBufferInfo(CurrentPointer As Long, BufferLength As Long, ReceiveBytes)
    CurrentPointer = CurrentPointer + ReceiveBytes
    BufferLength = BufferLength - ReceiveBytes
End Sub

' バッファを初期化する
Private Sub InitializeBuffer(responseBuffer As Variant, CurrentPointer As Long, BufferLength As Long)
    Erase responseBuffer
    CurrentPointer = 0
    BufferLength = UBound(responseBuffer) + 1
End Sub

' フラグメントを連結してUTF-8バイト列から文字列へ復号する
Private Function MakeResponseMessage(collect As Collection) As String
    Dim ary As Variant
    Dim total As Long
    Dim b() As Byte
    Dim pos As Long
    Dim i As Long

    For Each ary In collect
        total = total + (UBound(ary) + 1)
    Next

    If total = 0 Then
        MakeResponseMessage = ""
        Exit Function
    End If

    ReDim b(total - 1)
    pos = 0

    For Each ary In collect
        For i = 0 To UBound(ary)
            b(pos) = ary(i)
            pos = pos + 1
        Next
    Next

    MakeResponseMessage = Utf8BytesToString(b)
End Function

' UTF-8バイト列を文字列へ変換する
Private Function Utf8BytesToString(b() As Byte) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 1          'バイナリとして扱う
    stm.Open
    stm.Write b

    stm.Position = 0
    stm.Type = 2          'テキストとして扱う
    stm.Charset = "utf-8"
    Utf8BytesToString = stm.ReadText(-1)

    stm.Close
End Function

' 正常クローズを送信し、応答ステータスを確認する
Public Function CloseWebSocket() As Boolean
    Dim result As Long
    result = WinHttpWebSocketClose(http.websockethandle, _
                        WINHTTP_WEB_SOCKET_SUCCESS_CLOSE_STATUS, 0, 0)
    If result <> NO_ERROR Then GoTo quit

    ' サーバーから返されたクローズステータスを確認する
    Dim CloseReasonBuffer(123) As Byte
    Dim BufferSize As Long: BufferSize = UBound(CloseReasonBuffer) + 1
    Dim status As Integer
    Dim CloseReasonLength As Long

    result = WinHttpWebSocketQueryCloseStatus(http.websockethandle, status, _
                        CloseReasonBuffer(0), BufferSize, CloseReasonLength)
    If result <> NO_ERROR Then GoTo quit
    If status <> WINHTTP_WEB_SOCKET_SUCCESS_CLOSE_STATUS Then GoTo quit

    CloseWebSocket = True
Exit Function

quit:
    If result <> NO_ERROR Then Debug.Print result

    ' クローズ理由が返されている場合に備えて文字列化する
    'Debug.Print "The server closed the connection with status code: " & Status
    Dim strCloseReason As String
    Dim i As Long
    For i = LBound(CloseReasonBuffer) To UBound(CloseReasonBuffer)
        If CloseReasonBuffer(i) <> 0 Then
            strCloseReason = strCloseReason & Chr(CloseReasonBuffer(i))
        End If
    Next

    If Len(strCloseReason) > 0 Then
        'Debug.Print " and reason: " & strCloseReason
    End If
    CloseWebSocket = False
End Function

