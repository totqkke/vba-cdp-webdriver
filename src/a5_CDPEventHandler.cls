VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a5_CDPEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
' 応答待機中のレスポンス情報を保持する
Public NetworkResponses As Collection

Private json_ As a2_JSONHandler
Private basicInfo_ As a3_BasicInfos

' Network idle 判定用の状態を保持する
Public NetworkInFlight As Long
Public NetworkLastActivity As Double

' DOM 変化判定用の状態を保持する
Public DomUpdated As Boolean
Public DomLastUpdated As Double

' ダイアログ・ダウンロード関連情報を保持する
Public DialogInfoDic As Object
Public DownloadInfoDic As Object
Public FileChooserInfoDic As Object
Public LifecycleEvents As Collection
Public PageInfos As Collection
Public IframeInfos As Collection

Public IsCurrentTargetScheduledNavigation As Boolean
Public IsCurrentTargetStartedLoading As Boolean
Public IsCurrentTargetClearedScheduledNavigation As Boolean
Public IsCurrentTargetNavigated As Boolean

Public IsIsolatedFrameScheduledNavigation As Boolean
Public IsIsolatedFrameStartedLoading As Boolean
Public IsIsolatedFrameClearedScheduledNavigation As Boolean
Public IsIsolatedFrameNavigated As Boolean

Public IsCurrentContextIdDestroyed As Boolean

' 依存オブジェクトと内部状態を初期化する
Public Sub Init(json As a2_JSONHandler, basicInfo As a3_BasicInfos)
    Set json_ = json
    Set basicInfo_ = basicInfo

    Set DialogInfoDic = CreateObject("Scripting.Dictionary")
    InitializeDialogInfo
    Set DownloadInfoDic = CreateObject("Scripting.Dictionary")
    ClearDownloadInfo

    Set FileChooserInfoDic = CreateObject("Scripting.Dictionary")
    ClearFileChooserInfo
    Set LifecycleEvents = New Collection

    Set PageInfos = New Collection
    Set IframeInfos = New Collection

    NetworkInFlight = 0
    NetworkLastActivity = Timer()

    DomUpdated = False
    DomLastUpdated = 0
    Set NetworkResponses = New Collection

End Sub

' 保持中のレスポンス一覧をクリアする
Public Sub ClearNetworkResponses()
    Set NetworkResponses = New Collection
End Sub
' 各種イベント発生時の状態を更新する
Public Sub GetInfo(EventInfo As String)

    Dim eventName As String
    eventName = json_.GetValue(EventInfo, "method")

    If InStr(eventName, "frame") > 0 Then
        Dim Id As String

        If InStr(eventName, "Navigated") = 0 Then
            Id = json_.GetValue(EventInfo, "params", "frameId")
        Else
            Id = json_.GetValue(EventInfo, "params", "frame", "id")
        End If

        Dim isCurrentTarget As Boolean
        isCurrentTarget = (Id = basicInfo_.targetId)

        Dim isIsolatedFrameId
        isIsolatedFrameId = (Id = basicInfo_.IsolatedFrameId)
    End If

    Select Case eventName

        Case "Page.javascriptDialogOpening"
            SetDialogInfo EventInfo

        Case "Page.javascriptDialogClosed"
            InitializeDialogInfo
            Set DownloadInfoDic = CreateObject("Scripting.Dictionary")
            ClearDownloadInfo

        Case "Page.fileChooserOpened"
            SetFileChooserInfo EventInfo

        Case "Page.lifecycleEvent"
            AddLifecycleEvent EventInfo

        Case "Target.targetCreated"
            AddTargetInfo EventInfo

        Case "Target.targetInfoChanged"
            ChangeTargetInfo EventInfo

        Case "Target.targetDestroyed"
            RemoveTargetInfo EventInfo

        Case "Page.frameScheduledNavigation" '非推奨となっているがこれが返ってくる。
            If InStr(EventInfo, "anchorClick") > 0 Then
                If isCurrentTarget Then IsCurrentTargetScheduledNavigation = True
                If isIsolatedFrameId Then
                    IsIsolatedFrameScheduledNavigation = True
                End If
            End If

        Case "Page.frameStartedLoading"
            If isCurrentTarget Then IsCurrentTargetStartedLoading = True
            If isIsolatedFrameId Then
                IsIsolatedFrameStartedLoading = True
                Sleep 300
            End If

'        Case "Page.frameClearedScheduledNavigation" '非推奨となっているがこれが返ってくる。
'            If isCurrentTarget Then IsCurrentTargetClearedScheduledNavigation = True
'            If isIsolatedFrameId Then IsIsolatedFrameClearedScheduledNavigation = True

        Case "Page.frameNavigated"
            If isCurrentTarget Then IsCurrentTargetNavigated = True
            If isIsolatedFrameId Then
                IsIsolatedFrameNavigated = True
            End If

        Case "Browser.downloadWillBegin"
            DownloadInfoDic("HasDownload") = True
            DownloadInfoDic("Guid") = json_.GetValue(EventInfo, "params", "guid")
            DownloadInfoDic("SuggestedFilename") = json_.GetValue(EventInfo, "params", "suggestedFilename")
            DownloadInfoDic("Url") = json_.GetValue(EventInfo, "params", "url")
            DownloadInfoDic("State") = "inProgress"

        Case "Page.downloadWillBegin"
            DownloadInfoDic("HasDownload") = True
            DownloadInfoDic("Guid") = json_.GetValue(EventInfo, "params", "guid")
            DownloadInfoDic("SuggestedFilename") = json_.GetValue(EventInfo, "params", "suggestedFilename")
            DownloadInfoDic("Url") = json_.GetValue(EventInfo, "params", "url")
            DownloadInfoDic("State") = "inProgress"

        Case "Browser.downloadProgress"
            DownloadInfoDic("Guid") = json_.GetValue(EventInfo, "params", "guid")
            DownloadInfoDic("State") = json_.GetValue(EventInfo, "params", "state")
            DownloadInfoDic("ReceivedBytes") = json_.GetValue(EventInfo, "params", "receivedBytes")
            DownloadInfoDic("TotalBytes") = json_.GetValue(EventInfo, "params", "totalBytes")
            If InStr(EventInfo, "filePath") > 0 Then
                DownloadInfoDic("FilePath") = json_.GetValue(EventInfo, "params", "filePath")
            End If

        Case "Page.downloadProgress"
            DownloadInfoDic("Guid") = json_.GetValue(EventInfo, "params", "guid")
            DownloadInfoDic("State") = json_.GetValue(EventInfo, "params", "state")
            DownloadInfoDic("ReceivedBytes") = json_.GetValue(EventInfo, "params", "receivedBytes")
            DownloadInfoDic("TotalBytes") = json_.GetValue(EventInfo, "params", "totalBytes")
            If InStr(EventInfo, "filePath") > 0 Then
                DownloadInfoDic("FilePath") = json_.GetValue(EventInfo, "params", "filePath")
            End If

        Case "Runtime.executionContextDestroyed"
            Dim contextId As Long
            contextId = json_.GetValue(EventInfo, "params", "executionContextId")
            If contextId = basicInfo_.ExecutionContextId Then
                IsCurrentContextIdDestroyed = True
            End If

        Case "Network.requestWillBeSent"
            NetworkInFlight = NetworkInFlight + 1
            NetworkLastActivity = Timer()

        Case "Network.loadingFinished"
            If NetworkInFlight > 0 Then NetworkInFlight = NetworkInFlight - 1
            NetworkLastActivity = Timer()

        Case "Network.loadingFailed"
            If NetworkInFlight > 0 Then NetworkInFlight = NetworkInFlight - 1
            NetworkLastActivity = Timer()

        Case "Network.requestServedFromCache"
            If NetworkInFlight > 0 Then NetworkInFlight = NetworkInFlight - 1
            NetworkLastActivity = Timer()

        Case "DOM.documentUpdated"
            DomUpdated = True
            DomLastUpdated = Timer()

        Case "Network.responseReceived"
            Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
            d.Add "RequestId", CStr(json_.GetValue(EventInfo, "params", "requestId"))
            d.Add "Url", CStr(json_.GetValue(EventInfo, "params", "response", "url"))
            d.Add "Status", CLng(json_.GetValue(EventInfo, "params", "response", "status"))
            d.Add "MimeType", CStr(json_.GetValue(EventInfo, "params", "response", "mimeType"))
            NetworkResponses.Add d
    End Select
End Sub
' アラートダイアログ情報を反映する
Private Sub SetDialogInfo(ByVal EventInfo As String)
    DialogInfoDic("IsExistDialog") = True
    DialogInfoDic("DialogMessage") = json_.GetValue(EventInfo, "params", "message")
    DialogInfoDic("DialogType") = json_.GetValue(EventInfo, "params", "type")
    If json_.GetValue(EventInfo, "params", "hasBrowserHandler") = "True" Then
        DialogInfoDic("HasBrowserHandler") = True
    Else
        DialogInfoDic("HasBrowserHandler") = False
    End If
    DialogInfoDic("DefaultPrompt") = json_.GetValue(EventInfo, "params", "defaultPrompt")
End Sub

' ダイアログ情報辞書を初期化する
Private Sub InitializeDialogInfo()
    DialogInfoDic.RemoveAll
    DialogInfoDic.Add "IsExistDialog", False
    DialogInfoDic.Add "DialogMessage", ""
    DialogInfoDic.Add "DialogType", ""
    DialogInfoDic.Add "HasBrowserHandler", False
    DialogInfoDic.Add "DefaultPrompt", ""
End Sub

' Target 種別ごとに管理対象コレクションへ追加する
Private Sub AddTargetInfo(ByVal EventInfo As String)
    Dim tempDic As Object
    Set tempDic = CreateObject("Scripting.Dictionary")
    tempDic.Add "targetId", json_.GetValue(EventInfo, "params", "targetInfo", "targetId")
    tempDic.Add "title", json_.GetValue(EventInfo, "params", "targetInfo", "title")
    tempDic.Add "url", json_.GetValue(EventInfo, "params", "targetInfo", "url")
    tempDic.Add "browserContextId", json_.GetValue(EventInfo, "params", "targetInfo", "browserContextId")

    Dim targetType As String
    targetType = json_.GetValue(EventInfo, "params", "targetInfo", "type")
    Select Case targetType
      Case "page"
        AddInfoOf PageInfos, tempDic, EventInfo
      Case "iframe"
        AddInfoOf IframeInfos, tempDic, EventInfo
    End Select
End Sub
' Target 情報をコレクションへ追加または既存情報を更新する
Private Sub AddInfoOf(infos As Collection, tempDic As Object, _
                                    ByVal EventInfo As String)

    ' targetCreated で更新相当の内容が届くケースを吸収する
    ' 既存ページも targetCreated として届くケースを吸収する
    Dim i As Long
    For i = 1 To infos.count
        If infos(i)("targetId") = tempDic("targetId") Then
            infos(i)("title") = tempDic("title")
            infos(i)("url") = tempDic("url")
            infos(i)("browserContextId") = tempDic("browserContextId")
            Exit Sub
        End If
    Next
    infos.Add tempDic
End Sub

' 既存 Target 情報の変更を反映する
Private Sub ChangeTargetInfo(ByVal EventInfo As String)
    Dim targetType As String
    targetType = json_.GetValue(EventInfo, "params", "targetInfo", "type")

    Select Case targetType
      Case "page"
        ChangeInfoOf PageInfos, EventInfo
      Case "iframe"
        ChangeInfoOf IframeInfos, EventInfo
    End Select
End Sub

' 対象 Target の保持情報を更新する
Private Sub ChangeInfoOf(infos As Collection, _
                                        ByVal EventInfo As String)
    Dim targetId As String
    targetId = json_.GetValue(EventInfo, "params", "targetInfo", "targetId")
    Dim info As Variant
    For Each info In infos
        If info("targetId") = targetId Then
            info("title") = json_.GetValue(EventInfo, "params", "targetInfo", "title")
            info("url") = json_.GetValue(EventInfo, "params", "targetInfo", "url")
            info("browserContextId") = json_.GetValue(EventInfo, "params", "targetInfo", "browserContextId")
        End If
    Next
End Sub

' 削除された Target 情報をコレクションから除去する
Private Sub RemoveTargetInfo(ByVal EventInfo As String)
    Dim targetType As String
    targetType = json_.GetValue(EventInfo, "params", "targetInfo", "type")

    Select Case targetType
      Case "page"
        RemoveInfoOf PageInfos, EventInfo
      Case "iframe"
        RemoveInfoOf IframeInfos, EventInfo
    End Select
End Sub

' 対象 Target の保持情報を削除する
Private Sub RemoveInfoOf(infos As Collection, _
                                        ByVal EventInfo As String)
    Dim targetId As String
    targetId = json_.GetValue(EventInfo, "params", "targetInfo", "targetId")

    Dim i As Long
    For i = 1 To infos.count
        If infos(i)("targetId") = targetId Then
            infos.Remove i
        End If
    Next
End Sub

' 保持中の Target 情報をデバッグ出力する
Public Sub ForDebug_ShowtargetInfo()
    Dim i  As Long
    'Debug.Print "PageInfos↓==============="
    For i = 1 To PageInfos.count
        'Debug.Print i & "：" & PageInfos(i)("targetId")
    Next
    'Debug.Print "IframeInfos↓==============="
    For i = 1 To IframeInfos.count
        'Debug.Print i & "：" & IframeInfos(i)("targetId")
    Next
    'Debug.Print "========================"
End Sub

' ダウンロード状態を初期化する
Public Sub ClearDownloadInfo()
 DownloadInfoDic.RemoveAll
 DownloadInfoDic.Add "HasDownload", False
 DownloadInfoDic.Add "Guid", ""
 DownloadInfoDic.Add "SuggestedFilename", ""
 DownloadInfoDic.Add "Url", ""
 DownloadInfoDic.Add "State", ""
 DownloadInfoDic.Add "FilePath", ""
 DownloadInfoDic.Add "ReceivedBytes", 0
 DownloadInfoDic.Add "TotalBytes", 0
End Sub

' ファイル選択ダイアログ情報を初期化する
Public Sub ClearFileChooserInfo()
    If FileChooserInfoDic Is Nothing Then
        Set FileChooserInfoDic = CreateObject("Scripting.Dictionary")
    End If

    FileChooserInfoDic.RemoveAll
    FileChooserInfoDic.Add "BackendNodeId", 0
    FileChooserInfoDic.Add "Mode", ""
End Sub

' ライフサイクルイベント履歴をクリアする
Public Sub ClearLifecycleEvents()
    Set LifecycleEvents = New Collection
End Sub

' ファイル選択ダイアログ情報を反映する
Private Sub SetFileChooserInfo(EventInfo As String)
    If FileChooserInfoDic Is Nothing Then
        Set FileChooserInfoDic = CreateObject("Scripting.Dictionary")
    End If

    FileChooserInfoDic("BackendNodeId") = CLng(json_.GetValue(EventInfo, "params", "backendNodeId"))
    FileChooserInfoDic("Mode") = CStr(json_.GetValue(EventInfo, "params", "mode"))
End Sub

' ライフサイクルイベントを履歴へ追加する
Private Sub AddLifecycleEvent(EventInfo As String)
    If LifecycleEvents Is Nothing Then Set LifecycleEvents = New Collection

    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    dic.Add "Name", CStr(json_.GetValue(EventInfo, "params", "name"))
    dic.Add "FrameId", CStr(json_.GetValue(EventInfo, "params", "frameId"))
    dic.Add "Timestamp", CStr(json_.GetValue(EventInfo, "params", "timestamp"))
    LifecycleEvents.Add dic
End Sub


