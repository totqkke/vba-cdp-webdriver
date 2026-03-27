VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWebDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 画像形式を表す列挙体
Public Enum Image
    png = 1
    jpeg
    webp
End Enum

' ページソース取得対象を表す列挙体
Public Enum TargetIs
    html_outerhtml = 1
    html_innerhtml
    body_outerhtml
    body_innerhtml
    body_innertext
End Enum

' Windowサイズの列挙体
Public Enum WindowSize
    normal = 1
    minimize
    maximize
End Enum

' 現在のページ URL を返す
Public Property Get URL() As String
End Property

Public Property Get QueryString() As Object
End Property

Public Function GetValueFromQueryString(KeyName As String) As String
End Function

Public Function GetSecureRandomNumber(count As Long) As Variant
End Function

Public Function GetURLSafeRundomString(charNum As Long) As String
End Function

Public Function GetBase64URLEncodedSHA256(str As String) As String
End Function

Public Function EncodeURIConpornent(str As String) As String
End Function

' 指定 URL を開く
Public Sub OpenURL(targetURL As String)
End Sub

' JavaScript コードを実行する
Public Function ExecuteScript(script As String) As String
End Function

' Id 指定で要素を取得する
Public Function FindElementById(Id As String) As IWebElement
End Function

' XPath 指定で要素を取得する
Public Function FindElementByXPath(xPath As String) As IWebElement
End Function

' Name 指定で要素を取得する
Public Function FindElementByName(name As String) As IWebElement
End Function

' タグ名指定で要素コレクションを取得する
Public Function FindElementsByTag(tagName As String) As IWebElements
End Function

' XPath 指定で要素コレクションを取得する
Public Function FindElementsByXPath(xPath As String) As IWebElements
End Function

' クラス名指定で要素コレクションを取得する
Public Function FindElementsByClassName(className As String) As IWebElements
End Function

' XPath 指定要素の存在有無を判定する
Public Function IsExistsElementByXpath(xPath As String) As Boolean
End Function

' Id 指定要素の存在有無を判定する
Public Function IsExistsElementById(Id As String) As Boolean
End Function

' Name 指定要素の存在有無を判定する
Public Function IsExistsElementByName(name As String) As Boolean
End Function

' XPath 指定要素の出現を待機する
Public Sub WaitForElementToAppearByXpath(xPath As String, maxWaitTime_sec As Long)
End Sub

' Id 指定要素の出現を待機する
Public Sub WaitForElementToAppearById(Id As String, maxWaitTime_sec As Long)
End Sub

' Name 指定要素の出現を待機する
Public Sub WaitForElementToAppearByName(name As String, maxWaitTime_sec As Long)
End Sub

' Alert ダイアログを受諾して閉じる
Public Sub HandleAlertDialog()
End Sub

' WinAPI の Sleep を呼び出す
Public Sub SleepByWinAPI(milliSecond As Long)
End Sub

' Index 指定でタブを切り替える
Public Sub SwitchTabByIndex(index As Long)
End Sub

' 指定値に一致するフレームへ切り替える
Public Sub SwitchFrameByNameOrURLOrIndex(value As Variant)
End Sub

' iframe 要素を基準にフレームを切り替える
Public Sub SwitchFrameByIframeElement(IframeElement As IWebElement)
End Sub

' タイトル一致でタブを切り替える
Public Sub SwitchTabByTitle(tabTitle As String)
End Sub

' 現在開いているタブタイトル一覧を返す
Public Function GetTabList() As String
End Function
' 現在のタブ数を返す
Public Function GetTabCount() As Long
End Function
' 指定 index のタブを閉じる
Public Sub CloseTab(index As Long)
End Sub

' 指定タイトルのタブを閉じる
Public Sub CloseTabByTitle(tabTitle As String)
End Sub

' デフォルトフレームへ戻す
Public Sub SwitchFrameToDefault()
End Sub

' 切り替え可能なフレーム情報をデバッグ出力する
Public Sub DebugPrintAccessibleFrameInfo()
End Sub

' 現在セッションのウィンドウを閉じる
Public Sub CloseWindow()
End Sub

' ウィンドウを指定量スクロールする
Public Sub ScrollWindow(yPixel As Long, Optional xPixel As Long = 0)
End Sub

' ブラウザを前面化する
Public Sub SetForeGroundWindow()
End Sub

' 指定対象のページソースを返す
Public Function PageSource(target As TargetIs) As String
End Function

' ブラウザを戻る操作する
Public Sub Back()
End Sub

' ブラウザを進む操作する
Public Sub Forward()
End Sub

' スクリーンショットを保存して保存先パスを返す
Public Function ScreenShotSaveAsFile(Optional folderPath As String = "", Optional fileName As String = "", Optional imageFormat As Image = Image.png) As String
End Function

Public Function ScreenShotPasteToSheet( _
 left As Single, top As Single, Optional width As Single = 0, Optional height As Single = 0, _
 Optional sheetName As String = "", _
 Optional imageFormat As Image = Image.png) As Shape
End Function

Public Sub ShowWindow(size As WindowSize)
End Sub

' 現在ページを PDF 保存して保存先パスを返す
Public Function PrintToPDFSaveAsFile(Optional folderPath As String = "", Optional fileName As String = "", Optional printBackground As Boolean = True, Optional landscape As Boolean = False) As String
End Function

' 新規ドキュメント評価用スクリプトを登録する
Public Function AddScriptToEvaluateOnNewDocument(source As String, Optional worldName As String = "", Optional includeCommandLineAPI As Boolean = False, Optional runImmediately As Boolean = False) As String
End Function

' 登録済みスクリプトを削除する
Public Sub RemoveScriptToEvaluateOnNewDocument(identifier As String)
End Sub

' ライフサイクルイベント通知を設定する
Public Sub SetLifecycleEventsEnabled(enabled As Boolean)
End Sub

' ファイル選択ダイアログ割り込みを設定する
Public Sub SetInterceptFileChooserDialog(enabled As Boolean)
End Sub

' 保持中のライフサイクルイベント一覧を返す
Public Function GetLifecycleEvents() As Collection
End Function

' ライフサイクルイベント一覧をクリアする
Public Sub ClearLifecycleEvents()
End Sub

' FileChooser の BackendNodeId を返す
Public Function GetFileChooserBackendNodeId() As Long
End Function

' FileChooser の mode を返す
Public Function GetFileChooserMode() As String
End Function

' FileChooser 情報をクリアする
Public Sub ClearFileChooserInfo()
End Sub

' ダウンロード監視を開始する
Public Sub DownloadWatchStart(downloadDir As String, Optional allowAndName As Boolean = False)
End Sub

' ダウンロード完了まで待機する
Public Function DownloadWatchWaitCompleted(maxWaitTime_sec As Long) As Boolean
End Function

' 提案ファイル名を返す
Public Function DownloadWatchSuggestedFilename() As String
End Function

' ダウンロードファイルパスを返す
Public Function DownloadWatchFilePath() As String
End Function

' FileChooser にファイルを設定する
Public Sub FileChooserSetFile(filePath As String, Optional maxWaitTime_sec As Long = 10)
End Sub

' DOM 変化が発生するまで待機する
Public Sub WaitForDomMutation(maxWaitTime_sec As Long)
End Sub

' Network idle になるまで待機する
Public Sub WaitForNetworkIdle(maxWaitTime_sec As Long)
End Sub

' Network ドメインを無効化する
Public Sub NetworkDisable()
End Sub

' 保持中のレスポンス一覧をクリアする
Public Sub ClearNetworkResponses()
End Sub

' 指定 URL を含むレスポンス到着を待機する
Public Function WaitForResponse(urlContains As String, timeoutSec As Long) As String
End Function

' requestId に対応するレスポンス本文を返す
Public Function GetResponseBody(requestId As String) As String
End Function

' Cookies (Storage domain)
Public Function GetCookies() As String
End Function

' Cookie 1 件を設定する
Public Function SetCookie( _
    name As String, _
    value As String, _
    Optional URL As String = "", _
    Optional domain As String = "", _
    Optional path As String = "/", _
    Optional secure As Boolean = False, _
    Optional httpOnly As Boolean = False, _
    Optional sameSite As String = "", _
    Optional expires As Double = 0 _
) As Boolean
End Function

' Cookie 配列を設定する
Public Function SetCookies(cookieParamsJsonArray As String) As Boolean
End Function

' Cookie を削除する
Public Sub ClearCookies()
End Sub

' LocalStorage / SessionStorage (DOMStorage domain)
Public Function LocalStorageGetAll() As String
End Function

' LocalStorage の指定値を取得する
Public Function LocalStorageGetItem(key As String) As String
End Function

' LocalStorage へ項目を設定する
Public Sub LocalStorageSetItem(key As String, value As String)
End Sub

' LocalStorage から項目を削除する
Public Sub LocalStorageRemoveItem(key As String)
End Sub

' LocalStorage をクリアする
Public Sub LocalStorageClear()
End Sub

' SessionStorage 全項目を取得する
Public Function SessionStorageGetAll() As String
End Function

' SessionStorage の指定値を取得する
Public Function SessionStorageGetItem(key As String) As String
End Function

' SessionStorage へ項目を設定する
Public Sub SessionStorageSetItem(key As String, value As String)
End Sub

' SessionStorage から項目を削除する
Public Sub SessionStorageRemoveItem(key As String)
End Sub

' SessionStorage をクリアする
Public Sub SessionStorageClear()
End Sub

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

' Alert / Confirm / Prompt を受諾する
Public Function AcceptAlertDialog(Optional promptText As String = "") As String
End Function

' Alert / Confirm / Prompt を却下する
Public Function DismissAlertDialog(Optional promptText As String = "") As String
End Function

' ネットワーク割り込みスクリプトを有効化する
Public Sub EnableNetworkInterception()
End Sub

' ネットワーク割り込みスクリプトを無効化する
Public Sub DisableNetworkInterception()
End Sub

' 遮断対象 URL パターンを追加する
Public Sub AddBlockedURLPattern(urlPattern As String)
End Sub

' 遮断対象 URL パターンをクリアする
Public Sub ClearBlockedURLPatterns()
End Sub

' モックレスポンスを追加する
Public Sub AddMockResponse( _
    urlPattern As String, _
    statusCode As Long, _
    body As String, _
    Optional contentType As String = "text/plain" _
)
End Sub

'第1引数のCDPメソッドのコードの実行
Public Function ExecuteCDP(CDPMethodJson As String) As String
End Function

' モックレスポンス一覧をクリアする
Public Sub ClearMockResponses()
End Sub

