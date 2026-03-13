Attribute VB_Name = "SampleModule"
Option Explicit

'========================================================
' 共通ヘルパ
'========================================================

' EdgeDriver / ChromeDriver の切り替えポイント
Private Function NewDriver() As IWebDriver
    Dim drv As IWebDriver
    Set drv = New ChromeDriver   ' ChromeDriver にしたい場合はここを差し替え
    Set NewDriver = drv
End Function
'一括実行
Public Sub RunAllSamples()
    Dim currentProc As String
    
    On Error GoTo EH
    
    currentProc = "Sample_01_Basic_Open_And_PageSource"
    Debug.Print "[START] " & currentProc
    Sample_01_Basic_Open_And_PageSource
    Debug.Print "[OK] " & currentProc
    
    currentProc = "Sample_02_Find_Input_Click"
    Debug.Print "[START] " & currentProc
    Sample_02_Find_Input_Click
    Debug.Print "[OK] " & currentProc
    
    currentProc = "Sample_03_FindElements_Class_And_XPath"
    Debug.Print "[START] " & currentProc
    Sample_03_FindElements_Class_And_XPath
    Debug.Print "[OK] " & currentProc
    
    currentProc = "Sample_04_Attribute_Property_Checked"
    Debug.Print "[START] " & currentProc
    Sample_04_Attribute_Property_Checked
    Debug.Print "[OK] " & currentProc
    
    currentProc = "Sample_05_Waits"
    Debug.Print "[START] " & currentProc
    Sample_05_Waits
    Debug.Print "[OK] " & currentProc
    
    currentProc = "Sample_06_Frame_And_Tab"
    Debug.Print "[START] " & currentProc
    Sample_06_Frame_And_Tab
    Debug.Print "[OK] " & currentProc
    
    currentProc = "Sample_07_Alert"
    Debug.Print "[START] " & currentProc
    Sample_07_Alert
    Debug.Print "[OK] " & currentProc
    
    currentProc = "Sample_08_Download"
    Debug.Print "[START] " & currentProc
    Sample_08_Download
    Debug.Print "[OK] " & currentProc
    
    currentProc = "Sample_08B_Download_And_SaveAs"
    Debug.Print "[START] " & currentProc
    Sample_08B_Download_And_SaveAs
    Debug.Print "[OK] " & currentProc
    
    currentProc = "Sample_09_FileChooser"
    Debug.Print "[START] " & currentProc
    Sample_09_FileChooser
    Debug.Print "[OK] " & currentProc
    
    currentProc = "Sample_10_Network_Response"
    Debug.Print "[START] " & currentProc
    Sample_10_Network_Response
    Debug.Print "[OK] " & currentProc
    
    currentProc = "Sample_11_Screenshot_And_Pdf"
    Debug.Print "[START] " & currentProc
    Sample_11_Screenshot_And_Pdf
    Debug.Print "[OK] " & currentProc
    
    currentProc = "Sample_12_Cookies_And_Storage"
    Debug.Print "[START] " & currentProc
    Sample_12_Cookies_And_Storage
    Debug.Print "[OK] " & currentProc
    
    currentProc = "Sample_13_LifecycleEvents"
    Debug.Print "[START] " & currentProc
    Sample_13_LifecycleEvents
    Debug.Print "[OK] " & currentProc
    
    currentProc = "Sample_14_CloseTab_ByIndex"
    Debug.Print "[START] " & currentProc
    Sample_14_CloseTab_ByIndex
    Debug.Print "[OK] " & currentProc
    
    currentProc = "Sample_15_CloseTabByTitle_Basic"
    Debug.Print "[START] " & currentProc
    Sample_15_CloseTabByTitle_Basic
    Debug.Print "[OK] " & currentProc
    
    currentProc = "Sample_16_Switch_Then_Close_ByTitle"
    Debug.Print "[START] " & currentProc
    Sample_16_Switch_Then_Close_ByTitle
    Debug.Print "[OK] " & currentProc
    
    currentProc = "Sample_17_CloseTabs_OneByOne"
    Debug.Print "[START] " & currentProc
    Sample_17_CloseTabs_OneByOne
    Debug.Print "[OK] " & currentProc

    currentProc = "Sample_18_Find_By_Css"
    Debug.Print "[START] " & currentProc
    Sample_18_Find_By_Css
    Debug.Print "[OK] " & currentProc

    currentProc = "Sample_19_Find_By_Text"
    Debug.Print "[START] " & currentProc
    Sample_19_Find_By_Text
    Debug.Print "[OK] " & currentProc

    currentProc = "Sample_20_Deep_Search_By_Driver"
    Debug.Print "[START] " & currentProc
    Sample_20_Deep_Search_By_Driver
    Debug.Print "[OK] " & currentProc

    currentProc = "Sample_21_Element_Child_Locators"
    Debug.Print "[START] " & currentProc
    Sample_21_Element_Child_Locators
    Debug.Print "[OK] " & currentProc

    currentProc = "Sample_22_Element_Deep_Locators"
    Debug.Print "[START] " & currentProc
    Sample_22_Element_Deep_Locators
    Debug.Print "[OK] " & currentProc

    currentProc = "Sample_23_Dialog_Control"
    Debug.Print "[START] " & currentProc
    Sample_23_Dialog_Control
    Debug.Print "[OK] " & currentProc

    currentProc = "Sample_24_Network_Block"
    Debug.Print "[START] " & currentProc
    Sample_24_Network_Block
    Debug.Print "[OK] " & currentProc

    currentProc = "Sample_25_Network_Mock"
    Debug.Print "[START] " & currentProc
    Sample_25_Network_Mock
    Debug.Print "[OK] " & currentProc
    
    currentProc = "Sample_26_TargetId_RelatedAPIs"
    Debug.Print "[START] " & currentProc
    Sample_26_TargetId_RelatedAPIs
    Debug.Print "[OK] " & currentProc

    currentProc = "Sample_27_FrameId_RelatedAPIs"
    Debug.Print "[START] " & currentProc
    Sample_27_FrameId_RelatedAPIs
    Debug.Print "[OK] " & currentProc

    currentProc = "Sample_28_ObjectId_RelatedAPIs"
    Debug.Print "[START] " & currentProc
    Sample_28_ObjectId_RelatedAPIs
    Debug.Print "[OK] " & currentProc

    currentProc = "Sample_29_RequestId_RelatedAPIs"
    Debug.Print "[START] " & currentProc
    Sample_29_RequestId_RelatedAPIs
    Debug.Print "[OK] " & currentProc

    currentProc = "Sample_30_ShadowDom_Deep_By_Driver"
    Debug.Print "[START] " & currentProc
    Sample_30_ShadowDom_Deep_By_Driver
    Debug.Print "[OK] " & currentProc

    currentProc = "Sample_31_ShadowDom_Deep_By_Element"
    Debug.Print "[START] " & currentProc
    Sample_31_ShadowDom_Deep_By_Element
    Debug.Print "[OK] " & currentProc
    
    currentProc = "Sample_33_WebElementsToTextArray1D"
    Debug.Print "[START] " & currentProc
    Sample_33_WebElementsToTextArray1D
    Debug.Print "[OK] " & currentProc

    currentProc = "Sample_34_WebElementsToAttributeArray1D"
    Debug.Print "[START] " & currentProc
    Sample_34_WebElementsToAttributeArray1D
    Debug.Print "[OK] " & currentProc

    currentProc = "Sample_35_JoinWebElementsText"
    Debug.Print "[START] " & currentProc
    Sample_35_JoinWebElementsText
    Debug.Print "[OK] " & currentProc

    currentProc = "Sample_36_TableElementToTSV"
    Debug.Print "[START] " & currentProc
    Sample_36_TableElementToTSV
    Debug.Print "[OK] " & currentProc

    currentProc = "Sample_37_Array2DToRange"
    Debug.Print "[START] " & currentProc
    Sample_37_Array2DToRange
    Debug.Print "[OK] " & currentProc

    currentProc = "Sample_38_QueryString_Basic"
    Debug.Print "[START] " & currentProc
    Sample_38_QueryString_Basic
    Debug.Print "[OK] " & currentProc

    currentProc = "Sample_39_Utility_PublicAPIs_Basic"
    Debug.Print "[START] " & currentProc
    Sample_39_Utility_PublicAPIs_Basic
    Debug.Print "[OK] " & currentProc

    currentProc = "Sample_40_Window_And_PasteShot_Basic"
    Debug.Print "[START] " & currentProc
    Sample_40_Window_And_PasteShot_Basic
    Debug.Print "[OK] " & currentProc

    currentProc = "Sample_41_SelectBox_APIs"
    Debug.Print "[START] " & currentProc
    Sample_41_SelectBox_APIs
    Debug.Print "[OK] " & currentProc

    currentProc = "Sample_42_Input_Helper_APIs"
    Debug.Print "[START] " & currentProc
    Sample_42_Input_Helper_APIs
    Debug.Print "[OK] " & currentProc

    currentProc = "Sample_43_Dispatch_Event_APIs"
    Debug.Print "[START] " & currentProc
    Sample_43_Dispatch_Event_APIs

    Debug.Print "[DONE] RunAllSamples"
    Exit Sub

EH:
    Err.Raise Err.Number, "RunAllSamples/" & currentProc & "/" & Err.source, Err.Description
End Sub

'========================================================
' 01. 基本操作
' 何ができるか:
'   - ブラウザ起動
'   - URLを開く
'   - JavaScript実行
'   - 現在URL / ページソース取得
'========================================================
'========================================================
' 01. 基本操作
' 何ができるか:
' - ブラウザを起動して指定URLを開ける
' - 現在開いているURLを取得できる
' - JavaScriptをその場で実行できる
' - body の HTML を取得できる
'
' このサンプルの確認ポイント:
' - OpenURL でページを開けること
' - URL プロパティで現在URLが取得できること
' - ExecuteScript で document.title を取れること
' - PageSource で body の HTML を取れること
'========================================================
Public Sub Sample_01_Basic_Open_And_PageSource()
    Dim drv As IWebDriver
    Set drv = NewDriver()

    ' まず外部サイトを1ページ開く
    ' 一番基本的な「ブラウザを開いてページを表示する」操作
    drv.OpenURL "https://example.com"

    ' 現在URLを取得する
    ' OpenURL で開いた先が drv.URL で参照できることを確認する
    Debug.Print "URL=" & drv.URL

    ' JavaScript を直接実行する
    ' ここでは title を取得して、スクリプト実行の基本動作を確認する
    Debug.Print "Title=" & drv.ExecuteScript("document.title;")

    ' body 内部の HTML を取得する
    ' body 全体が欲しい場合は body_outerhtml を使う
    ' このサンプルでは body_innerhtml を使って中身だけ確認する
    Debug.Print "Body HTML=" & drv.PageSource(TargetIs.body_innerhtml)

    drv.CloseWindow
End Sub

'========================================================
' 02. 要素取得・入力・クリック
' 何ができるか:
' - ID で要素を取得できる
' - XPath で要素を取得できる
' - input に文字入力できる
' - button をクリックできる
' - 結果要素の textContent を読める
'
' このサンプルの確認ポイント:
' - FindElementById → SetValue が動くこと
' - FindElementById → Click が動くこと
' - XPath で結果要素を取得して文字列確認できること
'========================================================
Public Sub Sample_02_Find_Input_Click()
    Dim drv As IWebDriver
    Set drv = NewDriver()

    ' data URL で簡単なテストページを生成する
    ' txt1 に値を入れて Run ボタンを押すと out に結果が出る
    drv.OpenURL "data:text/html," & _
                "<html><body>" & _
                "<input id='txt1' name='txt1' value=''>" & _
                "<button id='btn1' onclick=""document.getElementById('out').textContent=document.getElementById('txt1').value;"">Run</button>" & _
                "<div id='out'></div>" & _
                "</body></html>"

    ' IDで入力欄を取得し、値を設定する
    ' ここでは txt1 に hello を入れる
    drv.FindElementById("txt1").SetValue "hello"

    ' IDでボタンを取得してクリックする
    ' クリック後、ページ内の JavaScript が走って out に文字列が入る
    drv.FindElementById("btn1").Click

    ' XPath で結果領域を取得し、画面反映結果を確認する
    Debug.Print "OUT=" & drv.FindElementByXPath("//div[@id=""out""]").GetTextContent

    drv.CloseWindow
End Sub

'========================================================
' 03. 複数要素取得
' 何ができるか:
' - class 名で複数要素を取得できる
' - XPath で複数要素を取得できる
' - For Each で順番に要素を読める
'
' このサンプルの確認ポイント:
' - FindElementsByClassName が複数件返すこと
' - FindElementsByXPath が複数件返すこと
' - それぞれの要素から textContent を読めること
'========================================================
Public Sub Sample_03_FindElements_Class_And_XPath()
    Dim drv As IWebDriver
    Set drv = NewDriver()

    ' class="c" の div と、class="r" の li を用意する
    ' 複数要素取得の動作確認専用ページ
    drv.OpenURL "data:text/html," & _
                "<html><body>" & _
                "<div class='c'>A</div><div class='c'>B</div><div class='c'>C</div>" & _
                "<ul><li class='r'>R1</li><li class='r'>R2</li><li class='r'>R3</li></ul>" & _
                "</body></html>"

    Dim elems1 As IWebElements
    Dim elems2 As IWebElements
    Dim e As IWebElement

    ' ClassName で複数取得する
    ' .c 相当の要素が3件取れる想定
    Set elems1 = drv.FindElementsByClassName("c")
    For Each e In elems1
        Debug.Print "class=c -> " & e.GetTextContent
    Next

    ' XPath で複数取得する
    ' li[@class='r'] が3件取れる想定
    Set elems2 = drv.FindElementsByXPath("//li[@class=""r""]")
    For Each e In elems2
        Debug.Print "xpath=//li[@class=""r""] -> " & e.GetTextContent
    Next

    drv.CloseWindow
End Sub

'========================================================
' 04. 属性 / プロパティ / チェック状態
' 何ができるか:
' - Checked の読み書きができる
' - GetAttribute / SetAttribute が使える
' - GetProperty / SetProperty が使える
'
' このサンプルの確認ポイント:
' - checkbox の Checked を True にして読めること
' - 属性の追加・取得ができること
' - value プロパティの取得・更新ができること
'========================================================
Public Sub Sample_04_Attribute_Property_Checked()
    Dim drv As IWebDriver
    Set drv = NewDriver()

    ' checkbox と text input を1個ずつ置く
    drv.OpenURL "data:text/html," & _
                "<html><body>" & _
                "<input id='cb1' type='checkbox'>" & _
                "<input id='txt1' value='abc'>" & _
                "</body></html>"

    Dim cb As IWebElement
    Dim txt As IWebElement

    Set cb = drv.FindElementById("cb1")
    Set txt = drv.FindElementById("txt1")

    ' Checked プロパティを書き込む
    ' checkbox がチェック状態になることを確認する
    cb.Checked = True

    ' Checked を読み戻して確認する
    Debug.Print "Checked=" & cb.Checked

    ' HTML 属性 value を読む
    ' input の初期値 abc が取得できることを確認する
    Debug.Print "txt.value(attr)=" & txt.GetAttribute("value")

    ' 独自属性 data-test を追加する
    ' 属性の追記ができることを確認する
    Call txt.SetAttribute("data-test", "sample")
    Debug.Print "txt.data-test=" & txt.GetAttribute("data-test")

    ' DOM プロパティ value を読む
    ' 属性取得との違い確認用
    Debug.Print "txt.value(prop)=" & CStr(txt.GetProperty("value"))

    ' DOM プロパティ value を変更する
    Call txt.SetProperty("value", "xyz")
    Debug.Print "txt.value(prop after set)=" & CStr(txt.GetProperty("value"))

    drv.CloseWindow
End Sub

'========================================================
' 05. 待機系
' 何ができるか:
' - 要素の表示待ちができる
' - 要素の有効化待ちができる
' - DOM変化待ちができる
' - Network Idle 待ちができる
'
' このサンプルの確認ポイント:
' - WaitUntilVisible が動くこと
' - WaitUntilEnabled が動くこと
' - WaitForDomMutation が動くこと
' - WaitForNetworkIdle が動くこと
'========================================================
Public Sub Sample_05_Waits()
    Dim drv As IWebDriver
    Set drv = NewDriver()

    ' 最初は非表示かつ disabled のボタンを置く
    drv.OpenURL "data:text/html," & _
                "<html><body>" & _
                "<button id='btn1' disabled style='display:none'>go</button>" & _
                "</body></html>"

    ' 1秒後に表示し、disabled を外す
    ' 表示待ち / 有効待ち の確認用
    drv.ExecuteScript "setTimeout(function(){" & _
                      "var b=document.getElementById('btn1');" & _
                      "b.style.display='block';" & _
                      "b.disabled=false;" & _
                      "},1000);"

    Dim btn As IWebElement
    Set btn = drv.FindElementById("btn1")

    ' 見える状態になるまで待つ
    btn.WaitUntilVisible 5

    ' 有効状態になるまで待つ
    btn.WaitUntilEnabled 5

    ' DOM変化待ち:
    ' 1秒後に div を1個追加して、DOM更新イベント待ちを確認する
    drv.ExecuteScript "setTimeout(function(){document.body.appendChild(document.createElement('div'));},1000);"
    drv.WaitForDomMutation 5

    ' Network Idle待ち:
    ' 外部ページへ移動してから fetch を1回発生させ、
    ' 通信が落ち着くまで待つ
    drv.OpenURL "https://example.com"
    drv.ExecuteScript "try{fetch('https://example.com/?t='+Date.now()).catch(function(){});}catch(e){}"
    drv.WaitForNetworkIdle 10

    drv.CloseWindow
End Sub

'========================================================
' 06. フレーム / タブ
' 何ができるか:
' - iframe へ切り替えられる
' - デフォルトフレームへ戻れる
' - タブ一覧を取得できる
' - タイトル指定でタブ切替できる
'
' このサンプルの確認ポイント:
' - SwitchFrameByNameOrURLOrIndex が動くこと
' - SwitchFrameToDefault が動くこと
' - GetTabList でタブ一覧が取れること
' - SwitchTabByTitle が動くこと
'========================================================
Public Sub Sample_06_Frame_And_Tab()
    Dim drv As IWebDriver
    Set drv = NewDriver()

    ' name=f1 の iframe を1個持つページを開く
    ' iframe 内には id=inner の input を置いておく
    drv.OpenURL "data:text/html," & _
                "<html><body>" & _
                "<iframe name='f1' srcdoc=""<html><body><input id='inner' value='inside'></body></html>""></iframe>" & _
                "</body></html>"

    ' iframe へ切り替える
    drv.SwitchFrameByNameOrURLOrIndex "f1"

    ' iframe 内要素が取れることを確認する
    Debug.Print "inner=" & CStr(drv.FindElementById("inner").GetProperty("value"))

    ' 親フレームへ戻る
    drv.SwitchFrameToDefault

    ' 新しいタブを開く
    drv.ExecuteScript "window.open('data:text/html,<html><head><title>TAB2</title></head><body>tab2</body></html>','_blank');"
    drv.SleepByWinAPI 1000

    ' 現在のタブ一覧を確認する
    Debug.Print "Tabs=" & drv.GetTabList

    ' タイトルが見える環境なら TAB2 に切り替える
    On Error Resume Next
    drv.SwitchTabByTitle "TAB2"
    On Error GoTo 0

    drv.CloseWindow
End Sub

'========================================================
' 07. アラート
' 何ができるか:
' - alert を開いて閉じられる
' - クリック後に制御が戻りにくい alert 系の対処例になる
'
' このサンプルの確認ポイント:
' - ClickAndThenAlertDialogErase が動くこと
' - alert を閉じた後も処理継続できること
'========================================================
Public Sub Sample_07_Alert()
    Dim drv As IWebDriver
    Set drv = NewDriver()

    drv.OpenURL "data:text/html," & _
                "<html><body>" & _
                "<button id='b1' onclick=""alert('hello');"">alert</button>" & _
                "</body></html>"

    ' ボタンを押すと alert が出る
    ' このメソッドはクリックと alert 消去をまとめて行う
    drv.FindElementById("b1").ClickAndThenAlertDialogErase

    drv.CloseWindow
End Sub

'========================================================
' 08. ダウンロード
' 何ができるか:
' - ダウンロード開始を監視できる
' - ダウンロード完了を待てる
' - 実際に保存されたファイル名・保存先を取得できる
'
' このサンプルの確認ポイント:
' - DownloadWatchStart が動くこと
' - DownloadWatchWaitCompleted が完了まで待てること
' - SuggestedFilename / FilePath を取得できること
'========================================================
Public Sub Sample_08_Download()
    Dim drv As IWebDriver
    Set drv = NewDriver()

    Dim outDir As String
    outDir = ThisWorkbook.path & "\download_test"
    EnsureFolderExists outDir

    ' download 属性付きのリンクを用意する
    ' data URL を保存する簡易ダウンロード例
    drv.OpenURL "data:text/html," & _
                "<html><body>" & _
                "<a id='dl' download='hello.txt' href='data:text/plain,hello'>download</a>" & _
                "</body></html>"

    ' ダウンロード監視を開始する
    ' allowAndName=False なのでブラウザ側の保存名のまま保存される
    drv.DownloadWatchStart outDir, False

    ' ダウンロードリンクをクリックする
    drv.FindElementById("dl").Click

    ' 完了まで待つ
    Call drv.DownloadWatchWaitCompleted(30)

    ' ブラウザが提案したファイル名を確認する
    Debug.Print "Suggested=" & drv.DownloadWatchSuggestedFilename

    ' 実際に保存されたファイルパスを確認する
    Debug.Print "FilePath=" & drv.DownloadWatchFilePath

    drv.CloseWindow
End Sub

'========================================================
' 08B. ダウンロードして別名保存
' 何ができるか:
' - ダウンロード完了後に保存名を固定できる
' - DownloadAndSaveAs ヘルパの使用例になる
'
' このサンプルの確認ポイント:
' - ダウンロード後に renamed_hello.txt へリネームできること
'========================================================
Public Sub Sample_08B_Download_And_SaveAs()
    Dim drv As IWebDriver
    Set drv = NewDriver()

    Dim outDir As String
    outDir = ThisWorkbook.path & "\download_test2"
    EnsureFolderExists outDir

    drv.OpenURL "data:text/html," & _
                "<html><body>" & _
                "<a id='dl' download='hello.txt' href='data:text/plain,hello'>download</a>" & _
                "</body></html>"

    ' DownloadAndSaveAs は
    ' 1) ダウンロード監視開始
    ' 2) 要素クリック
    ' 3) 完了待ち
    ' 4) 必要なら別名へ変更
    ' をまとめて行うヘルパ
    Debug.Print "Saved=" & DownloadAndSaveAs(drv, drv.FindElementById("dl"), outDir, "renamed_hello.txt")

    drv.CloseWindow
End Sub

'========================================================
' 09. FileChooser
' 何ができるか:
' - file input のネイティブダイアログを出さずにファイルを設定できる
' - CDP ベースの FileChooser 制御を確認できる
'
' 注意:
' - filePath は実在するファイルに置き換えること
' - ブラウザが非アクティブな状態では file chooser 処理が停止する
'   （ユーザー操作制約のため）
' - 実行中はブラウザを前面にしたままにすること

' このサンプルの確認ポイント:
' - SetInterceptFileChooserDialog が動くこと
' - FileChooserSetFile でファイル設定できること
'========================================================
Public Sub Sample_09_FileChooser()
    Dim drv As IWebDriver
    Set drv = NewDriver()

    ' 必ず実在ファイルへ置き換える
    ' フォルダではなくファイルパスを入れること
    Dim filePath As String
    filePath = ThisWorkbook.path & "\download_test2\sample.txt"

    drv.OpenURL "data:text/html," & _
                "<html><body><input id='f' type='file'></body></html>"

    ' file chooser interception を有効化する
    drv.SetInterceptFileChooserDialog True

    ' file input をクリックして file chooser イベントを発生させる
    drv.FindElementById("f").DispatchCDPMouseEvent LeftClick_XY, Center

    ' ネイティブダイアログを操作せず、CDP 経由でファイルをセットする
    drv.FileChooserSetFile filePath

    drv.CloseWindow
End Sub
'========================================================
' 10. ネットワーク応答待ち
' 何ができるか:
' - URL 部分一致でレスポンス到着待ちができる
' - requestId を取得できる
' - requestId から response body を取得できる
'
' このサンプルの確認ポイント:
' - WaitForResponse が対象レスポンスを捕まえられること
' - GetResponseBody で本文を取得できること
'========================================================
Public Sub Sample_10_Network_Response()
    Dim drv As IWebDriver
    Set drv = NewDriver()

    ' 通常ページを開く
    drv.OpenURL "https://example.com"

    ' 以前の応答記録を消す
    drv.ClearNetworkResponses

    ' api=1 を含む URL へ fetch を投げる
    ' このあと WaitForResponse("api=1") で待ち合わせる
    drv.ExecuteScript "try{fetch('https://example.com/?api=1&ts='+Date.now()).catch(function(){});}catch(e){}"

    Dim rid As String
    rid = drv.WaitForResponse("api=1", 10)

    ' どの requestId が捕まったかを確認
    Debug.Print "RequestId=" & rid

    ' response body を取得して確認
    Debug.Print "Body=" & drv.GetResponseBody(rid)

    drv.CloseWindow
End Sub

'========================================================
' 11. Screenshot / PDF
' 何ができるか:
' - スクリーンショットを画像ファイルとして保存できる
' - 現在ページを PDF 保存できる
'
' このサンプルの確認ポイント:
' - ScreenShotSaveAsFile が保存先パスを返すこと
' - PrintToPDFSaveAsFile が保存先パスを返すこと
'========================================================
Public Sub Sample_11_Screenshot_And_Pdf()
    Dim drv As IWebDriver
    Set drv = NewDriver()

    Dim outDir As String
    outDir = ThisWorkbook.path & "\capture_test"
    EnsureFolderExists outDir

    drv.OpenURL "https://example.com"

    ' スクリーンショットを PNG で保存する
    Debug.Print "Shot=" & drv.ScreenShotSaveAsFile(outDir, "page_shot", Image.png)

    ' PDF として保存する
    Debug.Print "PDF=" & drv.PrintToPDFSaveAsFile(outDir, "page_pdf", True, False)

    drv.CloseWindow
End Sub

'========================================================
' 12. Cookies / LocalStorage / SessionStorage
' 何ができるか:
' - Cookie を設定 / 取得 / 削除できる
' - LocalStorage を読み書きできる
' - SessionStorage を読み書きできる
'
' このサンプルの確認ポイント:
' - SetCookie / GetCookies / ClearCookies が動くこと
' - LocalStorageSetItem / GetItem / GetAll / Clear が動くこと
' - SessionStorageSetItem / GetItem / GetAll / Clear が動くこと
'========================================================
Public Sub Sample_12_Cookies_And_Storage()
    Dim drv As IWebDriver
    Set drv = NewDriver()

    drv.OpenURL "https://example.com"

    ' -----------------------------
    ' Cookie の操作
    ' -----------------------------
    ' まず既存 Cookie を消してクリーンな状態から始める
    drv.ClearCookies

    ' Cookie を1個設定する
    Call drv.SetCookie("vba_cookie", "abc123", "https://example.com", "", "/", False, False, "", 0)

    ' 現在の Cookie 一覧を確認する
    Debug.Print "Cookies=" & drv.GetCookies

    ' 後片付けとして Cookie を消す
    drv.ClearCookies

    ' -----------------------------
    ' LocalStorage の操作
    ' -----------------------------
    ' 一度空にしてから値を書き込む
    drv.LocalStorageClear
    drv.LocalStorageSetItem "k1", "v1"

    ' 単一値取得
    Debug.Print "LocalStorage(k1)=" & drv.LocalStorageGetItem("k1")

    ' 全件取得
    Debug.Print "LocalStorageAll=" & drv.LocalStorageGetAll

    ' -----------------------------
    ' SessionStorage の操作
    ' -----------------------------
    drv.SessionStorageClear
    drv.SessionStorageSetItem "s1", "sv1"

    ' 単一値取得
    Debug.Print "SessionStorage(s1)=" & drv.SessionStorageGetItem("s1")

    ' 全件取得
    Debug.Print "SessionStorageAll=" & drv.SessionStorageGetAll

    drv.CloseWindow
End Sub

'========================================================
' 13. Lifecycle Events
' 何ができるか:
' - lifecycle event の収集を有効化できる
' - 収集済みイベント一覧を取り出せる
' - 収集結果をクリアできる
'
' このサンプルの確認ポイント:
' - SetLifecycleEventsEnabled True で収集が始まること
' - OpenURL 後にイベントが溜まること
' - GetLifecycleEvents で名前 / frameId / timestamp を読めること
'========================================================
Public Sub Sample_13_LifecycleEvents()
    Dim drv As IWebDriver
    Set drv = NewDriver()

    ' 収集済みイベントを一度クリアする
    drv.ClearLifecycleEvents

    ' lifecycle event 収集を有効化する
    drv.SetLifecycleEventsEnabled True

    ' ページ遷移させてイベントを発生させる
    drv.OpenURL "https://example.com"

    Dim evs As Collection
    Dim ev As Variant

    Set evs = drv.GetLifecycleEvents
    For Each ev In evs
        Debug.Print ev("Name"), ev("FrameId"), ev("Timestamp")
    Next

    ' 後片付けとして無効化する
    drv.SetLifecycleEventsEnabled False
    drv.CloseWindow
End Sub

'========================================================
' CloseTab / CloseTabByTitle 用 共通ヘルパ
' 何をしているか:
' - GetTabList からタブ数を数える
' - タイトル一覧の中に特定タイトルがあるかを確認する
' - CloseTab 系サンプル専用の driver を作る
'========================================================

Private Function CountTabsForCloseSample(ByVal tabList As String) As Long
    ' GetTabList はタイトルをカンマ区切りで返す想定なので、
    ' カンマ分割して件数を数える
    If Len(tabList) = 0 Then
        CountTabsForCloseSample = 0
    Else
        CountTabsForCloseSample = UBound(Split(tabList, ",")) + 1
    End If
End Function

Private Function HasTabTitleForCloseSample(ByVal tabList As String, ByVal title As String) As Boolean
    ' タイトル一覧の中に、指定タイトルが残っているかを判定する
    Dim arr As Variant
    Dim i As Long

    If Len(tabList) = 0 Then Exit Function

    arr = Split(tabList, ",")
    For i = LBound(arr) To UBound(arr)
        If CStr(arr(i)) = title Then
            HasTabTitleForCloseSample = True
            Exit Function
        End If
    Next
End Function

Private Function NewDriverForCloseSample() As IWebDriver
    ' CloseTab 系サンプル専用の生成関数
    ' ChromeDriver を使いたい場合はここだけ差し替えればよい
    Dim drv As IWebDriver
    Set drv = New EdgeDriver
    Set NewDriverForCloseSample = drv
End Function

'========================================================
' 14. CloseTab(index)
' 何ができるか:
' - タブ番号を指定して、そのタブだけ閉じられる
'
' 使いどころ:
' - タイトルではなく「左から何番目か」で閉じたい場合
' - 空タイトルタブが混ざる環境でも index 基準で制御したい場合
'
' このサンプルの確認ポイント:
' - GetTabList の件数が減ること
' - 指定 index のタブだけ閉じられること
'========================================================
Public Sub Sample_14_CloseTab_ByIndex()
    Dim drv As IWebDriver
    Set drv = NewDriverForCloseSample()

    ' 1枚目: タイトル付きタブ
    drv.OpenURL "data:text/html,<html><head><title>BASE</title></head><body>base</body></html>"

    ' 2枚目, 3枚目: about:blank を追加
    ' 環境によってはタイトルが空で返るので、
    ' このサンプルでは index 指定で閉じる
    drv.ExecuteScript "window.open('about:blank','_blank');"
    drv.ExecuteScript "window.open('about:blank','_blank');"
    drv.SleepByWinAPI 1000

    Dim tabsBefore As String
    tabsBefore = drv.GetTabList
    Debug.Print "Before=" & tabsBefore
    Debug.Print "BeforeCount=" & CountTabsForCloseSample(tabsBefore)

    ' 左から2番目のタブを閉じる
    drv.CloseTab 2
    drv.SleepByWinAPI 700

    Dim tabsAfter As String
    tabsAfter = drv.GetTabList
    Debug.Print "After=" & tabsAfter
    Debug.Print "AfterCount=" & CountTabsForCloseSample(tabsAfter)

    drv.CloseWindow
End Sub

'========================================================
' 15. CloseTabByTitle(tabTitle)
' 何ができるか:
' - タイトル完全一致で目的タブだけ閉じられる
'
' 使いどころ:
' - 「BASE という名前のタブだけ閉じたい」ような場合
'
' このサンプルの確認ポイント:
' - BASE タブだけ閉じられること
' - 空タイトルタブが残っていても BASE が消えること
'========================================================
Public Sub Sample_15_CloseTabByTitle_Basic()
    Dim drv As IWebDriver
    Set drv = NewDriverForCloseSample()

    ' BASE タブを開く
    drv.OpenURL "data:text/html,<html><head><title>BASE</title></head><body>base</body></html>"

    ' 空タイトルタブを1枚追加
    drv.ExecuteScript "window.open('about:blank','_blank');"
    drv.SleepByWinAPI 1000

    Dim tabsBefore As String
    tabsBefore = drv.GetTabList
    Debug.Print "Before=" & tabsBefore

    ' タイトル完全一致で BASE を閉じる
    drv.CloseTabByTitle "BASE"
    drv.SleepByWinAPI 700

    Dim tabsAfter As String
    tabsAfter = drv.GetTabList
    Debug.Print "After=" & tabsAfter
    Debug.Print "BASE Exists After=" & HasTabTitleForCloseSample(tabsAfter, "BASE")

    drv.CloseWindow
End Sub

'========================================================
' 16. 目的タブへ移動してからタイトルで閉じる
' 何ができるか:
' - タイトル指定で目的タブへ移動できる
' - そのタブをタイトル指定で閉じられる
'
' 使いどころ:
' - 複数タブの中から対象を明示的にアクティブにしてから閉じたい場合
'
' このサンプルの確認ポイント:
' - TAB2 へ切り替えられること
' - TAB2 を閉じた後、タブ一覧から消えること
'========================================================
Public Sub Sample_16_Switch_Then_Close_ByTitle()
    Dim drv As IWebDriver
    Set drv = NewDriverForCloseSample()

    ' 1枚目
    drv.OpenURL "data:text/html,<html><head><title>BASE</title></head><body>base</body></html>"

    ' 2枚目: タイトル付きタブ
    drv.ExecuteScript "window.open('data:text/html,<html><head><title>TAB2</title></head><body>tab2</body></html>','_blank');"
    drv.SleepByWinAPI 1200

    Debug.Print "Tabs=" & drv.GetTabList

    ' タイトルが見える環境なら TAB2 へ移動
    On Error Resume Next
    drv.SwitchTabByTitle "TAB2"
    On Error GoTo 0

    ' TAB2 を閉じる
    On Error Resume Next
    drv.CloseTabByTitle "TAB2"
    On Error GoTo 0

    drv.SleepByWinAPI 700
    Debug.Print "Tabs After Close=" & drv.GetTabList

    drv.CloseWindow
End Sub

'========================================================
' 17. タブを1枚ずつ減らす
' 何ができるか:
' - 現在のタブ数を見ながら段階的に閉じられる
' - 「最後のタブを順に閉じる」ような処理例になる
'
' 使いどころ:
' - 複数タブを徐々に整理したい場合
' - 一括クローズではなく、状態を見ながら閉じたい場合
'
' このサンプルの確認ポイント:
' - CloseTab を繰り返して件数が減ること
' - 毎回 GetTabList で状態確認できること
'========================================================
Public Sub Sample_17_CloseTabs_OneByOne()
    Dim drv As IWebDriver
    Set drv = NewDriverForCloseSample()

    ' BASE + about:blank 2枚の合計3タブ構成にする
    drv.OpenURL "data:text/html,<html><head><title>BASE</title></head><body>base</body></html>"
    drv.ExecuteScript "window.open('about:blank','_blank');"
    drv.ExecuteScript "window.open('about:blank','_blank');"
    drv.SleepByWinAPI 1000

    Dim tabs As String
    Dim cnt As Long

    tabs = drv.GetTabList
    cnt = CountTabsForCloseSample(tabs)
    Debug.Print "Start Tabs=" & tabs
    Debug.Print "Start Count=" & cnt

    ' 1回目:
    ' 最後のタブを閉じる
    If cnt >= 2 Then
        drv.CloseTab cnt
        drv.SleepByWinAPI 700
    End If

    tabs = drv.GetTabList
    cnt = CountTabsForCloseSample(tabs)
    Debug.Print "After 1st Close=" & tabs
    Debug.Print "After 1st Count=" & cnt

    ' 2回目:
    ' さらに最後のタブを閉じる
    If cnt >= 2 Then
        drv.CloseTab cnt
        drv.SleepByWinAPI 700
    End If

    tabs = drv.GetTabList
    cnt = CountTabsForCloseSample(tabs)
    Debug.Print "After 2nd Close=" & tabs
    Debug.Print "After 2nd Count=" & cnt

    drv.CloseWindow
End Sub
'========================================================
' 18. CSS セレクタ検索
' 何ができるか:
' - CSS セレクタで単一要素を取得できる
' - CSS セレクタで複数要素をまとめて取得できる
' - Class / ID / 子孫セレクタなど、CSS 記法ベースで探索できる
'
' 使いどころ:
' - XPath より CSS セレクタの方が書きやすいページ
' - class / id / 階層指定で素直に取れるページ
'
' このサンプルで確認すること:
' - FindElementByCss が先頭1件を返すこと
' - FindElementsByCss が複数件を列挙できること
'========================================================
Public Sub Sample_18_Find_By_Css()
    Dim drv As IWebDriver
    Set drv = NewDriver()
    
    ' テスト用の簡単なHTMLを開く
    ' .box クラスの要素を3個並べておく
    drv.OpenURL "data:text/html," & _
                "<html><body>" & _
                "<div class='box'>A</div>" & _
                "<div class='box'>B</div>" & _
                "<div class='box'>C</div>" & _
                "</body></html>"
    
    ' 単一取得:
    ' CSS セレクタ .box に一致する最初の1件を取得する
    Debug.Print "First .box = " & drv.FindElementByCss(".box").GetTextContent
    
    ' 複数取得:
    ' .box に一致する要素をすべて取得して順番に確認する
    Dim elems As IWebElements
    Dim e As IWebElement
    Set elems = drv.FindElementsByCss(".box")
    
    For Each e In elems
        Debug.Print ".box -> " & e.GetTextContent
    Next
    
    drv.CloseWindow
End Sub

'========================================================
' 19. テキスト検索
' 何ができるか:
' - 要素の textContent を基準に単一要素を取得できる
' - 同じ textContent を持つ要素を複数取得できる
' - partialMatch=True で部分一致検索もできる
'
' 使いどころ:
' - ID / class が無く、表示文字列だけで探したいページ
' - 「ボタンの表示名」や「ラベル文字列」で要素を探したいケース
'
' このサンプルで確認すること:
' - 完全一致で1件取れること
' - 部分一致で複数件取れること
'========================================================
Public Sub Sample_19_Find_By_Text()
    Dim drv As IWebDriver
    Set drv = NewDriver()
    
    drv.OpenURL "data:text/html," & _
                "<html><body>" & _
                "<div>AAA</div>" & _
                "<div>BBB</div>" & _
                "<div>CCC-01</div>" & _
                "<div>CCC-02</div>" & _
                "</body></html>"
    
    ' 完全一致:
    ' 表示文字列が BBB の要素を取る
    Debug.Print "Text=BBB -> " & drv.FindElementByText("BBB").GetTextContent
    
    ' 部分一致:
    ' 「CCC」を含む要素を全部取る
    Dim elems As IWebElements
    Dim e As IWebElement
    Set elems = drv.FindElementsByText("CCC", True)
    
    For Each e In elems
        Debug.Print "Text contains CCC -> " & e.GetTextContent
    Next
    
    drv.CloseWindow
End Sub

'========================================================
' 20. Driver からの Deep 検索
' 何ができるか:
' - Shadow DOM の内側まで含めて CSS 検索できる
' - Shadow DOM の内側まで含めてテキスト検索できる
'
' 使いどころ:
' - Web Components を使っているページ
' - 通常の querySelector では見つからない Shadow DOM 内要素の取得
'
' このサンプルで確認すること:
' - FindElementByCssDeep が Shadow DOM 内の最初の要素を返すこと
' - FindElementByTextDeep が Shadow DOM 内テキストを取れること
' - FindElementsByCssDeep が複数件を列挙できること
'========================================================
Public Sub Sample_20_Deep_Search_By_Driver()
    Dim drv As IWebDriver
    Set drv = NewDriver()
    
    drv.OpenURL "data:text/html," & _
                "<html><body>" & _
                "<div id='host1'></div>" & _
                "<div id='host2'></div>" & _
                "</body></html>"
    
    ' host1 / host2 に Shadow DOM を作成して、その中に要素を入れる
    drv.ExecuteScript "var h1=document.getElementById('host1');" & _
                      "var s1=h1.attachShadow({mode:'open'});" & _
                      "s1.innerHTML=""<div class='in-shadow'>A1</div><div class='in-shadow'>B1</div>"";" & _
                      "var h2=document.getElementById('host2');" & _
                      "var s2=h2.attachShadow({mode:'open'});" & _
                      "s2.innerHTML=""<span class='in-shadow'>A2</span><span class='in-shadow'>B2</span>"";"
    
    ' Deep CSS:
    ' Shadow DOM をまたいで .in-shadow を検索する
    Debug.Print "Deep CSS first = " & drv.FindElementByCssDeep(".in-shadow").GetTextContent
    
    ' Deep Text:
    ' Shadow DOM をまたいで表示文字列 B2 を検索する
    Debug.Print "Deep Text exact = " & drv.FindElementByTextDeep("B2").GetTextContent
    
    ' Deep CSS の複数件取得
    Dim elems As IWebElements
    Dim e As IWebElement
    Set elems = drv.FindElementsByCssDeep(".in-shadow")
    
    For Each e In elems
        Debug.Print "Deep .in-shadow -> " & e.GetTextContent
    Next
    
    drv.CloseWindow
End Sub

'========================================================
' 21. 要素配下の CSS / Text 検索
' 何ができるか:
' - ある親要素の配下だけを対象に CSS 検索できる
' - ある親要素の配下だけを対象に Text 検索できる
' - 部分一致テキスト検索もできる
'
' 使いどころ:
' - ページ全体ではなく「このブロックの中だけ」を探したいケース
' - テーブル1行 / カード1枚 / ダイアログ1個の内側だけで探したいケース
'
' このサンプルで確認すること:
' - parent.FindElementByCss が親配下だけを対象にすること
' - parent.FindElementsByText が親配下だけを対象にすること
'========================================================
Public Sub Sample_21_Element_Child_Locators()
    Dim drv As IWebDriver
    Set drv = NewDriver()
    
    drv.OpenURL "data:text/html," & _
                "<html><body>" & _
                "<div id='wrap'>" & _
                "<span class='c'>alpha-001</span>" & _
                "<span class='c'>beta-002</span>" & _
                "<span class='c'>beta-003</span>" & _
                "</div>" & _
                "</body></html>"
    
    ' 親コンテナを取る
    Dim root As IWebElement
    Set root = drv.FindElementById("wrap")
    
    ' 親配下だけで CSS 検索
    Debug.Print "Child CSS first = " & root.FindElementByCss(".c").GetTextContent
    
    ' 親配下だけで Text 部分一致検索
    Debug.Print "Child Text partial(beta) first = " & root.FindElementByText("beta", True).GetTextContent
    
    Dim elems As IWebElements
    Dim e As IWebElement
    
    ' 親配下だけで CSS 複数取得
    Set elems = root.FindElementsByCss(".c")
    For Each e In elems
        Debug.Print "Child CSS .c -> " & e.GetTextContent
    Next
    
    ' 親配下だけで Text 部分一致の複数取得
    Set elems = root.FindElementsByText("beta", True)
    For Each e In elems
        Debug.Print "Child Text partial(beta) -> " & e.GetTextContent
    Next
    
    drv.CloseWindow
End Sub

'========================================================
' 22. 要素配下の Deep 検索
' 何ができるか:
' - 親要素の内側だけを対象に Shadow DOM を含む CSS 検索ができる
' - 親要素の内側だけを対象に Shadow DOM を含む Text 検索ができる
'
' 使いどころ:
' - ダイアログや特定コンポーネントの内部 Shadow DOM を探したいケース
' - ページ全体ではなく「このコンポーネントの中だけ」を Deep 検索したいケース
'
' このサンプルで確認すること:
' - area.FindElementByCssDeep が area 配下の Shadow DOM を見に行けること
' - area.FindElementsByTextDeep が複数件取れること
'========================================================
Public Sub Sample_22_Element_Deep_Locators()
    Dim drv As IWebDriver
    Set drv = NewDriver()
    
    drv.OpenURL "data:text/html," & _
                "<html><body>" & _
                "<section id='area'>" & _
                "<div id='hostA'></div>" & _
                "</section>" & _
                "</body></html>"
    
    ' area 配下にある hostA の Shadow DOM に要素を入れる
    drv.ExecuteScript "var h=document.getElementById('hostA');" & _
                      "var s=h.attachShadow({mode:'open'});" & _
                      "s.innerHTML=""<div class='deep-item'>deep-001</div><div class='deep-item'>deep-002</div>"";"
    
    Dim area As IWebElement
    Set area = drv.FindElementById("area")
    
    ' 親要素 area 配下を対象に Deep CSS 検索
    Debug.Print "Elem Deep CSS first = " & area.FindElementByCssDeep(".deep-item").GetTextContent
    
    ' 親要素 area 配下を対象に Deep Text 完全一致検索
    Debug.Print "Elem Deep Text exact = " & area.FindElementByTextDeep("deep-002").GetTextContent
    
    Dim elems As IWebElements
    Dim e As IWebElement
    
    ' 親要素 area 配下を対象に Deep CSS 複数取得
    Set elems = area.FindElementsByCssDeep(".deep-item")
    For Each e In elems
        Debug.Print "Elem Deep CSS -> " & e.GetTextContent
    Next
    
    ' 親要素 area 配下を対象に Deep Text 部分一致複数取得
    Set elems = area.FindElementsByTextDeep("deep", True)
    For Each e In elems
        Debug.Print "Elem Deep Text partial(deep) -> " & e.GetTextContent
    Next
    
    drv.CloseWindow
End Sub

'========================================================
' 23. Dialog 制御
' 何ができるか:
' - prompt に文字列を入れて OK できる
' - confirm を Cancel できる
' - 結果文字列を画面上に反映させて確認できる
'
' 使いどころ:
' - Alert / Confirm / Prompt の分岐を自動化したいケース
' - Prompt に入力値を渡したいケース
'
' このサンプルで確認すること:
' - AcceptAlertDialog("文字列") で prompt に値を渡せること
' - DismissAlertDialog で confirm をキャンセルできること
'========================================================
Public Sub Sample_23_Dialog_Control()
    Dim drv As IWebDriver
    Set drv = NewDriver()
    
    drv.OpenURL "data:text/html," & _
                "<html><body>" & _
                "<button id='btnPrompt' onclick=""var v=prompt('name?','');document.getElementById('out').textContent=(v===null?'PROMPT_CANCEL':'PROMPT:'+v);"">prompt</button>" & _
                "<button id='btnConfirm' onclick=""document.getElementById('out').textContent=confirm('ok?')?'CONFIRM_OK':'CONFIRM_CANCEL';"">confirm</button>" & _
                "<div id='out'></div>" & _
                "</body></html>"
    
    ' prompt を開く
    drv.FindElementById("btnPrompt").Click
    
    ' prompt に COPILOT を入れて OK
    Call drv.AcceptAlertDialog("COPILOT")
    drv.SleepByWinAPI 500
    
    ' 画面側で受け取った値を確認
    Debug.Print "Prompt result = " & drv.FindElementById("out").GetTextContent
    
    ' confirm を開く
    drv.FindElementById("btnConfirm").Click
    
    ' confirm を Cancel
    Call drv.DismissAlertDialog
    drv.SleepByWinAPI 500
    
    ' キャンセル結果を確認
    Debug.Print "Confirm result = " & drv.FindElementById("out").GetTextContent
    
    drv.CloseWindow
End Sub

'========================================================
' 24. Network 介入: Block
' 何ができるか:
' - 特定 URL パターンをブロックできる
' - ブロック対象一覧をクリアできる
' - 介入機能を無効化できる
'
' 使いどころ:
' - 不要な通信を止めたいケース
' - 広告 / 解析 / 特定 API への通信を抑止したいケース
'
' このサンプルで確認すること:
' - blocked.json を含む URL がエラーになること
' - 画面側の fetch が catch 側に入ること
'========================================================
Public Sub Sample_24_Network_Block()
    Dim drv As IWebDriver
    Set drv = NewDriver()
    
    drv.OpenURL "data:text/html,<html><body>network block</body></html>"
    
    ' hook を有効化
    drv.EnableNetworkInterception
    
    ' 以前のブロック条件を消してから今回の条件を追加
    drv.ClearBlockedURLPatterns
    drv.AddBlockedURLPattern "blocked.json"
    
    ' blocked.json を含む URL に fetch して、
    ' 成功なら OK:...
    ' 失敗なら ERR:...
    ' を window.__netOut に格納する
    drv.ExecuteScript "window.__netOut='';" & _
                      "fetch('https://example.com/blocked.json')" & _
                      ".then(function(r){return r.text();})" & _
                      ".then(function(t){window.__netOut='OK:' + t;})" & _
                      ".catch(function(e){window.__netOut='ERR:' + e.message;});"
    
    drv.SleepByWinAPI 1000
    
    ' 期待値:
    ' ERR:Blocked by VBA interceptor
    Debug.Print "Blocked fetch result = " & drv.ExecuteScript("window.__netOut;")
    
    ' 後片付け
    drv.ClearBlockedURLPatterns
    drv.DisableNetworkInterception
    drv.CloseWindow
End Sub

'========================================================
' 25. Network 介入: Mock Response
' 何ができるか:
' - 特定 URL パターンに対してモック応答を返せる
' - モック一覧をクリアできる
' - 介入機能を無効化できる
'
' 使いどころ:
' - 本物の API を叩かずに画面動作だけ確認したいケース
' - 返却内容を固定してテストしたいケース
'
' このサンプルで確認すること:
' - mock.json への fetch が 200 / 固定 body で返ること
' - 画面側の fetch 結果を文字列で確認できること
'========================================================
Public Sub Sample_25_Network_Mock()
    Dim drv As IWebDriver
    Set drv = NewDriver()
    
    drv.OpenURL "data:text/html,<html><body>network mock</body></html>"
    
    ' hook を有効化
    drv.EnableNetworkInterception
    
    ' 以前のモック条件を消してから今回のモックを追加
    drv.ClearMockResponses
    drv.AddMockResponse "mock.json", 200, "{""status"":""ok""}", "application/json"
    
    ' mock.json を含む URL に fetch する
    ' 本来は example.com に通信するが、
    ' hook により固定レスポンスが返る想定
    drv.ExecuteScript "window.__netOut='';" & _
                      "fetch('https://example.com/mock.json')" & _
                      ".then(function(r){" & _
                      "  return r.text().then(function(t){" & _
                      "    window.__netOut='STATUS:' + r.status + ' BODY:' + t;" & _
                      "  });" & _
                      "})" & _
                      ".catch(function(e){window.__netOut='ERR:' + e.message;});"
    
    drv.SleepByWinAPI 1000
    
    ' 期待値:
    ' STATUS:200 BODY:{""status"":""ok""}
    Debug.Print "Mock fetch result = " & drv.ExecuteScript("window.__netOut;")
    
    ' 後片付け
    drv.ClearMockResponses
    drv.DisableNetworkInterception
    drv.CloseWindow
End Sub

'========================================================
' 26. targetId 系テスト
' 何を確認するか:
' - タブ切替
' - タイトルでタブ切替
' - タイトルでタブクローズ
' - indexでタブクローズ
'
' このサンプルで内部的に踏むもの:
' - targetId を使うタブ再接続系
' - targetId を使う CloseTab 系
'========================================================
Public Sub Sample_26_TargetId_RelatedAPIs()
    Dim drv As IWebDriver
    Set drv = NewDriver()
    
    ' 1枚目
    drv.OpenURL "data:text/html,<html><head><title>BASE</title></head><body>base</body></html>"
    
    ' 2枚目
    drv.ExecuteScript _
    "var w=window.open('','_blank');" & _
    "if(w){" & _
    "  w.document.open();" & _
    "  w.document.write('<!DOCTYPE html><html><head><meta charset=""utf-8""><title>TAB2</title></head><body>tab2</body></html>');" & _
    "  w.document.close();" & _
    "}"

    
    ' 3枚目
    drv.ExecuteScript _
    "var w=window.open('','_blank');" & _
    "if(w){" & _
    "  w.document.open();" & _
    "  w.document.write('<!DOCTYPE html><html><head><meta charset=""utf-8""><title>TAB3</title></head><body>tab2</body></html>');" & _
    "  w.document.close();" & _
    "}"

    
    drv.SleepByWinAPI 1200
    
    ' 現在のタブ一覧確認
    Debug.Print "Tabs(before) = " & drv.GetTabList
    
    ' index 指定で切替
    drv.SwitchTabByIndex 2
    Debug.Print "After SwitchTabByIndex(2) URL = " & drv.URL
    
    ' title 指定で切替
    drv.SwitchTabByTitle "TAB3"
    Debug.Print "After SwitchTabByTitle(TAB3) URL = " & drv.URL
    
    ' title 指定で閉じる
    drv.CloseTabByTitle "TAB2"
    drv.SleepByWinAPI 700
    Debug.Print "Tabs(after CloseTabByTitle TAB2) = " & drv.GetTabList
    
    ' index 指定で閉じる
    drv.CloseTab 2
    drv.SleepByWinAPI 700
    Debug.Print "Tabs(after CloseTab 2) = " & drv.GetTabList
    
    drv.CloseWindow
End Sub

'========================================================
' 27. frameId 系テスト
' 何を確認するか:
' - name 指定で frame 切替
' - iframe 要素指定で frame 切替
' - デフォルト frame へ戻れること
'
' このサンプルで内部的に踏むもの:
' - frameId を使う CreateIsolatedFrameWorld 系
'========================================================
Public Sub Sample_27_FrameId_RelatedAPIs()
    Dim drv As IWebDriver
    Set drv = NewDriver()
    
    drv.OpenURL "data:text/html," & _
                "<html><body>" & _
                "<iframe id='f1' name='frameA' srcdoc=""<html><body><input id='inner1' value='inside-A'></body></html>""></iframe>" & _
                "<iframe id='f2' name='frameB' srcdoc=""<html><body><input id='inner2' value='inside-B'></body></html>""></iframe>" & _
                "</body></html>"
    
    ' name 指定で切替
    drv.SwitchFrameByNameOrURLOrIndex "frameA"
    Debug.Print "frameA inner1 = " & CStr(drv.FindElementById("inner1").GetProperty("value"))
    
    ' 親へ戻る
    drv.SwitchFrameToDefault
    
    ' iframe 要素指定で切替
    Dim ifr As IWebElement
    Set ifr = drv.FindElementById("f2")
    drv.SwitchFrameByIframeElement ifr
    Debug.Print "frameB inner2 = " & CStr(drv.FindElementById("inner2").GetProperty("value"))
    
    ' 親へ戻る
    drv.SwitchFrameToDefault
    
    drv.CloseWindow
End Sub

'========================================================
' 28. objectId 系テスト
' 何を確認するか:
' - Click
' - SetValue
' - GetTextContent
' - CDP マウスクリック
'
' このサンプルで内部的に踏むもの:
' - objectId を使う Focus
' - objectId を使う GetBoxModel
' - objectId を使う各種 callFunctionOn 系
'========================================================
Public Sub Sample_28_ObjectId_RelatedAPIs()
    Dim drv As IWebDriver
    Set drv = NewDriver()
    
    drv.OpenURL "data:text/html," & _
                "<html><body>" & _
                "<input id='txt1' value=''>" & _
                "<button id='btn1' onclick=""document.getElementById('out').textContent=document.getElementById('txt1').value;"">Run</button>" & _
                "<button id='btn2' onclick=""document.getElementById('out2').textContent='mouse-clicked';"">CDPClick</button>" & _
                "<div id='out'></div>" & _
                "<div id='out2'></div>" & _
                "</body></html>"
    
    Dim txt As IWebElement
    Dim btn1 As IWebElement
    Dim btn2 As IWebElement
    
    Set txt = drv.FindElementById("txt1")
    Set btn1 = drv.FindElementById("btn1")
    Set btn2 = drv.FindElementById("btn2")
    
    ' SetValue は内部で Focus(objectId) を踏む
    txt.SetValue "OBJECT-ID-TEST"
    
    ' Click は内部で Focus(objectId) を踏む
    btn1.Click
    Debug.Print "out = " & drv.FindElementById("out").GetTextContent
    
    ' DispatchCDPMouseEvent は内部で GetBoxModel(objectId) を踏む
    btn2.DispatchCDPMouseEvent LeftClick_XY, Center
    drv.SleepByWinAPI 500
    Debug.Print "out2 = " & drv.FindElementById("out2").GetTextContent
    
    drv.CloseWindow
End Sub

'========================================================
' 29. requestId 系テスト
' 何を確認するか:
' - WaitForResponse
' - GetResponseBody
'
' このサンプルで内部的に踏むもの:
' - requestId を使う Network.getResponseBody
'========================================================
Public Sub Sample_29_RequestId_RelatedAPIs()
    Dim drv As IWebDriver
    Set drv = NewDriver()
    
    drv.OpenURL "https://example.com"
    
    ' 古い応答記録をクリア
    drv.ClearNetworkResponses
    
    ' 待ち合わせ対象の fetch を投げる
    drv.ExecuteScript "try{fetch('https://example.com/?api=reqtest&ts='+Date.now()).catch(function(){});}catch(e){}"
    
    ' requestId 取得
    Dim rid As String
    rid = drv.WaitForResponse("api=reqtest", 10)
    Debug.Print "RequestId = " & rid
    
    ' requestId を使って body を取得
    Debug.Print "Body = " & drv.GetResponseBody(rid)
    
    drv.CloseWindow

End Sub
'========================================================
' 30. Shadow DOM を Driver から Deep 検索
' 何ができるか:
' - Driver 直下から Shadow DOM 内の要素を CSS で取得できる
' - Driver 直下から Shadow DOM 内の要素を Text で取得できる
' - Deep 系の複数取得ができる
'
' 使いどころ:
' - 通常の FindElementByCss / FindElementByText では届かない
'   Shadow DOM 内要素を直接取りたいケース
'
' このサンプルで確認すること:
' - FindElementByCssDeep で #innerBtn が取れること
' - FindElementByTextDeep で表示文字列から要素が取れること
' - FindElementsByCssDeep / FindElementsByTextDeep で複数件列挙できること
'========================================================
Public Sub Sample_30_ShadowDom_Deep_By_Driver()
    Dim drv As IWebDriver
    Set drv = NewDriver()

    drv.OpenURL "data:text/html," & _
    "<html><body>" & _
    "<div id='host1'></div>" & _
    "<div id='host2'></div>" & _
    "</body></html>"

    ' host1 / host2 に Shadow DOM を作成し、
    ' その内部に Deep 検索対象の要素を配置する
    drv.ExecuteScript _
    "var h1=document.getElementById('host1');" & _
    "var s1=h1.attachShadow({mode:'open'});" & _
    "s1.innerHTML=""" & _
    "<div class='deep-item'>deep-A1</div>" & _
    "<button id='innerBtn'>Shadow DOM 内のボタン</button>" & _
    "<div class='deep-item'>deep-A2</div>" & _
    """;" & _
    "var h2=document.getElementById('host2');" & _
    "var s2=h2.attachShadow({mode:'open'});" & _
    "s2.innerHTML=""" & _
    "<span class='deep-item'>deep-B1</span>" & _
    "<span class='deep-item'>deep-B2</span>" & _
    """;"

    ' Deep CSS: id で Shadow DOM 内ボタンを取得する
    Debug.Print "Driver Deep CSS #innerBtn = " & drv.FindElementByCssDeep("#innerBtn").GetTextContent

    ' Deep Text: 表示文字列で Shadow DOM 内ボタンを取得する
    Debug.Print "Driver Deep Text exact = " & drv.FindElementByTextDeep("Shadow DOM 内のボタン").GetTextContent

    Dim elems As IWebElements
    Dim e As IWebElement

    ' Deep CSS の複数取得
    Set elems = drv.FindElementsByCssDeep(".deep-item")
    For Each e In elems
        Debug.Print "Driver Deep CSS .deep-item -> " & e.GetTextContent
    Next

    ' Deep Text の複数取得（部分一致）
    Set elems = drv.FindElementsByTextDeep("deep-", True)
    For Each e In elems
        Debug.Print "Driver Deep Text partial(deep-) -> " & e.GetTextContent
    Next

    drv.CloseWindow
End Sub

'========================================================
' 31. Shadow DOM を Element 配下から Deep 検索
' 何ができるか:
' - 親要素の配下だけを対象に Shadow DOM 内要素を CSS で取得できる
' - 親要素の配下だけを対象に Shadow DOM 内要素を Text で取得できる
' - Element ベースの Deep 複数取得ができる
'
' 使いどころ:
' - ページ全体ではなく、特定コンテナ / 特定コンポーネントの内側だけを
'   Shadow DOM 含みで探したいケース
'
' このサンプルで確認すること:
' - area.FindElementByCssDeep が area 配下の Shadow DOM を見に行けること
' - area.FindElementByTextDeep が文字列一致で取れること
' - area.FindElementsByCssDeep / area.FindElementsByTextDeep が複数件取れること
'========================================================
Public Sub Sample_31_ShadowDom_Deep_By_Element()
    Dim drv As IWebDriver
    Set drv = NewDriver()

    drv.OpenURL "data:text/html," & _
    "<html><body>" & _
    "<section id='area1'><div id='hostA'></div></section>" & _
    "<section id='area2'><div id='hostB'></div></section>" & _
    "</body></html>"

    ' area1 配下の hostA にだけ Shadow DOM を作る
    ' area2 側には別内容を置いて、親要素配下限定で取れることを確認する
    drv.ExecuteScript _
    "var hA=document.getElementById('hostA');" & _
    "var sA=hA.attachShadow({mode:'open'});" & _
    "sA.innerHTML=""" & _
    "<div class='part'>alpha-deep-01</div>" & _
    "<div class='part'>alpha-deep-02</div>" & _
    "<button id='btnArea'>Area1 Shadow Button</button>" & _
    """;" & _
    "var hB=document.getElementById('hostB');" & _
    "var sB=hB.attachShadow({mode:'open'});" & _
    "sB.innerHTML=""" & _
    "<div class='part'>beta-deep-01</div>" & _
    """;"

    Dim area As IWebElement
    Set area = drv.FindElementById("area1")

    ' area1 配下だけを対象に Deep CSS 検索
    Debug.Print "Element Deep CSS #btnArea = " & area.FindElementByCssDeep("#btnArea").GetTextContent

    ' area1 配下だけを対象に Deep Text 検索
    Debug.Print "Element Deep Text exact = " & area.FindElementByTextDeep("Area1 Shadow Button").GetTextContent

    Dim elems As IWebElements
    Dim e As IWebElement

    ' area1 配下だけを対象に Deep CSS 複数取得
    Set elems = area.FindElementsByCssDeep(".part")
    For Each e In elems
        Debug.Print "Element Deep CSS .part -> " & e.GetTextContent
    Next

    ' area1 配下だけを対象に Deep Text 部分一致複数取得
    Set elems = area.FindElementsByTextDeep("alpha-deep", True)
    For Each e In elems
        Debug.Print "Element Deep Text partial(alpha-deep) -> " & e.GetTextContent
    Next

    drv.CloseWindow
End Sub
'========================================================
' 30. table 要素 → 二次元配列
' 何ができるか:
' - table 要素を受け取って二次元配列へ変換できる
' - th / td の順で1行分を配列化できる
'
' このサンプルの確認ポイント:
' - TableElementToArray2D が二次元配列を返すこと
' - 行列の内容が期待どおりに入ること
'========================================================
Public Sub Sample_32_TableElementToArray2D()
    Dim drv As IWebDriver
    Set drv = NewDriver()
    
    drv.OpenURL "data:text/html," & _
                "<html><body>" & _
                "<table id='tbl1' border='1'>" & _
                "<tr><th>ID</th><th>Name</th><th>Dept</th></tr>" & _
                "<tr><td>1</td><td>Alice</td><td>Sales</td></tr>" & _
                "<tr><td>2</td><td>Bob</td><td>Dev</td></tr>" & _
                "<tr><td>3</td><td>Carol</td><td>HR</td></tr>" & _
                "</table>" & _
                "</body></html>"
    
    Dim tbl As IWebElement
    Set tbl = drv.FindElementById("tbl1")
    
    Dim arr As Variant
    arr = TableElementToArray2D(tbl)
    
    If IsEmpty(arr) Then
        Debug.Print "TableElementToArray2D = Empty"
        drv.CloseWindow
        Exit Sub
    End If
    
    Dim r As Long
    Dim c As Long
    
    Debug.Print "Rows = " & UBound(arr, 1)
    Debug.Print "Cols = " & UBound(arr, 2)
    
    For r = 1 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
            Debug.Print "arr(" & r & "," & c & ") = " & arr(r, c)
        Next c
    Next r
    
    drv.CloseWindow
End Sub

'========================================================
' 33. 要素コレクション → 1次元 text 配列
' 何ができるか:
' - 複数要素の textContent を 1 次元配列で取得できる
'========================================================
Public Sub Sample_33_WebElementsToTextArray1D()
    Dim drv As IWebDriver
    Set drv = NewDriver()
    
    drv.OpenURL "data:text/html," & _
                "<html><body>" & _
                "<ul>" & _
                "<li class='item'>A</li>" & _
                "<li class='item'>B</li>" & _
                "<li class='item'>C</li>" & _
                "</ul>" & _
                "</body></html>"
    
    Dim elems As IWebElements
    Set elems = drv.FindElementsByCss(".item")
    
    Dim arr As Variant
    arr = WebElementsToTextArray1D(elems)
    
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        Debug.Print "TextArr(" & i & ") = " & arr(i)
    Next
    
    drv.CloseWindow
End Sub

'========================================================
' 34. 要素コレクション → 属性値 1次元配列
' 何ができるか:
' - 複数要素の属性値を 1 次元配列で取得できる
'========================================================
Public Sub Sample_34_WebElementsToAttributeArray1D()
    Dim drv As IWebDriver
    Set drv = NewDriver()
    
    drv.OpenURL "data:text/html," & _
                "<html><body>" & _
                "<a class='lnk' href='https://example.com/a'>A</a>" & _
                "<a class='lnk' href='https://example.com/b'>B</a>" & _
                "<a class='lnk' href='https://example.com/c'>C</a>" & _
                "</body></html>"
    
    Dim elems As IWebElements
    Set elems = drv.FindElementsByCss(".lnk")
    
    Dim arr As Variant
    arr = WebElementsToAttributeArray1D(elems, "href")
    
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        Debug.Print "HrefArr(" & i & ") = " & arr(i)
    Next
    
    drv.CloseWindow
End Sub

'========================================================
' 35. 要素コレクション text の連結
' 何ができるか:
' - 複数要素の textContent を区切り文字で連結できる
'========================================================
Public Sub Sample_35_JoinWebElementsText()
    Dim drv As IWebDriver
    Set drv = NewDriver()
    
    drv.OpenURL "data:text/html," & _
                "<html><body>" & _
                "<div class='msg'>Line1</div>" & _
                "<div class='msg'>Line2</div>" & _
                "<div class='msg'>Line3</div>" & _
                "</body></html>"
    
    Dim elems As IWebElements
    Set elems = drv.FindElementsByCss(".msg")
    
    Debug.Print "Joined(vbLf) =" & vbCrLf & JoinWebElementsText(elems, vbLf)
    Debug.Print "Joined(|) = " & JoinWebElementsText(elems, "|")
    
    drv.CloseWindow
End Sub

'========================================================
' 36. table 要素 → TSV 文字列
' 何ができるか:
' - table 要素を TSV 文字列へ変換できる
'========================================================
Public Sub Sample_36_TableElementToTSV()
    Dim drv As IWebDriver
    Set drv = NewDriver()
    
    drv.OpenURL "data:text/html," & _
                "<html><body>" & _
                "<table id='tbl1' border='1'>" & _
                "<tr><th>ID</th><th>Name</th><th>Dept</th></tr>" & _
                "<tr><td>1</td><td>Alice</td><td>Sales</td></tr>" & _
                "<tr><td>2</td><td>Bob</td><td>Dev</td></tr>" & _
                "</table>" & _
                "</body></html>"
    
    Dim tbl As IWebElement
    Set tbl = drv.FindElementById("tbl1")
    
    Debug.Print "TSV =" & vbCrLf & TableElementToTSV(tbl)
    
    drv.CloseWindow
End Sub

'========================================================
' 37. 2次元配列 → Range 出力
' 何ができるか:
' - 2 次元配列をシートへ一括出力できる
'========================================================
Public Sub Sample_37_Array2DToRange()
    Dim arr(1 To 3, 1 To 3) As Variant
    
    arr(1, 1) = "ID"
    arr(1, 2) = "Name"
    arr(1, 3) = "Dept"
    
    arr(2, 1) = 1
    arr(2, 2) = "Alice"
    arr(2, 3) = "Sales"
    
    arr(3, 1) = 2
    arr(3, 2) = "Bob"
    arr(3, 3) = "Dev"
    
    ' アクティブシートの A1 から出力する
    Array2DToRange arr, ActiveSheet.Range("A1")
    
    Debug.Print "Array2DToRange wrote 3x3 to " & ActiveSheet.name & "!A1"
End Sub
'========================================================
' 38. QueryString / GetValueFromQueryString
' 何ができるか:
' - 現在URLの QueryString を Dictionary として取得できる
' - 指定キーの値を直接取得できる
'
' このサンプルの確認ポイント:
' - QueryString に a / b が入ること
' - GetValueFromQueryString で各値を取得できること
'========================================================
Public Sub Sample_38_QueryString_Basic()
    Dim drv As IWebDriver
    Set drv = NewDriver()
    
    ' QueryString 付きURLを開く
    drv.OpenURL "https://example.com/?a=1&b=hello"
    
    ' QueryString 全体を取得する
    Dim qs As Object
    Set qs = drv.QueryString
    
    ' 個別キーを取得して確認する
    Debug.Print "a = " & drv.GetValueFromQueryString("a")
    Debug.Print "b = " & drv.GetValueFromQueryString("b")
    
    ' Dictionary 側の件数も確認する
    Debug.Print "QueryString.Count = " & qs.count
    
    ' 現在URLの確認
    Debug.Print "URL = " & drv.URL
    
    drv.CloseWindow
End Sub

'========================================================
' 39. Utility Public APIs
' 何ができるか:
' - URL エンコード文字列を取得できる
' - URL Safe なランダム文字列を生成できる
' - SHA-256 を Base64URL 形式で取得できる
' - 暗号学的乱数バイト列を取得できる
'
' このサンプルの確認ポイント:
' - EncodeURIConpornent が空でないこと
' - GetURLSafeRundomString が指定長で返ること
' - GetBase64URLEncodedSHA256 が空でないこと
' - GetSecureRandomNumber の配列長が指定値になること
'========================================================
Public Sub Sample_39_Utility_PublicAPIs_Basic()
    Dim drv As IWebDriver
    Set drv = NewDriver()
    
    ' URL エンコード結果を確認する
    Debug.Print "EncodeURIConpornent = " & drv.EncodeURIConpornent("a b+c")
    
    ' URL Safe なランダム文字列を確認する
    Debug.Print "GetURLSafeRundomString = " & drv.GetURLSafeRundomString(16)
    
    ' SHA-256 の Base64URL 結果を確認する
    Debug.Print "SHA256(base64url) = " & drv.GetBase64URLEncodedSHA256("abc")
    
    ' 暗号学的乱数バイト列の長さを確認する
    Dim v As Variant
    v = drv.GetSecureRandomNumber(8)
    Debug.Print "SecureRandom byte count = " & (UBound(v) - LBound(v) + 1)
    
    drv.CloseWindow
End Sub

'========================================================
' 40. ShowWindow / ScreenShotPasteToSheet
' 何ができるか:
' - ブラウザウィンドウを最大化 / 通常表示へ切り替えできる
' - スクリーンショットをシートへ貼り付けできる
'
' このサンプルの確認ポイント:
' - ShowWindow(WindowSize.maximize) が動くこと
' - ShowWindow(WindowSize.normal) が動くこと
' - ScreenShotPasteToSheet が Shape を返すこと
'========================================================
Public Sub Sample_40_Window_And_PasteShot_Basic()
    Dim drv As IWebDriver
    Set drv = NewDriver()
    
    ' 表示確認用ページを開く
    drv.OpenURL "https://example.com"
    
    ' 最大化してから通常表示へ戻す
    drv.ShowWindow WindowSize.maximize
    drv.SleepByWinAPI 1000
    drv.ShowWindow WindowSize.normal
    
    ' アクティブシートへスクリーンショットを貼り付ける
    Dim shp As Shape
    Set shp = drv.ScreenShotPasteToSheet(10, 10, 0, 0, ActiveSheet.name, Image.png)
    
    ' 返却された Shape 情報を確認する
    Debug.Print "Shape.Name = " & shp.name
    Debug.Print "Shape.Width = " & shp.width
    Debug.Print "Shape.Height = " & shp.height
    
    drv.CloseWindow
End Sub

'========================================================
' 41. Select 要素系 API
' 何ができるか:
' - SelectItemInSelectBoxByText で選択肢を切り替えできる
' - GetSelectedValue で現在選択 value を取得できる
' - GetSelectedTextContent で現在選択 text を取得できる
'
' 使いどころ:
' - select 要素の選択状態を text 基準で変えたいケース
' - 選択後の value / 表示文字列の両方を確認したいケース
'
' このサンプルの確認ポイント:
' - SelectItemInSelectBoxByText が動くこと
' - GetSelectedValue が value を返すこと
' - GetSelectedTextContent が表示文字列を返すこと
'========================================================
Public Sub Sample_41_SelectBox_APIs()
    Dim drv As IWebDriver
    Set drv = NewDriver()
    
    drv.OpenURL "data:text/html," & _
                "<html><body>" & _
                "<select id='sel1'>" & _
                "<option value='A'>Alpha</option>" & _
                "<option value='B'>Beta</option>" & _
                "<option value='C'>Gamma</option>" & _
                "</select>" & _
                "</body></html>"
    
    Dim sel As IWebElement
    Set sel = drv.FindElementById("sel1")
    
    ' 表示文字列 "Beta" を持つ option を選択する
    sel.SelectItemInSelectBoxByText "Beta"
    
    ' 現在の value を確認する
    Debug.Print "SelectedValue = " & sel.GetSelectedValue
    
    ' 現在の表示文字列を確認する
    Debug.Print "SelectedText = " & sel.GetSelectedTextContent
    
    drv.CloseWindow
End Sub

'========================================================
' 42. 入力補助 API
' 何ができるか:
' - SetTextContent で textContent を直接書き換えできる
' - SetValueViaPropertySetter で value setter 経由の入力ができる
' - FocusAndEnterKey で Enter キー入力を発生できる
'
' 使いどころ:
' - 通常の SetValue ではなく property setter 経由で値を入れたいケース
' - contenteditable / textContent ベースの要素を更新したいケース
' - Enter キー押下を明示的に発生させたいケース
'
' このサンプルの確認ポイント:
' - SetTextContent が div の textContent を更新すること
' - SetValueViaPropertySetter が input value を更新すること
' - FocusAndEnterKey で Enter 押下結果が反映されること
'========================================================
Public Sub Sample_42_Input_Helper_APIs()
    Dim drv As IWebDriver
    Set drv = NewDriver()
    
    drv.OpenURL "data:text/html," & _
                "<html><body>" & _
                "<div id='txtDiv'>before</div>" & _
                "<input id='txt1' value=''>" & _
                "<div id='out'></div>" & _
                "<script>" & _
                "document.getElementById('txt1').addEventListener('keydown',function(e){" & _
                " if(e.key==='Enter'){document.getElementById('out').textContent='ENTER:' + this.value;}" & _
                "});" & _
                "</script>" & _
                "</body></html>"
    
    Dim div1 As IWebElement
    Dim txt1 As IWebElement
    
    Set div1 = drv.FindElementById("txtDiv")
    Set txt1 = drv.FindElementById("txt1")
    
    ' div の textContent を直接書き換える
    div1.SetTextContent "after-textContent"
    Debug.Print "txtDiv = " & div1.GetTextContent
    
    ' input の value を property setter 経由で更新する
    Call txt1.SetValueViaPropertySetter("setter-input")
    Debug.Print "txt1.value = " & CStr(txt1.GetProperty("value"))
    
    ' focus して Enter を送る
    txt1.FocusAndEnterKey
    drv.SleepByWinAPI 500
    
    ' Enter 押下結果を確認する
    Debug.Print "out = " & drv.FindElementById("out").GetTextContent
    
    drv.CloseWindow
End Sub

'========================================================
' 43. 要素イベント発火 API
' 何ができるか:
' - DispatchCDPKeyEvent でキーイベントを発火できる
' - DispatchJSInputEvent / DispatchJSChangeEvent / DispatchJSFocusEvent が使える
' - DispatchJSMouseEvent / DispatchJSKeyboardEvent が使える
'
' 使いどころ:
' - 値変更後に input / change / focus イベントを明示的に流したいケース
' - JS ベースの監視ロジックに対してイベントだけを発火したいケース
' - キーイベント / マウスイベントの最低限の動作確認をしたいケース
'
' このサンプルの確認ポイント:
' - JS input / change / focus イベントが記録されること
' - JS mouse / keyboard イベントが記録されること
' - CDP key event 実行後も処理継続できること
'========================================================
Public Sub Sample_43_Dispatch_Event_APIs()
    Dim drv As IWebDriver
    Set drv = NewDriver()
    
    drv.OpenURL "data:text/html," & _
                "<html><body>" & _
                "<input id='txt1' value=''>" & _
                "<button id='btn1'>BTN</button>" & _
                "<div id='log'></div>" & _
                "<script>" & _
                "window.__evlog=[];" & _
                "function addLog(s){window.__evlog.push(s);document.getElementById('log').textContent=window.__evlog.join('|');}" & _
                "var t=document.getElementById('txt1');" & _
                "var b=document.getElementById('btn1');" & _
                "t.addEventListener('focus',function(){addLog('focus');});" & _
                "t.addEventListener('input',function(){addLog('input');});" & _
                "t.addEventListener('change',function(){addLog('change');});" & _
                "t.addEventListener('keydown',function(e){addLog('keydown:' + e.key);});" & _
                "b.addEventListener('click',function(){addLog('click');});" & _
                "</script>" & _
                "</body></html>"
    
    Dim txt1 As IWebElement
    Dim btn1 As IWebElement
    
    Set txt1 = drv.FindElementById("txt1")
    Set btn1 = drv.FindElementById("btn1")
    
    ' 先に value を入れておく
    txt1.SetValue "event-test"
    
    ' JS focus / input / change を発火する
    txt1.DispatchJSFocusEvent
    txt1.DispatchJSInputEvent
    txt1.DispatchJSChangeEvent
    
    ' JS keyboard event を発火する
    txt1.DispatchJSKeyboardEvent KeyDown, Enter_Key
    
    ' CDP keyboard event を発火する
    txt1.DispatchCDPKeyEvent KeyDown_, Enter_Key
    txt1.DispatchCDPKeyEvent KeyUp_, Enter_Key
    
    ' JS mouse click を発火する
    btn1.DispatchJSMouseEvent LeftClick
    
    drv.SleepByWinAPI 500
    
    ' 発火したイベント列を確認する
    Debug.Print "EventLog = " & drv.FindElementById("log").GetTextContent
    
    drv.CloseWindow
End Sub
