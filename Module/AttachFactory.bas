Attribute VB_Name = "AttachFactory"
Option Explicit

Public g_AttachExistingBrowser_ As Boolean
Public g_AttachTargetTitle_ As String

Public Function NewAttachedEdgeDriver(Optional ByVal targetTitle As String = "") As IWebDriver
    Dim drv As IWebDriver

    g_AttachExistingBrowser_ = True
    g_AttachTargetTitle_ = targetTitle

    On Error GoTo EH
    Set drv = New EdgeDriver
    Set NewAttachedEdgeDriver = drv

CleanUp:
    g_AttachExistingBrowser_ = False
    g_AttachTargetTitle_ = ""
    Exit Function

EH:
    g_AttachExistingBrowser_ = False
    g_AttachTargetTitle_ = ""
    Err.Raise Err.Number, Err.source, Err.Description
End Function

Public Function NewAttachedChromeDriver(Optional ByVal targetTitle As String = "") As IWebDriver
    Dim drv As IWebDriver

    g_AttachExistingBrowser_ = True
    g_AttachTargetTitle_ = targetTitle

    On Error GoTo EH
    Set drv = New ChromeDriver
    Set NewAttachedChromeDriver = drv

CleanUp:
    g_AttachExistingBrowser_ = False
    g_AttachTargetTitle_ = ""
    Exit Function

EH:
    g_AttachExistingBrowser_ = False
    g_AttachTargetTitle_ = ""
    Err.Raise Err.Number, Err.source, Err.Description
End Function

