VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------
' SeleniumWrapperVBA v1.0.4
'
' Auther:   er-ri (https://github.com/er-ri/selenium-wrapper-vba)
' Date:     2022/03
' License:  MIT (https://opensource.org/licenses/MIT)
'------------------------------------------------------------------
'
' =================================================================
' #Class Options
' =================================================================
Option Explicit

Public Enum BrowserTypesList
    InternetExplorer = 1
    Edge = 2
    Chrome = 3
    Firefox = 4
End Enum

Private m_BrowserType As BrowserTypesList

' Options for Internet Explorer
Private m_AttachToEdgeChrome As Boolean
Private m_EdgeExecutablePath As String
Private m_IntroduceFlakinessByIgnoringSecurityDomains As Boolean
Private m_IgnoreZoomSetting As Boolean

' Options for Chrome
Private m_ChromeArguments As Collection

' Options for Firefox
Private m_FirefoxArguments As Collection

' Options for Edge
Private m_EdgeArguments As Collection

'------------------------------------------------------------------
' Encapsulation
'------------------------------------------------------------------
Public Property Get BrowserType() As BrowserTypesList
    BrowserType = m_BrowserType
End Property

Public Property Let BrowserType(ByVal new_BrowserType As BrowserTypesList)
    m_BrowserType = new_BrowserType
End Property

Public Property Get AttachToEdgeChrome() As Boolean
    AttachToEdgeChrome = m_AttachToEdgeChrome
End Property

Public Property Let AttachToEdgeChrome(ByVal new_AttachToEdgeChrome As Boolean)
    m_AttachToEdgeChrome = new_AttachToEdgeChrome
End Property

Public Property Get EdgeExecutablePath() As String
    EdgeExecutablePath = m_EdgeExecutablePath
End Property

Public Property Let EdgeExecutablePath(ByVal new_EdgeExecutablePath As String)
    m_EdgeExecutablePath = new_EdgeExecutablePath
End Property

Public Property Get IntroduceFlakinessByIgnoringSecurityDomains() As Boolean
    IntroduceFlakinessByIgnoringSecurityDomains = m_IntroduceFlakinessByIgnoringSecurityDomains
End Property

Public Property Let IntroduceFlakinessByIgnoringSecurityDomains(ByVal new_IntroduceFlakinessByIgnoringSecurityDomains As Boolean)
    m_IntroduceFlakinessByIgnoringSecurityDomains = new_IntroduceFlakinessByIgnoringSecurityDomains
End Property

Public Property Get IgnoreZoomSetting() As Boolean
    IgnoreZoomSetting = m_IgnoreZoomSetting
End Property

Public Property Let IgnoreZoomSetting(ByVal new_IgnoreZoomSetting As Boolean)
    m_IgnoreZoomSetting = new_IgnoreZoomSetting
End Property

Public Property Get ChromeArguments() As Collection
    Set ChromeArguments = m_ChromeArguments
End Property

Public Property Let ChromeArguments(ByVal new_ChromeArguments As Collection)
    Set m_ChromeArguments = new_ChromeArguments
End Property

Public Property Get FirefoxArguments() As Collection
    Set FirefoxArguments = m_FirefoxArguments
End Property

Public Property Let FirefoxArguments(ByVal new_FirefoxArguments As Collection)
    Set m_FirefoxArguments = new_FirefoxArguments
End Property

Public Property Get EdgeArguments() As Collection
    Set EdgeArguments = m_EdgeArguments
End Property

Public Property Let EdgeArguments(ByVal new_EdgeArguments As Collection)
    Set m_EdgeArguments = new_EdgeArguments
End Property
'------------------------------------------------------------------
' Constructor
'------------------------------------------------------------------
Private Sub Class_Initialize()
    Set m_ChromeArguments = New Collection
    Set m_FirefoxArguments = New Collection
    Set m_EdgeArguments = New Collection
End Sub

'------------------------------------------------------------------
' Class Method
'------------------------------------------------------------------
Public Function ConstructCapabilities() As Dictionary
    Dim capabilitiesRoot As New Dictionary
    Dim alwaysMatchRoot As New Dictionary
 
    Select Case True
    Case Me.BrowserType = InternetExplorer
        Dim ieOptionsRoot As New Dictionary
        ieOptionsRoot.Add "ie.edgechromium", Me.AttachToEdgeChrome
        ieOptionsRoot.Add "ie.edgepath", Me.EdgeExecutablePath
        ieOptionsRoot.Add "ignoreProtectedModeSettings", Me.IntroduceFlakinessByIgnoringSecurityDomains
        ieOptionsRoot.Add "ignoreZoomSetting", Me.IgnoreZoomSetting
        
        alwaysMatchRoot.Add "browserName", "internet explorer"
        alwaysMatchRoot.Add "se:ieOptions", ieOptionsRoot
        capabilitiesRoot.Add "alwaysMatch", alwaysMatchRoot
    Case Me.BrowserType = Edge
        Dim edgeOptionsRoot As New Dictionary
        edgeOptionsRoot.Add "args", EdgeArguments
    
        alwaysMatchRoot.Add "browserName", "msedge"
        alwaysMatchRoot.Add "ms:edgeOptions", edgeOptionsRoot
        capabilitiesRoot.Add "alwaysMatch", alwaysMatchRoot
    Case Me.BrowserType = Chrome
        Dim chromeOptionsRoot As New Dictionary
        chromeOptionsRoot.Add "args", ChromeArguments
        
        alwaysMatchRoot.Add "browserName", "chrome"
        alwaysMatchRoot.Add "goog:chromeOptions", chromeOptionsRoot
        capabilitiesRoot.Add "alwaysMatch", alwaysMatchRoot
    Case Me.BrowserType = Firefox
        Dim firefoxOptionsRoot As New Dictionary
        firefoxOptionsRoot.Add "args", FirefoxArguments
        
        alwaysMatchRoot.Add "browserName", "firefox"
        alwaysMatchRoot.Add "moz:firefoxOptions", firefoxOptionsRoot
        capabilitiesRoot.Add "alwaysMatch", alwaysMatchRoot
    End Select
    
    Set ConstructCapabilities = capabilitiesRoot
End Function
