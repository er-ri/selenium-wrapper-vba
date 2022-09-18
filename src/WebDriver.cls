VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebDriver"
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
' #Class WebDriver
' =================================================================
Option Explicit

' Standard WebDriver commands.
Public W3C_NEW_SESSION
Public W3C_DELETE_SESSION
Public W3C_STATUS
Public W3C_GET_TIMEOUTS
Public W3C_SET_TIMEOUTS
Public W3C_GO
Public W3C_GET_CURRENT_URL
Public W3C_BACK
Public W3C_FORWARD
Public W3C_REFRESH
Public W3C_GET_TITLE
Public W3C_GET_WINDOW_HANDLE
Public W3C_CLOSE_WINDOW
Public W3C_SWITCH_TO_WINDOW
Public W3C_GET_WINDOW_HANDLES
Public W3C_NEW_WINDOW
Public W3C_SWITCH_TO_FRAME
Public W3C_SWITCH_TO_PARENT_FRAME
Public W3C_GET_WINDOW_RECT
Public W3C_SET_WINDOW_RECT
Public W3C_MAXIMIZE_WINDOW
Public W3C_MINIMIZE_WINDOW
Public W3C_FULLSCREEN_WINDOW
Public W3C_FIND_ELEMENT
Public W3C_FIND_ELEMENTS
Public W3C_FIND_ELEMENT_FROM_ELEMENT
Public W3C_FIND_ELEMENTS_FROM_ELEMENT
Public W3C_FIND_ELEMENT_FROM_SHADOW_ROOT
Public W3C_FIND_ELEMENTS_FROM_SHADOW_ROOT
Public W3C_GET_ACTIVE_ELEMENT
Public W3C_IS_ELEMENT_SELECTED
Public W3C_GET_ELEMENT_ATTRIBUTE
Public W3C_GET_ELEMENT_PROPERTY
Public W3C_GET_ELEMENT_CSS_VALUE
Public W3C_GET_ELEMENT_TEXT
Public W3C_GET_ELEMENT_TAG_NAME
Public W3C_GET_ELEMENT_RECT
Public W3C_IS_ELEMENT_ENABLED
Public W3C_ELEMENT_CLICK
Public W3C_ELEMENT_CLEAR
Public W3C_ELEMENT_SEND_KEYS
Public W3C_GET_PAGE_SOURCE
Public W3C_EXECUTE_SCRIPT
Public W3C_EXECUTE_ASYNC_SCRIPT
Public W3C_GET_ALL_COOKIES
Public W3C_GET_NAMED_COOKIE
Public W3C_ADD_COOKIE
Public W3C_DELETE_COOKIE
Public W3C_DELETE_ALL_COOKIES
Public W3C_PERFORM_ACTIONS
Public W3C_RELEASE_ACTIONS
Public W3C_DISMISS_ALERT
Public W3C_ACCEPT_ALERT
Public W3C_GET_ALERT_TEXT
Public W3C_SEND_ALERT_TEXT
Public W3C_TAKE_SCREENSHOT
Public W3C_TAKE_ELEMENT_SCREENSHOT

Private Const ELEMENT_KEY = "element-6066-11e4-a52e-4f735466cecf"

Public Enum By
    ClassName = 1
    ID = 2
    name = 3
    LinkText = 4
    PartialLinkText = 5
    TagName = 6
    XPath = 7
End Enum

Private m_RemoteEndUrl As String
Private m_PID As String
Private m_Capabilities As Dictionary

'------------------------------------------------------------------
' Encapsulation
'------------------------------------------------------------------
Public Property Get RemoteEndUrl() As String
    RemoteEndUrl = m_RemoteEndUrl
End Property

Public Property Let RemoteEndUrl(ByVal new_RemoteEndUrl As String)
    m_RemoteEndUrl = new_RemoteEndUrl
End Property

Public Property Get PID() As String
    PID = m_PID
End Property

Public Property Let PID(ByVal new_PID As String)
    m_PID = new_PID
End Property

Public Property Let Capabilities(ByVal new_Capabilities As Dictionary)
    Set m_Capabilities = new_Capabilities
End Property

Public Property Get Capabilities() As Dictionary
    Set Capabilities = m_Capabilities
End Property

'------------------------------------------------------------------
' Constructor
'------------------------------------------------------------------
Private Sub Class_Initialize()
    W3C_NEW_SESSION = Array("POST", "/session")
    W3C_DELETE_SESSION = Array("DELETE", "/session/{session id}")
    W3C_STATUS = Array("GET", "/status")
    W3C_GET_TIMEOUTS = Array("GET", "/session/{session id}/timeouts")
    W3C_SET_TIMEOUTS = Array("POST", "/session/{session id}/timeouts")
    W3C_GO = Array("POST", "/session/{session id}/url")
    W3C_BACK = Array("POST", "/session/{session id}/back")
    W3C_REFRESH = Array("POST", "/session/{session id}/refresh")
    W3C_GET_WINDOW_HANDLE = Array("GET", "/session/{session id}/window")
    W3C_NEW_WINDOW = Array("POST", "/session/{session id}/window/new")
    W3C_SWITCH_TO_WINDOW = Array("POST", "/session/{session id}/window")
    W3C_SWITCH_TO_FRAME = Array("POST", "/session/{session id}/frame")
    W3C_GET_WINDOW_RECT = Array("GET", "/session/{session id}/window/rect")
    W3C_MAXIMIZE_WINDOW = Array("POST", "/session/{session id}/window/maximize")
    W3C_FULLSCREEN_WINDOW = Array("POST", "/session/{session id}/window/fullscreen")
    W3C_FIND_ELEMENTS = Array("POST", "/session/{session id}/elements")
    W3C_FIND_ELEMENTS_FROM_ELEMENT = Array("POST", "/session/{session id}/element/{element id}/elements")
    W3C_FIND_ELEMENT_FROM_SHADOW_ROOT = Array("POST", "/session/{session id}/shadow/{shadow id}/element")
    W3C_FIND_ELEMENTS_FROM_SHADOW_ROOT = Array("POST", "/session/{session id}/shadow/{shadow id}/elements")
    W3C_IS_ELEMENT_SELECTED = Array("GET", "/session/{session id}/element/{element id}/selected")
    W3C_GET_ELEMENT_PROPERTY = Array("GET", "/session/{session id}/element/{element id}/property/{name}")
    W3C_GET_ELEMENT_TEXT = Array("GET", "/session/{session id}/element/{element id}/text")
    W3C_GET_ELEMENT_RECT = Array("GET", "/session/{session id}/element/{element id}/rect")
    W3C_ELEMENT_CLICK = Array("POST", "/session/{session id}/element/{element id}/click")
    W3C_ELEMENT_SEND_KEYS = Array("POST", "/session/{session id}/element/{element id}/value")
    W3C_EXECUTE_SCRIPT = Array("POST", "/session/{session id}/execute/sync")
    W3C_GET_ALL_COOKIES = Array("GET", "/session/{session id}/cookie")
    W3C_ADD_COOKIE = Array("POST", "/session/{session id}/cookie")
    W3C_DELETE_ALL_COOKIES = Array("DELETE", "/session/{session id)/cookie")
    W3C_RELEASE_ACTIONS = Array("DELETE", "/session/{session id}/actions")
    W3C_ACCEPT_ALERT = Array("POST", "/session/{session id}/alert/accept")
    W3C_SEND_ALERT_TEXT = Array("POST", "/session/{session id}/alert/text")
    W3C_TAKE_ELEMENT_SCREENSHOT = Array("GET", "/session/{session id}/element/{element id}/screenshot")
    W3C_GET_CURRENT_URL = Array("GET", "/session/{session id}/url")
    W3C_FORWARD = Array("POST", "/session/{session id}/forward")
    W3C_GET_TITLE = Array("GET", "/session/{session id}/title")
    W3C_CLOSE_WINDOW = Array("DELETE", "/session/{session id}/window")
    W3C_GET_WINDOW_HANDLES = Array("GET", "/session/{session id}/window/handles")
    W3C_SWITCH_TO_PARENT_FRAME = Array("POST", "/session/{session id}/frame/parent")
    W3C_SET_WINDOW_RECT = Array("POST", "/session/{session id}/window/rect")
    W3C_MINIMIZE_WINDOW = Array("POST", "/session/{session id}/window/minimize")
    W3C_FIND_ELEMENT = Array("POST", "/session/{session id}/element")
    W3C_FIND_ELEMENT_FROM_ELEMENT = Array("POST", "/session/{session id}/element/{element id}/element")
    W3C_GET_ACTIVE_ELEMENT = Array("GET", "/session/{session id}/element/active")
    W3C_GET_ELEMENT_ATTRIBUTE = Array("GET", "/session/{session id}/element/{element id}/attribute/{name}")
    W3C_GET_ELEMENT_CSS_VALUE = Array("GET", "/session/{session id}/element/{element id}/css/{property name}")
    W3C_GET_ELEMENT_TAG_NAME = Array("GET", "/session/{session id}/element/{element id}/name")
    W3C_IS_ELEMENT_ENABLED = Array("GET", "/session/{session id}/element/{element id}/enabled")
    W3C_ELEMENT_CLEAR = Array("POST", "/session/{session id}/element/{element id}/clear")
    W3C_GET_PAGE_SOURCE = Array("GET", "/session/{session id}/source")
    W3C_EXECUTE_ASYNC_SCRIPT = Array("POST", "/session/{session id}/execute/async")
    W3C_GET_NAMED_COOKIE = Array("GET", "/session/{session id}/cookie/{name}")
    W3C_DELETE_COOKIE = Array("DELETE", "/session/{session id}/cookie/{name}")
    W3C_PERFORM_ACTIONS = Array("POST", "/session/{session id}/actions")
    W3C_DISMISS_ALERT = Array("POST", "/session/{session id}/alert/dismiss")
    W3C_GET_ALERT_TEXT = Array("GET", "/session/{session id}/alert/text")
    W3C_TAKE_SCREENSHOT = Array("GET", "/session/{session id}/screenshot")
    W3C_TAKE_ELEMENT_SCREENSHOT = Array("GET", "/session/{session id}/element/{element id}/screenshot")
End Sub

'------------------------------------------------------------------
' Browser Implementation
'------------------------------------------------------------------
Public Sub Edge(Optional ByVal driverPath As String = "msedgedriver.exe", _
                Optional ByVal port As Integer = 9516)
    Dim cmdLaunch As String
    
    cmdLaunch = driverPath & " --port=" & CStr(port)
    RemoteEndUrl = "http://localhost:" & CStr(port)
    PID = Shell(cmdLaunch, vbMinimizedNoFocus)
End Sub

Public Sub Chrome(Optional ByVal driverPath As String = "chromedriver.exe", _
                  Optional ByVal port As Integer = 9515)
    Dim cmdLaunch As String
    
    cmdLaunch = driverPath & " --port=" & CStr(port)
    RemoteEndUrl = "http://localhost:" & CStr(port)
    PID = Shell(cmdLaunch, vbMinimizedNoFocus)
End Sub

Public Sub Firefox(Optional ByVal driverPath As String = "geckodriver.exe", _
                   Optional ByVal port As Integer = 4444)
    Dim cmdLaunch As String
    
    cmdLaunch = driverPath & " --port " & CStr(port)
    RemoteEndUrl = "http://localhost:" & CStr(port)
    PID = Shell(cmdLaunch, vbMinimizedNoFocus)
End Sub

Public Sub InternetExplorer(Optional ByVal driverPath As String = "IEDriverServer.exe", _
                            Optional ByVal port = 5555)
    Dim cmdLaunch As String
    
    cmdLaunch = driverPath & " /port=" & CStr(port)
    RemoteEndUrl = "http://localhost:" & CStr(port)
    PID = Shell(cmdLaunch, vbMinimizedNoFocus)
End Sub

Public Sub Quit()
    Shell "taskkill /pid " & PID
End Sub

'------------------------------------------------------------------
' W3C Standard Endpoints
' https://www.w3.org/TR/webdriver/
'------------------------------------------------------------------
' Optimize for similar commands?
Public Sub OpenBrowser(Optional ByVal browserOptions As Options = Nothing)
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
    
    If browserOptions Is Nothing Then
        requestBody.Add "capabilities", New Dictionary
    Else
        requestBody.Add "capabilities", browserOptions.ConstructCapabilities()
    End If
    
    Set responseBody = SendCommand2Webdriver(W3C_NEW_SESSION, urlParameters, requestBody)

    Capabilities = responseBody("value")
End Sub

Public Sub CloseBrowser()
    Dim urlParameters As New Dictionary
        
    urlParameters.Add "{session id}", Capabilities("sessionId")
    
    SendCommand2Webdriver W3C_DELETE_SESSION, urlParameters
End Sub

Public Function GetStatus()
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
    
    Set responseBody = SendCommand2Webdriver(W3C_STATUS, urlParameters, requestBody)
    
    Set GetStatus = responseBody("value")
End Function

Public Function GetTimeouts()
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
    
    urlParameters.Add "{session id}", Capabilities("sessionId")
    
    Set responseBody = SendCommand2Webdriver(W3C_GET_TIMEOUTS, urlParameters, requestBody)

    Set GetTimeouts = responseBody("value")
End Function

Public Sub SetTimeouts(Optional ByVal script As Long = 30000, _
                            Optional ByVal pageLoad As Long = 300000, _
                            Optional ByVal implicit As Long = 0)
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary
        
    urlParameters.Add "{session id}", Capabilities("sessionId")
    requestBody.Add "script", script
    requestBody.Add "pageLoad", pageLoad
    requestBody.Add "implicit", implicit    ' a time to wait in milliseconds for the element location strategy
    
    SendCommand2Webdriver W3C_SET_TIMEOUTS, urlParameters, requestBody
End Sub

Public Sub NavigateTo(ByVal url As String)
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary
        
    urlParameters.Add "{session id}", Capabilities("sessionId")
    requestBody.Add "url", url
    
    SendCommand2Webdriver W3C_GO, urlParameters, requestBody
End Sub

Public Function GetCurrentURL() As String
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
    
    urlParameters.Add "{session id}", Capabilities("sessionId")
    
    Set responseBody = SendCommand2Webdriver(W3C_GET_CURRENT_URL, urlParameters, requestBody)
    
    GetCurrentURL = responseBody("value")
End Function

Public Sub Back()
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary
        
    urlParameters.Add "{session id}", Capabilities("sessionId")
    
    SendCommand2Webdriver W3C_BACK, urlParameters, requestBody
End Sub

Public Sub Forward()
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary
    
    urlParameters.Add "{session id}", Capabilities("sessionId")
    
    SendCommand2Webdriver W3C_FORWARD, urlParameters, requestBody
End Sub

Public Sub Refresh()
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary
    
    urlParameters.Add "{session id}", Capabilities("sessionId")
    
    SendCommand2Webdriver W3C_REFRESH, urlParameters, requestBody
End Sub

Public Function GetTitle() As String
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
    
    urlParameters.Add "{session id}", Capabilities("sessionId")
    
    Set responseBody = SendCommand2Webdriver(W3C_GET_TITLE, urlParameters, requestBody)
    
    GetTitle = responseBody("value")
End Function

Public Function GetWindowHandle()
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
        
    urlParameters.Add "{session id}", Capabilities("sessionId")
    
    Set responseBody = SendCommand2Webdriver(W3C_GET_WINDOW_HANDLE, urlParameters, requestBody)
    
    GetWindowHandle = responseBody("value")
End Function

Public Sub CloseWindow()
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary
    
    urlParameters.Add "{session id}", Capabilities("sessionId")
    
    SendCommand2Webdriver W3C_CLOSE_WINDOW, urlParameters, requestBody
End Sub

Public Sub SwitchToWindow(ByVal handle As String)
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary
    
    urlParameters.Add "{session id}", Capabilities("sessionId")
    requestBody.Add "handle", handle

    SendCommand2Webdriver W3C_SWITCH_TO_WINDOW, urlParameters, requestBody
End Sub

Public Function GetWindowHandles()
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
    
    urlParameters.Add "{session id}", Capabilities("sessionId")
            
    Set responseBody = SendCommand2Webdriver(W3C_GET_WINDOW_HANDLES, urlParameters, requestBody)
    
    Set GetWindowHandles = responseBody("value")
End Function

' Return: JSON Object
' "handle": The value of handle.
' "type": Let type be "tab" if the newly created window shares an OS-level window with the current browsing context, or "window" otherwise.
Public Function NewWindow() As Dictionary
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
    
    urlParameters.Add "{session id}", Capabilities("sessionId")
    
    Set responseBody = SendCommand2Webdriver(W3C_NEW_WINDOW, urlParameters, requestBody)
    
    Set NewWindow = responseBody("value")
End Function

Public Function SwitchToFrame(ByVal frame As WebElement) As Boolean
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
    
    urlParameters.Add "{session id}", frame.driver.Capabilities("sessionId")
    requestBody.Add "id", frame.elementJson
    
    Set responseBody = SendCommand2Webdriver(W3C_SWITCH_TO_FRAME, urlParameters, requestBody)
    
    SwitchToFrame = True
End Function

Public Function SwitchToParentFrame() As Boolean
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
    
    urlParameters.Add "{session id}", Capabilities("sessionId")
    
    Set responseBody = SendCommand2Webdriver(W3C_SWITCH_TO_PARENT_FRAME, urlParameters, requestBody)
    
    SwitchToParentFrame = True
End Function

Public Function GetWindowRect() As Dictionary
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
    
    urlParameters.Add "{session id}", Capabilities("sessionId")
    Set responseBody = SendCommand2Webdriver(W3C_GET_WINDOW_RECT, urlParameters, requestBody)
    
    Set GetWindowRect = responseBody("value")
End Function

Public Function SetWindowRect(Optional ByVal x As Integer, _
                              Optional ByVal y As Integer, _
                              Optional ByVal width As Integer, _
                              Optional ByVal height As Integer)
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
    
    urlParameters.Add "{session id}", Capabilities("sessionId")
    
    If x <> 0 Then requestBody.Add "x", x
    If y <> 0 Then requestBody.Add "y", y
    If width <> 0 Then requestBody.Add "width", width
    If height <> 0 Then requestBody.Add "height", height
    
    Set responseBody = SendCommand2Webdriver(W3C_SET_WINDOW_RECT, urlParameters, requestBody)
    
    Set SetWindowRect = responseBody("value")
End Function

' Return: JSON Object
' "x": the screenX attribute of the window object
' "y": the screenY attribute of the window object
' "width": the width of the outer dimensions of the top-level browsing context, including browser chrome etc...
' "height": the height of the outer dimensions of the top-level browsing context, including browser chrome etc...
Public Function MaximizeWindow() As Dictionary
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
    
    urlParameters.Add "{session id}", Capabilities("sessionId")
    
    Set responseBody = SendCommand2Webdriver(W3C_MAXIMIZE_WINDOW, urlParameters, requestBody)
    
    Set MaximizeWindow = responseBody("value")
End Function

' Return: JSON Object
' "x": the screenX attribute of the window object
' "y": the screenY attribute of the window object
' "width": the width of the outer dimensions of the top-level browsing context, including browser chrome etc...
' "height": the height of the outer dimensions of the top-level browsing context, including browser chrome etc...
Public Function MinimizeWindow()
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
    
    urlParameters.Add "{session id}", Capabilities("sessionId")
    
    Set responseBody = SendCommand2Webdriver(W3C_MINIMIZE_WINDOW, urlParameters, requestBody)
    
    Set MinimizeWindow = responseBody("value")
End Function

Public Function FullscreenWindow()
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
    
    urlParameters.Add "{session id}", Capabilities("sessionId")
    
    Set responseBody = SendCommand2Webdriver(W3C_FULLSCREEN_WINDOW, urlParameters, requestBody)
        
    Set FullscreenWindow = responseBody("value")
End Function

Public Function FindElement(ByVal by_strategy As By, _
                            ByVal target As String) As WebElement
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
    
    urlParameters.Add "{session id}", Capabilities("sessionId")
    Set requestBody = SelectLocationStrategies(by_strategy, target)

    Set responseBody = SendCommand2Webdriver(W3C_FIND_ELEMENT, urlParameters, requestBody)
    
    Set FindElement = Convert2WebElement(responseBody("value"))
End Function

Public Function FindElements(ByVal by_strategy As By, _
                            ByVal target As String) As Variant
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
    
    Dim elements() As WebElement
       
    urlParameters.Add "{session id}", Capabilities("sessionId")
    Set requestBody = SelectLocationStrategies(by_strategy, target)
    
    Set responseBody = SendCommand2Webdriver(W3C_FIND_ELEMENTS, urlParameters, requestBody)
    
    Dim counter As Integer
    Dim elementJson As Dictionary
    ReDim elements(responseBody("value").Count - 1)
    
    counter = 0
    For Each elementJson In responseBody("value")
        ReDim Preserve elements(counter)
        Set elements(counter) = Convert2WebElement(elementJson)
        counter = counter + 1
    Next
    
    FindElements = elements
End Function

Public Function FindElementFromElement(ByVal element As WebElement, _
                                       ByVal by_strategy As By, _
                                       ByVal target As String) As WebElement
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
        
    urlParameters.Add "{session id}", Capabilities("sessionId")
    urlParameters.Add "{element id}", element.elementJson(ELEMENT_KEY)
    
    Set requestBody = SelectLocationStrategies(by_strategy, target)
    
    Set responseBody = SendCommand2Webdriver(W3C_FIND_ELEMENT_FROM_ELEMENT, urlParameters, requestBody)
    
    Set FindElementFromElement = Convert2WebElement(responseBody("value"))
End Function

Public Function FindElementsFromElement(ByVal element As WebElement, _
                                       ByVal by_strategy As By, _
                                       ByVal target As String) As Variant
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
    
    Dim elements() As WebElement
    
    urlParameters.Add "{session id}", Capabilities("sessionId")
    urlParameters.Add "{element id}", element.elementJson(ELEMENT_KEY)
    
    Set requestBody = SelectLocationStrategies(by_strategy, target)
    
    Set responseBody = SendCommand2Webdriver(W3C_FIND_ELEMENTS_FROM_ELEMENT, urlParameters, requestBody)
    
    Dim counter As Integer
    Dim elementJson As Dictionary
    ReDim elements(responseBody("value").Count - 1)
    
    counter = 0
    For Each elementJson In responseBody("value")
        ReDim Preserve elements(counter)
        Set elements(counter) = Convert2WebElement(elementJson)
        counter = counter + 1
    Next
    
    FindElementsFromElement = elements
End Function

Public Function GetElementAttribute(ByVal element As WebElement, _
                                    ByVal name As String)
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
        
    urlParameters.Add "{session id}", Capabilities("sessionId")
    urlParameters.Add "{element id}", element.elementJson(ELEMENT_KEY)
    urlParameters.Add "{name}", name
    
    Set responseBody = SendCommand2Webdriver(W3C_GET_ELEMENT_ATTRIBUTE, urlParameters, requestBody)
    
    GetElementAttribute = responseBody("value")
End Function

Public Function GetElementText(ByVal element As WebElement) As String
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
    
    urlParameters.Add "{session id}", Capabilities("sessionId")
    urlParameters.Add "{element id}", element.elementJson(ELEMENT_KEY)
    
    Set responseBody = SendCommand2Webdriver(W3C_GET_ELEMENT_TEXT, urlParameters, requestBody)
    
    GetElementText = responseBody("value")
End Function

Public Sub ElementClick(ByVal element As WebElement)
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary

    urlParameters.Add "{session id}", Capabilities("sessionId")
    urlParameters.Add "{element id}", element.elementJson(ELEMENT_KEY)

    SendCommand2Webdriver W3C_ELEMENT_CLICK, urlParameters, requestBody
End Sub

Public Sub ElementClear(ByVal element As WebElement)
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary
    
    urlParameters.Add "{session id}", Capabilities("sessionId")
    urlParameters.Add "{element id}", element.elementJson(ELEMENT_KEY)
    
    SendCommand2Webdriver W3C_ELEMENT_CLEAR, urlParameters, requestBody
End Sub

Public Sub ElementSendKeys(text As String, ByVal element As WebElement)
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary
    
    urlParameters.Add "{session id}", Capabilities("sessionId")
    urlParameters.Add "{element id}", element.elementJson(ELEMENT_KEY)
    
    requestBody.Add "text", text
    
    SendCommand2Webdriver W3C_ELEMENT_SEND_KEYS, urlParameters, requestBody
End Sub

Public Function GetPageSource()
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary

    urlParameters.Add "{session id}", Capabilities("sessionId")
    
    Set responseBody = SendCommand2Webdriver(W3C_GET_PAGE_SOURCE, urlParameters, requestBody)
    
    GetPageSource = responseBody("value")
End Function

Public Function ExecuteScript(ByVal script As String, ParamArray args() As Variant) As Variant
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
    
    urlParameters.Add "{session id}", Capabilities("sessionId")
    
    requestBody.Add "script", script
    requestBody.Add "args", args
    
    Set responseBody = SendCommand2Webdriver(W3C_EXECUTE_SCRIPT, urlParameters, requestBody)
    
    ExecuteScript = responseBody("value")
End Function

Public Function ExecuteAsyncScript(ByVal script As String, ParamArray args() As Variant) As Variant
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
    
    urlParameters.Add "{session id}", Capabilities("sessionId")
    
    requestBody.Add "script", script
    requestBody.Add "args", args
    
    Set responseBody = SendCommand2Webdriver(W3C_EXECUTE_ASYNC_SCRIPT, urlParameters, requestBody)
    
    ExecuteAsyncScript = responseBody("value")
End Function

Public Sub TakeScreenshot(ByVal savePath As String)
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
    
    urlParameters.Add "{session id}", Capabilities("sessionId")
    
    Set responseBody = SendCommand2Webdriver(W3C_TAKE_SCREENSHOT, urlParameters, requestBody)
    
    DecodeBase64ToFile responseBody("value"), savePath
End Sub

Public Sub TakeElementScreenshot(ByVal element As WebElement, ByVal savePath As String)
    Dim urlParameters As New Dictionary, _
        requestBody As New Dictionary, _
        responseBody As Dictionary
    
    urlParameters.Add "{session id}", Capabilities("sessionId")
    urlParameters.Add "{element id}", element.elementJson(ELEMENT_KEY)
    
    requestBody.Add "scroll", True
    
    Set responseBody = SendCommand2Webdriver(W3C_TAKE_ELEMENT_SCREENSHOT, urlParameters, requestBody)

    DecodeBase64ToFile responseBody("value"), savePath
End Sub

'------------------------------------------------------------------
' Communication Interface between 'Local end' & 'Remote end'
'------------------------------------------------------------------
Private Function SendCommand2Webdriver(ByVal driverCommand As Variant, _
                Optional urlParameters As Dictionary = Nothing, _
                Optional requestBody As Dictionary = Nothing) As Dictionary
    Dim method As String: method = driverCommand(0)
    Dim path As String: path = driverCommand(1)
    
    ' Replace url parameters with user defined values.
    ' {session id}, {element id}, etc..
    Dim paramKey As Variant
    For Each paramKey In urlParameters
        path = Replace(path, paramKey, urlParameters(paramKey))
    Next
    
    Dim xmlHttp As Object
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    
    xmlHttp.Open method, RemoteEndUrl + path
    xmlHttp.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    xmlHttp.setRequestHeader "Cache-Control", "no-cache"
    xmlHttp.Send JsonConverter.ConvertToJson(requestBody)

    Do While (xmlHttp.readyState <> 4)
        DoEvents
    Loop

    If xmlHttp.Status = 200 Then
        Set SendCommand2Webdriver = JsonConverter.ParseJson(xmlHttp.responseText)
    Else
        Log4Driver xmlHttp.responseText
        Err.Raise vbObjectError + 513, "WebDriver", xmlHttp.responseText
    End If
    
    Set xmlHttp = Nothing
End Function

'------------------------------------------------------------------
' Other Functions
'------------------------------------------------------------------
' Locator strategies
' Ref:
' https://github.com/SeleniumHQ/selenium/blob/13d8f8be751001d44df8e5f1797518f4fb4dec6b/java/client/src/org/openqa/selenium/remote/http/W3CHttpCommandCodec.java#L187
Private Function SelectLocationStrategies(ByVal Strategy As By, ByVal value As String) As Dictionary
    Dim searchJson As New Dictionary
    
    Select Case True
        Case Strategy = By.ClassName
            searchJson.Add "using", "css selector"
            searchJson.Add "value", "." & value
        Case Strategy = By.ID
            searchJson.Add "using", "css selector"
            searchJson.Add "value", "[id=""" + value + """]"
        Case Strategy = By.name
            searchJson.Add "using", "css selector"
            searchJson.Add "value", "[name='" & value & "']"
        Case Strategy = By.LinkText
            searchJson.Add "using", "link text"
        Case Strategy = By.PartialLinkText
            searchJson.Add "using", "partial link text"
        Case Strategy = By.TagName
            searchJson.Add "using", "css selector"
            searchJson.Add "value", value
        Case Strategy = By.XPath
            searchJson.Add "using", "xpath"
            searchJson.Add "value", value
    End Select
    
    Set SelectLocationStrategies = searchJson
End Function

Private Function Convert2WebElement(ByVal elementJson As Dictionary) As WebElement
    Dim element As New WebElement
    element.driver = Me
        
    element.elementJson = elementJson
            
    Set Convert2WebElement = element
End Function

Private Sub DecodeBase64ToFile(ByVal Base64String As String, ByVal savePath As String)
    Dim base64Byte() As Byte

    With CreateObject("MSXML2.DOMDocument").createElement("b64")
        .DataType = "bin.base64"
        .text = Base64String
        base64Byte = .nodeTypedValue
    End With
    
    Dim FileNumber As Long: FileNumber = FreeFile
    
    Open savePath For Binary Access Write As #FileNumber
    Put #FileNumber, 1, base64Byte
    Close #FileNumber
End Sub

Private Sub Log4Driver(ByVal errorMesssage As String)
    Dim INT_FILE As Integer
    INT_FILE = FreeFile()
    
    Open ThisWorkbook.path + "./log4driver" & format(DateTime.Now, "yyyy-MM-dd") & ".txt" For Append As #INT_FILE
    Print #INT_FILE, format(DateTime.Now, "yyyy-MM-dd hh:mm:ss") + errorMesssage
    
    Close #INT_FILE
End Sub
