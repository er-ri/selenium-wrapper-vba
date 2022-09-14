# SeleniumWrapperVBA
A Selenium-WebDriver-based browser automation framework implemented for VBA.

*  [日本語](https://github.com/er-ri/selenium-wrapper-vba/blob/main/README.JA.md)

<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li><a href="#about-the-project">About The Project</a></li>
    <li><a href="#requirements">Requirements</a></li>
    <li><a href="#getting-started">Getting Started</a></li>
    <li><a href="#usage">Usage</a></li>
      <ul>
        <li><a href="#element-retrieval">Element Retrieval</a></li>
        <li><a href="#timeouts">Timeouts</a></li>
        <li><a href="#working-with-iframe">Working with iframe</a></li>
        <li><a href="#working-with-multiple-windows">Working with multiple windows</a></li>
        <li><a href="#execute-javascript">Execute JavaScript</a></li>
        <li><a href="#execute-async-javascript">Execute Async JavaScript</a></li>
        <li><a href="#take-screenshot">Take Screenshot</a></li>
        <li><a href="#take-element-screenshot">Take Element Screenshot</a></li>
        <li><a href="#enable-edge-ie-mode">Enable Edge IE-mode</a></li>
        <li><a href="#headless-mode">Headless mode</a></li>
        <li><a href="#customize-user-agent">Customize User-Agent</a></li>
      </ul>
    <li><a href="#roadmap">Roadmap</a></li>
    <li><a href="#license">License</a></li>
    <li><a href="#contribution">Contribution</a></li>
    <li><a href="#references">References</a></li>
  </ol>
</details>

## About The Project
The project implements the `endpoint node command` defined in [W3C WebDriver specification](https://www.w3.org/TR/webdriver/#endpoints) through VBA. You can use the project to do browser automation without installing a programming language such as Python, Java, etc. However, excel and a browser-specific driver are required.

## Requirements
1. Excel 32bit or 64bit
2. Browser's driver(Supported Browsers: Firefox, Chrome, Edge and Internet Explorer)

##  Getting Started
1. Download the browser-specific drivers, [`geckodriver`](https://github.com/mozilla/geckodriver/releases),
[`chromedriver`](https://chromedriver.chromium.org/), 
[`edgedriver`](https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/), or
[`iedriver`](https://www.selenium.dev/downloads/)
2. Import `Options.cls`, `WebDriver.cls`, `WebElement.cls` and `JsonConverter.bas` into your Excel. (Open VBA Editor, `Alt + F11`; File > Import File) 
   * where `JsonConverter.bas`, a JSON Parser for VBA created and maintained by [@timhall](https://github.com/timhall). For more details, see [`here`](https://github.com/VBA-tools/VBA-JSON).
3. Include a reference to "Microsoft Scripting Runtime". (Tools->References Check "`Microsoft Scripting Runtime`")

#### Note
* Add browser's driver in the system `PATH`, or you can also specify the path when launching the corresponding browser's driver.
* Error's message will be output at the same directory of the Excel workbook with the name of "log4driver{YYYY-MM-DD}.txt", where {YYYY-MM-DD} is current date.
* Some configurations are required before using `iedriver`, see [here](https://www.selenium.dev/documentation/ie_driver_server/#required-configuration) for more details about the configurations. 
 
#### Example
```vba
Sub Example()
    Dim driver As New WebDriver

    driver.Chrome
    driver.OpenBrowser
    driver.NavigateTo "https://www.python.org/"
    driver.MaximizeWindow
    driver.FindElement(By.ID, "id-search-field").SendKeys "machine learning"
    driver.FindElement(By.ID, "submit").Click
    driver.TakeScreenshot ThisWorkbook.path + "./screenshot.png"
    driver.MinimizeWindow
    driver.CloseBrowser
    
    driver.Quit
    Set driver = Nothing
End Sub
```

## Usage
### Element Retrieval
#### Find Element
```vba
    ' Locator Strategies:
    Dim e1, e2, e3, e4, e5, e6, e7 As WebElement
    set e1 = driver.FindElement(By.ID, "id")
    set e2 = driver.FindElement(By.ClassName, "blue")
    set e3 = driver.FindElement(By.Name, "name")
    set e4 = driver.FindElement(By.LinkText, "www.google.com")
    set e5 = driver.FindElement(By.PartialLinkText, "www.googl")
    set e6 = driver.FindElement(By.TagName, "div")
    set e7 = driver.FindElement(By.XPath, "/html/body/div[1]/div[3]")
```

#### Find Elements
```vba
    Dim elements() As WebElement
    elements = driver.FindElements(By.TagName, "a")
    
    Dim element As Variant
    For Each element In elements
        ' Do your stuff
    Next element
```

#### Find Element Frome Element
```vba
    Dim elementRoot As WebElement
    Set elementRoot = driver.FindElement(By.ID, "root1")
    Dim elementChild As WebElement
    Set elementChild = driver.FindElementFromElement(elmentRoot, By.TagName, "div")
```

#### Find Elements Frome Element
```vba
    Dim elementRoot As WebElement
    Set elementRoot = driver.FindElement(By.ID, "root1")
    Dim elementChildren() As WebElement
    elementChildren() = driver.FindElementsFromElement(elmentRoot, By.TagName, "p")
```

### Timeouts
#### Get Timeouts
```vba
    Dim timeoutsDict As Dictionary
    Set timeoutsDict = driver.GetTimeouts()
    Debug.Print timeoutsDict("script")    ' 30000
    Debug.Print timeoutsDict("pageLoad")  ' 300000
    Debug.Print timeoutsDict("implicit")  ' 0
```

#### Set Timeouts
```vba
    ' Set "script":40000,"pageLoad":500000,"implicit":15000
    driver.SetTimeouts 40000, 500000, 15000
```

### Working with iframe
```vba
    Set iframe1 = driver.FindElement(By.ID, "iframe1")
    driver.SwitchToFrame iframe1
    ' Perform some operations...
    driver.SwitchToParentFrame    ' switch back
```

### Working with multiple windows
```vba
    ' Get current windows's handle.
    Dim hwnd As String
    hwnd = driver.GetWindowHandle
    ' Get the handles of all the windows.
    Dim hWnds As New Collection
    Set hWnds = driver.GetWindowHandles
    ' Switch to another window.
    driver.SwitchToWindow (driver.GetWindowHandles(2))
```

### Execute JavaScript
```vba
    ' No parameters, no return value.
    driver.ExecuteScript "alert('Hello world!');"
    ' Accept parameters, no return value.
    driver.ExecuteScript "alert('Proof: ' + arguments[0] + arguments[1]);", "1+1=", 2
    ' Accept parameters, return the result. 
    Dim result As Long
    result = driver.ExecuteScript("let result = arguments[0] + arguments[1];return result;", 1, 2)
    Debug.Print result  ' 3
```

### Execute Async JavaScript
```vba
    ' No parameters, no return value.
    driver.ExecuteAsyncScript "alert('Hello world!');"
    ' Accept parameters, no return value.
    driver.ExecuteAsyncScript "alert('Proof: ' + arguments[0] + arguments[1]);", "1+1=", 2
    ' Accept parameters, return the result. 
    Dim result As Long
    result = driver.ExecuteAsyncScript("let result = arguments[0] + arguments[1];return result;", 1, 2)
    Debug.Print result  ' 3
```

### Screenshot
#### Take Screenshot
```vba
    ' Take the current webpage screenshot and save it to the specific path.
    driver.TakeScreenshot ThisWorkbook.path + "./1.png"
```

#### Take Element Screenshot
```vba
    ' Take the element screenshot directly.
    driver.FindElement(By.ID, "selenium_logo").TakeScreenshot ThisWorkbook.path + "./logo.png"
    ' or
    Dim seleniumLogo As WebElement
    Set seleniumLogo = driver.FindElement(By.ID, "selenium_logo")
    seleniumLogo.TakeScreenshot ThisWorkbook.path + "./logo.png"
```

### Enable Edge IE-mode
```vba
    Dim driver As New WebDriver
    Dim ieOptions As New Options
    ieOptions.BrowserType = InternetExplorer
    ieOptions.IntroduceFlakinessByIgnoringSecurityDomains = True    ' Optional
    ieOptions.IgnoreZoomSetting = True  ' Optional
    ieOptions.AttachToEdgeChrome = True
    ieOptions.EdgeExecutablePath = "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
    
    driver.InternetExplorer "C:\WebDriver\IEDriverServer_Win32_4.0.0\IEDriverServer.exe"
    driver.OpenBrowser ieOptions
```

### Headless mode
#### Start Chrome in headless mode
```vba
    Dim driver As New WebDriver
    Dim chromeOptions As New Options
    chromeOptions.BrowserType = Chrome
    chromeOptions.ChromeArguments.add "--headless"

    driver.Chrome "C:\WebDriver\chromedriver_win32\chromedriver.exe"
    driver.OpenBrowser chromeOptions
```

#### Start Firefox in headless mode
```vba
    Dim driver As New WebDriver
    Dim firefoxOptions As New Options
    firefoxOptions.BrowserType = Firefox
    firefoxOptions.FirefoxArguments.Add "-headless"

    driver.Firefox "C:\WebDriver\Firefox\geckodriver.exe"
    driver.OpenBrowser firefoxOptions
```

### Customize User-Agent
#### Customize User-Agent in Chrome
```vba
    Dim driver As New WebDriver
    Dim chromeOptions As New Options
    
    driver.Chrome
    chromeOptions.BrowserType = Chrome
    chromeOptions.ChromeArguments.Add "--user-agent=my customized user-agent"
    driver.OpenBrowser chromeOptions
    driver.NavigateTo "https://www.whatismybrowser.com/detect/what-is-my-user-agent/"
```

## Roadmap
| Endpoint Node   Command        | Function Name           | Element Function Name |
|--------------------------------|-------------------------|-----------------------|
| New Session                    | OpenBrowser             |                       |
| Delete Session                 | CloseBrowser            |                       |
| Status                         | GetStatus               |                       |
| Get Timeouts                   | GetTimeouts             |                       |
| Set Timeouts                   | SetTimeouts             |                       |
| Navigate To                    | NavigateTo              |                       |
| Get Current URL                | GetCurrentURL           |                       |
| Back                           | Back                    |                       |
| Forward                        | Forward                 |                       |
| Refresh                        | Refresh                 |                       |
| Get Title                      | GetTitle                |                       |
| Get Window Handle              | GetWindowHandle         |                       |
| Close Window                   | CloseWindow             |                       |
| Switch To Window               | SwitchToWindow          |                       |
| Get Window Handles             | GetWindowHandles        |                       |
| New Window                     | NewWindow               |                       |
| Switch To Frame                | SwitchToFrame           |                       |
| Switch To Parent Frame         | SwitchToParentFrame     |                       |
| Get Window Rect                | GetWindowRect           |                       |
| Set Window Rect                | SetWindowRect           |                       |
| Maximize Window                | MaximizeWindow          |                       |
| Minimize Window                | MinimizeWindow          |                       |
| Fullscreen Window              | FullscreenWindow        |                       |
| Get Active Element             | Not yet                 |                       |
| Get Element Shadow Root        | Not yet                 |                       |
| Find Element                   | FindElement             |                       |
| Find Elements                  | FindElements            |                       |
| Find Element From Element      | FindElementFromElement  | FindElement           |
| Find Elements From Element     | FindElementsFromElement | FindElements          |
| Find Element From Shadow Root  | Not yet                 |                       |
| Find Elements From Shadow Root | Not yet                 |                       |
| Is Element Selected            | Not yet                 |                       |
| Get Element Attribute          | GetElementAttribute     | GetAttribute          |
| Get Element Property           | Not yet                 |                       |
| Get Element CSS Value          | Not yet                 |                       |
| Get Element Text               | GetElementText          | GetText               |
| Get Element Tag Name           | Not yet                 |                       |
| Get Element Rect               | Not yet                 |                       |
| Is Element Enabled             | Not yet                 |                       |
| Get Computed Role              | Not yet                 |                       |
| Get Computed Label             | Not yet                 |                       |
| Element Click                  | ElementClick            | Click                 |
| Element Clear                  | ElementClear            | Clear                 |
| Element Send Keys              | ElementSendKeys         | SendKeys              |
| Get Page Source                | GetPageSource           |                       |
| Execute Script                 | ExecuteScript           |                       |
| Execute Async Script           | ExecuteAsyncScript      |                       |
| Get All Cookies                | Not yet                 |                       |
| Get Named Cookie               | Not yet                 |                       |
| Add Cookie                     | Not yet                 |                       |
| Delete Cookie                  | Not yet                 |                       |
| Delete All Cookies             | Not yet                 |                       |
| Perform Actions                | Not yet                 |                       |
| Release Actions                | Not yet                 |                       |
| Dismiss Alert                  | Not yet                 |                       |
| Accept Alert                   | Not yet                 |                       |
| Get Alert Text                 | Not yet                 |                       |
| Send Alert Text                | Not yet                 |                       |
| Take Screenshot                | TakeScreenshot          |                       |
| Take Element Screenshot        | TakeElementScreenshot   | TakeScreenshot        |
| Print Page                     | Not yet                 |                       |
* Browser Capabilities are not listed above.

## License
Distributed under the MIT License. See `LICENSE.txt` for more information.

## Contribution
Any suggestions for improvement or contribution to this project are appreciated! Creating an issue or pull request!

## References
1. W3C WebDriver Working Draft:
   * https://www.w3.org/TR/webdriver/
2. The Selenium Browser Automation Project
   * https://www.selenium.dev/documentation/webdriver/
3. The W3C WebDriver Spec, A Simplified Guide:
   * https://github.com/jlipps/simple-wd-spec
4. geckodriver, WebDriver Reference
   * https://developer.mozilla.org/en-US/docs/Web/WebDriver
5. Capabilities & ChromeOptions
   * https://chromedriver.chromium.org/capabilities
6. Capabilities and EdgeOptions
   * https://docs.microsoft.com/en-us/microsoft-edge/webdriver-chromium/capabilities-edge-options
