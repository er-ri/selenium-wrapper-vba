# SeleniumWrapperVBA
エクセルVBA用のSelenium-WebDriverベースのブラウザー自動化フレームワーク

*  [English](https://github.com/er-ri/selenium-wrapper-vba/blob/main/README.md)
*  [中文](https://github.com/er-ri/selenium-wrapper-vba/blob/main/README.zh-CN.md)

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
      </ul>
    <li><a href="#license">License</a></li>
    <li><a href="#contribution">Contribution</a></li>
    <li><a href="#references">References</a></li>
  </ol>
</details>

## About The Project
このプロジェクトはプログラミング言語（例えPython、Javaなど）をインストールせずにブラウザの自動化を実現できる。必要なのはエクセルとブラウザドライバになります。

## Requirements
1. エクセル３２ビット又は６４ビット
2. ブラウザのウェブドライバ（サポートするブラウザ：Firefox、Chrome、EdgeとInternet Explorer）

## Getting Started
1. ブラウザのウェブドライバをダウンロードする、[`geckodriver`](https://github.com/mozilla/geckodriver/releases),
[`chromedriver`](https://chromedriver.chromium.org/), 
[`edgedriver`](https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/), or
[`iedriver`](https://www.selenium.dev/downloads/)
2. `Options.cls`、 `WebDriver.cls`、 `WebElement.cls`と`JsonConverter.bas`をエクセルにインポートする。(Open VBA Editor, `Alt + F11`; File > Import File) 
   * `JsonConverter.bas`は[@timhall](https://github.com/timhall)のプロジェクトとなります。詳細は[`こちら`](https://github.com/VBA-tools/VBA-JSON)に確認してください。
3. プロジェクト参照のところに"Microsoft Scripting Runtime"をチェックする。(Tools->References Check "`Microsoft Scripting Runtime`")

#### Note
* ウェブドライバのディレクトリを`PATH`に設定する。　設定しなかったらブラウザを起動するときにドライバの絶対パスを指定することもできます。
* エラーが発生するときにエラーメッセージが同ワークブックのフォルダに"log4driver{YYYY-MM-DD}.txt"というログファイルが作られます。{YYYY-MM-DD}は本日の日付です。
* `iedriver`を使用する前に設定が必要です。具体的な設定手順は[こちら](https://www.selenium.dev/documentation/ie_driver_server/#required-configuration)です。（英語）
 
#### Example
```vba
Sub Example()
    Dim driver As New WebDriver

    driver.Chrome
    driver.OpenBrowser
    driver.NavigateTo "https://www.google.com"
    driver.FindElement(By.name, "q").SendKeys "selenium wrapper vba"
    driver.FindElement(By.XPath, "/html/body/div[1]/div[3]/form/div[1]/div[1]/div[2]/div[2]/div[5]/center/input[1]").Click
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
    elementChildren() = driver.FindElementFromElement(elmentRoot, By.TagName, "p")
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

## License
MITライセンス。詳細は`LICENSE.txt`を確認してください.

## Contribution
ご意見やご不明な点がございましたら、お気軽に問い合わせください。

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
