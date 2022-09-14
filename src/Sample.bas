Attribute VB_Name = "Sample"
Option Explicit

Public driver As WebDriver

Sub Example()
    Set driver = New WebDriver

    driver.Chrome
    driver.OpenBrowser
    driver.MaximizeWindow
    driver.NavigateTo "https://www.selenium.dev/"
    driver.FindElement(By.XPath, "/html/body/header/nav/div/ul/li[4]/a").Click
    driver.TakeScreenshot ThisWorkbook.path + "./screenshot1.png"

    driver.CloseBrowser
    driver.Quit
    Set driver = Nothing
End Sub

Sub Example2()
    Set driver = New WebDriver

    driver.Edge
    driver.OpenBrowser
    driver.NavigateTo "https://www.python.org/"
    driver.MaximizeWindow
    driver.FindElement(By.ID, "id-search-field").SendKeys "machine learning"
    driver.FindElement(By.ID, "submit").Click
    driver.TakeScreenshot ThisWorkbook.path + "./screenshot2.png"
    driver.MinimizeWindow
    driver.CloseBrowser
    
    driver.Quit
    Set driver = Nothing
End Sub
