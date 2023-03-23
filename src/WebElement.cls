VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------
' SeleniumWrapperVBA
'
' Auther:   er-ri (https://github.com/er-ri/selenium-wrapper-vba)
' Date:     2022/03
' License:  MIT (https://opensource.org/licenses/MIT)
'------------------------------------------------------------------
'
' =================================================================
' #Class WebElement
' =================================================================
Option Explicit

Private m_Driver As WebDriver
Private m_elementJson As Dictionary

'------------------------------------------------------------------
' Encapsulation
'------------------------------------------------------------------
Public Property Get driver() As WebDriver
    Set driver = m_Driver
End Property

Public Property Let driver(ByVal new_Driver As WebDriver)
    Set m_Driver = new_Driver
End Property

Public Property Get elementJson() As Dictionary
    Set elementJson = m_elementJson
End Property

Public Property Let elementJson(ByVal new_elementJson As Dictionary)
    Set m_elementJson = new_elementJson
End Property

'------------------------------------------------------------------
' Element Command
'------------------------------------------------------------------
Public Function FindElement(ByVal by_strategy As By, _
                            ByVal target As String) As WebElement
    Set FindElement = driver.FindElementFromElement(Me, by_strategy, target)
End Function

Public Function FindElements(ByVal element As WebElement, _
                             ByVal by_strategy As By, _
                             ByVal target As String) As Variant
    FindElements = driver.FindElementsFromElement(Me, by_strategy, target)
End Function

Public Sub SendKeys(text As String)
    driver.ElementSendKeys text, Me
End Sub

Public Sub Clear()
    driver.ElementClear Me
End Sub

Public Sub Click()
    driver.ElementClick Me
End Sub

Public Function GetAttribute(ByVal name As String)
    GetAttribute = driver.GetElementAttribute(Me, name)
End Function

Public Function GetProperty(ByVal name As String)
    GetProperty = driver.GetElementProperty(Me, name)
End Function

Public Function GetText()
    GetText = driver.GetElementText(Me)
End Function

Public Sub TakeScreenshot(ByVal savePath As String)
    driver.TakeElementScreenshot Me, savePath
End Sub

