Imports System
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports NUnit.Framework
Imports Selenium
Imports Lean.Test.Automation.Framework.LibraryGlobal.libGlobal
Imports Lean.Test.Automation.API
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Remote
Imports OpenQA
Imports OpenQA.Selenium.IE
Imports OpenQA.Selenium.Chrome
Imports OpenQA.Selenium.Firefox
Imports OpenQA.Selenium.Safari
Imports System.Drawing.Imaging
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports OpenQA.Selenium.Interactions
Imports OpenQA.Selenium.Support.UI
Imports OpenQA.Selenium.Proxy
Imports OpenQA.Selenium.Appium.Interfaces
Imports OpenQA.Selenium.Appium.MultiTouch

Namespace LibraryAppium
    Public Class ScreenShotRemoteWebDriver1
        Inherits RemoteWebDriver
        Implements ITakesScreenshot
        Public Sub New(uri As Uri, dc As DesiredCapabilities)
            MyBase.New(uri, dc)
        End Sub

        Public Function GetScreenshots() As OpenQA.Selenium.Screenshot
            Dim screenshotResponse As Response = Me.Execute(DriverCommand.Screenshot, Nothing)
            Dim base64 As String = screenshotResponse.Value.ToString()
            Return New OpenQA.Selenium.Screenshot(base64)
        End Function
    End Class
    
    Public Class SeleniumAppiumHelper

        'Public driver As ScreenShotRemoteWebDriver
        Private capabilities As New DesiredCapabilities
        Private Shared testServerAddress As New Uri("http://127.0.0.1:4723/wd/hub")
        Private Shared INIT_TIMEOUT_SEC As TimeSpan = TimeSpan.FromSeconds(360)
        Public Enum Platform
            IOS
            Android
        End Enum
        
        Public Sub SetupTest()
            Select Case p_PlatformName
                Case "IOS"
                    Try

                        ''Dim apkpath As [String] = "D:\latest Android Apps\Test.apk"
                        ''Dim app As New File(apkpath)
                        ''capabilities.SetCapability("app", app.getAbsolutePath())
                        capabilities.SetCapability("platformName", "iOS")
                        capabilities.SetCapability("deviceName", "iPhone de Sergio")
                        capabilities.SetCapability("udid", "afdc1f22b9b7a0120b508987e09dac20e1427098")
                        capabilities.SetCapability("browser_Name", "iOS")
                        capabilities.SetCapability("platformVersion", "10.0.2")
                        capabilities.SetCapability("platform", "Mac")
                        'capabilities.SetCapability("appPackage", "com.expiretest")
                        'capabilities.SetCapability("appActivity", ".LoginActivity")
                        capabilities.SetCapability("app", "C:\LeanTestAutomation\app\TestApp.app")

                        objAppium = New RemoteWebDriver(testServerAddress, capabilities, INIT_TIMEOUT_SEC)
                    Catch ex As Exception

                    End Try
                Case "Android"
                    'capabilities.SetCapability("browserName", "")
                    capabilities.SetCapability("appium-version", "1.4.16.1")
                    capabilities.SetCapability("platformName", "Android")
                    capabilities.SetCapability("platformVersion", "")
                    capabilities.SetCapability("deviceName", p_Device)
                    capabilities.SetCapability("autoWebview", "False")
                    capabilities.SetCapability("app", p_pathUrlApp)

                    'capabilities.SetCapability("FullReset", False)
                    capabilities.SetCapability("noReset", True)
                    'capabilities.SetCapability("", "")
                    'capability.SetCapability("", "")
                    objAppium = New RemoteWebDriver(testServerAddress, capabilities, INIT_TIMEOUT_SEC)

                    'Dim pathServer As String = "C:\LeanTestAutomation\API\" 'System.IO.Directory.GetCurrentDirectory
                    ''Dim chromeOptions As New OpenQA.Selenium.Chrome.ChromeOptions
                    ''Dim proxy As New OpenQA.Selenium.Proxy
                    'Try
                    '    ' objAppium.Manage.Timeouts().SetPageLoadTimeout(System.TimeSpan.FromSeconds(p_timeout))
                    '    SelectBrowser(p_browserType)
                    '    Select Case p_appProcessName

                    '        Case "chrome"
                    '            objAppium = New ChromeDriver(pathServer) ', chromeOptions)
                    '        Case "iexplore"
                    '            objAppium = New InternetExplorerDriver(pathServer)
                    '        Case "firefox"
                    '            objAppium = New FirefoxDriver()
                    '        Case Else

                    '    End Select
                    'Catch ex As Exception
                    '    'handler error
                    'End Try
            End Select
        End Sub
        Public Sub TeardownTest()
            Try
                objAppium.Quit()
            Catch ex As Exception
                ' Ignore errors if unable to close the browser
            End Try
            ' Assert.AreEqual("", verificationErrors.ToString())
        End Sub
    End Class

    'class contains all methods to interactin with windows and elements
    Public Class InteractionAppium
    	'Private driver As New WDBrowserStackHelper
    	Dim libGlobal As LibraryGlobal.LibGlobal
        Function Open(url As String) As Boolean
            Try
                objAppium.Navigate.GoToUrl(url)
                Return True
            Catch ex As Exception
            	
                LibraryGlobal.LibGlobal.HandlerError("LibraryobjAppium.InteractionWDBrowseerStack.Open: " & url & " Error:" & ex.Message & " - " & ex.InnerException.ToString)
                Return False
            End Try
        End Function
        Function WaitForPageToLoad(Optional timeOut As Integer = 0) As Boolean
            timeOut = IIf(timeOut <> 0, timeOut, p_timeout)
            Try
                Console.WriteLine("WaitForPageToLoad: Timeout = " & p_timeout)
                objAppium.Manage.Timeouts().SetPageLoadTimeout(System.TimeSpan.FromSeconds(p_timeout))
            Catch ex As Exception
                LibraryGlobal.LibGlobal.HandlerError("LibraryobjAppium.InteractionWDBrowseerStack.WaitForPageToLoad: Error: " & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function

        'VERIFICAR*******************************************************
        Function OpenWindow(url As String, WindowID As String) As Boolean
            Try
                'objAppium.OpenWindow(url, WindowID)
                Return True
            Catch ex As Exception
                LibraryGlobal.LibGlobal.HandlerError("Library.Selenium.RC.Selenium.Interaction.OpenWindow: " & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
        'VERIFICAR*******************************************************
        Function SelectWindow(windowID As String) As Boolean
            Try
                'objSeleniumRC.SelectWindow(windowID)
                Return True
            Catch ex As Exception
                LibraryGlobal.LibGlobal.HandlerError("Library.Selenium.RC.Selenium.Interaction.SelectWindow: " & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
        Function Exist(element As String, waitMilliseconds As Integer, Optional typeIdentification As typeIdentification = typeIdentification.xpath) As Boolean

            Try
                Thread.Sleep(waitMilliseconds)
                Select Case typeIdentification
                    Case typeIdentification.xpath
                        If objAppium.FindElement(By.XPath(element)).Displayed Then
                            Return True
                        Else
                            Return False
                        End If
                    Case typeIdentification.name
                        If objAppium.FindElement(By.Name(element)).Displayed Then
                            Return True
                        Else
                            Return False
                        End If
                    Case typeIdentification.id
                        If objAppium.FindElement(By.Id(element)).Displayed Then
                            Return True
                        Else
                            Return False
                        End If
                End Select

            Catch ex As Exception
                'HandlerError("Library.Selenium.RC.Interaction.exist:" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
        Sub HigthLigth(element As String)
            'TODO higthLigth
        End Sub
        Function WaitExist(element As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath) As Boolean
            Try
                Test.HandlerMessage()
                For i = 0 To p_timeout - 1
                    Try
                        Select Case typeIdentification
                            Case typeIdentification.name
                                If objAppium.FindElement(By.Name(element)).Displayed Then
                                    Return True
                                End If
                            Case typeIdentification.id
                                If objAppium.FindElement(By.Id(element)).Displayed Then
                                    Return True
                                End If
                            Case typeIdentification.xpath
                                If objAppium.FindElement(By.XPath(element)).Displayed Then
                                    Return True
                                End If
                            Case typeIdentification.linkText
                                If objAppium.FindElement(By.LinkText(element)).Displayed Then
                                    Return True
                                End If
                        End Select
                    Catch ex As Exception
                        Console.Write("Element: " & element & " - Time: " & i & vbCrLf)
                        Thread.Sleep(1000) ' clique em pause, F11 e em seguida arraste o cursor para o End function
                    End Try
                Next
                'p_timeout = timeOut
                Test.HandlerMessage()
                Throw New Exception("waitLoadElement: Element=" & element & " not found")
                Return False
            Catch ex As Exception
                Throw New Exception(ex.Message)
                Return False
            End Try
        End Function
        'VERIFICAR*******************************************************
        Function GetHtmlSource()
            Try
                Return objAppium.PageSource.ToString()
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
        Sub Clear(element As String)
            Try
                Dim webElement As IWebElement = objAppium.FindElement(By.XPath(element))
                webElement.Clear()
            Catch ex As Exception

            End Try
        End Sub
        Sub TypeKeys(element As String, value As String)
            Try

                Dim webElement As IWebElement = objAppium.FindElement(By.XPath(element))
                webElement.SendKeys(value)
            Catch ex As Exception

            End Try
        End Sub
        Sub WindowMaximize()
            Try
                objAppium.Manage.Window.Maximize()
            Catch ex As Exception

            End Try
        End Sub
        'VERIFICAR*******************************************************
        Sub WindowFocus()
            Try
                ' objSeleniumRC.WindowFocus()
            Catch ex As Exception

            End Try
        End Sub
        Sub Refresh()
            Try
                objAppium.Navigate.Refresh()
            Catch ex As Exception

            End Try
        End Sub
        Function MouseOver(element As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath, Optional waitElement As Boolean = True) As Boolean
            Dim webElement As IWebElement = Nothing
            Try
                If String.IsNullOrEmpty(element) Then Return False
                If waitElement Then
                    If WaitExist(element, typeIdentification) Then

                        Try
                            Select Case typeIdentification
                                Case typeIdentification.name
                                    webElement = objAppium.FindElement(By.Name(element))
                                Case typeIdentification.id
                                    webElement = objAppium.FindElement(By.Id(element))
                                Case typeIdentification.xpath
                                    webElement = objAppium.FindElement(By.XPath(element))
                                Case typeIdentification.linkText
                                    webElement = objAppium.FindElement(By.LinkText(element))
                            End Select

                            'webElement.Location
                            Dim xx = webElement.Location
                            pc_dv.MouseMove(xx.X + 20, xx.Y + 74)
                            Return True
                        Catch ex As Exception
                            'Throw New Exception(ex.Message)
                            Return False
                        End Try
                    End If
                End If
                Return False
            Catch ex As Exception
                HandlerError("LibrarySeleniumRC.Interaction.click")
                Throw New Exception(ex.Message)
                Return False
            End Try
        End Function
        Sub Flick(x As Integer, y As Integer)
            Try
                Dim tou = New RemoteTouchScreen(objAppium)
                tou.Flick(x, y)
            Catch ex As Exception
            End Try
        End Sub
        Sub Down(x As Integer, y As Integer)
            Try
                Dim tou = New RemoteTouchScreen(objAppium)
                tou.Down(x, y)
            Catch ex As Exception
            End Try
        End Sub
        Sub Up(x As Integer, y As Integer)
            Try
                Dim tou = New RemoteTouchScreen(objAppium)
                tou.Up(x, y)
            Catch ex As Exception
            End Try
        End Sub
        Sub Move(x As Integer, y As Integer)
            Try
                Dim tou = New RemoteTouchScreen(objAppium)
                tou.Move(x, y)
            Catch ex As Exception
            End Try
        End Sub
        Sub DragAndDrop(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)
            Try
                Dim tou = New RemoteTouchScreen(objAppium)
                tou.Down(x1, y1)
                tou.Move(x2, y2)
                tou.Up(x2, y2)
            Catch ex As Exception
            End Try
        End Sub

        Function Click(element As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath, Optional waitElement As Boolean = True) As Boolean
            Dim webElement As IWebElement = Nothing
            Try
                If String.IsNullOrEmpty(element) Then Return False
                If waitElement Then
                    If WaitExist(element, typeIdentification) Then

                        Try
                            Select Case typeIdentification
                                Case typeIdentification.name
                                    webElement = objAppium.FindElement(By.Name(element))
                                Case typeIdentification.id
                                    webElement = objAppium.FindElement(By.Id(element))
                                Case typeIdentification.xpath
                                    webElement = objAppium.FindElement(By.XPath(element))
                                Case typeIdentification.linkText
                                    webElement = objAppium.FindElement(By.LinkText(element))
                            End Select


                            'webElement.Location
                            Dim xx = webElement.Location
                            pc_dv.MouseMove(xx.X + 20, xx.Y + 74)
                            webElement.Click()
                            Return True
                        Catch ex As Exception
                            'Throw New Exception(ex.Message)
                            Return False
                        End Try
                    End If
                End If
                Return False
            Catch ex As Exception
                HandlerError("LibrarySeleniumRC.Interaction.click")
                Throw New Exception(ex.Message)
                Return False
            End Try
        End Function
        Function DoubleClick(element As String) As Boolean

            Try
                If String.IsNullOrEmpty(element) Then Return False
                If WaitExist(element) Then
                    Try
                        Dim act As Actions = Nothing
                        act.DoubleClick(objAppium.FindElement(By.XPath(element))).Build().Perform()
                        Return True
                    Catch ex As Exception
                        HandlerError("LibrarySeleniumRC.Interaction.click")
                        Return False
                    End Try
                End If
                Return False
            Catch ex As Exception

            End Try
            HandlerError("LibrarySeleniumRC.Interaction.click")
            Return False
        End Function

        Public Function Type(element As String, value As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath) As Boolean
            Dim counTry As Integer = 0
            Try
                If String.IsNullOrEmpty(element) Then Return True
                If WaitExist(element, typeIdentification) Then
                    Try
                        Dim webElement As IWebElement = Nothing
                        Select Case typeIdentification
                            Case typeIdentification.xpath
                                webElement = objAppium.FindElement(By.XPath(element))
                            Case typeIdentification.name
                                webElement = objAppium.FindElement(By.Name(element))
                            Case typeIdentification.id
                                webElement = objAppium.FindElement(By.Id(element))
                        End Select
                        Try
                            webElement.Click()
                        Catch ex As Exception
                            Throw New Exception(ex.Message)
                            Return False
                        End Try
                        'Do While Not webElement.GetAttribute("value") = Trim(value)
                        webElement.Clear()
                        webElement.SendKeys(value)
                        Test.Wait(200)
                        'counTry += 1
                        'If counTry > 3 Then Throw New Exception("O valor digitado não corresponde ao valor apresentado pela tela")
                        'Loop
                        Return True
                    Catch ex As Exception

                    End Try
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("LibraryInteractionRC.Interaction.Type")
                Throw New Exception(ex.Message)
                Return False
            End Try
            Return Nothing
        End Function

        Function SelectValue(element As String, value As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath) As Boolean
            If String.IsNullOrEmpty(element) Then Return True
            If String.IsNullOrEmpty(value) Then Return True
            Dim captureValue As String = Nothing
            If WaitExist(element, typeIdentification) Then
                Dim webElement As IWebElement = Nothing
                Try
                    Do
                        Try

                            Select Case typeIdentification
                                Case typeIdentification.xpath
                                    webElement = objAppium.FindElement(By.XPath(element))
                                Case typeIdentification.name
                                    webElement = objAppium.FindElement(By.Name(element))
                                Case typeIdentification.id
                                    webElement = objAppium.FindElement(By.Id(element))
                            End Select
                            webElement.SendKeys(value)
                            Test.Wait(100)
                            captureValue = webElement.Text
                        Catch ex As Exception
                            Try
                                webElement.Click()
                                Test.SendKey(value)
                                Test.SendKey("{ENTER}")
                            Catch ex1 As Exception

                            End Try
                        End Try
                    Loop While captureValue = value 'esta errado, precisa ser corrigido. PRECISA SER <> e egar o valor correto
                    Return True
                Catch ex As Exception
                    Throw New Exception(ex.Message)
                    Return False
                End Try
            End If
            Return False
        End Function
        Function GetText(element As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath) As String
            If Not WaitExist(element, typeIdentification) Then Return Nothing
            Try
                Dim webElement As IWebElement = Nothing
                Select Case typeIdentification
                    Case typeIdentification.name
                        webElement = objAppium.FindElement(By.Name(element))
                    Case typeIdentification.id
                        webElement = objAppium.FindElement(By.Id(element))
                    Case typeIdentification.xpath
                        webElement = objAppium.FindElement(By.XPath(element))
                End Select
                Return webElement.GetAttribute("text")
            Catch ex As Exception
                'Throw New Exception(ex.Message)
                Return "Warning: Element not found! - Element: " & element
            End Try
        End Function
        Function GetValue(element As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath) As String
            'If Not WaitExist(element, typeIdentification) Then Return Nothing
            Try
                Dim webElement As IWebElement = Nothing
                Select Case typeIdentification
                    Case typeIdentification.name
                        webElement = objAppium.FindElement(By.Name(element))
                    Case typeIdentification.id
                        webElement = objAppium.FindElement(By.Id(element))
                    Case typeIdentification.xpath
                        webElement = objAppium.FindElement(By.XPath(element))
                End Select
                Return webElement.GetAttribute("value")
            Catch ex As Exception
                HandlerError("LibraryInteractionRC.Interaction.selectValue")
                Return Nothing
            End Try
        End Function

        Function GetTextPopup(Optional click As Boolean = True) As String
            'Dim msg As String = Nothing
            'Try
            '    Dim alert = objAppium.SwitchTo().Alert()
            '    msg = alert.Text
            '    If click Then alert.Accept()
            '    Return msg
            'Catch ex As Exception
            '    HandlerError("LibraryInteractionRC.Interaction.selectValue")
            '    Return Nothing
            'End Try
            Return Nothing
        End Function
        Function GetSelectedValue(element As String) As String
            If String.IsNullOrEmpty(element) Then Return Nothing
            If WaitExist(element) Then
                Try
                    Dim webElement As IWebElement = objAppium.FindElement(By.XPath(element))
                    Return webElement.GetAttribute("Value")
                Catch ex As Exception
                    HandlerError("LibraryInteractionRC.Interaction.selectValue")
                    Return Nothing
                End Try
            End If
            Return False
        End Function
        Function isEditable(element As String) As Boolean
            If String.IsNullOrEmpty(element) Then Return False
            If WaitExist(element) Then
                Try
                    If objAppium.FindElement(By.Id(element)).Enabled Then
                        Console.WriteLine("isEditable: " & element)
                        Return True
                    Else
                        Return False
                    End If
                Catch ex As Exception
                    HandlerError("LibraryInteractionRC.Interaction.waitLoadElement")
                    Return False
                End Try
            End If
            Return False
        End Function
        'VERIFICAR******************************************************
        Function GetCellData(element As String, Optional col As Integer = 0, Optional row As Integer = 0, Optional value As String = Nothing) As String
            Dim text As String = Nothing
            Try
                If Not WaitExist(element) Then Return Nothing
                For r = row To 100
                    For c = col To 100
                        Try
                            text = objSeleniumRC.GetTable(element & "." & r & "." & c)
                        Catch ex As Exception
                            Exit For
                        End Try
                        If text = value Then Return text
                    Next
                Next
                Return Nothing
            Catch ex As Exception
                HandlerError("LibraryInteractionRC.Interaction.GetCellData")
                Return Nothing
            End Try
        End Function
        'VERIFICAR******************************************************
        Function GetCellDataByReference(element As String, reference As String, colRefer As Integer, colTarget As Integer, Optional rowRefer As Integer = 0) As String

            Dim text As String = Nothing
            Try
                If Not WaitExist(element) Then Return Nothing
                For r = rowRefer To 100
                    Try
                        text = objSeleniumRC.GetTable(element & "." & r & "." & colRefer)
                        If text = reference Then
                            Return objSeleniumRC.GetTable(element & "." & r & "." & colTarget)
                        End If
                        Return text
                    Catch ex As Exception
                        Exit For
                    End Try
                Next
                Return Nothing
            Catch ex As Exception
                HandlerError("LibraryInteractionRC.Interaction.GetCellDataByReference")
                Return Nothing
            End Try
        End Function
        Public Sub TeardownTest()
            Try
                Do While Not IsNothing(objAppium)
                    objAppium.Quit()
                    objAppium = Nothing
                    Test.Wait(1000)
                Loop

            Catch ex As Exception
                ' Ignore errors if unable to close the browser
            End Try
            ' Assert.AreEqual("", verificationErrors.ToString())
        End Sub
        Sub Reflesh()
            Try
                objSeleniumWD.Navigate.Refresh()
                Dim objApp As New OpenQA.Selenium.Appium.MultiTouch.MultiAction
                Dim touch As New TouchAction(objAppium)


            Catch ex As Exception

            End Try
        End Sub
    End Class
End Namespace