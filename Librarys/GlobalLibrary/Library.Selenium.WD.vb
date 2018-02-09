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
Imports System.Windows.Forms
Imports OpenQA.Selenium.Opera
Imports OpenQA.Selenium.Edge

Namespace LibrarySeleniumWD
    Public Class ScreenShotRemoteWebDriver
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
        Public Class SeleniumWDHelper
        'Public driver As ScreenShotRemoteWebDriver
        Private capability As New DesiredCapabilities

     
        Public Sub SetupTest()
            Dim pathServer As String = "C:\LeanTestAutomation\API\" 'System.IO.Directory.GetCurrentDirectory
            'Dim chromeOptions As New OpenQA.Selenium.Chrome.ChromeOptions
            'Dim proxy As New OpenQA.Selenium.Proxy
            Try
                ' objSeleniumWD.Manage.Timeouts().SetPageLoadTimeout(System.TimeSpan.FromSeconds(p_timeout))
                SelectBrowser(p_browserType)
                Select Case p_appProcessName

                    Case "chrome"
                        Dim chromeOptions = New ChromeOptions()
                        'pc_util.StopProcess("chrome")
                        pc_util.StopProcess("chromedriver")

                        'chromeOptions.addArguments("test-type")
                        chromeOptions.AddArguments("start-maximized")
                        'chromeOptions.addArguments("--js-flags=--expose-gc")
                        'chromeOptions.addArguments("--enable-precise-memory-info")
                        chromeOptions.AddArguments("--disable-popup-blocking")
                        'chromeOptions.addArguments("--disable-default-apps")
                        chromeOptions.AddArguments("disable-infobars")
                        objSeleniumWD = New ChromeDriver(pathServer, chromeOptions)
                    Case "iexplore"
                        objSeleniumWD = New InternetExplorerDriver(pathServer)
                    Case "firefox"
                        Dim driverService = FirefoxDriverService.CreateDefaultService()
                        driverService.FirefoxBinaryPath = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
                        driverService.HideCommandPromptWindow = True
                        driverService.SuppressInitialDiagnosticInformation = True
                        objSeleniumWD = New FirefoxDriver(driverService, New FirefoxOptions(), TimeSpan.FromSeconds(60))
                    Case "opera"
                        objSeleniumWD = New OperaDriver(pathServer)
                    Case "edge"
                        objSeleniumWD = New EdgeDriver(pathServer)
                    Case Else
                        Test.Wait(1000)
                End Select
            Catch ex As Exception
                p_errorDescription = "Menssage error: " & ex.Message.ToString
            End Try
        End Sub
        Public Sub TeardownTest()
            Try
                objSeleniumWD.Quit()
            Catch ex As Exception
                ' Ignore errors if unable to close the browser
            End Try
            ' Assert.AreEqual("", verificationErrors.ToString())
        End Sub
    End Class

    'class contains all methods to interactin with windows and elements
    Public Class InteractionWD
    	
        Private screenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
        Private screenHeight As Integer = Screen.PrimaryScreen.Bounds.Height - 30

        'Private driver As New WDBrowserStackHelper
        Function Open(url As String) As Boolean
            Try
                objSeleniumWD.Navigate.GoToUrl(url)
                Return True
            Catch ex As Exception
                HandlerError("LibraryobjSeleniumWD.InteractionWDBrowseerStack.Open: " & url & " Error:" & ex.Message & " - " & ex.InnerException.ToString)
                Return False
            End Try
        End Function
        Function GetURL() As String
            Try
                Return objSeleniumWD.Url
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
        Function WaitForPageToLoad(Optional timeOut As Integer = 0) As Boolean
            timeOut = IIf(timeOut <> 0, timeOut, p_timeout)
            Try
                Console.WriteLine("WaitForPageToLoad: Timeout = " & p_timeout)
                objSeleniumWD.Manage.Timeouts().SetPageLoadTimeout(System.TimeSpan.FromSeconds(p_timeout))

            Catch ex As Exception
                HandlerError("LibraryobjSeleniumWD.InteractionWDBrowseerStack.WaitForPageToLoad: Error: " & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function

        'VERIFICAR*******************************************************
        Function OpenWindow(url As String, WindowID As String) As Boolean
            Try
                'objSeleniumWD.OpenWindow(url, WindowID)
                Return True
            Catch ex As Exception
                HandlerError("Library.Selenium.RC.Selenium.Interaction.OpenWindow: " & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
        'VERIFICAR*******************************************************
        Function SelectWindow(windowHandles As windowHandles) As Boolean
        	Try
        		Select Case windowHandles
        				
        			Case LibraryGlobal.LibGlobal.windowHandles.FirstOrDefault
                        objSeleniumWD.SwitchTo().Window(objSeleniumWD.WindowHandles().LastOrDefault())
                    Case LibraryGlobal.LibGlobal.windowHandles.First
                        objSeleniumWD.SwitchTo().Window(objSeleniumWD.WindowHandles().First())
                    Case LibraryGlobal.LibGlobal.windowHandles.Last
                        objSeleniumWD.SwitchTo().Window(objSeleniumWD.WindowHandles().Last)
                End Select

                Return True
            Catch ex As Exception
                HandlerError("Library.Selenium.RC.Selenium.Interaction.SelectWindow: " & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
        Function SelectFrame(frameName As String) As Boolean
        	Try
                objSeleniumWD.SwitchTo().Frame(frameName)
                Return True
            Catch ex As Exception
                HandlerError("Library.Selenium.WD.Selenium.Interaction.SelectFrame: " & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
        Function Exist(element As String, waitMilliseconds As Integer, Optional typeIdentification As typeIdentification = typeIdentification.xpath) As Boolean

            Try
                Thread.Sleep(waitMilliseconds)
                Select Case typeIdentification
                    Case typeIdentification.xpath
                        If objSeleniumWD.FindElement(By.XPath(element)).Displayed Then
                            Return True
                        Else
                            Return False
                        End If
                    Case typeIdentification.name
                        If objSeleniumWD.FindElement(By.Name(element)).Displayed Then
                            Return True
                        Else
                            Return False
                        End If
                    Case typeIdentification.id
                        If objSeleniumWD.FindElement(By.Id(element)).Displayed Then
                            Return True
                        Else
                            Return False
                        End If
                    Case typeIdentification.linkText
                        If objSeleniumWD.FindElement(By.LinkText(element)).Displayed Then
                            Return True
                        Else
                            Return False
                        End If
                    Case typeIdentification.css
                        If objSeleniumWD.FindElement(By.CssSelector(element)).Displayed Then
                            Return True
                        Else
                            Return False
                        End If
                End Select
            Catch ex As Exception
                HandlerError("Library.Selenium.RC.Interaction.exist:" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
        Sub HigthLigth(element As String)
            Dim jsDriver = DirectCast(objSeleniumWD, IJavaScriptExecutor)
            Dim highlightJavascript As String = "$(arguments[0]).css({ ""border-width"" : ""2px"", ""border-style"" : ""solid"", ""border-color"" : ""red"" });"
            jsDriver.ExecuteScript(highlightJavascript, New Object() {"xpath=" & element})
        End Sub
        Function WaitExist(element As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath) As Boolean
            Try
                For i = 0 To p_timeout - 1
                    Try
                        Select Case typeIdentification
                            Case typeIdentification.name
                                If objSeleniumWD.FindElement(By.Name(element)).Displayed Then
                                    Return True
                                End If
                            Case typeIdentification.id
                                If objSeleniumWD.FindElement(By.Id(element)).Displayed Then
                                    Return True
                                End If
                            Case typeIdentification.xpath
                                If objSeleniumWD.FindElement(By.XPath(element)).Displayed Then
                                    Return True
                                End If
                            Case typeIdentification.linkText
                                If objSeleniumWD.FindElement(By.LinkText(element)).Displayed Then
                                    Return True
                                End If
                            Case typeIdentification.css
                                If objSeleniumWD.FindElement(By.CssSelector(element)).Enabled Then
                                    Return True
                                End If
                        End Select

                        Test.Wait(1000)
                        Console.Write("Element: " & element & " - Time: " & i & vbCrLf)
                    Catch ex As Exception
                        Console.Write("Element: " & element & " - Time: " & i & vbCrLf)
                        Test.Wait(1000) ' clique em pause, F11 e em seguida arraste o cursor para o End function
                    End Try
                Next
                'p_timeout = timeOut
                Throw New Exception("waitLoadElement: Element=" & element & " not found")
                Return False
            Catch ex As Exception
                p_errorDescription = "Menssage error: " & ex.Message.ToString
                Throw New Exception(ex.Message)
                Return False
            End Try
        End Function
        'VERIFICAR*******************************************************
        Function GetHtmlSource()
            Try
                Return objSeleniumWD.PageSource.ToString()
            Catch ex As Exception
                p_errorDescription = "Menssage error: " & ex.Message.ToString
                Return Nothing
            End Try
        End Function
        Sub Clear(element As String)
            Try
                Dim webElement As IWebElement = objSeleniumWD.FindElement(By.XPath(element))
                webElement.Clear()
            Catch ex As Exception
                p_errorDescription = "Menssage error: " & ex.Message.ToString
                Throw New Exception("Clear element error: " & p_errorDescription)
            End Try
        End Sub

        Sub WindowMaximize()
            Try
                objSeleniumWD.Manage.Window.Maximize()
            Catch ex As Exception
                p_errorDescription = "Menssage error: " & ex.Message.ToString
                Throw New Exception("Windows maximize error: " & p_errorDescription)
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
                objSeleniumWD.Navigate.Refresh()
            Catch ex As Exception

            End Try
        End Sub
        Function MouseOver(element As String, value As Boolean, Optional typeIdentification As typeIdentification = typeIdentification.xpath, Optional waitElement As Boolean = True) As Boolean
            If Not value Then Exit Function
            Dim webElement As IWebElement = Nothing
            Try
                If String.IsNullOrEmpty(element) Then Return False
                If waitElement Then
                    If WaitExist(element, typeIdentification) Then

                        Try
                            Select Case typeIdentification
                                Case typeIdentification.name
                                    webElement = objSeleniumWD.FindElement(By.Name(element))
                                Case typeIdentification.id
                                    webElement = objSeleniumWD.FindElement(By.Id(element))
                                Case typeIdentification.xpath
                                    webElement = objSeleniumWD.FindElement(By.XPath(element))
                                Case typeIdentification.linkText
                                    webElement = objSeleniumWD.FindElement(By.LinkText(element))
                            End Select

                            'webElement.Location
                            MouseOverElement(webElement)
                            Return True
                        Catch ex As Exception
                            Return False
                        End Try
                    End If
                End If
                Return False
            Catch ex As Exception
                p_errorDescription = "Menssage error: " & ex.Message.ToString
                HandlerError("LibrarySeleniumRC.Interaction.click")
                Throw New Exception(ex.Message)
                Return False
            End Try
        End Function

        Function Click(element As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath, Optional waitElement As Boolean = True) As Boolean
            Dim webElement As IWebElement = Nothing
            Try
                If String.IsNullOrEmpty(element) Then Return False
                If waitElement Then
                    If WaitExist(element, typeIdentification) Then

                        Try
                            Select Case typeIdentification
                                Case typeIdentification.name
                                    webElement = objSeleniumWD.FindElement(By.Name(element))
                                Case typeIdentification.id
                                    webElement = objSeleniumWD.FindElement(By.Id(element))
                                Case typeIdentification.xpath
                                    webElement = objSeleniumWD.FindElement(By.XPath(element))
                                Case typeIdentification.linkText
                                    webElement = objSeleniumWD.FindElement(By.LinkText(element))
                                Case typeIdentification.css
                                    webElement = objSeleniumWD.FindElement(By.CssSelector(element))
                            End Select
                            MouseOverElement(webElement)
                            webElement.Click()
                            Return True
                        Catch ex As Exception
                            p_errorDescription = "Menssage error: " & ex.Message.ToString
                            'Throw New Exception(ex.Message)
                            Return False
                        End Try

                    End If
                End If
                Return False
            Catch ex As Exception
                p_errorDescription = "Menssage error: " & ex.Message.ToString
                HandlerError("LibrarySeleniumRC.Interaction.click")
                Throw New Exception(ex.Message)
                Return False
            End Try
        End Function
        Function DoubleClick(element As String, typeIdentification As typeIdentification) As Boolean

            Try
                If String.IsNullOrEmpty(element) Then Return False
                If WaitExist(element) Then
                    Try
                        Dim act As Actions = Nothing
                        Select Case typeIdentification
                            Case LibraryGlobal.LibGlobal.typeIdentification.id
                                act.DoubleClick(objSeleniumWD.FindElement(By.Id(element))).Build().Perform()
                            Case LibraryGlobal.LibGlobal.typeIdentification.name
                                act.DoubleClick(objSeleniumWD.FindElement(By.Name(element))).Build().Perform()
                            Case LibraryGlobal.LibGlobal.typeIdentification.xpath
                                act.DoubleClick(objSeleniumWD.FindElement(By.XPath(element))).Build().Perform()
                            Case LibraryGlobal.LibGlobal.typeIdentification.linkText
                                act.DoubleClick(objSeleniumWD.FindElement(By.LinkText(element))).Build().Perform()
                            Case LibraryGlobal.LibGlobal.typeIdentification.css
                                act.DoubleClick(objSeleniumWD.FindElement(By.CssSelector(element))).Build().Perform()
                        End Select

                        Return True
                    Catch ex As Exception
                        p_errorDescription = "Menssage error: " & ex.Message.ToString
                        HandlerError("LibrarySeleniumRC.Interaction.click")
                        Return False
                    End Try
                End If
                Return False
            Catch ex As Exception
                p_errorDescription = "Menssage error: " & ex.Message.ToString
            End Try
            HandlerError("LibrarySeleniumRC.Interaction.click")
            Return False
        End Function

        Public Function Type(element As String, value As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath) As Boolean
            Dim webElement As IWebElement = Nothing
            If String.IsNullOrEmpty(value) Then Return True
            Try
                If WaitExist(element, typeIdentification) Then
                    Try
                        Select Case typeIdentification
                            Case typeIdentification.xpath
                                webElement = objSeleniumWD.FindElement(By.XPath(element))
                            Case typeIdentification.name
                                webElement = objSeleniumWD.FindElement(By.Name(element))
                            Case typeIdentification.id
                                webElement = objSeleniumWD.FindElement(By.Id(element))
                            Case typeIdentification.css
                                webElement = objSeleniumWD.FindElement(By.CssSelector(element))
                        End Select
                        Try
                            MouseOverElement(webElement)
                        Catch ex As Exception
                            p_errorDescription = "Menssage error: " & ex.Message.ToString
                            Throw New Exception(ex.Message)
                            Return False
                        End Try
                        Try
                            If Not String.IsNullOrEmpty(value) Then webElement.Clear()
                            If value = "@Nothing" Then
                                webElement.Click()
                                webElement.Clear()
                                Return True
                            End If
                        Catch ex As Exception
                        End Try
                        If value <> "@Nothing" Then
                            webElement.Click()
                            webElement.SendKeys(value)
                        End If
                        Return True
                    Catch ex As Exception
                        webElement.SendKeys(value)
                    End Try
                Else
                    Return False
                End If
            Catch ex As Exception
                p_errorDescription = "Menssage error: " & ex.Message.ToString
                HandlerError("LibraryInteractionWD.Interaction.Type")
                Throw New Exception(ex.Message)
                Return False
            End Try
            Return Nothing
        End Function

        Function SelectValue(ImgPathOrCoordinatesXY As String, value As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath) As Boolean
            If String.IsNullOrEmpty(ImgPathOrCoordinatesXY) Then Return True
            If String.IsNullOrEmpty(value) Then Return True
            Dim captureValue As String = Nothing
            If WaitExist(ImgPathOrCoordinatesXY, typeIdentification) Then
                Dim webElement As IWebElement = Nothing
                Try
                    If typeIdentification = LibraryGlobal.LibGlobal.typeIdentification.leanTest Then
                        Return pc_dv.SelectValue(ImgPathOrCoordinatesXY, value)
                    End If
                    Select Case typeIdentification
                        Case typeIdentification.xpath
                            webElement = objSeleniumWD.FindElement(By.XPath(ImgPathOrCoordinatesXY))
                        Case typeIdentification.name
                            webElement = objSeleniumWD.FindElement(By.Name(ImgPathOrCoordinatesXY))
                        Case typeIdentification.id
                            webElement = objSeleniumWD.FindElement(By.Id(ImgPathOrCoordinatesXY))
                        Case typeIdentification.linkText
                            webElement = objSeleniumWD.FindElement(By.LinkText(ImgPathOrCoordinatesXY))
                    End Select
                    Try
                        MouseOverElement(webElement)
                    Catch ex As Exception
                        p_errorDescription = "Menssage error: " & ex.Message.ToString
                        Throw New Exception(ex.Message)
                        Return False
                    End Try
                    If value <> "@Nothing" Then
                        webElement.SendKeys(value)
                    End If

                Catch ex As Exception
                    Try
                        MouseOverElement(webElement)
                        webElement.Click()
                        Test.SendKey(value)
                        Test.SendKey("{ENTER}")
                        Return True
                    Catch ex1 As Exception
                    End Try
                End Try
            End If
            Return False
        End Function
        Function GetText(element As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath) As String
            If String.IsNullOrEmpty(element) Then Return Nothing
            If Not WaitExist(element, typeIdentification) Then Return Nothing
            Try
                Dim webElement As IWebElement = Nothing
                Select Case typeIdentification
                    Case typeIdentification.name
                        webElement = objSeleniumWD.FindElement(By.Name(element))
                    Case typeIdentification.id
                        webElement = objSeleniumWD.FindElement(By.Id(element))
                    Case typeIdentification.xpath
                        webElement = objSeleniumWD.FindElement(By.XPath(element))
                End Select
                'webElement.Location
                MouseOverElement(webElement)
                Return webElement.Text
            Catch ex As Exception
                p_errorDescription = "Menssage error: " & ex.Message.ToString
                HandlerError("LibraryInteractionWD.Interaction.selectValue")
                Return Nothing
            End Try
        End Function
        Function GetValue(element As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath) As String
            If String.IsNullOrEmpty(element) Then Return Nothing
            'If Not WaitExist(element, typeIdentification) Then Return Nothing
            Try
                Dim webElement As IWebElement = Nothing
                Select Case typeIdentification
                    Case typeIdentification.name
                        webElement = objSeleniumWD.FindElement(By.Name(element))
                    Case typeIdentification.id
                        webElement = objSeleniumWD.FindElement(By.Id(element))
                    Case typeIdentification.xpath
                        webElement = objSeleniumWD.FindElement(By.XPath(element))
                    Case LibraryGlobal.LibGlobal.typeIdentification.linkText
                        webElement = objSeleniumWD.FindElement(By.LinkText(element))
                End Select
                'webElement.Location
                MouseOverElement(webElement)
                Return webElement.GetAttribute("value")
            Catch ex As Exception
                p_errorDescription = "Menssage error: " & ex.Message.ToString
                HandlerError("LibraryInteractionRC.Interaction.selectValue")
                Return Nothing
            End Try
        End Function
       

        Function GetTextPopup(Optional click As Boolean = True, Optional frameName As String = Nothing) As String
            Dim msg As String = Nothing
            Try
                Dim alert

                If String.IsNullOrEmpty(frameName) Then
                    alert = objSeleniumWD.SwitchTo().Alert()
                Else
                    alert = objSeleniumWD.SwitchTo.Frame(frameName).SwitchTo.Alert
                End If
                msg = alert.Text
                If click Then alert.Accept()
                Return msg
            Catch ex As Exception
                p_errorDescription = "Menssage error: " & ex.Message.ToString
                HandlerError("LibraryInteractionRC.Interaction.selectValue")
                Return Nothing
            End Try
        End Function
        Function GetSelectedValue(element As String) As String
            If String.IsNullOrEmpty(element) Then Return Nothing
            If WaitExist(element) Then
                Try
                    Dim webElement As IWebElement = objSeleniumWD.FindElement(By.XPath(element))
                    MouseOverElement(webElement)
                    Return webElement.GetAttribute("Value")
                Catch ex As Exception
                    p_errorDescription = "Menssage error: " & ex.Message.ToString
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
                    If objSeleniumWD.FindElement(By.Id(element)).Enabled Then
                        Console.WriteLine("isEditable: " & element)
                        Return True
                    Else
                        Return False
                    End If
                Catch ex As Exception
                    p_errorDescription = "Menssage error: " & ex.Message.ToString
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
                p_errorDescription = "Menssage error: " & ex.Message.ToString
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
                p_errorDescription = "Menssage error: " & ex.Message.ToString
                HandlerError("LibraryInteractionRC.Interaction.GetCellDataByReference")
                Return Nothing
            End Try
        End Function
        Public Sub TeardownTest()
            Try
                Do While Not IsNothing(objSeleniumWD)
                    objSeleniumWD.Quit()
                    objSeleniumWD = Nothing
                    Test.Wait(1000)
                Loop
            Catch ex As Exception
                ' Ignore errors if unable to close the browser
            End Try
            ' Assert.AreEqual("", verificationErrors.ToString())
        End Sub
        Private Sub MouseOverElement(webElement As IWebElement)
            Try
                'webElement.Location
                Dim xy = webElement.Location
                If xy.X > screenWidth Then xy.X = xy.X
                If xy.Y > screenHeight Then xy.Y = screenHeight - 40 Else xy.Y = xy.Y + 74
                pc_dv.MouseMove(xy.X + 20, xy.Y)
            Catch ex As Exception
                p_errorDescription = "Menssage error: " & ex.Message.ToString
            End Try
        End Sub
        Sub Reflesh()
            Try
                objSeleniumWD.Navigate.Refresh()
            Catch ex As Exception
                p_errorDescription = "Menssage error: " & ex.Message.ToString
            End Try
        End Sub
        Function GetAttribute(element As String, attribute As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath) As String
            If String.IsNullOrEmpty(element) Then Return Nothing
            If String.IsNullOrEmpty(attribute) Then Return Nothing

            Dim captureValue As String = Nothing
            If WaitExist(element, typeIdentification) Then
                Dim webElement As IWebElement = Nothing
                Try
                    Select Case typeIdentification
                        Case typeIdentification.xpath
                            webElement = objSeleniumWD.FindElement(By.XPath(element))
                        Case typeIdentification.name
                            webElement = objSeleniumWD.FindElement(By.Name(element))
                        Case typeIdentification.id
                            webElement = objSeleniumWD.FindElement(By.Id(element))
                        Case typeIdentification.linkText
                            webElement = objSeleniumWD.FindElement(By.LinkText(element))
                    End Select
                    Try
                        MouseOverElement(webElement)
                        Return webElement.GetAttribute(attribute)
                    Catch ex As Exception
                        p_errorDescription = "Menssage error: " & ex.Message.ToString
                        Throw New Exception(ex.Message)
                        Return Nothing
                    End Try
                Catch ex As Exception
                    Return Nothing
                End Try
            End If
            Return Nothing
        End Function
    End Class
End Namespace