Imports System
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports NUnit.Framework
Imports Selenium
Imports SeleniumToolkit
Imports OpenQA
Imports OpenQA.Selenium.IE
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Imports Lean.Test.Automation.API
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Remote
Imports System.Reflection
Imports System.Windows.Forms
Imports System.Drawing.Imaging

Imports System.IO
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports Shell32

Module variables
    Public Enum ActiveTool
        SeleniumWD
        SeleniumWDBrowserStack
        SilkTest
        Appium
        LeanTest
        SeleniumRC
    End Enum
    '******************************************************************************************
    'variables
    Public p_ClassName As String 'coded ui
    Public p_WinName As String 'coded ui
    Public p_WinTitle As String 'coded ui
    Public p_ObjName As String 'coded ui
    Public p_CurrentDiretory As String

    Public p_IDScenario As Integer
    Public p_ScenarioName As String
    Public p_ScenarioDescription As String
    Public p_IDTest As String
    Public p_TestName As String
    Public p_DescriptionTest As String
    Public p_pathTest As String
    Public p_pathLab As String
    Public p_CountTest As Integer
    Public p_CountScenario As Integer
    Public p_RunIDScenario As Integer
    Public p_TableTest As String
    Public p_OrdemTest As String
    Public p_IDRun As String
    Public p_IsLoop As Boolean
    Public p_ClearResult As Boolean
    Public p_ResetApplication As Boolean
    Public p_errorDescription As String

    'variables of test
    Public libGlobal As New LibraryGlobal.LibGlobal
    Public Test As New LibraryGlobal.LibGlobal
    Public p_firstRun As Boolean
    Public p_IDScenarioTest As String
    Public p_StatusTest As String
    Public p_StatusScenario As String
    Public p_Release As String
    Public p_Cycle As String
    Public p_FullPathPlan As String
    Public p_FullPathExec As String
    Public p_Run As Object
    Public p_IDTestInstance As String
    Public p_Step As String
    Public p_ExpectedResuts As String
    Public p_ActualResuts As String

    '******************************************************************************************
    'load values
    Public p_executionTime As String
    Public p_pathSeleniuRC As String
    Public p_timeout As Integer
    Public p_browserType As String
    Public p_ToolName As ActiveTool
    Public p_pathUrlApp As String
    Public p_Device As String
    Public p_PlatformName As String
    Public p_ServerMobile As String
    Public p_pathEvidence As String
    Public p_speedyExecution As Integer
    Public p_pathFileTemplateLogEvidenceDoc As String
    Public p_pathFileStatusScenario As String
    Public p_pathStatusReportAllScenarios As String
    Public p_pathSaveLog As String
    Public p_pathAppChrome As String
    Public p_pathAppFirefox As String
    Public p_pathAppIexplorer As String
    Public p_pathAppOpera As String
    Public p_pathAppEdge As String
    Public p_appProcessName As String
    Public p_PublishQC As Boolean
    Public p_GenerateLogTest As Boolean
    Public p_Highlight As Boolean
    Public p_PathXML As String
    Public p_PackageName As String
    '******************************************************************************************
    'Selenium framework
    Public p_ServerRC As String
    Public p_portRC As Integer
    Public p_ServerWD As String
    Public objSeleniumRC 'As ISelenium
    Public objSeleniumWD As IWebDriver
    'Public objSeleniumWDBS As ScreenShotRemoteWebDriver
    Public objAppium As RemoteWebDriver
    'Public p_Platform As Platform
    '*******************************************************************************************
    'Componets framework
    Public pc_util As New Utilities
    'Public pc_log As New LogsExecution
    Public pc_db As New DBAdapter
    Public pc_te As New TestEvidence
    Public pc_rp As New ExportWord
    Public xml As New XMLFile
    Public p_email As New Email
    Public p_qc 'As New QCUtilities
    Public pc_dv As New Devices

End Module
Namespace LibraryGlobal
    Public Class CallTest
        Private LibGlobal As LibGlobal
        Public Shared Function RunTest(ScriptName As String) As Boolean
            LibGlobal.LogExecution("RunTest: " & ScriptName, LibGlobal.StateLog.Start)
            Try
                Dim type1 As Type = Type.GetType("Lean.Test.Automation.Framework." & ScriptName & "." & ScriptName)
                Dim constructor As ConstructorInfo = type1.GetConstructor(Type.EmptyTypes)
                Dim ClassObject As Object = constructor.Invoke(New Object() {})
                Dim Method As MethodInfo = type1.GetMethod("Run")
                Dim Value As Object = Method.Invoke(ClassObject, Nothing)
                LogExecution("RunTest: " & ScriptName, LibGlobal.StateLog.Finish)
                Return Value
            Catch ex As Exception
                LogExecution("RunTest error: " & ex.Message, LibGlobal.StateLog.Exception)
                MsgBox("RunTest " & ScriptName & "' error: " & ex.Message.ToString)
                Return False
            End Try
        End Function
    End Class
    Public Class LibGlobal
        Public Enum StateLog
            Start
            Finish
            Exception
            Information
        End Enum

        Public Shared Sub LogExecution(MethodName As String, State As StateLog, Optional isTestlog As Boolean = False)
            Try
                Dim statelocal As String = Nothing
                Select Case State
                    Case StateLog.Start
                        statelocal = "Start"
                    Case StateLog.Finish
                        statelocal = "Finish"
                    Case StateLog.Exception
                        statelocal = "Exception"
                        If isTestlog Then Test.TestLog("Executar passo de teste", Replace(MethodName, "'", ""), "Passo executado com falha", typelog.Failed)
                        HandlerError(MethodName)
                    Case StateLog.Information
                        statelocal = "Information"
                End Select
                Console.WriteLine("*******************************************************************************")
                Console.WriteLine(statelocal & " - " & MethodName & " - " & Now)
                Console.WriteLine(vbCrLf)
            Catch ex As Exception

            End Try
        End Sub
        Sub ExecuteCommandSQLAccess(sql As String)
            pc_db.ExecuteCommandSQL(sql, True, p_PathXML)
        End Sub
        Public Enum TypeBrowser
            InternetExplorer
            Firefox
            Chrome
            Opera
            Edge
        End Enum
        Public Enum typelog
            Passed
            Failed
            Warning
            Done
            Blocked
            Information
            NA
        End Enum
        Public Enum windowHandles
            FirstOrDefault
            First
            Last
        End Enum
        Public Enum typeIdentification
            id
            linkText
            name
            xpath
            css
            leanTest
        End Enum
        Private librarySeleniumWDInteraction As New LibrarySeleniumWD.InteractionWD
        Private libraryAppiumInteration As New LibraryAppium.InteractionAppium

        Public Shared Function SelectBrowser(browser As TypeBrowser) As String
            Dim appBrowser As String
            LogExecution("SelectBrowser", StateLog.Start)
            Try
                Select Case browser
                    Case TypeBrowser.Chrome
                        p_appProcessName = "chrome"
                        appBrowser = p_pathAppChrome
                    Case TypeBrowser.Firefox
                        p_appProcessName = "firefox"
                        appBrowser = p_pathAppFirefox
                    Case TypeBrowser.InternetExplorer
                        p_appProcessName = "iexplore"
                        appBrowser = p_pathAppIexplorer
                    Case TypeBrowser.Opera
                        p_appProcessName = "opera"
                        appBrowser = p_pathAppOpera
                    Case TypeBrowser.Edge
                        p_appProcessName = "edge"
                        appBrowser = p_pathAppEdge
                    Case Else
                        LogExecution("SelectBrowser: Nothing", StateLog.Exception)
                        appBrowser = Nothing
                        Return Nothing
                End Select
                If p_ResetApplication Then
                    'pc_util.StopProcess(p_appProcessName) 'refazer *************************************************
                    LogExecution("SelectBrowser: ResetApplication=" & p_ResetApplication, StateLog.Information)
                End If
                LogExecution("SelectBrowser: TypeBrowser=" & p_appProcessName, StateLog.Finish)
                Return appBrowser
            Catch ex As Exception
                HandlerError("libGlobal.SelectBrowser")
                Return Nothing
            End Try
        End Function
        Function StartTests() As Boolean
            LogExecution("StartTests", StateLog.Start)
            Try
                Select Case p_ToolName
                    Case ActiveTool.SeleniumWD 'selenium web driver
                        pc_util.StopProcess("chrome")
                        pc_util.StopProcess("chromedriver")
                        Dim wd = New LibrarySeleniumWD.SeleniumWDHelper
                        wd.SetupTest()
                        objSeleniumWD.Manage.Window.Maximize()
                        LogExecution("StartTests: SeleniumWD", StateLog.Finish)
                        Return True
                    Case ActiveTool.Appium
                        Dim app = New LibraryAppium.SeleniumAppiumHelper
                        app.SetupTest()
                        Return True
                    Case Else
                        LogExecution("SartTests: Case Else", StateLog.Exception)
                        Return True
                End Select
            Catch ex As Exception
                LogExecution("SartTests error: " & ex.Message, StateLog.Exception)
                Return False
            End Try
        End Function

        Public Shared Function EndTestTable() As Boolean
            LogExecution("EndTestTable", StateLog.Start)
            Try
                pc_db.EndExecutionTable(CInt(p_IDScenarioTest), p_StatusTest, p_IDTest)
                LogExecution("EndTestTable", StateLog.Finish)
                Return True
            Catch ex As Exception
                LogExecution("EndTestTable error: " & ex.Message, StateLog.Exception)
                Return False
            End Try
        End Function
        '*********************************************************************************************************************************
        'ENDTEST
        Public Function EndTest(GenerateLog As Boolean) As Boolean
            Dim savePath As String
            Dim lenID As Integer = Len(CStr(p_IDScenario))
            Dim pathTemp As String = "\SCEN"
            pathTemp = pathTemp & Strings.StrDup(5 - lenID, "0") & p_IDScenario

            LogExecution("EndTest: IDScenario: " & p_IDScenario & " - Test name: " & p_TestName & " - StatusTest: " & p_StatusTest, StateLog.Start)

            Try
                pc_db.EndExecution(p_TableTest, p_IDTest)
                LogExecution("EndTest.EndExecution: " & p_TableTest & " - IDTest" & p_IDTest, StateLog.Information)

                savePath = p_pathEvidence & p_pathLab & "\" & pathTemp
                savePath = Replace(savePath, "\\", "\")
                If GenerateLog Then
                    LogExecution("EndTest.GenerateLog in path: " & savePath, StateLog.Information)
                    pc_rp.ExportLogsTestToWord(p_pathFileTemplateLogEvidenceDoc, savePath, p_TableTest, p_IDTest, xml.Read("typeDB"))
                End If
                LogExecution("EndTest", StateLog.Finish)
                Return True
            Catch ex As Exception
                LogExecution("EndTest error: " & ex.Message, StateLog.Exception)
                Return False
            End Try
        End Function
        'End Scenario
        Shared Sub EndScenario(Optional idSceanario As Integer = 0, Optional scenarioName As String = Nothing)
            Dim savePath As String
            Dim lenID As Integer = Len(CStr(p_IDScenario))
            Dim pathTemp As String = "\SCEN"
            Dim idScen As Integer
            Dim scenarioNam As String
            Try
                LogExecution("EndScenario: " & p_IDScenario, StateLog.Start)

                idScen = IIf(idSceanario = 0, p_IDScenario, idSceanario)
                scenarioNam = IIf(scenarioName = Nothing, p_ScenarioName, scenarioName)

                If idScen <> 0 Then pc_db.EndExecutionScenario(idScen, p_StatusScenario)
                p_StatusScenario = Nothing

                pathTemp = pathTemp & Strings.StrDup(5 - lenID, "0") & idScen
                savePath = p_pathEvidence & p_pathLab & "\" & pathTemp
                savePath = Replace(savePath, "\\", "\")

                If p_GenerateLogTest Then
                    LogExecution("EndScenario.GenerateLogTest", StateLog.Information)
                    pc_rp.ExportLogsSceanarioToWord(p_pathFileStatusScenario, savePath, idScen, "0")
                End If
                LogExecution("EndScenario", StateLog.Finish)
            Catch ex As Exception
                LogExecution("EndScenario error: " & ex.Message, StateLog.Exception)
            End Try
        End Sub
        '*********************************************************************************************************************************
        'RunTestScenario - this function run all tests by scenario
        Shared Function RunTestScenario(Optional IDScenario As String = "0") As Boolean
            LogExecution("RunTestScenario: IDScenario=" & IDScenario, StateLog.Start)
            Try
                p_CountScenario = pc_db.RunTestScenario(CStr(IDScenario), 1)
                LogExecution("RunTestScenario: p_CountScenario=" & p_CountScenario, StateLog.Information)
                Return True
            Catch ex As Exception
                LogExecution("RunTestScenario: error=" & ex.Message, StateLog.Exception)
                Return False
            End Try
        End Function
        Sub LoadVariables()
            p_pathSeleniuRC = xml.Read("pathSeleniumRC")
            p_pathEvidence = xml.Read("pathEvidence")
            p_pathFileTemplateLogEvidenceDoc = xml.Read("pathTemplateLogTest") 'log evidence
            p_pathFileStatusScenario = xml.Read("pathStatusReportScenario")
            p_pathStatusReportAllScenarios = xml.Read("pathStatusReportAllScenarios")
            p_pathSaveLog = xml.Read("pathSaveLog")
            p_pathAppChrome = xml.Read("pathAppChrome")
            p_pathAppFirefox = xml.Read("pathAppFirefox")
            p_pathAppIexplorer = xml.Read("pathAppIexplorer")
            p_pathAppOpera = xml.Read("pathAppOpera")
            p_pathAppEdge = xml.Read("pathAppEdge")
            p_ServerRC = xml.Read("pathSeleniumRC")
            p_portRC = xml.Read("portRC")
            p_ServerWD = xml.Read("pathSeleniumWD")

            p_timeout = CInt(pc_db.GetParameter("TimeOut")) '10 'seconds
            p_ToolName = pc_db.GetParameter("ToolName") 'ActiveTool.SeleniumRC
            p_speedyExecution = pc_db.GetParameter("ExecutionSpeedy")

            p_ResetApplication = pc_db.GetParameter("ResetApp") 'True 'true fecha todos os processos abertos
            If Not String.IsNullOrEmpty((pc_db.GetParameter("ClearResult"))) Then p_ClearResult = (pc_db.GetParameter("ClearResult")) 'True
            p_GenerateLogTest = pc_db.GetParameter("GerateLog") 'False
            p_Highlight = pc_db.GetParameter("HighLight") 'True
            p_browserType = pc_db.GetParameter("Browser") '2 'TypeBrowser.Chrome
            p_PackageName = pc_db.GetParameter("Demand")
            p_Device = pc_db.GetParameter("Device")
            p_PlatformName = pc_db.GetParameter("PlatformName")
            p_executionTime = "\Execution_" & Format(Now, "yyyy_MM_dd")
            'p_PackageName = pc_db.GetValue("Select Top 1 * from LeanTestDemands", "Demand")
        End Sub
        '******************************************************************************************************
        'METHODS
        Function GetURL() As String
            Test.Wait(p_speedyExecution)
            Try
                LogExecution("Get url", StateLog.Start)
                Select Case p_ToolName
                    Case ActiveTool.SeleniumWD
                        Return librarySeleniumWDInteraction.GetURL()
                    Case ActiveTool.Appium
                    	'libraryAppiumInteration.Open(url)
                    	Return Nothing
                    Case Else
                        Return Nothing
                End Select
                LogExecution("Open: url", StateLog.Finish)
            Catch ex As Exception
                LogExecution("Open url", StateLog.Exception)
                Return Nothing
            End Try
        End Function

        Function Open(url As String) As Boolean
            Test.Wait(p_speedyExecution)
            Try
                LogExecution("Open url: " & url, StateLog.Start)
                Select Case p_ToolName
                    Case ActiveTool.SeleniumWD
                        librarySeleniumWDInteraction.Open(url)
                    Case ActiveTool.Appium
                        libraryAppiumInteration.Open(url)
                    Case Else

                End Select
                LogExecution("Open: url=" & url, StateLog.Finish)
                Return True
            Catch ex As Exception
                LogExecution("Open url: " & url, StateLog.Exception)
                Return False
            End Try
        End Function


        Function GetAttribute(element As String, attribute As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath) As String
            Dim value As String = Nothing
            Test.Wait(p_speedyExecution)
            Try
                LogExecution("GetAttribute: " & attribute, StateLog.Start)
                Select Case p_ToolName
                    Case ActiveTool.SeleniumWD
                        value = librarySeleniumWDInteraction.GetAttribute(element, attribute, typeIdentification)
                    Case ActiveTool.Appium

                    Case Else

                End Select
                LogExecution("GetAttribute: " & attribute, StateLog.Finish)
                Return value
            Catch ex As Exception
                LogExecution("GetAttribute: " & attribute, StateLog.Exception)
                Return Nothing
            End Try
        End Function
        Function Encrypt(text As String) As String
            Try
                Dim encryp = New Encryption
                Return encryp.Cryptografy(text)
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
        Function Decrypt(text As String) As String
            Try
                Dim encryp = New Encryption
                Return encryp.Decrypt(text)
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
        Function OpenWindow(url As String, IDWindow As String) As Boolean
            LogExecution("OpenWindow", StateLog.Start)
            Test.Wait(p_speedyExecution)
            Try
                Select Case p_ToolName
                    Case p_ToolName
                        librarySeleniumWDInteraction.OpenWindow(url, IDWindow)
                    Case ActiveTool.Appium
                        libraryAppiumInteration.OpenWindow(url, IDWindow)
                    Case Else

                End Select
                LogExecution("OpenWindow", StateLog.Finish)
                Return True
            Catch ex As Exception
                LogExecution("OpenWindow error: " & ex.Message, StateLog.Exception)
                Return False
            End Try
        End Function
        Function WaitForPageToLoad(Optional TimeOut As Integer = 30000) As Boolean
            LogExecution("WaitForPageToLoad", StateLog.Start)
            Test.Wait(p_speedyExecution)

            Try
                Select Case p_ToolName
                    Case ActiveTool.SeleniumWD
                        librarySeleniumWDInteraction.WaitForPageToLoad(TimeOut)
                    Case ActiveTool.Appium
                        libraryAppiumInteration.WaitForPageToLoad(TimeOut)
                    Case Else
                End Select
                LogExecution("WaitForPageToLoad", StateLog.Finish)
                Return True
            Catch ex As Exception
                LogExecution("WaitForPageToLoad", StateLog.Exception)
                Return False
            End Try
        End Function
        Function SelectWindow(windowHandles As windowHandles) As Boolean
            LogExecution("SelectWindow: " & windowHandles, StateLog.Start)
            Test.Wait(p_speedyExecution)

            Try
                Select Case p_ToolName
                    'Case ActiveTool.SeleniumRC
                    '    librarySeleniumRCInteraction.SelectWindow(windowHandles)
                    Case ActiveTool.SeleniumWD
                        librarySeleniumWDInteraction.SelectWindow(windowHandles)
                    Case Else

                End Select
                LogExecution("SelectWindow: " & windowHandles, StateLog.Finish)
                Return True
            Catch ex As Exception
                LogExecution("SelectWindow error : " & ex.Message, StateLog.Exception)
                Return False
            End Try
        End Function
        Function SelectFrame(frameName As String) As Boolean
            LogExecution("SelectFrame: " & frameName, StateLog.Start)
            Test.Wait(p_speedyExecution)
            Try
                librarySeleniumWDInteraction.SelectFrame(frameName)
                LogExecution("SelectFrame: " & frameName, StateLog.Finish)
                Return True
            Catch ex As Exception
                LogExecution("SelectFrame error : " & ex.Message, StateLog.Exception)
                Return False
            End Try
        End Function
        Function GetCellData(element As String, value As String, Optional col As Integer = 0, Optional row As Integer = 0) As String
            Dim text As String = Nothing
            LogExecution("GetCellData", StateLog.Start)
            Test.Wait(p_speedyExecution)

            Try
                Select Case p_ToolName
                    'Case ActiveTool.SeleniumRC
                    '    text = librarySeleniumRCInteraction.GetCellData(element, col, row, value)
                    Case ActiveTool.SeleniumWD
                        text = librarySeleniumWDInteraction.GetCellData(element, col, row, value)
                        'Case ActiveTool.SeleniumWDBrowserStack
                        '    text = librarySeleniumWDBSInteraction.GetCellData(element, col, row, value)
                    Case ActiveTool.Appium
                        libraryAppiumInteration.GetCellData(element, col, row, value)
                    Case Else
                        LogExecution("GetCellData", StateLog.Finish)
                        Return Nothing
                End Select
                LogExecution("GetCellData", StateLog.Finish)
                Return text
            Catch ex As Exception
                LogExecution("GetCellData error: " & ex.Message, StateLog.Exception)
                Return False
            End Try
        End Function
        'Private Function [Set](ImgPathOrCoordinatesXY As String, value As String)
        '    Try
        '        Dim dv = New Devices
        '        dv.Type(ImgPathOrCoordinatesXY, value, p_speedyExecution)
        '        Return True
        '    Catch ex As Exception
        '        Return Nothing
        '    End Try
        'End Function
        Function [Set](elementOrImgPathOrCoordinatesXY As String, value As String, Optional fieldOut As String = Nothing, Optional typeIdentification As typeIdentification = typeIdentification.xpath) As Boolean
            LogExecution("Set: Element=" & elementOrImgPathOrCoordinatesXY & " - value:" & value & fieldOut, StateLog.Start)
            Wait(p_speedyExecution)
            If Mid(value, 1, 1) = "@" Then
                LogExecution("Set.HandlerValue value: " & value, StateLog.Information)
                value = HandlerValue(value)
            End If
            Try
                If typeIdentification = LibGlobal.typeIdentification.leanTest Then
                    pc_dv.Type(elementOrImgPathOrCoordinatesXY, value, p_speedyExecution)
                Else
                    Select Case p_ToolName
                        'Case ActiveTool.SeleniumRC
                        '    librarySeleniumRCInteraction.Type(elementOrImgPathOrCoordinatesXY, value)
                        Case ActiveTool.SeleniumWD
                            librarySeleniumWDInteraction.Type(elementOrImgPathOrCoordinatesXY, value, typeIdentification)
                            'Case ActiveTool.SeleniumWDBrowserStack
                            '    librarySeleniumWDBSInteraction.Type(elementOrImgPathOrCoordinatesXY, value)
                        Case ActiveTool.Appium
                            libraryAppiumInteration.Type(elementOrImgPathOrCoordinatesXY, value, typeIdentification)
                    End Select
                End If
                'caso seja neessário armazenar o valor digitado na tabela
                If Not String.IsNullOrEmpty(fieldOut) And Not String.IsNullOrEmpty(value) And typeIdentification <> LibGlobal.typeIdentification.leanTest Then
                    Test.GetText(elementOrImgPathOrCoordinatesXY, fieldOut, typeIdentification, True)
                    LogExecution("Set.GetValueAndSave element: " & elementOrImgPathOrCoordinatesXY & " - FieldOut:" & fieldOut, StateLog.Information)
                ElseIf typeIdentification = LibGlobal.typeIdentification.leanTest Then
                    If Not String.IsNullOrEmpty(value) Then Call SetValueOutput(fieldOut, value)
                End If
                LogExecution("Set: Element=" & elementOrImgPathOrCoordinatesXY & " - value:" & value & " - FieldOut:" & fieldOut, StateLog.Finish)
                Return True
            Catch ex As Exception
                LogExecution("Set error: " & ex.Message, StateLog.Exception)
                TestLog("Set element error", "Set element: " & elementOrImgPathOrCoordinatesXY & " - Value: " & value, ex.Message, typelog.Failed)
            End Try
        End Function
        Function HandlerValue(Value As String) As String
            Dim Values As String() = Split(Value, ";")
            Dim valueTest As String = Values(0)
            Dim parameter1 As String = Nothing
            Dim parameter2 As String = Nothing
            Dim generate = New GenerateDataTest
            Dim returnValue As String

            LogExecution("HandlerValue value: " & Value, StateLog.Start)
            Try
                parameter1 = Values(1)
                parameter2 = Values(2)
            Catch ex As Exception
            End Try
            Try
                Select Case UCase(valueTest)
                    Case UCase("@Date")
                        Dim amount As Integer = IIf(Not String.IsNullOrEmpty(parameter1), parameter1, 0)
                        Select Case Left(parameter1, 1)
                            Case "+"
                                parameter1 = Right(parameter1, Len(parameter1) - 1)
                                returnValue = DateAdd(DateInterval.Day, CDbl(parameter1), Today)
                                LogExecution("HandlerValue.Return: " & returnValue, StateLog.Finish)
                                Return returnValue
                            Case "-"
                                parameter1 = Right(parameter1, Len(parameter1) - 1)
                                returnValue = DateAdd(DateInterval.Day, -CDbl(parameter1), Today)
                                LogExecution("HandlerValue.Return: " & returnValue, StateLog.Finish)
                                Return returnValue
                            Case Else
                                returnValue = Date.Today
                                LogExecution("HandlerValue.Return: " & returnValue, StateLog.Finish)
                                Return returnValue
                        End Select
                    Case UCase("@CPF")
                        returnValue = generate.GenerateCPF
                        LogExecution("HandlerValue.Return: " & returnValue, StateLog.Finish)
                        Return returnValue
                    Case UCase("@Guid")
                        returnValue = generate.GenerateGUID
                        LogExecution("HandlerValue.Return: " & returnValue, StateLog.Finish)
                        Return returnValue
                    Case UCase("@Pasword")
                        Dim pass As Integer = IIf(Not String.IsNullOrEmpty(parameter1), parameter1, 6)
                        returnValue = generate.GeneratePassword(pass)
                        LogExecution("HandlerValue.Return: " & returnValue, StateLog.Finish)
                        Return returnValue
                    Case UCase("@Number")
                        Dim min As String = IIf(Not String.IsNullOrEmpty(parameter1), parameter1, 1)
                        Dim max As String = IIf(Not String.IsNullOrEmpty(parameter2), parameter2, 1)
                        returnValue = generate.GenerateRadonNumber(min, max)
                        LogExecution("HandlerValue.Return: " & returnValue, StateLog.Finish)
                        Return returnValue
                    Case UCase("@Word")
                        Dim len As Integer = IIf(Not String.IsNullOrEmpty(parameter1), parameter1, 1)
                        returnValue = generate.GenerateWord(len, parameter2)
                        LogExecution("HandlerValue.Return: " & returnValue, StateLog.Finish)
                        Return returnValue
                    Case UCase("@Email")
                        returnValue = generate.GenerateWord(6) & "@" & generate.GenerateWord(4) & ".com"
                        LogExecution("HandlerValue.Return: " & returnValue, StateLog.Finish)
                        Return returnValue
                    Case Else
                        returnValue = ""
                        LogExecution("HandlerValue.Return: " & returnValue, StateLog.Finish)
                        Return ""
                End Select
            Catch ex As Exception
                LogExecution("HandlerValue error: ", StateLog.Exception)
                Return Nothing
            End Try
        End Function
        Enum eScroll
            UP
            Down
        End Enum
        Sub Scroll(scroll As eScroll)
            LogExecution("Scroll", StateLog.Start)
            Test.Wait(p_speedyExecution)

            Try
                Select Case scroll
                    Case eScroll.UP
                        SendKey("{PGUP}")
                    Case eScroll.Down
                        SendKey("{PGDN}")
                End Select
                LogExecution("Scroll", StateLog.Finish)
            Catch ex As Exception
                HandlerError("SetValue")
            End Try
        End Sub
        Function IsEditable(element As String) As Boolean
            Dim value As Boolean
            LogExecution("IsEditable element: " & element, StateLog.Start)
            Wait(p_speedyExecution)
            Try
                Select Case p_ToolName
                    'Case ActiveTool.SeleniumRC
                    '    value = librarySeleniumRCInteraction.isEditable(element)
                    '    LogExecution("IsEditable value: " & value, StateLog.Finish)
                    '    Return value
                    Case ActiveTool.SeleniumWD
                        'value = librarySeleniumWDInteraction(element)
                        LogExecution("IsEditable value: " & value, StateLog.Finish)
                        Return value
                        'Case ActiveTool.SeleniumWDBrowserStack
                        '    value = librarySeleniumWDBSInteraction.isEditable(element)
                        '    LogExecution("IsEditable value: " & value, StateLog.Finish)
                        '    Return value
                    Case ActiveTool.Appium
                        value = libraryAppiumInteration.isEditable(element)
                End Select
                Return True
            Catch ex As Exception
                LogExecution("IsEditable error: " & ex.Message, StateLog.Exception)
                Return False
            End Try
        End Function
        Private Function SelectLean(ImgPathOrCoordinatesXY As String, value As String, indexItem As Integer)
            Try
                Dim dv = New Devices
                dv.SelectValue(ImgPathOrCoordinatesXY, value, indexItem, p_speedyExecution)
                Return True
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
        Function [Select](elementOrImgPathOrCoordinatesXY As String, value As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath, Optional indexItem As Integer = 1) As String
            If String.IsNullOrEmpty(value) Then Return True
            Test.Wait(p_speedyExecution)
            LogExecution("[Select] element: " & elementOrImgPathOrCoordinatesXY & " - Value: " & value, StateLog.Start)

            If Mid(value, 1, 1) = "@" Then value = HandlerValue(value)
            Try
                If typeIdentification = LibGlobal.typeIdentification.leanTest Then
                    SelectLean(elementOrImgPathOrCoordinatesXY, value, indexItem)
                Else
                    Select Case p_ToolName
                        'Case ActiveTool.SeleniumRC
                        '    librarySeleniumRCInteraction.SelectValue(elementOrImgPathOrCoordinatesXY, value)
                        'Case ActiveTool.SeleniumWDBrowserStack
                        '    librarySeleniumWDBSInteraction.SelectValue(elementOrImgPathOrCoordinatesXY, value)
                        Case ActiveTool.SeleniumWD
                            librarySeleniumWDInteraction.SelectValue(elementOrImgPathOrCoordinatesXY, value, typeIdentification)
                        Case ActiveTool.Appium
                            libraryAppiumInteration.SelectValue(elementOrImgPathOrCoordinatesXY, value, typeIdentification)
                    End Select
                End If
                LogExecution("[Select] element: " & elementOrImgPathOrCoordinatesXY & " - Value: " & value, StateLog.Finish)
                Return True
            Catch ex As Exception
                LogExecution("[Select] error: " & ex.Message, StateLog.Exception)
                Throw New Exception(ex.Message)
            End Try
        End Function
        Private Function Click(ImgPathOrCoordinatesXY As String, waitElement As Boolean)
            Try
                Dim dv = New Devices
                If waitElement Then WaitExist(ImgPathOrCoordinatesXY, typeIdentification.leanTest)
                If Not dv.Click(ImgPathOrCoordinatesXY, p_speedyExecution) Then
                    Throw New Exception("Element not found: " & ImgPathOrCoordinatesXY)
                Else
                    Return True
                End If
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Function Click(elementOrImgPathOrCoordinatesXY As String, Value As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath, Optional waitElement As Boolean = True) As Boolean
            Try
                Value = CBool(Value)
                If Not CBool(Value) Then Return True
            Catch ex As Exception
                Return True
            End Try
            LogExecution("Click element: " & elementOrImgPathOrCoordinatesXY & " - value: " & Value, StateLog.Start)
            Wait(p_speedyExecution)
            Try
                If typeIdentification = LibGlobal.typeIdentification.leanTest Then
                    Click(elementOrImgPathOrCoordinatesXY, waitElement)
                Else
                    Select Case p_ToolName
                        'Case ActiveTool.SeleniumRC
                        '    librarySeleniumRCInteraction.click(elementOrImgPathOrCoordinatesXY)
                        Case ActiveTool.SeleniumWD
                            librarySeleniumWDInteraction.Click(elementOrImgPathOrCoordinatesXY, typeIdentification, waitElement)
                            'Case ActiveTool.SeleniumWDBrowserStack
                            '    librarySeleniumWDBSInteraction.click(elementOrImgPathOrCoordinatesXY)
                        Case ActiveTool.Appium
                            libraryAppiumInteration.Click(elementOrImgPathOrCoordinatesXY, typeIdentification)
                    End Select
                End If
                LogExecution("Click: Element=" & elementOrImgPathOrCoordinatesXY & " - value:" & Value, StateLog.Finish)
            Catch ex As Exception
                LogExecution("libGlobal.Click: " & ex.Message, StateLog.Exception)
                Throw New Exception(ex.Message)
            End Try
            HandlerMessage()
        End Function
        Private Function MouseOver(ImgPathOrCoordinatesXY As String) As Boolean
            Try
                Dim dv = New Devices
                Return dv.MouseOver(ImgPathOrCoordinatesXY)
            Catch ex As Exception
                Return False
            End Try
        End Function
        Function MouseOver(elementOrImgPathOrCoordinatesXY As String, value As Boolean, Optional typeIdentification As typeIdentification = LibGlobal.typeIdentification.xpath) As Boolean
            Try
                value = CBool(value)
            Catch ex As Exception
                Return True
            End Try
            LogExecution("MouseOver element: " & elementOrImgPathOrCoordinatesXY, StateLog.Start)
            Wait(p_speedyExecution)
            Try
                If typeIdentification = LibGlobal.typeIdentification.leanTest Then
                    MouseOver(elementOrImgPathOrCoordinatesXY)
                Else
                    Select Case p_ToolName
                        Case ActiveTool.SeleniumRC
                            'librarySeleniumRCInteraction.DoubleClick(element)
                        Case ActiveTool.SeleniumWD
                            librarySeleniumWDInteraction.MouseOver(elementOrImgPathOrCoordinatesXY, value, typeIdentification)
                        Case ActiveTool.SeleniumWDBrowserStack
                            'librarySeleniumWDBSInteraction.DoubleClick(element)
                        Case ActiveTool.Appium
                            'libraryAppiumInteration.DoubleClick(element)
                    End Select
                End If
                LogExecution("MouseOver element: " & elementOrImgPathOrCoordinatesXY, StateLog.Finish)
            Catch ex As Exception
                LogExecution("libGlobal.MouseOver: " & ex.Message, StateLog.Exception)
                Throw New Exception(ex.Message)
            End Try
            Return True
        End Function

        Private Function DoubleClick(ImgPathOrCoordinatesXY As String, waitElement As Boolean)
            Try
                Dim dv = New Devices
                If waitElement Then WaitExist(ImgPathOrCoordinatesXY, typeIdentification.leanTest)
                If Not dv.DoubleClick(ImgPathOrCoordinatesXY, p_speedyExecution) Then
                    Throw New Exception("Element not found: " & ImgPathOrCoordinatesXY)
                Else
                    Return True
                End If
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function

        Function DoubleClick(ImgPathOrCoordinatesXY As String, value As Boolean, Optional typeIdentification As typeIdentification = typeIdentification.xpath, Optional waitElement As Boolean = True) As Boolean
            Try
                value = CBool(value)
            Catch ex As Exception
                Return True
            End Try
            LogExecution("DoubleClick element: " & ImgPathOrCoordinatesXY, StateLog.Start)
            Wait(p_speedyExecution)
            Try
                If typeIdentification = LibGlobal.typeIdentification.leanTest Then
                    DoubleClick(ImgPathOrCoordinatesXY, waitElement)
                Else
                    Select Case p_ToolName
                        Case ActiveTool.SeleniumWD
                            librarySeleniumWDInteraction.DoubleClick(ImgPathOrCoordinatesXY, typeIdentification)
                        Case ActiveTool.Appium
                            libraryAppiumInteration.DoubleClick(ImgPathOrCoordinatesXY)
                    End Select
                End If
                LogExecution("DoubleClick element: " & ImgPathOrCoordinatesXY, StateLog.Finish)
            Catch ex As Exception
                LogExecution("libGlobal.Click: " & ex.Message, StateLog.Exception)
                Throw New Exception(ex.Message)
            End Try
            Return True
        End Function
        Function GetTextOCR(coordinates As String) As String
            Wait(p_speedyExecution)

            Dim ocr As New OCR
            Dim sc = New API.ScreenShot
            Dim text As String
            Dim pathAccess As String() = xml.Read("pathAccess", "").Split("\")
            Dim LocalPath As String = "C:\LeanTestAutomation\Scripts\Execution\" & pathAccess(4) & "\captureImg_" & p_TableTest & "\"

            Try
                text = ocr.GetTextOCR(sc.CaptureScreenByCoordinates(LocalPath, "img", coordinates))
                Return text
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
        Public Function GetPositionByImage(pathImage2 As String, Optional acceptancePercentage As String = "1") As String
            Try
                Dim scr = New API.ScreenShot
                Return scr.GetPositionByImage(pathImage2, acceptancePercentage)
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
        Private Function GetText(ImgPathOrCoordinatesXY As String)
            Try
                Dim dv = New Devices
                Return dv.GetText(ImgPathOrCoordinatesXY, p_speedyExecution)
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
        Function GetText(ImgPathOrCoordinatesXY As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath) As String
            Dim text As String = Nothing
            LogExecution("GetText: element=" & ImgPathOrCoordinatesXY, StateLog.Start)
            Wait(p_speedyExecution)
            Try
                If typeIdentification = LibGlobal.typeIdentification.leanTest Then
                    Return GetText(ImgPathOrCoordinatesXY)
                Else
                    Select Case p_ToolName
                        'Case ActiveTool.SeleniumRC
                        '    text = librarySeleniumRCInteraction.GetText(ImgPathOrCoordinatesXY)
                        'Case ActiveTool.SeleniumWDBrowserStack
                        '    text = librarySeleniumWDBSInteraction.GetText(ImgPathOrCoordinatesXY)
                        Case ActiveTool.SeleniumWD
                            text = librarySeleniumWDInteraction.GetText(ImgPathOrCoordinatesXY, typeIdentification)
                        Case ActiveTool.Appium
                            text = libraryAppiumInteration.GetText(ImgPathOrCoordinatesXY, typeIdentification)
                        Case Else
                            text = Nothing
                            Return Nothing
                    End Select
                End If
                LogExecution("GetText value: " & text, StateLog.Finish)
                Return text
            Catch ex As Exception
                LogExecution("GetText error: " & ex.Message, StateLog.Exception)
                Return Nothing
            End Try
        End Function
        Function GetValue(element As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath) As String
            Dim value As String
            LogExecution("GetValue: element=" & element, StateLog.Start)
            Wait(p_speedyExecution)

            Try
                Select Case p_ToolName
                    'Case ActiveTool.SeleniumRC
                    '    value = librarySeleniumRCInteraction.GetValue(element)
                    'Case ActiveTool.SeleniumWDBrowserStack
                    '    value = librarySeleniumWDBSInteraction.GetValue(element)
                    Case ActiveTool.SeleniumWD
                        value = librarySeleniumWDInteraction.GetValue(element, typeIdentification)
                    Case ActiveTool.Appium
                        value = libraryAppiumInteration.GetValue(element, typeIdentification)
                    Case Else
                        value = Nothing
                End Select
                LogExecution("GetValue value: " & value, StateLog.Finish)
                Return value
            Catch ex As Exception
                LogExecution("GetValue error: " & ex.Message, StateLog.Exception)
                Return Nothing
            End Try
        End Function
        Function GetTextAndSave(element As String, field As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath) As String
            LogExecution("GetTextAndSave element: " & element, StateLog.Start)
            Wait(p_speedyExecution)

            Dim msg As String
            Dim timeout = p_timeout

            Try
                If Not CBool(GetValueOutput(p_IDTest, p_TableTest, field)) Then Return Nothing
            Catch ex As Exception
                SetValueOutput(field, "")
            End Try

            p_timeout = 1
            Try
                Select Case p_ToolName
                    'Case ActiveTool.SeleniumRC
                    '    msg = librarySeleniumRCInteraction.GetText(element)
                    '    Call SetValueOutput(field, msg)
                    'Case ActiveTool.SeleniumWDBrowserStack
                    '    msg = librarySeleniumWDBSInteraction.GetText(element)
                    '    Call SetValueOutput(field, msg)
                    Case ActiveTool.SeleniumWD
                        msg = librarySeleniumWDInteraction.GetText(element, typeIdentification)
                        Call SetValueOutput(field, msg)
                    Case ActiveTool.Appium
                        msg = libraryAppiumInteration.GetText(element, typeIdentification)
                        Call SetValueOutput(field, msg)
                    Case Else
                        p_timeout = timeout
                        Return Nothing
                End Select
                LogExecution("GetTextAndSave text: " & msg, StateLog.Finish)
                p_timeout = timeout
                Return msg
            Catch ex As Exception
                p_timeout = timeout
                LogExecution("GetValue: " & ex.Message, StateLog.Exception)
                Return Nothing
            End Try
        End Function
        Function GetTextOCRAndSave(coordinates As String, field As String) As String
            LogExecution("GetTexOCRtAndSave name: " & field, StateLog.Start)
            Wait(p_speedyExecution)

            Dim msg As String
            Dim timeout = p_timeout

            Try
                If Not CBool(GetValueOutput(p_IDTest, p_TableTest, field)) Then Return Nothing
            Catch ex As Exception
                SetValueOutput(field, "")
            End Try
            p_timeout = 1
            Try
                msg = GetTextOCR(coordinates)
                Call SetValueOutput(field, msg)
                p_timeout = timeout
                LogExecution("GetTexOCRtAndSave text: " & msg, StateLog.Finish)
                Return msg
            Catch ex As Exception
                p_timeout = timeout
                LogExecution("GetTexOCRtAndSave: " & ex.Message, StateLog.Exception)
                Return Nothing
            End Try
        End Function


        Function GetSelectedValueAndSave(element As String, field As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath) As String
            Dim msg As String
            LogExecution("GetSelectedValueAndSave element: " & element & " - Field: " & field, StateLog.Start)
            Wait(p_speedyExecution)

            Try
                Select Case p_ToolName
                    'Case ActiveTool.SeleniumRC
                    '    msg = librarySeleniumRCInteraction.GetSelectedValue(element)
                    '    Call SetValueOutput(field, msg)
                    'Case ActiveTool.SeleniumWDBrowserStack
                    '    msg = librarySeleniumWDBSInteraction.GetSelectedValue(element)
                    '    Call SetValueOutput(field, msg)
                    Case ActiveTool.SeleniumWD
                        msg = librarySeleniumWDInteraction.GetSelectedValue(element)
                        Call SetValueOutput(field, msg)
                    Case ActiveTool.Appium
                        msg = libraryAppiumInteration.GetSelectedValue(element)
                        Call SetValueOutput(field, msg)
                    Case Else
                        LogExecution("GetSelectedValueAndSave value: ", StateLog.Finish)
                        Return Nothing
                End Select
                LogExecution("GetSelectedValueAndSave value: " & msg, StateLog.Finish)
                Return msg
            Catch ex As Exception
                'LogExecution("GetSelectedValueAndSave: " & ex.Message, StateLog.Exception)
                Return Nothing
            End Try
        End Function
        Function GetText(ImgPathOrCoordinatesXY As String, field As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath, Optional isType As Boolean = False) As String
            Dim msg As String
            LogExecution("GetValueAndSave element: " & ImgPathOrCoordinatesXY & " - Field: " & field, StateLog.Start)
            Wait(p_speedyExecution)
            If Not isType Then
                Try
                    If Not CBool(GetValueOutput(p_IDTest, p_TableTest, field)) Then Return Nothing
                Catch ex As Exception
                    SetValueOutput(field, "")
                End Try
            End If
            Try
                If typeIdentification = LibGlobal.typeIdentification.leanTest Then
                    msg = GetText(ImgPathOrCoordinatesXY)
                    If Not String.IsNullOrEmpty(msg) Then Call SetValueOutput(field, msg)
                Else
                    Select Case p_ToolName
                        'Case ActiveTool.SeleniumRC
                        '    msg = librarySeleniumRCInteraction.GetValue(ImgPathOrCoordinatesXY)
                        '    If Not String.IsNullOrEmpty(msg) Then Call SetValueOutput(field, msg)
                        'Case ActiveTool.SeleniumWDBrowserStack
                        '    msg = librarySeleniumWDBSInteraction.GetValue(ImgPathOrCoordinatesXY)
                        '    If Not String.IsNullOrEmpty(msg) Then Call SetValueOutput(field, msg)
                        Case ActiveTool.SeleniumWD
                            msg = librarySeleniumWDInteraction.GetValue(ImgPathOrCoordinatesXY, typeIdentification)
                            If Not String.IsNullOrEmpty(msg) Then Call SetValueOutput(field, msg)
                        Case ActiveTool.Appium
                            msg = libraryAppiumInteration.GetValue(ImgPathOrCoordinatesXY, typeIdentification)
                            If Not String.IsNullOrEmpty(msg) Then Call SetValueOutput(field, msg)
                        Case Else
                            LogExecution("GetValueAndSave value:", StateLog.Finish)
                            Return Nothing
                    End Select
                End If
                LogExecution("GetValueAndSave value: " & msg, StateLog.Finish)
                Return msg
            Catch ex As Exception
                'LogExecution("GetValue: " & ex.Message, StateLog.Exception)
                Return Nothing
            End Try
        End Function
        Function GetSelectedValue(element As String) As String
            Dim msg As String = Nothing
            LogExecution("GetSelectedValue element: " & element, StateLog.Start)
            Try
                Select Case p_ToolName
                    'Case ActiveTool.SeleniumRC
                    '    msg = librarySeleniumRCInteraction.GetSelectedValue(element)
                    'Case ActiveTool.SeleniumWDBrowserStack
                    '    msg = librarySeleniumWDBSInteraction.GetSelectedValue(element)
                    Case ActiveTool.Appium
                        msg = libraryAppiumInteration.GetSelectedValue(element)
                    Case Else
                        LogExecution("GetSelectedValue value: " & msg, StateLog.Finish)
                        Return Nothing
                End Select
                LogExecution("GetSelectedValue value: " & msg, StateLog.Finish)
                Return msg
            Catch ex As Exception
                LogExecution("GetSelectedValue error: " & ex.Message, StateLog.Exception)
                Return Nothing
            End Try
        End Function

        Function Exist(element As String, timeMilliSeconds As Integer, Optional typeIdentification As typeIdentification = typeIdentification.xpath) As Boolean
            Dim value As Boolean
            LogExecution("Exist element: " & element & " time out: " & timeMilliSeconds, StateLog.Start)
            Wait(p_speedyExecution)

            Try
                Select Case p_ToolName
                    'Case ActiveTool.SeleniumRC
                    '    value = librarySeleniumRCInteraction.Exist(element, timeMilliSeconds)
                    'Case ActiveTool.SeleniumWDBrowserStack
                    '    value = librarySeleniumWDBSInteraction.Exist(element, timeMilliSeconds)
                    Case ActiveTool.SeleniumWD
                        value = librarySeleniumWDInteraction.Exist(element, timeMilliSeconds, typeIdentification)
                    Case ActiveTool.Appium
                        value = libraryAppiumInteration.Exist(element, timeMilliSeconds, typeIdentification)
                End Select
                LogExecution("Exist: " & value, StateLog.Finish)
                Return value
            Catch ex As Exception
                HandlerError("Exist" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
        'Function WaitExist(coordinatesAndOrPathImageToCompare As String, acceptancePercentage As String, Optional milliSecondsWait As Integer = 500) As Boolean
        '    Try
        '        For i = 0 To p_timeout - 1
        '            Try
        '                If acceptancePercentage = "0" Then acceptancePercentage = "1,0"
        '                If WaitExistImage(coordinatesAndOrPathImageToCompare, acceptancePercentage, milliSecondsWait) Then
        '                    Return True
        '                Else
        '                    Test.Wait(1000)
        '                End If
        '                Console.Write("Element: " & coordinatesAndOrPathImageToCompare & " - Time: " & i & vbCrLf)
        '            Catch ex As Exception
        '                Console.Write("Element: " & coordinatesAndOrPathImageToCompare & " - Time: " & i & vbCrLf)
        '                Test.Wait(1000) ' clique em pause, F11 e em seguida arraste o cursor para o End function
        '            End Try
        '        Next
        '        Throw New Exception("waitLoadElement: Element=" & coordinatesAndOrPathImageToCompare & " not found")
        '        Return False
        '    Catch ex As Exception
        '        Throw New Exception(ex.Message)
        '        Return False
        '    End Try
        'End Function
        Function WaitNotExist(element As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath) As Boolean
            Try
                For i = 0 To p_speedyExecution
                    If Not WaitExist(element, typeIdentification) Then
                        Return True
                    End If
                Next
                Return False
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
        Function WaitExist(coordinatesAndOrPathImageToCompare As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath, Optional acceptancePercentage As String = "0,9", Optional SecondsWait As Integer = 20, Optional SearchImage As Boolean = True) As Boolean
            LogExecution("WaitExist: element=" & coordinatesAndOrPathImageToCompare, StateLog.Start)
            Wait(p_speedyExecution)

            Try
                'corrigir throw dos demais cases. somente wd corrigido
                If typeIdentification = LibGlobal.typeIdentification.leanTest Then
                    Return WaitExistImage(coordinatesAndOrPathImageToCompare, acceptancePercentage, SecondsWait, SearchImage)
                Else
                    Select Case p_ToolName
                        'Case ActiveTool.SeleniumRC
                        '    Return librarySeleniumRCInteraction.WaitExist(coordinatesAndOrPathImageToCompare)
                        'Case ActiveTool.SeleniumWDBrowserStack
                        '    Return librarySeleniumWDBSInteraction.WaitExist(coordinatesAndOrPathImageToCompare)
                        Case ActiveTool.SeleniumWD
                            Return librarySeleniumWDInteraction.WaitExist(coordinatesAndOrPathImageToCompare, typeIdentification)
                        Case ActiveTool.Appium
                            Return libraryAppiumInteration.WaitExist(coordinatesAndOrPathImageToCompare, typeIdentification)
                    End Select
                End If
            Catch ex As Exception
                LogExecution("WaitExist error: " & ex.Message, StateLog.Exception)
                Return ex.Message
            End Try
            Return False
        End Function
        Function WaitNotExist(coordinatesAndOrPathImageToCompare As String, Optional initialTimeMilliseconds As Integer = 500, Optional value As Boolean = True, Optional acceptancePercentage As String = "0,9", Optional milliSecondsWait As Integer = 0) As Boolean
            Try
                For i = 0 To p_speedyExecution
                    If Not WaitExistImage(coordinatesAndOrPathImageToCompare, value, acceptancePercentage) Then
                        Return True
                    End If
                Next
                Return False
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
        Private Function WaitExistImage(coordinatesAndOrPathImageToCompare As String, Optional acceptancePercentage As String = "0,9", Optional SecondsWait As Integer = 20, Optional SearchImage As Boolean = True) As Boolean
            Try
                If SecondsWait = 0 Then SecondsWait = p_timeout
                Dim scr = New API.ScreenShot
                LogExecution("WaitExistImage: " & coordinatesAndOrPathImageToCompare, StateLog.Information)
                Wait(p_speedyExecution)
                Return scr.WaitExistImage(coordinatesAndOrPathImageToCompare, acceptancePercentage, SecondsWait, SearchImage)
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Function GetPositionXY(pathImage As String, Optional acceptancePercentage As String = "0,9")
            Try
                Dim scr = New API.ScreenShot
                Dim coord As String() = scr.GetPositionByImage(pathImage, acceptancePercentage).Split(",")
                Return coord(0) & "," & coord(1)
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        Function OpenApp(appPath As String) As Boolean
            Wait(p_speedyExecution)
            Try
                LogExecution("Open appPath: " & appPath, StateLog.Start)
                pc_util.StartProcess(appPath)
                LogExecution("Open: appPath=" & appPath, StateLog.Finish)
                Return True
            Catch ex As Exception
                LogExecution("Open appPath: " & appPath, StateLog.Exception)
                Return False
            End Try
        End Function
        Function CloseApp(apphName As String) As Boolean
            Wait(p_speedyExecution)
            Try
                LogExecution("Close app: " & apphName, StateLog.Start)
                pc_util.StopProcess(apphName)
                LogExecution("Close: apphName=" & apphName, StateLog.Finish)
                Return True '
            Catch ex As Exception
                LogExecution("Close apphName: " & apphName, StateLog.Exception)
                Return False
            End Try
        End Function
        Public Sub Wait(milliSeconds As Integer)
            Thread.Sleep(milliSeconds)
        End Sub
        Public Sub MouseMove(x As String, y As String)
            Wait(p_speedyExecution)

            Dim dv = New Devices
            dv.MouseMove(x, y)
        End Sub
        Public Sub MouseShow(x As String, y As String)
            Wait(p_speedyExecution)
            Dim dv = New Devices
            dv.MouseShow()
        End Sub
        Public Sub MouseHide(x As String, y As String)
            Wait(p_speedyExecution)
            Dim dv = New Devices
            dv.MouseHide()
        End Sub
        Public Sub MouseClick(x As Integer, y As Integer, Optional value As Boolean = True)
            If Not value Then Exit Sub
            Wait(p_speedyExecution)
            Dim dv = New Devices
            Dim p = New MousePositions
            dv.MouseClick(x, y)
        End Sub
        Public Sub MouseDragAndDrop(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)
            Wait(p_speedyExecution)
            Dim dv = New Devices
            dv.MouseDragDrop(x1, y1, x2, y2)
        End Sub
        Public Sub SendKey(value As String, Optional AmountSendKey As Integer = 1, Optional WaitTimeMilliseconds As Integer = 200, Optional outField As String = Nothing)
            LogExecution("SendKey value: " & value & " - outField:" & outField, StateLog.Start)
            Wait(p_speedyExecution)

            value = IIf(Mid(value, 1, 1) = "@", HandlerValue(value), value)
            Select Case p_ToolName
                Case ActiveTool.Appium
                    For i = 1 To AmountSendKey
                        Wait(WaitTimeMilliseconds)
                        objAppium.Keyboard.SendKeys(value)
                    Next
                Case Else
                    Dim dv = New Devices
                    For i = 1 To AmountSendKey
                        Wait(WaitTimeMilliseconds)
                        dv.SendKeyWait(value)
                    Next
            End Select
            If Not String.IsNullOrEmpty(outField) And Not String.IsNullOrEmpty(value) Then
                Test.SetValueOutput(outField, value)
                LogExecution("SendKey value: " & value & " - outField:" & outField, StateLog.Finish)
            End If
        End Sub
        Public Sub Reflesh()
            Select Case ActiveTool.SeleniumWD
                Case ActiveTool.SeleniumWD
                    librarySeleniumWDInteraction.Reflesh()
                Case ActiveTool.Appium
                    libraryAppiumInteration.Refresh()
            End Select
        End Sub
        Public Shared Function HandlerError(sourceHandler As String) As Integer
            Try
                If Err.Number <> 0 Then
                    'pc_log.WriteLog("[Exception] - " & Now & " - Event :(" & sourceHandler & ")", False)
                Else
                    'pc_log.WriteLog("[Message] - " & Now & " - Event :(" & sourceHandler & ")", False)
                End If
                Err.Number = 0
                Return 1
            Catch ex As Exception
                Return 1
            End Try
        End Function
        'Public Function TestLog(stepName As String, ExpectedResults As String, Check As String) As Boolean
        '    'colocar aqui a validação do locator
        '    Dim ActualResults As String = Nothing
        '    Dim errors As String = Nothing
        '    LogExecution("TestLog Step: " & stepName & " - Expected Results: " & ExpectedResults & " - Check : " & Check, StateLog.Start)
        '    If Not String.IsNullOrEmpty(Check) Then
        '        Dim sourceHTML As String = UCase(Test.GetHtmlSource)
        '        If InStr(sourceHTML, UCase(Check)) Then
        '            ActualResults = "Step executado com sucesso"
        '            p_StatusTest = "Passed"
        '            LogExecution("TestLog.WaitExist: True - Actual Results: " & ActualResults, StateLog.Information)
        '        Else
        '            ActualResults = "Step executado com falha"
        '            errors = "Step executado com falha, element not found!"
        '            p_StatusTest = "Failed"
        '            CreateDefect(stepName, ExpectedResults, ActualResults)
        '            LogExecution("TestLog.WaitExist: False - Actual Results: " & ActualResults & "Element not found", StateLog.Information)
        '        End If
        '    Else
        '        p_StatusTest = "Passed"
        '    End If
        '    Dim strLocalPath As String

        '    Try
        '        Dim lenID As Integer = Len(CStr(p_IDScenario))
        '        Dim pathTemp As String = "\SCEN"
        '        pathTemp = pathTemp & Strings.StrDup(5 - lenID, "0") & p_IDScenario

        '        strLocalPath = p_pathEvidence & p_pathLab & "\" & pathTemp & "\" & p_TestName & "\Screenshot\"
        '        strLocalPath = Replace(strLocalPath, "\\", "\")

        '        'gera print screen pelo framework ou usando a API do browsersake remoto
        '        Select Case p_ToolName
        '            Case ActiveTool.SeleniumWDBrowserStack
        '                Dim util = New Utilities
        '                Dim dir = New Directorys
        '                Dim screenshot As OpenQA.Selenium.Screenshot = objSeleniumWDBS.GetScreenshot()
        '                Dim nameIMG As String = Nothing
        '                Dim temp As String = Format(Now, "MMddHHmm") & Now.Millisecond

        '                strLocalPath = strLocalPath & "\"
        '                strLocalPath = strLocalPath.Replace("\\", "\")

        '                nameIMG = util.ConvertToAlphanumeric(stepName) & "_" & temp & ".png"
        '                nameIMG = Right(nameIMG, 80)

        '                If Not System.IO.Directory.Exists(strLocalPath) Then
        '                    If Not dir.CreateDiretory(strLocalPath) Then MsgBox("TestLog error in create file path")
        '                End If
        '                strLocalPath = strLocalPath & nameIMG
        '                screenshot.SaveAsFile(strLocalPath, System.Drawing.Imaging.ImageFormat.Png)
        '                Dim Sql = "INSERT INTO LeanTestLogs (StepName, ExpectedResults, ActualResults, StatusExecution, PathEvidence, IDScenario, IDTest, IDRun, CreationDate) values (" &
        '                "'" & "[" & p_TestName & "] - " & stepName & "', '" & ExpectedResults & "', '" & Replace(ActualResults, "'", "''") & "', '" & p_StatusTest & "', '" & strLocalPath & "', " & p_IDScenario & ", '" & p_IDTest & "', '" & p_IDRun & "', '" & Now & "')"
        '                pc_db.ExecuteCommandSQL(Sql)
        '                'Case ActiveTool.SeleniumWD

        '            Case Else
        '                strLocalPath = strLocalPath.Replace(" \", "\").Replace("\ ", "\")
        '                strLocalPath = pc_te.TestLog("[" & p_TestName & "] - " & stepName, ExpectedResults, ActualResults, p_StatusTest, p_IDScenario, p_IDTest, strLocalPath, p_IDRun)

        '        End Select
        '        If p_StatusScenario <> "Failed" Then
        '            p_StatusScenario = p_StatusTest
        '        End If
        '        If p_StatusTest = "Failed" Then
        '            p_StatusScenario = "Failed"
        '        End If
        '        'publish QC
        '        If p_PublishQC Then
        '            Dim fullpath = Path.GetDirectoryName(strLocalPath)
        '            Dim name = Path.GetFileName(strLocalPath)
        '            If name <> Nothing Then
        '                p_qc.AddAttach(p_Run, fullpath, name, ExpectedResults)
        '                p_qc.AddStepRun(p_Run, p_StatusTest, stepName, ExpectedResults, ExpectedResults, ActualResults)
        '            End If
        '        End If
        '        LogExecution("TestLog", StateLog.Finish)
        '        Return True
        '    Catch ex As Exception
        '        LogExecution("TestLog: " & ex.Message, StateLog.Exception, True)
        '        Throw New Exception("TestLog error: " & ex.Message & " - " & errors)
        '    End Try
        'End Function
        Private Function VerifyExistsDefect(summary As String) As Boolean
            'Dim sql As String = "Select * from LeanTestDefects Where summary='" & summary & "' and idTest ='" & p_IDTest & "' and IDScenario='" & p_IDScenario & "' and Status<> 'Closed'"
            Dim sql As String = "Select * from LeanTestDefects Where summary='" & summary & "' and IDScenario='" & p_IDScenario & "' and Status <> 'Closed'"
            Try
                If Not String.IsNullOrEmpty(pc_db.GetValue(sql, "Summary")) Then Return True Else Return False
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        Private Sub CreateDefect(stepTest As String, ExpectedResults As String, actualResults As String)

            LogExecution("CreateDefect", StateLog.Start)
            Dim email = New Email
            Dim net = New NetworkInformation
            Dim sql As String
            Dim cols As String
            Dim db = New DBAdapter
            Dim file = New FileTexts

            Dim defect As String = file.ReadFile("C:\LeanTestAutomation\Template\templateDefectEmail.leantest")
            Dim emailSend As String = db.GetValue("Select * from LeanTestDemands", "EmailsReporte", True, Nothing, Nothing)

            Dim summary, description, CreationDate, status, severity, priority, detectedBy, AssignTo, typeIncident, demand, detectedRelease, detectedCycle,
                targetRelease, targetCycle, Application, Functionality, area, solutionType, solution, comments, idTool, idTest,
                idScenario, dataTest, StepsTest, ImportedData, Attachments As String

            Dim scen As String = "SCEN"

            scen += Strings.StrDup(4 - Len(CStr(p_IDScenario)), "0") & p_IDScenario

            'summary = "Incident - Test: " & p_TestName & " - Step: " & stepTest & " - IDTest:" & p_IDTest & " - IDScenario - " & p_IDScenario
            summary = "[Incident Automation Test] - Incidente encontrado no cenário " & scen & " - " & p_ScenarioName & " - [Test: " & p_TestName & "]"
            If VerifyExistsDefect(summary) Then Exit Sub

            cols = pc_db.GetColumnsTable(p_TableTest)
            cols = Mid(cols, 5, Len(cols) - 4)
            Dim fields As String = Nothing
            Dim values As String = Nothing
            Dim countCols() = cols.Split(",")
            Dim sqlData As String = "Select * From " & p_TableTest & " Where IDTest ='" & p_IDTest & "'"
            For i = 0 To UBound(countCols)
                If Mid(Trim(countCols(i)), 1, 1) = "v" Then
                    If Not InStr(countCols(i), "vbtn") Then
                        values = pc_db.GetValue(sqlData, countCols(i), False)
                        If Not String.IsNullOrEmpty(values) Then fields += countCols(i) & "=" & values & vbCrLf
                    End If
                End If
            Next
            dataTest = Left(fields, Len(fields) - 1) 'remove vbcrlf do final
            ImportedData = "False"
            pc_db.OpenTestLogByIDScenario(p_IDScenario)
            Try
                StepsTest = Nothing
                Attachments = Nothing
                Do While pc_db.drAccess.Read
                    StepsTest += "<b>Step: </b>" & pc_db.Fieldt("StepName") & "<br/>" &
                                 "<b>Expected Results: </b>" & pc_db.Fieldt("ExpectedResults") & "<br/>" &
                                 "<b>Actual Results: </b>" & pc_db.Fieldt("ActualResults") & "<br/>" &
                                 "<b>Status Execution: </b>" & pc_db.Fieldt("StatusExecution") & "<p>"
                    Attachments += pc_db.Fieldt("PathEvidence") & ";"
                Loop
                StepsTest = StepsTest.Replace("'", "")
                If Not String.IsNullOrEmpty(Attachments) Then Attachments = Left(Attachments, Len(Attachments) - 1) 'remove o ultimo ;

                description = "<b>AUTMATION TEST - Incidente identificado automaticamente</b>" & vbCrLf &
                       "<b>Scenario: </b>" & scen & p_ScenarioName & "<br/>" &
                       "<b>Test: </b>" & p_TestName & "<br/>" &
                       "<b>Decription: </b>" & p_DescriptionTest & "<br/>" &
                       "<b>Step:</b> " & stepTest & "<br/>" &
                       "<b>Expected Result: </b>" & ExpectedResults & "<br/>" &
                       "<b>Actual Result: </b>" & actualResults & "<p>" &
                       "<b>Informations aditionais incident: </b>" & p_errorDescription & "<br/>" &
                       "<b>Execution date: </b>" & Now & "<br/>" &
                       "<b>Machine execution: </b>" & net.GetHostName & "<br/>" &
                       "<b>Executed by: </b>" & net.GetCurrentUser & "<br/>" &
                       "<b>Steps: </b>" & StepsTest
                CreationDate = Now
                status = "New"
                severity = "Media"
                priority = "Media"
                detectedBy = net.GetCurrentUser
                AssignTo = ""
                typeIncident = "Incident"
                demand = db.GetValue("Select * from LeanTestDemands", "Demand", True, Nothing, Nothing)
                detectedRelease = p_Release
                detectedCycle = p_Cycle
                targetRelease = p_Release
                targetCycle = p_Cycle
                Application = ""
                Functionality = ""
                area = ""
                solutionType = ""
                solution = ""
                comments = ""
                idTool = ""
                idTest = p_IDTest
                idScenario = p_IDScenario

                description = description.Replace("'", "")

                sql = "Insert Into LeanTestDefects " &
                              "(IDDefect, Summary, CreationDate, Status, Severity, DetectBy, AssignTo, TypeIncident, Demand, " &
                              "DetectedRelease, DetectedCycle, TargetRelease, TargetCycle, Application, Functionality, Area, SolutionType, " &
                              "Description, Solution, Comments, IDTool, IDTest, IDScenario, DataTest, StepsTest, ImportedDate, Attachments) values " &
                              "('','" & summary & "', '" & CreationDate & "', '" & status & "', '" & severity & "', '" & detectedBy & "', '" & AssignTo & "', '" &
                              typeIncident & "', '" & demand & "', '" & detectedRelease & "', '" & detectedCycle & "', '" & targetRelease & "', '" &
                              targetCycle & "', '" & Application & "', '" & Functionality & "', '" & area & "', '" & solutionType & "', '" &
                              description & "', '" & solution & "', '" & comments & "', '" & idTool & "', '" & idTest & "', '" & idScenario & "', '" &
                              dataTest & "', '" & StepsTest & "', '" & ImportedData & "', '" & Attachments & "')"
                pc_db.ExecuteCommandSQL(sql)

                defect = defect.Replace("@package", demand)
                defect = defect.Replace("@machine", net.GetHostName)
                defect = defect.Replace("@idDefect", "new")
                defect = defect.Replace("@summary", summary)
                defect = defect.Replace("@status", "New")
                defect = defect.Replace("@severity", severity)
                defect = defect.Replace("@priority", priority)
                defect = defect.Replace("@detectedBy", detectedBy)
                defect = defect.Replace("@assginTo", AssignTo)
                defect = defect.Replace("@solutionType", "")
                defect = defect.Replace("@detectedRelease", detectedRelease)
                defect = defect.Replace("@detectedCycle", detectedCycle)
                defect = defect.Replace("@targetRelease", detectedRelease)
                defect = defect.Replace("@targetCycle", detectedCycle)
                defect = defect.Replace("@incidentType", typeIncident)
                defect = defect.Replace("@area", area)
                defect = defect.Replace("@application", Application)
                defect = defect.Replace("@functionality", Functionality)
                defect = defect.Replace("@description", description)
                defect = defect.Replace("@solution", solution)
                defect = defect.Replace("@Comments", comments)
                defect = defect.Replace("@creationDate", Now)

                email.to = emailSend
                email.subject = summary
                email.body = defect
                email.priority = System.Net.Mail.MailPriority.High
                email.isBodyHTML = True
                email.Attachments = Attachments
                LogExecution("Send email defect to " & emailSend, StateLog.Start)
                'email.Send()
                LogExecution("Send email defect to " & emailSend, StateLog.Finish)
                LogExecution("CreateDefect", StateLog.Finish)
            Catch ex As Exception
                LogExecution("CreateDefect", StateLog.Exception)
            End Try

        End Sub
        Public Function TestLog(stepName As String, ExpectedResults As String, ActualResults As String, Optional Check As String = Nothing) As Boolean
            Dim statusTest As typelog
            If String.IsNullOrEmpty(Check) Then statusTest = typelog.Passed Else statusTest = typelog.Done
            TestLog(stepName, ExpectedResults, ActualResults, statusTest, Check)
        End Function
        Public Function TestLog(stepName As String, ExpectedResults As String, ActualResults As String, TypeLog As typelog, Optional Check As String = Nothing) As Boolean

            Dim ActualResultsLocal As String = Nothing
            Dim errors As String = Nothing
            Dim util = New Utilities
            Dim nameIMG As String = Format(Now, "yyyyMMddHHmmsss") & Now.Millisecond & ".png"
            Dim dir = New Directorys
            Dim htmlSource As String = Nothing
            Dim lenIDScen As Integer = Len(CStr(p_IDScenario))
            Dim lenIDTest As Integer = Len(CStr(p_OrdemTest))
            Dim pathTempTest As String
            Dim flag As Boolean = False
            'Dim    pathlocal As String = System.IO.Directory.GetCurrentDirectory
            Dim pathTemp As String = "\SCEN"
            pathTemp = pathTemp & Strings.StrDup(4 - lenIDScen, "0") & p_IDScenario
            pathTempTest = Strings.StrDup(3 - lenIDTest, "0") & p_OrdemTest

            Dim pathAccess As String() = xml.Read("pathAccess", "").Split("\")
            Dim LocalPath As String = "C:\LeanTestAutomation\Scripts\Execution\" & pathAccess(4) & "\" & p_pathLab & "\" & pathTemp & "\" & p_TableTest & "\"
            LocalPath = Replace(LocalPath, "\\", "\")

            If Not System.IO.Directory.Exists(LocalPath) Then
                If Not dir.CreateDiretory(LocalPath) Then MsgBox("TestLog error in create file path")
            End If

            Try
                LogExecution("TestLog Step: " & stepName & " - Expected Results: " & ExpectedResults & " - Check : " & Check, StateLog.Start)
                If Not String.IsNullOrEmpty(Check) Then
                    ActualResultsLocal = CheckPointTest(Check, ExpectedResults, , False) 'nao gera o log e retorna o status para que seja tratado aqui.
                    Select Case ActualResultsLocal
                        Case "Passed"
                            ActualResults = "Step executado com sucesso"
                            p_StatusTest = "Passed"
                            'LogExecution("TestLog.WaitExist: True - Actual Results: " & ActualResults, StateLog.Information)
                        Case "Failed"
                            ActualResults = "Passo executado com falha em " & p_TableTest
                            errors = "Check executado com falha! Check: " & Check
                            p_StatusTest = "Failed"

                            ActualResults = "Passo executado com falha em " & p_TableTest & " erro: " & p_errorDescription
                            CreateDefect(stepName, ExpectedResults, Replace(ActualResults, "'", ""))
                            'LogExecution("TestLog.WaitExist: False - Actual Results: " & ActualResults & "Element not found", StateLog.Information)
                    End Select
                Else
                    Select Case TypeLog
                        Case LibGlobal.typelog.Passed
                            ActualResults = "Step executado com sucesso"
                            p_StatusTest = "Passed"
                        Case LibGlobal.typelog.Failed
                            ActualResults = "Passo executado com falha em " & p_TableTest & " erro: " & p_errorDescription
                            p_StatusTest = "Failed"
                            Call CreateDefect(stepName, ExpectedResults, Replace(ActualResults, "'", ""))
                        Case LibGlobal.typelog.Done
                            ActualResults = "Passo executado com atenção"
                            p_StatusTest = "Done"
                        Case LibGlobal.typelog.Warning
                            ActualResults = "Passo executado com atenção"
                            p_StatusTest = "Warning"
                        Case LibGlobal.typelog.Blocked
                            p_StatusTest = "Blocked"
                        Case LibGlobal.typelog.NA
                            p_StatusTest = "N/A"
                        Case LibGlobal.typelog.Information
                            p_StatusTest = "Information"
                        Case Else
                            p_StatusTest = "Done"
                    End Select
                End If
                LogExecution("TestLog: stepName=" & stepName & " - Status:" & p_StatusTest, StateLog.Information)
                'gera print screen pelo framework ou usando a API do browsersake remoto
                'verifica se existe mensagem aberta
               
                Select Case p_ToolName
                    Case ActiveTool.SeleniumWDBrowserStack
                        'htmlSource = Test.GetHtmlSource()
                        'If String.IsNullOrEmpty(htmlSource) Then
                        '    Test.GetTextPopup(True)
                        'End If
                        'Dim screenshot As OpenQA.Selenium.Screenshot = objSeleniumWDBS.GetScreenshot()
                        'flag = True
                    Case ActiveTool.SeleniumWD
                        htmlSource = Test.GetHtmlSource()
                        If String.IsNullOrEmpty(htmlSource) Then
                            Test.GetTextPopup(True)
                        End If
                        LocalPath = pc_te.TestLog(stepName, ExpectedResults, ActualResults, p_StatusTest, p_IDScenario, p_IDTest, LocalPath, p_IDRun)
                        'flag = True
                        flag = False
                    Case ActiveTool.Appium
                        Dim screenshot As OpenQA.Selenium.Screenshot = DirectCast(objAppium, ITakesScreenshot).GetScreenshot()

                        screenshot.SaveAsFile(LocalPath & "\" & nameIMG, System.Drawing.Imaging.ImageFormat.Png)
                        LocalPath = LocalPath & "\" & nameIMG
                        flag = True
                    Case Else
                        flag = True
                        '    LocalPath = pc_te.TestLog("[" & p_TestName & "] - " & stepName, ExpectedResults, ActualResults, p_StatusTest, p_IDScenario, p_IDTest, LocalPath, p_IDRun)
                        LocalPath = pc_te.TestLog(stepName, ExpectedResults, ActualResults, p_StatusTest, p_IDScenario, p_IDTest, LocalPath, p_IDRun)
                End Select
                If flag Then
                    Dim Sql = "INSERT INTO LeanTestLogs (StepName, ExpectedResults, ActualResults, StatusExecution, PathEvidence, IDScenario, IDTest, IDRun, CreationDate) values (" &
                      "'" & stepName & "', '" & Replace(ExpectedResults, "'", "") & "', '" & Replace(ActualResults, "'", "") & "', '" & p_StatusTest & "', '" & LocalPath & "', " & p_IDScenario & ", '" & p_IDTest & "', '" & p_IDRun & "', '" & Now & "')"
                    pc_db.ExecuteCommandSQL(Sql)
                End If

                If p_StatusScenario <> "Failed" Then
                    p_StatusScenario = p_StatusTest
                End If
                If p_StatusTest = "Failed" Then
                    p_StatusScenario = "Failed"
                End If
                'publish QC
                If p_PublishQC Then
                    Dim fullpath = Path.GetDirectoryName(LocalPath)
                    Dim name = Path.GetFileName(LocalPath)
                    If name <> Nothing Then
                        p_qc.AddAttach(p_Run, fullpath, name, ExpectedResults)
                        p_qc.AddStepRun(p_Run, p_StatusTest, stepName, ExpectedResults, ExpectedResults, ActualResults)
                    End If
                End If
                LogExecution("TestLog", StateLog.Finish)
                Return True
            Catch ex As Exception
                LogExecution("TestLog: " & ex.Message, StateLog.Exception, False)
                Return False
            End Try
        End Function

        Public Sub SetValueOutput(field As String, value As String)
            LogExecution("SetValueOutput: field=" & field & " value=" & value, StateLog.Start)
            Try
                pc_db.SetOutput(p_IDTest, p_TableTest, Replace(field, "'", ""), value)
                LogExecution("SetValueOutput", StateLog.Finish)
            Catch ex As Exception
                LogExecution("SetValueOutput: " & ex.Message, StateLog.Exception)
            End Try
        End Sub
        Public Function GetValueOutput(idTest As String, tableTest As String, field As String)
            Dim values As String = Nothing
            LogExecution("GetValueOutput: " & tableTest & "." & field & "=" & values, StateLog.Start)
            Try
                values = pc_db.GetOutput(idTest, tableTest, field)
                LogExecution("GetValueOutput", StateLog.Finish)
                Return values
            Catch ex As Exception
                LogExecution("GetValue: " & ex.Message, StateLog.Exception)
                Return Nothing
            End Try
        End Function
        Function GetHtmlSource() As String
            Try
                Select Case p_ToolName
                    'Case ActiveTool.SeleniumRC
                    '    Return librarySeleniumRCInteraction.GetHtmlSource()
                    'Case ActiveTool.SeleniumWDBrowserStack
                    '    Return librarySeleniumWDBSInteraction.GetHtmlSource()
                    Case ActiveTool.SeleniumWD
                        Return librarySeleniumWDInteraction.GetHtmlSource()
                    Case ActiveTool.Appium
                        Return libraryAppiumInteration.GetHtmlSource()
                    Case Else
                        Return Nothing
                End Select
            Catch ex As Exception
                HandlerError("libGlobal.GetHtmlSource: " & ex.StackTrace & " -" & ex.Message)
                Return Nothing
            End Try
        End Function


        Function GetTextPopup(Optional click As Boolean = True, Optional frameName As String = Nothing) As String
            Dim text As String = Nothing
            LogExecution("GetTextPopup", StateLog.Start)
            Try
                Select Case p_ToolName
                    Case ActiveTool.SeleniumRC
                        'Return librarySeleniumRCInteraction.GetHtmlSource()
                    Case ActiveTool.SeleniumWDBrowserStack
                        'Return librarySeleniumWDBSInteraction.GetHtmlSource
                    Case ActiveTool.SeleniumWD
                        text = librarySeleniumWDInteraction.GetTextPopup(click, frameName)
                    Case ActiveTool.Appium
                        text = libraryAppiumInteration.GetTextPopup(click)
                    Case Else
                        LogExecution("GetTextPopup text:", StateLog.Finish)
                        Return Nothing
                End Select
                LogExecution("GetTextPopup text: " & text, StateLog.Finish)
                Return text
            Catch ex As Exception
                LogExecution("GetTextPopup: " & ex.Message, StateLog.Exception)
                Return Nothing
            End Try
        End Function
        'verificar shared
        Function HandlerMessage(Optional frameName As String = Nothing) As String

            Select Case p_ToolName
                Case ActiveTool.SeleniumWD
                    Dim htmlSource As String = GetHtmlSource()
                    Dim MessageCapured As String = Nothing

                    If String.IsNullOrEmpty(htmlSource) Then MessageCapured = GetTextPopup(True, frameName) 'caso nao consiga recuperar o html, porque tem popup
                    If String.IsNullOrEmpty(MessageCapured) Then Return Nothing

                    Try
                        Dim ReturnAction As String = pc_db.VerifyMessage(MessageCapured)
                        'action retuned by database
                        Select Case ReturnAction
                            Case 0 'mensagem sucesso
                                Test.TestLog("Mensagem de Sistema", "Sistema apresenta mensagem", "Sistema apresentou a mensagem: " & MessageCapured, typelog.Passed)
                            Case 1 'falha
                                Test.TestLog("Mensagem de Sistema", "Mensagem inesperada", "Sistema apresentou a mensagem: " & MessageCapured, typelog.Failed)
                                'Throw New Exception("Sistema apresentou a mensagem: " & MessageCapured)
                            Case Else
                                Return MessageCapured
                        End Select
                        Return MessageCapured
                    Catch ex As Exception
                        Throw New Exception(ex.Message)
                    End Try
                Case ActiveTool.Appium
                    Dim title As String = Nothing
                    Dim message As String = Nothing
                    Dim element As String = Nothing

                    If Not Exist("br.com.vivo:id/title", 500, typeIdentification.id) Then
                        'If Not Exist("com.android.packageinstaller:id/permission_message", 500, typeIdentification.id) Then
                        '    Return Nothing
                        'Else
                        '    title = GetText("com.android.packageinstaller:id/permission_message", typeIdentification.id)
                        '    message = GetText("com.android.packageinstaller:id/permission_message", typeIdentification.id)
                        '    element = "com.android.packageinstaller:id/permission_allow_button"
                        'End If
                    Else
                        'title = GetText("br.com.vivo:id/title", typeIdentification.id)
                        'message = GetText("br.com.vivo:id/message", typeIdentification.id)
                        'element = "br.com.vivo:id/positive_button"
                    End If

                    If String.IsNullOrEmpty(message) Then Return Nothing

                    Try
                        Dim ReturnAction As String = pc_db.VerifyMessage(message)
                        'action retuned by database
                        Select Case ReturnAction
                            Case 0 'mensagem sucesso
                                Test.TestLog("Mensagem de aplicativo", "Aplicativo apresenta mensagem", "Aplicativo apresentou a mensagem: " & message, typelog.Passed)
                            Case 1 'falha
                                Test.TestLog("Mensagem de aplicativo", "Mensagem inesperada", "Aplicativo apresentou a mensagem: " & message, typelog.Failed)
                            Case Else
                                Test.Click(element, True, typeIdentification.id)
                                Return message
                        End Select
                        Test.Click(element, True, typeIdentification.id)
                        Return message
                    Catch ex As Exception
                        Throw New Exception(ex.Message)
                    End Try
            End Select
            Return Nothing
        End Function
        Public Function HandlerMessageApp(Optional positiveButton As Boolean = True, Optional negativeButton As Boolean = False)
            'br.com.vivo:id/title
            'br.com.vivo:id/message
            'br.com.vivo:id/negative_button
            'br.com.vivo:id/positive_button

            Dim title As String = GetValue("br.com.vivo:id/title", typeIdentification.id)
            Dim message As String = GetValue("br.com.vivo:id/message", typeIdentification.id)

            Test.Click("br.com.vivo:id/positive_button", positiveButton, typeIdentification.id)
            Test.Click("br.com.vivo:id/negative_button", negativeButton, typeIdentification.id)
            Return Nothing
        End Function
        Shared Function CreateStructureQC()
            Try
                p_qc.ConectQC("https://10.244.194.91/qcbin/", "QC_BRAZIL", "QA", "souzasl", "senhaIncorret@")
                Dim oTest = p_qc.TestPlanAdd(p_pathTest, "Test", "")
                Dim oCreatedTestSet = p_qc.TestLabTestSetAdd(p_pathLab, p_ScenarioName)
                If (oTest Is Nothing) Then
                    MsgBox("Error creating lab, test not found:")
                Else
                    If String.IsNullOrEmpty(p_IDTestInstance) Then
                        p_IDTestInstance = 0
                        p_IDTestInstance = p_qc.TestLabAddTestToTestSet(oCreatedTestSet, oTest, p_IDTestInstance)
                        pc_db.SetOutput(p_IDTest, p_TableTest, "IDTool", p_IDTestInstance)
                    End If
                    p_Run = p_qc.AddRunToInstanceTest(oCreatedTestSet, p_IDTestInstance, "Passed")
                End If
            Catch ex As Exception
                HandlerError("LibraryGlobal.libGlobal.ConectQC: " & ex.StackTrace & " - " & ex.Message)
            End Try
            Return Nothing
        End Function
        'CHECKPOINT ##########################################################################################################
        Public Function CheckPointElement(element As String, Optional typeIdentification As typeIdentification = LibGlobal.typeIdentification.xpath)
            Try
                Return Test.WaitExist(element, typeIdentification)
            Catch ex As Exception
                Return False
            End Try
        End Function
        Public Function CheckPointElementByValue(element As String, value As String, Optional typeIdentification As typeIdentification = typeIdentification.xpath)
            Dim timeout As Integer = p_timeout
            Try
                Do
                    If Test.Exist(element, 200, typeIdentification) Then
                        If value = Test.GetText(element, typeIdentification) Then
                            Return True
                            Exit Do
                        End If
                    End If
                    timeout -= 1
                    Test.Wait(1000)
                    If timeout = 0 Then
                        Return True
                        Exit Do
                    End If
                Loop While Test.WaitExist(element, typeIdentification)
                Return False
            Catch ex As Exception
                Return False
            End Try
        End Function
        Public Function CheckPointTest(checkpoint As String, expectedResults As String, Optional field_out As String = Nothing, Optional IsTestLog As Boolean = True) As String
            LogExecution("CheckPointTest: checkpoint=" & checkpoint, StateLog.Start)
            LogExecution("CheckPointTest: expectedResults=" & expectedResults, StateLog.Information)

            Dim output1 As String = Nothing
            Dim output2 As String = Nothing

            'caso o checkpoint seja vazio
            If String.IsNullOrEmpty(checkpoint) Then
                output1 = HandlerMessage()
                Return False
            End If

            'caso apareça um alert inesperado
            Test.Wait(1000)
            Dim html1 As String = Test.GetHtmlSource
            Dim msgbox As String = Nothing
            If String.IsNullOrEmpty(html1) Then
                msgbox = HandlerMessage()
                If msgbox <> "undefined" Then output1 = msgbox
            End If

            Dim TypeVerify As String = Nothing
            Dim element As String = Nothing
            Dim valueTest As String = Nothing

            Dim vals() As String = Split(checkpoint, ";")
            Dim statusTest As String
            Dim IsCapture As Boolean = True
            statusTest = "Passed"
            Try
                Try
                    element = vals(0)
                    valueTest = vals(1)
                    TypeVerify = vals(2)
                Catch ex As Exception
                End Try
                'condições de uso
                Select Case UCase(vals(2))
                    Case UCase("ExistMSG") 'valida a uma mensagem esperada
                        ';mensagem esperada;ExistMSG
                        Test.Wait(6000)
                        output1 = HandlerMessage()
                        If valueTest <> output1 Then statusTest = "Failed"

                    Case UCase("ValidateValueInElement")  'valida o valor captura em um elemento com o valor esperado
                        '"element;"valoresperado";ValidateValueInElement"
                        Test.WaitExist(element) 'ESPERA O ELEMENTO
                        Try
                            output1 = Test.GetText(element)
                            output2 = Test.GetValue(element)
                        Catch ex As Exception
                            output1 = Test.GetValue(element)
                            output2 = Test.GetValue(element)
                        End Try
                        output1 = IIf(Not String.IsNullOrEmpty(output1), output1, output2)
                        If output1 <> valueTest Then statusTest = "Failed"

                    Case UCase("ValidateValueInFieldWithElement") ' valida se um valor existente e armazenado em banco é igual a um capturado na tela em execução
                        'element;vValor;ValidateValueInFieldWithElement
                        valueTest = Test.GetValueOutput(p_IDTest, p_TableTest, valueTest)
                        Test.WaitExist(element) 'ESPERA O ELEMENTO
                        Try
                            output1 = Test.GetText(element)
                            output2 = Test.GetValue(element)
                        Catch ex As Exception
                            output1 = Test.GetValue(element)
                            output2 = Test.GetValue(element)
                        End Try
                        output1 = IIf(Not String.IsNullOrEmpty(output1), output1, output2)
                        If output1 <> valueTest Then statusTest = "Failed"

                    Case UCase("ValidateValueInFieldWithElementIsDifferent") ' valida se um valor existente e armazenado em banco é igual a um capturado na tela em execução
                        'element;vValor;ValidateValueInFieldWithElementIsDifferent
                        valueTest = Test.GetValueOutput(p_IDTest, p_TableTest, valueTest)
                        Test.WaitExist(element) 'ESPERA O ELEMENTO e valida o o valor captura é diferente do valor armazenado na base
                        Try
                            output1 = Test.GetText(element)
                            output2 = Test.GetValue(element)
                        Catch ex As Exception
                            output1 = Test.GetValue(element)
                            output2 = Test.GetValue(element)
                        End Try
                        output1 = IIf(Not String.IsNullOrEmpty(output1), output1, output2)
                        If output1 = valueTest Then statusTest = "Failed"

                    Case UCase("FieldIsNothing") 'VALIDA SE O CAMPO É VAZIO
                        '"vField;;FieldIsNothing" 
                        valueTest = Test.GetValueOutput(p_IDTest, p_TableTest, element)
                        If String.IsNullOrEmpty(valueTest) Then statusTest = "Passed" Else statusTest = "Failed"

                    Case UCase("FieldNotIsNothing") 'VALIDA SE O CAMPO NAO É VAZIO, pode ser usado com a captura de um campo na tela, banco de dados, elemento xml, etc
                        '"vField;;FieldNotIsNothing"
                        valueTest = Test.GetValueOutput(p_IDTest, p_TableTest, element)
                        If Not String.IsNullOrEmpty(valueTest) Then statusTest = "Passed" Else statusTest = "Failed"

                    Case UCase("FieldWithField") 'VALIDA SE OS VALORES DOS CAMPOS SAÕ IGUAIS. usado para validar dois valores, onde um é previanmente armazenado ou transferido de outra tabela e o outro valor capturado em tempo de execução.
                        'vField1;vField2;FieldWithField  
                        Dim valueTest1 As String = Test.GetValueOutput(p_IDTest, p_TableTest, element)
                        Dim valueTest2 As String = Test.GetValueOutput(p_IDTest, p_TableTest, valueTest)
                        If valueTest1 = valueTest2 Then statusTest = "Passed" Else statusTest = "Failed"

                    Case UCase("ValueWithField") 'VALIDA SE OS VALORES DOS CAMPOS SAÕ IGUAIS. usado para validar dois valores, onde um é previanmente armazenado ou transferido de outra tabela e o outro valor capturado em tempo de execução.
                        Dim util = New Utilities
                        'value;vField;ValueWithField  
                        Dim valueTest1 As String = element
                        Dim valueTest2 As String = Test.GetValueOutput(p_IDTest, p_TableTest, valueTest)
                        If util.ConvertToAlphanumeric(valueTest1, ) = util.ConvertToAlphanumeric(valueTest2) Then statusTest = "Passed" Else statusTest = "Failed"

                    Case UCase("FindText") 'VALIDA SE O TEXTO EXISTE NA TELA PO EXEMPLO "Numero do Pedido Gerado"
                        '";hello natalia;FindText" 
                        If String.IsNullOrEmpty(element) And Mid(valueTest, 1, 1) = "" Then valueTest = Test.GetValueOutput(p_IDTest, p_TableTest, Right(valueTest, Len(valueTest) - 1)) 'caso vValor
                        For i = 0 To p_timeout
                            Dim html As String = Test.GetHtmlSource
                            If InStr(UCase(html), UCase(valueTest)) Then
                                statusTest = "Passed"
                                Exit For
                            Else
                                statusTest = "Failed"
                            End If
                            IsCapture = False
                            Test.Wait(1000)
                        Next
                        '************************************************************************************************************
                        'AJUSTAR PARA VERIFICAR SE ELEMENTO EXISTE POR ID, NAME, LINK, ETC.

                    Case UCase("ElementExist") 'VALIDA SE UM ELEMENTO DA TELA EXISTE EXEMPLO IDName, //div[3]@ID='nome', etc
                        '"element;;ElementExist"
                        If Test.Exist(element, 500) Then statusTest = "Passed" Else statusTest = "Failed"
                    Case UCase("ElementNotExist")  'VALIDA SE UM ELEMENTO NAO EXISTE NA TELA
                        '"element;;ElementNotExist"
                        If Not Test.Exist(element, 500) Then statusTest = "Passed" Else statusTest = "Failed"
                End Select

                'cria o log do test*************************************************************************************************************************
                If IsTestLog Then 'crir o log quando este metodo nao for chamado pelo Testlog. neste caso o metodo chamador realiza  log
                    If statusTest = "Passed" Then
                        Test.TestLog("CheckPoint", "Resultado esperado: " & expectedResults, "Passo executado com sucesso!", typelog.Passed)
                    Else
                        Test.TestLog("Checkpoint", "Resultado esperado: " & expectedResults, "Passo executado com falha! Menssagem: " & output1, typelog.Failed)
                    End If
                End If
                LogExecution("CheckPoint", StateLog.Finish)
                Return statusTest
            Catch ex As Exception
                LogExecution("Checkpoint - Resultado esperado: " & expectedResults & "Passo executado com falha! Menssagem: " & ex.Message, StateLog.Exception)
                Throw New Exception(ex.Message)
            End Try
        End Function
        'END CHECKPOINT#######################################################################################################
        Sub EndExecution()
            Dim email = New Email
            Dim db = New DBAdapter
            Dim dir = New Directorys
            Dim sql As String = "select * from LeanTestDemands"

            Dim flagGenerateReport As String = db.GetParameter("GenerateReport")
            Dim flagSendEmail As String = db.GetParameter("sendByEmail")
            Dim flagPublishResults As String = db.GetParameter("publishAllResultsInPlenus")

            Dim pathAccess As String() = xml.Read("pathAccess", "").Split("\")
            Dim PublishPathResults As String = db.GetParameter("PathToPostResult") '\\server\PlenusExecution\

            PublishPathResults += "\\Scripts\Execution\" & pathAccess(4) & "\"
            PublishPathResults = PublishPathResults.Replace("\\", "\")

            Dim LocalPath As String = "C:\LeanTestAutomation\Scripts\Execution\" & pathAccess(4)
            Dim LocalPathBackUpExecution As String = "C:\LeanTestAutomation\Scripts\ExecutionBackUpExecution\" & pathAccess(4)
            LocalPath = Replace(LocalPath, "\\", "\")

            'load all values by LeanTestDemands
            Try
                LogExecution("EndExecution: Scenario" & p_ScenarioName & " - Status execution scenario= " & p_StatusScenario, StateLog.Start)
                EndScenario()

                Dim sqlCount As String = "Select count(*) as countTest From LeanTestScenarios Where StatusExecution  = 'No Run'"
                sqlCount = pc_db.GetValue(sqlCount, "countTest", True)
                If CBool(flagGenerateReport) Then 'generator reporte
                    If sqlCount = "0" Then
                        sqlCount = pc_rp.ExportLogsSceanariosToWord(p_pathStatusReportAllScenarios, p_pathEvidence & "\" & p_pathLab, "0", "0") 'retorna o caminho
                        If CBool(flagSendEmail) Then 'send email
                            email.to = db.GetValue(sql, "EmailsReporte")
                            email.body = db.GetValue(sql, "Description")
                            email.subject = "Status Report of package " & db.GetValue(sql, "Demand")
                            email.isBodyHTML = True
                            If Not String.IsNullOrEmpty(sqlCount) Then email.priority = Net.Mail.MailPriority.Normal
                            email.Attachments = sqlCount
                            email.Send()
                        End If
                    End If
                End If

                If CBool(flagPublishResults) Then ' if publish results is true
                    If Not String.IsNullOrEmpty(PublishPathResults) Then ' if path exists
                        'evidences
                        dir.CopyDiretory(LocalPath & "\" & p_pathLab, PublishPathResults & "\" & p_pathLab, True)

                        dir.DeleteDiretory(LocalPath & "\" & p_pathLab, True)
                        'dataTest
                        dir.CopyDiretory(LocalPath & "\" & "\DataTest.mdb", PublishPathResults & "\DataTest.mdb", True)
                        db.ClearResultsExecution(0)
                    End If
                End If

                'backup
                If Not String.IsNullOrEmpty(p_pathLab) Then
                    LocalPath += "\" & p_pathLab & "\"
                    LocalPath = LocalPath.Replace("\\", "\")
                    LocalPathBackUpExecution += "\" & p_pathLab & "\"
                    LocalPathBackUpExecution = LocalPathBackUpExecution.Replace("\\", "\")

                    dir.CopyDiretory(LocalPath, LocalPathBackUpExecution, False)
                    dir.CopyFiles(Replace(LocalPath, p_pathLab, "") & "DataTest.mdb", LocalPathBackUpExecution & "DataTest.mdb", True)
                End If
            Catch ex As Exception
            End Try
            LogExecution("EndExecution", StateLog.Finish)
            pc_db.DisConnectDB()
        End Sub
        Sub DeleteAllCookies()
            Try
                objSeleniumWD.Manage.Cookies.DeleteAllCookies()
            Catch ex As Exception

            End Try
        End Sub
        Function IsDisponible(url As String, element As String) As Boolean
            Try
                p_pathUrlApp = url
                Test.Open(p_pathUrlApp)
                Return WaitExist(element)
            Catch ex As Exception
                Return False
            End Try
        End Function
        Sub ClipBoardSetText(text As String)
            Try
                Clipboard.Clear()
                Clipboard.SetText(text)
            Catch ex As Exception
            End Try
        End Sub
        Sub ClipBoardClear()
            Try
                Clipboard.Clear()
            Catch ex As Exception
            End Try
        End Sub
        Function ClipBoardGetText()
            Try
                If Clipboard.ContainsText Then
                    Return Clipboard.GetText
                Else
                    Return Nothing
                End If
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
        Function ClipBoarContainsText()
            Try
                Return Clipboard.ContainsText
            Catch ex As Exception
                Return Nothing
            End Try
        End Function


        Sub TeardownTest()
            ' Dim sh = New Shell
            LogExecution("TeardownTest", StateLog.Start)
            Try
                Select Case p_ToolName
                    'Case ActiveTool.SeleniumRC
                    '    librarySeleniumRCInteraction.TeardownTest()
                    'Case ActiveTool.SeleniumWDBrowserStack
                    '    librarySeleniumWDBSInteraction.TeardownTest()
                    Case ActiveTool.SeleniumWD
                        librarySeleniumWDInteraction.TeardownTest()
                    Case ActiveTool.Appium
                        libraryAppiumInteration.TeardownTest()
                    Case Else
                        LogExecution("GetTextPopup text:", StateLog.Finish)
                End Select
                LogExecution("TeardownTest", StateLog.Finish)

            Catch ex As Exception
                LogExecution("GetTextPopup: " & ex.Message, StateLog.Exception)
            End Try
            ' sh.UndoMinimizeALL()
        End Sub
        Function CopyDiretory(pathSource As String, pathTarget As String, Optional overwriteWithBackup As Boolean = False) As Boolean
            Try
                Dim dir = New Directorys
                Return dir.CopyDiretory(pathSource, pathTarget, overwriteWithBackup)
            Catch ex As Exception
                Return False
            End Try
        End Function
        Function CopyFiles(sourceFileName As String, destFileName As String, Optional overWrite As Boolean = True) As Boolean
            Try
                Dim dir = New Directorys
                Return dir.CopyFiles(sourceFileName, destFileName, overWrite)
            Catch ex As Exception
                Return False
            End Try
        End Function
        Function DeleteDiretory(Path As String, Optional Recursive As Boolean = True) As Boolean
            Try
                Dim dir = New Directorys
                Return dir.DeleteDiretory(Path, Recursive)
            Catch ex As Exception
                Return False
            End Try
        End Function
        Function DeleteFile(Path As String) As Boolean
            Try
                Dim dir = New Directorys
                Return dir.DeleteFile(Path)
            Catch ex As Exception
                Return False
            End Try
        End Function
        Function Exists(Path As String) As Boolean
            Try
                Dim dir = New Directorys
                Return dir.Exists(Path)
            Catch ex As Exception
                Return False
            End Try
        End Function
        Function GetDirectoryName(fullPath As String) As Boolean
            Try
                Dim dir = New Directorys
                Return dir.GetDirectoryName(fullPath)
            Catch ex As Exception
                Return False
            End Try
        End Function
        Function GetExtension(Path As String) As Boolean
            Try
                Dim dir = New Directorys
                Return dir.GetExtension(Path)
            Catch ex As Exception
                Return False
            End Try
        End Function
        Function GetFileNameWithExtention(fullPath As String) As Boolean
            Try
                Dim dir = New Directorys
                Return dir.GetFileNameWithExtention(fullPath)
            Catch ex As Exception
                Return False
            End Try
        End Function
        Function Rename(pathFile As String, pathNewFile As String) As Boolean
            Try
                Dim dir = New Directorys
                dir.Rename(pathFile, pathNewFile)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        Function ConvertToAlphanumeric(Text, changeCharacterTo) As String
            Try
                Dim util = New Utilities
                Return util.ConvertToAlphanumeric(Text, changeCharacterTo)
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
        Function ConvertToNumeric(Text, changeCharacterTo) As String
            Try
                Dim util = New Utilities
                Return util.ConvertToNumeric(Text)
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
        Function RemoveNullCharacters(Text) As String
            Try
                Dim util = New Utilities
                Return util.RemoveNullCharacters(Text)
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
        Function ExistProcess(ProcessName As String) As String
            Try
                Dim util = New Utilities
                Return util.ExistProcess(ProcessName)
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
        Function ReadFile(PathFile As String) As String
            Try
                Dim file = New FileTexts
                Return file.ReadFile(PathFile)
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
        Function WriteAllText(pathSave As String, AllText As String) As String
            Try
                Dim file = New FileTexts
                file.WriteAllText(pathSave, AllText)
                Return True
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
        Function WriteFile(Text As String, pathSave As String) As String
            Try
                Dim file = New FileTexts
                file.WriteFile(Text, pathSave)
                Return True
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
        Function WriteFileAppendText(line As String, pathSave As String) As String
            Try
                Dim file = New FileTexts
                file.WriteFileAppendText(line, pathSave)
                Return True
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
        Function FileCompare(file1 As String, file2 As String, fileOutput As String, separator As String) As String
            Try
                Dim file = New FileTexts
                file.FileCompare(file1, file2, fileOutput, separator)
                Return True
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
        Function XMLRead(pathElement As String, pathXML As String)
            Try
                Dim xml = New XMLFile
                Return xml.Read(pathElement, pathXML)
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
        Function XMLSetValue(Node As String, pathXML As String)
            Try
                Dim xml = New XMLFile
                xml.setValue(Node, pathXML)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        Sub Flick(x As Integer, y As Integer)
            Try
                libraryAppiumInteration.Flick(x, y)
            Catch ex As Exception
            End Try
        End Sub
        Sub Down(x As Integer, y As Integer)
            Try
                libraryAppiumInteration.Down(x, y)
            Catch ex As Exception
            End Try
        End Sub
        Sub Up(x As Integer, y As Integer)
            Try
                libraryAppiumInteration.Up(x, y)
            Catch ex As Exception
            End Try
        End Sub
        Sub Move(x As Integer, y As Integer)
            Try
                libraryAppiumInteration.Move(x, y)
            Catch ex As Exception
            End Try
        End Sub
        Sub DragAndDrop(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)
            Try
                libraryAppiumInteration.DragAndDrop(x1, y1, x2, y2)
            Catch ex As Exception
            End Try
        End Sub

        Friend Sub [Set](v As String)
            Throw New NotImplementedException()
        End Sub
    End Class
End Namespace