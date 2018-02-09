'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 05/02/2018 14:59:29 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation
'Environment:..... Automation Project
'Application...... Practico Live
'Functionality:... Fiscal
'Master Test:..... No Defined
'TableTest:....... test_Practico_Live_Fiscal
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_Practico_Live_Fiscal
    Public Class test_Practico_Live_Fiscal
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
								p_pathUrlApp = "C:\Users\primecontrol.girlan\AppData\Local\Apps\2.0\Q498DANR.8BA\N0XOMWML.692\live..tion_0000000000000000_0006.0005_91774542e426de1b\PracticoLive.exe"'xml.Read(""urlLocal"") 'Create urlLocal element in XML
								Test.OpenApp(p_pathUrlApp)
								Test.TestLog("Acessar o sistema Practico Live", "Sistema deve apresentar a funcionalidade Fiscal", "Sistema apresentou a funcionalidade Fiscal com sucesso", typelog.Passed)
							End If
							Test.OpenApp("C:\Users\primecontrol.girlan\AppData\Local\Apps\2.0\Q498DANR.8BA\N0XOMWML.692\live..tion_0000000000000000_0006.0005_91774542e426de1b\PracticoLive.exe") 'Open App 
                            Test.TestLog("Informar valores", "Sistema deve solicitar os valores conforme critérios de entrada", "Sistema permitiu a inclusão de valores com sucesso", typelog.Passed)
                            System.Threading.Thread.Sleep(5000)
                            If CBool(vbtnAcionar_Ok) Then
                                Test.TestLog("Clicar em Acionar Ok", "Clicar em Acionar Ok e verificar o resultado esperado", "Clique em Acionar Ok com sucesso", typelog.Passed)
                                Test.Click("686,434,73,33;C:\LeanTestAutomation\Scripts\lean.test.automation.framework\imgCaptured\img_20180205150230.bmp", vbtnAcionar_Ok, typeIdentification.leanTest) 'click Acionar Ok
                                Test.TestLog("Resultado após clique em Acionar Ok", "Resultado após clique em Acionar Ok", "Resultado verificado com sucesso", typelog.Passed)
                            End If
                            If CBool(vbtnok_Ambiente) Then
                                Test.TestLog("Clicar em ok Ambiente", "Clicar em ok Ambiente e verificar o resultado esperado", "Clique em ok Ambiente com sucesso", typelog.Passed)
                                Test.SendKey("{ENTER}") 'SendKey {ENTER}
                                Test.TestLog("Resultado após clique em ok Ambiente", "Resultado após clique em ok Ambiente", "Resultado verificado com sucesso", typelog.Passed)
                                System.Threading.Thread.Sleep(9000)
                                Test.Wait(9000)
                                'Cliando em Fiscal
                                Test.Click("296,31,56,23;C:\LeanTestAutomation\Scripts\lean.test.automation.framework\imgCaptured\img_20180205154407.bmp", vbtnok_Ambiente, typeIdentification.leanTest)
                                System.Threading.Thread.Sleep(5000)
                                'Clicando em Gerador de Arquivos Fiscais
                                Test.Click("283,58,88,68;C:\LeanTestAutomation\Scripts\lean.test.automation.framework\imgCaptured\img_20180205154547.bmp", vbtnok_Ambiente, typeIdentification.leanTest)
                                System.Threading.Thread.Sleep(5000)
                                'Clicando no seletor 
                                Test.Click("533,263,20,27;C:\LeanTestAutomation\Scripts\lean.test.automation.framework\imgCaptured\img_20180205165149.bmp", vbtnok_Ambiente, typeIdentification.leanTest)
                                System.Threading.Thread.Sleep(5000)
                                'Preenchendo o campo Tipo de arquivo
                                Test.SendKey("Nota fiscal Paulista")
                                Test.SendKey("{TAB}")
                                Test.SendKey("{RIGHT}")

                            End If

                            'Checkpoint
                            Test.CheckPointTest(p_CheckPoint1, p_ExpectedResult)
                            'end test                         
                            Test.EndTest(p_GenerateLogTest)
                            If p_IsLoop Then StartTest() Else p_CountTest = 0
                        Catch ex As Exception
							p_errorDescription = "Menssage error: " & ex.Message.ToString
							Test.TestLog("Passo executado", "Execução do passo com sucesso", "Passo executado com falha! Message: " & p_errorDescription, typelog.Failed)
							EndTestTable()
                       Test.EndTest(p_GenerateLogTest)
                            If p_IsLoop Then StartTest() Else p_CountTest = 0
                        End Try
                    Loop
                    EndTestTable()
                    Return True
                Else
                    Test.TestLog("Teste executado", "Teste executado com sucesso", "Teste executado com falha! StartTest = False", typelog.Failed)
                    EndTestTable()
                    Return False
                End If
            Catch ex As Exception
                p_errorDescription = "Menssage error: " & ex.Message.ToString
				HandlerError("test_Practico_Live_Fiscal.test_Practico_Live_Fiscal.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vSenha_Loguin, vbtnAcionar_Ok,vbtnok_Ambiente, vIsOpenSystem As String

        Private Function StartTest() As Boolean
            Dim strQueryOut1, strQueryOut2, strQueryOut3, strQueryOut4, strQueryOut5, strQueryOut6 as string
            Try
                p_CountTest = pc_db.OpenTestTable(p_TableTest, p_IDScenario) 'opening the test table containing all the test cases
                If p_CountTest <> 0 Then
                    p_IDScenario = pc_db.Fieldt("IDScenario") 'set IDSceario
                    p_IDTest = pc_db.Fieldt("IDTest") 'set IDTest
                    p_OrdemTest = pc_db.Fieldt("Ordem")
                    p_TestName = pc_db.Fieldt("TestName")
                    p_DescriptionTest = pc_db.Fieldt("Description")
                    p_IDRun = pc_db.Fieldt("IDRun")
                    p_ExpectedResult = pc_db.Fieldt("ExpectedResult")
                    p_IDTestInstance = pc_db.Fieldt("IDTool")
					p_CheckPoint1 = pc_db.Fieldt("CheckPoint1")

                    'Data Transfer Parameters
                    strQueryOut1 = pc_db.Fieldt("QueryInput1")
                    strQueryOut2 = pc_db.Fieldt("QueryInput2")
                    strQueryOut3 = pc_db.Fieldt("QueryInput3")
                    strQueryOut4 = pc_db.Fieldt("QueryInput4")
					strQueryOut5 = pc_db.Fieldt("QueryInput5")
					strQueryOut6 = pc_db.Fieldt("QueryInput6")
                   
                    'transfer values between tables
					If String.IsNullOrEmpty(strQueryOut1) Then pc_db.TransferDataInTablesArray(strQueryOut1, p_TableTest, p_IDScenario, p_IDTest)
                    If String.IsNullOrEmpty(strQueryOut2) Then pc_db.TransferDataInTablesArray(strQueryOut1, p_TableTest, p_IDScenario, p_IDTest)
                    If String.IsNullOrEmpty(strQueryOut3) Then pc_db.TransferDataInTablesArray(strQueryOut1, p_TableTest, p_IDScenario, p_IDTest)
                    If String.IsNullOrEmpty(strQueryOut4) Then pc_db.TransferDataInTablesArray(strQueryOut1, p_TableTest, p_IDScenario, p_IDTest)
                    If String.IsNullOrEmpty(strQueryOut5) Then pc_db.TransferDataInTablesArray(strQueryOut1, p_TableTest, p_IDScenario, p_IDTest)
                    If String.IsNullOrEmpty(strQueryOut6) Then pc_db.TransferDataInTablesArray(strQueryOut1, p_TableTest, p_IDScenario, p_IDTest)
					'
                    p_CountTest = pc_db.OpenTestTable(p_TableTest, p_IDScenario)
                    vSenha_Loguin = pc_db.Fieldt("vSenha_Loguin")
					vbtnAcionar_Ok = pc_db.Fieldt("vbtnAcionar_Ok")
					vbtnok_Ambiente = pc_db.Fieldt("vbtnok_Ambiente")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_Practico_Live_Fiscal.test_Practico_Live_Fiscal.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
