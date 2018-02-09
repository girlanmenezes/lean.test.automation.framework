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

Imports Lean.Test.Automation.Framework.IniciarAplicacao



Namespace test_Practico_Live_Fiscal
    Public Class test_Practico_Live_Fiscal
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            'Dim Ini = New IniciarAplicacao()
                            'If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
                            'p_pathUrlApp = "C:\Users\primecontrol.girlan\AppData\Local\Apps\2.0\WG5AHKQ8.B0G\DCEWTA4Q.63N\live..tion_0000000000000000_0006.0005_91774542e426de1b\PracticoLive.exe" 'xml.Read(""urlLocal"") 'Create urlLocal element in XML
                            'Test.OpenApp(p_pathUrlApp)
                            'Test.TestLog("Acessar o sistema Practico Live", "Sistema deve apresentar a funcionalidade Fiscal", "Sistema apresentou a funcionalidade Fiscal com sucesso", typelog.Passed)
                            'End If
                            'Test.OpenApp("C:\Users\primecontrol.girlan\AppData\Local\Apps\2.0\WG5AHKQ8.B0G\DCEWTA4Q.63N\live..tion_0000000000000000_0006.0005_91774542e426de1b\PracticoLive.exe") 'Open App 
                            'Test.TestLog("Informar valores", "Sistema deve solicitar os valores conforme critérios de entrada", "Sistema permitiu a inclusão de valores com sucesso", typelog.Passed)
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
                            End If

                            If CBool(vbtn_Fiscal) Then 'Cliando em Fiscal
                                Test.WaitExist("296,31,56,23;C:\LeanTestAutomation\Scripts\lean.test.automation.framework\imgCaptured\img_20180205154407.bmp", typeIdentification.leanTest, "0,9", 20, True)
                                Test.TestLog("Clicar em Fiscal", "Clicar em Fiscal e verificar o resultado esperado", "Clique em Fiscal com sucesso", typelog.Passed)
                                Test.Click("296,31,56,23;C:\LeanTestAutomation\Scripts\lean.test.automation.framework\imgCaptured\img_20180205154407.bmp", vbtnok_Ambiente, typeIdentification.leanTest)
                                System.Threading.Thread.Sleep(5000)
                                Test.TestLog("Resultado após clique em Fiscal", "Resultado após clique em Fiscal", "Resultado verificado com sucesso", typelog.Passed)

                            End If

                            If CBool(vbtn_GerarFiscal) Then 'Clicando em Gerador de Arquivos Fiscais
                                Test.TestLog("Clicar em Gerador Arq Fiscais", "Clicar em Gerador Arq Fiscai e verificar o resultado esperado", "Clique em Gerador Arq Fiscai com sucesso", typelog.Passed)
                                Test.Click("283,58,88,68;C:\LeanTestAutomation\Scripts\lean.test.automation.framework\imgCaptured\img_20180205154547.bmp", vbtnok_Ambiente, typeIdentification.leanTest)
                                System.Threading.Thread.Sleep(5000)
                                Test.TestLog("Resultado após clique em Gerador Arq Fiscais", "Resultado após clique em Gerador Arq Fiscais", "Resultado verificado com sucesso", typelog.Passed)
                            End If

                            If CBool(vbtn_TipoArq) Then 'Clicando no seletor 
                                Test.TestLog("Clicar e preencher o Tipo Arquivo", "Clicar e preencher o Tipo Arquivo e verificar o resultado esperado", "Clique e preencher o Tipo Arquivo com sucesso", typelog.Passed)
                                Test.Click("533,263,20,27;C:\LeanTestAutomation\Scripts\lean.test.automation.framework\imgCaptured\img_20180205165149.bmp", vbtnok_Ambiente, typeIdentification.leanTest)
                                System.Threading.Thread.Sleep(5000)
                                Test.SendKey("Nota fiscal Paulista") 'Preenchendo o campo Tipo de arquivo
                                Test.SendKey("{ENTER}")
                                Test.Wait(3000)
                                Test.TestLog("Resultado após clique e preencher o Tipo Arquivo", "Resultado após clique e preencher o Tipo Arquivo", "Resultado verificado com sucesso", typelog.Passed)
                            End If

                            If CBool(vData) Then 'Data 
                                Test.TestLog("Clicar e preencher o campo data", "Clicar e preencher o campo data e verificar o resultado esperado", "Clique e preencher o campo data com sucesso", typelog.Passed)
                                Test.SendKey("{TAB}") 'mudando foco para o campo data
                                Test.SendKey(vData_Inicial) 'Preencher data inicial
                                Test.Wait("2000")
                                Test.SendKey("{TAB}") 'mudando foco para o campo data final
                                Test.SendKey(vData_Final) ' Preencher data final
                                Test.TestLog("Resultado após clique e preencher campo data", "Resultado após clique e preencher o campo data", "Resultado verificado com sucesso", typelog.Passed)

                            End If

                            If CBool(vbtn_UniNego) Then
                                Test.TestLog("Clicar e selecionar uma unidade de negocio", "Clicar e selecionar uma unidade de negocio e verificar o resultado esperado", "Clique e selecionar uma unidade de negocio com sucesso", typelog.Passed)
                                Test.Wait("2000")
                                Test.SendKey("{TAB}")
                                Test.SendKey("{ENTER}")
                                ' Test.WaitExist("331,158,219,26;C:\LeanTestAutomation\Scripts\lean.test.automation.framework\imgCaptured\img_20180208112841.bmp", valor, typeIdentification.leanTest)
                                Test.SendKey("{TAB}")
                                Test.SendKey(vbtn_unidadenegocio)
                                Test.Click("388,186,53,42;C:\LeanTestAutomation\Scripts\lean.test.automation.framework\imgCaptured\img_20180208113340.bmp", vbtn_UniNego, typeIdentification.leanTest)
                                Test.Click("348,342,25,21;C:\LeanTestAutomation\Scripts\lean.test.automation.framework\imgCaptured\img_20180208140803.bmp", vbtn_UniNego, typeIdentification.leanTest)
                                Test.Wait(2000)
                                Test.Click("521,188,55,39;C:\LeanTestAutomation\Scripts\lean.test.automation.framework\imgCaptured\img_20180208141017.bmp", vbtn_UniNego, typeIdentification.leanTest)
                                Test.TestLog("Resultado após clique e selecionar uma unidade de negocio", "Resultado após clique e selecionar uma unidade de negocio", "Resultado verificado com sucesso", typelog.Passed)
                            End If

                            If CBool(vGerar_Arquivo) Then
                                Test.TestLog("Clicar em Gerar Arquivo", "Clicar Gerar Arquivo e verificar o resultado esperado", "Clique em Gerar Arquivo com sucesso", typelog.Passed)
                                Test.Wait(3000)
                                Test.Click("251,180,98,42;C:\LeanTestAutomation\Scripts\lean.test.automation.framework\imgCaptured\img_20180208141452.bmp", vGerar_Arquivo, typeIdentification.leanTest)
                                Test.TestLog("Resultado após clique em Gerar Arquivo", "Resultado após clique em Gerar Arquivo", "Resultado verificado com sucesso", typelog.Passed)
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
        Public Shared vSenha_Loguin, vbtnAcionar_Ok, vbtnok_Ambiente, vGerar_Arquivo, vData_Inicial, vbtn_UniNego, vData, vbtn_unidadenegocio, vData_Final, vbtn_Fiscal, vbtn_GerarFiscal, vbtn_TipoArq, vIsOpenSystem As String

        Private Function StartTest() As Boolean
            Dim strQueryOut1, strQueryOut2, strQueryOut3, strQueryOut4, strQueryOut5, strQueryOut6 As String
            IniciarAplicacao.IniciarAP()
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
                    vbtn_TipoArq = pc_db.Fieldt("vbtn_TipoArq")
                    vbtn_GerarFiscal = pc_db.Fieldt("vbtn_GerarFiscal")
                    vbtn_Fiscal = pc_db.Fieldt("vbtn_Fiscal")
                    vData_Final = pc_db.Fieldt("vData_Final")
                    vData_Inicial = pc_db.Fieldt("vData_Inicial")
                    vbtn_unidadenegocio = pc_db.Fieldt("vbtn_unidadenegocio")
                    vData = pc_db.Fieldt("vData")
                    vbtn_UniNego = pc_db.Fieldt("vbtn_UniNego")
                    vGerar_Arquivo = pc_db.Fieldt("vGerar_Arquivo")


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
