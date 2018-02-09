Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Imports Lean.Test.Automation.Framework.LibraryGlobal.CallTest

Module ScriptMaster
    Sub Main()

        'Dim api = New test_api.test_api
        'api.Run()

        LogExecution("ScriptMaster.Main: Start execution scenarios in " & Now, StateLog.Start)
        Test.LoadVariables()
        Dim newScenario As Integer = 0
        Dim lastIDScenario As Integer = 0
        Dim IDScenarios As String = pc_db.GetParameter("IDScenarios")
        Dim lastScenarioName As String
        '*************************************************************************************************
        'start execution by tests*************************************************************************
        If p_ClearResult Then pc_db.ClearResultsExecution(IDScenarios) 'clear all results after start test
        Try
            If Not RunTestScenario(IDScenarios) Then Throw New Exception("Execution Main.RunTestScenario error")
            Do While p_CountScenario > 0
                lastIDScenario = p_IDScenario 'na primeira execução, nao deve carregar o ID do cenario
                lastScenarioName = p_ScenarioName
                p_TableTest = pc_db.FieldScen("TableTest")
                p_IDScenario = pc_db.FieldScen("IDScenario")
                p_IsLoop = CBool(pc_db.FieldScen("ISLoop"))
                p_IDScenarioTest = pc_db.FieldScen("ID")
                p_ScenarioName = pc_db.FieldScen("Scenario")
                p_pathTest = pc_db.FieldScen("PathTest")
                p_pathLab = pc_db.FieldScen("PathLab") & p_executionTime
                p_Release = pc_db.FieldScen("Release")
                p_Cycle = pc_db.FieldScen("Cycle")
                p_OrdemTest = pc_db.FieldScen("Ordem")

                If newScenario <> p_IDScenario Then
                    If newScenario <> 0 Then
                        EndScenario(lastIDScenario, lastScenarioName)
                        Test.TeardownTest()
                    End If
                    pc_db.StartExecutionScenario(p_IDScenario)
                    p_ResetApplication = True
                Else
                    p_ResetApplication = False
                End If
                newScenario = p_IDScenario
                '*************************************************************************************************
                'RUN TEST*****************************************************************************************

                If Not RunTest(p_TableTest) Then p_StatusScenario = "Failed"
                If p_StatusScenario = "Failed" Then
                    'update all test by scenario to Blocked
                    pc_db.SetBlockedTest(p_IDScenario)
                End If
                'FIM *********************************************************************************************
                '*************************************************************************************************
                If Not RunTestScenario(IDScenarios) Then Throw New Exception("Execution Main.RunTestScenario error")
            Loop
            Test.EndExecution()
        Catch ex As Exception
            HandlerError("ScriptMaster.Main: " & ex.InnerException.ToString & " - " & ex.Message)
            MsgBox(ex.Message)
        End Try
        Test.TeardownTest()
        LogExecution("ScriptMaster.Main: End execution scenarios in " & Now & "Status Scenario: " & p_StatusTest, StateLog.Finish)
    End Sub
End Module