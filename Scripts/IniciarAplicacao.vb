Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal

Public Class IniciarAplicacao
    Public Shared Function IniciarAP() As Boolean
        Try
            Test.OpenApp("C:\Users\primecontrol.girlan\AppData\Local\Apps\2.0\WG5AHKQ8.B0G\DCEWTA4Q.63N\live..tion_0000000000000000_0006.0005_91774542e426de1b\PracticoLive.exe") 'Open App
            Test.TestLog("Acessar o sistema Practico Live", "Sistema deve apresentar a funcionalidade Fiscal", "Sistema apresentou a funcionalidade Fiscal com sucesso", typelog.Passed)
            Return True
        Catch ex As Exception
            Test.TestLog("Erro ao Abrir aplicação", "Erro ao Abrir aplicação", "Erro ao Abrir aplicação", typelog.Failed)
            Test.EndTest(p_GenerateLogTest)
            Return False
        End Try
    End Function

End Class
