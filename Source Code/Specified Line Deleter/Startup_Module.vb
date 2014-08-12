Imports System.Diagnostics
Imports System.IO

Module Startup_Module

    Private Sub Error_Handler(ByVal message As String)
        Try
            Dim Display_Message1 As New Display_Message(message)
            Display_Message1.ShowDialog()
        Catch ex As Exception
            MsgBox("An error occurred in Specified Line Deleter's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Sub main()
        Dim ApplicationName As String
        ApplicationName = "Specified Line Deleter"
        Try
            Dim aModuleName As String = Diagnostics.Process.GetCurrentProcess.MainModule.ModuleName
            Dim aProcName As String = System.IO.Path.GetFileNameWithoutExtension(aModuleName)

            If Process.GetProcessesByName(aProcName).Length > 3 Then
                Dim Display_Message1 As New Display_Message("Another Instance of " & ApplicationName & " is already running. Only three instances of " & ApplicationName & " is allowed to run at any time")
                Display_Message1.ShowDialog()
                Application.Exit()
            Else
                Dim Splash As New Splash_Screen
                Splash.Show()
                While Splash.dataloaded = False
                End While
                Dim monitor As New Main_Screen(Splash)
                monitor.ShowDialog()
                monitor.Visible = False
                While monitor.dataloaded = False
                End While
                monitor.Visible = True
                Application.Exit()
            End If
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while launching " & ApplicationName & ". The program will attempt to recover shortly.")
        End Try
    End Sub
End Module
