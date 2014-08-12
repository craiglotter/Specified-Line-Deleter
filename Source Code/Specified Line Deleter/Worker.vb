Imports System.IO
Imports System.Text


Public Class Worker

    Inherits System.ComponentModel.Component

    ' Declares the variables you will use to hold your thread objects.

    Public WorkerThread As System.Threading.Thread

    Public filequeue1 As String
    Private filecount As Long
    Private blankcount As Long
    Private fullcount As Long
    Public searchterm As String
    Private filereader As StreamReader
    Private filewriter As StreamWriter

    Public Event WorkerFileProcessing(ByVal filename As String, ByVal queue As Integer)
    Public Event WorkerStatusMessage(ByVal message As String, ByVal statustag As Integer)
    Public Event WorkerError(ByVal Message As Exception)
    Public Event WorkerFileCount(ByVal Result As Long, ByVal count As Integer)
    Public Event WorkerComplete(ByVal queue As Integer)




#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)

    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        
    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub Error_Handler(ByVal message As Exception)
        Try
            If (Not WorkerThread.ThreadState.ToString.IndexOf("Aborted") > -1) And (Not WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1) Then
                RaiseEvent WorkerError(message)
            End If
        Catch ex As Exception
            MsgBox("An error occurred in Specified Line Deleter's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub Activity_Logger(ByVal identifier_msg As String)
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & Format(Now(), "yyyyMMdd") & "_Activity_Log.txt", True)

            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg)


            filewriter.Flush()
            filewriter.Close()

        Catch exc As Exception
            Error_Handler(exc)
        End Try
    End Sub

    Public Sub ChooseThreads(ByVal threadNumber As Integer)
        Try
            ' Determines which thread to start based on the value it receives.
            Select Case threadNumber
                Case 1
                    WorkerThread = New System.Threading.Thread(AddressOf WorkerFileCount_Routine)
                    WorkerThread.Start()
                Case 2
                    filereader.Close()
            End Select
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub





    Private Sub WorkerFileCount_Routine()
        RaiseEvent WorkerStatusMessage("Running Search and Delete Check", 1)
        Activity_Logger("-- Process Started --")
        filecount = 0
        fullcount = 0
        blankcount = 0

        RaiseEvent WorkerFileCount(filecount, 1)
        RaiseEvent WorkerFileCount(fullcount, 2)
        RaiseEvent WorkerFileCount(blankcount, 3)

        Try
            FileCountRunner(filequeue1)
        Catch ex As Exception
            Error_Handler(ex)
        End Try

        RaiseEvent WorkerFileCount(filecount, 1)
        RaiseEvent WorkerFileCount(fullcount, 2)
        RaiseEvent WorkerFileCount(blankcount, 3)

        RaiseEvent WorkerStatusMessage("Search and Delete Check Completed", 1)
        Activity_Logger("-- Process Ended --")
        RaiseEvent WorkerComplete(0)
    End Sub

    Private Sub FileCountRunner(ByVal filename As String)
        Try
            Activity_Logger("File: " & filename)
            filereader = New StreamReader(filename, True)
            filewriter = New StreamWriter(filename & "_XTEMPX.txt", False)
            RaiseEvent WorkerStatusMessage("Examining: " & filename, 2)
            Dim linetocheck As String
            Dim previousline As ArrayList = New ArrayList
            linetocheck = Nothing
            'previousline = ""
            While filereader.Peek <> -1
                If (Not WorkerThread.ThreadState.ToString.IndexOf("Aborted") > -1) And (Not WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1) Then


                linetocheck = filereader.ReadLine
                    previousline.Add(linetocheck)

                filecount = filecount + 1
                RaiseEvent WorkerFileCount(filecount, 1)

                If linetocheck.Length > 0 And Not linetocheck = "" Then
                Else
                    blankcount = blankcount + 1
                    RaiseEvent WorkerFileCount(blankcount, 3)
                End If


                Else
                Exit While
                End If
            End While
            filereader.Close()

            If previousline.Count > 0 Then


                '      previousline.Sort()
                '   Dim prev, curr As String
                '   prev = ""
                '   curr = ""
                For Each res As String In previousline
                    '    prev = curr
                    '    curr = res
                    If res.IndexOf(searchterm) = -1 Then
                        filewriter.WriteLine(res)
                    Else
                        fullcount = fullcount + 1
                        RaiseEvent WorkerFileCount(fullcount, 2)
                        Activity_Logger("Removed Item: " & res)
                    End If
                Next
            Else
                RaiseEvent WorkerFileCount(fullcount, 2)
            End If

            filewriter.Close()

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub 'ProcessDirectory

End Class
