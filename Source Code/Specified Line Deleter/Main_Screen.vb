Imports System.IO

Public Class Main_Screen
    Inherits System.Windows.Forms.Form

    Dim WithEvents Worker1 As Worker

    Private workerbusy As Boolean = False
    Private steps As Integer = 0
    'steps 0: process not launched
    'steps 1: line count

    Public Delegate Sub WorkerComplete_h()
    Public Delegate Sub WorkerError_h(ByVal Message As Exception)
    Public Delegate Sub WorkerFileCount_h(ByVal Result As Long)
    Public Delegate Sub WorkerStatusMessage_h(ByVal message As String, ByVal statustag As Integer)
    Public Delegate Sub WorkerFileProcessing_h(ByVal filename As String, ByVal queue As Integer)

    Private application_exit As Boolean = False
    Private shutting_down As Boolean = False
    Private splash_loader As Splash_Screen
    Public dataloaded As Boolean = False
    Private error_reporting_level

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerCompleteHandler
        AddHandler Worker1.WorkerError, AddressOf WorkerErrorHandler
        AddHandler Worker1.WorkerFileCount, AddressOf WorkerFileCountHandler
        AddHandler Worker1.WorkerStatusMessage, AddressOf WorkerStatusMessageHandler
        AddHandler Worker1.WorkerFileProcessing, AddressOf WorkerFileProcessingHandler

    End Sub

    Public Sub New(ByVal splash As Splash_Screen)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        splash_loader = splash
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerCompleteHandler
        AddHandler Worker1.WorkerError, AddressOf WorkerErrorHandler
        AddHandler Worker1.WorkerFileCount, AddressOf WorkerFileCountHandler
        AddHandler Worker1.WorkerStatusMessage, AddressOf WorkerStatusMessageHandler
        AddHandler Worker1.WorkerFileProcessing, AddressOf WorkerFileProcessingHandler
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ButtonFolderBrowse As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtBaseFolder As System.Windows.Forms.TextBox
    Friend WithEvents FullError As System.Windows.Forms.CheckBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents txtProcessLaunched As System.Windows.Forms.Label
    Friend WithEvents ButtonOperationLaunch As System.Windows.Forms.Button
    Friend WithEvents txtStatus As System.Windows.Forms.TextBox
    Friend WithEvents txtFileCount As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents txtFolderCount As System.Windows.Forms.Label
    Friend WithEvents txtStatusBar1 As System.Windows.Forms.TextBox
    Friend WithEvents Button_Pause1 As System.Windows.Forms.Button
    Friend WithEvents Button_ExitThread As System.Windows.Forms.Button
    Friend WithEvents txtFileCount2 As System.Windows.Forms.Label
    Friend WithEvents txtFileCount3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtSearchTerm As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Main_Screen))
        Me.txtBaseFolder = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.ButtonFolderBrowse = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.FullError = New System.Windows.Forms.CheckBox
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ButtonOperationLaunch = New System.Windows.Forms.Button
        Me.Button_Pause1 = New System.Windows.Forms.Button
        Me.Button_ExitThread = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtSearchTerm = New System.Windows.Forms.TextBox
        Me.txtProcessLaunched = New System.Windows.Forms.Label
        Me.txtStatus = New System.Windows.Forms.TextBox
        Me.txtFileCount = New System.Windows.Forms.Label
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.txtFolderCount = New System.Windows.Forms.Label
        Me.txtStatusBar1 = New System.Windows.Forms.TextBox
        Me.txtFileCount2 = New System.Windows.Forms.Label
        Me.txtFileCount3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'txtBaseFolder
        '
        Me.txtBaseFolder.AllowDrop = True
        Me.txtBaseFolder.BackColor = System.Drawing.Color.White
        Me.txtBaseFolder.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtBaseFolder.ForeColor = System.Drawing.Color.Black
        Me.txtBaseFolder.Location = New System.Drawing.Point(8, 56)
        Me.txtBaseFolder.Name = "txtBaseFolder"
        Me.txtBaseFolder.ReadOnly = True
        Me.txtBaseFolder.Size = New System.Drawing.Size(248, 20)
        Me.txtBaseFolder.TabIndex = 4
        Me.txtBaseFolder.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtBaseFolder, "The file to be parsed")
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(8, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(280, 24)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "BROWSE FOR THE FILE TO BE PARSED FOR THE INPUTTED SEARCH TERM AND HIT THE PROCESS" & _
        " BUTTON TO LAUNCH THE OPERATION"
        '
        'ButtonFolderBrowse
        '
        Me.ButtonFolderBrowse.BackColor = System.Drawing.Color.Gainsboro
        Me.ButtonFolderBrowse.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonFolderBrowse.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonFolderBrowse.Location = New System.Drawing.Point(264, 56)
        Me.ButtonFolderBrowse.Name = "ButtonFolderBrowse"
        Me.ButtonFolderBrowse.Size = New System.Drawing.Size(56, 20)
        Me.ButtonFolderBrowse.TabIndex = 7
        Me.ButtonFolderBrowse.Text = "BROWSE"
        Me.ToolTip1.SetToolTip(Me.ButtonFolderBrowse, "Launches the File Browser Dialog")
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(8, 40)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(136, 16)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "TEXT FILE TO BE PROCESSED:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'FullError
        '
        Me.FullError.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.FullError.Location = New System.Drawing.Point(312, 14)
        Me.FullError.Name = "FullError"
        Me.FullError.Size = New System.Drawing.Size(12, 12)
        Me.FullError.TabIndex = 9
        Me.ToolTip1.SetToolTip(Me.FullError, "If Checked, Error Handling Routine enters Full Exception Mode")
        '
        'ButtonOperationLaunch
        '
        Me.ButtonOperationLaunch.BackColor = System.Drawing.Color.Gainsboro
        Me.ButtonOperationLaunch.Enabled = False
        Me.ButtonOperationLaunch.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonOperationLaunch.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonOperationLaunch.Location = New System.Drawing.Point(32, 136)
        Me.ButtonOperationLaunch.Name = "ButtonOperationLaunch"
        Me.ButtonOperationLaunch.Size = New System.Drawing.Size(88, 20)
        Me.ButtonOperationLaunch.TabIndex = 10
        Me.ButtonOperationLaunch.Text = "Process"
        Me.ToolTip1.SetToolTip(Me.ButtonOperationLaunch, "Launches Specified Line Deleter Operation")
        '
        'Button_Pause1
        '
        Me.Button_Pause1.BackColor = System.Drawing.Color.Gainsboro
        Me.Button_Pause1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button_Pause1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_Pause1.Location = New System.Drawing.Point(128, 136)
        Me.Button_Pause1.Name = "Button_Pause1"
        Me.Button_Pause1.Size = New System.Drawing.Size(56, 20)
        Me.Button_Pause1.TabIndex = 17
        Me.Button_Pause1.Text = "PAUSE"
        Me.ToolTip1.SetToolTip(Me.Button_Pause1, "Pauses Specified Line Deleter Operation")
        Me.Button_Pause1.Visible = False
        '
        'Button_ExitThread
        '
        Me.Button_ExitThread.BackColor = System.Drawing.Color.Gainsboro
        Me.Button_ExitThread.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button_ExitThread.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_ExitThread.Location = New System.Drawing.Point(192, 136)
        Me.Button_ExitThread.Name = "Button_ExitThread"
        Me.Button_ExitThread.Size = New System.Drawing.Size(56, 20)
        Me.Button_ExitThread.TabIndex = 18
        Me.Button_ExitThread.Text = "EXIT"
        Me.ToolTip1.SetToolTip(Me.Button_ExitThread, "Exits Specified Line Deleter Operation")
        Me.Button_ExitThread.Visible = False
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.DimGray
        Me.Label1.Location = New System.Drawing.Point(244, 1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 9)
        Me.Label1.TabIndex = 21
        Me.Label1.Text = "BUILD 20060201.1"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.ToolTip1.SetToolTip(Me.Label1, "BUILD 20060201.1")
        '
        'txtSearchTerm
        '
        Me.txtSearchTerm.AllowDrop = True
        Me.txtSearchTerm.BackColor = System.Drawing.Color.White
        Me.txtSearchTerm.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSearchTerm.ForeColor = System.Drawing.Color.Black
        Me.txtSearchTerm.Location = New System.Drawing.Point(8, 96)
        Me.txtSearchTerm.Name = "txtSearchTerm"
        Me.txtSearchTerm.Size = New System.Drawing.Size(248, 20)
        Me.txtSearchTerm.TabIndex = 23
        Me.txtSearchTerm.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtSearchTerm, "The case-sensitive search term")
        '
        'txtProcessLaunched
        '
        Me.txtProcessLaunched.BackColor = System.Drawing.Color.Red
        Me.txtProcessLaunched.ForeColor = System.Drawing.Color.Black
        Me.txtProcessLaunched.Location = New System.Drawing.Point(8, 176)
        Me.txtProcessLaunched.Name = "txtProcessLaunched"
        Me.txtProcessLaunched.Size = New System.Drawing.Size(312, 16)
        Me.txtProcessLaunched.TabIndex = 12
        Me.txtProcessLaunched.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtStatus
        '
        Me.txtStatus.BackColor = System.Drawing.Color.Red
        Me.txtStatus.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStatus.ForeColor = System.Drawing.Color.Gray
        Me.txtStatus.Location = New System.Drawing.Point(40, 272)
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.ReadOnly = True
        Me.txtStatus.Size = New System.Drawing.Size(280, 10)
        Me.txtStatus.TabIndex = 15
        Me.txtStatus.Text = ""
        Me.txtStatus.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtFileCount
        '
        Me.txtFileCount.BackColor = System.Drawing.Color.Red
        Me.txtFileCount.ForeColor = System.Drawing.Color.Black
        Me.txtFileCount.Location = New System.Drawing.Point(8, 208)
        Me.txtFileCount.Name = "txtFileCount"
        Me.txtFileCount.Size = New System.Drawing.Size(312, 16)
        Me.txtFileCount.TabIndex = 16
        Me.txtFileCount.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.Filter = "All Files|*.*"
        Me.OpenFileDialog1.Title = "Select the file to check"
        '
        'txtFolderCount
        '
        Me.txtFolderCount.BackColor = System.Drawing.Color.Red
        Me.txtFolderCount.ForeColor = System.Drawing.Color.Black
        Me.txtFolderCount.Location = New System.Drawing.Point(8, 256)
        Me.txtFolderCount.Name = "txtFolderCount"
        Me.txtFolderCount.Size = New System.Drawing.Size(312, 16)
        Me.txtFolderCount.TabIndex = 13
        Me.txtFolderCount.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtStatusBar1
        '
        Me.txtStatusBar1.BackColor = System.Drawing.Color.Red
        Me.txtStatusBar1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtStatusBar1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStatusBar1.ForeColor = System.Drawing.Color.Black
        Me.txtStatusBar1.Location = New System.Drawing.Point(8, 192)
        Me.txtStatusBar1.Name = "txtStatusBar1"
        Me.txtStatusBar1.ReadOnly = True
        Me.txtStatusBar1.Size = New System.Drawing.Size(310, 13)
        Me.txtStatusBar1.TabIndex = 14
        Me.txtStatusBar1.Text = ""
        '
        'txtFileCount2
        '
        Me.txtFileCount2.BackColor = System.Drawing.Color.Red
        Me.txtFileCount2.ForeColor = System.Drawing.Color.Black
        Me.txtFileCount2.Location = New System.Drawing.Point(8, 224)
        Me.txtFileCount2.Name = "txtFileCount2"
        Me.txtFileCount2.Size = New System.Drawing.Size(312, 16)
        Me.txtFileCount2.TabIndex = 19
        Me.txtFileCount2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtFileCount3
        '
        Me.txtFileCount3.BackColor = System.Drawing.Color.Red
        Me.txtFileCount3.ForeColor = System.Drawing.Color.Black
        Me.txtFileCount3.Location = New System.Drawing.Point(8, 240)
        Me.txtFileCount3.Name = "txtFileCount3"
        Me.txtFileCount3.Size = New System.Drawing.Size(312, 16)
        Me.txtFileCount3.TabIndex = 20
        Me.txtFileCount3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(8, 80)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(136, 16)
        Me.Label2.TabIndex = 22
        Me.Label2.Text = "REMOVE LINE CONTAINING:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'Main_Screen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Red
        Me.ClientSize = New System.Drawing.Size(328, 286)
        Me.Controls.Add(Me.txtSearchTerm)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtFileCount3)
        Me.Controls.Add(Me.txtFileCount2)
        Me.Controls.Add(Me.Button_ExitThread)
        Me.Controls.Add(Me.Button_Pause1)
        Me.Controls.Add(Me.txtFileCount)
        Me.Controls.Add(Me.txtStatus)
        Me.Controls.Add(Me.txtStatusBar1)
        Me.Controls.Add(Me.txtBaseFolder)
        Me.Controls.Add(Me.txtFolderCount)
        Me.Controls.Add(Me.txtProcessLaunched)
        Me.Controls.Add(Me.ButtonOperationLaunch)
        Me.Controls.Add(Me.FullError)
        Me.Controls.Add(Me.ButtonFolderBrowse)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.ForeColor = System.Drawing.Color.Black
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(336, 320)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(336, 320)
        Me.Name = "Main_Screen"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Specified Line Deleter"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Error_Handler(ByVal message As String)
        Try

            Dim Display_Message1 As New Display_Message(message)
            Display_Message1.ShowDialog()

        Catch ex As Exception
            MsgBox("An error occurred in Specified Line Deleter's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub Error_Handler(ByVal exception As Exception)
        Try
            Dim message As String
            If FullError.Checked = True Then
                message = exception.ToString
            Else
                message = exception.Message.ToString
            End If

            If Not (exception.GetType.Name = "ThreadAbortException") Then
                Dim Display_Message1 As New Display_Message(message)
                Display_Message1.ShowDialog()
            End If

        Catch ex As Exception
            MsgBox("An error occurred in Specified Line Deleter's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub ButtonFolderBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonFolderBrowse.Click
        Try
            Dim result As DialogResult
            If Not txtBaseFolder.Text = "" And txtBaseFolder.Text Is Nothing = False Then


                Dim filecheck As System.IO.FileInfo = New System.IO.FileInfo(txtBaseFolder.Text)
                If filecheck.Exists = True Then
                    OpenFileDialog1.FileName = txtBaseFolder.Text
                End If
            End If
            result = OpenFileDialog1.ShowDialog()
            If result = DialogResult.OK Or result = DialogResult.Yes Then
                txtBaseFolder.Text = OpenFileDialog1.FileName
                ButtonOperationLaunch.Enabled = True
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub




    Private Sub SendMessage(ByVal labelname As String, ByVal message As String)
        Try
            Dim controllist As ControlCollection = Me.Controls
            Dim cont As Control

            For Each cont In controllist
                If cont.Name = labelname Then
                    cont.Text = message
                    cont.Refresh()
                    Exit For
                End If
            Next
        Catch ex As Exception
            If (Not Worker1.WorkerThread.ThreadState.ToString.IndexOf("Aborted") > -1) And (Not Worker1.WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1) Then
                Error_Handler(message & vbCrLf & labelname & vbCrLf & Worker1.WorkerThread.ThreadState.ToString)
                Error_Handler(ex)
            End If
        End Try
    End Sub

    Private Sub ButtonOperationLaunch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonOperationLaunch.Click
        ButtonOperationLaunch.Enabled = False
        SendMessage("txtProcessLaunched", "")
        SendMessage("txtStatus", "")
        SendMessage("txtStatusBar1", "")
        SendMessage("txtFolderCount", "")
        SendMessage("txtFileCount", "")

        SendMessage("txtProcessLaunched", "Launched: " & Format(Now(), "dd/MM/yyyy hh:mm:ss tt"))
        Worker1.filequeue1 = txtBaseFolder.Text
        Worker1.searchterm = txtSearchTerm.Text

        steps = 1
        workerbusy = True
        Worker1.ChooseThreads(1)
        Button_Pause1.Visible = True
        Button_Pause1.Refresh()

    End Sub

    Public Sub WorkerStatusMessageHandler(ByVal message As String, ByVal statustag As Integer)
        Try
            If statustag = 1 Then
                SendMessage("txtStatus", message)
            Else
                SendMessage("txtStatusBar1", message)
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerErrorHandler(ByVal Message As Exception)
        Try
            Error_Handler(Message)
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub


    Public Sub WorkerFileCountHandler(ByVal Result As Long, ByVal line As Integer)
        Try
            Select Case line
                Case 1
                    If (Not Worker1.WorkerThread.ThreadState.ToString.IndexOf("Aborted") > -1) And (Not Worker1.WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1) Then
                        SendMessage("txtFileCount", "Total Line Count: " & Result.ToString)
                    End If
                Case 2
                    If (Not Worker1.WorkerThread.ThreadState.ToString.IndexOf("Aborted") > -1) And (Not Worker1.WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1) Then
                        SendMessage("txtFileCount2", "Removed Line Count: " & Result.ToString)
                    End If
                Case 3
                    If (Not Worker1.WorkerThread.ThreadState.ToString.IndexOf("Aborted") > -1) And (Not Worker1.WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1) Then
                        SendMessage("txtFileCount3", "Blank Line Count: " & Result.ToString)
                    End If
            End Select

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerCompleteHandler(ByVal queue As Integer)
        Try
            Dim eventhandled As Boolean = False

            Button_Pause1.Visible = False
            Button_Pause1.Refresh()
            workerbusy = False
            ButtonOperationLaunch.Enabled = True
            SendMessage("txtFolderCount", "Completed: " & Format(Now(), "dd/MM/yyyy hh:mm:ss tt"))
            'SendMessage("txtStatusBar1", "")
            ' SendMessage("txtStatus", "File Line Count Examination Complete")
            '  If txtFileCount2.Text.EndsWith(": 0") = False Then


            If MsgBox("Process Complete. Do you wish to replace your existing text file with the newly created file?", MsgBoxStyle.YesNo, "Process Complete") = MsgBoxResult.Yes Then
                Try
                    File.Delete(txtBaseFolder.Text)
                    File.Move(txtBaseFolder.Text & "_XTEMPX.txt", txtBaseFolder.Text)
                Catch ex As Exception
                    Error_Handler(ex)
                End Try
            Else
                If MsgBox("Do you wish to keep the temporary file (" & txtBaseFolder.Text & "_XTEMPX.txt" & ") that was generated in the process?", MsgBoxStyle.YesNo, "Process Complete") = MsgBoxResult.No Then
                    Try
                        File.Delete(txtBaseFolder.Text & "_XTEMPX.txt")
                    Catch ex As Exception
                        Error_Handler(ex)
                    End Try
                End If
            End If
            '  Else
            '  File.Delete(txtBaseFolder.Text & "_XTEMPX.txt")
            ' End If
            eventhandled = True
            txtStatusBar1.Select(0, 0)
            txtBaseFolder.Select(0, 0)
            txtBaseFolder.Select()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerFileProcessingHandler(ByVal filename As String, ByVal queue As Integer)
        Try


        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Main_Screen_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Try
            Worker1.Dispose()
        Catch ex As Exception
            Error_Handler(ex)
        End Try

    End Sub

    Private Sub Button_Pause_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Pause1.Click
        Try

            If workerbusy = True Then
                Worker1.WorkerThread.Suspend()
                Button_Pause1.Text = "RESUME"
                Button_Pause1.Refresh()

                Button_ExitThread.Visible = True
                Button_ExitThread.Refresh()
                workerbusy = False
            Else

                Worker1.WorkerThread.Resume()
                Button_Pause1.Text = "PAUSE"
                Button_Pause1.Refresh()

                Button_ExitThread.Visible = False
                Button_ExitThread.Refresh()
                workerbusy = True
            End If

        Catch ex As Exception

            Error_Handler(ex)

        End Try
    End Sub

    Private Sub exit_threads()
        Try

            If Worker1.WorkerThread.ThreadState.ToString.IndexOf("Suspended") > -1 Or Worker1.WorkerThread.ThreadState.ToString.IndexOf("SuspendRequested") > -1 Then
                Worker1.WorkerThread.Resume()
            End If

            If Worker1.WorkerThread.ThreadState.ToString.IndexOf("WaitSleepJoin") > -1 Then
                Worker1.WorkerThread.Interrupt()
            End If

            If Worker1.WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1 Then
                Worker1.WorkerThread.ResetAbort()
            End If

            ' SendMessage("txtStatus", "Aborting worker thread")
            Worker1.WorkerThread.Abort()
            SendMessage("txtFolderCount", "Aborted: " & Format(Now(), "dd/MM/yyyy hh:mm:ss tt"))

            SendMessage("txtStatus", "Worker thread aborted")
            Button_ExitThread.Visible = False
            Button_ExitThread.Refresh()
            Button_Pause1.Visible = False
            Button_Pause1.Text = "PAUSE"
            ButtonOperationLaunch.Enabled = True

            Worker1.ChooseThreads(2)

            'Worker1.Dispose()
            'Worker1 = New Worker
            'AddHandler Worker1.WorkerComplete, AddressOf WorkerCompleteHandler
            'AddHandler Worker1.WorkerError, AddressOf WorkerErrorHandler
            'AddHandler Worker1.WorkerFileCount, AddressOf WorkerFileCountHandler
            'AddHandler Worker1.WorkerStatusMessage, AddressOf WorkerStatusMessageHandler
            'AddHandler Worker1.WorkerFileProcessing, AddressOf WorkerFileProcessingHandler

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Button_ExitThread_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_ExitThread.Click
        exit_threads()
    End Sub

    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dataloaded = True
        splash_loader.Visible = False
    End Sub

    Private Sub txtBaseFolder_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtBaseFolder.DragEnter
        Try
            If e.Data.GetDataPresent(DataFormats.FileDrop) Then
                e.Effect = DragDropEffects.All
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub txtBaseFolder_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtBaseFolder.DragDrop
        Try
            If e.Data.GetDataPresent(DataFormats.FileDrop) Then
                Dim MyFiles() As String
                Dim i As Integer

                ' Assign the files to an array.
                MyFiles = e.Data.GetData(DataFormats.FileDrop)
                ' Loop through the array and add the files to the list.
                'For i = 0 To MyFiles.Length - 1
                If MyFiles.Length > 0 Then
                    Dim finfo As FileInfo = New FileInfo(MyFiles(0))
                    If finfo.Exists = True Then
                        txtBaseFolder.Text = (MyFiles(0))
                        ButtonOperationLaunch.Enabled = True
                    End If
                End If
                'Next
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub
End Class
