

Imports System.IO
Public Class New_StockSheet
    ' Public DefaultFolder As String = Form1.DefaultLongPath ' Folder for all departments

    Public DefaultConfigPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "AConfig\")
    Public DefaultLongPath As String = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + Path.DirectorySeparatorChar + "Assessment\"
    Public FormConfigPath As String = DefaultConfigPath
    Public RootFolderPath As String = DefaultLongPath ' The folder for the Department names
    Public SheetConfigPath As String = "" ' The Assessment Config folder
    Public ArchiveFolder As String = "" 'The Assessment Archive folder
    Public ArchivePath As String = "" ' The Archive Folder path Plus the assessment filename
    Public PathSet As Boolean = False
    Public DeptPath As String = ""




    ' Public DefaultShortPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)) + Path.DirectorySeparatorChar
    Public LastAssessment As String = ""
    Public CurrentPath As String = ""
    Public NewDept As Boolean = False
    Public NewDeptName As String = ""
    Public NumberOfItems As Integer = 10
    Public Autosave As String = "True"
    Public Response As Integer = 0
    Public todayDate As String = DateTime.Now.ToString("dd-MM-yyyy")

    Public Sub NewDepartment() Handles MyBase.Load
        Txt_DefFolder.Text = RootFolderPath
        LoadFormConfig(FormConfigPath)
        Cmb_DeptName.Select()
    End Sub
    'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ
    'Configuration load routines
    'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ
    Private Sub LoadFormConfig(ConfigPath As String)
        Dim Record_Field() As String = Nothing
        Dim Itemdetails As String = ""
        Dim fPath As String = ConfigPath
        Dim deptPath As String = ""
        Dim Auto_Save As String = ""
        Dim Zoom As Integer = 0
        Dim cPass As Integer = 0

        Try
            Form1.CheckFolderExists(fPath)

            Dim fileExists As Boolean
            fileExists = My.Computer.FileSystem.FileExists(fPath)

            If fileExists Then
                Using freader As New System.IO.StreamReader(fPath)
                    Do While Not freader.EndOfStream
                        Itemdetails = freader.ReadLine
                        Record_Field = Split(Itemdetails, ",") '  'Split the .CSV file into  fields
                    Loop
                    freader.Close()
                End Using

                If UBound(Record_Field) = 0 Then 'The Default folder hasn't been set up yet
                    'SaveFormConfig(FormConfigPath)
                ElseIf UBound(Record_Field) = 1 Then  ' A Department has been set up. Load the department Configs
                    rootfolderPath = Record_Field(0)
                    Txt_DefFolder.Text = rootfolderPath
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, "Error loading  Config file")
        End Try

        ' See if there are any departments set up 
        'If so, load the department names into the Department combo box
        Dim subFolder As String() = Directory.GetDirectories(rootfolderPath)
        Dim DeptName As String = ""

        For Each DeptName In subFolder
            Record_Field = Split(DeptName, "\")
            DeptName = Record_Field(UBound(Record_Field))
            If DeptName <> "Archive" Then
                Cmb_DeptName.Items.Add(DeptName)
            End If
        Next
        If DeptName = "" Then 'There are no departments set up yet
            NewDept = True
        End If

    End Sub

    Private Sub LoadSheetConfig(ConfigPath As String)
        Dim Record_Field() As String = Nothing
        Dim Itemdetails As String = ""
        Dim fPath As String = ConfigPath
        Dim deptPath As String = ""
        ' Dim Auto_Save As String = ""
        Dim Zoom As Integer = 0
        Dim cPass As Integer = 0

        Try
            Form1.CheckFolderExists(fPath)

            Dim fileExists As Boolean
            fileExists = My.Computer.FileSystem.FileExists(fPath)

            If fileExists Then

                Using freader As New System.IO.StreamReader(fPath)
                    Do While Not freader.EndOfStream
                        Itemdetails = freader.ReadLine
                        Record_Field = Split(Itemdetails, ",") '  'Split the .CSV file into  fields
                    Loop
                    freader.Close()
                End Using

                If UBound(Record_Field) >= 3 Then ' We have loaded the department configs
                    Cmb_DeptName.Text = Record_Field(0)
                    Autosave = Record_Field(2)
                    Zoom = Record_Field(3)
                Else 'The department hasn't been set up yet

                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, "Error loading  Config file")
        End Try

    End Sub
    'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ
    'Input routines
    'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ
    Private Sub DefFolder(sender As Object, e As EventArgs) Handles Txt_DefFolder.Click
        FolderBrowserDialog1.InitialDirectory = Form1.DefaultLongPath
        FolderBrowserDialog1.Description = "Select a Folder"
        FolderBrowserDialog1.ShowNewFolderButton = True
        If FolderBrowserDialog1.ShowDialog = DialogResult.OK Then
            Txt_DefFolder.Text = FolderBrowserDialog1.SelectedPath & "\"
            rootfolderPath = FolderBrowserDialog1.SelectedPath & "\" ' This is the folder for all Departments
            Cmb_DeptName.Select()
        End If
    End Sub
    Private Sub SheetName(sender As Object, e As KeyEventArgs) Handles Cmb_DeptName.KeyUp
        Dim kCode As Integer = e.KeyCode
        If kCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Dim fPath = ""
            If Cmb_DeptName.Text = "" Then
                Beep()
                Exit Sub
            End If

            Try
                NewDeptName = Cmb_DeptName.Text
                DeptPath = rootfolderPath & NewDeptName & "\"
                CurrentPath = DeptPath & todayDate & ".csv"
                SheetConfigPath = FormConfigPath & NewDeptName & "\Config.sys"
                Dim FolderExists = Directory.Exists(DeptPath)

                If NewDept = True Then
                    NewDept = False
                    fPath = Form1.RootFolderPath & NewDeptName & "\" & todayDate & ".csv"
                    Form1.DeptName = NewDeptName
                    Form1.CurrentPath = fPath
                    CreateBlankSheet(fPath)
                    Btn_Save.Select()
                    Exit Sub
                End If

                If FolderExists = False Then ' This is not an existing department
                    Response = MsgBox("Do you want to load a template", vbYesNo Or vbQuestion, "Creating New Stock Sheet")
                    If Response = vbYes Then
                        LoadTemplate(RootFolderPath)
                    ElseIf Response = vbNo Then
                        fPath = Form1.RootFolderPath & NewDeptName & "\" & todayDate & ".csv"

                        Form1.DeptName = NewDeptName
                        Form1.CurrentPath = fPath
                        CreateBlankSheet(fPath)
                    End If

                Else ' It's an existing Department, so 

                    LoadSheetConfig(SheetConfigPath)
                    ' And see if there is an assesment file in the folder
                    'get the names of the files in the folder.(There should be only one)
                    Dim fName = ""
                    Dim fil = Directory.GetFiles(DeptPath)
                    Dim NrFiles = 0
                    For Each fName In fil
                        OpenFileDialog1.FileName = fName
                        fPath = OpenFileDialog1.FileName
                        NrFiles += 1
                    Next

                    If NrFiles > 0 Then
                        Form1.LV_Stock.Hide()
                        Form1.EditStockSheet = True
                        'Form1.LoadStocklist(fPath)
                        Form1.LV_Stock.Hide() 'Make sure it stays hidden  
                    End If
                End If
                NewDeptName = Cmb_DeptName.Text
                Btn_Save.Select()
            Catch ex As Exception
                MsgBox(ex.Message, vbCritical, "Error loading Configuration")
            End Try

        End If

    End Sub

    Private Sub Txt_NrOfItems_TextChanged(sender As Object, e As KeyEventArgs)
        Dim kCode As Integer = e.KeyCode
        If kCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Btn_Save.Select()
        End If
    End Sub
    'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ
    'Save routines
    'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ
    Private Sub SaveFormConfig()
        Dim e As EventArgs = Nothing
        Dim sender As Object = "NewDepartment"
        Dim fPath As String = ""
        Dim ConfigString As String = ""

        rootfolderPath = Txt_DefFolder.Text
        NewDeptName = Cmb_DeptName.Text
        If NumberOfItems = 0 Then NumberOfItems = 50

        'Write the form Configuration to the FormConfig folder

        fPath = rootfolderPath
        Form1.CheckFolderExists(fPath)
        If LastAssessment = "" Then
            LastAssessment = rootfolderPath
        End If
        ConfigString = (FormConfigPath & "," & LastAssessment)
        Try

            Using fwriter As New System.IO.StreamWriter(fPath)
                fwriter.WriteLine(ConfigString)
                fwriter.Close()
            End Using
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, "Sorry")
        End Try

    End Sub
    Private Sub SaveSheetConfig()
        Dim e As EventArgs = Nothing
        Dim sender As Object = "NewDepartment"

        Dim ConfigString As String = ""

        rootfolderPath = Txt_DefFolder.Text
        NewDeptName = Cmb_DeptName.Text
        If NumberOfItems = 0 Then NumberOfItems = 10

        'Write the Sheet Configuration to the Sheet Config folder

        Form1.CheckFolderExists(SheetConfigPath)
        ConfigString = NewDeptName & "," & NumberOfItems & "," & Autosave & ",0"

        Try
            Using fwriter As New System.IO.StreamWriter(SheetConfigPath)
                fwriter.WriteLine(ConfigString)
                fwriter.Close()
            End Using

        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, "Oop, that didn't work.")
        End Try

    End Sub

    'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ
    'Set up sheets routines
    'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ
    Private Sub CreateBlankSheet(fpath As String)
        Form1.CheckFolderExists(fpath)

        Dim BlankLine As String = ",,,,,,,"
        Try
            Using fwriter As New System.IO.StreamWriter(fpath)
                For i = 1 To NumberOfItems
                    fwriter.WriteLine(BlankLine)
                Next
                fwriter.Close()
            End Using
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, "Error writing Stock List")
        End Try

    End Sub
    Private Sub LoadTemplate(fpath As String)
        Dim Itemdetails As String = ""
        Dim Record_Field() As String = Nothing
        Dim Zoom As Integer = 0

        OpenFileDialog1.InitialDirectory = Form1.RootFolderPath
        OpenFileDialog1.FileName = fpath
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            fpath = OpenFileDialog1.FileName
        End If

        Form1.CheckFolderExists(DeptPath)
        FileCopy(fpath, CurrentPath)
        Form1.EditStockSheet = True
        Form1.CurrentPath = CurrentPath
        Form1.SheetConfigPath = SheetConfigPath

        'Get the configs for the Template sheet
        Dim p As Integer = InStr(fpath, "csv")

        Dim cPath As String = Mid(fpath, 1, p - 12) & "Archive\Config\Config.sys"

        Using freader As New System.IO.StreamReader(cPath)
            Do While Not freader.EndOfStream
                Itemdetails = freader.ReadLine
                Record_Field = Split(Itemdetails, ",") '  'Split the .CSV file into  fields
            Loop
            freader.Close()
        End Using

        If UBound(Record_Field) >= 3 Then ' We have loaded the department configs
            NumberOfItems = Record_Field(1)
            Autosave = Record_Field(2)
            Zoom = Record_Field(3)
        End If

        Form1.CheckFolderExists(SheetConfigPath)

        Dim configstring = NewDeptName & "," & Autosave & "," & Zoom
        Try
            Using fwriter As New System.IO.StreamWriter(SheetConfigPath)
                fwriter.WriteLine(configstring)
                fwriter.Close()
            End Using

        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, "Oop, that didn't work.")
        End Try


    End Sub

    'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ
    'Exit routines
    'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ
    Private Sub Btn_Save_Click(sender As Object, e As EventArgs) Handles Btn_Save.Click
        Dim NewDeptName As String = Cmb_DeptName.Text
        Dim fPath As String = ""

        fPath = RootFolderPath & NewDeptName & "\Archive\"
        Dim FolderExists = Directory.Exists(fPath)
        If FolderExists = False Then
            MkDir(fPath)
        End If

        CurrentPath = rootfolderPath & NewDeptName & "\" & todayDate & ".csv"
        SheetConfigPath = DefaultConfigPath & NewDeptName & "\Config.sys"

        SaveSheetConfig()

        'Transfer the settings to the Main Form
        Form1.CurrentPath = CurrentPath
        Form1.FileToPrint = CurrentPath
        Form1.EditStockSheet = True
        Form1.EditToolStripMenuItem.Visible = False
        Form1.DoAssessmentToolStripMenuItem.Visible = False
        Form1.DeptName = NewDeptName
        Form1.NumberOfRows = NumberOfItems
        Form1.EditStockSheet = True
        Form1.DoAssessment = False
        Form1.NewDepartment = True
        'Load the new department
        Form1.LV_Stock.Hide()
        Form1.SheetConfigPath = SheetConfigPath
        Form1.ArchivePath = RootFolderPath & NewDeptName & "\Archive\"
        Form1.SetupSheet()
        Form1.LoadStocklist(CurrentPath)
        Form1.LV_Stock.Show()
        Form1.EditCell(0, 0)
    End Sub
    Private Sub Btn_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Cancel.Click
        Txt_DefFolder.Text = ""
        Cmb_DeptName.Text = ""
        Me.Close()
    End Sub

End Class