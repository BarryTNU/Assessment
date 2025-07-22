Imports System.DirectoryServices.ActiveDirectory
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows

Public Class Form1

    Public DefaultLongPath As String = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + Path.DirectorySeparatorChar + "Assessment\"
    Public DefaultConfigPath As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + Path.DirectorySeparatorChar + "Assessment\"

    Public FormConfigPath As String = DefaultConfigPath
    Public RootFolderPath As String = "" ' The folder for the Department names
    Public SheetConfigPath As String = "" ' The Assessment Config folder
    Public ArchiveFolder As String = "" 'The Assessment Archive folder
    Public ArchivePath As String = "" ' The Archive Folder path Plus the assessment filename
    Public PathSet As Boolean = False
    Public DeptPath As String = ""
    Public CurrentPath As String = ""

    Public sender As Object = Nothing
    Public e As Object = Nothing
    Public FileName As String = ""
    Public NewAssessment As Boolean = False
    Public NewDepartment As Boolean = False


    Public currentRow As Integer = 0
    Public currentCol As Integer = 0
    Public nextRow As Integer = 0
    Public editor As TextBox = Nothing
    Public EditColumnStart As Integer = 3
    Public EditColumnEnd As Integer = 8
    Public editBox As New TextBox()
    Public NumberOfRows As Integer = 10
    Public NumberOfItems As Integer = 0

    Public Zoom As Integer = 0
    Public DeptName As String = ""
    Public EditStockSheet As Boolean = False
    Public DoAssessment As Boolean = False
    Public NewSheet As Boolean = False
    Public LastAssessment As String = ""
    Public AutoSave As Boolean = True
    Public Edit_Assessment As Boolean = False
    Public SheetSaved As Boolean = False
    Public AssessmentComplete As Boolean = False
    Public todayDate As String = DateTime.Now.ToString("dd-MM-yyyy")
    Public FileToPrint As String = ""
    Public FormSize As String = ""

    Const VK_NUMLOCK As Integer = &H90
    Const KEYEVENTF_EXTENDEDKEY As UInteger = &H1
    Const KEYEVENTF_KEYUP As UInteger = &H2

    ReadOnly ColourBlack = System.Drawing.Color.Black
    ReadOnly ColourRed = System.Drawing.Color.Red
    ReadOnly ColourGreen = System.Drawing.Color.LawnGreen
    ReadOnly ColourAmber = System.Drawing.Color.Orange
    ReadOnly ColourWhite = System.Drawing.Color.White

    'Set up Timer to handle single and double clicks
    Private clickTimer As New Timer()
    Private clickedItem As ListViewItem = Nothing


    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    Public Sub MainForm() Handles MyBase.Load ' runs on form load
        Me.StartPosition = FormStartPosition.CenterScreen
        Dim screenBounds As Rectangle = Screen.PrimaryScreen.WorkingArea
        Me.KeyPreview = True
        ForceNumLockOn() 'Turn NumLock on.  Needed to use the Number keypad
        FormSize = Str(Me.Width) & "," & Str(Me.Height) ' Save the form size for use later
        Me.Width = LV_Stock.Width + 50 ' 
        Me.Height = LV_Stock.Height
        Me.Left = (screenBounds.Width - Me.Width) \ 2
        Me.Top = (screenBounds.Height - Me.Height) \ 2
        Me.Text = CurrentPath
        Me.Controls.Add(editBox)
        editBox.Visible = False

        AutoSave = True
        CloseToolStripMenuItem.Visible = False
        SaveToolStripMenuItem.Visible = False
        SaveAsToolStripMenuItem.Visible = False
        DoAssessmentToolStripMenuItem.Visible = False
        EditToolStripMenuItem.Visible = False
        SaveSeparator.Visible = False
        SaveAsSeparator.Visible = False
        EditSeparator.Visible = False
        DoAssessSeparator.Visible = False
        ExitSeparator.Visible = False
        ToolStripSeparator6.Visible = False
        SettingsToolStripMenuItem.Visible = False

        LV_Stock.View = View.Details
        LV_Stock.FullRowSelect = True
        LV_Stock.OwnerDraw = True
        GetFormConfig()

        ' Set up the timer slightly above system double-click time
        clickTimer.Interval = SystemInformation.DoubleClickTime + 10
        AddHandler clickTimer.Tick, AddressOf OnSingleClickConfirmed
        ' Hook mouse events
        AddHandler LV_Stock.MouseClick, AddressOf StockList_MouseClick
        AddHandler LV_Stock.MouseDoubleClick, AddressOf CellText

    End Sub

    Sub GetFormConfig() ' Sets the form defaults

        Dim Record_Field() As String = Nothing
        Dim FolderPath As String '=""

        CheckFolderExists(FormConfigPath)
        Dim fileCount As Integer = Directory.GetFiles(FormConfigPath, "Config.sys").Length
        FormConfigPath += "Config.sys"
        If fileCount = 0 Then ' The config file has not been set up
            Dim configstring As String = DefaultLongPath & "," & DefaultLongPath
            Try
                Using fwriter As New System.IO.StreamWriter(FormConfigPath, False)
                    fwriter.WriteLine(configstring)
                    fwriter.Close()
                End Using

            Catch ex As Exception
                MsgBox(ex.Message, vbCritical, "Error writing  Configuration file")
            End Try
        End If
        Dim fPath As String = FormConfigPath

        Try
            Using freader As New System.IO.StreamReader(fPath)
                Do While Not freader.EndOfStream
                    FolderPath = freader.ReadLine
                    Record_Field = Split(FolderPath, ",")
                Loop
                freader.Close()
            End Using
            If Record_Field(0) <> "" Then
                RootFolderPath = Record_Field(0)
            End If

            If RootFolderPath = "" Then
                RootFolderPath = DefaultLongPath
            End If

            If UBound(Record_Field) > 0 Then
                DeptPath = Record_Field(UBound(Record_Field))
            Else
                DeptPath = RootFolderPath
            End If

            CheckFolderExists(RootFolderPath)

        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, "Error loading Folder Path")
        End Try

        If Directory.Exists(RootFolderPath) Then ' There are some departments set up
            fileCount = Directory.GetDirectories(RootFolderPath).Length
            If fileCount > 0 Then '
                AssessmentToolStripMenuItem.Visible = True
            End If
        End If

        Me.Text = RootFolderPath

        Dim ImagePath As String = DefaultConfigPath & "Background.pic"
        If File.Exists(ImagePath) Then ' There is a background image
            Dim bmp As Bitmap
            Using fs As New FileStream(ImagePath, FileMode.Open, FileAccess.Read)
                bmp = New Bitmap(fs)
                Me.BackgroundImage = bmp
            End Using
        End If

    End Sub

    Public Sub LoadSheetConfig(SheetConfigPath As String) ' Sets the Sheet Configuration
        'Load the Sheet defaults'
        Dim itemdetails As String
        Dim Record_Field()
        Dim Auto_Save As String = ""
        Dim i As Integer ' = 0

        If SheetConfigPath = "" Then Exit Sub
        Try
            i = 0
            Using fReader As New StreamReader(SheetConfigPath)
                Do While Not fReader.EndOfStream
                    itemdetails = fReader.ReadLine
                    Record_Field = Split(itemdetails, ",") '  'Split the .CSV file into  fields
                    If Record_Field(0) = "" Then Continue Do
                    DeptName = Record_Field(0)
                    NumberOfRows = Record_Field(1)
                    Auto_Save = Record_Field(2)
                    Zoom = Record_Field(3)
                Loop
                fReader.Close()
            End Using
        Catch ex As Exception
            'MsgBox(ex.Message, vbCritical, "Error loading Config file")
        End Try

        If Auto_Save = "True" Or Auto_Save = "" Then
            AutoSave = True
        Else
            AutoSave = False
        End If

        If AutoSave = False Then
            SaveToolStripMenuItem.Visible = True
            SaveAsToolStripMenuItem.Visible = True
        Else
            SaveToolStripMenuItem.Visible = False
            SaveAsToolStripMenuItem.Visible = False
        End If

        ZoomToolStripMenuItem.Visible = True
        SettingsToolStripMenuItem.Visible = True

    End Sub

    Public Sub SetupSheet()
        Dim Rows As Integer = 2

        Try
            With LV_Stock

                .View = View.Details
                .OwnerDraw = True
                .FullRowSelect = True
                .Hide()
                .Clear() 'Clear the previous entries
                .FullRowSelect = True
                .Dock = DockStyle.Fill

                Dim List_header = "Stock Item Name  " & "," & "Size" & "," & "Value" & "," & "Opening" & "," & "Issues" & "," & "Closing" & "," & "Sold" & "," & "Yield"

                ' The header titles are passed to the subroutine as List_Header, and returned as an array
                Dim Record_Field() As String = Split(List_header, ",")
                .Columns.Insert(0, Record_Field(0), 260, CType(HorizontalAlignment.Center, Forms.HorizontalAlignment))
                .Columns.Insert(1, Record_Field(1), 90, CType(HorizontalAlignment.Center, Forms.HorizontalAlignment))
                .Columns.Insert(2, Record_Field(2), 90, CType(HorizontalAlignment.Center, Forms.HorizontalAlignment))
                .Columns.Insert(3, Record_Field(3), 90, CType(HorizontalAlignment.Center, Forms.HorizontalAlignment))
                .Columns.Insert(4, Record_Field(4), 90, CType(HorizontalAlignment.Center, Forms.HorizontalAlignment))
                .Columns.Insert(5, Record_Field(5), 90, CType(HorizontalAlignment.Center, Forms.HorizontalAlignment))
                .Columns.Insert(6, Record_Field(6), 90, CType(HorizontalAlignment.Center, Forms.HorizontalAlignment))
                .Columns.Insert(7, Record_Field(7), 90, CType(HorizontalAlignment.Center, Forms.HorizontalAlignment))
                .Refresh()
                .View = View.Details

                'Fill the List with spaces
                For i = 1 To Rows 'i is row number. Row 0 is already added
                    .Items.Add(New ListViewItem(New String() {"", "", "", "", "", "", "", ""}))
                Next

            End With

        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, "Error setting up form")
        End Try


    End Sub

    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'File Loading and Saving Handlers
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

    'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ
    'Form housekeeping routines
    'ZZZZZZZZZZZZZZZZZZZZZZZZZZZ
    Public Sub CheckFolderExists(FName As String)
        Dim Letter As String ' = ""
        If FName = "" Then Exit Sub
        For i = Len(FName) To 0 Step -1
            Letter = Mid(FName, i, 1)
            If Letter = "\" Then
                FName = Mid(FName, 1, i)
                Exit For
            End If
        Next
        Dim FolderExists = Directory.Exists(FName)
        If FolderExists = False Then
            MkDir(FName)
        End If
    End Sub


    Sub SaveFormConfig(ConfigString As String)
        Dim fPath As String = FormConfigPath

        'Write the form Configuration to the Config folder
        If fPath = "" Then Exit Sub
        CheckFolderExists(fPath)
        Try
            Using fwriter As New System.IO.StreamWriter(fPath, False)
                fwriter.WriteLine(ConfigString)
                fwriter.Close()
            End Using
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, "Error writing  Form Config file")
        End Try

    End Sub
    Sub SaveSheetConfig(sheetconfigpath As String)
        Dim fpath As String ' = ""
        Dim ConfigString As String
        If sheetconfigpath = "" Then Exit Sub
        'Write the Sheet Configuration to the Sheet Config folder
        fpath = sheetconfigpath
        CheckFolderExists(fpath)

        If Zoom > 6 Then Zoom = 6
        ConfigString = DeptName & "," & NumberOfRows & "," & AutoSave & "," & Zoom

        Try
            Using fwriter As New System.IO.StreamWriter(fpath)
                fwriter.WriteLine(ConfigString)
                fwriter.Close()
            End Using

        Catch ex As Exception
            'MsgBox(ex.Message, vbCritical, "Oops, that didn't work.")
        End Try

    End Sub

    Public Sub GetFilePaths(fPath As String)
        ' Dim response As Integer ' = 0

        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        'Set up the Working folders etc
        'Set up the default folders
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX 
        'This is the only place these folders are set
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

        If fPath = "" Then Exit Sub
        Dim Folder() = Split(fPath, "\")
        Dim NrOfFolders As Integer = UBound(Folder)

        If NrOfFolders > 0 Then
            'Extract the Department name and path
            'Folder is the complete file path including the file name
            'So the department name is the 2nd last item
            ' And if this is the Archive folder, then the dept name is the 3rd last item
            RootFolderPath = Folder(0) & "\" & Folder(1) & "\" & Folder(2) & "\" ' RootFolderPath is the path to the folder which holds the department names
            FileName = Folder(NrOfFolders) 'The Assessment Name
            DeptName = Folder(UBound(Folder) - 1)
            If DeptName = "Archive" Then '
                DeptName = Folder(UBound(Folder) - 2)
            End If
            DeptPath = RootFolderPath & DeptName & "\"
            ArchiveFolder = DeptPath & "Archive\"
            ArchivePath = ArchiveFolder & FileName
            SheetConfigPath = DefaultConfigPath & DeptName & "\Config.sys"
        End If

        CurrentPath = fPath
        Dim Sheetpath As String = RootFolderPath & DeptName & "\"
        Dim ConfigString As String = RootFolderPath & "," & Sheetpath

        If ConfigString <> "" Then
            SaveFormConfig(ConfigString)
        End If

        PathSet = True
    End Sub

    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'Load Routines
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

    Public Sub GetStockListPath(sender As Object, e As EventArgs) Handles AssessmentToolStripMenuItem.Click
        ' In this routine we choose the Department we wish to Assess
        Dim response As Integer ' = 0
        Dim fPath As String ' = ""
        Dim NrFiles As Integer ' = 0


        DoAssessment = False ' cancel these flags
        EditStockSheet = False

        If SheetConfigPath <> "" Then
            SaveSheetConfig(SheetConfigPath)
        End If

        Try

            If Directory.Exists(DeptPath) Then
                Dim fil As String() = Directory.GetFiles(DeptPath)
                Dim fName As String = ""
                For Each fName In fil
                    OpenFileDialog1.FileName = fName
                    LastAssessment = OpenFileDialog1.FileName
                    NrFiles += 1
                Next
                OpenFileDialog1.InitialDirectory = LastAssessment
            Else
                OpenFileDialog1.InitialDirectory = RootFolderPath
            End If

            LV_Stock.Hide() 'Hide the sheet while we are loading it
            OpenFileDialog1.Filter = "Assessment Files (*.csv)|*.csv|Backup Files (*.bak)|*.bak"
            OpenFileDialog1.FileName = LastAssessment
            OpenFileDialog1.Multiselect = False

            response = OpenFileDialog1.ShowDialog()
            If response = vbOK Then
                FileName = OpenFileDialog1.FileName
                If System.IO.File.Exists(OpenFileDialog1.FileName) Then
                    fPath = FileName
                Else    ' File doesn't exist (may occur in some cases)
                    Exit Sub
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, "Error Opening Assessment")
        End Try

        GetFilePaths(fPath)

        PathSet = True
        LoadSheetConfig(SheetConfigPath)
        SetupSheet()
        LV_Stock.Hide() 'Hide the sheet wile loading
        LoadStocklist(fPath)

    End Sub

    Public Sub LoadStocklist(FileName As String)
        Dim screenBounds As Rectangle = Screen.PrimaryScreen.WorkingArea
        Dim fPath As String = ""
        Dim response As Integer = vbYes
        Dim todayDate As String = DateTime.Now.ToString("dd-MM-yyyy")
        Dim ItemDetails As String
        Dim Record_Field()
        Dim i As Integer = 0
        Dim OpeningStock As String = ""
        Dim ClosingStock As String = ""

        'In GetstocklistPath we have chosen the Department we wish to do the Assessment for and
        'Selected a file

        'There are three possibilities. either we are                                                          ;
        '1: Loading a previously completed Assessment file, intending to start another Assessment. or we are
        '2: Loading an assessment which has been Archived, to either look at it, or make changes to it.
        '   (this is prohibited, so the original assessment is retained, and the updated one is saved as a .bak file)
        '3: Or there was no pre-existing assessment in the Department folder, so we are starting a new Assessment sheet

        NumberOfItems = 0

        If FileName = "" Then Exit Sub
        fPath = FileName
        Dim x As Integer = Len(fPath)
        If x <> 0 Then
            DeptPath = Mid(fPath, 1, x - 14)
        End If

        If editor IsNot Nothing Then
            LV_Stock.Controls.Remove(editor)
            editor.Dispose()
        End If

        Me.Hide() ' Hide the main form while the Assessment is being loaded

        If NewDepartment = True Or InStr(fPath, "Archive") <> 0 Then GoTo CurrentAssessment

        response = MsgBox("This is a current Assessment." & Chr(13) & "Do you intend to start a New Assessment", vbYesNo Or vbQuestion Or vbDefaultButton2, "")
        If response = vbYes Then
            NewAssessment = True
            DoAssessment = True
            EditToolStripMenuItem.Visible = True
            DoAssessmentToolStripMenuItem.Visible = False
        Else
            NewAssessment = False
            EditStockSheet = False
            DoAssessment = False
            EditToolStripMenuItem.Visible = True
            DoAssessmentToolStripMenuItem.Visible = True
            'Leave open the option to either edit the Stock details, or do an assessment
        End If

CurrentAssessment:

        With LV_Stock
            NumberOfRows = 0
            Try
                i = 0
                Using fReader As New StreamReader(fPath) ' Get the Assessment data
                    Do While Not fReader.EndOfStream
                        ItemDetails = fReader.ReadLine
                        If ItemDetails <> "" Then
                            Record_Field = Split(ItemDetails, ",") '  'Split the .CSV file into  fields
                            Dim q As Integer = UBound(Record_Field)

                            If q <> 7 Then 'Check that it's a valid Assessment file
                                MsgBox("This is not an Assessment file" & Chr(13) & "Please choose an Assessment file", vbOKOnly Or vbExclamation, "")
                                fReader.Close()
                                Me.Show() 'Show the main form before exiting
                                Exit Sub
                            End If

                            ClosingStock = Record_Field(5)
                            For p = 0 To UBound(Record_Field)  ' Break the string into fields
                                .Items(i).SubItems(p).Text = Record_Field(p)
                            Next

                            If .Items(i).SubItems(2).Text <> "" Then
                                Dim StockValue As Single = Val(.Items(i).SubItems(2).Text)
                                .Items(i).SubItems(2).Text = (StockValue).ToString("C")
                                Dim StockYield As Single = Val(.Items(i).SubItems(7).Text)
                                .Items(i).SubItems(7).Text = (StockYield).ToString("C")
                            End If

                            If NewAssessment = True And ClosingStock <> "" Then
                                'We are opening a previous assessment, and starting a new one
                                ' So Move closing stock to opening stock and clear Sold and Yield cells
                                .Items(i).SubItems(3).Text = ClosingStock
                                .Items(i).SubItems(4).Text = ""
                                .Items(i).SubItems(5).Text = ""
                                .Items(i).SubItems(6).Text = ""
                                .Items(i).SubItems(7).Text = ""
                            ElseIf NewDepartment = True Then
                                'Clear all but first 3 fields,
                                .Items(i).SubItems(3).Text = ""
                                .Items(i).SubItems(4).Text = ""
                                .Items(i).SubItems(5).Text = ""
                                .Items(i).SubItems(6).Text = ""
                                .Items(i).SubItems(7).Text = ""
                            End If

                            If .Items(i).SubItems(0).Text <> "" Then
                                NumberOfItems += 1 'This is the number of stock items
                            End If

                            i += 1

                            'Add another row
                            .Items.Add(New ListViewItem(New String() {"", "", "", "", "", "", "", ""}))
                            NumberOfRows += 1
                        End If
                    Loop
                    fReader.Close()
                    NewDepartment = False
                End Using

            Catch ex As Exception
                MsgBox(ex.Message, vbCritical, "Subscript out of range")
                EditStockSheet = False
                DoAssessmentToolStripMenuItem.Visible = False
                EditToolStripMenuItem.Visible = False
                FileToPrint = "" ' The file we send to the printer
                PrintToolStripMenuItem.Visible = False

            End Try

            CalculateYield()

            Me.Show()
            If NewDepartment = False Then
                GetFilePaths(fPath)
                LoadSheetConfig(SheetConfigPath)
            ElseIf NewDepartment = True Then
                EditStockSheet = True
                DoAssessmentToolStripMenuItem.Visible = True
                EditToolStripMenuItem.Visible = False
                NewDepartment = False
            End If

            .Visible = True

            CurrentPath = fPath
            FileToPrint = fPath ' The file we send to the printer
            PrintToolStripMenuItem.Visible = True

        End With

        If editor IsNot Nothing Then
            LV_Stock.Controls.Remove(editor)
            editor.Dispose()
        End If

        If AutoSave = False Then
            SaveToolStripMenuItem.Visible = True
            SaveAsToolStripMenuItem.Visible = True
        End If

        i = InStr(fPath, "Archive")
        If i > 0 Then ' If we are editing a completed assessment then leave the original, and save it as a backup
            Dim TempPath As String = Replace(fPath, "csv", "bak")
            SaveSheet(TempPath) 'Save the sheet backup before we start editing it
        Else 'we have loaded a previously completed assessment from the Root folder, so assumedly we are doing a new assessment.
            Dim q As Integer = InStr(fPath, "csv")
            Dim f_Name = Mid(fPath, q - 11)
            FileCopy(fPath, DeptPath & "Archive\" & f_Name) ' Copy the previous assessment to the Archive folder.
            CurrentPath = DeptPath & todayDate & ".csv" ' And start a new one

            If InStr(CurrentPath, todayDate) = False Then
                Rename(fPath, CurrentPath)
            End If
        End If

        Dim aName As String = Replace(CurrentPath, RootFolderPath, "")
        aName = Replace(aName, ".csv", "")
        Me.Text = "Assessment for " & aName ' Put the current department as the screen header

        NewSheet = False
        NewAssessment = False
        NewDepartment = False
        CloseToolStripMenuItem.Visible = True
        ZoomToolStripMenuItem.Visible = True
        SettingsToolStripMenuItem.Visible = True
        PrintToolStripMenuItem.Visible = True
        PrintAssessmentToolStripMenuItem.Visible = True
        PrintWorkSheetToolStripMenuItem.Visible = True

        'Center form onscreen
        screenBounds = Screen.PrimaryScreen.WorkingArea
        Me.Left = (screenBounds.Width - Me.Width) \ 2
        Me.Top = (screenBounds.Height - Me.Height) \ 2

        'Zoom the form  to it's previous size
        Dim Z As Integer = Zoom
        For i = 0 To Z
            Form_Zoom.ZoomIn(sender, e, Me, LV_Stock)
        Next

        If DoAssessment = True Then ' Highlight the Stock Issues and closing fields
            EditCellText(0, 4)
        ElseIf EditStockSheet = True Then ' Highlight the Item Name & Details fields
            EditCellText(0, 0)
        End If

        Me.Show() ' And show the main form after the assessment is loaded
        LV_Stock.Show()

    End Sub

    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'Save routines
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    Private Sub SaveAndMoveNext()
        Dim txt As String = LV_Stock.Items(currentRow).SubItems(currentCol).Text

        LV_Stock.Items(currentRow).SubItems(currentCol).Text = editor.Text
        If currentCol = 2 Then 'this is the item cost
            Dim ItemCost As Decimal
            If Decimal.TryParse(LV_Stock.Items(currentRow).SubItems(currentCol).Text, ItemCost) Then
                LV_Stock.Items(currentRow).SubItems(currentCol).Text = ItemCost.ToString("C")
            End If
        End If

        If EditStockSheet = True Then
            EditColumnStart = 0 : EditColumnEnd = 3
        ElseIf DoAssessment = True Then
            EditColumnStart = 4 : EditColumnEnd = 5
        End If

        If currentCol = 2 Or currentCol = 5 And editor.Text <> "" Then
            CalculateYield()
        End If

        If AutoSave = True Then
            If currentCol = 2 Or currentCol = 3 Or currentCol = 5 Then 'Only save if there's enough data
                Dim s_ender As String = "AutoSave"
                SaveStockList(s_ender, CurrentPath)
            End If
        End If
        currentCol += 1

        If currentCol >= EditColumnStart And currentCol > EditColumnEnd Then
            currentCol = EditColumnStart
            currentRow += 1
        End If

        If currentRow >= LV_Stock.Items.Count - 1 And DoAssessment = True Then
            currentRow = 0
            If editor IsNot Nothing Then
                LV_Stock.Controls.Remove(editor)
                editor.Dispose()
            End If

        End If

        If DoAssessment = True Then
            If LV_Stock.Items(currentRow).SubItems(0).Text = "" And (currentCol + 1) = 5 Then
                AssessmentComplete = True

                ' Remove previous editor if any
                If editor IsNot Nothing Then
                    LV_Stock.Controls.Remove(editor)
                    editor.Dispose()
                End If
            Else

                EditCellText(currentRow, currentCol)
            End If
        End If


        If EditStockSheet = True Then

            If nextRow + 1 >= LV_Stock.Items.Count - 1 Then
                LV_Stock.Items.Add(New ListViewItem(New String() {"", "", "", "", "", "", "", ""}))
                NumberOfRows = LV_Stock.Items.Count
                SaveSheetConfig(SheetConfigPath)
                EditCellText(currentRow, currentCol)
            Else
                EditCellText(currentRow, currentCol)
            End If
        End If

    End Sub

    Public Sub SaveStockList(S_ender As Object, StocklistPath As String)
        ' Dim i As Integer ' = 0
        ' Dim f As Integer ' = 0
        Dim fPath As String = StocklistPath
        Dim todayDate As String = DateTime.Now.ToString("dd-MM-yyyy") ' Format: dd-MM-yyyy_HH-mm 

        If StocklistPath = "" Then Exit Sub

        If StocklistPath = "" Then 'Probably never going to happen
            MsgBox("Please load a Stock Sheet", MsgBoxStyle.OkOnly, "Stock Sheet Is Empty")
            Exit Sub
        End If

        If AutoSave = True Then
            SaveSheet(fPath)
            Exit Sub
        End If

        If S_ender Is "Save" Then ' 
            If StocklistPath <> "" Then
                SaveSheet(StocklistPath)
                Exit Sub
            End If
        ElseIf S_ender Is "Exit" Then
            fPath = StocklistPath
            SaveSheet(fPath)
            Me.Close()
            Exit Sub
        ElseIf S_ender Is "SaveAs" Then
            OpenFileDialog1.Title = "Save As"
            OpenFileDialog1.DefaultExt = ".csv"
            OpenFileDialog1.FileName = StocklistPath

            Try
                If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
                    fPath = SaveFileDialog1.FileName
                    Dim fP As Integer = InStr(fPath, ".")
                    If fP <> 0 Then
                        fPath = Mid(fPath, 1, fP - 1) & ".csv"
                    Else
                        fPath += ".csv"
                    End If
                    SaveSheet(fPath)
                    Exit Sub
                End If

            Catch ex As Exception
                MsgBox(ex.Message, vbCritical, "Error Parsing File")
            End Try
        End If

    End Sub

    Public Sub SaveSheet(fPath As String)
        Dim row As String = ""
        Dim i As Integer
        Dim FileName As String
        Dim NumberOfItems As Integer = 0
        Dim q As Boolean = False

        If fPath = "" Then Exit Sub
        If InStr(fPath, "Archive") <> 0 Then
            fPath = Replace(fPath, "csv", "bak") 'save the  same file to the Archive folder
        End If

        FileToPrint = fPath

        With LV_Stock
            Try
                Using fwriter As New System.IO.StreamWriter(fPath)
                    For Each item As ListViewItem In .Items
                        If item.SubItems(0).Text = Nothing Then
                            Exit For
                        End If


                        Dim sValue As String = item.SubItems(2).Text ' Strip off the $  and , from the item value
                        Dim oldValue As String = sValue
                        sValue = Replace(sValue, "$", "")
                        sValue = Replace(sValue, ",", "")
                        item.SubItems(2).Text = sValue

                        Dim sYield As String = item.SubItems(7).Text
                        Dim oldYield As String = sYield 'Save the old yield
                        sYield = Replace(sYield, "$", "") ' Strip off the $  and , from the item 
                        sYield = Replace(sYield, ",", "") ' Messes with the calculation. Also saving a comma is prohibited as the comma is the field separator
                        item.SubItems(7).Text = sYield

                        row = item.SubItems(0).Text 'This is the first item

                        For i = 1 To 7 'Starting on item 2
                            If item.SubItems(i).Text = "Total Yield" Then
                                item.SubItems(i).Text = ""
                                item.SubItems(i + 1).Text = ""
                            End If
                            row += "," & item.SubItems(i).Text
                        Next

                        'Dim q As Integer = InStr(row, "Total Yield")

                        fwriter.WriteLine(row)
                        row = ""
                        NumberOfItems += 1
                        item.SubItems(2).Text = oldValue 'Put it back again
                        item.SubItems(7).Text = oldYield 'Put it back again
                    Next

                    If NumberOfItems < 10 Then ' Put in a few blank lines to  make the screen look better
                        For i = NumberOfItems To 10
                            fwriter.WriteLine(",,,,,,,")
                        Next
                    End If

                    fwriter.Close()
                End Using
                SheetSaved = True

                If EditStockSheet = True Then
                    EditToolStripMenuItem.Visible = False
                    DoAssessmentToolStripMenuItem.Visible = True
                ElseIf DoAssessment = True Then
                    EditToolStripMenuItem.Visible = True
                    DoAssessmentToolStripMenuItem.Visible = False
                End If

                q = InStr(fPath, "Archive")
                If q = True Then
                    Exit Sub
                End If

            Catch ex As Exception
                MsgBox(ex.Message, vbCritical, "Error writing Stock List")
            End Try
        End With

        Try
            If DoAssessment = True Then ' Move all files except the current assessment to the archive folder
                Dim Spl() As String = Nothing
                For Each file As String In Directory.GetFiles(DeptPath) ' delete all files except the current one
                    Spl = Split(file, "\")
                    FileName = Spl(UBound(Spl))
                    q = InStr(FileName, todayDate)
                    If q = False Then
                        FileCopy(DeptPath & FileName, ArchiveFolder & FileName) 'Copy the file to Archive
                        Kill(DeptPath & FileName) 'Delete the file from the Department folder
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, "Error moving files")
        End Try

        PrintToolStripMenuItem.Visible = True
        SheetSaved = True
    End Sub

    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'Form Navigation and drawing event handlers  
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    Private Sub ListView1DrawColumnHeader(sender As Object, e As DrawListViewColumnHeaderEventArgs) Handles LV_Stock.DrawColumnHeader
        ' Set custom background color
        e.Graphics.FillRectangle(New SolidBrush(Color.Black), e.Bounds)

        ' Set custom text color and alignment
        Dim sf As New StringFormat()
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center
        e.Graphics.DrawString(e.Header.Text, LV_Stock.Font, Brushes.White, e.Bounds, sf)
        '  Draw border
        e.Graphics.DrawRectangle(Pens.Red, e.Bounds)
    End Sub


    Private Sub ListView_DrawItem(sender As Object, e As DrawListViewItemEventArgs) Handles LV_Stock.DrawItem
        ' Required to make OwnerDraw work properly
        e.DrawDefault = False
    End Sub

    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        LV_Stock.Width = Me.ClientSize.Width - 20
        LV_Stock.Height = Me.ClientSize.Height - 20
    End Sub

    Private Sub LVistView_DrawSubItem(sender As Object, e As DrawListViewSubItemEventArgs) Handles LV_Stock.DrawSubItem

        Dim subItemText As String = e.SubItem.Text.Trim()
        Dim isEmpty As Boolean = String.IsNullOrEmpty(subItemText)
        Dim backColor As Color = Color.White
        Using brush As New SolidBrush(backColor)
            e.Graphics.FillRectangle(brush, e.Bounds)
        End Using

        TextRenderer.DrawText(e.Graphics, e.SubItem.Text, LV_Stock.Font, e.Bounds, e.SubItem.ForeColor, TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)

    End Sub

    Private Sub ColourCellBlock(sender As Object, e As DrawListViewSubItemEventArgs) Handles LV_Stock.DrawSubItem
        Dim row As Integer = e.ItemIndex
        Dim col As Integer = e.ColumnIndex
        Dim blockStartCol As Integer ' = 0
        Dim blockEndCol As Integer ' = 0

        ' Define the block of cells to color
        Dim blockStartRow As Integer = 0
        Dim blockEndRow As Integer = NumberOfRows - 1
        If EditStockSheet = True Then
            DoAssessment = False
            blockStartCol = 0
            blockEndCol = 3
        ElseIf DoAssessment = True Then
            EditStockSheet = False
            blockStartCol = 4
            blockEndCol = 5
        Else
            Exit Sub
        End If
        ' Apply background color if within block
        If row >= blockStartRow AndAlso row <= blockEndRow AndAlso col >= blockStartCol AndAlso col <= blockEndCol Then
            If LV_Stock.Items(row).SubItems(0).Text <> "" And EditStockSheet = False Then
                e.Graphics.FillRectangle(Brushes.LightGreen, e.Bounds)
            ElseIf EditStockSheet = True Then
                e.Graphics.FillRectangle(Brushes.LightGreen, e.Bounds)
            End If

        Else
            e.Graphics.FillRectangle(New SolidBrush(LV_Stock.BackColor), e.Bounds)
        End If

        ' Draw the text
        TextRenderer.DrawText(e.Graphics, e.SubItem.Text, LV_Stock.Font, e.Bounds, e.SubItem.ForeColor, TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)

    End Sub

    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'Cell editing handlers
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

    'Handle  Single and double Mouse Click on form
    Private Sub StockList_MouseClick(sender As Object, e As MouseEventArgs)
        clickedItem = LV_Stock.GetItemAt(e.X, e.Y)
        clickTimer.Start()
    End Sub
    Private Sub CellText(sender As Object, e As MouseEventArgs) ''DoubleClick
        clickTimer.Stop()
        Dim info = LV_Stock.HitTest(e.X, e.Y)
        Dim colIndex = info.Item.SubItems.IndexOf(info.SubItem)
        Dim selectedItem As ListViewItem = LV_Stock.SelectedItems(0)
        Dim rowIndex As Integer = selectedItem.Index
        EditCellText(rowIndex, colIndex)
    End Sub

    Public Sub EditCellText(Rowindex As Integer, colindex As Integer)
        Dim LastItem As Integer = 0
        Dim response As Integer = 0
        LV_Stock.Select()

        If colindex < 4 Then 'We've clicked in the Items name and details field
            EditStockSheet = True
            DoAssessment = False
            DoAssessmentToolStripMenuItem.Visible = True
            EditToolStripMenuItem.Visible = False
            NumberOfRows = LV_Stock.Items.Count
            Dim i As Integer = 0
            For i = 0 To NumberOfRows - 1
                For p = 0 To 7
                    LV_Stock.Items(i).SubItems(p).Text = LV_Stock.Items(i).SubItems(p).Text ' Triggers a sheet redraw
                    If LV_Stock.Items(i).SubItems(6).Text = "Total Yield" Then
                        LV_Stock.Items(i).SubItems(6).Text = ""
                        LV_Stock.Items(i).SubItems(7).Text = ""
                    End If
                Next
            Next
        ElseIf colindex = 4 Or colindex = 5 Then 'We've clicked in the stock issues or closing stock field
            EditStockSheet = False
            DoAssessment = True
            DoAssessmentToolStripMenuItem.Visible = False
            EditToolStripMenuItem.Visible = True
            NumberOfRows = LV_Stock.Items.Count
            Dim i As Integer = 0
            For i = 0 To NumberOfRows - 1
                For p = 0 To 7
                    LV_Stock.Items(i).SubItems(p).Text = LV_Stock.Items(i).SubItems(p).Text ' Triggers a sheet redraw
                Next
            Next

        End If

        For i = 0 To LV_Stock.Items.Count - 1
            If LV_Stock.Items(i).SubItems(0).Text <> "" Then
                LastItem += 1
            End If
        Next
        If Rowindex >= LastItem Then ' Ensure we can't add items too low in the list
            Rowindex = LastItem
        End If

        Dim q As Integer = InStr(CurrentPath, "Archive")
        If NewAssessment = False And q > 0 Then
            Dim fName As String = Replace(FileName, "csv", "bak")
            Dim Message As String = "Do you intend to modify it? " & Chr(13) & "If so, the modified file will be saved as " & Chr(13) & ArchiveFolder & fName

            response = MsgBox(Message, vbYesNo Or vbCritical Or MsgBoxStyle.DefaultButton2, "This Assessment has been Archived")
            If response = vbNo Then
                If editor IsNot Nothing Then
                    LV_Stock.Controls.Remove(editor)
                    editor.Dispose()
                End If
                Exit Sub
            Else
                NewAssessment = True
            End If
        End If

        With LV_Stock
            If .SelectedItems.Count > 0 Then

                If Rowindex < 0 Then Rowindex = 0
                If .Items(Rowindex).SubItems(0).Text = "" And Rowindex >= 0 And DoAssessment = True Then 'we are trying to select down the sheet
                    Beep()
                    Rowindex = 0 'Don't go below number of items
                End If
            End If
        End With

        If DoAssessment = False And EditStockSheet = False Then Exit Sub

        If NewAssessment <> True And DeptName = "" Then
            MsgBox("You must load or set up a Stock Sheet" & vbCr & "before you can edit it", vbOKOnly, "No Active Stock Sheet")
            Exit Sub
        End If

        If EditStockSheet = True Then
            DoAssessment = False
            EditColumnStart = 0
            EditColumnEnd = 3
        ElseIf DoAssessment = True Then
            EditStockSheet = False
            EditColumnStart = 4
            EditColumnEnd = 5
        End If

        If EditStockSheet = True Then
            If LV_Stock.Items(Rowindex).SubItems(0).Text = "" Then
                EditCell(Rowindex, 0)
            Else
                EditCell(Rowindex, colindex)
            End If
        ElseIf DoAssessment = True Then
            EditCell(Rowindex, colindex)
        End If

    End Sub

    Private Sub OnSingleClickConfirmed(sender As Object, e As EventArgs)
        clickTimer.Stop()
        DoAssessment = False
        EditStockSheet = False
        DoAssessmentToolStripMenuItem.Visible = True
        EditToolStripMenuItem.Visible = True

        If editor IsNot Nothing Then
            LV_Stock.Controls.Remove(editor)
            editor.Dispose()
        End If

        For Each item As ListViewItem In LV_Stock.Items
            ' Set the background color of the main item
            item.BackColor = Color.White
            ' Set the background color of all subitems
            For Each subitem As ListViewItem.ListViewSubItem In item.SubItems
                subitem.BackColor = Color.White
            Next
        Next
        ' Refresh the ListView to apply changes
        LV_Stock.Refresh()

    End Sub

    'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ


    Public Sub EditCell(rowIndex As Integer, colIndex As Integer)
        Dim bounds As Rectangle = LV_Stock.Items(rowIndex).SubItems(colIndex).Bounds
        ' Remove previous editor if any
        If editor IsNot Nothing Then
            LV_Stock.Controls.Remove(editor)
            editor.Dispose()
        End If

        If colIndex = 0 Then
            bounds = LV_Stock.Items(rowIndex).GetBounds(ItemBoundsPortion.Label)
        End If

        ' Create a new TextBox
        editor = New TextBox()
        editor.Bounds = bounds
        editor.BackColor = ColourWhite

        editor.Text = LV_Stock.Items(rowIndex).SubItems(colIndex).Text
        LV_Stock.Controls.Add(editor)
        editor.Focus()
        editor.Select(editor.Text.Length, 0)

        ' Handle Tab for moving to next cell
        AddHandler editor.KeyDown, AddressOf Editor_KeyDown

        currentRow = rowIndex
        currentCol = colIndex
    End Sub
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'KeyPress handlers
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

    Private Sub Form1_MouseWheel(sender As Object, e As MouseEventArgs) Handles LV_Stock.MouseWheel
        If editor IsNot Nothing Then
            LV_Stock.Controls.Remove(editor)
            editor.Dispose()
        End If
    End Sub

    'Turn the NumLock on if it isn't already on

    <DllImport("user32.dll")>
    Public Shared Function GetKeyState(ByVal nVirtKey As Integer) As Short
    End Function

    <DllImport("user32.dll")>
    Public Shared Sub keybd_event(ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As UInteger, ByVal dwExtraInfo As UIntPtr)
    End Sub
    Private Sub ForceNumLockOn() ' Needed to use number keypad
        Dim keyState As Boolean = (GetKeyState(VK_NUMLOCK) And 1) = 1
        If Not keyState Then
            ' Num Lock is OFF, simulate key press to turn it ON
            keybd_event(VK_NUMLOCK, 0, KEYEVENTF_EXTENDEDKEY, UIntPtr.Zero)
            keybd_event(VK_NUMLOCK, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, UIntPtr.Zero)
        End If
    End Sub

    'Handle TAB key for moving between cells
    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean

        If editor IsNot Nothing Then
            If keyData = Keys.Tab Then
                ' Handle Tab key and Enter key
                If editor.Text = "" And currentCol <> 1 And currentCol <> 3 Then
                    Beep()

                End If
                nextRow = currentRow + 1
                SaveAndMoveNext()
            End If
        End If

        Return False

    End Function

    Private Sub Editor_KeyDown(sender As Object, e As KeyEventArgs)

        Dim kCode As Integer = e.KeyCode
        ' Dim Txt = (LV_Stock.Items(currentRow).SubItems(currentCol).Text)

        If kCode = 8 Then Exit Sub 'Back key

        If DoAssessment = True Then
            'numbers keys on top row and keypad
            If kCode < 48 Or kCode > 57 AndAlso kCode < 96 Or kCode > 105 AndAlso kCode <> 110 AndAlso kCode <> 190 Then ' Can only accept numbers and  and .
                e.SuppressKeyPress = True
            End If
        ElseIf EditStockSheet = True Then
            If currentCol > 1 Then 'First 2 columns are AlphaNumeric
                If kCode < 48 Or kCode > 57 AndAlso kCode < 96 Or kCode > 105 AndAlso kCode <> 110 AndAlso kCode <> 190 Then ' Can only accept numbers and  and .
                    e.SuppressKeyPress = True
                End If
            End If

        End If

        If kCode = Keys.Enter Then
            If currentCol = 1 And editor.Text = "" Then
                If LV_Stock.Items(currentRow - 1).SubItems(1).Text <> "" Then
                    editor.Text = LV_Stock.Items(currentRow - 1).SubItems(1).Text
                End If
            End If

            If editor.Text = "" And currentCol <> 1 And currentCol <> 3 Then
                Beep()
                e.SuppressKeyPress = True
                Exit Sub
            End If
            e.SuppressKeyPress = True
            nextRow = currentRow + 1
            SaveAndMoveNext()
        End If

    End Sub


    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'Assessment Calculation
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

    Public Sub CalculateYield()
        If EditStockSheet = True Then Exit Sub
        Dim Stock_Value As Single = 0
        Dim Stock_Opening As Single = 0
        Dim Stock_Issues As Single = 0
        Dim Stock_Closing As Single = 0
        Dim Stock_Sold As Single = 0
        Dim Stock_Yield As Single = 0
        Dim Row As Integer = 0
        Dim RowIndex As Integer = 0
        Dim ItemCount As Integer = 0
        Dim TotalYield As Single = 0
        Dim NrOfItems = LV_Stock.Items.Count

        With LV_Stock
            Try
                For RowIndex = 0 To NrOfItems - 1
                    If .Items(RowIndex).SubItems(0).Text <> "" Then ' Exclude any empty rows
                        ItemCount += 1
                        Dim sYield As String = .Items(RowIndex).SubItems(2).Text
                        sYield = Replace(sYield, "$", "")
                        sYield = Replace(sYield, ",", "") 'Strip off the $ and any commas

                        Stock_Value = Val(sYield)
                        If (Stock_Value) = 0 Then
                            Dim iName As String = .Items(RowIndex).SubItems(0).Text
                            MsgBox("Please enter an Item Value", vbOK Or vbCritical, "There is no Value entered for " & UCase(iName))
                            EditCell(RowIndex, 1)
                            Exit Sub
                        End If

                        Stock_Opening = Val(.Items(RowIndex).SubItems(3).Text)
                        Stock_Issues = Val(.Items(RowIndex).SubItems(4).Text)
                        Stock_Closing = Val(.Items(RowIndex).SubItems(5).Text)

                        If Stock_Closing <> 0 Then
                            Stock_Sold = (Stock_Opening + Stock_Issues) - Stock_Closing
                            Stock_Yield = Stock_Sold * Stock_Value
                            Dim Sold As Single = Math.Floor(Stock_Sold * 100) / 100
                            Stock_Sold = Math.Round(Sold, 1)
                            .Items(RowIndex).SubItems(6).Text = Stock_Sold
                            .Items(RowIndex).SubItems(7).Text = (Stock_Yield).ToString("C")
                        End If
                        If Stock_Yield <> 0 Then 'check if there is closing stock
                            TotalYield += Stock_Yield
                            Stock_Yield = 0
                        End If

                    End If

                    If Stock_Sold < 0 Then 'warn if negative
                        Beep()
                    End If
                Next

            Catch ex As Exception
                MsgBox(ex.Message, vbCritical, "Error at Row " & RowIndex)
            End Try
        End With

        LV_Stock.Items(ItemCount + 1).SubItems(6).Text = "Total Yield"
        LV_Stock.Items(ItemCount + 1).SubItems(7).Text = TotalYield.ToString("C")

    End Sub


    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    'Tool Strip Items handlers
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    Private Sub AutoSaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AutoSaveToolStripMenuItem.Click
        If AutoSave = True Then
            AutoSave = False
            SheetSaved = False 'Cancel this if parts of the sheet have been saved previously
            AutoSaveToolStripMenuItem.Text = "Auto Save On"
            SaveAsToolStripMenuItem.Visible = True
            SaveAsSeparator.Visible = True
            SaveToolStripMenuItem.Visible = True
            SaveSeparator.Visible = True
        Else
            AutoSave = True
            AutoSaveToolStripMenuItem.Text = "Auto Save Off"
            SaveAsToolStripMenuItem.Visible = False
            SaveAsSeparator.Visible = False
            SaveToolStripMenuItem.Visible = False
            SaveSeparator.Visible = False
        End If
    End Sub

    Private Sub Assessment(sender As Object, e As EventArgs) Handles DoAssessmentToolStripMenuItem.Click
        If editor IsNot Nothing Then
            LV_Stock.Controls.Remove(editor)
            editor.Dispose()
        End If

        DoAssessmentToolStripMenuItem.Visible = False
        EditToolStripMenuItem.Visible = True
        DoAssessSeparator.Visible = False
        EditSeparator.Visible = True
        DoAssessment = True
        EditStockSheet = False

        NumberOfRows = LV_Stock.Items.Count
        Dim i As Integer ' = 0
        For i = 0 To NumberOfRows - 1
            For p = 0 To 7
                LV_Stock.Items(i).SubItems(p).Text = LV_Stock.Items(i).SubItems(p).Text ' Triggers a sheet redraw
            Next
        Next
        EditCell(0, 4)

    End Sub
    Private Sub EditSheet(sender As Object, e As EventArgs) Handles EditToolStripMenuItem.Click
        If editor IsNot Nothing Then
            LV_Stock.Controls.Remove(editor)
            editor.Dispose()
        End If

        EditStockSheet = True
        DoAssessment = False
        DoAssessmentToolStripMenuItem.Visible = True
        DoAssessSeparator.Visible = True
        EditToolStripMenuItem.Visible = False
        EditSeparator.Visible = False
        If CurrentPath <> "" Then
            Edit_Assessment = True
        End If
        ' Dim i As Integer' = 0
        For i = 0 To LV_Stock.Items.Count - 1
            For p = 0 To 7
                LV_Stock.Items(i).SubItems(p).Text = LV_Stock.Items(i).SubItems(p).Text ' Triggers a sheet redraw
            Next
        Next

        EditCell(0, 0)
    End Sub
    Private Sub ZoomInToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ZoomInToolStripMenuItem.Click
        If editor IsNot Nothing Then
            LV_Stock.Controls.Remove(editor)
            editor.Dispose()
        End If

        If Zoom <= 5 Then
            Form_Zoom.ZoomIn(sender, e, Me, LV_Stock)
            Zoom += 1
        End If

    End Sub

    Private Sub ZoomOutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ZoomOutToolStripMenuItem.Click
        If editor IsNot Nothing Then
            LV_Stock.Controls.Remove(editor)
            editor.Dispose()
        End If

        If Zoom >= 1 Then
            Form_Zoom.ZoomOut(sender, e, Me, LV_Stock)
            Zoom -= 1
        End If

    End Sub

    Public Sub OpenNewDepartment(sender As Object, e As EventArgs) Handles NewDepartmentToolStripMenuItem.Click
        New_StockSheet.Show()
        New_StockSheet.BringToFront()
    End Sub
    Sub Save(s_ender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        If editor IsNot Nothing Then
            LV_Stock.Controls.Remove(editor)
            editor.Dispose()
        End If
        SaveStockList("Save", CurrentPath)
    End Sub

    Private Sub SaveAs(s_ender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        If editor IsNot Nothing Then
            LV_Stock.Controls.Remove(editor)
            editor.Dispose()
        End If
        SaveStockList("SaveAs", CurrentPath)
    End Sub


    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        If editor IsNot Nothing Then
            LV_Stock.Controls.Remove(editor)
            editor.Dispose()
        End If
        If AutoSave = True And CurrentPath <> "" Then
            SaveStockList("Exit", CurrentPath)
        End If

        If SheetSaved = False And CurrentPath <> "" Then
            Dim Response As Integer

            Response = MsgBox("DO YOU WANT TO SAVE?", vbYesNoCancel Or vbCritical, "ASSESSMENT HAS NOT BEEN SAVED")
            If Response = vbYes Then
                AutoSave = True
                SaveStockList("Exit", CurrentPath)
            ElseIf Response = vbCancel Then
                Exit Sub
            End If
        End If

        SaveSheetConfig(SheetConfigPath)
        Me.Close()
        Me.Dispose()

    End Sub

    Private Sub PrintToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintToolStripMenuItem.Click
        PrintAssessmentToolStripMenuItem.Visible = True
        PrintWorkSheetToolStripMenuItem.Visible = True
    End Sub

    Private Sub AssessmentToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintAssessmentToolStripMenuItem.Click
        If FileToPrint <> "" Then
            Print_Assessment.PrintSheet("Assessment", FileToPrint)
        End If
    End Sub

    Private Sub WorkSheetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintWorkSheetToolStripMenuItem.Click
        If FileToPrint <> "" Then
            Print_Assessment.PrintSheet("WorkSheet", FileToPrint)
        End If
    End Sub

    Private Sub ManualToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ManualToolStripMenuItem.Click
        Me.Hide()
        Manual.Show()
    End Sub

    Private Sub ScreenBackground(sender As Object, e As EventArgs) Handles ScreenBackgroundToolStripMenuItem.Click
        Dim fPath As String = ""

        Try
            OpenFileDialog1.InitialDirectory = "c:\"
            OpenFileDialog1.Filter = "Photo Files (*.jpg)|*.jpg|All Files (*.*)| *.*"
            OpenFileDialog1.FileName = "Choose a Background Image."
            OpenFileDialog1.Multiselect = False
            OpenFileDialog1.RestoreDirectory = True

            Dim response = OpenFileDialog1.ShowDialog()
            If response = vbOK Then
                FileName = OpenFileDialog1.FileName
                If System.IO.File.Exists(OpenFileDialog1.FileName) Then
                    fPath = FileName
                    ' Copy the Image to the Config folder
                    Dim ImagePath As String = DefaultConfigPath & "Background.pic"
                    FileCopy(fPath, ImagePath)
                End If
            End If

            Dim BmMain = New Bitmap(fPath) 'Convert the file to a bitmap and  Display the results.
            Me.BackgroundImage = BmMain

        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, "Error loading Picture")
        End Try

        Close_Click(sender, e)

    End Sub

    Private Sub Close_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        LV_Stock.Hide()
        AutoSave = True
        ZoomToolStripMenuItem.Visible = False
        CloseToolStripMenuItem.Visible = False
        SaveToolStripMenuItem.Visible = False
        SaveAsToolStripMenuItem.Visible = False
        DoAssessmentToolStripMenuItem.Visible = False
        EditToolStripMenuItem.Visible = False
        SaveSeparator.Visible = False
        SaveAsSeparator.Visible = False
        EditSeparator.Visible = False
        DoAssessSeparator.Visible = False
        ExitSeparator.Visible = False
        ToolStripSeparator6.Visible = False
        PrintToolStripMenuItem.Visible = False
        SettingsToolStripMenuItem.Visible = False
        Dim Fields() = Split(FormSize, ",")
        Me.Width = Fields(0) ' Reset the form size
        Me.Height = Fields(1)
        Me.Text = "Stock Assessment"
    End Sub

    Private Sub AboutStockAssessment_Click(sender As Object, e As EventArgs) Handles AboutStockAssessmentToolStripMenuItem.Click
        Dim Message As String = "Please report any bugs or errors to the author." & vbCrLf & "CamSoft, Australia:  CamsoftAu@gmail.com"
        MsgBox(Message, vbOKOnly, "Stock Assessment is Freeware.")
    End Sub
End Class
