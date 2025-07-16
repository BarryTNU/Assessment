
Imports System.Drawing.Printing
Imports System.IO
Imports System.Windows.Controls
'Imports System.Windows.Controls
'Imports System.Windows.Documents
'Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Public Class Manual

    Public FormConfigPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "AConfig\")
    Public Zoom As Integer = 0
    Public Text_Changed As Boolean = False


    ReadOnly ColourBlack = System.Drawing.Color.Black
    ReadOnly ColourRed = System.Drawing.Color.Red
    ReadOnly ColourGreen = System.Drawing.Color.LawnGreen
    ReadOnly ColourAmber = System.Drawing.Color.Orange
    ReadOnly ColourWhite = System.Drawing.Color.White
    Public Sub Manual_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadManual()
        TextBox.ReadOnly = True
        ToolStripMenuItem1.Visible = False

        Lst_FontStyle.Visible = False
        Lst_FontSize.Visible = False


        For Each f As FontFamily In FontFamily.Families
            Lst_FontStyle.Items.Add(f.Name)
        Next

        ' Add common font sizes
        For size As Integer = 8 To 30 Step 2
            Lst_FontSize.Items.Add(size.ToString())
        Next

        'Center form onscreen
        Dim screenBounds As Rectangle = Screen.PrimaryScreen.WorkingArea
        Me.Left = (screenBounds.Width - Me.Width) \ 2
        Me.Top = (screenBounds.Height - Me.Height) \ 2
        Me.Show()

    End Sub


    Sub ZoomInToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ZoomInToolStripMenuItem.Click
        Zoom += 5
        If Zoom > 50 Then Zoom = 50
        Zoom_TextBox.ZoomIn(Zoom)
    End Sub

    Sub ZoomOutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ZoomOutToolStripMenuItem.Click
        Zoom -= 5
        If Zoom < 0 Then Zoom = 0
        Zoom_TextBox.ZoomOut(Zoom)
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        If Text_Changed = True Then
            SaveManual()
            Text_Changed = False
        End If
        Form1.Show()
        Me.Dispose()
    End Sub

    Private Sub Save_Manual(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        SaveManual()
    End Sub
    Private Sub SaveManual()
        Dim fPath As String = FormConfigPath & "Manual\Manual.txt"
        Dim mText As String = TextBox.Text
        Form1.CheckFolderExists(fPath)

        Try
            Using fwriter As New System.IO.StreamWriter(fPath)
                fwriter.WriteLine(mText)
                fwriter.Close()
            End Using
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, "Error Saving Manual")
        End Try

        Text_Changed = False

    End Sub
    Private Sub LoadManual()
        Dim fPath As String = FormConfigPath & "Manual\Manual.txt"
        Dim mText As String = TextBox.Text
        Form1.CheckFolderExists(fPath)

        Try
            Using fReader As New StreamReader(fPath)
                Do While Not fReader.EndOfStream
                    mText = fReader.ReadToEnd
                Loop
                TextBox.Text = mText
                fReader.Close()
            End Using
        Catch ex As Exception ' We will end up here if the Manual hasn't been saved to the Config folder; So we load the included text file
            ' MsgBox(ex.Message, vbCritical, "Error loading Manual")
            TextBox.Text = Manual() ' Get the Manual text
            SaveManual() ' and save it on the Config folder so we can edit it if required
        End Try

    End Sub

    Private Sub TextBox_DoubleClick(sender As Object, e As MouseEventArgs) Handles TextBox.MouseDoubleClick

        If TextBox.ReadOnly = True Then
            TextBox.ReadOnly = False
            Me.BackColor = ColourGreen
            ToolStrip1.Visible = True
            TextBox.Height = TextBox.Height - (ToolStrip1.Height + 2)
        Else
            TextBox.ReadOnly = True
            Me.BackColor = ColourWhite
            ToolStrip1.Visible = False
            TextBox.Height = TextBox.Height + (ToolStrip1.Height + 2)
        End If
        TextBox.Select()
        SaveManual()
    End Sub

    Private Sub Mouse_Click(sender As Object, e As MouseEventArgs) Handles TextBox.MouseClick
        Lst_FontStyle.Visible = False
        Lst_FontSize.Visible = False
    End Sub
    Private Sub BtnBold_Click(sender As Object, e As EventArgs) Handles Btn_Bold.Click
        Dim currentFont As Font = TextBox.SelectionFont
        If currentFont IsNot Nothing Then
            Dim newFontStyle As FontStyle = currentFont.Style Xor FontStyle.Bold
            TextBox.SelectionFont = New Font(currentFont, newFontStyle)
        End If
    End Sub
    Private Sub BtnItalic_Click(sender As Object, e As EventArgs) Handles Btn_Italic.Click
        Dim currentFont = TextBox.SelectionFont
        If currentFont IsNot Nothing Then
            Dim newFontStyle = currentFont.Style Xor FontStyle.Italic
            TextBox.SelectionFont = New Font(currentFont, newFontStyle)
        End If
    End Sub

    Private Sub BtnUnderline_Click(sender As Object, e As EventArgs) Handles Btn_Underline.Click
        Dim currentFont As Font = TextBox.SelectionFont
        If currentFont IsNot Nothing Then
            Dim newFontStyle As FontStyle = currentFont.Style Xor FontStyle.Underline
            TextBox.SelectionFont = New Font(currentFont, newFontStyle)
        End If
    End Sub
    Private Sub Btn_Font_Click(sender As Object, e As EventArgs)
        Lst_FontStyle.Visible = True
        Lst_FontStyle.SelectedItem = TextBox.Font
        Lst_FontSize.Visible = True
    End Sub
    Private Sub Select_Font(sender As Object, e As EventArgs) Handles Lst_FontStyle.SelectedIndexChanged
        Try
            Dim newFont As String = Lst_FontStyle.SelectedItem
            Dim FontSize = TextBox.SelectionFont.Size
            Dim Font_style = TextBox.SelectionFont.Style
            TextBox.SelectionFont = New Font(newFont, FontSize, Font_style)
            Lst_FontStyle.Visible = False
        Catch ex As Exception

            MsgBox(ex.Message, vbCritical, "Error Selecting Font Style")
        End Try
    End Sub
    Private Sub Cmb_FontSize_Click(sender As Object, e As EventArgs) Handles Lst_FontSize.SelectedIndexChanged
        Try

            Dim Font As String = TextBox.SelectionFont.Size
            Dim FontSize As Single = Lst_FontSize.SelectedItem
            Dim Font_style = TextBox.SelectionFont.Style
            TextBox.SelectionFont = New Font(Font, FontSize, Font_style)
            Lst_FontSize.Visible = False

        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, "Error Selecting Font Size")
        End Try

    End Sub

    Private Sub TextBox_TextChanged(sender As Object, e As EventArgs) Handles TextBox.TextChanged
        Text_Changed = True
    End Sub

    Private Sub FontStyle_Click(sender As Object, e As EventArgs) Handles Lbl_FontStyle.Click
        Lst_FontStyle.Visible = True
    End Sub

    Private Sub FontSize_Click(sender As Object, e As EventArgs) Handles Lbl_FontSize.Click
        Lst_FontSize.Visible = True
    End Sub

    'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ
    'Print Routine
    'ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ
    ' This variable will hold the text to be printed.
    Private textToPrint As String
    Private currentCharIndex As Integer = 0

    Private Sub PrintButton_Click(sender As Object, e As EventArgs) Handles PrintToolStripMenuItem.Click
        textToPrint = TextBox.Text
        currentCharIndex = 0
        PrintDocument1.Print()
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim font As New Font("Arial", 12)
        Dim printArea As RectangleF = e.MarginBounds
        Dim charsFitted As Integer
        Dim linesFilled As Integer

        ' Draw the text page by page
        e.Graphics.MeasureString(textToPrint.Substring(currentCharIndex), font, printArea.Size, StringFormat.GenericDefault, charsFitted, linesFilled)
        e.Graphics.DrawString(textToPrint.Substring(currentCharIndex), font, Brushes.Black, printArea)

        currentCharIndex += charsFitted

        ' More pages?
        If currentCharIndex < textToPrint.Length Then
            e.HasMorePages = True
        Else
            e.HasMorePages = False
        End If
    End Sub

    Function Manual()
        Dim Txt As String = ""
        Txt = "STOCK  ASSESSMENT  INSTRUCTIONS

Stock Assessment has been written to handle assessing stock in locations such as Bars, Restaurants, etc, which sell stock items in portions.
However it will work for any business with similar requirements.  
For example shops which sell fabric, and have to assess part bolts of cloth, 
or engineering shops, which might sell part of a length of steel. 
In fact any organisation which does regular stock takes, and needs to assess the value of stock sold.

When Stock Assessment is first loaded you will see a screen with just two items.
FILE  and  MANUAL
Clicking MANUAL will bring up this document, which should be read and understood before proceeding.

After you are familiar with how Stock Assessment works Click FILE.

There are two options
OPEN, and EXIT
Click OPEN, and select DEPARTMENT.

The Department form opens, with the cursor on the Department Name field.
Enter the name of the department. 
In a hotel, this will be 'Pool Bar', or 'House Bar', or 'Public Bar', etc, but is can be anything which allows you to distinguish this  department from other departments you may have.

By default the File Location is set to 'C:\ProgramData\Assessement\'.
You should leave it as this, unless you have a compelling reason to change it; for example if you wish to store the files on a server, or in the cloud.

As this is the first Department,  a  blank Stock sheet will be created and saved.
Click SAVE, and the form will close and you will return to the Assessment program with the Stock sheet open, ready for you to start entering data.

Don't enter too many stock items at this stage, just a few to try out the program and see if it suits your needs.  More can be added at any time.


The Stock Item Name can include detail's such as size, quality, etc. It is completely free form.
The Size field is optional; it's not included in any calculations etc.
Note:  If you do not enter a Size, and the previous size field is not empty, this previous value will be copied to the current size field.  If you want to leave a blank field, just enter a Space. 
The Value field is required.  Note that this is what you expect to receive from the stock item when it is sold;  it is not the item cost.
The Opening stock can be entered now, if you have this information, or it can be done at the Assessment time.

As you enter stock items the sheet will grow to accommodate the number of items.
This is set at ten initially, and as you near ten items, additional rows will be added.  There is no functional limit to the number of items the Assessment will handle, but from a practical point of view, the number of items should be kept to what you can assess in a short period, say under one hour.  Create additional departments if you have a large number of stock items.

Give some serious thought to this before you start, as it's time consuming to change later.

Having set up your first department, you should now close the program and start again, as though you are doing an assessment.

Click FILE, and EXIT.

Having restarted the program, Click FILE / OPEN.

You now have two choices,  ASSESSMENT, and DEPARTMENT. click ASSESSMENT, and you will see a file selector dialog, with the last Assessment selected. Accept this by clicking Open.

Note that there is just one file in the folder, which has the date of the last assessment as it's name. There will only ever by one, or on  rare occasions two, files in this folder. Every time the sheet is saved all files except the current one are moved to the Archive folder.  Click on this folder to see what files it contains.  You can reload any of these files to view them, however you cannot change a file which has been archived.  Instead, if you edit an archived file, it will be saved as a *.bak file, and the original will remain unchanged.

Note: If the file name is today's date, it is assumed you are continuing a previously started assessment.  If the file name is a previous date, it is possible that you are either continuing a previous assessment, or starting a new one. In this case you are asked if you intend to start a New Assessment. Clicking Yes will roll the Assessment over, moving the Closing stock to the Opening stock, with the Issues and Closing fields blank.
Note also: That you cannot roll over today's assessment. The only way to roll over is to wait until the next day, or the next week or month.

You now have the stock sheet ready to either add more stock items, or start an assessment.

Before you can do an assessment you must 'Assess' your stock. ie, Find how many stock items have been issued, and how many are remaining.  To assist you in this you can print a Work Sheet.  Click PRINT, and select WORK SHEET. 

This is a list of all stock items and their details, with blank spaces for you to add Stock Issues  and Stock Closing (Remaining).
The Stock Closing can be a whole number, or a decimal.  In a bar, stock is assessed in tenths (ie 0.4), so the Stock Closing could be 3.4 for example.

Having Assessed your stock, you are ready to  enter the data.

With the stock sheet open, click FILE / DO ASSESSMENT.

The Issues and Closing cells are highlighted, with the first item ready to start entering data.  As you enter data and press 'Enter' or 'Tab' the sheet is saved, and the edit field moves to the next cell. This continues until all the amounts have been entered.

Note: You can change to Stock Item editing or Assessment mode by double clicking on the sheet. A double click in any of the first 4 cells will start Editing mode, and double clicking in the Issues or Closing cells will start Assessment mode. Also, once you are in either mode, double clicking in a cell will move focus to that cell, and open the editor.

The editor is selective as to which characters in will accept. In the Item Name and Size fields it will accept letters and numbers In the other fields it wil only accept numbers, and the decimal point.
In all fields it will accept Enter, Tab, and Back Space.

If you make a mistake entering data, you can correct it by double clicking on the cell, and re-entering the data.
Note: The contents of the cell are not automatically deleted. You must delete the old content before adding new.

As you enter Issues and Closing stock the sheet is saved, and the stock sold and the Yield is calculated.  The total yield is shown at the bottom of the List. This should be compared with the return from the sale of the stock to determine if stock is being stolen, sold too cheaply, or if there is some other problem in the department.


Additional features

ZOOM. The form can be zoomed in or out to make it easier to read or enter data.
The Manual can also be Zoomed.

SAVING. By default the sheet will AutoSave as data is entered.  You can turn AutoSave off by clicking SETTINGS / AUTOSAVE OFF. You will then have the option to SAVE or SAVE AS. This could be useful if you want to, for example, save the sheet to a thumb drive or other location.

If AutoSave is Off, and you try to exit without saving, you will be given the option to Save.

All settings in Stock Assessment are 'Sticky'. So if AutoSave is Off when you Exit, it will be Off when you resume. Also your Zoom settings are retained, and the form will be the same size when you resume.

The last department assessed is the default department when you resume, and can be loaded by simply clicking Open on the File Dialog.

If you wish to assess a different department you should choose it from the File Dialog. This will then become the default department when you resume.

The Manual is also a rudimentary Text Editor.  Double clicking on the form will put it into edit mode.  You can add comments, or change the Manual to suit yourself.  You can't change the save location. The editor doesn't automatically save; you must save before exiting.

You can send the manual to your default printer by clicking PRINT


It has been shown that in a liquor sales scenario in particular, the only way to prevent loss of income is to do regular assessments.  The staff should be aware that assessments are done, and the results of each assessment should be communicated to the staff.  

(The assessment in fact doesn't need to be all that accurate. Just the fact that the staff know it is being done is enough to keep them honest.)"

        Return Txt
    End Function


End Class