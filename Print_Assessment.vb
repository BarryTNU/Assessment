
Imports System.Drawing.Printing
Imports System.IO

Public Class Print_Assessment

    ' At the class level
    Public linesToPrint As List(Of String)
    Private currentLineIndex As Integer = 0
    Private PageHeader As String = ""

    ' File-related variables ...
    Private RootFolder As String = form1.rootfolderpath
    Dim ItemNumber As Integer
    Private PrintTarget As String

    ' Printing-related variables..
    Private currentCharIndex As Integer = 0
    Private LineCount As Integer = 0
    Private PageCount As Integer = 0
    Private SheetStyle As String
    Private lineIndex As Integer = 0
    Private PageToPrint As String = ""
    Private Separator As String = ""
    Private Data() As String
    Private Item As String = ""

    'Stock data
    Private Item_Name As String = ""
    Private Item_Size As String = ""
    Private Item_cost As String = ""
    Private Item_OpeningStock As String = ""
    Private Item_Issues As String = ""
    Private Item_ClosingStock As String = ""
    Private Item_Sold As String = ""
    Private Item_Yield As String = ""
    Private TotalYield As Single = 0


    Public Sub PrintSheet(Style As String, FileToPrint As String)
        SheetStyle = Style
        PrintTarget = FileToPrint
        lineIndex = 0
        PrintDocument1.DefaultPageSettings.Landscape = False
        PrintDocument1.PrintController = New StandardPrintController
        'Set the Print area' This is for A4 paper
        PrintDocument1.DefaultPageSettings.Margins = New Printing.Margins(10, 10, 30, 30) ' Left, Right, Top, Bottom
        PrintDoc()
    End Sub
    Sub PrintDoc()
        Dim LinesPerPage As Integer = 40
        Dim ItemCount As Integer = File.ReadAllLines(PrintTarget).Length

        PageToPrint = ""
        LineCount = 0

        PageHeader = ""
        LineCount = 0
        'Set up Page Headings

        If SheetStyle = "Assessment" Then
            Separator = ("----------------------------------------------------------------------")
            PageHeader += Environment.NewLine & "Print Time: " & Format(TimeOfDay, "hh:mm:ss").PadRight(15) & "STOCK ASSESSMENT FOR " & UCase(Form1.DeptName)
            PageHeader += Environment.NewLine & ("")
            PageHeader += Environment.NewLine & "ITEM".PadRight(15) & "SIZE".PadRight(8) & "VALUE".PadRight(8) & "OPENING".PadRight(8) & "ISSUES".PadRight(8) & "CLOSING".PadRight(8) & "SOLD".PadRight(8) & "YIELD"
            PageHeader += Environment.NewLine & Separator
        ElseIf SheetStyle = "WorkSheet" Then
            Separator = ("------------------------------------------------------------")
            PageHeader += Environment.NewLine & "Print Time: " & Format(TimeOfDay, "hh:mm:ss").PadRight(10) & "ASSESSMENT WORKSHEET FOR " & UCase(Form1.DeptName)
            PageHeader += Environment.NewLine & ("")
            PageHeader += Environment.NewLine & "ITEM".PadRight(15) & "SIZE".PadRight(10) & "VALUE".PadRight(10) & "OPENING".PadRight(8) & "|" & "ISSUES".PadRight(7) & "|" & "CLOSING".PadRight(6) & "|"
            PageHeader += Environment.NewLine & Separator
        ElseIf SheetStyle = "Manual" Then
            Separator = ("----------------------------------------------------------------------")
            PageHeader += Environment.NewLine & "Print Time: " & Format(TimeOfDay, "hh:mm:ss").PadRight(15) & "MANUAL FOR STOCK ASSESSMENT"
            PageHeader += Environment.NewLine & ("")
            PageHeader += Environment.NewLine & Separator
        End If

        'Start setting up the page

        ItemNumber = 0


        Using fReader As New StreamReader(PrintTarget)

            Do While ItemNumber < ItemCount 'Not fReader.EndOfStream
                Item = fReader.ReadLine
                ItemNumber += 1
                Data = Split(Item, ",") 'Split the .CSV file into  fields
                Dim q = UBound(Data)
                If q <> 7 Then Exit Sub

                Try
                    Item_Name = Data(0)
                    Item_Size = Data(1)
                    Item_cost = "$" & Data(2)
                    Item_OpeningStock = Data(3)
                    Item_Issues = Data(4)
                    Item_ClosingStock = Data(5)
                    Item_Sold = Data(6)
                    Item_Yield = "$" & Data(7)
                    Dim Yield = Data(7)

                    If Yield <> Nothing Then
                        TotalYield += Yield
                    End If

                    If SheetStyle = "Assessment" And Item_Name <> Nothing Then
                        PageToPrint += Environment.NewLine & (Item_Name.PadRight(15) & Item_Size.PadRight(8) & Item_cost.PadRight(8) & Item_OpeningStock.PadRight(8) & Item_Issues.PadRight(8) & Item_ClosingStock.PadRight(8) & Item_Sold.PadRight(8) & Item_Yield)
                        PageToPrint += Environment.NewLine & Separator
                        LineCount += 2
                    ElseIf SheetStyle = "WorkSheet" And Item_Name <> Nothing Then
                        PageToPrint += Environment.NewLine & (Item_Name.PadRight(15) & Item_Size.PadRight(10) & Item_cost.PadRight(10) & Item_OpeningStock.PadRight(8) & "|".PadRight(8) & "|".PadRight(8)) & "|"
                        PageToPrint += Environment.NewLine & Separator
                        LineCount += 2
                    ElseIf SheetStyle = "Manual" Then
                        PageToPrint = printtarget
                    End If

                Catch ex As Exception
                    MsgBox(ex.Message, vbCritical, "Error reading file")
                End Try

                If ItemNumber >= ItemCount And SheetStyle = "Assessment" Then
                    'Print the Total Yield at the bottom of the page
                    PageToPrint += Environment.NewLine & "" 'Print a blank line
                    PageToPrint += Environment.NewLine & Space(48) & "Total Yield  " & TotalYield.ToString("C")
                    'Make sure we start next Print with a clean sheet
                    Item_Name = ""
                    Item_Size = ""
                    Item_cost = ""
                    Item_OpeningStock = ""
                    Item_Issues = ""
                    Item_ClosingStock = ""
                    Item_Sold = ""
                    Item_Yield = ""
                    TotalYield = 0
                End If
            Loop

            ' Split all lines into a list for paged printing
            linesToPrint = PageToPrint.Split(ControlChars.Lf).ToList()
            currentLineIndex = 0

            fReader.Close()
            ' Send the string to the printer
            PrintDocument1.Print()
        End Using

    End Sub


    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim PrintFont As New Font("Courier New", 12, FontStyle.Bold)
        Dim lineHeight As Single = PrintFont.GetHeight(e.Graphics)
        Dim x As Single = e.MarginBounds.Left
        Dim y As Single = e.MarginBounds.Top
        Dim b As Single = e.MarginBounds.Bottom
        PageCount += 1
        Dim PageTitle As String = Environment.NewLine & "Print Date: " & Format(Today, "dd/MM/yyyy").PadRight(20) & "Page: " & CStr(PageCount).PadLeft(3)
        Dim Header As String = PageTitle & vbLf & PageHeader
        ' Draw page header
        Dim HeaderFont As New Font("Courier New", 12, FontStyle.Bold)

        e.Graphics.DrawString(Header, HeaderFont, Brushes.Black, x, y)

        y += lineHeight * 8 ' Extra space below header

        ' Print lines until we run out of vertical space
        While currentLineIndex < linesToPrint.Count
            If y + lineHeight > e.MarginBounds.Bottom Then
                e.HasMorePages = True
                Exit Sub
            End If

            e.Graphics.DrawString(linesToPrint(currentLineIndex), printFont, Brushes.Black, x, y)
            y += lineHeight
            currentLineIndex += 1
        End While

        ' If we got here, no more pages
        e.HasMorePages = False
        PageCount = 0
    End Sub


End Class