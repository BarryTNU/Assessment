Imports System.Drawing.Printing

Public Class MultiPagePrint


    ' Class-level variables
    Private fullText As String              ' Entire text to be printed
    Private currentCharIndex As Integer     ' Index of the current character being printed

    ' Call this to start printing
    Public Sub StartPrinting()
        ' Simulate a long text for demonstration
        fullText = New String("This is line of sample text. ", 500) ' repeat string to simulate long content
        currentCharIndex = 0

        ' PrintDocument1.Print()
    End Sub

    ' Handles the printing of each page
    Private Sub PrintDocument1_PrintPage(sender As Object, e As PrintPageEventArgs) ' Handles PrintDocument1.PrintPage
        Dim font As New Font("Arial", 12)
        Dim layoutArea As RectangleF = e.MarginBounds
        Dim stringFormat As New StringFormat()

        ' Measure how many characters fit on the page
        Dim charsFitted As Integer
        Dim linesFilled As Integer

        e.Graphics.MeasureString(fullText.Substring(currentCharIndex), font, layoutArea.Size, stringFormat, charsFitted, linesFilled)

        ' Print the portion that fits
        e.Graphics.DrawString(fullText.Substring(currentCharIndex), font, Brushes.Black, layoutArea, stringFormat)

        ' Advance the character index
        currentCharIndex += charsFitted

        ' Check if more pages are needed
        e.HasMorePages = (currentCharIndex < fullText.Length)

        ' Cleanup
        font.Dispose()
    End Sub




End Class
