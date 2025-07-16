

Imports System.Windows
Imports System.Windows.Controls

Module Zoom_TextBox

    Public currentFontSize As Single = 12.0F
    ' Public ScreenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
    ' Public ScreenHeight As Integer = Screen.PrimaryScreen.Bounds.Height
    Public FormHeight As Integer = Screen.PrimaryScreen.Bounds.Height
    Public TextBoxHeight = Manual.TextBox.Height
    Public EditingAdjustment As Integer = 100

    Sub ZoomIn(Zoom As Integer)
        Dim screenBounds As Rectangle = Screen.PrimaryScreen.WorkingArea
        Dim totalWidth As Integer = Manual.Width
        Dim totalHeight As Integer = Manual.Height
        totalWidth += Zoom * 5
        totalHeight += Zoom * 2
        If totalWidth < screenBounds.Width And totalHeight < screenBounds.Height Then
            ' --- Resize the form   ---
            Manual.Width = totalWidth
            Manual.Height = totalHeight
            Manual.TextBox.Height = totalHeight - EditingAdjustment

            Manual.TextBox.SelectAll()
            Dim fontSize As Single = Manual.TextBox.SelectionFont.Size
            If fontSize + 2 <= 30 Then 'set maximum size of font
                Manual.TextBox.SelectionFont = New Font(Manual.TextBox.SelectionFont.FontFamily, fontSize + 2)
            End If
            Manual.TextBox.DeselectAll()
        End If

        'Center form onscreen
        Manual.Left = (screenBounds.Width - Manual.Width) \ 2
        Manual.Top = (screenBounds.Height - Manual.Height) \ 2
        Manual.Show()

    End Sub

    Sub ZoomOut(Zoom As Integer)
        Dim totalWidth As Integer = Manual.Width
        Dim totalHeight As Integer = Manual.Height
        Dim screenBounds As Rectangle = Screen.PrimaryScreen.WorkingArea

        totalWidth -= Zoom * 5
        totalHeight -= Zoom * 2
        If totalWidth > 900 Then
            'Resize the form
            Manual.Width = totalWidth
            Manual.Height = totalHeight
            Manual.TextBox.Height = totalHeight - EditingAdjustment
            Manual.TextBox.SelectAll()
            Dim fontSize As Single = Manual.TextBox.SelectionFont.Size
            If fontSize - 2 > 12 Then 'Set minimum size of font
                Manual.TextBox.SelectionFont = New Font(Manual.TextBox.SelectionFont.FontFamily, fontSize - 2)
            End If
            Manual.TextBox.DeselectAll()
        End If

        'Center form onscreen
        Manual.Left = (screenBounds.Width - Manual.Width) \ 2
        Manual.Top = (screenBounds.Height - Manual.Height) \ 2
        Manual.Show()

    End Sub

End Module
