Imports System.IO
Imports System.Windows
Module Form_Zoom

    Public currentFontSize As Single = 12.0F
    Public ScreenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
    Public ScreenHeight As Integer = Screen.PrimaryScreen.Bounds.Height

    Sub ZoomIn(sender As Object, e As EventArgs, zForm As Form, LV As ListView)

        Dim totalWidth As Integer = 0 '
        Dim totalHeight As Integer = 0

        With zForm

            If LV.Width <= .MaximumSize.Width Then 'We have not yet reached maximum size
                currentFontSize += 1.4F
                ChangeFontSize(zForm, currentFontSize)

                .Hide()

                LV.Font = New Font(Form1.LV_Stock.Font.FontFamily, CInt(currentFontSize))
                For Each col As ColumnHeader In LV.Columns
                    col.Width += CInt(currentFontSize)
                    totalWidth += col.Width
                Next
            Else
                .Show()
                Exit Sub
            End If

            Dim itemHeight As Integer = LV.GetItemRect(0).Height
            totalHeight = (itemHeight * LV.Items.Count)
            If totalHeight >= ScreenHeight - 300 Then totalHeight = ScreenHeight - 300

            ' --- Add some padding for scrollbars and borders ---
            totalWidth += SystemInformation.VerticalScrollBarWidth + 20
            totalHeight += SystemInformation.HorizontalScrollBarHeight + 60

            ' --- Resize the form  ---

            .Width = totalWidth
            .Height = totalHeight + 20 'ScreenHeight

            'Center form onscreen
            Dim screenBounds As Rectangle = Screen.PrimaryScreen.WorkingArea
            .Left = (screenBounds.Width - .Width) \ 2
            .Top = (screenBounds.Height - .Height) \ 2
            .Show()

        End With

    End Sub

    Sub ZoomOut(sender As Object, e As EventArgs, zForm As Form, LV As ListView)
        Dim totalWidth As Integer = 0
        Dim totalHeight As Integer = 0

        With zForm

            If LV.Width > .MinimumSize.Width Then 'We have not reached minimum  size

                currentFontSize -= 1.4F
                If currentFontSize < 10 Then
                    currentFontSize = 10
                End If

                ChangeFontSize(zForm, currentFontSize)

                Dim w As Integer = .Width
                .Hide()

                LV.Font = New Font(Form1.LV_Stock.Font.FontFamily, CInt(currentFontSize))
                For Each col As ColumnHeader In LV.Columns
                    col.Width -= CInt(currentFontSize)
                    totalWidth += col.Width
                Next
            Else
                Exit Sub

            End If

            Dim itemHeight As Integer = LV.GetItemRect(0).Height
            totalHeight = (itemHeight * LV.Items.Count)

            ' --- Add some padding for scrollbars and borders ---
            totalWidth += SystemInformation.VerticalScrollBarWidth + 20
            totalHeight += SystemInformation.HorizontalScrollBarHeight + 60
            If totalHeight >= ScreenHeight - 300 Then totalHeight = ScreenHeight - 300

            ' --- Resize the form   ---
            .Width = totalWidth
            .Height = totalHeight + 20 'ScreenHeight

            'Center form onscreen
            Dim screenBounds As Rectangle = Screen.PrimaryScreen.WorkingArea
            .Left = (screenBounds.Width - .Width) \ 2
            .Top = (screenBounds.Height - .Height) \ 2
            .Show()

        End With

    End Sub


    Private Sub ChangeFontSize(zform As Form, newSize As Single)
        For Each ctrl As Control In zform.Controls
            ctrl.Font = New Font(ctrl.Font.FontFamily, newSize, ctrl.Font.Style)
        Next
    End Sub
End Module

