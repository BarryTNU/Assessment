Imports System
Imports System.IO

Module Form_Zoom
    Public currentFontSize As Single = 12.0F
    Public FormWidth As Integer = 800
    Public FormHeight As Integer = 400
    Public ScreenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
    Public ScreenHeight As Integer = Screen.PrimaryScreen.Bounds.Height

    Sub ZoomIn(sender As Object, e As EventArgs, zForm As Form, LV As ListView)
        currentFontSize += 1.4F
        If currentFontSize > 16 Then
            ' currentFontSize = 16
        End If
        ChangeFontSize(zForm, currentFontSize)
        Form1.MenuStrip1.Font += 1.4F


        With zForm

            If .Width > ScreenWidth - 100 Or .width < 800 Then
                Exit Sub
            End If

            If FormHeight > ScreenHeight Then
                FormHeight = ScreenHeight - 100
            End If

            LV.Font = New Font(Form1.LV_Stock.Font.FontFamily, CInt(currentFontSize))
            For Each col As ColumnHeader In LV.Columns
                col.Width += CInt(currentFontSize) * 2
            Next
            Dim innerWidth As Integer = LV.ClientSize.Width

            .Width = innerWidth + 300
            If .Width > ScreenWidth Then
                .Width = ScreenWidth
            End If
            .Height = Form1.LV_Stock.Items.Count * 50
            If Form1.Height > ScreenHeight - 50 Then
                Form1.Height = ScreenHeight - 50
            End If

            'Center form onscreen
            Dim screenBounds As Rectangle = Screen.PrimaryScreen.WorkingArea
            .Left = (screenBounds.Width - .Width) \ 2
            .Top = (screenBounds.Height - .Height) \ 2
        End With
    End Sub

    Sub ZoomOut(sender As Object, e As EventArgs, zForm As Form, LV As ListView)

        currentFontSize -= 1.4F
        If currentFontSize < 10 Then
            ' currentFontSize = 10
        End If

        ChangeFontSize(zForm, currentFontSize)
        Form1.MenuStrip1.Font += 1.4F


        With zForm

            If .Width > ScreenWidth - 100 Or .width < 800 Then
                Exit Sub
            End If

            If FormHeight > ScreenHeight Then
                FormHeight = ScreenHeight - 100
            End If

            LV.Font = New Font(Form1.LV_Stock.Font.FontFamily, CInt(currentFontSize))
            For Each col As ColumnHeader In LV.Columns
                col.Width += CInt(currentFontSize) * 2
            Next
            Dim innerWidth As Integer = LV.ClientSize.Width

            .Width = innerWidth + 300
            If .Width > ScreenWidth Then
                .Width = ScreenWidth
            End If
            .Height = Form1.LV_Stock.Items.Count * 50
            If Form1.Height > ScreenHeight - 50 Then
                Form1.Height = ScreenHeight - 50
            End If

            'Center form onscreen
            Dim screenBounds As Rectangle = Screen.PrimaryScreen.WorkingArea
            .Left = (screenBounds.Width - .Width) \ 2
            .Top = (screenBounds.Height - .Height) \ 2
        End With
    End Sub

    Sub Zoom_Out(sender As Object, e As EventArgs, zForm As Form, LV As ListView)

        With zForm
            currentFontSize -= 1.4F
            If currentFontSize < 12 Then
                currentFontSize = 12
            End If

            If .ClientSize.Width < 1050 Then
                Exit Sub
            End If

            LV.Font = New Font(.Font.FontFamily, CInt(currentFontSize))
            For Each col As ColumnHeader In LV.Columns
                col.Width -= CInt(currentFontSize) * 2
            Next
            Dim innerWidth As Integer = LV.Width

            .Width = innerWidth - 195
            If .Width < 900 Then
                .Width = 950
            End If
            .Height = Form1.LV_Stock.Items.Count * 50
            If Form1.Height > ScreenHeight - 50 Then
                Form1.Height = ScreenHeight - 50
            End If

            'Center form onscreen
            Dim screenBounds As Rectangle = Screen.PrimaryScreen.WorkingArea
            .Left = (screenBounds.Width - .Width) \ 2
            .Top = (screenBounds.Height - .Height) \ 2
        End With

    End Sub


    Private Sub ChangeFontSize(zform As Form, newSize As Single)
        For Each ctrl As Control In zform.Controls
            ctrl.Font = New Font(ctrl.Font.FontFamily, newSize, ctrl.Font.Style)
        Next
    End Sub
End Module
