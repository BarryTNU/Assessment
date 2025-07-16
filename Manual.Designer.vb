<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Manual
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Manual))
        MenuStrip1 = New MenuStrip()
        FileToolStripMenuItem = New ToolStripMenuItem()
        ToolStripSeparator1 = New ToolStripSeparator()
        PrintToolStripMenuItem = New ToolStripMenuItem()
        ToolStripSeparator3 = New ToolStripSeparator()
        SaveToolStripMenuItem = New ToolStripMenuItem()
        ToolStripSeparator4 = New ToolStripSeparator()
        ExitToolStripMenuItem = New ToolStripMenuItem()
        ToolStripMenuItem1 = New ToolStripMenuItem()
        ZoomToolStripMenuItem = New ToolStripMenuItem()
        ToolStripSeparator5 = New ToolStripSeparator()
        ToolStripSeparator6 = New ToolStripSeparator()
        ZoomInToolStripMenuItem = New ToolStripMenuItem()
        ToolStripSeparator7 = New ToolStripSeparator()
        ZoomOutToolStripMenuItem = New ToolStripMenuItem()
        TextBox = New RichTextBox()
        Btn_Bold = New ToolStripButton()
        Btn_Underline = New ToolStripButton()
        toolStripSeparator = New ToolStripSeparator()
        toolStripSeparator2 = New ToolStripSeparator()
        ToolStrip1 = New ToolStrip()
        Btn_Italic = New ToolStripButton()
        Lbl_FontSize = New ToolStripLabel()
        Lbl_FontStyle = New ToolStripLabel()
        Lst_FontStyle = New ListBox()
        Lst_FontSize = New ListBox()
        PrintDocument1 = New Printing.PrintDocument()
        MenuStrip1.SuspendLayout()
        ToolStrip1.SuspendLayout()
        SuspendLayout()
        ' 
        ' MenuStrip1
        ' 
        MenuStrip1.BackColor = SystemColors.HotTrack
        MenuStrip1.Font = New Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        MenuStrip1.Items.AddRange(New ToolStripItem() {FileToolStripMenuItem, ToolStripMenuItem1, ZoomToolStripMenuItem})
        MenuStrip1.Location = New Point(0, 0)
        MenuStrip1.Name = "MenuStrip1"
        MenuStrip1.Padding = New Padding(9, 3, 0, 3)
        MenuStrip1.Size = New Size(1143, 31)
        MenuStrip1.TabIndex = 0
        MenuStrip1.Text = "MenuStrip1"
        ' 
        ' FileToolStripMenuItem
        ' 
        FileToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {ToolStripSeparator1, PrintToolStripMenuItem, ToolStripSeparator3, SaveToolStripMenuItem, ToolStripSeparator4, ExitToolStripMenuItem})
        FileToolStripMenuItem.ForeColor = SystemColors.ButtonHighlight
        FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        FileToolStripMenuItem.Size = New Size(49, 25)
        FileToolStripMenuItem.Text = "File"
        ' 
        ' ToolStripSeparator1
        ' 
        ToolStripSeparator1.Name = "ToolStripSeparator1"
        ToolStripSeparator1.Size = New Size(177, 6)
        ' 
        ' PrintToolStripMenuItem
        ' 
        PrintToolStripMenuItem.Name = "PrintToolStripMenuItem"
        PrintToolStripMenuItem.Size = New Size(180, 26)
        PrintToolStripMenuItem.Text = "Print"
        ' 
        ' ToolStripSeparator3
        ' 
        ToolStripSeparator3.Name = "ToolStripSeparator3"
        ToolStripSeparator3.Size = New Size(177, 6)
        ' 
        ' SaveToolStripMenuItem
        ' 
        SaveToolStripMenuItem.Name = "SaveToolStripMenuItem"
        SaveToolStripMenuItem.Size = New Size(180, 26)
        SaveToolStripMenuItem.Text = "Save"
        ' 
        ' ToolStripSeparator4
        ' 
        ToolStripSeparator4.Name = "ToolStripSeparator4"
        ToolStripSeparator4.Size = New Size(177, 6)
        ' 
        ' ExitToolStripMenuItem
        ' 
        ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        ExitToolStripMenuItem.Size = New Size(180, 26)
        ExitToolStripMenuItem.Text = "Exit"
        ' 
        ' ToolStripMenuItem1
        ' 
        ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        ToolStripMenuItem1.Size = New Size(12, 25)
        ' 
        ' ZoomToolStripMenuItem
        ' 
        ZoomToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {ToolStripSeparator5, ToolStripSeparator6, ZoomInToolStripMenuItem, ToolStripSeparator7, ZoomOutToolStripMenuItem})
        ZoomToolStripMenuItem.ForeColor = SystemColors.ButtonHighlight
        ZoomToolStripMenuItem.Name = "ZoomToolStripMenuItem"
        ZoomToolStripMenuItem.Size = New Size(67, 25)
        ZoomToolStripMenuItem.Text = "Zoom"
        ' 
        ' ToolStripSeparator5
        ' 
        ToolStripSeparator5.Name = "ToolStripSeparator5"
        ToolStripSeparator5.Size = New Size(154, 6)
        ' 
        ' ToolStripSeparator6
        ' 
        ToolStripSeparator6.Name = "ToolStripSeparator6"
        ToolStripSeparator6.Size = New Size(154, 6)
        ' 
        ' ZoomInToolStripMenuItem
        ' 
        ZoomInToolStripMenuItem.Name = "ZoomInToolStripMenuItem"
        ZoomInToolStripMenuItem.Size = New Size(157, 26)
        ZoomInToolStripMenuItem.Text = "Zoom In"
        ' 
        ' ToolStripSeparator7
        ' 
        ToolStripSeparator7.Name = "ToolStripSeparator7"
        ToolStripSeparator7.Size = New Size(154, 6)
        ' 
        ' ZoomOutToolStripMenuItem
        ' 
        ZoomOutToolStripMenuItem.Name = "ZoomOutToolStripMenuItem"
        ZoomOutToolStripMenuItem.Size = New Size(157, 26)
        ZoomOutToolStripMenuItem.Text = "Zoom Out"
        ' 
        ' TextBox
        ' 
        TextBox.Dock = DockStyle.Bottom
        TextBox.Font = New Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        TextBox.ForeColor = SystemColors.MenuText
        TextBox.Location = New Point(0, 35)
        TextBox.Margin = New Padding(4)
        TextBox.Name = "TextBox"
        TextBox.ReadOnly = True
        TextBox.Size = New Size(1143, 595)
        TextBox.TabIndex = 1
        TextBox.Text = ""
        ' 
        ' Btn_Bold
        ' 
        Btn_Bold.Alignment = ToolStripItemAlignment.Right
        Btn_Bold.DisplayStyle = ToolStripItemDisplayStyle.Text
        Btn_Bold.Font = New Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Btn_Bold.Image = CType(resources.GetObject("Btn_Bold.Image"), Image)
        Btn_Bold.ImageTransparentColor = Color.Magenta
        Btn_Bold.Name = "Btn_Bold"
        Btn_Bold.Size = New Size(49, 25)
        Btn_Bold.Text = "Bold"
        ' 
        ' Btn_Underline
        ' 
        Btn_Underline.Alignment = ToolStripItemAlignment.Right
        Btn_Underline.DisplayStyle = ToolStripItemDisplayStyle.Text
        Btn_Underline.Font = New Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Btn_Underline.Image = CType(resources.GetObject("Btn_Underline.Image"), Image)
        Btn_Underline.ImageTransparentColor = Color.Magenta
        Btn_Underline.Name = "Btn_Underline"
        Btn_Underline.Size = New Size(90, 25)
        Btn_Underline.Text = "Underline"
        ' 
        ' toolStripSeparator
        ' 
        toolStripSeparator.Name = "toolStripSeparator"
        toolStripSeparator.Size = New Size(6, 28)
        ' 
        ' toolStripSeparator2
        ' 
        toolStripSeparator2.Name = "toolStripSeparator2"
        toolStripSeparator2.Size = New Size(6, 28)
        ' 
        ' ToolStrip1
        ' 
        ToolStrip1.Items.AddRange(New ToolStripItem() {Btn_Italic, Btn_Bold, Btn_Underline, Lbl_FontSize, Lbl_FontStyle, toolStripSeparator, toolStripSeparator2})
        ToolStrip1.Location = New Point(0, 31)
        ToolStrip1.Name = "ToolStrip1"
        ToolStrip1.Size = New Size(1143, 28)
        ToolStrip1.TabIndex = 2
        ToolStrip1.Visible = False
        ' 
        ' Btn_Italic
        ' 
        Btn_Italic.Alignment = ToolStripItemAlignment.Right
        Btn_Italic.DisplayStyle = ToolStripItemDisplayStyle.Text
        Btn_Italic.Font = New Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Btn_Italic.ImageTransparentColor = Color.Magenta
        Btn_Italic.Name = "Btn_Italic"
        Btn_Italic.Size = New Size(52, 25)
        Btn_Italic.Text = "Italic"
        ' 
        ' Lbl_FontSize
        ' 
        Lbl_FontSize.Alignment = ToolStripItemAlignment.Right
        Lbl_FontSize.Font = New Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Lbl_FontSize.Name = "Lbl_FontSize"
        Lbl_FontSize.Size = New Size(79, 25)
        Lbl_FontSize.Text = "Font Size"
        ' 
        ' Lbl_FontStyle
        ' 
        Lbl_FontStyle.Alignment = ToolStripItemAlignment.Right
        Lbl_FontStyle.Font = New Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Lbl_FontStyle.Name = "Lbl_FontStyle"
        Lbl_FontStyle.Size = New Size(86, 25)
        Lbl_FontStyle.Text = "Font Style"
        ' 
        ' Lst_FontStyle
        ' 
        Lst_FontStyle.FormattingEnabled = True
        Lst_FontStyle.HorizontalScrollbar = True
        Lst_FontStyle.ItemHeight = 21
        Lst_FontStyle.Location = New Point(681, 61)
        Lst_FontStyle.Name = "Lst_FontStyle"
        Lst_FontStyle.Size = New Size(198, 382)
        Lst_FontStyle.TabIndex = 3
        Lst_FontStyle.Visible = False
        ' 
        ' Lst_FontSize
        ' 
        Lst_FontSize.FormattingEnabled = True
        Lst_FontSize.ItemHeight = 21
        Lst_FontSize.Location = New Point(885, 61)
        Lst_FontSize.Name = "Lst_FontSize"
        Lst_FontSize.Size = New Size(50, 256)
        Lst_FontSize.TabIndex = 4
        Lst_FontSize.Visible = False
        ' 
        ' Manual
        ' 
        AutoScaleDimensions = New SizeF(10F, 21F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1143, 630)
        ControlBox = False
        Controls.Add(Lst_FontSize)
        Controls.Add(Lst_FontStyle)
        Controls.Add(ToolStrip1)
        Controls.Add(TextBox)
        Controls.Add(MenuStrip1)
        Font = New Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        MainMenuStrip = MenuStrip1
        Margin = New Padding(4)
        Name = "Manual"
        Text = "Stock Assessment Manual"
        MenuStrip1.ResumeLayout(False)
        MenuStrip1.PerformLayout()
        ToolStrip1.ResumeLayout(False)
        ToolStrip1.PerformLayout()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents TextBox As RichTextBox
    Friend WithEvents Btn_Bold As ToolStripButton
    Friend WithEvents Btn_Underline As ToolStripButton
    Friend WithEvents toolStripSeparator As ToolStripSeparator
    Friend WithEvents toolStripSeparator2 As ToolStripSeparator
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents Btn_Italic As ToolStripButton
    Friend WithEvents FileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents PrintToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ZoomToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ZoomInToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents Lbl_FontStyle As ToolStripLabel
    Friend WithEvents Lbl_FontSize As ToolStripLabel
    Friend WithEvents Lst_FontStyle As ListBox
    Friend WithEvents Lst_FontSize As ListBox
    Friend WithEvents SaveToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator4 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator5 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator6 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator7 As ToolStripSeparator
    Friend WithEvents ZoomOutToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents PrintDocument1 As Drawing.Printing.PrintDocument
End Class
