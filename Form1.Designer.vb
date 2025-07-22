<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        MenuStrip1 = New MenuStrip()
        OpenToolStripMenuItem = New ToolStripMenuItem()
        ToolStripSeparator6 = New ToolStripSeparator()
        NewStockListToolStripMenuItem = New ToolStripMenuItem()
        AssessmentToolStripMenuItem = New ToolStripMenuItem()
        ToolStripSeparator7 = New ToolStripSeparator()
        NewDepartmentToolStripMenuItem = New ToolStripMenuItem()
        ToolStripSeparator1 = New ToolStripSeparator()
        CloseToolStripMenuItem = New ToolStripMenuItem()
        SaveSeparator = New ToolStripSeparator()
        SaveToolStripMenuItem = New ToolStripMenuItem()
        SaveAsSeparator = New ToolStripSeparator()
        SaveAsToolStripMenuItem = New ToolStripMenuItem()
        EditSeparator = New ToolStripSeparator()
        EditToolStripMenuItem = New ToolStripMenuItem()
        DoAssessSeparator = New ToolStripSeparator()
        DoAssessmentToolStripMenuItem = New ToolStripMenuItem()
        ExitSeparator = New ToolStripSeparator()
        ExitToolStripMenuItem = New ToolStripMenuItem()
        PrintToolStripMenuItem = New ToolStripMenuItem()
        ToolStripSeparator9 = New ToolStripSeparator()
        PrintAssessmentToolStripMenuItem = New ToolStripMenuItem()
        ToolStripSeparator8 = New ToolStripSeparator()
        PrintWorkSheetToolStripMenuItem = New ToolStripMenuItem()
        ManualToolStripMenuItem = New ToolStripMenuItem()
        SettingsToolStripMenuItem = New ToolStripMenuItem()
        ToolStripSeparator12 = New ToolStripSeparator()
        AutoSaveToolStripMenuItem = New ToolStripMenuItem()
        ScreenBackgroundToolStripMenuItem = New ToolStripMenuItem()
        ZoomToolStripMenuItem = New ToolStripMenuItem()
        ToolStripSeparator11 = New ToolStripSeparator()
        ZoomInToolStripMenuItem = New ToolStripMenuItem()
        ToolStripSeparator10 = New ToolStripSeparator()
        ZoomOutToolStripMenuItem = New ToolStripMenuItem()
        OpenFileDialog1 = New OpenFileDialog()
        SaveFileDialog1 = New SaveFileDialog()
        LV_Stock = New ListView()
        AboutStockAssessmentToolStripMenuItem = New ToolStripMenuItem()
        MenuStrip1.SuspendLayout()
        SuspendLayout()
        ' 
        ' MenuStrip1
        ' 
        MenuStrip1.BackColor = SystemColors.Control
        MenuStrip1.Font = New Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        MenuStrip1.Items.AddRange(New ToolStripItem() {OpenToolStripMenuItem, PrintToolStripMenuItem, ManualToolStripMenuItem, SettingsToolStripMenuItem, ZoomToolStripMenuItem})
        MenuStrip1.LayoutStyle = ToolStripLayoutStyle.HorizontalStackWithOverflow
        MenuStrip1.Location = New Point(0, 0)
        MenuStrip1.Name = "MenuStrip1"
        MenuStrip1.Padding = New Padding(8, 3, 0, 3)
        MenuStrip1.Size = New Size(912, 31)
        MenuStrip1.TabIndex = 1
        MenuStrip1.Text = "MenuStrip1"
        ' 
        ' OpenToolStripMenuItem
        ' 
        OpenToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {ToolStripSeparator6, NewStockListToolStripMenuItem, ToolStripSeparator1, CloseToolStripMenuItem, SaveSeparator, SaveToolStripMenuItem, SaveAsSeparator, SaveAsToolStripMenuItem, EditSeparator, EditToolStripMenuItem, DoAssessSeparator, DoAssessmentToolStripMenuItem, ExitSeparator, ExitToolStripMenuItem})
        OpenToolStripMenuItem.ForeColor = SystemColors.ActiveCaptionText
        OpenToolStripMenuItem.Name = "OpenToolStripMenuItem"
        OpenToolStripMenuItem.Size = New Size(49, 25)
        OpenToolStripMenuItem.Text = "File"
        ' 
        ' ToolStripSeparator6
        ' 
        ToolStripSeparator6.Name = "ToolStripSeparator6"
        ToolStripSeparator6.Size = New Size(199, 6)
        ' 
        ' NewStockListToolStripMenuItem
        ' 
        NewStockListToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {AssessmentToolStripMenuItem, ToolStripSeparator7, NewDepartmentToolStripMenuItem})
        NewStockListToolStripMenuItem.Name = "NewStockListToolStripMenuItem"
        NewStockListToolStripMenuItem.Size = New Size(202, 26)
        NewStockListToolStripMenuItem.Text = "Open"
        ' 
        ' AssessmentToolStripMenuItem
        ' 
        AssessmentToolStripMenuItem.Name = "AssessmentToolStripMenuItem"
        AssessmentToolStripMenuItem.Size = New Size(172, 26)
        AssessmentToolStripMenuItem.Text = "Assessment"
        AssessmentToolStripMenuItem.Visible = False
        ' 
        ' ToolStripSeparator7
        ' 
        ToolStripSeparator7.Name = "ToolStripSeparator7"
        ToolStripSeparator7.Size = New Size(169, 6)
        ' 
        ' NewDepartmentToolStripMenuItem
        ' 
        NewDepartmentToolStripMenuItem.Name = "NewDepartmentToolStripMenuItem"
        NewDepartmentToolStripMenuItem.Size = New Size(172, 26)
        NewDepartmentToolStripMenuItem.Text = "Department"
        ' 
        ' ToolStripSeparator1
        ' 
        ToolStripSeparator1.Name = "ToolStripSeparator1"
        ToolStripSeparator1.Size = New Size(199, 6)
        ' 
        ' CloseToolStripMenuItem
        ' 
        CloseToolStripMenuItem.Name = "CloseToolStripMenuItem"
        CloseToolStripMenuItem.Size = New Size(202, 26)
        CloseToolStripMenuItem.Text = "Close"
        ' 
        ' SaveSeparator
        ' 
        SaveSeparator.Name = "SaveSeparator"
        SaveSeparator.Size = New Size(199, 6)
        ' 
        ' SaveToolStripMenuItem
        ' 
        SaveToolStripMenuItem.Name = "SaveToolStripMenuItem"
        SaveToolStripMenuItem.Size = New Size(202, 26)
        SaveToolStripMenuItem.Text = "Save"
        SaveToolStripMenuItem.Visible = False
        ' 
        ' SaveAsSeparator
        ' 
        SaveAsSeparator.Name = "SaveAsSeparator"
        SaveAsSeparator.Size = New Size(199, 6)
        ' 
        ' SaveAsToolStripMenuItem
        ' 
        SaveAsToolStripMenuItem.Name = "SaveAsToolStripMenuItem"
        SaveAsToolStripMenuItem.Size = New Size(202, 26)
        SaveAsToolStripMenuItem.Text = "Save As"
        ' 
        ' EditSeparator
        ' 
        EditSeparator.Name = "EditSeparator"
        EditSeparator.Size = New Size(199, 6)
        ' 
        ' EditToolStripMenuItem
        ' 
        EditToolStripMenuItem.Name = "EditToolStripMenuItem"
        EditToolStripMenuItem.Size = New Size(202, 26)
        EditToolStripMenuItem.Text = "Edit Stock Items"
        EditToolStripMenuItem.Visible = False
        ' 
        ' DoAssessSeparator
        ' 
        DoAssessSeparator.Name = "DoAssessSeparator"
        DoAssessSeparator.Size = New Size(199, 6)
        ' 
        ' DoAssessmentToolStripMenuItem
        ' 
        DoAssessmentToolStripMenuItem.Name = "DoAssessmentToolStripMenuItem"
        DoAssessmentToolStripMenuItem.Size = New Size(202, 26)
        DoAssessmentToolStripMenuItem.Text = "Do Assessment"
        DoAssessmentToolStripMenuItem.Visible = False
        ' 
        ' ExitSeparator
        ' 
        ExitSeparator.Name = "ExitSeparator"
        ExitSeparator.Size = New Size(199, 6)
        ' 
        ' ExitToolStripMenuItem
        ' 
        ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        ExitToolStripMenuItem.Size = New Size(202, 26)
        ExitToolStripMenuItem.Text = "Exit"
        ' 
        ' PrintToolStripMenuItem
        ' 
        PrintToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {ToolStripSeparator9, PrintAssessmentToolStripMenuItem, ToolStripSeparator8, PrintWorkSheetToolStripMenuItem})
        PrintToolStripMenuItem.ForeColor = SystemColors.ActiveCaptionText
        PrintToolStripMenuItem.Name = "PrintToolStripMenuItem"
        PrintToolStripMenuItem.Size = New Size(59, 25)
        PrintToolStripMenuItem.Text = "Print"
        PrintToolStripMenuItem.Visible = False
        ' 
        ' ToolStripSeparator9
        ' 
        ToolStripSeparator9.Name = "ToolStripSeparator9"
        ToolStripSeparator9.Size = New Size(165, 6)
        ' 
        ' PrintAssessmentToolStripMenuItem
        ' 
        PrintAssessmentToolStripMenuItem.Name = "PrintAssessmentToolStripMenuItem"
        PrintAssessmentToolStripMenuItem.Size = New Size(168, 26)
        PrintAssessmentToolStripMenuItem.Text = "Assessment"
        ' 
        ' ToolStripSeparator8
        ' 
        ToolStripSeparator8.Name = "ToolStripSeparator8"
        ToolStripSeparator8.Size = New Size(165, 6)
        ' 
        ' PrintWorkSheetToolStripMenuItem
        ' 
        PrintWorkSheetToolStripMenuItem.Name = "PrintWorkSheetToolStripMenuItem"
        PrintWorkSheetToolStripMenuItem.Size = New Size(168, 26)
        PrintWorkSheetToolStripMenuItem.Text = "Work Sheet"
        ' 
        ' ManualToolStripMenuItem
        ' 
        ManualToolStripMenuItem.Alignment = ToolStripItemAlignment.Right
        ManualToolStripMenuItem.ForeColor = SystemColors.ActiveCaptionText
        ManualToolStripMenuItem.Name = "ManualToolStripMenuItem"
        ManualToolStripMenuItem.Size = New Size(80, 25)
        ManualToolStripMenuItem.Text = "Manual"
        ' 
        ' SettingsToolStripMenuItem
        ' 
        SettingsToolStripMenuItem.Alignment = ToolStripItemAlignment.Right
        SettingsToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {ToolStripSeparator12, AutoSaveToolStripMenuItem, ScreenBackgroundToolStripMenuItem, AboutStockAssessmentToolStripMenuItem})
        SettingsToolStripMenuItem.ForeColor = SystemColors.ActiveCaptionText
        SettingsToolStripMenuItem.Name = "SettingsToolStripMenuItem"
        SettingsToolStripMenuItem.Size = New Size(84, 25)
        SettingsToolStripMenuItem.Text = "Settings"
        SettingsToolStripMenuItem.Visible = False
        ' 
        ' ToolStripSeparator12
        ' 
        ToolStripSeparator12.Name = "ToolStripSeparator12"
        ToolStripSeparator12.Size = New Size(262, 6)
        ' 
        ' AutoSaveToolStripMenuItem
        ' 
        AutoSaveToolStripMenuItem.Name = "AutoSaveToolStripMenuItem"
        AutoSaveToolStripMenuItem.Size = New Size(265, 26)
        AutoSaveToolStripMenuItem.Text = "Auto Save Off"
        ' 
        ' ScreenBackgroundToolStripMenuItem
        ' 
        ScreenBackgroundToolStripMenuItem.Name = "ScreenBackgroundToolStripMenuItem"
        ScreenBackgroundToolStripMenuItem.Size = New Size(265, 26)
        ScreenBackgroundToolStripMenuItem.Text = "Screen Background"
        ' 
        ' ZoomToolStripMenuItem
        ' 
        ZoomToolStripMenuItem.Alignment = ToolStripItemAlignment.Right
        ZoomToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {ToolStripSeparator11, ZoomInToolStripMenuItem, ToolStripSeparator10, ZoomOutToolStripMenuItem})
        ZoomToolStripMenuItem.ForeColor = SystemColors.ActiveCaptionText
        ZoomToolStripMenuItem.Name = "ZoomToolStripMenuItem"
        ZoomToolStripMenuItem.Size = New Size(67, 25)
        ZoomToolStripMenuItem.Text = "Zoom"
        ZoomToolStripMenuItem.Visible = False
        ' 
        ' ToolStripSeparator11
        ' 
        ToolStripSeparator11.Name = "ToolStripSeparator11"
        ToolStripSeparator11.Size = New Size(154, 6)
        ' 
        ' ZoomInToolStripMenuItem
        ' 
        ZoomInToolStripMenuItem.ForeColor = SystemColors.ActiveCaptionText
        ZoomInToolStripMenuItem.Name = "ZoomInToolStripMenuItem"
        ZoomInToolStripMenuItem.Size = New Size(157, 26)
        ZoomInToolStripMenuItem.Text = "Zoom In"
        ' 
        ' ToolStripSeparator10
        ' 
        ToolStripSeparator10.Name = "ToolStripSeparator10"
        ToolStripSeparator10.Size = New Size(154, 6)
        ' 
        ' ZoomOutToolStripMenuItem
        ' 
        ZoomOutToolStripMenuItem.Name = "ZoomOutToolStripMenuItem"
        ZoomOutToolStripMenuItem.Size = New Size(157, 26)
        ZoomOutToolStripMenuItem.Text = "Zoom Out"
        ' 
        ' OpenFileDialog1
        ' 
        OpenFileDialog1.FileName = "OpenFileDialog1"
        ' 
        ' LV_Stock
        ' 
        LV_Stock.BorderStyle = BorderStyle.None
        LV_Stock.Dock = DockStyle.Fill
        LV_Stock.Font = New Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        LV_Stock.GridLines = True
        LV_Stock.HeaderStyle = ColumnHeaderStyle.Nonclickable
        LV_Stock.LabelWrap = False
        LV_Stock.Location = New Point(0, 31)
        LV_Stock.Name = "LV_Stock"
        LV_Stock.OwnerDraw = True
        LV_Stock.ShowItemToolTips = True
        LV_Stock.Size = New Size(912, 604)
        LV_Stock.TabIndex = 2
        LV_Stock.UseCompatibleStateImageBehavior = False
        LV_Stock.Visible = False
        ' 
        ' AboutStockAssessmentToolStripMenuItem
        ' 
        AboutStockAssessmentToolStripMenuItem.Name = "AboutStockAssessmentToolStripMenuItem"
        AboutStockAssessmentToolStripMenuItem.Size = New Size(265, 26)
        AboutStockAssessmentToolStripMenuItem.Text = "About Stock Assessment"
        ' 
        ' Form1
        ' 
        AutoScaleMode = AutoScaleMode.None
        AutoSize = True
        BackColor = SystemColors.GradientActiveCaption
        BackgroundImageLayout = ImageLayout.Zoom
        ClientSize = New Size(912, 635)
        ControlBox = False
        Controls.Add(LV_Stock)
        Controls.Add(MenuStrip1)
        DoubleBuffered = True
        Font = New Font("Segoe UI Black", 20.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        ForeColor = SystemColors.ActiveCaptionText
        FormBorderStyle = FormBorderStyle.FixedSingle
        MainMenuStrip = MenuStrip1
        Margin = New Padding(4)
        MaximizeBox = False
        MaximumSize = New Size(1800, 1100)
        MinimumSize = New Size(914, 637)
        Name = "Form1"
        SizeGripStyle = SizeGripStyle.Show
        Text = "Stock Assessment"
        MenuStrip1.ResumeLayout(False)
        MenuStrip1.PerformLayout()
        ResumeLayout(False)
        PerformLayout()
    End Sub
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents OpenToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents NewStockListToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SaveToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ZoomToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ZoomInToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ZoomOutToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents EditToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SettingsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AutoSaveToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AssessmentToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents DoAssessmentToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents PrintToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ManualToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents NewDepartmentToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents PrintAssessmentToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents PrintWorkSheetToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents LV_Stock As ListView
    Friend WithEvents ToolStripSeparator6 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator7 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator9 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator8 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator12 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator11 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator10 As ToolStripSeparator
    Friend WithEvents SaveSeparator As ToolStripSeparator
    Friend WithEvents SaveAsSeparator As ToolStripSeparator
    Friend WithEvents EditSeparator As ToolStripSeparator
    Friend WithEvents DoAssessSeparator As ToolStripSeparator
    Friend WithEvents ExitSeparator As ToolStripSeparator
    Friend WithEvents SaveAsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ScreenBackgroundToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents CloseToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AboutStockAssessmentToolStripMenuItem As ToolStripMenuItem

End Class
