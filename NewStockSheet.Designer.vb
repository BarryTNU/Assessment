<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class New_StockSheet
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
        Lbl_SheetName = New Label()
        Btn_Save = New Button()
        Btn_Cancel = New Button()
        Lbl_DefFolder = New Label()
        Txt_DefFolder = New TextBox()
        OpenFileDialog1 = New OpenFileDialog()
        FolderBrowserDialog1 = New FolderBrowserDialog()
        Cmb_DeptName = New ComboBox()
        SuspendLayout()
        ' 
        ' Lbl_SheetName
        ' 
        Lbl_SheetName.AutoSize = True
        Lbl_SheetName.Font = New Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Lbl_SheetName.Location = New Point(8, 71)
        Lbl_SheetName.Name = "Lbl_SheetName"
        Lbl_SheetName.Size = New Size(152, 21)
        Lbl_SheetName.TabIndex = 1
        Lbl_SheetName.Text = "Department Name"
        ' 
        ' Btn_Save
        ' 
        Btn_Save.BackColor = Color.Lime
        Btn_Save.Font = New Font("Segoe UI", 11.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Btn_Save.Location = New Point(186, 120)
        Btn_Save.Name = "Btn_Save"
        Btn_Save.Size = New Size(75, 42)
        Btn_Save.TabIndex = 16
        Btn_Save.Text = "SAVE"
        Btn_Save.UseVisualStyleBackColor = False
        ' 
        ' Btn_Cancel
        ' 
        Btn_Cancel.BackColor = Color.Red
        Btn_Cancel.FlatAppearance.BorderSize = 2
        Btn_Cancel.Font = New Font("Segoe UI", 11.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Btn_Cancel.Location = New Point(344, 120)
        Btn_Cancel.Name = "Btn_Cancel"
        Btn_Cancel.Size = New Size(85, 42)
        Btn_Cancel.TabIndex = 17
        Btn_Cancel.Text = "CANCEL"
        Btn_Cancel.UseVisualStyleBackColor = False
        ' 
        ' Lbl_DefFolder
        ' 
        Lbl_DefFolder.AutoSize = True
        Lbl_DefFolder.Font = New Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Lbl_DefFolder.Location = New Point(8, 30)
        Lbl_DefFolder.Name = "Lbl_DefFolder"
        Lbl_DefFolder.Size = New Size(128, 21)
        Lbl_DefFolder.TabIndex = 13
        Lbl_DefFolder.Text = "Folder Location"
        ' 
        ' Txt_DefFolder
        ' 
        Txt_DefFolder.Font = New Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Txt_DefFolder.Location = New Point(186, 22)
        Txt_DefFolder.Name = "Txt_DefFolder"
        Txt_DefFolder.Size = New Size(243, 29)
        Txt_DefFolder.TabIndex = 12
        ' 
        ' OpenFileDialog1
        ' 
        OpenFileDialog1.FileName = "OpenFileDialog1"
        ' 
        ' Cmb_DeptName
        ' 
        Cmb_DeptName.Font = New Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Cmb_DeptName.FormattingEnabled = True
        Cmb_DeptName.Location = New Point(186, 67)
        Cmb_DeptName.Name = "Cmb_DeptName"
        Cmb_DeptName.Size = New Size(203, 29)
        Cmb_DeptName.TabIndex = 18
        ' 
        ' New_StockSheet
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        AutoValidate = AutoValidate.Disable
        CancelButton = Btn_Cancel
        ClientSize = New Size(451, 185)
        ControlBox = False
        Controls.Add(Cmb_DeptName)
        Controls.Add(Lbl_DefFolder)
        Controls.Add(Txt_DefFolder)
        Controls.Add(Btn_Cancel)
        Controls.Add(Btn_Save)
        Controls.Add(Lbl_SheetName)
        MaximizeBox = False
        MinimizeBox = False
        Name = "New_StockSheet"
        StartPosition = FormStartPosition.CenterScreen
        Text = "New Stock Sheet"
        ResumeLayout(False)
        PerformLayout()
    End Sub
    Friend WithEvents Lbl_SheetName As Label
    Friend WithEvents Btn_Save As Button
    Friend WithEvents Btn_Cancel As Button
    Friend WithEvents Lbl_DefFolder As Label
    Friend WithEvents Txt_DefFolder As TextBox
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents Ldl_AutoSave As Label
    Friend WithEvents Lst_AutoSave As ListBox
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents Cmb_DeptName As ComboBox
End Class
