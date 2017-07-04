<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmXlsToCsv
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
        Me.label5 = New System.Windows.Forms.Label()
        Me.label4 = New System.Windows.Forms.Label()
        Me.cmbSheet = New System.Windows.Forms.ComboBox()
        Me.label3 = New System.Windows.Forms.Label()
        Me.txtCsv = New System.Windows.Forms.TextBox()
        Me.label2 = New System.Windows.Forms.Label()
        Me.label1 = New System.Windows.Forms.Label()
        Me.txtFolder = New System.Windows.Forms.TextBox()
        Me.txtXlx = New System.Windows.Forms.TextBox()
        Me.btnBrowsFolder = New System.Windows.Forms.Button()
        Me.btnConvert = New System.Windows.Forms.Button()
        Me.btnBrows = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.SuspendLayout()
        '
        'label5
        '
        Me.label5.AutoSize = True
        Me.label5.BackColor = System.Drawing.Color.AliceBlue
        Me.label5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label5.Location = New System.Drawing.Point(19, 56)
        Me.label5.Name = "label5"
        Me.label5.Size = New System.Drawing.Size(213, 26)
        Me.label5.TabIndex = 23
        Me.label5.Text = "Convert Excel to CSV"
        '
        'label4
        '
        Me.label4.AutoSize = True
        Me.label4.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.label4.Location = New System.Drawing.Point(20, 138)
        Me.label4.Name = "label4"
        Me.label4.Size = New System.Drawing.Size(72, 13)
        Me.label4.TabIndex = 22
        Me.label4.Text = "Sheet Name :"
        '
        'cmbSheet
        '
        Me.cmbSheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSheet.FormattingEnabled = True
        Me.cmbSheet.Location = New System.Drawing.Point(130, 135)
        Me.cmbSheet.Name = "cmbSheet"
        Me.cmbSheet.Size = New System.Drawing.Size(198, 21)
        Me.cmbSheet.TabIndex = 21
        '
        'label3
        '
        Me.label3.AutoSize = True
        Me.label3.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.label3.Location = New System.Drawing.Point(20, 229)
        Me.label3.Name = "label3"
        Me.label3.Size = New System.Drawing.Size(90, 13)
        Me.label3.TabIndex = 20
        Me.label3.Text = "Output file name :"
        '
        'txtCsv
        '
        Me.txtCsv.Location = New System.Drawing.Point(130, 226)
        Me.txtCsv.Name = "txtCsv"
        Me.txtCsv.Size = New System.Drawing.Size(198, 20)
        Me.txtCsv.TabIndex = 19
        '
        'label2
        '
        Me.label2.AutoSize = True
        Me.label2.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.label2.Location = New System.Drawing.Point(20, 202)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(77, 13)
        Me.label2.TabIndex = 18
        Me.label2.Text = "Output Folder :"
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.label1.Location = New System.Drawing.Point(20, 106)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(66, 13)
        Me.label1.TabIndex = 17
        Me.label1.Text = "Input Excel :"
        '
        'txtFolder
        '
        Me.txtFolder.Location = New System.Drawing.Point(130, 199)
        Me.txtFolder.Name = "txtFolder"
        Me.txtFolder.Size = New System.Drawing.Size(239, 20)
        Me.txtFolder.TabIndex = 16
        '
        'txtXlx
        '
        Me.txtXlx.Location = New System.Drawing.Point(130, 106)
        Me.txtXlx.Name = "txtXlx"
        Me.txtXlx.Size = New System.Drawing.Size(239, 20)
        Me.txtXlx.TabIndex = 15
        '
        'btnBrowsFolder
        '
        Me.btnBrowsFolder.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.btnBrowsFolder.Location = New System.Drawing.Point(375, 196)
        Me.btnBrowsFolder.Name = "btnBrowsFolder"
        Me.btnBrowsFolder.Size = New System.Drawing.Size(57, 23)
        Me.btnBrowsFolder.TabIndex = 14
        Me.btnBrowsFolder.Text = "Browse"
        Me.btnBrowsFolder.UseVisualStyleBackColor = False
        '
        'btnConvert
        '
        Me.btnConvert.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.btnConvert.Location = New System.Drawing.Point(328, 252)
        Me.btnConvert.Name = "btnConvert"
        Me.btnConvert.Size = New System.Drawing.Size(104, 42)
        Me.btnConvert.TabIndex = 13
        Me.btnConvert.Text = "CONVERT"
        Me.btnConvert.UseVisualStyleBackColor = False
        '
        'btnBrows
        '
        Me.btnBrows.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.btnBrows.Location = New System.Drawing.Point(375, 106)
        Me.btnBrows.Name = "btnBrows"
        Me.btnBrows.Size = New System.Drawing.Size(57, 23)
        Me.btnBrows.TabIndex = 24
        Me.btnBrows.Text = "Browse"
        Me.btnBrows.UseVisualStyleBackColor = False
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'frmXlsToCsv
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(521, 350)
        Me.Controls.Add(Me.btnBrows)
        Me.Controls.Add(Me.label5)
        Me.Controls.Add(Me.label4)
        Me.Controls.Add(Me.cmbSheet)
        Me.Controls.Add(Me.label3)
        Me.Controls.Add(Me.txtCsv)
        Me.Controls.Add(Me.label2)
        Me.Controls.Add(Me.label1)
        Me.Controls.Add(Me.txtFolder)
        Me.Controls.Add(Me.txtXlx)
        Me.Controls.Add(Me.btnBrowsFolder)
        Me.Controls.Add(Me.btnConvert)
        Me.Name = "frmXlsToCsv"
        Me.Text = "Excel To Csv"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Private WithEvents label5 As Label
    Private WithEvents label4 As Label
    Private WithEvents cmbSheet As ComboBox
    Private WithEvents label3 As Label
    Private WithEvents txtCsv As TextBox
    Private WithEvents label2 As Label
    Private WithEvents label1 As Label
    Private WithEvents txtFolder As TextBox
    Private WithEvents txtXlx As TextBox
    Private WithEvents btnBrowsFolder As Button
    Private WithEvents btnConvert As Button
    Private WithEvents btnBrows As Button
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
End Class
