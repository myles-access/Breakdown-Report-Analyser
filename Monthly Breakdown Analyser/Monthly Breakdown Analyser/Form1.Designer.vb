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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.runButton = New System.Windows.Forms.Button()
        Me.fileSelectButton = New System.Windows.Forms.Button()
        Me.fileLabel = New System.Windows.Forms.Label()
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'runButton
        '
        Me.runButton.Font = New System.Drawing.Font("Calibri", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.runButton.Location = New System.Drawing.Point(12, 13)
        Me.runButton.Name = "runButton"
        Me.runButton.Size = New System.Drawing.Size(454, 175)
        Me.runButton.TabIndex = 0
        Me.runButton.Text = "Generate Report Analysis"
        Me.runButton.UseVisualStyleBackColor = True
        '
        'fileSelectButton
        '
        Me.fileSelectButton.BackgroundImage = CType(resources.GetObject("fileSelectButton.BackgroundImage"), System.Drawing.Image)
        Me.fileSelectButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.fileSelectButton.Location = New System.Drawing.Point(12, 234)
        Me.fileSelectButton.Name = "fileSelectButton"
        Me.fileSelectButton.Size = New System.Drawing.Size(126, 130)
        Me.fileSelectButton.TabIndex = 1
        Me.fileSelectButton.UseVisualStyleBackColor = True
        '
        'fileLabel
        '
        Me.fileLabel.AutoSize = True
        Me.fileLabel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.fileLabel.ForeColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.fileLabel.Location = New System.Drawing.Point(144, 288)
        Me.fileLabel.Name = "fileLabel"
        Me.fileLabel.Size = New System.Drawing.Size(126, 24)
        Me.fileLabel.TabIndex = 2
        Me.fileLabel.Text = "No file selected"
        '
        'OpenFileDialog
        '
        Me.OpenFileDialog.FileName = "OpenFileDialog1"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(32, 194)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 27)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "close xl"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 22.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(478, 376)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.fileLabel)
        Me.Controls.Add(Me.fileSelectButton)
        Me.Controls.Add(Me.runButton)
        Me.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form1"
        Me.Text = "Breakdown Report Analyser"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents runButton As Button
    Friend WithEvents fileSelectButton As Button
    Friend WithEvents fileLabel As Label
    Friend WithEvents OpenFileDialog As OpenFileDialog
    Friend WithEvents Button1 As Button
End Class
