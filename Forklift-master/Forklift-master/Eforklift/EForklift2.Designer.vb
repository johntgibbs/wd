<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Eforklift2
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
        Me.xit = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.List1 = New System.Windows.Forms.ListBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Text1 = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'xit
        '
        Me.xit.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.xit.Location = New System.Drawing.Point(751, 0)
        Me.xit.Name = "xit"
        Me.xit.Size = New System.Drawing.Size(120, 79)
        Me.xit.TabIndex = 0
        Me.xit.Text = "Xit"
        Me.xit.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(399, 238)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(158, 25)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Options Menu"
        '
        'List1
        '
        Me.List1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.List1.FormattingEnabled = True
        Me.List1.ItemHeight = 24
        Me.List1.Location = New System.Drawing.Point(404, 283)
        Me.List1.Name = "List1"
        Me.List1.Size = New System.Drawing.Size(345, 196)
        Me.List1.TabIndex = 2
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(404, 504)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(196, 56)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Accept"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Text1
        '
        Me.Text1.Font = New System.Drawing.Font("Courier New", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Text1.Location = New System.Drawing.Point(2, 0)
        Me.Text1.Multiline = True
        Me.Text1.Name = "Text1"
        Me.Text1.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.Text1.Size = New System.Drawing.Size(367, 479)
        Me.Text1.TabIndex = 4
        Me.Text1.Visible = False
        '
        'Eforklift2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(883, 584)
        Me.Controls.Add(Me.Text1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.List1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.xit)
        Me.Name = "Eforklift2"
        Me.Text = "EFL Options Menu"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents xit As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents List1 As System.Windows.Forms.ListBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Text1 As System.Windows.Forms.TextBox
End Class
