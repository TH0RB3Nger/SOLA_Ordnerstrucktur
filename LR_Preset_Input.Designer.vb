<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class LR_Preset_Input
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
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

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Button_Teen = New System.Windows.Forms.Button()
        Me.Button_Kids = New System.Windows.Forms.Button()
        Me.Button_Cancel = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox_Kürzel = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Button_Teen
        '
        Me.Button_Teen.Enabled = False
        Me.Button_Teen.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_Teen.Location = New System.Drawing.Point(104, 167)
        Me.Button_Teen.Name = "Button_Teen"
        Me.Button_Teen.Size = New System.Drawing.Size(115, 46)
        Me.Button_Teen.TabIndex = 0
        Me.Button_Teen.Text = "Teen Sola"
        Me.Button_Teen.UseVisualStyleBackColor = True
        '
        'Button_Kids
        '
        Me.Button_Kids.Enabled = False
        Me.Button_Kids.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_Kids.Location = New System.Drawing.Point(252, 167)
        Me.Button_Kids.Name = "Button_Kids"
        Me.Button_Kids.Size = New System.Drawing.Size(115, 46)
        Me.Button_Kids.TabIndex = 1
        Me.Button_Kids.Text = "Kids Sola"
        Me.Button_Kids.UseVisualStyleBackColor = True
        '
        'Button_Cancel
        '
        Me.Button_Cancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_Cancel.Location = New System.Drawing.Point(411, 167)
        Me.Button_Cancel.Name = "Button_Cancel"
        Me.Button_Cancel.Size = New System.Drawing.Size(115, 46)
        Me.Button_Cancel.TabIndex = 2
        Me.Button_Cancel.Text = "Cancel"
        Me.Button_Cancel.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(70, 51)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(361, 20)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Bitte Geben sie ein eindeutiges Namenskürzel an:"
        '
        'TextBox_Kürzel
        '
        Me.TextBox_Kürzel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_Kürzel.Location = New System.Drawing.Point(474, 48)
        Me.TextBox_Kürzel.Name = "TextBox_Kürzel"
        Me.TextBox_Kürzel.Size = New System.Drawing.Size(87, 26)
        Me.TextBox_Kürzel.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(90, 112)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(450, 20)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Bitte wähle das Sola für das die Presets Erstellt werden Sollen:"
        '
        'LR_Preset_Input
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(631, 261)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox_Kürzel)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button_Cancel)
        Me.Controls.Add(Me.Button_Kids)
        Me.Controls.Add(Me.Button_Teen)
        Me.Name = "LR_Preset_Input"
        Me.Text = "Angaben für Lightroom"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button_Teen As Button
    Friend WithEvents Button_Kids As Button
    Friend WithEvents Button_Cancel As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents TextBox_Kürzel As TextBox
    Friend WithEvents Label2 As Label
End Class
