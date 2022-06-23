<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Auswahl_Quality
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
        Me.Textbox_Pfad = New System.Windows.Forms.TextBox()
        Me.Cancel = New System.Windows.Forms.Button()
        Me.Button_RAW = New System.Windows.Forms.Button()
        Me.Button_LQ = New System.Windows.Forms.Button()
        Me.Button_HQ = New System.Windows.Forms.Button()
        Me.MsgText = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Textbox_Pfad
        '
        Me.Textbox_Pfad.Location = New System.Drawing.Point(32, 70)
        Me.Textbox_Pfad.Name = "Textbox_Pfad"
        Me.Textbox_Pfad.Size = New System.Drawing.Size(659, 20)
        Me.Textbox_Pfad.TabIndex = 13
        '
        'Cancel
        '
        Me.Cancel.Enabled = False
        Me.Cancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cancel.Location = New System.Drawing.Point(542, 119)
        Me.Cancel.Name = "Cancel"
        Me.Cancel.Size = New System.Drawing.Size(149, 39)
        Me.Cancel.TabIndex = 12
        Me.Cancel.Text = "Cancel"
        Me.Cancel.UseVisualStyleBackColor = True
        '
        'Button_RAW
        '
        Me.Button_RAW.Enabled = False
        Me.Button_RAW.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_RAW.Location = New System.Drawing.Point(372, 119)
        Me.Button_RAW.Name = "Button_RAW"
        Me.Button_RAW.Size = New System.Drawing.Size(149, 39)
        Me.Button_RAW.TabIndex = 11
        Me.Button_RAW.Text = "RAW"
        Me.Button_RAW.UseVisualStyleBackColor = True
        '
        'Button_LQ
        '
        Me.Button_LQ.Enabled = False
        Me.Button_LQ.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_LQ.Location = New System.Drawing.Point(202, 119)
        Me.Button_LQ.Name = "Button_LQ"
        Me.Button_LQ.Size = New System.Drawing.Size(149, 39)
        Me.Button_LQ.TabIndex = 10
        Me.Button_LQ.Text = "LowQuality (LQ)"
        Me.Button_LQ.UseVisualStyleBackColor = True
        '
        'Button_HQ
        '
        Me.Button_HQ.Enabled = False
        Me.Button_HQ.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_HQ.Location = New System.Drawing.Point(32, 119)
        Me.Button_HQ.Name = "Button_HQ"
        Me.Button_HQ.Size = New System.Drawing.Size(149, 39)
        Me.Button_HQ.TabIndex = 9
        Me.Button_HQ.Text = "HighQuality (HQ)"
        Me.Button_HQ.UseVisualStyleBackColor = True
        '
        'MsgText
        '
        Me.MsgText.AutoSize = True
        Me.MsgText.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MsgText.Location = New System.Drawing.Point(41, 9)
        Me.MsgText.Name = "MsgText"
        Me.MsgText.Size = New System.Drawing.Size(622, 40)
        Me.MsgText.TabIndex = 8
        Me.MsgText.Text = "Es konnte nicht erkant werden um welches Preset es sich handelt!" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Bitte wählen si" &
    "e anhand des Dateinamens am ende des Pfads den Richtigen Type aus."
        Me.MsgText.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Custom_Msgbox
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(730, 186)
        Me.Controls.Add(Me.Textbox_Pfad)
        Me.Controls.Add(Me.Cancel)
        Me.Controls.Add(Me.Button_RAW)
        Me.Controls.Add(Me.Button_LQ)
        Me.Controls.Add(Me.Button_HQ)
        Me.Controls.Add(Me.MsgText)
        Me.Name = "Custom_Msgbox"
        Me.Text = "Help"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Textbox_Pfad As TextBox
    Friend WithEvents Cancel As Button
    Friend WithEvents Button_RAW As Button
    Friend WithEvents Button_LQ As Button
    Friend WithEvents Button_HQ As Button
    Friend WithEvents MsgText As Label
End Class
