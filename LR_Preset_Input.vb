Public Class LR_Preset_Input
    Inherits Form
    Dim sPattern As String
    Dim bClose As Boolean = False
    Public Property Kürzel As String
        Get
            Return Me.TextBox_Kürzel.Text
        End Get
        Set(value As String)
            Me.TextBox_Kürzel.Text = value
        End Set
    End Property
    Public Property Kürzel_Regex As String
        Get
            Return sPattern
        End Get
        Set(value As String)
            sPattern = value
        End Set
    End Property
    Private Sub TextBox_Kürzel_TextChanged(sender As Object, e As EventArgs) Handles TextBox_Kürzel.MouseLeave
        'Const sPattern As String = "^[a-zA-Z]+(([',. -][a-zA-Z ])?[a-zA-Z]*)*$"
        Dim ValidateStrOutput As (boolErgebnis As Boolean, OutString As String)
        ValidateStrOutput = Main_Form.ValidateStr(sender.Text, sPattern)

        Me.TextBox_Kürzel.Text = ValidateStrOutput.OutString
        Me.Button_Teen.Enabled = ValidateStrOutput.boolErgebnis
        Me.Button_Kids.Enabled = ValidateStrOutput.boolErgebnis

    End Sub

    Private Sub Button_Teen_Click(sender As Object, e As EventArgs) Handles Button_Teen.Click
        Me.DialogResult = 1
        bClose = True
        Me.Close()
    End Sub

    Private Sub Button_Kids_Click(sender As Object, e As EventArgs) Handles Button_Kids.Click
        Me.DialogResult = 2
        bClose = True
        Me.Close()
    End Sub

    Private Sub Button_Cancel_Click(sender As Object, e As EventArgs) Handles Button_Cancel.Click
        Me.DialogResult = 5
        Me.TextBox_Kürzel.Text = ""
        bClose = True
        Me.Close()
    End Sub

    Private Sub LR_Preset_Input_Load(sender As Object, e As EventArgs) Handles MyBase.Closing
        If Not bClose Then Me.DialogResult = 5

    End Sub
End Class