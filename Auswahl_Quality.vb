Public Class Auswahl_Quality
    Inherits Form

    Private Sub Button_HQ_Click(sender As Object, e As EventArgs) Handles Button_HQ.Click
        Me.DialogResult = 3
        Me.Close()
    End Sub

    Private Sub Button_LQ_Click(sender As Object, e As EventArgs) Handles Button_LQ.Click
        Me.DialogResult = 5
        Me.Close()
    End Sub

    Private Sub Button_RAW_Click(sender As Object, e As EventArgs) Handles Button_RAW.Click
        Me.DialogResult = 7
        Me.Close()
    End Sub

    Private Sub Cancel_Click(sender As Object, e As EventArgs) Handles Cancel.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub


End Class