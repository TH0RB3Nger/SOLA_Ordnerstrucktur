Imports System.Text.RegularExpressions
Public Class Tools
    Dim bSolaJahrMan As Boolean
    Public Property SolaJahrMan As Boolean
        Get
            Return bSolaJahrMan
        End Get
        Set(value As Boolean)
            bSolaJahrMan = value
        End Set
    End Property

    Function SolaJahr(bTeens As Boolean, bkids As Boolean) As String
        Dim sJahr As String
        If bTeens And Not bkids Then
            sJahr = CStr(Year(Main_Form.DateTimePickerTeen.Value))

        ElseIf Not bTeens And bkids Then
            sJahr = CStr(Year(Main_Form.DateTimePickerKids.Value))
        ElseIf bTeens And bkids Then
            Dim temp(1)
            temp(0) = Year(Main_Form.DateTimePickerTeen.Value)
            temp(1) = Year(Main_Form.DateTimePickerKids.Value)
            If temp(0) = temp(1) Then
                sJahr = CStr(temp(0))
            Else
                If MsgBox("Für das Teen und Kids Sola sind verschiedene Jahre ausgewählt!" & vbNewLine & "Teen: " & temp(0) & vbNewLine & "Kids: " & temp(1), vbOKCancel, "Achtung") = MsgBoxResult.Cancel Then
                    Exit Function
                ElseIf Not bSolaJahrMan Then
                    sJahr = InputBox("Bitte geben sie das Sola Jahr Manuell ein:", "Sola Jahr", CStr(temp(0)))
                    bSolaJahrMan = True
                End If
            End If
        Else
            MsgBox("Es wurde kein SOLA ausgewählt !", vbExclamation)
            sJahr = ""
        End If
        Return sJahr
    End Function

    Function Select_Data_Preset() As (Valid As Integer, Year As String, Sola As String, Kürzel As String)
        'Dim sJear As String
        'Dim sSOLA As String
        'Dim sKürzel As String
        'Dim bValid As Boolean = False
        'If MsgBox("Soll für das Teen Sola vorbereitet werden ?", vbYesNo) = vbYes Then sSOLA = "Teens" Else sSOLA = "Kids"
        'sKürzel = InputBox("Bitte gib dein Kürzel ein")
        ''Regex überprüfung
        'sJear = CStr(Format(Now, "yy"))
        'If String.IsNullOrEmpty(sJear) Or String.IsNullOrEmpty(sSOLA) Or String.IsNullOrEmpty(sKürzel) Then
        '    MsgBox("Ungültige Eingabe", vbAbort)
        'Else
        '    bValid = True
        'End If
        'Return (bValid, sJear, sSOLA, sKürzel)

        Dim sJear As String
        Dim Dialog_Data_input As New LR_Preset_Input
        Dim sSola As String = ""
        Dim iValid As Integer = 0
        Dim sKürzel As String = ""
        Dialog_Data_input.ShowDialog()
        Select Case Dialog_Data_input.DialogResult
            Case 1 'Teen
                sSola = "Teens"
            Case 2 'Kids
                sSola = "Kids"
            Case Else
                sSola = ""
                iValid = 5
                Return (iValid, sJear, sSola, sKürzel)
                Exit Function
        End Select
        sKürzel = Dialog_Data_input.Kürzel
        sJear = CStr(Format(Now, "yy"))
        If String.IsNullOrEmpty(sJear) Or String.IsNullOrEmpty(sSola) Or String.IsNullOrEmpty(sKürzel) Then
            'MsgBox("Ungültige Eingabe", vbAbort)
            iValid = 0
        Else
            iValid = 1
        End If
        Return (iValid, sJear, sSola, sKürzel)
    End Function
    Function Edit_Jear_in_Presets(Pfad As String, Select_Data As (Valid As Boolean, Year As String, Sola As String, Kürzel As String))
        Dim sFileOutput As String()
        Dim OK_Writeline(3) As (Name As String, Found As Boolean)
        Dim sFehlerBeschreibung As String = ""
        Dim bFault_Writelines As Boolean = False
        Dim sQuality_long As String = ""
        Dim sQuality_short As String = ""
        Dim myMsgbox As New Auswahl_Quality

        Select Case True
            Case Pfad.IndexOf("HighQuality_HQ") > -1
                sQuality_long = "HighQuality (HQ)"
                sQuality_short = "HQ"
            Case Pfad.IndexOf("LowQuality_LQ") > -1
                sQuality_long = "LowQuality (LQ)"
                sQuality_short = "LQ"
            Case Pfad.IndexOf("RAW") > -1
                sQuality_long = "RAW"
                sQuality_short = "RAW"
            Case Else
                myMsgbox.Text = "Help"
                myMsgbox.Textbox_Pfad.Text = Pfad
                myMsgbox.ShowDialog()
                Select Case myMsgbox.DialogResult
                    Case 3
                        sQuality_long = "HighQuality (HQ)"
                    Case 5
                        sQuality_long = "LowQuality (LQ)"
                    Case 7
                        sQuality_long = "RAW"
                    Case vbCancel
                        Return False
                        Exit Function
                    Case Else
                        Return False
                        Exit Function
                End Select
        End Select


        sFileOutput = System.IO.File.ReadAllLines(Pfad)
        For iIndex = 0 To sFileOutput.Length - 1
            If sFileOutput(iIndex).IndexOf("internalName =") > -1 Then
                sFileOutput(iIndex) = vbTab & "internalName = " & Chr(34) & "SOLA" & Select_Data.Year & "_" & Select_Data.Sola & "_" & sQuality_long & Chr(34) & ","
                OK_Writeline(0).Name = "internalName"
                OK_Writeline(0).Found = True
                Continue For
            End If
            If sFileOutput(iIndex).IndexOf("title =") > -1 Then
                sFileOutput(iIndex) = vbTab & "title = " & Chr(34) & "SOLA" & Select_Data.Year & "_" & Select_Data.Sola & "_" & sQuality_long & Chr(34) & ","
                OK_Writeline(1).Name = "title"
                OK_Writeline(1).Found = True
                Continue For
            End If
            If sFileOutput(iIndex).IndexOf("tokenCustomString =") > -1 Then
                sFileOutput(iIndex) = vbTab & vbTab & "tokenCustomString = " & Chr(34) & Select_Data.Kürzel & Chr(34) & ","
                OK_Writeline(2).Name = "tokenCustomString"
                OK_Writeline(2).Found = True
                Continue For
            End If
            If sFileOutput(iIndex).IndexOf("tokens =") > -1 Then
                sFileOutput(iIndex) = vbTab & vbTab & "tokens = " & Chr(34) & "{{date_YY}}-{{date_MM}}-{{date_DD}}__{{date_Hour}}-{{date_Minute}}-{{date_Second}}__SOLA" & Select_Data.Year & "_" & Select_Data.Sola & "_{{custom_token}}_" & sQuality_short & "_{{image_name}}" & Chr(34) & ","
                OK_Writeline(3).Name = "tokens"
                OK_Writeline(3).Found = True
                Continue For
            End If
        Next

        For Each element In OK_Writeline
            If Not element.Found Then
                sFehlerBeschreibung = sFehlerBeschreibung & " " & Chr(34) & element.Name & Chr(34) & " Konte nicht gefunden werden" & vbNewLine
                bFault_Writelines = True
            End If
        Next

        If Not bFault_Writelines Then
            System.IO.File.WriteAllLines(Pfad, sFileOutput)
            Return True
            Exit Function
        Else
            MsgBox(sFehlerBeschreibung, vbCritical, "Fehler !")
        End If
        Return False
    End Function
    Function NameNotEmpty(x As Char, y As Char) As Boolean
        Dim s
        Dim text As String
        If y = "F" Then
            s = "Foto"
        ElseIf y = "V" Then
            s = "Video"
        Else
            Exit Function
        End If
        If x = "T" Then
            For i = 1 To 10
                text = Main_Form.GroupBoxTeens.Controls("TextBox" & x & s & i).text
                If text.Length > 0 Then
                    NameNotEmpty = True
                    Exit Function
                End If
            Next
        End If
        If x = "K" Then
            For i = 1 To 10
                If Main_Form.GroupBoxKids.Controls("TextBox" & x & s & i).Text = "" Then
                    NameNotEmpty = True
                    Exit Function
                End If
            Next
        End If
        NameNotEmpty = False

    End Function
    Function ValidateStrKürzel(iStr As String) As (boolErgebnis As Boolean, inputString As String)
        ' Dim sPattern As String = "/^[a-zA-Z]+(([',. -][a-zA-Z ])?[a-zA-Z]*)*$/gi" 
        Dim rx As New Regex("^(?:(?:[a-zA-Z]{1,2}(?:\.))+)$")
        Dim Matches As MatchCollection = rx.Matches(iStr)
        If Matches.Count > 0 Then
            Return (True, Matches(0).Value)
        Else
            Return (False, "")
        End If
    End Function
    Function ValidateStr(iStr As String, Pattern As String) As (boolErgebnis As Boolean, inputString As String)
        ' Dim sPattern As String = "/^[a-zA-Z]+(([',. -][a-zA-Z ])?[a-zA-Z]*)*$/gi" 
        Dim rx As Regex
        If rx.IsMatch(iStr, Pattern, RegexOptions.IgnoreCase Or RegexOptions.Multiline) Then
            Return (True, iStr)
        Else
            Return (False, "")
        End If
    End Function
    Function DeletArrayGaps(Array)
        Dim tempArray(UBound(Array)) As String
        Dim j = 0
        For i = 0 To UBound(Array)
            If Array(i) <> "" Then
                ReDim Preserve tempArray(j)
                tempArray(j) = Array(i)

                j = j + 1
            End If
        Next
        ReDim Preserve tempArray(UBound(Array))
        Return tempArray
    End Function
End Class
