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
    'Sola Jahr automatisch zurückgeben
    Function SolaJahr(bTeens As Boolean, bkids As Boolean) As String
        Dim sJahr As String
        If bTeens And Not bkids Then    'anwahl Teensola
            sJahr = CStr(Year(Main_Form.DateTimePickerTeen.Value))  'auslesen des Startdatums Teen vom Form

        ElseIf Not bTeens And bkids Then    'anwahl kidssola
            sJahr = CStr(Year(Main_Form.DateTimePickerKids.Value))  'auslesen des Startdatums Kids vom Form
        ElseIf bTeens And bkids Then        'anwahl Kids und Teen Sola
            Dim temp(1)
            temp(0) = Year(Main_Form.DateTimePickerTeen.Value)  'auslesen des Startdatums Teen vom Form
            temp(1) = Year(Main_Form.DateTimePickerKids.Value)  'auslesen des Startdatums Kids vom Form
            If temp(0) = temp(1) Then       'vergleich das beide Startdaten im Selben Jahr liegen
                sJahr = CStr(temp(0))
            Else    'Meldung das Veschiedene Jahre gewählt wurden
                If MsgBox("Für das Teen und Kids Sola sind verschiedene Jahre ausgewählt!" & vbNewLine & "Teen: " & temp(0) & vbNewLine & "Kids: " & temp(1), vbOKCancel, "Achtung") = MsgBoxResult.Cancel Then
                    Exit Function
                ElseIf Not bSolaJahrMan Then
                    sJahr = InputBox("Bitte geben sie das Sola Jahr Manuell ein:", "Sola Jahr", CStr(temp(0)))      'manuele eingabe
                    bSolaJahrMan = True
                End If
            End If
        Else
            MsgBox("Es wurde kein SOLA ausgewählt !", vbExclamation)
            sJahr = ""
        End If
        Return sJahr
    End Function
    'auswertung des Rückgabewerts des LR_Preset_input Forms
    Function Select_Data_Preset() As (Valid As Integer, Year As String, Sola As String, Kürzel As String)

        Dim sJear As String
        Dim Dialog_Data_input As New LR_Preset_Input
        Dim sSola As String = ""
        Dim iValid As Integer = 0
        Dim sKürzel As String = ""
        'Form als Dialog anzeigen
        Dialog_Data_input.ShowDialog()
        'Rückgabe auswerten
        Select Case Dialog_Data_input.DialogResult
            Case 1                      'auswahl Teen
                sSola = "Teens"
            Case 2                      'auswahl Kids
                sSola = "Kids"
            Case Else                   'keine Wahl 
                sSola = ""
                iValid = 5
                Return (iValid, sJear, sSola, sKürzel)
                Exit Function
        End Select
        sKürzel = Dialog_Data_input.Kürzel      'eingetragenes Kürzel
        sJear = CStr(Format(Now, "yy"))         'aktuelles Jahr 
        'Pfüfen der Rückgabe Strings
        If String.IsNullOrEmpty(sJear) Or String.IsNullOrEmpty(sSola) Or String.IsNullOrEmpty(sKürzel) Then
            'MsgBox("Ungültige Eingabe", vbAbort)
            iValid = 0
        Else
            iValid = 1
        End If
        Return (iValid, sJear, sSola, sKürzel)      'Rückgabe 
    End Function
    'Bearbeiten der LR Presets auf der Festplatte
    Function Edit_Jear_in_Presets(Pfad As String, Select_Data As (Valid As Boolean, Year As String, Sola As String, Kürzel As String))
        Dim sFileOutput As String()
        Dim OK_Writeline(3) As (Name As String, Found As Boolean)
        Dim sFehlerBeschreibung As String = ""
        Dim bFault_Writelines As Boolean = False
        Dim sQuality_long As String = ""
        Dim sQuality_short As String = ""
        Dim myMsgbox As New Auswahl_Quality
        'Prüfen um welches Preset es sich in dem Pfad handelt
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
            Case Else                   'keine zuordnung des Presets möglich
                'Manuele abfrage des Presets
                With myMsgbox
                    .Text = "Help"
                    .Textbox_Pfad.Text = Pfad
                    .ShowDialog()
                    Select Case .DialogResult                       'manuele eingabe auswerten
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
                End With
        End Select

        'Presets Bearbeiten
        sFileOutput = System.IO.File.ReadAllLines(Pfad)             'Preset einlesen
        For iIndex = 0 To sFileOutput.Length - 1
            If sFileOutput(iIndex).IndexOf("internalName =") > -1 Then      'suche nach "internalName =" in der aktuellen zeile
                'Komplette Zeile ersetzen
                sFileOutput(iIndex) = vbTab & "internalName = " & Chr(34) & "SOLA" & Select_Data.Year & "_" & Select_Data.Sola & "_" & sQuality_long & Chr(34) & ","
                OK_Writeline(0).Name = "internalName"
                OK_Writeline(0).Found = True
                Continue For
            End If
            If sFileOutput(iIndex).IndexOf("title =") > -1 Then      'suche nach "title =" in der aktuellen zeile
                'Komplette Zeile ersetzen
                sFileOutput(iIndex) = vbTab & "title = " & Chr(34) & "SOLA" & Select_Data.Year & "_" & Select_Data.Sola & "_" & sQuality_long & Chr(34) & ","
                OK_Writeline(1).Name = "title"
                OK_Writeline(1).Found = True
                Continue For
            End If
            If sFileOutput(iIndex).IndexOf("tokenCustomString =") > -1 Then      'suche nach "tokenCustomString =" in der aktuellen zeile
                'Komplette Zeile ersetzen
                sFileOutput(iIndex) = vbTab & vbTab & "tokenCustomString = " & Chr(34) & Select_Data.Kürzel & Chr(34) & ","
                OK_Writeline(2).Name = "tokenCustomString"
                OK_Writeline(2).Found = True
                Continue For
            End If
            If sFileOutput(iIndex).IndexOf("tokens =") > -1 Then                  'suche nach "tokens =" in der aktuellen zeile
                'Komplette Zeile ersetzen
                sFileOutput(iIndex) = vbTab & vbTab & "tokens = " & Chr(34) & "{{date_YY}}-{{date_MM}}-{{date_DD}}__{{date_Hour}}-{{date_Minute}}-{{date_Second}}__SOLA" & Select_Data.Year & "_" & Select_Data.Sola & "_{{custom_token}}_" & sQuality_short & "_{{image_name}}" & Chr(34) & ","
                OK_Writeline(3).Name = "tokens"
                OK_Writeline(3).Found = True
                Continue For
            End If
        Next
        'Prüfen ob alles gefunden und Bearbeitet wurde
        For Each element In OK_Writeline
            If Not element.Found Then
                sFehlerBeschreibung = sFehlerBeschreibung & " " & Chr(34) & element.Name & Chr(34) & " Konte nicht gefunden werden" & vbNewLine    'Fehlerbeschreibung
                bFault_Writelines = True
            End If
        Next

        If Not bFault_Writelines Then
            System.IO.File.WriteAllLines(Pfad, sFileOutput)         'geänderte Datei Schreiben (Speichern)
            Return True
            Exit Function
        Else
            MsgBox(sFehlerBeschreibung, vbCritical, "Fehler !")     'Fehler Zeigen
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
        If x = "T" Then     'auswahl T
            For i = 1 To 10     'Alle Eingabe Felder Prüfen
                text = Main_Form.GroupBoxTeens.Controls("TextBox" & x & s & i).text
                If String.IsNullOrWhiteSpace(text) And String.IsNullOrEmpty(text) Then  'Prüfen auf Inhalt
                    Return True
                    Exit Function
                End If
            Next
        End If
        If x = "K" Then     'auswahl K
            For i = 1 To 10     'Alle Eingabe Felder Prüfen
                text = Main_Form.GroupBoxKids.Controls("TextBox" & x & s & i).text
                If String.IsNullOrWhiteSpace(text) And String.IsNullOrEmpty(text) Then  'Prüfen auf Inhalt
                    Return True
                    Exit Function
                End If
            Next
        End If
        Return False

    End Function
    'Kürzel mit RegX Prüfen 
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
    'Leere Array Elemente ans ende Schiben
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
