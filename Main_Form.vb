
Imports System.Resources

'Made By Thorben Renfordt
'TODO Funktionen die auserhalb genutzt werden auslagern
'TODO Code Komentieren!
Public Class Main_Form
    'TODO: Speichern und Laden Datumswerte unteschiedlich !!
    'Allgemeine Variablen (Daten)
    Const Version As String = "Alpha-v0.2.1"
    Const sDateFormat As String = "dd-MM-yyyy"
    Dim sSolaJahr As String
    Dim sNameTFoto(9), sNameTVideo(9), sNameKFoto(9), sNameKVideo(9)
    Dim bTFoto, bTVideo, bTShowfiles, bTInstagramm, bTGrafik, bTAudio, bTOrga, bTAllgemein
    Dim bKFoto, bKVideo, bKShowfiles, bKInstagramm, bKGrafik, bKAudio, bKOrga, bKAllgemein
    Dim dTTag(7), dKTag(7), sTTag(7), sKTag(7)
    Dim bTeens, bKids
    Dim bSolaJahrMan As Boolean
    Dim DL As Char = ";"
    Dim Tools As New Tools
    Dim Pfad

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.LinkLabel_Version.Text = Version
        Dim dMinDate As DateTime = DateAdd(DateInterval.DayOfYear, CDbl(DateDiff(DateInterval.Day, Now(), CDate(CStr(Year(Now()) & "/" & Month(Now()) - 1 & "/" & 1)))), Now())
        Dim dMaxDate As DateTime = DateAdd(DateInterval.Year, 5, Now())
        With Me.DateTimePickerTeen
            .Format = DateTimePickerFormat.Custom
            .CustomFormat = sDateFormat
            .MinDate = dMinDate
            .MaxDate = dMaxDate
            .Checked = False

        End With
        With Me.DateTimePickerKids
            .Format = DateTimePickerFormat.Custom
            .CustomFormat = sDateFormat
            .MinDate = dMinDate
            .MaxDate = dMaxDate
            .Checked = False
        End With

    End Sub

    Private Sub DataSave_Click(sender As Object, e As EventArgs) Handles DataSave.Click
        Dim sf As SaveFileDialog = New SaveFileDialog()
        Dim sPfad, sContent, sKopf
        sf.Filter = "CSV|*.csv"
        sf.FilterIndex = 2
        sf.RestoreDirectory = True
        sSolaJahr = Tools.SolaJahr(bTeens, bKids)

        If sf.ShowDialog() = DialogResult.OK Then
            sPfad = sf.FileName
            Dim Writer As New System.IO.StreamWriter(sf.FileName)
            sKopf = "SolaJahr" & DL & "Teens" & DL & "Kids" & DL & "NameTeenFotograf1" & DL & "NameTeenFotograf2" & DL & "NameTeenFotograf3" & DL & "NameTeenFotograf4" & DL & "NameTeenFotograf5" & DL & "NameTeenFotograf6" & DL & "NameTeenFotograf7" & DL & "NameTeenFotograf8" & DL & "NameTeenFotograf9" & DL & "NameTeenFotograf10"
            sKopf = sKopf & DL & "NameTeenVideograf1" & DL & "NameTeenVideograf2" & DL & "NameTeenVideograf3" & DL & "NameTeenVideograf4" & DL & "NameTeenVideograf5" & DL & "NameTeenVideograf6" & DL & "NameTeenVideograf7" & DL & "NameTeenVideograf8" & DL & "NameTeenVideograf9" & DL & "NameTeenVideograf10"
            sKopf = sKopf & DL & "NameKidsFotograf1" & DL & "NameKidsFotograf2" & DL & "NameKidsFotograf3" & DL & "NameKidsFotograf4" & DL & "NameKidsFotograf5" & DL & "NameKidsFotograf6" & DL & "NameKidsFotograf7" & DL & "NameKidsFotograf8" & DL & "NameKidsFotograf9" & DL & "NameKidsFotograf10"
            sKopf = sKopf & DL & "NameKidsVideograf1" & DL & "NameKidsVideograf2" & DL & "NameKidsVideograf3" & DL & "NameKidsVideograf4" & DL & "NameKidsVideograf5" & DL & "NameKidsVideograf6" & DL & "NameKidsVideograf7" & DL & "NameKidsVideograf8" & DL & "NameKidsVideograf9" & DL & "NameKidsVideograf10"
            sKopf = sKopf & DL & "AuswahlTeenFoto" & DL & "AuswahlTeenVideo" & DL & "AuswahlTeenShowfiles" & DL & "AuswahlTeenInstagramm" & DL & "AuswahlTeenGrafik" & DL & "AuswahlTeenAudio" & DL & "AuswahlTeenOrga" & DL & "AuswahlTeenAllgemein"
            sKopf = sKopf & DL & "AuswahlKidsFoto" & DL & "AuswahlKidsVideo" & DL & "AuswahlKidsShowfiles" & DL & "AuswahlKidsInstagramm" & DL & "AuswahlKidsGrafik" & DL & "AuswahlKidsAudio" & DL & "AuswahlKidsOrga" & DL & "AuswahlKidsAllgemein"
            sKopf = sKopf & DL & "TeenStartDatum" & DL & "KidsStartDatum"

            sContent = sSolaJahr & DL & bTeens & DL & bKids
            For i = 0 To 9
                sContent = sContent & DL & sNameTFoto(i)
            Next
            For i = 0 To 9
                sContent = sContent & DL & sNameTVideo(i)
            Next
            For i = 0 To 9
                sContent = sContent & DL & sNameKFoto(i)
            Next
            For i = 0 To 9
                sContent = sContent & DL & sNameKVideo(i)
            Next
            sContent = sContent & DL & bTFoto & DL & bTVideo & DL & bTShowfiles & DL & bTInstagramm & DL & bTGrafik & DL & bTAudio & DL & bTOrga & DL & bTAllgemein
            sContent = sContent & DL & bKFoto & DL & bKVideo & DL & bKShowfiles & DL & bKInstagramm & DL & bKGrafik & DL & bKAudio & DL & bKOrga & DL & bKAllgemein

            sContent = sContent & DL & CStr(dTTag(0)) & DL & CStr(dKTag(0))

            ' If FileLineCount(File) = 0 Then
            Writer.WriteLine(sKopf)

            ' End If
            Writer.WriteLine(sContent)


            Writer.Close()
        End If


    End Sub

    Private Sub DataLoad_Click(sender As Object, e As EventArgs) Handles DataLoad.Click

        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim strFileName As String
        Dim iField As Integer = 0
        fd.Title = "Konfiguration Wählen"
        'fd.InitialDirectory = Application.ExecutablePath
        fd.Filter = "CSV|*.csv"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.FileName
            Using Reader As New Microsoft.VisualBasic.FileIO.TextFieldParser(strFileName)
                Reader.TextFieldType = FileIO.FieldType.Delimited
                Reader.SetDelimiters(DL)
                Reader.ReadLine()

                Dim currentRow As String()
                Try
                    currentRow = Reader.ReadFields()
                    sSolaJahr = currentRow(0)
                    bTeens = CBool(currentRow(1))
                    bKids = CBool(currentRow(2))
                    For i = 0 To 9
                        sNameTFoto(i) = currentRow(3 + i)
                    Next
                    For i = 0 To 9
                        sNameTVideo(i) = currentRow(13 + i)
                    Next
                    For i = 0 To 9
                        sNameKFoto(i) = currentRow(23 + i)
                    Next
                    For i = 0 To 9
                        sNameKVideo(i) = currentRow(33 + i)
                    Next
                    bTFoto = CBool(currentRow(43))
                    bTVideo = CBool(currentRow(44))
                    bTShowfiles = CBool(currentRow(45))
                    bTInstagramm = CBool(currentRow(46))
                    bTGrafik = CBool(currentRow(47))
                    bTAudio = CBool(currentRow(48))
                    bTOrga = CBool(currentRow(49))
                    bTAllgemein = CBool(currentRow(50))
                    bKFoto = CBool(currentRow(51))
                    bKVideo = CBool(currentRow(52))
                    bKShowfiles = CBool(currentRow(53))
                    bKInstagramm = CBool(currentRow(54))
                    bKGrafik = CBool(currentRow(55))
                    bKAudio = CBool(currentRow(56))
                    bKOrga = CBool(currentRow(57))
                    bKAllgemein = CBool(currentRow(58))
                    dTTag(0) = CDate(currentRow(59))
                    dKTag(0) = CDate(currentRow(60))
                    'Dim currentField As String
                    'For Each currentField In currentRow
                    'sContent(iField) = currentField
                    'iField = iField + 1
                    'Next
                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message & " Is not valid and will be Skipped.")
                End Try

            End Using
            NewDataShow()

        End If
    End Sub
    Sub NewDataShow()
        CheckBoxTeen.Checked = bTeens
        CheckBoxKids.Checked = bKids
        DateTimePickerTeen.Value = dTTag(0)
        DateTimePickerKids.Value = dKTag(0)
        CheckedListBoxTeen.SetItemChecked(0, bTFoto)
        CheckedListBoxTeen.SetItemChecked(1, bTVideo)
        CheckedListBoxTeen.SetItemChecked(2, bTShowfiles)
        CheckedListBoxTeen.SetItemChecked(3, bTInstagramm)
        CheckedListBoxTeen.SetItemChecked(4, bTGrafik)
        CheckedListBoxTeen.SetItemChecked(5, bTAudio)
        CheckedListBoxTeen.SetItemChecked(6, bTOrga)
        CheckedListBoxTeen.SetItemChecked(7, bTAllgemein)
        CheckedListBoxKids.SetItemChecked(0, bKFoto)
        CheckedListBoxKids.SetItemChecked(1, bKVideo)
        CheckedListBoxKids.SetItemChecked(2, bKShowfiles)
        CheckedListBoxKids.SetItemChecked(3, bKInstagramm)
        CheckedListBoxKids.SetItemChecked(4, bKGrafik)
        CheckedListBoxKids.SetItemChecked(5, bKAudio)
        CheckedListBoxKids.SetItemChecked(6, bKOrga)
        CheckedListBoxKids.SetItemChecked(7, bKAllgemein)
        For i = 1 To 10
            Me.GroupBoxTeens.Controls("TextBoxTFoto" & i).Text = sNameTFoto(i - 1)
            Me.GroupBoxTeens.Controls("TextBoxTVideo" & i).Text = sNameTVideo(i - 1)
            Me.GroupBoxKids.Controls("TextBoxKFoto" & i).Text = sNameKFoto(i - 1)
            Me.GroupBoxKids.Controls("TextBoxKVideo" & i).Text = sNameKVideo(i - 1)
        Next

    End Sub



    Private Sub Button1_Click_Install_LR_Preset(sender As Object, e As EventArgs) Handles Button1.Click
        'Exportiert die Lightrom Presets
        Const sDestPfad_devlop_preset As String = "\Adobe\Lightroom\Develop Presets"
        Const sExtention_devlop_preset As String = ".xmp"
        Const sExtention_export_Preset As String = ".lrtemplate"
        Const sDestPfad_export_Preset As String = "\Adobe\Lightroom\Export Presets\User Presets"
        Const maxRetry As Integer = 3
        Dim Select_Data_Valid As (IsValid As Integer, NotValid As Integer, cancel As Integer) = (1, 0, 5)
        Dim iRetry As Integer = 0
        Dim Select_Data As (Valid As Integer, Year As String, Sola As String, Kürzel As String) = (False, "", "", "")
        Dim i As Integer = 0
Retry:
        Dim bOK_create As Boolean
        Dim bOK_change As Boolean
        Dim aOK_create(4) As Boolean
        Dim aOK_change(4) As Boolean
        Dim appdata As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        Dim sPfadHQ As String = appdata & sDestPfad_export_Preset & "\" & "HighQuality_HQ_" & sExtention_export_Preset
        Dim sPfadLQ As String = appdata & sDestPfad_export_Preset & "\" & "LowQuality_LQ_" & sExtention_export_Preset
        Dim sPfadRAW As String = appdata & sDestPfad_export_Preset & "\" & "RAW" & sExtention_export_Preset
        Dim sPfad_devlop_SOLA_Draußen As String = appdata & sDestPfad_devlop_preset & "\" & "SOLA_Draußen" & sExtention_devlop_preset
        Dim sPfad_devlop_SOLA_Veranstaltungszelt As String = appdata & sDestPfad_devlop_preset & "\" & "SOLA_Veranstaltungszelt" & sExtention_devlop_preset
        Dim sOk_create(4) As String
        Dim sOk_change(4) As String
        Dim x As Integer = -1
        Dim y As Integer = -1
        Dim doRetry As Boolean = False


        Select_Data = Tools.Select_Data_Preset()
        Select Case Select_Data.Valid
            Case Select_Data_Valid.IsValid
                Exit Select
            Case Select_Data_Valid.NotValid
                Select_Data = Tools.Select_Data_Preset()
                Select Case Select_Data.Valid
                    Case Select_Data_Valid.IsValid
                        Exit Select
                    Case Select_Data_Valid.NotValid
                        Select_Data = Tools.Select_Data_Preset()
                        Select Case Select_Data.Valid
                            Case Select_Data_Valid.IsValid
                                Exit Select
                            Case Select_Data_Valid.NotValid
                                MsgBox("Ungültige eingaben", vbCritical)
                                Exit Sub
                            Case Select_Data_Valid.cancel
                                Exit Sub
                        End Select
                    Case Select_Data_Valid.cancel
                        Exit Sub
                End Select
            Case Select_Data_Valid.cancel
                Exit Sub
        End Select




        System.IO.File.WriteAllBytes(sPfadHQ, My.Resources.Resource1.HighQuality__HQ_)
        aOK_change(0) = Tools.Edit_Jear_in_Presets(sPfadHQ, Select_Data)
        System.IO.File.WriteAllBytes(sPfadLQ, My.Resources.Resource1.LowQuality__LQ_)
        aOK_change(1) = Tools.Edit_Jear_in_Presets(sPfadLQ, Select_Data)
        System.IO.File.WriteAllBytes(sPfadRAW, My.Resources.Resource1.RAW)
        aOK_change(2) = Tools.Edit_Jear_in_Presets(sPfadRAW, Select_Data)

        If System.IO.File.Exists(sPfadHQ) And System.IO.File.Exists(sPfadLQ) And System.IO.File.Exists(sPfadRAW) Then
            bOK_create = True
            For i = 0 To 2
                aOK_create(i) = True
            Next
        Else
            aOK_create(0) = System.IO.File.Exists(sPfadHQ)
            aOK_create(1) = System.IO.File.Exists(sPfadLQ)
            aOK_create(2) = System.IO.File.Exists(sPfadRAW)
        End If

        For i = 0 To 2
            If aOK_change(i) Then
                bOK_change = True
            Else
                bOK_change = False
                Exit For
            End If
        Next


        System.IO.File.WriteAllBytes(sPfad_devlop_SOLA_Draußen, My.Resources.Resource1.draußen)
        System.IO.File.WriteAllBytes(sPfad_devlop_SOLA_Veranstaltungszelt, My.Resources.Resource1.Veranstaltungszelt)
        If System.IO.File.Exists(sPfad_devlop_SOLA_Draußen) And System.IO.File.Exists(sPfad_devlop_SOLA_Veranstaltungszelt) Then
            For i = 3 To 4
                aOK_create(i) = True
            Next
            If bOK_create And bOK_change Then
                MsgBox("Exportvorlagen und Presets erfolgreich kopiert." & vbNewLine & "Exportvorlagen Erfolgreich angepasst", MsgBoxStyle.Information, "LR Preset erfolgreich kopiert")
            Else
                GoTo Diag
            End If

        Else
            aOK_create(3) = System.IO.File.Exists(sPfad_devlop_SOLA_Draußen)
            aOK_create(4) = System.IO.File.Exists(sPfad_devlop_SOLA_Veranstaltungszelt)

            GoTo Diag
        End If
        Exit Sub
Diag:

        If iRetry < maxRetry Then
            If Not bOK_create Then
                For Each Bool In aOK_create
                    x = x + 1
                    If Bool Then sOk_create(x) = "OK" Else sOk_create(x) = "NOK"
                Next
                If MsgBox("Fehler beim erstellen der vorgaben" & vbNewLine & "Exportvorgabe HQ: " & sOk_create(0) & vbNewLine _
                        & "Exportvorgabe LQ: " & sOk_create(1) & vbNewLine _
                        & "Exportvorgabe RAW: " & sOk_create(2) & vbNewLine _
                        & "Entwicklungsvorgabe Draußen: " & sOk_create(3) & vbNewLine _
                        & "Entwicklungsvorgabe Veranstaltungszelt: " & sOk_create(4), MsgBoxStyle.RetryCancel, "Fehler") = MsgBoxResult.Retry Then

                    doRetry = True
                End If
            End If

            If Not bOK_change Then
                For Each Bool As Boolean In aOK_change
                    y = y + 1
                    If Bool Then sOk_change(y) = "OK" Else sOk_change(x) = "NOK"
                Next
                If MsgBox("Fehler beim ändern der vorgaben" & vbNewLine & "Exportvorgabe HQ: " & sOk_change(0) & vbNewLine _
                            & "Exportvorgabe LQ: " & sOk_change(1) & vbNewLine _
                            & "Exportvorgabe RAW: " & sOk_change(2) & vbNewLine, MsgBoxStyle.RetryCancel, "Fehler") = MsgBoxResult.Retry Then

                    doRetry = True
                End If
            End If
            If doRetry Then
                iRetry = iRetry + 1
                GoTo Retry
            End If
        Else
            MsgBox("Fehler beim erstellen der vorgaben" & vbNewLine & "Exportvorgabe HQ: " & sOk_create(0) & vbNewLine _
            & "Exportvorgabe LQ: " & sOk_create(1) & vbNewLine _
            & "Exportvorgabe RAW: " & sOk_create(2) & vbNewLine _
            & "Entwicklungsvorgabe Draußen: " & sOk_create(3) & vbNewLine _
            & "Entwicklungsvorgabe Veranstaltungszelt: " & sOk_create(4) & vbNewLine _
            & "Fehler beim ändern der vorgaben" & vbNewLine & "Exportvorgabe HQ: " & sOk_change(0) & vbNewLine _
            & "Exportvorgabe LQ: " & sOk_change(1) & vbNewLine _
            & "Exportvorgabe RAW: " & sOk_change(2), MsgBoxStyle.OkCancel, "Fehler")
        End If
    End Sub



    Private Sub DateTimePickerTeen_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePickerTeen.ValueChanged
        Tools.SolaJahrMan = False
        Berechne_Woche(sender.Value, True)
    End Sub



    Private Sub DateTimePickerKids_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePickerKids.ValueChanged
        bSolaJahrMan = False
        Berechne_Woche(sender.Value, False)
    End Sub



    Sub Berechne_Woche(startDate As Date, Teen As Boolean) 'TODO Umbauen damit es als Funktion in die Tools Klasse past
        Dim dTag As Date
        If Teen Then
            dTTag(0) = Format(startDate, sDateFormat)  '"yy-MM-dd"
            sTTag(0) = CStr(dTTag(0))
            For i = 1 To 7
                dTag = DateAdd("d", i, startDate)
                dTTag(i) = Format(dTag, sDateFormat)
                sTTag(i) = CStr(dTTag(i))
            Next
        Else
            dKTag(0) = Format(startDate, sDateFormat)
            sKTag(0) = CStr(dKTag(0))
            For i = 1 To 7
                dTag = DateAdd("d", i, startDate)
                dKTag(i) = Format(dTag, sDateFormat)
                sKTag(i) = CStr(dKTag(i))
            Next
        End If
    End Sub

    Private Sub ButtonStrucktur_Click(sender As Object, e As EventArgs) Handles ButtonStrucktur.Click
        If Not String.IsNullOrEmpty(Pfad) Then
            Ordnerstrucktur_Erstellen()
        Else
            If MsgBox("Es wurde kein Ordner ausgewählt!" & vbCrLf & "Ordner jetzt wählen ?", vbOKCancel) = vbOK Then
                Ordnerauswahl_ini()
            End If
        End If

    End Sub



    Private Sub TextBoxTFoto1_TextChanged(sender As Object, e As EventArgs) Handles TextBoxTFoto1.Leave, TextBoxTFoto2.Leave, TextBoxTFoto3.Leave, TextBoxTFoto4.Leave, TextBoxTFoto5.Leave, TextBoxTFoto6.Leave, TextBoxTFoto7.Leave, TextBoxTFoto8.Leave, TextBoxTFoto9.Leave, TextBoxTFoto10.Leave, TextBoxTVideo1.Leave, TextBoxTVideo2.Leave, TextBoxTVideo3.Leave, TextBoxTVideo4.Leave, TextBoxTVideo5.Leave, TextBoxTVideo6.Leave, TextBoxTVideo7.Leave, TextBoxTVideo8.Leave, TextBoxTVideo9.Leave, TextBoxTVideo10.Leave, TextBoxKFoto1.Leave, TextBoxKFoto2.Leave, TextBoxKFoto3.Leave, TextBoxKFoto4.Leave, TextBoxKFoto5.Leave, TextBoxKFoto6.Leave, TextBoxKFoto7.Leave, TextBoxKFoto8.Leave, TextBoxKFoto9.Leave, TextBoxKFoto10.Leave, TextBoxKVideo1.Leave, TextBoxKVideo2.Leave, TextBoxKVideo3.Leave, TextBoxKVideo4.Leave, TextBoxKVideo5.Leave, TextBoxKVideo6.Leave, TextBoxKVideo7.Leave, TextBoxKVideo8.Leave, TextBoxKVideo9.Leave, TextBoxKVideo10.Leave
        Const sPattern As String = "^[a-zA-Z]+(([',. -][a-zA-Z ])?[a-zA-Z]*)*$" 'TODO: Aktuell sind keine Zahlen hinter dem Namen erlaubt
        If Not Tools.ValidateStr(sender.Text, sPattern).boolErgebnis Then
            If Not String.IsNullOrEmpty(sender.Text) Then
                MsgBox("Ungültige Eingabe", MsgBoxStyle.OkOnly)
            End If
            sender.Text = ""
            Exit Sub
        End If
        Dim Sola, Bereich, sNr
        Dim iNr As Integer
        Sola = sender.Name(7)
        Bereich = Strings.Mid(sender.Name, 9, 1)
        sNr = Strings.Right(sender.Name, 2)
        If Not IsNumeric(sNr(0)) Then
            sNr = sNr(1)
            iNr = Char.GetNumericValue(sNr)
        Else
            iNr = CInt(sNr)
        End If

        If Sola = "T" Then
            If Bereich = "F" Then
                sNameTFoto(iNr - 1) = sender.Text
            ElseIf Bereich = "V" Then
                sNameTVideo(iNr - 1) = sender.Text
            End If

        ElseIf Sola = "K" Then
            If Bereich = "F" Then
                sNameKFoto(iNr - 1) = sender.Text
            ElseIf Bereich = "V" Then
                sNameKVideo(iNr - 1) = sender.Text
            End If
        End If
        Freigabe_Erstellen()
    End Sub
    Sub Ordnerauswahl_ini()
        Pfad = Ordnerwahl("Bitte Verzeichnis zum Erstellen der Ordnerstrucktur Wählen")
        LabelPfad.Text = Pfad
        Freigabe_Erstellen()
    End Sub

    Private Sub ButtonPicPfad_Click(sender As Object, e As EventArgs) Handles ButtonPicPfad.Click
        Ordnerauswahl_ini()
    End Sub

    Private Sub LinkLabel_Version_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel_Version.LinkClicked
        System.Diagnostics.Process.Start("https://github.com/TH0RB3Nger/SOLA_Ordnerstrucktur/releases")
        LinkLabel_Version.LinkVisited = True
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        'Link zu dem Sola Wiedenest Instagram Account
        System.Diagnostics.Process.Start("https://www.instagram.com/sola_wiedenest/")
        LinkLabel2.LinkVisited = True
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        'Link zu dem Torsten Instagrm Account
        System.Diagnostics.Process.Start("https://www.instagram.com/der.torsten/")
        LinkLabel1.LinkVisited = True
    End Sub
    Private Sub CheckBoxTeen_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxTeen.CheckedChanged
        'Wenn aktiviert werden die Elemente freigegeben und der Zustand in der Variable gespeichert
        If CheckBoxTeen.Checked Then
            DateTimePickerTeen.Enabled = True
            CheckedListBoxTeen.Enabled = True
            bTeens = True
        Else
            DateTimePickerTeen.Enabled = False
            CheckedListBoxTeen.Enabled = False
            bTeens = False
        End If
        EnableNamenTeens()

    End Sub
    Private Sub CheckBoxKids_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxKids.CheckedChanged
        'Wenn aktiviert werden die Elemente freigegeben und der Zustand in der Variable gespeichert
        If CheckBoxKids.Checked Then
            DateTimePickerKids.Enabled = True
            CheckedListBoxKids.Enabled = True
            bKids = True
        Else
            DateTimePickerKids.Enabled = False
            CheckedListBoxKids.Enabled = False
            bKids = False
        End If
        EnableNamenKids()

    End Sub
    Private Sub CheckedListBoxTeen_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBoxTeen.SelectedIndexChanged
        'Auswerten der Auswahl und ablegen in den Variablen
        bTFoto = CheckedListBoxTeen.GetItemChecked(0)
        bTVideo = CheckedListBoxTeen.GetItemChecked(1)
        bTShowfiles = CheckedListBoxTeen.GetItemChecked(2)
        bTInstagramm = CheckedListBoxTeen.GetItemChecked(3)
        bTGrafik = CheckedListBoxTeen.GetItemChecked(4)
        bTAudio = CheckedListBoxTeen.GetItemChecked(5)
        bTOrga = CheckedListBoxTeen.GetItemChecked(6)
        bTAllgemein = CheckedListBoxTeen.GetItemChecked(7)
        EnableNamenTeens()
    End Sub
    Private Sub CheckedListBoxKids_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBoxKids.SelectedIndexChanged
        'Auswerten der Auswahl und ablegen in den Variablen
        bKFoto = CheckedListBoxKids.GetItemChecked(0)
        bKVideo = CheckedListBoxKids.GetItemChecked(1)
        bKShowfiles = CheckedListBoxKids.GetItemChecked(2)
        bKInstagramm = CheckedListBoxKids.GetItemChecked(3)
        bKGrafik = CheckedListBoxKids.GetItemChecked(4)
        bKAudio = CheckedListBoxKids.GetItemChecked(5)
        bKOrga = CheckedListBoxKids.GetItemChecked(6)
        bKAllgemein = CheckedListBoxKids.GetItemChecked(7)
        EnableNamenKids()
    End Sub
    Private Sub EnableNamenTeens()
        'Namensfelder aktivieren wenn entsprechende auswahl gegeben ist
        If bTeens Then
            For i = 1 To 10
                If bTFoto Then
                    Me.GroupBoxTeens.Controls("TextBoxTFoto" & i).Enabled = True
                Else
                    Me.GroupBoxTeens.Controls("TextBoxTFoto" & i).Enabled = False
                End If
                If bTVideo Then
                    Me.GroupBoxTeens.Controls("TextBoxTVideo" & i).Enabled = True
                Else
                    Me.GroupBoxTeens.Controls("TextBoxTVideo" & i).Enabled = False
                End If
            Next
        Else
            For i = 1 To 10
                Me.GroupBoxTeens.Controls("TextBoxTFoto" & i).Enabled = False
                Me.GroupBoxTeens.Controls("TextBoxTVideo" & i).Enabled = False
            Next
        End If
        Freigabe_Erstellen()
    End Sub
    Private Sub EnableNamenKids()
        'Namensfelder aktivieren wenn entsprechende auswahl gegeben ist
        If bKids Then
            For i = 1 To 10
                If bKFoto Then
                    Me.GroupBoxKids.Controls("TextBoxKFoto" & i).Enabled = True
                Else
                    Me.GroupBoxKids.Controls("TextBoxKFoto" & i).Enabled = False
                End If
                If bKVideo Then
                    Me.GroupBoxKids.Controls("TextBoxKVideo" & i).Enabled = True
                Else
                    Me.GroupBoxKids.Controls("TextBoxKVideo" & i).Enabled = False
                End If
            Next
        Else
            For i = 1 To 10
                Me.GroupBoxKids.Controls("TextBoxKFoto" & i).Enabled = False
                Me.GroupBoxKids.Controls("TextBoxKVideo" & i).Enabled = False
            Next
        End If
        Freigabe_Erstellen()
    End Sub
    Function Ordnerwahl(msg As String) As String
        ''Auswahldialog für die Ordnerwahl

        OokiiDialog.Multiselect = False
        OokiiDialog.ShowNewFolderButton = True
        OokiiDialog.Description = msg
        OokiiDialog.UseDescriptionForTitle = True

        OokiiDialog.ShowDialog()
        If String.IsNullOrEmpty(OokiiDialog.SelectedPath) Then
            MsgBox("Es wurde kein Ordner ausgewählt!", vbExclamation)
            Return vbEmpty
            Exit Function
        End If
        Return OokiiDialog.SelectedPath

    End Function
    Private Sub Freigabe_Erstellen()
        If Not (Pfad = "" Or String.IsNullOrEmpty(Pfad)) Then
            If bTeens And Not bKids Then
                If (bTShowfiles Or bTInstagramm Or bTGrafik Or bTAudio Or bTOrga Or bTAllgemein) Or (bTFoto And Tools.NameNotEmpty("T", "F")) Or (bTVideo And Tools.NameNotEmpty("T", "V")) Then
                    ButtonStrucktur.Enabled = True
                Else
                    ButtonStrucktur.Enabled = False
                End If
            ElseIf bKids And Not bTeens Then
                If (bKShowfiles Or bKInstagramm Or bKGrafik Or bKAudio Or bKOrga Or bKAllgemein) Or (bKFoto And Tools.NameNotEmpty("K", "F")) Or (bKVideo And Tools.NameNotEmpty("K", "V")) Then
                    ButtonStrucktur.Enabled = True
                Else
                    ButtonStrucktur.Enabled = False
                End If
            ElseIf bTeens And bKids Then
                If (bTShowfiles Or bTInstagramm Or bTGrafik Or bTAudio Or bTOrga Or bTAllgemein) Or (bTFoto And Tools.NameNotEmpty("T", "F")) Or (bTVideo And Tools.NameNotEmpty("T", "V")) _
                    And (bKShowfiles Or bKInstagramm Or bKGrafik Or bKAudio Or bKOrga Or bKAllgemein) Or (bKFoto And Tools.NameNotEmpty("K", "F")) Or (bKVideo And Tools.NameNotEmpty("K", "V")) Then
                    ButtonStrucktur.Enabled = True
                Else
                    ButtonStrucktur.Enabled = False
                End If
            Else
                ButtonStrucktur.Enabled = False
            End If
        Else
            ButtonStrucktur.Enabled = False
        End If
    End Sub


    Sub Ordnerstrucktur_Erstellen()
        'Thorben Renfordt

        'Allgemeine Variablen deklaration:
        Dim i
        Dim oFSO As Object
        Dim sPfad As String
        Dim iOrdner As Integer

        'Variablen Inizialisation
        oFSO = CreateObject("Scripting.FileSystemObject")
        i = 0

        'Stammverzeichniss Wählen
        Delete_Gaps_in_Name()
        Dim j
        Dim Tag
        Dim iZeile
        Dim Tdone(7), Kdone(7)
        Dim sString1
        Dim sPerOrdnerFoto(3)
        Dim FDL
        FDL = "\"
        sSolaJahr = Tools.SolaJahr(bTeens, bKids)
        sPerOrdnerFoto(0) = "01_ImportRAW"
        sPerOrdnerFoto(1) = "02_ExportJPEG_HQ"
        sPerOrdnerFoto(2) = "03_ExportJPEG_LQ"
        sPerOrdnerFoto(3) = "04_ExportRAW"

        MkDir(Pfad & FDL & "Sola_" & sSolaJahr)
        sPfad = Pfad & FDL & "Sola_" & sSolaJahr
        iOrdner = 1
        If bTeens Then
            Chek_Tag_Array(sTTag, DateTimePickerTeen, True)
            MkDir(sPfad & FDL & "01_Teens")
            sPfad = sPfad & FDL & "01_Teens"
            iOrdner = iOrdner + 1
            'Range("B2").Value = "02_Kids"
            iZeile = 0
            For i = 1 To 8
                sString1 = "0" & i
                sPfad = Pfad & FDL & "Sola_" & sSolaJahr & FDL & "01_Teens"
                Select Case True
            'Foto Ordner
                    Case bTFoto And Not Tdone(0)
                        sString1 = sString1 & "_Foto"
                        Tdone(0) = True
                        MkDir(sPfad & FDL & sString1)
                        iOrdner = iOrdner + 1
                        sPfad = sPfad & FDL & sString1

                        For Tag = 1 To 8
                            MkDir(sPfad & FDL & sTTag(Tag - 1) & "_Tag_" & Tag)
                            iOrdner = iOrdner + 1
                            sPfad = sPfad & FDL & sTTag(Tag - 1) & "_Tag_" & Tag
                            MkDir(sPfad & FDL & "01_Bilder des Tages" & Tag & "_HQ")
                            iOrdner = iOrdner + 1
                            MkDir(sPfad & FDL & "02_Bilder des Tages" & Tag & "_LQ")
                            iOrdner = iOrdner + 1
                            MkDir(sPfad & FDL & "03_Auswahl Bilderclip")
                            iOrdner = iOrdner + 1
                            MkDir(sPfad & FDL & "04_Auswahl Musik")
                            iOrdner = iOrdner + 1
                            For j = 1 To 10
                                If Not sNameTFoto(j - 1) = "" Then
                                    If j + 4 > 9 Then
                                        MkDir(sPfad & FDL & j + 4 & "_" & sNameTFoto(j - 1))
                                        iOrdner = iOrdner + 1
                                        sPfad = sPfad & FDL & j + 4 & "_" & sNameTFoto(j - 1)
                                    Else
                                        sPfad = sPfad & FDL & "0" & j + 4 & "_" & sNameTFoto(j - 1)
                                        MkDir(sPfad)
                                        iOrdner = iOrdner + 1
                                        'sPfad = sPfad & FDL & "0" & j + 4 & "_" & sNameTFoto(j - 1)
                                    End If

                                    For k = 0 To 3
                                        MkDir(sPfad & FDL & sPerOrdnerFoto(k))
                                        iOrdner = iOrdner + 1
                                    Next
                                    'sPfad = sPfad & FDL & sTTag(Tag - 1) & "_Tag" & Tag
                                    sPfad = Pfad & FDL & "Sola_" & sSolaJahr & FDL & "01_Teens" & FDL & sString1 & FDL & sTTag(Tag - 1) & "_Tag_" & Tag
                                End If
                            Next
                            sPfad = Pfad & FDL & "Sola_" & sSolaJahr & FDL & "01_Teens" & FDL & sString1
                        Next
                        MkDir(sPfad & FDL & "LR Kataloge")
                        iOrdner = iOrdner + 1
                        sPfad = sPfad & FDL & "LR Kataloge"
                        For j = 1 To 10
                            If Not sNameTFoto(j - 1) = "" Then
                                If j > 9 Then
                                    MkDir(sPfad & FDL & j & "_" & sNameTFoto(j - 1))
                                    iOrdner = iOrdner + 1
                                Else
                                    MkDir(sPfad & FDL & "0" & j & "_" & sNameTFoto(j - 1))
                                    iOrdner = iOrdner + 1
                                End If
                            End If
                        Next
                        sPfad = sPfad & FDL & "01_Teens"
            'Video Ordner
                    Case bTVideo And Not Tdone(1)
                        'sPfad = sHauptPfad & FDL & "Sola_" & sSolaJahr & FDL & "01_Teens"
                        sString1 = sString1 & "_Video"

                        MkDir(sPfad & FDL & sString1)
                        iOrdner = iOrdner + 1
                        sPfad = sPfad & FDL & sString1


                        For Tag = 1 To 8
                            MkDir(sPfad & FDL & sTTag(Tag - 1) & "_Tag_" & Tag)
                            iOrdner = iOrdner + 1
                            sPfad = sPfad & FDL & sTTag(Tag - 1) & "_Tag_" & Tag
                            MkDir(sPfad & FDL & "01_Rohvideos")
                            iOrdner = iOrdner + 1
                            MkDir(sPfad & FDL & "02_Projektdatein")
                            iOrdner = iOrdner + 1
                            MkDir(sPfad & FDL & "03_Audio-Musik")
                            iOrdner = iOrdner + 1
                            sPfad = sPfad & FDL & "01_Rohvideos"
                            For j = 1 To 10
                                If Not sNameTVideo(j - 1) = "" Then
                                    If j > 9 Then
                                        MkDir(sPfad & FDL & j & "_" & sNameTVideo(j - 1))
                                        iOrdner = iOrdner + 1
                                    Else
                                        MkDir(sPfad & FDL & "0" & j & "_" & sNameTVideo(j - 1))
                                        iOrdner = iOrdner + 1
                                    End If
                                End If
                            Next
                            sPfad = Pfad & FDL & "Sola_" & sSolaJahr & FDL & "01_Teens" & FDL & sString1

                        Next
                        sPfad = sPfad & FDL & "01_Teens"
                        Tdone(1) = True
                'Showfiles
                    Case bTShowfiles And Not Tdone(2)
                        'sPfad = sHauptPfad & FDL & "Sola_" & sSolaJahr & FDL & "01_Teens"
                        sString1 = sString1 & "_Showfiles"
                        MkDir(sPfad & FDL & sString1)
                        iOrdner = iOrdner + 1
                        sPfad = sPfad & FDL & sString1
                        For Tag = 1 To 8
                            MkDir(sPfad & FDL & sTTag(Tag - 1) & "_Tag_" & Tag)
                            iOrdner = iOrdner + 1
                        Next
                        'sPfad = sPfad & FDL & "01_Teens"
                        Tdone(2) = True
                 'instagramm
                    Case bTInstagramm And Not Tdone(3)
                        sString1 = sString1 & "_Instagramm"
                        MkDir(sPfad & FDL & sString1)
                        iOrdner = iOrdner + 1
                        Tdone(3) = True
                 'Grafik
                    Case bTGrafik And Not Tdone(4)
                        sString1 = sString1 & "_Grafik"
                        MkDir(sPfad & FDL & sString1)
                        iOrdner = iOrdner + 1
                        Tdone(4) = True
                 'Audio
                    Case bTAudio And Not Tdone(5)
                        sString1 = sString1 & "_Audio"
                        MkDir(sPfad & FDL & sString1)
                        iOrdner = iOrdner + 1
                        sPfad = sPfad & FDL & sString1
                        For Tag = 1 To 8
                            MkDir(sPfad & FDL & sTTag(Tag - 1) & "_Tag_" & Tag)
                            iOrdner = iOrdner + 1
                        Next
                        'sPfad = sPfad & FDL & "01_Teens"
                        Tdone(5) = True
                 'Orga
                    Case bTOrga And Not Tdone(6)
                        sString1 = sString1 & "_Orga"
                        MkDir(sPfad & FDL & sString1)
                        iOrdner = iOrdner + 1
                        Tdone(6) = True
                 'Allgemein
                    Case bTAllgemein And Not Tdone(7)
                        sString1 = sString1 & "_Allgemein"
                        MkDir(sPfad & FDL & sString1)
                        iOrdner = iOrdner + 1
                        Tdone(7) = True



                End Select
                'bTFoto,                                bTVideo,                                    bTShowfiles,                                    bTInstagramm,                                       bTGrafik,                                   bTAudio,                                    bTOrga,                                 bTAllgemein
                If (bTFoto And Tdone(0) Or Not bTFoto) And (bTVideo And Tdone(1) Or Not bTVideo) And (bTShowfiles And Tdone(2) Or Not bTShowfiles) And (bTInstagramm And Tdone(3) Or Not bTInstagramm) And (bTGrafik And Tdone(4) Or Not bTGrafik) And (bTAudio And Tdone(5) Or Not bTAudio) And (bTOrga And Tdone(6) Or Not bTOrga) And (bTAllgemein And Tdone(7) Or Not bTAllgemein) Then
                    Exit For
                End If
            Next
        End If
        sPfad = Pfad & FDL & "Sola_" & sSolaJahr
        If bKids Then
            Chek_Tag_Array(sKTag, DateTimePickerKids, False)
            MkDir(sPfad & FDL & "02_Kids")
            iOrdner = iOrdner + 1
            sPfad = sPfad & FDL & "02_Kids"
            'Range("B2").Value = "02_Kids"
            iZeile = 0
            For i = 1 To 8
                sString1 = "0" & i
                sPfad = Pfad & FDL & "Sola_" & sSolaJahr & FDL & "02_Kids"
                Select Case True
            'Foto Ordner
                    Case bKFoto And Not Kdone(0)
                        sString1 = sString1 & "_Foto"
                        Kdone(0) = True
                        MkDir(sPfad & FDL & sString1)
                        iOrdner = iOrdner + 1
                        sPfad = sPfad & FDL & sString1

                        For Tag = 1 To 8
                            MkDir(sPfad & FDL & dKTag(Tag - 1) & "_Tag_" & Tag)
                            iOrdner = iOrdner + 1
                            sPfad = sPfad & FDL & dKTag(Tag - 1) & "_Tag_" & Tag
                            MkDir(sPfad & FDL & "01_Bilder des Tages" & Tag & "_HQ")
                            iOrdner = iOrdner + 1
                            MkDir(sPfad & FDL & "02_Bilder des Tages" & Tag & "_LQ")
                            iOrdner = iOrdner + 1
                            MkDir(sPfad & FDL & "03_Auswahl Bilderclip")
                            iOrdner = iOrdner + 1
                            MkDir(sPfad & FDL & "04_Auswahl Musik")
                            iOrdner = iOrdner + 1
                            For j = 1 To 10
                                If Not sNameKFoto(j - 1) = "" Then
                                    If j + 4 > 9 Then
                                        MkDir(sPfad & FDL & j + 4 & "_" & sNameKFoto(j - 1))
                                        iOrdner = iOrdner + 1
                                        sPfad = sPfad & FDL & j + 4 & "_" & sNameKFoto(j - 1)
                                    Else
                                        sPfad = sPfad & FDL & "0" & j + 4 & "_" & sNameKFoto(j - 1)
                                        MkDir(sPfad)
                                        iOrdner = iOrdner + 1
                                        'sPfad = sPfad & FDL & "0" & j + 4 & "_" & sNameTFoto(j - 1)
                                    End If

                                    For k = 0 To 3
                                        MkDir(sPfad & FDL & sPerOrdnerFoto(k))
                                        iOrdner = iOrdner + 1
                                    Next
                                    'sPfad = sPfad & FDL & sTTag(Tag - 1) & "_Tag" & Tag
                                    sPfad = Pfad & FDL & "Sola_" & sSolaJahr & FDL & "02_Kids" & FDL & sString1 & FDL & sKTag(Tag - 1) & "_Tag_" & Tag
                                End If
                            Next
                            sPfad = Pfad & FDL & "Sola_" & sSolaJahr & FDL & "02_Kids" & FDL & sString1
                        Next
                        MkDir(sPfad & FDL & "LR Kataloge")
                        iOrdner = iOrdner + 1
                        sPfad = sPfad & FDL & "LR Kataloge"
                        For j = 1 To 10
                            If Not sNameKFoto(j - 1) = "" Then
                                If j > 9 Then
                                    MkDir(sPfad & FDL & j & "_" & sNameKFoto(j - 1))
                                    iOrdner = iOrdner + 1
                                Else
                                    MkDir(sPfad & FDL & "0" & j & "_" & sNameKFoto(j - 1))
                                    iOrdner = iOrdner + 1
                                End If
                            End If
                        Next
                        sPfad = sPfad & FDL & "02_Kids"
            'Video Ordner
                    Case bKVideo And Not Kdone(1)
                        'sPfad = sHauptPfad & FDL & "Sola_" & sSolaJahr & FDL & "01_Teens"
                        sString1 = sString1 & "_Video"

                        MkDir(sPfad & FDL & sString1)
                        iOrdner = iOrdner + 1
                        sPfad = sPfad & FDL & sString1


                        For Tag = 1 To 8
                            MkDir(sPfad & FDL & sKTag(Tag - 1) & "_Tag_" & Tag)
                            iOrdner = iOrdner + 1
                            sPfad = sPfad & FDL & sKTag(Tag - 1) & "_Tag_" & Tag
                            MkDir(sPfad & FDL & "01_Rohvideos")
                            iOrdner = iOrdner + 1
                            MkDir(sPfad & FDL & "02_Projektdatein")
                            iOrdner = iOrdner + 1
                            MkDir(sPfad & FDL & "03_Audio-Musik")
                            iOrdner = iOrdner + 1
                            sPfad = sPfad & FDL & "01_Rohvideos"
                            For j = 1 To 10
                                If Not sNameKVideo(j - 1) = "" Then
                                    If j > 9 Then
                                        MkDir(sPfad & FDL & j & "_" & sNameKVideo(j - 1))
                                        iOrdner = iOrdner + 1
                                    Else
                                        MkDir(sPfad & FDL & "0" & j & "_" & sNameKVideo(j - 1))
                                        iOrdner = iOrdner + 1
                                    End If
                                End If
                            Next
                            sPfad = Pfad & FDL & "Sola_" & sSolaJahr & FDL & "02_Kids" & FDL & sString1

                        Next
                        sPfad = sPfad & FDL & "02_Kids"
                        Kdone(1) = True
                'Showfiles
                    Case bKShowfiles And Not Kdone(2)
                        'sPfad = sHauptPfad & FDL & "Sola_" & sSolaJahr & FDL & "01_Teens"
                        sString1 = sString1 & "_Showfiles"
                        MkDir(sPfad & FDL & sString1)
                        iOrdner = iOrdner + 1
                        sPfad = sPfad & FDL & sString1
                        For Tag = 1 To 8
                            MkDir(sPfad & FDL & sKTag(Tag - 1) & "_Tag_" & Tag)
                            iOrdner = iOrdner + 1
                        Next
                        'sPfad = sPfad & FDL & "01_Teens"
                        Kdone(2) = True
                 'instagramm
                    Case bKInstagramm And Not Kdone(3)
                        sString1 = sString1 & "_Instagramm"
                        MkDir(sPfad & FDL & sString1)
                        iOrdner = iOrdner + 1
                        Kdone(3) = True
                 'Grafik
                    Case bKGrafik And Not Kdone(4)
                        sString1 = sString1 & "_Grafik"
                        MkDir(sPfad & FDL & sString1)
                        iOrdner = iOrdner + 1
                        Kdone(4) = True
                 'Audio
                    Case bKAudio And Not Kdone(5)
                        sString1 = sString1 & "_Audio"
                        MkDir(sPfad & FDL & sString1)
                        iOrdner = iOrdner + 1
                        sPfad = sPfad & FDL & sString1
                        For Tag = 1 To 8
                            MkDir(sPfad & FDL & sKTag(Tag - 1) & "_Tag_" & Tag)
                            iOrdner = iOrdner + 1
                        Next
                        'sPfad = sPfad & FDL & "01_Teens"
                        Kdone(5) = True
                 'Orga
                    Case bKOrga And Not Kdone(6)
                        sString1 = sString1 & "_Orga"
                        MkDir(sPfad & FDL & sString1)
                        iOrdner = iOrdner + 1
                        Kdone(6) = True
                 'Allgemein
                    Case bKAllgemein And Not Kdone(7)
                        sString1 = sString1 & "_Allgemein"
                        MkDir(sPfad & FDL & sString1)
                        iOrdner = iOrdner + 1
                        Kdone(7) = True

                End Select
                'bTFoto,                                bTVideo,                                    bTShowfiles,                                    bTInstagramm,                                       bTGrafik,                                   bTAudio,                                    bTOrga,                                 bTAllgemein
                If (bKFoto And Kdone(0) Or Not bKFoto) And (bKVideo And Kdone(1) Or Not bKVideo) And (bKShowfiles And Kdone(2) Or Not bKShowfiles) And (bKInstagramm And Kdone(3) Or Not bKInstagramm) And (bKGrafik And Kdone(4) Or Not bKGrafik) And (bKAudio And Kdone(5) Or Not bKAudio) And (bKOrga And Kdone(6) Or Not bKOrga) And (bKAllgemein And Kdone(7) Or Not bKAllgemein) Then
                    Exit For
                End If
            Next
        End If
        If bTeens And Not bKids Then
            If (bTFoto And Tdone(0) Or Not bTFoto) And (bTVideo And Tdone(1) Or Not bTVideo) And (bTShowfiles And Tdone(2) Or Not bTShowfiles) And (bTInstagramm And Tdone(3) Or Not bTInstagramm) And (bTGrafik And Tdone(4) Or Not bTGrafik) And (bTAudio And Tdone(5) Or Not bTAudio) And (bTOrga And Tdone(6) Or Not bTOrga) And (bTAllgemein And Tdone(7) Or Not bTAllgemein) Then
                MsgBox("Ordnerstrucktur Wurde Erstellt" & vbNewLine & "Es wurden " & iOrdner & " erstellt", vbOK)
            End If
        ElseIf bKids And Not bTeens Then
            If (bKFoto And Kdone(0) Or Not bKFoto) And (bKVideo And Kdone(1) Or Not bKVideo) And (bKShowfiles And Kdone(2) Or Not bKShowfiles) And (bKInstagramm And Kdone(3) Or Not bKInstagramm) And (bKGrafik And Kdone(4) Or Not bKGrafik) And (bKAudio And Kdone(5) Or Not bKAudio) And (bKOrga And Kdone(6) Or Not bKOrga) And (bKAllgemein And Kdone(7) Or Not bKAllgemein) Then
                MsgBox("Ordnerstrucktur Wurde Erstellt" & vbNewLine & "Es wurden " & iOrdner & " erstellt", vbOK)
            End If
        ElseIf bTeens And bKids Then
            If ((bTFoto And Tdone(0) Or Not bTFoto) And (bTVideo And Tdone(1) Or Not bTVideo) And (bTShowfiles And Tdone(2) Or Not bTShowfiles) And (bTInstagramm And Tdone(3) Or Not bTInstagramm) And (bTGrafik And Tdone(4) Or Not bTGrafik) And (bTAudio And Tdone(5) Or Not bTAudio) And (bTOrga And Tdone(6) Or Not bTOrga) And (bTAllgemein And Tdone(7) Or Not bTAllgemein)) And ((bKFoto And Kdone(0) Or Not bKFoto) And (bKVideo And Kdone(1) Or Not bKVideo) And (bKShowfiles And Kdone(2) Or Not bKShowfiles) And (bKInstagramm And Kdone(3) Or Not bKInstagramm) And (bKGrafik And Kdone(4) Or Not bKGrafik) And (bKAudio And Kdone(5) Or Not bKAudio) And (bKOrga And Kdone(6) Or Not bKOrga) And (bKAllgemein And Kdone(7) Or Not bKAllgemein)) Then
                MsgBox("Ordnerstrucktur Wurde Erstellt" & vbNewLine & "Es wurden " & iOrdner & " Ordner erstellt", vbOK)
            End If
        End If


    End Sub
    Sub Chek_Tag_Array(Tag As Array, Controle As Object, Teen As Boolean)
        For i = 0 To Tag.Length - 1
            If IsNothing(Tag(i)) Then
                Berechne_Woche(Controle.Value, Teen)
            End If
        Next
    End Sub
    Sub Delete_Gaps_in_Name()
        Dim nameTF(9) As String
        Dim nameTV(9) As String
        Dim nameKF(9) As String
        Dim nameKV(9) As String

        Dim j = 0
        For i = 1 To 10
            nameTF(i - 1) = Me.GroupBoxTeens.Controls("TextBoxTFoto" & i).Text
        Next
        For i = 1 To 10
            nameTV(i - 1) = Me.GroupBoxTeens.Controls("TextBoxTVideo" & i).Text
        Next
        For i = 1 To 10
            nameKF(i - 1) = Me.GroupBoxKids.Controls("TextBoxKFoto" & i).Text
        Next
        For i = 1 To 10
            nameKV(i - 1) = Me.GroupBoxKids.Controls("TextBoxKVideo" & i).Text
        Next
        nameTF = Tools.DeletArrayGaps(nameTF)
        nameTV = Tools.DeletArrayGaps(nameTV)
        nameKF = Tools.DeletArrayGaps(nameKF)
        nameKV = Tools.DeletArrayGaps(nameKV)
        For i = 1 To 10
            Me.GroupBoxTeens.Controls("TextBoxTFoto" & i).Text = nameTF(i - 1)
        Next
        For i = 1 To 10
            Me.GroupBoxTeens.Controls("TextBoxTVideo" & i).Text = nameTV(i - 1)
        Next
        For i = 1 To 10
            Me.GroupBoxKids.Controls("TextBoxKFoto" & i).Text = nameKF(i - 1)
        Next
        For i = 1 To 10
            Me.GroupBoxKids.Controls("TextBoxKVideo" & i).Text = nameKV(i - 1)
        Next
    End Sub

End Class