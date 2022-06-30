
Imports System.Resources

'Made By Thorben Renfordt
Public Class Main_Form
    'Allgemeine Variablen (Daten)
    Const Version As String = "Alpha-v0.2.1"                                                    'Version
    Const sDateFormat As String = "dd-MM-yyyy"                                                  'Formatierung vom Datum
    Dim sSolaJahr As String                                                                     'Jahr in dem das Sola Stattfindet
    Dim sNameTFoto(9), sNameTVideo(9), sNameKFoto(9), sNameKVideo(9)                            'Namen der Mittarbeiter im Stringarray
    Dim bTFoto, bTVideo, bTShowfiles, bTInstagramm, bTGrafik, bTAudio, bTOrga, bTAllgemein      'Auswahlen für Ordner Teensola
    Dim bKFoto, bKVideo, bKShowfiles, bKInstagramm, bKGrafik, bKAudio, bKOrga, bKAllgemein      'Auswahlen für Ordner Kidsola
    Dim dTTag(7), dKTag(7), sTTag(7), sKTag(7)                                                  'Datum fon jedem Tag des Solas (als DT und String)
    Dim bTeens, bKids                                                                           'Auswahl der Solas (Teen/Kids)
    Dim bSolaJahrMan As Boolean                                                                 'Sola Jahr wurde Manuell ausgewählt
    Dim DL As Char = ";"                                                                        'Trennzeichen für die csv datei
    Dim Tools As New Tools                                                                      'Verweis auf weitere funktionen (Tools.vb)
    Dim Pfad                                                                                    'Ausgewählter Pfad


    'Initialisieren des Forms
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Version anzeigen
        Me.LinkLabel_Version.Text = Version
        'Max und Min für die Datumsauswahl festlegen
        Dim dMinDate As DateTime = DateAdd(DateInterval.DayOfYear, CDbl(DateDiff(DateInterval.Day, Now(), CDate(CStr(Year(Now()) & "/" & Month(Now()) - 1 & "/" & 1)))), Now())
        Dim dMaxDate As DateTime = DateAdd(DateInterval.Year, 5, Now())
        'Datumsauswahl Formatieren und Min/Max Festlegen
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
    'Speichern der Konfiguration (csv Datei) in einem ausgewälten Pfad
    Private Sub DataSave_Click(sender As Object, e As EventArgs) Handles DataSave.Click
        Delete_Gaps_in_Name()                   'Lere einträge der namen aufrücken
        Dim sf As SaveFileDialog = New SaveFileDialog()
        Dim sPfad, sContent, sKopf
        'Dialog Speichern unter konfigurieren
        With sf
            .Filter = "CSV|*.csv"                       'Dateiendung
            .RestoreDirectory = True                    'Start bei zulezt gewältem Ordner
            .CheckPathExists = True                     'Zeigt eine Warnung wenn der Pfad nicht existiert
            .Title = "Konfiguration Speichern unter:"
        End With
        sSolaJahr = Tools.SolaJahr(bTeens, bKids)
        If String.IsNullOrEmpty(sSolaJahr) Then         'Auswahl des Solas Cheken (Teens/Kids) und bei Keiner anwahl
            Exit Sub                                    'Sub Beenden
        End If
        'Dialog Speichern unter anzeigen und bei verlassen mit OK 
        'csv datei Schreiben
        If sf.ShowDialog() = DialogResult.OK Then
            sPfad = sf.FileName
            Dim Writer As New System.IO.StreamWriter(sf.FileName)
            'Kopfzeile der csv datei (Spaltenbeschreibung)
            sKopf = "SolaJahr" & DL & "Teens" & DL & "Kids" & DL & "NameTeenFotograf1" & DL & "NameTeenFotograf2" & DL & "NameTeenFotograf3" & DL & "NameTeenFotograf4" & DL & "NameTeenFotograf5" & DL & "NameTeenFotograf6" & DL & "NameTeenFotograf7" & DL & "NameTeenFotograf8" & DL & "NameTeenFotograf9" & DL & "NameTeenFotograf10"
            sKopf = sKopf & DL & "NameTeenVideograf1" & DL & "NameTeenVideograf2" & DL & "NameTeenVideograf3" & DL & "NameTeenVideograf4" & DL & "NameTeenVideograf5" & DL & "NameTeenVideograf6" & DL & "NameTeenVideograf7" & DL & "NameTeenVideograf8" & DL & "NameTeenVideograf9" & DL & "NameTeenVideograf10"
            sKopf = sKopf & DL & "NameKidsFotograf1" & DL & "NameKidsFotograf2" & DL & "NameKidsFotograf3" & DL & "NameKidsFotograf4" & DL & "NameKidsFotograf5" & DL & "NameKidsFotograf6" & DL & "NameKidsFotograf7" & DL & "NameKidsFotograf8" & DL & "NameKidsFotograf9" & DL & "NameKidsFotograf10"
            sKopf = sKopf & DL & "NameKidsVideograf1" & DL & "NameKidsVideograf2" & DL & "NameKidsVideograf3" & DL & "NameKidsVideograf4" & DL & "NameKidsVideograf5" & DL & "NameKidsVideograf6" & DL & "NameKidsVideograf7" & DL & "NameKidsVideograf8" & DL & "NameKidsVideograf9" & DL & "NameKidsVideograf10"
            sKopf = sKopf & DL & "AuswahlTeenFoto" & DL & "AuswahlTeenVideo" & DL & "AuswahlTeenShowfiles" & DL & "AuswahlTeenInstagramm" & DL & "AuswahlTeenGrafik" & DL & "AuswahlTeenAudio" & DL & "AuswahlTeenOrga" & DL & "AuswahlTeenAllgemein"
            sKopf = sKopf & DL & "AuswahlKidsFoto" & DL & "AuswahlKidsVideo" & DL & "AuswahlKidsShowfiles" & DL & "AuswahlKidsInstagramm" & DL & "AuswahlKidsGrafik" & DL & "AuswahlKidsAudio" & DL & "AuswahlKidsOrga" & DL & "AuswahlKidsAllgemein"
            sKopf = sKopf & DL & "TeenStartDatum" & DL & "KidsStartDatum"
            'Daten für die 2. Zeile Sammeln
            sContent = sSolaJahr & DL & bTeens & DL & bKids         'Sola Jahr, anwahl Teens, anwahl Kids
            For i = 0 To sNameTFoto.Length - 1
                sContent = sContent & DL & sNameTFoto(i)            'Die Namen der Fotografen vom Teensola
            Next
            For i = 0 To sNameTVideo.Length - 1
                sContent = sContent & DL & sNameTVideo(i)           'Die Namen der Videografen vom Teensola
            Next
            For i = 0 To sNameKFoto.Length - 1
                sContent = sContent & DL & sNameKFoto(i)            'Die Namen der Fotografen vom Kidssola
            Next
            For i = 0 To sNameKVideo.Length - 1
                sContent = sContent & DL & sNameKVideo(i)           'Die Namen der Videografen vom Kidssola
            Next
            '   anwahlen                Foto          Video         Showfiles           Instagramm          Grafik          Audio          Orga          Allgemein
            sContent = sContent & DL & bTFoto & DL & bTVideo & DL & bTShowfiles & DL & bTInstagramm & DL & bTGrafik & DL & bTAudio & DL & bTOrga & DL & bTAllgemein
            sContent = sContent & DL & bKFoto & DL & bKVideo & DL & bKShowfiles & DL & bKInstagramm & DL & bKGrafik & DL & bKAudio & DL & bKOrga & DL & bKAllgemein
            '                          Beginn Teensola       Beginn Kidssola
            sContent = sContent & DL & CStr(dTTag(0)) & DL & CStr(dKTag(0))

            ' If FileLineCount(File) = 0 Then
            Writer.WriteLine(sKopf)         'Datei erstellen und Schreiben der Kopfzeile

            ' End If
            Writer.WriteLine(sContent)      'Datei erstellen und Schreiben der Daten


            Writer.Close()                  'Datei Schlißen
        End If


    End Sub
    'Laden einer Gespeicherten Konfiguration
    Private Sub DataLoad_Click(sender As Object, e As EventArgs) Handles DataLoad.Click

        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim strFileName As String
        Dim iField As Integer = 0
        'Konfigurieren des Datei Öffnen Dialogs
        With fd
            .Title = "Konfiguration Wählen"
            '.InitialDirectory = Application.ExecutablePath
            .Filter = "CSV|*.csv"                       'Dateiendung
            .RestoreDirectory = True                    'Start bei zulezt gewältem Ordner
            .CheckFileExists = True                     'Zeigt eine warnung wenn die datei nicht existiert
        End With

        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.FileName
            Using Reader As New Microsoft.VisualBasic.FileIO.TextFieldParser(strFileName)       'Reader erstellen
                'Reader Konfigurieren
                With Reader
                    .TextFieldType = FileIO.FieldType.Delimited         'anwahl datei mit Trenzeichen getrenten daten
                    .SetDelimiters(DL)                                  'Trennzeichen Festlegen
                    .ReadLine()                                         'Kopfzeile Lesen
                End With
                Dim currentRow As String()
                Try
                    currentRow = Reader.ReadFields()                    'Daten Lesen und in Array laden
                    sSolaJahr = currentRow(0)                           'Sola Jahr auslesen
                    bTeens = CBool(currentRow(1))                       'anwahl Sola Teens
                    bKids = CBool(currentRow(2))                        'anwahl Sola Kids
                    For i = 0 To 9
                        sNameTFoto(i) = currentRow(3 + i)               'Namen der Fotografen vom Teensola
                    Next
                    For i = 0 To 9
                        sNameTVideo(i) = currentRow(13 + i)             'Namen der Videografen vom Teensola
                    Next
                    For i = 0 To 9
                        sNameKFoto(i) = currentRow(23 + i)              'Namen der Fotografen vom Kidssola
                    Next
                    For i = 0 To 9
                        sNameKVideo(i) = currentRow(33 + i)             'Namen der Videografen vom Kidssola
                    Next
                    bTFoto = CBool(currentRow(43))                      'anwahl Foto Teens
                    bTVideo = CBool(currentRow(44))                     'anwahl Video Teens
                    bTShowfiles = CBool(currentRow(45))                 'anwahl Showfiles Teens
                    bTInstagramm = CBool(currentRow(46))                'anwahl Instergramm Teens
                    bTGrafik = CBool(currentRow(47))                    'anwahl Grafik Teens
                    bTAudio = CBool(currentRow(48))                     'anwahl Audio Teens
                    bTOrga = CBool(currentRow(49))                      'anwahl Orga Teens
                    bTAllgemein = CBool(currentRow(50))                 'anwahl Allgemein Teens
                    bKFoto = CBool(currentRow(51))                      'anwahl Foto Kids
                    bKVideo = CBool(currentRow(52))                     'anwahl Video Kids
                    bKShowfiles = CBool(currentRow(53))                 'anwahl Showfiles Kids
                    bKInstagramm = CBool(currentRow(54))                'anwahl Instergramm Kids
                    bKGrafik = CBool(currentRow(55))                    'anwahl Grafik Kids
                    bKAudio = CBool(currentRow(56))                     'anwahl Audio Kids
                    bKOrga = CBool(currentRow(57))                      'anwahl Orga Kids
                    bKAllgemein = CBool(currentRow(58))                 'anwahl Allgemein Kids
                    dTTag(0) = CDate(currentRow(59))                    'Start Datum Teen Sola
                    dKTag(0) = CDate(currentRow(60))                    'Start Datum kids Sola

                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException         'Error handling
                    MsgBox("Line " & ex.Message & " Is not valid and will be Skipped.")
                End Try

            End Using
            NewDataShow()       'Daten anzeigen

        End If
    End Sub
    'Anzeige im Form aktualiesieren
    Sub NewDataShow()
        CheckBoxTeen.Checked = bTeens                                           'anwahl Teensola
        CheckBoxKids.Checked = bKids                                            'anwahl Kidssola
        DateTimePickerTeen.Value = dTTag(0)                                     'Startdatum Teen Sola
        DateTimePickerKids.Value = dKTag(0)                                     'Startdatum Kids Sola
        With CheckedListBoxTeen                                                 'Ordneranwahlen Teen
            .SetItemChecked(0, bTFoto)
            .SetItemChecked(1, bTVideo)
            .SetItemChecked(2, bTShowfiles)
            .SetItemChecked(3, bTInstagramm)
            .SetItemChecked(4, bTGrafik)
            .SetItemChecked(5, bTAudio)
            .SetItemChecked(6, bTOrga)
            .SetItemChecked(7, bTAllgemein)
        End With
        With CheckedListBoxKids                                                 'Ordneranwahlen Kids
            .SetItemChecked(0, bTFoto)
            .SetItemChecked(1, bTVideo)
            .SetItemChecked(2, bTShowfiles)
            .SetItemChecked(3, bTInstagramm)
            .SetItemChecked(4, bTGrafik)
            .SetItemChecked(5, bTAudio)
            .SetItemChecked(6, bTOrga)
            .SetItemChecked(7, bTAllgemein)
        End With

        For i = 1 To 10
            With Me.GroupBoxTeens                                               'Namen Teensola eintragen
                .Controls("TextBoxTFoto" & i).Text = sNameTFoto(i - 1)
                .Controls("TextBoxTVideo" & i).Text = sNameTVideo(i - 1)
            End With
            With Me.GroupBoxKids                                                'Namen Kidssola eintragen
                .Controls("TextBoxKFoto" & i).Text = sNameKFoto(i - 1)
                .Controls("TextBoxKVideo" & i).Text = sNameKVideo(i - 1)
            End With
        Next

    End Sub


    'Lightroom Preset's Installieren und Presets anpassen
    Private Sub Button1_Click_Install_LR_Preset(sender As Object, e As EventArgs) Handles Button1.Click
        Const sDestPfad_devlop_preset As String = "\Adobe\Lightroom\Develop Presets"        'Pfad für die Entwicklungspresets
        Const sExtention_devlop_preset As String = ".xmp"                                   'Dateiändung Entwicklungspresets
        Const sExtention_export_Preset As String = ".lrtemplate"                            'Dateiändung Exportpreset
        Const sDestPfad_export_Preset As String = "\Adobe\Lightroom\Export Presets\User Presets"    'Pfad für die Exportpresets
        Const maxRetry As Integer = 3                                                       'Anzahl versuche
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

        'Dialog zur abfrage vom Kürzel und Sola (Teen/Kids) anzeigen
        Select_Data = Tools.Select_Data_Preset()
        Select Case Select_Data.Valid   'Rückgabewerte cheken
            Case Select_Data_Valid.IsValid  'Rückgabe OK
                Exit Select
            Case Select_Data_Valid.NotValid 'Rückgabe NOK
                Select_Data = Tools.Select_Data_Preset()    'Dialog erneut anzeigen
                Select Case Select_Data.Valid   'Rückgabewerte cheken
                    Case Select_Data_Valid.IsValid  'Rückgabe OK
                        Exit Select
                    Case Select_Data_Valid.NotValid 'Rückgabe NOK
                        Select_Data = Tools.Select_Data_Preset()    'Dialog erneut anzeigen
                        Select Case Select_Data.Valid
                            Case Select_Data_Valid.IsValid 'Rückgabe OK
                                Exit Select
                            Case Select_Data_Valid.NotValid  'Rückgabe NOK
                                MsgBox("Ungültige eingaben", vbCritical) 'Auswahl abbrechen mit Meldung
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



        'Daten auf die Festplatte Schreiben (Export Presets)
        System.IO.File.WriteAllBytes(sPfadHQ, My.Resources.Resource1.HighQuality__HQ_)          'Preset HQ
        aOK_change(0) = Tools.Edit_Jear_in_Presets(sPfadHQ, Select_Data)                        'Preset anpassen
        System.IO.File.WriteAllBytes(sPfadLQ, My.Resources.Resource1.LowQuality__LQ_)           'Preset LQ
        aOK_change(1) = Tools.Edit_Jear_in_Presets(sPfadLQ, Select_Data)                        'Preset anpassen
        System.IO.File.WriteAllBytes(sPfadRAW, My.Resources.Resource1.RAW)                      'Preset RAW
        aOK_change(2) = Tools.Edit_Jear_in_Presets(sPfadRAW, Select_Data)                       'Preset anpassen

        'Überprüfen ob alle datein vorhanden sind (Preset HQ, Preset LQ, Preset RAW)
        If System.IO.File.Exists(sPfadHQ) And System.IO.File.Exists(sPfadLQ) And System.IO.File.Exists(sPfadRAW) Then
            bOK_create = True   'Datein Erfolgreich erstellt
            For i = 0 To 2
                aOK_create(i) = True
            Next
        Else        'auswertung welche Datei(en) Fehlen
            aOK_create(0) = System.IO.File.Exists(sPfadHQ)
            aOK_create(1) = System.IO.File.Exists(sPfadLQ)
            aOK_create(2) = System.IO.File.Exists(sPfadRAW)
        End If

        For i = 0 To 2      'auswertung ob die Presets erfolgreich geändert wurden
            If aOK_change(i) Then
                bOK_change = True
            Else
                bOK_change = False
                Exit For
            End If
        Next

        'Daten auf die Festplatte Schreiben (Entwicklungs Presets)
        System.IO.File.WriteAllBytes(sPfad_devlop_SOLA_Draußen, My.Resources.Resource1.draußen)                                     'Sola Draußen
        System.IO.File.WriteAllBytes(sPfad_devlop_SOLA_Veranstaltungszelt, My.Resources.Resource1.Veranstaltungszelt)               'Sola Veranstaltungszelt

        'Überprüfen ob alle datein vorhanden sind (Sola Draußen, Sola Veranstaltungszelt)
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
Diag:       'Diagnose welche Daten fehlen bzw. nicht angepasst sind

        If iRetry < maxRetry Then
            If Not bOK_create Then      'auswerten ob fehler beim Erstellen vorliegen
                For Each Bool In aOK_create
                    x = x + 1
                    If Bool Then sOk_create(x) = "OK" Else sOk_create(x) = "NOK"
                Next
                'Diagnose Meldung generieren mit Retry funktion
                If MsgBox("Fehler beim erstellen der vorgaben" & vbNewLine & "Exportvorgabe HQ: " & sOk_create(0) & vbNewLine _
                        & "Exportvorgabe LQ: " & sOk_create(1) & vbNewLine _
                        & "Exportvorgabe RAW: " & sOk_create(2) & vbNewLine _
                        & "Entwicklungsvorgabe Draußen: " & sOk_create(3) & vbNewLine _
                        & "Entwicklungsvorgabe Veranstaltungszelt: " & sOk_create(4), MsgBoxStyle.RetryCancel, "Fehler") = MsgBoxResult.Retry Then

                    doRetry = True
                End If
            End If

            If Not bOK_change Then      'auswerten ob fehler beim ändern Der Presets vorliegen
                For Each Bool As Boolean In aOK_change
                    y = y + 1
                    If Bool Then sOk_change(y) = "OK" Else sOk_change(x) = "NOK"
                Next
                'Diagnose Meldung generieren mit Retry funktion
                If MsgBox("Fehler beim ändern der vorgaben" & vbNewLine & "Exportvorgabe HQ: " & sOk_change(0) & vbNewLine _
                            & "Exportvorgabe LQ: " & sOk_change(1) & vbNewLine _
                            & "Exportvorgabe RAW: " & sOk_change(2) & vbNewLine, MsgBoxStyle.RetryCancel, "Fehler") = MsgBoxResult.Retry Then

                    doRetry = True
                End If
            End If
            If doRetry Then
                iRetry = iRetry + 1     'Retry Zähler
                GoTo Retry              'neuer Versuch
            End If
        Else
            'Finale Meldung wenn Keine neuversuche mehr Möglich sind
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


    'Auswahl des Datums (Bei Wertänderung)
    Private Sub DateTimePickerTeen_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePickerTeen.ValueChanged
        Tools.SolaJahrMan = False           'Manuele auswahl Rücksetzen
        Berechne_Woche(sender.Value, True)  'Berechnen der einzelnen Tage (Datumswerte)
    End Sub
    Private Sub DateTimePickerKids_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePickerKids.ValueChanged
        bSolaJahrMan = False                'Manuele auswahl Rücksetzen
        Berechne_Woche(sender.Value, False) 'Berechnen der einzelnen Tage (Datumswerte)
    End Sub

    'Berechnen des Datums für jeden Tag auf dem Sola 
    Sub Berechne_Woche(startDate As Date, Teen As Boolean) 'TODO Umbauen damit es als Funktion in die Tools Klasse past
        Dim dTag As Date
        If Teen Then                                    'Teensola
            dTTag(0) = Format(startDate, sDateFormat)   'Globales Format anwenden
            sTTag(0) = CStr(dTTag(0))                   'Startdatum in String wandeln
            For i = 1 To 7                              'Restliche Tage Berechnen
                dTag = DateAdd("d", i, startDate)       'Tag(e) addieren
                dTTag(i) = Format(dTag, sDateFormat)    'Globales Format anwenden
                sTTag(i) = CStr(dTTag(i))               'Datum in String wandeln
            Next
        Else                                            'Kidssola
            dKTag(0) = Format(startDate, sDateFormat)   'Globales Format anwenden
            sKTag(0) = CStr(dKTag(0))                   'Startdatum in String wandeln
            For i = 1 To 7                              'Restliche Tage Berechnen
                dTag = DateAdd("d", i, startDate)       'Tag(e) addieren
                dKTag(i) = Format(dTag, sDateFormat)    'Globales Format anwenden
                sKTag(i) = CStr(dKTag(i))               'Datum in String wandeln
            Next
        End If
    End Sub
    'Button Ordnerstrucktur Erstellen
    Private Sub ButtonStrucktur_Click(sender As Object, e As EventArgs) Handles ButtonStrucktur.Click
        Delete_Gaps_in_Name()                   'Lere einträge der namen aufrücken
        If Not String.IsNullOrEmpty(Pfad) Then  'Prüfen ob Pfad ausgewählt wurde
            Ordnerstrucktur_Erstellen()         'Strucktur Erstellen
        Else
            If MsgBox("Es wurde kein Ordner ausgewählt!" & vbCrLf & "Ordner jetzt wählen ?", vbOKCancel) = vbOK Then
                Ordnerauswahl_ini()             'Pfadauswahl
            End If
        End If

    End Sub
    'Beim Verlassen der Namensfelder mittels Regix Namen überprüfen
    Private Sub TextBoxTFoto1_TextChanged(sender As Object, e As EventArgs) Handles TextBoxTFoto1.Leave, TextBoxTFoto2.Leave, TextBoxTFoto3.Leave, TextBoxTFoto4.Leave, TextBoxTFoto5.Leave, TextBoxTFoto6.Leave, TextBoxTFoto7.Leave, TextBoxTFoto8.Leave, TextBoxTFoto9.Leave, TextBoxTFoto10.Leave, TextBoxTVideo1.Leave, TextBoxTVideo2.Leave, TextBoxTVideo3.Leave, TextBoxTVideo4.Leave, TextBoxTVideo5.Leave, TextBoxTVideo6.Leave, TextBoxTVideo7.Leave, TextBoxTVideo8.Leave, TextBoxTVideo9.Leave, TextBoxTVideo10.Leave, TextBoxKFoto1.Leave, TextBoxKFoto2.Leave, TextBoxKFoto3.Leave, TextBoxKFoto4.Leave, TextBoxKFoto5.Leave, TextBoxKFoto6.Leave, TextBoxKFoto7.Leave, TextBoxKFoto8.Leave, TextBoxKFoto9.Leave, TextBoxKFoto10.Leave, TextBoxKVideo1.Leave, TextBoxKVideo2.Leave, TextBoxKVideo3.Leave, TextBoxKVideo4.Leave, TextBoxKVideo5.Leave, TextBoxKVideo6.Leave, TextBoxKVideo7.Leave, TextBoxKVideo8.Leave, TextBoxKVideo9.Leave, TextBoxKVideo10.Leave
        Const sPattern As String = "^[a-zA-Z]+(([',. -][a-zA-Z ])?[a-zA-Z]*)*$" 'TODO: Aktuell sind keine Zahlen hinter dem Namen erlaubt
        If Not Tools.ValidateStr(sender.Text, sPattern).boolErgebnis Then   'String überprüfen 
            If Not String.IsNullOrEmpty(sender.Text) Then                   'Prüfen ob Feld Leer ist
                MsgBox("Ungültige Eingabe", MsgBoxStyle.OkOnly)             'Meldung ausgeben
            End If
            sender.Text = ""
            'Exit Sub
        End If
        Dim Sola, Bereich, sNr
        Dim iNr As Integer
        Sola = sender.Name(7)                           'auslesen zu welchem Sola das feld gehört (Teen/Kids)
        Bereich = Strings.Mid(sender.Name, 9, 1)        'auslesen in welchem bereich das Feld ist (Foto/Video)
        sNr = Strings.Right(sender.Name, 2)             'auslesen welche nummer das Feld hat
        If Not IsNumeric(sNr(0)) Then                   'Prüfen ob noch Buchstaben in der nummer sind
            sNr = sNr(1)                                'Buchstaben entfernen
            iNr = Char.GetNumericValue(sNr)             'String in Integer Wandeln
        Else
            iNr = CInt(sNr)                             'String in Integer wandeln
        End If
        'auswertung der Felder
        If Sola = "T" Then                              'auswerten des Teen Solas
            If Bereich = "F" Then                       'auswerten des Bereichs Foto
                sNameTFoto(iNr - 1) = sender.Text       'Name in variable "Namen Teen Foto" schreiben
            ElseIf Bereich = "V" Then                   'auswerten des Bereichs Video
                sNameTVideo(iNr - 1) = sender.Text      'Name in variable "Namen Teen Video" schreiben
            End If

        ElseIf Sola = "K" Then                          'auswerten des Kids Solas
            If Bereich = "F" Then                       'auswerten des Bereichs Foto
                sNameKFoto(iNr - 1) = sender.Text       'Name in variable "Namen Kids Foto" schreiben
            ElseIf Bereich = "V" Then                   'auswerten des Bereichs Video
                sNameKVideo(iNr - 1) = sender.Text      'Name in variable "Namen Kids Video" schreiben
            End If
        End If
        Freigabe_Erstellen()                            'Freigaben der Bearbeitbar der einzelnen anwahlen Felder und Buttons
    End Sub
    'Auswahl des Ziels für die Ordnerstrucktur (Startordner)
    Sub Ordnerauswahl_ini()
        Pfad = Ordnerwahl("Bitte Verzeichnis zum Erstellen der Ordnerstrucktur Wählen")     'Ordner auswahl Starten
        LabelPfad.Text = Pfad                                                               'Auswahl anzeigen
        Freigabe_Erstellen()                                                                'Freigaben der Bearbeitbar der einzelnen anwahlen Felder und Buttons
    End Sub
    'Button Orfnerauswahl
    Private Sub ButtonPicPfad_Click(sender As Object, e As EventArgs) Handles ButtonPicPfad.Click
        Delete_Gaps_in_Name()                   'Lere einträge der namen aufrücken
        Ordnerauswahl_ini()     'Auswahl des Ziels für die Ordnerstrucktur (Startordner)
    End Sub
    'Öffnet Verlinkung für die Versionsnummer
    Private Sub LinkLabel_Version_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel_Version.LinkClicked
        System.Diagnostics.Process.Start("https://github.com/TH0RB3Nger/SOLA_Ordnerstrucktur/releases")     'link zu Github Releases öffnen
        LinkLabel_Version.LinkVisited = True                                                                'zeigt das Link geöffnet wurde
    End Sub
    'Öffnet Verlinkung für Sola Wiedenest
    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        System.Diagnostics.Process.Start("https://www.instagram.com/sola_wiedenest/")       'Link Instagramm Sola Wiedenest
        LinkLabel2.LinkVisited = True                                                       'zeigt das Link geöffnet wurde
    End Sub
    'Öffnet Verlinkung zu Torsten 
    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start("https://www.instagram.com/der.torsten/")       'Link Instagramm Torsten
        LinkLabel1.LinkVisited = True                                                    'zeigt das Link geöffnet wurde
    End Sub
    'anwahlen Teensola werden Freigegeben
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
        EnableNamenTeens()                  'Namensfelder Teensola werden Freigegeben

    End Sub
    'anwahlen Kidssola werden Freigegeben
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
        EnableNamenKids()                   'Namensfelder Kidssola werden Freigegeben

    End Sub
    'Auswertung der anwahlen Teensola
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
        EnableNamenTeens()                                          'Namensfelder Teensola werden Freigegeben
    End Sub
    'Auswertung der anwahlen Kidssola
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
        EnableNamenKids()                                               'Namensfelder Kidssola werden Freigegeben
    End Sub
    'Freigabe der Namensfelder Teensola
    Private Sub EnableNamenTeens()
        'Namensfelder aktivieren wenn entsprechende auswahl gegeben ist
        With Me.GroupBoxTeens
            If bTeens Then      'Sola Teens angewählt
                For i = 1 To 10
                    If bTFoto Then      'Foto angewählt
                        .Controls("TextBoxTFoto" & i).Enabled = True
                    Else
                        .Controls("TextBoxTFoto" & i).Enabled = False
                    End If
                    If bTVideo Then     'Video angewählt
                        .Controls("TextBoxTVideo" & i).Enabled = True
                    Else
                        .Controls("TextBoxTVideo" & i).Enabled = False
                    End If
                Next
            Else
                For i = 1 To 10
                    .Controls("TextBoxTFoto" & i).Enabled = False
                    .Controls("TextBoxTVideo" & i).Enabled = False
                Next
            End If
        End With
        Freigabe_Erstellen()        'Freigaben der Bearbeitbar der einzelnen anwahlen Felder und Buttons
    End Sub
    Private Sub EnableNamenKids()
        'Namensfelder aktivieren wenn entsprechende auswahl gegeben ist
        With Me.GroupBoxKids
            If bKids Then   'anwahl Kidssola
                For i = 1 To 10
                    If bKFoto Then  'anwahl Foto
                        .Controls("TextBoxKFoto" & i).Enabled = True
                    Else
                        .Controls("TextBoxKFoto" & i).Enabled = False
                    End If
                    If bKVideo Then 'anwahl Video
                        .Controls("TextBoxKVideo" & i).Enabled = True
                    Else
                        .Controls("TextBoxKVideo" & i).Enabled = False
                    End If
                Next
            Else
                For i = 1 To 10
                    .Controls("TextBoxKFoto" & i).Enabled = False
                    .Controls("TextBoxKVideo" & i).Enabled = False
                Next
            End If
        End With
        Freigabe_Erstellen()        'Freigaben der Bearbeitbar der einzelnen anwahlen Felder und Buttons
    End Sub
    'Ordnerauswahl für Ordnerstrucktur
    Function Ordnerwahl(msg As String) As String
        ''Auswahldialog für die Ordnerwahl
        With OokiiDialog
            'Ordnerauswahl dialog Konfigurieren
            .Multiselect = False                'Merfachauswahl
            .ShowNewFolderButton = True         'Neuen Ordner erstallen möglich
            .Description = msg                  'Titel
            .UseDescriptionForTitle = True      'Beschribung aktivieren

            .ShowDialog()                       'auswahl anzeigen
            If String.IsNullOrEmpty(.SelectedPath) Then                     'Prüfen das etwas ausgewählt wurde
                MsgBox("Es wurde kein Ordner ausgewählt!", vbExclamation)   'Meldung bei keiner auswahl
                Return vbEmpty      'Leeren String Zurückgeben
                Exit Function
            End If
            Return .SelectedPath    'ausgewählten Pfad Zurückgeben
        End With
    End Function
    'Freigabe für Buttons 
    Private Sub Freigabe_Erstellen()
        If Not (Pfad = "" Or String.IsNullOrEmpty(Pfad)) Then       'ausgewälten Pfad Prüfen
            If bTeens And Not bKids Then                            'anwahl Teensola ohne Kidssola
                'anwahlen und Namensfelder Teensola Prüfen
                If (bTShowfiles Or bTInstagramm Or bTGrafik Or bTAudio Or bTOrga Or bTAllgemein) Or (bTFoto And Tools.NameNotEmpty("T", "F")) Or (bTVideo And Tools.NameNotEmpty("T", "V")) Then
                    ButtonStrucktur.Enabled = True                  'Button Ordnerstrucktur erstellen Freigeben
                Else
                    ButtonStrucktur.Enabled = False                 'Button Ordnerstrucktur erstellen Sperren
                End If
            ElseIf bKids And Not bTeens Then                        'anwahl Kidssola ohne Teensola
                'anwahlen und Namensfelder Kidsola Prüfen
                If (bKShowfiles Or bKInstagramm Or bKGrafik Or bKAudio Or bKOrga Or bKAllgemein) Or (bKFoto And Tools.NameNotEmpty("K", "F")) Or (bKVideo And Tools.NameNotEmpty("K", "V")) Then
                    ButtonStrucktur.Enabled = True                  'Button Ordnerstrucktur erstellen Freigeben
                Else
                    ButtonStrucktur.Enabled = False                 'Button Ordnerstrucktur erstellen Sperren
                End If
            ElseIf bTeens And bKids Then                            'anwahl Kidssola und Teensola
                'anwahlen und Namensfelder Prüfen
                If (bTShowfiles Or bTInstagramm Or bTGrafik Or bTAudio Or bTOrga Or bTAllgemein) Or (bTFoto And Tools.NameNotEmpty("T", "F")) Or (bTVideo And Tools.NameNotEmpty("T", "V")) _
                    And (bKShowfiles Or bKInstagramm Or bKGrafik Or bKAudio Or bKOrga Or bKAllgemein) Or (bKFoto And Tools.NameNotEmpty("K", "F")) Or (bKVideo And Tools.NameNotEmpty("K", "V")) Then
                    ButtonStrucktur.Enabled = True                  'Button Ordnerstrucktur erstellen Freigeben
                Else
                    ButtonStrucktur.Enabled = False                 'Button Ordnerstrucktur erstellen Sperren
                End If
            Else
                ButtonStrucktur.Enabled = False                 'Button Ordnerstrucktur erstellen Sperren
            End If
        Else
            ButtonStrucktur.Enabled = False                 'Button Ordnerstrucktur erstellen Sperren
        End If
    End Sub

    'Ertellt die Konfigurierte Ordnerstrucktur
    Sub Ordnerstrucktur_Erstellen()

        'Allgemeine Variablen deklaration:
        Dim i
        Dim oFSO As Object
        Dim sPfad As String
        Dim iOrdner As Integer          'Zählt Erstellte Ordner

        'Variablen Inizialisation
        oFSO = CreateObject("Scripting.FileSystemObject")
        i = 0


        Dim j
        Dim Tag
        Dim iZeile
        Dim Tdone(7), Kdone(7)
        Dim sString1
        Dim sPerOrdnerFoto() As String = {"01_ImportRAW", "02_ExportJPEG_HQ", "03_ExportJPEG_LQ", "04_ExportRAW"}       'Unterordner Für jeden Fotografen
        Dim FDL As String = "\"
        sSolaJahr = Tools.SolaJahr(bTeens, bKids)

        MkDir(Pfad & FDL & "Sola_" & sSolaJahr)                         'Stammordner Erstellen (Sola_XX)
        sPfad = Pfad & FDL & "Sola_" & sSolaJahr
        iOrdner = 1
        If bTeens Then                                                  'anwahl Teensola
            Chek_Tag_Array(sTTag, DateTimePickerTeen, True)             'Prüft das Array auf Volständigkeit ansonsten wird neu Berechnet
            MkDir(sPfad & FDL & "01_Teens")                             'Ordner Teens erstellen
            sPfad = sPfad & FDL & "01_Teens"
            iOrdner = iOrdner + 1
            'Range("B2").Value = "02_Kids"
            iZeile = 0
            For i = 1 To 8                                                  'Index
                sString1 = "0" & i                                          'Führende null hizufügen
                sPfad = Pfad & FDL & "Sola_" & sSolaJahr & FDL & "01_Teens"
                Select Case True
            'Foto Ordner angewählt und nicht bereits erledigt
                    Case bTFoto And Not Tdone(0)
                        sString1 = sString1 & "_Foto"
                        MkDir(sPfad & FDL & sString1)                       'Fotoordner erstellt
                        Tdone(0) = True
                        iOrdner = iOrdner + 1
                        sPfad = sPfad & FDL & sString1

                        For Tag = 1 To 8                                                'Tag 1 bis 8
                            MkDir(sPfad & FDL & sTTag(Tag - 1) & "_Tag_" & Tag)         'Erstellt den Tagesordner (dd-mm-yyyy_Tag_X)
                            iOrdner = iOrdner + 1
                            sPfad = sPfad & FDL & sTTag(Tag - 1) & "_Tag_" & Tag
                            MkDir(sPfad & FDL & "01_Bilder_des_Tages_" & Tag & "_HQ")    'Erstellt den Ordner "01_Bilder_des_Tages_X_HQ"
                            iOrdner = iOrdner + 1
                            MkDir(sPfad & FDL & "02_Bilder_des_Tages_" & Tag & "_LQ")   'Erstellt den Ordner "02_Bilder_des_Tages_X_LQ"
                            iOrdner = iOrdner + 1
                            MkDir(sPfad & FDL & "03_Auswahl Bilderclip")                'Erstellt den Ordner "03_Auswahl Bilderclip"
                            iOrdner = iOrdner + 1
                            MkDir(sPfad & FDL & "04_Auswahl Musik")                     'Erstellt den Ordner "04_Auswahl Musik"
                            iOrdner = iOrdner + 1
                            For j = 1 To 10                                             'Für Jeden Fotografen
                                If Not sNameTFoto(j - 1) = "" Then                      'Prüfen das Fotograf vorhanden ist
                                    If j + 4 > 9 Then                                   'Bei 5 bis 9 mit führender null sonst Ohne
                                        MkDir(sPfad & FDL & j + 4 & "_" & sNameTFoto(j - 1))    'Ordner für Fotografen erstellen "NR_Name"
                                        iOrdner = iOrdner + 1
                                        sPfad = sPfad & FDL & j + 4 & "_" & sNameTFoto(j - 1)
                                    Else
                                        sPfad = sPfad & FDL & "0" & j + 4 & "_" & sNameTFoto(j - 1)
                                        MkDir(sPfad)                                             'Ordner für Fotografen erstellen "NR_Name"
                                        iOrdner = iOrdner + 1
                                        'sPfad = sPfad & FDL & "0" & j + 4 & "_" & sNameTFoto(j - 1)
                                    End If

                                    For k = 0 To sPerOrdnerFoto.Length - 1                       'Erstellen der Persönlichen festgelegten Ordner (01_ImportRAW, 02_ExportJPEG_HQ, ...)
                                        MkDir(sPfad & FDL & sPerOrdnerFoto(k))
                                        iOrdner = iOrdner + 1
                                    Next
                                    'sPfad = sPfad & FDL & sTTag(Tag - 1) & "_Tag" & Tag
                                    sPfad = Pfad & FDL & "Sola_" & sSolaJahr & FDL & "01_Teens" & FDL & sString1 & FDL & sTTag(Tag - 1) & "_Tag_" & Tag
                                End If
                            Next
                            sPfad = Pfad & FDL & "Sola_" & sSolaJahr & FDL & "01_Teens" & FDL & sString1
                        Next
                        MkDir(sPfad & FDL & "LR Kataloge")              'Ordner für Lightroomkataloge erstellen
                        iOrdner = iOrdner + 1
                        sPfad = sPfad & FDL & "LR Kataloge"
                        For j = 1 To 10                                 'Unterordner für Jeden Fotografen erstellen
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
            'Video Ordner angewählt und nicht bereits erledigt
                    Case bTVideo And Not Tdone(1)
                        'sPfad = sHauptPfad & FDL & "Sola_" & sSolaJahr & FDL & "01_Teens"
                        sString1 = sString1 & "_Video"

                        MkDir(sPfad & FDL & sString1)                       'Video Ordner Erstellen
                        iOrdner = iOrdner + 1
                        sPfad = sPfad & FDL & sString1


                        For Tag = 1 To 8                                    'Tages Ordner Erstellen Tag 1 bis 8
                            MkDir(sPfad & FDL & sTTag(Tag - 1) & "_Tag_" & Tag)
                            iOrdner = iOrdner + 1
                            sPfad = sPfad & FDL & sTTag(Tag - 1) & "_Tag_" & Tag
                            MkDir(sPfad & FDL & "01_Rohvideos")             'Ordner "01_Rohvideos" erstellen
                            iOrdner = iOrdner + 1
                            MkDir(sPfad & FDL & "02_Projektdatein")         'Ordner "02_Projektdatein" erstellen
                            iOrdner = iOrdner + 1
                            MkDir(sPfad & FDL & "03_Audio-Musik")           'Ordner "03_Audio-Musik" erstellen
                            iOrdner = iOrdner + 1
                            sPfad = sPfad & FDL & "01_Rohvideos"
                            For j = 1 To 10                                 'Ordner für jeden Vidiografen erstellen
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
                'Showfiles Ordner ausgewählt und nicht erledigt
                    Case bTShowfiles And Not Tdone(2)
                        'sPfad = sHauptPfad & FDL & "Sola_" & sSolaJahr & FDL & "01_Teens"
                        sString1 = sString1 & "_Showfiles"
                        MkDir(sPfad & FDL & sString1)                       'Ordner Showfiles erstellen
                        iOrdner = iOrdner + 1
                        sPfad = sPfad & FDL & sString1
                        For Tag = 1 To 8                                    'Tagesordner 1 bis 8 erstellen
                            MkDir(sPfad & FDL & sTTag(Tag - 1) & "_Tag_" & Tag)
                            iOrdner = iOrdner + 1
                        Next
                        'sPfad = sPfad & FDL & "01_Teens"
                        Tdone(2) = True
                 'Instagramm Ordner ausgewählt und nicht erledigt
                    Case bTInstagramm And Not Tdone(3)
                        sString1 = sString1 & "_Instagramm"
                        MkDir(sPfad & FDL & sString1)                       'Ordner Instergramm erstellen
                        iOrdner = iOrdner + 1
                        Tdone(3) = True
                 'Grafik Ordner ausgewählt und nicht erledigt  
                    Case bTGrafik And Not Tdone(4)
                        sString1 = sString1 & "_Grafik"
                        MkDir(sPfad & FDL & sString1)                       'Ordner Grafik erstellen
                        iOrdner = iOrdner + 1
                        Tdone(4) = True
                 'Audio Ordner ausgewählt und nicht erledigt
                    Case bTAudio And Not Tdone(5)
                        sString1 = sString1 & "_Audio"
                        MkDir(sPfad & FDL & sString1)                       'Ordner Audio erstellen
                        iOrdner = iOrdner + 1
                        sPfad = sPfad & FDL & sString1
                        For Tag = 1 To 8                                    'Tagesordner 1 bis 8 erstellen
                            MkDir(sPfad & FDL & sTTag(Tag - 1) & "_Tag_" & Tag)
                            iOrdner = iOrdner + 1
                        Next
                        'sPfad = sPfad & FDL & "01_Teens"
                        Tdone(5) = True
                 'Orga Ordner ausgewählt und nicht erledigt
                    Case bTOrga And Not Tdone(6)
                        sString1 = sString1 & "_Orga"
                        MkDir(sPfad & FDL & sString1)                       'Ordner Orga erstellen
                        iOrdner = iOrdner + 1
                        Tdone(6) = True
                 'Allgemein ausgewählt und nicht erledigt
                    Case bTAllgemein And Not Tdone(7)
                        sString1 = sString1 & "_Allgemein"
                        MkDir(sPfad & FDL & sString1)                       'Ordner Allgemein erstellen
                        iOrdner = iOrdner + 1
                        Tdone(7) = True



                End Select
                'Pfüfen ob alles was angewält ist auch als erledigt makiert ist
                'bTFoto,                                bTVideo,                                    bTShowfiles,                                    bTInstagramm,                                       bTGrafik,                                   bTAudio,                                    bTOrga,                                 bTAllgemein
                If (bTFoto And Tdone(0) Or Not bTFoto) And (bTVideo And Tdone(1) Or Not bTVideo) And (bTShowfiles And Tdone(2) Or Not bTShowfiles) And (bTInstagramm And Tdone(3) Or Not bTInstagramm) And (bTGrafik And Tdone(4) Or Not bTGrafik) And (bTAudio And Tdone(5) Or Not bTAudio) And (bTOrga And Tdone(6) Or Not bTOrga) And (bTAllgemein And Tdone(7) Or Not bTAllgemein) Then
                    Exit For
                End If
            Next
        End If
        sPfad = Pfad & FDL & "Sola_" & sSolaJahr
        If bKids Then       'anwahl Kidssola
            Chek_Tag_Array(sKTag, DateTimePickerKids, False)                'Prüft das Array auf Volständigkeit ansonsten wird neu Berechnet
            MkDir(sPfad & FDL & "02_Kids")                                  'Ordner 02_Kids erstellt 
            iOrdner = iOrdner + 1
            sPfad = sPfad & FDL & "02_Kids"
            'Range("B2").Value = "02_Kids"
            iZeile = 0                                                      'TODO ?etfernen?
            For i = 1 To 8                                                  'Index
                sString1 = "0" & i
                sPfad = Pfad & FDL & "Sola_" & sSolaJahr & FDL & "02_Kids"
                Select Case True
            'Foto Ordner angewählt und nicht erledigt
                    Case bKFoto And Not Kdone(0)
                        sString1 = sString1 & "_Foto"
                        Kdone(0) = True
                        MkDir(sPfad & FDL & sString1)                       'Ordner Foto erstellen
                        iOrdner = iOrdner + 1
                        sPfad = sPfad & FDL & sString1

                        For Tag = 1 To 8                                    'Tag 1 bis 8
                            MkDir(sPfad & FDL & dKTag(Tag - 1) & "_Tag_" & Tag)                       'Erstellt den Tagesordner (dd-mm-yyyy_Tag_X)
                            iOrdner = iOrdner + 1
                            sPfad = sPfad & FDL & dKTag(Tag - 1) & "_Tag_" & Tag
                            MkDir(sPfad & FDL & "01_Bilder_des_Tages_" & Tag & "_HQ")                       'Ordner "01_Bilder_des_Tages_XX_HQ" erstellen
                            iOrdner = iOrdner + 1
                            MkDir(sPfad & FDL & "02_Bilder_des_Tages_" & Tag & "_LQ")                       'Ordner "02_Bilder_des_Tages_XX_LQ" erstellen
                            iOrdner = iOrdner + 1
                            MkDir(sPfad & FDL & "03_Auswahl_Bilderclip")                       'Ordner 03_Auswahl_Bilderclip erstellen
                            iOrdner = iOrdner + 1
                            MkDir(sPfad & FDL & "04_Auswahl_Musik")                       'Ordner 04_Auswahl_Musik erstellen
                            iOrdner = iOrdner + 1
                            For j = 1 To 10                                 'Ordner für jeden Fotografen erstellen
                                If Not sNameKFoto(j - 1) = "" Then
                                    If j + 4 > 9 Then
                                        MkDir(sPfad & FDL & j + 4 & "_" & sNameKFoto(j - 1))
                                        iOrdner = iOrdner + 1
                                        sPfad = sPfad & FDL & j + 4 & "_" & sNameKFoto(j - 1)
                                    Else
                                        sPfad = sPfad & FDL & "0" & j + 4 & "_" & sNameKFoto(j - 1)
                                        MkDir(sPfad)
                                        iOrdner = iOrdner + 1
                                    End If

                                    For k = 0 To 3
                                        MkDir(sPfad & FDL & sPerOrdnerFoto(k))      'Erstellen der Persönlichen festgelegten Ordner (01_ImportRAW, 02_ExportJPEG_HQ, ...)
                                        iOrdner = iOrdner + 1
                                    Next
                                    'sPfad = sPfad & FDL & sTTag(Tag - 1) & "_Tag" & Tag
                                    sPfad = Pfad & FDL & "Sola_" & sSolaJahr & FDL & "02_Kids" & FDL & sString1 & FDL & sKTag(Tag - 1) & "_Tag_" & Tag
                                End If
                            Next
                            sPfad = Pfad & FDL & "Sola_" & sSolaJahr & FDL & "02_Kids" & FDL & sString1
                        Next
                        MkDir(sPfad & FDL & "LR Kataloge")                       'Ordner LR Kataloge erstellen
                        iOrdner = iOrdner + 1
                        sPfad = sPfad & FDL & "LR Kataloge"
                        For j = 1 To 10
                            If Not sNameKFoto(j - 1) = "" Then
                                If j > 9 Then                       'Ordner für jeden Fotografen erstellen
                                    MkDir(sPfad & FDL & j & "_" & sNameKFoto(j - 1))
                                    iOrdner = iOrdner + 1
                                Else
                                    MkDir(sPfad & FDL & "0" & j & "_" & sNameKFoto(j - 1))
                                    iOrdner = iOrdner + 1
                                End If
                            End If
                        Next
                        sPfad = sPfad & FDL & "02_Kids"
            'Video Ordner angewählt und nicht erledigt
                    Case bKVideo And Not Kdone(1)
                        sString1 = sString1 & "_Video"

                        MkDir(sPfad & FDL & sString1)                       'Ordner Video erstellen
                        iOrdner = iOrdner + 1
                        sPfad = sPfad & FDL & sString1


                        For Tag = 1 To 8                                    'Tagesordner 1 bis 8 
                            MkDir(sPfad & FDL & sKTag(Tag - 1) & "_Tag_" & Tag)                       'Ordner Tagesordner erstellen
                            iOrdner = iOrdner + 1
                            sPfad = sPfad & FDL & sKTag(Tag - 1) & "_Tag_" & Tag
                            MkDir(sPfad & FDL & "01_Rohvideos")                       'Ordner 01_Rohvideos erstellen
                            iOrdner = iOrdner + 1
                            MkDir(sPfad & FDL & "02_Projektdatein")                       'Ordner 02_Projektdatein erstellen
                            iOrdner = iOrdner + 1
                            MkDir(sPfad & FDL & "03_Audio-Musik")                       'Ordner 03_Audio-Musik erstellen
                            iOrdner = iOrdner + 1
                            sPfad = sPfad & FDL & "01_Rohvideos"                       'Ordner 01_Rohvideos erstellen
                            For j = 1 To 10                                             'Ordner für jeden Videografen erstellen
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
                'Showfiles Ordner angewählt und nicht erledigt
                    Case bKShowfiles And Not Kdone(2)
                        'sPfad = sHauptPfad & FDL & "Sola_" & sSolaJahr & FDL & "01_Teens"
                        sString1 = sString1 & "_Showfiles"
                        MkDir(sPfad & FDL & sString1)                       'Ordner Showfiles erstellen
                        iOrdner = iOrdner + 1
                        sPfad = sPfad & FDL & sString1
                        For Tag = 1 To 8                       'Tagesordner 1 bis 8 erstellen
                            MkDir(sPfad & FDL & sKTag(Tag - 1) & "_Tag_" & Tag)
                            iOrdner = iOrdner + 1
                        Next
                        'sPfad = sPfad & FDL & "01_Teens"
                        Kdone(2) = True
                 'Instagramm Ordner angewählt und nicht erledigt
                    Case bKInstagramm And Not Kdone(3)
                        sString1 = sString1 & "_Instagramm"
                        MkDir(sPfad & FDL & sString1)                       'Ordner Instagramm erstellen
                        iOrdner = iOrdner + 1
                        Kdone(3) = True
                 'Grafik Ordner angewählt und nicht erledigt
                    Case bKGrafik And Not Kdone(4)
                        sString1 = sString1 & "_Grafik"
                        MkDir(sPfad & FDL & sString1)                       'Ordner Grafik erstellen
                        iOrdner = iOrdner + 1
                        Kdone(4) = True
                 'Audio Ordner angewählt und nicht erledigt
                    Case bKAudio And Not Kdone(5)
                        sString1 = sString1 & "_Audio"
                        MkDir(sPfad & FDL & sString1)                       'Ordner Audio erstellen
                        iOrdner = iOrdner + 1
                        sPfad = sPfad & FDL & sString1
                        For Tag = 1 To 8                       'Tagesordner 1 bis 8 erstellen
                            MkDir(sPfad & FDL & sKTag(Tag - 1) & "_Tag_" & Tag)
                            iOrdner = iOrdner + 1
                        Next
                        'sPfad = sPfad & FDL & "01_Teens"
                        Kdone(5) = True
                 'Orga Ordner angewählt und nicht erledigt
                    Case bKOrga And Not Kdone(6)
                        sString1 = sString1 & "_Orga"
                        MkDir(sPfad & FDL & sString1)                       'Ordner Orga erstellen
                        iOrdner = iOrdner + 1
                        Kdone(6) = True
                 'Allgemein Ordner angewählt und nicht erledigt
                    Case bKAllgemein And Not Kdone(7)
                        sString1 = sString1 & "_Allgemein"
                        MkDir(sPfad & FDL & sString1)                       'Ordner Allgemein erstellen
                        iOrdner = iOrdner + 1
                        Kdone(7) = True

                End Select
                'Pfüfen ob alles was angewält ist auch als erledigt makiert ist
                'bTFoto,                                bTVideo,                                    bTShowfiles,                                    bTInstagramm,                                       bTGrafik,                                   bTAudio,                                    bTOrga,                                 bTAllgemein
                If (bKFoto And Kdone(0) Or Not bKFoto) And (bKVideo And Kdone(1) Or Not bKVideo) And (bKShowfiles And Kdone(2) Or Not bKShowfiles) And (bKInstagramm And Kdone(3) Or Not bKInstagramm) And (bKGrafik And Kdone(4) Or Not bKGrafik) And (bKAudio And Kdone(5) Or Not bKAudio) And (bKOrga And Kdone(6) Or Not bKOrga) And (bKAllgemein And Kdone(7) Or Not bKAllgemein) Then
                    Exit For
                End If
            Next
        End If
        'Finale Prüfung ob alle aufgaben erledigt wurden je nach anwahl
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
    'Prüft das Array der Datumswerte auf Volständigkeit ansonsten wird neu Berechnet
    Sub Chek_Tag_Array(Tag As Array, Controle As Object, Teen As Boolean)
        For i = 1 To Tag.Length - 1
            If IsNothing(Tag(i)) Then
                Berechne_Woche(Controle.Value, Teen)        'Berechnet die Tage
                Exit Sub
            End If
        Next
    End Sub
    'Lere einträge der namen aufrücken
    Sub Delete_Gaps_in_Name()
        sNameTFoto = Tools.DeletArrayGaps(sNameTFoto)               'Lere arrayelemente ans ende schieben
        sNameTVideo = Tools.DeletArrayGaps(sNameTVideo)               'Lere arrayelemente ans ende schieben
        sNameKFoto = Tools.DeletArrayGaps(sNameKFoto)               'Lere arrayelemente ans ende schieben
        sNameKVideo = Tools.DeletArrayGaps(sNameKVideo)               'Lere arrayelemente ans ende schieben
        'Aufgerückte namen Anzeigen
        For i = 1 To 10
            Me.GroupBoxTeens.Controls("TextBoxTFoto" & i).Text = sNameTFoto(i - 1)
        Next
        For i = 1 To 10
            Me.GroupBoxTeens.Controls("TextBoxTVideo" & i).Text = sNameTVideo(i - 1)
        Next
        For i = 1 To 10
            Me.GroupBoxKids.Controls("TextBoxKFoto" & i).Text = sNameKFoto(i - 1)
        Next
        For i = 1 To 10
            Me.GroupBoxKids.Controls("TextBoxKVideo" & i).Text = sNameKVideo(i - 1)
        Next
    End Sub

End Class