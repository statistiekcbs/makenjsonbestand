Attribute VB_Name = "tools"

Public Function Opzoeken_code(elementNr, RijOfKolomKop, sOmschr As String) As String
    
'als het een kolomkop is, is het altijd verslagperiode
If RijOfKolomKop = "Kolom" Then
    Opzoeken_code = geef_code("verslagperiode", sOmschr)
Else
    If sJSONElement(elementNr) = "kengetallen" Then
        Opzoeken_code = geef_code("kengetal", sOmschr)
    End If
    If sJSONElement(elementNr) = "beleidsindicatoren" Then
        Opzoeken_code = geef_code("beleidsindicator", sOmschr)
    End If

End If
    
If Opzoeken_code = "" Then MsgBox "Er is iets misgegaan met het opzoeken van de code voor " & sOmschr


End Function

Public Function maak_iso_datum(sDatum As String) As String
On Error GoTo EH
Dim dDatum As Date
dDatum = CDate(sDatum)

    CurrentTime = Year(dDatum) & "/" & Month(dDatum) & "/" & Day(dDatum) & " " & Hour(dDatum) & ":" & Minute(dDatum) & ":" & Second(dDatum)
    maak_iso_datum = Application.WorksheetFunction.Text(CurrentTime, "yyyy-mm-ddThh:MM:ss") & "+02:00"

Exit Function
EH:
    MsgBox "datum informatieblad niet herkend, graag zo: dd-mm-jjjj", vbCritical, "datum fout"
    End
End Function


Public Function geef_code(Soort As String, waarde As String) As String
'"kengetal", .Cells(regel, array_instellingen(18)).Value)
Dim iSoortNr As Integer
Dim iCodeNr As Integer

For i = 1 To 5
    If codeSoort(i) = Soort Then iSoortNr = i
Next

For a = 1 To 20
    If omschrijving(iSoortNr, a) = waarde Then iCodeNr = a
Next

geef_code = code(iSoortNr, iCodeNr)



End Function




Public Function ZoekBestand(iFilenr As Integer) As Boolean
'opzoeken pad en bestand mbv verkenner voor bestand
On Error GoTo Err_ZoekBestand
Dim varFileName, strFile As String
Dim strFile1 As String, strFile2 As String
Dim bln As Boolean
 
    'openen dialoogbox van de verkenner en teruggeven gekozen bestand
    varFileName = Application.GetOpenFilename
    
    'alleen indien een bestand gekozen doorgaan
    If Len(varFileName) > 0 And Not (varFileName = False) Then
        strFile = CStr(varFileName)
        ThisWorkbook.Activate
        Sheets(c_shMain).Select
        'schermbeveiliging uitzetten
        ProtectSheets False, c_shMain
        
        bln = False
        If iFilenr = 1 Then
            strFile1 = GetFileNameFromPath(strFile, True)
            strFile2 = GetFileNameFromPath(Worksheets(c_shMain).txtFile2.Value, True)
            If strFile1 <> strFile2 Then
                bln = True
                'complete naam invullen in vak
                Worksheets(c_shMain).txtFile1.Value = strFile
                'of alleen het pad neerzetten?, nee compleet pad, die wordt verderop gebruikt
                'Worksheets(c_shMain).txtFile1.Value = GetPathFromFileName(strFile)
            End If
            
        Else
            strFile1 = GetFileNameFromPath(Worksheets(c_shMain).txtFile1.Value, True)
            strFile2 = GetFileNameFromPath(strFile, True)
            If strFile1 <> strFile2 Then
                bln = True
                Worksheets(c_shMain).txtFile2.Value = strFile
            End If
        End If
        
        If Not bln Then
            'foutmelding geven indien bestandsnamen gelijk zijn, Excel kan geen 2 gelijke bestandsnamen openen!
            MsgBox "Je kunt geen 2 bestanden controleren met dezelfde naam !", vbOKOnly, c_strTitle
            If iFilenr = 1 Then
                Worksheets(c_shMain).txtFile1.Value = ""
            Else
                Worksheets(c_shMain).txtFile2.Value = ""
            End If
                
        Else
            'leegmaken en vullen tabbladen overzicht
            Call LeegmakenTabOverzichten(iFilenr)
            If iFilenr = 1 Then
                Call InvullenTabOverzichten(strFile, c_ColSheetName1, c_ColSheettext1)
            Else
                Call InvullenTabOverzichten(strFile, c_ColSheetName2, c_ColSheettext2)
            End If
        End If
        
    End If
    
Exit_ZoekBestand:
    'schermbeveiliging weer aanzetten
    ProtectSheets True, c_shMain
    Exit Function
    
Err_ZoekBestand:
    MsgBox "fout in invullen bestandsgegevens" & vbCrLf & _
        Err.Number & " " & Err.Description, vbOKOnly, c_strTitle
    'Resume Next
    Resume Exit_ZoekBestand
End Function

Public Function GetFileName(Optional strStartDir As String, Optional strFilterFiles As String, _
    Optional strTitle As String) As String
'strFilterFiles meegeven als bijvoorbeeld : "*.xls, *.xlsx, *.csv, *.txt"
'indien fout in filter dan worden alle bestanden getoond
On Error GoTo Err_f
Dim fDialog As Office.FileDialog
Dim strFile As String
Dim strDrive As String, strDriveletter As String

    If Len(strTitle) = 0 Then strTitle = "Selecteer een bestand"
    If Len(strStartDir) = 0 Then strStartDir = ActiveWorkbook.Path

    ' Set up the File Dialog.
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)

    With fDialog

        If Len(strStartDir) > 0 Then .InitialFileName = strStartDir

        ' Allow user to make single selections in dialog box
        .AllowMultiSelect = False
        ' Set the title of the dialog box.
        .Title = strTitle

        ' Clear out the current filters, and add our own.
        .Filters.Clear
        If Len(strFilterFiles) > 0 Then
            .Filters.Add "Bestanden : ", strFilterFiles
        Else
            .Filters.Add "All Files", "*.*"
        End If

        ' Show the dialog box.
        ' If the .Show method returns True, then user picked at least one file.
        ' If the .Show method returns False, the user clicked Cancel.
        If .Show = True Then
            strFile = .SelectedItems(1)
        Else
            strFile = ""
        End If
    End With
    
    'omzetten filename naar compleet pad indien met letters gewerkt
    If Mid(strFile, 2, 1) = ":" Then
        strDriveletter = Left(strFile, 2)
        strDrive = GetDriveFullname(strDriveletter, False)
        If Len(strDrive) > 0 Then
            strFile = Replace(strFile, strDriveletter, strDrive)
        End If
    End If

    GetFileName = strFile

    Set fDialog = Nothing
    Exit Function

Err_f:
    If Not fDialog Is Nothing Then
        fDialog.Filters.Clear
        fDialog.Filters.Add "All Files", "*.*"
        Resume Next
    End If
    Exit Function
    Resume 0
End Function

Public Function GetDriveFullname(ByVal strDriveletter As String, Optional blnMessage As Boolean = False) As String
Dim objFSO As Object, objShell As Object, objFile As Object
Dim strFile As Variant, strFileOut As Variant
Dim strNet As String
Dim strMessage As String

'net use Z:  levert:
'Lokale naam             Z:
'Externe naam            \\cbsp.nl\Infrastructuur\Apps\Productie
'Type netwerkbron        Schijf
'De opdracht is voltooid.

    If Right(strDriveletter, 1) <> ":" Then strDriveletter = strDriveletter & ":"
    If Len(strDriveletter) > 2 Then
        strMessage = strDriveletter & " is geen juiste driveletter"
    
    Else
        Set objShell = CreateObject("Wscript.Shell")
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
        strFile = GetTempDir & "\kanweg.bat"
        strFileOut = GetTempDir & "\netuse.txt"
        
        If objFSO.FileExists(strFile) Then objFSO.DeleteFile strFile
        If objFSO.FileExists(strFileOut) Then objFSO.DeleteFile strFileOut
                
        strNet = "NET USE " & strDriveletter & " > " & strFileOut
        Set objFile = objFSO.CreateTextFile(strFile)
        objFile.Write strNet
        objFile.Close
        
        'de bat-file uitvoeren en dus de uitvoer van de net use in een string opvangen
        'aangepast zodat de dosbox niet te zien is
        objShell.Run strFile, 0, True
        strNet = objFSO.OpenTextFile(strFileOut).ReadAll()
        
        If InStr(1, strNet, "Externe naam", vbTextCompare) > 0 Then
            strNet = Mid(strNet, InStr(1, strNet, "\\cbsp.nl"))
            strNet = Left(strNet, InStr(1, strNet, vbCrLf) - 1)
            GetDriveFullname = strNet
            
        Else
            strMessage = "Driveletter niet in gebruik"
        End If
        
        If Len(strMessage) > 0 And blnMessage Then MsgBox strMessage, vbOKOnly, "GetDriveFullname"
    End If

Exit_f:
    If objFSO.FileExists(strFile) Then objFSO.DeleteFile strFile
    If objFSO.FileExists(strFileOut) Then objFSO.DeleteFile strFileOut
    Set objShell = Nothing
    Set objFile = Nothing
    Set objFSO = Nothing
    Set objShell = Nothing
End Function
Public Function GetFileNameFromPath(ByVal strFilename As String, _
    Optional ByVal blnExtension As Boolean = True) As String
Dim i As Integer
Dim t As Integer
Dim strName As String

    strName = strFilename
    t = Len(strFilename)
    For i = t To 1 Step -1
        If Mid(strFilename, i, 1) = "\" Then
            strName = Right(strFilename, t - i)
            Exit For
        End If
    Next i
    If blnExtension = False Then
        t = Len(strName)
        For i = t To 1 Step -1
            If Mid(strName, i, 1) = "." Then
                strName = Left(strName, i - 1)
                Exit For
            End If
        Next i
    End If
    GetFileNameFromPath = strName
    
End Function

Public Function GetPathFromFileName(ByVal strFilename As String) As String
Dim i As Integer
Dim t As Integer
Dim strDir As String

    t = Len(strFilename)
    For i = t To 1 Step -1
        If Mid(strFilename, i, 1) = "\" Then
            strDir = Left(strFilename, i)
            Exit For
        End If
    Next i
    GetPathFromFileName = strDir
    
End Function

