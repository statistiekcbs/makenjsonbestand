Attribute VB_Name = "main_prog"
Public sExportNaam As String
Public sVorigeText As String

Sub zoek_bestand()
    Dim startdir As String
    Dim sFilter As String
    sFilter = "*.xls, *.xlsx, *.xlsb, *.xlsm"
    startdir = ActiveWorkbook.Path
    sVolledigPadBestand = GetFileName(startdir, sFilter)
    sPadBestand = GetPathFromFileName(sVolledigPadBestand)
    sNaamBestand = GetFileNameFromPath(sVolledigPadBestand, False) 'naam, zonder extensie
    sNaamBestand = Replace(sNaamBestand, "krd", "iv3")
    sNaamBestand = Replace(sNaamBestand, "KRD", "iv3")
    sExportNaam = sPadBestand & "\" & sNaamBestand & ".json"
    
    Sheets("start").Cells(17, 3).Value = sVolledigPadBestand
    Sheets("start").Cells(19, 3).Value = sExportNaam
    
End Sub

Public Sub maakJSON_v2()

On Error GoTo EH

Dim sVarNaam(100) As String
Dim sVarWas(100) As String
Dim sVarWordt(100) As String
Dim sVarFormat(100) As String

Dim sheetNaam As String
Dim sopmerking As String
Dim Rekeningkant As String
Dim sOvlaag As String
Dim sjaar As String
Dim rijkop As String
Dim kolkop As String
Dim sOpenbaar As String
Dim bOpenbaar As Boolean
Dim bKeerDuizend As Boolean

Dim sFin_Pakket As String
Dim sExport_softw As String


Dim stekst As String
Dim fso As New FileSystemObject

Dim sOvNaam, sOvNummer, sPeriode, sStatus, sCPnaam, sCPtels, CPmail, sDatum As String

Dim totaal_te_lezen_regels As Integer
Dim BezigMet As Integer
Dim iOvLaag As Integer

Dim oBronSheet As Worksheet


sVorigeText = ""
sDitWb_naam = ActiveWorkbook.Name

Call laad_instellingen 'namen van de sheets en zo ophalen
Call Lees_codes 'lijstje maken van codes en omschrijvingen

'inlezen variabelen
bKeerDuizend = False
With Sheets("Start")
    sBronNaam = .Range("NaamBronBestand").Value
    sExportNaam = .Range("NaamExport").Value
    sOpenbaar = .Range("Details_ja_nee").Value
    sFin_Pakket = .Range("Fin_Pakket").Value
    sExport_softw = .Range("Export_softw").Value
    If .Range("keerDuizend").Value = "Ja" Then bKeerDuizend = True
End With

If sBronNaam = "" Then sFoutmelding = "Er is geen naam van een bronbestand opgegeven." & vbNewLine
If sExportNaam = "" Then sFoutmelding = sFoutmelding & "Er is geen naam van een exportbestand opgegeven." & vbNewLine
If sBronNaam <> "" And Not fso.FileExists(sBronNaam) Then sFoutmelding = sFoutmelding & "Het bronbestand bestaat niet" & vbNewLine
If sOpenbaar = "" Then sFoutmelding = sFoutmelding & "Er is niet opgegeven of de details openbaar mogen worden." & vbNewLine
If sFin_Pakket = "Vul in" Then sFoutmelding = sFoutmelding & "Er is niet opgegeven welk financiële pakket u gebruikt." & vbNewLine
If sExport_softw = "Vul in" Then sFoutmelding = sFoutmelding & "Er is niet opgegeven welke software u gebruikt om de data te exporteren." & vbNewLine

'Debug.Print sFin_Pakket

If sFoutmelding <> "" Then
    MsgBox "Er zijn fouten opgetreden:" & vbNewLine & vbNewLine & sFoutmelding, vbCritical, "Fout geconstateerd"
    Exit Sub
End If

If sOpenbaar = "ja" Then sOpenbaar = "true"
If sOpenbaar = "nee" Then sOpenbaar = "false"

'openen bron
Set oBron = Workbooks.Open(sBronNaam, False)

sWb_BronNaam = ActiveWorkbook.Name

If Not TabbladBestaat("4.Informatie") Then
    MsgBox "Tabblad informatie niet gevonden, script afgebroken"
    Exit Sub
End If

With oBron.Sheets("4.Informatie")
    DoEvents
    
    'verzamelen meta informatie.
    'bij gemeenten en GR's komt de instellng overheidslaag voor, bij provincies niet.
    'dus eerst de ovlaag instellen als provincie, mocht anders blijken overschrijven we dat gewoon.
    'zelfde geldt voor de status
    sOvlaag = "Provincie"
    sStatus = "Realisatie"
    'omdat de info op verschillende regels staat bij verschillende bestanden, langs de koppen lopen op zoek naar de juiste info:
    For regel = 4 To 25
        If .Cells(regel, 2).Value = "Overheidslaag" And .Cells(regel, 3).Value <> "" Then sOvlaag = .Cells(regel, 3).Value
        If .Cells(regel, 2).Value = "Naam" And .Cells(regel, 3).Value <> "" Then sOvNaam = .Cells(regel, 3).Value
        If .Cells(regel, 2).Value = "Nummer" And .Cells(regel, 3).Value <> "" Then sOvNummer = .Cells(regel, 3).Value
        If .Cells(regel, 2).Value = "Jaar" And .Cells(regel, 3).Value <> "" Then sjaar = .Cells(regel, 3).Value
        If .Cells(regel, 2).Value = "Periode" And .Cells(regel, 3).Value <> "" Then sPeriode = .Cells(regel, 3).Value
        If .Cells(regel, 2).Value = "Status" And .Cells(regel, 3).Value <> "" Then sStatus = .Cells(regel, 3).Value
        If .Cells(regel, 2).Value = "Naam: " And .Cells(regel, 3).Value <> "" Then sCPnaam = .Cells(regel, 3).Value
        If .Cells(regel, 2).Value = "Telefoon: " And .Cells(regel, 3).Value <> "" Then sCPtel = .Cells(regel, 3).Value
        If .Cells(regel, 2).Value = "E-mail: " And .Cells(regel, 3).Value <> "" Then sCPmail = .Cells(regel, 3).Value
        If .Cells(regel, 2).Value = "Datum: " And .Cells(regel, 3).Value <> "" Then sDatum = .Cells(regel, 3).Value
    Next
    Workbooks(sDitWb_naam).Sheets("Opzoek").Cells(1, 2).Value = CInt(sjaar)
    Call Lees_codes 'lijstje maken van codes en omschrijvingen

    iJaar = CInt(sjaar)
    iPeriode = CInt(sPeriode)
    If sDatum = "" Then sDatum = "1-1-2020"
    sDatum = maak_iso_datum(sDatum)
    If LCase(sOvlaag) = "provincie" Then iOvLaag = 3
    If LCase(sOvlaag) = "gr" Then iOvLaag = 5
    If LCase(sOvlaag) = "gemeente" Then iOvLaag = 6
    
    Call SchrijfMeta(sOvlaag, sOvNaam, sOvNummer, iJaar, iPeriode, sStatus, sOpenbaar, sFin_Pakket, sExport_softw, sCPnaam, sCPtel, sCPmail, sDatum)

    'zoeken naar opmerkingen...
    eind_opm = 50 'alvast een maximum meegeven
    For regel = 20 To 50
        If InStr(.Cells(regel, 2).Value, "Ruimte voor toelichting") > 0 Then begin_opm = regel + 1
        If InStr(.Cells(regel, 2).Value, "Boekwinst/verlies") > 0 Then eind_opm = regel - 1
    Next
    bVerzamelGeschreven = False
    For regel = begin_opm To eind_opm 'door de opmerkingen heen lopen
        If .Cells(regel, 2).Value <> "" Then
            If bVerzamelGeschreven = False Then
                Call SchrijfVerzamel("opmerkingen")
                bVerzamelGeschreven = True
            End If
            sopmerking = .Cells(regel, 2).Value
            bEentjeGevonden = True
            Call SchrijfOpmerking(sopmerking, False)
        End If
    Next
    If bEentjeGevonden = True Then
            Call SchrijfOpmerking(sopmerking, True)
'        Else
'            Call SchrijfSluitRegel(False)
        End If
    If bVerzamelGeschreven = True Then Call SchrijfSluitRegel(False)
    
End With

Call SchrijfKopData 'hier wordt het element data geschreven

Workbooks(sDitWb_naam).Activate

Call geef_instelling_2(sOvlaag, sjaar, iOvLaag)
'Debug.Print array_instellingen
totaal_te_lezen_regels = 0
For element = 1 To MaxAantalElementen(iOvLaag)
    ivanaf = 0: itot = 0
    If Instelling_nummers(element, RIJNR_TM) <> "" Then itot = CInt(Instelling_nummers(element, RIJNR_TM))
    If Instelling_nummers(element, RIJNR_VA) <> "" Then ivanaf = CInt(Instelling_nummers(element, RIJNR_VA))
    totaal_te_lezen_regels = totaal_te_lezen_regels + (itot - ivanaf)
Next
BezigMet = 0: Call statusBar_bijwerken(totaal_te_lezen_regels, BezigMet, False)


For elementNr = 1 To MaxAantalElementen(iOvLaag) 'lasten, baten, balans_lasten, etc
    elementNaam = sJSONElement(elementNr)
    
    ' Wijzigingen (04-07-2025) - VOHM
    ' sheet 12.Beleidsindicatoren is niet meer aanwezig in het Excel sjabloon (iv3-model) vanaf boekjaar 2026
    sheetNaam = Instelling_nummers(elementNr, TAB_NAAM)
    Set oBronSheet = Nothing
    On Error Resume Next
    Set oBronSheet = oBron.Sheets(sheetNaam)
    On Error GoTo EH
    If oBronSheet Is Nothing Then
        ' indien een sheet niet bestaat in het Excel sjabloon, dan schrijven we een leeg element
        Call SchrijfVerzamel(elementNaam)
        GoTo VolgendeElement
    End If
    ' Einde wijzigingen (04-07-2025)
    
    With oBron.Sheets(sheetNaam)
        Call SchrijfVerzamel(elementNaam)
        'eerst kijken of er iets in ingevuld in een element, dan pas element-record schrijven
        For regel = CInt(Instelling_nummers(elementNr, RIJNR_VA)) To CInt(Instelling_nummers(elementNr, RIJNR_TM))
            BezigMet = BezigMet + 1: Call statusBar_bijwerken(totaal_te_lezen_regels, BezigMet, False)
            For kolom = CInt(Instelling_nummers(elementNr, KOLNR_VA)) To CInt(Instelling_nummers(elementNr, KOLNR_TM))
                bTotaalGevonden = False
                If CInt(Instelling_nummers(elementNr, KOLNR_KOP_CODES)) <> 0 Then
                    If InStr(.Cells(regel, CInt(Instelling_nummers(elementNr, KOLNR_KOP_CODES))).Value, "Totaal") <> 0 Then bTotaalGevonden = True
                End If
                If CInt(Instelling_nummers(elementNr, RIJNR_KOP_OMSCHR)) <> 0 Then
                    If InStr(.Cells(CInt(Instelling_nummers(elementNr, RIJNR_KOP_OMSCHR)), kolom).Value, "Totaal") <> 0 Then bTotaalGevonden = True
                End If
                If Trim(.Cells(regel, kolom).Value) <> "" And .Cells(regel, kolom).Value <> 0 And bTotaalGevonden = False Then
                    'er is een gevulde cel gevonden, dus moeten we de waarde en de codes ophalen
                    If bKeerDuizend = True And elementNaam <> "kengetallen" And elementNaam <> "beleidsindicatoren" Then
                        waarde = .Cells(regel, kolom).Value * 1000
                    Else
                        waarde = .Cells(regel, kolom).Value
                    End If
                    'bij sommige tabbladen zijn geen codes opgenomen. Hiervoor zal de code opgezocht moeten worden
                    If Instelling_nummers(elementNr, KOLNR_KOP_CODES) = "0" Then
                        rijkop = Opzoeken_code(elementNr, "Rij", .Cells(regel, CInt(Instelling_nummers(elementNr, KOLNR_KOP_OMSCHR))))
                    Else
                        rijkop = .Cells(regel, CInt(Instelling_nummers(elementNr, KOLNR_KOP_CODES)))
                    End If
                    If Instelling_nummers(elementNr, RIJNR_KOP_CODES) = "0" Then
                        kolkop = Opzoeken_code(elementNr, "Kolom", .Cells(CInt(Instelling_nummers(elementNr, RIJNR_KOP_OMSCHR)), kolom))
                    Else
                        kolkop = .Cells(CInt(Instelling_nummers(elementNr, RIJNR_KOP_CODES)), kolom)
                    End If
                    bEentjeGevonden = True
                    Call SchrijfRecord(sRijKopSleutel(elementNr), rijkop, _
                                        sKolomKopSleutel(elementNr), kolkop, _
                                        sBedragSleutel(elementNr), waarde, _
                                        bBedragIsString(elementNr), False, False)
                                       
                End If
            Next kolom
        Next regel
        If bEentjeGevonden = True Then
            Call SchrijfRecord(sRijKopSleutel(elementNr), rijkop, _
                               sKolomKopSleutel(elementNr), kolkop, _
                               sBedragSleutel(elementNr), waarde, _
                               bBedragIsString(elementNr), True, False)
        Else
            Call SchrijfSluitRegel(False)
        End If
    End With
    
VolgendeElement:
    If elementNr = MaxAantalElementen(iOvLaag) Then 'als het de laatste is, moet er geen komma achter het haakje
        Call SchrijfSluitRegel(True)
    Else
        Call SchrijfSluitRegel(False)
    End If
Next elementNr

Call Afsluiting
oBron.Close False

MsgBox "Klaarrrrrrrr"

Call statusBar_bijwerken(totaal_te_lezen_regels, BezigMet, True)

Exit Sub
EH:
'waarschijnlijk een onjuist jaar of zo.
sBericht = "Verwerking is vastgelopen." & vbNewLine & _
            "Er zijn verschillende mogelijkheden." & vbNewLine & _
            "Het gekozen bestand is van een provincie of gemeente van vóór 2017." & vbNewLine & _
            "Het gekozen bestand is van een gemeenschappelijke Regeling van vóór 2018." & vbNewLine & _
            "Het gekozen bestand is geen valide iv3 excel." & vbNewLine & _
            "Het tabblad informatie is niet goed ingevuld of bestaat niet" & vbNewLine & _
            vbNewLine & _
            "Script gestopt, er is geen JSON bestand aangemaakt."
MsgBox sBericht, vbCritical, "Vaudt"
fso.DeleteFile sExportNaam

End Sub

Public Function TabbladBestaat(sheetNaam As String) As Boolean
     TabbladBestaat = False
      For Each WS In Worksheets
        If sheetNaam = WS.Name Then
          TabbladBestaat = True
          Exit Function
        End If
      Next WS
End Function
Public Function SchrijfSluitRegel(bAllerlaatsteRegel As Boolean)
Dim fso As New FileSystemObject
Set objTekstBestand = fso.OpenTextFile(sExportNaam, 8)
    If bAllerlaatsteRegel = True Then
        objTekstBestand.WriteLine ("]")
    Else
        objTekstBestand.WriteLine ("],")
    End If
    sVorigeText = ""
End Function
Public Function SchrijfRegel(blaatsteRegel, cat, taakv, balcode, standper, bedr, Optional bAllerlaatsteRegel As Boolean)
Dim stekst As String
Dim fso As New FileSystemObject

'rekeningkant: String,  (baten, lasten, balans)
'                  categorie: String,
'                  taakveld: String,
'                  balanscode: String,
'                  standper: String, (1jan ultimo)
'                  Bedrag:Integer

stekst = "{"
    If cat <> "" Then stekst = stekst & Chr(34) & "categorie" & Chr(34) & ":" & Chr(34) & cat & Chr(34) & ","
    If taakv <> "" Then stekst = stekst & Chr(34) & "taakveld" & Chr(34) & ":" & Chr(34) & taakv & Chr(34) & ","
    If balcode <> "" Then stekst = stekst & Chr(34) & "balanscode" & Chr(34) & ":" & Chr(34) & balcode & Chr(34) & ","
    If standper <> "" Then stekst = stekst & Chr(34) & "standper" & Chr(34) & ":" & Chr(34) & standper & Chr(34) & ","
    stekst = stekst & Chr(34) & "bedrag" & Chr(34) & ":" & bedr & "}"

If sVorigeText = "" Then
    sVorigeText = stekst
Else
    Set objTekstBestand = fso.OpenTextFile(sExportNaam, 8)
    If blaatsteRegel = False Then
        objTekstBestand.WriteLine (sVorigeText & ",")
        sVorigeText = stekst
    Else
        If bAllerlaatsteRegel = True Then
            objTekstBestand.WriteLine (sVorigeText & "]")
        Else
            objTekstBestand.WriteLine (sVorigeText & "],")
        End If
        sVorigeText = ""
    End If
End If

'Debug.Print sTekst
DoEvents
End Function
Public Function SchrijfKen_Beleid(blaatsteRegel, kengetal, beleids, verslagper, bedr, Optional bAllerlaatsteRegel As Boolean)
Dim stekst As String
Dim fso As New FileSystemObject

'rekeningkant: String,  (kengetal of beleidsindicator)
'                  verslagperiode: String,
'                  Bedrag:String

stekst = "{"
    If kengetal <> "" Then stekst = stekst & Chr(34) & "kengetal" & Chr(34) & ":" & Chr(34) & cat & Chr(34) & ","
    If taakv <> "" Then stekst = stekst & Chr(34) & "beleidsindicator" & Chr(34) & ":" & Chr(34) & taakv & Chr(34) & ","
    If balcode <> "" Then stekst = stekst & Chr(34) & "verslagperiode" & Chr(34) & ":" & Chr(34) & balcode & Chr(34) & ","
    stekst = stekst & Chr(34) & "bedrag" & Chr(34) & ":" & Chr(34) & bedr & Chr(34) & "}"

If sVorigeText = "" Then
    sVorigeText = stekst
Else
    Set objTekstBestand = fso.OpenTextFile(sExportNaam, 8)
    If blaatsteRegel = False Then
        objTekstBestand.WriteLine (sVorigeText & ",")
        sVorigeText = stekst
    Else
        If bAllerlaatsteRegel = True Then
            objTekstBestand.WriteLine (sVorigeText & "]")
        Else
            objTekstBestand.WriteLine (sVorigeText & "],")
        End If
        sVorigeText = ""
    End If
End If

'Debug.Print sTekst
DoEvents
End Function

Public Function SchrijfOpmerking(sopmerking As String, blaatsteRegel As Boolean)
Dim stekst As String
Dim fso As New FileSystemObject

stekst = "{" & Chr(34) & "tekst" & Chr(34) & ":" & Chr(34) & sopmerking & Chr(34) & "}"
    
If sVorigeText = "" Then
    sVorigeText = stekst
Else
    Set objTekstBestand = fso.OpenTextFile(sExportNaam, 8)
    If blaatsteRegel = False Then
        objTekstBestand.WriteLine (sVorigeText & ",")
        sVorigeText = stekst
    Else
        objTekstBestand.WriteLine (sVorigeText)
        sVorigeText = ""
    End If
End If

End Function


Public Function SchrijfRecord(sSleutel_1 As String, sWaarde_1 As String, sSleutel_2 As String, sWaarde_2 As String, _
                            sSleutelBedrag As String, Bedrag As Variant, bBedragIsString As Boolean, blaatsteRegel As Boolean, _
                            bAllerlaatsteRegel As Boolean)
                            
Dim stekst As String
Dim fso As New FileSystemObject

If bBedragIsString = False Then
    stekst = "{"
        stekst = stekst & Chr(34) & sSleutel_1 & Chr(34) & ":" & Chr(34) & sWaarde_1 & Chr(34) & ","
        stekst = stekst & Chr(34) & sSleutel_2 & Chr(34) & ":" & Chr(34) & sWaarde_2 & Chr(34) & ","
        stekst = stekst & Chr(34) & sSleutelBedrag & Chr(34) & ":" & Round(CDbl(Bedrag), 0) & "}"
Else
    stekst = "{"
        stekst = stekst & Chr(34) & sSleutel_1 & Chr(34) & ":" & Chr(34) & sWaarde_1 & Chr(34) & ","
        stekst = stekst & Chr(34) & sSleutel_2 & Chr(34) & ":" & Chr(34) & sWaarde_2 & Chr(34) & ","
        stekst = stekst & Chr(34) & sSleutelBedrag & Chr(34) & ":" & Chr(34) & Bedrag & Chr(34) & "}"
End If

If sVorigeText = "" Then
    sVorigeText = stekst
Else
    Set objTekstBestand = fso.OpenTextFile(sExportNaam, 8)
    If blaatsteRegel = False Then
        objTekstBestand.WriteLine (sVorigeText & ",")
        sVorigeText = stekst
    Else
'        If bAllerlaatsteRegel = True Then
'            objTekstBestand.WriteLine (sVorigeText & "]")
'        Else
'            objTekstBestand.WriteLine (sVorigeText & "],")
'        End If
        objTekstBestand.WriteLine (sVorigeText)
        sVorigeText = ""
    End If
End If

'Debug.Print sTekst
DoEvents
End Function





Public Function SchrijfVerzamel(kant)
Dim stekst As String
Dim fso As New FileSystemObject
Set objTekstBestand = fso.OpenTextFile(sExportNaam, 8)
stekst = Chr(34) & kant & Chr(34) & ":["

objTekstBestand.WriteLine (stekst)

'Debug.Print sTekst

End Function


Public Function SchrijfMeta(sOvlaag, sOvNaam, sOvNummer, iJaar, iPeriode, sStatus, sDetails, sFinPak, sExpTool, sCPnaam, sCPtel, sCPmail, sDatum)
Dim stekst As String
Dim fso As New FileSystemObject

If fso.FileExists(sExportNaam) Then
    fso.DeleteFile (sExportNaam)
End If
Set f = fso.CreateTextFile(sExportNaam)
f.Close

Set objTekstBestand = fso.OpenTextFile(sExportNaam, 8)

If sDetails = "Ja" Then
    sDetails = "true"
Else
    sDetails = "false"
End If

stekst = "{" & Chr(34) & "metadata" & Chr(34) & ":{"
stekst = stekst & Chr(34) & "overheidslaag" & Chr(34) & ":" & Chr(34) & sOvlaag & Chr(34) & ","
stekst = stekst & Chr(34) & "overheidsnummer" & Chr(34) & ":" & Chr(34) & sOvNummer & Chr(34) & ","
stekst = stekst & Chr(34) & "overheidsnaam" & Chr(34) & ":" & Chr(34) & sOvNaam & Chr(34) & ","
stekst = stekst & Chr(34) & "boekjaar" & Chr(34) & ":" & iJaar & ","
stekst = stekst & Chr(34) & "periode" & Chr(34) & ":" & iPeriode & ","
stekst = stekst & Chr(34) & "status" & Chr(34) & ":" & Chr(34) & sStatus & Chr(34) & ","
stekst = stekst & Chr(34) & "datum" & Chr(34) & ":" & Chr(34) & sDatum & Chr(34) & ","
stekst = stekst & Chr(34) & "details_openbaar" & Chr(34) & ":" & sDetails & ","
stekst = stekst & Chr(34) & "financieel_pakket" & Chr(34) & ":" & Chr(34) & sFinPak & Chr(34) & ","
stekst = stekst & Chr(34) & "export_software" & Chr(34) & ":" & Chr(34) & sExpTool & Chr(34) & "},"

objTekstBestand.WriteLine (stekst)

stekst = Chr(34) & "contact" & Chr(34) & ":{"
stekst = stekst & Chr(34) & "naam" & Chr(34) & ":" & Chr(34) & sCPnaam & Chr(34) & ","
stekst = stekst & Chr(34) & "telefoon" & Chr(34) & ":" & Chr(34) & sCPtel & Chr(34) & ","
stekst = stekst & Chr(34) & "email" & Chr(34) & ":" & Chr(34) & sCPmail & Chr(34) & "},"
objTekstBestand.WriteLine (stekst)

End Function

Public Function SchrijfKopData()
Dim stekst As String
Dim fso As New FileSystemObject

Set objTekstBestand = fso.OpenTextFile(sExportNaam, 8)

    stekst = Chr(34) & "data" & Chr(34) & ":{"
    objTekstBestand.WriteLine (stekst)


End Function



Public Function Afsluiting()
Dim fso As New FileSystemObject
Set objTekstBestand = fso.OpenTextFile(sExportNaam, 8)
objTekstBestand.WriteLine ("}}")
End Function

Public Function statusBar_bijwerken(Totaal As Integer, BezigMet As Integer, klaar As Boolean)
On Error GoTo EH_status
Dim CurrentStatus As Integer
Dim NumberOfBars As Integer
Dim pctDone As Integer
Dim lastrow As Long, i As Long

'(Step 1) Display Status Bar
NumberOfBars = 80
'Application.StatusBar = "[" & Space(NumberOfBars) & "]"
If BezigMet > Totaal Then BezigMet = Totaal
'(Step 2) Periodically update your Status Bar
    CurrentStatus = Int((BezigMet / Totaal) * NumberOfBars)
    pctDone = Round(CurrentStatus / NumberOfBars * 100, 0)
    Application.StatusBar = "Voortgang: [" & String(CurrentStatus, "|") & _
                            Space(NumberOfBars - CurrentStatus) & "]" & _
                            "   " & pctDone & "% Compleet"
    DoEvents
If klaar = True Then Application.StatusBar = ""

Exit Function
EH_status:
MsgBox "vastgelopen in statusbarverwerking"

End Function
