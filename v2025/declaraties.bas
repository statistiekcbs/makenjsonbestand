Attribute VB_Name = "declaraties"
Public code(5, 20) As String
Public omschrijving(5, 20) As String
Public codeSoort(5) As String
Public sJSONElement(20) As String 'de verschillende afdelingen binnen de jsondata
Public sRijKopSleutel(20) As String 'sleutelnaam van de rijcodes voor JSON output
Public sKolomKopSleutel(20) As String 'sleutelnaam van de kolomcodes voor JSON output
Public sBedragSleutel(20) As String 'sleutelnaam van de waarde voor JSON output
Public bBedragIsString(20) As Boolean 'bedragen zijn longs, behalve kengetallen en beleidsindicatoren
Public Instelling_nummers(10, 20) As String 'per element met begin en eind regel en koppen e.d.

Public sDitWb_naam As String
Public iJaar As Integer
Public iPeriode As Integer
Public sDatum As String
Public sStatus As String
Public HUIDIG_PAD As String
Public HUIDIGE_EXCEL As String


Public Const TAB_NAAM = 1
Public Const RIJNR_KOP_CODES = 2
Public Const RIJNR_KOP_OMSCHR = 3
Public Const KOLNR_KOP_CODES = 4
Public Const KOLNR_KOP_OMSCHR = 5
Public Const RIJNR_VA = 6
Public Const RIJNR_TM = 7
Public Const KOLNR_VA = 8
Public Const KOLNR_TM = 9

Public MaxAantalElementen(10) As Integer 'afhankelijk van de ovlaag
Public sExportNaam As String
Public sVorigeText As String


Public ov_info(100, 15) As String ' alle informatie van het tabblad ov_lijst
Public Const OV_LAAG = 1
Public Const OV_NR = 2
Public Const OV_NAAM = 3
Public Const DETAILS = 4
Public Const FIN_PAK = 5
Public Const EXP_TOOL = 6
Public Const CP_NAAM = 7
Public Const CP_TEL = 8
Public Const CP_MAIL = 9


Public Function laad_instellingen()

'export elementen
sJSONElement(1) = "lasten"
sJSONElement(2) = "balans_lasten"
sJSONElement(3) = "baten"
sJSONElement(4) = "balans_baten"
sJSONElement(5) = "balans_standen"
sJSONElement(6) = "kengetallen"
sJSONElement(7) = "beleidsindicatoren"

'als het een provincie of gr betreft, zullen de tabbladen 6 en 7 niet voorkomen.
MaxAantalElementen(3) = 5
MaxAantalElementen(5) = 5
MaxAantalElementen(6) = 7


sRijKopSleutel(1) = "taakveld": sKolomKopSleutel(1) = "categorie": sBedragSleutel(1) = "bedrag": bBedragIsString(1) = False
sRijKopSleutel(2) = "balanscode": sKolomKopSleutel(2) = "categorie": sBedragSleutel(2) = "bedrag": bBedragIsString(2) = False
sRijKopSleutel(3) = "taakveld": sKolomKopSleutel(3) = "categorie": sBedragSleutel(3) = "bedrag": bBedragIsString(3) = False
sRijKopSleutel(4) = "balanscode": sKolomKopSleutel(4) = "categorie": sBedragSleutel(4) = "bedrag": bBedragIsString(4) = False
sRijKopSleutel(5) = "balanscode": sKolomKopSleutel(5) = "standper": sBedragSleutel(5) = "bedrag": bBedragIsString(5) = False
sRijKopSleutel(6) = "kengetal": sKolomKopSleutel(6) = "verslagperiode": sBedragSleutel(6) = "waarde": bBedragIsString(6) = True
sRijKopSleutel(7) = "beleidsindicator": sKolomKopSleutel(7) = "verslagperiode": sBedragSleutel(7) = "waarde": bBedragIsString(7) = True

HUIDIG_PAD = ActiveWorkbook.Path
HUIDIGE_EXCEL = ActiveWorkbook.Name

End Function

Public Function geef_instelling_2(sOvlaag As String, sjaar As String, iOvLaag As Integer) As Boolean
Dim beginReg As Integer
Dim beginKol As Integer
Dim MaxTeLezenKol As Integer

MaxTeLezenKol = 9

With Sheets("Opzoek")
    'aan de hand van ovlaag en jaar het juiste blok zoeken..
    For reg = 1 To 100
        If .Cells(reg, 1).Value = "Jaren" Then
            For kol = 1 To 100
                If CStr(.Cells(reg, kol).Value) = sjaar Then
                    beginKol = kol
                    Exit For
                End If
            Next kol
        End If
        If .Cells(reg, 1).Value = sOvlaag Then beginReg = reg
    Next reg
    If beginReg = 0 Or beginKol = 0 Then Er_Bericht = "jaar niet gevonden": GoTo Er_instellingen
    'inlezen arrays voor elk JSON element.
    For Element_nr = 1 To MaxAantalElementen(iOvLaag)
        'nu moeten we van elk van de 7 elementen de juiste regel vinden
        For rij = beginReg To beginReg + MaxAantalElementen(iOvLaag)
            If LCase(.Cells(rij, 2).Value) = LCase(sJSONElement(Element_nr)) Then
                'juiste rij gevonden, reeks invullen
                volgnr = 0
                For kol = beginKol To beginKol + MaxTeLezenKol
                    volgnr = volgnr + 1
                    Instelling_nummers(Element_nr, volgnr) = CStr(.Cells(rij, kol).Value)
                Next kol
            End If
        Next rij
    Next Element_nr
    
    
End With
'MsgBox beginReg & "   " & beginKol
Exit Function
Er_instellingen:
MsgBox Er_Bericht, vbCritical, "Fout in instellingen zoeken"


End Function


Public Function Lees_codes() As Boolean


codeSoort(1) = "kengetal"
codeSoort(2) = "beleidsindicator"
codeSoort(3) = "verslagperiode"

With Workbooks(sDitWb_naam).Sheets("Opzoek")
    
    For kolom = 1 To 50
        If InStr(.Cells(1, kolom), "Kengetallen Vertaling") > 0 Then
            nr = 0
            For regel = 3 To 14
                If .Cells(regel, kolom).Value <> "" Then
                    nr = nr + 1
                    omschrijving(1, nr) = .Cells(regel, kolom).Value
                    code(1, nr) = .Cells(regel, kolom + 1).Value
                End If
            Next
        End If
        
        If InStr(.Cells(1, kolom), "beleidsindicatoren Vertaling") > 0 Then
            nr = 0
            For regel = 3 To 14
                If .Cells(regel, kolom).Value <> "" Then
                    nr = nr + 1
                    omschrijving(2, nr) = .Cells(regel, kolom).Value
                    code(2, nr) = .Cells(regel, kolom + 1).Value
                End If
            Next
        End If
        
        If InStr(.Cells(1, kolom), "verslagperiode Vertaling") > 0 Then
            nr = 0
            For regel = 3 To 14
                If .Cells(regel, kolom).Value <> "" Then
                    nr = nr + 1
                    omschrijving(3, nr) = .Cells(regel, kolom).Value
                    code(3, nr) = .Cells(regel, kolom + 1).Value
                End If
            Next
        End If
        
    Next kolom
End With

Lees_codes = True

End Function


