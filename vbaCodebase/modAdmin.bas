Attribute VB_Name = "modAdmin"
Option Explicit

' Remplit al�atoirement toutes les cellules d'�paisseur (officielles et rattrapage) du rouleau actif
' @but : Simuler des mesures d'�paisseur pour test ou d�mo
' @param Aucun
' @return Aucun
' @pr� : Les plages nomm�es d'�paisseur doivent exister et PRODUCTION_WS doit �tre initialis�
Public Sub FillThicknessCellsRandom()
    Dim ws As Worksheet
    Set ws = PRODUCTION_WS
    Dim thickNames As Variant
    thickNames = Array("leftThicknessCels", "rightThicknessCels", "leftSecThicknessCels", "rightSecThicknessCels")
    Dim i As Integer, thickName As String, refString As String
    For i = LBound(thickNames) To UBound(thickNames)
        thickName = thickNames(i)
        If NameExists(thickName) Then
            refString = ThisWorkbook.Names(thickName).RefersTo
            If refString = "=FAUX" Or refString = "=FALSE" Then
                ' Ne rien faire
            Else
                refString = Replace(refString, "=", "")
                Dim addresses As Variant
                addresses = Split(refString, ",")
                Dim addr As Variant
                For Each addr In addresses
                    Dim cell As Range
                    On Error Resume Next
                    Set cell = ws.Range(addr)
                    If Not cell Is Nothing Then
                        cell.Value = Round(4.4 + Rnd() * (7.6 - 4.4), 2)
                    End If
                    On Error GoTo 0
                Next addr
            End If
        End If
    Next i
    Call FormatRollLayout
End Sub

' Efface toutes les cellules d'�paisseur officielles du rouleau actif
' @but : R�initialiser les mesures d'�paisseur saisies (hors rattrapage)
' @param Aucun
' @return Aucun
' @pr� : Les plages nomm�es leftThicknessCels et rightThicknessCels doivent exister et PRODUCTION_WS doit �tre initialis�
Public Sub ClearThicknessCells()
    Dim ws As Worksheet
    Set ws = PRODUCTION_WS
    Dim thickNames As Variant
    thickNames = Array("leftThicknessCels", "rightThicknessCels")
    Dim i As Integer, thickName As String, refString As String
    For i = LBound(thickNames) To UBound(thickNames)
        thickName = thickNames(i)
        If NameExists(thickName) Then
            refString = ThisWorkbook.Names(thickName).RefersTo
            If refString = "=FAUX" Or refString = "=FALSE" Then
                ' Ne rien faire
            Else
                refString = Replace(refString, "=", "")
                Dim addresses As Variant
                addresses = Split(refString, ",")
                Dim addr As Variant
                For Each addr In addresses
                    Dim cell As Range
                    On Error Resume Next
                    Set cell = ws.Range(addr)
                    If Not cell Is Nothing Then
                        cell.Value = ""
                    End If
                    On Error GoTo 0
                Next addr
            End If
        End If
    Next i
    Call FormatRollLayout
    Call RewriteActiveRollLengths
End Sub


' Vide toutes les cellules de la zone active du rouleau (activeRollArea)
' @but : R�initialiser compl�tement la zone de saisie du rouleau
' @param Aucun
' @return Aucun
' @pr� : La plage nomm�e activeRollArea doit exister et PRODUCTION_WS doit �tre initialis�
Public Sub ClearAllActiveRollArea()
    Dim ws As Worksheet: Set ws = PRODUCTION_WS
    If ws Is Nothing Then Exit Sub

    Dim wasProtected As Boolean: wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect

    Dim rngActive As Range
    Set rngActive = ws.Range(ThisWorkbook.Names("activeRollArea").RefersTo)
    rngActive.Value = ""

    If wasProtected Then ws.Protect

    Call FormatRollLayout
    Call RewriteActiveRollLengths
End Sub

' R��crit les m�trages (1, 2, 3, ...) dans les colonnes de longueur de la zone active
' @but : Restaurer les valeurs de m�trage apr�s un vidage complet de la zone active
' @param Aucun
' @return Aucun
' @pr� : Les plages nomm�es lengthCols et activeRollArea doivent exister et PRODUCTION_WS doit �tre initialis�
Public Sub RewriteActiveRollLengths()
    Dim ws As Worksheet: Set ws = PRODUCTION_WS
    If ws Is Nothing Then Exit Sub

    Dim wasProtected As Boolean: wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect

    Dim rngActive As Range
    Set rngActive = ws.Range(ThisWorkbook.Names("activeRollArea").RefersTo)

    Dim rngLength As Range
    Set rngLength = Application.Intersect(ws.Range(ThisWorkbook.Names("lengthCols").RefersTo), rngActive)
    If rngLength Is Nothing Then Exit Sub

    ' D�verrouille les cellules de longueur si besoin
    rngLength.Locked = False

    Dim firstRow As Long: firstRow = rngActive.Rows(1).Row
    Dim cell As Range
    For Each cell In rngLength.Cells
        cell.Value = cell.Row - firstRow + 1
    Next cell

    If wasProtected Then ws.Protect

    Call FormatRollLayout
    
End Sub 


Public Sub SaveShiftFromButton()
    Dim s As Shift
    Set s = New Shift
    s.LoadFromSheet PRODUCTION_WS
    s.AppendToDataShifts Worksheets("dataShifts")
    ' MsgBox "Le shift a bien �t� ajout� � l'onglet dataShifts !", vbInformation
End Sub 