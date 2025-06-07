Attribute VB_Name = "modGlobalsCtrl"
Option Explicit

' @but : Vérifie la conformité de tous les contrôles globaux par rapport à leurs min/max
' @param motif (ByRef) : motif détaillé en cas de non-conformité
' @return Boolean (True si conforme, False sinon)
' @pré : Les plages nommées doivent être initialisées
Public Function IsGlobalsCtrlConform(Optional ByRef motif As String = "") As Boolean
    Dim isConform As Boolean: isConform = True
    motif = ""
    ' Micronnaire
    Dim i As Integer, val As Variant
    For i = 1 To 3
        val = ThisWorkbook.Names("micG" & i).RefersToRange.Value
        If val = "" Or Not IsNumeric(val) Then
            isConform = False
            motif = motif & "micG" & i & " non renseigné ou non numérique | "
        ElseIf val < ThisWorkbook.Names("micronnaireMin").RefersToRange.Value Or val > ThisWorkbook.Names("micronnaireMax").RefersToRange.Value Then
            isConform = False
            motif = motif & "micG" & i & " hors tolérance | "
        End If
        val = ThisWorkbook.Names("micD" & i).RefersToRange.Value
        If val = "" Or Not IsNumeric(val) Then
            isConform = False
            motif = motif & "micD" & i & " non renseigné ou non numérique | "
        ElseIf val < ThisWorkbook.Names("micronnaireMin").RefersToRange.Value Or val > ThisWorkbook.Names("micronnaireMax").RefersToRange.Value Then
            isConform = False
            motif = motif & "micD" & i & " hors tolérance | "
        End If
    Next i
    ' Bain
    val = ThisWorkbook.Names("bain").RefersToRange.Value
    If val = "" Or Not IsNumeric(val) Then
        isConform = False
        motif = motif & "Bain non renseigné ou non numérique | "
    ElseIf val < ThisWorkbook.Names("bainMin").RefersToRange.Value Or val > ThisWorkbook.Names("bainMax").RefersToRange.Value Then
        isConform = False
        motif = motif & "Bain hors tolérance | "
    End If
    ' Masse surfacique
    Dim masseNames As Variant: masseNames = Array("masseSurfaciqueGG", "masseSurfaciqueGC", "masseSurfaciqueDC", "masseSurfaciqueDD")
    Dim j As Integer
    For j = 0 To 3
        val = ThisWorkbook.Names(masseNames(j)).RefersToRange.Value
        If val = "" Or Not IsNumeric(val) Then
            isConform = False
            motif = motif & masseNames(j) & " non renseigné ou non numérique | "
        ElseIf val < ThisWorkbook.Names("masseSurfMin").RefersToRange.Value Or val > ThisWorkbook.Names("masseSurfMax").RefersToRange.Value Then
            isConform = False
            motif = motif & masseNames(j) & " hors tolérance | "
        End If
    Next j
    ' LOI
    If UCase(Trim(ThisWorkbook.Names("loi").RefersToRange.Value)) <> "OK" Then
        isConform = False
        motif = motif & "LOI non conforme | "
    End If
    IsGlobalsCtrlConform = isConform
End Function

Public Sub TestGlobalsCtrlConform()
    Dim motif As String
    If IsGlobalsCtrlConform(motif) Then
        MsgBox "Contrôles globaux CONFORMES !", vbInformation
    Else
        MsgBox "NON CONFORME : " & motif, vbExclamation
    End If
End Sub

Public Sub SetLOI_OK()
    ' Met la valeur OK dans la cellule LOI (nommée) pour valider le contrôle
    Dim rng As Range
    Set rng = ThisWorkbook.Names("loi").RefersToRange
    If UCase(Trim(rng.Value)) = "OK" Then
        If MsgBox("L'échantillon LOI a déjà été donné, êtes-vous sûr ?", vbYesNo + vbQuestion) <> vbYes Then
            Exit Sub
        End If
    End If
    rng.Value = "OK"
End Sub

Public Sub ExportGlobalsCtrlToSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("dataGlbCtrls")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "La feuille 'dataGlbCtrls' n'existe pas.", vbCritical
        Exit Sub
    End If

    Vérifier si les contrôles ont déjà été sauvegardés (plage fusionnée)
    On teste la première cellule de la plage fusionnée pour éviter l'incompatibilité de type
    If PRODUCTION_WS.Range("AR60:AV60").Cells(1, 1).Value = "Contrôles Sauvegardés" Then
        ' MsgBox "Ces contrôles ont déjà été sauvegardés", vbExclamation
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim isEmpty As Boolean: isEmpty = (ws.Cells(1, 1).Value = "")

    ' Calcul des moyennes
    Dim micG1 As Variant, micG2 As Variant, micG3 As Variant
    Dim micD1 As Variant, micD2 As Variant, micD3 As Variant
    micG1 = ThisWorkbook.Names("micG1").RefersToRange.Value
    micG2 = ThisWorkbook.Names("micG2").RefersToRange.Value
    micG3 = ThisWorkbook.Names("micG3").RefersToRange.Value
    micD1 = ThisWorkbook.Names("micD1").RefersToRange.Value
    micD2 = ThisWorkbook.Names("micD2").RefersToRange.Value
    micD3 = ThisWorkbook.Names("micD3").RefersToRange.Value
    Dim moyenneMicG As Variant, moyenneMicD As Variant
    If IsNumeric(micG1) And IsNumeric(micG2) And IsNumeric(micG3) Then
        moyenneMicG = Round((CDbl(micG1) + CDbl(micG2) + CDbl(micG3)) / 3, 2)
    Else
        moyenneMicG = ""
    End If
    If IsNumeric(micD1) And IsNumeric(micD2) And IsNumeric(micD3) Then
        moyenneMicD = Round((CDbl(micD1) + CDbl(micD2) + CDbl(micD3)) / 3, 2)
    Else
        moyenneMicD = ""
    End If

    ' Écrire les en-têtes si la feuille est vide
    If isEmpty Then
        ws.Cells(1, 1).Value = "globalsCtrlID"
        ws.Cells(1, 2).Value = "shiftID"
        ws.Cells(1, 3).Value = "moyenneMicG"
        ws.Cells(1, 4).Value = "moyenneMicD"
        ws.Cells(1, 5).Value = "micG1"
        ws.Cells(1, 6).Value = "micG2"
        ws.Cells(1, 7).Value = "micG3"
        ws.Cells(1, 8).Value = "micD1"
        ws.Cells(1, 9).Value = "micD2"
        ws.Cells(1, 10).Value = "micD3"
        ws.Cells(1, 11).Value = "masseSurfaciqueGG"
        ws.Cells(1, 12).Value = "masseSurfaciqueGC"
        ws.Cells(1, 13).Value = "masseSurfaciqueDC"
        ws.Cells(1, 14).Value = "masseSurfaciqueDD"
        ws.Cells(1, 15).Value = "bain"
        ws.Cells(1, 16).Value = "loi"
        ws.Cells(1, 17).Value = "productRollID"
        ws.Cells(1, 18).Value = "saveDateTime"
        lastRow = 1
    End If

    Dim nextRow As Long: nextRow = lastRow + 1
    ws.Cells(nextRow, 1).Value = ThisWorkbook.Names("globalsCtrlID").RefersToRange.Value
    ws.Cells(nextRow, 2).Value = ThisWorkbook.Names("shiftID").RefersToRange.Value
    ws.Cells(nextRow, 3).Value = moyenneMicG
    ws.Cells(nextRow, 4).Value = moyenneMicD
    ws.Cells(nextRow, 5).Value = micG1
    ws.Cells(nextRow, 6).Value = micG2
    ws.Cells(nextRow, 7).Value = micG3
    ws.Cells(nextRow, 8).Value = micD1
    ws.Cells(nextRow, 9).Value = micD2
    ws.Cells(nextRow, 10).Value = micD3
    ws.Cells(nextRow, 11).Value = ThisWorkbook.Names("masseSurfaciqueGG").RefersToRange.Value
    ws.Cells(nextRow, 12).Value = ThisWorkbook.Names("masseSurfaciqueGC").RefersToRange.Value
    ws.Cells(nextRow, 13).Value = ThisWorkbook.Names("masseSurfaciqueDC").RefersToRange.Value
    ws.Cells(nextRow, 14).Value = ThisWorkbook.Names("masseSurfaciqueDD").RefersToRange.Value
    ws.Cells(nextRow, 15).Value = ThisWorkbook.Names("bain").RefersToRange.Value
    ws.Cells(nextRow, 16).Value = ThisWorkbook.Names("loi").RefersToRange.Value
    ws.Cells(nextRow, 17).Value = ThisWorkbook.Names("productRollID").RefersToRange.Value
    ws.Cells(nextRow, 18).Value = Now

    Call SetGlobalsCtrlSaved
End Sub

Public Sub ClearGlobalsCtrlValues()
    Dim ctrlNames As Variant
    ctrlNames = Array("micG1", "micG2", "micG3", "micD1", "micD2", "micD3", _
        "masseSurfaciqueGG", "masseSurfaciqueGC", "masseSurfaciqueDC", "masseSurfaciqueDD", _
        "bain", "loi")
    Dim i As Integer
    For i = LBound(ctrlNames) To UBound(ctrlNames)
        ThisWorkbook.Names(ctrlNames(i)).RefersToRange.Value = ""
    Next i
    Call ResetGlobalsCtrlSaved
End Sub

Public Sub SetGlobalsCtrlSaved()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PROD")
    Dim wasProtected As Boolean: wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect
    ws.Range("AR60:AV60").Value = "Contrôles Sauvegardés"
    ws.Range("AU59").Value = ws.Range("AU59").Value + 1
    ' ws.Range("AR60:AV60").Interior.Color = RGB(0, 176, 80) ' Vert Excel
    ' ws.Range("AK60:BC60").Interior.Color = RGB(0, 176, 80)
    If wasProtected Then ws.Protect
End Sub

Public Sub ResetGlobalsCtrlSaved()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PROD")
    Dim wasProtected As Boolean: wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect 
    ws.Range("AR60:AV60").Value = ""
    ' ws.Range("AR60:AV60").Interior.Color = RGB(77, 147, 217) ' Bleu #4D93D9
    ' ws.Range("AK60:BC60").Interior.Color = RGB(77, 147, 217)
    If wasProtected Then ws.Protect
End Sub