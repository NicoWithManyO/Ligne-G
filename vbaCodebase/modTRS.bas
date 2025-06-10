Attribute VB_Name = "modTRS"
Option Explicit

' Copie les données des cellules spécifiées vers la première ligne vide à partir de AC78
' @but : Copier les données de AM47:AO47, AQ47:AW47 et BB47:BC47 vers la première ligne vide à partir de AC78
' @return : aucun
Public Sub CopyDataToFirstEmptyRow()
    Dim ws As Worksheet
    Set ws = PRODUCTION_WS
    If ws Is Nothing Then Exit Sub
    
    ' Demander confirmation
    Dim rep As VbMsgBoxResult
    rep = MsgBox("Voulez-vous réellement déclarer " & ws.Range("BB47").Value & "min en temps perdu (" & ws.Range("AM47").Value & ") ?", vbYesNo + vbQuestion, "Confirmation")
    If rep <> vbYes Then Exit Sub
    
    ' Déprotéger la feuille si nécessaire
    Dim wasProtected As Boolean
    wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect
    
    ' Trouver la première ligne vide à partir de AC79
    Dim firstEmptyRow As Long
    firstEmptyRow = 78
    Do While ws.Range("AC" & firstEmptyRow).Value <> ""
        firstEmptyRow = firstEmptyRow + 1
    Loop
    
    ' Copier les données
    ' AM47:AO47 -> AC[firstEmptyRow]
    ws.Range("AC" & firstEmptyRow).Value = ws.Range("AM47").Value
    
    ' AQ47:AW47 -> AD[firstEmptyRow]:AF[firstEmptyRow] (fusionnées)
    ws.Range("AD" & firstEmptyRow & ":AF" & firstEmptyRow).Merge
    ws.Range("AD" & firstEmptyRow).Value = ws.Range("AQ47").Value
    
    ' BB47:BC47 -> AG[firstEmptyRow]
    ws.Range("AG" & firstEmptyRow).Value = ws.Range("BB47").Value
    
    ' Vider les cellules d'origine (sauf BB47:BC47)
    ws.Range("AM47:AO47").Value = ""
    ws.Range("AQ47:AW47").Value = ""
    ws.Range("AY47").Value = ""
    ws.Range("BA47").Value = ""
    
    ' Reproter la feuille si elle était protégée
    If wasProtected Then ws.Protect
    
    MsgBox "Temps perdu déclaré avec succès.", vbInformation
End Sub

' Vide la plage de cellules AC78:AG99
' @but : Effacer toutes les données dans la plage AC78:AG99
' @return : aucun
Public Sub ClearTimeLostRange()
    Dim ws As Worksheet
    Set ws = PRODUCTION_WS
    If ws Is Nothing Then Exit Sub
    
    ' Demander confirmation
    Dim rep As VbMsgBoxResult
    rep = MsgBox("Voulez-vous réellement supprimer tout temps perdu déclaré ?", vbYesNo + vbQuestion, "Confirmation")
    If rep <> vbYes Then Exit Sub
    
    ' Déprotéger la feuille si nécessaire
    Dim wasProtected As Boolean
    wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect
    
    ' Vider la plage
    ws.Range("AC78:AG99").Value = ""
    
    ' Reproter la feuille si elle était protégée
    If wasProtected Then ws.Protect
    
    MsgBox "La plage de temps perdu a été vidée.", vbInformation
End Sub


