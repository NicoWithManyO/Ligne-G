Attribute VB_Name = "modTRS"
Option Explicit

' Copie les donn�es des cellules sp�cifi�es vers la premi�re ligne vide � partir de AC78
' @but : Copier les donn�es de AM47:AO47, AQ47:AW47 et BB47:BC47 vers la premi�re ligne vide � partir de AC78
' @return : aucun
Public Sub CopyDataToFirstEmptyRow()
    Dim ws As Worksheet
    Set ws = PRODUCTION_WS
    If ws Is Nothing Then Exit Sub
    
    ' Demander confirmation
    Dim rep As VbMsgBoxResult
    rep = MsgBox("Voulez-vous r�ellement d�clarer " & ws.Range("BB47").Value & "min en temps perdu (" & ws.Range("AM47").Value & ") ?", vbYesNo + vbQuestion, "Confirmation")
    If rep <> vbYes Then Exit Sub
    
    ' D�prot�ger la feuille si n�cessaire
    Dim wasProtected As Boolean
    wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect
    
    ' Trouver la premi�re ligne vide � partir de AC79
    Dim firstEmptyRow As Long
    firstEmptyRow = 78
    Do While ws.Range("AC" & firstEmptyRow).Value <> ""
        firstEmptyRow = firstEmptyRow + 1
    Loop
    
    ' Copier les donn�es
    ' AM47:AO47 -> AC[firstEmptyRow]
    ws.Range("AC" & firstEmptyRow).Value = ws.Range("AM47").Value
    
    ' AQ47:AW47 -> AD[firstEmptyRow]:AF[firstEmptyRow] (fusionn�es)
    ws.Range("AD" & firstEmptyRow & ":AF" & firstEmptyRow).Merge
    ws.Range("AD" & firstEmptyRow).Value = ws.Range("AQ47").Value
    
    ' BB47:BC47 -> AG[firstEmptyRow]
    ws.Range("AG" & firstEmptyRow).Value = ws.Range("BB47").Value
    
    ' Vider les cellules d'origine (sauf BB47:BC47)
    ws.Range("AM47:AO47").Value = ""
    ws.Range("AQ47:AW47").Value = ""
    ws.Range("AY47").Value = ""
    ws.Range("BA47").Value = ""
    
    ' Reproter la feuille si elle �tait prot�g�e
    If wasProtected Then ws.Protect
    
    MsgBox "Temps perdu d�clar� avec succ�s.", vbInformation
End Sub

' Vide la plage de cellules AC78:AG99
' @but : Effacer toutes les donn�es dans la plage AC78:AG99
' @return : aucun
Public Sub ClearTimeLostRange()
    Dim ws As Worksheet
    Set ws = PRODUCTION_WS
    If ws Is Nothing Then Exit Sub
    
    ' Demander confirmation
    Dim rep As VbMsgBoxResult
    rep = MsgBox("Voulez-vous r�ellement supprimer tout temps perdu d�clar� ?", vbYesNo + vbQuestion, "Confirmation")
    If rep <> vbYes Then Exit Sub
    
    ' D�prot�ger la feuille si n�cessaire
    Dim wasProtected As Boolean
    wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect
    
    ' Vider la plage
    ws.Range("AC78:AG99").Value = ""
    
    ' Reproter la feuille si elle �tait prot�g�e
    If wasProtected Then ws.Protect
    
    MsgBox "La plage de temps perdu a �t� vid�e.", vbInformation
End Sub


