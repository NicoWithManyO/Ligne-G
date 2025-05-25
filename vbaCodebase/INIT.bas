Attribute VB_Name = "INIT"

Option Explicit

' Initialise tous les composants n�cessaires apr�s le chargement des modules
' @pre : les modules doivent �tre charg�s
' @return : aucun
Public Sub initializeComponents()
    Debug.Print "[initializeComponents] D�but de l'initialisation"
    
    ' Initialisation de la feuille de production
    Set PRODUCTION_WS = ThisWorkbook.Sheets("PROD")
    If PRODUCTION_WS Is Nothing Then
        Debug.Print "[initializeComponents] ERREUR : Feuille PROD non trouv�e"
        Exit Sub
    End If


    ' Initialisation des ranges
    Call initShiftRanges
    Call defineRollNamedRanges
    Call FormatRollLayout
    
    Debug.Print "[initializeComponents] Initialisation termin�e"

End Sub

Public Sub SetTargetLength(ws As Worksheet, targetLength As Double)
    Application.EnableEvents = False
    ws.Unprotect
    ws.Range(TARGET_LENGTH_ADDR).Value = targetLength
    ws.Range(TARGET_LENGTH_ADDR).Locked = True
    ws.Protect
    Debug.Print "[SetTargetLength] Nouvelle longueur cible = " & targetLength
    Call initializeComponents
    Application.EnableEvents = True
End Sub

