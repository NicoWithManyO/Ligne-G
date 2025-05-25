Attribute VB_Name = "INIT"

Option Explicit

' Initialise tous les composants nécessaires après le chargement des modules
' @pre : les modules doivent être chargés
' @return : aucun
Public Sub initializeComponents()
    Debug.Print "[initializeComponents] Début de l'initialisation"
    
    ' Initialisation de la feuille de production
    Set PRODUCTION_WS = ThisWorkbook.Sheets("PROD")
    If PRODUCTION_WS Is Nothing Then
        Debug.Print "[initializeComponents] ERREUR : Feuille PROD non trouvée"
        Exit Sub
    End If


    ' Initialisation des ranges
    Call initShiftRanges
    Call initOFRanges
    Call defineRollNamedRanges
    Call FormatRollLayout
    Call initCtrlLimitValues
    
    Debug.Print "[initializeComponents] Initialisation terminée"

    Call IsRollConformDefects
    Call saveDetectedDefects

    Call IsRollConformThickness
    Call saveDetectedThickness
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

Public Sub SetOFNumber(ws As Worksheet, ofNumber As Long)
    Application.EnableEvents = False
    ws.Unprotect
    ws.Range(RANGE_OF_NUMBER).Value = ofNumber
    ws.Range(RANGE_OF_NUMBER).Locked = True
    ws.Protect
    Debug.Print "[SetOFNumber] Nouveau numéro OF = " & ofNumber
    Application.EnableEvents = True
End Sub

Public Sub SetCutOFNumber(ws As Worksheet, cutOfNumber As Long)
    Application.EnableEvents = False
    ws.Unprotect
    ws.Range(RANGE_CUT_OF_NUMBER).Value = cutOfNumber
    ws.Range(RANGE_CUT_OF_NUMBER).Locked = True
    ws.Protect
    Debug.Print "[SetCutOFNumber] Nouveau numéro OF de coupe = " & cutOfNumber
    Application.EnableEvents = True
End Sub

