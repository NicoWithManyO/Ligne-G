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
    Call modRangesInit.initShiftRanges
    Call modRangesInit.defineRollNamedRanges
    Call FormatRollLayout
    Debug.Print "[initializeComponents] Initialisation terminée"

End Sub

