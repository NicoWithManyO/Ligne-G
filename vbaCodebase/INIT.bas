Attribute VB_Name = "INIT"

Option Explicit

' Initialise tous les composants nécessaires après le chargement des modules
' @pre : les modules doivent être chargés
' @return : aucun
Public Sub initializeComponents()
    Debug.Print "[initializeComponents] Début de l'initialisation"
    MODE_PERMISSIF = True
    
    ' Initialisation de la feuille de production
    Set PRODUCTION_WS = ThisWorkbook.Sheets("PROD")
    If PRODUCTION_WS Is Nothing Then
        Debug.Print "[initializeComponents] ERREUR : Feuille PROD non trouvée"
        Exit Sub
    End If

    Call ReadModePermissifFromSheet
    ' Initialisation des ranges
    Call initShiftRanges
    Call initOFRanges
    Call initProductRollRanges
    Call defineRollNamedRanges
    
    Call initGlobalsCtrlRanges

    Call FormatRollLayout
    Call initCtrlLimitValues
    
    Debug.Print "[initializeComponents] Initialisation terminée"

    ' Call IsRollConformDefects
    ' Call saveDetectedDefects

    ' Call IsRollConformThickness
    ' Call saveDetectedThickness

    ' Ajout : réécriture des métrages à chaque init
    Call RewriteActiveRollLengths
End Sub
