Attribute VB_Name = "INIT"

Option Explicit

' Initialise tous les composants n√©cessaires apr√®s le chargement des modules
' @pre : les modules doivent √™tre charg√©s
' @return : aucun
Public Sub initializeComponents()
    Debug.Print "[initializeComponents] D√©but de l'initialisation"
    MODE_PERMISSIF = True
    
    ' Initialisation de la feuille de production
    Set PRODUCTION_WS = ThisWorkbook.Sheets("PROD")
    If PRODUCTION_WS Is Nothing Then
        Debug.Print "[initializeComponents] ERREUR : Feuille PROD non trouv√©e"
        Exit Sub
    End If

    Call ReadModePermissifFromSheet
    ' Initialisation des ranges
    Call initShiftRanges
    Call initOFRanges
    Call initProductRollRanges
    Call defineRollNamedRanges

    Call FormatRollLayout
    Call initCtrlLimitValues
    
    Debug.Print "[initializeComponents] Initialisation termin√©e"

    Call IsRollConformDefects
    Call saveDetectedDefects

    Call IsRollConformThickness
    Call saveDetectedThickness

    ' Ajout : rÈÈcriture des mÈtrages ‡ chaque init
    Call RewriteActiveRollLengths
End Sub
