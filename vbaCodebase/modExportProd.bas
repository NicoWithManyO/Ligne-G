Attribute VB_Name = "modExportProd"
Option Explicit

' Sauvegarde un rouleau depuis la feuille PROD vers dataSave
' @but : Sauvegarde les données d'un rouleau depuis la feuille de production vers dataSave
' @return : aucun
Public Sub saveRollFromProd()
    If PRODUCTION_WS Is Nothing Then
        Debug.Print "[saveRollFromProd] ERREUR : PRODUCTION_WS non initialisé"
        Exit Sub
    End If
    
    ' Création d'une instance de Roll
    Dim myRoll As New Roll
    
    ' Chargement des données depuis PROD
    myRoll.LoadFromSheet PRODUCTION_WS
    
    ' Sauvegarde dans dataRolls
    myRoll.SaveToSheet ThisWorkbook.Sheets("dataRolls")
    
    Debug.Print "[saveRollFromProd] Rouleau sauvegardé : " & myRoll.ID
End Sub 