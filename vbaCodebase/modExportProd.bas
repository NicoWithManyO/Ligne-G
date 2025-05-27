Attribute VB_Name = "modExportProd"
Option Explicit

' Sauvegarde un rouleau depuis la feuille PROD vers dataSave
' @but : Sauvegarde les donn�es d'un rouleau depuis la feuille de production vers dataSave
' @return : aucun
Public Sub saveRollFromProd()
    If PRODUCTION_WS Is Nothing Then
        Debug.Print "[saveRollFromProd] ERREUR : PRODUCTION_WS non initialis�"
        Exit Sub
    End If
    
    ' Cr�ation d'une instance de Roll
    Dim myRoll As New Roll
    
    ' Chargement des donn�es depuis PROD
    myRoll.LoadFromSheet PRODUCTION_WS
    
    ' Sauvegarde dans dataRolls
    myRoll.SaveToSheet ThisWorkbook.Sheets("dataRolls")
    
    Debug.Print "[saveRollFromProd] Rouleau sauvegard� : " & myRoll.ID
End Sub 