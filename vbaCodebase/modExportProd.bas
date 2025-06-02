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

' Sauvegarde les données du rouleau dans un fichier texte
' @but : Sauvegarde toutes les données du rouleau dans un fichier texte formaté
' @param filePath : Chemin du fichier où sauvegarder les données
' @return : aucun
Public Sub SaveRollToFile(filePath As String)
    If PRODUCTION_WS Is Nothing Then
        Debug.Print "[SaveRollToFile] ERREUR : PRODUCTION_WS non initialisé"
        Exit Sub
    End If
    
    ' Création d'une instance de Roll
    Dim myRoll As New Roll
    
    ' Chargement des données depuis PROD
    myRoll.LoadFromSheet PRODUCTION_WS
    
    ' Création du fichier
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file As Object: Set file = fso.CreateTextFile(filePath, True)
    
    ' Écriture des données de base
    file.WriteLine "=== Données du rouleau ==="
    file.WriteLine "ID: " & myRoll.ID
    file.WriteLine "OF: " & myRoll.OF
    file.WriteLine "Numéro: " & myRoll.Number
    file.WriteLine "Poste: " & myRoll.FabricationShift
    file.WriteLine "Opérateur: " & myRoll.FabricationOperator
    file.WriteLine "OF en cours: " & myRoll.OFInProgress
    file.WriteLine "Longueur: " & myRoll.Length
    file.WriteLine "Poids tube: " & myRoll.PipeWeight
    file.WriteLine "Poids total: " & myRoll.TotalWeight
    file.WriteLine "Poids: " & myRoll.Weight
    file.WriteLine "Statut: " & myRoll.Status
    file.WriteLine "Défauts: " & myRoll.Defects
    
    ' Écriture des épaisseurs
    file.WriteLine ""
    file.WriteLine "=== Épaisseurs ==="
    
    ' Parcourir toutes les épaisseurs par position
    Dim positions As Variant: positions = Array("Gauche", "Droite")
    Dim pos As Variant
    
    For Each pos In positions
        file.WriteLine pos & ":"
        Dim thickness As Object
        For Each thickness In myRoll.Thicknesses(pos)
            Dim line As String
            line = "  " & thickness("rowOffset") & "m: " & Format(thickness("value"), "0.00")
            If thickness.Exists("rattrapageValue") Then
                line = line & " | " & Format(thickness("rattrapageValue"), "0.00")
            End If
            file.WriteLine line
        Next thickness
    Next pos
    
    ' Fermeture du fichier
    file.Close
    
    Debug.Print "[SaveRollToFile] Rouleau sauvegardé dans : " & filePath
End Sub

' Lit les données de la feuille de production et sauvegarde le rouleau
' @but : Lit les données du rouleau depuis la feuille de production et les sauvegarde
' @return : aucun
Public Sub ReadAndSaveRoll()
    If PRODUCTION_WS Is Nothing Then
        Debug.Print "[ReadAndSaveRoll] ERREUR : PRODUCTION_WS non initialisé"
        Exit Sub
    End If
    
    ' Création d'une instance de Roll
    Dim myRoll As New Roll
    
    ' Les données sont automatiquement chargées et sauvegardées dans le constructeur
    ' mais on peut ajouter des vérifications supplémentaires ici si nécessaire
    
    Debug.Print "[ReadAndSaveRoll] Rouleau lu et sauvegardé : " & myRoll.ID
End Sub 