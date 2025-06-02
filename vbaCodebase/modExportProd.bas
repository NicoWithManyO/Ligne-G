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

' Sauvegarde les donn�es du rouleau dans un fichier texte
' @but : Sauvegarde toutes les donn�es du rouleau dans un fichier texte format�
' @param filePath : Chemin du fichier o� sauvegarder les donn�es
' @return : aucun
Public Sub SaveRollToFile(filePath As String)
    If PRODUCTION_WS Is Nothing Then
        Debug.Print "[SaveRollToFile] ERREUR : PRODUCTION_WS non initialis�"
        Exit Sub
    End If
    
    ' Cr�ation d'une instance de Roll
    Dim myRoll As New Roll
    
    ' Chargement des donn�es depuis PROD
    myRoll.LoadFromSheet PRODUCTION_WS
    
    ' Cr�ation du fichier
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file As Object: Set file = fso.CreateTextFile(filePath, True)
    
    ' �criture des donn�es de base
    file.WriteLine "=== Donn�es du rouleau ==="
    file.WriteLine "ID: " & myRoll.ID
    file.WriteLine "OF: " & myRoll.OF
    file.WriteLine "Num�ro: " & myRoll.Number
    file.WriteLine "Poste: " & myRoll.FabricationShift
    file.WriteLine "Op�rateur: " & myRoll.FabricationOperator
    file.WriteLine "OF en cours: " & myRoll.OFInProgress
    file.WriteLine "Longueur: " & myRoll.Length
    file.WriteLine "Poids tube: " & myRoll.PipeWeight
    file.WriteLine "Poids total: " & myRoll.TotalWeight
    file.WriteLine "Poids: " & myRoll.Weight
    file.WriteLine "Statut: " & myRoll.Status
    file.WriteLine "D�fauts: " & myRoll.Defects
    
    ' �criture des �paisseurs
    file.WriteLine ""
    file.WriteLine "=== �paisseurs ==="
    
    ' Parcourir toutes les �paisseurs par position
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
    
    Debug.Print "[SaveRollToFile] Rouleau sauvegard� dans : " & filePath
End Sub

' Lit les donn�es de la feuille de production et sauvegarde le rouleau
' @but : Lit les donn�es du rouleau depuis la feuille de production et les sauvegarde
' @return : aucun
Public Sub ReadAndSaveRoll()
    If PRODUCTION_WS Is Nothing Then
        Debug.Print "[ReadAndSaveRoll] ERREUR : PRODUCTION_WS non initialis�"
        Exit Sub
    End If
    
    ' Cr�ation d'une instance de Roll
    Dim myRoll As New Roll
    
    ' Les donn�es sont automatiquement charg�es et sauvegard�es dans le constructeur
    ' mais on peut ajouter des v�rifications suppl�mentaires ici si n�cessaire
    
    Debug.Print "[ReadAndSaveRoll] Rouleau lu et sauvegard� : " & myRoll.ID
End Sub 