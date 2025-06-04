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
    
    ' Vérifier si l'ID existe déjà
    Dim wsDataRolls As Worksheet
    Set wsDataRolls = ThisWorkbook.Sheets("dataRolls")
    Dim lastRow As Long
    lastRow = wsDataRolls.Cells(wsDataRolls.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow ' Commencer à 2 pour ignorer l'en-tête
        If wsDataRolls.Cells(i, 1).Value = myRoll.ID Then
            MsgBox "Un rouleau avec l'ID " & myRoll.ID & " existe déjà.", vbExclamation
            Exit Sub
        End If
    Next i
    
    ' Sauvegarde dans dataRolls
    myRoll.SaveToSheet wsDataRolls
    
    Debug.Print "[saveRollFromProd] Rouleau sauvegardé : " & myRoll.ID
    Debug.Print "[saveRollFromProd] Status du rouleau : " & myRoll.Status

    ' Si le status est conforme, on incrémente le numéro de roll
    If UCase(myRoll.Status) = "CONFORME" Then
        Dim currentRollNumber As Long
        currentRollNumber = PRODUCTION_WS.Range(RANGE_PRODUCTROLL_NUMBER).Value
        Debug.Print "[saveRollFromProd] Numéro de roll actuel : " & currentRollNumber
        Call SetRollNumber(PRODUCTION_WS, currentRollNumber + 1)
        Debug.Print "[saveRollFromProd] Numéro de roll incrémenté : " & (currentRollNumber + 1)
    Else
        Debug.Print "[saveRollFromProd] Status non conforme : " & myRoll.Status & " - Pas d'incrémentation"
    End If

    ' Gestion des poids
    Dim wasProtected As Boolean: wasProtected = PRODUCTION_WS.ProtectContents
    If wasProtected Then PRODUCTION_WS.Unprotect

    ' Copier BK82 vers BH80 si BH80 est vide
    If IsEmpty(PRODUCTION_WS.Range("BH80").Value) Then
        Dim bk82Value As Variant
        bk82Value = PRODUCTION_WS.Range("BK82").Value
        If Not IsEmpty(bk82Value) Then
            PRODUCTION_WS.Range("BH80").Value = bk82Value
        End If
    End If

    ' Vider BH81 et BK82
    PRODUCTION_WS.Range("BH81").Value = ""
    PRODUCTION_WS.Range("BK82").Value = ""

    If wasProtected Then PRODUCTION_WS.Protect

    ' Vider la zone active et réécrire les mètres
    Call ClearAllActiveRollArea

    ' Message de confirmation
    MsgBox "Le rouleau " & myRoll.ID & " a bien été sauvegardé : " & myRoll.Status, vbInformation
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