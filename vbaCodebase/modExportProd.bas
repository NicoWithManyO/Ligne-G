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
    
    ' V�rifier si l'ID existe d�j�
    Dim wsDataRolls As Worksheet
    Set wsDataRolls = ThisWorkbook.Sheets("dataRolls")
    Dim lastRow As Long
    lastRow = wsDataRolls.Cells(wsDataRolls.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow ' Commencer � 2 pour ignorer l'en-t�te
        If wsDataRolls.Cells(i, 1).Value = myRoll.ID Then
            MsgBox "Un rouleau avec l'ID " & myRoll.ID & " existe d�j�.", vbExclamation
            Exit Sub
        End If
    Next i
    
    ' Sauvegarde dans dataRolls
    myRoll.SaveToSheet wsDataRolls
    
    Debug.Print "[saveRollFromProd] Rouleau sauvegard� : " & myRoll.ID
    Debug.Print "[saveRollFromProd] Status du rouleau : " & myRoll.Status

    ' Si le status est conforme, on incr�mente le num�ro de roll
    If UCase(myRoll.Status) = "CONFORME" Then
        Dim currentRollNumber As Long
        currentRollNumber = PRODUCTION_WS.Range(RANGE_PRODUCTROLL_NUMBER).Value
        Debug.Print "[saveRollFromProd] Num�ro de roll actuel : " & currentRollNumber
        Call SetRollNumber(PRODUCTION_WS, currentRollNumber + 1)
        Debug.Print "[saveRollFromProd] Num�ro de roll incr�ment� : " & (currentRollNumber + 1)
    Else
        Debug.Print "[saveRollFromProd] Status non conforme : " & myRoll.Status & " - Pas d'incr�mentation"
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

    ' Vider la zone active et r��crire les m�tres
    Call ClearAllActiveRollArea

    ' Message de confirmation
    MsgBox "Le rouleau " & myRoll.ID & " a bien �t� sauvegard� : " & myRoll.Status, vbInformation
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