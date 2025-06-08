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
    
    ' Déprotéger la feuille au début si besoin
    Dim wasProtected As Boolean
    wasProtected = PRODUCTION_WS.ProtectContents
    If wasProtected Then PRODUCTION_WS.Unprotect
    
    ' Création d'une instance de Roll
    Dim myRoll As New Roll
    
    ' Chargement des données depuis PROD
    myRoll.LoadFromSheet PRODUCTION_WS
    
    ' Vérifier que BH80, BH81 et BH82 sont renseignées
    If PRODUCTION_WS.Range("BH80").Value = "" Or PRODUCTION_WS.Range("BH81").Value = "" Or PRODUCTION_WS.Range(RANGE_REAL_LENGTH).Value = "" Then
        MsgBox "Merci de renseigner les champs : Masse du tube, Masse totale et Longueur avant de sauvegarder.", vbExclamation
        GoTo SafeExit
    End If
    
    ' Vérifier que toutes les épaisseurs sont présentes
    Dim missingMeasurements As String
    Dim rollLength As Double
    rollLength = PRODUCTION_WS.Range(RANGE_REAL_LENGTH).Value
    If Not IsNumeric(rollLength) Or rollLength <= 0 Then
        rollLength = PRODUCTION_WS.Range(TARGET_LENGTH_ADDR).Value
    End If
    If Not AreAllThicknessesPresent(missingMeasurements) Then
        MsgBox "Merci de renseigner toutes les épaisseurs requises pour un rouleau de " & rollLength & "m avant de sauvegarder :" & vbCrLf & missingMeasurements, vbExclamation
        GoTo SafeExit
    End If
    
    ' Vérifier que les infos de poste (shift) sont renseignées
    Dim missingShiftFields As String
    missingShiftFields = ""
    If PRODUCTION_WS.Range("shiftDate").Value = "" Then missingShiftFields = missingShiftFields & "- Date du poste" & vbCrLf
    If PRODUCTION_WS.Range("shiftOperateur").Value = "" Then missingShiftFields = missingShiftFields & "- Opérateur" & vbCrLf
    If PRODUCTION_WS.Range("shiftVaccation").Value = "" Then missingShiftFields = missingShiftFields & "- Vacation" & vbCrLf
    If PRODUCTION_WS.Range("shiftID").Value = "" Then missingShiftFields = missingShiftFields & "- ID du poste" & vbCrLf
    If PRODUCTION_WS.Range("shiftMachinePrisePoste").Value = "" Then missingShiftFields = missingShiftFields & "- Machine prise de poste" & vbCrLf
    If PRODUCTION_WS.Range("shiftDuree").Value = "" Then missingShiftFields = missingShiftFields & "- Durée du poste" & vbCrLf
    If missingShiftFields <> "" Then
        MsgBox "Merci de renseigner les informations de poste suivantes avant de sauvegarder :" & vbCrLf & missingShiftFields, vbExclamation
        GoTo SafeExit
    End If
    
    ' Vérifier si la longueur du rouleau est différente de la longueur cible
    Dim cible As String
    If IsNumeric(PRODUCTION_WS.Range(TARGET_LENGTH_ADDR).Value) And PRODUCTION_WS.Range(TARGET_LENGTH_ADDR).Value <> "" Then
        cible = PRODUCTION_WS.Range(TARGET_LENGTH_ADDR).Value & "m"
    Else
        cible = "non renseignée"
    End If
    If myRoll.Length <> PRODUCTION_WS.Range(TARGET_LENGTH_ADDR).Value Then
        If Not MODE_PERMISSIF Then
            MsgBox "La longueur du rouleau (" & myRoll.Length & "m) est différente de la longueur cible (" & cible & ")." & vbCrLf & _
                   "La sauvegarde est refusée car le mode permissif n'est pas activé.", vbExclamation, "Différence de longueur"
            Debug.Print "[saveRollFromProd] Export refusé : longueur différente et mode permissif désactivé."
            GoTo SafeExit
        Else
            Dim lengthDiffMsg As String
            lengthDiffMsg = "La longueur du rouleau (" & myRoll.Length & "m) est différente de la longueur cible (" & cible & ")." & vbCrLf & _
                            "Voulez-vous tout de même sauvegarder ce rouleau ?"
            If MsgBox(lengthDiffMsg, vbYesNo + vbQuestion, "Différence de longueur") <> vbYes Then
                Debug.Print "[saveRollFromProd] Export annulé par l'utilisateur (différence de longueur)."
                GoTo SafeExit
            End If
        End If
    End If
    
    ' Demander confirmation avant sauvegarde
    Dim confirmMsg As String
    confirmMsg = "Confirmer la sauvegarde du rouleau :" & vbCrLf & _
                 "ID : " & myRoll.ID & vbCrLf & _
                 "Longueur : " & myRoll.Length & "m." & vbCrLf & _
                 "Statut : " & myRoll.Status
    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Confirmation export rouleau") <> vbYes Then
        Debug.Print "[saveRollFromProd] Export annulé par l'utilisateur."
        GoTo SafeExit
    End If
    
    ' Vérifier si l'ID existe déjà
    Dim wsDataRolls As Worksheet
    Set wsDataRolls = ThisWorkbook.Sheets("dataRolls")
    Dim lastRow As Long
    lastRow = wsDataRolls.Cells(wsDataRolls.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow ' Commencer à 2 pour ignorer l'en-tête
        If wsDataRolls.Cells(i, 1).Value = myRoll.ID Then
            MsgBox "Un rouleau avec l'ID " & myRoll.ID & " existe déjà.", vbExclamation
            GoTo SafeExit
        End If
    Next i
    
    ' --- Gestion conditionnelle des contrôles globaux dans le Roll ---
    Dim at59Val As Variant
    at59Val = PRODUCTION_WS.Range("AT59").Value
    If UCase(Trim(at59Val)) = "" Then
        ' Calcul des moyennes MicG et MicD
        Dim micG1 As Variant, micG2 As Variant, micG3 As Variant
        Dim micD1 As Variant, micD2 As Variant, micD3 As Variant
        micG1 = ThisWorkbook.Names("micG1").RefersToRange.Value
        micG2 = ThisWorkbook.Names("micG2").RefersToRange.Value
        micG3 = ThisWorkbook.Names("micG3").RefersToRange.Value
        micD1 = ThisWorkbook.Names("micD1").RefersToRange.Value
        micD2 = ThisWorkbook.Names("micD2").RefersToRange.Value
        micD3 = ThisWorkbook.Names("micD3").RefersToRange.Value
        
        If IsNumeric(micG1) And IsNumeric(micG2) And IsNumeric(micG3) Then
            myRoll.MicG = Round((CDbl(micG1) + CDbl(micG2) + CDbl(micG3)) / 3, 2)
        Else
            myRoll.MicG = ""
        End If
        If IsNumeric(micD1) And IsNumeric(micD2) And IsNumeric(micD3) Then
            myRoll.MicD = Round((CDbl(micD1) + CDbl(micD2) + CDbl(micD3)) / 3, 2)
        Else
            myRoll.MicD = ""
        End If
        ' Masse surfacique G et D
        Dim masseGG As Variant, masseDD As Variant
        masseGG = ThisWorkbook.Names("masseSurfaciqueGG").RefersToRange.Value
        masseDD = ThisWorkbook.Names("masseSurfaciqueDD").RefersToRange.Value
        If IsNumeric(masseGG) Then
            myRoll.MasseSurfaciqueG = masseGG
        Else
            myRoll.MasseSurfaciqueG = ""
        End If
        If IsNumeric(masseDD) Then
            myRoll.MasseSurfaciqueD = masseDD
        Else
            myRoll.MasseSurfaciqueD = ""
        End If
        ' Ensimage
        myRoll.Ensimage = PRODUCTION_WS.Range("bain").Value
        PRODUCTION_WS.Range("AT59").Value = myRoll.ID
    End If
    ' --- Fin gestion conditionnelle des contrôles globaux ---
    
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
        ' Mettre AQ59 à OK uniquement si conforme
        ' If PRODUCTION_WS.ProtectContents Then PRODUCTION_WS.Unprotect
        ' PRODUCTION_WS.Range("AQ59").Value = "OK"
    Else
        Debug.Print "[saveRollFromProd] Status non conforme : " & myRoll.Status & " - Pas d'incrémentation"
    End If

    ' Gestion des poids
    ' Copier BK82 vers BH80 si BK82 a une valeur, sinon vider BH80
    If Not IsEmpty(PRODUCTION_WS.Range("BK82").Value) And PRODUCTION_WS.Range("BK82").Value <> "" Then
        PRODUCTION_WS.Range("BH80").Value = PRODUCTION_WS.Range("BK82").Value
    Else
        PRODUCTION_WS.Range("BH80").Value = ""
    End If
    PRODUCTION_WS.Range("BK82").Value = ""

    ' Vider BH81
    PRODUCTION_WS.Range("BH81").Value = ""

    ' Vider la zone active et réécrire les mètres
    Call ClearAllActiveRollArea
    Call ExportGlobalsCtrlToSheet
    ' Message de confirmation
    MsgBox "Le rouleau " & myRoll.ID & " a bien été sauvegardé : " & myRoll.Status, vbInformation

SafeExit:
    ' Reprotéger la feuille si besoin
    If wasProtected Then PRODUCTION_WS.Protect
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