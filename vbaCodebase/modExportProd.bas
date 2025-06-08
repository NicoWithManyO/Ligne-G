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
    
    ' D�prot�ger la feuille au d�but si besoin
    Dim wasProtected As Boolean
    wasProtected = PRODUCTION_WS.ProtectContents
    If wasProtected Then PRODUCTION_WS.Unprotect
    
    ' Cr�ation d'une instance de Roll
    Dim myRoll As New Roll
    
    ' Chargement des donn�es depuis PROD
    myRoll.LoadFromSheet PRODUCTION_WS
    
    ' V�rifier que BH80, BH81 et BH82 sont renseign�es
    If PRODUCTION_WS.Range("BH80").Value = "" Or PRODUCTION_WS.Range("BH81").Value = "" Or PRODUCTION_WS.Range(RANGE_REAL_LENGTH).Value = "" Then
        MsgBox "Merci de renseigner les champs : Masse du tube, Masse totale et Longueur avant de sauvegarder.", vbExclamation
        GoTo SafeExit
    End If
    
    ' V�rifier que toutes les �paisseurs sont pr�sentes
    Dim missingMeasurements As String
    Dim rollLength As Double
    rollLength = PRODUCTION_WS.Range(RANGE_REAL_LENGTH).Value
    If Not IsNumeric(rollLength) Or rollLength <= 0 Then
        rollLength = PRODUCTION_WS.Range(TARGET_LENGTH_ADDR).Value
    End If
    If Not AreAllThicknessesPresent(missingMeasurements) Then
        MsgBox "Merci de renseigner toutes les �paisseurs requises pour un rouleau de " & rollLength & "m avant de sauvegarder :" & vbCrLf & missingMeasurements, vbExclamation
        GoTo SafeExit
    End If
    
    ' V�rifier que les infos de poste (shift) sont renseign�es
    Dim missingShiftFields As String
    missingShiftFields = ""
    If PRODUCTION_WS.Range("shiftDate").Value = "" Then missingShiftFields = missingShiftFields & "- Date du poste" & vbCrLf
    If PRODUCTION_WS.Range("shiftOperateur").Value = "" Then missingShiftFields = missingShiftFields & "- Op�rateur" & vbCrLf
    If PRODUCTION_WS.Range("shiftVaccation").Value = "" Then missingShiftFields = missingShiftFields & "- Vacation" & vbCrLf
    If PRODUCTION_WS.Range("shiftID").Value = "" Then missingShiftFields = missingShiftFields & "- ID du poste" & vbCrLf
    If PRODUCTION_WS.Range("shiftMachinePrisePoste").Value = "" Then missingShiftFields = missingShiftFields & "- Machine prise de poste" & vbCrLf
    If PRODUCTION_WS.Range("shiftDuree").Value = "" Then missingShiftFields = missingShiftFields & "- Dur�e du poste" & vbCrLf
    If missingShiftFields <> "" Then
        MsgBox "Merci de renseigner les informations de poste suivantes avant de sauvegarder :" & vbCrLf & missingShiftFields, vbExclamation
        GoTo SafeExit
    End If
    
    ' V�rifier si la longueur du rouleau est diff�rente de la longueur cible
    Dim cible As String
    If IsNumeric(PRODUCTION_WS.Range(TARGET_LENGTH_ADDR).Value) And PRODUCTION_WS.Range(TARGET_LENGTH_ADDR).Value <> "" Then
        cible = PRODUCTION_WS.Range(TARGET_LENGTH_ADDR).Value & "m"
    Else
        cible = "non renseign�e"
    End If
    If myRoll.Length <> PRODUCTION_WS.Range(TARGET_LENGTH_ADDR).Value Then
        If Not MODE_PERMISSIF Then
            MsgBox "La longueur du rouleau (" & myRoll.Length & "m) est diff�rente de la longueur cible (" & cible & ")." & vbCrLf & _
                   "La sauvegarde est refus�e car le mode permissif n'est pas activ�.", vbExclamation, "Diff�rence de longueur"
            Debug.Print "[saveRollFromProd] Export refus� : longueur diff�rente et mode permissif d�sactiv�."
            GoTo SafeExit
        Else
            Dim lengthDiffMsg As String
            lengthDiffMsg = "La longueur du rouleau (" & myRoll.Length & "m) est diff�rente de la longueur cible (" & cible & ")." & vbCrLf & _
                            "Voulez-vous tout de m�me sauvegarder ce rouleau ?"
            If MsgBox(lengthDiffMsg, vbYesNo + vbQuestion, "Diff�rence de longueur") <> vbYes Then
                Debug.Print "[saveRollFromProd] Export annul� par l'utilisateur (diff�rence de longueur)."
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
        Debug.Print "[saveRollFromProd] Export annul� par l'utilisateur."
        GoTo SafeExit
    End If
    
    ' V�rifier si l'ID existe d�j�
    Dim wsDataRolls As Worksheet
    Set wsDataRolls = ThisWorkbook.Sheets("dataRolls")
    Dim lastRow As Long
    lastRow = wsDataRolls.Cells(wsDataRolls.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow ' Commencer � 2 pour ignorer l'en-t�te
        If wsDataRolls.Cells(i, 1).Value = myRoll.ID Then
            MsgBox "Un rouleau avec l'ID " & myRoll.ID & " existe d�j�.", vbExclamation
            GoTo SafeExit
        End If
    Next i
    
    ' --- Gestion conditionnelle des contr�les globaux dans le Roll ---
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
    ' --- Fin gestion conditionnelle des contr�les globaux ---
    
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
        ' Mettre AQ59 � OK uniquement si conforme
        ' If PRODUCTION_WS.ProtectContents Then PRODUCTION_WS.Unprotect
        ' PRODUCTION_WS.Range("AQ59").Value = "OK"
    Else
        Debug.Print "[saveRollFromProd] Status non conforme : " & myRoll.Status & " - Pas d'incr�mentation"
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

    ' Vider la zone active et r��crire les m�tres
    Call ClearAllActiveRollArea
    Call ExportGlobalsCtrlToSheet
    ' Message de confirmation
    MsgBox "Le rouleau " & myRoll.ID & " a bien �t� sauvegard� : " & myRoll.Status, vbInformation

SafeExit:
    ' Reprot�ger la feuille si besoin
    If wasProtected Then PRODUCTION_WS.Protect
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