VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Shift"
Option Explicit

' Propri�t�s Shift
' ID : Identifiant unique du poste
Public ID As Variant
' DatePoste : Date du poste
Public DatePoste As Variant
' Operateur : Nom de l'op�rateur
Public Operateur As Variant
' Vaccation : Type de vacation (matin, apr�s-midi, nuit)
Public Vaccation As Variant
' Duree : Dur�e du poste en heures
Public Duree As Variant
' MachinePrisePoste : Num�ro de la machine au d�but du poste
Public MachinePrisePoste As Variant
' LgEnrouleePrisePoste : Longueur enroul�e au d�but du poste
Public LgEnrouleePrisePoste As Variant
' MachineFinPoste : Num�ro de la machine � la fin du poste
Public MachineFinPoste As Variant
' LgEnrouleeFinPoste : Longueur enroul�e � la fin du poste
Public LgEnrouleeFinPoste As Variant
' Commentaires : Commentaires additionnels sur le poste
Public Commentaires As Variant
' SaveDateTime : Date et heure de la sauvegarde
Public SaveDateTime As Variant

' Charge les donn�es du shift depuis la feuille
' @but : Initialise les propri�t�s du shift avec les valeurs de la feuille
' @param ws : Feuille contenant les donn�es du shift
' @return : aucun
Public Sub LoadFromSheet(ws As Worksheet)
    ID = ws.Range(RANGE_SHIFT_ID).Value
    DatePoste = ws.Range(RANGE_SHIFT_DATE).Value
    Operateur = ws.Range(RANGE_SHIFT_OPERATEUR).Value
    Vaccation = ws.Range(RANGE_SHIFT_VACCATION).Value
    Duree = ws.Range(RANGE_SHIFT_DUREE).Value
    MachinePrisePoste = ws.Range(RANGE_SHIFT_MACHINE_PRISE_POSTE).Value
    LgEnrouleePrisePoste = ws.Range(RANGE_SHIFT_LG_ENROULEE_PRISE_POSTE).Value
    MachineFinPoste = ws.Range(RANGE_SHIFT_MACHINE_FIN_POSTE).Value
    LgEnrouleeFinPoste = ws.Range(RANGE_SHIFT_LG_ENROULEE_FIN_POSTE).Value
    
    ' Concat�ne les lignes des commentaires
    Dim commentRange As Range
    Dim commentCell As Range
    Dim commentText As String
    
    Set commentRange = ws.Range(RANGE_SHIFT_COMMENTAIRES)
    commentText = ""
    
    For Each commentCell In commentRange
        If Not IsEmpty(commentCell.Value) Then
            If commentText <> "" Then commentText = commentText & " | "
            commentText = commentText & commentCell.Value
        End If
    Next commentCell
    
    Commentaires = commentText
    SaveDateTime = Now
End Sub

' Sauvegarde les donn�es du shift dans la feuille
' @but : �crit les propri�t�s du shift dans la feuille
' @param ws : Feuille o� sauvegarder les donn�es
' @return : aucun
Public Sub SaveToSheet(ws As Worksheet)
    ws.Range(RANGE_SHIFT_ID).Value = ID
    ws.Range(RANGE_SHIFT_DATE).Value = DatePoste
    ws.Range(RANGE_SHIFT_OPERATEUR).Value = Operateur
    ws.Range(RANGE_SHIFT_VACCATION).Value = Vaccation
    ws.Range(RANGE_SHIFT_DUREE).Value = Duree
    ws.Range(RANGE_SHIFT_MACHINE_PRISE_POSTE).Value = MachinePrisePoste
    ws.Range(RANGE_SHIFT_LG_ENROULEE_PRISE_POSTE).Value = LgEnrouleePrisePoste
    ws.Range(RANGE_SHIFT_MACHINE_FIN_POSTE).Value = MachineFinPoste
    ws.Range(RANGE_SHIFT_LG_ENROULEE_FIN_POSTE).Value = LgEnrouleeFinPoste
    
    ' R�partit les commentaires sur les lignes
    Dim commentRange As Range
    Dim commentLines() As String
    Dim i As Long
    
    Set commentRange = ws.Range(RANGE_SHIFT_COMMENTAIRES)
    commentLines = Split(Commentaires, " | ")
    
    ' Efface d'abord la plage
    commentRange.ClearContents
    
    ' R�partit les lignes
    For i = 0 To UBound(commentLines)
        If i <= commentRange.Rows.Count - 1 Then
            commentRange.Cells(i + 1, 1).Value = commentLines(i)
        End If
    Next i
    
    ' ws.Range(RANGE_SHIFT_SAVE_DATE_TIME).Value = SaveDateTime
End Sub

' Ajoute les donn�es du shift � la fin de la feuille data_shifts
' @but : Enregistre un nouveau shift dans l'historique
' @param ws : Feuille data_shifts o� ajouter les donn�es
' @return : aucun
Public Sub AppendToDataShifts(ws As Worksheet)
    Dim lastRow As Long
    Dim nextRow As Long
    Dim i As Long
    Dim shiftExists As Boolean
    
    ' V�rifie si le shift existe d�j�
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    shiftExists = False
    
    For i = 2 To lastRow  ' Commence � 2 pour sauter l'en-t�te
        If ws.Cells(i, 1).Value = ID Then
            shiftExists = True
            MsgBox "Un poste avec l'ID " & ID & " existe d�j� ...", vbExclamation, "Poste d�j� existant"
            Exit Sub
        End If
    Next i
    
    ' V�rifie tous les champs obligatoires
    If ID = "" Or IsEmpty(ID) Then
        MsgBox "L'ID du poste est obligatoire.", vbExclamation, "Donn�es manquantes"
        Exit Sub
    End If
    
    If DatePoste = "" Or IsEmpty(DatePoste) Then
        MsgBox "La date du poste est obligatoire.", vbExclamation, "Donn�es manquantes"
        Exit Sub
    End If
    
    If Operateur = "" Or IsEmpty(Operateur) Then
        MsgBox "L'op�rateur est obligatoire.", vbExclamation, "Donn�es manquantes"
        Exit Sub
    End If
    
    If Vaccation = "" Or IsEmpty(Vaccation) Then
        MsgBox "La vacation est obligatoire.", vbExclamation, "Donn�es manquantes"
        Exit Sub
    End If
    
    If Duree = "" Or IsEmpty(Duree) Then
        MsgBox "La dur�e du poste est obligatoire.", vbExclamation, "Donn�es manquantes"
        Exit Sub
    End If
    
    If MachinePrisePoste = "" Or IsEmpty(MachinePrisePoste) Then
        MsgBox "L'�tat de la machine en prise de poste est obligatoire.", vbExclamation, "Donn�es manquantes"
        Exit Sub
    End If
    
    If MachineFinPoste = "" Or IsEmpty(MachineFinPoste) Then
        MsgBox "L'�tat de la machine en fin de poste est obligatoire.", vbExclamation, "Donn�es manquantes"
        Exit Sub
    End If
    
    ' V�rifie les longueurs si les machines sont d�marr�es
    If MachinePrisePoste = "D�marr�e" And (LgEnrouleePrisePoste = "" Or IsEmpty(LgEnrouleePrisePoste)) Then
        MsgBox "La machine est d�marr�e en prise de poste mais aucune longueur n'est saisie.", vbExclamation, "Donn�es manquantes"
        Exit Sub
    End If
    
    If MachineFinPoste = "D�marr�e" And (LgEnrouleeFinPoste = "" Or IsEmpty(LgEnrouleeFinPoste)) Then
        MsgBox "La machine est d�marr�e en fin de poste mais aucune longueur n'est saisie.", vbExclamation, "Donn�es manquantes"
        Exit Sub
    End If
    
    ' Demande confirmation avant de sauvegarder
    If MsgBox("Voulez-vous enregistrer le poste de " & Operateur & " du " & DatePoste & " (" & Vaccation & ") ?", vbQuestion + vbYesNo, "Confirmation") = vbNo Then
        Exit Sub
    End If
    
    ' Si la feuille est vide, commence � la ligne 2 (ligne 1 pour les en-t�tes)
    If lastRow = 1 And ws.Cells(1, 1).Value = "" Then
        nextRow = 2
    Else
        nextRow = lastRow + 1
    End If

    ws.Cells(nextRow, 1).Value = ID
    ws.Cells(nextRow, 2).Value = DatePoste
    ws.Cells(nextRow, 3).Value = Operateur
    ws.Cells(nextRow, 4).Value = Vaccation
    ws.Cells(nextRow, 5).Value = Duree
    ws.Cells(nextRow, 6).Value = MachinePrisePoste
    ' Longueur prise de poste uniquement si machine d�marr�e
    If MachinePrisePoste = "D�marr�e" Then
        ws.Cells(nextRow, 7).Value = LgEnrouleePrisePoste
    Else
        ws.Cells(nextRow, 7).Value = ""
    End If
    ws.Cells(nextRow, 8).Value = MachineFinPoste
    ' Longueur fin de poste uniquement si machine d�marr�e
    If MachineFinPoste = "D�marr�e" Then
        ws.Cells(nextRow, 9).Value = LgEnrouleeFinPoste
    Else
        ws.Cells(nextRow, 9).Value = ""
    End If
    ws.Cells(nextRow, 10).Value = Commentaires
    ws.Cells(nextRow, 11).Value = Now ' Ajout de la date/heure de sauvegarde
    
    MsgBox "Le poste a �t� enregistr� avec succ�s.", vbInformation, "Enregistrement r�ussi"
End Sub
