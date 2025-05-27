VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Shift"
Option Explicit

' Propriétés Shift
' ID : Identifiant unique du poste
Public ID As Variant
' DatePoste : Date du poste
Public DatePoste As Variant
' Operateur : Nom de l'opérateur
Public Operateur As Variant
' Vaccation : Type de vacation (matin, après-midi, nuit)
Public Vaccation As Variant
' Duree : Durée du poste en heures
Public Duree As Variant
' MachinePrisePoste : Numéro de la machine au début du poste
Public MachinePrisePoste As Variant
' LgEnrouleePrisePoste : Longueur enroulée au début du poste
Public LgEnrouleePrisePoste As Variant
' MachineFinPoste : Numéro de la machine à la fin du poste
Public MachineFinPoste As Variant
' LgEnrouleeFinPoste : Longueur enroulée à la fin du poste
Public LgEnrouleeFinPoste As Variant
' Commentaires : Commentaires additionnels sur le poste
Public Commentaires As Variant
' SaveDateTime : Date et heure de la sauvegarde
Public SaveDateTime As Variant

' Charge les données du shift depuis la feuille
' @but : Initialise les propriétés du shift avec les valeurs de la feuille
' @param ws : Feuille contenant les données du shift
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
    
    ' Concatène les lignes des commentaires
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

' Sauvegarde les données du shift dans la feuille
' @but : Écrit les propriétés du shift dans la feuille
' @param ws : Feuille où sauvegarder les données
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
    
    ' Répartit les commentaires sur les lignes
    Dim commentRange As Range
    Dim commentLines() As String
    Dim i As Long
    
    Set commentRange = ws.Range(RANGE_SHIFT_COMMENTAIRES)
    commentLines = Split(Commentaires, " | ")
    
    ' Efface d'abord la plage
    commentRange.ClearContents
    
    ' Répartit les lignes
    For i = 0 To UBound(commentLines)
        If i <= commentRange.Rows.Count - 1 Then
            commentRange.Cells(i + 1, 1).Value = commentLines(i)
        End If
    Next i
    
    ws.Range(RANGE_SHIFT_SAVE_DATE_TIME).Value = SaveDateTime
End Sub

' Ajoute les données du shift à la fin de la feuille data_shifts
' @but : Enregistre un nouveau shift dans l'historique
' @param ws : Feuille data_shifts où ajouter les données
' @return : aucun
Public Sub AppendToDataShifts(ws As Worksheet)
    Dim lastRow As Long
    Dim nextRow As Long
    Dim i As Long
    Dim shiftExists As Boolean
    
    ' Vérifie si le shift existe déjà
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    shiftExists = False
    
    For i = 2 To lastRow  ' Commence à 2 pour sauter l'en-tête
        If ws.Cells(i, 1).Value = ID Then
            shiftExists = True
            MsgBox "Un poste avec l'ID " & ID & " existe déjà ...", vbExclamation, "Poste déjà existant"
            Exit Sub
        End If
    Next i
    
    ' Vérifie tous les champs obligatoires
    If ID = "" Or IsEmpty(ID) Then
        MsgBox "L'ID du poste est obligatoire.", vbExclamation, "Données manquantes"
        Exit Sub
    End If
    
    If DatePoste = "" Or IsEmpty(DatePoste) Then
        MsgBox "La date du poste est obligatoire.", vbExclamation, "Données manquantes"
        Exit Sub
    End If
    
    If Operateur = "" Or IsEmpty(Operateur) Then
        MsgBox "L'opérateur est obligatoire.", vbExclamation, "Données manquantes"
        Exit Sub
    End If
    
    If Vaccation = "" Or IsEmpty(Vaccation) Then
        MsgBox "La vacation est obligatoire.", vbExclamation, "Données manquantes"
        Exit Sub
    End If
    
    If Duree = "" Or IsEmpty(Duree) Then
        MsgBox "La durée du poste est obligatoire.", vbExclamation, "Données manquantes"
        Exit Sub
    End If
    
    If MachinePrisePoste = "" Or IsEmpty(MachinePrisePoste) Then
        MsgBox "L'état de la machine en prise de poste est obligatoire.", vbExclamation, "Données manquantes"
        Exit Sub
    End If
    
    If MachineFinPoste = "" Or IsEmpty(MachineFinPoste) Then
        MsgBox "L'état de la machine en fin de poste est obligatoire.", vbExclamation, "Données manquantes"
        Exit Sub
    End If
    
    ' Vérifie les longueurs si les machines sont démarrées
    If MachinePrisePoste = "Démarrée" And (LgEnrouleePrisePoste = "" Or IsEmpty(LgEnrouleePrisePoste)) Then
        MsgBox "La machine est démarrée en prise de poste mais aucune longueur n'est saisie.", vbExclamation, "Données manquantes"
        Exit Sub
    End If
    
    If MachineFinPoste = "Démarrée" And (LgEnrouleeFinPoste = "" Or IsEmpty(LgEnrouleeFinPoste)) Then
        MsgBox "La machine est démarrée en fin de poste mais aucune longueur n'est saisie.", vbExclamation, "Données manquantes"
        Exit Sub
    End If
    
    ' Demande confirmation avant de sauvegarder
    If MsgBox("Voulez-vous enregistrer le poste de " & Operateur & " du " & DatePoste & " (" & Vaccation & ") ?", vbQuestion + vbYesNo, "Confirmation") = vbNo Then
        Exit Sub
    End If
    
    ' Si la feuille est vide, commence à la ligne 2 (ligne 1 pour les en-têtes)
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
    ' Longueur prise de poste uniquement si machine démarrée
    If MachinePrisePoste = "Démarrée" Then
        ws.Cells(nextRow, 7).Value = LgEnrouleePrisePoste
    Else
        ws.Cells(nextRow, 7).Value = ""
    End If
    ws.Cells(nextRow, 8).Value = MachineFinPoste
    ' Longueur fin de poste uniquement si machine démarrée
    If MachineFinPoste = "Démarrée" Then
        ws.Cells(nextRow, 9).Value = LgEnrouleeFinPoste
    Else
        ws.Cells(nextRow, 9).Value = ""
    End If
    ws.Cells(nextRow, 10).Value = Commentaires
    ws.Cells(nextRow, 11).Value = Now ' Ajout de la date/heure de sauvegarde
    
    MsgBox "Le poste a été enregistré avec succès.", vbInformation, "Enregistrement réussi"
End Sub
