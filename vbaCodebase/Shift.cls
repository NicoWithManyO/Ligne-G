VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Shift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Propriétés Shift
' ID : Identifiant unique du poste
Public ID As String
' DatePoste : Date du poste
Public DatePoste As Date
' Operateur : Nom de l'opérateur
Public Operateur As String
' Vaccation : Type de vacation (matin, après-midi, nuit)
Public Vaccation As String
' Duree : Durée du poste en heures
Public Duree As Double
' MachinePrisePoste : Numéro de la machine au début du poste
Public MachinePrisePoste As String
' LgEnrouleePrisePoste : Longueur enroulée au début du poste
Public LgEnrouleePrisePoste As Double
' MachineFinPoste : Numéro de la machine à la fin du poste
Public MachineFinPoste As String
' LgEnrouleeFinPoste : Longueur enroulée à la fin du poste
Public LgEnrouleeFinPoste As Double
' Commentaires : Commentaires additionnels sur le poste
Public Commentaires As String
' SaveDateTime : Date et heure de la sauvegarde
Public SaveDateTime As Date
' Contrôles globaux
Public ctrlMoyenneMicG As Double
Public ctrlMoyenneMicD As Double
Public ctrlMicG1 As Double
Public ctrlMicG2 As Double
Public ctrlMicG3 As Double
Public ctrlMicD1 As Double
Public ctrlMicD2 As Double
Public ctrlMicD3 As Double
Public ctrlMasseSurfaciqueGG As Double
Public ctrlMasseSurfaciqueGC As Double
Public ctrlMasseSurfaciqueDC As Double
Public ctrlMasseSurfaciqueDD As Double
Public ctrlBain As Double
Public ctrlLoi As Double
Public ctrlSavedOnRoll As Double

Public lgOK As Double
Public lgNOK As Double
Public lgScrap As Double
Public possibleLg As Double
' TRS : Taux de Rendement Synthétique (% de lgOK par rapport à possibleLg)
Public TRS As Double

' Charge les données du shift depuis la feuille
' @but : Initialise les propriétés du shift avec les valeurs de la feuille
' @param ws : Feuille contenant les données du shift
' @return : aucun
Public Sub LoadFromSheet(ws As Worksheet)
    ' Charger les données de base
    ID = ws.Range("shiftID").Value
    DatePoste = ws.Range("shiftDate").Value
    Operateur = ws.Range("shiftOperateur").Value
    Vaccation = ws.Range("shiftVaccation").Value
    Duree = ws.Range("shiftDuree").Value
    MachinePrisePoste = ws.Range("shiftMachinePrisePoste").Value
    LgEnrouleePrisePoste = ws.Range("shiftLgEnrouleePrisePoste").Value
    MachineFinPoste = ws.Range("shiftMachineFinPoste").Value
    LgEnrouleeFinPoste = ws.Range("shiftLgEnrouleeFinPoste").Value

    ' Charger les longueurs spécifiques depuis les adresses fixes
    Call LoadLengthsFromSheet(ws)

    ' Calcul du TRS (Taux de Rendement Synthétique)
    Call CalculateTRS
End Sub

' Sauvegarde les données du shift dans la feuille
' @but : écrit les propriétés du shift dans la feuille
' @param ws : Feuille où sauvegarder les données
' @return : aucun
Public Sub SaveToSheet(ws As Worksheet)
    ' Sauvegarder les données de base
    ws.Range("shiftID").Value = ID
    ws.Range("shiftDate").Value = DatePoste
    ws.Range("shiftOperateur").Value = Operateur
    ws.Range("shiftVaccation").Value = Vaccation
    ws.Range("shiftDuree").Value = Duree
    ws.Range("shiftMachinePrisePoste").Value = MachinePrisePoste
    ws.Range("shiftLgEnrouleePrisePoste").Value = LgEnrouleePrisePoste
    ws.Range("shiftMachineFinPoste").Value = MachineFinPoste
    ws.Range("shiftLgEnrouleeFinPoste").Value = LgEnrouleeFinPoste

    ' Répartit les commentaires sur les lignes
    Call SplitCommentsToRows(ws)
End Sub

' Répartit les lignes
Private Sub SplitCommentsToRows(ws As Worksheet)
    Dim comments As String
    comments = ws.Range("shiftComments").Value
    If comments = "" Then Exit Sub

    Dim lines() As String
    lines = Split(comments, vbCrLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        ws.Cells(i + 1, 1).Value = lines(i)
    Next i
End Sub

' Ajoute les données du shift à la fin de la feuille data_shifts
' @but : Enregistre un nouveau shift dans l'historique
' @param ws : Feuille data_shifts où ajouter les données
' @return : aucun
Public Sub AppendToDataShifts(ws As Worksheet)
    ' Vérification de l'état de la machine et des temps de démarrage
    If MachinePrisePoste = "A l'Arrêt" Then
        Dim hasStartupTime As Boolean
        hasStartupTime = False
        Dim cell As Range
        For Each cell In PRODUCTION_WS.Range("AC78:AC100")
            If cell.Value = "Démarrage" Then
                hasStartupTime = True
                Exit For
            End If
        Next cell
        
        If Not hasStartupTime Then
            If MsgBox("La machine était à l'arrêt à la prise de poste, mais aucun temps de démarrage déclaré, continuer quand même ?", vbQuestion + vbYesNo, "Avertissement") = vbNo Then
                Exit Sub
            End If
        End If
    End If

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
    
    ' Vérification des contrôles globaux obligatoires
    If Not AreAllGlobalsCtrlFilled(True) Then
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
    ' If MachinePrisePoste = "Démarrée" Then
    ws.Cells(nextRow, 7).Value = LgEnrouleePrisePoste
    ' Else
    '     ws.Cells(nextRow, 10).Value = ""
    ' End If
    ws.Cells(nextRow, 8).Value = MachineFinPoste
    ' Longueur fin de poste uniquement si machine démarrée
    ' If MachineFinPoste = "Démarrée" Then
    ws.Cells(nextRow, 9).Value = LgEnrouleeFinPoste
    ' Else
    '     ws.Cells(nextRow, 9).Value = ""
    ' End If
    
    ws.Cells(nextRow, 10).Value = Commentaires
    ws.Cells(nextRow, 11).Value = Now ' Ajout de la date/heure de sauvegarde

    ' Remplir les contrôles globaux
    ctrlMoyenneMicG = ""
    ctrlMoyenneMicD = ""
    Dim micG1 As Double, micG2 As Double, micG3 As Double
    Dim micD1 As Double, micD2 As Double, micD3 As Double
    micG1 = ThisWorkbook.Names("micG1").RefersToRange.Value
    micG2 = ThisWorkbook.Names("micG2").RefersToRange.Value
    micG3 = ThisWorkbook.Names("micG3").RefersToRange.Value
    micD1 = ThisWorkbook.Names("micD1").RefersToRange.Value
    micD2 = ThisWorkbook.Names("micD2").RefersToRange.Value
    micD3 = ThisWorkbook.Names("micD3").RefersToRange.Value
    If IsNumeric(micG1) And IsNumeric(micG2) And IsNumeric(micG3) Then
        ctrlMoyenneMicG = Round((CDbl(micG1) + CDbl(micG2) + CDbl(micG3)) / 3, 2)
    End If
    If IsNumeric(micD1) And IsNumeric(micD2) And IsNumeric(micD3) Then
        ctrlMoyenneMicD = Round((CDbl(micD1) + CDbl(micD2) + CDbl(micD3)) / 3, 2)
    End If
    ctrlMicG1 = micG1
    ctrlMicG2 = micG2
    ctrlMicG3 = micG3
    ctrlMicD1 = micD1
    ctrlMicD2 = micD2
    ctrlMicD3 = micD3
    ctrlMasseSurfaciqueGG = ThisWorkbook.Names("masseSurfaciqueGG").RefersToRange.Value
    ctrlMasseSurfaciqueGC = ThisWorkbook.Names("masseSurfaciqueGC").RefersToRange.Value
    ctrlMasseSurfaciqueDC = ThisWorkbook.Names("masseSurfaciqueDC").RefersToRange.Value
    ctrlMasseSurfaciqueDD = ThisWorkbook.Names("masseSurfaciqueDD").RefersToRange.Value
    ctrlBain = ThisWorkbook.Names("bain").RefersToRange.Value
    ctrlLoi = ThisWorkbook.Names("loi").RefersToRange.Value
    Dim at59Val As Double
    at59Val = PRODUCTION_WS.Range("AT59").Value
    If at59Val = "" Or IsEmpty(at59Val) Then
        ctrlSavedOnRoll = ""
    Else
        ctrlSavedOnRoll = at59Val
    End If

    ' Vider la plage des commentaires sur PROD
    PRODUCTION_WS.Range(RANGE_SHIFT_COMMENTAIRES).Value = ""

    ' Déclaration unique de colStart
    Dim colStart As Long
    ' Gestion automatique des en-têtes pour les colonnes ctrl*
    If ws.Cells(1, 1).Value = "" Then
        ws.Cells(1, 1).Value = "ID"
        ws.Cells(1, 2).Value = "DatePoste"
        ws.Cells(1, 3).Value = "Operateur"
        ws.Cells(1, 4).Value = "Vaccation"
        ws.Cells(1, 5).Value = "Duree"
        ws.Cells(1, 6).Value = "MachinePrisePoste"
        ws.Cells(1, 7).Value = "LgEnrouleePrisePoste"
        ws.Cells(1, 8).Value = "MachineFinPoste"
        ws.Cells(1, 9).Value = "LgEnrouleeFinPoste"
        ws.Cells(1, 10).Value = "Commentaires"
        ws.Cells(1, 11).Value = "SaveDateTime"
        colStart = 12
        ws.Cells(1, colStart).Value = "ctrlMoyenneMicG"
        ws.Cells(1, colStart + 1).Value = "ctrlMoyenneMicD"
        ws.Cells(1, colStart + 2).Value = "ctrlMicG1"
        ws.Cells(1, colStart + 3).Value = "ctrlMicG2"
        ws.Cells(1, colStart + 4).Value = "ctrlMicG3"
        ws.Cells(1, colStart + 5).Value = "ctrlMicD1"
        ws.Cells(1, colStart + 6).Value = "ctrlMicD2"
        ws.Cells(1, colStart + 7).Value = "ctrlMicD3"
        ws.Cells(1, colStart + 8).Value = "ctrlMasseSurfaciqueGG"
        ws.Cells(1, colStart + 9).Value = "ctrlMasseSurfaciqueGC"
        ws.Cells(1, colStart + 10).Value = "ctrlMasseSurfaciqueDC"
        ws.Cells(1, colStart + 11).Value = "ctrlMasseSurfaciqueDD"
        ws.Cells(1, colStart + 12).Value = "ctrlBain"
        ws.Cells(1, colStart + 13).Value = "ctrlLoi"
        ws.Cells(1, colStart + 14).Value = "ctrlSavedOnRoll"
        colStart = colStart + 15
        ws.Cells(1, colStart).Value = "lgOK"
        ws.Cells(1, colStart + 1).Value = "lgNOK"
        ws.Cells(1, colStart + 2).Value = "lgScrap"
        ws.Cells(1, colStart + 3).Value = "possibleLg"
        ws.Cells(1, colStart + 4).Value = "TRS"
    End If
    ' Export dans la feuille à la fin de la ligne
    colStart = ws.Cells(nextRow, ws.Columns.Count).End(xlToLeft).Column + 1
    ws.Cells(nextRow, colStart).Value = ctrlMoyenneMicG
    ws.Cells(nextRow, colStart + 1).Value = ctrlMoyenneMicD
    ws.Cells(nextRow, colStart + 2).Value = ctrlMicG1
    ws.Cells(nextRow, colStart + 3).Value = ctrlMicG2
    ws.Cells(nextRow, colStart + 4).Value = ctrlMicG3
    ws.Cells(nextRow, colStart + 5).Value = ctrlMicD1
    ws.Cells(nextRow, colStart + 6).Value = ctrlMicD2
    ws.Cells(nextRow, colStart + 7).Value = ctrlMicD3
    ws.Cells(nextRow, colStart + 8).Value = ctrlMasseSurfaciqueGG
    ws.Cells(nextRow, colStart + 9).Value = ctrlMasseSurfaciqueGC
    ws.Cells(nextRow, colStart + 10).Value = ctrlMasseSurfaciqueDC
    ws.Cells(nextRow, colStart + 11).Value = ctrlMasseSurfaciqueDD
    ws.Cells(nextRow, colStart + 12).Value = ctrlBain
    ws.Cells(nextRow, colStart + 13).Value = ctrlLoi
    ws.Cells(nextRow, colStart + 14).Value = ctrlSavedOnRoll
    ws.Cells(nextRow, colStart + 15).Value = lgOK
    ws.Cells(nextRow, colStart + 16).Value = lgNOK
    ws.Cells(nextRow, colStart + 17).Value = lgScrap
    ws.Cells(nextRow, colStart + 18).Value = possibleLg
    ws.Cells(nextRow, colStart + 19).Value = TRS

    Call ClearGlobalsCtrlValues
    Call ClearTimeLostRange
    call ShiftsRotate
    call ShiftMachineRotate
    PRODUCTION_WS.Range("BG77").Value = "isFirst"
    PRODUCTION_WS.Range("AB49").Value = ""
    PRODUCTION_WS.Range("AD50").Value = "" 
    MsgBox "Le poste a été enregistré avec succès.", vbInformation, "Enregistrement réussi"
End Sub
