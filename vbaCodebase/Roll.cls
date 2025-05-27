VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Roll"
Option Explicit

' Propriétés Roll
' ID : Identifiant unique du rouleau
Public ID As Variant
' OF : Numéro d'ordre de fabrication
Public OF As Variant
' Number : Numéro du rouleau
Public Number As Variant
' FabricationShift : Poste de fabrication
Public FabricationShift As Variant
' FabricationOperator : Opérateur de fabrication
Public FabricationOperator As Variant
' OFInProgress : OF en cours
Public OFInProgress As Variant
' TargetLength : Longueur cible
Public TargetLength As Variant
' Destination : Destination du rouleau (Prod, Decoupe, Dechet)
Public Destination As Variant
' PipeWeight : Masse du tube
Public PipeWeight As Variant
' TotalWeight : Masse totale
Public TotalWeight As Variant
' Weight : Masse
Public Weight As Variant
' Length : Longueur
Public Length As Variant
' Status : État du rouleau
Public Status As Variant
' SaveDateTime : Date et heure de la sauvegarde
Public SaveDateTime As Variant
' Defects : Défauts détectés
Public Defects As Variant


' Charge les données depuis la feuille
' @but : Charge les données du rouleau depuis la feuille de production
' @param ws : Feuille de travail source
' @return : aucun
Public Sub LoadFromSheet(ws As Worksheet)
    ' Chargement des propriétés de base
    Dim idValue As Variant
    idValue = ws.Range(RANGE_PRODUCTROLL_ID).Value
    If IsEmpty(idValue) Or idValue = "" Then
        ID = ws.Range(RANGE_CUT_OF_NUMBER).Value & "-" & Format(Now, "ddmmyy-hhmm")
    Else
        ID = idValue
    End If
    OF = ws.Range(RANGE_PRODUCTROLL_OF).Value
    Number = ws.Range(RANGE_PRODUCTROLL_NUMBER).Value
    FabricationShift = ws.Range(RANGE_PRODUCTROLL_FABRICATION_SHIFT).Value
    FabricationOperator = ws.Range(RANGE_SHIFT_OPERATEUR).Value
    OFInProgress = ws.Range(RANGE_OF_NUMBER).Value
    Status = ws.Range(RANGE_PRODUCTROLL_STATUS).Value
    PipeWeight = ws.Range(RANGE_PRODUCTROLL_PIPE_WEIGHT).Value
    TotalWeight = ws.Range(RANGE_PRODUCTROLL_TOTAL_WEIGHT).Value
    Length = ws.Range(RANGE_PRODUCTROLL_LENGTH).Value
    Defects = ws.Range(RANGE_PRODUCTROLL_DEFECTS).Value
    
    ' Calcul du poids (TotalWeight - PipeWeight)
    If Not IsEmpty(TotalWeight) And Not IsEmpty(PipeWeight) Then
        Weight = TotalWeight - PipeWeight
    End If
End Sub

' Sauvegarde les données dans la feuille
' @but : Sauvegarde les données du rouleau dans la feuille de production
' @param ws : Feuille de travail cible
' @return : aucun
Public Sub SaveToSheet(ws As Worksheet)
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = ID
    ws.Cells(nextRow, 2).Value = OF
    ws.Cells(nextRow, 3).Value = Number
    ws.Cells(nextRow, 4).Value = FabricationShift
    ws.Cells(nextRow, 5).Value = Status
    ws.Cells(nextRow, 6).Value = PipeWeight
    ws.Cells(nextRow, 7).Value = TotalWeight
    ws.Cells(nextRow, 8).Value = Weight
    ws.Cells(nextRow, 9).Value = Length
    ws.Cells(nextRow, 10).Value = Defects
    ws.Cells(nextRow, 11).Value = Now()
End Sub
