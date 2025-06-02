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
' Thicknesses : Épaisseurs mesurées (3m et 8m)
Public Thicknesses As Object
' Mesures dynamiques (toujours présentes, vides si inutilisées)
Public Mesure1 As String
Public Mesure2 As String
Public Mesure3 As String
Public Mesure4 As String
Public Mesure5 As String
Public Mesure6 As String
Public Mesure7 As String
Public Mesure8 As String
Public Mesure9 As String
Public Mesure10 As String


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
    
    ' Chargement des mesures tous les 5m (jusqu'à 10 mesures)
    Dim thicknesses As Object
    Set thicknesses = LoadAllThicknesses()
    
    Dim i As Integer, m As Integer
    Dim mesureStr As String
    For m = 1 To 10
        mesureStr = GetMesureString(thicknesses, m)
        Select Case m
            Case 1: Me.Mesure1 = mesureStr
            Case 2: Me.Mesure2 = mesureStr
            Case 3: Me.Mesure3 = mesureStr
            Case 4: Me.Mesure4 = mesureStr
            Case 5: Me.Mesure5 = mesureStr
            Case 6: Me.Mesure6 = mesureStr
            Case 7: Me.Mesure7 = mesureStr
            Case 8: Me.Mesure8 = mesureStr
            Case 9: Me.Mesure9 = mesureStr
            Case 10: Me.Mesure10 = mesureStr
        End Select
    Next m
End Sub

' Génère la chaîne de mesure pour la m-ième mesure (tous les 5m)
Private Function GetMesureString(thicknesses As Object, mesureIndex As Integer) As String
    ' mesureIndex = 1 pour 5m, 2 pour 10m, ...
    Dim offset As Integer: offset = (mesureIndex - 1) * 5 + 1
    Dim gVals(1 To 3) As String, dVals(1 To 3) As String
    Dim gCount As Integer: gCount = 0
    Dim dCount As Integer: dCount = 0
    Dim t As Object
    Dim i As Integer
    For Each t In thicknesses("Gauche")
        If t("rowOffset") = offset Or t("rowOffset") = offset + 1 Or t("rowOffset") = offset + 2 Then
            gCount = gCount + 1
            gVals(gCount) = Format(t("value"), "0.00")
            If t.Exists("rattrapageValue") Then
                gVals(gCount) = gVals(gCount) & "/" & Format(t("rattrapageValue"), "0.00")
            End If
            If gCount = 3 Then Exit For
        End If
    Next t
    For Each t In thicknesses("Droite")
        If t("rowOffset") = offset Or t("rowOffset") = offset + 1 Or t("rowOffset") = offset + 2 Then
            dCount = dCount + 1
            dVals(dCount) = Format(t("value"), "0.00")
            If t.Exists("rattrapageValue") Then
                dVals(dCount) = dVals(dCount) & "/" & Format(t("rattrapageValue"), "0.00")
            End If
            If dCount = 3 Then Exit For
        End If
    Next t
    Dim result As String
    ' Trouver le rowOffset minimal parmi les 6 valeurs
    Dim minOffset As Integer: minOffset = 0
    Dim offsets(1 To 6) As Integer
    Dim idx As Integer: idx = 0
    For Each t In thicknesses("Gauche")
        If t("rowOffset") = offset Or t("rowOffset") = offset + 1 Or t("rowOffset") = offset + 2 Then
            idx = idx + 1
            offsets(idx) = t("rowOffset")
            If idx = 3 Then Exit For
        End If
    Next t
    For Each t In thicknesses("Droite")
        If t("rowOffset") = offset Or t("rowOffset") = offset + 1 Or t("rowOffset") = offset + 2 Then
            idx = idx + 1
            offsets(idx) = t("rowOffset")
            If idx = 6 Then Exit For
        End If
    Next t
    If idx > 0 Then
        minOffset = offsets(1)
        For i = 2 To idx
            If offsets(i) < minOffset Then minOffset = offsets(i)
        Next i
    End If
    result = Trim(Join(gVals, " ") & " " & Join(dVals, " "))
    If minOffset > 0 Then
        GetMesureString = minOffset & "m : " & result
    Else
        GetMesureString = result
    End If
End Function

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
    
    ' Sauvegarde des mesures dans les colonnes 12 à 21
    Dim m As Integer
    For m = 1 To 10
        ws.Cells(nextRow, 11 + m).Value = CallByName(Me, "Mesure" & m, VbGet)
    Next m
End Sub

' Formate les épaisseurs en chaîne de caractères pour la sauvegarde
' @but : Convertit les données d'épaisseurs en chaîne formatée
' @return String : Chaîne formatée des épaisseurs
Private Function FormatThicknessesToString() As String
    Dim result As String
    result = ""
    Dim pos As Variant
    Dim positions As Variant: positions = Array("Gauche", "Droite")
    For Each pos In positions
        result = result & pos & "["
        result = result & FormatThicknessCollection(Thicknesses(pos))
        result = result & "]"
    Next pos
    FormatThicknessesToString = result
End Function

' Formate une collection d'épaisseurs en chaîne
' @but : Convertit une collection d'épaisseurs en chaîne formatée
' @param coll : Collection d'épaisseurs à formater
' @return String : Chaîne formatée des épaisseurs
Private Function FormatThicknessCollection(coll As Collection) As String
    Dim result As String
    result = ""
    
    Dim thickness As Object
    For Each thickness In coll
        If result <> "" Then result = result & ","
        result = result & thickness("rowOffset") & ":" & Format(thickness("value"), "0.00")
        If thickness.Exists("rattrapageValue") Then
            result = result & "/" & Format(thickness("rattrapageValue"), "0.00")
        End If
    Next thickness
    
    FormatThicknessCollection = result
End Function

Private Sub Class_Initialize()
    ' Initialiser les collections
    Set Defects = New Collection
    
    ' Charger les données depuis la feuille de production
    If Not PRODUCTION_WS Is Nothing Then
        LoadFromSheet PRODUCTION_WS
        Debug.Print "[Roll] Rouleau initialisé : " & ID
    Else
        Debug.Print "[Roll] ERREUR : PRODUCTION_WS non initialisé"
    End If
End Sub
