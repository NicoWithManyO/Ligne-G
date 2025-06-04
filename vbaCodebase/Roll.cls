VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Roll"
Option Explicit

' Propri�t�s Roll
' ID : Identifiant unique du rouleau
Public ID As Variant
' OF : Num�ro d'ordre de fabrication
Public OF As Variant
' Number : Num�ro du rouleau
Public Number As Variant
' FabricationShift : Poste de fabrication
Public FabricationShift As Variant
' FabricationOperator : Op�rateur de fabrication
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
' Status : �tat du rouleau
Public Status As Variant
' SaveDateTime : Date et heure de la sauvegarde
Public SaveDateTime As Variant
' Defects : D�fauts d�tect�s
Public Defects As Variant
' Thicknesses : �paisseurs mesur�es (3m et 8m)
Public Thicknesses As Object
' Mesures dynamiques (toujours pr�sentes, vides si inutilis�es)
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
' Gm2 : Masse lin�aire
Public Gm2 As Variant

' Structure pour d�finir une colonne
Private Type ColumnConfig
    Header As String
    PropertyName As String
End Type

' Configuration des colonnes
Private Function GetColumnConfig() As Variant
    Dim config(1 To 22) As Object
    Dim i As Integer

    For i = 1 To 22
        Set config(i) = CreateObject("Scripting.Dictionary")
    Next i

    config(1)("Header") = "ID": config(1)("PropertyName") = "ID"
    config(2)("Header") = "OF": config(2)("PropertyName") = "OF"
    config(3)("Header") = "Num�ro": config(3)("PropertyName") = "Number"
    config(4)("Header") = "Poste": config(4)("PropertyName") = "FabricationShift"
    config(5)("Header") = "Statut": config(5)("PropertyName") = "Status"
    config(6)("Header") = "Poids tube": config(6)("PropertyName") = "PipeWeight"
    config(7)("Header") = "Poids total": config(7)("PropertyName") = "TotalWeight"
    config(8)("Header") = "Poids": config(8)("PropertyName") = "Weight"
    config(9)("Header") = "Longueur": config(9)("PropertyName") = "Length"
    config(10)("Header") = "g/m�": config(10)("PropertyName") = "Gm2"
    config(11)("Header") = "Date sauvegarde": config(11)("PropertyName") = "SaveDateTime"
    config(12)("Header") = "D�fauts": config(12)("PropertyName") = "Defects"
    config(13)("Header") = "Mesure 1": config(13)("PropertyName") = "Mesure1"
    config(14)("Header") = "Mesure 2": config(14)("PropertyName") = "Mesure2"
    config(15)("Header") = "Mesure 3": config(15)("PropertyName") = "Mesure3"
    config(16)("Header") = "Mesure 4": config(16)("PropertyName") = "Mesure4"
    config(17)("Header") = "Mesure 5": config(17)("PropertyName") = "Mesure5"
    config(18)("Header") = "Mesure 6": config(18)("PropertyName") = "Mesure6"
    config(19)("Header") = "Mesure 7": config(19)("PropertyName") = "Mesure7"
    config(20)("Header") = "Mesure 8": config(20)("PropertyName") = "Mesure8"
    config(21)("Header") = "Mesure 9": config(21)("PropertyName") = "Mesure9"
    config(22)("Header") = "Mesure 10": config(22)("PropertyName") = "Mesure10"

    GetColumnConfig = config
End Function

' Charge les donn�es depuis la feuille
' @but : Charge les donn�es du rouleau depuis la feuille de production
' @param ws : Feuille de travail source
' @return : aucun
Public Sub LoadFromSheet(ws As Worksheet)
    ' Chargement des propri�t�s de base
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
    
    ' Calcul du g/m�
    If Not IsEmpty(Weight) And Not IsEmpty(Length) And Length <> 0 Then
        Gm2 = Weight / Length
    Else
        Gm2 = ""
    End If
    
    ' Chargement des mesures tous les 5m (jusqu'� 10 mesures)
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

' G�n�re la cha�ne de mesure pour la m-i�me mesure (tous les 5m)
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

' Sauvegarde les donn�es dans la feuille
' @but : Sauvegarde les donn�es du rouleau dans la feuille de production
' @param ws : Feuille de travail cible
' @return : aucun
Public Sub SaveToSheet(ws As Worksheet)
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' R�cup�rer la configuration des colonnes
    Dim config As Variant
    config = GetColumnConfig()
    
    ' Si la ligne 1 est vide, r��crire les en-t�tes
    Dim i As Integer, enteteVide As Boolean
    enteteVide = True
    For i = 1 To UBound(config)
        If ws.Cells(1, i).Value <> "" Then
            enteteVide = False
            Exit For
        End If
    Next i
    If enteteVide Then
        For i = 1 To UBound(config)
            ws.Cells(1, i).Value = config(i)("Header")
        Next i
    End If
    
    ' Sauvegarder les donn�es
    For i = 1 To UBound(config)
        If config(i)("PropertyName") = "SaveDateTime" Then
            ws.Cells(nextRow, i).Value = Now()
        Else
            ws.Cells(nextRow, i).Value = CallByName(Me, config(i)("PropertyName"), VbGet)
        End If
    Next i
End Sub

' Formate les �paisseurs en cha�ne de caract�res pour la sauvegarde
' @but : Convertit les donn�es d'�paisseurs en cha�ne format�e
' @return String : Cha�ne format�e des �paisseurs
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

' Formate une collection d'�paisseurs en cha�ne
' @but : Convertit une collection d'�paisseurs en cha�ne format�e
' @param coll : Collection d'�paisseurs � formater
' @return String : Cha�ne format�e des �paisseurs
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
    
    ' Charger les donn�es depuis la feuille de production
    If Not PRODUCTION_WS Is Nothing Then
        LoadFromSheet PRODUCTION_WS
        Debug.Print "[Roll] Rouleau initialis� : " & ID
    Else
        Debug.Print "[Roll] ERREUR : PRODUCTION_WS non initialis�"
    End If
End Sub
