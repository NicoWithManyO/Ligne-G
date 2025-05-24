Attribute VB_Name = "PROD"

Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit
    Application.EnableEvents = False

    ' Construction dynamique de la plage d'épaisseurs (toutes les sous-plages existantes)
    Dim allParts As Collection: Set allParts = New Collection
    If NameExists("leftThicknessCels") Then allParts.Add Range("leftThicknessCels")
    If NameExists("rightThicknessCels") Then allParts.Add Range("rightThicknessCels")
    If NameExists("leftSecThicknessCels") Then allParts.Add Range("leftSecThicknessCels")
    If NameExists("rightSecThicknessCels") Then allParts.Add Range("rightSecThicknessCels")

    Dim allThickness As Range, r As Variant
    For Each r In allParts
        If allThickness Is Nothing Then
            Set allThickness = r
        Else
            Set allThickness = Union(allThickness, r)
        End If
    Next r

    If allThickness Is Nothing Then GoTo SafeExit
    PRODUCTION_WS.Unprotect
    ' Applique le style uniquement aux cellules modifiées qui sont dans la plage d'épaisseurs
    Dim cell As Range
    For Each cell In Target.Cells
        If Not Intersect(cell, allThickness) Is Nothing Then
            Call ApplyThicknessStyle(cell)
        End If
    Next cell
    PRODUCTION_WS.Protect
SafeExit:
    Application.EnableEvents = True
End Sub

Public Sub ApplyThicknessStyle(cell As Range)
    Dim ws As Worksheet
    Set ws = cell.Worksheet

    ws.Unprotect

    If IsEmpty(cell.Value) Or Trim(cell.Value) = "" Then
        ' Cas cellule vide : fond bleu, texte bleu
        cell.Interior.Color = RGB(0, 112, 192)   ' Bleu
        cell.Font.Color = RGB(255, 255, 255)
    Else
        Dim v As Double
        v = Val(cell.Value)
        If v < 4 Then
            ' Rouge, texte blanc
            cell.Interior.Color = RGB(255, 0, 0)
            cell.Font.Color = RGB(255, 255, 255)
        ElseIf (v >= 4 And v < 5) Or v > 9 Then
            ' Vert, texte orange
            cell.Interior.Color = RGB(0, 176, 80)
            cell.Font.Color = RGB(255, 192, 0)
        ElseIf v >= 5 And v <= 9 Then
            ' Vert, texte blanc
            cell.Interior.Color = RGB(0, 176, 80)
            cell.Font.Color = RGB(255, 255, 255)
        End If
    End If

    ws.Protect
End Sub

