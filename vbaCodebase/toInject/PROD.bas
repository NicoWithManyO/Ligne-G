Attribute VB_Name = "PROD"

Option Explicit

' === Adresse de la cellule contenant la longueur cible du rouleau
' (Constante globale utilisée)

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

    ' Si la cellule modifiée est la longueur cible
    If Not Intersect(Target, Range(TARGET_LENGTH_ADDR)) Is Nothing Then
        ' Lancer l'initialisation des composants
        Call initializeComponents
    End If

SafeExit:
    Application.EnableEvents = True
End Sub

