Attribute VB_Name = "PROD"

Option Explicit

' === Adresse de la cellule contenant la longueur cible du rouleau
' (Constante globale utilisée)
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit
    Application.EnableEvents = False
    
    Debug.Print "[Worksheet_Change] Cellule modifiée : " & Target.Address & " - Valeur : " & Target.Value
    
    ' Déprotéger la feuille si nécessaire
    Dim wasProtected As Boolean
    wasProtected = Me.ProtectContents
    If wasProtected Then Me.Unprotect

    ' Vérifie si la cellule modifiée est dans la zone active
    Dim rngActive As Range
    On Error Resume Next
    Set rngActive = Me.Range(ThisWorkbook.Names("activeRollArea").RefersTo)
    On Error GoTo 0
    If Not rngActive Is Nothing Then
        If Not Intersect(Target, rngActive) Is Nothing Then
            Call UpdateRollConformState
        End If
    End If

    ' Vérifie si la cellule modifiée est une des cellules de machine
    If Target.Address = Range(RANGE_SHIFT_MACHINE_PRISE_POSTE).Address Then
        Debug.Print "[Worksheet_Change] Machine prise poste modifiée"
        ' Mise à jour de la cellule de longueur prise de poste
        Select Case Target.Value
            Case "Démarrée"
                Debug.Print "[Worksheet_Change] État : Démarrée"
                Range("AF61").Locked = False
                Range("AF61").Interior.Color = &HF8E9DA  ' #DAE9F8 en BGR
                Range("AF61").Font.Color = &H985C21     ' #215C98 en BGR
            Case "A l'Arrêt"
                Debug.Print "[Worksheet_Change] État : À l'arrêt"
                Range("AF61").Locked = True
                Range("AF61").Interior.Color = &HF2F2F2  ' #F2F2F2 (gris)
                Range("AF61").Font.Color = &HF2F2F2     ' #F2F2F2 (gris)
                Range("AF61").Value = ""
        End Select
    ElseIf Target.Address = Range(RANGE_SHIFT_MACHINE_FIN_POSTE).Address Then
        Debug.Print "[Worksheet_Change] Machine fin poste modifiée"
        ' Mise à jour de la cellule de longueur fin de poste
        Select Case Target.Value
            Case "Démarrée"
                Debug.Print "[Worksheet_Change] État : Démarrée"
                Range("AF64").Locked = False
                Range("AF64").Interior.Color = &HF8E9DA  ' #DAE9F8 en BGR
                Range("AF64").Font.Color = &H985C21     ' #215C98 en BGR
            Case "A l'Arrêt"
                Debug.Print "[Worksheet_Change] État : À l'arrêt"
                Range("AF64").Locked = True
                Range("AF64").Interior.Color = &HF2F2F2  ' #F2F2F2 (gris)
                Range("AF64").Font.Color = &HF2F2F2     ' #F2F2F2 (gris)
                Range("AF64").Value = ""
        End Select
    End If

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
    
    ' Applique le style uniquement aux cellules modifiées qui sont dans la plage d'épaisseurs
    Dim cell As Range
    For Each cell In Target.Cells
        If Not Intersect(cell, allThickness) Is Nothing Then
            Call ApplyThicknessStyle(cell)
        End If
    Next cell

    ' Si la cellule modifiée est la longueur cible
    If Not Intersect(Target, Range(TARGET_LENGTH_ADDR)) Is Nothing Then
        ' Lancer l'initialisation des composants
        Call initializeComponents
    End If

    ' Vérifie si la cellule modifiée est une des cellules de contrôle global
    Dim ctrlNames As Variant
    ctrlNames = Array("micG1", "micG2", "micG3", "micD1", "micD2", "micD3", _
        "masseSurfaciqueGG", "masseSurfaciqueGC", "masseSurfaciqueDC", "masseSurfaciqueDD", _
        "bain")
    Dim i As Integer
    For i = LBound(ctrlNames) To UBound(ctrlNames)
        If Not Intersect(Target, Range(ctrlNames(i))) Is Nothing Then
            Range("AR60:AV60").Value = ""
            Exit For
        End If
    Next i

SafeExit:
    ' Reproter la feuille si elle était protégée
    If wasProtected Then Me.Protect
    Application.EnableEvents = True
End Sub

