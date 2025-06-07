Attribute VB_Name = "PROD"

Option Explicit

' === Adresse de la cellule contenant la longueur cible du rouleau
' (Constante globale utilis�e)
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit
    Application.EnableEvents = False
    
    Debug.Print "[Worksheet_Change] Cellule modifi�e : " & Target.Address & " - Valeur : " & Target.Value
    
    ' D�prot�ger la feuille si n�cessaire
    Dim wasProtected As Boolean
    wasProtected = Me.ProtectContents
    If wasProtected Then Me.Unprotect

    ' V�rifie si la cellule modifi�e est dans la zone active
    Dim rngActive As Range
    On Error Resume Next
    Set rngActive = Me.Range(ThisWorkbook.Names("activeRollArea").RefersTo)
    On Error GoTo 0
    If Not rngActive Is Nothing Then
        If Not Intersect(Target, rngActive) Is Nothing Then
            Call UpdateRollConformState
        End If
    End If

    ' V�rifie si la cellule modifi�e est une des cellules de machine
    If Target.Address = Range(RANGE_SHIFT_MACHINE_PRISE_POSTE).Address Then
        Debug.Print "[Worksheet_Change] Machine prise poste modifi�e"
        ' Mise � jour de la cellule de longueur prise de poste
        Select Case Target.Value
            Case "D�marr�e"
                Debug.Print "[Worksheet_Change] �tat : D�marr�e"
                Range("AF61").Locked = False
                Range("AF61").Interior.Color = &HF8E9DA  ' #DAE9F8 en BGR
                Range("AF61").Font.Color = &H985C21     ' #215C98 en BGR
            Case "A l'Arr�t"
                Debug.Print "[Worksheet_Change] �tat : � l'arr�t"
                Range("AF61").Locked = True
                Range("AF61").Interior.Color = &HF2F2F2  ' #F2F2F2 (gris)
                Range("AF61").Font.Color = &HF2F2F2     ' #F2F2F2 (gris)
                Range("AF61").Value = ""
        End Select
    ElseIf Target.Address = Range(RANGE_SHIFT_MACHINE_FIN_POSTE).Address Then
        Debug.Print "[Worksheet_Change] Machine fin poste modifi�e"
        ' Mise � jour de la cellule de longueur fin de poste
        Select Case Target.Value
            Case "D�marr�e"
                Debug.Print "[Worksheet_Change] �tat : D�marr�e"
                Range("AF64").Locked = False
                Range("AF64").Interior.Color = &HF8E9DA  ' #DAE9F8 en BGR
                Range("AF64").Font.Color = &H985C21     ' #215C98 en BGR
            Case "A l'Arr�t"
                Debug.Print "[Worksheet_Change] �tat : � l'arr�t"
                Range("AF64").Locked = True
                Range("AF64").Interior.Color = &HF2F2F2  ' #F2F2F2 (gris)
                Range("AF64").Font.Color = &HF2F2F2     ' #F2F2F2 (gris)
                Range("AF64").Value = ""
        End Select
    End If

    ' Construction dynamique de la plage d'�paisseurs (toutes les sous-plages existantes)
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
    
    ' Applique le style uniquement aux cellules modifi�es qui sont dans la plage d'�paisseurs
    Dim cell As Range
    For Each cell In Target.Cells
        If Not Intersect(cell, allThickness) Is Nothing Then
            Call ApplyThicknessStyle(cell)
        End If
    Next cell

    ' Si la cellule modifi�e est la longueur cible
    If Not Intersect(Target, Range(TARGET_LENGTH_ADDR)) Is Nothing Then
        ' Lancer l'initialisation des composants
        Call initializeComponents
    End If

    ' V�rifie si la cellule modifi�e est une des cellules de contr�le global
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
    ' Reproter la feuille si elle �tait prot�g�e
    If wasProtected Then Me.Protect
    Application.EnableEvents = True
End Sub

