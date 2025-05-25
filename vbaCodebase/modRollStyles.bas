Attribute VB_Name = "modRollStyles"
Option Explicit

' Couleurs principales (en BGR pour VBA)
Public Const COLOR_BG_WHITE As Long = &HFFFFFF
Public Const COLOR_BG_GRAY As Long = &H808080 ' #808080 (gris fonc�, zone inactive)
Public Const COLOR_BG_GRAY_LIGHT As Long = &HA6A6A6 ' #A6A6A6 (gris clair, lengthCols)
Public Const COLOR_BG_BLUE_LIGHT As Long = &HF8E9DA ' #DAE9F8 en BGR
Public Const COLOR_TXT_BLUE As Long = &H985C21 ' #215C98 en BGR
Public Const COLOR_TXT_WHITE As Long = &HFFFFFF
Public Const COLOR_TXT_RED As Long = &H0000FF ' #FF0000 en BGR
Public Const COLOR_TXT_ORANGE As Long = &HC0FF00 ' #FFC000 en BGR

' Applique le style "zone active" (fond blanc, texte bleu)
Public Sub ApplyActiveStyle(rng As Range)
    rng.Interior.Color = COLOR_BG_WHITE
    rng.Font.Color = COLOR_TXT_BLUE
    rng.Locked = True
End Sub

' Applique le style "lengthCols" (fond gris clair, texte bleu)
Public Sub ApplyLengthStyle(rng As Range)
    rng.Interior.Color = COLOR_BG_GRAY_LIGHT
    rng.Font.Color = COLOR_TXT_BLUE
End Sub

' Applique le style "thickness" selon la valeur (vide, OK, NOK, orange si [4,5[ ou >9)
Public Sub ApplyThicknessStyle(cell As Range)
    Dim ws As Worksheet
    Set ws = cell.Worksheet

    ' D�prot�ger si n�cessaire
    If ws.ProtectContents Then
        ws.Unprotect
    End If

    If IsEmpty(cell.Value) Or Trim(cell.Value) = "" Then
        ' Cas cellule vide : fond bleu, texte bleu
        cell.Interior.Color = COLOR_BG_BLUE_LIGHT   ' #DAE9F8 en BGR
        cell.Font.Color = COLOR_TXT_BLUE   ' #215C98 en BGR
        
        ' Si c'est une cellule officielle qui devient vide, on d�sactive la cellule de rattrapage
        If Not Intersect(cell, Range("leftThicknessCels")) Is Nothing Or _
           Not Intersect(cell, Range("rightThicknessCels")) Is Nothing Then
            ' D�terminer la cellule de rattrapage potentielle
            Dim rattrapageCell As Range
            Dim isLastRow As Boolean
            isLastRow = cell.Row = Range("activeRollArea").Rows(Range("activeRollArea").Rows.Count).Row
            
            If isLastRow Then
                Set rattrapageCell = cell.Offset(-1, 0)
            Else
                Set rattrapageCell = cell.Offset(1, 0)
            End If
            
            ' V�rifier si la cellule de rattrapage existe dans les plages appropri�es
            Dim isValidRattrapage As Boolean
            isValidRattrapage = False
            On Error Resume Next
            isValidRattrapage = (Not Intersect(rattrapageCell, Range("leftSecThicknessCels")) Is Nothing) Or _
                                (Not Intersect(rattrapageCell, Range("rightSecThicknessCels")) Is Nothing)
            On Error GoTo 0
            If isValidRattrapage Then
                ' D�sactiver la cellule de rattrapage
                rattrapageCell.Locked = True
                rattrapageCell.Interior.Color = COLOR_BG_WHITE
                rattrapageCell.Font.Color = COLOR_TXT_WHITE
            End If
        End If
    Else
        Dim v As Double
        v = Val(cell.Value)
        If v < Range("ctrlMinThickness").Value Then
            ' Rouge, texte blanc
            cell.Interior.Color = RGB(255, 0, 0)
            cell.Font.Color = COLOR_TXT_WHITE
            
            ' V�rifier si c'est une cellule de mesure (pas de rattrapage)
            If Not Intersect(cell, Range("leftSecThicknessCels")) Is Nothing Or _
               Not Intersect(cell, Range("rightSecThicknessCels")) Is Nothing Then
                ' Ne rien faire pour les cellules de rattrapage
            Else
                ' D�terminer la cellule de rattrapage potentielle
                isLastRow = cell.Row = Range("activeRollArea").Rows(Range("activeRollArea").Rows.Count).Row
                
                If isLastRow Then
                    Set rattrapageCell = cell.Offset(-1, 0)
                Else
                    Set rattrapageCell = cell.Offset(1, 0)
                End If
                
                ' V�rifier si la cellule de rattrapage existe dans les plages appropri�es
                isValidRattrapage = False
                On Error Resume Next
                isValidRattrapage = (Not Intersect(rattrapageCell, Range("leftSecThicknessCels")) Is Nothing) Or _
                                    (Not Intersect(rattrapageCell, Range("rightSecThicknessCels")) Is Nothing)
                On Error GoTo 0
                If isValidRattrapage Then
                    ' La cellule de rattrapage existe, on peut l'activer
                    rattrapageCell.Locked = False
                    Call ApplyThicknessStyle(rattrapageCell)
                End If
            End If
        ElseIf (v < Range("ctrlWarnThickness").Value) Or (v > 9) Then
            ' Vert, texte orange
            cell.Interior.Color = RGB(0, 176, 80)
            cell.Font.Color = RGB(255, 192, 0)
            
            ' Si la cellule devient OK, on d�sactive la cellule de rattrapage
            If Not Intersect(cell, Range("leftThicknessCels")) Is Nothing Or _
               Not Intersect(cell, Range("rightThicknessCels")) Is Nothing Then
                ' D�terminer la cellule de rattrapage potentielle
                isLastRow = cell.Row = Range("activeRollArea").Rows(Range("activeRollArea").Rows.Count).Row
                
                If isLastRow Then
                    Set rattrapageCell = cell.Offset(-1, 0)
                Else
                    Set rattrapageCell = cell.Offset(1, 0)
                End If
                
                ' V�rifier si la cellule de rattrapage existe dans les plages appropri�es
                isValidRattrapage = False
                On Error Resume Next
                isValidRattrapage = (Not Intersect(rattrapageCell, Range("leftSecThicknessCels")) Is Nothing) Or _
                                    (Not Intersect(rattrapageCell, Range("rightSecThicknessCels")) Is Nothing)
                On Error GoTo 0
                If isValidRattrapage Then
                    ' D�sactiver la cellule de rattrapage
                    rattrapageCell.Locked = True
                    rattrapageCell.Interior.Color = COLOR_BG_WHITE
                    rattrapageCell.Font.Color = COLOR_TXT_WHITE
                End If
            End If
        Else
            ' Vert, texte blanc
            cell.Interior.Color = RGB(0, 176, 80)
            cell.Font.Color = COLOR_TXT_WHITE
            
            ' Si la cellule devient OK, on d�sactive la cellule de rattrapage
            If Not Intersect(cell, Range("leftThicknessCels")) Is Nothing Or _
               Not Intersect(cell, Range("rightThicknessCels")) Is Nothing Then
                ' D�terminer la cellule de rattrapage potentielle
                isLastRow = cell.Row = Range("activeRollArea").Rows(Range("activeRollArea").Rows.Count).Row
                
                If isLastRow Then
                    Set rattrapageCell = cell.Offset(-1, 0)
                Else
                    Set rattrapageCell = cell.Offset(1, 0)
                End If
                
                ' V�rifier si la cellule de rattrapage existe dans les plages appropri�es
                isValidRattrapage = False
                On Error Resume Next
                isValidRattrapage = (Not Intersect(rattrapageCell, Range("leftSecThicknessCels")) Is Nothing) Or _
                                    (Not Intersect(rattrapageCell, Range("rightSecThicknessCels")) Is Nothing)
                On Error GoTo 0
                If isValidRattrapage Then
                    ' D�sactiver la cellule de rattrapage
                    rattrapageCell.Locked = True
                    rattrapageCell.Interior.Color = COLOR_BG_WHITE
                    rattrapageCell.Font.Color = COLOR_TXT_WHITE
                End If
            End If
        End If
    End If

    ' Reproter si elle �tait prot�g�e au d�part
    If ws.ProtectContents Then
        ws.Protect
    End If
End Sub

' Applique le style "zone inactive" (fond gris, texte gris)
Public Sub ApplyInactiveStyle(rng As Range)
    rng.Interior.Color = COLOR_BG_GRAY
    rng.Font.Color = COLOR_BG_GRAY
    rng.Locked = True
End Sub

' V�rifie si un nom existe dans le classeur
Public Function NameExists(nom As String) As Boolean
    Dim n As Name
    NameExists = False
    For Each n In ThisWorkbook.Names
        If n.Name = nom Or n.Name Like "*" & nom Then
            NameExists = True
            Exit Function
        End If
    Next n
End Function

Public Sub FormatRollLayout()
    Dim ws As Worksheet
    Set ws = PRODUCTION_WS
    If ws Is Nothing Then Exit Sub
    ws.Unprotect

    ' 1. Zone inactive (fond et texte gris, verrouill�e)
    Dim rngInactive As Range
    On Error Resume Next
    Set rngInactive = Evaluate(ThisWorkbook.Names("inactiveRollArea").RefersTo)
    On Error GoTo 0
    If Not rngInactive Is Nothing Then
        Call ApplyInactiveStyle(rngInactive)
    End If

    ' 2. Zone active (fond blanc, texte bleu, verrouill�e)
    Dim rngActive As Range
    On Error Resume Next
    Set rngActive = Evaluate(ThisWorkbook.Names("activeRollArea").RefersTo)
    On Error GoTo 0
    If Not rngActive Is Nothing Then
        Call ApplyActiveStyle(rngActive)
    End If

    ' 3. Colonnes lengthCols dans la zone active (fond gris clair, texte bleu, verrouill�e)
    Dim rngLength As Range
    On Error Resume Next
    Set rngLength = Application.Intersect( _
        Evaluate(ThisWorkbook.Names("lengthCols").RefersTo), _
        rngActive)
    On Error GoTo 0
    If Not rngLength Is Nothing Then
        Call ApplyLengthStyle(rngLength)
    End If

    ' 4. Cellules de mesure officielles (left/rightThicknessCels) : bleu clair si vide, d�verrouill�es
    Dim thickNames As Variant
    thickNames = Array("leftThicknessCels", "rightThicknessCels")
    Dim i As Integer
    Dim thickName As String, refString As String
    
    ' D�prot�ger une seule fois avant la boucle si n�cessaire
    If ws.ProtectContents Then
        ws.Unprotect
    End If
    
    For i = LBound(thickNames) To UBound(thickNames)
        thickName = thickNames(i)
        If NameExists(thickName) Then
            refString = ThisWorkbook.Names(thickName).RefersTo
            If refString = "=FAUX" Or refString = "=FALSE" Then
                ' Ne rien faire
            Else
                refString = Replace(refString, "=", "")
                Dim addresses As Variant
                addresses = Split(refString, ",")
                Dim addr As Variant
                For Each addr In addresses
                    Dim cell As Range
                    Set cell = ws.Range(addr)
                    Call ApplyThicknessStyle(cell)
                    cell.Locked = False
                Next addr
            End If
        End If
    Next i
    
    ' Reproter apr�s la boucle si elle �tait prot�g�e au d�part
    If ws.ProtectContents Then
        ws.Protect
    End If

    ' 5. D�verrouille les colonnes de d�fauts dans la zone active
    Dim defNames As Variant
    defNames = Array("leftDefaultsCol", "centerDefaultsCol", "rightDefaultsCol")
    Dim defName As String, rngDef As Range
    For i = LBound(defNames) To UBound(defNames)
        defName = defNames(i)
        If NameExists(defName) Then
            On Error Resume Next
            Set rngDef = Application.Intersect( _
                Evaluate(ThisWorkbook.Names(defName).RefersTo), _
                rngActive)
            On Error GoTo 0
            If Not rngDef Is Nothing Then
                rngDef.Locked = False
                rngDef.Font.Color = COLOR_TXT_RED
            End If
        End If
    Next i

    ws.Protect
End Sub
