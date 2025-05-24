Attribute VB_Name = "modRollStyles"
Option Explicit

' Couleurs principales (en BGR pour VBA)
Public Const COLOR_BG_WHITE As Long = &HFFFFFF
Public Const COLOR_BG_GRAY As Long = &H808080 ' #808080
Public Const COLOR_BG_BLUE_LIGHT As Long = &HF8E9DA ' #DAE9F8 en BGR
Public Const COLOR_TXT_BLUE As Long = &H985C21 ' #215C98 en BGR
Public Const COLOR_TXT_WHITE As Long = &HFFFFFF
Public Const COLOR_TXT_RED As Long = &H0000FF ' #FF0000 en BGR
Public Const COLOR_TXT_ORANGE As Long = &H00A5FF ' #FFA500 en BGR

' Applique le style "zone active" (fond blanc, texte bleu)
Public Sub ApplyActiveStyle(rng As Range)
    rng.Interior.Color = COLOR_BG_WHITE
    rng.Font.Color = COLOR_TXT_BLUE
    rng.Locked = True
End Sub

' Applique le style "lengthCols" (fond gris, texte bleu)
Public Sub ApplyLengthStyle(rng As Range)
    rng.Interior.Color = COLOR_BG_GRAY
    rng.Font.Color = COLOR_TXT_BLUE
End Sub

' Applique le style "thickness" selon la valeur (vide, OK, NOK, orange si [4,5[ ou >9)
Public Sub ApplyThicknessStyle(cell As Range)
    Const COLOR_TXT_ORANGE As Long = &H00A5FF ' #FFA500 en BGR
    If Trim(cell.Value & "") = "" Then
        cell.Interior.Color = COLOR_BG_BLUE_LIGHT
        cell.Font.Color = COLOR_TXT_BLUE
        Debug.Print "[ApplyThicknessStyle] vide : " & cell.Address
    Else
        Dim v As Double
        v = Val(cell.Value)
        Debug.Print "[ApplyThicknessStyle] valeur : " & cell.Address & " = '" & cell.Value & "' (Val=" & v & ", Type=" & TypeName(cell.Value) & ")"
        If v >= 4 And v < 5 Then
            cell.Interior.Color = RGB(0, 176, 80)   ' Vert
            cell.Font.Color = COLOR_TXT_ORANGE
        ElseIf v > 9 Then
            cell.Interior.Color = RGB(0, 176, 80)   ' Vert
            cell.Font.Color = COLOR_TXT_ORANGE
        ElseIf v >= 4 Then
            cell.Interior.Color = RGB(0, 176, 80)   ' Vert
            cell.Font.Color = COLOR_TXT_WHITE
        ElseIf v > 0 Then
            cell.Interior.Color = RGB(255, 0, 0)    ' Rouge
            cell.Font.Color = COLOR_TXT_WHITE
        Else
            Debug.Print "[ApplyThicknessStyle] non numérique ou <=0 : " & cell.Address & " = '" & cell.Value & "'"
        End If
    End If
End Sub

' Applique le style "zone inactive" (fond gris, texte gris)
Public Sub ApplyInactiveStyle(rng As Range)
    rng.Interior.Color = COLOR_BG_GRAY
    rng.Font.Color = COLOR_BG_GRAY
    rng.Locked = True
End Sub

' Vérifie si un nom existe dans le classeur
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

    ' 1. Zone inactive (fond et texte gris, verrouillée)
    Dim rngInactive As Range
    On Error Resume Next
    Set rngInactive = Evaluate(ThisWorkbook.Names("inactiveRollArea").RefersTo)
    On Error GoTo 0
    If Not rngInactive Is Nothing Then
        Call ApplyInactiveStyle(rngInactive)
    End If

    ' 2. Zone active (fond blanc, texte bleu, verrouillée)
    Dim rngActive As Range
    On Error Resume Next
    Set rngActive = Evaluate(ThisWorkbook.Names("activeRollArea").RefersTo)
    On Error GoTo 0
    If Not rngActive Is Nothing Then
        Call ApplyActiveStyle(rngActive)
    End If

    ' 3. Colonnes lengthCols dans la zone active (fond gris, texte bleu, verrouillée)
    Dim rngLength As Range
    On Error Resume Next
    Set rngLength = Application.Intersect( _
        Evaluate(ThisWorkbook.Names("lengthCols").RefersTo), _
        rngActive)
    On Error GoTo 0
    If Not rngLength Is Nothing Then
        Call ApplyLengthStyle(rngLength)
    End If

    ' 4. Cellules de mesure officielles (left/rightThicknessCels) : bleu clair si vide, déverrouillées
    Dim thickNames As Variant
    thickNames = Array("leftThicknessCels", "rightThicknessCels")
    Dim i As Integer
    Dim thickName As String, refString As String
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

    ' 5. Déverrouille les colonnes de défauts dans la zone active
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