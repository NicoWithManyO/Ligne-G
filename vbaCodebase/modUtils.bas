Attribute VB_Name = "modUtils"
Option Explicit

Public Sub PromptAndSetTargetLength()
    Dim ws As Worksheet
    Set ws = PRODUCTION_WS
    Dim userInput As Variant
    userInput = InputBox("Modifier la longueur cible (1 � 50m) ?", "Set Target Length")
    If userInput = "" Then Exit Sub
    If Not IsNumeric(userInput) Then
        MsgBox "Une valeur num�rique est attendue", vbExclamation
        Exit Sub
    End If
    Dim val As Double
    val = CDbl(userInput)
    If val < 1 Or val > 50 Then
        MsgBox "La valeur doit �tre comprise entre 1 et 50", vbExclamation
        Exit Sub
    End If
    Call SetTargetLength(ws, val)
End Sub

Public Sub PromptAndSetOFNumber()
    Dim ws As Worksheet
    Set ws = PRODUCTION_WS
    Dim userInput As Variant
    userInput = InputBox("Modifier le num�ro OF ?", "Set OF Number")
    If userInput = "" Then Exit Sub
    If Not IsNumeric(userInput) Then
        MsgBox "Une valeur num�rique est attendue", vbExclamation
        Exit Sub
    End If
    Dim val As Long
    val = CLng(userInput)
    If val < 1 Then
        MsgBox "La valeur doit �tre sup�rieure � 0", vbExclamation
        Exit Sub
    End If
    Call SetOFNumber(ws, val)
End Sub

Public Sub PromptAndSetCutOFNumber()
    Dim ws As Worksheet
    Set ws = PRODUCTION_WS
    Dim userInput As Variant
    userInput = InputBox("Modifier le num�ro OF de coupe ?", "Set Cut OF Number")
    If userInput = "" Then Exit Sub
    If Not IsNumeric(userInput) Then
        MsgBox "Une valeur num�rique est attendue", vbExclamation
        Exit Sub
    End If
    Dim val As Long
    val = CLng(userInput)
    If val < 1 Then
        MsgBox "La valeur doit �tre sup�rieure � 0", vbExclamation
        Exit Sub
    End If
    Call SetCutOFNumber(ws, val)
End Sub

' Limite la zone de d�filement � la ligne 120
Public Sub LimitScrollArea()
    Dim ws As Worksheet
    Set ws = PRODUCTION_WS
    If ws Is Nothing Then Exit Sub
    
    ' D�prot�ger si n�cessaire
    If ws.ProtectContents Then
        ws.Unprotect
    End If
    
    ' Trouver la derni�re colonne utilis�e
    Dim lastCol As String
    lastCol = Split(ws.Cells(1, ws.Columns.Count).End(xlToLeft).Address, "$")(1)
    
    ' D�finir la zone de d�filement jusqu'� la ligne 120
    ws.ScrollArea = "AA50:" & lastCol & "120"
    
    ' Reproter si elle �tait prot�g�e au d�part
    If ws.ProtectContents Then
        ws.Protect
    End If
End Sub 

' Met la date du jour dans la cellule shiftDate
Public Sub SetTodayDate()
    If PRODUCTION_WS Is Nothing Then Exit Sub
    
    ' D�prot�ger si n�cessaire
    If PRODUCTION_WS.ProtectContents Then
        PRODUCTION_WS.Unprotect
    End If
    
    ' Mettre la date du jour
    Range("shiftDate").Value = Date
    
    ' Reproter si elle �tait prot�g�e au d�part
    If PRODUCTION_WS.ProtectContents Then
        PRODUCTION_WS.Protect
    End If
End Sub


