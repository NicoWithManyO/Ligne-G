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

Public Sub SetTargetLength(ws As Worksheet, targetLength As Double)
    Application.EnableEvents = False
    ws.Unprotect
    ws.Range(TARGET_LENGTH_ADDR).Value = targetLength
    ws.Range(TARGET_LENGTH_ADDR).Locked = True
    ws.Protect
    Debug.Print "[SetTargetLength] Nouvelle longueur cible = " & targetLength
    Call initializeComponents
    Application.EnableEvents = True
End Sub

Public Sub SetOFNumber(ws As Worksheet, ofNumber As Long)
    Application.EnableEvents = False
    ws.Unprotect
    ws.Range(RANGE_OF_NUMBER).Value = ofNumber
    ws.Range(RANGE_OF_NUMBER).Locked = True
    ws.Protect
    Debug.Print "[SetOFNumber] Nouveau num�ro OF = " & ofNumber
    Application.EnableEvents = True
End Sub

Public Sub SetCutOFNumber(ws As Worksheet, cutOfNumber As Long)
    Application.EnableEvents = False
    ws.Unprotect
    ws.Range(RANGE_CUT_OF_NUMBER).Value = cutOfNumber
    ws.Range(RANGE_CUT_OF_NUMBER).Locked = True
    ws.Protect
    Debug.Print "[SetCutOFNumber] Nouveau num�ro OF de coupe = " & cutOfNumber
    Application.EnableEvents = True
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