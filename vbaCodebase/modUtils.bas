Attribute VB_Name = "modUtils"
Option Explicit

Public Sub PromptAndSetTargetLength()
    Dim ws As Worksheet
    Set ws = PRODUCTION_WS
    Dim userInput As Variant
    userInput = InputBox("Enter a new target length (1 to 50 m):", "Set Target Length")
    If userInput = "" Then Exit Sub
    If Not IsNumeric(userInput) Then
        MsgBox "Please enter a numeric value.", vbExclamation
        Exit Sub
    End If
    Dim val As Double
    val = CDbl(userInput)
    If val < 1 Or val > 50 Then
        MsgBox "Value must be between 1 and 50.", vbExclamation
        Exit Sub
    End If
    Call SetTargetLength(ws, val)
End Sub 