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

Public Sub PromptAndSetRollNumber()
    Dim ws As Worksheet
    Set ws = PRODUCTION_WS
    Dim userInput As Variant
    userInput = InputBox("Modifier le num�ro de rouleau ?", "Set Roll Number")
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
    Call SetRollNumber(ws, val)
End Sub

Public Sub PromptAndSetModePermissif()
    Dim rep As VbMsgBoxResult
    rep = MsgBox("Activer le mode permissif ? (OUI = autorise la d�coupe/non-conforme sur confirmation, NON = refuse strictement)", vbYesNo + vbQuestion, "Mode permissif")
    If rep = vbYes Then
        MODE_PERMISSIF = True
    Else
        MODE_PERMISSIF = False
    End If
    WriteModePermissifToSheet
    MsgBox "Mode permissif : " & IIf(MODE_PERMISSIF, "OUI", "NON"), vbInformation
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

Public Sub SetRollNumber(ws As Worksheet, rollNumber As Long)
    Application.EnableEvents = False
    ws.Unprotect
    ws.Range("BH78").Value = rollNumber
    ws.Range("BH78").Locked = True
    ws.Protect
    Debug.Print "[SetRollNumber] Nouveau num�ro de roll = " & rollNumber
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

Public Sub ReadModePermissifFromSheet()
    Dim wsParams As Worksheet
    Set wsParams = ThisWorkbook.Sheets("params")
    Dim val As String
    val = UCase(Trim(wsParams.Range("E1").Value))
    If val = "OUI" Or val = "TRUE" Or val = "1" Then
        MODE_PERMISSIF = True
    Else
        MODE_PERMISSIF = False
    End If
End Sub

Public Sub WriteModePermissifToSheet()
    Dim wsParams As Worksheet
    Set wsParams = ThisWorkbook.Sheets("params")
    If MODE_PERMISSIF Then
        wsParams.Range("E1").Value = "OUI"
    Else
        wsParams.Range("E1").Value = "NON"
    End If
End Sub

Public Sub SetModePermissif(val As Boolean)
    MODE_PERMISSIF = val
    WriteModePermissifToSheet
End Sub


' V�rifie si un nom existe dans le classeur
' @but : V�rifier l'existence d'un nom d�fini dans le classeur
' @param nom (String) : nom � v�rifier
' @return Boolean : True si le nom existe, False sinon
' @pr� : Aucun
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


' Envoi � la d�coupe, en for�ant "NON CONFORME" (BK=84=""), puis lance l'export
' @but : Vider le contenu de la cellule BK84
' @param Aucun
' @return Aucun
' @pr� : PRODUCTION_WS doit �tre initialis�
Public Sub ToDecoupe()
    If PRODUCTION_WS Is Nothing Then Exit Sub
    
    ' Bo�te de dialogue de confirmation
    Dim response As VbMsgBoxResult
    response = MsgBox("�tes-vous s�r de vouloir envoyer ce rouleau vers la d�coupe ?", vbYesNo + vbQuestion, "Confirmation")
    
    ' Si l'utilisateur clique sur "Non", on annule l'op�ration
    If response = vbNo Then
        Debug.Print "[ToDecoupe] Op�ration annul�e par l'utilisateur"
        Exit Sub
    End If
    
    ' D�prot�ger si n�cessaire
    Dim wasProtected As Boolean
    wasProtected = PRODUCTION_WS.ProtectContents
    If wasProtected Then PRODUCTION_WS.Unprotect
    
    ' Vider la cellule BK84
    PRODUCTION_WS.Range("BK84").Value = ""
    
    ' Reproter si elle �tait prot�g�e au d�part
    If wasProtected Then PRODUCTION_WS.Protect
    
    Debug.Print "[ClearCellBK84] Cellule BK84 vid�e"
    call saveRollFromProd
End Sub