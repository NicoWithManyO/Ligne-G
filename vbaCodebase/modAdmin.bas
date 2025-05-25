Attribute VB_Name = "modAdmin"
Option Explicit

Public Sub FillThicknessCellsRandom()
    Dim ws As Worksheet
    Set ws = PRODUCTION_WS
    Dim thickNames As Variant
    thickNames = Array("leftThicknessCels", "rightThicknessCels", "leftSecThicknessCels", "rightSecThicknessCels")
    Dim i As Integer, thickName As String, refString As String
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
                    On Error Resume Next
                    Set cell = ws.Range(addr)
                    If Not cell Is Nothing Then
                        cell.Value = Round(4.4 + Rnd() * (7.6 - 4.4), 2)
                    End If
                    On Error GoTo 0
                Next addr
            End If
        End If
    Next i
    Call FormatRollLayout
End Sub

Public Sub ClearThicknessCells()
    Dim ws As Worksheet
    Set ws = PRODUCTION_WS
    Dim thickNames As Variant
    thickNames = Array("leftThicknessCels", "rightThicknessCels", "leftSecThicknessCels", "rightSecThicknessCels")
    Dim i As Integer, thickName As String, refString As String
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
                    On Error Resume Next
                    Set cell = ws.Range(addr)
                    If Not cell Is Nothing Then
                        cell.Value = ""
                    End If
                    On Error GoTo 0
                Next addr
            End If
        End If
    Next i
    Call FormatRollLayout
End Sub 