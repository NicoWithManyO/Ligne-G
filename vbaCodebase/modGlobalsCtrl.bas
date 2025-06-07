Attribute VB_Name = "modGlobalsCtrl"
Option Explicit

' @but : Vérifie la conformité de tous les contrôles globaux par rapport à leurs min/max
' @param motif (ByRef) : motif détaillé en cas de non-conformité
' @return Boolean (True si conforme, False sinon)
' @pré : Les plages nommées doivent être initialisées
Public Function IsGlobalsCtrlConform(Optional ByRef motif As String = "") As Boolean
    Dim isConform As Boolean: isConform = True
    motif = ""
    ' Micronnaire
    Dim i As Integer, val As Double
    For i = 1 To 3
        val = ThisWorkbook.Names("micG" & i).RefersToRange.Value
        If val < ThisWorkbook.Names("micronnaireMin").RefersToRange.Value Or val > ThisWorkbook.Names("micronnaireMax").RefersToRange.Value Then
            isConform = False
            motif = motif & "micG" & i & " hors tolérance | "
        End If
        val = ThisWorkbook.Names("micD" & i).RefersToRange.Value
        If val < ThisWorkbook.Names("micronnaireMin").RefersToRange.Value Or val > ThisWorkbook.Names("micronnaireMax").RefersToRange.Value Then
            isConform = False
            motif = motif & "micD" & i & " hors tolérance | "
        End If
    Next i
    ' Bain
    val = ThisWorkbook.Names("bain").RefersToRange.Value
    If val < ThisWorkbook.Names("bainMin").RefersToRange.Value Or val > ThisWorkbook.Names("bainMax").RefersToRange.Value Then
        isConform = False
        motif = motif & "Bain hors tolérance | "
    End If
    ' Masse surfacique
    Dim masseNames As Variant: masseNames = Array("masseSurfaciqueGG", "masseSurfaciqueGC", "masseSurfaciqueDC", "masseSurfaciqueDD")
    Dim j As Integer
    For j = 0 To 3
        val = ThisWorkbook.Names(masseNames(j)).RefersToRange.Value
        If val < ThisWorkbook.Names("masseSurfMin").RefersToRange.Value Or val > ThisWorkbook.Names("masseSurfMax").RefersToRange.Value Then
            isConform = False
            motif = motif & masseNames(j) & " hors tolérance | "
        End If
    Next j
    ' LOI
    If UCase(Trim(ThisWorkbook.Names("loi").RefersToRange.Value)) <> "OK" Then
        isConform = False
        motif = motif & "LOI non conforme | "
    End If
    IsGlobalsCtrlConform = isConform

End Function

Public Sub TestGlobalsCtrlConform()
    Dim motif As String
    If IsGlobalsCtrlConform(motif) Then
        MsgBox "Contrôles globaux CONFORMES !", vbInformation
    Else
        MsgBox "NON CONFORME : " & motif, vbExclamation
    End If
End Sub 