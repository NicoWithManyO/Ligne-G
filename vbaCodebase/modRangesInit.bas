Attribute VB_Name = "modRangesInit"
Option Explicit

' Constantes pour les noms des ranges
Const RANGE_SHIFT_ID As String = "shiftID"
Const RANGE_SHIFT_DATE As String = "shiftDate"
Const RANGE_SHIFT_OPERATEUR As String = "shiftOperateur"
Const RANGE_SHIFT_VACCATION As String = "shiftVaccation"
Const RANGE_SHIFT_DUREE As String = "shiftDuree"
Const RANGE_SHIFT_MACHINE_PRISE_POSTE As String = "shiftMachinePrisePoste"
Const RANGE_SHIFT_LG_ENROULEE_PRISE_POSTE As String = "shiftLgEnrouleePrisePoste"
Const RANGE_SHIFT_MACHINE_FIN_POSTE As String = "shiftMachineFinPoste"
Const RANGE_SHIFT_LG_ENROULEE_FIN_POSTE As String = "shiftLgEnrouleeFinPoste"
Const RANGE_SHIFT_COMMENTAIRES As String = "shiftCommentaires"
Const RANGE_OF_NUMBER As String = "OFNumber"
Const RANGE_CUT_OF_NUMBER As String = "CutOFNumber"

' Constantes pour les ranges du rouleau
Const ROLL_START_ROW As Long = 68
Const ROLL_END_ROW As Long = 167
Const ROLL_START_COL As String = "AJ"
Const ROLL_END_COL As String = "BD"
Const ROLL_MEASURE_INTERVAL As Long = 5  ' Mesures tous les 5m
Const ROLL_MEASURE_OFFSET As Long = 3    ' Première mesure à 3m

' Constantes pour les ranges des valeurs limites
Const RANGE_CTRL_MIN_THICKNESS As String = "ctrlMinThickness"  ' Seuil rouge (BK59)
Const RANGE_CTRL_WARN_THICKNESS As String = "ctrlWarnThickness"  ' Seuil orange (BJ59)

' Initialise les ranges nommées pour le suivi des shifts
' @pre : PRODUCTION_WS doit être initialisé
' @return : aucun
Public Sub initShiftRanges()
    If PRODUCTION_WS Is Nothing Then
        Debug.Print "[initShiftRanges] ERREUR : PRODUCTION_WS non initialisé"
        Exit Sub
    End If
    
    ' Suppression des ranges existantes pour éviter les doublons
    On Error Resume Next
    ThisWorkbook.Names(RANGE_SHIFT_ID).Delete
    ThisWorkbook.Names(RANGE_SHIFT_DATE).Delete
    ThisWorkbook.Names(RANGE_SHIFT_OPERATEUR).Delete
    ThisWorkbook.Names(RANGE_SHIFT_VACCATION).Delete
    ThisWorkbook.Names(RANGE_SHIFT_DUREE).Delete
    ThisWorkbook.Names(RANGE_SHIFT_MACHINE_PRISE_POSTE).Delete
    ThisWorkbook.Names(RANGE_SHIFT_LG_ENROULEE_PRISE_POSTE).Delete
    ThisWorkbook.Names(RANGE_SHIFT_MACHINE_FIN_POSTE).Delete
    ThisWorkbook.Names(RANGE_SHIFT_LG_ENROULEE_FIN_POSTE).Delete
    ThisWorkbook.Names(RANGE_SHIFT_COMMENTAIRES).Delete
    On Error GoTo 0
    
    ' Création des nouvelles ranges
    ThisWorkbook.Names.Add Name:=RANGE_SHIFT_ID, RefersTo:=PRODUCTION_WS.Range("AC55")
    ThisWorkbook.Names.Add Name:=RANGE_SHIFT_DATE, RefersTo:=PRODUCTION_WS.Range("AD54")
    ThisWorkbook.Names.Add Name:=RANGE_SHIFT_OPERATEUR, RefersTo:=PRODUCTION_WS.Range("AD56")
    ThisWorkbook.Names.Add Name:=RANGE_SHIFT_VACCATION, RefersTo:=PRODUCTION_WS.Range("AD58")
    ThisWorkbook.Names.Add Name:=RANGE_SHIFT_DUREE, RefersTo:=PRODUCTION_WS.Range("AG58")
    
    ThisWorkbook.Names.Add Name:=RANGE_SHIFT_MACHINE_PRISE_POSTE, RefersTo:=PRODUCTION_WS.Range("AC61")
    ThisWorkbook.Names.Add Name:=RANGE_SHIFT_LG_ENROULEE_PRISE_POSTE, RefersTo:=PRODUCTION_WS.Range("AF61")
    ThisWorkbook.Names.Add Name:=RANGE_SHIFT_MACHINE_FIN_POSTE, RefersTo:=PRODUCTION_WS.Range("AC64")
    ThisWorkbook.Names.Add Name:=RANGE_SHIFT_LG_ENROULEE_FIN_POSTE, RefersTo:=PRODUCTION_WS.Range("AF64")
    ThisWorkbook.Names.Add Name:=RANGE_SHIFT_COMMENTAIRES, RefersTo:=PRODUCTION_WS.Range("AC67:AG71")
    
    ' Affichage des adresses des ranges
    Debug.Print "[initShiftRanges] -> " & RANGE_SHIFT_ID & " : AC55"
    Debug.Print "[initShiftRanges] -> " & RANGE_SHIFT_DATE & " : AD54"
    Debug.Print "[initShiftRanges] -> " & RANGE_SHIFT_OPERATEUR & " : AD56"
    Debug.Print "[initShiftRanges] -> " & RANGE_SHIFT_VACCATION & " : AD58"
    Debug.Print "[initShiftRanges] -> " & RANGE_SHIFT_DUREE & " : AG58"
    Debug.Print "[initShiftRanges] -> " & RANGE_SHIFT_MACHINE_PRISE_POSTE & " : AC61"
    Debug.Print "[initShiftRanges] -> " & RANGE_SHIFT_LG_ENROULEE_PRISE_POSTE & " : AF61"
    Debug.Print "[initShiftRanges] -> " & RANGE_SHIFT_MACHINE_FIN_POSTE & " : AC64"
    Debug.Print "[initShiftRanges] -> " & RANGE_SHIFT_LG_ENROULEE_FIN_POSTE & " : AF64"
    Debug.Print "[initShiftRanges] -> " & RANGE_SHIFT_COMMENTAIRES & " : AC67:AG71"
End Sub

' Définit les plages nommées pour le rouleau
' @pre : PRODUCTION_WS doit être initialisé
' @return : aucun
Public Sub defineRollNamedRanges()
    If PRODUCTION_WS Is Nothing Then
        Debug.Print "[defineRollNamedRanges] ERREUR : PRODUCTION_WS non initialisé"
        Exit Sub
    End If
    
    ' Initialisation de la cellule de longueur cible
    Set TARGET_LENGTH_CELL = PRODUCTION_WS.Range(TARGET_LENGTH_ADDR)
    Debug.Print "[defineRollNamedRanges] Longueur cible : " & TARGET_LENGTH_CELL.Value
    If TARGET_LENGTH_CELL Is Nothing Then
        Debug.Print "[defineRollNamedRanges] ERREUR : Cellule de longueur cible non trouvée : " & TARGET_LENGTH_ADDR
        Exit Sub
    End If
    
    ' === Configuration des colonnes ===
    Dim colConfig As Object: Set colConfig = CreateObject("Scripting.Dictionary")
    With colConfig
        .Add "lengthCols", Array(2, 10, 12, 20)      ' Colonnes de longueur
        .Add "leftThicknessCols", Array(4, 6, 8)      ' Colonnes d'épaisseur gauche
        .Add "rightThicknessCols", Array(14, 16, 18)  ' Colonnes d'épaisseur droite
        .Add "leftDefaultsCol", Array(1)              ' Colonne défaut gauche
        .Add "centerDefaultsCol", Array(11)           ' Colonne défaut centre
        .Add "rightDefaultsCol", Array(21)            ' Colonne défaut droite
    End With
    
    ' === Suppression des anciens noms ===
    Dim namesToDelete As Variant
    namesToDelete = Array("maxRollArea", "activeRollArea", "inactiveRollArea", _
                         "lengthCols", "leftThicknessCols", "rightThicknessCols", _
                         "leftDefaultsCol", "centerDefaultsCol", "rightDefaultsCol", _
                         "leftThicknessCels", "rightThicknessCels", _
                         "leftSecThicknessCels", "rightSecThicknessCels", "allThicknessCels")
    
    On Error Resume Next
    Dim name As Variant
    For Each name In namesToDelete
        ThisWorkbook.Names(name).Delete
    Next name
    On Error GoTo 0
    
    ' === Définition des zones principales ===
    Dim maxRange As Range
    Set maxRange = PRODUCTION_WS.Range(ROLL_START_COL & ROLL_START_ROW & ":" & ROLL_END_COL & ROLL_END_ROW)
    ThisWorkbook.Names.Add Name:="maxRollArea", RefersTo:=maxRange
    Debug.Print "[defineRollNamedRanges] -> maxRollArea : " & maxRange.Address
    
    ' === Zone active basée sur la longueur cible ===
    Dim targetLen As Long: targetLen = CLng(TARGET_LENGTH_CELL.Value)
    Dim activeRange As Range: Set activeRange = maxRange.Resize(targetLen)
    ThisWorkbook.Names.Add Name:="activeRollArea", RefersTo:=activeRange
    Debug.Print "[defineRollNamedRanges] -> activeRollArea : " & activeRange.Address
    
    ' === Zone inactive si nécessaire ===
    If targetLen < maxRange.Rows.Count Then
        Dim inactiveRange As Range
        Set inactiveRange = maxRange.Offset(targetLen).Resize(maxRange.Rows.Count - targetLen)
        ThisWorkbook.Names.Add Name:="inactiveRollArea", RefersTo:=inactiveRange
        Debug.Print "[defineRollNamedRanges] -> inactiveRollArea : " & inactiveRange.Address
    Else
        ' Toujours créer inactiveRollArea, mais la référencer à =FAUX
        ThisWorkbook.Names.Add Name:="inactiveRollArea", RefersTo:="=FAUX"
        Debug.Print "[defineRollNamedRanges] -> inactiveRollArea (FAUX)"
    End If
    
    ' === Création des ranges de colonnes ===
    Dim key As Variant, offsets As Variant, colOffset As Variant
    Dim rngUnion As Range
    For Each key In colConfig.Keys
        Set rngUnion = Nothing
        offsets = colConfig(key)
        
        For Each colOffset In offsets
            If rngUnion Is Nothing Then
                Set rngUnion = maxRange.Columns(colOffset)
            Else
                Set rngUnion = Union(rngUnion, maxRange.Columns(colOffset))
            End If
        Next colOffset
        
        ThisWorkbook.Names.Add Name:=key, RefersTo:=rngUnion
        Debug.Print "[defineRollNamedRanges] -> " & key & " : " & rngUnion.Address
    Next key
    
    ' === Création des cellules de mesure d'épaisseur ===
    Dim measureCells As Object: Set measureCells = CreateObject("Scripting.Dictionary")
    measureCells.Add "left", New Collection
    measureCells.Add "right", New Collection
    measureCells.Add "leftSec", New Collection
    measureCells.Add "rightSec", New Collection
    
    ' Détermination des lignes de mesure
    Dim measureRows As Collection: Set measureRows = New Collection
    If targetLen = 1 Then
        measureRows.Add 1 ' Officielle sur 1m
    ElseIf targetLen = 2 Then
        measureRows.Add 1 ' Officielle sur 1m
    ElseIf targetLen >= 3 Then
        Dim row As Long
        For row = ROLL_MEASURE_OFFSET To targetLen Step ROLL_MEASURE_INTERVAL
            measureRows.Add row
        Next row
    End If
    
    ' Création des cellules de mesure
    Dim mRow As Variant
    For Each mRow In measureRows
        ' Cellules principales
        For Each colOffset In colConfig("leftThicknessCols")
            measureCells("left").Add maxRange.Cells(mRow, colOffset)
        Next colOffset
        For Each colOffset In colConfig("rightThicknessCols")
            measureCells("right").Add maxRange.Cells(mRow, colOffset)
        Next colOffset

        ' Cellules de rattrapage
        If targetLen = 2 And mRow = 1 Then
            ' Pour 2m, rattrapage sur 2m
            For Each colOffset In colConfig("leftThicknessCols")
                measureCells("leftSec").Add maxRange.Cells(2, colOffset)
            Next colOffset
            For Each colOffset In colConfig("rightThicknessCols")
                measureCells("rightSec").Add maxRange.Cells(2, colOffset)
            Next colOffset
        ElseIf targetLen >= 3 Then
            If mRow + 1 <= targetLen Then
                For Each colOffset In colConfig("leftThicknessCols")
                    measureCells("leftSec").Add maxRange.Cells(mRow + 1, colOffset)
                Next colOffset
                For Each colOffset In colConfig("rightThicknessCols")
                    measureCells("rightSec").Add maxRange.Cells(mRow + 1, colOffset)
                Next colOffset
            ElseIf mRow - 1 >= 1 Then
                For Each colOffset In colConfig("leftThicknessCols")
                    measureCells("leftSec").Add maxRange.Cells(mRow - 1, colOffset)
                Next colOffset
                For Each colOffset In colConfig("rightThicknessCols")
                    measureCells("rightSec").Add maxRange.Cells(mRow - 1, colOffset)
                Next colOffset
            End If
        End If
    Next mRow
    
    ' Création des ranges finales
    Dim finalRanges As Object: Set finalRanges = CreateObject("Scripting.Dictionary")
    Dim cellType As Variant, cell As Variant
    Dim i As Long
    
    For Each cellType In Array("left", "right", "leftSec", "rightSec")
        If measureCells(cellType).Count > 0 Then
            Set rngUnion = measureCells(cellType)(1)
            For i = 2 To measureCells(cellType).Count
                Set rngUnion = Union(rngUnion, measureCells(cellType)(i))
            Next i
            finalRanges.Add cellType, rngUnion
        End If
    Next cellType
    
    ' Création des noms de range pour les cellules de mesure
    If Not finalRanges("left") Is Nothing Then
        ThisWorkbook.Names.Add Name:="leftThicknessCels", RefersTo:=finalRanges("left")
        Debug.Print "[defineRollNamedRanges] -> leftThicknessCels : " & finalRanges("left").Address
    End If
    
    If Not finalRanges("right") Is Nothing Then
        ThisWorkbook.Names.Add Name:="rightThicknessCels", RefersTo:=finalRanges("right")
        Debug.Print "[defineRollNamedRanges] -> rightThicknessCels : " & finalRanges("right").Address
    End If
    
    If finalRanges.Exists("leftSec") Then
        If Not finalRanges("leftSec") Is Nothing Then
            ThisWorkbook.Names.Add Name:="leftSecThicknessCels", RefersTo:=finalRanges("leftSec")
            Debug.Print "[defineRollNamedRanges] -> leftSecThicknessCels : " & finalRanges("leftSec").Address
        End If
    End If
    
    If finalRanges.Exists("rightSec") Then
        If Not finalRanges("rightSec") Is Nothing Then
            ThisWorkbook.Names.Add Name:="rightSecThicknessCels", RefersTo:=finalRanges("rightSec")
            Debug.Print "[defineRollNamedRanges] -> rightSecThicknessCels : " & finalRanges("rightSec").Address
        End If
    End If
    
    ' Création de la range unifiée pour toutes les cellules de mesure
    Dim allThickness As Range
    Set allThickness = Nothing
    If finalRanges.Exists("left") Then
        Set allThickness = finalRanges("left")
    End If
    If finalRanges.Exists("right") Then
        If Not finalRanges("right") Is Nothing Then
            If allThickness Is Nothing Then
                Set allThickness = finalRanges("right")
            Else
                Set allThickness = Union(allThickness, finalRanges("right"))
            End If
        End If
    End If
    If finalRanges.Exists("leftSec") Then
        If Not finalRanges("leftSec") Is Nothing Then
            If allThickness Is Nothing Then
                Set allThickness = finalRanges("leftSec")
            Else
                Set allThickness = Union(allThickness, finalRanges("leftSec"))
            End If
        End If
    End If
    If finalRanges.Exists("rightSec") Then
        If Not finalRanges("rightSec") Is Nothing Then
            If allThickness Is Nothing Then
                Set allThickness = finalRanges("rightSec")
            Else
                Set allThickness = Union(allThickness, finalRanges("rightSec"))
            End If
        End If
    End If
    If Not allThickness Is Nothing Then
        ThisWorkbook.Names.Add Name:="allThicknessCels", RefersTo:=allThickness
        Debug.Print "[defineRollNamedRanges] -> allThicknessCels : " & allThickness.Address
    End If
    
    ' Vérification que toutes les plages ont été créées
    Dim allRanges As Variant
    allRanges = Array("maxRollArea", "activeRollArea", "inactiveRollArea", _
                     "lengthCols", "leftThicknessCols", "rightThicknessCols", _
                     "leftDefaultsCol", "centerDefaultsCol", "rightDefaultsCol", _
                     "leftThicknessCels", "rightThicknessCels", _
                     "leftSecThicknessCels", "rightSecThicknessCels", "allThicknessCels")
    
    For Each name In allRanges
        On Error Resume Next
        Dim testRange As Range
        Set testRange = ThisWorkbook.Names(name).RefersToRange
        On Error GoTo 0
        
        If testRange Is Nothing Then
            Debug.Print "[defineRollNamedRanges] ATTENTION : Plage non créée : " & name
        End If
    Next name
End Sub

' Initialise les ranges nommées pour les valeurs limites de contrôle
' @pre : PRODUCTION_WS doit être initialisé
' @return : aucun
Public Sub initCtrlLimitValues()
    If PRODUCTION_WS Is Nothing Then
        Debug.Print "[initCtrlLimitValues] ERREUR : PRODUCTION_WS non initialisé"
        Exit Sub
    End If
    
    ' Suppression des ranges existantes pour éviter les doublons
    On Error Resume Next
    ThisWorkbook.Names(RANGE_CTRL_MIN_THICKNESS).Delete
    ThisWorkbook.Names(RANGE_CTRL_WARN_THICKNESS).Delete
    On Error GoTo 0
    
    ' Création des nouvelles ranges
    ThisWorkbook.Names.Add Name:=RANGE_CTRL_MIN_THICKNESS, RefersTo:=PRODUCTION_WS.Range("BK59")  ' Seuil rouge
    ThisWorkbook.Names.Add Name:=RANGE_CTRL_WARN_THICKNESS, RefersTo:=PRODUCTION_WS.Range("BJ59")  ' Seuil orange
    
    ' Affichage des adresses des ranges
    Debug.Print "[initCtrlLimitValues] -> " & RANGE_CTRL_MIN_THICKNESS & " : BK59 (seuil rouge)"
    Debug.Print "[initCtrlLimitValues] -> " & RANGE_CTRL_WARN_THICKNESS & " : BJ59 (seuil orange)"
End Sub

' Initialise les ranges nommées pour les numéros OF
' @pre : PRODUCTION_WS doit être initialisé
' @return : aucun
Public Sub initOFRanges()
    If PRODUCTION_WS Is Nothing Then
        Debug.Print "[initOFRanges] ERREUR : PRODUCTION_WS non initialisé"
        Exit Sub
    End If
    
    ' Suppression des ranges existantes pour éviter les doublons
    On Error Resume Next
    ThisWorkbook.Names(RANGE_OF_NUMBER).Delete
    ThisWorkbook.Names(RANGE_CUT_OF_NUMBER).Delete
    On Error GoTo 0
    
    ' Création des nouvelles ranges
    ThisWorkbook.Names.Add Name:=RANGE_OF_NUMBER, RefersTo:=PRODUCTION_WS.Range("BH69")
    ThisWorkbook.Names.Add Name:=RANGE_CUT_OF_NUMBER, RefersTo:=PRODUCTION_WS.Range("BH73")
    
    ' Affichage des adresses des ranges
    Debug.Print "[initOFRanges] -> " & RANGE_OF_NUMBER & " : BH69"
    Debug.Print "[initOFRanges] -> " & RANGE_CUT_OF_NUMBER & " : BH73"
End Sub

