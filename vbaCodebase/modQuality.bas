Attribute VB_Name = "modQuality"
Option Explicit

' V�rifie la conformit� du rouleau uniquement sur les d�fauts (pas l'�paisseur)
' @but : Retourne True si le rouleau est conforme sur les d�fauts, False sinon. Motif en out param.
' @param motif (ByRef, optionnel) : cha�ne de motif de non-conformit�
' @return Boolean : True si conforme, False sinon
' @pr� : PRODUCTION_WS doit �tre initialis� et les plages de d�fauts doivent exister
Public Function IsRollConformDefects(Optional ByRef motif As String = "") As Boolean
    Dim ws As Worksheet: Set ws = PRODUCTION_WS
    If ws Is Nothing Then motif = "Feuille PROD non initialis�e": IsRollConformDefects = False: Exit Function

    ' D�verrouiller la feuille si besoin
    Dim wasProtected As Boolean: wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect

    ' Dictionnaire pour compter les d�fauts
    Dim defCounts As Object: Set defCounts = CreateObject("Scripting.Dictionary")
    Dim defMax As Object: Set defMax = CreateObject("Scripting.Dictionary")
    
    ' Liste des d�fauts � contr�ler et leur seuil (lecture dynamique depuis le tableau)
    Dim iRow As Long
    For iRow = 54 To 59
        Dim nomDefaut As String
        nomDefaut = Trim(ws.Range("BG" & iRow).Value)
        Dim seuil As Variant
        seuil = ws.Range("BH" & iRow).Value
        If nomDefaut <> "" And seuil <> "-" And Not IsEmpty(seuil) Then
            defMax.Add nomDefaut, seuil
        End If
    Next iRow
    
    ' Initialiser les compteurs
    Dim defName As Variant
    For Each defName In defMax.Keys
        defCounts(defName) = 0
    Next defName
    
    ' Parcourir la zone active des colonnes d�fauts
    Dim rngActive As Range: Set rngActive = ws.Range(ThisWorkbook.Names("activeRollArea").RefersTo)
    Dim defCols As Variant: defCols = Array("leftDefaultsCol", "centerDefaultsCol", "rightDefaultsCol")
    Dim rngDef As Range, cell As Range
    Dim i As Integer
    For i = LBound(defCols) To UBound(defCols)
        Debug.Print "[IsRollConformDefects] Test colonne d�faut : " & defCols(i)
        If NameExists(CStr(defCols(i))) Then
            Debug.Print "[IsRollConformDefects]   -> Nom existe : " & defCols(i)
            Set rngDef = Application.Intersect(ws.Range(ThisWorkbook.Names(defCols(i)).RefersTo), rngActive)
            If Not rngDef Is Nothing Then
                Debug.Print "[IsRollConformDefects]   -> Intersect OK : " & rngDef.Address
                For Each cell In rngDef.Cells
                    If Trim(cell.Value) <> "" Then
                        Debug.Print "[IsRollConformDefects]     Cellule : " & cell.Address & " - Valeur : " & cell.Value
                        If defMax.Exists(cell.Value) Then
                            defCounts(cell.Value) = defCounts(cell.Value) + 1
                            Debug.Print "[IsRollConformDefects]     -> Compt� : " & cell.Value & " (total : " & defCounts(cell.Value) & ")"
                        End If
                    End If
                Next cell
            Else
                Debug.Print "[IsRollConformDefects]   -> Intersect = Nothing pour " & defCols(i)
            End If
        Else
            Debug.Print "[IsRollConformDefects]   -> Nom NON trouv� : " & defCols(i)
        End If
    Next i
    
    ' V�rifier la conformit�
    motif = ""
    Dim isConform As Boolean: isConform = True
    For Each defName In defMax.Keys
        If defCounts(defName) > defMax(defName) Then
            motif = motif & defName & " : " & defCounts(defName) & " (max " & defMax(defName) & ") | " 
            isConform = False
        End If
    Next defName
    If motif <> "" Then
        motif = "D�fauts d�passant le seuil : " & motif
    End If
    ws.Range("BK85").Value = isConform
    If wasProtected Then ws.Protect
    IsRollConformDefects = isConform
End Function

' Sauvegarde les d�fauts d�tect�s dans la feuille de production
' @but : Parcourt les colonnes de d�fauts et sauvegarde les d�fauts d�tect�s dans une cellule d�di�e
' @param Aucun
' @return Aucun
' @pr� : PRODUCTION_WS doit �tre initialis� et les plages de d�fauts doivent exister
Public Sub saveDetectedDefects()
    Dim ws As Worksheet: Set ws = PRODUCTION_WS
    If ws Is Nothing Then Exit Sub

    ' D�verrouiller la feuille si besoin
    Dim wasProtected As Boolean: wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect

    Dim rngActive As Range: Set rngActive = ws.Range(ThisWorkbook.Names("activeRollArea").RefersTo)
    Dim defCols As Variant: defCols = Array("leftDefaultsCol", "centerDefaultsCol", "rightDefaultsCol")
    Dim positions As Variant: positions = Array("Gauche", "Centre", "Droite")
    Dim i As Integer, rowOffset As Long
    Dim rngDef As Range, cell As Range
    Dim defectsList As Collection: Set defectsList = New Collection

    For i = LBound(defCols) To UBound(defCols)
        If NameExists(CStr(defCols(i))) Then
            Set rngDef = Application.Intersect(ws.Range(ThisWorkbook.Names(defCols(i)).RefersTo), rngActive)
            If Not rngDef Is Nothing Then
                For Each cell In rngDef.Cells
                    If Trim(cell.Value) <> "" Then
                        rowOffset = cell.Row - rngActive.Rows(1).Row + 1
                        defectsList.Add positions(i) & " " & rowOffset & "m " & cell.Value
                    End If
                Next cell
            End If
        End If
    Next i

    Dim result As String: result = ""
    Dim d As Variant
    For Each d In defectsList
        If result <> "" Then result = result & " / "
        result = result & d
    Next d
    ws.Range("BG85").Value = result
    If wasProtected Then ws.Protect
End Sub

' V�rifie la conformit� du rouleau sur l'�paisseur
' @but : Retourne True si le rouleau est conforme sur l'�paisseur, False sinon. Motif en out param.
' @param motif (ByRef, optionnel) : cha�ne de motif de non-conformit�
' @return Boolean : True si conforme, False sinon
' @pr� : PRODUCTION_WS doit �tre initialis� et les plages d'�paisseur doivent exister
Public Function IsRollConformThickness(Optional ByRef motif As String = "") As Boolean
    Dim ws As Worksheet: Set ws = PRODUCTION_WS
    If ws Is Nothing Then motif = "Feuille PROD non initialis�e": IsRollConformThickness = False: Exit Function

    ' D�verrouiller la feuille si besoin
    Dim wasProtected As Boolean: wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect

    Dim rngActive As Range: Set rngActive = ws.Range(ThisWorkbook.Names("activeRollArea").RefersTo)
    Dim thickNames As Variant: thickNames = Array("leftThicknessCels", "rightThicknessCels")
    Dim positions As Variant: positions = Array("Gauche", "Droite")
    Dim secNames As Variant: secNames = Array("leftSecThicknessCels", "rightSecThicknessCels")
    Dim i As Integer, rowOffset As Long
    Dim rngThick As Range, cell As Range
    Dim NOKBloquant As Integer: NOKBloquant = 0
    Dim motifList As String: motifList = ""
    Dim ctrlMin As Double: ctrlMin = CDbl(ws.Range("ctrlMinThickness").Value)

    Dim isConform As Boolean: isConform = True
    For i = LBound(thickNames) To UBound(thickNames)
        Debug.Print "[IsRollConformThickness] Test plage : " & thickNames(i)
        If NameExists(CStr(thickNames(i))) Then
            Dim refString As String
            refString = ThisWorkbook.Names(thickNames(i)).RefersTo
            refString = Replace(refString, "=", "")
            Dim addresses As Variant
            addresses = Split(refString, ",")
            Dim addr As Variant
            For Each addr In addresses
                Set cell = ws.Range(addr)
                If IsNumeric(cell.Value) And cell.Value <> "" And CDbl(cell.Value) < ctrlMin Then
                    rowOffset = cell.Row - rngActive.Rows(1).Row + 1
                    ' Chercher la cellule de rattrapage
                    Dim rattrapageCell As Range
                    Dim isLastRow As Boolean
                    isLastRow = cell.Row = rngActive.Rows(rngActive.Rows.Count).Row
                    If isLastRow Then
                        Set rattrapageCell = cell.Offset(-1, 0)
                    Else
                        Set rattrapageCell = cell.Offset(1, 0)
                    End If
                    Dim isBloquant As Boolean: isBloquant = False
                    If NameExists(CStr(secNames(i))) Then
                        Dim foundInRattrapage As Boolean: foundInRattrapage = False
                        Dim refStringR As String
                        refStringR = ThisWorkbook.Names(secNames(i)).RefersTo
                        refStringR = Replace(refStringR, "=", "")
                        Dim addressesR As Variant
                        addressesR = Split(refStringR, ",")
                        Dim addrR As Variant
                        For Each addrR In addressesR
                            If rattrapageCell.Address = ws.Range(addrR).Address And rattrapageCell.Worksheet.Name = ws.Range(addrR).Worksheet.Name Then
                                foundInRattrapage = True
                                Exit For
                            End If
                        Next addrR
                        If foundInRattrapage Then
                            If IsNumeric(rattrapageCell.Value) And CDbl(rattrapageCell.Value) < ctrlMin Then
                                isBloquant = True
                                isConform = False ' Paire NOK trouv�e
                            End If
                        Else
                            isBloquant = True
                            isConform = False ' Paire NOK trouv�e (pas de rattrapage)
                        End If
                    Else
                        isBloquant = True
                        isConform = False ' Paire NOK trouv�e (pas de plage de rattrapage)
                    End If
                    If isBloquant Then
                        NOKBloquant = NOKBloquant + 1
                        If motifList <> "" Then motifList = motifList & " / "
                        motifList = motifList & positions(i) & " " & rowOffset & "m NOK=" & Format(cell.Value, "0.00")
                    End If
                End If
            Next addr
        Else
            Debug.Print "[IsRollConformThickness]   -> Nom NON trouv� : " & thickNames(i)
        End If
    Next i
    If NOKBloquant > 3 Then isConform = False
    motif = motifList
    ws.Range("BK86").Value = isConform
    If wasProtected Then ws.Protect
    IsRollConformThickness = isConform
End Function

' Sauvegarde les �paisseurs d�tect�es dans la feuille de production
' @but : Parcourt les colonnes d'�paisseur et sauvegarde les valeurs NOK dans une cellule d�di�e
' @param Aucun
' @return Aucun
' @pr� : PRODUCTION_WS doit �tre initialis� et les plages d'�paisseur doivent exister
Public Sub saveDetectedThickness()
    Dim ws As Worksheet: Set ws = PRODUCTION_WS
    If ws Is Nothing Then Exit Sub

    ' D�verrouiller la feuille si besoin
    Dim wasProtected As Boolean: wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect

    Dim rngActive As Range: Set rngActive = ws.Range(ThisWorkbook.Names("activeRollArea").RefersTo)
    Dim thickNames As Variant: thickNames = Array("leftThicknessCels", "rightThicknessCels")
    Dim positions As Variant: positions = Array("Gauche", "Droite")
    Dim secNames As Variant: secNames = Array("leftSecThicknessCels", "rightSecThicknessCels")
    Dim i As Integer, rowOffset As Long
    Dim NOKList As Collection: Set NOKList = New Collection
    Dim ctrlMin As Double: ctrlMin = CDbl(ws.Range("ctrlMinThickness").Value)
    Dim rngThick As Range, cell As Range

    For i = LBound(thickNames) To UBound(thickNames)
        If NameExists(CStr(thickNames(i))) Then
            Dim refString As String
            refString = ThisWorkbook.Names(thickNames(i)).RefersTo
            refString = Replace(refString, "=", "")
            Dim addresses As Variant
            addresses = Split(refString, ",")
            Dim addr As Variant
            For Each addr In addresses
                Set cell = ws.Range(addr)
                If IsNumeric(cell.Value) And cell.Value <> "" And CDbl(cell.Value) < ctrlMin Then
                    rowOffset = cell.Row - rngActive.Rows(1).Row + 1
                    Dim txt As String: txt = positions(i) & " " & rowOffset & "m " & Format(cell.Value, "0.00")
                    ' Chercher la cellule de rattrapage
                    Dim rattrapageCell As Range
                    Dim isLastRow As Boolean
                    isLastRow = cell.Row = rngActive.Rows(rngActive.Rows.Count).Row
                    If isLastRow Then
                        Set rattrapageCell = cell.Offset(-1, 0)
                    Else
                        Set rattrapageCell = cell.Offset(1, 0)
                    End If
                    Dim foundInRattrapage As Boolean: foundInRattrapage = False
                    If NameExists(CStr(secNames(i))) Then
                        Dim refStringR As String
                        refStringR = ThisWorkbook.Names(secNames(i)).RefersTo
                        refStringR = Replace(refStringR, "=", "")
                        Dim addressesR As Variant
                        addressesR = Split(refStringR, ",")
                        Dim addrR As Variant
                        For Each addrR In addressesR
                            If rattrapageCell.Address = ws.Range(addrR).Address And rattrapageCell.Worksheet.Name = ws.Range(addrR).Worksheet.Name Then
                                foundInRattrapage = True
                                txt = txt & " | " & Format(rattrapageCell.Value, "0.00")
                                Exit For
                            End If
                        Next addrR
                    End If
                    NOKList.Add txt
                End If
            Next addr
        End If
    Next i

    Dim result As String: result = ""
    Dim d As Variant
    For Each d In NOKList
        If result <> "" Then result = result & " / "
        result = result & d
    Next d
    ws.Range("BG86").Value = result
    If wasProtected Then ws.Protect
End Sub

