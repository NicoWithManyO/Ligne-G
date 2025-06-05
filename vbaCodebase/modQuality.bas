Attribute VB_Name = "modQuality"
Option Explicit

' Vérifie la conformité du rouleau uniquement sur les défauts (pas l'épaisseur)
' @but : Retourne True si le rouleau est conforme sur les défauts, False sinon. Motif en out param.
' @param motif (ByRef, optionnel) : chaîne de motif de non-conformité
' @return Boolean : True si conforme, False sinon
' @pré : PRODUCTION_WS doit être initialisé et les plages de défauts doivent exister
Public Function IsRollConformDefects(Optional ByRef motif As String = "") As Boolean
    Dim ws As Worksheet: Set ws = PRODUCTION_WS
    If ws Is Nothing Then motif = "Feuille PROD non initialisée": IsRollConformDefects = False: Exit Function

    ' Déverrouiller la feuille si besoin
    Dim wasProtected As Boolean: wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect

    ' Dictionnaire pour compter les défauts
    Dim defCounts As Object: Set defCounts = CreateObject("Scripting.Dictionary")
    Dim defMax As Object: Set defMax = CreateObject("Scripting.Dictionary")
    
    ' Liste des défauts à contrôler et leur seuil (lecture dynamique depuis le tableau)
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
    
    ' Parcourir la zone active des colonnes défauts
    Dim rngActive As Range: Set rngActive = ws.Range(ThisWorkbook.Names("activeRollArea").RefersTo)
    Dim defCols As Variant: defCols = Array("leftDefaultsCol", "centerDefaultsCol", "rightDefaultsCol")
    Dim rngDef As Range, cell As Range
    Dim i As Integer
    For i = LBound(defCols) To UBound(defCols)
        Debug.Print "[IsRollConformDefects] Test colonne défaut : " & defCols(i)
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
                            Debug.Print "[IsRollConformDefects]     -> Compté : " & cell.Value & " (total : " & defCounts(cell.Value) & ")"
                        End If
                    End If
                Next cell
            Else
                Debug.Print "[IsRollConformDefects]   -> Intersect = Nothing pour " & defCols(i)
            End If
        Else
            Debug.Print "[IsRollConformDefects]   -> Nom NON trouvé : " & defCols(i)
        End If
    Next i
    
    ' Vérifier la conformité
    motif = ""
    Dim isConform As Boolean: isConform = True
    For Each defName In defMax.Keys
        If defCounts(defName) > defMax(defName) Then
            motif = motif & defName & " : " & defCounts(defName) & " (max " & defMax(defName) & ") | " 
            isConform = False
        End If
    Next defName
    If motif <> "" Then
        motif = "Défauts dépassant le seuil : " & motif
    End If
    ws.Range("BK85").Value = isConform
    If wasProtected Then ws.Protect
    IsRollConformDefects = isConform
End Function

' Sauvegarde les défauts détectés dans la feuille de production
' @but : Parcourt les colonnes de défauts et sauvegarde les défauts détectés dans une cellule dédiée
' @param Aucun
' @return Aucun
' @pré : PRODUCTION_WS doit être initialisé et les plages de défauts doivent exister
Public Sub saveDetectedDefects()
    Dim ws As Worksheet: Set ws = PRODUCTION_WS
    If ws Is Nothing Then Exit Sub

    ' Déverrouiller la feuille si besoin
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

' Vérifie la conformité du rouleau sur l'épaisseur
' @but : Retourne True si le rouleau est conforme sur l'épaisseur, False sinon. Motif en out param.
' @param motif (ByRef, optionnel) : chaîne de motif de non-conformité
' @return Boolean : True si conforme, False sinon
' @pré : PRODUCTION_WS doit être initialisé et les plages d'épaisseur doivent exister
Public Function IsRollConformThickness(Optional ByRef motif As String = "") As Boolean
    Dim ws As Worksheet: Set ws = PRODUCTION_WS
    If ws Is Nothing Then motif = "Feuille PROD non initialisée": IsRollConformThickness = False: Exit Function

    ' Déverrouiller la feuille si besoin
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
                                isConform = False ' Paire NOK trouvée
                            End If
                        Else
                            isBloquant = True
                            isConform = False ' Paire NOK trouvée (pas de rattrapage)
                        End If
                    Else
                        isBloquant = True
                        isConform = False ' Paire NOK trouvée (pas de plage de rattrapage)
                    End If
                    If isBloquant Then
                        NOKBloquant = NOKBloquant + 1
                        If motifList <> "" Then motifList = motifList & " / "
                        motifList = motifList & positions(i) & " " & rowOffset & "m NOK=" & Format(cell.Value, "0.00")
                    End If
                End If
            Next addr
        Else
            Debug.Print "[IsRollConformThickness]   -> Nom NON trouvé : " & thickNames(i)
        End If
    Next i
    If NOKBloquant > 3 Then isConform = False
    motif = motifList
    ws.Range("BK86").Value = isConform
    If wasProtected Then ws.Protect
    IsRollConformThickness = isConform
End Function

' Sauvegarde les épaisseurs détectées dans la feuille de production
' @but : Parcourt les colonnes d'épaisseur et sauvegarde les valeurs NOK dans une cellule dédiée
' @param Aucun
' @return Aucun
' @pré : PRODUCTION_WS doit être initialisé et les plages d'épaisseur doivent exister
Public Sub saveDetectedThickness()
    Dim ws As Worksheet: Set ws = PRODUCTION_WS
    If ws Is Nothing Then Exit Sub

    ' Déverrouiller la feuille si besoin
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

' Charge toutes les épaisseurs dans un tableau
' @but : Parcourt les colonnes d'épaisseur et stocke toutes les valeurs dans un tableau
' @param Aucun
' @return Object : Dictionary contenant les épaisseurs groupées par position
' @pré : PRODUCTION_WS doit être initialisé et les plages d'épaisseur doivent exister
Public Function LoadAllThicknesses() As Object
    Dim ws As Worksheet: Set ws = PRODUCTION_WS
    If ws Is Nothing Then Set LoadAllThicknesses = CreateObject("Scripting.Dictionary"): Exit Function

    ' Déverrouiller la feuille si besoin
    Dim wasProtected As Boolean: wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect

    Dim rngActive As Range: Set rngActive = ws.Range(ThisWorkbook.Names("activeRollArea").RefersTo)
    Dim thickNames As Variant: thickNames = Array("leftThicknessCels", "rightThicknessCels")
    Dim positions As Variant: positions = Array("Gauche", "Droite")
    Dim secNames As Variant: secNames = Array("leftSecThicknessCels", "rightSecThicknessCels")
    Dim i As Integer, rowOffset As Long
    Dim cell As Range
    
    ' Créer la structure de données principale
    Dim thicknessData As Object: Set thicknessData = CreateObject("Scripting.Dictionary")
    thicknessData.Add "Gauche", New Collection
    thicknessData.Add "Droite", New Collection

    For i = LBound(thickNames) To UBound(thickNames)
        Debug.Print "[LoadAllThicknesses] Test plage : " & thickNames(i)
        If NameExists(CStr(thickNames(i))) Then
            Dim refString As String
            refString = ThisWorkbook.Names(thickNames(i)).RefersTo
            refString = Replace(refString, "=", "")
            Dim addresses As Variant
            addresses = Split(refString, ",")
            Dim addr As Variant
            For Each addr In addresses
                Set cell = ws.Range(addr)
                rowOffset = cell.Row - rngActive.Rows(1).Row + 1
                Dim thicknessInfo As Object: Set thicknessInfo = CreateObject("Scripting.Dictionary")
                thicknessInfo.Add "rowOffset", rowOffset
                thicknessInfo.Add "value", cell.Value ' Peut être vide ou non numérique
                thicknessInfo.Add "isConform", True
                
                ' Chercher la cellule de rattrapage
                Dim rattrapageCell As Range
                Dim isLastRow As Boolean
                isLastRow = cell.Row = rngActive.Rows(rngActive.Rows.Count).Row
                If isLastRow Then
                    Set rattrapageCell = cell.Offset(-1, 0)
                Else
                    Set rattrapageCell = cell.Offset(1, 0)
                End If
                
                ' Vérifier si la cellule de rattrapage existe dans la plage secondaire
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
                            thicknessInfo.Add "rattrapageValue", rattrapageCell.Value ' Peut être vide ou non numérique
                            Exit For
                        End If
                    Next addrR
                End If
                
                ' Ajouter à la collection appropriée
                thicknessData(positions(i)).Add thicknessInfo
            Next addr
        Else
            Debug.Print "[LoadAllThicknesses]   -> Nom NON trouvé : " & thickNames(i)
        End If
    Next i

    If wasProtected Then ws.Protect
    Set LoadAllThicknesses = thicknessData
End Function

' Procédure de test pour LoadAllThicknesses
' @but : Afficher toutes les épaisseurs trouvées dans la fenêtre de débogage
Public Sub TestLoadAllThicknesses()
    Dim thicknesses As Object
    Set thicknesses = LoadAllThicknesses()
    
    Debug.Print "=== Test de LoadAllThicknesses ==="
    
    ' Afficher les épaisseurs par position
    Dim positions As Variant: positions = Array("Gauche", "Droite")
    Dim pos As Variant
    
    For Each pos In positions
        Debug.Print "=== Épaisseurs " & pos & " ==="
        Dim thickness As Object
        For Each thickness In thicknesses(pos)
            Debug.Print "  Offset : " & thickness("rowOffset") & "m"
            Debug.Print "  Valeur : " & Format(thickness("value"), "0.00")
            If thickness.Exists("rattrapageValue") Then
                Debug.Print "  Rattrapage : " & Format(thickness("rattrapageValue"), "0.00")
            End If
            Debug.Print "  ---"
        Next thickness
    Next pos
    
    Debug.Print "=== Fin du test ==="
End Sub

' Vérifie que toutes les épaisseurs sont présentes en fonction de la longueur
' @but : Vérifie que toutes les mesures d'épaisseur requises sont présentes pour chaque ligne
' @param ByRef missingMeasurements As String : Liste des mesures manquantes
' @return Boolean : True si toutes les mesures sont présentes, False sinon
' @pré : PRODUCTION_WS doit être initialisé
Public Function AreAllThicknessesPresent(ByRef missingMeasurements As String) As Boolean
    Dim ws As Worksheet: Set ws = PRODUCTION_WS
    
    ' Essayer d'abord d'utiliser la longueur réelle (BH82)
    Dim realLength As Double
    realLength = ws.Range(RANGE_REAL_LENGTH).Value
    
    ' Si pas de longueur réelle, utiliser la longueur cible
    If Not IsNumeric(realLength) Or realLength <= 0 Then
        realLength = ws.Range(TARGET_LENGTH_ADDR).Value
        Debug.Print "[AreAllThicknessesPresent] Utilisation de la longueur cible = " & realLength & "m"
    Else
        Debug.Print "[AreAllThicknessesPresent] Utilisation de la longueur réelle = " & realLength & "m"
    End If
    
    ' Vérifier que la longueur est valide
    If Not IsNumeric(realLength) Or realLength <= 0 Then
        Debug.Print "[AreAllThicknessesPresent] ERREUR : Longueur invalide = " & realLength
        missingMeasurements = "Longueur invalide"
        AreAllThicknessesPresent = False
        Exit Function
    End If
    
    ' Calculer les positions des mesures requises
    Dim requiredMeasurements As Collection: Set requiredMeasurements = New Collection
    Dim currentPos As Double: currentPos = ROLL_MEASURE_OFFSET
    
    ' Cas spéciaux pour les rouleaux courts
    If realLength = 1 Then
        Debug.Print "[AreAllThicknessesPresent] Rouleau de 1m : mesure à 1m"
        requiredMeasurements.Add 1
    ElseIf realLength = 2 Then
        Debug.Print "[AreAllThicknessesPresent] Rouleau de 2m : mesure à 1m"
        requiredMeasurements.Add 1
    Else
        Debug.Print "[AreAllThicknessesPresent] Rouleau de " & realLength & "m : mesures tous les " & ROLL_MEASURE_INTERVAL & "m à partir de " & ROLL_MEASURE_OFFSET & "m"
        Do While currentPos <= realLength
            requiredMeasurements.Add currentPos
            Debug.Print "[AreAllThicknessesPresent]   -> Ajout position " & currentPos & "m"
            currentPos = currentPos + ROLL_MEASURE_INTERVAL
        Loop
    End If

    ' Vérifier les mesures présentes
    Dim missingList As String: missingList = ""
    Dim positions As Variant: positions = Array("Gauche", "Droite")
    Dim posName As Variant
    Dim hasAllMeasurements As Boolean: hasAllMeasurements = True

    Dim ctrlMin As Double
    ctrlMin = CDbl(ws.Range("ctrlMinThickness").Value)

    ' Pour chaque position (Gauche et Droite)
    For Each posName In positions
        Dim missingForPosition As String: missingForPosition = ""
        Dim thickRange As Range
        Dim secRange As Range

        If posName = "Gauche" Then
            Set thickRange = ws.Range("leftThicknessCels")
            Set secRange = ws.Range("leftSecThicknessCels")
        Else
            Set thickRange = ws.Range("rightThicknessCels")
            Set secRange = ws.Range("rightSecThicknessCels")
        End If

        Dim pos As Variant
        For Each pos In requiredMeasurements
            Dim allColsOK As Boolean: allColsOK = True
            Dim colIdx As Integer: colIdx = 0

            ' Pour chaque colonne de la ligne concernée
            Dim cell As Range
            For Each cell In thickRange.Cells
                If cell.Row - ROLL_START_ROW + 1 = pos Then
                    colIdx = colIdx + 1
                    Dim val As Variant
                    val = cell.Value

                    If val = "" Or Not IsNumeric(val) Then
                        allColsOK = False
                    ElseIf val < ctrlMin Then
                        ' Chercher la cellule de rattrapage correspondante (même colonne, même ligne)
                        Dim secCell As Range
                        Dim sc As Range
                        Set secCell = Nothing
                        For Each sc In secRange.Cells
                            If (sc.Row = cell.Row + 1 Or sc.Row = cell.Row - 1) And sc.Column = cell.Column Then
                                Set secCell = sc
                                Exit For
                            End If
                        Next sc
                        If secCell Is Nothing Then
                            allColsOK = False
                        ElseIf secCell.Value = "" Then
                            allColsOK = False
                        End If
                    End If
                End If
            Next cell

            If Not allColsOK Then
                If missingForPosition <> "" Then missingForPosition = missingForPosition & ", "
                missingForPosition = missingForPosition & pos & "m"
            End If
        Next pos

        If missingForPosition <> "" Then
            hasAllMeasurements = False
            If missingList <> "" Then missingList = missingList & " | "
            missingList = missingList & posName & " : " & missingForPosition
        End If
    Next posName

    ' Mettre à jour le paramètre de sortie
    missingMeasurements = missingList
    
    ' Retourner True si toutes les mesures sont présentes pour toutes les positions
    Debug.Print "[AreAllThicknessesPresent] Résultat final : " & IIf(hasAllMeasurements, "Toutes les mesures sont présentes", "Mesures manquantes : " & missingList)
    AreAllThicknessesPresent = hasAllMeasurements
End Function

' Procédure de test simple pour AreAllThicknessesPresent
' @but : Tester la fonction AreAllThicknessesPresent sur la feuille courante
' @param Aucun
' @return Aucun
Public Sub TestAreAllThicknessesPresent()
    Dim missing As String
    Dim result As Boolean
    result = AreAllThicknessesPresent(missing)
    If result Then
        MsgBox "Toutes les épaisseurs requises sont présentes.", vbInformation
    Else
        MsgBox "Mesures d'épaisseur manquantes : " & missing, vbExclamation
    End If
End Sub



