Attribute VB_Name = "LOADER"

Option Explicit

Const THIS_MODULE As String = "LOADER"
Const vbaCodebaseDir As String = "vbaCodebase"
Const vbaToInjectDir As String = "vbaCodebase\toInject"



' Charge tous les modules .bas et .cls depuis le dossier vbaCodebase
' @pre : le dossier doit exister � c�t� du classeur Excel
' @return : aucun
' Charge tous les modules (.bas, .cls) depuis le dossier /vbaCodebase
' Supprime tous les composants VBA sauf ce module (THIS_MODULE)
Public Sub loadModulesFromFolder()

    Dim path As String
    path = ThisWorkbook.Path & Application.PathSeparator & "vbaCodebase" & Application.PathSeparator

    Dim fileSystem As Object: Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Dim file As Object
    Dim moduleFolder As Object: Set moduleFolder = fileSystem.GetFolder(path)

    ' === Supprimer tous les modules sauf celui-ci ===
    Dim vbComp As VBIDE.VBComponent
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Name <> THIS_MODULE Then
            On Error Resume Next
            ThisWorkbook.VBProject.VBComponents.Remove vbComp
            Debug.Print "[loadModulesFromFolder] removed -> " & vbComp.Name
            On Error GoTo 0
        End If
    Next vbComp

    ' === Importer tous les .bas et .cls ===
    Dim ext As String
    For Each file In moduleFolder.Files
        ext = LCase(fileSystem.GetExtensionName(file.Name))
        If ext = "bas" Or ext = "cls" Then
            ThisWorkbook.VBProject.VBComponents.Import file.Path
            Debug.Print "[loadModulesFromFolder] imported -> " & file.Name
        End If
    Next file

End Sub


' Importe tous les modules .bas pr�sents dans /vbaCodebase/toInject/
' - Injecte ThisWorkbook.bas dans le module objet ThisWorkbook
' - Les autres .bas sont import�s normalement (remplacement si existe)
Public Sub injectModulesFromToInject()
    Dim vbProj As VBIDE.VBProject: Set vbProj = ThisWorkbook.VBProject
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dossier As String: dossier = ThisWorkbook.Path & "\vbaCodebase\toInject\"

    If Not fso.FolderExists(dossier) Then
        MsgBox "Dossier introuvable : " & dossier, vbCritical
        Exit Sub
    End If

    ' === Construction dynamique : nom onglet ? CodeName
    Dim sheetMap As Object: Set sheetMap = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        sheetMap.Add LCase(ws.Name), ws.CodeName
    Next ws

    ' === Traitement des fichiers .bas
    Dim fichier As Object
    Dim baseName As String
    Dim cheminFichier As String
    Dim targetModule As String

    For Each fichier In fso.GetFolder(dossier).Files
        If LCase(fso.GetExtensionName(fichier.Name)) = "bas" Then
            baseName = LCase(fso.GetBaseName(fichier.Name))
            cheminFichier = fichier.Path

            Select Case baseName
                Case "thisworkbook"
                    Call injectCodeFromBasFile(cheminFichier, "ThisWorkbook")

                Case Else
                    If sheetMap.Exists(baseName) Then
                        targetModule = sheetMap(baseName)
                        Call injectCodeFromBasFile(cheminFichier, targetModule)
                    Else
                        ' Fichier .bas classique : suppression + import
                        On Error Resume Next
                        vbProj.VBComponents.Remove vbProj.VBComponents(baseName)
                        On Error GoTo 0
                        vbProj.VBComponents.Import cheminFichier
                        Debug.Print "[injectModulesFromToInject] import� -> " & fichier.Name
                    End If
            End Select
        End If
    Next fichier

    MsgBox "Injection termin�e -> modules charg�s depuis toInject", vbInformation
End Sub


' Support
Public Sub injectCodeFromBasFile(cheminFichier As String, targetModule As String)
    Dim vbComp As VBIDE.VBComponent
    Set vbComp = ThisWorkbook.VBProject.VBComponents(targetModule)
    If vbComp Is Nothing Then Exit Sub
    vbComp.CodeModule.DeleteLines 1, vbComp.CodeModule.CountOfLines
    vbComp.CodeModule.AddFromFile cheminFichier
End Sub