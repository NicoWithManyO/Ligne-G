Attribute VB_Name = "GLOBALS"

' === GLOBALS.bas ===
Option Explicit
Public Const TO_INJECT_PATH As String = "vbaCodebase\toInject\"

' === Feuille principale de production
    Public PRODUCTION_WS As Worksheet

' === Adresse de la zone maximale du rouleau
Public Const MAX_ROLL_ADDR As String = "AJ68:BD167"

' === Adresse de la cellule contenant la longueur cible du rouleau
Public Const TARGET_LENGTH_ADDR As String = "BH71"
' === Objet Range pointant vers cette cellule
Public TARGET_LENGTH_CELL As Range

' === Ranges Shift ===
Public Const RANGE_SHIFT_ID As String = "shiftID"
Public Const RANGE_SHIFT_DATE As String = "shiftDate"
Public Const RANGE_SHIFT_OPERATEUR As String = "shiftOperateur"
Public Const RANGE_SHIFT_VACCATION As String = "shiftVaccation"
Public Const RANGE_SHIFT_DUREE As String = "shiftDuree"
Public Const RANGE_SHIFT_MACHINE_PRISE_POSTE As String = "shiftMachinePrisePoste"
Public Const RANGE_SHIFT_LG_ENROULEE_PRISE_POSTE As String = "shiftLgEnrouleePrisePoste"
Public Const RANGE_SHIFT_MACHINE_FIN_POSTE As String = "shiftMachineFinPoste"
Public Const RANGE_SHIFT_LG_ENROULEE_FIN_POSTE As String = "shiftLgEnrouleeFinPoste"
Public Const RANGE_SHIFT_COMMENTAIRES As String = "shiftCommentaires"

' === Ranges OF ===
Public Const RANGE_OF_NUMBER As String = "OFNumber"
Public Const RANGE_CUT_OF_NUMBER As String = "CutOFNumber"

' === Ranges Rouleau ===
Public Const ROLL_START_ROW As Long = 68
Public Const ROLL_END_ROW As Long = 167
Public Const ROLL_START_COL As String = "AJ"
Public Const ROLL_END_COL As String = "BD"
Public Const ROLL_MEASURE_INTERVAL As Long = 5
Public Const ROLL_MEASURE_OFFSET As Long = 3

' === Ranges Limites ===
Public Const RANGE_CTRL_MIN_THICKNESS As String = "ctrlMinThickness"
Public Const RANGE_CTRL_WARN_THICKNESS As String = "ctrlWarnThickness"

